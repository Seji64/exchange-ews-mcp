using System.Net;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using ExchangeEwsMcp.Auth;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.DependencyInjection;

namespace ExchangeEwsMcp.Tests;

public sealed partial class OAuthFlowTests : IDisposable
{
    private const string RedirectUri = "https://openwebui.test/oauth/callback";
    private readonly ServerFactory _factory = new();
    private HttpClient? _client;

    // Created on first use so a test can change settings first.
    private HttpClient _http => _client ??= _factory.CreateClient(new WebApplicationFactoryClientOptions { AllowAutoRedirect = false, HandleCookies = true });

    public void Dispose()
    {
        _client?.Dispose();
        _factory.Dispose();
    }

    [Fact]
    public async Task Unauthenticated_mcp_request_is_challenged_with_resource_metadata()
    {
        var response = await _http.PostAsync("/mcp", JsonRpc("initialize", InitializeParams()));

        Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
        var challenge = response.Headers.WwwAuthenticate.ToString();
        Assert.True(challenge.Contains("resource_metadata=\"http://localhost/.well-known/oauth-protected-resource/mcp\""), challenge);

        var prm = await GetJson("/.well-known/oauth-protected-resource/mcp");
        Assert.Equal("http://localhost/mcp", prm.GetProperty("resource").GetString());
        Assert.Equal("http://localhost", prm.GetProperty("authorization_servers")[0].GetString());
    }

    [Fact]
    public async Task Authorization_server_metadata_is_published()
    {
        var metadata = await GetJson("/.well-known/oauth-authorization-server");
        Assert.Equal("http://localhost", metadata.GetProperty("issuer").GetString());
        Assert.Equal("http://localhost/oauth/register", metadata.GetProperty("registration_endpoint").GetString());
        Assert.Contains("S256", metadata.GetProperty("code_challenge_methods_supported").EnumerateArray().Select(e => e.GetString()));
    }

    [Fact]
    public async Task Full_flow_with_dynamic_registration_consent_and_refresh()
    {
        var (clientId, clientSecret) = await RegisterAsync();
        var tokens = await AuthorizeAsync(clientId, clientSecret, expectConsent: true);

        // The upstream code was redeemed with PKCE and the client secret.
        var upstreamRequest = _factory.Idp.TokenRequests.Single();
        Assert.Equal("authorization_code", upstreamRequest["grant_type"]);
        Assert.Equal("upstream-secret", upstreamRequest["client_secret"]);
        Assert.Equal("http://localhost/oauth/callback", upstreamRequest["redirect_uri"]);
        Assert.Contains("https://mail.contoso.com/EWS.AccessAsUser.All", upstreamRequest["scope"]);
        Assert.False(string.IsNullOrEmpty(upstreamRequest["code_verifier"]));

        // MCP access works.
        var toolNames = await ListToolsAsync(tokens.AccessToken);
        Assert.Contains("search_mail", toolNames);
        Assert.Contains("create_draft", toolNames);
        Assert.Contains("list_events", toolNames);
        Assert.DoesNotContain("send_draft", toolNames);

        // Refresh rotates the refresh token.
        var refreshed = await TokenRequestAsync(new()
        {
            ["grant_type"] = "refresh_token",
            ["refresh_token"] = tokens.RefreshToken,
            ["client_id"] = clientId,
            ["client_secret"] = clientSecret,
        });
        Assert.Equal(HttpStatusCode.OK, refreshed.StatusCode);
        var newTokens = await ReadTokens(refreshed);
        Assert.NotEqual(tokens.RefreshToken, newTokens.RefreshToken);
        Assert.NotEmpty(await ListToolsAsync(newTokens.AccessToken));

        var reused = await TokenRequestAsync(new()
        {
            ["grant_type"] = "refresh_token",
            ["refresh_token"] = tokens.RefreshToken,
            ["client_id"] = clientId,
            ["client_secret"] = clientSecret,
        });
        Assert.Equal(HttpStatusCode.BadRequest, reused.StatusCode);
    }

    [Fact]
    public async Task Allow_listed_redirect_uri_skips_consent()
    {
        _factory.Settings["OAuth:AllowedRedirectUriPrefixes:0"] = "https://openwebui.test/";
        var (clientId, clientSecret) = await RegisterAsync();
        var tokens = await AuthorizeAsync(clientId, clientSecret, expectConsent: false);
        Assert.NotEmpty(await ListToolsAsync(tokens.AccessToken));
    }

    [Fact]
    public async Task Token_request_with_wrong_pkce_verifier_or_reused_code_fails()
    {
        var (clientId, clientSecret) = await RegisterAsync();
        var verifier = Crypto.RandomToken(32);
        var code = await GetAuthorizationCodeAsync(clientId, verifier, expectConsent: true);

        var wrong = await TokenRequestAsync(new()
        {
            ["grant_type"] = "authorization_code",
            ["code"] = code,
            ["redirect_uri"] = RedirectUri,
            ["code_verifier"] = "not-the-verifier",
            ["client_id"] = clientId,
            ["client_secret"] = clientSecret,
        });
        Assert.Equal(HttpStatusCode.BadRequest, wrong.StatusCode);

        // The code was consumed by the failed attempt and cannot be used again.
        var retry = await TokenRequestAsync(new()
        {
            ["grant_type"] = "authorization_code",
            ["code"] = code,
            ["redirect_uri"] = RedirectUri,
            ["code_verifier"] = verifier,
            ["client_id"] = clientId,
            ["client_secret"] = clientSecret,
        });
        Assert.Equal(HttpStatusCode.BadRequest, retry.StatusCode);
    }

    [Fact]
    public async Task Confidential_client_must_authenticate()
    {
        var (clientId, _) = await RegisterAsync();
        var response = await TokenRequestAsync(new()
        {
            ["grant_type"] = "refresh_token",
            ["refresh_token"] = "x",
            ["client_id"] = clientId,
            ["client_secret"] = "wrong",
        });
        Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
    }

    [Fact]
    public async Task Callback_without_browser_binding_cookie_is_rejected()
    {
        var (clientId, _) = await RegisterAsync();
        var upstreamUrl = await StartAuthorizationAsync(clientId, Crypto.PkceChallenge(Crypto.RandomToken()), expectConsent: true);
        var state = QueryHelpers.ParseQuery(new Uri(upstreamUrl).Query)["state"].ToString();

        using var otherBrowser = _factory.CreateClient(new WebApplicationFactoryClientOptions { AllowAutoRedirect = false });
        var response = await otherBrowser.GetAsync($"/oauth/callback?code=upstream-code&state={state}");
        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    [Fact]
    public async Task Upstream_token_is_refreshed_when_expired_and_grant_revoked_when_refresh_fails()
    {
        _factory.Idp.AccessTokenLifetimeSeconds = 60; // within the refresh skew → immediately refreshed
        var (clientId, clientSecret) = await RegisterAsync();
        var tokens = await AuthorizeAsync(clientId, clientSecret, expectConsent: true);

        var tokenService = _factory.Services.GetRequiredService<TokenService>();
        var grant = tokenService.ValidateAccessToken(tokens.AccessToken)!;
        Assert.NotNull(grant);

        var ewsToken = await tokenService.GetUpstreamAccessTokenAsync(grant.Id, CancellationToken.None);
        Assert.Equal("ews-access-2", ewsToken);
        var refresh = _factory.Idp.TokenRequests.Last();
        Assert.Equal("refresh_token", refresh["grant_type"]);
        Assert.Equal("ews-refresh-1", refresh["refresh_token"]);

        _factory.Idp.FailRefresh = true;
        await Assert.ThrowsAsync<GrantExpiredException>(() => tokenService.GetUpstreamAccessTokenAsync(grant.Id, CancellationToken.None));
        Assert.Null(tokenService.ValidateAccessToken(tokens.AccessToken));
    }

    [Fact]
    public async Task Registration_rejects_disallowed_redirect_uris()
    {
        var response = await _http.PostAsync("/oauth/register", JsonContent(new { redirect_uris = new[] { "http://evil.test/cb" } }));
        Assert.Equal(HttpStatusCode.BadRequest, response.StatusCode);
    }

    // ---- helpers -------------------------------------------------------------------------------

    private sealed record Tokens(string AccessToken, string RefreshToken);

    private async Task<(string ClientId, string ClientSecret)> RegisterAsync()
    {
        var response = await _http.PostAsync("/oauth/register", JsonContent(new
        {
            client_name = "Open WebUI",
            redirect_uris = new[] { RedirectUri },
            token_endpoint_auth_method = "client_secret_post",
        }));
        Assert.Equal(HttpStatusCode.Created, response.StatusCode);
        var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync()).RootElement;
        return (json.GetProperty("client_id").GetString()!, json.GetProperty("client_secret").GetString()!);
    }

    private async Task<Tokens> AuthorizeAsync(string clientId, string clientSecret, bool expectConsent, string redirectUri = RedirectUri)
    {
        var verifier = Crypto.RandomToken(32);
        var code = await GetAuthorizationCodeAsync(clientId, verifier, expectConsent, redirectUri);
        var response = await TokenRequestAsync(new()
        {
            ["grant_type"] = "authorization_code",
            ["code"] = code,
            ["redirect_uri"] = redirectUri,
            ["code_verifier"] = verifier,
            ["client_id"] = clientId,
            ["client_secret"] = clientSecret,
        });
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        return await ReadTokens(response);
    }

    private async Task<string> GetAuthorizationCodeAsync(string clientId, string verifier, bool expectConsent, string redirectUri = RedirectUri)
    {
        var upstreamUrl = await StartAuthorizationAsync(clientId, Crypto.PkceChallenge(verifier), expectConsent, redirectUri);
        var upstreamQuery = QueryHelpers.ParseQuery(new Uri(upstreamUrl).Query);
        Assert.Equal("upstream-client", upstreamQuery["client_id"]);
        Assert.Equal("S256", upstreamQuery["code_challenge_method"]);
        _factory.Idp.ExpectedNonce = upstreamQuery["nonce"];

        var callback = await _http.GetAsync($"/oauth/callback?code=upstream-code&state={Uri.EscapeDataString(upstreamQuery["state"]!)}");
        Assert.Equal(HttpStatusCode.Redirect, callback.StatusCode);
        var clientRedirect = callback.Headers.Location!;
        Assert.StartsWith(redirectUri, clientRedirect.ToString());
        var query = QueryHelpers.ParseQuery(clientRedirect.Query);
        Assert.Equal("client-state", query["state"]);
        Assert.Equal("http://localhost", query["iss"]);
        return query["code"]!;
    }

    private async Task<string> StartAuthorizationAsync(string clientId, string challenge, bool expectConsent, string redirectUri = RedirectUri)
    {
        var url = QueryHelpers.AddQueryString("/oauth/authorize", new Dictionary<string, string?>
        {
            ["response_type"] = "code",
            ["client_id"] = clientId,
            ["redirect_uri"] = redirectUri,
            ["state"] = "client-state",
            ["code_challenge"] = challenge,
            ["code_challenge_method"] = "S256",
            ["resource"] = "http://localhost/mcp",
        });
        var response = await _http.GetAsync(url);
        if (expectConsent)
        {
            Assert.Equal(HttpStatusCode.OK, response.StatusCode);
            var html = await response.Content.ReadAsStringAsync();
            var tx = TxRegex().Match(html).Groups[1].Value;
            response = await _http.PostAsync("/oauth/consent", new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["tx"] = WebUtility.HtmlDecode(tx),
                ["decision"] = "allow",
            }));
        }
        Assert.Equal(HttpStatusCode.Redirect, response.StatusCode);
        var location = response.Headers.Location!.ToString();
        Assert.StartsWith("https://login.example.test/tenant/oauth2/v2.0/authorize", location);
        return location;
    }

    private Task<HttpResponseMessage> TokenRequestAsync(Dictionary<string, string> form) =>
        _http.PostAsync("/oauth/token", new FormUrlEncodedContent(form));

    private static async Task<Tokens> ReadTokens(HttpResponseMessage response)
    {
        var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync()).RootElement;
        Assert.Equal("Bearer", json.GetProperty("token_type").GetString());
        return new Tokens(json.GetProperty("access_token").GetString()!, json.GetProperty("refresh_token").GetString()!);
    }

    private async Task<List<string>> ListToolsAsync(string accessToken)
    {
        var init = await McpRequestAsync(accessToken, "initialize", InitializeParams());
        Assert.Equal("exchange-ews-mcp", init.GetProperty("result").GetProperty("serverInfo").GetProperty("name").GetString());
        var tools = await McpRequestAsync(accessToken, "tools/list", new { });
        return tools.GetProperty("result").GetProperty("tools").EnumerateArray().Select(t => t.GetProperty("name").GetString()!).ToList();
    }

    private async Task<JsonElement> McpRequestAsync(string accessToken, string method, object parameters)
    {
        var (response, body) = await ServerFactory.McpAsync(_http, accessToken, method, parameters);
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        return body!.Value;
    }

    private static object InitializeParams() => new
    {
        protocolVersion = "2025-06-18",
        capabilities = new { },
        clientInfo = new { name = "test", version = "1.0" },
    };

    private static StringContent JsonRpc(string method, object parameters) =>
        JsonContent(new { jsonrpc = "2.0", id = 1, method, @params = parameters });

    private static StringContent JsonContent(object body) =>
        new(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

    private async Task<JsonElement> GetJson(string url)
    {
        var response = await _http.GetAsync(url);
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        return JsonDocument.Parse(await response.Content.ReadAsStringAsync()).RootElement;
    }

    [GeneratedRegex("name=\"tx\" value=\"([^\"]+)\"")]
    private static partial Regex TxRegex();
}
