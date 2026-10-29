using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using ExchangeEwsMcp.Auth;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Testing;
using Microsoft.Extensions.DependencyInjection;

namespace ExchangeEwsMcp.Tests;

public sealed class ServerFactory : WebApplicationFactory<Program>
{
    private readonly string _dataDirectory = Path.Combine(Path.GetTempPath(), "ews-mcp-tests-" + Guid.NewGuid().ToString("N"));

    public FakeIdentityProvider Idp { get; } = new();

    public Dictionary<string, string?> Settings { get; } = new()
    {
        ["Server:PublicBaseUrl"] = "http://localhost",
        ["IdentityProvider:Authority"] = FakeIdentityProvider.Authority,
        ["IdentityProvider:ClientId"] = "upstream-client",
        ["IdentityProvider:ClientSecret"] = "upstream-secret",
        ["IdentityProvider:Scopes:0"] = "openid",
        ["IdentityProvider:Scopes:1"] = "offline_access",
        ["IdentityProvider:Scopes:2"] = "https://mail.contoso.com/EWS.AccessAsUser.All",
        ["Exchange:EwsUrl"] = "https://mail.contoso.com/EWS/Exchange.asmx",
        ["OAuth:StaticClients:0:ClientId"] = "static-client",
        ["OAuth:StaticClients:0:ClientSecret"] = "static-secret",
        ["OAuth:StaticClients:0:RedirectUris:0"] = "https://openwebui.test/oauth/static/callback",
    };

    protected override void ConfigureWebHost(IWebHostBuilder builder)
    {
        Settings["Server:DataDirectory"] = _dataDirectory;
        foreach (var (key, value) in Settings)
        {
            builder.UseSetting(key, value);
        }
        builder.ConfigureServices(services =>
        {
            services.AddHttpClient(UpstreamIdentityProvider.HttpClientName).ConfigurePrimaryHttpMessageHandler(() => Idp);
        });
    }

    /// <summary>Creates a grant directly (the browser flow is covered by OAuthFlowTests) and returns an access token for it.</summary>
    public string CreateAccessToken()
    {
        var store = Services.GetRequiredService<OAuthStore>();
        var client = store.RegisterClient("test", ["https://openwebui.test/cb"], confidential: false, out _);
        var grant = new Grant
        {
            Id = Crypto.RandomToken(16),
            ClientId = client.ClientId,
            Subject = "oid-1",
            UserPrincipalName = "alice@contoso.com",
            Email = "alice@contoso.com",
            DisplayName = "Alice",
            UpstreamAccessToken = "ews-token-1",
            UpstreamAccessTokenExpiresAt = DateTimeOffset.UtcNow.AddHours(1),
            CreatedAt = DateTimeOffset.UtcNow,
        };
        store.SaveGrant(grant);
        return Services.GetRequiredService<TokenService>().Issue(grant).AccessToken;
    }

    /// <summary>Sends a JSON-RPC request to the MCP endpoint and returns the response envelope.</summary>
    public static async Task<(HttpResponseMessage Response, JsonElement? Body)> McpAsync(HttpClient http, string accessToken, string method, object parameters)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, "/mcp")
        {
            Content = new StringContent(JsonSerializer.Serialize(new { jsonrpc = "2.0", id = 1, method, @params = parameters }), Encoding.UTF8, "application/json"),
        };
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
        request.Headers.Accept.ParseAdd("application/json");
        request.Headers.Accept.ParseAdd("text/event-stream");
        request.Headers.Add("MCP-Protocol-Version", "2025-06-18");
        var response = await http.SendAsync(request);
        var body = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode) return (response, null);
        if (response.Content.Headers.ContentType?.MediaType == "text/event-stream")
        {
            body = body.Split('\n').Last(l => l.StartsWith("data:"))["data:".Length..].Trim();
        }
        return (response, JsonDocument.Parse(body).RootElement);
    }

    protected override void Dispose(bool disposing)
    {
        base.Dispose(disposing);
        try
        {
            Directory.Delete(_dataDirectory, recursive: true);
        }
        catch (IOException)
        {
        }
    }
}
