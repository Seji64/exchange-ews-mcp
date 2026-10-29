using System.Net;
using System.Text;
using System.Text.Json.Serialization;
using ExchangeEwsMcp.Configuration;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Options;

namespace ExchangeEwsMcp.Auth;

/// <summary>
/// A small OAuth 2.1 authorization server (RFC 8414 metadata, RFC 7591 dynamic client registration,
/// authorization code + PKCE, refresh tokens) that brokers the user sign-in to the upstream identity provider.
/// </summary>
public static class OAuthServerEndpoints
{
    private const string BrowserBindingCookie = "ews_mcp_oauth";
    private static readonly TimeSpan TransactionLifetime = TimeSpan.FromMinutes(15);

    public static IEndpointRouteBuilder MapOAuthServer(this IEndpointRouteBuilder app)
    {
        app.MapGet("/.well-known/oauth-authorization-server", Metadata);
        var oauth = app.MapGroup("/oauth").RequireRateLimiting("oauth");
        oauth.MapPost("/register", Register).DisableAntiforgery();
        oauth.MapGet("/authorize", Authorize);
        oauth.MapPost("/consent", Consent).DisableAntiforgery();
        oauth.MapGet("/callback", Callback);
        oauth.MapPost("/token", Token).DisableAntiforgery();
        return app;
    }

    private static IResult Metadata(IOptions<ServerOptions> server, IOptions<OAuthServerOptions> oauth)
    {
        var issuer = server.Value.Issuer;
        var metadata = new Dictionary<string, object>
        {
            ["issuer"] = issuer,
            ["authorization_endpoint"] = issuer + "/oauth/authorize",
            ["token_endpoint"] = issuer + "/oauth/token",
            ["response_types_supported"] = new[] { "code" },
            ["response_modes_supported"] = new[] { "query" },
            ["grant_types_supported"] = new[] { "authorization_code", "refresh_token" },
            ["code_challenge_methods_supported"] = new[] { "S256" },
            ["token_endpoint_auth_methods_supported"] = new[] { "none", "client_secret_post", "client_secret_basic" },
            ["scopes_supported"] = new[] { McpScopes.Exchange },
            ["authorization_response_iss_parameter_supported"] = true,
        };
        if (oauth.Value.AllowDynamicClientRegistration)
        {
            metadata["registration_endpoint"] = issuer + "/oauth/register";
        }
        return Results.Json(metadata);
    }

    // ---- Dynamic client registration (RFC 7591) ------------------------------------------------

    private static IResult Register([FromBody] RegistrationRequest? request, OAuthStore store, IOptions<OAuthServerOptions> oauth, ILoggerFactory loggers)
    {
        if (!oauth.Value.AllowDynamicClientRegistration)
        {
            return OAuthError("access_denied", "Dynamic client registration is disabled.", HttpStatusCode.Forbidden);
        }
        if (request?.RedirectUris is not { Length: > 0 } redirectUris)
        {
            return OAuthError("invalid_redirect_uri", "At least one redirect_uri is required.");
        }
        foreach (var uri in redirectUris)
        {
            if (!IsAcceptableRedirectUri(uri, oauth.Value))
            {
                return OAuthError("invalid_redirect_uri", $"Redirect URI '{uri}' is not allowed.");
            }
        }
        if (request.GrantTypes is { Length: > 0 } grantTypes && !grantTypes.Contains("authorization_code"))
        {
            return OAuthError("invalid_client_metadata", "Only the authorization_code grant is supported.");
        }

        var authMethod = request.TokenEndpointAuthMethod ?? "client_secret_basic";
        if (authMethod is not ("none" or "client_secret_post" or "client_secret_basic"))
        {
            return OAuthError("invalid_client_metadata", $"Unsupported token_endpoint_auth_method '{authMethod}'.");
        }

        var name = request.ClientName is { Length: > 0 } n ? n[..Math.Min(n.Length, 100)] : null;
        var client = store.RegisterClient(name, redirectUris, confidential: authMethod != "none", out var secret);
        loggers.CreateLogger(typeof(OAuthServerEndpoints)).LogInformation(
            "Registered OAuth client {ClientId} ({Name}) with redirect URIs {RedirectUris}", client.ClientId, name, redirectUris);

        var response = new Dictionary<string, object?>
        {
            ["client_id"] = client.ClientId,
            ["client_id_issued_at"] = DateTimeOffset.UtcNow.ToUnixTimeSeconds(),
            ["client_name"] = name,
            ["redirect_uris"] = redirectUris,
            ["grant_types"] = new[] { "authorization_code", "refresh_token" },
            ["response_types"] = new[] { "code" },
            ["token_endpoint_auth_method"] = authMethod,
            ["scope"] = McpScopes.Exchange,
        };
        if (secret is not null)
        {
            response["client_secret"] = secret;
            response["client_secret_expires_at"] = 0;
        }
        return Results.Json(response, statusCode: StatusCodes.Status201Created);
    }

    internal static bool IsAcceptableRedirectUri(string uri, OAuthServerOptions options)
    {
        if (!Uri.TryCreate(uri, UriKind.Absolute, out var parsed) || !string.IsNullOrEmpty(parsed.Fragment))
        {
            return false;
        }
        var isLoopback = parsed.Scheme == Uri.UriSchemeHttp && parsed.IsLoopback;
        if (parsed.Scheme != Uri.UriSchemeHttps && !isLoopback)
        {
            return false;
        }
        return options.AllowedRedirectUriPrefixes.Length == 0 || IsAllowListed(uri, options);
    }

    internal static bool IsAllowListed(string uri, OAuthServerOptions options) =>
        options.AllowedRedirectUriPrefixes.Any(p => uri.StartsWith(p, StringComparison.OrdinalIgnoreCase)
            // "https://a.test" must not match "https://a.test.evil.example/".
            && (p.EndsWith('/') || uri.Length == p.Length || uri[p.Length] is '/' or ':' or '?'));

    // ---- Authorization endpoint ----------------------------------------------------------------

    private static async Task<IResult> Authorize(
        HttpContext http,
        OAuthStore store,
        UpstreamIdentityProvider upstream,
        IOptions<OAuthServerOptions> oauth,
        IOptions<ServerOptions> server,
        TimeProvider time)
    {
        var q = http.Request.Query;
        string? clientId = q["client_id"], redirectUri = q["redirect_uri"];

        // Errors concerning the client or redirect URI must not be redirected (RFC 6749 4.1.2.1).
        var client = string.IsNullOrEmpty(clientId) ? null : store.FindClient(clientId);
        if (client is null)
        {
            return ErrorPage("Unknown client", "The application requesting access is not registered with this server.");
        }
        if (string.IsNullOrEmpty(redirectUri) && client.RedirectUris.Count == 1)
        {
            redirectUri = client.RedirectUris[0];
        }
        if (string.IsNullOrEmpty(redirectUri) || !client.RedirectUris.Contains(redirectUri, StringComparer.Ordinal))
        {
            return ErrorPage("Invalid redirect URI", "The redirect URI does not match the registered client.");
        }

        string? state = q["state"];
        if (q["response_type"] != "code")
        {
            return RedirectError(redirectUri, state, "unsupported_response_type", "Only response_type=code is supported.", server.Value);
        }
        string? codeChallenge = q["code_challenge"];
        if (string.IsNullOrEmpty(codeChallenge) || q["code_challenge_method"] != "S256")
        {
            return RedirectError(redirectUri, state, "invalid_request", "PKCE with code_challenge_method=S256 is required.", server.Value);
        }
        string? resource = q["resource"];
        if (!string.IsNullOrEmpty(resource) && !IsOwnResource(resource, server.Value))
        {
            return RedirectError(redirectUri, state, "invalid_target", "Unknown resource.", server.Value);
        }

        var tx = new AuthTransaction(
            Id: Crypto.RandomToken(),
            ClientId: client.ClientId,
            RedirectUri: redirectUri,
            State: state,
            CodeChallenge: codeChallenge,
            Scope: q["scope"],
            BrowserBinding: Crypto.RandomToken(),
            UpstreamCodeVerifier: Crypto.RandomToken(48),
            Nonce: Crypto.RandomToken(16),
            ExpiresAt: time.GetUtcNow().Add(TransactionLifetime));
        store.SaveTransaction(tx);

        http.Response.Cookies.Append(BrowserBindingCookie, tx.BrowserBinding, new CookieOptions
        {
            HttpOnly = true,
            Secure = server.Value.Issuer.StartsWith("https:", StringComparison.OrdinalIgnoreCase),
            SameSite = SameSiteMode.Lax,
            Path = "/oauth",
            MaxAge = TransactionLifetime,
        });

        // Consent protects against "confused deputy" attacks: the upstream IdP may sign the user in silently,
        // so with open registration the user must confirm which application receives access to the mailbox.
        if (!client.IsStatic && !IsAllowListed(redirectUri, oauth.Value))
        {
            return ConsentPage(client, tx);
        }

        return Results.Redirect(await upstream.BuildAuthorizeUrlAsync(tx.Id, tx.Nonce, tx.UpstreamCodeVerifier, http.RequestAborted));
    }

    private static bool IsOwnResource(string resource, ServerOptions server)
    {
        var r = resource.TrimEnd('/');
        return string.Equals(r, server.McpResource, StringComparison.OrdinalIgnoreCase)
            || string.Equals(r, server.Issuer, StringComparison.OrdinalIgnoreCase);
    }

    private static async Task<IResult> Consent(HttpContext http, OAuthStore store, UpstreamIdentityProvider upstream, IOptions<ServerOptions> server)
    {
        var form = await http.Request.ReadFormAsync(http.RequestAborted);
        var tx = form["tx"].ToString() is { Length: > 0 } id ? store.FindTransaction(id) : null;
        if (tx is null || !IsBoundToBrowser(http, tx))
        {
            return ErrorPage("Request expired", "This authorization request is no longer valid. Please start again from your application.");
        }
        if (form["decision"] != "allow")
        {
            store.DeleteTransaction(tx.Id);
            return RedirectError(tx.RedirectUri, tx.State, "access_denied", "The user denied the request.", server.Value);
        }
        return Results.Redirect(await upstream.BuildAuthorizeUrlAsync(tx.Id, tx.Nonce, tx.UpstreamCodeVerifier, http.RequestAborted));
    }

    private static bool IsBoundToBrowser(HttpContext http, AuthTransaction tx) =>
        http.Request.Cookies.TryGetValue(BrowserBindingCookie, out var binding)
        && binding is not null
        && Crypto.FixedTimeEquals(binding, tx.BrowserBinding);

    // ---- Upstream callback ---------------------------------------------------------------------

    private static async Task<IResult> Callback(
        HttpContext http,
        OAuthStore store,
        UpstreamIdentityProvider upstream,
        IOptions<OAuthServerOptions> oauth,
        IOptions<ServerOptions> server,
        TimeProvider time,
        ILoggerFactory loggers)
    {
        var logger = loggers.CreateLogger(typeof(OAuthServerEndpoints));
        var q = http.Request.Query;
        var tx = q["state"].ToString() is { Length: > 0 } id ? store.FindTransaction(id) : null;
        if (tx is null || !IsBoundToBrowser(http, tx))
        {
            return ErrorPage("Request expired", "This sign-in is no longer valid. Please start again from your application.");
        }
        store.DeleteTransaction(tx.Id);
        http.Response.Cookies.Delete(BrowserBindingCookie, new CookieOptions { Path = "/oauth" });

        if (!string.IsNullOrEmpty(q["error"]))
        {
            logger.LogWarning("Upstream sign-in failed: {Error} {Description}", q["error"].ToString(), q["error_description"].ToString());
            return RedirectError(tx.RedirectUri, tx.State, "access_denied", "Sign-in at the identity provider failed.", server.Value);
        }
        if (string.IsNullOrEmpty(q["code"]))
        {
            return RedirectError(tx.RedirectUri, tx.State, "server_error", "Identity provider returned no code.", server.Value);
        }

        UpstreamTokens tokens;
        UpstreamUser user;
        try
        {
            tokens = await upstream.RedeemCodeAsync(q["code"]!, tx.UpstreamCodeVerifier, http.RequestAborted);
            user = UpstreamIdentityProvider.ReadUser(tokens);
        }
        catch (UpstreamAuthException e)
        {
            logger.LogWarning("Upstream code redemption failed: {Error} {Description}", e.Error, e.Description);
            return RedirectError(tx.RedirectUri, tx.State, "server_error", "Could not obtain a token from the identity provider.", server.Value);
        }
        if (tokens.IdToken is not null && user.Nonce != tx.Nonce)
        {
            return RedirectError(tx.RedirectUri, tx.State, "server_error", "Nonce mismatch.", server.Value);
        }
        if (tokens.RefreshToken is null)
        {
            logger.LogWarning("Identity provider returned no refresh token for {User}; add 'offline_access' to the scopes. " +
                "The user will need to sign in again when the access token expires.", user.UserPrincipalName ?? user.Subject);
        }

        var now = time.GetUtcNow();
        var grant = new Grant
        {
            Id = Crypto.RandomToken(16),
            ClientId = tx.ClientId,
            Subject = user.Subject,
            UserPrincipalName = user.UserPrincipalName,
            Email = user.Email,
            DisplayName = user.DisplayName,
            UpstreamAccessToken = tokens.AccessToken,
            UpstreamAccessTokenExpiresAt = tokens.ExpiresAt,
            UpstreamRefreshToken = tokens.RefreshToken,
            RefreshTokenExpiresAt = now,
            CreatedAt = now,
        };
        store.SaveGrant(grant);
        logger.LogInformation("User {User} authorized client {ClientId}.", user.UserPrincipalName ?? user.Subject, tx.ClientId);

        var code = store.CreateAuthCode(new AuthCode(grant.Id, tx.ClientId, tx.RedirectUri, tx.CodeChallenge, now.Add(oauth.Value.AuthorizationCodeLifetime)));
        var parameters = new Dictionary<string, string?> { ["code"] = code, ["iss"] = server.Value.Issuer };
        if (tx.State is not null) parameters["state"] = tx.State;
        return Results.Redirect(QueryHelpers.AddQueryString(tx.RedirectUri, parameters));
    }

    // ---- Token endpoint ------------------------------------------------------------------------

    private static async Task<IResult> Token(HttpContext http, OAuthStore store, TokenService tokens, TimeProvider time)
    {
        if (!http.Request.HasFormContentType)
        {
            return OAuthError("invalid_request", "Expected application/x-www-form-urlencoded body.");
        }
        var form = await http.Request.ReadFormAsync(http.RequestAborted);
        var (client, clientError) = AuthenticateClient(http, form, store);
        if (client is null)
        {
            return clientError!;
        }

        switch (form["grant_type"].ToString())
        {
            case "authorization_code":
            {
                var code = store.RedeemAuthCode(form["code"].ToString());
                if (code is null || code.ClientId != client.ClientId)
                {
                    return OAuthError("invalid_grant", "Invalid or expired authorization code.");
                }
                var redirectUri = form["redirect_uri"].ToString();
                if (!string.IsNullOrEmpty(redirectUri) && redirectUri != code.RedirectUri)
                {
                    return OAuthError("invalid_grant", "redirect_uri mismatch.");
                }
                var verifier = form["code_verifier"].ToString();
                if (string.IsNullOrEmpty(verifier) || !Crypto.FixedTimeEquals(Crypto.PkceChallenge(verifier), code.CodeChallenge))
                {
                    return OAuthError("invalid_grant", "PKCE verification failed.");
                }
                var grant = store.FindGrant(code.GrantId);
                if (grant is null || grant.Revoked)
                {
                    return OAuthError("invalid_grant", "Authorization no longer valid.");
                }
                return TokenResponse(tokens.Issue(grant));
            }
            case "refresh_token":
            {
                var refreshToken = form["refresh_token"].ToString();
                var grant = string.IsNullOrEmpty(refreshToken) ? null : store.FindGrantByRefreshToken(refreshToken);
                if (grant is null || grant.Revoked || grant.ClientId != client.ClientId || grant.RefreshTokenExpiresAt < time.GetUtcNow())
                {
                    return OAuthError("invalid_grant", "Invalid or expired refresh token.");
                }
                return TokenResponse(tokens.Issue(grant));
            }
            default:
                return OAuthError("unsupported_grant_type", "Supported: authorization_code, refresh_token.");
        }
    }

    private static (OAuthClient? Client, IResult? Error) AuthenticateClient(HttpContext http, IFormCollection form, OAuthStore store)
    {
        string? clientId = form["client_id"], clientSecret = form["client_secret"];

        string? header = http.Request.Headers.Authorization;
        if (header is not null && header.StartsWith("Basic ", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                var decoded = Encoding.UTF8.GetString(Convert.FromBase64String(header["Basic ".Length..].Trim()));
                var separator = decoded.IndexOf(':');
                if (separator > 0)
                {
                    clientId = Uri.UnescapeDataString(decoded[..separator]);
                    clientSecret = Uri.UnescapeDataString(decoded[(separator + 1)..]);
                }
            }
            catch (FormatException)
            {
                return (null, OAuthError("invalid_client", "Malformed Basic authorization header.", HttpStatusCode.Unauthorized));
            }
        }

        var client = string.IsNullOrEmpty(clientId) ? null : store.FindClient(clientId);
        if (client is null)
        {
            return (null, OAuthError("invalid_client", "Unknown client.", HttpStatusCode.Unauthorized));
        }
        if (client.IsConfidential
            && (string.IsNullOrEmpty(clientSecret) || !Crypto.FixedTimeEquals(Crypto.Sha256(clientSecret), client.ClientSecretHash!)))
        {
            return (null, OAuthError("invalid_client", "Client authentication failed.", HttpStatusCode.Unauthorized));
        }
        return (client, null);
    }

    private static IResult TokenResponse(IssuedTokens issued)
    {
        var body = new Dictionary<string, object>
        {
            ["access_token"] = issued.AccessToken,
            ["token_type"] = "Bearer",
            ["expires_in"] = issued.ExpiresIn,
            ["refresh_token"] = issued.RefreshToken,
            ["scope"] = McpScopes.Exchange,
        };
        return new NoStoreJsonResult(body);
    }

    // ---- Helpers -------------------------------------------------------------------------------

    private static IResult OAuthError(string error, string description, HttpStatusCode status = HttpStatusCode.BadRequest) =>
        new NoStoreJsonResult(new Dictionary<string, string> { ["error"] = error, ["error_description"] = description }, (int)status);

    private static IResult RedirectError(string redirectUri, string? state, string error, string description, ServerOptions server)
    {
        var parameters = new Dictionary<string, string?>
        {
            ["error"] = error,
            ["error_description"] = description,
            ["iss"] = server.Issuer,
        };
        if (state is not null) parameters["state"] = state;
        return Results.Redirect(QueryHelpers.AddQueryString(redirectUri, parameters));
    }

    private static IResult ConsentPage(OAuthClient client, AuthTransaction tx)
    {
        var name = WebUtility.HtmlEncode(client.Name ?? "An MCP client");
        var host = WebUtility.HtmlEncode(new Uri(tx.RedirectUri).Authority);
        var body = $"""
            <h1>Allow access to your mailbox?</h1>
            <p><strong>{name}</strong> wants to access your Exchange mailbox (mail, calendar, contacts) on your behalf.</p>
            <p>After you sign in you will be sent back to <code>{host}</code>. Only continue if you recognize this address
            and started this request yourself.</p>
            <form method="post" action="consent">
              <input type="hidden" name="tx" value="{WebUtility.HtmlEncode(tx.Id)}">
              <button type="submit" name="decision" value="allow">Allow and sign in</button>
              <button type="submit" name="decision" value="deny" class="secondary">Deny</button>
            </form>
            """;
        return HtmlPage("Allow access", body);
    }

    private static IResult ErrorPage(string title, string message) =>
        HtmlPage(title, $"<h1>{WebUtility.HtmlEncode(title)}</h1><p>{WebUtility.HtmlEncode(message)}</p>", StatusCodes.Status400BadRequest);

    private static IResult HtmlPage(string title, string body, int status = StatusCodes.Status200OK) =>
        new HtmlResult($$"""
            <!doctype html>
            <html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
            <title>{{WebUtility.HtmlEncode(title)}} – Exchange EWS MCP</title>
            <style>
              body{font-family:system-ui,sans-serif;max-width:32rem;margin:4rem auto;padding:0 1rem;line-height:1.5;color:#1f2328;background:#fff}
              h1{font-size:1.4rem} code{background:#f0f0f0;padding:.1rem .3rem;border-radius:4px}
              button{font:inherit;padding:.5rem 1rem;margin-right:.5rem;border-radius:6px;border:1px solid #0b5cad;background:#0b5cad;color:#fff;cursor:pointer}
              button.secondary{background:#fff;color:#0b5cad}
              @media (prefers-color-scheme: dark){body{background:#0d1117;color:#e6edf3} code{background:#21262d} }
            </style></head><body>{{body}}</body></html>
            """, status);

    private sealed class HtmlResult(string html, int status) : IResult
    {
        public Task ExecuteAsync(HttpContext http)
        {
            http.Response.StatusCode = status;
            http.Response.ContentType = "text/html; charset=utf-8";
            http.Response.Headers.CacheControl = "no-store";
            http.Response.Headers["X-Frame-Options"] = "DENY";
            http.Response.Headers.ContentSecurityPolicy = "default-src 'none'; style-src 'unsafe-inline'; form-action 'self' https:; frame-ancestors 'none'";
            return http.Response.WriteAsync(html);
        }
    }

    private sealed class NoStoreJsonResult(object body, int status = StatusCodes.Status200OK) : IResult
    {
        public Task ExecuteAsync(HttpContext http)
        {
            http.Response.Headers.CacheControl = "no-store";
            http.Response.Headers.Pragma = "no-cache";
            return Results.Json(body, statusCode: status).ExecuteAsync(http);
        }
    }

    private sealed class RegistrationRequest
    {
        [JsonPropertyName("redirect_uris")]
        public string[]? RedirectUris { get; set; }

        [JsonPropertyName("client_name")]
        public string? ClientName { get; set; }

        [JsonPropertyName("token_endpoint_auth_method")]
        public string? TokenEndpointAuthMethod { get; set; }

        [JsonPropertyName("grant_types")]
        public string[]? GrantTypes { get; set; }
    }
}

public static class McpScopes
{
    public const string Exchange = "exchange";
}
