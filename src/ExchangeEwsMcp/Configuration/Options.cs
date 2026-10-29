using System.ComponentModel.DataAnnotations;

namespace ExchangeEwsMcp.Configuration;

/// <summary>General hosting settings.</summary>
public sealed class ServerOptions
{
    public const string Section = "Server";

    /// <summary>
    /// The externally visible base URL of this server (e.g. https://mcp.contoso.com).
    /// Used as OAuth issuer, for redirect URIs and in the protected resource metadata.
    /// </summary>
    [Required]
    public string PublicBaseUrl { get; set; } = "";

    /// <summary>Directory for the SQLite database and the data protection key ring.</summary>
    [Required]
    public string DataDirectory { get; set; } = "data";

    /// <summary>Path the MCP endpoint is mapped to.</summary>
    public string McpPath { get; set; } = "/mcp";

    public string Issuer => PublicBaseUrl.TrimEnd('/');

    public string McpResource => Issuer + McpPath;
}

/// <summary>
/// The upstream OpenID Connect provider that issues the EWS access tokens:
/// Microsoft Entra ID (Hybrid Modern Authentication) or AD FS (Exchange on-premises Modern Authentication).
/// </summary>
public sealed class IdentityProviderOptions
{
    public const string Section = "IdentityProvider";

    /// <summary>
    /// OIDC authority. The discovery document is loaded from {Authority}/.well-known/openid-configuration.
    /// Entra ID: https://login.microsoftonline.com/{tenant-id}/v2.0 — AD FS: https://adfs.contoso.com/adfs
    /// </summary>
    [Required]
    public string Authority { get; set; } = "";

    [Required]
    public string ClientId { get; set; } = "";

    [Required]
    public string ClientSecret { get; set; } = "";

    /// <summary>
    /// Scopes requested upstream. For Entra ID HMA this must contain the EWS scope of the on-premises
    /// Exchange URL, e.g. "https://mail.contoso.com/EWS.AccessAsUser.All", plus "offline_access".
    /// </summary>
    public string[] Scopes { get; set; } = [];

    /// <summary>Optional "resource" parameter (AD FS style), e.g. "https://mail.contoso.com/".</summary>
    public string? Resource { get; set; }
}

/// <summary>Settings of the built-in OAuth 2.1 authorization server that MCP clients talk to.</summary>
public sealed class OAuthServerOptions
{
    public const string Section = "OAuth";

    public TimeSpan AccessTokenLifetime { get; set; } = TimeSpan.FromHours(1);

    public TimeSpan RefreshTokenLifetime { get; set; } = TimeSpan.FromDays(30);

    public TimeSpan AuthorizationCodeLifetime { get; set; } = TimeSpan.FromMinutes(5);

    /// <summary>Allow MCP clients to register themselves (RFC 7591). Open WebUI's "OAuth 2.1" mode needs this.</summary>
    public bool AllowDynamicClientRegistration { get; set; } = true;

    /// <summary>
    /// If non-empty, redirect URIs of dynamically registered clients must start with one of these prefixes
    /// (e.g. "https://openwebui.contoso.com/"), and the consent page is skipped. Strongly recommended in production.
    /// </summary>
    public string[] AllowedRedirectUriPrefixes { get; set; } = [];

    /// <summary>Pre-registered clients (e.g. for Open WebUI's "OAuth 2.1 (Static)" mode).</summary>
    public List<StaticClientOptions> StaticClients { get; set; } = [];
}

public sealed class StaticClientOptions
{
    public string ClientId { get; set; } = "";

    /// <summary>Optional. Without a secret the client is a public client (PKCE only).</summary>
    public string? ClientSecret { get; set; }

    public string? Name { get; set; }

    public string[] RedirectUris { get; set; } = [];
}

/// <summary>Exchange Web Services settings.</summary>
public sealed class ExchangeOptions
{
    public const string Section = "Exchange";

    /// <summary>EWS endpoint, e.g. https://mail.contoso.com/EWS/Exchange.asmx</summary>
    [Required]
    public string EwsUrl { get; set; } = "";

    /// <summary>Time zone used to interpret and render dates (IANA or Windows id).</summary>
    public string TimeZone { get; set; } = "Europe/Berlin";

    /// <summary>
    /// When false (default) the server never sends mail: everything ends up as a draft and meetings are
    /// saved without sending invitations. When true, additional tools to send drafts / invitations are exposed.
    /// </summary>
    public bool AllowSend { get; set; }

    /// <summary>Upper bound for list/search results.</summary>
    public int MaxResults { get; set; } = 50;

    /// <summary>Maximum number of characters of a message body returned to the model.</summary>
    public int MaxBodyLength { get; set; } = 20000;

    /// <summary>Maximum attachment size (bytes) the server will download and return.</summary>
    public int MaxAttachmentBytes { get; set; } = 5 * 1024 * 1024;

    /// <summary>Optional EWS request trace logging (very verbose, contains mail content).</summary>
    public bool TraceEws { get; set; }
}
