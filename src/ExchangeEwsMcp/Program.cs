using System.Reflection;
using System.Threading.RateLimiting;
using ExchangeEwsMcp.Auth;
using ExchangeEwsMcp.Configuration;
using ExchangeEwsMcp.Exchange;
using ExchangeEwsMcp.Tools;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Options;
using ModelContextProtocol;
using ModelContextProtocol.AspNetCore.Authentication;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;

if (args.Contains("--healthcheck"))
{
    // Used by the Docker HEALTHCHECK (the runtime image ships without curl/wget).
    var port = Environment.GetEnvironmentVariable("ASPNETCORE_HTTP_PORTS")?.Split(';')[0] ?? "8080";
    using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(4) };
    try
    {
        return (await http.GetAsync($"http://127.0.0.1:{port}/healthz")).IsSuccessStatusCode ? 0 : 1;
    }
    catch (HttpRequestException)
    {
        return 1;
    }
}

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddOptions<ServerOptions>().Bind(builder.Configuration.GetSection(ServerOptions.Section))
    .ValidateDataAnnotations()
    .Validate(o => Uri.TryCreate(o.PublicBaseUrl, UriKind.Absolute, out var u) && u.AbsolutePath == "/",
        "Server:PublicBaseUrl must be an absolute URL without path, e.g. https://mcp.contoso.com")
    .ValidateOnStart();
var oauthOptions = builder.Configuration.GetSection(OAuthServerOptions.Section).Get<OAuthServerOptions>() ?? new OAuthServerOptions();
var passthrough = oauthOptions.Mode == AuthMode.Passthrough;

builder.Services.AddOptions<IdentityProviderOptions>().Bind(builder.Configuration.GetSection(IdentityProviderOptions.Section))
    .ValidateDataAnnotations()
    .Validate(o => o.Scopes.Length > 0, "IdentityProvider:Scopes must contain the EWS scope.")
    .Validate(o => passthrough || (o.ClientId.Length > 0 && o.ClientSecret.Length > 0),
        "IdentityProvider:ClientId and ClientSecret are required in OAuth:Mode=Broker.")
    .ValidateOnStart();
builder.Services.AddOptions<OAuthServerOptions>().Bind(builder.Configuration.GetSection(OAuthServerOptions.Section));
builder.Services.AddOptions<ExchangeOptions>().Bind(builder.Configuration.GetSection(ExchangeOptions.Section))
    .ValidateDataAnnotations()
    .ValidateOnStart();

var serverOptions = builder.Configuration.GetSection(ServerOptions.Section).Get<ServerOptions>() ?? new ServerOptions();
var identityProvider = builder.Configuration.GetSection(IdentityProviderOptions.Section).Get<IdentityProviderOptions>() ?? new IdentityProviderOptions();
var exchangeOptions = builder.Configuration.GetSection(ExchangeOptions.Section).Get<ExchangeOptions>() ?? new ExchangeOptions();

builder.Services.AddDataProtection()
    .SetApplicationName("exchange-ews-mcp")
    .PersistKeysToFileSystem(new DirectoryInfo(Path.Combine(serverOptions.DataDirectory, "keys")));

builder.Services.AddSingleton(TimeProvider.System);
builder.Services.AddMemoryCache();
builder.Services.AddHttpContextAccessor();
builder.Services.AddHttpClient(UpstreamIdentityProvider.HttpClientName, c => c.Timeout = TimeSpan.FromSeconds(30));
builder.Services.AddSingleton<OAuthStore>();
builder.Services.AddSingleton<UpstreamIdentityProvider>();
builder.Services.AddSingleton<TokenService>();
builder.Services.AddSingleton<ExchangeServiceFactory>();
builder.Services.AddScoped<EwsContext>();

builder.Services.AddAuthentication(o =>
    {
        o.DefaultAuthenticateScheme = AccessTokenAuthenticationHandler.SchemeName;
        o.DefaultChallengeScheme = McpAuthenticationDefaults.AuthenticationScheme;
    })
    .AddScheme<Microsoft.AspNetCore.Authentication.AuthenticationSchemeOptions, AccessTokenAuthenticationHandler>(
        AccessTokenAuthenticationHandler.SchemeName, null)
    .AddMcp(o =>
    {
        o.ResourceMetadata = new()
        {
            Resource = serverOptions.McpResource,
            // In passthrough mode the client signs in at the identity provider directly and requests the EWS scopes.
            AuthorizationServers = { passthrough ? identityProvider.Authority : serverOptions.Issuer },
            ScopesSupported = passthrough ? [.. identityProvider.Scopes] : [McpScopes.Exchange],
            ResourceName = "Exchange EWS MCP",
        };
    });
builder.Services.AddAuthorization();

builder.Services.AddRateLimiter(o =>
{
    o.RejectionStatusCode = StatusCodes.Status429TooManyRequests;
    o.AddPolicy("oauth", http => RateLimitPartition.GetFixedWindowLimiter(
        http.Connection.RemoteIpAddress?.ToString() ?? "unknown",
        _ => new FixedWindowRateLimiterOptions { PermitLimit = 60, Window = TimeSpan.FromMinutes(1) }));
});

var version = typeof(Program).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion.Split('+')[0] ?? "0.0.0";
var mcp = builder.Services.AddMcpServer(o =>
    {
        o.ServerInfo = new Implementation { Name = "exchange-ews-mcp", Title = "Exchange EWS", Version = version };
        o.ServerInstructions = exchangeOptions.AccessLevel switch
        {
            AccessLevel.ReadOnly => "This server is read-only: it can search and read, but not create or change anything.\n",
            AccessLevel.NoDelete => "Deleting is disabled on this server.\n",
            _ => "",
        } + (exchangeOptions.AllowSend
            ? """
              Access to the user's on-premises Exchange mailbox (mail, calendar, address book, automatic replies).
              Mail is created as drafts. Sending (send_draft, send_meeting_invitations, sendInvitations/sendUpdates/sendResponse)
              is allowed ONLY after the user explicitly confirmed the specific item. Times without offset are in the server time zone (see whoami).
              """
            : """
              Access to the user's on-premises Exchange mailbox (mail, calendar, address book, automatic replies).
              This server never sends mail: messages are saved as drafts and meetings are saved without sending invitations,
              so the user can review and send them in Outlook. Times without offset are in the server time zone (see whoami).
              """);
    })
    .WithHttpTransport(o => o.Stateless = true)
    .WithTools<MailTools>()
    .WithTools<CalendarTools>()
    .WithTools<PeopleTools>()
    .WithTools<OutOfOfficeTools>()
    .WithRequestFilters(filters => filters.AddCallToolFilter(next => async (context, ct) =>
    {
        try
        {
            return await next(context, ct);
        }
        catch (Exception e) when (e is GrantExpiredException or ServiceResponseException or ServiceRequestException or ServiceLocalException)
        {
            var logger = context.Services?.GetService<ILoggerFactory>()?.CreateLogger("ExchangeEwsMcp.Tools");
            logger?.LogWarning(e, "Tool {Tool} failed", context.Params?.Name);
            // Token refresh failures surface from inside the EWS library, possibly wrapped.
            var expired = e as GrantExpiredException ?? e.InnerException as GrantExpiredException;
            return new CallToolResult
            {
                IsError = true,
                Content = [new TextContentBlock { Text = expired?.Message ?? EwsHelpers.DescribeEwsError(e) }],
            };
        }
    }));
if (exchangeOptions.AllowSend)
{
    mcp.WithTools<SendTools>();
}

// Guard rails: drop the tools the access level does not allow, based on their MCP annotations.
builder.Services.PostConfigure<McpServerOptions>(o =>
{
    foreach (var tool in o.ToolCollection?.ToList() ?? [])
    {
        if (!ToolPolicy.IsAllowed(tool.ProtocolTool.Annotations, exchangeOptions.AccessLevel))
        {
            o.ToolCollection!.Remove(tool);
        }
    }
});

var app = builder.Build();

// The server is usually run behind a reverse proxy. All absolute URLs (issuer, redirect URIs, resource metadata)
// are derived from Server:PublicBaseUrl, so make the request look like it arrived on the public origin.
var publicUri = new Uri(app.Services.GetRequiredService<IOptions<ServerOptions>>().Value.Issuer);
app.Use((context, next) =>
{
    context.Request.Scheme = publicUri.Scheme;
    context.Request.Host = HostString.FromUriComponent(publicUri.Authority);
    return next(context);
});

app.UseRateLimiter();
app.UseAuthentication();
app.UseAuthorization();

app.MapGet("/healthz", () => Results.Text("ok"));
if (!passthrough)
{
    app.MapOAuthServer();
}
app.MapMcp(serverOptions.McpPath).RequireAuthorization();

await app.RunAsync();
return 0;

public partial class Program;
