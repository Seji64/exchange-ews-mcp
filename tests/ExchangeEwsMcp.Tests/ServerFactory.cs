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
