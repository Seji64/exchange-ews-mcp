using System.Net;
using System.Text.Json;

namespace ExchangeEwsMcp.Tests;

/// <summary>OAuth:Mode=Passthrough — the MCP client brings its own EWS token (e.g. Open WebUI "OAuth 2.1 (Static)").</summary>
public sealed class PassthroughTests : IAsyncLifetime
{
    private readonly FakeEwsServer _ews = new();
    private ServerFactory _factory = null!;
    private HttpClient _http = null!;

    public Task InitializeAsync()
    {
        _factory = new ServerFactory();
        _factory.Settings["OAuth:Mode"] = "Passthrough";
        _factory.Settings["IdentityProvider:ClientId"] = null;
        _factory.Settings["IdentityProvider:ClientSecret"] = null;
        _factory.Settings["Exchange:EwsUrl"] = _ews.Url;
        _http = _factory.CreateClient(new() { AllowAutoRedirect = false });
        return Task.CompletedTask;
    }

    public async Task DisposeAsync()
    {
        _http.Dispose();
        await _factory.DisposeAsync();
        await _ews.DisposeAsync();
    }

    [Fact]
    public async Task Resource_metadata_points_to_the_identity_provider_and_broker_endpoints_are_off()
    {
        var prm = JsonDocument.Parse(await _http.GetStringAsync("/.well-known/oauth-protected-resource/mcp")).RootElement;
        Assert.Equal(FakeIdentityProvider.Authority, prm.GetProperty("authorization_servers")[0].GetString());
        Assert.Equal("https://mail.contoso.com", prm.GetProperty("resource").GetString());
        Assert.Contains("https://mail.contoso.com/EWS.AccessAsUser.All",
            prm.GetProperty("scopes_supported").EnumerateArray().Select(s => s.GetString()));

        Assert.Equal(HttpStatusCode.NotFound, (await _http.GetAsync("/.well-known/oauth-authorization-server")).StatusCode);
        Assert.Equal(HttpStatusCode.NotFound, (await _http.PostAsync("/oauth/register", new StringContent("{}"))).StatusCode);
    }

    [Fact]
    public async Task Token_is_passed_through_to_ews()
    {
        _ews.Responses["FindItem"] = """
            <m:FindItemResponse><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:RootFolder IndexedPagingOffset="1" TotalItemsInView="1" IncludesLastItemInRange="true">
                <t:Items><t:Message><t:ItemId Id="AAMk1" ChangeKey="C" /><t:Subject>Back soon</t:Subject></t:Message></t:Items>
              </m:RootFolder>
            </m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>
            """;
        var token = Token(DateTimeOffset.UtcNow.AddHours(1));

        var response = await CallToolAsync(token, "search_mail", new { });

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var body = await response.Content.ReadAsStringAsync();
        Assert.True(body.Contains("Back soon"), body);
        Assert.Equal("Bearer " + token, _ews.Requests.Single().Authorization);
    }

    [Fact]
    public async Task Expired_token_is_rejected_so_the_client_refreshes()
    {
        var response = await CallToolAsync(Token(DateTimeOffset.UtcNow.AddMinutes(-1)), "whoami", new { });
        Assert.Equal(HttpStatusCode.Unauthorized, response.StatusCode);
        Assert.Contains("resource_metadata=", response.Headers.WwwAuthenticate.ToString());
        Assert.Empty(_ews.Requests);
    }

    private static string Token(DateTimeOffset expires) => FakeIdentityProvider.Jwt(new()
    {
        ["aud"] = "https://mail.contoso.com",
        ["oid"] = "oid-1",
        ["upn"] = "alice@contoso.com",
        ["exp"] = expires.ToUnixTimeSeconds(),
    });

    private async Task<HttpResponseMessage> CallToolAsync(string token, string name, object arguments) =>
        (await ServerFactory.McpAsync(_http, token, "tools/call", new { name, arguments })).Response;
}
