using System.Net;
using System.Text.Json;

namespace ExchangeEwsMcp.Tests;

/// <summary>Calls tools through the MCP endpoint against a fake EWS server.</summary>
public sealed class EwsToolTests : IAsyncLifetime
{
    private readonly FakeEwsServer _ews = new();
    private ServerFactory _factory = null!;
    private HttpClient _http = null!;
    private string _accessToken = null!;

    public async Task InitializeAsync()
    {
        _factory = new ServerFactory();
        _factory.Settings["Exchange:EwsUrl"] = _ews.Url;
        _factory.Settings["Exchange:TimeZone"] = "Europe/Berlin";
        _http = _factory.CreateClient();

        _accessToken = _factory.CreateAccessToken();
        await Task.CompletedTask;
    }

    public async Task DisposeAsync()
    {
        _http.Dispose();
        await _factory.DisposeAsync();
        await _ews.DisposeAsync();
    }

    [Fact]
    public async Task Search_mail_sends_bearer_token_and_returns_local_times()
    {
        _ews.Responses["FindItem"] = """
            <m:FindItemResponse>
              <m:ResponseMessages>
                <m:FindItemResponseMessage ResponseClass="Success">
                  <m:ResponseCode>NoError</m:ResponseCode>
                  <m:RootFolder IndexedPagingOffset="1" TotalItemsInView="1" IncludesLastItemInRange="true">
                    <t:Items>
                      <t:Message>
                        <t:ItemId Id="AAMkMsg1" ChangeKey="CQAAAB" />
                        <t:Subject>Quarterly budget</t:Subject>
                        <t:DateTimeReceived>2026-10-28T09:15:00Z</t:DateTimeReceived>
                        <t:HasAttachments>true</t:HasAttachments>
                        <t:Importance>High</t:Importance>
                        <t:IsDraft>false</t:IsDraft>
                        <t:DisplayTo>Alice Example</t:DisplayTo>
                        <t:Preview>Hi Alice, here are the numbers</t:Preview>
                        <t:From><t:Mailbox><t:Name>Bob Builder</t:Name><t:EmailAddress>bob@contoso.com</t:EmailAddress><t:RoutingType>SMTP</t:RoutingType></t:Mailbox></t:From>
                        <t:IsRead>false</t:IsRead>
                      </t:Message>
                    </t:Items>
                  </m:RootFolder>
                </m:FindItemResponseMessage>
              </m:ResponseMessages>
            </m:FindItemResponse>
            """;

        var result = await CallToolAsync("search_mail", new { query = "budget", maxResults = 5 });

        Assert.False(result.TryGetProperty("isError", out var isError) && isError.GetBoolean(), result.ToString());
        var text = result.GetProperty("content")[0].GetProperty("text").GetString()!;
        var json = JsonDocument.Parse(text).RootElement;
        var item = json.GetProperty("items")[0];
        Assert.Equal("AAMkMsg1", item.GetProperty("id").GetString());
        Assert.Equal("Quarterly budget", item.GetProperty("subject").GetString());
        Assert.Equal("bob@contoso.com", item.GetProperty("from").GetProperty("email").GetString());
        Assert.Equal("2026-10-28T10:15:00+01:00", item.GetProperty("received").GetString());
        Assert.False(item.GetProperty("isRead").GetBoolean());

        var (authorization, body) = _ews.Requests.Single();
        Assert.Equal("Bearer ews-token-1", authorization);
        Assert.Contains("<m:QueryString>budget</m:QueryString>", body);
        Assert.Contains("Id=\"inbox\"", body);
    }

    [Fact]
    public async Task Ews_errors_are_reported_to_the_model()
    {
        _ews.Responses["GetItem"] = """
            <m:GetItemResponse>
              <m:ResponseMessages>
                <m:GetItemResponseMessage ResponseClass="Error">
                  <m:MessageText>The specified object was not found in the store.</m:MessageText>
                  <m:ResponseCode>ErrorItemNotFound</m:ResponseCode>
                  <m:DescriptiveLinkKey>0</m:DescriptiveLinkKey>
                  <m:Items />
                </m:GetItemResponseMessage>
              </m:ResponseMessages>
            </m:GetItemResponse>
            """;

        var result = await CallToolAsync("read_mail", new { id = "AAMkMissing" });

        Assert.True(result.GetProperty("isError").GetBoolean());
        Assert.Contains("not found", result.GetProperty("content")[0].GetProperty("text").GetString());
    }

    [Fact]
    public async Task Unauthorized_ews_response_explains_configuration_problem()
    {
        _ews.StatusCode = 401;
        var result = await CallToolAsync("list_events", new { });
        Assert.True(result.GetProperty("isError").GetBoolean());
        Assert.Contains("401", result.GetProperty("content")[0].GetProperty("text").GetString());
    }

    [Fact]
    public async Task Sending_is_rejected_when_disabled()
    {
        var result = await CallToolAsync("create_event", new
        {
            subject = "Sync",
            start = "2026-11-03T10:00",
            end = "2026-11-03T10:30",
            requiredAttendees = new[] { "bob@contoso.com" },
            sendInvitations = true,
        });
        Assert.True(result.GetProperty("isError").GetBoolean());
        Assert.Contains("Sending is disabled", result.GetProperty("content")[0].GetProperty("text").GetString());
        Assert.Empty(_ews.Requests);
    }

    private async Task<JsonElement> CallToolAsync(string name, object arguments)
    {
        var (response, body) = await ServerFactory.McpAsync(_http, _accessToken, "tools/call", new { name, arguments });
        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        return body!.Value.GetProperty("result");
    }
}
