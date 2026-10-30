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
    public async Task Shared_mailbox_requests_are_anchored_to_that_mailbox()
    {
        _ews.Responses["FindItem"] = """
            <m:FindItemResponse><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:RootFolder IndexedPagingOffset="0" TotalItemsInView="0" IncludesLastItemInRange="true"><t:Items /></m:RootFolder>
            </m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>
            """;

        _ews.Responses["ResolveNames"] = """
            <m:ResolveNamesResponse><m:ResponseMessages><m:ResolveNamesResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:ResolutionSet TotalItemsInView="1" IncludesLastItemInRange="true"><t:Resolution><t:Mailbox>
                <t:Name>Techmaster</t:Name><t:EmailAddress>techmaster@contoso.com</t:EmailAddress>
                <t:RoutingType>SMTP</t:RoutingType><t:MailboxType>Mailbox</t:MailboxType>
              </t:Mailbox></t:Resolution></m:ResolutionSet>
            </m:ResolveNamesResponseMessage></m:ResponseMessages></m:ResolveNamesResponse>
            """;

        // The model passes the alias; it is resolved to the SMTP address once.
        var first = await CallToolAsync("search_mail", new { mailbox = "techmaster" });
        await CallToolAsync("search_mail", new { mailbox = "techmaster" });
        await CallToolAsync("search_mail", new { });

        Assert.False(first.TryGetProperty("isError", out var isError) && isError.GetBoolean(), first.ToString());
        Assert.Equal(["", "techmaster@contoso.com", "techmaster@contoso.com", ""], _ews.AnchorMailboxes);
        Assert.Contains("<t:EmailAddress>techmaster@contoso.com</t:EmailAddress>", _ews.Requests.ElementAt(1).Body);
    }

    [Fact]
    public async Task Draft_gets_the_outlook_on_the_web_signature_when_requested()
    {
        _ews.Responses["GetUserConfiguration"] = """
            <m:GetUserConfigurationResponse><m:ResponseMessages><m:GetUserConfigurationResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:UserConfiguration>
                <t:UserConfigurationName Name="OWA.UserOptions"><t:DistinguishedFolderId Id="root" /></t:UserConfigurationName>
                <t:ItemId Id="AAMkCfg" ChangeKey="C" />
                <t:Dictionary><t:DictionaryEntry>
                  <t:DictionaryKey><t:Type>String</t:Type><t:Value>signaturehtml</t:Value></t:DictionaryKey>
                  <t:DictionaryValue><t:Type>String</t:Type><t:Value>&lt;b&gt;Alice Example&lt;/b&gt;</t:Value></t:DictionaryValue>
                </t:DictionaryEntry></t:Dictionary>
              </m:UserConfiguration>
            </m:GetUserConfigurationResponseMessage></m:ResponseMessages></m:GetUserConfigurationResponse>
            """;
        _ews.Responses["CreateItem"] = """
            <m:CreateItemResponse><m:ResponseMessages><m:CreateItemResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:Items><t:Message><t:ItemId Id="AAMkDraft" ChangeKey="C" /></t:Message></m:Items>
            </m:CreateItemResponseMessage></m:ResponseMessages></m:CreateItemResponse>
            """;

        var result = await CallToolAsync("create_draft", new { subject = "Hi", body = "Hello Bob", to = new[] { "bob@contoso.com" }, includeSignature = true });

        Assert.False(result.TryGetProperty("isError", out var isError) && isError.GetBoolean(), result.ToString());
        var createItem = _ews.Requests.First(r => r.Body.Contains(":CreateItem ")).Body;
        Assert.Contains("Hello Bob", createItem);
        Assert.Contains("&lt;b&gt;Alice Example&lt;/b&gt;", createItem);

        // The signature is cached (Exchange:SignatureCacheDuration, default 10 minutes).
        await CallToolAsync("create_draft", new { subject = "Again", body = "Hello again", to = new[] { "bob@contoso.com" }, includeSignature = true });
        Assert.Single(_ews.Requests, r => r.Body.Contains(":GetUserConfiguration"));
    }

    [Fact]
    public async Task Expand_group_resolves_the_list_and_expands_nested_lists_once()
    {
        _ews.Responses["ResolveNames"] = """
            <m:ResolveNamesResponse><m:ResponseMessages><m:ResolveNamesResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:ResolutionSet TotalItemsInView="1" IncludesLastItemInRange="true"><t:Resolution><t:Mailbox>
                <t:Name>Team</t:Name><t:EmailAddress>team@contoso.com</t:EmailAddress>
                <t:RoutingType>SMTP</t:RoutingType><t:MailboxType>PublicDL</t:MailboxType>
              </t:Mailbox></t:Resolution></m:ResolutionSet>
            </m:ResolveNamesResponseMessage></m:ResponseMessages></m:ResolveNamesResponse>
            """;
        // Every list contains Alice and the nested list "sub" (which also contains itself: cycle).
        _ews.Responses["ExpandDL"] = """
            <m:ExpandDLResponse><m:ResponseMessages><m:ExpandDLResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:DLExpansion TotalItemsInView="2" IncludesLastItemInRange="true">
                <t:Mailbox><t:Name>Alice</t:Name><t:EmailAddress>alice@contoso.com</t:EmailAddress><t:RoutingType>SMTP</t:RoutingType><t:MailboxType>Mailbox</t:MailboxType></t:Mailbox>
                <t:Mailbox><t:Name>Sub</t:Name><t:EmailAddress>sub@contoso.com</t:EmailAddress><t:RoutingType>SMTP</t:RoutingType><t:MailboxType>PublicDL</t:MailboxType></t:Mailbox>
              </m:DLExpansion>
            </m:ExpandDLResponseMessage></m:ResponseMessages></m:ExpandDLResponse>
            """;

        var result = await CallToolAsync("expand_group", new { group = "Team", recursive = true });

        var json = JsonDocument.Parse(result.GetProperty("content")[0].GetProperty("text").GetString()!).RootElement;
        Assert.Equal("team@contoso.com", json.GetProperty("group").GetString());
        Assert.True(json.GetProperty("complete").GetBoolean());
        Assert.Equal(["alice@contoso.com", "sub@contoso.com"], json.GetProperty("members").EnumerateArray().Select(m => m.GetProperty("email").GetString()));
        Assert.Equal(3, _ews.Requests.Count); // ResolveNames, ExpandDL team, ExpandDL sub
    }

    [Fact]
    public async Task Item_from_a_shared_mailbox_is_retried_there_when_the_mailbox_was_not_passed()
    {
        _ews.Responses["GetItem"] = """
            <m:GetItemResponse><m:ResponseMessages><m:GetItemResponseMessage ResponseClass="Error">
              <m:MessageText>The specified object was not found in the store.</m:MessageText>
              <m:ResponseCode>ErrorItemNotFound</m:ResponseCode><m:DescriptiveLinkKey>0</m:DescriptiveLinkKey><m:Items />
            </m:GetItemResponseMessage></m:ResponseMessages></m:GetItemResponse>
            """;
        _ews.Responses["ConvertId"] = """
            <m:ConvertIdResponse><m:ResponseMessages><m:ConvertIdResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:AlternateId xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="t:AlternateIdType"
                Format="EwsId" Id="AAMkShared" Mailbox="techmaster@contoso.com" />
            </m:ConvertIdResponseMessage></m:ResponseMessages></m:ConvertIdResponse>
            """;
        _ews.Responses["ResolveNames"] = """
            <m:ResolveNamesResponse><m:ResponseMessages><m:ResolveNamesResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:ResolutionSet TotalItemsInView="1" IncludesLastItemInRange="true"><t:Resolution><t:Mailbox>
                <t:Name>Alice</t:Name><t:EmailAddress>alice@contoso.com</t:EmailAddress>
                <t:RoutingType>SMTP</t:RoutingType><t:MailboxType>Mailbox</t:MailboxType>
              </t:Mailbox></t:Resolution></m:ResolutionSet>
            </m:ResolveNamesResponseMessage></m:ResponseMessages></m:ResolveNamesResponse>
            """;
        _ews.AnchoredResponses["GetItem"] = """
            <m:GetItemResponse><m:ResponseMessages><m:GetItemResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:Items><t:Message><t:ItemId Id="AAMkShared" ChangeKey="C" /><t:Subject>Shared hello</t:Subject><t:Body BodyType="Text">Hi</t:Body></t:Message></m:Items>
            </m:GetItemResponseMessage></m:ResponseMessages></m:GetItemResponse>
            """;

        var result = await CallToolAsync("read_mail", new { id = "AAMkShared" });

        Assert.False(result.TryGetProperty("isError", out var isError) && isError.GetBoolean(), result.ToString());
        Assert.Contains("Shared hello", result.GetProperty("content")[0].GetProperty("text").GetString());
        Assert.Equal("techmaster@contoso.com", _ews.AnchorMailboxes.Last());
    }

    [Theory]
    [InlineData(false, "New text", true)]
    [InlineData(true, "Old quoted text", null)]
    public async Task Read_mail_returns_only_the_new_part_unless_history_is_requested(bool includeQuotedHistory, string expected, bool? omitted)
    {
        _ews.Responses["GetItem"] = """
            <m:GetItemResponse><m:ResponseMessages><m:GetItemResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:Items><t:Message><t:ItemId Id="AAMk1" ChangeKey="C" /><t:Subject>Re: Plan</t:Subject>
                <t:Body BodyType="Text">New text

            From: Bob
            Old quoted text with a lot of history</t:Body>
                <t:UniqueBody BodyType="Text">New text</t:UniqueBody>
              </t:Message></m:Items>
            </m:GetItemResponseMessage></m:ResponseMessages></m:GetItemResponse>
            """;

        var result = await CallToolAsync("read_mail", new { id = "AAMk1", includeQuotedHistory });

        var json = JsonDocument.Parse(result.GetProperty("content")[0].GetProperty("text").GetString()!).RootElement;
        Assert.Contains(expected, json.GetProperty("body").GetString());
        Assert.Equal(omitted, json.TryGetProperty("quotedHistoryOmitted", out var o) ? o.GetBoolean() : null);
    }

    [Fact]
    public async Task Mark_mail_uses_one_get_and_one_update_for_all_items()
    {
        string Messages(string op, string inner, string tail = "") => string.Concat(new[] { "A1", "A2", "A3" }.Select(id =>
            $"<m:{op}ResponseMessage ResponseClass=\"Success\"><m:ResponseCode>NoError</m:ResponseCode><m:Items><t:Message><t:ItemId Id=\"{id}\" ChangeKey=\"C\" />{inner}</t:Message></m:Items>{tail}</m:{op}ResponseMessage>"));
        _ews.Responses["GetItem"] = $"<m:GetItemResponse><m:ResponseMessages>{Messages("GetItem", "<t:IsRead>false</t:IsRead>")}</m:ResponseMessages></m:GetItemResponse>";
        _ews.Responses["UpdateItem"] = $"<m:UpdateItemResponse><m:ResponseMessages>{Messages("UpdateItem", "", "<m:ConflictResults><t:Count>0</t:Count></m:ConflictResults>")}</m:ResponseMessages></m:UpdateItemResponse>";

        var result = await CallToolAsync("mark_mail", new { ids = new[] { "A1", "A2", "A3" }, isRead = true });

        Assert.Equal("Updated 3 item(s).", result.GetProperty("content")[0].GetProperty("text").GetString());
        Assert.Equal(2, _ews.Requests.Count);
    }

    [Fact]
    public async Task Folder_paths_are_resolved_once()
    {
        _ews.Responses["FindFolder"] = """
            <m:FindFolderResponse><m:ResponseMessages><m:FindFolderResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:RootFolder TotalItemsInView="1" IncludesLastItemInRange="true"><t:Folders><t:Folder><t:FolderId Id="FOLDER1" ChangeKey="C" /></t:Folder></t:Folders></m:RootFolder>
            </m:FindFolderResponseMessage></m:ResponseMessages></m:FindFolderResponse>
            """;
        _ews.Responses["FindItem"] = """
            <m:FindItemResponse><m:ResponseMessages><m:FindItemResponseMessage ResponseClass="Success">
              <m:ResponseCode>NoError</m:ResponseCode>
              <m:RootFolder IndexedPagingOffset="0" TotalItemsInView="0" IncludesLastItemInRange="true"><t:Items /></m:RootFolder>
            </m:FindItemResponseMessage></m:ResponseMessages></m:FindItemResponse>
            """;

        await CallToolAsync("search_mail", new { folder = "Inbox/Projects" });
        await CallToolAsync("search_mail", new { folder = "Inbox/Projects" });

        Assert.Single(_ews.Requests, r => r.Body.Contains(":FindFolder "));
        Assert.All(_ews.Requests.Where(r => r.Body.Contains(":FindItem ")), r => Assert.Contains("FOLDER1", r.Body));
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
        var text = result.GetProperty("content")[0].GetProperty("text").GetString();
        Assert.Contains("401", text);
        Assert.Contains("x-ms-diagnostics: 2000003;reason=\"Invalid audience\"", text);
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
