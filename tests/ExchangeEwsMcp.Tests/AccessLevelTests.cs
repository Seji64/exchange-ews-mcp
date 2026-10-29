using ExchangeEwsMcp.Configuration;
using ExchangeEwsMcp.Tools;
using ModelContextProtocol.Protocol;

namespace ExchangeEwsMcp.Tests;

public sealed class AccessLevelTests
{
    [Theory]
    [InlineData("ReadOnly", false, new[] { "search_mail", "read_mail", "list_events", "get_availability", "whoami", "get_out_of_office" },
        new[] { "create_draft", "move_mail", "mark_mail", "set_out_of_office", "delete_draft", "delete_event", "respond_to_meeting" })]
    [InlineData("NoDelete", false, new[] { "search_mail", "create_draft", "update_draft", "move_mail", "create_event", "update_event", "set_out_of_office" },
        new[] { "delete_draft", "delete_event", "send_draft" })]
    [InlineData("NoDelete", true, new[] { "create_draft", "send_draft", "send_meeting_invitations" }, new[] { "delete_draft", "delete_event" })]
    [InlineData("ReadOnly", true, new[] { "search_mail" }, new[] { "send_draft", "create_draft" })]
    [InlineData("Full", false, new[] { "create_draft", "delete_draft", "delete_event" }, new[] { "send_draft" })]
    public async Task Tools_are_filtered_by_access_level(string level, bool allowSend, string[] expected, string[] forbidden)
    {
        using var factory = new ServerFactory();
        factory.Settings["Exchange:AccessLevel"] = level;
        factory.Settings["Exchange:AllowSend"] = allowSend.ToString();
        using var http = factory.CreateClient();
        var token = factory.CreateAccessToken();

        var (_, body) = await ServerFactory.McpAsync(http, token, "tools/list", new { });
        var tools = body!.Value.GetProperty("result").GetProperty("tools").EnumerateArray().Select(t => t.GetProperty("name").GetString()).ToList();

        Assert.All(expected, t => Assert.Contains(t, tools));
        Assert.All(forbidden, t => Assert.DoesNotContain(t, tools));

        // Removed tools cannot be called either.
        var (_, call) = await ServerFactory.McpAsync(http, token, "tools/call", new { name = forbidden[0], arguments = new { } });
        Assert.True(call!.Value.TryGetProperty("error", out _) || call.Value.GetProperty("result").GetProperty("isError").GetBoolean());
    }

    [Theory]
    [InlineData("move_mail")]
    [InlineData("respond_to_meeting")]
    public async Task No_delete_level_refuses_operations_that_end_in_deleted_items(string tool)
    {
        await using var ews = new FakeEwsServer();
        using var factory = new ServerFactory();
        factory.Settings["Exchange:AccessLevel"] = "NoDelete";
        factory.Settings["Exchange:EwsUrl"] = ews.Url;
        using var http = factory.CreateClient();

        object arguments = tool == "move_mail"
            ? new { ids = new[] { "AAMk1" }, folder = "Deleted Items" }
            : new { id = "AAMk1", response = "decline" };
        var (_, body) = await ServerFactory.McpAsync(http, factory.CreateAccessToken(), "tools/call", new { name = tool, arguments });

        var result = body!.Value.GetProperty("result");
        Assert.True(result.GetProperty("isError").GetBoolean());
        Assert.Contains("eleting is disabled", result.GetProperty("content")[0].GetProperty("text").GetString());
        Assert.Empty(ews.Requests);
    }

    [Fact]
    public void Unannotated_tools_count_as_destructive()
    {
        Assert.False(ToolPolicy.IsAllowed(null, AccessLevel.NoDelete));
        Assert.False(ToolPolicy.IsAllowed(new ToolAnnotations(), AccessLevel.ReadOnly));
        Assert.True(ToolPolicy.IsAllowed(null, AccessLevel.Full));
    }
}
