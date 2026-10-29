using System.Collections.Concurrent;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Server;
using Microsoft.AspNetCore.Hosting.Server.Features;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace ExchangeEwsMcp.Tests;

/// <summary>A minimal EWS endpoint returning canned SOAP responses per operation.</summary>
public sealed class FakeEwsServer : IAsyncDisposable
{
    private readonly WebApplication _app;

    public ConcurrentQueue<(string? Authorization, string Body)> Requests { get; } = new();

    /// <summary>Operation name (e.g. "FindItem") → SOAP body content placed inside s:Body.</summary>
    public ConcurrentDictionary<string, string> Responses { get; } = new();

    public int StatusCode { get; set; } = 200;

    public string Url { get; }

    public FakeEwsServer()
    {
        var builder = WebApplication.CreateSlimBuilder();
        builder.WebHost.UseUrls("http://127.0.0.1:0");
        builder.Logging.ClearProviders();
        _app = builder.Build();
        _app.MapPost("/EWS/Exchange.asmx", async (HttpContext http) =>
        {
            var body = await new StreamReader(http.Request.Body).ReadToEndAsync();
            Requests.Enqueue((http.Request.Headers.Authorization.ToString(), body));
            if (StatusCode != 200)
            {
                http.Response.StatusCode = StatusCode;
                return;
            }
            var operation = Responses.Keys.FirstOrDefault(op => body.Contains($":{op} ") || body.Contains($":{op}>"));
            http.Response.ContentType = "text/xml; charset=utf-8";
            await http.Response.WriteAsync(Envelope(operation is null ? "" : Responses[operation]));
        });
        _app.StartAsync().GetAwaiter().GetResult();
        var address = _app.Services.GetRequiredService<IServer>().Features.Get<IServerAddressesFeature>()!.Addresses.First();
        Url = address + "/EWS/Exchange.asmx";
    }

    private static string Envelope(string body) => $"""
        <?xml version="1.0" encoding="utf-8"?>
        <s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">
          <s:Header>
            <h:ServerVersionInfo MajorVersion="15" MinorVersion="2" MajorBuildNumber="2562" MinorBuildNumber="17" Version="V2017_07_11"
              xmlns:h="http://schemas.microsoft.com/exchange/services/2006/types" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" />
          </s:Header>
          <s:Body xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
            {body}
          </s:Body>
        </s:Envelope>
        """;

    public async ValueTask DisposeAsync() => await _app.DisposeAsync();
}
