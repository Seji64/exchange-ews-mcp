using System.Collections.Concurrent;
using System.Net;
using System.Text;
using System.Text.Json;
using ExchangeEwsMcp.Auth;

namespace ExchangeEwsMcp.Tests;

/// <summary>Simulates the upstream OIDC provider (discovery + token endpoint).</summary>
public sealed class FakeIdentityProvider : HttpMessageHandler
{
    public const string Authority = "https://login.example.test/tenant/v2.0";

    public ConcurrentQueue<Dictionary<string, string>> TokenRequests { get; } = new();

    public string? ExpectedNonce { get; set; }

    public int AccessTokenLifetimeSeconds { get; set; } = 3600;

    public bool FailRefresh { get; set; }

    private int _counter;

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
    {
        var url = request.RequestUri!.ToString();
        if (url == Authority + "/.well-known/openid-configuration")
        {
            return Json(new
            {
                issuer = Authority,
                authorization_endpoint = "https://login.example.test/tenant/oauth2/v2.0/authorize",
                token_endpoint = "https://login.example.test/tenant/oauth2/v2.0/token",
            });
        }
        if (url == "https://login.example.test/tenant/oauth2/v2.0/token")
        {
            var form = (await request.Content!.ReadAsStringAsync(ct)).Split('&')
                .Select(p => p.Split('=', 2))
                .ToDictionary(p => WebUtility.UrlDecode(p[0]), p => WebUtility.UrlDecode(p[1]));
            TokenRequests.Enqueue(form);

            if (form["grant_type"] == "refresh_token" && FailRefresh)
            {
                return Json(new { error = "invalid_grant", error_description = "expired" }, HttpStatusCode.BadRequest);
            }
            var n = Interlocked.Increment(ref _counter);
            return Json(new
            {
                token_type = "Bearer",
                access_token = $"ews-access-{n}",
                refresh_token = $"ews-refresh-{n}",
                expires_in = AccessTokenLifetimeSeconds,
                id_token = Jwt(new Dictionary<string, object?>
                {
                    ["oid"] = "user-oid-1",
                    ["preferred_username"] = "alice@contoso.com",
                    ["name"] = "Alice Example",
                    ["nonce"] = ExpectedNonce,
                }),
            });
        }
        return new HttpResponseMessage(HttpStatusCode.NotFound);
    }

    public static string Jwt(Dictionary<string, object?> claims) =>
        Crypto.Base64Url(Encoding.UTF8.GetBytes("{\"alg\":\"none\"}")) + "." +
        Crypto.Base64Url(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(claims))) + ".";

    private static HttpResponseMessage Json(object body, HttpStatusCode status = HttpStatusCode.OK) =>
        new(status) { Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json") };
}
