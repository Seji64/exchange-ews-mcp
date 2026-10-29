using ExchangeEwsMcp.Auth;
using ExchangeEwsMcp.Configuration;
using ExchangeEwsMcp.Exchange;
using ExchangeEwsMcp.Tools;

namespace ExchangeEwsMcp.Tests;

public class UnitTests
{
    [Theory]
    [InlineData("Europe/Berlin", "W. Europe Standard Time")]
    [InlineData("W. Europe Standard Time", "W. Europe Standard Time")]
    [InlineData("UTC", "UTC")]
    public void Time_zone_is_resolved_to_windows_id(string configured, string expected)
    {
        Assert.Equal(expected, ExchangeServiceFactory.ResolveTimeZone(configured).Id);
    }

    [Fact]
    public void Dates_without_offset_are_local_and_with_offset_are_utc()
    {
        var local = EwsContext.ParseDateTimeValue("2026-11-03T14:30", "x");
        Assert.Equal(DateTimeKind.Unspecified, local.Kind);
        Assert.Equal(new DateTime(2026, 11, 3, 14, 30, 0), local);

        var utc = EwsContext.ParseDateTimeValue("2026-11-03T14:30:00+01:00", "x");
        Assert.Equal(DateTimeKind.Utc, utc.Kind);
        Assert.Equal(new DateTime(2026, 11, 3, 13, 30, 0), utc);

        Assert.Equal(DateTimeKind.Utc, EwsContext.ParseDateTimeValue("2026-11-03T14:30:00Z", "x").Kind);
        Assert.Equal(new DateTime(2026, 11, 3), EwsContext.ParseDateTimeValue("2026-11-03", "x"));
        Assert.Throws<ModelContextProtocol.McpException>(() => EwsContext.ParseDateTimeValue("next tuesday", "x"));
    }

    [Fact]
    public void Aqs_query_combines_filters()
    {
        var aqs = MailTools.BuildAqs("budget", "Alice \"A\"", "Q3", "2026-10-01", "2026-11-01T00:00", true, true);
        Assert.Equal("budget from:\"Alice A\" subject:\"Q3\" received>=2026-10-01 received<2026-11-01 isread:false hasattachment:true", aqs);
    }

    [Theory]
    [InlineData("https://openwebui.contoso.com/oauth/callback", true)]
    [InlineData("http://localhost:3000/cb", true)]
    [InlineData("http://127.0.0.1/cb", true)]
    [InlineData("http://openwebui.contoso.com/cb", false)]
    [InlineData("https://evil.test/cb", false)]
    [InlineData("https://openwebui.contoso.com/cb#frag", false)]
    [InlineData("https://openwebui.contoso.com.evil.test/cb", false)]
    [InlineData("https://bare.contoso.com/cb", true)]
    [InlineData("https://bare.contoso.com.evil.test/cb", false)]
    [InlineData("not a url", false)]
    public void Redirect_uri_validation(string uri, bool expected)
    {
        var options = new OAuthServerOptions { AllowedRedirectUriPrefixes = ["https://openwebui.contoso.com/", "https://bare.contoso.com", "http://localhost", "http://127.0.0.1"] };
        Assert.Equal(expected, OAuthServerEndpoints.IsAcceptableRedirectUri(uri, options));
    }

    [Fact]
    public void Redirect_uri_without_allow_list_requires_https_or_loopback()
    {
        var options = new OAuthServerOptions();
        Assert.True(OAuthServerEndpoints.IsAcceptableRedirectUri("https://any.test/cb", options));
        Assert.False(OAuthServerEndpoints.IsAcceptableRedirectUri("http://any.test/cb", options));
    }

    [Fact]
    public void User_is_read_from_entra_and_adfs_id_tokens()
    {
        var entra = UpstreamIdentityProvider.ReadUser(new UpstreamTokens("at", DateTimeOffset.UtcNow, null, FakeIdentityProvider.Jwt(new()
        {
            ["oid"] = "oid-1", ["sub"] = "sub-1", ["preferred_username"] = "alice@contoso.com", ["name"] = "Alice", ["nonce"] = "n",
        })));
        Assert.Equal(("oid-1", "alice@contoso.com", "alice@contoso.com", "Alice", "n"), (entra.Subject, entra.UserPrincipalName, entra.Email, entra.DisplayName, entra.Nonce));

        var adfs = UpstreamIdentityProvider.ReadUser(new UpstreamTokens("at", DateTimeOffset.UtcNow, null, FakeIdentityProvider.Jwt(new()
        {
            ["sub"] = "sub-2", ["upn"] = "bob@contoso.com", ["unique_name"] = @"CONTOSO\bob",
        })));
        Assert.Equal(("sub-2", "bob@contoso.com", "bob@contoso.com"), (adfs.Subject, adfs.UserPrincipalName, adfs.Email));
    }

    [Fact]
    public void Pkce_challenge_matches_rfc7636_example()
    {
        Assert.Equal("E9Melhoa2OwvFrEMTJguCHaoeK1t8URWbuGJSstw-cM", Crypto.PkceChallenge("dBjftJeZ4CVP-mB92K27uhbUJU1p1r_wW1gFWFOEjXk"));
    }
}
