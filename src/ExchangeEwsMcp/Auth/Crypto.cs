using System.Security.Cryptography;
using System.Text;

namespace ExchangeEwsMcp.Auth;

internal static class Crypto
{
    /// <summary>Returns a URL-safe random string with the given amount of entropy.</summary>
    public static string RandomToken(int bytes = 32) => Base64Url(RandomNumberGenerator.GetBytes(bytes));

    public static string Sha256(string value) => Base64Url(SHA256.HashData(Encoding.UTF8.GetBytes(value)));

    /// <summary>RFC 7636 S256 code challenge for a verifier.</summary>
    public static string PkceChallenge(string verifier) => Sha256(verifier);

    public static bool FixedTimeEquals(string a, string b) =>
        CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(a), Encoding.UTF8.GetBytes(b));

    public static string Base64Url(ReadOnlySpan<byte> data) =>
        Convert.ToBase64String(data).TrimEnd('=').Replace('+', '-').Replace('/', '_');

    public static byte[] FromBase64Url(string value)
    {
        var s = value.Replace('-', '+').Replace('_', '/');
        s += (s.Length % 4) switch { 2 => "==", 3 => "=", _ => "" };
        return Convert.FromBase64String(s);
    }
}
