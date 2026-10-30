using ExchangeEwsMcp.Auth;
using ExchangeEwsMcp.Configuration;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.Options;

namespace ExchangeEwsMcp.Exchange;

/// <summary>
/// Holds attachments for a short time under single-use random URLs, so another tool (e.g. Fileshed in Open WebUI)
/// can fetch the file without its content passing through the model.
/// </summary>
public sealed class DownloadStore(IMemoryCache cache, IOptions<ServerOptions> server)
{
    public sealed record File(string Name, string ContentType, byte[] Content);

    private static readonly TimeSpan Lifetime = TimeSpan.FromMinutes(10);

    public string Add(File file)
    {
        var token = Crypto.RandomToken();
        cache.Set(Key(token), file, Lifetime);
        return $"{server.Value.Issuer}/downloads/{token}";
    }

    /// <summary>Returns the file once; the URL is invalid afterwards.</summary>
    public File? Take(string token)
    {
        if (!cache.TryGetValue(Key(token), out File? file)) return null;
        cache.Remove(Key(token));
        return file;
    }

    private static string Key(string token) => "download:" + token;
}
