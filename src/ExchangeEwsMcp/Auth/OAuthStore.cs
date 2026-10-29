using System.Text.Json;
using ExchangeEwsMcp.Configuration;
using Microsoft.AspNetCore.DataProtection;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Options;

namespace ExchangeEwsMcp.Auth;

public sealed record OAuthClient(string ClientId, string? ClientSecretHash, string? Name, IReadOnlyList<string> RedirectUris)
{
    public bool IsConfidential => ClientSecretHash is not null;
}

/// <summary>A pending authorization request (between /oauth/authorize and /oauth/callback).</summary>
public sealed record AuthTransaction(
    string Id,
    string ClientId,
    string RedirectUri,
    string? State,
    string CodeChallenge,
    string? Scope,
    string BrowserBinding,
    string UpstreamCodeVerifier,
    string Nonce,
    DateTimeOffset ExpiresAt);

/// <summary>An authorization code issued to an MCP client.</summary>
public sealed record AuthCode(string GrantId, string ClientId, string RedirectUri, string CodeChallenge, DateTimeOffset ExpiresAt);

/// <summary>
/// A user's authorization of an MCP client. Holds the (encrypted) upstream tokens used to call EWS.
/// </summary>
public sealed class Grant
{
    public required string Id { get; init; }
    public required string ClientId { get; init; }
    public required string Subject { get; set; }
    public string? UserPrincipalName { get; set; }
    public string? Email { get; set; }
    public string? DisplayName { get; set; }
    public string? UpstreamAccessToken { get; set; }
    public DateTimeOffset UpstreamAccessTokenExpiresAt { get; set; }
    public string? UpstreamRefreshToken { get; set; }
    public string? RefreshTokenHash { get; set; }
    public DateTimeOffset RefreshTokenExpiresAt { get; set; }
    public DateTimeOffset CreatedAt { get; init; }
    public bool Revoked { get; set; }
}

/// <summary>SQLite-backed persistence for clients, pending authorizations, codes and grants.</summary>
public sealed class OAuthStore
{
    private readonly string _connectionString;
    private readonly IDataProtector _protector;
    private readonly TimeProvider _time;

    public OAuthStore(IOptions<ServerOptions> server, IDataProtectionProvider dataProtection, TimeProvider time)
    {
        Directory.CreateDirectory(server.Value.DataDirectory);
        var path = Path.Combine(server.Value.DataDirectory, "exchange-ews-mcp.db");
        _connectionString = new SqliteConnectionStringBuilder { DataSource = path }.ToString();
        _protector = dataProtection.CreateProtector("ExchangeEwsMcp.UpstreamTokens.v1");
        _time = time;
        Initialize();
    }

    private DateTimeOffset Now => _time.GetUtcNow();

    private void Initialize()
    {
        using var db = Open();
        Execute(db, """
            PRAGMA journal_mode=WAL;
            CREATE TABLE IF NOT EXISTS clients (
                client_id TEXT PRIMARY KEY,
                client_secret_hash TEXT NULL,
                name TEXT NULL,
                redirect_uris TEXT NOT NULL,
                created_at INTEGER NOT NULL);
            CREATE TABLE IF NOT EXISTS auth_transactions (
                id TEXT PRIMARY KEY,
                data TEXT NOT NULL,
                expires_at INTEGER NOT NULL);
            CREATE TABLE IF NOT EXISTS auth_codes (
                code_hash TEXT PRIMARY KEY,
                grant_id TEXT NOT NULL,
                client_id TEXT NOT NULL,
                redirect_uri TEXT NOT NULL,
                code_challenge TEXT NOT NULL,
                expires_at INTEGER NOT NULL);
            CREATE TABLE IF NOT EXISTS grants (
                id TEXT PRIMARY KEY,
                client_id TEXT NOT NULL,
                subject TEXT NOT NULL,
                upn TEXT NULL,
                email TEXT NULL,
                display_name TEXT NULL,
                upstream_access_token TEXT NULL,
                upstream_access_expires_at INTEGER NOT NULL,
                upstream_refresh_token TEXT NULL,
                refresh_token_hash TEXT NULL,
                refresh_expires_at INTEGER NOT NULL,
                created_at INTEGER NOT NULL,
                revoked INTEGER NOT NULL DEFAULT 0);
            CREATE INDEX IF NOT EXISTS ix_grants_refresh ON grants(refresh_token_hash);
            """);
    }

    private SqliteConnection Open()
    {
        var db = new SqliteConnection(_connectionString);
        db.Open();
        return db;
    }

    private static void Execute(SqliteConnection db, string sql, params (string Name, object? Value)[] parameters)
    {
        using var cmd = Command(db, sql, parameters);
        cmd.ExecuteNonQuery();
    }

    private static SqliteCommand Command(SqliteConnection db, string sql, params (string Name, object? Value)[] parameters)
    {
        var cmd = db.CreateCommand();
        cmd.CommandText = sql;
        foreach (var (name, value) in parameters)
        {
            cmd.Parameters.AddWithValue(name, value ?? DBNull.Value);
        }
        return cmd;
    }

    private string? Protect(string? value) => value is null ? null : _protector.Protect(value);

    private string? Unprotect(string? value)
    {
        if (value is null) return null;
        try
        {
            return _protector.Unprotect(value);
        }
        catch (System.Security.Cryptography.CryptographicException)
        {
            // Key ring lost or rotated out: treat as missing so the user re-authenticates.
            return null;
        }
    }

    // ---- Clients -------------------------------------------------------------------------------

    public OAuthClient RegisterClient(string? name, IReadOnlyList<string> redirectUris, bool confidential, out string? clientSecret)
    {
        var clientId = Guid.NewGuid().ToString("N");
        clientSecret = confidential ? Crypto.RandomToken() : null;
        var client = new OAuthClient(clientId, clientSecret is null ? null : Crypto.Sha256(clientSecret), name, redirectUris);
        using var db = Open();
        Execute(db, "INSERT INTO clients (client_id, client_secret_hash, name, redirect_uris, created_at) VALUES ($id, $secret, $name, $uris, $now)",
            ("$id", client.ClientId), ("$secret", client.ClientSecretHash), ("$name", name),
            ("$uris", JsonSerializer.Serialize(redirectUris)), ("$now", Now.ToUnixTimeSeconds()));
        return client;
    }

    public OAuthClient? FindClient(string clientId)
    {
        using var db = Open();
        using var cmd = Command(db, "SELECT client_secret_hash, name, redirect_uris FROM clients WHERE client_id = $id", ("$id", clientId));
        using var r = cmd.ExecuteReader();
        if (!r.Read()) return null;
        return new OAuthClient(
            clientId,
            r.IsDBNull(0) ? null : r.GetString(0),
            r.IsDBNull(1) ? null : r.GetString(1),
            JsonSerializer.Deserialize<string[]>(r.GetString(2)) ?? []);
    }

    // ---- Authorization transactions ------------------------------------------------------------

    public void SaveTransaction(AuthTransaction tx)
    {
        using var db = Open();
        Execute(db, "DELETE FROM auth_transactions WHERE expires_at < $now; DELETE FROM auth_codes WHERE expires_at < $now;",
            ("$now", Now.ToUnixTimeSeconds()));
        Execute(db, "INSERT INTO auth_transactions (id, data, expires_at) VALUES ($id, $data, $exp)",
            ("$id", tx.Id), ("$data", _protector.Protect(JsonSerializer.Serialize(tx))), ("$exp", tx.ExpiresAt.ToUnixTimeSeconds()));
    }

    public AuthTransaction? FindTransaction(string id)
    {
        using var db = Open();
        using var cmd = Command(db, "SELECT data FROM auth_transactions WHERE id = $id AND expires_at >= $now",
            ("$id", id), ("$now", Now.ToUnixTimeSeconds()));
        var data = cmd.ExecuteScalar() as string;
        var json = Unprotect(data);
        return json is null ? null : JsonSerializer.Deserialize<AuthTransaction>(json);
    }

    public void DeleteTransaction(string id)
    {
        using var db = Open();
        Execute(db, "DELETE FROM auth_transactions WHERE id = $id", ("$id", id));
    }

    // ---- Authorization codes -------------------------------------------------------------------

    public string CreateAuthCode(AuthCode code)
    {
        var value = Crypto.RandomToken();
        using var db = Open();
        Execute(db, "INSERT INTO auth_codes (code_hash, grant_id, client_id, redirect_uri, code_challenge, expires_at) VALUES ($h, $g, $c, $r, $cc, $exp)",
            ("$h", Crypto.Sha256(value)), ("$g", code.GrantId), ("$c", code.ClientId), ("$r", code.RedirectUri),
            ("$cc", code.CodeChallenge), ("$exp", code.ExpiresAt.ToUnixTimeSeconds()));
        return value;
    }

    /// <summary>Looks up and deletes an authorization code (codes are single use).</summary>
    public AuthCode? RedeemAuthCode(string code)
    {
        using var db = Open();
        using var tx = db.BeginTransaction();
        AuthCode? result = null;
        using (var cmd = Command(db, "SELECT grant_id, client_id, redirect_uri, code_challenge, expires_at FROM auth_codes WHERE code_hash = $h", ("$h", Crypto.Sha256(code))))
        {
            cmd.Transaction = tx;
            using var r = cmd.ExecuteReader();
            if (r.Read())
            {
                result = new AuthCode(r.GetString(0), r.GetString(1), r.GetString(2), r.GetString(3), DateTimeOffset.FromUnixTimeSeconds(r.GetInt64(4)));
            }
        }
        using (var del = Command(db, "DELETE FROM auth_codes WHERE code_hash = $h", ("$h", Crypto.Sha256(code))))
        {
            del.Transaction = tx;
            del.ExecuteNonQuery();
        }
        tx.Commit();
        return result is not null && result.ExpiresAt >= Now ? result : null;
    }

    // ---- Grants --------------------------------------------------------------------------------

    public void SaveGrant(Grant g)
    {
        using var db = Open();
        Execute(db, """
            INSERT INTO grants (id, client_id, subject, upn, email, display_name, upstream_access_token, upstream_access_expires_at,
                                upstream_refresh_token, refresh_token_hash, refresh_expires_at, created_at, revoked)
            VALUES ($id, $client, $sub, $upn, $email, $name, $at, $atexp, $rt, $rth, $rtexp, $created, $revoked)
            ON CONFLICT(id) DO UPDATE SET
                subject = excluded.subject, upn = excluded.upn, email = excluded.email, display_name = excluded.display_name,
                upstream_access_token = excluded.upstream_access_token, upstream_access_expires_at = excluded.upstream_access_expires_at,
                upstream_refresh_token = excluded.upstream_refresh_token, refresh_token_hash = excluded.refresh_token_hash,
                refresh_expires_at = excluded.refresh_expires_at, revoked = excluded.revoked
            """,
            ("$id", g.Id), ("$client", g.ClientId), ("$sub", g.Subject), ("$upn", g.UserPrincipalName), ("$email", g.Email),
            ("$name", g.DisplayName), ("$at", Protect(g.UpstreamAccessToken)), ("$atexp", g.UpstreamAccessTokenExpiresAt.ToUnixTimeSeconds()),
            ("$rt", Protect(g.UpstreamRefreshToken)), ("$rth", g.RefreshTokenHash), ("$rtexp", g.RefreshTokenExpiresAt.ToUnixTimeSeconds()),
            ("$created", g.CreatedAt.ToUnixTimeSeconds()), ("$revoked", g.Revoked ? 1 : 0));
    }

    public Grant? FindGrant(string id) => QueryGrant("id = $v", id);

    public Grant? FindGrantByRefreshToken(string refreshToken) => QueryGrant("refresh_token_hash = $v", Crypto.Sha256(refreshToken));

    private Grant? QueryGrant(string where, string value)
    {
        using var db = Open();
        using var cmd = Command(db, $"""
            SELECT id, client_id, subject, upn, email, display_name, upstream_access_token, upstream_access_expires_at,
                   upstream_refresh_token, refresh_token_hash, refresh_expires_at, created_at, revoked
            FROM grants WHERE {where}
            """, ("$v", value));
        using var r = cmd.ExecuteReader();
        if (!r.Read()) return null;
        string? S(int i) => r.IsDBNull(i) ? null : r.GetString(i);
        return new Grant
        {
            Id = r.GetString(0),
            ClientId = r.GetString(1),
            Subject = r.GetString(2),
            UserPrincipalName = S(3),
            Email = S(4),
            DisplayName = S(5),
            UpstreamAccessToken = Unprotect(S(6)),
            UpstreamAccessTokenExpiresAt = DateTimeOffset.FromUnixTimeSeconds(r.GetInt64(7)),
            UpstreamRefreshToken = Unprotect(S(8)),
            RefreshTokenHash = S(9),
            RefreshTokenExpiresAt = DateTimeOffset.FromUnixTimeSeconds(r.GetInt64(10)),
            CreatedAt = DateTimeOffset.FromUnixTimeSeconds(r.GetInt64(11)),
            Revoked = r.GetInt64(12) != 0,
        };
    }
}
