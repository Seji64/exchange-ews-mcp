using ExchangeEwsMcp.Configuration;
using ModelContextProtocol.Protocol;

namespace ExchangeEwsMcp.Tools;

internal static class ToolPolicy
{
    /// <summary>Unset hints count as writing / destructive, as in the MCP specification.</summary>
    public static bool IsAllowed(ToolAnnotations? annotations, AccessLevel level) => level switch
    {
        AccessLevel.ReadOnly => annotations?.ReadOnlyHint == true,
        AccessLevel.NoDelete => annotations?.ReadOnlyHint == true || annotations?.DestructiveHint == false,
        _ => true,
    };
}
