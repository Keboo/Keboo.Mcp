namespace Keboo.Mcp.Configuration;

internal sealed class TeamsGraphOptions
{
    public const string SectionName = "KebooMcp:Graph";

    public string? TenantId { get; set; }

    public string? ClientId { get; set; }

    public TeamsGraphAuthenticationMode AuthenticationMode { get; set; } = TeamsGraphAuthenticationMode.InteractiveBrowser;

    public string TokenCacheName { get; set; } = "Keboo.Mcp";

    public int DefaultChatPageSize { get; set; } = 50;

    public int MaxChatPageSize { get; set; } = 100;
}

internal enum TeamsGraphAuthenticationMode
{
    InteractiveBrowser,
    DeviceCode
}
