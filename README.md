# Keboo.Mcp

<!-- mcp-name: io.github.keboo/keboo-mcp -->

Local stdio MCP server for Microsoft Teams chat workflows.

## Current tools

The first tool set is focused on Microsoft Teams chats:

| Tool | Purpose |
| --- | --- |
| `ListChats` | Lists the signed-in user's chats with chat IDs, members, and preview metadata. |
| `SendChatMessage` | Sends a plain-text message to an existing Teams chat by chat ID. |
| `SendDirectMessage` | Sends a plain-text direct message to a user by work-account UPN/email or Microsoft Entra user ID. Existing 1:1 chats are reused; otherwise one is created. |

## Requirements

1. .NET 10 SDK
2. A Microsoft Entra app registration for delegated Microsoft Graph access
3. A work or school Microsoft 365 account

> Microsoft Teams chat APIs used here don't support personal Microsoft accounts for these delegated flows.

## Required Microsoft Graph delegated permissions

Grant these delegated permissions to the app registration used by the server:

- `User.Read`
- `Chat.Read`
- `Chat.Create`
- `ChatMessage.Send`

`ListChats` uses `Chat.Read`. `SendDirectMessage` uses `Chat.Create` plus `ChatMessage.Send`. The server uses delegated auth, so actions happen as the signed-in user.

## Configure the server

The server reads configuration from either user secrets or environment variables.

### Option 1: user secrets

```powershell
dotnet user-secrets set "KebooMcp:Graph:TenantId" "<tenant-id>" --project .\Keboo.Mcp\Keboo.Mcp.csproj
dotnet user-secrets set "KebooMcp:Graph:ClientId" "<client-id>" --project .\Keboo.Mcp\Keboo.Mcp.csproj
dotnet user-secrets set "KebooMcp:Graph:AuthenticationMode" "InteractiveBrowser" --project .\Keboo.Mcp\Keboo.Mcp.csproj
```

### Option 2: environment variables

```powershell
$env:KebooMcp__Graph__TenantId = "<tenant-id>"
$env:KebooMcp__Graph__ClientId = "<client-id>"
$env:KebooMcp__Graph__AuthenticationMode = "InteractiveBrowser"
```

### Authentication modes

| Value | Behavior |
| --- | --- |
| `InteractiveBrowser` | Opens a local browser for sign-in. This is the default and best fit for local desktop use. |
| `DeviceCode` | Prints a device-code prompt to stderr for environments where opening a browser isn't practical. |

The token cache is persisted locally under the default cache name `Keboo.Mcp`. You can override it with `KebooMcp__Graph__TokenCacheName`.

## Run locally

```powershell
dotnet run --project .\Keboo.Mcp\Keboo.Mcp.csproj
```

## Run from NuGet

```powershell
dnx Keboo.Mcp@0.0.1 --yes
```

The NuGet package includes an embedded `.mcp/server.json` manifest so MCP clients can discover the stdio transport and prompt for the required `KebooMcp__Graph__TenantId` and `KebooMcp__Graph__ClientId` environment variables.

## Example MCP client configuration

This repo includes a root `.mcp.json` file so clients that support repository-scoped MCP configuration can discover the server automatically.

If you need to add it manually, use this shape:

```json
{
  "servers": {
    "keboo-mcp": {
      "type": "stdio",
      "command": "dotnet",
      "args": [
        "run",
        "--project",
        "Keboo.Mcp\\Keboo.Mcp.csproj"
      ],
      "env": {
        "KebooMcp__Graph__AuthenticationMode": "InteractiveBrowser"
      }
    }
  }
}
```
