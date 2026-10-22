# Installation

## 1. Get the executable

Download `outlook-mcp.exe` from the [releases page](https://github.com/Coherent-market-Insight/outlook-mcp/releases/latest), or build it:

```powershell
git clone https://github.com/Coherent-market-Insight/outlook-mcp.git
cd outlook-mcp
dotnet publish src/OutlookMcp -c Release -r win-x64 -p:PublishSingleFile=true -o dist
```

Put the file somewhere permanent (it is referenced by path from client configs), for example `%LOCALAPPDATA%\outlook-mcp\outlook-mcp.exe`.

Run `outlook-mcp.exe doctor`. It reports the Windows version, whether Outlook's COM class is registered, how long the connection took, every account with its type, and how fast the inbox can be read.

## 2. Register with a client

`outlook-mcp install --client <name>` edits the right file, preserves everything else in it, writes a `.bak-outlook-mcp` backup, and updates the entry in place if you run it again. Flags for the server go after `--`.

| Client | `--client` | File edited | Key |
|---|---|---|---|
| Claude Code | `claude-code` | `~/.claude.json` (or `<--config-dir>\.claude.json`) | `mcpServers` |
| Claude Desktop | `claude-desktop` | `%APPDATA%\Claude\claude_desktop_config.json` | `mcpServers` |
| Cursor | `cursor` | `~/.cursor/mcp.json` | `mcpServers` |
| VS Code | `vscode` | `%APPDATA%\Code\User\mcp.json` | `servers` |
| Gemini CLI | `gemini` | `~/.gemini/settings.json` | `mcpServers` |
| OpenCode | `opencode` | `~/.config/opencode/opencode.json` | `mcp` |
| Codex CLI | `codex` | `~/.codex/config.toml` | `[mcp_servers.outlook]` |

Examples:

```powershell
outlook-mcp install --client claude-code
outlook-mcp install --client claude-code --config-dir C:\Users\me\.work-profile   # a second profile
outlook-mcp install --client cursor -- --read-only
outlook-mcp install --client vscode -- --toolsets mail,calendar --allowed-domains example.com
```

Restart the client (or reconnect the server from its MCP panel) afterwards.

### Manual configuration

Stdio entry, shown for the `mcpServers` style:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "C:\\Users\\me\\AppData\\Local\\outlook-mcp\\outlook-mcp.exe",
      "args": ["--read-only"]
    }
  }
}
```

VS Code uses `"servers"` and wants `"type": "stdio"`. OpenCode uses `"mcp": { "outlook": { "type": "local", "command": ["...exe", "--read-only"], "enabled": true } }`. Codex uses TOML:

```toml
[mcp_servers.outlook]
command = 'C:\Users\me\AppData\Local\outlook-mcp\outlook-mcp.exe'
args = ["--read-only"]
```

## 3. Presets

| Goal | Arguments |
|---|---|
| Triage only, no changes at all | `--read-only` |
| Everything except sending | `--disable-send` |
| Mail only, minimal tokens | `--toolsets mail --read-only` |
| Calendar assistant | `--toolsets calendar` |
| Allow sending, but only internally | `--allowed-domains yourcompany.com` |

## 4. The skill

`outlook-mcp install-skill` writes `SKILL.md` to `~/.claude/skills/outlook/` (use `--config-dir` for another profile, or `--dir` for any skills folder). It tells an agent how to use the tools cheaply (list, then read; batch changes) and safely (drafts first, explicit sends, no invented addresses). Agents that do not support skills can be given the same text as project instructions: the file is [skills/outlook/SKILL.md](../skills/outlook/SKILL.md).

## Uninstall

Remove the `outlook` entry from the client config (the backup next to it holds the previous state) and delete the executable. Nothing else is written outside the temp folder used for saved attachments (`%TEMP%\outlook-mcp`).
