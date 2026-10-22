MCP server for the **Outlook desktop client on Windows**: mail, calendar, contacts and tasks across every account in your profile, built to be token-lean and safe by default.

### Get started

1. Download **`outlook-mcp.exe`** below and keep it somewhere permanent.
2. Run `outlook-mcp.exe doctor` to check that it can see your Outlook.
3. Run `outlook-mcp.exe install --client claude-code`, then restart the app. Other clients: `claude-desktop`, `cursor`, `vscode`, `gemini`, `opencode`, `codex`.

Optional: `outlook-mcp.exe install-skill` installs the bundled agent skill. Add `-- --read-only` after the install command for a read-only setup. Full instructions are in [docs/INSTALL.md](https://github.com/Coherent-market-Insight/outlook-mcp/blob/main/docs/INSTALL.md).

**Requirements:** Windows 10 or 11 and classic Outlook (not the new Outlook for Windows).

**Verify the download** with the `.sha256` files. The executable is not code-signed, so Windows SmartScreen may show an "unrecognized app" warning on first run.

---

