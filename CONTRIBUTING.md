# Contributing

Thanks for helping. This project is small on purpose: it should stay fast, safe by default, and cheap in tokens.

## Setup

You need Windows, the .NET 10 SDK, and (only for the live suite) classic Outlook with a profile you do not mind testing against.

```powershell
git clone https://github.com/Coherent-market-Insight/outlook-mcp.git
cd outlook-mcp
dotnet build
dotnet test tests/OutlookMcp.Tests
```

The unit tests need no Outlook and run in under a second.

## Adding or changing a tool

1. Put it in the matching file under `src/OutlookMcp/Tools/` as a public static method.
2. Mark it with `[ToolMeta("name", "toolset", Access.X)]` and give the method and **every parameter** a short `[Description]`. `Access` decides the MCP annotations and what `--read-only` and `--disable-send` remove, so choose honestly: anything that changes data is `Write`, anything that transmits is `Send`, anything that removes is `Delete`.
3. Run all COM work inside `Sta.Run<string>(() => ...)`. Never touch COM objects outside it.
4. Keep output compact: one line per item, ids from `Ol.Id`, bodies through `Fmt.CleanBody`, long text through `Fmt.Window`. Write tools answer with one confirmation line.
5. Throw `ToolError` with a message the model can act on. Do not return stack traces.
6. Use `dynamic` carefully: assign results to typed variables (`string x = ...`) before using them, because an expression involving a dynamic argument is itself dynamic.
7. Add or update tests. The registry tests enforce naming, descriptions, and a token budget for the whole tool surface, so a new tool may require trimming wording elsewhere.

## Token budget

Tool definitions are paid for on every request. `RegistryTests.Full_tool_surface_stays_within_the_token_budget` fails if the full surface grows past 30,000 characters. Prefer extending an existing tool with an optional parameter over adding a new tool, and keep descriptions to what a model needs to choose and call it.

## The live suite

`node scripts/live-test.mjs` runs against your real Outlook. It creates items tagged `[outlook-mcp test]` and removes only the items it created, by id. It never searches and deletes by subject. Mail is only sent when `OUTLOOK_MCP_TEST_SEND_TO` is set, and only to that address. Never point it at a profile you cannot afford to touch.

## Style

- Match the surrounding code; the repository's `.editorconfig` covers formatting.
- Comments explain why, not what.
- No em dashes in user-facing strings or tool descriptions.

## Pull requests

Keep them focused. Describe the behaviour change, how you tested it, and note any effect on tool definitions or token cost. By contributing you agree your work is released under the [MIT license](LICENSE).
