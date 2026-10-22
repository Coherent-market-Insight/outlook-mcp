## What and why

## How it was tested

- [ ] `dotnet test tests/OutlookMcp.Tests` passes
- [ ] Live suite run against a test profile (if COM behaviour changed)

## Tool surface

- [ ] No new tool, or the new tool has a description on the method and on every parameter
- [ ] `Access` level is honest (`Write`, `Send`, `Delete`)
- [ ] `docs/TOOLS.md` regenerated if tools changed (`outlook-mcp tools --markdown`)
- [ ] Effect on token cost noted (the registry test enforces the overall budget)
