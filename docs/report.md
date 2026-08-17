# excel-vba-mcp — Plugin Distribution Findings

**Date:** 2026-08-17
**Scope:** Why the `excel-vba-mcp` plugin does not load/appear correctly when installed from the `HappypsychoX/Skills-Marketplace` marketplace in Claude Desktop.
**Intended location:** `docs/report.md`

## Summary

The marketplace entry and the plugin manifest are both syntactically valid. The plugin fails not because of a listing/registration problem but because **its declared MCP server command points at a compiled binary that does not exist in the installed plugin**.

A Claude Code / Claude Desktop plugin is delivered as a **git clone of the plugin repo**. GitHub **release assets are not part of that clone**. The manifest references `bin/win-x64/ExcelVbaMcp.exe`, but that path exists only in a GitHub release, so on any installed machine the path never resolves and the MCP server cannot start.

## What was verified

- **Marketplace manifest** (`.claude-plugin/marketplace.json` in `HappypsychoX/Skills-Marketplace`): contains a correct `excel-vba-mcp` entry pointing at `HappypsychoX/excel-vba-mcp` via the `github` source. No branch is pinned, so the source's **default branch** is used.
- **Plugin manifest** (`.claude-plugin/plugin.json`): valid. Declares one MCP server:
  - `command`: `${CLAUDE_PLUGIN_ROOT}/bin/win-x64/ExcelVbaMcp.exe`
- **`.mcp.json`** (repo root): mirrors the same command as `./bin/win-x64/ExcelVbaMcp.exe`.
- **Repo contents (main branch):** only C# source under `src/ExcelVbaMcp.Server/` plus solution/config files. **There is no `bin/` directory** anywhere in the repo — nothing is built or committed.
- **Binary distribution:** the compiled `.exe` is published as a **GitHub release asset**, not committed (too large to commit directly).
- **Branch note:** both `main` and `master` currently serve the same `plugin.json`. Because the marketplace source does not pin a branch, whichever is the repo's default branch is what installs — worth confirming the stale one is removed or kept in sync.

## Root cause

The MCP command path resolves against `${CLAUDE_PLUGIN_ROOT}` (the cloned plugin directory). Since the `.exe` lives only in a GitHub release and never in the clone, the command target is missing after install. The plugin therefore cannot launch its server. There is no native plugin-install hook that automatically downloads a release asset into the plugin directory.

## Recommended fix — launcher wrapper (preferred)

Point the MCP `command` at a small **script committed to the repo** rather than directly at the `.exe`. On first run the script ensures the binary is present (downloading it from the pinned GitHub release into a cache directory if needed), then execs it.

Why this is the right fit here:
- Keeps the repo small; the release stays the binary host.
- Guarantees the manifest path resolves at install time (the script is in the clone).
- Version is pinned by the script, so installs are reproducible.

Trade-offs: the script must handle download, version pinning, and integrity verification (checksum), and cache invalidation when the version changes.

Sketch of the manifest change:

```json
{
  "mcpServers": {
    "excel-vba-mcp": {
      "command": "${CLAUDE_PLUGIN_ROOT}/bin/launch.cmd",
      "args": []
    }
  }
}
```

The `bin/launch.cmd` (or a PowerShell script it calls) would:
1. Resolve a cache dir (e.g. `%LOCALAPPDATA%\excel-vba-mcp\<version>\`).
2. If `ExcelVbaMcp.exe` is not already cached, download the pinned release asset and verify its checksum.
3. Exec the cached exe, forwarding stdio.

## Alternative options

- **Package-manager distribution (`npx`/`uvx`/`dnx`-style):** publish the server to a registry and make the MCP command a tool-runner invocation. Offloads download and caching to the package manager instead of a hand-rolled script. Fit depends on how you want to distribute a .NET binary.
- **Shrink and commit the binary:** `dotnet publish -r win-x64 -p:PublishSingleFile=true -p:PublishTrimmed=true --self-contained`, possibly with ReadyToRun tuning, may bring the file under GitHub's limits and small enough to commit. Simplest install story if size allows.
- **Git LFS:** possible but risky for plugins — if the installing client does not smudge LFS pointers during clone, the plugin gets a pointer file instead of the exe. Not recommended for this distribution model.

## If it isn't appearing in the marketplace list at all

Separate from the load failure: the marketplace manifest is correct and includes the entry, so a missing listing is most likely a **stale local cache**. Refresh the marketplace in Claude Desktop (update the marketplace, or remove and re-add it) to force a re-pull of `marketplace.json`.

## Suggested next steps for Codex

1. Add a committed launcher script under `bin/` and repoint `plugin.json` + `.mcp.json` at it.
2. Pin the release tag and asset filename in the launcher; verify with a checksum.
3. Add a CI workflow that builds `win-x64` on tag and uploads the asset the launcher expects, so the pinned version always exists.
4. Confirm the repo default branch and remove/sync the stale `main`/`master` duplicate.
5. After changes, refresh the marketplace in Claude Desktop and verify the server starts.
