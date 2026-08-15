# Excel VBA MCP Plugin Layout

This repository is a single root-level plugin for both Claude Code and Codex.

- `.claude-plugin/plugin.json` is the Claude Code marketplace entry point.
- `.codex-plugin/plugin.json` is the Codex marketplace entry point.
- `.mcp.json` is the shared bundled MCP configuration.
- `bin/win-x64/ExcelVbaMcp.exe` is the self-contained bundled server.

Keep the two manifests at the repository root and keep their versions in sync with `Directory.Build.props`. Do not reintroduce a nested plugin directory or a release downloader. Verify `.mcp.json` activation in each host after changing its command resolution.
