# Excel VBA MCP Plugin Layout

This repository is a single root-level plugin for both Claude Code and Codex.

- `.claude-plugin/plugin.json` is the Claude Code marketplace entry point.
- `.codex-plugin/plugin.json` is the Codex marketplace entry point.
- `.mcp.json` is the Codex MCP configuration and uses the documented `mcp_servers` wrapper.
- `.claude-plugin/plugin.json` carries Claude Code's inline MCP configuration using `CLAUDE_PLUGIN_ROOT`.
- `bin/win-x64/ExcelVbaMcp.exe` is the self-contained bundled server.

Keep the two manifests at the repository root and keep their versions in sync with `Directory.Build.props`. Do not reintroduce a nested plugin directory or a release downloader. Verify MCP activation in each host after changing either configuration.
