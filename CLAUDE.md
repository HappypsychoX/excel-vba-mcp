# Excel VBA MCP Plugin Layout

This repository is a single root-level plugin for both Claude Code and Codex.

- `.claude-plugin/plugin.json` is the Claude Code marketplace entry point.
- `.codex-plugin/plugin.json` is the Codex marketplace entry point.
- `.mcp.json` is the Codex MCP configuration and uses the plugin-schema `mcpServers` wrapper.
- `.claude-plugin/plugin.json` carries Claude Code's inline MCP configuration using `CLAUDE_PLUGIN_ROOT`.
- `bin/win-x64/ExcelVbaMcp.exe` exists only in generated/package output and is ignored by Git.

Keep the two manifests at the repository root and keep their versions in sync with `Directory.Build.props`. CI must publish the self-contained server into the completed plugin package at `bin/win-x64/ExcelVbaMcp.exe`; never commit that generated executable. Do not reintroduce a nested source plugin directory, release downloader, bootstrap script, or `%LOCALAPPDATA%` installation. Verify MCP activation in each host after changing either configuration.
