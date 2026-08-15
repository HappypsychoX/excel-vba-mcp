# Excel VBA MCP Server — Development Checklist

## Current Status

- **Updated:** August 15, 2026
- **Branch:** `main`
- **Phase 1 server version:** `0.1.3`; tagged plugin release pending

**Overall Phase 1:** in progress — the read-only server, CI packaging, first tagged release, and protocol smoke test are complete. Distribution now follows the bundled-plugin model: the plugin contains its `.mcp.json` and self-contained `ExcelVbaMcp.exe`; no first-use GitHub download is part of plugin activation.

Completed foundation work:

- Minimal .NET 10 stdio MCP server using the official `ModelContextProtocol` C# SDK.
- Exactly two read-only tools: `ping` and `get_version`.
- Self-contained, single-file `win-x64` publish and framework-free MCP smoke test.
- Root-level Claude Code and Codex plugin manifests with host-specific bundled MCP configuration.
- GitHub Actions restores, builds, publishes, smoke-tests, packages, and uploads the plugin ZIP.
- The next tagged release packages the plugin ZIP only.

No Excel, COM, VBIDE, workbook, or VBA automation has been implemented.

## Phase 1 — Bundle and Activate the Read-Only Server

- [x] Create a minimal MCP server project in C#/.NET 10 using stdio transport.
- [x] Build it as a self-contained, single-file Windows executable (`ExcelVbaMcp.exe`).
- [x] Have the server expose only the read-only `ping` and `get_version` tools.
- [x] Create and push the GitHub repository for the project.
- [x] Publish the Phase 1 server in the `v0.1.0` GitHub Release.
- [x] Include `ExcelVbaMcp.exe` inside root `bin/win-x64/`.
- [x] Add root `.mcp.json` for Codex and an inline MCP configuration for Claude Code.
- [x] Package a plugin ZIP containing its manifest, MCP configuration, and server executable.
- [x] Verify an MCP client can list exactly both tools and call `ping`/`get_version` against the bundled executable.
- [ ] Install the root plugin in clean Claude Code and Codex profiles; verify host activation, tool listing, calls, and clean shutdown.
- [ ] Confirm a second fresh session in each host reuses the installed plugin bundle without a network download.

### Superseded download bootstrap work

The former first-use downloader that placed a release executable in `%LOCALAPPDATA%\ExcelVbaMcp\` is intentionally removed. Plugin installation already saves the server alongside its manifest and `.mcp.json`; adding a second distribution mechanism would duplicate the plugin model.

### Phase 1 Success Criteria

A clean Codex profile can install the plugin, start its bundled server, list exactly `ping` and `get_version`, call both tools, end the MCP session cleanly, and start a later session from the same installed bundle without downloading a release.

**Status:** pending clean-profile Codex host validation. The server and bundle can be validated locally; host installation must still establish the runtime command-resolution behavior.

---

## Phase 2 — Connect to Excel

- [ ] Add Excel COM interop support.
- [ ] Detect running Excel instances.
- [ ] Add `list_workbooks`.
- [ ] Attach to an already-open workbook without taking ownership of Excel.
- [ ] Verify disconnecting from MCP does not close the user's Excel instance.
- [ ] Verify all COM references are released cleanly.
- [ ] Verify no orphaned `EXCEL.EXE` processes are created.

---

## Phase 3 — Read VBA

- [ ] Add VBIDE project access.
- [ ] Detect whether **Trust access to the VBA project object model** is enabled.
- [ ] Add `list_vba_components`.
- [ ] Add `read_vba_module`.
- [ ] Add `list_vba_procedures`.
- [ ] Add `read_vba_procedure`.
- [ ] Add `search_vba`.
- [ ] Test standard modules, class modules, worksheets, and `ThisWorkbook`.
- [ ] Return useful errors for protected/inaccessible VBA projects.

---

## Phase 4 — Safely Modify VBA

- [ ] Add `replace_vba_procedure`.
- [ ] Add `insert_vba_procedure`.
- [ ] Add `delete_vba_procedure`.
- [ ] Add `create_vba_module`.
- [ ] Add `rename_vba_component`.
- [ ] Add `delete_vba_component`.
- [ ] Require targeted edits rather than rewriting entire modules whenever possible.
- [ ] Keep workbook saving separate from code modification.
- [ ] Verify changes appear immediately in Excel/VBE.
- [ ] Verify repeated edits do not corrupt the workbook.

---

## Phase 5 — Backup and Recovery

- [ ] Add `backup_workbook`.
- [ ] Automatically create a backup before the first write operation in a session.
- [ ] Add an edit-session/transaction concept.
- [ ] Add `rollback_changes`.
- [ ] Prevent accidental overwriting of the original backup.
- [ ] Test forced failures during editing.
- [ ] Verify the original workbook can always be recovered.

---

## Phase 6 — Validation and Execution

- [ ] Investigate reliable VBA compile validation through VBIDE.
- [ ] Add `compile_vba_project` if practical.
- [ ] Add `run_macro`.
- [ ] Capture and return VBA runtime errors.
- [ ] Prevent arbitrary macro execution unless explicitly requested.
- [ ] Add optional save-after-success behavior.
- [ ] Keep `save_workbook` as an explicit tool.

---

## Phase 7 — Versioning and Updates

- [x] Add a server version number, centrally defined in `Directory.Build.props`.
- [x] Add the read-only `get_version` tool.
- [x] Define `v*` GitHub Release naming and packaging conventions in the README and workflow.
- [ ] Define plugin version upgrade, integrity, and rollback behavior.

---

## Phase 8 — Hardening

- [ ] Add structured logging.
- [ ] Log workbook/VBA modifications without unnecessarily storing VBA source.
- [ ] Add protection against modifying the wrong workbook.
- [ ] Detect unsaved workbook state before dangerous operations.
- [ ] Handle multiple Excel instances.
- [ ] Handle multiple workbooks with identical filenames.
- [ ] Test `.xlsm` and `.xlsb` workbooks.
- [ ] Test large VBA projects.
- [ ] Test abnormal MCP/agent termination.
- [ ] Confirm the server never directly modifies `vbaProject.bin` or other Office package internals.
- [ ] Document security requirements and known limitations.

---

## Core Rule

- [x] **Document the invariant that all workbook and VBA modifications go through Excel/COM/VBIDE. Never directly rewrite the `.xlsm`, `.xlsb`, or `vbaProject.bin` file structure.** Enforcement and integration testing begin when write capabilities are introduced.
