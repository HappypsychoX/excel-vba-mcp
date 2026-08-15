# Excel VBA MCP Server — Development Checklist

## Current Status

- **Updated:** August 15, 2026
- **Branch:** `main`
- **Phase 1 scaffold:** complete and pushed in `3d40a51`; CI runtime update in `3b6a3d7`

**Overall Phase 1:** in progress — the server, plugin scaffold, CI packaging, and local MCP validation are complete; download-on-first-use remains deferred.

Completed foundation work:

- Minimal .NET 10 stdio MCP server using the official `ModelContextProtocol` C# SDK.
- Exactly two read-only tools: `ping` and `get_version`.
- Self-contained, single-file `win-x64` publish validated locally.
- Framework-free MCP smoke test lists and calls both tools successfully.
- Validated Codex plugin manifest under `plugin/excel-vba-mcp/`.
- GitHub Actions restores, builds, publishes, smoke-tests, packages, and uploads the Windows ZIP.
- Tagged `v*` release automation is implemented but has not yet been exercised by publishing a tag.

No Excel, COM, VBIDE, workbook, or VBA automation has been implemented.

## Phase 1 — Prove Download-on-First-Use

- [x] Create a minimal MCP server project in C#/.NET 10 using stdio transport.
- [x] Build it as a self-contained, single-file Windows executable (`ExcelVbaMcp.exe`).
- [x] Have the server expose only the read-only `ping` and `get_version` tools.
- [x] Create and push the GitHub repository for the project.
- [ ] Publish `ExcelVbaMcp.exe` as a GitHub Release asset. The `v*` tag workflow is ready; no release tag has been published yet.
- [ ] Define the permanent local installation location; `%LOCALAPPDATA%\ExcelVbaMcp\` is a candidate, not yet a decision.
- [ ] Give the test agent instructions to check the chosen location for `ExcelVbaMcp.exe`.
- [ ] If the executable does not exist, have the agent locate the latest GitHub Release and download it.
- [ ] Have the agent save the executable to the permanent local location.
- [ ] Have the agent launch the downloaded MCP server.
- [x] Verify an MCP client can connect to the locally published executable, list exactly both tools, and call `ping`/`get_version`.
- [x] End the local MCP smoke-test session and verify the server shuts down cleanly.
- [ ] Start a second session and verify the agent finds and reuses the existing executable without downloading it again.
- [ ] Test from a clean PC or user profile to verify the entire first-use process from scratch.
- [ ] Document the exact bootstrap instructions that reliably worked.

### Phase 1 Success Criteria

A machine with no existing VBA MCP installation can be given the repository/release location, automatically download the server, save it permanently, launch it, use its test tool, and reuse the same installation later.

**Status:** Not yet met. Executable discovery, installation location, integrity verification, first-use download, reuse, update, and rollback behavior are intentionally deferred. Accordingly, the plugin does not contain a fragile `.mcp.json` executable path.

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
- [ ] Allow an agent to compare the installed version with the latest release.
- [ ] Download updates only when appropriate.
- [ ] Verify release downloads with SHA-256 hashes or code signing.
- [ ] Preserve the previous working version for rollback.
- [ ] Document the agent update procedure.

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
