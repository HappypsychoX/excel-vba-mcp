# Excel VBA MCP Server — Development Checklist

## Current Status

- **Updated:** August 16, 2026
- **Branch:** `main`
- **Phase 1 server version:** `0.1.3`
- **Phase 2 server version:** `0.2.0`

**Overall Phase 1:** complete — the read-only server, CI packaging, tagged plugin release, protocol smoke test, and Codex plugin-host acceptance test are complete. Distribution follows the bundled-plugin model: the plugin contains its `.mcp.json` and self-contained `ExcelVbaMcp.exe`; no first-use GitHub download is part of plugin activation.

Completed foundation work:

- Minimal .NET 10 stdio MCP server using the official `ModelContextProtocol` C# SDK.
- Exactly two read-only tools: `ping` and `get_version`.
- Self-contained, single-file `win-x64` publish and framework-free MCP smoke test.
- Root-level Claude Code and Codex plugin manifests with host-specific bundled MCP configuration.
- GitHub Actions restores, builds, tests, publishes, assembles, smoke-tests, and uploads the plugin ZIP.
- Tagged releases package the plugin ZIP only.
- The generated `bin/win-x64/ExcelVbaMcp.exe` is ignored and is never committed to normal Git history.

Phase 2 adds read-only workbook discovery for Excel instances that the user already opened. It is not an Excel-control feature: it never creates, opens, saves, closes, or quits Excel or a workbook. The implementation, 14 automated unit tests, Excel-absent protocol smoke test, and recorded real-Excel lifecycle harness have passed.

## Phase 1 — Bundle and Activate the Read-Only Server

- [x] Create a minimal MCP server project in C#/.NET 10 using stdio transport.
- [x] Build it as a self-contained, single-file Windows executable (`ExcelVbaMcp.exe`).
- [x] Have the server expose only the read-only `ping` and `get_version` tools.
- [x] Create and push the GitHub repository for the project.
- [x] Publish the Phase 1 server in the `v0.1.0` GitHub Release.
- [x] Include `ExcelVbaMcp.exe` at `bin/win-x64/` inside the completed plugin package.
- [x] Add root `.mcp.json` for Codex and an inline MCP configuration for Claude Code.
- [x] Package a plugin ZIP containing its manifest, MCP configuration, and server executable.
- [x] Verify an MCP client can list exactly both tools and call `ping`/`get_version` against the bundled executable.
- [x] Install the root plugin in a clean Codex profile; verify host activation, tool listing, calls, and clean shutdown.
- [x] Confirm a second fresh Codex session reuses the installed plugin bundle without a network download.
- [ ] Validate the root plugin in Claude Code; this cross-host compatibility check is not a Phase 1 completion gate.

### Superseded download bootstrap work

The former first-use downloader that placed a release executable in `%LOCALAPPDATA%\ExcelVbaMcp\` is intentionally removed. Plugin installation already saves the server alongside its manifest and `.mcp.json`; adding a second distribution mechanism would duplicate the plugin model.

### Development and release packaging

- [x] Keep generated `bin/win-x64/ExcelVbaMcp.exe` out of Git; old commits may retain the historical binary.
- [x] Restore, build, and test the solution from a clean source checkout before packaging.
- [x] Publish a self-contained, single-file `win-x64` executable to generated output.
- [x] Assemble `.claude-plugin/`, `.codex-plugin/`, `.mcp.json`, and `bin/win-x64/ExcelVbaMcp.exe` into the plugin package.
- [x] Smoke-test the staged executable and the executable extracted from the completed ZIP.
- [x] Upload the ZIP as a workflow artifact and attach the same ZIP to `v*` GitHub Releases.
- [ ] Perform manual clean-install activation checks in Codex and Claude Code when release contents or manifests change.

Never stage or commit the published executable. Installation obtains the complete plugin ZIP once; activation launches the executable already stored in the local plugin cache and performs no network download.

### Phase 1 Success Criteria

A clean Codex profile can install the plugin, start its bundled server, list exactly `ping` and `get_version`, call both tools, end the MCP session cleanly, and start a later session from the same installed bundle without downloading a release.

**Status:** passed with plugin version `0.1.3`. After manual plugin installation and plugin-scoped MCP enablement in `config.toml`, two fresh Codex sessions started the bundled executable from the installed cache, listed exactly `ping` and `get_version`, returned `pong` and `0.1.3`, shut down cleanly, and left no `ExcelVbaMcp.exe` process. The second session reused the installed bundle without a release download.

---

## Phase 2 — Connect to Excel

`list_workbooks` is read-only, idempotent, and closed-world. It returns `{ excelRunning, workbooks }`; each workbook has `name`, `fullPath` (or `null` when unsaved), `saved`, and `readOnly`. When Excel is absent it returns `excelRunning: false` with an empty list and does not start Excel. COM failures and an inaccessible or busy Excel instance are reported as errors rather than as “Excel is not running.”

- [x] Add Excel COM interop support. **Automated evidence:** Release build passed with zero warnings or errors; the embedded Excel and Office Core interop types avoid a separately installed PIA dependency.
- [x] Detect running Excel instances. **Automated evidence:** the unit suite includes a direct Windows ROT locator regression through the STA dispatcher.
- [x] Add `list_workbooks`. **Automated evidence:** the protocol smoke test passed with exactly `get_version`, `list_workbooks`, and `ping`, and validated the Excel-absent JSON response.
- [x] Attach to an already-open workbook without taking ownership of Excel. **Automated evidence:** tests cover the transient reader contract and cleanup after a mid-enumeration failure.
- [x] Verify disconnecting from MCP does not close the user's Excel instance. **Real-Excel evidence (August 16, 2026):** after two successful calls and MCP disconnect, the test owner confirmed `Phase2SavedOne`, `Phase2SavedTwo`, and unsaved `Book3` were still open and each had `Worksheets.Count > 0`.
- [x] Verify all COM references are released cleanly. **Automated evidence:** 14 passing unit tests exercise reverse-order release after normal and mid-enumeration-failure paths, plus dispatcher cancellation and shutdown. The recorded real-Excel run also passed the lifecycle gate.
- [x] Verify no orphaned `EXCEL.EXE` processes are created. **Real-Excel evidence (August 16, 2026):** the external pre-check found zero Excel processes; the test-owned PID `29344` was unchanged during two calls, and after the owner closed Excel the external check again found zero. The harness exited `0` with `Real Excel lifecycle integration test passed.`

### Phase 2 acceptance status

Phase 2 is **complete** for server version `0.2.0`: the Release build passed with zero warnings/errors, 14 unit tests passed, the Excel-absent protocol smoke test passed, and the August 16, 2026 real-Excel harness passed. The harness ran outside the sandbox in the same interactive Windows user/session/integrity context as Excel; sandboxed processes cannot inspect the desktop Excel Running Object Table. Multi-instance identity and safe workbook targeting remain Phase 8 limitations.

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
