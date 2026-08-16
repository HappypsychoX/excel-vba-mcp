# Phase 2 real-Excel verification

Run this procedure on Windows 10 or Windows 11 x64 with Microsoft Excel desktop installed. It is a manual acceptance gate, separate from the automated unit and protocol smoke tests. Run the harness outside a workspace sandbox and in the same interactive Windows user, session, and integrity context as Excel; the Windows Running Object Table is not visible across that desktop-security boundary.

## Preconditions

- Build and publish the server as described in the [README](../README.md).
- Close all Excel windows. In PowerShell, record the baseline process IDs:

  ```powershell
  Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id
  ```

- Start Excel yourself and open two known workbooks. Keep one saved workbook and create a new unsaved workbook (`Book1` is sufficient). Record the current `EXCEL.EXE` process IDs and workbook names.
- Do not use a server build running as a different Windows user or session from Excel.

## Procedure

1. Start the published MCP server through the smoke-test client or an MCP host, then call `list_workbooks`.
2. Verify the saved workbook reports its display name, full path, `saved: true`, and the expected `readOnly` value.
3. Verify the unsaved workbook reports its display name, `fullPath: null`, and `saved: false`.
4. Call `list_workbooks` repeatedly (at least ten times). Confirm the expected workbooks remain present and that no extra `EXCEL.EXE` PID appears.
5. End the MCP session cleanly. Confirm Excel remains open, both workbooks remain open, and Excel is still responsive. Make a harmless edit in the unsaved workbook, then undo it.
6. Close Excel manually, saving or discarding changes as appropriate. Confirm all Excel processes that existed for the test exit normally and that no orphaned `EXCEL.EXE` process remains.
7. If possible, repeat with two separately started Excel processes and duplicate workbook names. Record any workbook or instance ambiguity; this is an expected Phase 8 limitation, not permission to choose an instance for a mutation.

## Failure-enumeration check

Exercise the test seam that makes workbook enumeration fail after at least one workbook has been acquired. Confirm the operation returns a clear error and subsequent calls continue to work. The automated test suite covers deterministic cleanup for this path; the real-Excel run establishes that the same lifecycle is safe with Office.

## Evidence to record

Record the date, Windows and Excel versions, server version/commit, initial and final Excel process IDs, the returned workbook metadata, number of repeated calls, and pass/fail for each step. Attach that record to the release or pull request before checking the remaining Phase 2 lifecycle and orphan-process checklist items.

Do not mark the real-Excel lifecycle or orphan-process checks complete merely because unit tests and the smoke test pass.

## Recorded Phase 2 evidence

On August 16, 2026, the `0.2.0` server passed this harness outside the workspace sandbox. The external pre-check found zero `EXCEL.EXE` processes. The test owner started Excel PID `29344`, opened `Phase2SavedOne` and `Phase2SavedTwo`, and created unsaved `Book3`. Two `list_workbooks` calls succeeded with the PID set unchanged. After MCP disconnect, the owner confirmed all three workbooks remained open and usable (`Worksheets.Count > 0`), then closed the test-owned Excel instance. The external post-check again found zero `EXCEL.EXE` processes, and the harness exited `0` with `Real Excel lifecycle integration test passed.`
