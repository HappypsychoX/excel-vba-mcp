# Phase 2 — Connect to Excel

Phase 2 remains strictly read-only: connect transiently to Excel instances the user already opened, enumerate workbooks, and never create, open, save, close, or quit anything.

## 1. Define the Phase 2 contract

Add a structured `list_workbooks` response:

```text
{
  excelRunning: boolean,
  workbooks: [
    {
      name: string,
      fullPath: string | null,
      saved: boolean,
      readOnly: boolean
    }
  ]
}
```

Required behavior:

- When Excel is absent, return `excelRunning: false` and an empty collection without starting Excel.
- When Excel is running with no workbooks, return `excelRunning: true` and an empty collection.
- For unsaved workbooks, return the display name and a null path.
- Return a useful MCP error for a COM failure or inaccessible instance, distinct from “Excel is not running.”
- Mark the tool `ReadOnly`, `Idempotent`, and `OpenWorld = false`.

Phase 2 does not add a public `attach_workbook` tool. Attachment is an internal, temporary operation performed while listing workbooks.

## 2. Add the COM architecture

Introduce three separable components:

- `ExcelComDispatcher`: a dedicated STA thread that serializes all Excel COM work.
- `ExcelInstanceLocator`: finds already-running Excel applications through the Windows Running Object Table and OLE APIs.
- `ExcelWorkbookReader`: enumerates workbook metadata without modifying Excel.

Register these services through dependency injection in `ServerHost`. Keep COM logic out of the MCP tool class.

Follow these ownership and lifetime rules:

- Never call `new Excel.Application`.
- Never call `Application.Quit`, `Workbook.Close`, `Workbooks.Open`, or `Workbook.Save`.
- Do not retain Excel runtime callable wrappers between tool calls during Phase 2.
- Avoid COM `foreach`; use indexed access so enumerators cannot leak.
- Release workbook, collection, and application references deterministically in reverse order.
- Keep all COM acquisition, use, and release on the STA dispatcher thread.

Use early-bound Excel interop for maintainability, with embedded interop types so the plugin does not require separately installed primary interop assemblies. Continue publishing a self-contained `win-x64` executable.

## 3. Implement instance discovery and safe attachment

Discovery should:

1. Query the Running Object Table for existing Excel objects.
2. Resolve workbook entries back to their owning Excel application.
3. Deduplicate applications using a stable runtime identifier such as Excel’s window handle or process ID.
4. Enumerate each reachable application’s open workbooks.
5. Release every temporary COM reference before returning plain .NET records.

Comprehensive multi-instance handling remains part of Phase 8. Phase 2 must document any remaining limitations and must not silently select a workbook for future mutation. Prefer returning all discoverable workbooks with an instance identity where practical.

## 4. Add `list_workbooks`

Create a new injected tool class instead of expanding the static `PhaseOneTools` implementation.

The tool should:

- Accept no arguments.
- Call the workbook reader through the STA dispatcher.
- Return only plain serializable records, never COM objects.
- Honor cancellation while queued.
- Apply a bounded operation timeout and return a clear error if Excel is busy or unresponsive.
- Preserve the existing behavior of `ping` and `get_version`.

Update the smoke-test expectation to:

```text
get_version
list_workbooks
ping
```

The CI smoke test must tolerate machines without Excel by validating the well-formed “Excel not running” response.

## 5. Add automated and Excel-backed tests

Create a normal test project for logic that does not require Office, covering:

- No-running-instance mapping.
- Empty and populated workbook responses.
- Unsaved workbook path handling.
- Duplicate or ambiguous names.
- COM exception translation.
- Cancellation and dispatcher shutdown.
- Guaranteed cleanup when enumeration fails midway.

Create a Windows-only local integration harness for real Excel:

1. Record existing `EXCEL.EXE` process IDs.
2. Have the tester open Excel and two known workbooks, including one unsaved workbook.
3. Start the MCP server and call `list_workbooks`.
4. Verify the expected workbook metadata.
5. Repeat the call to expose reference leaks.
6. End the MCP session cleanly.
7. Verify Excel and its workbooks remain open and usable.
8. Close Excel manually and verify all original Excel processes exit.
9. Confirm no additional `EXCEL.EXE` process was created at any point.

Also test a failure during enumeration to ensure partial COM acquisition is cleaned up.

## 6. Update documentation and complete Phase 2

Update the README to describe:

- Windows and Microsoft Excel prerequisites.
- The new `list_workbooks` tool and response.
- The “attach but never own” lifecycle.
- Known multi-instance limitations deferred to Phase 8.
- How to run the real-Excel integration test.

Mark Phase 2 complete in the development checklist only after all acceptance checks pass.

## Implementation sequence

1. Add testable workbook result models and COM abstractions.
2. Implement the STA dispatcher and COM-release helpers.
3. Implement Running Object Table discovery and workbook enumeration.
4. Add the injected MCP tool.
5. Extend protocol smoke testing.
6. Run real-Excel lifecycle and process tests.
7. Update documentation, publish the bundled executable, and perform a clean-profile plugin test.

## Completion criteria

Phase 2 is complete when:

- `list_workbooks` reliably reports already-open workbooks.
- Calling it while Excel is absent does not start Excel and returns a valid empty response.
- Disconnecting MCP leaves the user’s Excel instance and workbooks open and usable.
- All COM references are released after each operation and during server shutdown.
- Repeated calls do not prevent Excel from exiting normally.
- The server never creates an orphaned `EXCEL.EXE` process.
- The build, publish, protocol smoke test, and real-Excel integration test all pass.
