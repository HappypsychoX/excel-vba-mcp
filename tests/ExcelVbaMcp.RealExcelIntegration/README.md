# Real Excel integration harness

This Windows-only harness validates the Phase 2 lifecycle against an Excel instance that the tester starts and owns. It is intentionally excluded from CI: GitHub-hosted runners do not supply a supported interactive Excel installation.

Before running it, manually create or choose two saved workbooks with distinct display names, then decide the display name Excel will assign to a new unsaved workbook (for example, `Book3`). Publish the server first, then run:

```powershell
dotnet run --project tests/ExcelVbaMcp.RealExcelIntegration --configuration Release -- `
  --server .\artifacts\publish\ExcelVbaMcp.exe `
  --workbook SavedOne.xlsx `
  --workbook SavedTwo.xlsx `
  --unsaved Book3
```

The harness records existing `EXCEL.EXE` process IDs, asks the tester to open Excel and the two saved workbooks plus the unsaved workbook, then calls `list_workbooks` twice. It verifies saved paths, the unsaved null path, the repeated response, and that the Excel process set did not change during the MCP session. It then asks the tester to confirm the workbooks remain usable, close Excel manually, and verifies that every Excel process started for the test has exited.

Do not use an existing workbook whose name is duplicated in another open Excel instance. The test treats duplicate returned names as a failure so the expected metadata is unambiguous.
