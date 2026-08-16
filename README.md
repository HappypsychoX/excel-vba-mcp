# Excel VBA MCP

Excel VBA MCP is a local [Model Context Protocol](https://modelcontextprotocol.io/) server for safely inspecting Excel workbooks and, in later phases, VBA. Phase 2 is deliberately read-only: it attaches briefly to Excel instances that the user already opened, reports their workbooks, and then releases its COM references. It never starts, opens, saves, closes, or quits Excel.

## Prerequisites

- Windows 10 or Windows 11 x64.
- Microsoft Excel desktop installed for the same interactive Windows user who runs the MCP host. Excel must already be open before `list_workbooks` can find it.
- .NET 10 SDK for development, build, and test. The published `win-x64` bundle is self-contained and does not require the SDK or a separately installed .NET runtime.

Phase 2 does not require enabling **Trust access to the VBA project object model**. That setting becomes relevant only when a later phase reads VBA through VBIDE.

The server connects only to Excel in the current interactive Windows user, session, and integrity context. It cannot inspect Excel running under another account, session, or desktop-security boundary. In particular, a sandboxed process cannot see the desktop Excel Running Object Table; run the real-Excel harness outside that sandbox with the same interactive desktop token as Excel.

## Tools

| Tool | Annotations | Purpose |
| --- | --- | --- |
| `ping` | Read-only | Confirms that the server is responsive. |
| `get_version` | Read-only | Returns the running assembly/package version. |
| `list_workbooks` | Read-only, idempotent, closed-world | Lists metadata for workbooks currently open in discoverable Excel instances. |

`list_workbooks` takes no arguments and returns plain serializable data with this shape:

```json
{
  "excelRunning": true,
  "workbooks": [
    {
      "name": "Budget.xlsx",
      "fullPath": "C:\\Work\\Budget.xlsx",
      "saved": true,
      "readOnly": false
    },
    {
      "name": "Book1",
      "fullPath": null,
      "saved": false,
      "readOnly": false
    }
  ]
}
```

If Excel is not running, the valid response is `{"excelRunning":false,"workbooks":[]}`. The call does not start Excel. If Excel is running with no open workbooks, it returns `excelRunning: true` and an empty array. An inaccessible, busy, or failed COM attachment is reported as an MCP tool error, rather than being represented as “Excel is not running.”

### Attach, but never own

Each call temporarily locates an existing Excel COM application, reads workbook metadata on the server’s dedicated STA dispatcher, converts it to .NET records, and releases temporary COM references before returning. The server never calls `new Excel.Application`, `Application.Quit`, `Workbook.Close`, `Workbooks.Open`, or `Workbook.Save`; it retains no Excel runtime-callable wrapper between calls. Ending the MCP session must leave the user’s Excel process and workbooks open and usable.

Do not treat `list_workbooks` as a workbook-selection mechanism for future write actions. It exposes no COM objects and no public `attach_workbook` tool.

### Multi-instance limitation

The server attempts to enumerate every Excel application discoverable through the Windows Running Object Table, but Phase 2 does not yet provide a durable instance identity or a way to target a particular workbook. Duplicate workbook names and multiple Excel processes can therefore be ambiguous, and some processes may not be discoverable through the ROT. Comprehensive multi-instance discovery, identity, and safe targeting are Phase 8 work; `list_workbooks` remains read-only until that work is complete.

## Architecture and repository layout

```text
.
|-- .github/workflows/ci.yml                 Build, package, and tagged-release automation
|-- .claude-plugin/plugin.json               Claude Code marketplace manifest
|-- .codex-plugin/plugin.json                Codex marketplace manifest
|-- .mcp.json                                Codex bundled MCP configuration (`mcpServers`)
|-- bin/win-x64/ExcelVbaMcp.exe              Generated locally; ignored by Git
|-- src/ExcelVbaMcp.Server/                  MCP host, COM dispatcher, locator, reader, and tools
|-- tests/                                   Protocol smoke test and unit tests
|-- docs/phase-2-real-excel-verification.md  Windows/Excel manual verification procedure
|-- Directory.Build.props                    Product and assembly version metadata
`-- Excel VBA MCP Server — Development Checklist.md
```

The repository contains the plugin manifests and server source, but a clean source checkout does not contain `bin/win-x64/ExcelVbaMcp.exe`. That generated file is ignored by Git. CI publishes the server and assembles the installable plugin ZIP, whose root folder contains both manifests, `.mcp.json`, and `bin/win-x64/ExcelVbaMcp.exe`.

In the completed package, Codex discovers `.codex-plugin/plugin.json`, which references the root `.mcp.json` companion file. That JSON file uses the plugin-schema `mcpServers` key; the similarly named `mcp_servers` key appears only in Codex TOML settings. Claude Code discovers `.claude-plugin/plugin.json`, which carries inline MCP configuration using `CLAUDE_PLUGIN_ROOT`. Both start the executable already bundled in the installed plugin; neither downloads a GitHub release on activation. [Claude Code plugin documentation](https://code.claude.com/docs/en/plugins-reference) and [Codex plugin documentation](https://developers.openai.com/plugins/build/plugins)

Logging is sent to stderr because stdout is reserved for MCP protocol messages.

## Build, test, and run locally

From a Developer PowerShell in the repository root:

```powershell
dotnet restore ExcelVbaMcp.slnx
dotnet build ExcelVbaMcp.slnx --configuration Release --no-restore
dotnet test tests/ExcelVbaMcp.Tests/ExcelVbaMcp.Tests.csproj `
  --configuration Release --no-build --no-restore
```

The unit tests use test doubles for Excel/COM and do not require Office. They cover the no-running-instance response, empty and populated results, unsaved workbook paths, duplicate names, COM-error translation, dispatcher cancellation/shutdown, and cleanup after a failed enumeration.

Publish the executable to generated output:

```powershell
dotnet publish src/ExcelVbaMcp.Server/ExcelVbaMcp.Server.csproj `
  --configuration Release `
  --runtime win-x64 `
  --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:DebugType=None `
  -p:DebugSymbols=false `
  --output artifacts/publish
```

Run the protocol smoke test against the published executable:

```powershell
dotnet run --project tests/ExcelVbaMcp.SmokeTest/ExcelVbaMcp.SmokeTest.csproj `
  --configuration Release `
  --no-build `
  -- artifacts/publish/ExcelVbaMcp.exe
```

The smoke test expects exactly `get_version`, `list_workbooks`, and `ping`. It is valid to run it on a machine without Excel: in that case it verifies the well-formed `excelRunning: false` response and confirms that the server does not launch Excel.

For the Windows/Excel lifecycle and process checks, follow [the Phase 2 real-Excel verification procedure](docs/phase-2-real-excel-verification.md). Those checks are manual acceptance gates; unit and smoke tests cannot prove that the user’s Excel processes remain usable or that no orphaned `EXCEL.EXE` process exists.

After publishing, the Windows-only harness can be run as follows (use two saved workbook display names that are unique among all currently open workbooks):

```powershell
dotnet run --project tests/ExcelVbaMcp.RealExcelIntegration `
  --configuration Release -- `
  --server .\artifacts\publish\ExcelVbaMcp.exe `
  --workbook SavedOne.xlsx `
  --workbook SavedTwo.xlsx `
  --unsaved Book3
```

The harness records `EXCEL.EXE` process IDs, asks the tester to open the workbooks, calls `list_workbooks` twice, checks the metadata and process set, then waits for the tester to confirm usability and manually close Excel. Run it outside any workspace sandbox and in Excel’s interactive desktop token. Its exit status is the evidence for the lifecycle and orphan-process checklist gates.

## Install and activation model

Installation obtains the complete packaged plugin once, including `bin/win-x64/ExcelVbaMcp.exe`, and stores it in the host's local plugin cache. Activation then launches that already-cached executable. Activation is network-free: there is no repository clone, bootstrap script, GitHub API call, release download, `%LOCALAPPDATA%` installation, or first-use downloader.

Plugin installation and plugin-server activation are separate settings. After installing and enabling the `excel-vba-mcp` plugin, enable its bundled MCP server in `~/.codex/config.toml` (on Windows, `%USERPROFILE%\.codex\config.toml`):

```toml
[plugins."excel-vba-mcp".mcp_servers."excel-vba-mcp"]
enabled = true
default_tools_approval_mode = "prompt"
```

Fully restart the Codex desktop app, then start a new session. The session should list `ping`, `get_version`, and `list_workbooks`.

The **Connect to a custom MCP** screen is a useful diagnostic but is not the plugin activation path: configuring the cached executable there creates a separate, manually managed MCP server and bypasses the bundled-plugin setting.

## Release process

1. Update `VersionPrefix`, `AssemblyVersion`, and `FileVersion` in `Directory.Build.props`.
2. Keep both root plugin manifests in sync with `Directory.Build.props`.
3. Build and test the solution, publish to `artifacts/publish`, and run the smoke test using the commands above.
4. Assemble a plugin staging directory containing `.claude-plugin/`, `.codex-plugin/`, `.mcp.json`, and the published executable at `bin/win-x64/ExcelVbaMcp.exe`.
5. Create the plugin ZIP, inspect and extract it, then smoke-test the executable from the extracted package.
6. Complete the manual real-Excel verification on a Windows machine with Excel before declaring a Phase 2 release complete.
7. Commit source, manifests, documentation, and workflow changes only. Never add the generated executable to Git. Create a matching annotated tag such as `v0.2.0`, then push.

Every push and pull request restores, builds, tests, publishes, assembles, and smoke-tests the server from source. CI produces one installable `excel-vba-mcp-plugin-win-x64.zip` with this required layout:

```text
excel-vba-mcp/
|-- .claude-plugin/
|-- .codex-plugin/
|-- .mcp.json
`-- bin/win-x64/ExcelVbaMcp.exe
```

CI smoke-tests the executable in the staging package and again after extracting the ZIP. The ZIP is uploaded as a workflow artifact. A `v*` tag creates or updates the matching GitHub Release and attaches that same ZIP.

## Scope and safety boundary

Phase 2 does not read VBA, modify VBA, save workbooks, run macros, or change any Excel state. All future workbook and VBA changes must go through Excel/COM/VBIDE rather than rewriting Office package internals such as `.xlsm`, `.xlsb`, or `vbaProject.bin`. See the [development checklist](Excel%20VBA%20MCP%20Server%20%E2%80%94%20Development%20Checklist.md) for the planned sequence and verification record.
