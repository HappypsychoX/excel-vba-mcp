# Excel VBA MCP

Excel VBA MCP is a local Model Context Protocol server intended to let MCP clients work safely with Microsoft Excel and VBA. Phase 1 establishes only the distributable server and plugin foundations. It deliberately performs no Excel automation.

## Phase 1 scope

The .NET 10 server uses the official `ModelContextProtocol` C# SDK over stdio and exposes exactly two read-only tools:

| Tool | Result | Purpose |
| --- | --- | --- |
| `ping` | `pong` | Confirm that the server is responsive. |
| `get_version` | Assembly/package version | Identify the running build. |

There are no prompts, resources, Excel APIs, workbook operations, or write-capable tools.

## Architecture and repository layout

```text
.
|-- .github/workflows/ci.yml                 Build, package, and tagged-release automation
|-- plugin/excel-vba-mcp/
|   `-- .codex-plugin/plugin.json            Valid Codex plugin metadata
|-- src/ExcelVbaMcp.Server/
|   |-- Program.cs                           Minimal process entry point
|   |-- ServerHost.cs                        Host, logging, stdio transport, and tool registration
|   |-- ServerVersion.cs                     Assembly/package version lookup
|   `-- Tools/PhaseOneTools.cs               The two Phase 1 tool definitions
|-- tests/ExcelVbaMcp.SmokeTest/             Framework-free MCP client smoke test
|-- Directory.Build.props                    Product and assembly version metadata
|-- Directory.Packages.props                 Central NuGet dependency versions
`-- Excel VBA MCP Server — Development Checklist.md
```

Logging is sent to stderr because stdout is reserved for MCP protocol messages.

## Build and run locally

Install the .NET 10 SDK, then run:

```powershell
dotnet restore ExcelVbaMcp.slnx
dotnet build ExcelVbaMcp.slnx --configuration Release --no-restore
dotnet run --project src/ExcelVbaMcp.Server/ExcelVbaMcp.Server.csproj
```

The final command waits for an MCP client on stdin/stdout; it is not an interactive shell.

To produce the same Windows artifact as CI:

```powershell
dotnet publish src/ExcelVbaMcp.Server/ExcelVbaMcp.Server.csproj `
  --configuration Release `
  --runtime win-x64 `
  --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  --output artifacts/publish
```

Run the end-to-end MCP smoke test against that executable:

```powershell
dotnet run --project tests/ExcelVbaMcp.SmokeTest/ExcelVbaMcp.SmokeTest.csproj `
  -- artifacts/publish/ExcelVbaMcp.exe
```

## Codex plugin status

`plugin/excel-vba-mcp/.codex-plugin/plugin.json` is a valid metadata-only plugin scaffold. It intentionally has no `.mcp.json`: a repository-relative path cannot reliably locate an executable downloaded to an arbitrary machine, and no bundled-executable resolver or verified installer exists yet.

Activation/bootstrap work must establish a durable install location, download and integrity-verification policy, update/rollback behavior, and a reliable command path before an MCP server entry is added. Until then, clients can launch a locally built or published executable through an explicit user configuration.

## Release process

1. Update `VersionPrefix`, `AssemblyVersion`, and `FileVersion` in `Directory.Build.props`.
2. Keep the version in `plugin/excel-vba-mcp/.codex-plugin/plugin.json` in sync.
3. Restore, build, publish, validate the plugin, and run the smoke test.
4. Commit the release and create a matching annotated tag such as `v0.1.0`.
5. Push the commit and tag.

Every push and pull request restores and builds the solution, publishes a self-contained single-file `win-x64` executable, packages it as `excel-vba-mcp-win-x64.zip`, and uploads the ZIP as a workflow artifact. A `v*` tag also creates or updates the corresponding GitHub release and attaches the ZIP for direct download.

## Phase 1 checklist

- [x] Create a minimal C#/.NET 10 stdio MCP server.
- [x] Use the official `ModelContextProtocol` C# SDK.
- [x] Expose only read-only `ping` and `get_version` tools.
- [x] Centralize assembly/package version metadata.
- [x] Add a valid Codex plugin scaffold without a fragile executable path.
- [x] Restore and build in GitHub Actions on pushes and pull requests.
- [x] Publish and package a self-contained, single-file `win-x64` executable.
- [x] Upload CI artifacts and attach ZIPs to `v*` GitHub releases.
- [x] Provide a protocol-level smoke test for tool discovery and invocation.
- [ ] Design and validate download-on-first-use installation/bootstrap.
- [ ] Test install, reuse, update, and rollback from a clean Windows profile.

The original, longer-term development checklist remains in [`Excel VBA MCP Server — Development Checklist.md`](./Excel%20VBA%20MCP%20Server%20%E2%80%94%20Development%20Checklist.md).

## Explicitly deferred

The following are out of Phase 1 and are not implemented:

- Excel automation and COM interop
- VBIDE access and trust-setting detection
- Workbook discovery, attachment, reading, saving, or other operations
- VBA discovery, reading, editing, validation, or execution
- Installer, download-on-first-use, executable discovery, integrity verification, updates, and rollback
- Plugin `.mcp.json` activation until executable resolution is reliable

See the development checklist for the planned sequence. All future workbook and VBA changes must go through Excel/COM/VBIDE rather than rewriting Office package internals.
