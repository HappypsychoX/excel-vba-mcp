# Excel VBA MCP

Excel VBA MCP is a local Model Context Protocol server intended to let MCP clients work safely with Microsoft Excel and VBA. Phase 1 establishes the distributable server and Codex plugin foundations only; it performs no Excel automation.

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
|   |-- .codex-plugin/plugin.json            Plugin metadata and MCP configuration reference
|   |-- .mcp.json                            Bundled MCP server definition
|   `-- bin/win-x64/ExcelVbaMcp.exe           Self-contained bundled server
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

The plugin is the distribution unit: when Codex installs it, the plugin manifest, `.mcp.json`, and `ExcelVbaMcp.exe` are saved together. The MCP definition starts the executable bundled in that plugin directory; it does not download a GitHub release on first use. Codex documents `.mcp.json` as the mechanism for distributing an MCP server with a plugin, with the manifest referring to it from the plugin root. [Codex plugin documentation](https://developers.openai.com/plugins/build/plugins)

Logging is sent to stderr because stdout is reserved for MCP protocol messages.

## Build and run locally

Install the .NET 10 SDK, then run:

```powershell
dotnet restore ExcelVbaMcp.slnx
dotnet build ExcelVbaMcp.slnx --configuration Release --no-restore
dotnet run --project src/ExcelVbaMcp.Server/ExcelVbaMcp.Server.csproj
```

The final command waits for an MCP client on stdin/stdout; it is not an interactive shell.

Publish the executable into the plugin bundle:

```powershell
dotnet publish src/ExcelVbaMcp.Server/ExcelVbaMcp.Server.csproj `
  --configuration Release `
  --runtime win-x64 `
  --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:DebugType=None `
  -p:DebugSymbols=false `
  --output plugin/excel-vba-mcp/bin/win-x64
```

Run the protocol smoke test against the bundled executable:

```powershell
dotnet run --project tests/ExcelVbaMcp.SmokeTest/ExcelVbaMcp.SmokeTest.csproj `
  --configuration Release `
  --no-build `
  -- plugin/excel-vba-mcp/bin/win-x64/ExcelVbaMcp.exe
```

## Install and activation model

Install the packaged plugin in Codex and enable it. The plugin brings the server executable with it, so there is no repository clone, bootstrap script, GitHub API call, or release download during first use.

`%LOCALAPPDATA%\ExcelVbaMcp\ExcelVbaMcp.exe` is not used by this distribution model. The plugin cache is the sole installation location; no standalone installer or direct-release download is planned.

The precise host-side activation test is still an acceptance check: install the plugin into a clean Codex profile and confirm that the host resolves the root-relative command in `.mcp.json`, lists exactly `ping` and `get_version`, calls both, and shuts the server down when the MCP session ends.

## Release process

1. Update `VersionPrefix`, `AssemblyVersion`, and `FileVersion` in `Directory.Build.props`.
2. Keep the version in `plugin/excel-vba-mcp/.codex-plugin/plugin.json` in sync.
3. Publish the server into `plugin/excel-vba-mcp/bin/win-x64/` using the command above.
4. Validate the plugin manifest and run the smoke test against the bundled executable.
5. Commit the source and bundled executable, create a matching annotated tag such as `v0.1.0`, then push.

Every push and pull request restores, builds, publishes, and smoke-tests the server. CI produces one installable `excel-vba-mcp-plugin-win-x64.zip` containing the plugin manifest, MCP definition, and executable. A `v*` tag creates or updates the GitHub release and attaches that ZIP.

## Phase 1 checklist

- [x] Create a minimal C#/.NET 10 stdio MCP server.
- [x] Use the official `ModelContextProtocol` C# SDK.
- [x] Expose only read-only `ping` and `get_version` tools.
- [x] Centralize assembly/package version metadata.
- [x] Bundle the executable and `.mcp.json` with a valid Codex plugin manifest.
- [x] Restore and build in GitHub Actions on pushes and pull requests.
- [x] Publish and package a self-contained, single-file `win-x64` executable.
- [x] Upload CI artifacts and attach ZIPs to `v*` GitHub releases.
- [x] Provide a protocol-level smoke test for tool discovery and invocation.
- [ ] Validate plugin installation and activation from a clean Codex profile.

The original, longer-term development checklist remains in [Excel VBA MCP Server — Development Checklist.md](./Excel%20VBA%20MCP%20Server%20%E2%80%94%20Development%20Checklist.md).

## Explicitly deferred

The following are out of Phase 1 and are not implemented:

- Excel automation and COM interop
- VBIDE access and trust-setting detection
- Workbook discovery, attachment, reading, saving, or other operations
- VBA discovery, reading, editing, validation, or execution
- Plugin upgrade, integrity verification, and rollback behavior

See the development checklist for the planned sequence. All future workbook and VBA changes must go through Excel/COM/VBIDE rather than rewriting Office package internals.
