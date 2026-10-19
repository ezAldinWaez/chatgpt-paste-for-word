# Contributing

This is a Windows desktop Word add-in with a local PowerShell clipboard converter. Before changing behavior, read [docs/architecture.md](docs/architecture.md) and run `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/test-source.ps1`.

For a full build, install 64-bit .NET Framework 4.x and Office/Visual Studio primary interop assemblies, then run `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/build.ps1`. `scripts/build.ps1` accepts `-OfficeAssembly` and `-ExtensibilityAssembly` paths if those assemblies are not discoverable on your machine. It embeds the used interop types; it does not bundle the primary interop assemblies. Word is required for integration testing.

Please include the Word version, Windows version, whether you used ChatGPT's Copy button or selected text, and whether the cursor was in an LTR or RTL paragraph when reporting paste bugs. A small redacted example of the source HTML/Markdown and resulting Word layout is more useful than a screenshot alone. Avoid posting private clipboard contents or `ribbon.log` without checking them first.

Keep the version in `VERSION`, all source assembly versions, the ribbon version, the installer display version, and the ribbon DLL filename in sync. The validation script checks these. Follow the manual Word checks in [docs/testing.md](docs/testing.md) before preparing a release.

Pull requests are welcome. Do not commit build output, local logs, downloaded Pandoc binaries, or Office interop assemblies.
