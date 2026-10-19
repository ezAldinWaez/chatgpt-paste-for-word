# Releasing

Release builds need Windows, the 64-bit .NET Framework compiler, and Office/Extensibility primary interop assemblies. See [Microsoft's PIA documentation](https://learn.microsoft.com/en-us/visualstudio/vsto/office-primary-interop-assemblies?view=visualstudio) for development-machine setup. `scripts/build.ps1` accepts explicit `-OfficeAssembly` and `-ExtensibilityAssembly` paths when auto-discovery cannot find them.

1. Update `VERSION`, matching source version strings, `CHANGELOG.md`, and the supported/tested version statements in `README.md`. Keep the Pandoc URL and both expected hashes in `src/Setup.cs` synchronized when updating Pandoc.
2. Run `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/package-release.ps1`. The command creates `dist/ChatGPT-Paste-for-Word-Setup-VERSION.exe` and `dist/SHA256SUMS.txt`; `dist/` is intentionally ignored by Git.
3. Complete [testing.md](testing.md) on a clean Windows/Word installation. Check the installer hash with `Get-FileHash -Algorithm SHA256 dist\ChatGPT-Paste-for-Word-Setup-VERSION.exe`.
4. Commit the source, tag it `vVERSION`, and push the branch and tag to the GitHub repository. Creating a GitHub release publishes the installer to others, so review the final artifacts and release notes first.
5. With an authenticated GitHub CLI, create a **draft** release and attach the installer, checksum file, README, and license:

   ```powershell
   gh release create vVERSION dist/ChatGPT-Paste-for-Word-Setup-VERSION.exe dist/SHA256SUMS.txt dist/README.md dist/LICENSE --draft --verify-tag --title "ChatGPT Paste for Word vVERSION" --notes-file CHANGELOG.md
   ```

   Replace `VERSION` with the number in the `VERSION` file. Review the draft on GitHub, then publish it.

The repository's validation workflow checks source consistency on pushes and pull requests. The Windows installer is built locally because the Ribbon depends on Office interop assemblies, and final behavior must be checked in a real Word session. Never commit downloaded Pandoc or Office binaries to this repository. The installer fetches Pandoc from its pinned upstream release and checks its hash during installation.
