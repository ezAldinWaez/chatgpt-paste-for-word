# Manual release checks

Run `scripts/package-release.ps1` and confirm that the installer, `README.md`, `LICENSE`, and `SHA256SUMS.txt` appear in `dist/`. Check the installer hash against the manifest. On a Windows machine with desktop Word:

1. Install for the current user. Confirm the **ChatGPT Paste** button appears on Word's Home ribbon and dedicated tab after restarting every Word process.
2. Copy a ChatGPT response using its Copy button. Paste a mixed article with headings, a list, a table, inline emphasis, inline code, and displayed equations. Confirm headings map to Word heading styles, equations remain editable, and body text follows the destination Normal style/theme font.
3. Repeat with selected ChatGPT text and with plain Markdown text. Confirm the visual button and Ctrl+Alt+Shift+V both work.
4. Paste into an Arabic RTL paragraph and into an LTR paragraph. Check paragraph direction, table column order, numbered headings/captions, punctuation next to Arabic text, decimal numbers, parentheses, and Latin scientific units. Show Word's paragraph marks for the spacing checks.
5. Run the same installer again and choose Update. Reopen Word and confirm the button still works.
6. Run Uninstall, including once while Word is open. Confirm the button/shortcut stop working, the startup entry and Word add-in registration are gone, and any locked add-in files are removed after Word exits.

Automated `scripts/test-source.ps1` checks version consistency, installer resource references, PowerShell syntax, and all nine Windows icon payloads (including their actual color depth). The GitHub workflow runs only these source checks because a hosted runner is not a substitute for a Word desktop integration test.

## v0.1.1 release validation — 2026-10-19

The Windows build passed source checks and compiled all three assemblies. Installer payloads matched the built files; assembly/product versions and the installer checksum matched. Windows loaded all nine icon sizes with transparent corners and the filled clipboard artwork. Both executables contained the custom icon, and both ribbon buttons returned the cached OLE image.

The actual ribbon paste implementation from the packaged DLL was exercised in a separate Word 16.0 instance using temporary documents: Markdown in LTR and RTL contexts, and Arabic HTML in an RTL context. Each fixture retained two editable equations, a heading, two list items, and one table. Pasted body styles and nonempty paragraph/table directions matched the destination. The clipboard was restored and test documents closed without saving.

This patch's checks used extracted installer payloads; clean-profile installation, update, and uninstall were not rerun.
