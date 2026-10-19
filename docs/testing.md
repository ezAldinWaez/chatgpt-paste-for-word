# Manual release checks

Run `scripts/package-release.ps1` and confirm that the installer, `README.md`, `LICENSE`, and `SHA256SUMS.txt` appear in `dist/`. Check the installer hash against the manifest. On a Windows machine with desktop Word:

1. Install for the current user. Confirm the **ChatGPT Paste** button appears on Word's Home ribbon and dedicated tab after restarting every Word process.
2. Copy a ChatGPT response using its Copy button. Paste a mixed article with headings, a list, a table, inline emphasis, inline code, and displayed equations. Confirm headings map to Word heading styles, equations remain editable, and body text follows the destination Normal style/theme font.
3. Repeat with selected ChatGPT text and with plain Markdown text. Confirm the visual button and Ctrl+Alt+Shift+V both work.
4. Paste into an Arabic RTL paragraph and into an LTR paragraph. Check paragraph direction, table column order, numbered headings/captions, punctuation next to Arabic text, decimal numbers, parentheses, and Latin scientific units. Show Word's paragraph marks for the spacing checks.
5. Run the same installer again and choose Update. Reopen Word and confirm the button still works.
6. Run Uninstall, including once while Word is open. Confirm the button/shortcut stop working, the startup entry and Word add-in registration are gone, and any locked add-in files are removed after Word exits.

Automated `scripts/test-source.ps1` checks version consistency, installer resource references, and PowerShell syntax. The GitHub workflow runs only these source checks because a hosted runner is not a substitute for a Word desktop integration test.
