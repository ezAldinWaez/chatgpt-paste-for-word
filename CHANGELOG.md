# Changelog

## v0.1.1 (Alpha) — 2026-10-19

- Added a custom clipboard/ChatGPT icon to Word's ribbon buttons, the tray utility, setup, and Windows Installed Apps.
- Preserved full color and transparency in all nine Windows icon sizes, with a PNG frame at 256 pixels.
- Added icon export tooling and validation to catch malformed icon payloads.

## v0.1.0 (Alpha) — 2026-10-19

Initial alpha release.

- Added a Word Home ribbon button and a dedicated ChatGPT Paste tab, alongside the keyboard shortcut.
- Preserved headings, lists, tables, inline formatting, preformatted code, and editable equations while applying destination Word styles and theme fonts.
- Improved Arabic/RTL paragraph and table handling, including punctuation and spacing around numeric labels and equations.
- Added current-user installation, update, and uninstall with cleanup of loaded add-in files after Word exits.
- Pinned Pandoc 3.9.0.2 and verified its download before installation.
