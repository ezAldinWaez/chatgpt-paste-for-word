# ChatGPT Paste for Word

**v0.1.1 · Alpha** — Paste ChatGPT responses into desktop Word with editable equations, destination styles, and Arabic/RTL support. Independent of OpenAI and Microsoft.

**[Download for Windows](https://github.com/ezAldinWaez/chatgpt-paste-for-word/releases/download/v0.1.1/ChatGPT-Paste-for-Word-Setup-0.1.1.exe)** · [Releases](https://github.com/ezAldinWaez/chatgpt-paste-for-word/releases) · [Changelog](CHANGELOG.md)

## Install and use

1. Run the installer and choose **Install** or **Update**. No administrator access needed.
2. Close every Word window, then reopen Word.
3. Copy a ChatGPT response, place the cursor in Word, and click **ChatGPT Paste** on the Home ribbon or ChatGPT Paste tab. Shortcut: **Ctrl+Alt+Shift+V**.

Requires **64-bit Windows** and **desktop Microsoft Word**; tested with Word 2021. First install needs internet to download verified [Pandoc 3.9.0.2](https://github.com/jgm/pandoc/releases/tag/3.9.0.2) (~41 MB). The installer is unsigned, so Windows may show an unknown-publisher warning.

## What it preserves

- Editable equations, headings, lists, tables, emphasis, and code.
- Destination Word styles and theme fonts.
- Cursor reading direction, including Arabic paragraphs and RTL table order.

Conversion runs locally; copied text is not uploaded. Word controls equation appearance, so spacing may differ from ChatGPT.

## Help and development

For formatting details or an inactive button, see [usage and troubleshooting](docs/usage.md). To remove the tool, run setup → **Uninstall**, or use Windows **Installed apps**.

[Contributing](CONTRIBUTING.md) · [Architecture](docs/architecture.md) · [Testing](docs/testing.md) · [Releasing](docs/releasing.md) · [MIT license](LICENSE)
