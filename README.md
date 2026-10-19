# ChatGPT Paste for Word

An independent Windows utility for copying formatted ChatGPT responses with equations into desktop Microsoft Word. It adds a **ChatGPT Paste** button to Word's **Home** ribbon and a dedicated **ChatGPT Paste** tab. The **Ctrl+Alt+Shift+V** shortcut also remains available. It is not an official OpenAI or Microsoft product.

## Install

1. Double-click `ChatGPT-Paste-for-Word-Setup-1.5.9.exe`.
2. Click **Install**. No administrator access is needed.
3. On a first install, setup downloads the official Pandoc 3.9.0.2 Windows release (~41 MB), checks its SHA-256 hash, and installs to your user profile. Updates reuse a verified local copy.

Requires 64-bit Windows, desktop Microsoft Word, and an internet connection during installation. Tested with Word 2021 on Windows. The installer is unsigned, so Windows may show an unknown-publisher warning.

If an earlier version is installed, click **Update** to replace it. **Close every Word window and reopen Word** after installation so the button loads the new version. Word can remain running in the background after its visible windows close; if the old button persists, save your documents and check for `WINWORD.EXE` in Task Manager before reopening Word. The shortcut now warns when the current Word window still has an older add-in loaded. The same installer offers **Uninstall**. Close Word first for immediate file removal; if Word is open, uninstall disables the add-in and startup shortcut immediately, then finishes deleting in-use files after Word closes.

## Use

Copy a ChatGPT response with the **Copy** button or by selecting text. Click where it should go in Word, then click **ChatGPT Paste** on the Home ribbon (or on the ChatGPT Paste tab). The keyboard shortcut **Ctrl+Alt+Shift+V** does the same thing. The utility keeps ChatGPT's HTML structure, converts recognized LaTeX equations to MathML, preserves existing MathML, and requests an HTML paste into Word. It maps imported body paragraphs to the destination document's **Normal** style and applies each pasted paragraph's destination style theme font instead of the font copied from ChatGPT. It retains heading styles, lists, tables, preformatted and inline code, inline emphasis, and editable equations. Inline code retains its monospace font. Pasted paragraphs and table columns follow the reading direction at the Word cursor; in RTL context, the first source table column appears on the right. Numeric labels isolated by ChatGPT's HTML, such as "الجدول 2.", keep their punctuation and spacing in Arabic captions; numbered Arabic headings and parenthesized equation references get a stable gap before the next Arabic word. Decimal numbers and Latin scientific units retain their left-to-right character order. Left or right alignment follows the cursor context, while centered and justified paragraphs remain as they were. If it finds no convertible equations, it leaves the clipboard as copied. When the copy has no HTML, it converts the Markdown text instead. Conversion runs locally; your copied text is not uploaded.

**Show/Hide ¶** can help inspect spacing, but its visible marks are separate from the underlying direction of Word text runs. For Arabic-only HTML spans, this version explicitly imports neutral punctuation such as parentheses with the surrounding Arabic run. Decimal values and Latin units remain separate left-to-right runs. If you want to hide the nonprinting marks, toggle **Home → ¶** or press **Ctrl+Shift+8**; that display setting does not change reading direction.

In testing, the Word ribbon button pasted an article with 25 editable equations, ten Word headings, nine list items, two tables, and preformatted code; no body paragraph kept Word's Normal (Web) style, and saved Normal text had no direct Times New Roman runs. Word may still display a resolved font name such as Times New Roman when a document's theme assigns that font to Arabic headings; those headings continue to use the destination heading theme. Numbered equations using LaTeX `\tag{...}` retain visible numbers inside editable equations. Results may vary with different ChatGPT content, copy formats, and Word versions.

If the button seems inactive, first close and reopen **all** Word windows. The ribbon add-in writes `ribbon.log` in `%LOCALAPPDATA%\ChatGPTPasteForWord` on each click; that file records a click and any conversion error without recording clipboard text.

Word renders editable equations with its own math font and spacing. The mathematical symbols are preserved, but the equation can look different from ChatGPT's display. A pixel-identical copy would require an image instead of editable Word math. See [Microsoft's MathML import documentation](https://learn.microsoft.com/en-us/office/math/mathml).

## Remove

Run the same setup file and click **Uninstall**, or open **Windows Settings → Apps → Installed apps**, find **ChatGPT Paste for Word (unofficial)**, and choose **Uninstall**. This also removes the startup entry. If Word still has an add-in file open, the uninstaller schedules that file for automatic deletion once Word exits. You can exit the running utility from its tray icon.

## Converter attribution

The installer downloads [Pandoc 3.9.0.2](https://github.com/jgm/pandoc/releases/tag/3.9.0.2) directly from the official release page. The expected SHA-256 of `pandoc-3.9.0.2-windows-x86_64.zip` is `C97542F2800F446E788D9F74237856D995421AD1BB3CC8324286840C5F272D3A`. Pandoc's copyright and license files are installed alongside its executable.

## Source and releases

The maintained source is under `src/`. Build and packaging scripts are under `scripts/`; see [CONTRIBUTING.md](CONTRIBUTING.md) for development requirements and [docs/releasing.md](docs/releasing.md) for the release process. The source is licensed under [MIT](LICENSE). Release installers are separate build artifacts rather than files committed to Git. This project is independent of OpenAI and Microsoft.
