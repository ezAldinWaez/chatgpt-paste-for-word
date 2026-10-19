# Architecture

`src/Setup.cs` is a current-user Windows Forms installer. Its executable embeds the tray utility, Ribbon DLL, converter script, and uninstall scripts as resources. It downloads the pinned Pandoc Windows ZIP from the upstream release, checks SHA-256, and installs a verified `pandoc.exe` under `%LOCALAPPDATA%\ChatGPTPasteForWord`. It registers the managed COM add-in for Word in the current user's 32- and 64-bit registry views and adds the tray utility to startup. It does not require administrator rights.

`src/ChatGPTPasteRibbon.cs` exposes the Ribbon button and a bridge used by the tray utility. `src/ChatGPTPasteHotkey.cs` owns the keyboard shortcut and forwards paste requests to the active Word add-in. `src/Prepare-Word-Paste.ps1` reads the current clipboard's HTML when present, or Markdown/plain text otherwise. It converts recognized TeX equations to MathML through Pandoc, keeps Word-compatible HTML structure, and places CF_HTML back on the clipboard for Word to paste. `src/WordPasteFormatting.cs` maps imported text to the destination document's styles and reading direction.

The Ribbon DLL embeds Office and Extensibility COM types at build time. Build machines need the relevant primary interop assemblies; user machines need desktop Word, not separate interop assemblies. The source is compiled with the .NET Framework C# compiler because the add-in is loaded by Windows desktop Office.

The generated master icon and its Windows/Office exports live under `assets/`. `scripts/prepare-icons.ps1` regenerates the exports using Windows/.NET. The build embeds the multi-resolution ICO in both executables and the transparent 32 pixel PNG in the Ribbon DLL. Both ribbon buttons use the same cached OLE picture through their `getImage` callback. Setup registers the installed executable as the Windows Installed Apps icon.

The converter temporarily changes the clipboard during a paste request. When conversion fails, it aims to leave the original clipboard unchanged. No pasted content is sent to a service. Pandoc is downloaded only during setup when a verified local copy is unavailable.

The installer is unsigned. Maintainers can add Authenticode signing in a future release without changing the source layout.
