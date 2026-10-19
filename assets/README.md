# Application icon

`chatgpt-paste.png` is the transparent master artwork generated with the built-in image generation tool. The design combines a clipboard with a ChatGPT tile in place of the usual paper sheet.

`chatgpt-paste.ico` contains 16, 20, 24, 32, 40, 48, 64, 128, and 256 pixel Windows icons. The 32 pixel PNG is embedded in the Word ribbon add-in. PNG exports at 16, 48, and 256 pixels are also included for reuse.

To regenerate these exports on Windows, run `powershell -NoProfile -ExecutionPolicy Bypass -File scripts/prepare-icons.ps1`. The checked-in exports mean building the application needs no image-generation service or extra image tools.

This icon identifies an independent utility; it does not imply OpenAI or Microsoft endorsement.

## Generation prompt

Use case: logo-brand. Asset type: production Windows application icon and Microsoft Word ribbon button icon for an independent utility named ChatGPT Paste for Word. Generate ONE polished square icon, no text. Visual concept based on a familiar Paste icon: a warm amber/gold clipboard with a clearly visible top clip at the rear, and in front at the lower right a compact deep teal rounded-square tile containing the recognizable white ChatGPT six-loop interwoven knot symbol, replacing the usual sheet of paper. The clipboard and ChatGPT tile form one compact silhouette. Flat modern desktop icon style, crisp broad clean edges, minimal subtle shading for separation, visually strong simple shapes, legible when reduced to 32 and 16 pixels. Front-facing, no perspective. Centered composition with about 8 percent transparent padding, fill most of the square. Truly transparent background, including outside the shapes. No words, no letters, no paper sheet, no Microsoft Word logo, no photographic texture, no surrounding frame, no drop shadow outside the silhouette, no checkerboard.
