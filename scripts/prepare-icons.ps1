# Convert the generated master PNG into native Windows and Office icon assets.
# Uses only Windows/.NET; run this when assets/chatgpt-paste.png changes.
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing
$root = Split-Path -Parent $PSScriptRoot
$assets = Join-Path $root 'assets'
$source = [Drawing.Image]::FromFile((Join-Path $assets 'chatgpt-paste.png'))
$sizes = @(16, 20, 24, 32, 40, 48, 64, 128, 256)
$entries = @()
try {
    foreach ($size in $sizes) {
        $bitmap = [Drawing.Bitmap]::new($size, $size, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.CompositingMode = [Drawing.Drawing2D.CompositingMode]::SourceCopy
            $graphics.InterpolationMode = [Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.PixelOffsetMode = [Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $graphics.DrawImage($source, [Drawing.Rectangle]::new(0, 0, $size, $size))
            if ($size -in @(16, 32, 48, 256)) {
                $bitmap.Save((Join-Path $assets "chatgpt-paste-$size.png"), [Drawing.Imaging.ImageFormat]::Png)
            }
            # Windows uses a PNG-compressed frame for the 256 pixel icon.
            if ($size -eq 256) {
                $largeFrame = [IO.MemoryStream]::new()
                try {
                    $bitmap.Save($largeFrame, [Drawing.Imaging.ImageFormat]::Png)
                    $entries += @{ Size = $size; Data = $largeFrame.ToArray() }
                }
                finally { $largeFrame.Dispose() }
                continue
            }
            # Write a 32-bit DIB directly: Icon.FromHandle().Save() quantizes
            # generated icons to 16 colors on .NET Framework.
            $rectangle = [Drawing.Rectangle]::new(0, 0, $size, $size)
            $locked = $bitmap.LockBits($rectangle, [Drawing.Imaging.ImageLockMode]::ReadOnly, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
            $pixels = [byte[]]::new($size * $size * 4)
            try {
                for ($row = 0; $row -lt $size; $row++) {
                    $rowPointer = [IntPtr]::Add($locked.Scan0, $row * $locked.Stride)
                    [Runtime.InteropServices.Marshal]::Copy($rowPointer, $pixels, ($size - 1 - $row) * $size * 4, $size * 4)
                }
            }
            finally { $bitmap.UnlockBits($locked) }
            $maskStride = [int]([Math]::Ceiling($size / 32.0) * 4)
            $mask = [byte[]]::new($maskStride * $size)
            for ($row = 0; $row -lt $size; $row++) {
                for ($column = 0; $column -lt $size; $column++) {
                    if ($pixels[($row * $size + $column) * 4 + 3] -eq 0) {
                        $maskIndex = $row * $maskStride + [int][Math]::Floor($column / 8.0)
                        $mask[$maskIndex] = $mask[$maskIndex] -bor (0x80 -shr ($column % 8))
                    }
                }
            }
            $stream = [IO.MemoryStream]::new()
            $dib = [IO.BinaryWriter]::new($stream)
            try {
                $dib.Write([uint32]40) # BITMAPINFOHEADER size
                $dib.Write([int]$size)
                $dib.Write([int]($size * 2)) # XOR bitmap plus AND mask
                $dib.Write([uint16]1)
                $dib.Write([uint16]32)
                $dib.Write([uint32]0) # BI_RGB
                $dib.Write([uint32]($pixels.Length + $mask.Length))
                foreach ($unused in 1..4) { $dib.Write([uint32]0) }
                $dib.Write($pixels)
                $dib.Write($mask)
                $entries += @{ Size = $size; Data = $stream.ToArray() }
            }
            finally { $dib.Dispose() }
        }
        finally { $graphics.Dispose(); $bitmap.Dispose() }
    }
    $output = [IO.File]::Create((Join-Path $assets 'chatgpt-paste.ico'))
    $writer = [IO.BinaryWriter]::new($output)
    try {
        $writer.Write([uint16]0)
        $writer.Write([uint16]1)
        $writer.Write([uint16]$entries.Count)
        $offset = 6 + 16 * $entries.Count
        foreach ($entry in $entries) {
            $dimension = if ($entry.Size -eq 256) { 0 } else { $entry.Size }
            $writer.Write([byte]$dimension)
            $writer.Write([byte]$dimension)
            $writer.Write([byte]0)
            $writer.Write([byte]0)
            $writer.Write([uint16]1)
            $writer.Write([uint16]32)
            $writer.Write([uint32]$entry.Data.Length)
            $writer.Write([uint32]$offset)
            $offset += $entry.Data.Length
        }
        foreach ($entry in $entries) { $writer.Write([byte[]]$entry.Data) }
    }
    finally { $writer.Dispose() }
}
finally { $source.Dispose() }
Write-Host 'Prepared transparent PNGs and the multi-resolution Windows icon.'
