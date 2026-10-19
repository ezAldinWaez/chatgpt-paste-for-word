param([switch]$Quiet)
$ErrorActionPreference = 'Stop'
try {
    Add-Type -AssemblyName System.Windows.Forms
    $clipData = [System.Windows.Forms.Clipboard]::GetDataObject()
    if ($null -eq $clipData) { throw 'Copy the ChatGPT response first.' }
    # CF_HTML is UTF-8 bytes. WinForms GetData(Html) decodes some clipboard
    # producers' bytes as Windows-1252, turning e.g. ’ into â€™.
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class HtmlClipboardBytes {
    [DllImport("user32.dll")] public static extern bool OpenClipboard(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool CloseClipboard();
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern uint RegisterClipboardFormat(string name);
    [DllImport("user32.dll")] public static extern IntPtr GetClipboardData(uint format);
    [DllImport("kernel32.dll")] public static extern IntPtr GlobalLock(IntPtr handle);
    [DllImport("kernel32.dll")] public static extern bool GlobalUnlock(IntPtr handle);
    [DllImport("kernel32.dll")] public static extern UIntPtr GlobalSize(IntPtr handle);
}
'@
    $html = ''
    $htmlFormat = [HtmlClipboardBytes]::RegisterClipboardFormat('HTML Format')
    if (-not [HtmlClipboardBytes]::OpenClipboard([IntPtr]::Zero)) { throw 'Could not open the clipboard. Please try again.' }
    try {
        $handle = [HtmlClipboardBytes]::GetClipboardData($htmlFormat)
        if ($handle -ne [IntPtr]::Zero) {
            $byteCount = [HtmlClipboardBytes]::GlobalSize($handle).ToUInt64()
            if ($byteCount -gt 0 -and $byteCount -lt 100000000) {
                $pointer = [HtmlClipboardBytes]::GlobalLock($handle)
                if ($pointer -eq [IntPtr]::Zero) { throw 'Could not read the HTML clipboard data.' }
                try {
                    $bytes = New-Object byte[] ([int]$byteCount)
                    [Runtime.InteropServices.Marshal]::Copy($pointer, $bytes, 0, [int]$byteCount)
                    $html = (New-Object System.Text.UTF8Encoding($false, $true)).GetString($bytes).TrimEnd([char]0)
                } finally { [void][HtmlClipboardBytes]::GlobalUnlock($handle) }
            }
        }
    } finally { [void][HtmlClipboardBytes]::CloseClipboard() }
    if ($html) {
        $fragmentMatch = [regex]::Match($html, '(?s)<!--StartFragment-->(.*?)<!--EndFragment-->')
        if (-not $fragmentMatch.Success) { throw 'The HTML clipboard does not contain a valid fragment.' }
        $source = $fragmentMatch.Groups[1].Value
        $inputFormat = 'html+tex_math_single_backslash+tex_math_dollars'
    } else {
        $source = [string]$clipData.GetData([System.Windows.Forms.DataFormats]::UnicodeText)
        $inputFormat = 'markdown+tex_math_single_backslash+tex_math_dollars'
    }
    if ([string]::IsNullOrWhiteSpace($source)) { throw 'No text was found. Copy the response first.' }
    # Pandoc's MathML writer rejects TeX equation tags at the end of display
    # math. Keep the number inside the editable equation as visible math text.
    $source = [regex]::Replace($source, '(?s)\\\[(.*?)\\\]', {
        param($match)
        [regex]::Replace($match.Value, '\\tag\*?\{([^{}]+)\}', '\qquad($1)')
    })
    $pandocPath = Join-Path $PSScriptRoot 'pandoc.exe'
    if (-not (Test-Path -LiteralPath $pandocPath)) { throw 'The bundled Pandoc converter is missing. Reinstall ChatGPT Paste for Word.' }
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = $pandocPath
    $processInfo.Arguments = '--from=' + $inputFormat + ' --to=html5 --mathml --ascii --wrap=none --fail-if-warnings'
    $processInfo.UseShellExecute = $false
    $processInfo.CreateNoWindow = $true
    $processInfo.RedirectStandardInput = $true
    $processInfo.RedirectStandardOutput = $true
    $processInfo.RedirectStandardError = $true
    $processInfo.StandardOutputEncoding = New-Object System.Text.UTF8Encoding($false)
    $processInfo.StandardErrorEncoding = New-Object System.Text.UTF8Encoding($false)
    $converter = New-Object System.Diagnostics.Process
    $converter.StartInfo = $processInfo
    [void]$converter.Start()
    $inputBytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($source)
    $converter.StandardInput.BaseStream.Write($inputBytes, 0, $inputBytes.Length)
    $converter.StandardInput.Close()
    $converted = $converter.StandardOutput.ReadToEnd()
    $errors = $converter.StandardError.ReadToEnd()
    $converter.WaitForExit()
    if ($converter.ExitCode -ne 0) { throw ('Conversion failed; your clipboard was left unchanged. ' + $errors) }
    if ($html) {
        # Keep ChatGPT's headings, lists, tables, blockquotes, and other HTML
        # untouched. Use Pandoc only for the MathML nodes, in source order.
        $originalBody = [regex]::Match($source, '(?is)^\s*<html\b[^>]*>\s*(?:<head\b[^>]*>.*?</head>\s*)?<body\b[^>]*>(.*)</body>\s*</html>\s*$')
        $structure = if ($originalBody.Success) { $originalBody.Groups[1].Value } else { $source }
        # Word discards HTML <bdi> isolation around short caption numbers,
        # which can attach punctuation to adjacent Arabic text. LRM/RLM are
        # invisible in Word and keep a numeric label such as "2." together.
        $structure = [regex]::Replace($structure, '(?is)<bdi\b[^>]*>([^<>]*)</bdi>', {
            param($match)
            $label = $match.Groups[1].Value
            if ($label -match '^\s*[0-9٠-٩۰-۹]+(?:[.,٫][0-9٠-٩۰-۹]+)*[.):]?\s*$') {
                return '&#x200E;' + $label + '&#x200F;'
            }
            return $match.Value
        })
        # Word treats an ordinary space at the start of the next HTML span as
        # optional. Keep the visible gap after RTL numeric caption labels.
        $structure = [regex]::Replace($structure, '(?is)(&#x200F;\s*</strong>\s*<span\b[^>]*>)[ \t]+', '$1&#160;')
        # Numeric section prefixes need both a stable LTR number and a real
        # gap before Arabic. Word otherwise joins e.g. "3.1" to the title.
        $structure = [regex]::Replace($structure, '(?is)(<h[1-6]\b[^>]*>(?:\s*<span\b[^>]*>)?)([0-9]+(?:\.[0-9]+)*\.?)[ \t]+(?=(?:&#x0*6|[\u0600-\u06ff]))', '$1&#x200E;$2&#x200F;&#160;')
        # Word can glue a parenthesized equation reference to the next Arabic
        # word even when the HTML has a normal space. Work only on text nodes;
        # math annotations, code, and HTML attributes remain untouched.
        $structure = [regex]::Replace($structure, '(?s)(?<=>)[^<]+(?=<)', {
            param($node)
            $value = $node.Value
            if ($value -notmatch '(?i)(?:&#x0*6[0-9a-f]{2};|[\u0600-\u06ff])') { return $value }
            return [regex]::Replace($value, '(?i)\(([0-9٠-٩]+)\)[ \t]+(?=(?:&#x0*6[0-9a-f]{2};|[\u0600-\u06ff]))', {
                param($reference)
                return '&#x200E;(' + $reference.Groups[1].Value + ')&#x200F;&#160;'
            })
        })
        # Word's HTML importer may put neutral punctuation in a separate LTR
        # run even when it encloses Arabic, e.g. the brackets in "الزمن (ساعة)".
        # Mark only simple Arabic-only spans RTL. Mixed spans with numbers or
        # Latin scientific notation must keep their normal LTR numeric runs.
        $structure = [regex]::Replace($structure, '(?is)<pre\b.*?</pre>|<code\b.*?</code>|<span(?<attrs>[^>]*)>(?<body>[^<>]*)</span>', {
            param($node)
            if (-not $node.Value.StartsWith('<span', [StringComparison]::OrdinalIgnoreCase)) { return $node.Value }
            $attributes = $node.Groups['attrs'].Value
            if ($attributes -match '(?i)\bdir\s*=') { return $node.Value }
            $body = $node.Groups['body'].Value
            $plain = [System.Net.WebUtility]::HtmlDecode($body)
            if ($plain -match '[\u0600-\u08ff\ufb50-\ufdff\ufe70-\ufeff]' -and
                $plain -notmatch '(?i)[a-z0-9\u0660-\u0669\u06f0-\u06f9]') {
                return '<span' + $attributes + ' dir="rtl">' + $body + '</span>'
            }
            return $node.Value
        })
        $mathNodes = [regex]::Matches($converted, '(?s)<math\b[^>]*>.*?</math>')
        $texPattern = '(?s)\\\[(.*?)\\\]|\\\((.*?)\\\)|\$\$(.*?)\$\$|\$(?!\$)(.*?)\$'
        $texNodes = [regex]::Matches($structure, $texPattern)
        $sourceMathCount = [regex]::Matches($structure, '<math(?:\s|>)').Count
        if ($sourceMathCount -gt 0 -and $texNodes.Count -eq 0 -and $sourceMathCount -eq $mathNodes.Count) {
            # A copied ChatGPT answer may already contain MathML. Retain its
            # HTML directly so RTL labels, styles, and tables survive.
            $converted = $structure
        } elseif ($mathNodes.Count -gt 0 -and $mathNodes.Count -eq $texNodes.Count) {
            $script:mathIndex = 0
            $converted = [regex]::Replace($structure, $texPattern, {
                param($match)
                $node = $mathNodes[$script:mathIndex].Value
                $script:mathIndex++
                return $node
            })
        }
    }
    # Word's HTML math importer can consume ordinary spaces after inline math.
    $converted = [regex]::Replace($converted, '(?s)(<math display="inline".*?</math></span>\s*<span>)\s+', '$1&#160;')
    $converted = [regex]::Replace($converted, '(?s)(<math display="inline".*?</math>(?:</span>)*)\s+(?=[^<\s])', '$1&#160;')
    # WinForms can down-convert CF_HTML text. Numeric entities keep the bytes
    # ASCII-only while Word renders the original Unicode characters.
    $converted = [regex]::Replace($converted, '[\uD800-\uDBFF][\uDC00-\uDFFF]|[^\x00-\x7F]', {
        param($match)
        $codePoint = if ($match.Value.Length -eq 2) { [char]::ConvertToUtf32($match.Value, 0) } else { [int][char]$match.Value[0] }
        '&#x{0:X};' -f $codePoint
    })
    $mathCount = [regex]::Matches($converted, '<math(?:\s|>)').Count
    if ($mathCount -eq 0) { exit 2 }
    $prefix = '<html><head><meta charset="utf-8"></head><body><!--StartFragment-->'
    $suffix = '<!--EndFragment--></body></html>'
    $headerTemplate = "Version:0.9`r`nStartHTML:{0:D10}`r`nEndHTML:{1:D10}`r`nStartFragment:{2:D10}`r`nEndFragment:{3:D10}`r`n"
    $encoding = [System.Text.Encoding]::UTF8
    $startHtml = $encoding.GetByteCount(($headerTemplate -f 0,0,0,0))
    $startFragment = $startHtml + $encoding.GetByteCount($prefix)
    $endFragment = $startFragment + $encoding.GetByteCount($converted)
    $endHtml = $endFragment + $encoding.GetByteCount($suffix)
    $clipboardHtml = ($headerTemplate -f $startHtml,$endHtml,$startFragment,$endFragment) + $prefix + $converted + $suffix
    $newClipboard = New-Object System.Windows.Forms.DataObject
    $newClipboard.SetData([System.Windows.Forms.DataFormats]::Html, $clipboardHtml)
    $originalText = [string]$clipData.GetData([System.Windows.Forms.DataFormats]::UnicodeText)
    if ($originalText) { $newClipboard.SetData([System.Windows.Forms.DataFormats]::UnicodeText, $originalText) }
    [System.Windows.Forms.Clipboard]::SetDataObject($newClipboard, $true)
    Write-Output ('Ready to paste into Word: ' + $mathCount + ' equations converted.')
} catch {
    if ($Quiet) {
        [Console]::Error.WriteLine($_.Exception.Message)
        exit 1
    }
    Add-Type -AssemblyName System.Windows.Forms
    [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Prepare Word Paste')
    exit 1
}
