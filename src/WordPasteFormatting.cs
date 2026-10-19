using System;
using System.Collections.Generic;

internal sealed class WordPasteContext
{
    internal int Start;
    internal int SelectedLength;
    internal int DocumentEnd;
    internal int ReadingOrder;
    internal int Alignment;
}

internal static class WordPasteFormatting
{
    private sealed class Exclusion
    {
        internal int Start;
        internal int End;
    }

    internal static WordPasteContext Capture(dynamic selection, dynamic document)
    {
        dynamic range = selection.Range;
        int order = Convert.ToInt32(selection.ParagraphFormat.ReadingOrder);
        if (order != 0 && order != 1) order = Convert.ToInt32(range.Paragraphs[1].ReadingOrder);
        if (order != 0 && order != 1) order = 1;
        return new WordPasteContext
        {
            Start = Convert.ToInt32(range.Start),
            SelectedLength = Convert.ToInt32(range.End) - Convert.ToInt32(range.Start),
            DocumentEnd = Convert.ToInt32(document.Content.End),
            ReadingOrder = order, // 0 = RTL; 1 = LTR
            Alignment = Convert.ToInt32(selection.ParagraphFormat.Alignment)
        };
    }

    internal static int Normalize(dynamic document, dynamic selection, WordPasteContext context)
    {
        dynamic after = selection.Range;
        int end = Math.Max(Convert.ToInt32(after.Start), Convert.ToInt32(after.End));
        int measuredEnd = context.Start + Convert.ToInt32(document.Content.End) - context.DocumentEnd + context.SelectedLength;
        end = Math.Max(end, measuredEnd);
        end = Math.Min(end, Convert.ToInt32(document.Content.End));
        if (end <= context.Start) return 0;

        dynamic normalStyle = document.Styles[-1]; // wdStyleNormal
        string normalName = (string)normalStyle.NameLocal;
        string webName = (string)document.Styles[-95].NameLocal; // wdStyleHtmlNormal
        string preName = (string)document.Styles[-102].NameLocal; // wdStyleHtmlPre
        dynamic codeStyle = document.Styles[-99]; // wdStyleHtmlCode
        dynamic normalFont = normalStyle.Font;
        dynamic pasted = document.Range(context.Start, end);
        int changed = 0;

        // Changing a paragraph style can invalidate Word's COM enumerator,
        // especially when called from the separate shortcut process.
        int paragraphCount = Convert.ToInt32(pasted.Paragraphs.Count);
        for (int paragraphIndex = 1; paragraphIndex <= paragraphCount; paragraphIndex++)
        {
            dynamic paragraph = pasted.Paragraphs[paragraphIndex];
            string styleName = (string)paragraph.Style.NameLocal;
            if (String.Equals(styleName, webName, StringComparison.OrdinalIgnoreCase))
            {
                paragraph.Style = normalStyle;
                styleName = normalName;
            }

            paragraph.ReadingOrder = context.ReadingOrder;
            int currentAlignment = Convert.ToInt32(paragraph.Alignment);
            if ((currentAlignment == 0 || currentAlignment == 2) &&
                (context.Alignment == 0 || context.Alignment == 2))
                paragraph.Alignment = context.Alignment;

            // Use the destination paragraph style for every pasted paragraph.
            // A heading keeps its heading font, while body and table text get
            // their own destination style's theme font rather than source HTML.
            if (String.Equals(styleName, preName, StringComparison.OrdinalIgnoreCase)) continue;
            dynamic styleFont = paragraph.Style.Font;
            string ascii = FontOrFallback((string)styleFont.NameAscii, (string)normalFont.NameAscii);
            string other = FontOrFallback((string)styleFont.NameOther, (string)normalFont.NameOther);
            string bi = FontOrFallback((string)styleFont.NameBi, (string)normalFont.NameBi);
            string farEast = FontOrFallback((string)styleFont.NameFarEast, (string)normalFont.NameFarEast);

            int from = Math.Max(context.Start, Convert.ToInt32(paragraph.Range.Start));
            int to = Math.Min(end, Convert.ToInt32(paragraph.Range.End));
            if (to <= from) continue;
            // Keep editable OMath and inline code in their original fonts.
            var exclusions = new List<Exclusion>();
            foreach (dynamic equation in paragraph.Range.OMaths)
            {
                int mathStart = Convert.ToInt32(equation.Range.Start);
                int mathEnd = Convert.ToInt32(equation.Range.End);
                exclusions.Add(new Exclusion { Start = mathStart, End = mathEnd });
            }
            int searchStart = from;
            while (searchStart < to)
            {
                dynamic codeSearch = document.Range(searchStart, to);
                dynamic find = codeSearch.Find;
                find.ClearFormatting();
                find.Style = codeStyle;
                find.Format = true;
                find.Text = "";
                find.Forward = true;
                find.Wrap = 0; // wdFindStop
                if (!Convert.ToBoolean(find.Execute())) break;
                int codeStart = Convert.ToInt32(codeSearch.Start);
                int codeEnd = Convert.ToInt32(codeSearch.End);
                if (codeEnd <= searchStart) break;
                exclusions.Add(new Exclusion { Start = codeStart, End = codeEnd });
                searchStart = codeEnd;
            }
            exclusions.Sort(delegate(Exclusion a, Exclusion b) { return a.Start.CompareTo(b.Start); });
            foreach (Exclusion excluded in exclusions)
            {
                if (excluded.End <= from) continue;
                if (excluded.Start > from) ApplyTheme(document, from, Math.Min(excluded.Start, to), ascii, other, bi, farEast);
                from = Math.Max(from, excluded.End);
                if (from >= to) break;
            }
            if (from < to) ApplyTheme(document, from, to, ascii, other, bi, farEast);
            changed++;
        }
        // Paragraph reading order does not set the visual order of table
        // columns. Keep the first source column on the cursor's leading side.
        int tableCount = Convert.ToInt32(pasted.Tables.Count);
        for (int tableIndex = 1; tableIndex <= tableCount; tableIndex++)
        {
            dynamic table = pasted.Tables[tableIndex];
            int tableStart = Convert.ToInt32(table.Range.Start);
            int tableEnd = Convert.ToInt32(table.Range.End);
            if (tableStart >= context.Start && tableEnd <= end)
                table.TableDirection = context.ReadingOrder; // 0 = RTL; 1 = LTR
        }
        return changed;
    }

    private static string FontOrFallback(string value, string fallback)
    {
        return String.IsNullOrWhiteSpace(value) ? fallback : value;
    }

    private static void ApplyTheme(dynamic document, int start, int end, string ascii, string other, string bi, string farEast)
    {
        // Word can replace a cross-process COM Range after a font property is
        // changed. Reacquire it for each theme slot using stable positions.
        if (!String.IsNullOrWhiteSpace(ascii)) document.Range(start, end).Font.NameAscii = ascii;
        if (!String.IsNullOrWhiteSpace(other)) document.Range(start, end).Font.NameOther = other;
        if (!String.IsNullOrWhiteSpace(bi)) document.Range(start, end).Font.NameBi = bi;
        if (!String.IsNullOrWhiteSpace(farEast)) document.Range(start, end).Font.NameFarEast = farEast;
    }
}
