using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace Libraries
{
    /// <summary>
    /// Word Find 実行時に Selection を動かさないよう、Content.Duplicate と選択復元を共通化する。
    /// </summary>
    public static class WordFindHelper
    {
        public static Range DuplicateContent(Document document)
        {
            if (document?.Content == null)
                return null;
            return document.Content.Duplicate;
        }

        public static void ConfigureSafeFind(Find find, string text)
        {
            if (find == null)
                return;
            find.ClearFormatting();
            find.Text = text ?? "";
            find.Forward = true;
            find.Wrap = WdFindWrap.wdFindStop;
            find.Format = false;
            find.Replacement.Text = "";
        }

        public static void PreserveSelection(Application app, Action action)
        {
            if (app == null)
            {
                action();
                return;
            }

            int selStart = -1;
            int selEnd = -1;
            try
            {
                Selection sel = app.Selection;
                selStart = sel.Start;
                selEnd = sel.End;
            }
            catch { }

            try
            {
                action();
            }
            finally
            {
                RestoreSelectionIfNeeded(app, selStart, selEnd);
            }
        }

        public static T PreserveSelection<T>(Application app, Func<T> func)
        {
            if (app == null)
                return func();

            int selStart = -1;
            int selEnd = -1;
            try
            {
                Selection sel = app.Selection;
                selStart = sel.Start;
                selEnd = sel.End;
            }
            catch { }

            try
            {
                return func();
            }
            finally
            {
                RestoreSelectionIfNeeded(app, selStart, selEnd);
            }
        }

        private static void RestoreSelectionIfNeeded(Application app, int selStart, int selEnd)
        {
            if (selStart < 0 || selEnd < 0)
                return;
            try
            {
                Selection after = app.Selection;
                if (after.Start != selStart || after.End != selEnd)
                    after.SetRange(selStart, selEnd);
            }
            catch { }
        }

        /// <summary>最初の一致 Range を返す。見つからなければ null（呼び出し元で ReleaseComObject）。</summary>
        public static Range FindFirstRange(Document document, Application app, string searchText)
        {
            return PreserveSelection(app, () => FindFirstRangeCore(document, searchText));
        }

        private static Range FindFirstRangeCore(Document document, string searchText)
        {
            Range searchRange = null;
            try
            {
                searchRange = DuplicateContent(document);
                if (searchRange == null)
                    return null;

                Find find = searchRange.Find;
                ConfigureSafeFind(find, searchText);
                if (!find.Execute())
                {
                    Marshal.ReleaseComObject(searchRange);
                    return null;
                }

                return searchRange;
            }
            catch
            {
                if (searchRange != null)
                {
                    try { Marshal.ReleaseComObject(searchRange); } catch { }
                }
                return null;
            }
        }

        public static int CountTextOccurrences(Document document, Application app, string searchText)
        {
            return PreserveSelection(app, () => CountTextOccurrencesCore(document, searchText));
        }

        private static int CountTextOccurrencesCore(Document document, string searchText)
        {
            Range searchRange = null;
            int count = 0;
            try
            {
                searchRange = DuplicateContent(document);
                if (searchRange == null)
                    return 0;

                Find find = searchRange.Find;
                ConfigureSafeFind(find, searchText);
                while (find.Execute())
                {
                    count++;
                    searchRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                }
            }
            catch { }
            finally
            {
                if (searchRange != null)
                {
                    try { Marshal.ReleaseComObject(searchRange); } catch { }
                }
            }
            return count;
        }

        /// <summary>document.xml 内の w:footnoteReference 数（8-3 ゲート・チェッカー用）。</summary>
        public static int CountFootnoteReferencesInXml(string wordOpenXml)
        {
            if (string.IsNullOrEmpty(wordOpenXml))
                return 0;

            var docPartMatch = Regex.Match(
                wordOpenXml,
                @"<pkg:part pkg:name=""/word/document\.xml""[^>]*>.*?</pkg:part>",
                RegexOptions.Singleline);
            string bodyXml = docPartMatch.Success ? docPartMatch.Value : wordOpenXml;
            return Regex.Matches(bodyXml, @"<w:footnoteReference\b").Count;
        }

        /// <summary>footnotes.xml 内の実脚注数（separator / continuationSeparator を除く）。</summary>
        public static int CountFootnoteDefinitionsInXml(string wordOpenXml)
        {
            if (string.IsNullOrEmpty(wordOpenXml))
                return 0;

            var footnotesPartMatch = Regex.Match(
                wordOpenXml,
                @"<pkg:part pkg:name=""/word/footnotes\.xml""[^>]*>.*?</pkg:part>",
                RegexOptions.Singleline);
            if (!footnotesPartMatch.Success)
                return 0;

            string footnotesXml = footnotesPartMatch.Value;
            int count = 0;
            var matches = Regex.Matches(footnotesXml, @"<w:footnote\b[^>]*>", RegexOptions.Singleline);
            foreach (Match m in matches)
            {
                string tag = m.Value;
                if (tag.IndexOf("w:type=\"separator\"", StringComparison.Ordinal) >= 0
                    || tag.IndexOf("w:type=\"continuationSeparator\"", StringComparison.Ordinal) >= 0
                    || tag.IndexOf("w:type=\"continuationNotice\"", StringComparison.Ordinal) >= 0)
                    continue;
                count++;
            }
            return count;
        }
    }
}
