using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Libraries;
using Microsoft.Office.Interop.Word;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題第1回 プロジェクト1。
    /// 1-1 は作業フォルダ内の PDF 存在、1-2〜1-5 は文書の最終状態で判定する。
    /// </summary>
    public class WordChecker1_1_PV1
    {
        private const string PdfFileName = "写真展案内完成.pdf";
        private const string BringHeading = "持参するもの";
        private const string BulletRangeStart = "カメラ";
        private const string BulletRangeEnd = "筆記具";
        private const int TargetListLevel = 2;
        private const float TargetTopMarginMm = 15f;
        private const float TopMarginTolerancePt = 2f;

        private static readonly string[][] ExhibitRows =
        {
            new[] { "作品名", "撮影地", "部門" },
            new[] { "朝の埠頭", "南港", "風景" },
            new[] { "赤いたんぽぽ", "河原", "自然" },
            new[] { "窓の灯", "商店街", "夕景" }
        };

        public bool CheckTask_1_1_01() => RunCheck(CheckTask_1_1_01_Impl);
        public bool CheckTask_1_1_02() => RunCheck(CheckTask_1_1_02_Impl);
        public bool CheckTask_1_1_03() => RunCheck(CheckTask_1_1_03_Impl);
        public bool CheckTask_1_1_04() => RunCheck(CheckTask_1_1_04_Impl);
        public bool CheckTask_1_1_05() => RunCheck(CheckTask_1_1_05_Impl);

        private bool RunCheck(Func<string, bool> task)
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(WordScoreExplanation.UnavailableText);
                return task(filePath);
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
        }

        private bool CheckTask_1_1_01_Impl(string filePath)
        {
            try
            {
                // 教材7-4/7-5と同様: Word と同じフォルダに指定名の出力があるか。開いているかは見ない。
                string dir = Path.GetDirectoryName(filePath);
                if (string.IsNullOrEmpty(dir))
                    return Miss(WordScoreExplanation.UnavailableText);

                string targetPdf = Path.Combine(dir, PdfFileName);
                if (File.Exists(targetPdf))
                    return true;

                return Miss("「写真展案内完成.pdf」が文書と同じフォルダにありません。");
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
        }

        private bool CheckTask_1_1_02_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                Section section = document.Sections[1];
                PageSetup setup = section.PageSetup;
                float topPt = setup.TopMargin;
                Marshal.ReleaseComObject(setup);
                Marshal.ReleaseComObject(section);

                float targetPt = TargetTopMarginMm * 72f / 25.4f;
                if (Math.Abs(topPt - targetPt) <= TopMarginTolerancePt)
                    return true;

                return Miss("文書の上の余白が15mmになっていません。");
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_1_03_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                // 問題文どおり「レベルを1つ下げ」→ 最終状態はリストレベル2（操作ログは使わない）。
                if (!TryGetBringItemsListLevels(document, out var levels, out string locateFail))
                    return Miss(locateFail);

                if (levels.Count == 0)
                    return Miss("「カメラ」から「筆記具」の箇条書きが見つかりません。");

                bool allLevel2 = true;
                for (int i = 0; i < levels.Count; i++)
                {
                    if (levels[i] != TargetListLevel)
                    {
                        allLevel2 = false;
                        break;
                    }
                }

                if (allLevel2)
                    return true;

                bool allSame = true;
                for (int i = 1; i < levels.Count; i++)
                {
                    if (levels[i] != levels[0])
                    {
                        allSame = false;
                        break;
                    }
                }

                if (allSame)
                    return Miss("「カメラ」から「筆記具」の箇条書きがレベル「" + levels[0] + "」になっています。");

                return Miss("「カメラ」から「筆記具」の箇条書きのレベルが揃っていません。");
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        /// <summary>
        /// 「持参するもの」の下、「カメラ」〜「筆記具」の箇条書きの ListLevelNumber を集める。
        /// </summary>
        private static bool TryGetBringItemsListLevels(
            Document document,
            out List<int> levels,
            out string failReason)
        {
            levels = new List<int>();
            failReason = null;

            bool foundHeading = false;
            bool foundEnd = false;
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                try
                {
                    string text = ParagraphText(paragraph);
                    if (!foundHeading)
                    {
                        if (text.IndexOf(BringHeading, StringComparison.Ordinal) >= 0)
                            foundHeading = true;
                        continue;
                    }

                    if (string.IsNullOrEmpty(text))
                        continue;

                    bool isList = false;
                    int level = 0;
                    ListFormat listFormat = null;
                    try
                    {
                        listFormat = paragraph.Range.ListFormat;
                        WdListType listType = listFormat.ListType;
                        isList = listType != WdListType.wdListNoNumbering;
                        if (isList)
                            level = listFormat.ListLevelNumber;
                    }
                    catch
                    {
                        isList = false;
                    }
                    finally
                    {
                        if (listFormat != null)
                            Marshal.ReleaseComObject(listFormat);
                    }

                    if (!isList)
                    {
                        if (levels.Count > 0)
                            break;
                        continue;
                    }

                    bool startsRange = text.IndexOf(BulletRangeStart, StringComparison.Ordinal) >= 0;
                    bool endsRange = text.IndexOf(BulletRangeEnd, StringComparison.Ordinal) >= 0;

                    if (levels.Count == 0 && !startsRange)
                        continue;

                    levels.Add(level);

                    if (endsRange)
                    {
                        foundEnd = true;
                        break;
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(paragraph);
                }
            }

            if (!foundHeading)
            {
                failReason = "「持参するもの」が見つかりません。";
                return false;
            }

            if (levels.Count == 0 || !foundEnd)
            {
                failReason = "「カメラ」から「筆記具」の箇条書きが見つかりません。";
                return false;
            }

            return true;
        }

        private bool CheckTask_1_1_04_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml))
                    return Miss(WordScoreExplanation.UnavailableText);

                string endnotesXml = GetPackagePartXml(xml, "/word/endnotes.xml");
                if (string.IsNullOrEmpty(endnotesXml))
                    return Miss("文末脚注がありません。");

                string endnoteId = FindNoteIdByBody(
                    endnotesXml,
                    @"<w:endnote\b[^>]*w:id=""([^""]+)""[^>]*>.*?</w:endnote>",
                    "人物を主題にした写真のこと");
                if (string.IsNullOrEmpty(endnoteId))
                    return Miss("文末脚注の内容が「人物を主題にした写真のこと」になっていません。");

                string bodyXml = GetPackagePartXml(xml, "/word/document.xml");
                if (string.IsNullOrEmpty(bodyXml))
                    bodyXml = xml;

                if (!HasReferenceImmediatelyAfter(bodyXml, "ポートレート", "endnoteReference", endnoteId))
                    return Miss("「ポートレート」の後ろに文末脚注が挿入されていません。");

                return true;
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_1_05_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                Paragraph heading = FindParagraphByText(document, "出品作");
                if (heading == null)
                    return Miss("見出し「出品作」がありません。");

                int headingEnd = heading.Range.End;
                Marshal.ReleaseComObject(heading);

                Table table = FindFirstTableAfter(document, headingEnd);
                if (table == null)
                    return Miss("見出し「出品作」の下が表になっていません。");

                try
                {
                    if (table.Rows.Count != ExhibitRows.Length || table.Columns.Count != 3)
                        return Miss("見出し「出品作」の下の表の行数または列数が合っていません。");

                    for (int r = 1; r <= ExhibitRows.Length; r++)
                    {
                        for (int c = 1; c <= 3; c++)
                        {
                            string actual = CellText(table.Cell(r, c));
                            if (!string.Equals(actual, ExhibitRows[r - 1][c - 1], StringComparison.Ordinal))
                                return Miss("見出し「出品作」の下の表の内容が合っていません。");
                        }
                    }

                    return true;
                }
                finally
                {
                    Marshal.ReleaseComObject(table);
                }
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static Application GetWordApp()
        {
            try
            {
                return (Application)Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                return null;
            }
        }

        private static Document GetDocument(Application wordApp, string filePath)
        {
            string fileName = Path.GetFileName(filePath);
            foreach (Document doc in wordApp.Documents)
            {
                if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                    doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                {
                    return doc;
                }
            }

            return null;
        }

        private static string GetCurrentWordFilePath()
        {
            Application wordApp = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp?.ActiveDocument != null)
                    return wordApp.ActiveDocument.FullName;
                return null;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                if (wordApp != null)
                    Marshal.ReleaseComObject(wordApp);
            }
        }

        private static Paragraph FindParagraphByText(Document document, string expected)
        {
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                if (ParagraphText(paragraph) == expected)
                    return paragraph;
                Marshal.ReleaseComObject(paragraph);
            }

            return null;
        }

        private static Table FindFirstTableAfter(Document document, int headingEnd)
        {
            Table best = null;
            int bestStart = int.MaxValue;
            foreach (Table table in document.Tables)
            {
                int start = table.Range.Start;
                if (start >= headingEnd && start < bestStart)
                {
                    if (best != null)
                        Marshal.ReleaseComObject(best);
                    best = table;
                    bestStart = start;
                }
                else
                {
                    Marshal.ReleaseComObject(table);
                }
            }

            return best;
        }

        private static string ParagraphText(Paragraph paragraph)
        {
            try
            {
                return (paragraph.Range.Text ?? string.Empty).Trim('\r', '\n', '\a', ' ', '　');
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string CellText(Cell cell)
        {
            try
            {
                return (cell.Range.Text ?? string.Empty).Trim('\r', '\n', '\a', ' ', '　');
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string GetPackagePartXml(string wordOpenXml, string partName)
        {
            string pattern = @"<pkg:part pkg:name=""" + Regex.Escape(partName) + @"""[^>]*>.*?</pkg:part>";
            Match match = Regex.Match(wordOpenXml, pattern, RegexOptions.Singleline);
            return match.Success ? match.Value : string.Empty;
        }

        private static string FindNoteIdByBody(string notesXml, string notePattern, string expectedBody)
        {
            foreach (Match match in Regex.Matches(notesXml, notePattern, RegexOptions.Singleline))
            {
                string tag = match.Value;
                if (tag.IndexOf("w:type=\"separator\"", StringComparison.Ordinal) >= 0 ||
                    tag.IndexOf("w:type=\"continuationSeparator\"", StringComparison.Ordinal) >= 0)
                {
                    continue;
                }

                string text = Regex.Replace(tag, @"<[^>]+>", "");
                if (text.Contains(expectedBody))
                    return match.Groups[1].Value;
            }

            return null;
        }

        private static bool HasReferenceImmediatelyAfter(
            string bodyXml,
            string targetText,
            string referenceLocalName,
            string noteId)
        {
            string replaced = Regex.Replace(
                bodyXml,
                @"<w:" + referenceLocalName + @"\b[^>]*w:id=""([^""]+)""[^>]*/?>",
                "[NOTE_$1]");
            string clean = Regex.Replace(replaced, @"<[^>]+>", "");
            return clean.Contains(targetText + "[NOTE_" + noteId + "]");
        }

        private static bool Miss(string reason)
        {
            WordScoreExplanation.Note(reason);
            return false;
        }
    }
}
