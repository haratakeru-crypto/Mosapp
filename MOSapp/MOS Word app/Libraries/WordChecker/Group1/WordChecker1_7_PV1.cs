using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Libraries;
using Microsoft.Office.Interop.Word;

namespace Libraries.Group1
{
    /// <summary>
    /// 類題第1回 プロジェクト7。7-1〜7-4 は最終状態のみで判定する。
    /// </summary>
    public class WordChecker1_7_PV1
    {
        private const string RevisionAcceptPhrase = "および車椅子";
        private const string RevisionRejectPhrase = "※先着20名";
        private const string InspectionDateLinePhrase = "点検日：20XX年9月12日";
        private const string TargetContentStatus = "点検中";
        private const string SmartArtAnchorLabel = "点検";
        private const string SmartArtSiblingLabel = "補充";
        private const string SmartArtChildLabel = "記録";
        private const int FontHalfPointsTopLevel = 3100;
        private const int FontHalfPointsChild = 1400;
        private const int FontHalfPointsTolerance = 100;

        public bool CheckTask_1_7_01() => RunCheck(CheckTask_1_7_01_Impl);
        public bool CheckTask_1_7_02() => RunCheck(CheckTask_1_7_02_Impl);
        public bool CheckTask_1_7_03() => RunCheck(CheckTask_1_7_03_Impl);
        public bool CheckTask_1_7_04() => RunCheck(CheckTask_1_7_04_Impl);

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

        private bool CheckTask_1_7_01_Impl(string filePath)
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

                string plain = GetDocumentPlainText(document);
                if (plain.IndexOf(RevisionAcceptPhrase, StringComparison.Ordinal) < 0)
                    return Miss("「および車椅子」が本文に反映されていません。");

                if (plain.IndexOf(RevisionRejectPhrase, StringComparison.Ordinal) >= 0)
                    return Miss("「※先着20名」が元に戻されていません。");

                if (HasPendingRevisionForPhrase(document, RevisionAcceptPhrase)
                    || HasPendingRevisionForPhrase(document, RevisionRejectPhrase))
                {
                    return Miss("指定の変更履歴が未処理のままです。");
                }

                string xml;
                try { xml = document.WordOpenXML ?? string.Empty; }
                catch { xml = string.Empty; }

                if (XmlHasTrackedChangeMarkupForPhrase(xml, RevisionAcceptPhrase)
                    || XmlHasTrackedChangeMarkupForPhrase(xml, RevisionRejectPhrase))
                {
                    return Miss("指定の変更履歴が確定していません。");
                }

                return true;
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_7_02_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            Paragraph paragraph = null;
            try
            {
                wordApp = GetWordApp();
                if (wordApp == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                paragraph = FindParagraphContaining(document, InspectionDateLinePhrase);
                if (paragraph == null)
                    return Miss("「点検日：20XX年9月12日」の行が見つかりません。");

                int paraStart = paragraph.Range.Start;
                int paraEnd = paragraph.Range.End;
                if (!IsWrapTightInParagraph(document, paraStart, paraEnd))
                    return Miss("点検日行付近の図形の文字列折り返しが［狭く］になっていません。");

                return true;
            }
            finally
            {
                if (paragraph != null)
                    Marshal.ReleaseComObject(paragraph);
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_7_03_Impl(string filePath)
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

                if (!TryGetContentStatus(document, out string status) || string.IsNullOrEmpty(status))
                    return Miss("文書のプロパティ［状態］が設定されていません。");

                if (!string.Equals(status.Trim(), TargetContentStatus, StringComparison.Ordinal))
                    return Miss("文書のプロパティ［状態］が「点検中」になっていません。");

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

        /// <summary>
        /// ［状態］= Content status。OpenXML (cp:contentStatus) を優先し、取れなければ COM。
        /// </summary>
        private static bool TryGetContentStatus(Document document, out string status)
        {
            status = null;
            if (document == null)
                return false;

            try
            {
                string xml = document.WordOpenXML;
                if (!string.IsNullOrEmpty(xml) && TryExtractContentStatusFromOpenXml(xml, out string fromXml)
                    && !string.IsNullOrEmpty(fromXml))
                {
                    status = fromXml;
                    return true;
                }
            }
            catch { }

            status = TryGetContentStatusFromCom(document);
            return !string.IsNullOrEmpty(status);
        }

        private static bool TryExtractContentStatusFromOpenXml(string xml, out string status)
        {
            status = null;
            if (string.IsNullOrEmpty(xml))
                return false;

            Match match = Regex.Match(
                xml,
                @"<(?:cp:)?contentStatus\b[^>]*>([\s\S]*?)</(?:cp:)?contentStatus>",
                RegexOptions.IgnoreCase);
            if (!match.Success)
                return false;

            status = System.Net.WebUtility.HtmlDecode(match.Groups[1].Value.Trim());
            return true;
        }

        private static string TryGetContentStatusFromCom(Document document)
        {
            string[] propertyNames =
            {
                "Content status",
                "Content Status",
                "ContentStatus",
                "状態"
            };

            foreach (string name in propertyNames)
            {
                try
                {
                    dynamic prop = ((dynamic)document.BuiltInDocumentProperties)[name];
                    string value = prop?.Value?.ToString();
                    if (!string.IsNullOrEmpty(value))
                        return value;
                }
                catch { }
            }

            return null;
        }

        private bool CheckTask_1_7_04_Impl(string filePath)
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

                string xml;
                try { xml = document.WordOpenXML ?? string.Empty; }
                catch { xml = string.Empty; }

                if (!SmartArtTask704Matches(xml, out string failReason))
                    return Miss(failReason);

                return true;
            }
            finally
            {
                if (document != null)
                    Marshal.ReleaseComObject(document);
            }
        }

        private static bool SmartArtTask704Matches(string xml, out string failReason)
        {
            failReason = "SmartArtの構成が正しくありません。";
            if (string.IsNullOrEmpty(xml))
                return false;

            string dataXml = ExtractDiagramPartXml(xml, "data");
            string drawingXml = ExtractDiagramPartXml(xml, "drawing");
            if (string.IsNullOrEmpty(dataXml) || string.IsNullOrEmpty(drawingXml))
            {
                failReason = "SmartArtのデータが見つかりません。";
                return false;
            }

            if (dataXml.IndexOf("<a:t>" + SmartArtAnchorLabel + "</a:t>", StringComparison.Ordinal) < 0
                || dataXml.IndexOf("<a:t>" + SmartArtSiblingLabel + "</a:t>", StringComparison.Ordinal) < 0
                || dataXml.IndexOf("<a:t>" + SmartArtChildLabel + "</a:t>", StringComparison.Ordinal) < 0)
            {
                failReason = "SmartArtに「点検」「補充」「記録」が揃っていません。";
                return false;
            }

            if (!TryGetDiagramPointModelId(dataXml, SmartArtSiblingLabel, out string siblingId)
                || !TryGetDiagramPointModelId(dataXml, SmartArtChildLabel, out string childId))
            {
                failReason = "SmartArtの「補充」「記録」を特定できません。";
                return false;
            }

            if (!DiagramHasConnection(dataXml, siblingId, childId))
            {
                failReason = "「補充」の下位に「記録」が追加されていません。";
                return false;
            }

            if (!DrawingFontSizeNearLabel(drawingXml, SmartArtSiblingLabel, FontHalfPointsTopLevel))
            {
                failReason = "「補充」のフォントサイズが同レベルの図形（31pt）と一致していません。";
                return false;
            }

            if (!DrawingFontSizeNearLabel(drawingXml, SmartArtChildLabel, FontHalfPointsChild))
            {
                failReason = "「記録」のフォントサイズが14ptになっていません。";
                return false;
            }

            return true;
        }

        private static string ExtractDiagramPartXml(string wordOpenXml, string partKind)
        {
            Match match = Regex.Match(
                wordOpenXml,
                @"<pkg:part pkg:name=""/word/diagrams/" + Regex.Escape(partKind) + @"[^""]*""[^>]*>.*?</pkg:part>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);
            return match.Success ? match.Value : string.Empty;
        }

        private static bool TryGetDiagramPointModelId(string dataXml, string label, out string modelId)
        {
            modelId = null;
            string token = "<a:t>" + label + "</a:t>";
            foreach (Match pt in Regex.Matches(
                dataXml,
                @"<dgm:pt modelId=""(?<id>\{[^""]+\})""[^>]*>(?<body>.*?)</dgm:pt>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase))
            {
                if (pt.Groups["body"].Value.IndexOf(token, StringComparison.Ordinal) < 0)
                    continue;

                modelId = pt.Groups["id"].Value;
                return true;
            }

            return false;
        }

        private static bool DiagramHasConnection(string dataXml, string srcId, string destId)
        {
            foreach (Match cxn in Regex.Matches(dataXml, @"<dgm:cxn\b[^>]*/>|<dgm:cxn\b[^>]*>", RegexOptions.IgnoreCase))
            {
                string tag = cxn.Value;
                if (tag.IndexOf("srcId=\"" + srcId + "\"", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                if (tag.IndexOf("destId=\"" + destId + "\"", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }

            return false;
        }

        private static bool DrawingFontSizeNearLabel(string drawingXml, string label, int expectedHalfPoints)
        {
            int idx = drawingXml.IndexOf(label, StringComparison.Ordinal);
            if (idx < 0)
                return false;

            int start = Math.Max(0, idx - 900);
            int length = Math.Min(1400, drawingXml.Length - start);
            string window = drawingXml.Substring(start, length);

            foreach (Match match in Regex.Matches(window, @"sz=""(?<sz>\d+)"""))
            {
                if (!int.TryParse(match.Groups["sz"].Value, out int halfPoints))
                    continue;

                if (Math.Abs(halfPoints - expectedHalfPoints) <= FontHalfPointsTolerance)
                    return true;
            }

            return false;
        }

        private static bool HasPendingRevisionForPhrase(Document document, string phrase)
        {
            Revisions revisions = null;
            try
            {
                revisions = document.Revisions;
                int count = revisions.Count;
                for (int i = 1; i <= count; i++)
                {
                    Revision revision = null;
                    try
                    {
                        revision = revisions[i];
                        string text = string.Empty;
                        try { text = revision.Range?.Text ?? string.Empty; } catch { }
                        text = NormalizeRevisionText(text);
                        if (text.IndexOf(phrase, StringComparison.Ordinal) >= 0)
                            return true;
                    }
                    finally
                    {
                        if (revision != null)
                            Marshal.ReleaseComObject(revision);
                    }
                }
            }
            catch { }
            finally
            {
                if (revisions != null)
                    Marshal.ReleaseComObject(revisions);
            }

            return false;
        }

        private static bool XmlHasTrackedChangeMarkupForPhrase(string xml, string phrase)
        {
            if (string.IsNullOrEmpty(xml))
                return false;

            int idx = xml.IndexOf(phrase, StringComparison.Ordinal);
            while (idx >= 0)
            {
                int start = Math.Max(0, idx - 400);
                string window = xml.Substring(start, idx - start);
                if (window.IndexOf("<w:ins", StringComparison.OrdinalIgnoreCase) >= 0
                    || window.IndexOf("<w:del", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }

                idx = xml.IndexOf(phrase, idx + phrase.Length, StringComparison.Ordinal);
            }

            return false;
        }

        private static string NormalizeRevisionText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;
            return text.Replace("\r", string.Empty).Replace("\n", string.Empty).Replace("\a", string.Empty);
        }

        private static string GetDocumentPlainText(Document document)
        {
            Range content = null;
            try
            {
                content = document.Content;
                return content?.Text ?? string.Empty;
            }
            finally
            {
                if (content != null)
                    Marshal.ReleaseComObject(content);
            }
        }

        private static bool IsWrapTightInParagraph(Document document, int paraStart, int paraEnd)
        {
            Shapes shapes = null;
            try
            {
                shapes = document.Shapes;
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    Shape shape = null;
                    WrapFormat wrap = null;
                    try
                    {
                        shape = shapes[i];
                        int anchor = -1;
                        try { anchor = shape.Anchor?.Start ?? -1; } catch { }
                        if (anchor < paraStart || anchor > paraEnd)
                            continue;

                        wrap = shape.WrapFormat;
                        if (wrap != null && wrap.Type == WdWrapType.wdWrapTight)
                            return true;
                    }
                    finally
                    {
                        if (wrap != null)
                            Marshal.ReleaseComObject(wrap);
                        if (shape != null)
                            Marshal.ReleaseComObject(shape);
                    }
                }
            }
            catch { }
            finally
            {
                if (shapes != null)
                    Marshal.ReleaseComObject(shapes);
            }

            return false;
        }

        private static Paragraph FindParagraphContaining(Document document, string part)
        {
            foreach (Paragraph paragraph in document.Paragraphs)
            {
                if (ParagraphContainsText(paragraph, part))
                    return paragraph;
                Marshal.ReleaseComObject(paragraph);
            }

            return null;
        }

        private static bool ParagraphContainsText(Paragraph paragraph, string part)
        {
            try
            {
                return (paragraph.Range.Text ?? string.Empty).IndexOf(part, StringComparison.Ordinal) >= 0;
            }
            catch
            {
                return false;
            }
        }

        private static Application GetWordApp()
        {
            try { return (Application)Marshal.GetActiveObject("Word.Application"); }
            catch { return null; }
        }

        private static Document GetDocument(Application wordApp, string filePath)
        {
            string fileName = Path.GetFileName(filePath);
            foreach (Document doc in wordApp.Documents)
            {
                if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                    || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
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
                return wordApp?.ActiveDocument?.FullName;
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

        private static bool Miss(string reason)
        {
            WordScoreExplanation.Note(reason);
            return false;
        }
    }
}
