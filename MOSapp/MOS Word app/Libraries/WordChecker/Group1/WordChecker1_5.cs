using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_5
    {
        public bool CheckTask_1_5_01() => RunCheck(CheckTask_1_5_01_Impl);
        public bool CheckTask_1_5_02() => RunCheck(CheckTask_1_5_02_Impl);
        public bool CheckTask_1_5_03() => RunCheck(CheckTask_1_5_03_Impl);
        public bool CheckTask_1_5_04() => RunCheck(CheckTask_1_5_04_Impl);
        public bool CheckTask_1_5_05() => RunCheck(CheckTask_1_5_05_Impl);
        public bool CheckTask_1_5_06() => RunCheck(CheckTask_1_5_06_Impl);
        public bool CheckTask_1_5_07() => RunCheck(CheckTask_1_5_07_Impl);
        public bool CheckTask_1_5_08() => RunCheck(CheckTask_1_5_08_Impl);

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

        private bool CheckTask_1_5_01_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);
                Range searchRange = WordFindHelper.DuplicateContent(document); Find find = searchRange.Find; WordFindHelper.ConfigureSafeFind(find, "5月21日より5日間の"); find.Execute();
                if (!find.Found)
                {
                    Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                    return Miss("「5月21日より5日間の…」の段落が見つかりません。");
                }
                Range paraRange = searchRange.Paragraphs[1].Range;
                int paraStart = paraRange.Start;
                int paraEnd = paraRange.End;

                // 「5月21日より5日間の...」段落の画像の折り返しが「上下」(wdWrapTopBottom) か
                bool result = IsWrapTopBottomInParagraph(document, paraStart, paraEnd);
                bool hasImage = HasImageInParagraph(document, paraRange, paraStart, paraEnd);

                Marshal.ReleaseComObject(paraRange);
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);
                // 5-1: 現在「上下」、または当該タスクの証跡で WrapTopBottom 操作あり（5-2 後はログで判定）
                bool logOk = LogReader.HasTaskEvidence(5, 1, "WrapTopBottom");
                if (result || logOk)
                    return true;

                // A: 未挿入 / B: 画像ありだが上下でない（＋記録なし。5-2 後の状態上書き時の説明用）
                if (!hasImage)
                    return Miss("対象の段落に画像が挿入されていません。");
                WordScoreExplanation.Note("画像の文字列の折り返しが「上下」になっていません。");
                return Miss("上下にした操作の記録もありません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>指定段落内に行内または浮動の画像があるか。</summary>
        private static bool HasImageInParagraph(Document document, Range paraRange, int paraStart, int paraEnd)
        {
            if (document == null || paraRange == null)
                return false;

            try
            {
                InlineShapes inlines = paraRange.InlineShapes;
                try
                {
                    if (inlines != null && inlines.Count > 0)
                        return true;
                }
                finally { if (inlines != null) Marshal.ReleaseComObject(inlines); }
            }
            catch { }

            Microsoft.Office.Interop.Word.Shapes shapes = null;
            try
            {
                shapes = document.Shapes;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Microsoft.Office.Interop.Word.Shape sh = null;
                    try
                    {
                        sh = shapes[i];
                        int anchor = sh.Anchor != null ? sh.Anchor.Start : -1;
                        if (anchor >= paraStart && anchor <= paraEnd)
                            return true;
                    }
                    catch { }
                    finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                }
            }
            catch { }
            finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
            return false;
        }

        /// <summary>指定段落内にアンカーされた浮動 Shape の折り返しが指定タイプか。</summary>
        private static bool IsWrapTypeInParagraph(Document document, int paraStart, int paraEnd, WdWrapType wrapType)
        {
            if (document == null) return false;
            Microsoft.Office.Interop.Word.Shapes shapes = null;
            try
            {
                shapes = document.Shapes;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Microsoft.Office.Interop.Word.Shape sh = null;
                    try
                    {
                        sh = shapes[i];
                        int anchor = sh.Anchor != null ? sh.Anchor.Start : -1;
                        if (anchor < paraStart || anchor > paraEnd) continue;
                        WrapFormat wf = sh.WrapFormat;
                        try
                        {
                            if (wf != null && wf.Type == wrapType)
                                return true;
                        }
                        finally { if (wf != null) Marshal.ReleaseComObject(wf); }
                    }
                    catch { }
                    finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                }
            }
            catch { }
            finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
            return false;
        }

        /// <summary>指定段落内にアンカーされた画像の折り返しが「上下」(wdWrapTopBottom) か。</summary>
        private static bool IsWrapTopBottomInParagraph(Document document, int paraStart, int paraEnd)
        {
            return IsWrapTypeInParagraph(document, paraStart, paraEnd, WdWrapType.wdWrapTopBottom);
        }

        /// <summary>指定段落内にアンカーされた画像の折り返しが「狭く」(wdWrapTight) か。</summary>
        private static bool IsWrapTightInParagraph(Document document, int paraStart, int paraEnd)
        {
            return IsWrapTypeInParagraph(document, paraStart, paraEnd, WdWrapType.wdWrapTight);
        }

        private bool CheckTask_1_5_02_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);
                Range searchRange = WordFindHelper.DuplicateContent(document); Find find = searchRange.Find; WordFindHelper.ConfigureSafeFind(find, "5月21日より5日間の"); find.Execute();
                if (!find.Found)
                {
                    Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                    return Miss("「5月21日より5日間の…」の段落が見つかりません。");
                }
                Range paraRange = searchRange.Paragraphs[1].Range;
                int paraStart = paraRange.Start;
                int paraEnd = paraRange.End;

                bool result = IsWrapTightInParagraph(document, paraStart, paraEnd);

                Marshal.ReleaseComObject(paraRange);
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);
                // 5-2: 最終状態が「狭く」であれば ○（操作証跡は不要。後続タスクでも折り返しは維持される想定）
                if (result)
                    return true;
                return Miss("画像の文字列の折り返しが「狭く」になっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_03_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                // 5-3: 「5月21日より5日間の」段落付近の画像に限定
                string imageXml = GetImageXmlNearText(xml, "5月21日より5日間の");
                if (string.IsNullOrEmpty(imageXml))
                    return Miss("「5月21日より5日間の…」の段落付近の画像が見つかりません。");

                // その画像に「水彩：スポンジ」が適用されているか
                if (imageXml.Contains("artisticWatercolorSponge"))
                    return true;
                return Miss("アート効果「水彩：スポンジ」が設定されていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_04_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                // 5-4: 「5月21日より5日間の」段落付近の画像に限定
                string imageXml = GetImageXmlNearText(xml, "5月21日より5日間の");
                if (string.IsNullOrEmpty(imageXml))
                    return Miss("「5月21日より5日間の…」の段落付近の画像が見つかりません。");

                // ぼかし25ポイント (317500 EMU)
                if (imageXml.Contains("softEdge") && imageXml.Contains("317500"))
                    return true;
                return Miss("図の効果「ぼかし25ポイント」が設定されていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_05_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                // 5-5: 「TOEICテスト対策セミナー」付近の画像（タイトル画像）に限定
                string imageXml = GetImageXmlNearText(xml, "TOEICテスト対策セミナー");
                if (string.IsNullOrEmpty(imageXml))
                    return Miss("「TOEICテスト対策セミナー…」のタイトル付近の画像が見つかりません。");

                // 面取り ハードエッジ
                if (imageXml.Contains("bevelT") && imageXml.Contains("hardEdge"))
                    return true;
                return Miss("図の効果「面取り ハードエッジ」が設定されていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_06_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                // 5-5 と同じく「TOEICテスト対策セミナー」付近の画像に限定
                string imageXml = GetImageXmlNearText(xml, "TOEICテスト対策セミナー");
                if (string.IsNullOrEmpty(imageXml))
                    return Miss("「TOEICテスト対策セミナー…」のタイトル付近の画像が見つかりません。");

                const string expected = "セミナー案内";
                string alt = TryGetPictureDescriptionFromImageXml(imageXml);
                if (alt == null)
                    alt = TryGetAlternativeTextNearText(document, "TOEICテスト対策セミナー") ?? "";

                string trimmed = (alt ?? "").Trim();
                if (string.Equals(trimmed, expected, StringComparison.Ordinal))
                    return true;
                if (string.IsNullOrEmpty(trimmed))
                    return Miss("代替テキストが入力されていません。");
                return Miss($"代替テキストが「{QuoteForStudent(trimmed)}」と入力されています。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>画像 XML の docPr / v:shape から代替テキスト（descr）を取得。属性自体が無いときは null。</summary>
        private static string TryGetPictureDescriptionFromImageXml(string imageXml)
        {
            if (string.IsNullOrEmpty(imageXml))
                return null;

            var docPr = System.Text.RegularExpressions.Regex.Match(
                imageXml,
                @"<(?:wp|wp14):docPr\b([^>]*)/?>",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);
            if (docPr.Success)
            {
                string attrs = docPr.Groups[1].Value;
                if (TryReadXmlAttribute(attrs, "descr", out string descr))
                    return descr;
                if (TryReadXmlAttribute(attrs, "description", out string description))
                    return description;
            }

            var vShape = System.Text.RegularExpressions.Regex.Match(
                imageXml,
                @"<v:shape\b([^>]*)>",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);
            if (vShape.Success && TryReadXmlAttribute(vShape.Groups[1].Value, "alt", out string alt))
                return alt;

            return null;
        }

        private static bool TryReadXmlAttribute(string attrs, string attrName, out string value)
        {
            value = null;
            if (string.IsNullOrEmpty(attrs) || string.IsNullOrEmpty(attrName))
                return false;
            var m = System.Text.RegularExpressions.Regex.Match(
                attrs,
                attrName + @"\s*=\s*""([^""]*)""",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (!m.Success)
                return false;
            value = m.Groups[1].Value;
            return true;
        }

        /// <summary>指定テキストを含む段落付近にアンカーされた Shape の AlternativeText。</summary>
        private static string TryGetAlternativeTextNearText(Document document, string targetText)
        {
            if (document == null || string.IsNullOrEmpty(targetText))
                return null;

            Range searchRange = null;
            Find find = null;
            try
            {
                searchRange = WordFindHelper.DuplicateContent(document);
                find = searchRange.Find;
                WordFindHelper.ConfigureSafeFind(find, targetText);
                find.Execute();
                if (!find.Found)
                    return null;

                Range paraRange = searchRange.Paragraphs[1].Range;
                int paraStart = paraRange.Start;
                int paraEnd = paraRange.End;
                // タイトル画像は直前段落に置かれることが多い
                int lookStart = paraStart;
                try
                {
                    if (searchRange.Paragraphs[1].Range.Start > document.Content.Start)
                    {
                        Paragraph prev = searchRange.Paragraphs[1].Previous(1);
                        if (prev != null)
                        {
                            lookStart = prev.Range.Start;
                            Marshal.ReleaseComObject(prev);
                        }
                    }
                }
                catch { }

                string found = null;
                Microsoft.Office.Interop.Word.Shapes shapes = null;
                try
                {
                    shapes = document.Shapes;
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        Microsoft.Office.Interop.Word.Shape sh = null;
                        try
                        {
                            sh = shapes[i];
                            int anchor = sh.Anchor != null ? sh.Anchor.Start : -1;
                            if (anchor < lookStart || anchor > paraEnd)
                                continue;
                            found = sh.AlternativeText ?? "";
                            break;
                        }
                        catch { }
                        finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                    }
                }
                finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }

                try
                {
                    InlineShapes inlines = paraRange.InlineShapes;
                    try
                    {
                        if (string.IsNullOrEmpty(found) && inlines != null && inlines.Count > 0)
                            found = inlines[1].AlternativeText ?? "";
                    }
                    finally { if (inlines != null) Marshal.ReleaseComObject(inlines); }
                }
                catch { }

                Marshal.ReleaseComObject(paraRange);
                return found;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
            }
        }

        private static string QuoteForStudent(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "（空）";
            string text = value.Replace("\r", "").Replace("\n", " ");
            const int maxLen = 40;
            if (text.Length <= maxLen)
                return text;
            return text.Substring(0, maxLen) + "…";
        }

        private bool CheckTask_1_5_07_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                // 5-3/5-4 と同じく「5月21日より5日間の」段落付近の画像に限定
                string imageXml = GetImageXmlNearText(xml, "5月21日より5日間の");
                if (string.IsNullOrEmpty(imageXml))
                    return Miss("「5月21日より5日間の…」の段落付近の画像が見つかりません。");

                // 装飾用: <adec:decorative xmlns:adec=".../2017/decorative" val="1"/>
                if (IsDecorativePictureXml(imageXml))
                    return true;
                return Miss("対象の画像の代替テキストが装飾用になっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>画像 XML 内に装飾用フラグ（adec:decorative val="1"）があるか。</summary>
        private static bool IsDecorativePictureXml(string imageXml)
        {
            if (string.IsNullOrEmpty(imageXml))
                return false;
            if (imageXml.IndexOf("2017/decorative", StringComparison.OrdinalIgnoreCase) < 0)
                return false;

            // 対象ブロック内の decorative 要素だけを見る（文書全体の別画像を拾わない）
            var m = System.Text.RegularExpressions.Regex.Match(
                imageXml,
                @"<(?:\w+:)?decorative\b([^>]*)/?>",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Singleline);
            if (!m.Success)
                return false;
            return System.Text.RegularExpressions.Regex.IsMatch(
                m.Groups[1].Value,
                @"\bval\s*=\s*""1""",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        private bool CheckTask_1_5_08_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return Miss(WordScoreExplanation.UnavailableText);

                // 5-8: 文書の「最後」の画像を特定する
                var matches = System.Text.RegularExpressions.Regex.Matches(xml, @"<w:drawing>.*?</w:drawing>", System.Text.RegularExpressions.RegexOptions.Singleline);
                if (matches.Count == 0)
                    return Miss("文末の画像が見つかりません。");

                // 最後の画像ブロックを取得
                string lastImageXml = matches[matches.Count - 1].Value;

                // 領域保持（foregroundMark）の座標 y1 を解析
                var yMatches = System.Text.RegularExpressions.Regex.Matches(lastImageXml, @"y1=""(\d+)""");
                bool hasTopMark = false;
                bool hasBottomMark = false;
                foreach (System.Text.RegularExpressions.Match m in yMatches)
                {
                    if (int.TryParse(m.Groups[1].Value, out int y))
                    {
                        if (y < 40000) hasTopMark = true;
                        if (y > 60000) hasBottomMark = true;
                    }
                }

                // 上下両方のエリアにマークがあり、背景削除が確定されていれば合格
                if (hasTopMark && hasBottomMark)
                    return true;
                return Miss("文末の画像の背景が削除されていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>
        /// 指定したテキストを含む段落内、またはその直前にある画像のXMLブロックを抽出します。
        /// </summary>
        private string GetImageXmlNearText(string fullXml, string targetText)
        {
            var paragraphs = System.Text.RegularExpressions.Regex.Matches(fullXml, @"<w:p\b[^>]*>.*?</w:p>", System.Text.RegularExpressions.RegexOptions.Singleline);

            for (int i = 0; i < paragraphs.Count; i++)
            {
                string pXml = paragraphs[i].Value;
                string pText = System.Text.RegularExpressions.Regex.Replace(pXml, @"<[^>]+>", "");

                if (pText.Contains(targetText))
                {
                    // 1. まず同じ段落内で画像を探す
                    var drawingMatch = System.Text.RegularExpressions.Regex.Match(pXml, @"<(w:drawing|v:shape).*?>.*?</\1>", System.Text.RegularExpressions.RegexOptions.Singleline);
                    if (drawingMatch.Success) return drawingMatch.Value;

                    // 2. なければ直前の段落を探す（タイトル画像などは直前の段落に置かれることが多いため）
                    if (i > 0)
                    {
                        string prevPXml = paragraphs[i - 1].Value;
                        var prevDrawingMatch = System.Text.RegularExpressions.Regex.Match(prevPXml, @"<(w:drawing|v:shape).*?>.*?</\1>", System.Text.RegularExpressions.RegexOptions.Singleline);
                        if (prevDrawingMatch.Success) return prevDrawingMatch.Value;
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine($"[GetImageXmlNearText] テキスト '{targetText}' 付近（直前含む）に画像が見つかりませんでした。");
            return "";
        }

        private static bool Miss(string reason)
        {
            WordScoreExplanation.Note(reason);
            return false;
        }

        private string GetCurrentWordFilePath()
        {
            Application wordApp = null;
            try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); if (wordApp.ActiveDocument != null) return wordApp.ActiveDocument.FullName; return null; }
            catch (COMException) { return null; }
            finally { if (wordApp != null) Marshal.ReleaseComObject(wordApp); }
        }
    }
}
