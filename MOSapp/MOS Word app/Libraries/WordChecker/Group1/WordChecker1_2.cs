using System;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_2
    {
        public bool CheckTask_1_2_01() => RunCheck(CheckTask_1_2_01_Impl);
        public bool CheckTask_1_2_02() => RunCheck(CheckTask_1_2_02_Impl);
        public bool CheckTask_1_2_03() => RunCheck(CheckTask_1_2_03_Impl);
        public bool CheckTask_1_2_04() => RunCheck(CheckTask_1_2_04_Impl);
        public bool CheckTask_1_2_05() => RunCheck(CheckTask_1_2_05_Impl);

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

        /// <summary>
        /// タスク2-1: 「青空文庫のURLはコチラ↓」を切り取って、「朗読を楽しみましょう！」の下の段落に貼り付け。
        /// </summary>
        private bool CheckTask_1_2_01_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch
                {
                    wordApp = new Application();
                    wordApp.Visible = true;
                }

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                bool cutParagraph = LogReader.HasTaskEvidence(2, 1, "CutParagraphSelection");
                bool cutExecuted = LogReader.HasTaskEvidence(2, 1, "Cut");
                bool pasteExecuted = LogReader.HasTaskEvidence(2, 1, "Paste");
                if (cutParagraph)
                    return Miss("段落ごと切り取った記録があるため、×になりました。");
                if (!cutExecuted && !pasteExecuted)
                    return Miss("切り取りまたは貼り付けの操作の記録がありません。");
                if (!cutExecuted)
                    return Miss("切り取りの操作の記録がありません。");
                if (!pasteExecuted)
                    return Miss("貼り付けの操作の記録がありません。");

                int count = WordFindHelper.CountTextOccurrences(document, wordApp, "青空文庫のURLはコチラ↓");

                Range urlRange = WordFindHelper.FindFirstRange(document, wordApp, "青空文庫のURLはコチラ↓");
                Range headingRange = WordFindHelper.FindFirstRange(document, wordApp, "朗読を楽しみましょう");
                Range linkRange = WordFindHelper.FindFirstRange(document, wordApp, "https://www.aozora.gr.jp/index.html");
                if (urlRange == null || headingRange == null || linkRange == null)
                {
                    string missReason = urlRange == null
                        ? "「青空文庫のURLはコチラ↓」が文書内にありません。"
                        : headingRange == null
                            ? "「朗読を楽しみましょう！」が文書内にありません。"
                            : "貼り付け位置の確認に必要な文字列が見つかりません。";
                    if (urlRange != null) Marshal.ReleaseComObject(urlRange);
                    if (headingRange != null) Marshal.ReleaseComObject(headingRange);
                    if (linkRange != null) Marshal.ReleaseComObject(linkRange);
                    return Miss(missReason);
                }

                int headingStart = headingRange.Start;
                bool isAfterHeading = urlRange.Start > headingStart;

                Paragraph linkPara = linkRange.Paragraphs[1];
                Paragraph abovePara = null;
                try
                {
                    object one = 1;
                    abovePara = linkPara.Previous(ref one) as Paragraph;
                }
                catch { }

                bool isJustAboveLink = false;
                if (abovePara != null)
                {
                    Range aboveRange = abovePara.Range;
                    isJustAboveLink = (urlRange.Start >= aboveRange.Start && urlRange.End <= aboveRange.End);
                    Marshal.ReleaseComObject(aboveRange);
                }

                Marshal.ReleaseComObject(urlRange);
                Marshal.ReleaseComObject(headingRange);
                Marshal.ReleaseComObject(linkRange);
                if (abovePara != null) Marshal.ReleaseComObject(abovePara);
                Marshal.ReleaseComObject(linkPara);

                if (count != 1)
                    WordScoreExplanation.Note($"「青空文庫のURLはコチラ↓」が文書内に「{count}」か所あります。");
                if (!isAfterHeading || !isJustAboveLink)
                    WordScoreExplanation.Note("「青空文庫のURLはコチラ↓」が「朗読を楽しみましょう！」の下の段落になっていません。");

                bool isCorrectPosition = isAfterHeading && isJustAboveLink;
                if (isCorrectPosition && count == 1)
                    return true;
                return false;
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        /// <summary>
        /// タスク2-2: 見出し「方法」の下にある箇条書きのレベルを「3」に変更。
        /// </summary>
        private bool CheckTask_1_2_02_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return Miss(WordScoreExplanation.UnavailableText); }

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                Range headingRange = FindExactText(document, "方法");
                if (headingRange == null)
                    return Miss("見出し「方法」がありません。");

                int level3Count = 0;
                int listCount = 0;
                var levels = new System.Collections.Generic.List<int>();
                try
                {
                    foreach (Paragraph p in document.Paragraphs)
                    {
                        if (p.Range.Start > headingRange.Start)
                        {
                            string text = (p.Range.Text ?? "").Trim('\r', '\n', '\a', ' ', '　');
                            if (text.Length > 0)
                            {
                                if (p.Range.ListFormat.ListType != WdListType.wdListNoNumbering)
                                {
                                    int level = p.Range.ListFormat.ListLevelNumber;
                                    listCount++;
                                    levels.Add(level);
                                    if (level == 3) level3Count++;
                                }
                                if (listCount >= 3) break;
                            }
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(headingRange);
                }

                if (listCount == 3 && level3Count == 3)
                    return true;

                if (listCount == 0)
                    return Miss("見出し「方法」の下に箇条書きがありません。");
                if (listCount != 3)
                    WordScoreExplanation.Note($"見出し「方法」の下の箇条書きの数が「{listCount}」になっています。");
                if (level3Count != 3 && levels.Count > 0)
                {
                    bool allSame = levels.TrueForAll(l => l == levels[0]);
                    if (allSame)
                        WordScoreExplanation.Note($"見出し「方法」の下の箇条書きがレベル「{levels[0]}」になっています。");
                    else
                        WordScoreExplanation.Note($"見出し「方法」の下の箇条書きのレベルが「{string.Join("」「", levels)}」になっています。");
                }
                return false;
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        /// <summary>
        /// タスク2-3: 「朗読を楽しみましょう！」の文字色を「青、アクセント1、黒+基本色25％」に変更。
        /// </summary>
        private bool CheckTask_1_2_03_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return Miss(WordScoreExplanation.UnavailableText); }

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                Range targetRange = FindExactText(document, "朗読を楽しみましょう！");
                if (targetRange == null)
                    return Miss("「朗読を楽しみましょう！」が文書内にありません。");

                try
                {
                    Font font = targetRange.Characters[1].Font;
                    try
                    {
                        bool isAccent1 = false;
                        float shade = 0f;
                        try
                        {
                            isAccent1 = font.TextColor.ObjectThemeColor == WdThemeColorIndex.wdThemeColorAccent1;
                        }
                        catch { }
                        try
                        {
                            object tintVal = font.TextColor.TintAndShade;
                            if (tintVal != null)
                            {
                                if (tintVal is float f) shade = f;
                                else if (tintVal is double d) shade = (float)d;
                                else shade = Convert.ToSingle(tintVal);
                            }
                        }
                        catch { }
                        bool isDarker25 = IsAccent1BlackBasic25Percent(font, targetRange, document, shade);
                        bool isBlue = false;
                        try { isBlue = IsResolvedColorBlue(font.TextColor); } catch { }
                        bool colorStateOk = isAccent1 && isDarker25 && isBlue;
                        if (colorStateOk)
                            return true;

                        WordScoreExplanation.Note("「朗読を楽しみましょう！」の文字の色が「青、アクセント1、黒+基本色25％」になっていません。");
                        return false;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(font);
                    }
                }
                catch
                {
                    return Miss(WordScoreExplanation.UnavailableText);
                }
                finally
                {
                    Marshal.ReleaseComObject(targetRange);
                }
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        /// <summary>0〜1 の相対輝度（ITU-R BT.601 係数、RGB 0〜255）。</summary>
        private static float FontRgbLuminance(int r, int g, int b)
        {
            r = Math.Max(0, Math.Min(255, r));
            g = Math.Max(0, Math.Min(255, g));
            b = Math.Max(0, Math.Min(255, b));
            return (0.299f * r + 0.587f * g + 0.114f * b) / 255f;
        }

        /// <summary>テーマ色で Font.Color が誤った解決値になるため使わない。TextColor.RGB の下位24bitのみ解釈する。</summary>
        private static bool TryGetFontResolvedLuminanceNoAutoColor(Font font, out float lum)
        {
            lum = 0f;
            try
            {
                int rgb = (int)font.TextColor.RGB;
                uint u = unchecked((uint)rgb) & 0xFFFFFFu;
                int r = (int)(u & 0xFF);
                int g = (int)((u >> 8) & 0xFF);
                int b = (int)((u >> 16) & 0xFF);
                lum = FontRgbLuminance(r, g, b);
                return true;
            }
            catch { }
            return false;
        }

        /// <summary>Flat OPC の WordOpenXML から word 本文相当のマークアップを連結する（先頭数百文字だけでは w:color が無い）。</summary>
        private static string ExpandWordOpenXmlForColorSearch(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return xml;
            var sb = new StringBuilder(xml.Length + 512);
            sb.Append(xml);
            foreach (Match m in Regex.Matches(xml, @"<pkg:xmlData[^>]*>([\s\S]*?)</pkg:xmlData>", RegexOptions.IgnoreCase))
                sb.Append('\n').Append(m.Groups[1].Value);
            return sb.ToString();
        }

        /// <summary>「黒+基本色25%」（アクセント1を約25%暗く）のみ true。50%暗色は拒否。</summary>
        private static bool IsAccent1BlackBasic25Percent(Font font, Range range, Document document, float tintAndShade)
        {
            if (IsTintAndShade25Percent(tintAndShade))
                return true;

            if (Math.Abs(tintAndShade) >= 0.001f)
                return false;

            if (TryIsAccent1BlackBasic25PercentFromOpenXml(range, out _))
                return true;

            try
            {
                if (TryGetDocumentThemeAccent1Rgb(document, out int br, out int bg, out int bb)
                    && TryGetFontResolvedLuminanceNoAutoColor(font, out float textLum))
                {
                    float baseLum = FontRgbLuminance(br, bg, bb);
                    if (textLum >= 0.7f)
                        return false;
                    float ratio = baseLum > 0.001f ? textLum / baseLum : 0f;
                    if (ratio < 0.60f)
                        return false;
                    return ratio >= 0.70f && ratio <= 0.80f;
                }
            }
            catch { }

            return false;
        }

        private static bool IsTintAndShade25Percent(float shade)
        {
            if (shade <= -0.35f)
                return false;
            return shade >= -0.31f && shade <= -0.19f;
        }

        private static bool IsThemeShade25Percent(int themeShade)
        {
            if (themeShade >= 0x70 && themeShade <= 0x90)
                return false;
            return themeShade >= 0xB0 && themeShade <= 0xC8;
        }

        private static bool IsLumMod25Percent(int lumModVal)
        {
            if (lumModVal >= 45000 && lumModVal <= 55000)
                return false;
            return lumModVal >= 70000 && lumModVal <= 85000;
        }

        /// <summary>WordOpenXML の w:color に accent1 + 25%暗色があるか。</summary>
        private static bool TryIsAccent1BlackBasic25PercentFromOpenXml(Range range, out string snippet)
        {
            snippet = "";
            try
            {
                string xml = range.WordOpenXML;
                if (string.IsNullOrEmpty(xml))
                    return false;
                string blob = WebUtility.HtmlDecode(ExpandWordOpenXmlForColorSearch(xml));
                snippet = blob.Length > 480 ? blob.Substring(0, 480) : blob;
                foreach (Match m in Regex.Matches(blob, @"<w:color\b[^>]*(?:/>|>)", RegexOptions.IgnoreCase))
                {
                    string tag = m.Value;
                    if (!Regex.IsMatch(tag, @"themeColor\s*=\s*""accent1""", RegexOptions.IgnoreCase))
                        continue;
                    Match mShade = Regex.Match(tag, @"themeShade\s*=\s*""([0-9A-Fa-f]+)""", RegexOptions.IgnoreCase);
                    if (mShade.Success)
                    {
                        int v = Convert.ToInt32(mShade.Groups[1].Value, 16);
                        if (IsThemeShade25Percent(v))
                            return true;
                        if (v >= 0x70 && v <= 0x90)
                            return false;
                    }
                    Match mMod = Regex.Match(tag, @"lumMod\s*=\s*""([0-9]+)""", RegexOptions.IgnoreCase);
                    if (mMod.Success && int.TryParse(mMod.Groups[1].Value, out int lumModVal))
                    {
                        if (IsLumMod25Percent(lumModVal))
                            return true;
                        if (lumModVal >= 45000 && lumModVal <= 55000)
                            return false;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>文書テーマのアクセント1（ギャラリー最上段）の RGB。Office の型を直接参照しないため dynamic で取得する。</summary>
        private static bool TryGetDocumentThemeAccent1Rgb(Document document, out int r, out int g, out int b)
        {
            r = g = b = 0;
            object dtObj = null;
            object schemeObj = null;
            object cfObj = null;
            try
            {
                dynamic doc = document;
                dtObj = doc.DocumentTheme;
                if (dtObj == null) return false;
                dynamic dt = dtObj;
                schemeObj = dt.ThemeColorScheme;
                if (schemeObj == null) return false;
                dynamic scheme = schemeObj;
                // MsoThemeColorIndex: アクセント1 = 5（Office 共通）
                cfObj = scheme.Colors(5);
                if (cfObj == null) return false;
                dynamic cf = cfObj;
                int rgb = (int)cf.RGB;
                uint u = unchecked((uint)rgb) & 0xFFFFFFu;
                r = (int)(u & 0xFF);
                g = (int)((u >> 8) & 0xFF);
                b = (int)((u >> 16) & 0xFF);
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (cfObj != null) Marshal.ReleaseComObject(cfObj);
                if (schemeObj != null) Marshal.ReleaseComObject(schemeObj);
                if (dtObj != null) Marshal.ReleaseComObject(dtObj);
            }
        }

        /// <summary>
        /// 解決後の色が青系かどうかを判定する。
        /// Office の環境差で RGB/BGR 解釈がずれる場合があるため、両方で青優位を許容する。
        /// </summary>
        private static bool IsResolvedColorBlue(dynamic textColor)
        {
            try
            {
                int rgb = (int)textColor.RGB;
                int r1 = rgb & 0xFF;
                int g1 = (rgb >> 8) & 0xFF;
                int b1 = (rgb >> 16) & 0xFF;
                bool blueByOrder1 = b1 >= r1 && b1 >= g1 && b1 > 0;

                int r2 = (rgb >> 16) & 0xFF;
                int g2 = (rgb >> 8) & 0xFF;
                int b2 = rgb & 0xFF;
                bool blueByOrder2 = b2 >= r2 && b2 >= g2 && b2 > 0;

                return blueByOrder1 || blueByOrder2;
            }
            catch { return false; }
        }

        private const string Task24ExpectedText = "ご参加お待ちしております";

        /// <summary>
        /// タスク2-4: 文書の一番下の図形に「ご参加お待ちしております」と入力。
        /// </summary>
        private bool CheckTask_1_2_04_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            Shape bottomShape = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return Miss(WordScoreExplanation.UnavailableText); }

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                bottomShape = FindBottomTextShape(document);
                if (bottomShape == null)
                    return Miss("文書の一番下の図形がありません。");

                string raw = "";
                try { raw = bottomShape.TextFrame.TextRange.Text ?? ""; } catch { }
                string display = NormalizeShapeText(raw);
                if (display.Contains(Task24ExpectedText))
                    return true;
                if (string.IsNullOrEmpty(display))
                    return Miss($"文書の一番下の図形に「{Task24ExpectedText}」がありません。");
                return Miss($"文書の一番下の図形の文字列が「{Quote(display)}」になっています。");
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (bottomShape != null) Marshal.ReleaseComObject(bottomShape);
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        /// <summary>
        /// タスク2-5: テキストボックス内の文字列を15ptの斜体に。
        /// 文言正誤に依存せず、一番下のテキスト付図形の書式を見る。
        /// </summary>
        private bool CheckTask_1_2_05_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            Shape targetShape = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return Miss(WordScoreExplanation.UnavailableText); }

                document = GetDocument(wordApp, filePath);
                if (document == null)
                    return Miss(WordScoreExplanation.UnavailableText);

                targetShape = FindBottomTextShape(document);
                if (targetShape == null)
                    return Miss("テキスト付図形がありません。");

                Font font = targetShape.TextFrame.TextRange.Characters[1].Font;
                float size = 0f;
                try { size = font.Size; } catch { }
                bool is15pt = Math.Abs(size - 15.0f) < 0.5f;
                int italicVal = 0;
                int boldVal = 0;
                int underlineVal = 0;
                try { italicVal = font.Italic; } catch { }
                try { boldVal = font.Bold; } catch { }
                try { underlineVal = (int)font.Underline; } catch { }
                bool isItalic = (italicVal != 0 && italicVal != 9999999);
                bool noBold = (boldVal == 0);
                bool noUnderline = (underlineVal == (int)WdUnderline.wdUnderlineNone);
                Marshal.ReleaseComObject(font);

                if (is15pt && isItalic && noBold && noUnderline)
                    return true;

                if (!is15pt)
                    WordScoreExplanation.Note($"テキストボックス内の文字サイズが「{Quote(size.ToString("0.#"))}」ptになっています。");
                if (!isItalic)
                    WordScoreExplanation.Note("テキストボックス内の文字が斜体になっていません。");
                if (!noBold)
                    WordScoreExplanation.Note("テキストボックス内の文字に太字が残っています。");
                if (!noUnderline)
                    WordScoreExplanation.Note("テキストボックス内の文字に下線が残っています。");
                return false;
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (targetShape != null) Marshal.ReleaseComObject(targetShape);
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        /// <summary>TextFrame を持つ図形のうち、ページ上で最も下（Top 最大）のものを返す。呼び出し側で Release すること。</summary>
        private static Shape FindBottomTextShape(Document document)
        {
            Shape best = null;
            float bestTop = float.MinValue;
            foreach (Shape shape in document.Shapes)
            {
                bool keep = false;
                try
                {
                    if (!HasUsableTextFrame(shape))
                        continue;
                    float top = shape.Top;
                    if (best == null || top >= bestTop)
                    {
                        if (best != null) Marshal.ReleaseComObject(best);
                        best = shape;
                        bestTop = top;
                        keep = true;
                    }
                }
                catch { }
                finally
                {
                    if (!keep)
                        Marshal.ReleaseComObject(shape);
                }
            }
            return best;
        }

        private static bool HasUsableTextFrame(Shape shape)
        {
            try
            {
                TextFrame tf = shape.TextFrame;
                if (tf == null) return false;
                try
                {
                    // HasText: msoTrue(-1) / msoFalse(0)。未対応図形は例外になり得る。
                    int hasText = tf.HasText;
                    if (hasText == 0)
                        return tf.TextRange != null;
                    return true;
                }
                catch
                {
                    return tf.TextRange != null;
                }
            }
            catch
            {
                return false;
            }
        }

        private static string NormalizeShapeText(string raw)
        {
            if (string.IsNullOrEmpty(raw))
                return "";
            return raw.Trim('\r', '\n', '\a', ' ', '　');
        }

        private static bool Miss(string reason)
        {
            WordScoreExplanation.Note(reason);
            return false;
        }

        private static string Quote(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "（空）";
            string text = value.Replace("\r", "").Replace("\n", " ");
            const int maxLen = 40;
            if (text.Length <= maxLen)
                return text;
            return text.Substring(0, maxLen) + "…";
        }

        private Document GetDocument(Application wordApp, string filePath)
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

        private Range FindExactText(Document doc, string searchText)
        {
            Application wordApp = null;
            try
            {
                wordApp = (Application)Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                return null;
            }
            try
            {
                return WordFindHelper.FindFirstRange(doc, wordApp, searchText);
            }
            finally
            {
                if (wordApp != null)
                    Marshal.ReleaseComObject(wordApp);
            }
        }

        private string GetCurrentWordFilePath()
        {
            Application wordApp = null;
            try
            {
                wordApp = (Application)Marshal.GetActiveObject("Word.Application");
                if (wordApp.ActiveDocument != null)
                {
                    return wordApp.ActiveDocument.FullName;
                }
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
    }
}
