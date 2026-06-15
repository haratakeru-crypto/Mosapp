using System;
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
        /// <summary>
        /// タスク2-1: 「青空文庫のURLはコチラ↓」の文字列を切り取って、見出し「朗読を楽しみましょう！」の下の段落に貼り付けます。
        /// </summary>
        public bool CheckTask_1_2_01()
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_2_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// タスク2-2: 見出し「方法」の下にある箇条書きのレベルを「3」に変更します。
        /// </summary>
        public bool CheckTask_1_2_02()
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_2_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// タスク2-3: 文書内の「朗読を楽しみましょう！」の文字の色を「青、アクセント1、黒+基本色25％」に変更します。
        /// </summary>
        public bool CheckTask_1_2_03()
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_2_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// タスク2-4: 文書の一番下の図形に、「ご参加お待ちしております」と入力します。
        /// </summary>
        public bool CheckTask_1_2_04()
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_2_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// タスク2-5: テキストボックス内の文字列を15ptの斜体にします。
        /// </summary>
        public bool CheckTask_1_2_05()
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckTask_1_2_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private bool CheckTask_1_2_01(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                try
                {
                    wordApp = (Application)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    wordApp = new Application();
                    wordApp.Visible = true;
                }

                document = null;
                string fileName = System.IO.Path.GetFileName(filePath);

                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        break;
                    }
                }

                if (document == null) return false;

                // 0. Cut / Paste が操作ログまたは証跡で記録されていること
                bool cutParagraph = LogReader.HasTaskEvidence(2, 1, "CutParagraphSelection");
                bool cutExecuted = LogReader.HasTaskEvidence(2, 1, "Cut");
                bool pasteExecuted = LogReader.HasTaskEvidence(2, 1, "Paste");
                if (cutParagraph) return false;
                if (!cutExecuted || !pasteExecuted) return false;

                int count = WordFindHelper.CountTextOccurrences(document, wordApp, "青空文庫のURLはコチラ↓");

                Range urlRange = WordFindHelper.FindFirstRange(document, wordApp, "青空文庫のURLはコチラ↓");
                Range headingRange = WordFindHelper.FindFirstRange(document, wordApp, "朗読を楽しみましょう");
                Range linkRange = WordFindHelper.FindFirstRange(document, wordApp, "https://www.aozora.gr.jp/index.html");
                if (urlRange == null || headingRange == null || linkRange == null)
                {
                    if (urlRange != null) Marshal.ReleaseComObject(urlRange);
                    if (headingRange != null) Marshal.ReleaseComObject(headingRange);
                    if (linkRange != null) Marshal.ReleaseComObject(linkRange);
                    return false;
                }

                // 「青空文庫のURLはコチラ↓」が見出しより後ろにあるか
                int headingStart = headingRange.Start;
                bool isAfterHeading = urlRange.Start > headingStart;

                // リンク段落の1つ上の段落に「青空文庫のURLはコチラ↓」が含まれるか（Paragraph.Range は段落記号を含むため「含まれる」で判定）
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

                bool isCorrectPosition = isAfterHeading && isJustAboveLink;
                return isCorrectPosition && (count == 1);
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_2_02(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                try
                {
                    wordApp = (Application)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    wordApp = new Application();
                    wordApp.Visible = true;
                }

                document = null;
                string fileName = System.IO.Path.GetFileName(filePath);

                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        break;
                    }
                }

                if (document == null) return false;

                Range headingRange = FindExactText(document, "方法");
                if (headingRange == null) return false;

                int level3Count = 0;
                int listCount = 0;
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
                                    listCount++;
                                    if (p.Range.ListFormat.ListLevelNumber == 3) level3Count++;
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
                return listCount == 3 && level3Count == 3;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_2_03(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                try
                {
                    wordApp = (Application)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    wordApp = new Application();
                    wordApp.Visible = true;
                }

                document = null;
                string fileName = System.IO.Path.GetFileName(filePath);

                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        break;
                    }
                }

                if (document == null) return false;

                // 採点対象「朗読を楽しみましょう！」の色を直接チェック（見本テキストは使わない）
                Range targetRange = FindExactText(document, "朗読を楽しみましょう！");

                if (targetRange == null)
                {
                    return false;
                }

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
                        catch { /* ObjectThemeColor が未実装の環境あり */ }
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
                        return colorStateOk;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(font);
                    }
                }
                catch
                {
                    return false;
                }
                finally
                {
                    Marshal.ReleaseComObject(targetRange);
                }
            }
            catch
            {
                return false;
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

        /// <summary>WordOpenXML の w:color に accent1 と themeShade（または lumMod による暗化）があるか。COM の TintAndShade が 0 でもここで判別できる。</summary>
        private static bool TryIsAccent1DarkerFromWordOpenXml(Range range, out string snippet)
        {
            return TryIsAccent1BlackBasic25PercentFromOpenXml(range, out snippet);
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

        private bool CheckTask_1_2_04(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                try
                {
                    wordApp = (Application)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    wordApp = new Application();
                    wordApp.Visible = true;
                }

                document = null;
                string fileName = System.IO.Path.GetFileName(filePath);

                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        break;
                    }
                }

                if (document == null) return false;

                foreach (Shape shape in document.Shapes)
                {
                    if (shape.TextFrame != null && shape.TextFrame.TextRange != null)
                    {
                        string text = shape.TextFrame.TextRange.Text ?? "";
                        if (text.Contains("ご参加お待ちしております"))
                        {
                            Marshal.ReleaseComObject(shape);
                            return true;
                        }
                    }
                    Marshal.ReleaseComObject(shape);
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_2_05(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { wordApp = new Application(); wordApp.Visible = true; }

                document = null;
                string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    { document = doc; break; }
                }

                if (document == null) return false;

                Shape targetShape = null;
                foreach (Shape shape in document.Shapes)
                {
                    if (shape.TextFrame != null && shape.TextFrame.TextRange != null)
                    {
                        if ((shape.TextFrame.TextRange.Text ?? "").Contains("ご参加お待ちしております"))
                        { targetShape = shape; break; }
                    }
                    Marshal.ReleaseComObject(shape);
                }

                if (targetShape == null) return false;

                try
                {
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
                    bool result = is15pt && isItalic && noBold && noUnderline;
                    return result;
                }
                finally
                {
                    Marshal.ReleaseComObject(targetShape);
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
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

