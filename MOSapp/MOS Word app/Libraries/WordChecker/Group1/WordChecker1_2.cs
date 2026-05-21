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
        /// タスク2-4: 文書の一番下の図形に、「いつでもご参加ください」と入力します。
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
                bool cutExecuted = LogReader.HasTaskEvidence(2, 1, "Cut");
                bool pasteExecuted = LogReader.HasTaskEvidence(2, 1, "Paste");
                if (!cutExecuted || !pasteExecuted) return false;

                // 1. 文書全体の「青空文庫のURLはコチラ↓」の数をカウント（1回なら切り取り成功）
                int count = 0;
                Range search = document.Content;
                search.Find.ClearFormatting();
                search.Find.Text = "青空文庫のURLはコチラ↓";
                while (search.Find.Execute())
                {
                    count++;
                    search.Collapse(WdCollapseDirection.wdCollapseEnd);
                }
                Marshal.ReleaseComObject(search);

                // 2. 移動先の位置関係をチェック
                Range urlRange = FindExactText(document, "青空文庫のURLはコチラ↓");
                Range headingRange = FindExactText(document, "朗読を楽しみましょう");
                // 青空文庫のリンク「https://www.aozora.gr.jp/index.html」を含む段落の1つ上に貼り付けられているかで判定
                Range linkRange = FindExactText(document, "https://www.aozora.gr.jp/index.html");
                if (urlRange == null || headingRange == null || linkRange == null)
                {
                    if (urlRange != null) Marshal.ReleaseComObject(urlRange);
                    if (headingRange != null) Marshal.ReleaseComObject(headingRange);
                    if (linkRange != null) Marshal.ReleaseComObject(linkRange);
                    return false;
                }

                // 「青空文庫のURLはコチラ↓」が見出しより後ろにあるか
                bool isAfterHeading = urlRange.Start > headingRange.Start;

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
                        int rgb = 0;
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
                        try { rgb = (int)font.TextColor.RGB; } catch { }
                        int themeRaw = -999;
                        try { themeRaw = (int)font.TextColor.ObjectThemeColor; } catch { }
                        // 「黒+基本色25%」: TintAndShade が効く環境は従来レンジ。Word によっては 0 のまま解決 RGB のみになるため、
                        // 文書テーマのアクセント1（最上段）より十分暗い解決色かどうかで補完する。
                        bool isDarker25ByTint = shade >= -0.31f && shade <= -0.19f;
                        bool isDarker25ByOpenXml = false;
                        string openXmlSnippet = "";
                        try
                        {
                            isDarker25ByOpenXml = TryIsAccent1DarkerFromWordOpenXml(targetRange, out openXmlSnippet);
                        }
                        catch { }
                        bool isDarker25ByLum = false;
                        float baseLum = -1f;
                        float textLum = -1f;
                        bool themeAccentReadOk = false;
                        try
                        {
                            themeAccentReadOk = TryGetDocumentThemeAccent1Rgb(document, out int br, out int bg, out int bb);
                            // COM の RGB はテーマ色で不正確なことがある（ログで textLum が異常に高い）。OpenXML 失敗時のみ輝度比較。
                            if (!isDarker25ByOpenXml && themeAccentReadOk && TryGetFontResolvedLuminanceNoAutoColor(font, out textLum))
                            {
                                baseLum = FontRgbLuminance(br, bg, bb);
                                // COM の RGB が壊れていると textLum が 0.85 超になる。明らかな誤値は輝度比較に使わない。
                                if (textLum < 0.7f)
                                    isDarker25ByLum = textLum < baseLum * 0.88f && textLum < baseLum - 0.02f;
                            }
                            else if (!isDarker25ByOpenXml && !themeAccentReadOk && TryGetFontResolvedLuminanceNoAutoColor(font, out textLum))
                            {
                                if (textLum < 0.7f)
                                    isDarker25ByLum = textLum < 0.52f;
                            }
                        }
                        catch { }
                        bool isDarker25 = isDarker25ByTint || (Math.Abs(shade) < 0.001f && (isDarker25ByOpenXml || isDarker25ByLum));
                        bool isBlue = false;
                        try { isBlue = IsResolvedColorBlue(font.TextColor); } catch { }
                        bool colorStateOk = isAccent1 && isDarker25 && isBlue;
                        // 色の一致を主判定とする（VSTO ログは補助。ログ未取得でも正しいテーマ色なら正解）
                        bool result = colorStateOk;
                        return result;
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

        /// <summary>WordOpenXML の w:color に accent1 と themeShade（または lumMod による暗化）があるか。COM の TintAndShade が 0 でもここで判別できる。</summary>
        private static bool TryIsAccent1DarkerFromWordOpenXml(Range range, out string snippet)
        {
            snippet = "";
            try
            {
                string xml = range.WordOpenXML;
                if (string.IsNullOrEmpty(xml))
                    return false;
                string blob = WebUtility.HtmlDecode(ExpandWordOpenXmlForColorSearch(xml));
                snippet = blob.Length > 480 ? blob.Substring(0, 480) : blob;
                // 同一 w:color 開始タグ内に accent1 と、シェード／明度変更のいずれかがあるか
                foreach (Match m in Regex.Matches(blob, @"<w:color\b[^>]*(?:/>|>)", RegexOptions.IgnoreCase))
                {
                    string tag = m.Value;
                    if (!Regex.IsMatch(tag, @"themeColor\s*=\s*""accent1""", RegexOptions.IgnoreCase))
                        continue;
                    Match mShade = Regex.Match(tag, @"themeShade\s*=\s*""([0-9A-Fa-f]+)""", RegexOptions.IgnoreCase);
                    if (mShade.Success)
                    {
                        int v = Convert.ToInt32(mShade.Groups[1].Value, 16);
                        if (v > 0)
                            return true;
                    }
                    Match mMod = Regex.Match(tag, @"lumMod\s*=\s*""([0-9]+)""", RegexOptions.IgnoreCase);
                    if (mMod.Success && int.TryParse(mMod.Groups[1].Value, out int lumModVal) && lumModVal > 0 && lumModVal < 100000)
                        return true;
                    Match mOff = Regex.Match(tag, @"lumOff\s*=\s*""([0-9]+)""", RegexOptions.IgnoreCase);
                    if (mOff.Success && int.TryParse(mOff.Groups[1].Value, out int lumOffVal) && lumOffVal != 0)
                        return true;
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
                        if (text.Contains("いつでもご参加ください"))
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
                        if ((shape.TextFrame.TextRange.Text ?? "").Contains("いつでもご参加ください"))
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

        /// <summary>
        /// 検索時のCOMエラーや改行コードの巻き込みを防ぐためのヘルパー。見つかった場合はそのRangeを返す（呼び出し元で解放すること）。見つからない場合はnull。
        /// </summary>
        private Range FindExactText(Document doc, string searchText)
        {
            Range range = doc.Content;
            range.Find.ClearFormatting();
            object findText = searchText;
            object matchCase = false;
            object matchWholeWord = false;
            object matchWildcards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object wrap = WdFindWrap.wdFindStop;
            object format = false;
            object replaceWith = Type.Missing;
            object replace = Type.Missing;
            object matchKashida = Type.Missing;
            object matchDiacritics = Type.Missing;
            object matchAlefHamza = Type.Missing;
            object matchControl = Type.Missing;
            bool found = range.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildcards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWith, ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
            if (!found)
            {
                Marshal.ReleaseComObject(range);
                return null;
            }
            return range;
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

