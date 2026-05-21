using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_1
    {
        public bool CheckTask_1_1_01()
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath)) return false;
                return CheckTask_1_1_01(filePath);
            }
            catch { return false; }
        }

        public bool CheckTask_1_1_02()
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath)) return false;
                return CheckTask_1_1_02(filePath);
            }
            catch { return false; }
        }

        public bool CheckTask_1_1_03()
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath)) return false;
                return CheckTask_1_1_03(filePath);
            }
            catch { return false; }
        }

        public bool CheckTask_1_1_04()
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath)) return false;
                return CheckTask_1_1_04(filePath);
            }
            catch { return false; }
        }

        public bool CheckTask_1_1_05()
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath)) return false;
                return CheckTask_1_1_05(filePath);
            }
            catch { return false; }
        }

        private bool CheckTask_1_1_01(string filePath)
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
                if (document == null) return false;

                // 1-1: 証跡で ShowAll 2回以上、かつ編集記号表示で正解（個別リセット後の旧全体ログ誤判定を防ぐ）
                bool showAllExecutedTwice = LogReader.HasTaskEvidenceAtLeast(1, 1, "ShowAll", 2);
                if (!showAllExecutedTwice) return false;

                bool showAll = wordApp.ActiveWindow.View.ShowAll;
                return showAll;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_1_02(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                document = GetDocument(wordApp, filePath);
                if (document == null) return false;

                Paragraph p1 = FindTargetParagraph(document, "社員のコンプライアンス意識の確立");
                if (p1 == null) return false;

                object countOne = 1;
                Paragraph p2 = (Paragraph)p1.Next(ref countOne);
                Paragraph p3 = p2 != null ? (Paragraph)p2.Next(ref countOne) : null;

                bool b1 = p1.Range.ListFormat.ListType == WdListType.wdListBullet;
                bool b2 = p2 != null && p2.Range.ListFormat.ListType == WdListType.wdListBullet;
                bool b3 = p3 != null && p3.Range.ListFormat.ListType == WdListType.wdListBullet;

                if (p3 != null) Marshal.ReleaseComObject(p3);
                if (p2 != null) Marshal.ReleaseComObject(p2);
                Marshal.ReleaseComObject(p1);

                return b1 && b2 && b3;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>1-3: 最後の段落にスタイル「参照2」が付いているかで判定する</summary>
        private bool CheckTask_1_1_03(string filePath)
        {
            System.Diagnostics.Debug.WriteLine("1-3 entry");
            Application wordApp = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                document = GetDocument(wordApp, filePath);
                if (document == null) { System.Diagnostics.Debug.WriteLine("1-3 exit: document null"); return false; }
                System.Diagnostics.Debug.WriteLine("1-3 GetDocument ok");

                int count = document.Paragraphs.Count;
                if (count < 1) { System.Diagnostics.Debug.WriteLine("1-3 exit: count<1"); return false; }
                System.Diagnostics.Debug.WriteLine("1-3 count=" + count);

                // 最後の段落を取得（末尾が空段落の場合はその手前）
                Paragraph target = document.Paragraphs[count];
                string lastText = "";
                try
                {
                    Range r = target.Range;
                    try { lastText = (r?.Text ?? "").Trim().Replace("\r", "").Replace("\n", "").Replace("\a", ""); }
                    finally { if (r != null) Marshal.ReleaseComObject(r); }
                }
                catch { }
                if (string.IsNullOrEmpty(lastText) && count >= 2)
                {
                    if (target != null) Marshal.ReleaseComObject(target);
                    target = document.Paragraphs[count - 1];
                }
                System.Diagnostics.Debug.WriteLine("1-3 target set");

                try
                {
                    string pName = GetParagraphStyleSafe(target);
                    string rName = GetRangeStyleSafe(target.Range);

                    System.Diagnostics.Debug.WriteLine("1-3 pName=[" + (pName ?? "") + "] rName=[" + (rName ?? "") + "]");

                    return pName.Contains("参照2") || pName.Contains("reference2") ||
                           rName.Contains("参照2") || rName.Contains("reference2");
                }
                finally
                {
                    Marshal.ReleaseComObject(target);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("1-3 exception: " + (ex?.Message ?? ""));
                return false;
            }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>1-4: 「はじめに」以降で「日本におけるCSR活動」を含む段落を探し、スタイル「見出し2」を判定する</summary>
        private bool CheckTask_1_1_04(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                document = GetDocument(wordApp, filePath);
                if (document == null) return false;

                Paragraph target = null;
                bool foundHajimeni = false;

                try
                {
                    foreach (Paragraph p in document.Paragraphs)
                    {
                        try
                        {
                            string text = p.Range.Text ?? "";
                            string normalizedText = text.Replace(" ", "").Replace("　", "")
                                                        .Replace("\r", "").Replace("\n", "").Replace("\a", "");

                            if (normalizedText.Contains("はじめに"))
                            {
                                foundHajimeni = true;
                            }
                            else if (foundHajimeni && normalizedText.Contains("日本におけるCSR活動"))
                            {
                                target = p;
                                break;
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                if (target == null) return false;

                try
                {
                    string pName = GetParagraphStyleSafe(target);
                    string rName = GetRangeStyleSafe(target.Range);

                    return pName.Contains("見出し2") || pName.Contains("heading2") ||
                           rName.Contains("見出し2") || rName.Contains("heading2");
                }
                finally
                {
                    Marshal.ReleaseComObject(target);
                }
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_1_05(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            // #region agent log
            try
            {
                string logPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? "", "..", "..", "..", "..", "debug-340a5e.log"));
                long ts = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
                string line = "{\"sessionId\":\"340a5e\",\"hypothesisId\":\"H0\",\"location\":\"WordChecker1_1.CheckTask_1_1_05\",\"message\":\"entry\",\"data\":{\"filePath\":\"" + (filePath ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"},\"timestamp\":" + ts + "}\n";
                File.AppendAllText(logPath, line, Encoding.UTF8);
            }
            catch { }
            // #endregion
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                document = GetDocument(wordApp, filePath);
                if (document == null) return false;

                // 1-5: 「会社の社会的責任(Corporate Social Responsibility　以下CSR)は…」の長い段落を探し、その段落の書式がクリアされていれば正解
                Paragraph target = null;
                bool foundHajimeni = false;
                const string targetParagraphStart = "会社の社会的責任(Corporate Social Responsibility　以下CSR)は";
                string normalizedTargetStart = targetParagraphStart.Replace(" ", "").Replace("　", "").Trim();

                try
                {
                    foreach (Paragraph p in document.Paragraphs)
                    {
                        try
                        {
                            string text = p.Range.Text ?? "";
                            string normalizedText = text.Replace(" ", "").Replace("　", "")
                                                        .Replace("\r", "").Replace("\n", "").Replace("\a", "").Replace("\t", "").Trim();

                            if (normalizedText.Contains("はじめに"))
                            {
                                foundHajimeni = true;
                            }
                            else if (foundHajimeni && normalizedText.Contains(normalizedTargetStart))
                            {
                                target = p;
                                break;
                            }
                        }
                        catch { }
                    }
                }
                catch { }

                // フォールバック: 「はじめに」の直後で見つからない場合、文書内で該当長文段落を探す
                if (target == null)
                {
                    try
                    {
                        foreach (Paragraph p in document.Paragraphs)
                        {
                            try
                            {
                                string text = p.Range.Text ?? "";
                                string normalizedText = text.Replace(" ", "").Replace("　", "")
                                                            .Replace("\r", "").Replace("\n", "").Replace("\a", "").Replace("\t", "").Trim();
                                if (normalizedText.Contains(normalizedTargetStart))
                                {
                                    target = p;
                                    break;
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }

                // #region agent log
                try
                {
                    long ts = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
                    string line = "{\"sessionId\":\"340a5e\",\"hypothesisId\":\"H1\",\"location\":\"WordChecker1_1.CheckTask_1_1_05\",\"message\":\"target\",\"data\":{\"targetNull\":" + (target == null ? "true" : "false") + "},\"timestamp\":" + ts + "}\n";
                    string logPath1 = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? "", "..", "..", "..", "..", "debug-340a5e.log"));
                    try { File.AppendAllText(logPath1, line, Encoding.UTF8); } catch { }
                    string logPath2 = null;
                    try { string logDir = Path.GetDirectoryName(filePath); if (!string.IsNullOrEmpty(logDir)) logPath2 = Path.Combine(logDir, "debug-340a5e.log"); } catch { }
                    if (!string.IsNullOrEmpty(logPath2)) try { File.AppendAllText(logPath2, line, Encoding.UTF8); } catch { }
                    if (target == null)
                    {
                        var samples = new System.Collections.Generic.List<string>();
                        try
                        {
                            int n = 0;
                            foreach (Paragraph p in document.Paragraphs)
                            {
                                if (n >= 40) break;
                                try
                                {
                                    string text = p.Range.Text ?? "";
                                    string normalizedText = text.Replace(" ", "").Replace("　", "").Replace("\r", "").Replace("\n", "").Replace("\a", "").Replace("\t", "").Trim();
                                    if (normalizedText.Length > 0) { samples.Add(normalizedText.Length <= 80 ? normalizedText : normalizedText.Substring(0, 80) + "..."); n++; }
                                }
                                catch { }
                            }
                        }
                        catch { }
                        string escaped = string.Join("|", samples).Replace("\\", "\\\\").Replace("\"", "\\\"");
                        string line2 = "{\"sessionId\":\"340a5e\",\"hypothesisId\":\"H1b\",\"message\":\"paragraphSamples\",\"data\":{\"samples\":\"" + escaped + "\"},\"timestamp\":" + ts + "}\n";
                        try { File.AppendAllText(logPath1, line2, Encoding.UTF8); } catch { }
                        if (!string.IsNullOrEmpty(logPath2)) try { File.AppendAllText(logPath2, line2, Encoding.UTF8); } catch { }
                    }
                }
                catch { }
                // #endregion

                if (target == null) return false;

                Font font = null;
                try { font = target.Range.Font; } catch { }
                bool isBold = false;
                bool isItalic = false;
                int underline = 0;
                bool isColorAutomatic = true;
                if (font != null)
                {
                    try { isBold = FontTriStateEqualsTrue(font.Bold); } catch { }
                    try { isItalic = FontTriStateEqualsTrue(font.Italic); } catch { }
                    try { underline = GetFontUnderlineValue(font); } catch { }
                    try { isColorAutomatic = (int)font.Color == (int)WdColor.wdColorAutomatic; } catch { }
                }
                // 下線: 0=なし, -1=未定義(wdUndefined) はクリア済み。正の値は下線あり
                bool isUnderline = underline > 0;

                string styleName = "";
                try { styleName = GetParagraphStyleSafe(target); } catch { }
                bool isNormal = IsNormalStyleName(styleName);

                // ログ優先: リボンの「書式のクリア」(ClearFormatting) が採点ログにあれば、書式がおおむねクリアされていれば正解
                bool logOk = LogReader.HasTaskEvidence(1, 5, "ClearFormatting");
                bool strictlyCleared = isNormal && !isBold && !isItalic && !isUnderline && isColorAutomatic;
                bool roughlyCleared = !isBold && !isItalic && !isUnderline && isColorAutomatic;
                bool passedLogRough = logOk && roughlyCleared;

                // #region agent log
                try
                {
                    string logPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? "", "..", "..", "..", "..", "debug-340a5e.log"));
                    long ts = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
                    string styleEsc = (styleName ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"");
                    string line = "{\"sessionId\":\"340a5e\",\"hypothesisId\":\"H2-H4\",\"location\":\"WordChecker1_1.CheckTask_1_1_05\",\"message\":\"conditions\",\"data\":{\"logOk\":" + (logOk ? "true" : "false") + ",\"roughlyCleared\":" + (roughlyCleared ? "true" : "false") + ",\"strictlyCleared\":" + (strictlyCleared ? "true" : "false") + ",\"passedLogRough\":" + (passedLogRough ? "true" : "false") + ",\"isBold\":" + (isBold ? "true" : "false") + ",\"isItalic\":" + (isItalic ? "true" : "false") + ",\"isUnderline\":" + (isUnderline ? "true" : "false") + ",\"isColorAutomatic\":" + (isColorAutomatic ? "true" : "false") + ",\"isNormal\":" + (isNormal ? "true" : "false") + ",\"styleName\":\"" + styleEsc + "\",\"underline\":" + underline + "},\"timestamp\":" + ts + "}\n";
                    File.AppendAllText(logPath, line, Encoding.UTF8);
                }
                catch { }
                // #endregion

                if (logOk && roughlyCleared)
                    return true;
                // ログなし時は従来の厳密なファイル判定
                try
                {
                    bool result = strictlyCleared;
                    // #region agent log
                    try
                    {
                        string logPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? "", "..", "..", "..", "..", "debug-340a5e.log"));
                        long ts = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
                        string line = "{\"sessionId\":\"340a5e\",\"hypothesisId\":\"H5\",\"location\":\"WordChecker1_1.CheckTask_1_1_05\",\"message\":\"return\",\"data\":{\"result\":" + (result ? "true" : "false") + "},\"timestamp\":" + ts + "}\n";
                        File.AppendAllText(logPath, line, Encoding.UTF8);
                    }
                    catch { }
                    // #endregion
                    return result;
                }
                finally
                {
                    if (font != null) Marshal.ReleaseComObject(font);
                    Marshal.ReleaseComObject(target);
                }
            }
            catch (Exception ex)
            {
                // #region agent log
                try
                {
                    string logPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? "", "..", "..", "..", "..", "debug-340a5e.log"));
                    long ts = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
                    string exEsc = (ex?.Message ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"");
                    string line = "{\"sessionId\":\"340a5e\",\"hypothesisId\":\"H5\",\"location\":\"WordChecker1_1.CheckTask_1_1_05\",\"message\":\"catch\",\"data\":{\"ex\":\"" + exEsc + "\"},\"timestamp\":" + ts + "}\n";
                    File.AppendAllText(logPath, line, Encoding.UTF8);
                }
                catch { }
                // #endregion
                return false;
            }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>
        /// オブジェクトから安全に .get_Style() または .Style を呼び出す（dynamic 廃止・RuntimeBinderException 防止）
        /// </summary>
        private object GetStyleObject(object target)
        {
            if (target == null) return null;
            try
            {
                if (target is Paragraph para)
                {
                    Range r = null;
                    try
                    {
                        r = para.Range;
                        if (r == null) return null;
                        try { return r.get_Style(); } catch { }
                        return GetStyleObjectByInvoke(r);
                    }
                    finally { if (r != null) Marshal.ReleaseComObject(r); }
                }

                if (target is Range range)
                {
                    try { return range.get_Style(); } catch { }
                    return GetStyleObjectByInvoke(range);
                }

                return GetStyleObjectByInvoke(target);
            }
            catch { }
            return null;
        }

        private static object GetStyleObjectByInvoke(object target)
        {
            if (target == null) return null;
            var flagsInv = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.InvokeMethod;
            var flagsProp = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty;
            try
            {
                object s = target.GetType().InvokeMember("get_Style", flagsInv, null, target, null);
                if (s != null) return s;
            }
            catch { }
            try
            {
                object s = target.GetType().InvokeMember("Style", flagsProp, null, target, null);
                if (s != null) return s;
            }
            catch { }
            return null;
        }

        /// <summary>COMオブジェクトから安全にスタイル名を取り出す</summary>
        private string GetStyleNameSafe(object styleObj)
        {
            if (styleObj == null) return "";

            string name = "";
            try
            {
                if (styleObj is string s)
                {
                    name = s;
                }
                else if (styleObj is Style styleObjAsStyle)
                {
                    try { name = styleObjAsStyle.NameLocal; } catch { }
                    if (string.IsNullOrEmpty(name))
                    {
                        try
                        {
                            var flags = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty;
                            object nameObj = styleObjAsStyle.GetType().InvokeMember("Name", flags, null, styleObjAsStyle, null);
                            if (nameObj != null) name = nameObj.ToString();
                        }
                        catch { }
                    }
                }

                if (string.IsNullOrEmpty(name))
                {
                    try
                    {
                        Type type = styleObj.GetType();
                        var flags = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty;
                        object nameLocalObj = type.InvokeMember("NameLocal", flags, null, styleObj, null);
                        if (nameLocalObj != null) name = nameLocalObj.ToString();
                    }
                    catch { }
                }
                if (string.IsNullOrEmpty(name))
                {
                    try
                    {
                        Type type = styleObj.GetType();
                        var flags = System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.GetProperty;
                        object nameObj = type.InvokeMember("Name", flags, null, styleObj, null);
                        if (nameObj != null) name = nameObj.ToString();
                    }
                    catch { }
                }

                if (string.IsNullOrEmpty(name))
                {
                    name = styleObj.ToString();
                }

                if (!string.IsNullOrEmpty(name) && name != "System.__ComObject")
                {
                    name = name.Replace(" ", "").Replace("　", "").ToLower();
                    name = name.Replace("０", "0").Replace("１", "1").Replace("２", "2").Replace("３", "3")
                               .Replace("４", "4").Replace("５", "5").Replace("６", "6").Replace("７", "7")
                               .Replace("８", "8").Replace("９", "9");
                    return name;
                }
            }
            catch { }
            return "";
        }

        /// <summary>指定範囲が目次(TOC)フィールド内かどうか</summary>
        private bool IsRangeInTOCField(Document doc, Range range)
        {
            if (doc == null || range == null) return false;
            try
            {
                foreach (Field f in doc.Fields)
                {
                    Range fr = null;
                    try
                    {
                        if (f.Type != WdFieldType.wdFieldTOC) continue;
                        fr = f.Result;
                        if (fr != null && range.Start >= fr.Start && range.End <= fr.End)
                            return true;
                    }
                    finally { if (fr != null) Marshal.ReleaseComObject(fr); }
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 目次を回避しつつ、Wordのネイティブ検索(Find)を使って正確に対象段落を取得する。
        /// 目次外で最初にヒットした段落を返す。全て目次内の場合は最後のヒットを返す。
        /// </summary>
        private Paragraph FindTargetParagraph(Document doc, string searchText)
        {
            Paragraph firstNonTOC = null;
            Paragraph lastAny = null;
            Range searchRange = doc.Content;
            searchRange.Find.ClearFormatting();
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

            try
            {
                while (searchRange.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildcards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWith, ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl))
                {
                    Paragraph p = searchRange.Paragraphs[1];
                    if (IsRangeInTOCField(doc, p.Range))
                    {
                        if (lastAny != null) Marshal.ReleaseComObject(lastAny);
                        lastAny = p;
                    }
                    else
                    {
                        if (firstNonTOC != null) Marshal.ReleaseComObject(firstNonTOC);
                        firstNonTOC = p;
                        if (lastAny != null) { Marshal.ReleaseComObject(lastAny); lastAny = null; }
                        break;
                    }
                    searchRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(searchRange);
            }
            if (firstNonTOC != null)
            {
                if (lastAny != null) Marshal.ReleaseComObject(lastAny);
                return firstNonTOC;
            }
            return lastAny;
        }

        /// <summary>Word の Font.Bold/Italic の TriState: 書式「あり」のとき true。wdTrue=-1, wdFalse=0 のため -1 と 1 を「あり」とする。</summary>
        private static bool FontTriStateEqualsTrue(object value)
        {
            if (value == null) return false;
            if (value is int i) return i == 1 || i == -1;
            if (value is bool b) return b;
            try { int v = Convert.ToInt32(value); return v == 1 || v == -1; }
            catch { return false; }
        }

        /// <summary>Font.Underline を int で取得（COM の戻り値に対応）</summary>
        private static int GetFontUnderlineValue(Font font)
        {
            if (font == null) return 0;
            try
            {
                object u = font.Underline;
                if (u == null) return 0;
                return Convert.ToInt32(u);
            }
            catch { return 0; }
        }

        /// <summary>「標準」系スタイルかどうか。書式クリア後に Word が付けるスタイル名も許容する。</summary>
        private static bool IsNormalStyleName(string styleName)
        {
            if (string.IsNullOrEmpty(styleName)) return true;
            string s = styleName.Trim().ToLowerInvariant();
            if (s.Length == 0) return true;
            return s.Contains("標準") ||
                   s.Contains("normal") ||
                   s.Contains("bodytext") ||
                   s.Contains("default") ||
                   s.Contains("char") ||
                   s.Contains("本文") ||
                   s.Contains("テキスト");
        }

        /// <summary>段落のスタイル名を安全に取得する</summary>
        private string GetParagraphStyleSafe(Paragraph target)
        {
            return GetStyleNameSafe(GetStyleObject(target));
        }

        /// <summary>範囲の先頭文字のスタイル名を安全に取得する</summary>
        private string GetRangeStyleSafe(Range range)
        {
            if (range == null) return "";
            try
            {
                if (range.Characters.Count > 0)
                {
                    return GetStyleNameSafe(GetStyleObject(range.Characters[1]));
                }
            }
            catch { }
            return "";
        }

        /// <summary>ファイル名から安全にドキュメントを取得する</summary>
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

