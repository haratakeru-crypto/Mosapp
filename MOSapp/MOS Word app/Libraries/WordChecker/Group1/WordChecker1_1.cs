using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_1
    {
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

                // 1-1: トグル証跡2回以上 かつ FinalOn/FinalOff の last-wins が FinalOn。
                // View.ShowAll は再オープンで消えるため採点式には使わない（その場＝一括）。
                // FinalOff で以前の FinalOn を無効化する（2回〇のあと3回目OFFで×）。
                bool showAllExecutedTwice = LogReader.HasTaskEvidenceAtLeast(1, 1, "ShowAll", 2);
                if (!showAllExecutedTwice)
                    return Miss("編集記号の表示について、必要な操作の記録がありません。");

                if (LogReader.IsLatestShowAllFinalOn(1, 1))
                    return true;

                return Miss("編集記号が表示されていません。");
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_1_02_Impl(string filePath)
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

                Paragraph p1 = FindTargetParagraph(document, "社員のコンプライアンス意識の確立");
                if (p1 == null)
                    return Miss("「社員のコンプライアンス意識の確立」の段落がありません。");

                Paragraph heading = FindTargetParagraph(document, "CSR活動のメリット");
                if (heading == null)
                {
                    Marshal.ReleaseComObject(p1);
                    return Miss("見出し「CSR活動のメリット」がありません。");
                }

                if (p1.Range.Start <= heading.Range.Start)
                {
                    Marshal.ReleaseComObject(heading);
                    Marshal.ReleaseComObject(p1);
                    return Miss("「社員のコンプライアンス意識の確立」が見出し「CSR活動のメリット」の下にありません。");
                }

                object countOne = 1;
                Paragraph p2 = (Paragraph)p1.Next(ref countOne);
                Paragraph p3 = p2 != null ? (Paragraph)p2.Next(ref countOne) : null;
                Paragraph p4 = p3 != null ? (Paragraph)p3.Next(ref countOne) : null;

                bool b1 = p1.Range.ListFormat.ListType == WdListType.wdListBullet;
                bool b2 = p2 != null && p2.Range.ListFormat.ListType == WdListType.wdListBullet;
                bool b3 = p3 != null && p3.Range.ListFormat.ListType == WdListType.wdListBullet;
                bool headingNotBullet = heading.Range.ListFormat.ListType != WdListType.wdListBullet;
                bool nextNotBullet = p4 == null || p4.Range.ListFormat.ListType != WdListType.wdListBullet;

                int bulletCountInRange = 0;
                Paragraph cur = heading;
                Paragraph stopAfter = p4 ?? p3;
                while (cur != null)
                {
                    if (cur.Range.ListFormat.ListType == WdListType.wdListBullet)
                        bulletCountInRange++;
                    if (stopAfter != null && cur.Range.Start == stopAfter.Range.Start)
                        break;
                    object one = 1;
                    Paragraph next = (Paragraph)cur.Next(ref one);
                    if (cur != heading) Marshal.ReleaseComObject(cur);
                    cur = next;
                }
                if (cur != null && cur != heading) Marshal.ReleaseComObject(cur);

                bool exactThree = bulletCountInRange == 3;
                bool result = b1 && b2 && b3 && headingNotBullet && nextNotBullet && exactThree;

                if (!result)
                {
                    if (!b1 || !b2 || !b3)
                        WordScoreExplanation.Note("見出し「CSR活動のメリット」の下の「社員のコンプライアンス意識の確立」から3つの段落が箇条書きになっていません。");
                    if (!headingNotBullet)
                        WordScoreExplanation.Note("見出し「CSR活動のメリット」が箇条書きになっています。");
                    if (!nextNotBullet)
                        WordScoreExplanation.Note("箇条書きが3つを超えています。");
                    if (!exactThree && (b1 && b2 && b3) && headingNotBullet && nextNotBullet)
                        WordScoreExplanation.Note($"見出し「CSR活動のメリット」の下の箇条書きの数が「{bulletCountInRange}」になっています。");
                }

                if (p4 != null) Marshal.ReleaseComObject(p4);
                if (p3 != null) Marshal.ReleaseComObject(p3);
                if (p2 != null) Marshal.ReleaseComObject(p2);
                Marshal.ReleaseComObject(heading);
                Marshal.ReleaseComObject(p1);

                return result;
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>1-3: 最後の段落（画像の左）にスタイル「参照2」が付いているかで判定する</summary>
        private bool CheckTask_1_1_03_Impl(string filePath)
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

                int count = document.Paragraphs.Count;
                if (count < 1)
                    return Miss("画像の左の段落がありません。");

                // 最後の段落を取得（末尾が空段落の場合はその手前）＝画像の左
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

                try
                {
                    string pName = GetParagraphStyleSafe(target);
                    string rName = GetRangeStyleSafe(target.Range);
                    bool ok = pName.Contains("参照2") || pName.Contains("reference2") ||
                              rName.Contains("参照2") || rName.Contains("reference2");
                    if (ok)
                        return true;

                    string display = GetParagraphStyleDisplayName(target);
                    if (string.IsNullOrWhiteSpace(display))
                        return Miss("画像の左の段落がスタイル「参照２」になっていません。");
                    return Miss($"画像の左の段落のスタイルが「{Quote(display)}」になっています。");
                }
                finally
                {
                    Marshal.ReleaseComObject(target);
                }
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>1-4: 小見出し「CSR活動の光と影」のスタイルが「見出し３」であることを判定する</summary>
        private bool CheckTask_1_1_04_Impl(string filePath)
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

                Paragraph target = FindTargetParagraph(document, "CSR活動の光と影");
                if (target == null)
                    return Miss("小見出し「CSR活動の光と影」がありません。");

                try
                {
                    string pName = GetParagraphStyleSafe(target);
                    string rName = GetRangeStyleSafe(target.Range);
                    bool ok = pName.Contains("見出し3") || pName.Contains("heading3") ||
                              rName.Contains("見出し3") || rName.Contains("heading3");
                    if (ok)
                        return true;

                    string display = GetParagraphStyleDisplayName(target);
                    if (string.IsNullOrWhiteSpace(display))
                        return Miss("「CSR活動の光と影」がスタイル「見出し３」になっていません。");
                    return Miss($"「CSR活動の光と影」のスタイルが「{Quote(display)}」になっています。");
                }
                finally
                {
                    Marshal.ReleaseComObject(target);
                }
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_1_05_Impl(string filePath)
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

                if (target == null)
                    return Miss("見出し「第１節　はじめに」の下の「会社の社会的責任」の段落がありません。");

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
                if (logOk && roughlyCleared)
                {
                    if (font != null) Marshal.ReleaseComObject(font);
                    Marshal.ReleaseComObject(target);
                    return true;
                }

                try
                {
                    if (strictlyCleared)
                        return true;

                    if (isBold)
                        WordScoreExplanation.Note("「会社の社会的責任」の段落に太字が残っています。");
                    if (isItalic)
                        WordScoreExplanation.Note("「会社の社会的責任」の段落に斜体が残っています。");
                    if (isUnderline)
                        WordScoreExplanation.Note("「会社の社会的責任」の段落に下線が残っています。");
                    if (!isColorAutomatic)
                        WordScoreExplanation.Note("「会社の社会的責任」の段落の文字の色が自動になっていません。");
                    if (!isNormal)
                    {
                        string display = GetParagraphStyleDisplayName(target);
                        if (string.IsNullOrWhiteSpace(display))
                            WordScoreExplanation.Note("「会社の社会的責任」の段落のスタイルが「標準」になっていません。");
                        else
                            WordScoreExplanation.Note($"「会社の社会的責任」の段落のスタイルが「{Quote(display)}」になっています。");
                    }

                    if (isBold || isItalic || isUnderline || !isColorAutomatic || !isNormal)
                        return false;

                    return Miss("「会社の社会的責任」の段落の書式がクリアされていません。");
                }
                finally
                {
                    if (font != null) Marshal.ReleaseComObject(font);
                    Marshal.ReleaseComObject(target);
                }
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
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

        /// <summary>COMオブジェクトから安全にスタイル名を取り出す（照合用・正規化済み）</summary>
        private string GetStyleNameSafe(object styleObj)
        {
            string name = GetStyleDisplayNameRaw(styleObj);
            if (string.IsNullOrEmpty(name))
                return "";

            name = name.Replace(" ", "").Replace("　", "").ToLower();
            name = name.Replace("０", "0").Replace("１", "1").Replace("２", "2").Replace("３", "3")
                       .Replace("４", "4").Replace("５", "5").Replace("６", "6").Replace("７", "7")
                       .Replace("８", "8").Replace("９", "9");
            return name;
        }

        /// <summary>学生向け表示用のスタイル名（正規化しない）。埋め込み時は Quote する。</summary>
        private string GetParagraphStyleDisplayName(Paragraph target)
        {
            return GetStyleDisplayNameRaw(GetStyleObject(target));
        }

        private static string GetStyleDisplayNameRaw(object styleObj)
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
                    return name.Trim();
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
            Application wordApp = null;
            try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = null; }

            return WordFindHelper.PreserveSelection(wordApp, () => FindTargetParagraphCore(doc, searchText));
        }

        private Paragraph FindTargetParagraphCore(Document doc, string searchText)
        {
            Paragraph firstNonTOC = null;
            Paragraph lastAny = null;
            Range searchRange = WordFindHelper.DuplicateContent(doc);
            if (searchRange == null)
                return null;

            Find find = searchRange.Find;
            WordFindHelper.ConfigureSafeFind(find, searchText);

            try
            {
                while (find.Execute())
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
