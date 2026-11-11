using System;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_10
    {
        private const string Task03ComboFailReason =
            "行頭文字がフォント「Webdings」・文字コード「120」の組み合わせになっていません。";

        public bool CheckTask_1_10_01() => RunCheck(CheckTask_1_10_01_Impl);
        public bool CheckTask_1_10_02() => RunCheck(CheckTask_1_10_02_Impl);
        public bool CheckTask_1_10_03() => RunCheck(CheckTask_1_10_03_Impl);
        public bool CheckTask_1_10_04() => RunCheck(CheckTask_1_10_04_Impl);
        public bool CheckTask_1_10_05() => RunCheck(CheckTask_1_10_05_Impl);

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

        private bool CheckTask_1_10_01_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            Range searchRange = null;
            Find find = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return Miss(WordScoreExplanation.UnavailableText); }

                document = GetDocument(wordApp, filePath);
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                // 表記ゆれは許容しない。「ウイルス」で Find し、その直前に「コンピュータ」（6文字）があるか（＝コンピュータウイルス）を厳密に判定
                const int computerLen = 6; // 「コンピュータ」の文字数
                searchRange = WordFindHelper.DuplicateContent(document);
                find = searchRange.Find;
                WordFindHelper.ConfigureSafeFind(find, "ウイルス");
                object findTextVirus = "ウイルス";
                object matchCase = false;
                object matchWholeWord = false;
                object matchWildcards = false;
                object matchSoundsLike = false;
                object matchAllWordForms = false;
                object forward = true;
                object wrap = WdFindWrap.wdFindStop;
                object format = false;
                object replaceWith = System.Reflection.Missing.Value;
                object replace = WdReplace.wdReplaceNone;
                object matchKashida = false;
                object matchDiacritics = false;
                object matchAlefHamza = false;
                object matchControl = false;

                bool atLeastOneReplacement = false;
                bool allVirusPrecededByComputer = true;
                int safetyCounter = 0;
                int foundCount = 0;

                while (find.Execute(ref findTextVirus, ref matchCase, ref matchWholeWord, ref matchWildcards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWith, ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl))
                {
                    safetyCounter++;
                    if (safetyCounter > 1000)
                    {
                        allVirusPrecededByComputer = false;
                        break;
                    }

                    if (!find.Found) break;
                    foundCount++;
                    Range virusRange = searchRange.Duplicate;
                    try
                    {
                        virusRange.Collapse(WdCollapseDirection.wdCollapseStart);
                        int moveResult = virusRange.MoveStart(WdUnits.wdCharacter, -computerLen);
                        if (moveResult != 0)
                        {
                            string before = virusRange.Text ?? "";
                            if (before.StartsWith("コンピュータ"))
                            {
                                atLeastOneReplacement = true;
                            }
                            else
                            {
                                allVirusPrecededByComputer = false;
                                break;
                            }
                        }
                        else
                        {
                            allVirusPrecededByComputer = false;
                            break;
                        }
                    }
                    finally
                    {
                        if (virusRange != null) Marshal.ReleaseComObject(virusRange);
                    }
                    searchRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    searchRange.Move(WdUnits.wdCharacter, 1);
                }

                if (atLeastOneReplacement && allVirusPrecededByComputer)
                    return true;

                if (foundCount == 0)
                    return Miss("「ウイルス」が「コンピュータウイルス」に置き換えられていません。");
                if (!allVirusPrecededByComputer)
                    return Miss("「ウイルス」が「コンピュータウイルス」に置き換えられていません。");
                return Miss("「ウイルス」が「コンピュータウイルス」に置き換えられていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally
            {
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_10_02_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;

            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return Miss(WordScoreExplanation.UnavailableText); }

                document = GetDocument(wordApp, filePath);
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                bool stateOk = IsTask02StateCorrect(document);
                bool logOk = LogReader.HasTaskEvidence(10, 2, "BulletDefineNew");

                if (stateOk && logOk)
                    return true;

                bool noted = false;
                if (!stateOk)
                {
                    WordScoreExplanation.Note("行頭文字がPCアイコンの画像になっていません。");
                    noted = true;
                }
                if (stateOk && !logOk)
                {
                    WordScoreExplanation.Note("行頭文字を定義した記録がありません。");
                    noted = true;
                }
                if (noted)
                    return false;
                return Miss("行頭文字がPCアイコンの画像になっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_10_03_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;

            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return Miss(WordScoreExplanation.UnavailableText); }

                document = GetDocument(wordApp, filePath);
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                string stateFailReason;
                bool stateOk = TryExplainTask03State(document, out stateFailReason);
                bool logOk = LogReader.HasTaskEvidence(10, 3, "BulletDefineNew");

                if (stateOk && logOk)
                    return true;

                bool noted = false;
                if (!stateOk)
                {
                    WordScoreExplanation.Note(stateFailReason ?? Task03ComboFailReason);
                    noted = true;
                }
                if (stateOk && !logOk)
                {
                    WordScoreExplanation.Note("行頭文字を定義した記録がありません。");
                    noted = true;
                }
                if (noted)
                    return false;
                return Miss(Task03ComboFailReason);
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool IsTask02StateCorrect(Document document)
        {
            string[] targets = { "推測できる簡単なパスワード", "身に覚えのないリンク", "サポート切れソフトウェア" };
            foreach (string target in targets)
            {
                if (!IsPictureBulletParagraph(document, target))
                    return false;
            }
            return true;
        }

        private static bool IsPictureBulletParagraph(Document document, string target)
        {
            Range searchRange = WordFindHelper.DuplicateContent(document);
            Find find = searchRange.Find;
            WordFindHelper.ConfigureSafeFind(find, target);

            object findText = target;
            object matchCase = false;
            object matchWholeWord = false;
            object matchWildcards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object wrap = WdFindWrap.wdFindStop;
            object format = false;
            object missing = System.Reflection.Missing.Value;

            bool found = find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildcards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            bool isTargetValid = false;
            if (found && find.Found)
            {
                Paragraphs paragraphs = null;
                Paragraph para = null;
                ListFormat lf = null;
                try
                {
                    paragraphs = searchRange.Paragraphs;
                    if (paragraphs != null && paragraphs.Count >= 1)
                    {
                        para = paragraphs[1];
                        lf = para.Range.ListFormat;

                        if (lf.ListType == WdListType.wdListPictureBullet)
                        {
                            isTargetValid = true;
                        }
                        else if (lf.ListType == WdListType.wdListBullet)
                        {
                            ListTemplate lt = lf.ListTemplate;
                            if (lt != null)
                            {
                                ListLevels levels = lt.ListLevels;
                                if (levels != null && levels.Count >= 1)
                                {
                                    ListLevel level = levels[1];
                                    InlineShape pb = null;
                                    try
                                    {
                                        pb = level.PictureBullet;
                                        if (pb != null)
                                            isTargetValid = true;
                                    }
                                    catch { }
                                    finally
                                    {
                                        if (pb != null) Marshal.ReleaseComObject(pb);
                                        if (level != null) Marshal.ReleaseComObject(level);
                                    }
                                }
                                if (levels != null) Marshal.ReleaseComObject(levels);
                                Marshal.ReleaseComObject(lt);
                            }
                        }
                    }
                }
                finally
                {
                    if (lf != null) Marshal.ReleaseComObject(lf);
                    if (para != null) Marshal.ReleaseComObject(para);
                    if (paragraphs != null) Marshal.ReleaseComObject(paragraphs);
                }
            }

            Marshal.ReleaseComObject(find);
            Marshal.ReleaseComObject(searchRange);
            return isTargetValid;
        }

        /// <summary>
        /// 10-3 の状態判定。失敗時は学生向け理由を reason に入れる。
        /// 優先: フォントのみNG→実値 / コードのみNG→120未設定 / 両方NG・取得不可→セット条件。
        /// </summary>
        private static bool TryExplainTask03State(Document document, out string reason)
        {
            reason = null;
            string[] targets = { "セキュリティ対策ソフト", "使わなくなった機器" };
            bool sawAnyBullet = false;

            foreach (string target in targets)
            {
                if (!TryGetBulletLevelInfo(document, target, out string fontName, out string numFormat, out bool isBullet))
                {
                    reason = Task03ComboFailReason;
                    return false;
                }

                if (!isBullet)
                {
                    reason = Task03ComboFailReason;
                    return false;
                }

                sawAnyBullet = true;
                bool isFontMatch = !string.IsNullOrEmpty(fontName)
                    && fontName.IndexOf("Webdings", StringComparison.OrdinalIgnoreCase) >= 0;
                bool isCodeMatch = !string.IsNullOrEmpty(numFormat)
                    && (numFormat.Contains("x")
                        || numFormat.Contains("\x0078")
                        || numFormat.Contains(((char)120).ToString())
                        || numFormat.Contains(((char)0xF078).ToString())
                        || numFormat.Contains(((char)61560).ToString()));

                string knownFont = (!string.IsNullOrEmpty(fontName) && fontName != "null") ? fontName : null;

                if (!isFontMatch || !isCodeMatch)
                {
                    if (!isFontMatch && isCodeMatch)
                    {
                        reason = !string.IsNullOrEmpty(knownFont)
                            ? $"行頭文字のフォントが「{knownFont}」になっています。"
                            : "行頭文字のフォントが「Webdings」になっていません。";
                    }
                    else if (isFontMatch && !isCodeMatch)
                    {
                        reason = "行頭文字の文字コードが「120」になっていません。";
                    }
                    else
                    {
                        reason = Task03ComboFailReason;
                    }
                    return false;
                }
            }

            if (!sawAnyBullet)
            {
                reason = Task03ComboFailReason;
                return false;
            }

            return true;
        }

        private static bool TryGetBulletLevelInfo(
            Document document,
            string target,
            out string fontName,
            out string numFormat,
            out bool isBullet)
        {
            fontName = null;
            numFormat = null;
            isBullet = false;

            Range searchRange = WordFindHelper.DuplicateContent(document);
            Find find = searchRange.Find;
            WordFindHelper.ConfigureSafeFind(find, target);

            object findText = target;
            object matchCase = false;
            object matchWholeWord = false;
            object matchWildcards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object wrap = WdFindWrap.wdFindStop;
            object format = false;
            object missing = System.Reflection.Missing.Value;

            bool found = find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildcards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            Paragraphs paragraphs = null;
            Paragraph para = null;
            ListFormat lf = null;
            try
            {
                if (!found || !find.Found)
                    return false;

                paragraphs = searchRange.Paragraphs;
                if (paragraphs == null || paragraphs.Count < 1)
                    return false;

                para = paragraphs[1];
                lf = para.Range.ListFormat;
                if (lf.ListType != WdListType.wdListBullet)
                    return true; // 段落はあるが記号箇条書きではない

                isBullet = true;
                ListTemplate lt = lf.ListTemplate;
                if (lt == null)
                    return true;

                ListLevels levels = lt.ListLevels;
                if (levels != null && levels.Count >= 1)
                {
                    ListLevel level = levels[1];
                    try
                    {
                        fontName = level.Font.Name ?? "null";
                        numFormat = level.NumberFormat ?? "null";
                    }
                    catch { }
                    finally
                    {
                        if (level != null) Marshal.ReleaseComObject(level);
                    }
                }
                if (levels != null) Marshal.ReleaseComObject(levels);
                Marshal.ReleaseComObject(lt);
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (lf != null) Marshal.ReleaseComObject(lf);
                if (para != null) Marshal.ReleaseComObject(para);
                if (paragraphs != null) Marshal.ReleaseComObject(paragraphs);
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
            }
        }

        private bool CheckTask_1_10_04_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;

            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return Miss(WordScoreExplanation.UnavailableText); }

                document = GetDocument(wordApp, filePath);
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                Range searchRange = WordFindHelper.DuplicateContent(document);
                Find find = searchRange.Find;
                WordFindHelper.ConfigureSafeFind(find, "事例");

                object missing = System.Reflection.Missing.Value;
                object forward = true;
                object wrap = WdFindWrap.wdFindStop;

                bool allItalic = true;
                bool foundAny = false;

                while (find.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref forward, ref wrap, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing))
                {
                    foundAny = true;
                    Font font = searchRange.Font;
                    bool isItalic = font.Italic != 0;
                    Marshal.ReleaseComObject(font);

                    if (!isItalic)
                    {
                        allItalic = false;
                        break;
                    }
                }

                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);

                if (foundAny && allItalic)
                    return true;
                if (!foundAny)
                    return Miss("「事例」が見つかりません。");
                return Miss("「事例」が斜体になっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_10_05_Impl(string filePath)
        {
            Application wordApp = null;
            Document document = null;

            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return Miss(WordScoreExplanation.UnavailableText); }

                document = GetDocument(wordApp, filePath);
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                PageSetup ps = document.Sections[1].PageSetup;
                const float pageTolerance = 5f;
                float pageWidth = (float)ps.PageWidth;
                float pageHeight = (float)ps.PageHeight;
                bool isB5 = Math.Abs(pageWidth - 516f) <= pageTolerance && Math.Abs(pageHeight - 729f) <= pageTolerance;

                Range lastRange = WordFindHelper.DuplicateContent(document);
                lastRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                lastRange.MoveStart(WdUnits.wdParagraph, -2);
                ParagraphFormat pf = lastRange.ParagraphFormat;

                bool isLineSpacing16 = pf.LineSpacingRule == WdLineSpacing.wdLineSpaceMultiple
                    && Math.Abs((float)pf.LineSpacing - 19.2f) <= 0.1f;
                string spacingLabel = DescribeLineSpacing(pf);

                bool noted = false;
                if (!isB5)
                {
                    string sizeLabel = DescribePageSize(pageWidth, pageHeight);
                    if (!string.IsNullOrEmpty(sizeLabel))
                        WordScoreExplanation.Note($"用紙サイズが「{sizeLabel}」になっています。");
                    else
                        WordScoreExplanation.Note("用紙サイズがB5になっていません。");
                    noted = true;
                }
                if (!isLineSpacing16)
                {
                    if (!string.IsNullOrEmpty(spacingLabel))
                        WordScoreExplanation.Note($"最後の2行の行間が「{spacingLabel}」になっています。");
                    else
                        WordScoreExplanation.Note("最後の2行の行間が「1.6行」になっていません。");
                    noted = true;
                }

                Marshal.ReleaseComObject(pf);
                Marshal.ReleaseComObject(lastRange);
                Marshal.ReleaseComObject(ps);

                if (isB5 && isLineSpacing16)
                    return true;
                if (noted)
                    return false;
                return Miss("用紙サイズと行間が指定どおりになっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        /// <summary>よくある用紙サイズを短いラベルにする。</summary>
        private static string DescribePageSize(float widthPt, float heightPt)
        {
            const float tol = 8f;
            if (Math.Abs(widthPt - 595f) <= tol && Math.Abs(heightPt - 842f) <= tol)
                return "A4";
            if (Math.Abs(widthPt - 516f) <= tol && Math.Abs(heightPt - 729f) <= tol)
                return "B5";
            if (Math.Abs(widthPt - 612f) <= tol && Math.Abs(heightPt - 792f) <= tol)
                return "レター";
            if (Math.Abs(widthPt - 515f) <= tol && Math.Abs(heightPt - 728f) <= tol)
                return "B5";
            return null;
        }

        /// <summary>学生向けに行間の現状を短いラベルにする。</summary>
        private static string DescribeLineSpacing(ParagraphFormat pf)
        {
            WdLineSpacing rule;
            float spacing;
            try
            {
                rule = pf.LineSpacingRule;
                spacing = (float)pf.LineSpacing;
            }
            catch
            {
                return null;
            }

            if (rule == WdLineSpacing.wdLineSpaceMultiple)
            {
                float multiple;
                if (spacing >= 1.0f && spacing <= 3.0f)
                    multiple = spacing;
                else if (spacing >= 10f && spacing <= 48f)
                    multiple = spacing / 12f;
                else
                    return null;
                float rounded = (float)Math.Round(multiple, 2);
                string num = Math.Abs(rounded * 10f - Math.Round(rounded * 10f)) < 0.001f
                    ? rounded.ToString("0.0", CultureInfo.InvariantCulture)
                    : rounded.ToString("0.##", CultureInfo.InvariantCulture);
                return num + "行";
            }
            if (rule == WdLineSpacing.wdLineSpaceSingle)
                return "1行";
            if (rule == WdLineSpacing.wdLineSpace1pt5)
                return "1.5行";
            if (rule == WdLineSpacing.wdLineSpaceDouble)
                return "2行";
            if (rule == WdLineSpacing.wdLineSpaceExactly)
                return "固定値";
            if (rule == WdLineSpacing.wdLineSpaceAtLeast)
                return "最小値";
            return null;
        }

        private static Document GetDocument(Application wordApp, string filePath)
        {
            string fileName = System.IO.Path.GetFileName(filePath);
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
