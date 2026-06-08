using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_10
    {
        public bool CheckTask_1_10_01() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_10_01(filePath); } catch { return false; } }
        public bool CheckTask_1_10_02() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_10_02(filePath); } catch { return false; } }
        public bool CheckTask_1_10_03() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_10_03(filePath); } catch { return false; } }
        public bool CheckTask_1_10_04() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_10_04(filePath); } catch { return false; } }
        public bool CheckTask_1_10_05() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_10_05(filePath); } catch { return false; } }

        private bool CheckTask_1_10_01(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            Range searchRange = null;
            Find find = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                        || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        break;
                    }
                }
                if (document == null) return false;

                // 表記ゆれは許容しない。「ウイルス」で Find し、その直前に「コンピュータ」（6文字）があるか（＝コンピュータウイルス）を厳密に判定
                const int computerLen = 6; // 「コンピュータ」の文字数
                searchRange = document.Content;
                find = searchRange.Find;
                find.ClearFormatting();
                find.Text = "ウイルス";
                object findTextVirus = "ウイルス";
                object matchCase = false;
                object matchWholeWord = false;
                object matchWildcards = false;
                object matchSoundsLike = false;
                object matchAllWordForms = false;
                object forward = true;
                // 文書末尾で検索を打ち切る（先頭にラップさせない）
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

                while (find.Execute(ref findTextVirus, ref matchCase, ref matchWholeWord, ref matchWildcards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWith, ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl))
                {
                    // 念のため無限ループ防止のセーフティ
                    safetyCounter++;
                    if (safetyCounter > 1000)
                    {
                        allVirusPrecededByComputer = false;
                        break;
                    }

                    if (!find.Found) break;
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

                return atLeastOneReplacement && allVirusPrecededByComputer;
            }
            catch { return false; }
            finally
            {
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_10_02(string filePath)
        {
            Application wordApp = null;
            Document document = null;

            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                        || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        break;
                    }
                }
                if (document == null) return false;

                bool isState02 = IsTask02StateCorrect(document);
                bool isState03 = IsTask03StateCorrect(document);

                int requiredCount = 0;
                if (isState02) requiredCount++;
                if (isState03) requiredCount++;

                // 教材としての品質担保（手抜き防止）：状態が正しい場合、その回数分「新しい行頭文字の定義」のログが必要
                if (requiredCount > 0 && !LogReader.HasTaskEvidenceForProjectAtLeast(10, "BulletDefineNew", requiredCount))
                {
                    return false;
                }

                return isState02;
            }
            catch { return false; }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_10_03(string filePath)
        {
            Application wordApp = null;
            Document document = null;

            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                        || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        break;
                    }
                }
                if (document == null) return false;

                bool isState02 = IsTask02StateCorrect(document);
                bool isState03 = IsTask03StateCorrect(document);

                int requiredCount = 0;
                if (isState02) requiredCount++;
                if (isState03) requiredCount++;

                // 教材としての品質担保（手抜き防止）：状態が正しい場合、その回数分「新しい行頭文字の定義」のログが必要
                if (requiredCount > 0 && !LogReader.HasTaskEvidenceForProjectAtLeast(10, "BulletDefineNew", requiredCount))
                {
                    return false;
                }

                return isState03;
            }
            catch { return false; }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool IsTask02StateCorrect(Document document)
        {
            // 3つの段落テキストをそれぞれ直接検索して画像箇条書き状態を厳格チェック
            string[] targets = { "推測できる簡単なパスワード", "身に覚えのないリンク", "サポート切れソフトウェア" };
            foreach (string target in targets)
            {
                Range searchRange = document.Content;
                Find find = searchRange.Find;
                find.ClearFormatting();
                find.Text = target;

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
                            
                            // 画像箇条書きかどうかの厳密な二重チェック
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
                                            {
                                                isTargetValid = true;
                                            }
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

                if (!isTargetValid) return false;
            }
            return true;
        }

        private bool IsTask03StateCorrect(Document document)
        {
            // 箇条書きブロックの最初と最後の段落を一意に特定して、Webdings 120箇条書き状態を厳格チェック
            string[] targets = { "セキュリティ対策ソフト", "使わなくなった機器" };
            foreach (string target in targets)
            {
                Range searchRange = document.Content;
                Find find = searchRange.Find;
                find.ClearFormatting();
                find.Text = target;

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
                            
                            if (lf.ListType == WdListType.wdListBullet)
                            {
                                ListTemplate lt = lf.ListTemplate;
                                if (lt != null)
                                {
                                    ListLevels levels = lt.ListLevels;
                                    if (levels != null && levels.Count >= 1)
                                    {
                                        ListLevel level = levels[1];
                                        try
                                        {
                                            string fontName = level.Font.Name ?? "null";
                                            string numFormat = level.NumberFormat ?? "null";

                                            bool isFontMatch = fontName.IndexOf("Webdings", StringComparison.OrdinalIgnoreCase) >= 0;
                                            bool isCodeMatch = numFormat.Contains("x") 
                                                            || numFormat.Contains("\x0078") 
                                                            || numFormat.Contains(((char)120).ToString()) 
                                                            || numFormat.Contains(((char)0xF078).ToString())
                                                            || numFormat.Contains(((char)61560).ToString());

                                            if (isFontMatch && isCodeMatch)
                                            {
                                                isTargetValid = true;
                                            }
                                        }
                                        catch { }
                                        finally
                                        {
                                            if (level != null) Marshal.ReleaseComObject(level);
                                        }
                                    }
                                    if (levels != null) Marshal.ReleaseComObject(levels);
                                    Marshal.ReleaseComObject(lt);
                                }
                            }
                        }
                    }
                    catch { }
                    finally
                    {
                        if (lf != null) Marshal.ReleaseComObject(lf);
                        if (para != null) Marshal.ReleaseComObject(para);
                        if (paragraphs != null) Marshal.ReleaseComObject(paragraphs);
                    }
                }

                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);

                if (!isTargetValid) return false;
            }
            return true;
        }

        private bool CheckTask_1_10_04(string filePath)
        {
            Application wordApp = null;
            Document document = null;

            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                        || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        break;
                    }
                }
                if (document == null) return false;

                Range searchRange = document.Content;
                Find find = searchRange.Find;
                find.ClearFormatting();
                find.Text = "事例";

                object missing = System.Reflection.Missing.Value;
                object forward = true;
                // 文書末尾に到達したらループを終了する
                object wrap = WdFindWrap.wdFindStop;

                bool allItalic = true;
                bool foundAny = false;

                // 文書内のすべての「事例」をスキャン
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

                return foundAny && allItalic;
            }
            catch { return false; }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_10_05(string filePath)
        {
            Application wordApp = null;
            Document document = null;

            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                        || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        break;
                    }
                }
                if (document == null) return false;

                // 1. 用紙サイズ (B5) のチェック
                PageSetup ps = document.Sections[1].PageSetup;
                const float pageTolerance = 5f;
                bool isB5 = Math.Abs((float)ps.PageWidth - 516f) <= pageTolerance && Math.Abs((float)ps.PageHeight - 729f) <= pageTolerance;
                
                // 2. 最後の2段落の行間 (1.6) のチェック
                Range lastRange = document.Content;
                lastRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                lastRange.MoveStart(WdUnits.wdParagraph, -2);
                ParagraphFormat pf = lastRange.ParagraphFormat;
                
                // 行間 1.6 倍（19.2pt）の厳格なチェック（誤差 ±0.1pt のみ許容）
                bool isLineSpacing16 = pf.LineSpacingRule == WdLineSpacing.wdLineSpaceMultiple && Math.Abs((float)pf.LineSpacing - 19.2f) <= 0.1f;
                
                Marshal.ReleaseComObject(pf);
                Marshal.ReleaseComObject(lastRange);
                Marshal.ReleaseComObject(ps);
                
                return isB5 && isLineSpacing16;
            }
            catch { return false; }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
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

