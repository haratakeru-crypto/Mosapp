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
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 表記ゆれは許容しない。「ウイルス」で Find し、その直前に「コンピュータ」（6文字）があるか（＝コンピュータウイルス）を厳密に判定
                const int computerLen = 6; // 「コンピュータ」の文字数
                Range searchRange = document.Content;
                Find find = searchRange.Find;
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
                    finally { Marshal.ReleaseComObject(virusRange); }
                    searchRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                    searchRange.Move(WdUnits.wdCharacter, 1);
                }
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);
                return atLeastOneReplacement && allVirusPrecededByComputer;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_10_02(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // リスト段落の行頭にインラインシェイプ（画像）があるか（PCアイコン等）
                foreach (Paragraph para in document.Paragraphs)
                {
                    try
                    {
                        if (para.Range.ListFormat.ListType != WdListType.wdListNoNumbering && para.Range.InlineShapes.Count >= 1)
                            return true;
                    }
                    finally { Marshal.ReleaseComObject(para); }
                }
                return false;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_10_03(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // リスト段落の行頭が Webdings フォントで文字コード 120 (Chr(120)) か
                foreach (Paragraph para in document.Paragraphs)
                {
                    try
                    {
                        if (para.Range.ListFormat.ListType == WdListType.wdListNoNumbering) continue;
                        if (para.Range.Characters.Count < 1) continue;
                        Range rng = para.Range.Characters[1];
                        string fontName = rng.Font?.Name ?? "";
                        string ch = rng.Text ?? "";
                        Marshal.ReleaseComObject(rng);
                        if (fontName.IndexOf("Webdings", StringComparison.OrdinalIgnoreCase) >= 0 && ch.Length > 0 && (int)ch[0] == 120)
                            return true;
                    }
                    finally { Marshal.ReleaseComObject(para); }
                }
                return false;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_10_04(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "事例"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                // 「事例」が斜体になっているかチェック
                Range foundRange = searchRange;
                Font font = foundRange.Font;
                bool result = font.Italic != 0;
                Marshal.ReleaseComObject(font); Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_10_05(string filePath)
        {
            // 行間 1.6 は VSTO ログで記録されているか優先して判定（ログがあれば行間は正解とみなす）
            if (LogReader.HasCommandExecuted("LineSpacing16"))
            {
                bool isB5 = CheckB5PageSize(filePath);
                return isB5;
            }
            // フォールバック: コードで B5 と行間 1.6 を判定
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                PageSetup ps = document.Sections[1].PageSetup;
                // B5サイズ（ポイント）に許容誤差を設ける
                const float pageTolerance = 5f;
                bool isB5 = Math.Abs((float)ps.PageWidth - 516f) <= pageTolerance && Math.Abs((float)ps.PageHeight - 729f) <= pageTolerance;
                Range lastRange = document.Content;
                lastRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                lastRange.MoveStart(WdUnits.wdParagraph, -2);
                ParagraphFormat pf = lastRange.ParagraphFormat;
                // 行間 1.6 倍付近の許容範囲を広げる（17〜22pt 程度）
                bool isLineSpacing16 = pf.LineSpacingRule == WdLineSpacing.wdLineSpaceMultiple && (float)pf.LineSpacing >= 17f && (float)pf.LineSpacing <= 22f;
                Marshal.ReleaseComObject(pf); Marshal.ReleaseComObject(lastRange); Marshal.ReleaseComObject(ps);
                return isB5 && isLineSpacing16;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckB5PageSize(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        break;
                    }
                }
                if (document == null) return false;
                PageSetup ps = document.Sections[1].PageSetup;
                const float pageTolerance = 5f;
                bool isB5 = Math.Abs((float)ps.PageWidth - 516f) <= pageTolerance && Math.Abs((float)ps.PageHeight - 729f) <= pageTolerance;
                Marshal.ReleaseComObject(ps);
                return isB5;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
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

