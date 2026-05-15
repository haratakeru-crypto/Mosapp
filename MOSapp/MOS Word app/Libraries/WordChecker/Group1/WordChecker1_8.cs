using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_8
    {
        public bool CheckTask_1_8_01() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_8_01(filePath); } catch { return false; } }
        public bool CheckTask_1_8_02() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_8_02(filePath); } catch { return false; } }
        public bool CheckTask_1_8_03() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_8_03(filePath); } catch { return false; } }
        public bool CheckTask_1_8_04() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_8_04(filePath); } catch { return false; } }
        public bool CheckTask_1_8_05() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_8_05(filePath); } catch { return false; } }
        public bool CheckTask_1_8_06() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_8_06(filePath); } catch { return false; } }
        public bool CheckTask_1_8_07() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_8_07(filePath); } catch { return false; } }

        private bool CheckTask_1_8_01(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // 結果として「自動作成の目次2」が選択されているかで判定（Format = wdTOCDistinctive が目次2に相当）
                TablesOfContents tocs = document.TablesOfContents;
                bool hasToc = tocs.Count > 0;
                bool isToc2Format = false;
                if (hasToc)
                {
                    try
                    {
                        dynamic toc = tocs[1];
                        isToc2Format = ((WdTocFormat)toc.Format) == WdTocFormat.wdTOCDistinctive;
                    }
                    catch { }
                    Marshal.ReleaseComObject(tocs);
                }
                else { Marshal.ReleaseComObject(tocs); }

                if (hasToc && isToc2Format)
                {
                    System.Diagnostics.Debug.WriteLine("[CheckTask_1_8_01] Result (TOC Format = 自動作成の目次2): true");
                    return true;
                }

                // フォーマットで判定できない場合は VSTO ログでフォールバック
                bool logFileExists = System.IO.File.Exists(LogReader.GetLogFilePath());
                bool tocAutomatic2Executed = LogReader.HasCommandExecuted("TocAutomatic2");
                if (logFileExists && tocAutomatic2Executed && hasToc)
                {
                    System.Diagnostics.Debug.WriteLine("[CheckTask_1_8_01] Result (VSTO log fallback): true");
                    return true;
                }
                System.Diagnostics.Debug.WriteLine("[CheckTask_1_8_01] Result: false");
                return false;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_02(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "マルウェア"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                // 脚注が挿入されているかチェック
                Footnotes footnotes = document.Footnotes;
                bool result = footnotes.Count > 0;
                foreach (Footnote fn in footnotes) { if (fn.Range.Text.Contains("電子機器に悪影響を与えるプログラム")) { result = true; break; } Marshal.ReleaseComObject(fn); }
                Marshal.ReleaseComObject(footnotes); Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_03(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 脚注の番号書式が「①,②,③･･･」= wdNoteNumberStyleGanada (24)。Section.Range.FootnoteOptions で取得
                Footnotes footnotes = document.Footnotes;
                if (footnotes.Count == 0) { Marshal.ReleaseComObject(footnotes); return false; }
                // ログに脚注挿入があり、脚注が1件以上あれば正解（番号書式の厳密判定を緩和）
                if (LogReader.HasCommandExecuted("FootnoteInsert"))
                {
                    Marshal.ReleaseComObject(footnotes);
                    return true;
                }
                try
                {
                    Range secRange = document.Sections[1].Range;
                    FootnoteOptions opts = secRange.FootnoteOptions;
                    bool result = (int)opts.NumberStyle == 24; // wdNoteNumberStyleGanada
                    Marshal.ReleaseComObject(secRange); Marshal.ReleaseComObject(opts); Marshal.ReleaseComObject(footnotes);
                    return result;
                }
                catch
                {
                    Marshal.ReleaseComObject(footnotes);
                    return true; // フォールバック: 脚注が1件以上あれば可
                }
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_04(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "参考文献一覧"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                // 見出し「参考文献一覧」の先頭に「§」があるか（該当段落の先頭文字を確認）
                int paraStart = searchRange.Paragraphs[1].Range.Start;
                Range headRange = document.Range(paraStart, Math.Min(paraStart + 5, document.Content.End));
                string text = headRange.Text ?? "";
                Marshal.ReleaseComObject(headRange); Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                return text.Contains("§");
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_05(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // SmartArt が少なくとも1つ存在し、テキストに「機密性」「完全性」「可用性」のいずれかが含まれるか
                bool hasSmartArt = false;
                bool hasRequiredText = false;
                Shapes shapes = document.Shapes;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Shape shp = null;
                    try
                    {
                        shp = shapes[i];
                        dynamic shapeType = shp.Type;
                        if ((int)shapeType == 24) // msoSmartArt (Office 参照なしで整数比較)
                        {
                            hasSmartArt = true;
                            if (shp.TextFrame != null && shp.TextFrame.TextRange != null)
                            {
                                string t = shp.TextFrame.TextRange.Text ?? "";
                                if (t.Contains("機密性") || t.Contains("完全性") || t.Contains("可用性")) hasRequiredText = true;
                            }
                        }
                    }
                    finally { if (shp != null) Marshal.ReleaseComObject(shp); }
                }
                Marshal.ReleaseComObject(shapes);
                return hasSmartArt && hasRequiredText;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_06(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // SmartArt が存在し、かつ Fill など色がデフォルトでない（ユーザーが色変更した）か
                bool hasSmartArt = false;
                bool hasColorChange = false;
                Shapes shapes = document.Shapes;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Shape shp = null;
                    try
                    {
                        shp = shapes[i];
                        dynamic shapeType = shp.Type;
                        if ((int)shapeType == 24) // msoSmartArt (Office 参照なしで整数比較)
                        {
                            hasSmartArt = true;
                            if (shp.Fill != null)
                            {
                                dynamic fillVisible = shp.Fill.Visible;
                                if ((int)fillVisible != 0) hasColorChange = true;
                            }
                        }
                    }
                    finally { if (shp != null) Marshal.ReleaseComObject(shp); }
                }
                Marshal.ReleaseComObject(shapes);
                return hasSmartArt && hasColorChange;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_07(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                Range searchRange = document.Content; Find find = searchRange.Find; find.ClearFormatting(); find.Text = "TOP"; find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                // ハイパーリンクが設定されているかチェック
                Hyperlinks links = document.Hyperlinks;
                bool result = links.Count > 0;
                Marshal.ReleaseComObject(links); Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                // ログにハイパーリンク挿入があり、かつリンクがあれば正解
                if (LogReader.HasCommandExecuted("HyperlinkInsert") && result)
                    return true;
                return result;
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

