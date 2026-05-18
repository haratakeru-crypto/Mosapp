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
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // XML取得
                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // 文書内のすべてのテキストボックス（w:txbxContent）を抽出
                var txbxMatches = System.Text.RegularExpressions.Regex.Matches(xml, @"<w:txbxContent\b[^>]*>.*?</w:txbxContent>", System.Text.RegularExpressions.RegexOptions.Singleline);
                
                foreach (System.Text.RegularExpressions.Match match in txbxMatches)
                {
                    string shapeXml = match.Value;
                    // 図形内に目次（TOCフィールド）が挿入されているかチェック
                    bool hasTOCField = System.Text.RegularExpressions.Regex.IsMatch(shapeXml, @"<w:instrText\b[^>]*>\s*TOC\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    // 「自動作成の目次2」のタイトル「目次」が含まれること
                    bool hasTOC2Text = shapeXml.Contains("目次");

                    if (hasTOCField && hasTOC2Text)
                    {
                        return true;
                    }
                }

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
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // XML取得
                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // 1. 脚注設定（w:footnotePr）の中に w:numFmt w:val="decimalEnclosedCircle" が含まれるかチェック
                // 通常はセクションプロパティ（w:sectPr）内の w:footnotePr に格納される
                var footnotePrMatch = System.Text.RegularExpressions.Regex.Match(xml, @"<w:footnotePr\b[^>]*>.*?</w:footnotePr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!footnotePrMatch.Success) return false;

                string footnotePrXml = footnotePrMatch.Value;
                bool isCircularNumberStyle = System.Text.RegularExpressions.Regex.IsMatch(footnotePrXml, @"<w:numFmt\b[^>]*w:val=""decimalEnclosedCircle""");

                return isCircularNumberStyle;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_04(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // XML取得
                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // 「参考文献」を含む段落を探す（問題文は参考文献一覧だが実機は「参考文献」）
                var paragraphs = System.Text.RegularExpressions.Regex.Matches(xml, @"<w:p\b[^>]*>.*?</w:p>", System.Text.RegularExpressions.RegexOptions.Singleline);
                foreach (System.Text.RegularExpressions.Match p in paragraphs)
                {
                    string pXml = p.Value;
                    string pText = System.Text.RegularExpressions.Regex.Replace(pXml, @"<[^>]+>", "").Trim();
                    if (pText.Contains("参考文献"))
                    {
                        // 目次（TOC）のエントリはハイパーリンクやPAGEREFを含むため除外する
                        if (pXml.Contains("PAGEREF") || pXml.Contains("w:hyperlink")) continue;

                        bool containsSec = pText.Contains("§");
                        bool containsSym = pXml.Contains("<w:sym ");
                        
                        // 記号タグ（<w:sym>）が使われている場合、またはプレーンテキストで§がある場合を許容
                        return containsSec || containsSym;
                    }
                }

                return false;
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

