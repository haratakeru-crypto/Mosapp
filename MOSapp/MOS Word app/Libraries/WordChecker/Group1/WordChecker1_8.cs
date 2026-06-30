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
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // 1. 脚注定義 (/word/footnotes.xml) から対象テキストを持つ脚注IDを特定する
                var footnotesPartMatch = System.Text.RegularExpressions.Regex.Match(xml, @"<pkg:part pkg:name=""/word/footnotes.xml""[^>]*>.*?</pkg:part>", System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!footnotesPartMatch.Success) return false;
                string footnotesXml = footnotesPartMatch.Value;

                var footnoteMatches = System.Text.RegularExpressions.Regex.Matches(footnotesXml, @"<w:footnote\b[^>]*w:id=""([^""]+)""[^>]*>.*?</w:footnote>", System.Text.RegularExpressions.RegexOptions.Singleline);
                string targetFootnoteId = null;
                foreach (System.Text.RegularExpressions.Match fn in footnoteMatches)
                {
                    string fnXml = fn.Value;
                    string fnText = System.Text.RegularExpressions.Regex.Replace(fnXml, @"<[^>]+>", "");
                    if (fnText.Contains("電子機器に悪影響を与えるプログラム"))
                    {
                        targetFootnoteId = fn.Groups[1].Value;
                        break;
                    }
                }

                if (string.IsNullOrEmpty(targetFootnoteId)) return false;

                // 2. 本文 (/word/document.xml) の中で「マルウェア」が含まれる最初の段落を特定する
                var documentPartMatch = System.Text.RegularExpressions.Regex.Match(xml, @"<pkg:part pkg:name=""/word/document.xml""[^>]*>.*?</pkg:part>", System.Text.RegularExpressions.RegexOptions.Singleline);
                string bodyXml = documentPartMatch.Success ? documentPartMatch.Value : xml;

                var paragraphs = System.Text.RegularExpressions.Regex.Matches(bodyXml, @"<w:p\b[^>]*>.*?</w:p>", System.Text.RegularExpressions.RegexOptions.Singleline);
                
                foreach (System.Text.RegularExpressions.Match p in paragraphs)
                {
                    string pXml = p.Value;
                    string pTextOnly = System.Text.RegularExpressions.Regex.Replace(pXml, @"<[^>]+>", "");
                    
                    if (pTextOnly.Contains("マルウェア"))
                    {
                        // 最初の「マルウェア」段落を発見！
                        // 脚注の参照タグを一旦 [FOOTNOTE_ID] に置換する
                        string replacedXml = System.Text.RegularExpressions.Regex.Replace(pXml, @"<w:footnoteReference\b[^>]*w:id=""([^""]+)""[^>]*>", "[FOOTNOTE_$1]");
                        // 他のXMLタグをすべて削除してプレーンテキスト化
                        string cleanText = System.Text.RegularExpressions.Regex.Replace(replacedXml, @"<[^>]+>", "");
                        
                        // "マルウェア" の直後に [FOOTNOTE_targetFootnoteId] が配置されているかを厳密に検証
                        string expectedToken = $"マルウェア[FOOTNOTE_{targetFootnoteId}]";
                        return cleanText.Contains(expectedToken);
                    }
                }

                return false;
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
                var footnotePrMatch = System.Text.RegularExpressions.Regex.Match(xml, @"<w:footnotePr\b[^>]*>.*?</w:footnotePr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!footnotePrMatch.Success) return false;

                string footnotePrXml = footnotePrMatch.Value;
                if (!System.Text.RegularExpressions.Regex.IsMatch(footnotePrXml, @"<w:numFmt\b[^>]*w:val=""decimalEnclosedCircle"""))
                    return false;

                // 2. 脚注参照は 8-2 で挿入した 1 個のみ（「挿入」で余分な参照が増えた場合は不合格）
                int refCount = WordFindHelper.CountFootnoteReferencesInXml(xml);
                if (refCount != 1)
                    return false;

                // 3. 脚注定義も 1 個（separator 等を除く）
                return WordFindHelper.CountFootnoteDefinitionsInXml(xml) == 1;
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
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // SmartArt関連のパッケージパートを抽出
                var partMatches = System.Text.RegularExpressions.Regex.Matches(xml, @"<pkg:part pkg:name=""/word/diagrams/([^""]+)""[^>]*>.*?</pkg:part>", System.Text.RegularExpressions.RegexOptions.Singleline);
                
                bool hasVennLayout = false;
                bool hasRequiredTexts = false;

                foreach (System.Text.RegularExpressions.Match match in partMatches)
                {
                    string partXml = match.Value;
                    string partName = match.Groups[1].Value;

                    if (partName.Contains("layout"))
                    {
                        if (partXml.Contains(@"urn:microsoft.com/office/officeart/2005/8/layout/venn1"))
                        {
                            hasVennLayout = true;
                        }
                    }
                    else if (partName.Contains("data"))
                    {
                        if (partXml.Contains("機密性") && partXml.Contains("完全性") && partXml.Contains("可用性"))
                        {
                            hasRequiredTexts = true;
                        }
                    }
                }

                return hasVennLayout && hasRequiredTexts;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_06(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // SmartArt関連のパッケージパートを抽出
                var partMatches = System.Text.RegularExpressions.Regex.Matches(xml, @"<pkg:part pkg:name=""/word/diagrams/([^""]+)""[^>]*>.*?</pkg:part>", System.Text.RegularExpressions.RegexOptions.Singleline);
                
                bool hasColorfulColors = false;

                foreach (System.Text.RegularExpressions.Match match in partMatches)
                {
                    string partXml = match.Value;
                    string partName = match.Groups[1].Value;

                    if (partName.Contains("colors"))
                    {
                        if (partXml.Contains(@"urn:microsoft.com/office/officeart/2005/8/colors/colorful2"))
                        {
                            hasColorfulColors = true;
                            break;
                        }
                    }
                }

                return hasColorfulColors;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_8_07(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // "TOP" を含む図形全体のブロック <mc:AlternateContent> を抽出する
                var alternates = System.Text.RegularExpressions.Regex.Matches(xml, @"<mc:AlternateContent\b[^>]*>.*?</mc:AlternateContent>", System.Text.RegularExpressions.RegexOptions.Singleline);
                bool hasTopHyperlink = false;

                foreach (System.Text.RegularExpressions.Match alt in alternates)
                {
                    string altXml = alt.Value;
                    if (altXml.Contains("TOP"))
                    {
                        // 図形（枠線）自体にハイパーリンクが設定されている場合、
                        // 互換用の VML タグ <v:shape ... href="#_top" が生成されます。
                        // これを検証することで、「文字ではなく図形に」「文頭への」リンクが貼られたかを厳密に判定します。
                        if (altXml.Contains("href=\"#_top\""))
                        {
                            hasTopHyperlink = true;
                            break;
                        }
                    }
                }

                return hasTopHyperlink;
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

