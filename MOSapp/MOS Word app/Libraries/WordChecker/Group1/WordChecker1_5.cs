using System;
using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_5
    {
        public bool CheckTask_1_5_01() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_01(filePath); } catch { return false; } }
        public bool CheckTask_1_5_02() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_02(filePath); } catch { return false; } }
        public bool CheckTask_1_5_03() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_03(filePath); } catch { return false; } }
        public bool CheckTask_1_5_04() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_04(filePath); } catch { return false; } }
        public bool CheckTask_1_5_05() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_05(filePath); } catch { return false; } }
        public bool CheckTask_1_5_06() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_06(filePath); } catch { return false; } }
        public bool CheckTask_1_5_07() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_07(filePath); } catch { return false; } }
        public bool CheckTask_1_5_08() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_5_08(filePath); } catch { return false; } }

        private bool CheckTask_1_5_01(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                Range searchRange = WordFindHelper.DuplicateContent(document); Find find = searchRange.Find; WordFindHelper.ConfigureSafeFind(find, "5月21日より5日間の"); find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                Range paraRange = searchRange.Paragraphs[1].Range;
                int paraStart = paraRange.Start;
                int paraEnd = paraRange.End;

                // 「5月21日より5日間の...」段落の画像の折り返しが「上下」(wdWrapTopBottom) か
                bool result = IsWrapTopBottomInParagraph(document, paraStart, paraEnd);

                Marshal.ReleaseComObject(paraRange);
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);
                // 5-1: 現在「上下」、または当該タスクの証跡で WrapTopBottom 操作あり（5-2 後はログで判定）
                bool logOk = LogReader.HasTaskEvidence(5, 1, "WrapTopBottom");
                return result || logOk;

            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>指定段落内にアンカーされた浮動 Shape の折り返しが指定タイプか。</summary>
        private static bool IsWrapTypeInParagraph(Document document, int paraStart, int paraEnd, WdWrapType wrapType)
        {
            if (document == null) return false;
            Microsoft.Office.Interop.Word.Shapes shapes = null;
            try
            {
                shapes = document.Shapes;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    Microsoft.Office.Interop.Word.Shape sh = null;
                    try
                    {
                        sh = shapes[i];
                        int anchor = sh.Anchor != null ? sh.Anchor.Start : -1;
                        if (anchor < paraStart || anchor > paraEnd) continue;
                        WrapFormat wf = sh.WrapFormat;
                        try
                        {
                            if (wf != null && wf.Type == wrapType)
                                return true;
                        }
                        finally { if (wf != null) Marshal.ReleaseComObject(wf); }
                    }
                    catch { }
                    finally { if (sh != null) Marshal.ReleaseComObject(sh); }
                }
            }
            catch { }
            finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
            return false;
        }

        /// <summary>指定段落内にアンカーされた画像の折り返しが「上下」(wdWrapTopBottom) か。</summary>
        private static bool IsWrapTopBottomInParagraph(Document document, int paraStart, int paraEnd)
        {
            return IsWrapTypeInParagraph(document, paraStart, paraEnd, WdWrapType.wdWrapTopBottom);
        }

        /// <summary>指定段落内にアンカーされた画像の折り返しが「狭く」(wdWrapTight) か。</summary>
        private static bool IsWrapTightInParagraph(Document document, int paraStart, int paraEnd)
        {
            return IsWrapTypeInParagraph(document, paraStart, paraEnd, WdWrapType.wdWrapTight);
        }

        private bool CheckTask_1_5_02(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                Range searchRange = WordFindHelper.DuplicateContent(document); Find find = searchRange.Find; WordFindHelper.ConfigureSafeFind(find, "5月21日より5日間の"); find.Execute();
                if (!find.Found) { Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange); return false; }
                Range paraRange = searchRange.Paragraphs[1].Range;
                int paraStart = paraRange.Start;
                int paraEnd = paraRange.End;

                bool result = IsWrapTightInParagraph(document, paraStart, paraEnd);

                Marshal.ReleaseComObject(paraRange);
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);
                // 5-2: 段落内が「狭く」かつ WrapTight 操作の証跡あり（5-3 以降も折り返しは維持される想定）
                bool logOk = LogReader.HasTaskEvidence(5, 2, "WrapTight");
                return result && logOk;

            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_03(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // XML解析
                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // 5-3: 「5月21日より5日間の」段落付近の画像に限定
                string imageXml = GetImageXmlNearText(xml, "5月21日より5日間の");
                
                // その画像に「水彩：スポンジ」が適用されているか
                return imageXml.Contains("artisticWatercolorSponge");
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_04(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // XML解析
                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // 5-4: 「5月21日より5日間の」段落付近の画像に限定
                string imageXml = GetImageXmlNearText(xml, "5月21日より5日間の");

                // ぼかし25ポイント (317500 EMU)
                return imageXml.Contains("softEdge") && imageXml.Contains("317500");
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_05(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // XML解析
                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // 5-5: 「TOEICテスト対策セミナー」付近の画像（タイトル画像）に限定
                string imageXml = GetImageXmlNearText(xml, "TOEICテスト対策セミナー");

                // 面取り ハードエッジ
                return imageXml.Contains("bevelT") && imageXml.Contains("hardEdge");
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_06(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 「TOEICテスト対策セミナー･･･」のタイトルの右側の画像の代替テキストが「セミナー案内」か
                bool result = false;
                foreach (Microsoft.Office.Interop.Word.Shape sh in document.Shapes)
                {
                    try
                    {
                        string alt = sh.AlternativeText ?? "";
                        if (alt.Trim().Contains("セミナー案内")) { result = true; break; }
                    }
                    catch { }
                }
                return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_07(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // XML解析による判定
                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // 装飾用のXML形式: <adec:decorative xmlns:adec=".../2017/decorative" val="1"/>
                bool isDecorative = xml.Contains("2017/decorative") && xml.Contains("val=\"1\"");

                return isDecorative;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_5_08(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return false; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // XML解析
                string xml = document.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return false;

                // 5-8: 文書の「最後」の画像を特定する
                var matches = System.Text.RegularExpressions.Regex.Matches(xml, @"<w:drawing>.*?</w:drawing>", System.Text.RegularExpressions.RegexOptions.Singleline);
                if (matches.Count == 0) return false;
                
                // 最後の画像ブロックを取得
                string lastImageXml = matches[matches.Count - 1].Value;

                // 領域保持（foregroundMark）の座標 y1 を解析
                var yMatches = System.Text.RegularExpressions.Regex.Matches(lastImageXml, @"y1=""(\d+)""");
                bool hasTopMark = false;
                bool hasBottomMark = false;
                foreach (System.Text.RegularExpressions.Match m in yMatches)
                {
                    if (int.TryParse(m.Groups[1].Value, out int y))
                    {
                        if (y < 40000) hasTopMark = true;
                        if (y > 60000) hasBottomMark = true;
                    }
                }

                // 上下両方のエリアにマークがあり、背景削除が確定されていれば合格
                return hasTopMark && hasBottomMark;



            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }



        /// <summary>
        /// 指定したテキストを含む段落内、またはその直前にある画像のXMLブロックを抽出します。
        /// </summary>
        private string GetImageXmlNearText(string fullXml, string targetText)
        {
            var paragraphs = System.Text.RegularExpressions.Regex.Matches(fullXml, @"<w:p\b[^>]*>.*?</w:p>", System.Text.RegularExpressions.RegexOptions.Singleline);
            
            for (int i = 0; i < paragraphs.Count; i++)
            {
                string pXml = paragraphs[i].Value;
                string pText = System.Text.RegularExpressions.Regex.Replace(pXml, @"<[^>]+>", "");
                
                if (pText.Contains(targetText))
                {
                    // 1. まず同じ段落内で画像を探す
                    var drawingMatch = System.Text.RegularExpressions.Regex.Match(pXml, @"<(w:drawing|v:shape).*?>.*?</\1>", System.Text.RegularExpressions.RegexOptions.Singleline);
                    if (drawingMatch.Success) return drawingMatch.Value;

                    // 2. なければ直前の段落を探す（タイトル画像などは直前の段落に置かれることが多いため）
                    if (i > 0)
                    {
                        string prevPXml = paragraphs[i - 1].Value;
                        var prevDrawingMatch = System.Text.RegularExpressions.Regex.Match(prevPXml, @"<(w:drawing|v:shape).*?>.*?</\1>", System.Text.RegularExpressions.RegexOptions.Singleline);
                        if (prevDrawingMatch.Success) return prevDrawingMatch.Value;
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine($"[GetImageXmlNearText] テキスト '{targetText}' 付近（直前含む）に画像が見つかりませんでした。");
            return "";
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

