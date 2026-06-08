using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_6
    {
        public bool CheckTask_1_6_01() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_01(filePath); } catch { return false; } }
        public bool CheckTask_1_6_02() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_02(filePath); } catch { return false; } }
        public bool CheckTask_1_6_03() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_03(filePath); } catch { return false; } }
        public bool CheckTask_1_6_04() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_04(filePath); } catch { return false; } }
        public bool CheckTask_1_6_05() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_05(filePath); } catch { return false; } }
        public bool CheckTask_1_6_06() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_06(filePath); } catch { return false; } }
        public bool CheckTask_1_6_07() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_6_07(filePath); } catch { return false; } }

        private bool CheckTask_1_6_01(string filePath)
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

                // 1. ターゲット特定: 「月別催事内容」の直後にある表を抽出
                string tableXml = GetTableXmlNearText(xml, "月別催事内容");
                if (string.IsNullOrEmpty(tableXml)) return false;

                // 2. 構造チェック
                int rowCount = System.Text.RegularExpressions.Regex.Matches(tableXml, "<w:tr[ >]").Count;
                var firstRowMatch = System.Text.RegularExpressions.Regex.Match(tableXml, @"<w:tr[ >].*?</w:tr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                int colCount = firstRowMatch.Success ? System.Text.RegularExpressions.Regex.Matches(firstRowMatch.Value, "<w:tc[ >]").Count : 0;

                // 「文字列の幅に合わせる」設定の確認
                // 表全体の幅が auto であり、かつセルの幅設定も auto であることを確認
                var tblPrMatch = System.Text.RegularExpressions.Regex.Match(tableXml, @"<w:tblPr\b[^>]*>.*?</w:tblPr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                string tblPrXml = tblPrMatch.Success ? tblPrMatch.Value : "";
                var tcPrMatch = System.Text.RegularExpressions.Regex.Match(tableXml, @"<w:tcPr\b[^>]*>.*?</w:tcPr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                string tcPrXml = tcPrMatch.Success ? tcPrMatch.Value : "";

                bool isAutoFit = tblPrXml.Contains(@"w:type=""auto""") && tcPrXml.Contains(@"w:type=""auto""");




                // 3. 分割状態の確認（6-2の影響を考慮）

                int table1EndIndex = xml.IndexOf(tableXml) + tableXml.Length;
                string followingXml = xml.Substring(table1EndIndex);
                var nextTableMatch = System.Text.RegularExpressions.Regex.Match(followingXml, @"<w:tbl\b[^>]*>.*?</w:tbl>", System.Text.RegularExpressions.RegexOptions.Singleline);
                bool hasSplitTable = nextTableMatch.Success && nextTableMatch.Index < 1000;

                // 行数判定: 「そのまま7行」または「分割されて1つ目が4行（かつ直後に表あり）」
                bool rowCountOk = (rowCount == 7) || (rowCount == 4 && hasSplitTable);


                // 4. 内容チェック: 分割されている場合は2つ目の表も対象に含める
                string combinedContentXml = tableXml + (hasSplitTable ? nextTableMatch.Value : "");
                bool contentOk = combinedContentXml.Contains("実施月") && combinedContentXml.Contains("富士の水だより");

                // 5. ログチェック
                bool logOk = LogReader.HasTaskEvidence(6, 1, "TableConvertTextToTable");

                return logOk && rowCountOk && colCount == 2 && contentOk && isAutoFit;

            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_02(string filePath)
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

                // 1. ターゲット特定: 見出し直後の最初の表を抽出
                string table1Xml = GetTableXmlNearText(xml, "月別催事内容");
                if (string.IsNullOrEmpty(table1Xml)) return false;

                // 2. 分割後の行数チェック
                // 5行目から分割 = 1つ目の表は 4行 になっているはず
                int rowCount1 = System.Text.RegularExpressions.Regex.Matches(table1Xml, "<w:tr[ >]").Count;

                // 3. 分割された2つ目の表の存在確認
                // table1Xml の終了位置のすぐ後に別の <w:tbl> があるか
                int table1EndIndex = xml.IndexOf(table1Xml) + table1Xml.Length;
                string followingXml = xml.Substring(table1EndIndex);
                var nextTableMatch = System.Text.RegularExpressions.Regex.Match(followingXml, @"<w:tbl\b[^>]*>.*?</w:tbl>", System.Text.RegularExpressions.RegexOptions.Singleline);

                // 直後（1000文字以内）に次の表が始まっていれば分割されているとみなす
                bool hasSplitTable = nextTableMatch.Success && nextTableMatch.Index < 1000;

                // 4. 判定（ログ記録が難しいため、構造のみで判定）
                // 1つ目の表が4行かつ、その直後に次の表が存在すれば合格
                return rowCount1 == 4 && hasSplitTable;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }


        private bool CheckTask_1_6_03(string filePath)
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

                // 1. ターゲット特定: 文書内で最初の表（「（1）4月の同時キャンペーン」配下の表）
                //    タスク仕様の「1つ目の表」= 文書全体で最初に出現する <w:tbl>
                var firstTableMatch = System.Text.RegularExpressions.Regex.Match(
                    xml, @"<w:tbl\b[^>]*>.*?</w:tbl>", System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!firstTableMatch.Success) return false;
                string tableXml = firstTableMatch.Value;

                // 2. 全行の <w:tr> を抽出
                var allRows = System.Text.RegularExpressions.Regex.Matches(
                    tableXml, @"<w:tr[ >].*?</w:tr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                if (allRows.Count < 2) return false;

                // 3. 1行目はちょうど 4 セル
                //    1行2列目を3列に分割 → 1行目セル数 = 元1セル + 分割3セル = 4
                //    元: [c1] [c2]  → 分割後: [c1] [c2a] [c2b] [c2c]
                int firstRowCellCount = System.Text.RegularExpressions.Regex.Matches(allRows[0].Value, "<w:tc[ >]").Count;
                bool firstRowOk = firstRowCellCount == 4;

                // 4. 2行目以降は元の 2 セルのまま
                //    （3分割でなく4分割した場合や、別の行を分割した場合の偽陽性を排除）
                bool otherRowsOk = true;
                for (int i = 1; i < allRows.Count; i++)
                {
                    int cells = System.Text.RegularExpressions.Regex.Matches(allRows[i].Value, "<w:tc[ >]").Count;
                    if (cells != 2) { otherRowsOk = false; break; }
                }

                // 5. 判定: TableSplitCells は idMso でフックできないため、XMLの最終状態のみで判定
                //          1行目=4セル かつ 他行=2セル のときのみ ○
                return firstRowOk && otherRowsOk;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_04(string filePath)
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

                // 1. ターゲット特定: 「北海道のうまいもの市場売上」の直後にある表を抽出
                //    （「2.」は段落番号で文字列ではないため、本文テキストで検索）
                string tableXml = GetTableXmlNearText(xml, "北海道のうまいもの市場売上");
                if (string.IsNullOrEmpty(tableXml)) return false;

                // 2. 1行目を抽出
                var firstRowMatch = System.Text.RegularExpressions.Regex.Match(
                    tableXml, @"<w:tr[ >].*?</w:tr>", System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!firstRowMatch.Success) return false;
                string rowXml = firstRowMatch.Value;

                // 3. <w14:textOutline> ブロックを抽出し、中身が accent2/ED7D31 か確認
                //    「輪郭の中身」だけを見ることで、フォント色との混在パターンを排除する
                var outlineMatch = System.Text.RegularExpressions.Regex.Match(
                    rowXml, @"<w14:textOutline\b[^>]*>.*?</w14:textOutline>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!outlineMatch.Success) return false;
                string outlineXml = outlineMatch.Value;

                bool outlineIsOrange =
                    System.Text.RegularExpressions.Regex.IsMatch(outlineXml, @"w14:schemeClr\s+w14:val=""accent2""")
                    || System.Text.RegularExpressions.Regex.IsMatch(
                        outlineXml, @"w14:srgbClr\s+w14:val=""ED7D31""",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // 4. 塗りつぶしも accent2/ED7D31 か確認
                //    赤枠プリセット = 塗り orange + 輪郭 orange、青枠プリセット = 塗り white + 輪郭 orange
                //    塗りつぶしは <w14:textFill> ブロック または <w:color>（単色時）で表現される
                bool fillIsOrange;
                var fillMatch = System.Text.RegularExpressions.Regex.Match(
                    rowXml, @"<w14:textFill\b[^>]*>.*?</w14:textFill>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (fillMatch.Success)
                {
                    string fillXml = fillMatch.Value;
                    fillIsOrange =
                        System.Text.RegularExpressions.Regex.IsMatch(fillXml, @"w14:schemeClr\s+w14:val=""accent2""")
                        || System.Text.RegularExpressions.Regex.IsMatch(
                            fillXml, @"w14:srgbClr\s+w14:val=""ED7D31""",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                }
                else
                {
                    // <w14:textFill> がない場合、Word は単色塗りを <w:color> で保存している
                    // （赤枠プリセットの正解パターンはこちら）
                    fillIsOrange =
                        System.Text.RegularExpressions.Regex.IsMatch(rowXml, @"<w:color\s+[^>]*w:themeColor=""accent2""")
                        || System.Text.RegularExpressions.Regex.IsMatch(
                            rowXml, @"<w:color\s+[^>]*w:val=""ED7D31""",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                }

                // 5. 追加効果（影・反射・光彩）が存在しないこと
                //    正解プリセットは「塗り + 輪郭」のみのシンプルな効果
                //    影や反射が付いている場合は別プリセット（青枠等）なので × にする
                bool hasAdditionalEffects =
                    rowXml.Contains("<w14:shadow")
                    || rowXml.Contains("<w14:reflection")
                    || rowXml.Contains("<w14:glow");

                // 6. 判定: 輪郭=orange かつ 塗り=orange かつ 追加効果なし
                return outlineIsOrange && fillIsOrange && !hasAdditionalEffects;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_05(string filePath)
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

                // 1. ターゲット特定: 「北海道のうまいもの市場売上」の直後にある表を抽出
                string tableXml = GetTableXmlNearText(xml, "北海道のうまいもの市場売上");
                if (string.IsNullOrEmpty(tableXml)) return false;

                // 2. <w:tblGrid> ブロックを抽出
                var tblGridMatch = System.Text.RegularExpressions.Regex.Match(
                    tableXml, @"<w:tblGrid\b[^>]*>.*?</w:tblGrid>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!tblGridMatch.Success) return false;
                string tblGridXml = tblGridMatch.Value;

                // 3. <w:gridCol w:w="数値"/> の値を全て取得（単位は twip = 1/20 pt）
                var gridCols = System.Text.RegularExpressions.Regex.Matches(
                    tblGridXml, @"<w:gridCol\s+w:w=""(\d+)""");
                if (gridCols.Count < 2) return false;

                // 4. 全ての w:w 値が同一（許容誤差 20twip ≒ 1pt）か確認
                //    Wordの「幅を揃える」機能は内部丸めで完全一致にならないケースがあるため誤差を許容
                int firstWidth;
                if (!int.TryParse(gridCols[0].Groups[1].Value, out firstWidth)) return false;
                const int tolerance = 20;
                for (int i = 1; i < gridCols.Count; i++)
                {
                    int width;
                    if (!int.TryParse(gridCols[i].Groups[1].Value, out width)) return false;
                    if (Math.Abs(width - firstWidth) > tolerance) return false;
                }

                // 5. 操作ログチェック: 「幅を揃える」ボタン（TableColumnsDistribute）が実行されたか
                //    Ribbon.xml で idMso="TableColumnsDistribute" をフックすることで記録される
                //    手動で列幅を同じ値に設定した場合はログが残らないので × となる
                if (!LogReader.HasTaskEvidence(6, 4, "TableColumnsDistribute")) return false;

                return true;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_06(string filePath)
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

                // 1. ターゲット特定: 「北海道のうまいもの市場売上」の直後にある表を抽出
                //    （文書内に複数表があっても、見出し直下の表に対象を絞る）
                string tableXml = GetTableXmlNearText(xml, "北海道のうまいもの市場売上");
                if (string.IsNullOrEmpty(tableXml)) return false;

                // 2. 1行目の <w:tr> を抽出
                var firstRowMatch = System.Text.RegularExpressions.Regex.Match(
                    tableXml, @"<w:tr[ >].*?</w:tr>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!firstRowMatch.Success) return false;
                string firstRowXml = firstRowMatch.Value;

                // 3. 1行目の <w:trPr> 内に <w:tblHeader> が存在するか確認
                //    （タイトル行繰り返し設定でのみ生成される一意なタグ）
                var trPrMatch = System.Text.RegularExpressions.Regex.Match(
                    firstRowXml, @"<w:trPr\b[^>]*>.*?</w:trPr>",
                    System.Text.RegularExpressions.RegexOptions.Singleline);
                if (!trPrMatch.Success) return false;
                string trPrXml = trPrMatch.Value;

                // <w:tblHeader/> または <w:tblHeader ... /> または <w:tblHeader></w:tblHeader> いずれも許容
                bool hasTblHeader = System.Text.RegularExpressions.Regex.IsMatch(
                    trPrXml, @"<w:tblHeader\b");

                return hasTblHeader;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_6_07(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 13行3列の表が存在し、1行目に「最高売上」「平均売上」「合計売上」が左から順に並んでいるかチェック
                Tables tables = document.Tables;
                bool result = false;
                foreach (Table table in tables)
                {
                    if (table.Rows.Count == 13 && table.Columns.Count == 3)
                    {
                        string row1Text = table.Rows[1].Range.Text;
                        int idx1 = row1Text.IndexOf("最高売上");
                        int idx2 = row1Text.IndexOf("平均売上");
                        int idx3 = row1Text.IndexOf("合計売上");

                        if (idx1 >= 0 && idx2 > idx1 && idx3 > idx2)
                        {
                            result = true;
                            break;
                        }
                    }
                    Marshal.ReleaseComObject(table);
                }
                Marshal.ReleaseComObject(tables); return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>
        /// 指定したテキストを含む段落の直後にある表のXMLブロックを抽出します。
        /// </summary>
        private string GetTableXmlNearText(string fullXml, string targetText)
        {
            var paragraphs = System.Text.RegularExpressions.Regex.Matches(fullXml, @"<w:p\b[^>]*>.*?</w:p>", System.Text.RegularExpressions.RegexOptions.Singleline);

            for (int i = 0; i < paragraphs.Count; i++)
            {
                string pXml = paragraphs[i].Value;
                string pText = System.Text.RegularExpressions.Regex.Replace(pXml, @"<[^>]+>", "");

                if (pText.Contains(targetText))
                {
                    // この段落の直後（10000文字以内）にある最初の表を探す
                    // 注: 見出しと表の間に他の段落（例: 「単位:千円」）が挟まる場合、
                    //     OOXMLの書式情報の量によっては500文字では届かないため余裕を持たせる
                    int startIndex = paragraphs[i].Index + paragraphs[i].Length;
                    var tableMatch = System.Text.RegularExpressions.Regex.Match(fullXml.Substring(startIndex), @"<w:tbl\b[^>]*>.*?</w:tbl>", System.Text.RegularExpressions.RegexOptions.Singleline);

                    if (tableMatch.Success && tableMatch.Index < 10000)
                    {
                        return tableMatch.Value;
                    }
                }
            }
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

