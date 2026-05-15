using System;
using System.IO;
using System.Runtime.InteropServices;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group2
{
    /// <summary>
    /// 比較ベースのチェック用ヘルパークラス
    /// CompletedフォルダとInitialフォルダを比較して正誤判定を行う
    /// </summary>
    public class ExcelCheckerComparisonHelper
    {
        private static JObject _config;

        static ExcelCheckerComparisonHelper()
        {
            LoadConfig();
        }

        private static void LoadConfig()
        {
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "config.json");
                if (File.Exists(configPath))
                {
                    string json = File.ReadAllText(configPath);
                    _config = JObject.Parse(json);
                }
            }
            catch
            {
                // config.jsonの読み込みに失敗した場合
            }
        }

        /// <summary>
        /// config.jsonから初期データファイルのパスを取得
        /// </summary>
        public static string GetInitialDataFilePath(int tabId, int projectId)
        {
            if (_config == null) return null;

            try
            {
                var projectData = _config["tabs"]?[tabId.ToString()]?["projects"]?[projectId.ToString()];
                return projectData?["initialDataFile"]?.ToString();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// config.jsonから完成データファイルのパスを取得
        /// </summary>
        public static string GetCompletedDataFilePath(int tabId, int projectId)
        {
            if (_config == null) return null;

            try
            {
                var projectData = _config["tabs"]?[tabId.ToString()]?["projects"]?[projectId.ToString()];
                return projectData?["completedDataFile"]?.ToString();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// ワークブックを取得
        /// </summary>
        public static Workbook GetWorkbook(Application excelApp, string filePath)
        {
            if (excelApp == null || string.IsNullOrEmpty(filePath))
                return null;

            try
            {
                string fileName = Path.GetFileName(filePath);
                
                // 既に開いているワークブックを検索
                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        return wb;
                    }
                }

                // 開いていない場合は開く
                if (File.Exists(filePath))
                {
                    return excelApp.Workbooks.Open(filePath, ReadOnly: true);
                }
            }
            catch
            {
                // エラーが発生した場合
            }
            return null;
        }

        /// <summary>
        /// ワークシートを検索
        /// </summary>
        public static Worksheet FindWorksheet(Workbook workbook, string sheetName)
        {
            if (workbook == null) return null;

            try
            {
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return worksheet;
                    }
                }
            }
            catch
            {
                // シート検索でエラーが発生した場合
            }
            return null;
        }

        /// <summary>
        /// 最初のテーブルを検索
        /// </summary>
        public static ListObject FindFirstTable(Worksheet worksheet)
        {
            if (worksheet == null) return null;

            try
            {
                if (worksheet.ListObjects.Count > 0)
                {
                    return worksheet.ListObjects[1];
                }
            }
            catch
            {
                // テーブル検索でエラーが発生した場合
            }
            return null;
        }

        /// <summary>
        /// ページ設定を比較
        /// </summary>
        public static bool ComparePageSetup(PageSetup setup1, PageSetup setup2)
        {
            if (setup1 == null || setup2 == null) return false;

            try
            {
                // 印刷方向
                if (setup1.Orientation != setup2.Orientation)
                    return false;

                // 印刷範囲
                string printArea1 = setup1.PrintArea ?? string.Empty;
                string printArea2 = setup2.PrintArea ?? string.Empty;
                if (NormalizeRange(printArea1) != NormalizeRange(printArea2))
                    return false;

                // 印刷タイトル行
                string titleRows1 = setup1.PrintTitleRows ?? string.Empty;
                string titleRows2 = setup2.PrintTitleRows ?? string.Empty;
                if (NormalizeRange(titleRows1) != NormalizeRange(titleRows2))
                    return false;

                // 余白（許容誤差1ポイント）
                const double tolerance = 1.0;
                if (Math.Abs(setup1.TopMargin - setup2.TopMargin) > tolerance) return false;
                if (Math.Abs(setup1.BottomMargin - setup2.BottomMargin) > tolerance) return false;
                if (Math.Abs(setup1.LeftMargin - setup2.LeftMargin) > tolerance) return false;
                if (Math.Abs(setup1.RightMargin - setup2.RightMargin) > tolerance) return false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// テーブルのプロパティを比較
        /// </summary>
        public static bool CompareTableProperties(ListObject table1, ListObject table2)
        {
            if (table1 == null || table2 == null) return false;

            try
            {
                // テーブルスタイルの比較
                if (table1.TableStyle != table2.TableStyle)
                    return false;

                // 縞模様（行）の比較
                if (table1.ShowTableStyleRowStripes != table2.ShowTableStyleRowStripes)
                    return false;

                // 縞模様（列）の比較
                if (table1.ShowTableStyleColumnStripes != table2.ShowTableStyleColumnStripes)
                    return false;

                // 最初の列の強調の比較
                if (table1.ShowTableStyleFirstColumn != table2.ShowTableStyleFirstColumn)
                    return false;

                // 最後の列の強調の比較
                if (table1.ShowTableStyleLastColumn != table2.ShowTableStyleLastColumn)
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// セルの書式を比較
        /// </summary>
        public static bool CompareCellFormat(Range cell1, Range cell2)
        {
            if (cell1 == null || cell2 == null) return false;

            try
            {
                // フォントの太字
                if (cell1.Font.Bold != cell2.Font.Bold)
                    return false;

                // フォントの斜体
                if (cell1.Font.Italic != cell2.Font.Italic)
                    return false;

                // フォントの下線
                if (cell1.Font.Underline != cell2.Font.Underline)
                    return false;

                // 取り消し線
                if (cell1.Font.Strikethrough != cell2.Font.Strikethrough)
                    return false;

                // 折り返し
                if (cell1.WrapText != cell2.WrapText)
                    return false;

                // 水平方向の配置
                if (cell1.HorizontalAlignment != cell2.HorizontalAlignment)
                    return false;

                // 垂直方向の配置
                if (cell1.VerticalAlignment != cell2.VerticalAlignment)
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 範囲文字列を正規化（$を除去、大文字に変換、スペースを除去）
        /// </summary>
        public static string NormalizeRange(string range)
        {
            if (string.IsNullOrEmpty(range)) return string.Empty;
            return range.Replace("$", "").Replace(" ", "").ToUpper();
        }

        /// <summary>
        /// ワークブックを安全に閉じる
        /// </summary>
        public static void CloseWorkbook(Workbook workbook, bool saveChanges = false)
        {
            if (workbook != null)
            {
                try
                {
                    workbook.Close(saveChanges);
                    Marshal.ReleaseComObject(workbook);
                }
                catch
                {
                    // エラーが発生した場合は無視
                }
            }
        }

        /// <summary>
        /// COMオブジェクトを安全に解放
        /// </summary>
        public static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                try
                {
                    Marshal.ReleaseComObject(obj);
                }
                catch
                {
                    // エラーが発生した場合は無視
                }
            }
        }
    }
}

