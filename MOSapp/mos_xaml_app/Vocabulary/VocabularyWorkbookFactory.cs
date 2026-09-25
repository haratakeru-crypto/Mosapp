using System;
using System.IO;
using System.Runtime.InteropServices;
using ExcelApp = Microsoft.Office.Interop.Excel.Application;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelListObject = Microsoft.Office.Interop.Excel.ListObject;
using ExcelXlChartType = Microsoft.Office.Interop.Excel.XlChartType;

namespace MOSExcelMogiApp.Vocabulary
{
    public static class VocabularyWorkbookFactory
    {
        public static string GetWorkbookPath()
        {
            string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "Vocabulary");
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "vocab_workbook.xlsx");
        }

        /// <summary>
        /// テーブル＋簡易グラフ入りの教材ブックを用意する（無ければ作成）。
        /// </summary>
        public static string EnsureWorkbook(ExcelApp excelApp)
        {
            string path = GetWorkbookPath();
            if (File.Exists(path))
                return path;

            ExcelWorkbook wb = null;
            try
            {
                wb = excelApp.Workbooks.Add();
                ExcelWorksheet ws = (ExcelWorksheet)wb.Worksheets[1];
                ws.Name = "単語帳";

                object[,] data =
                {
                    { "店舗", "担当者名", "アカウント", "勤続年数", "昇給" },
                    { "池袋店", "神山　仁", "kamiyama01", 5, "" },
                    { "池袋店", "加賀　亮太", "kaga02", 13, "" },
                    { "秋葉原店", "三島　瞳", "mishima03", 2, "" },
                    { "有楽町店", "水野　奈菜", "mizuno04", 15, "" },
                    { "有楽町店", "岡部　進", "okabe05", 4, "" },
                    { "有楽町店", "谷　幸一郎", "tani06", 6, "" },
                };

                ExcelRange start = ws.Range["B8"];
                ExcelRange end = (ExcelRange)ws.Cells[8 + 6, 2 + 4];
                ExcelRange tableRange = ws.Range[start, end];
                tableRange.Value2 = data;

                for (int r = 9; r <= 14; r++)
                {
                    ((ExcelRange)ws.Cells[r, 6]).Formula = $"=IF(E{r}>5,\"あり\",\"なし\")";
                }

                ExcelListObject table = ws.ListObjects.Add(
                    SourceType: Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange,
                    Source: tableRange,
                    XlListObjectHasHeaders: Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes);
                table.Name = "店舗一覧";

                var chartObjects = (Microsoft.Office.Interop.Excel.ChartObjects)ws.ChartObjects();
                var chartObj = chartObjects.Add(360, 120, 360, 220);
                var chart = chartObj.Chart;
                chart.SetSourceData(ws.Range["E8:E14"]);
                chart.ChartType = ExcelXlChartType.xlColumnClustered;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "勤続年数";

                ws.Range["B2"].Value2 = "MOS Excel 単語帳用シート（テーブルとグラフ）";
                ws.Range["B3"].Value2 = "キーワードに対応するタブ・ボタン／関数を探してください。";

                if (File.Exists(path))
                    File.Delete(path);
                wb.SaveAs(path);
                wb.Close(SaveChanges: false);
                wb = null;
                return path;
            }
            finally
            {
                if (wb != null)
                {
                    try { wb.Close(SaveChanges: false); } catch { }
                    try { Marshal.ReleaseComObject(wb); } catch { }
                }
            }
        }
    }
}
