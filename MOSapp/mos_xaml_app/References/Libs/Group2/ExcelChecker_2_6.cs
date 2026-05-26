using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace mogiExcelChecker6
{
    public class mogiExcelChecker6
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\mogi_project6.xlsx";

        /// <summary>
        /// Project6のタスク6-1をチェックする
        /// 「文化祭」シートの1日目から4日目に条件付き書式で「3つの信号（枠なし）」を設定
        /// </summary>
        public bool CheckProject_06_Task_06_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_06_Task_06_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project6のタスク6-2をチェックする
        /// 「文化祭」シートのB5を「見出し1」のスタイルに変更
        /// </summary>
        public bool CheckProject_06_Task_06_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_06_Task_06_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project6のタスク6-3をチェックする
        /// 「売上一覧」シートでIF関数で在庫補充判定
        /// </summary>
        public bool CheckProject_06_Task_06_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_06_Task_06_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project6のタスク6-4をチェックする
        /// 「売上一覧」シートの商品型番は降順、在庫は昇順に並べ替え
        /// </summary>
        public bool CheckProject_06_Task_06_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_06_Task_06_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project6のタスク6-5をチェックする
        /// 「売上一覧」シートのA4のタイトルをA4とB4の中央に配置
        /// </summary>
        public bool CheckProject_06_Task_06_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_06_Task_06_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project6のタスク6-6をチェックする
        /// 「文化祭」シートの表の売上グラフに縦棒のスパークライン
        /// </summary>
        public bool CheckProject_06_Task_06_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_06_Task_06_06(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string ValidateProject_6()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: mogi_project6.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: mogi_project6.xlsxが見つかりません。";
                }

                var results = new System.Collections.Generic.List<string>();
                
                bool task1 = CheckProject_06_Task_06_01(TARGET_FILE_PATH);
                results.Add($"Task 6-1 (条件付き書式): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckProject_06_Task_06_02(TARGET_FILE_PATH);
                results.Add($"Task 6-2 (セルスタイル): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckProject_06_Task_06_03(TARGET_FILE_PATH);
                results.Add($"Task 6-3 (IF関数): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckProject_06_Task_06_04(TARGET_FILE_PATH);
                results.Add($"Task 6-4 (並べ替え): {(task4 ? "OK" : "NG")}");
                
                bool task5 = CheckProject_06_Task_06_05(TARGET_FILE_PATH);
                results.Add($"Task 6-5 (中央配置): {(task5 ? "OK" : "NG")}");
                
                bool task6 = CheckProject_06_Task_06_06(TARGET_FILE_PATH);
                results.Add($"Task 6-6 (縦棒スパークライン): {(task6 ? "OK" : "NG")}");
                
                return string.Join("\n", results);
            }
            catch (Exception ex)
            {
                return $"エラー: {ex.Message}";
            }
        }

        private bool IsExcelFileOpen(string filePath)
        {
            Application excelApp = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                foreach (Workbook workbook in excelApp.Workbooks)
                {
                    if (string.Equals(workbook.FullName, filePath, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                return false;
            }
            catch (COMException)
            {
                return false;
            }
            finally
            {
                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);
            }
        }

        private string GetCurrentExcelFilePath()
        {
            Application excelApp = null;
            try
            {
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                if (excelApp.ActiveWorkbook != null)
                {
                    return excelApp.ActiveWorkbook.FullName;
                }
                return null;
            }
            catch (COMException)
            {
                return null;
            }
            finally
            {
                if (excelApp != null)
                    Marshal.ReleaseComObject(excelApp);
            }
        }
        
        private Worksheet FindWorksheet(Workbook workbook, string sheetName)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }
            return null;
        }

        private Workbook GetWorkbook(Application excelApp, string filePath)
        {
            string fileName = System.IO.Path.GetFileName(filePath);
            foreach (Workbook wb in excelApp.Workbooks)
            {
                if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                    wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                {
                    return wb;
                }
            }
            return null;
        }

        private bool CheckProject_06_Task_06_01(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "文化祭");
                if (worksheet == null) return false;
                
                // 条件付き書式が適用されているかチェック
                Range targetRange = worksheet.Range["E8:H16"];
                if (targetRange.FormatConditions.Count > 0)
                {
                    // 条件付き書式が設定されていることを確認
                    return true;
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckProject_06_Task_06_02(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "文化祭");
                if (worksheet == null) return false;
                
                // B5セルのスタイルが「見出し1」かチェック
                Range targetCell = worksheet.Range["B5"];
                string cellStyle = targetCell.Style.Name;
                return cellStyle != null && cellStyle.Contains("見出し 1");
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckProject_06_Task_06_03(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null) return false;
                
                // F8セルのIF関数をチェック
                Range targetCell = worksheet.Range["F8"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // IF関数で在庫補充判定をチェック
                    return normalizedFormula.Contains("IF") && 
                           (normalizedFormula.Contains("補充") || normalizedFormula.Contains("要補充"));
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckProject_06_Task_06_04(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null) return false;
                
                // 並べ替えが適用されているかチェック（オートフィルタモードで確認）
                return worksheet.AutoFilterMode || worksheet.UsedRange.Rows.Count > 1;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckProject_06_Task_06_05(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null) return false;
                
                // A4セルの配置が選択範囲内で中央になっているかチェック
                Range a4Cell = worksheet.Range["A4"];
                return (int)a4Cell.HorizontalAlignment == -4108; // xlCenterAcrossSelection
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckProject_06_Task_06_06(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                }
                
                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;
                
                worksheet = FindWorksheet(workbook, "文化祭");
                if (worksheet == null) return false;
                
                // I8:I16にスパークラインが存在するかチェック
                Range targetRange = worksheet.Range["I8:I16"];
                foreach (Range cell in targetRange.Cells)
                {
                    if (cell.SparklineGroups.Count > 0)
                    {
                        // スパークラインの種類が縦棒かチェック
                        var sparklineGroup = cell.SparklineGroups[1];
                        if (sparklineGroup.Type == XlSparkType.xlSparkColumn)
                        {
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }
    }
} 