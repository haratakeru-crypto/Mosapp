using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace mogiExcelChecker10
{
    public class mogiExcelChecker10
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\mogi_project10.xlsx";

        /// <summary>
        /// Project10のタスク10-1をチェックする
        /// 「売上一覧」シートのG6に「タグを表示してください。」というメモを表示
        /// </summary>
        public bool CheckProject_10_Task_10_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_10_Task_10_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project10のタスク10-2をチェックする
        /// 「売上一覧」シートのタグ列に商品型番-商品名と表示
        /// </summary>
        public bool CheckProject_10_Task_10_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_10_Task_10_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project10のタスク10-3をチェックする
        /// 「売上一覧」シートの印刷の向きを横向きに
        /// </summary>
        public bool CheckProject_10_Task_10_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_10_Task_10_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project10のタスク10-4をチェックする
        /// 「売上一覧」シートのI6のパソコンのご相談の文字に「http:pcostomer.jp」のリンクを挿入
        /// </summary>
        public bool CheckProject_10_Task_10_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_10_Task_10_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project10のタスク10-5をチェックする
        /// 「売上一覧」シートの商品の種類の列に商品名を重複せずにすべて表示
        /// </summary>
        public bool CheckProject_10_Task_10_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_10_Task_10_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project10のタスク10-6をチェックする
        /// 「担当リスト」シートの売上順担当者名の列に売上が高い順に並び替えて表示
        /// </summary>
        public bool CheckProject_10_Task_10_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_10_Task_10_06(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string ValidateProject_10()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: mogi_project10.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: mogi_project10.xlsxが見つかりません。";
                }

                var results = new System.Collections.Generic.List<string>();
                
                bool task1 = CheckProject_10_Task_10_01(TARGET_FILE_PATH);
                results.Add($"Task 10-1 (メモ追加): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckProject_10_Task_10_02(TARGET_FILE_PATH);
                results.Add($"Task 10-2 (タグ表示): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckProject_10_Task_10_03(TARGET_FILE_PATH);
                results.Add($"Task 10-3 (印刷向き): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckProject_10_Task_10_04(TARGET_FILE_PATH);
                results.Add($"Task 10-4 (ハイパーリンク): {(task4 ? "OK" : "NG")}");
                
                bool task5 = CheckProject_10_Task_10_05(TARGET_FILE_PATH);
                results.Add($"Task 10-5 (UNIQUE関数): {(task5 ? "OK" : "NG")}");
                
                bool task6 = CheckProject_10_Task_10_06(TARGET_FILE_PATH);
                results.Add($"Task 10-6 (SORT関数): {(task6 ? "OK" : "NG")}");
                
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

        private bool CheckProject_10_Task_10_01(string filePath)
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
                
                // G6セルにメモが追加されているかチェック
                Range targetCell = worksheet.Range["G6"];
                Comment comment = targetCell.Comment;
                if (comment != null)
                {
                    string commentText = comment.Text();
                    return commentText.Contains("タグを表示してください。");
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

        private bool CheckProject_10_Task_10_02(string filePath)
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
                
                // G7セルのCONCAT関数をチェック
                Range targetCell = worksheet.Range["G7"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // CONCAT関数でD7とE7を結合しているかチェック
                    return normalizedFormula.Contains("CONCAT(D7,\"-\",E7)") || 
                           normalizedFormula.Contains("=CONCAT(D7,\"-\",E7)") ||
                           normalizedFormula.Contains("CONCAT(D7") && normalizedFormula.Contains("E7");
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

        private bool CheckProject_10_Task_10_03(string filePath)
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
                
                // 印刷の向きが横向きかチェック
                PageSetup pageSetup = worksheet.PageSetup;
                return pageSetup.Orientation == XlPageOrientation.xlLandscape;
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

        private bool CheckProject_10_Task_10_04(string filePath)
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
                
                // I6セルのハイパーリンクをチェック
                Range targetCell = worksheet.Range["I6"];
                if (targetCell.Hyperlinks.Count > 0)
                {
                    Hyperlink hyperlink = targetCell.Hyperlinks[1];
                    string address = hyperlink.Address;
                    if (address != null && address.Contains("http:pcostomer.jp"))
                    {
                        return true;
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

        private bool CheckProject_10_Task_10_05(string filePath)
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
                
                // H7セルのUNIQUE関数をチェック
                Range targetCell = worksheet.Range["H7"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // UNIQUE関数でE7:E176を使用しているかチェック
                    return normalizedFormula.Contains("UNIQUE(E7:E176)") || 
                           normalizedFormula.Contains("=UNIQUE(E7:E176)");
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

        private bool CheckProject_10_Task_10_06(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "担当リスト");
                if (worksheet == null) return false;
                
                // D10セルのSORT関数をチェック
                Range targetCell = worksheet.Range["D10"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // SORT関数でA10:B15を売上の高い順に並び替えているかチェック
                    return normalizedFormula.Contains("SORT(A10:B15,2,-1)") || 
                           normalizedFormula.Contains("=SORT(A10:B15,2,-1)") ||
                           (normalizedFormula.Contains("SORT") && normalizedFormula.Contains("A10:B15"));
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