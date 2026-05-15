using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelChecker8
{
    public class ExcelChecker8
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\project8.xlsx";

        /// <summary>
        /// Project8のタスク8-1をチェックする
        /// シート[学生名簿]のセル範囲【C5:C24】に、「氏名」の名前を定義
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_08_Task_08_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_08_Task_08_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project8のタスク8-2をチェックする
        /// 名前「開催日」の名前付き範囲に移動し、日付を「5/10」に変更
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_08_Task_08_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_08_Task_08_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project8のタスク8-3をチェックする
        /// シート[売上報告]のセル【J5】に、SUM関数を使用して名前付き範囲「売上合計」を使用
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_08_Task_08_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_08_Task_08_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project8のタスク8-4をチェックする
        /// シート[学生名簿]の「学部学科」の列にCONCAT関数を使って「学部」と「学科」の文字列を結合
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_08_Task_08_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_08_Task_08_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project8のタスク8-5をチェックする
        /// シート[担当者リスト]の「メールアドレス」の列に、CONCAT関数を使って「アカウント」と「@rabbit.ac.jp」をつなげて表示
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_08_Task_08_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_08_Task_08_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project8のタスク8-6をチェックする
        /// シート[申込一覧]の「所属」の列に、CONCAT関数を使って「キャンパス-学部」と表示
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_08_Task_08_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_08_Task_08_06(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string ValidateProject_8()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: project8.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: project8.xlsxが見つかりません。";
                }

                var results = new System.Collections.Generic.List<string>();
                
                bool task1 = CheckProject_08_Task_08_01(TARGET_FILE_PATH);
                results.Add($"Task 8-1 (名前定義): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckProject_08_Task_08_02(TARGET_FILE_PATH);
                results.Add($"Task 8-2 (名前付き範囲変更): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckProject_08_Task_08_03(TARGET_FILE_PATH);
                results.Add($"Task 8-3 (SUM関数・名前付き範囲): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckProject_08_Task_08_04(TARGET_FILE_PATH);
                results.Add($"Task 8-4 (CONCAT関数・文字列結合): {(task4 ? "OK" : "NG")}");
                
                bool task5 = CheckProject_08_Task_08_05(TARGET_FILE_PATH);
                results.Add($"Task 8-5 (CONCAT関数・メール): {(task5 ? "OK" : "NG")}");
                
                bool task6 = CheckProject_08_Task_08_06(TARGET_FILE_PATH);
                results.Add($"Task 8-6 (CONCAT関数・所属): {(task6 ? "OK" : "NG")}");
                
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

        private bool CheckProject_08_Task_08_01(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
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
                
                // 名前「氏名」が定義されているかチェック
                foreach (Name name in workbook.Names)
                {
                    if (name.Name == "氏名")
                    {
                        // 参照先がC5:C24かチェック
                        string refersTo = name.RefersTo;
                        if (refersTo.Contains("C5:C24"))
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
                // workbookはCloseしない（既存のファイルを操作しているため）
            }
        }

        private bool CheckProject_08_Task_08_02(string filePath)
        {
            Application excelApp = null;
            Workbook workbook = null;
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
                
                // 名前「開催日」が定義されているかチェック
                foreach (Name name in workbook.Names)
                {
                    if (name.Name == "開催日")
                    {
                        // 参照先の値が「5/10」かチェック
                        try
                        {
                            Range range = name.RefersToRange;
                            if (range != null)
                            {
                                string value = range.Text.ToString();
                                if (value.Contains("5/10"))
                                {
                                    return true;
                                }
                            }
                        }
                        catch
                        {
                            // 参照先が取得できない場合は継続
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
                // workbookはCloseしない（既存のファイルを操作しているため）
            }
        }

        private bool CheckProject_08_Task_08_03(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "売上報告");
                if (worksheet == null) return false;
                
                // J5セルで名前付き範囲「売上合計」を使用したSUM関数をチェック
                Range targetCell = worksheet.Range["J5"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("SUM(売上合計)") || normalizedFormula.Contains("=SUM(売上合計)");
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

        private bool CheckProject_08_Task_08_04(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "学生名簿");
                if (worksheet == null) return false;
                
                // G5セルのCONCAT関数をチェック
                Range targetCell = worksheet.Range["G5"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("CONCAT(E5:F5)") || normalizedFormula.Contains("=CONCAT(E5:F5)");
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

        private bool CheckProject_08_Task_08_05(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "担当者リスト");
                if (worksheet == null) return false;
                
                // H5セルのCONCAT関数をチェック
                Range targetCell = worksheet.Range["H5"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("CONCAT(G5,\"@RABBIT.AC.JP\")") || normalizedFormula.Contains("=CONCAT(G5,\"@RABBIT.AC.JP\")");
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

        private bool CheckProject_08_Task_08_06(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "申込一覧");
                if (worksheet == null) return false;
                
                // G5セルのCONCAT関数をチェック（キャンパス-学部）
                Range targetCell = worksheet.Range["G5"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("CONCAT(B5,\"-\",E5)") || normalizedFormula.Contains("=CONCAT(B5,\"-\",E5)");
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