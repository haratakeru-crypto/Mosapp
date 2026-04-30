using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace mogiExcelChecker9
{
    public class mogiExcelChecker9
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\mogi_project9.xlsx";

        /// <summary>
        /// Project9のタスク9-1をチェックする
        /// 「担当者マスター」シートのE列に勤続年数が5より大きければ「あり」、そうでなければ「なし」と表示
        /// </summary>
        public bool CheckProject_09_Task_09_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_09_Task_09_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project9のタスク9-2をチェックする
        /// 「出張精算」シートの「手当金額」の列に距離が300㎞以上であれば「10000」、なければ「5000」と表示
        /// </summary>
        public bool CheckProject_09_Task_09_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_09_Task_09_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project9のタスク9-3をチェックする
        /// 「売上一覧」シートの在庫が13％以下であれば「在庫を補充」、そうでなければ空欄
        /// </summary>
        public bool CheckProject_09_Task_09_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_09_Task_09_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project9のタスク9-4をチェックする
        /// 「模試結果」シート内の表の通し番号を関数を使って1から自動で順に表示
        /// </summary>
        public bool CheckProject_09_Task_09_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_09_Task_09_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project9のタスク9-5をチェックする
        /// 「営業予定」シートの時間変更の列に9時から15分ごとに内容が終わるように変更
        /// </summary>
        public bool CheckProject_09_Task_09_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_09_Task_09_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project9のタスク9-6をチェックする
        /// 「担当リスト」シートの売上順担当者名の列に売上が高い順に並び替えて表示
        /// </summary>
        public bool CheckProject_09_Task_09_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_09_Task_09_06(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project9のタスク9-7をチェックする
        /// 「売上一覧」シートの税込価格の列に税込み金額を求める
        /// </summary>
        public bool CheckProject_09_Task_09_07()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_09_Task_09_07(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string ValidateProject_9()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: mogi_project9.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: mogi_project9.xlsxが見つかりません。";
                }

                var results = new System.Collections.Generic.List<string>();
                
                bool task1 = CheckProject_09_Task_09_01(TARGET_FILE_PATH);
                results.Add($"Task 9-1 (IF関数・勤続年数): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckProject_09_Task_09_02(TARGET_FILE_PATH);
                results.Add($"Task 9-2 (IF関数・手当金額): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckProject_09_Task_09_03(TARGET_FILE_PATH);
                results.Add($"Task 9-3 (IF関数・在庫補充): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckProject_09_Task_09_04(TARGET_FILE_PATH);
                results.Add($"Task 9-4 (SEQUENCE関数・通し番号): {(task4 ? "OK" : "NG")}");
                
                bool task5 = CheckProject_09_Task_09_05(TARGET_FILE_PATH);
                results.Add($"Task 9-5 (SEQUENCE関数・時間): {(task5 ? "OK" : "NG")}");
                
                bool task6 = CheckProject_09_Task_09_06(TARGET_FILE_PATH);
                results.Add($"Task 9-6 (SORT関数・並び替え): {(task6 ? "OK" : "NG")}");
                
                bool task7 = CheckProject_09_Task_09_07(TARGET_FILE_PATH);
                results.Add($"Task 9-7 (税込価格計算): {(task7 ? "OK" : "NG")}");
                
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

        private bool CheckProject_09_Task_09_01(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "担当者マスター");
                if (worksheet == null) return false;
                
                // E11セルのIF関数をチェック
                Range targetCell = worksheet.Range["E11"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("IF(D11>5,\"あり\",\"なし\")") || 
                           normalizedFormula.Contains("=IF(D11>5,\"あり\",\"なし\")");
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

        private bool CheckProject_09_Task_09_02(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "出張精算");
                if (worksheet == null) return false;
                
                // G8セルのIF関数をチェック
                Range targetCell = worksheet.Range["G8"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("IF(F8>=300,10000,5000)") || 
                           normalizedFormula.Contains("=IF(F8>=300,10000,5000)");
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

        private bool CheckProject_09_Task_09_03(string filePath)
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
                
                // H7セルのIF関数をチェック
                Range targetCell = worksheet.Range["H7"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("IF(G7<=0.13,\"在庫を補充\",\"\")") || 
                           normalizedFormula.Contains("=IF(G7<=0.13,\"在庫を補充\",\"\")");
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

        private bool CheckProject_09_Task_09_04(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "模試結果");
                if (worksheet == null) return false;
                
                // B8セルのSEQUENCE関数をチェック
                Range targetCell = worksheet.Range["B8"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("SEQUENCE(22,1,1,1)") || 
                           normalizedFormula.Contains("=SEQUENCE(22,1,1,1)");
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

        private bool CheckProject_09_Task_09_05(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "営業予定");
                if (worksheet == null) return false;
                
                // D10セルのSEQUENCE関数をチェック
                Range targetCell = worksheet.Range["D10"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("SEQUENCE(6,1,\"09:00\",\"00:15\")") || 
                           normalizedFormula.Contains("=SEQUENCE(6,1,\"09:00\",\"00:15\")");
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

        private bool CheckProject_09_Task_09_06(string filePath)
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
                    return normalizedFormula.Contains("SORT(A10:B15,2,-1)") || 
                           normalizedFormula.Contains("=SORT(A10:B15,2,-1)");
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

        private bool CheckProject_09_Task_09_07(string filePath)
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
                
                // J7セルの税込価格の計算をチェック
                Range targetCell = worksheet.Range["J7"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("F7*$L$11") || 
                           normalizedFormula.Contains("=F7*$L$11");
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