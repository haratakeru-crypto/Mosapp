using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace mogiExcelChecker8
{
    public class mogiExcelChecker8
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\mogi_project8.xlsx";

        /// <summary>
        /// Project8のタスク8-1をチェックする
        /// キャンパス別試験結果シートのB6:C26に「学生」という名前を付ける
        /// </summary>
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
        /// 「消費税」という名前付き範囲に移動し、数字を「10％」と変更
        /// </summary>
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
        /// 「文化祭」シートのJ7に全部活の金額を合計した数値を「各売上合計」を使い表示
        /// </summary>
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
        /// キャンパス別試験結果シートの表内の学部学科の列に学部と学科を付け加える
        /// </summary>
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
        /// 「担当者マスター」シートの表のアカウントとアドレス「@win.jp」を組み合わせてメールアドレスの列に表示
        /// </summary>
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
        /// 「売上一覧」シートのタグ列に商品型番-商品名と表示
        /// </summary>
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
                    return "警告: mogi_project8.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: mogi_project8.xlsxが見つかりません。";
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
                results.Add($"Task 8-6 (CONCAT関数・タグ): {(task6 ? "OK" : "NG")}");
                
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
                
                // 名前「学生」が定義されているかチェック
                foreach (Name name in workbook.Names)
                {
                    if (name.Name == "学生")
                    {
                        // 参照先がB6:C26かチェック
                        string refersTo = name.RefersTo;
                        if (refersTo.Contains("B6:C26"))
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
                
                // 名前「消費税」が定義されているかチェック
                foreach (Name name in workbook.Names)
                {
                    if (name.Name == "消費税")
                    {
                        // 参照先の値が「10%」かチェック
                        try
                        {
                            Range range = name.RefersToRange;
                            if (range != null)
                            {
                                double value = range.Value2;
                                if (value == 0.1)
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
                
                worksheet = FindWorksheet(workbook, "文化祭");
                if (worksheet == null) return false;
                
                // J7セルで名前付き範囲「各売上合計」を使用したSUM関数をチェック
                Range targetCell = worksheet.Range["J8"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("SUM(各売上合計)") || normalizedFormula.Contains("=SUM(各売上合計)");
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
                
                worksheet = FindWorksheet(workbook, "キャンパス別試験結果");
                if (worksheet == null) return false;
                
                // G7セルのCONCAT関数をチェック
                Range targetCell = worksheet.Range["G7"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("CONCAT(E7:F7)") || normalizedFormula.Contains("=CONCAT(E7:F7)");
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
                
                worksheet = FindWorksheet(workbook, "担当者マスター");
                if (worksheet == null) return false;
                
                // F11セルのCONCAT関数をチェック
                Range targetCell = worksheet.Range["F11"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("CONCAT(C11,\"@WIN.JP\")") || normalizedFormula.Contains("=CONCAT(C11,\"@WIN.JP\")");
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
                
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null) return false;
                
                // G7セルのCONCAT関数をチェック（商品型番-商品名）
                Range targetCell = worksheet.Range["G7"];
                string formula = targetCell.Formula;
                
                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    return normalizedFormula.Contains("CONCAT(D7,\"-\",E7)") || normalizedFormula.Contains("=CONCAT(D7,\"-\",E7)");
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