using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace mogiExcelChecker2
{
    public class mogiExcelChecker2
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\mogi_project2.xlsx";

        /// <summary>
        /// Project2のタスク2-1をチェックする
        /// キャンパス別試験結果シートの表内の学部学科の列に学部と学科を付け加えて、表を完成させる
        /// </summary>
        public bool CheckProject_02_Task_02_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_02_Task_02_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project2のタスク2-2をチェックする
        /// キャンパス別試験結果シートの氏名の列C7:C26に左インデントを1つ追加
        /// </summary>
        public bool CheckProject_02_Task_02_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_02_Task_02_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project2のタスク2-3をチェックする
        /// キャンパス別試験結果シートのセルB4内の文字をB4とC4の中央に配置
        /// </summary>
        public bool CheckProject_02_Task_02_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_02_Task_02_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project2のタスク2-4をチェックする
        /// キャンパス別試験結果シートのB6:C26に「学生」という名前を付ける
        /// </summary>
        public bool CheckProject_02_Task_02_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_02_Task_02_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project2のタスク2-5をチェックする
        /// 「消費税」という名前付き範囲に移動し、数字を「10％」と変更
        /// </summary>
        public bool CheckProject_02_Task_02_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_02_Task_02_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string ValidateProject_2()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: mogi_project2.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: mogi_project2.xlsxが見つかりません。";
                }

                var results = new System.Collections.Generic.List<string>();
                
                bool task1 = CheckProject_02_Task_02_01(TARGET_FILE_PATH);
                results.Add($"Task 2-1 (CONCAT関数): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckProject_02_Task_02_02(TARGET_FILE_PATH);
                results.Add($"Task 2-2 (左インデント): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckProject_02_Task_02_03(TARGET_FILE_PATH);
                results.Add($"Task 2-3 (中央配置): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckProject_02_Task_02_04(TARGET_FILE_PATH);
                results.Add($"Task 2-4 (名前定義): {(task4 ? "OK" : "NG")}");
                
                bool task5 = CheckProject_02_Task_02_05(TARGET_FILE_PATH);
                results.Add($"Task 2-5 (名前付き範囲変更): {(task5 ? "OK" : "NG")}");
                
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

        private bool CheckProject_02_Task_02_01(string filePath)
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
                    // CONCAT関数でE7:F7を結合しているかチェック
                    return normalizedFormula.Contains("CONCAT(E7:F7)") || 
                           normalizedFormula.Contains("=CONCAT(E7:F7)") ||
                           (normalizedFormula.Contains("CONCAT") && normalizedFormula.Contains("E7") && normalizedFormula.Contains("F7"));
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

        private bool CheckProject_02_Task_02_02(string filePath)
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
                
                // C7:C26の範囲でインデントが1に設定されているかチェック
                Range targetRange = worksheet.Range["C7:C26"];
                foreach (Range cell in targetRange.Cells)
                {
                    if (cell.IndentLevel != 1)
                    {
                        return false;
                    }
                }
                return true;
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

        private bool CheckProject_02_Task_02_03(string filePath)
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
                
                // B4セルの配置が選択範囲内で中央になっているかチェック
                Range b4Cell = worksheet.Range["B4"];
                return (int)b4Cell.HorizontalAlignment == -4108; // xlCenterAcrossSelection
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

        private bool CheckProject_02_Task_02_04(string filePath)
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
                try
                {
                    Range namedRange = workbook.Names.Item("学生").RefersToRange;
                    string address = namedRange.Address;
                    // B6:C26の範囲が名前「学生」として定義されているかチェック
                    return address.Contains("B6:C26") || address.Contains("$B$6:$C$26");
                }
                catch
                {
                    return false;
                }
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

        private bool CheckProject_02_Task_02_05(string filePath)
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
                
                // 名前「消費税」の値が10%かチェック
                try
                {
                    Range namedRange = workbook.Names.Item("消費税").RefersToRange;
                    var value = namedRange.Value2;
                    if (value != null)
                    {
                        string valueStr = value.ToString();
                        return valueStr.Contains("10%") || valueStr.Contains("0.1") || valueStr.Contains("10％");
                    }
                }
                catch
                {
                    return false;
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
    }
} 