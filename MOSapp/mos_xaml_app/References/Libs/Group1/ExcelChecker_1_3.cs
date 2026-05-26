using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelChecker3
{
    public class ExcelChecker3
    {
        private const string TARGET_FILE_PATH = @"C:\MOSTest\Excel365\project3.xlsx";

        /// <summary>
        /// Project3のタスク3-1をチェックする
        /// シート[下半期売上]のセル【A2】に、「タイトル」のセルのスタイルを設定
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_03_Task_03_01()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_03_Task_03_01(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project3のタスク3-2をチェックする
        /// シート[社員リスト]のセル範囲【B5:B44】に左インデントを2文字分設定
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_03_Task_03_02()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_03_Task_03_02(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project3のタスク3-3をチェックする
        /// シート[社員リスト]のセル【A2】の文字の配置をセル範囲【A2:F2】の中央にする
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_03_Task_03_03()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_03_Task_03_03(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project3のタスク3-4をチェックする
        /// シート[担当者別売上]の表に、セル範囲【H5:K19】をコピーし、列幅を保持して貼り付け
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_03_Task_03_04()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_03_Task_03_04(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project3のタスク3-5をチェックする
        /// シート[業務予定]のセル範囲【C5:C11】の時間に、取り消し線を設定
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_03_Task_03_05()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_03_Task_03_05(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project3のタスク3-6をチェックする
        /// シート[参加者一覧]の表のタイトル「セミナー参加者リスト」のセルの結合と文字列の配置を解除
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_03_Task_03_06()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_03_Task_03_06(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Project3のタスク3-7をチェックする
        /// シート[参加者一覧]の表の中から、「氏名」が「風間 健太郎」と「平井 元」の行を削除
        /// </summary>
        /// <returns>チェック結果</returns>
        public bool CheckProject_03_Task_03_07()
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                return CheckProject_03_Task_03_07(filePath);
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string ValidateProject_3()
        {
            try
            {
                if (IsExcelFileOpen(TARGET_FILE_PATH))
                {
                    return "警告: project3.xlsxは既に開いています。ファイルを閉じてから再実行してください。";
                }

                if (!File.Exists(TARGET_FILE_PATH))
                {
                    return "エラー: project3.xlsxが見つかりません。";
                }

                var results = new System.Collections.Generic.List<string>();
                
                bool task1 = CheckProject_03_Task_03_01(TARGET_FILE_PATH);
                results.Add($"Task 3-1 (セルのスタイル): {(task1 ? "OK" : "NG")}");
                
                bool task2 = CheckProject_03_Task_03_02(TARGET_FILE_PATH);
                results.Add($"Task 3-2 (左インデント): {(task2 ? "OK" : "NG")}");
                
                bool task3 = CheckProject_03_Task_03_03(TARGET_FILE_PATH);
                results.Add($"Task 3-3 (選択範囲内で中央): {(task3 ? "OK" : "NG")}");
                
                bool task4 = CheckProject_03_Task_03_04(TARGET_FILE_PATH);
                results.Add($"Task 3-4 (列幅を保持してコピー): {(task4 ? "OK" : "NG")}");
                
                bool task5 = CheckProject_03_Task_03_05(TARGET_FILE_PATH);
                results.Add($"Task 3-5 (取り消し線): {(task5 ? "OK" : "NG")}");
                
                bool task6 = CheckProject_03_Task_03_06(TARGET_FILE_PATH);
                results.Add($"Task 3-6 (セルの結合解除): {(task6 ? "OK" : "NG")}");
                
                bool task7 = CheckProject_03_Task_03_07(TARGET_FILE_PATH);
                results.Add($"Task 3-7 (行削除): {(task7 ? "OK" : "NG")}");
                
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
                // 実行中のExcelアプリケーションを取得
                excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                
                // アクティブなワークブックのパスを取得
                if (excelApp.ActiveWorkbook != null)
                {
                    return excelApp.ActiveWorkbook.FullName;
                }
                
                return null;
            }
            catch (COMException)
            {
                // Excelが起動していない場合
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

        private bool CheckProject_03_Task_03_01(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "下半期売上");
                if (worksheet == null) return false;
                
                Range targetCell = worksheet.Range["A2"];
                string cellStyle = targetCell.Style.Name;
                return cellStyle != null && cellStyle.Contains("タイトル");
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

        private bool CheckProject_03_Task_03_02(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "社員リスト");
                if (worksheet == null) return false;
                
                Range targetRange = worksheet.Range["B5:B44"];
                foreach (Range cell in targetRange.Cells)
                {
                    if (cell.IndentLevel != 2)
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

        private bool CheckProject_03_Task_03_03(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "社員リスト");
                if (worksheet == null) return false;
                
                Range a2Cell = worksheet.Range["A2"];
                return (int)a2Cell.HorizontalAlignment == -4108;
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

        private bool CheckProject_03_Task_03_04(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "担当者別売上");
                if (worksheet == null) return false;
                
                Range sourceRange = worksheet.Range["H5:K19"];
                Range targetRange = worksheet.Range["A5:D19"];
                
                for (int i = 1; i <= sourceRange.Rows.Count; i++)
                {
                    for (int j = 1; j <= sourceRange.Columns.Count; j++)
                    {
                        var sourceValue = sourceRange.Cells[i, j].Value2;
                        var targetValue = targetRange.Cells[i, j].Value2;
                        if (sourceValue != targetValue)
                        {
                            return false;
                        }
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

        private bool CheckProject_03_Task_03_05(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "業務予定");
                if (worksheet == null) return false;
                
                Range targetRange = worksheet.Range["C5:C11"];
                foreach (Range cell in targetRange.Cells)
                {
                    if (!Convert.ToBoolean(cell.Font.Strikethrough))
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

        private bool CheckProject_03_Task_03_06(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "参加者一覧");
                if (worksheet == null) return false;
                
                Range targetCell = worksheet.Range["B2"];
                return targetCell.MergeArea.Cells.Count == 1;
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

        private bool CheckProject_03_Task_03_07(string filePath)
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
                
                worksheet = FindWorksheet(workbook, "参加者一覧");
                if (worksheet == null) return false;
                
                Range usedRange = worksheet.UsedRange;
                int nameColumnIndex = -1;
                for (int col = 1; col <= usedRange.Columns.Count; col++)
                {
                    var cellValue = usedRange.Cells[1, col].Value2;
                    if (cellValue != null && cellValue.ToString().Contains("氏名"))
                    {
                        nameColumnIndex = col;
                        break;
                    }
                }
                
                if (nameColumnIndex == -1) return false;
                
                for (int row = 2; row <= usedRange.Rows.Count; row++)
                {
                    var cellValue = usedRange.Cells[row, nameColumnIndex].Value2;
                    if (cellValue != null)
                    {
                        string name = cellValue.ToString();
                        if (name.Contains("風間 健太郎") || name.Contains("平井 元"))
                        {
                            return false;
                        }
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
    }
} 