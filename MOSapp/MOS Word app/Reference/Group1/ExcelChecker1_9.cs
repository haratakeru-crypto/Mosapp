using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group1
{
    public class ExcelChecker1_9
    {
        public bool CheckExcel(string filePath)
        {
            return true;
        }

        public bool CheckTask_1_9_01()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_9_01_Impl(filePath);
        }

        public bool CheckTask_1_9_02()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_9_02_Impl(filePath);
        }

        public bool CheckTask_1_9_03()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_9_03_Impl(filePath);
        }

        public bool CheckTask_1_9_04()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_9_04_Impl(filePath);
        }

        public bool CheckTask_1_9_05()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_9_05_Impl(filePath);
        }

        public bool CheckTask_1_9_06()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_9_06_Impl(filePath);
        }

        public bool CheckTask_1_9_07()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_9_07_Impl(filePath);
        }

        private bool CheckTask_1_9_01_Impl(string filePath)
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

                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);

                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }

                if (workbook == null)
                {
                    return false;
                }

                // タスク9-1: シート［売上報告］の数式を表示
                worksheet = FindWorksheet(workbook, "売上報告");
                if (worksheet == null) return false;

                worksheet.Activate();
                Window window = excelApp.ActiveWindow;
                
                // 数式が表示されているかチェック
                return window.DisplayFormulas;
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

        private bool CheckTask_1_9_02_Impl(string filePath)
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

                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);

                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }

                if (workbook == null)
                {
                    return false;
                }

                // タスク9-2: シート［受注明細］の表を「商品ID」の昇順、「商品ID」が同じ場合は「金額」の降順に並べ替え
                worksheet = FindWorksheet(workbook, "受注明細");
                if (worksheet == null) return false;

                // 並べ替えが正しく行われているかチェック
                // データ範囲を特定して、商品IDの昇順と金額の降順をチェック
                Range usedRange = worksheet.UsedRange;
                int lastRow = usedRange.Rows.Count;
                
                // 最低2行のデータがあることを確認
                if (lastRow < 6) return false;

                string previousProductId = "";
                double previousAmount = double.MaxValue;
                bool isFirstRow = true;

                for (int i = 5; i <= lastRow; i++) // 5行目からデータ開始と仮定
                {
                    Range productIdCell = worksheet.Cells[i, 2] as Range; // B列: 商品ID
                    Range amountCell = worksheet.Cells[i, 5] as Range; // E列: 金額

                    if (productIdCell == null || amountCell == null) continue;
                    
                    object productIdValue = productIdCell.Value2;
                    object amountValue = amountCell.Value2;
                    
                    if (productIdValue == null) break; // データの終わり
                    
                    string currentProductId = productIdValue.ToString();
                    double currentAmount = 0;
                    
                    if (amountValue != null)
                    {
                        double.TryParse(amountValue.ToString(), out currentAmount);
                    }

                    if (!isFirstRow)
                    {
                        // 商品IDが同じ場合、金額が降順（前の行>=現在の行）かチェック
                        if (currentProductId == previousProductId)
                        {
                            if (previousAmount < currentAmount)
                            {
                                return false; // 金額の降順が守られていない
                            }
                        }
                        // 商品IDが異なる場合、IDが昇順かチェック
                        else
                        {
                            if (string.Compare(previousProductId, currentProductId, StringComparison.Ordinal) > 0)
                            {
                                return false; // 商品IDの昇順が守られていない
                            }
                            previousAmount = double.MaxValue; // 新しい商品IDグループ
                        }
                    }

                    previousProductId = currentProductId;
                    previousAmount = currentAmount;
                    isFirstRow = false;
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

        private bool CheckTask_1_9_03_Impl(string filePath)
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

                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);

                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }

                if (workbook == null)
                {
                    return false;
                }

                // タスク9-3: シート［下半期売上］の「7月」から「12月」の列に、アイコンセット「3つの矢印（色分け）」を設定
                worksheet = FindWorksheet(workbook, "下半期売上");
                if (worksheet == null) return false;

                Range targetRange = worksheet.Range["D5:I12"]; // 7月から12月のデータ範囲
                
                // 条件付き書式があるかチェック
                if (targetRange.FormatConditions.Count > 0)
                {
                    foreach (FormatCondition fc in targetRange.FormatConditions)
                    {
                        // アイコンセットの条件付き書式をチェック (Type = 6 はxlIconSet)
                        if ((int)fc.Type == 6)
                        {
                            try
                            {
                                IconSetCondition iconSet = fc as IconSetCondition;
                                if (iconSet != null)
                                {
                                    // 3つの矢印（色分け）アイコンセットをチェック
                                    if (iconSet.IconSet.ID == (int)XlIconSet.xl3Arrows ||
                                        iconSet.IconSet.ID == (int)XlIconSet.xl3ArrowsGray)
                                    {
                                        return true;
                                    }
                                }
                            }
                            catch
                            {
                                // IconSetConditionへのキャストが失敗した場合は次へ
                            }
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

        private bool CheckTask_1_9_04_Impl(string filePath)
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

                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);

                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }

                if (workbook == null)
                {
                    return false;
                }

                // タスク9-4: シート［下半期売上］の「7月」から「12月」の列に、条件付き書式で「3000」より大きいセルに「濃い黄色の文字、黄色の背景」を設定
                worksheet = FindWorksheet(workbook, "下半期売上");
                if (worksheet == null) return false;

                Range targetRange = worksheet.Range["D5:I12"]; // 7月から12月のデータ範囲
                
                // 条件付き書式があるかチェック
                if (targetRange.FormatConditions.Count > 0)
                {
                    foreach (FormatCondition fc in targetRange.FormatConditions)
                    {
                        // セルの値による条件付き書式をチェック (Type = 1 はxlCellValue)
                        if ((int)fc.Type == (int)XlFormatConditionType.xlCellValue)
                        {
                            try
                            {
                                // Formula1に"3000"が含まれているかチェック
                                string formula1 = fc.Formula1 as string;
                                if (formula1 != null && formula1.Contains("3000"))
                                {
                                    // Operatorが「より大きい」であることをチェック
                                    if ((int)fc.Operator == (int)XlFormatConditionOperator.xlGreater)
                                    {
                                        return true;
                                    }
                                }
                            }
                            catch
                            {
                                // 次の条件をチェック
                            }
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

        private bool CheckTask_1_9_05_Impl(string filePath)
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

                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);

                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }

                if (workbook == null)
                {
                    return false;
                }

                // タスク9-5: シート「受注明細」のヘッダーの右に現在の日付を挿入
                worksheet = FindWorksheet(workbook, "受注明細");
                if (worksheet == null) return false;

                // ヘッダーの右側に日付コードがあるかチェック
                string rightHeader = worksheet.PageSetup.RightHeader as string;
                if (rightHeader != null)
                {
                    // &[日付]または&Dが含まれているかチェック
                    return rightHeader.Contains("&D") || rightHeader.Contains("&[日付]");
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

        private bool CheckTask_1_9_06_Impl(string filePath)
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

                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);

                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }

                if (workbook == null)
                {
                    return false;
                }

                // タスク9-6: シート「受注明細」のフッターの右に「P/N」を挿入（Pはページ番号、Nは総ページ数）
                worksheet = FindWorksheet(workbook, "受注明細");
                if (worksheet == null) return false;

                // フッターの右側にページ番号/総ページ数のコードがあるかチェック
                string rightFooter = worksheet.PageSetup.RightFooter as string;
                if (rightFooter != null)
                {
                    // &P/&N または &[ページ番号]/&[総ページ数] が含まれているかチェック
                    return (rightFooter.Contains("&P") && rightFooter.Contains("&N") && rightFooter.Contains("/")) ||
                           (rightFooter.Contains("&[ページ番号]") && rightFooter.Contains("&[総ページ数]"));
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

        private bool CheckTask_1_9_07_Impl(string filePath)
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

                workbook = null;
                string fileName = System.IO.Path.GetFileName(filePath);

                foreach (Workbook wb in excelApp.Workbooks)
                {
                    if (wb.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                        wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = wb;
                        break;
                    }
                }

                if (workbook == null)
                {
                    return false;
                }

                // タスク9-7: アクセシビリティチェック - 負の数に黒いマイナス記号を表示する形式を選択
                worksheet = FindWorksheet(workbook, "下半期売上");
                if (worksheet == null) return false;

                // H7セルの数値書式をチェック
                Range targetCell = worksheet.Range["H7"];
                string numberFormat = targetCell.NumberFormat as string;
                
                if (numberFormat != null)
                {
                    // 負の数値の表示形式が赤色でないことを確認
                    // 標準的な負の数の形式: -1234 または (1234)
                    // 赤色の形式には [Red] が含まれる
                    return !numberFormat.Contains("[Red]") && !numberFormat.Contains("[赤]");
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

        private string GetCurrentExcelFilePath()
        {
            try
            {
                Application excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                if (excelApp.ActiveWorkbook != null)
                {
                    return excelApp.ActiveWorkbook.FullName;
                }
            }
            catch
            {
                // Excel not running or no active workbook
            }
            return null;
        }

        private Worksheet FindWorksheet(Workbook workbook, string worksheetName)
        {
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name.Equals(worksheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return ws;
                }
            }
            return null;
        }
    }
}