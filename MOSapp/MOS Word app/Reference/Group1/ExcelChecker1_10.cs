using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group1
{
    public class ExcelChecker1_10
    {
        public bool CheckExcel(string filePath)
        {
            return true;
        }

        public bool CheckTask_1_10_01()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_10_01_Impl(filePath);
        }

        public bool CheckTask_1_10_02()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_10_02_Impl(filePath);
        }

        public bool CheckTask_1_10_03()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_10_03_Impl(filePath);
        }

        public bool CheckTask_1_10_04()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_10_04_Impl(filePath);
        }

        public bool CheckTask_1_10_05()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_10_05_Impl(filePath);
        }

        public bool CheckTask_1_10_06()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_10_06_Impl(filePath);
        }

        public bool CheckTask_1_10_07()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_10_07_Impl(filePath);
        }

        public bool CheckTask_1_10_08()
        {
            string filePath = GetCurrentExcelFilePath();
            if (string.IsNullOrEmpty(filePath))
                return false;
            return CheckTask_1_10_08_Impl(filePath);
        }

        private bool CheckTask_1_10_01_Impl(string filePath)
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

                // タスク10-1: シート［担当者リスト］の「昇給」の列に、関数で勤続年数が5より大きければ「あり」、そうでなければ「なし」
                worksheet = FindWorksheet(workbook, "担当者リスト");
                if (worksheet == null) return false;

                // G5セルのIF関数をチェック
                Range targetCell = worksheet.Range["G5"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // IF関数とF5>5の条件、「あり」「なし」の値をチェック
                    return normalizedFormula.Contains("IF(") &&
                           normalizedFormula.Contains("F5>5") &&
                           normalizedFormula.Contains("あり") &&
                           normalizedFormula.Contains("なし") &&
                           !normalizedFormula.Contains("$");
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

        private bool CheckTask_1_10_02_Impl(string filePath)
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

                // タスク10-2: シート［出張精算］の「手当金額」の列に、関数で「距離」が300㎞以上であれば「10000」、そうでなければ「5000」
                worksheet = FindWorksheet(workbook, "出張精算");
                if (worksheet == null) return false;

                // G5セルのIF関数をチェック
                Range targetCell = worksheet.Range["G5"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // IF関数とE5>=300の条件、10000、5000の値をチェック
                    return normalizedFormula.Contains("IF(") &&
                           normalizedFormula.Contains("E5>=300") &&
                           normalizedFormula.Contains("10000") &&
                           normalizedFormula.Contains("5000") &&
                           !normalizedFormula.Contains("$");
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

        private bool CheckTask_1_10_03_Impl(string filePath)
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

                // タスク10-3: シート［売上一覧］の「補充の有無」の列に、関数で「在庫」が13%以下であれば「在庫を補充」、そうでなければ何も表示しない
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null) return false;

                // G4セルのIF関数をチェック（G5ではなくG4から開始）
                Range targetCell = worksheet.Range["G4"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // IF関数とF5<=13%または<=0.13の条件、「在庫を補充」の値、空文字列をチェック
                    bool hasIF = normalizedFormula.Contains("IF(");
                    bool hasCondition = normalizedFormula.Contains("F5<=13%") || normalizedFormula.Contains("F5<=0.13") ||
                                       normalizedFormula.Contains("F4<=13%") || normalizedFormula.Contains("F4<=0.13");
                    bool hasMessage = normalizedFormula.Contains("在庫を補充");
                    bool hasEmptyString = normalizedFormula.Contains("\"\"");
                    bool noAbsoluteRef = !normalizedFormula.Contains("$");
                    
                    return hasIF && hasCondition && hasMessage && hasEmptyString && noAbsoluteRef;
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

        private bool CheckTask_1_10_04_Impl(string filePath)
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

                // タスク10-4: シート「担当者リスト」の「No」の列に、関数で1から順に22行分入力
                worksheet = FindWorksheet(workbook, "担当者リスト");
                if (worksheet == null) return false;

                // A5セルのSEQUENCE関数をチェック
                Range targetCell = worksheet.Range["A5"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // SEQUENCE関数と22行、1列、開始値1、目盛り1をチェック
                    return normalizedFormula.Contains("SEQUENCE(") &&
                           normalizedFormula.Contains("22") &&
                           normalizedFormula.Contains("1,1,1");
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

        private bool CheckTask_1_10_05_Impl(string filePath)
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

                // タスク10-5: シート「業務予定」の「開始時間」の列の数式を変更して、10時から30分おきになるようにする
                worksheet = FindWorksheet(workbook, "業務予定");
                if (worksheet == null) return false;

                // C4セルのSEQUENCE関数をチェック（目盛りが0.5になっているか）
                Range targetCell = worksheet.Range["C4"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // SEQUENCE関数と目盛り0.5をチェック
                    return normalizedFormula.Contains("SEQUENCE(") &&
                           normalizedFormula.Contains("0.5");
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

        private bool CheckTask_1_10_06_Impl(string filePath)
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

                // タスク10-6: シート「売上集計」のセル【D6】を開始位置として、関数で表を「売上合計」の高い順に並べ替えて表示
                // シート名が「営業予定」または「売上集計」の可能性があるので両方チェック
                worksheet = FindWorksheet(workbook, "売上集計");
                if (worksheet == null)
                {
                    worksheet = FindWorksheet(workbook, "営業予定");
                }
                if (worksheet == null) return false;

                // D6セルのSORT関数をチェック
                Range targetCell = worksheet.Range["D6"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // SORT関数と並べ替えインデックス2、並べ替え順序-1（降順）をチェック
                    return normalizedFormula.Contains("SORT(") &&
                           normalizedFormula.Contains("A6:B14") &&
                           normalizedFormula.Contains("2") &&
                           normalizedFormula.Contains("-1");
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

        private bool CheckTask_1_10_07_Impl(string filePath)
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

                // タスク10-7: シート「売上一覧」の「税込価格」の列に、「税込価格」を算出（単価と税率の乗算、税率はセル【K4】を参照）
                worksheet = FindWorksheet(workbook, "売上一覧");
                if (worksheet == null) return false;

                // I4セルの数式をチェック（E4*$K$4の絶対参照）
                Range targetCell = worksheet.Range["I4"];
                string formula = targetCell.Formula as string;

                if (formula != null)
                {
                    string normalizedFormula = formula.Replace(" ", "").ToUpper();
                    // E4*$K$4の絶対参照をチェック
                    return normalizedFormula.Contains("E4") &&
                           normalizedFormula.Contains("$K$4") &&
                           normalizedFormula.Contains("*");
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

        private bool CheckTask_1_10_08_Impl(string filePath)
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

                // タスク10-8: シート「在庫管理」のセル「B4」を基準に、テキストファイル「在庫管理表.txt」をインポート
                worksheet = FindWorksheet(workbook, "在庫管理");
                if (worksheet == null) return false;

                // B4付近にQueryTableがあるかチェック
                if (worksheet.QueryTables.Count > 0)
                {
                    foreach (QueryTable qt in worksheet.QueryTables)
                    {
                        // クエリテーブルの接続文字列にテキストファイルが含まれているかチェック
                        string connection = qt.Connection as string;
                        if (connection != null && connection.Contains(".txt"))
                        {
                            // テーブルの開始位置がB4付近かチェック
                            Range destination = qt.Destination;
                            if (destination != null)
                            {
                                if (destination.Address.Contains("$B$4") || destination.Row == 4)
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }

                // ListObjectsもチェック（Power Queryでインポートした場合）
                if (worksheet.ListObjects.Count > 0)
                {
                    foreach (ListObject lo in worksheet.ListObjects)
                    {
                        Range headerRange = lo.HeaderRowRange;
                        if (headerRange != null && headerRange.Row >= 4 && headerRange.Column == 2)
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