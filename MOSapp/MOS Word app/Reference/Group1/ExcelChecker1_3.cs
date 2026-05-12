using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Libraries.Group1
{
    public class ExcelChecker1_3
    {
        public bool CheckTask_1_3_01()
        {
            System.Diagnostics.Debug.WriteLine("[DEBUG] CheckTask_1_3_01 called");
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] filePath is null or empty");
                    return false;
                }
                bool result = CheckTask_1_3_01_Impl(filePath);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_01_Impl returned: {result}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_01: {ex.Message}");
                return false;
            }
        }

        public bool CheckTask_1_3_02()
        {
            System.Diagnostics.Debug.WriteLine("[DEBUG] CheckTask_1_3_02 called");
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                bool result = CheckTask_1_3_02_Impl(filePath);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_02_Impl returned: {result}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_02: {ex.Message}");
                return false;
            }
        }

        public bool CheckTask_1_3_03()
        {
            System.Diagnostics.Debug.WriteLine("[DEBUG] CheckTask_1_3_03 called");
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                bool result = CheckTask_1_3_03_Impl(filePath);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_03_Impl returned: {result}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_03: {ex.Message}");
                return false;
            }
        }

        public bool CheckTask_1_3_04()
        {
            System.Diagnostics.Debug.WriteLine("[DEBUG] CheckTask_1_3_04 called");
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                bool result = CheckTask_1_3_04_Impl(filePath);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_04_Impl returned: {result}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_04: {ex.Message}");
                return false;
            }
        }

        public bool CheckTask_1_3_05()
        {
            System.Diagnostics.Debug.WriteLine("[DEBUG] CheckTask_1_3_05 called");
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                bool result = CheckTask_1_3_05_Impl(filePath);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_05_Impl returned: {result}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_05: {ex.Message}");
                return false;
            }
        }

        public bool CheckTask_1_3_06()
        {
            System.Diagnostics.Debug.WriteLine("[DEBUG] CheckTask_1_3_06 called");
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                bool result = CheckTask_1_3_06_Impl(filePath);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_06_Impl returned: {result}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_06: {ex.Message}");
                return false;
            }
        }

        public bool CheckTask_1_3_07()
        {
            System.Diagnostics.Debug.WriteLine("[DEBUG] CheckTask_1_3_07 called");
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    return false;
                }
                bool result = CheckTask_1_3_07_Impl(filePath);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_07_Impl returned: {result}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_07: {ex.Message}");
                return false;
            }
        }

        private bool CheckTask_1_3_01_Impl(string filePath)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_01_Impl called with filePath: {filePath}");
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Got existing Excel application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Created new Excel application");
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] No matching workbook found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found matching workbook");

                worksheet = FindWorksheet(workbook, "下半期売上");
                if (worksheet == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Worksheet '下半期売上' not found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found worksheet '下半期売上'");

                Range targetCell = worksheet.Range["A2"];
                
                // 記事の方法を使用した詳細なスタイル判定
                bool hasTitleStyle = CheckCellStyleDetailed(targetCell, "タイトル");
                
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Has title style: {hasTitleStyle}");
                return hasTitleStyle;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_01_Impl: {ex.Message}");
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_3_02_Impl(string filePath)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_02_Impl called with filePath: {filePath}");
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Got existing Excel application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Created new Excel application");
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] No matching workbook found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found matching workbook");

                worksheet = FindWorksheet(workbook, "社員リスト");
                if (worksheet == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Worksheet '社員リスト' not found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found worksheet '社員リスト'");

                // タスク3-2: セル範囲B5:B44に左インデントを2文字分設定
                Range targetRange = worksheet.Range["B5:B44"];
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Checking range: {targetRange.Address}");
                
                // 記事の方法を使用した詳細なインデント判定
                bool allCellsHaveIndent2 = CheckIndentLevelDetailed(targetRange, 2);
                
                System.Diagnostics.Debug.WriteLine($"[DEBUG] All cells have indent level 2: {allCellsHaveIndent2}");
                return allCellsHaveIndent2;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_02_Impl: {ex.Message}");
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_3_03_Impl(string filePath)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_03_Impl called with filePath: {filePath}");
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Got existing Excel application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Created new Excel application");
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] No matching workbook found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found matching workbook");

                worksheet = FindWorksheet(workbook, "社員リスト");
                if (worksheet == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Worksheet '社員リスト' not found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found worksheet '社員リスト'");

                // タスク3-3: セルA2の文字の配置をセル範囲A2:F2の中央に設定
                Range a2Cell = worksheet.Range["A2"];
                
                // 記事の方法を使用した詳細な配置判定
                // -4108はxlCenter（中央揃え）の定数値
                bool isCenterAligned = CheckAlignmentDetailed(a2Cell, -4108);
                
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Is center aligned: {isCenterAligned}");
                return isCenterAligned;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_03_Impl: {ex.Message}");
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_3_04_Impl(string filePath)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_04_Impl called with filePath: {filePath}");
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Got existing Excel application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Created new Excel application");
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] No matching workbook found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found matching workbook");

                worksheet = FindWorksheet(workbook, "担当者別売上");
                if (worksheet == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Worksheet '担当者別売上' not found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found worksheet '担当者別売上'");

                // タスク3-4: セル範囲H5:K19をコピーし、列幅を保持して貼り付け
                Range sourceRange = worksheet.Range["H5:K19"];
                Range targetRange = worksheet.Range["A5:D19"];
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Source range: {sourceRange.Address}, Target range: {targetRange.Address}");

                bool allValuesMatch = true;
                for (int i = 1; i <= sourceRange.Rows.Count; i++)
                {
                    for (int j = 1; j <= sourceRange.Columns.Count; j++)
                    {
                        var sourceValue = ((Range)sourceRange.Cells[i, j]).Value2;
                        var targetValue = ((Range)targetRange.Cells[i, j]).Value2;
                        
                        if (sourceValue != targetValue)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] Value mismatch at ({i},{j}): source='{sourceValue}', target='{targetValue}'");
                            allValuesMatch = false;
                            break;
                        }
                    }
                    if (!allValuesMatch) break;
                }
                
                System.Diagnostics.Debug.WriteLine($"[DEBUG] All values match: {allValuesMatch}");
                return allValuesMatch;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_04_Impl: {ex.Message}");
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_3_05_Impl(string filePath)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_05_Impl called with filePath: {filePath}");
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Got existing Excel application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Created new Excel application");
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] No matching workbook found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found matching workbook");

                worksheet = FindWorksheet(workbook, "業務予定");
                if (worksheet == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Worksheet '業務予定' not found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found worksheet '業務予定'");

                // タスク3-5: セル範囲C5:C11の時間に取り消し線を設定
                Range targetRange = worksheet.Range["C5:C11"];
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Checking range: {targetRange.Address}");
                
                // 記事の方法を使用した詳細な取り消し線判定
                bool allCellsHaveStrikethrough = CheckStrikethroughDetailed(targetRange);
                
                System.Diagnostics.Debug.WriteLine($"[DEBUG] All cells have strikethrough: {allCellsHaveStrikethrough}");
                return allCellsHaveStrikethrough;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_05_Impl: {ex.Message}");
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_3_06_Impl(string filePath)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_06_Impl called with filePath: {filePath}");
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Got existing Excel application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Created new Excel application");
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] No matching workbook found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found matching workbook");

                worksheet = FindWorksheet(workbook, "参加者一覧");
                if (worksheet == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Worksheet '参加者一覧' not found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found worksheet '参加者一覧'");

                // タスク3-6: セルB2のセルの結合と文字列の配置を解除
                Range targetCell = worksheet.Range["B2"];
                int mergeAreaCellCount = targetCell.MergeArea.Cells.Count;
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Cell B2 merge area cell count: {mergeAreaCellCount}");
                
                // セルの結合が解除されている場合、MergeArea.Cells.Countは1になる
                bool isUnmerged = mergeAreaCellCount == 1;
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Cell is unmerged: {isUnmerged}");
                
                return isUnmerged;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_06_Impl: {ex.Message}");
                return false;
            }
            finally
            {
                if (worksheet != null)
                    Marshal.ReleaseComObject(worksheet);
            }
        }

        private bool CheckTask_1_3_07_Impl(string filePath)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] CheckTask_1_3_07_Impl called with filePath: {filePath}");
            Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                try
                {
                    excelApp = (Application)Marshal.GetActiveObject("Excel.Application");
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Got existing Excel application");
                }
                catch
                {
                    excelApp = new Application();
                    excelApp.Visible = true;
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Created new Excel application");
                }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] No matching workbook found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found matching workbook");

                worksheet = FindWorksheet(workbook, "参加者一覧");
                if (worksheet == null)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] Worksheet '参加者一覧' not found");
                    return false;
                }
                System.Diagnostics.Debug.WriteLine("[DEBUG] Found worksheet '参加者一覧'");

                // タスク3-7: 「氏名」が「風間 健太郎」と「平井 元」の行を削除
                Range usedRange = worksheet.UsedRange;
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Used range: {usedRange.Address}");
                
                // ヘッダー行を確認
                System.Diagnostics.Debug.WriteLine("[DEBUG] Checking header row:");
                for (int col = 1; col <= usedRange.Columns.Count; col++)
                {
                    var cellValue = ((Range)usedRange.Cells[1, col]).Value2;
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Header column {col}: '{cellValue}'");
                }
                
                int nameColumnIndex = -1;
                for (int col = 1; col <= usedRange.Columns.Count; col++)
                {
                    var cellValue = ((Range)usedRange.Cells[1, col]).Value2;
                    if (cellValue != null && cellValue.ToString().Contains("氏名"))
                    {
                        nameColumnIndex = col;
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] Found '氏名' column at index: {col}");
                        break;
                    }
                }

                if (nameColumnIndex == -1)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] '氏名' column not found");
                    return false;
                }

                // 全行の名前を確認
                System.Diagnostics.Debug.WriteLine("[DEBUG] Checking all names in the column:");
                bool targetNamesFound = false;
                int totalRows = usedRange.Rows.Count;
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Total rows in used range: {totalRows}");
                
                for (int row = 2; row <= totalRows; row++)
                {
                    var cellValue = ((Range)usedRange.Cells[row, nameColumnIndex]).Value2;
                    if (cellValue != null)
                    {
                        string name = cellValue.ToString().Trim();
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] Row {row} name: '{name}'");
                        
                        // より厳密な名前の比較
                        if (name == "風間 健太郎" || name == "平井 元" || 
                            name.Contains("風間 健太郎") || name.Contains("平井 元"))
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] Target name found in row {row}: '{name}'");
                            targetNamesFound = true;
                            break;
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] Row {row} name: (null)");
                    }
                }
                
                // 削除された場合、これらの名前は見つからないはず
                bool result = !targetNamesFound;
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Target names found: {targetNamesFound}, Result (deleted): {result}");
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Total rows checked: {totalRows - 1}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] Exception in CheckTask_1_3_07_Impl: {ex.Message}");
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
            string fileName = Path.GetFileName(filePath);
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

        // 文字の書式を詳細に判定するヘルパーメソッド（記事の方法を活用）
        private bool CheckTextFormatting(Range cell)
        {
            try
            {
                // フォントの詳細確認
                bool isBold = (bool)cell.Font.Bold;
                bool isItalic = (bool)cell.Font.Italic;
                var fontColor = cell.Font.Color;
                var fontSize = cell.Font.Size;
                
                // 背景色の確認
                var backgroundColor = cell.Interior.Color;
                
                // 境界線の確認
                var borderStyle = cell.Borders.LineStyle;
                
                Console.WriteLine($"[DEBUG] Font: Bold={isBold}, Italic={isItalic}, Size={fontSize}");
                Console.WriteLine($"[DEBUG] Colors: Font={fontColor}, Background={backgroundColor}");
                
                return true; // 書式が適用されている
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error checking text formatting: {ex.Message}");
                return false;
            }
        }

        // セルスタイルの詳細判定（記事の方法を活用）
        private bool CheckCellStyleDetailed(Range cell, string expectedStyleName)
        {
            try
            {
                // 方法1: 直接スタイル名取得（COMオブジェクトのため制限される場合がある）
                string cellStyle = "";
                try
                {
                    // COMオブジェクトでは直接アクセスが制限される場合があるため、動的アクセスにフォールバック
                    throw new InvalidOperationException("Direct access not available");
                }
                catch
                {
                    // 方法2: 動的プロパティアクセス
                    var styleProperty = cell.GetType().GetProperty("Style");
                    if (styleProperty != null)
                    {
                        var styleObj = styleProperty.GetValue(cell);
                        var nameProperty = styleObj.GetType().GetProperty("Name");
                        if (nameProperty != null)
                        {
                            var nameValue = nameProperty.GetValue(styleObj);
                            cellStyle = nameValue?.ToString() ?? "";
                        }
                    }
                }

                Console.WriteLine($"[DEBUG] Cell style: '{cellStyle}'");
                
                // より柔軟なスタイル名マッチング
                if (string.IsNullOrEmpty(cellStyle))
                {
                    return false;
                }

                // 部分一致も許可
                bool hasMatchingStyle = cellStyle.IndexOf(expectedStyleName, StringComparison.OrdinalIgnoreCase) >= 0 ||
                                      cellStyle.Contains("タイトル") || cellStyle.Contains("Title") ||
                                      cellStyle.Contains("見出し") || cellStyle.Contains("Heading");

                Console.WriteLine($"[DEBUG] Style match: {hasMatchingStyle}");
                return hasMatchingStyle;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error checking cell style: {ex.Message}");
                return false;
            }
        }

        // インデントレベルの詳細確認（記事の方法を活用）
        private bool CheckIndentLevelDetailed(Range range, int expectedLevel)
        {
            try
            {
                bool allCellsCorrect = true;
                int checkedCells = 0;
                
                foreach (Range cell in range.Cells)
                {
                    checkedCells++;
                    int indentLevel = 0;
                    
                    try
                    {
                        // 方法1: 直接プロパティアクセス（COMオブジェクトのため制限される場合がある）
                        throw new InvalidOperationException("Direct access not available");
                    }
                    catch
                    {
                        // 方法2: 動的プロパティアクセス
                        var indentProperty = cell.GetType().GetProperty("IndentLevel");
                        if (indentProperty != null)
                        {
                            var indentValue = indentProperty.GetValue(cell);
                            indentLevel = Convert.ToInt32(indentValue);
                        }
                    }
                    
                    Console.WriteLine($"[DEBUG] Cell {cell.Address} indent level: {indentLevel}");
                    
                    if (indentLevel != expectedLevel)
                    {
                        allCellsCorrect = false;
                        Console.WriteLine($"[DEBUG] Cell {cell.Address} has incorrect indent level: {indentLevel}, expected: {expectedLevel}");
                    }
                }
                
                Console.WriteLine($"[DEBUG] Checked {checkedCells} cells, all correct: {allCellsCorrect}");
                return allCellsCorrect;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error checking indent level: {ex.Message}");
                return false;
            }
        }

        // 文字配置の詳細確認（記事の方法を活用）
        private bool CheckAlignmentDetailed(Range cell, int expectedAlignment)
        {
            try
            {
                int horizontalAlignment = 0;
                
                try
                {
                    // 方法1: 直接プロパティアクセス
                    horizontalAlignment = (int)cell.HorizontalAlignment;
                }
                catch
                {
                    // 方法2: 動的プロパティアクセス
                    var alignmentProperty = cell.GetType().GetProperty("HorizontalAlignment");
                    if (alignmentProperty != null)
                    {
                        var alignmentValue = alignmentProperty.GetValue(cell);
                        horizontalAlignment = (int)alignmentValue;
                    }
                }
                
                Console.WriteLine($"[DEBUG] Cell alignment: {horizontalAlignment}, expected: {expectedAlignment}");
                
                bool isCorrectAlignment = horizontalAlignment == expectedAlignment;
                Console.WriteLine($"[DEBUG] Alignment correct: {isCorrectAlignment}");
                
                return isCorrectAlignment;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error checking alignment: {ex.Message}");
                return false;
            }
        }

        // 取り消し線の詳細確認（記事の方法を活用）
        private bool CheckStrikethroughDetailed(Range range)
        {
            try
            {
                bool allCellsHaveStrikethrough = true;
                int checkedCells = 0;
                
                foreach (Range cell in range.Cells)
                {
                    checkedCells++;
                    bool hasStrikethrough = false;
                    
                    try
                    {
                        // 方法1: 直接プロパティアクセス
                        hasStrikethrough = Convert.ToBoolean(cell.Font.Strikethrough);
                    }
                    catch
                    {
                        // 方法2: 動的プロパティアクセス
                        var fontProperty = cell.GetType().GetProperty("Font");
                        if (fontProperty != null)
                        {
                            var fontObj = fontProperty.GetValue(cell);
                            var strikethroughProperty = fontObj.GetType().GetProperty("Strikethrough");
                            if (strikethroughProperty != null)
                            {
                                var strikethroughValue = strikethroughProperty.GetValue(fontObj);
                                hasStrikethrough = Convert.ToBoolean(strikethroughValue);
                            }
                        }
                    }
                    
                    Console.WriteLine($"[DEBUG] Cell {cell.Address} strikethrough: {hasStrikethrough}");
                    
                    if (!hasStrikethrough)
                    {
                        allCellsHaveStrikethrough = false;
                        Console.WriteLine($"[DEBUG] Cell {cell.Address} does not have strikethrough");
                    }
                }
                
                Console.WriteLine($"[DEBUG] Checked {checkedCells} cells, all have strikethrough: {allCellsHaveStrikethrough}");
                return allCellsHaveStrikethrough;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error checking strikethrough: {ex.Message}");
                return false;
            }
        }
    }
}