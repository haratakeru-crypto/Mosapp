using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Linq;

// ==========================================
// COM Interop のエイリアス（曖昧さ回避）
// ==========================================
using Excel = Microsoft.Office.Interop.Excel;

// ==========================================
// 追加: Open XML SDK の名前空間
// ==========================================
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet; // 描画(Spreadsheet)
using A = DocumentFormat.OpenXml.Drawing; // 描画(共通)

namespace Libraries.Group1
{
    public class ExcelChecker1_4
    {
        // ==========================================
        // 公開メソッド
        // ==========================================
        public bool CheckTask_1_4_01() => RunCheck(CheckTask_1_4_01_Impl, "Task 4-1 (Sparkline)");
        public bool CheckTask_1_4_02() => RunCheck(CheckTask_1_4_02_Impl, "Task 4-2 (Stacked Col Chart)");
        public bool CheckTask_1_4_03() => RunCheck(CheckTask_1_4_03_Impl, "Task 4-3 (3D Pie Chart)");
        
        // 4-4のみOpenXML実装へ切り替え
        public bool CheckTask_1_4_04() => CheckTask_1_4_04_OpenXml();

        // 共通エラーハンドリング (COM用)
        private bool RunCheck(Func<string, bool> checkImpl, string taskName)
        {
            try
            {
                string filePath = GetCurrentExcelFilePath();
                if (string.IsNullOrEmpty(filePath))
                {
                    Console.WriteLine($"[DEBUG] {taskName}: File path not found.");
                    return false;
                }
                return checkImpl(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Exception in {taskName}: {ex.Message}");
                return false;
            }
        }

        // ==========================================
        // タスク 4-1: 縦棒スパークライン
        // ==========================================
        private bool CheckTask_1_4_01_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "上半期売上", (ws) =>
            {
                Excel.SparklineGroups groups = ws.Cells.SparklineGroups;
                if (groups.Count == 0) return false;

                foreach (Excel.SparklineGroup group in groups)
                {
                    if (group.Type != Excel.XlSparkType.xlSparkColumn) continue;
                    if (group.Count != 6) continue;

                    try
                    {
                        Excel.Range location = group.Location;
                        string addr = location.Address.Replace("$", "").Replace(" ", "");
                        
                        if (!addr.Contains("I5") || !addr.Contains("I10")) continue;
                        
                        string source = group.SourceData.Replace("$", "").Replace(" ", "").ToUpper();
                        if (!source.Contains("B5") || !source.Contains("G10")) continue;

                        Console.WriteLine("[DEBUG] Task 4-1 Passed.");
                        return true;
                    }
                    catch { continue; }
                }

                Console.WriteLine("[DEBUG] Task 4-1 Failed: No matching Sparkline found.");
                return false;
            });
        }

        // ==========================================
        // タスク 4-2: 積み上げ縦棒グラフ
        // ==========================================
        private bool CheckTask_1_4_02_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "5年間売上", (ws) =>
            {
                Excel.ChartObjects charts = (Excel.ChartObjects)ws.ChartObjects();
                if (charts.Count == 0) return false;

                foreach (Excel.ChartObject co in charts)
                {
                    Excel.Chart chart = co.Chart;
                    if (chart.ChartType != Excel.XlChartType.xlColumnStacked) continue;

                    Excel.SeriesCollection seriesColl = (Excel.SeriesCollection)chart.SeriesCollection();
                    if (seriesColl.Count != 2) continue;

                    // 作成時の選択範囲をチェック (プロジェクト4, タスク2)
                    string targetSheet = "5年間売上";
                    string loggedSelection = ExcelLogReader.GetChartCreationSelection(4, 2, 1, chart.Name, targetSheet);
                    if (loggedSelection == null) return false;

                    if (!IsSelectionCorrect(loggedSelection, targetSheet, "A4:C10"))
                    {
                        return false;
                    }

                    return true;
                }

                return false;
            });
        }

        // ==========================================
        // タスク 4-3: 3-D円グラフ
        // ==========================================
        private bool CheckTask_1_4_03_Impl(string filePath)
        {
            return CheckTaskBasic(filePath, "下半期売上", (ws) =>
            {
                Excel.ChartObjects charts = (Excel.ChartObjects)ws.ChartObjects();
                if (charts.Count == 0) return false;

                double boundaryX = ws.Range["I1"].Left;

                foreach (Excel.ChartObject co in charts)
                {
                    Excel.Chart chart = co.Chart;
                    int type = (int)chart.ChartType;

                    bool isPie = (type == -4102) || (type == 5) || (type == -4103) || (type == (int)Excel.XlChartType.xl3DPie);
                    if (!isPie) continue;

                    Excel.SeriesCollection sc = (Excel.SeriesCollection)chart.SeriesCollection();
                    if (sc.Count != 1) continue;

                    if (co.Left > boundaryX + 10)
                    {
                        // 作成時の選択範囲をチェック (プロジェクト4, タスク3)
                        string targetSheet = "下半期売上";
                        string loggedSelection = ExcelLogReader.GetChartCreationSelection(4, 3, 1, chart.Name, targetSheet);
                        if (loggedSelection == null)
                        {
                            Console.WriteLine($"[DEBUG] Task 4-3 Failed: Chart creation log not found for {chart.Name}.");
                            return false;
                        }

                        if (!IsSelectionCorrect(loggedSelection, targetSheet, "A4:A10,H4:H10"))
                        {
                            return false;
                        }

                        return true;
                    }
                }

                return false;
            });
        }

        // ==========================================
        // タスク 4-4: 代替テキスト (Open XML SDK 完全実装版)
        // ==========================================
        private bool CheckTask_1_4_04_OpenXml()
        {
            // COM Interopを使用して、現在開いているExcelファイルのフルパスを取得
            string originalPath = null;
            try
            {
                Excel.Application excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                string targetName = "project4.xlsx";
                
                foreach (Excel.Workbook wb in excelApp.Workbooks)
                {
                    if (wb.Name.Equals(targetName, StringComparison.OrdinalIgnoreCase))
                    {
                        originalPath = wb.FullName;
                        Console.WriteLine($"[DEBUG] Task 4-4: Found file path: {originalPath}");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Task 4-4: Failed to get Excel application. {ex.Message}");
                return false;
            }

            if (string.IsNullOrEmpty(originalPath) || !File.Exists(originalPath))
            {
                Console.WriteLine($"[DEBUG] Task 4-4: File not found. Path: {originalPath}");
                return false;
            }

            // 【重要】Excelが開いているとロックされるため、一時ファイルにコピーして検証する
            string tempPath = Path.GetTempFileName();
            try
            {
                File.Copy(originalPath, tempPath, true);
                Console.WriteLine($"[DEBUG] Task 4-4: Copied to temp file: {tempPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Task 4-4: Copy failed. {ex.Message}");
                return false;
            }

            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(tempPath, false))
                {
                    WorkbookPart wbPart = document.WorkbookPart;
                    // シート名と代替テキストの目標値
                    string targetSheetName = NormalizeString("商品別売上");
                    string targetAltText = NormalizeString("有楽町店の売上グラフ");

                    // 1. シートを特定
                    Sheet sheet = wbPart.Workbook.Descendants<Sheet>()
                        .FirstOrDefault(s => NormalizeString(s.Name).Equals(targetSheetName, StringComparison.OrdinalIgnoreCase));

                    if (sheet == null)
                    {
                        Console.WriteLine($"[DEBUG] Task 4-4 Failed: Sheet '{targetSheetName}' not found.");
                        return false;
                    }

                    // 2. WorksheetPartを取得
                    if (!(wbPart.GetPartById(sheet.Id) is WorksheetPart wsPart)) return false;

                    // 3. DrawingsPartの確認 (グラフはここにある)
                    if (wsPart.DrawingsPart == null)
                    {
                        Console.WriteLine("[DEBUG] Task 4-4 Failed: No DrawingsPart found (No charts/images).");
                        return false;
                    }

                    Xdr.WorksheetDrawing wsDrawing = wsPart.DrawingsPart.WorksheetDrawing;

                    // 4. すべてのGraphicFrame (グラフのコンテナ) を走査
                    // TwoCellAnchor / OneCellAnchor の区別なく取得
                    var graphicFrames = wsDrawing.Descendants<Xdr.GraphicFrame>();

                    foreach (var gf in graphicFrames)
                    {
                        // 5. これが「グラフ」であることをURIで確認
                        // 画像(Picture)ではなく、ChartのURIを持っているか
                        var graphicData = gf.Graphic?.GraphicData;
                        if (graphicData == null || graphicData.Uri != "http://schemas.openxmlformats.org/drawingml/2006/chart")
                        {
                            continue; // グラフ以外はスキップ
                        }

                        // 6. NonVisualDrawingProperties (cNvPr) の取得
                        // 階層: graphicFrame -> nvGraphicFramePr -> cNvPr
                        var nvPr = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties;
                        if (nvPr == null) continue;

                        // 7. Title(タイトル) と Description(説明) の両方を取得して結合
                        // ※Excelのバージョンによってどちらに入るか異なるため両方見る
                        string title = nvPr.Title?.Value ?? "";
                        string description = nvPr.Description?.Value ?? "";
                        
                        string combinedText = title + description; // 連結
                        string normalizedFullText = NormalizeString(combinedText);
                        
                        Console.WriteLine($"[DEBUG] Chart Found (ID:{nvPr.Id}). AltText Content: '{normalizedFullText}'");

                        // 代替テキストが空でないこと、かつ目標テキストと完全一致することを確認
                        if (!string.IsNullOrEmpty(normalizedFullText) && 
                            normalizedFullText.Equals(targetAltText, StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine("[DEBUG] Task 4-4 Passed: Exact match found in OpenXML.");
                            return true;
                        }
                        else
                        {
                            Console.WriteLine($"[DEBUG] Chart (ID:{nvPr.Id}) - Mismatch. Expected: '{targetAltText}', Found: '{normalizedFullText}'");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Task 4-4 OpenXML Error: {ex.Message}");
                return false;
            }
            finally
            {
                // 一時ファイルの削除
                try { File.Delete(tempPath); } catch { }
            }

            Console.WriteLine("[DEBUG] Task 4-4 Failed: Target AltText not found in any chart.");
            return false;
        }

        // ==========================================
        // 共通ヘルパー
        // ==========================================

        private bool CheckTaskBasic(string filePath, string sheetName, Func<Excel.Worksheet, bool> checkLogic)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;
            try
            {
                try { excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { return false; }

                workbook = GetWorkbook(excelApp, filePath);
                if (workbook == null) return false;

                worksheet = FindWorksheetFuzzy(workbook, sheetName);
                if (worksheet == null) return false;

                return checkLogic(worksheet);
            }
            catch { return false; }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
            }
        }

        private Excel.Workbook GetWorkbook(Excel.Application excelApp, string filePath)
        {
            string targetName = "project4.xlsx"; 
            string fileName = Path.GetFileName(filePath); 

            foreach (Excel.Workbook wb in excelApp.Workbooks)
            {
                if (wb.Name.Equals(targetName, StringComparison.OrdinalIgnoreCase) || 
                    wb.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                {
                    return wb;
                }
            }
            return null;
        }

        private Excel.Worksheet FindWorksheetFuzzy(Excel.Workbook workbook, string targetName)
        {
            string normalizedTarget = NormalizeString(targetName);
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (NormalizeString(sheet.Name).Equals(normalizedTarget, StringComparison.OrdinalIgnoreCase))
                    return sheet;
            }
            return null;
        }

        private bool IsSelectionCorrect(string loggedSelection, string targetSheet, string targetAddress)
        {
            if (string.IsNullOrEmpty(loggedSelection)) return false;

            // 1. シート名の検証 (全角半角・引用符無視)
            string normLogged = NormalizeString(loggedSelection).ToUpper();
            string normTargetSheet = NormalizeString(targetSheet).ToUpper();

            if (!normLogged.Contains(normTargetSheet + "!"))
            {
                return false;
            }

            // 2. セル範囲の抽出
            string selection = loggedSelection.Contains("!") ? loggedSelection.Substring(loggedSelection.IndexOf("!") + 1) : loggedSelection;
            
            // 絶対参照の $ を削除
            selection = selection.Replace("$", "");

            string normSelection = NormalizeString(selection).ToUpper();
            string normTargetAddr = NormalizeString(targetAddress).ToUpper();

            // カンマ区切りの複数範囲に対応するため、各エリアをソートして比較
            var loggedAreas = normSelection.Split(',').OrderBy(a => a).ToList();
            var targetAreas = normTargetAddr.Split(',').OrderBy(a => a).ToList();

            if (loggedAreas.Count != targetAreas.Count) return false;

            for (int i = 0; i < loggedAreas.Count; i++)
            {
                if (loggedAreas[i] != targetAreas[i]) return false;
            }

            return true;
        }

        private string NormalizeString(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            
            // 全角→半角変換
            char[] chars = input.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] >= '０' && chars[i] <= '９')
                {
                    chars[i] = (char)(chars[i] - '０' + '0');
                }
                else if (chars[i] >= 'Ａ' && chars[i] <= 'Ｚ')
                {
                    chars[i] = (char)(chars[i] - 'Ａ' + 'A');
                }
                else if (chars[i] >= 'ａ' && chars[i] <= 'ｚ')
                {
                    chars[i] = (char)(chars[i] - 'ａ' + 'a');
                }
            }
            string normalized = new string(chars).ToUpper().Replace("'", "");
            
            // 空白・制御文字削除
            return Regex.Replace(normalized, @"\s+", "");
        }

        private string GetCurrentExcelFilePath()
        {
            // プロジェクトの構造に合わせてパスを返してください
            // 例: return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "project4.xlsx");
            return "project4.xlsx"; 
        }
    }
}
