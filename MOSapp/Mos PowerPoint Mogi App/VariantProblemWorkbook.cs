using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace MOS_PowerPoint_app
{
    /// <summary>類題1の問題文と解答手順（PowerPoint類題.xlsx の1行）。</summary>
    public sealed class VariantTaskText
    {
        public int ProjectId { get; set; }
        public int TaskId { get; set; }
        public string Description { get; set; }
        public string AnswerSteps { get; set; }
    }

    /// <summary>
    /// 類題の問題文と解答手順を xlsx から読む。類題セット1の行だけが入っている。
    /// </summary>
    public static class VariantProblemWorkbook
    {
        public const string FileName = "PowerPoint類題.xlsx";

        private static readonly object CacheGate = new object();
        private static string _cachedPath;
        private static DateTime _cachedWriteTimeUtc;
        private static List<VariantTaskText> _cachedRows = new List<VariantTaskText>();

        public static string ResolvePath()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            string referencesPath = Path.Combine(baseDirectory, "References", FileName);
            if (File.Exists(referencesPath))
                return referencesPath;

            string flatPath = Path.Combine(baseDirectory, FileName);
            return File.Exists(flatPath) ? flatPath : referencesPath;
        }

        /// <summary>起動時に xlsx を先読みし、アプリバー生成で OpenXml が走らないようにする。</summary>
        public static void Preload()
        {
            try
            {
                LoadRows();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[VariantProblemWorkbook] preload failed: {ex.Message}");
            }
        }

        public static IReadOnlyList<int> GetProjectIds()
        {
            return LoadRows()
                .Select(row => row.ProjectId)
                .Distinct()
                .OrderBy(id => id)
                .ToList();
        }

        public static int GetTaskCount(int projectId)
        {
            return LoadRows().Count(row => row.ProjectId == projectId);
        }

        public static IReadOnlyList<VariantTaskText> GetTasks(int projectId)
        {
            return LoadRows()
                .Where(row => row.ProjectId == projectId)
                .OrderBy(row => row.TaskId)
                .ToList();
        }

        private static List<VariantTaskText> LoadRows()
        {
            string path = ResolvePath();
            DateTime writeTimeUtc = File.Exists(path)
                ? File.GetLastWriteTimeUtc(path)
                : DateTime.MinValue;

            lock (CacheGate)
            {
                if (string.Equals(_cachedPath, path, StringComparison.OrdinalIgnoreCase) &&
                    _cachedWriteTimeUtc == writeTimeUtc)
                {
                    return _cachedRows;
                }

                _cachedPath = path;
                _cachedWriteTimeUtc = writeTimeUtc;
                _cachedRows = ReadRows(path);
                return _cachedRows;
            }
        }

        private static List<VariantTaskText> ReadRows(string path)
        {
            var rows = new List<VariantTaskText>();
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                System.Diagnostics.Debug.WriteLine($"[VariantProblemWorkbook] file not found: {path}");
                return rows;
            }

            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var document = SpreadsheetDocument.Open(stream, false))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;
                    Sheet sheet = workbookPart?.Workbook?.Sheets?.Elements<Sheet>().FirstOrDefault();
                    if (sheet == null || !(workbookPart.GetPartById(sheet.Id) is WorksheetPart worksheetPart))
                        return rows;

                    var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable
                        ?.Elements<SharedStringItem>()
                        .Select(item => item.InnerText ?? string.Empty)
                        .ToList();

                    bool headerFound = false;
                    int projectColumn = 0;
                    int taskColumn = 0;
                    int descriptionColumn = 0;
                    int answerStepsColumn = 0;

                    foreach (Row row in worksheetPart.Worksheet.Descendants<Row>())
                    {
                        var cells = ReadRow(row, sharedStrings);
                        if (!headerFound)
                        {
                            if (!TryMapHeaders(cells, out projectColumn, out taskColumn, out descriptionColumn, out answerStepsColumn))
                                continue;

                            headerFound = true;
                            continue;
                        }

                        if (!TryParseId(CellText(cells, projectColumn), out int projectId) ||
                            !TryParseId(CellText(cells, taskColumn), out int taskId))
                        {
                            continue;
                        }

                        rows.Add(new VariantTaskText
                        {
                            ProjectId = projectId,
                            TaskId = taskId,
                            Description = CellText(cells, descriptionColumn),
                            AnswerSteps = CellText(cells, answerStepsColumn)
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[VariantProblemWorkbook] read failed: {ex.Message}");
                rows.Clear();
            }

            return rows;
        }

        private static bool TryMapHeaders(
            Dictionary<int, string> cells,
            out int projectColumn,
            out int taskColumn,
            out int descriptionColumn,
            out int answerStepsColumn)
        {
            projectColumn = 0;
            taskColumn = 0;
            descriptionColumn = 0;
            answerStepsColumn = 0;

            foreach (var pair in cells)
            {
                string header = (pair.Value ?? string.Empty).Trim();
                if (header == "プロジェクト")
                    projectColumn = pair.Key;
                else if (header == "タスク")
                    taskColumn = pair.Key;
                else if (header == "問題文")
                    descriptionColumn = pair.Key;
                else if (header == "解答手順")
                    answerStepsColumn = pair.Key;
            }

            return projectColumn > 0 && taskColumn > 0 && descriptionColumn > 0 && answerStepsColumn > 0;
        }

        private static Dictionary<int, string> ReadRow(Row row, IList<string> sharedStrings)
        {
            var cells = new Dictionary<int, string>();
            foreach (Cell cell in row.Elements<Cell>())
            {
                int column = ColumnIndex(cell.CellReference);
                if (column <= 0)
                    continue;
                cells[column] = GetCellText(cell, sharedStrings);
            }

            return cells;
        }

        private static string CellText(Dictionary<int, string> cells, int column)
        {
            return cells.TryGetValue(column, out string text) ? text ?? string.Empty : string.Empty;
        }

        private static string GetCellText(Cell cell, IList<string> sharedStrings)
        {
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                if (int.TryParse(cell.CellValue?.InnerText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index) &&
                    sharedStrings != null &&
                    index >= 0 &&
                    index < sharedStrings.Count)
                {
                    return sharedStrings[index] ?? string.Empty;
                }

                return string.Empty;
            }

            if (cell.DataType != null && cell.DataType.Value == CellValues.InlineString)
                return cell.InlineString?.InnerText ?? string.Empty;

            return cell.CellValue?.InnerText ?? string.Empty;
        }

        private static int ColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
                return 0;

            int index = 0;
            foreach (char ch in cellReference)
            {
                if (!char.IsLetter(ch))
                    break;
                index = (index * 26) + (char.ToUpperInvariant(ch) - 'A' + 1);
            }

            return index;
        }

        private static bool TryParseId(string text, out int id)
        {
            id = 0;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            text = text.Trim();
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out id))
                return id > 0;

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
            {
                id = (int)Math.Round(number);
                return id > 0 && Math.Abs(number - id) < 0.001;
            }

            return false;
        }
    }
}

