using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace New_MOSWordVSTOAddIn
{
    /// <summary>
    /// 採点必須コマンドを project/task 付きで <c>mos_word_log.txt</c> に記録する。
    /// </summary>
    internal static class WordEvidenceHelper
    {
        private const string Task2_1CutTargetText = "青空文庫のURLはコチラ↓";

        private static readonly (int ProjectId, int TaskId, string CommandId)[] EvidenceTargets =
        {
            (1, 1, "ShowAll"),
            // 1-1-5: 環境により FontClearFormatting 等の idMso が無効のため、Ribbon では ClearFormatting のみフック
            (1, 5, "ClearFormatting"),
            (2, 1, "Cut"),
            (2, 1, "Paste"),
            (3, 1, "PageMarginsModerate"),
            (3, 3, "PageOrientationPortraitLandscape"),
            (4, 3, "ReviewDeleteComment"),
            (4, 3, "ReviewResolveComment"),
            (4, 4, "StyleSetLineSimple"),
            (4, 5, "Watermark"),
            (4, 5, "WatermarkMenu"),
            (4, 5, "GalleryWatermark"),
            (4, 5, "WatermarkCustomDialog"),
            (4, 6, "PageBorders"),
            (5, 1, "WrapTopBottom"),
            (5, 2, "WrapTight"),
            (6, 1, "TableConvertTextToTable"),
            (6, 4, "TableColumnsDistribute"),
            (7, 1, "UpgradeDocument"),
            (7, 2, "SetDocumentCompany"),
            (7, 3, "IntegralHeader"),
            (7, 4, "FileSaveAsTxt"),
            (7, 5, "FileSaveAsDocm"),
            (9, 5, "AcceptAllChangesInDocAndStopTracking"),
            (10, 2, "BulletDefineNew"),
            (10, 3, "BulletDefineNew"),
        };

        public static void LogCommandWithEvidence(string commandId)
        {
            int activeProjectId = TryGetActiveProjectId();

            foreach (var entry in EvidenceTargets)
            {
                if (!string.Equals(entry.CommandId, commandId, StringComparison.OrdinalIgnoreCase))
                    continue;

                if (activeProjectId > 0)
                {
                    if (entry.ProjectId != activeProjectId)
                        continue;
                }
                else if (!UsesFixedProjectWhenActiveUnknown(commandId))
                    continue;

                Logger.LogTaskEvidence(entry.ProjectId, entry.TaskId, commandId);
            }
        }

        /// <summary>2-1: 切り取り選択を記録し、段落単位なら CutParagraphSelection 証跡を付ける。記録した場合 true（通常 Cut は付けない）。</summary>
        public static bool TryLogInvalidParagraphCut()
        {
            try
            {
                int activeProjectId = TryGetActiveProjectId();
                if (activeProjectId > 0 && activeProjectId != 2)
                    return false;

                var app = Globals.ThisAddIn?.Application;
                if (app?.Selection == null)
                    return false;

                CutSelectionAnalysis analysis = AnalyzeCutSelection(app.Selection);
                string detail = "paraCount=" + analysis.ParaCount
                    + "|wholePara=" + (analysis.WholeParagraph ? "1" : "0")
                    + "|selEnd=" + analysis.SelEnd
                    + "|paraEnd=" + analysis.ParaEnd
                    + "|targetEnd=" + analysis.TargetEnd
                    + "|sel=" + (analysis.SelText ?? "").Replace("|", "/");
                Logger.LogOperation("CutSelection", detail);

                if (!analysis.IsParagraphCut)
                    return false;

                Logger.LogTaskEvidence(2, 1, "CutParagraphSelection");
                return true;
            }
            catch { }

            return false;
        }

        private struct CutSelectionAnalysis
        {
            public bool IsParagraphCut;
            public int ParaCount;
            public bool WholeParagraph;
            public string SelText;
            public int SelStart;
            public int SelEnd;
            public int ParaStart;
            public int ParaEnd;
            public int TargetStart;
            public int TargetEnd;
        }

        private static CutSelectionAnalysis AnalyzeCutSelection(Selection sel)
        {
            var a = new CutSelectionAnalysis
            {
                SelText = NormalizeCutSelectionText(sel.Text),
                ParaCount = sel.Paragraphs.Count
            };

            if (a.ParaCount >= 2)
            {
                a.IsParagraphCut = true;
                return a;
            }

            if (a.ParaCount != 1)
                return a;

            Paragraph para = sel.Paragraphs[1];
            Range paraRange = null;
            try
            {
                paraRange = para.Range;
                Range selRange = sel.Range;
                a.ParaStart = paraRange.Start;
                a.ParaEnd = paraRange.End;
                a.SelStart = selRange.Start;
                a.SelEnd = selRange.End;
                a.WholeParagraph = a.SelStart == a.ParaStart && a.SelEnd == a.ParaEnd;

                if (TryGetTargetTextRangeInParagraph(paraRange, out int tStart, out int tEnd))
                {
                    a.TargetStart = tStart;
                    a.TargetEnd = tEnd;
                }

                if (a.WholeParagraph)
                {
                    a.IsParagraphCut = true;
                    return a;
                }

                if (a.TargetEnd > 0 && a.SelStart == a.TargetStart && a.SelEnd == a.TargetEnd)
                {
                    a.IsParagraphCut = false;
                    return a;
                }

                if (string.Equals(a.SelText, Task2_1CutTargetText, StringComparison.Ordinal) && a.SelEnd < a.ParaEnd)
                {
                    a.IsParagraphCut = false;
                    return a;
                }

                if (a.TargetEnd > 0 && a.SelEnd > a.TargetEnd)
                {
                    a.IsParagraphCut = true;
                    return a;
                }

                if ((a.SelText ?? "").Length > Task2_1CutTargetText.Length)
                {
                    a.IsParagraphCut = true;
                    return a;
                }
            }
            finally
            {
                if (paraRange != null) Marshal.ReleaseComObject(paraRange);
                Marshal.ReleaseComObject(para);
            }

            return a;
        }

        private static bool TryGetTargetTextRangeInParagraph(Range paraRange, out int targetStart, out int targetEnd)
        {
            targetStart = 0;
            targetEnd = 0;
            Range search = null;
            try
            {
                search = paraRange.Duplicate;
                Find find = search.Find;
                find.ClearFormatting();
                find.Text = Task2_1CutTargetText;
                find.Forward = true;
                find.Wrap = WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                if (!find.Execute())
                    return false;
                targetStart = search.Start;
                targetEnd = search.End;
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (search != null) Marshal.ReleaseComObject(search);
            }
        }

        private static string NormalizeCutSelectionText(string text)
        {
            return (text ?? "").TrimEnd('\r', '\n', '\a', '\v');
        }

        /// <summary>
        /// 保存後は ActiveDocument が「朗読会.*」となり ProjectN をファイル名から取れない。
        /// 合成コマンドは EvidenceTargets の project/task で記録する。
        /// </summary>
        private static bool UsesFixedProjectWhenActiveUnknown(string commandId)
        {
            return string.Equals(commandId, "ShowAll", StringComparison.OrdinalIgnoreCase)
                || string.Equals(commandId, "FileSaveAsTxt", StringComparison.OrdinalIgnoreCase)
                || string.Equals(commandId, "FileSaveAsDocm", StringComparison.OrdinalIgnoreCase);
        }

        private static int TryGetActiveProjectId()
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app?.ActiveDocument == null)
                    return 0;

                string fullName;
                try { fullName = app.ActiveDocument.FullName; }
                catch { return 0; }

                if (string.IsNullOrEmpty(fullName))
                    return 0;

                string name = Path.GetFileNameWithoutExtension(fullName);
                var m = Regex.Match(name, @"project\s*(\d+)", RegexOptions.IgnoreCase);
                if (m.Success && int.TryParse(m.Groups[1].Value, out int id))
                    return id;
            }
            catch { }

            return 0;
        }
    }
}
