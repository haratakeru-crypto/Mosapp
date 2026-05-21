using System;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace New_MOSWordVSTOAddIn
{
    /// <summary>
    /// 採点必須コマンドを project/task 付きで <c>mos_word_log.txt</c> に記録する。
    /// </summary>
    internal static class WordEvidenceHelper
    {
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
            (5, 1, "WrapInline"),
            (5, 2, "WrapSquare"),
            (6, 1, "TableConvertTextToTable"),
            (6, 4, "TableColumnsDistribute"),
            (7, 1, "UpgradeDocument"),
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
                {
                    continue;
                }

                Logger.LogTaskEvidence(entry.ProjectId, entry.TaskId, commandId);
            }
        }

        /// <summary>
        /// 保存後は ActiveDocument が「朗読会.*」となり ProjectN をファイル名から取れない。
        /// 合成コマンドは EvidenceTargets の project/task で記録する。
        /// </summary>
        private static bool UsesFixedProjectWhenActiveUnknown(string commandId)
        {
            return string.Equals(commandId, "FileSaveAsTxt", StringComparison.OrdinalIgnoreCase)
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
