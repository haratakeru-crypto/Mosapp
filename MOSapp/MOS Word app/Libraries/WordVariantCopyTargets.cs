using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MOS_Word_app
{
    /// <summary>
    /// 類題で問題文の下線・青文字クリックコピーが付く語句／タスク。
    /// 破壊ゲートの Paste 許可と UI 表示で共有する。
    /// </summary>
    public static class WordVariantCopyTargets
    {
        /// <summary>xlsx の「」内を "…" にして下線・青文字にする明示リスト（キー: "P-T"）。</summary>
        public static readonly IReadOnlyDictionary<string, string[]> Phrases =
            new Dictionary<string, string[]>(StringComparer.Ordinal)
            {
                ["1-1"] = new[] { "写真展案内完成" },
                ["1-4"] = new[] { "人物を主題にした写真のこと" },
                ["4-4"] = new[] { "開場時刻を再確認" },
                ["5-3"] = new[] { "鉢植え" },
                ["6-1"] = new[] { "star@example.jp" },
                ["6-2"] = new[] { "夜の部は19時開始です" },
                ["6-4"] = new[] { "ぜひ" },
                ["7-4"] = new[] { "記録" },
                ["8-1"] = new[] { "食器" }
            };

        private static readonly Regex InputQuotePattern = new Regex(
            @"(?:「([^」]*)」(?:、[^「]*?)?)+と入力",
            RegexOptions.Compiled);

        /// <summary>クリックコピー対象なら Paste を破壊ゲートで許可する。</summary>
        public static bool AllowsPaste(int projectId, int taskId)
        {
            if (projectId <= 0 || taskId <= 0)
                return false;

            string key = projectId + "-" + taskId;
            if (Phrases.ContainsKey(key))
                return true;

            foreach (VariantTaskText task in VariantProblemWorkbook.GetTasks(projectId))
            {
                if (task.TaskId != taskId)
                    continue;
                return DescriptionHasCopyableQuotes(task.Description);
            }

            return false;
        }

        public static bool TryGetPhrases(int projectId, int taskId, out string[] phrases)
        {
            return Phrases.TryGetValue(projectId + "-" + taskId, out phrases);
        }

        /// <summary>「…」と入力 パターン、または既に "…" がある問題文。</summary>
        public static bool DescriptionHasCopyableQuotes(string description)
        {
            if (string.IsNullOrEmpty(description))
                return false;
            if (description.IndexOf('"') >= 0)
                return true;
            return InputQuotePattern.IsMatch(description);
        }
    }
}
