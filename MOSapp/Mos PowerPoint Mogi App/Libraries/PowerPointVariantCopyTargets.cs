using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MOS_PowerPoint_app
{
    /// <summary>
    /// 類題で問題文の下線・青文字クリックコピーが付く語句／タスク。
    /// </summary>
    public static class PowerPointVariantCopyTargets
    {
        /// <summary>xlsx の「」内／素の語句を "…" にして下線・青文字にする明示リスト（キー: "P-T"）。</summary>
        public static readonly IReadOnlyDictionary<string, string[]> Phrases =
            new Dictionary<string, string[]>(StringComparer.Ordinal)
            {
                ["1-5"] = new[] { "干潟の危険は詳しく" },
                ["4-3"] = new[] { "教育" },
                ["4-6"] = new[] { "画像付きスライド" },
                ["5-4"] = new[] { "minato-ongaku.example" },
                ["6-2"] = new[] { "33.12", "25.04" },
            };

        private static readonly Regex InputQuotePattern = new Regex(
            @"(?:「([^」]*)」(?:、[^「]*?)?)+と入力",
            RegexOptions.Compiled);

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
