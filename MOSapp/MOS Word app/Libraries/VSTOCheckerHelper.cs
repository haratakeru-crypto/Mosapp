using System;
using System.Collections.Generic;

namespace Libraries
{
    /// <summary>
    /// VSTOで採点すべき問題を判定するヘルパークラス
    /// </summary>
    public static class VSTOCheckerHelper
    {
        /// <summary>
        /// VSTOで採点すべき問題IDのセット
        /// 形式: "プロジェクト-タスク" (例: "1-1", "2-1")
        /// </summary>
        private static readonly HashSet<string> VSTORequiredTaskIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "1-1",  // 編集記号の表示/非表示 (ShowAll)
            "2-1",  // 文字列の切り取り・貼り付け (Cut, Paste)
            "4-2",  // コメントへの返信 (ReviewCommentReply)
            "4-3",  // コメントの解決 (ReviewResolveComment)
            "6-1",  // 文字列を表に変換 (TableConvertTextToTable)
            "7-1",  // 互換モードの解除 (UpgradeDocument)
            "7-4",  // テキスト保存 (FileSaveAs)
            "7-5",  // マクロ有効保存 (FileSaveAs)
            "8-1"   // 目次の挿入 (TocAutomatic2)
        };

        /// <summary>
        /// 指定した問題IDがVSTOで採点すべき問題かどうかを判定
        /// </summary>
        /// <param name="groupId">グループID（プロジェクト番号）</param>
        /// <param name="projectId">プロジェクトID</param>
        /// <param name="taskId">タスクID</param>
        /// <returns>VSTOで採点すべき問題の場合true</returns>
        public static bool RequiresVSTOCheck(int groupId, int projectId, int taskId)
        {
            string taskKey = $"{projectId}-{taskId}";
            return VSTORequiredTaskIds.Contains(taskKey);
        }

        /// <summary>
        /// 指定した問題IDがVSTOで採点すべき問題かどうかを判定（文字列形式）
        /// </summary>
        /// <param name="taskKey">問題ID（形式: "プロジェクト-タスク"、例: "1-1"）</param>
        /// <returns>VSTOで採点すべき問題の場合true</returns>
        public static bool RequiresVSTOCheck(string taskKey)
        {
            if (string.IsNullOrWhiteSpace(taskKey))
                return false;

            return VSTORequiredTaskIds.Contains(taskKey.Trim());
        }

        /// <summary>
        /// VSTOで採点すべき問題IDのリストを取得
        /// </summary>
        /// <returns>問題IDのリスト</returns>
        public static IEnumerable<string> GetVSTORequiredTaskIds()
        {
            return VSTORequiredTaskIds;
        }
    }
}

