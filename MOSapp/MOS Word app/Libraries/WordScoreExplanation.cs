using System;
using System.Collections.Generic;
using System.Text;

namespace Libraries
{
    /// <summary>
    /// 採点の正誤は変えず、×の理由だけを学生向けの文にする。
    /// チェッカーは Note、呼び出し側は Clear → 採点 → ResolveFailReason。
    /// チェッカー DLL とホストで別アセンブリになるため、理由は AppDomain 経由で共有する。
    /// </summary>
    public static class WordScoreExplanation
    {
        public const string RequirementMissText = "求められている状態になっていないため、×になりました。";
        public const string UnavailableText = "この課題は確認できませんでした。";

        /// <summary>破壊的操作の動作名（未知）。一覧組み立て用。</summary>
        public const string UnknownStudentAction = "操作をした";

        /// <summary>破壊的操作の理由が取れなかったときの全文。</summary>
        public const string DestructiveStudentText =
            "この課題の作業中に、この課題では使わない操作をした記録があるため、×になりました。";

        private const string CheckerReasonKey = "WordScoreExplanation.CheckerReason";

        /// <summary>タスク採点の直前に呼び、前の課題の理由を消す。</summary>
        public static void ClearCheckerReason()
        {
            AppDomain.CurrentDomain.SetData(CheckerReasonKey, null);
        }

        /// <summary>チェッカーが false を返す直前に、学生向けの理由を1件足す。</summary>
        public static void Note(string reason)
        {
            if (string.IsNullOrWhiteSpace(reason))
                return;
            string line = reason.Trim();
            string existing = AppDomain.CurrentDomain.GetData(CheckerReasonKey) as string;
            AppDomain.CurrentDomain.SetData(
                CheckerReasonKey,
                string.IsNullOrEmpty(existing) ? line : existing + Environment.NewLine + line);
        }

        /// <summary>記録された理由を取り出し、次の課題に持ち越さない。</summary>
        public static string TakeCheckerReason()
        {
            string reason = AppDomain.CurrentDomain.GetData(CheckerReasonKey) as string;
            AppDomain.CurrentDomain.SetData(CheckerReasonKey, null);
            return reason;
        }

        /// <summary>
        /// ×のときの学生向け理由を組み立てる。合格時は空文字。
        /// ゲート失敗とチェッカー失敗が両方あるときは改行でつなぐ。
        /// </summary>
        public static string ResolveFailReason(
            bool passed,
            bool gateFailed = false,
            IList<string> gateStudentActions = null,
            bool unavailable = false)
        {
            string checkerReason = TakeCheckerReason();
            if (passed)
                return "";

            if (unavailable)
                return UnavailableText;

            string checkerText = string.IsNullOrWhiteSpace(checkerReason)
                ? null
                : checkerReason.Trim();
            string gateText = gateFailed
                ? FormatDestructiveStudentReasons(gateStudentActions)
                : null;

            if (checkerText != null && gateText != null)
                return checkerText + Environment.NewLine + gateText;
            if (gateText != null)
                return gateText;
            if (checkerText != null)
                return checkerText;
            return RequirementMissText;
        }

        /// <summary>
        /// 破壊的操作の動作名一覧を学生向け文にする。前置きは先頭1文のみ。
        /// </summary>
        public static string FormatDestructiveStudentReasons(IList<string> actions)
        {
            if (actions == null || actions.Count == 0)
                return DestructiveStudentText;

            var cleaned = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string action in actions)
            {
                if (string.IsNullOrWhiteSpace(action))
                    continue;
                string line = action.Trim();
                if (seen.Add(line))
                    cleaned.Add(line);
            }

            if (cleaned.Count == 0)
                return DestructiveStudentText;

            if (cleaned.Count == 1)
                return FormatRecordedAction(cleaned[0]);

            var sb = new StringBuilder();
            sb.Append(DestructiveStudentText);
            foreach (string action in cleaned)
            {
                sb.Append(Environment.NewLine);
                sb.Append("・");
                sb.Append(action);
            }
            return sb.ToString();
        }

        public static string FormatRecordedAction(string action)
        {
            return "この課題の作業中に、この課題では使わない" + action + "記録があるため、×になりました。";
        }

        /// <summary>[Op] 種別を学生向けの短い動作名（「〜した」）にする。</summary>
        public static string DescribeOperationType(string operationType)
        {
            if (string.IsNullOrWhiteSpace(operationType))
                return UnknownStudentAction;

            switch (operationType.Trim())
            {
                case "Cut":
                case "CutSelection":
                    return "切り取りをした";
                case "Paste":
                    return "貼り付けをした";
                case "InsertSectionBreak":
                    return "セクション区切りを入れた";
                case "ReviewNewComment":
                    return "コメントを追加した";
                case "InsertTable":
                    return "表を追加した";
                case "InsertPicture":
                    return "画像を挿入した";
                case "Watermark":
                    return "透かしを変えた";
                case "PageBorders":
                    return "ページ罫線を変えた";
                case "HeaderFooterEdit":
                    return "ヘッダーまたはフッターを編集した";
                case "FileSaveAs":
                    return "名前を付けて保存をした";
                default:
                    return UnknownStudentAction;
            }
        }

        /// <summary>スナップショット差分トークン（または特殊ゲート理由）を学生向け動作名にする。</summary>
        public static string DescribeSnapshotOrGateToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
                return UnknownStudentAction;

            string t = token.Trim();
            if (t.StartsWith("SectionsCount", StringComparison.OrdinalIgnoreCase))
                return "セクション区切りを変えた";
            if (t.StartsWith("BodyTextLength", StringComparison.OrdinalIgnoreCase)
                || t.StartsWith("VisibleBodyTextLength", StringComparison.OrdinalIgnoreCase))
                return "文書の本文を変えた";
            if (t.StartsWith("InlineShapesCount", StringComparison.OrdinalIgnoreCase))
                return "画像を追加または削除した";
            if (t.StartsWith("FloatingShapesCount", StringComparison.OrdinalIgnoreCase))
                return "図形を追加または削除した";
            if (t.StartsWith("TablesCount", StringComparison.OrdinalIgnoreCase))
                return "表を追加または削除した";
            if (t.StartsWith("CommentsCount", StringComparison.OrdinalIgnoreCase))
                return "コメントを追加または削除した";
            if (t.StartsWith("HeaderFooterFingerprint", StringComparison.OrdinalIgnoreCase))
                return "ヘッダーまたはフッターを変えた";
            if (t.StartsWith("CompatibilityMode", StringComparison.OrdinalIgnoreCase))
                return "文書の互換モードを変えた";
            if (t.StartsWith("WatermarkFingerprint", StringComparison.OrdinalIgnoreCase))
                return "透かしを変えた";
            if (t.StartsWith("PageBorderFingerprint", StringComparison.OrdinalIgnoreCase))
                return "ページ罫線を変えた";
            if (t.StartsWith("FootnoteReferenceCount", StringComparison.OrdinalIgnoreCase))
                return "脚注を変えた";
            if (t.IndexOf("CutParagraphSelection", StringComparison.OrdinalIgnoreCase) >= 0)
                return "段落ごと切り取った";
            if (t.IndexOf("ReviewDeleteComment", StringComparison.OrdinalIgnoreCase) >= 0)
                return "コメントを解決せずに削除した";
            if (t.IndexOf("ReviewResolveComment", StringComparison.OrdinalIgnoreCase) >= 0)
                return "コメントを削除せずに解決した";
            if (t.IndexOf("disallowed operation", StringComparison.OrdinalIgnoreCase) >= 0)
                return UnknownStudentAction;
            if (t.IndexOf("logged destructive", StringComparison.OrdinalIgnoreCase) >= 0)
                return UnknownStudentAction;
            return UnknownStudentAction;
        }
    }
}
