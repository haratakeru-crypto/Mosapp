using System;

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
            string gateInternalReason = null,
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
                ? FormatGateStudentReason(gateInternalReason)
                : null;

            if (checkerText != null && gateText != null)
                return checkerText + Environment.NewLine + gateText;
            if (gateText != null)
                return gateText;
            if (checkerText != null)
                return checkerText;
            return RequirementMissText;
        }

        /// <summary>ゲート内部理由を学生向け文にする（詳細マッピングは今後拡充）。</summary>
        public static string FormatGateStudentReason(string gateInternalReason)
        {
            if (string.IsNullOrWhiteSpace(gateInternalReason))
                return DestructiveStudentText;

            // 骨格段階では種別を分けず共通文。内部理由はログ用に残す想定。
            return DestructiveStudentText;
        }
    }
}
