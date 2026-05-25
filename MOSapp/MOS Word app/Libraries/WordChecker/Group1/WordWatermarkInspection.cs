using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace Libraries.Group1
{
    /// <summary>
    /// 4-5 透かし判定。ギャラリー「下書き1」は表示文字「下書き」＋斜め、「下書き2」は「下書き」＋横書き。
    /// WordChecker と VSTO ポーリングで同条件を共有する。
    /// </summary>
    public static class WordWatermarkInspection
    {
        private const int ContextCharsBefore = 2000;
        private const int ContextCharsAfter = 1500;

        private static readonly string[] ForbiddenWatermarkKeywords =
        {
            "社外秘", "至急", "案内", "転送厳禁", "CONFIDENTIAL", "URGENT", "DO NOT COPY"
        };

        private static readonly string[] DiagonalRotationLiteralMarkers =
        {
            "rotation:315", "rotation:-315", "rotation:-45", "rotation:45",
            "rot=\"315\"", "rot=\"-315\"", "rot=\"-45\"", "rot=\"45\""
        };

        private static readonly Regex RotationInStyleRegex = new Regex(
            @"rotation:\s*(-?\d+(?:\.\d+)?)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public static string NormalizeXml(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return string.Empty;
            try
            {
                return xml.Normalize(NormalizationForm.FormKC);
            }
            catch
            {
                return xml;
            }
        }

        /// <summary>［デザイン］透かし「下書き1」: 文字「下書き」かつ斜め。社外秘・至急・下書き2（横）は false。</summary>
        public static bool IsDraft1Watermark(string normalizedXml)
        {
            if (string.IsNullOrEmpty(normalizedXml))
                return false;
            if (HasForbiddenWatermark(normalizedXml))
                return false;
            if (normalizedXml.IndexOf("下書き", StringComparison.Ordinal) < 0)
                return false;
            if (IsDraft2HorizontalWatermark(normalizedXml))
                return false;
            return ContainsDraft1DiagonalNearDraftText(normalizedXml);
        }

        /// <summary>透かし「下書き2」: 文字「下書き」かつ横書き（斜めでない）。</summary>
        public static bool IsDraft2HorizontalWatermark(string normalizedXml)
        {
            if (string.IsNullOrEmpty(normalizedXml))
                return false;
            if (normalizedXml.IndexOf("下書き", StringComparison.Ordinal) < 0)
                return false;

            foreach (int index in FindAllIndices(normalizedXml, "下書き"))
            {
                string context = GetContext(normalizedXml, index, ContextCharsBefore, ContextCharsAfter);
                if (ContextHasDiagonalRotation(context))
                    continue;
                if (ContextLooksHorizontalWatermark(context))
                    return true;
            }

            return false;
        }

        public static bool HasForbiddenWatermark(string normalizedXml)
        {
            if (string.IsNullOrEmpty(normalizedXml))
                return false;
            foreach (string keyword in ForbiddenWatermarkKeywords)
            {
                if (normalizedXml.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }

        /// <summary>4-7 後: 禁止透かし・下書き1/2 相当の透かしが残っていない。</summary>
        public static bool IsWatermarkClearedForTask47FollowUp(string normalizedXml)
        {
            if (string.IsNullOrEmpty(normalizedXml))
                return true;
            if (HasForbiddenWatermark(normalizedXml))
                return false;
            if (IsDraft1Watermark(normalizedXml))
                return false;
            if (IsDraft2HorizontalWatermark(normalizedXml))
                return false;
            return normalizedXml.IndexOf("下書き", StringComparison.Ordinal) < 0;
        }

        private static bool ContainsDraft1DiagonalNearDraftText(string normalizedXml)
        {
            foreach (int index in FindAllIndices(normalizedXml, "下書き"))
            {
                string context = GetContext(normalizedXml, index, ContextCharsBefore, ContextCharsAfter);
                if (ContextHasDiagonalRotation(context))
                    return true;
            }
            return false;
        }

        private static bool ContextHasDiagonalRotation(string context)
        {
            if (string.IsNullOrEmpty(context))
                return false;

            foreach (string marker in DiagonalRotationLiteralMarkers)
            {
                if (context.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }

            foreach (Match match in RotationInStyleRegex.Matches(context))
            {
                if (!TryParseRotationDegrees(match.Groups[1].Value, out double degrees))
                    continue;
                if (IsDiagonalRotationDegrees(degrees))
                    return true;
            }

            return false;
        }

        private static bool ContextLooksHorizontalWatermark(string context)
        {
            if (string.IsNullOrEmpty(context))
                return false;

            if (context.IndexOf("rotation:0", StringComparison.OrdinalIgnoreCase) >= 0
                || context.IndexOf("rotation:0deg", StringComparison.OrdinalIgnoreCase) >= 0)
                return true;

            foreach (Match match in RotationInStyleRegex.Matches(context))
            {
                if (TryParseRotationDegrees(match.Groups[1].Value, out double degrees)
                    && Math.Abs(degrees) < 2)
                    return true;
            }

            // 横書きプリセットは textDirection が lrTb で rotation 指定が無いことが多い
            if (context.IndexOf("lrTb", StringComparison.OrdinalIgnoreCase) >= 0
                && context.IndexOf("下書き", StringComparison.Ordinal) >= 0)
                return true;

            return false;
        }

        private static bool TryParseRotationDegrees(string value, out double degrees)
        {
            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out degrees);
        }

        /// <summary>Word 透かしの斜め（下書き1）: 0° 付近以外の典型角。</summary>
        private static bool IsDiagonalRotationDegrees(double degrees)
        {
            degrees = NormalizeDegrees(degrees);
            if (degrees < 2)
                return false;
            // 45°, 135°, 225°, 315° 付近（±8°）
            double[] diagonalAngles = { 45, 135, 225, 315 };
            foreach (double target in diagonalAngles)
            {
                double diff = Math.Abs(degrees - target);
                if (diff <= 8 || Math.Abs(diff - 360) <= 8)
                    return true;
            }
            return false;
        }

        private static double NormalizeDegrees(double degrees)
        {
            degrees %= 360;
            if (degrees < 0)
                degrees += 360;
            return degrees;
        }

        private static string GetContext(string text, int anchorIndex, int charsBefore, int charsAfter)
        {
            int start = Math.Max(0, anchorIndex - charsBefore);
            int end = Math.Min(text.Length, anchorIndex + charsAfter);
            return text.Substring(start, end - start);
        }

        private static IEnumerable<int> FindAllIndices(string text, string value)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(value))
                yield break;

            int index = 0;
            while (index < text.Length)
            {
                index = text.IndexOf(value, index, StringComparison.Ordinal);
                if (index < 0)
                    yield break;
                yield return index;
                index += value.Length;
            }
        }
    }
}
