using System;
using System.Collections.Generic;
using System.Linq;

namespace MosPracticeClient
{
    public enum ScoringRangeKind
    {
        Projects1To3,
        Projects4To6,
        Projects7To8,
        Projects7To9,
        Projects1To7,
        Projects1To8,
        Projects1To9,
        All,
        Custom
    }

    public enum ScoringRangeProfile
    {
        Standard,
        WordVariant
    }

    public static class ScoringRangeSelection
    {
        public static IReadOnlyList<int> Resolve(
            ScoringRangeKind kind,
            IReadOnlyCollection<int> availableProjectIds,
            IEnumerable<int> customProjectIds = null)
        {
            var available = new HashSet<int>(availableProjectIds ?? Array.Empty<int>());
            IEnumerable<int> candidates;
            switch (kind)
            {
                case ScoringRangeKind.Projects1To3:
                    candidates = Enumerable.Range(1, 3);
                    break;
                case ScoringRangeKind.Projects4To6:
                    candidates = Enumerable.Range(4, 3);
                    break;
                case ScoringRangeKind.Projects7To8:
                    candidates = Enumerable.Range(7, 2);
                    break;
                case ScoringRangeKind.Projects7To9:
                    candidates = Enumerable.Range(7, 3);
                    break;
                case ScoringRangeKind.Projects1To7:
                    candidates = Enumerable.Range(1, 7);
                    break;
                case ScoringRangeKind.Projects1To8:
                    candidates = Enumerable.Range(1, 8);
                    break;
                case ScoringRangeKind.Projects1To9:
                    candidates = Enumerable.Range(1, 9);
                    break;
                case ScoringRangeKind.Custom:
                    candidates = customProjectIds ?? Array.Empty<int>();
                    break;
                default:
                    candidates = available;
                    break;
            }

            return candidates
                .Where(available.Contains)
                .Distinct()
                .OrderBy(id => id)
                .ToList();
        }

        public static bool ShouldOffer(
            ScoringRangeKind kind,
            IReadOnlyCollection<int> availableProjectIds,
            ScoringRangeProfile profile = ScoringRangeProfile.Standard)
        {
            if (!IsKindInProfile(kind, profile))
                return false;

            if (kind == ScoringRangeKind.All || kind == ScoringRangeKind.Custom)
                return true;

            var resolved = Resolve(kind, availableProjectIds);
            if (resolved.Count == 0)
                return false;

            if (kind == ScoringRangeKind.Projects1To7
                || kind == ScoringRangeKind.Projects1To8
                || kind == ScoringRangeKind.Projects1To9)
            {
                var all = Resolve(ScoringRangeKind.All, availableProjectIds);
                if (resolved.SequenceEqual(all))
                    return false;
            }

            return true;
        }

        public static string ShortenRangeLabel(string rangeLabel)
        {
            string label = (rangeLabel ?? "").Trim();
            if (label.Length == 0) return "";
            if (label == "すべてのプロジェクト") return "すべて";
            const string prefix = "プロジェクト";
            if (label.StartsWith(prefix, StringComparison.Ordinal))
                return "P" + label.Substring(prefix.Length);
            return label;
        }

        public static bool IsKindInProfile(ScoringRangeKind kind, ScoringRangeProfile profile)
        {
            if (profile == ScoringRangeProfile.WordVariant)
            {
                return kind == ScoringRangeKind.Projects1To3
                    || kind == ScoringRangeKind.Projects4To6
                    || kind == ScoringRangeKind.Projects7To8
                    || kind == ScoringRangeKind.All
                    || kind == ScoringRangeKind.Custom;
            }

            return kind == ScoringRangeKind.Projects1To7
                || kind == ScoringRangeKind.Projects1To8
                || kind == ScoringRangeKind.Projects1To9
                || kind == ScoringRangeKind.All
                || kind == ScoringRangeKind.Custom;
        }

        public static string FormatRangeLabel(ScoringRangeKind kind, IReadOnlyList<int> selectedIds)
        {
            switch (kind)
            {
                case ScoringRangeKind.Projects1To3:
                    return "プロジェクト1～3";
                case ScoringRangeKind.Projects4To6:
                    return "プロジェクト4～6";
                case ScoringRangeKind.Projects7To8:
                    return "プロジェクト7・8";
                case ScoringRangeKind.Projects7To9:
                    return "プロジェクト7～9";
                case ScoringRangeKind.Projects1To7:
                    return "プロジェクト1～7";
                case ScoringRangeKind.Projects1To8:
                    return "プロジェクト1～8";
                case ScoringRangeKind.Projects1To9:
                    return "プロジェクト1～9";
                case ScoringRangeKind.All:
                    return "すべてのプロジェクト";
                default:
                    if (selectedIds == null || selectedIds.Count == 0)
                        return "カスタム";
                    return "プロジェクト" + string.Join(",", selectedIds);
            }
        }
    }
}
