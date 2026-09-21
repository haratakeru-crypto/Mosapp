using System;
using System.Collections.Generic;
using System.Linq;

namespace MosPracticeClient
{
    public enum ScoringRangeKind
    {
        Projects1To3,
        Projects4To6,
        Projects7To9,
        Projects1To9,
        All,
        Custom
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
                case ScoringRangeKind.Projects7To9:
                    candidates = Enumerable.Range(7, 3);
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

        public static string FormatRangeLabel(ScoringRangeKind kind, IReadOnlyList<int> selectedIds)
        {
            switch (kind)
            {
                case ScoringRangeKind.Projects1To3:
                    return "プロジェクト1～3";
                case ScoringRangeKind.Projects4To6:
                    return "プロジェクト4～6";
                case ScoringRangeKind.Projects7To9:
                    return "プロジェクト7～9";
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
