using System;
using System.Windows;
using System.Windows.Automation;

namespace MOSExcelMogiApp.Vocabulary
{
    /// <summary>
    /// Excel リボンのコンテキストタブを UI Automation で検知・座標取得する（表示・進行専用）。
    /// </summary>
    public static class VocabularyRibbonTabProbe
    {
        static readonly string[] TableDesignNames =
        {
            "テーブルデザイン",
            "テーブル デザイン",
            "Table Design",
            "TableDesign",
        };

        static readonly string[] ChartDesignNames =
        {
            "グラフのデザイン",
            "グラフ デザイン",
            "Chart Design",
            "ChartDesign",
        };

        static readonly string[] HomeTabNames =
        {
            "ホーム",
            "Home",
        };

        public static bool IsTableDesignTabSelected(IntPtr excelHwnd)
        {
            return IsNamedTabSelected(excelHwnd, TableDesignNames);
        }

        public static bool IsChartDesignTabSelected(IntPtr excelHwnd)
        {
            return IsNamedTabSelected(excelHwnd, ChartDesignNames);
        }

        /// <summary>テーブルデザインタブの画面物理ピクセル矩形。見つからなければ null。</summary>
        public static Rect? TryGetTableDesignTabScreenRect(IntPtr excelHwnd)
        {
            return TryGetNamedTabScreenRect(excelHwnd, TableDesignNames);
        }

        public static Rect? TryGetChartDesignTabScreenRect(IntPtr excelHwnd)
        {
            return TryGetNamedTabScreenRect(excelHwnd, ChartDesignNames);
        }

        /// <summary>ホームタブへ切り替え（2/2 開始時にデザインタブ既選択を解除するため）。</summary>
        public static bool TryActivateHomeTab(IntPtr excelHwnd)
        {
            return TrySelectNamedTab(excelHwnd, HomeTabNames);
        }

        static bool IsNamedTabSelected(IntPtr excelHwnd, string[] names)
        {
            var tab = FindNamedTab(excelHwnd, names);
            if (tab == null) return false;
            return IsSelected(tab);
        }

        static Rect? TryGetNamedTabScreenRect(IntPtr excelHwnd, string[] names)
        {
            var tab = FindNamedTab(excelHwnd, names);
            if (tab == null) return null;
            try
            {
                var br = tab.Current.BoundingRectangle;
                if (br.IsEmpty || double.IsNaN(br.Width) || br.Width < 4 || br.Height < 4)
                    return null;
                return new Rect(br.X, br.Y, br.Width, br.Height);
            }
            catch
            {
                return null;
            }
        }

        static bool TrySelectNamedTab(IntPtr excelHwnd, string[] names)
        {
            var tab = FindNamedTab(excelHwnd, names);
            if (tab == null) return false;
            try
            {
                object patternObj;
                if (tab.TryGetCurrentPattern(SelectionItemPattern.Pattern, out patternObj)
                    && patternObj is SelectionItemPattern sip)
                {
                    sip.Select();
                    return true;
                }
            }
            catch { }

            try
            {
                object invokeObj;
                if (tab.TryGetCurrentPattern(InvokePattern.Pattern, out invokeObj)
                    && invokeObj is InvokePattern ip)
                {
                    ip.Invoke();
                    return true;
                }
            }
            catch { }

            return false;
        }

        static AutomationElement FindNamedTab(IntPtr excelHwnd, string[] names)
        {
            if (excelHwnd == IntPtr.Zero || names == null || names.Length == 0)
                return null;

            try
            {
                var root = AutomationElement.FromHandle(excelHwnd);
                if (root == null) return null;

                var tabCondition = new PropertyCondition(
                    AutomationElement.ControlTypeProperty, ControlType.TabItem);

                var tabs = root.FindAll(TreeScope.Descendants, tabCondition);
                if (tabs == null || tabs.Count == 0) return null;

                AutomationElement best = null;
                double bestScore = -1;

                foreach (AutomationElement tab in tabs)
                {
                    string name = "";
                    try { name = tab.Current.Name ?? ""; } catch { continue; }
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (!NameMatches(name, names)) continue;

                    // 画面上に見えているタブを優先（オフスクリーン除外）
                    double score = 0;
                    try
                    {
                        if (tab.Current.IsOffscreen) continue;
                        var br = tab.Current.BoundingRectangle;
                        if (br.IsEmpty || br.Width < 4) continue;
                        score = br.Width * br.Height;
                        // リボン帯（ウィンドウ上部）を優先
                        if (br.Y < 200) score += 10000;
                    }
                    catch { continue; }

                    if (score > bestScore)
                    {
                        bestScore = score;
                        best = tab;
                    }
                }

                return best;
            }
            catch
            {
                return null;
            }
        }

        static bool NameMatches(string actual, string[] candidates)
        {
            string a = actual.Trim();
            foreach (var c in candidates)
            {
                if (string.IsNullOrEmpty(c)) continue;
                if (string.Equals(a, c, StringComparison.OrdinalIgnoreCase))
                    return true;
                if (a.IndexOf(c, StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
            }
            return false;
        }

        static bool IsSelected(AutomationElement tab)
        {
            try
            {
                object patternObj;
                if (tab.TryGetCurrentPattern(SelectionItemPattern.Pattern, out patternObj)
                    && patternObj is SelectionItemPattern sip)
                {
                    return sip.Current.IsSelected;
                }
            }
            catch { }

            try
            {
                if (tab.Current.HasKeyboardFocus) return true;
            }
            catch { }

            return false;
        }
    }
}
