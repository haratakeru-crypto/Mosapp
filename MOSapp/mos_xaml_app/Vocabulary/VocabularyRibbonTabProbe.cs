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

        /// <summary>リボン上のボタンまたはグループ名の画面物理ピクセル矩形。見えているものだけ。</summary>
        public static Rect? TryGetRibbonControlScreenRect(IntPtr excelHwnd, string controlName)
        {
            if (excelHwnd == IntPtr.Zero || string.IsNullOrWhiteSpace(controlName))
                return null;

            try
            {
                var root = AutomationElement.FromHandle(excelHwnd);
                if (root == null) return null;

                var condition = new OrCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.SplitButton),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Group));
                var found = root.FindAll(TreeScope.Descendants, condition);
                if (found == null || found.Count == 0) return null;

                string wanted = controlName.Trim();
                AutomationElement best = null;
                int bestRank = int.MinValue;
                foreach (AutomationElement el in found)
                {
                    string name = "";
                    try { name = el.Current.Name ?? ""; } catch { continue; }
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (name.IndexOf(wanted, StringComparison.OrdinalIgnoreCase) < 0) continue;

                    System.Windows.Rect br;
                    try
                    {
                        if (el.Current.IsOffscreen) continue;
                        br = el.Current.BoundingRectangle;
                    }
                    catch { continue; }
                    if (br.IsEmpty || br.Width < 8 || br.Height < 8 || br.Width > 520 || br.Height > 160)
                        continue;

                    int rank = (int)Math.Min(8000, br.Width * br.Height);
                    if (string.Equals(name.Trim(), wanted, StringComparison.OrdinalIgnoreCase))
                        rank += 20000;
                    try
                    {
                        var kind = el.Current.ControlType;
                        if (kind == ControlType.Button || kind == ControlType.SplitButton)
                            rank += 10000;
                    }
                    catch { }
                    if (rank > bestRank)
                    {
                        bestRank = rank;
                        best = el;
                    }
                }

                if (best == null) return null;
                var box = best.Current.BoundingRectangle;
                return new Rect(box.X, box.Y, box.Width, box.Height);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>リボン TabItem 名（ページレイアウト、挿入 など）の画面物理ピクセル矩形。</summary>
        public static Rect? TryGetTabScreenRect(IntPtr excelHwnd, string tabName)
        {
            if (string.IsNullOrWhiteSpace(tabName)) return null;
            string name = tabName.Trim();
            string withoutTab = name.EndsWith("タブ", StringComparison.Ordinal)
                ? name.Substring(0, name.Length - 2).Trim()
                : name;
            return TryGetNamedTabScreenRect(excelHwnd, new[] { name, withoutTab });
        }

        /// <summary>
        /// タブの位置。TabItem で見つからなければ、リボン上部のボタン等から同じ名前を探す。
        /// </summary>
        public static Rect? TryGetTabScreenRectLoose(IntPtr excelHwnd, string tabName)
        {
            var rect = TryGetTabScreenRect(excelHwnd, tabName);
            if (rect.HasValue) return rect;
            if (excelHwnd == IntPtr.Zero || string.IsNullOrWhiteSpace(tabName)) return null;

            try
            {
                var root = AutomationElement.FromHandle(excelHwnd);
                if (root == null) return null;
                System.Windows.Rect rootBox;
                try { rootBox = root.Current.BoundingRectangle; }
                catch { return null; }

                var condition = new OrCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
                var found = root.FindAll(TreeScope.Descendants, condition);
                if (found == null) return null;

                string wanted = NormalizeTabName(tabName);
                foreach (AutomationElement el in found)
                {
                    try
                    {
                        if (el.Current.IsOffscreen) continue;
                        if (!string.Equals(NormalizeTabName(el.Current.Name), wanted, StringComparison.OrdinalIgnoreCase))
                            continue;
                        var br = el.Current.BoundingRectangle;
                        if (br.IsEmpty || br.Width < 8 || br.Height < 8) continue;
                        if (!rootBox.IsEmpty && br.Y > rootBox.Y + 200) continue;
                        return new Rect(br.X, br.Y, br.Width, br.Height);
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        delegate bool EnumChildProc(IntPtr hwnd, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool EnumChildWindows(IntPtr parent, EnumChildProc callback, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr hWnd);

        static readonly string[] BackNames = { "戻る", "Back" };
        static readonly string[] BackstagePageNames =
        {
            "ホーム", "新規", "開く", "情報", "上書き保存", "名前を付けて保存", "コピーを保存",
            "印刷", "共有", "エクスポート", "発行", "閉じる", "アカウント", "フィードバック", "オプション",
            "Home", "New", "Open", "Info", "Save", "Save As", "Save a Copy", "Print", "Share",
            "Export", "Publish", "Close", "Account", "Feedback", "Options",
        };

        /// <summary>ファイルタブ（バックステージ）を表示するウィンドウ。無ければ Zero。</summary>
        static IntPtr FindBackstageHost(IntPtr excelHwnd)
        {
            if (excelHwnd == IntPtr.Zero) return IntPtr.Zero;
            IntPtr found = IntPtr.Zero;
            var sb = new System.Text.StringBuilder(64);
            try
            {
                EnumChildWindows(excelHwnd, (h, _) =>
                {
                    sb.Clear();
                    GetClassName(h, sb, sb.Capacity);
                    if (string.Equals(sb.ToString(), "FullpageUIHost", StringComparison.OrdinalIgnoreCase)
                        && IsWindowVisible(h))
                    {
                        found = h;
                        return false;
                    }
                    return true;
                }, IntPtr.Zero);
            }
            catch { }
            return found;
        }

        /// <summary>ファイルタブ（バックステージ）が開いているか。</summary>
        public static bool IsBackstageOpen(IntPtr excelHwnd)
        {
            if (FindBackstageHost(excelHwnd) != IntPtr.Zero) return true;
            return FindBackstageElement(excelHwnd, BackNames, exact: true) != null;
        }

        /// <summary>バックステージの左側で選択中の項目名（情報 など）。</summary>
        public static string TryGetBackstageSelectedPage(IntPtr excelHwnd)
        {
            try
            {
                var root = BackstageRoot(excelHwnd);
                if (root == null) return null;
                var items = root.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.IsSelectionItemPatternAvailableProperty, true));
                if (items == null) return null;
                foreach (AutomationElement el in items)
                {
                    string name;
                    try
                    {
                        if (el.Current.IsOffscreen) continue;
                        name = (el.Current.Name ?? "").Trim();
                    }
                    catch { continue; }
                    if (!NameMatchesExact(name, BackstagePageNames)) continue;
                    if (IsSelectedStrict(el)) return name;
                }
            }
            catch { }
            return null;
        }

        /// <summary>バックステージの項目（情報、戻る など）の画面物理ピクセル矩形。</summary>
        public static Rect? TryGetBackstageItemScreenRect(IntPtr excelHwnd, string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            string[] names = NameMatchesExact(name, BackNames) ? BackNames : new[] { name.Trim() };
            var el = FindBackstageElement(excelHwnd, names, exact: true);
            if (el == null) return null;
            try
            {
                var br = el.Current.BoundingRectangle;
                if (br.IsEmpty || br.Width < 8 || br.Height < 8) return null;
                return new Rect(br.X, br.Y, br.Width, br.Height);
            }
            catch { return null; }
        }

        static AutomationElement BackstageRoot(IntPtr excelHwnd)
        {
            try
            {
                IntPtr host = FindBackstageHost(excelHwnd);
                return AutomationElement.FromHandle(host != IntPtr.Zero ? host : excelHwnd);
            }
            catch { return null; }
        }

        /// <summary>バックステージの左側（ナビゲーション）にある名前の要素。</summary>
        static AutomationElement FindBackstageElement(IntPtr excelHwnd, string[] names, bool exact)
        {
            if (excelHwnd == IntPtr.Zero) return null;
            try
            {
                var root = BackstageRoot(excelHwnd);
                if (root == null) return null;
                System.Windows.Rect rootBox;
                try { rootBox = root.Current.BoundingRectangle; }
                catch { rootBox = System.Windows.Rect.Empty; }

                var condition = new OrCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Hyperlink));
                var found = root.FindAll(TreeScope.Descendants, condition);
                if (found == null) return null;

                foreach (AutomationElement el in found)
                {
                    try
                    {
                        if (el.Current.IsOffscreen) continue;
                        string name = (el.Current.Name ?? "").Trim();
                        if (exact ? !NameMatchesExact(name, names) : !NameMatches(name, names)) continue;
                        var br = el.Current.BoundingRectangle;
                        if (br.IsEmpty || br.Width < 8 || br.Height < 8) continue;
                        // 左側のナビゲーションだけを見る（本文の同名リンクを避ける）。
                        if (!rootBox.IsEmpty && br.X > rootBox.X + 420) continue;
                        return el;
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        static bool NameMatchesExact(string actual, string[] candidates)
        {
            string a = (actual ?? "").Trim();
            foreach (var c in candidates)
            {
                if (string.Equals(a, c, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        /// <summary>リボンで選択中のタブ名。シート見出しは除く。取れなければ null。</summary>
        public static string TryGetSelectedTabName(IntPtr excelHwnd)
        {
            if (excelHwnd == IntPtr.Zero) return null;
            try
            {
                var root = AutomationElement.FromHandle(excelHwnd);
                if (root == null) return null;
                System.Windows.Rect rootBox;
                try { rootBox = root.Current.BoundingRectangle; }
                catch { return null; }

                var tabs = root.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
                if (tabs == null) return null;

                foreach (AutomationElement tab in tabs)
                {
                    string name;
                    System.Windows.Rect br;
                    try
                    {
                        if (tab.Current.IsOffscreen) continue;
                        name = tab.Current.Name ?? "";
                        br = tab.Current.BoundingRectangle;
                    }
                    catch { continue; }
                    if (string.IsNullOrWhiteSpace(name) || br.IsEmpty || br.Width < 4) continue;
                    // シート見出し（下端）を除き、リボン帯のタブだけを見る。
                    if (!rootBox.IsEmpty && br.Y > rootBox.Y + 260) continue;
                    if (IsSelectedStrict(tab))
                        return name.Trim();
                }
            }
            catch { }
            return null;
        }

        /// <summary>タブ名が候補と同じか。空白と末尾の「タブ」は無視する。</summary>
        public static bool TabNameEquals(string actual, string expected)
        {
            string a = NormalizeTabName(actual);
            string e = NormalizeTabName(expected);
            if (a.Length == 0 || e.Length == 0) return false;
            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase)
                   || a.IndexOf(e, StringComparison.OrdinalIgnoreCase) >= 0
                   || e.IndexOf(a, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool IsHomeTabName(string name)
        {
            foreach (var h in HomeTabNames)
            {
                if (TabNameEquals(name, h)) return true;
            }
            return false;
        }

        static string NormalizeTabName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "";
            string n = name.Replace(" ", "").Replace("　", "").Trim();
            if (n.EndsWith("タブ", StringComparison.Ordinal))
                n = n.Substring(0, n.Length - 2);
            return n;
        }

        static bool IsSelectedStrict(AutomationElement tab)
        {
            try
            {
                object patternObj;
                if (tab.TryGetCurrentPattern(SelectionItemPattern.Pattern, out patternObj)
                    && patternObj is SelectionItemPattern sip)
                    return sip.Current.IsSelected;
            }
            catch { }
            return false;
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
