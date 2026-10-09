using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using Newtonsoft.Json;

namespace MOSExcelMogiApp.Vocabulary
{
    /// <summary>
    /// 設定画面で合わせたコーチマークの枠を、画面ごとのキーで保存する。
    /// 物理ピクセルではなく、Excel ウィンドウか画面全体に対する比率で持つ。
    /// </summary>
    public static class CoachHoleOverrideStore
    {
        public const string BasisExcel = "excel";
        public const string BasisScreen = "screen";

        sealed class Entry
        {
            [JsonProperty("basis")] public string Basis { get; set; }
            [JsonProperty("left")] public double Left { get; set; }
            [JsonProperty("top")] public double Top { get; set; }
            [JsonProperty("width")] public double Width { get; set; }
            [JsonProperty("height")] public double Height { get; set; }
        }

        [StructLayout(LayoutKind.Sequential)]
        struct RECT { public int Left, Top, Right, Bottom; }

        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(int nIndex);

        static readonly object Gate = new object();
        static Dictionary<string, Entry> _map;

        static string FilePath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MOSapp",
            "CoachHoleOverrides.json");

        static Rect? Frame(string basis, IntPtr excelHwnd)
        {
            if (string.Equals(basis, BasisScreen, StringComparison.OrdinalIgnoreCase))
            {
                int x = GetSystemMetrics(76), y = GetSystemMetrics(77);
                int w = GetSystemMetrics(78), h = GetSystemMetrics(79);
                if (w < 100 || h < 100) return null;
                return new Rect(x, y, w, h);
            }
            if (excelHwnd == IntPtr.Zero || !GetWindowRect(excelHwnd, out RECT r)) return null;
            if (r.Right - r.Left < 100 || r.Bottom - r.Top < 100) return null;
            return new Rect(r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top);
        }

        static void EnsureLoaded()
        {
            if (_map != null) return;
            try
            {
                _map = File.Exists(FilePath)
                    ? JsonConvert.DeserializeObject<Dictionary<string, Entry>>(File.ReadAllText(FilePath))
                    : null;
            }
            catch { _map = null; }
            if (_map == null) _map = new Dictionary<string, Entry>(StringComparer.Ordinal);
        }

        static void SaveFile()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(FilePath));
                File.WriteAllText(FilePath, JsonConvert.SerializeObject(_map, Formatting.Indented));
            }
            catch { }
        }

        public static bool Has(string key)
        {
            if (string.IsNullOrEmpty(key)) return false;
            lock (Gate)
            {
                EnsureLoaded();
                return _map.ContainsKey(key);
            }
        }

        /// <summary>保存した枠を、いまの画面の物理ピクセルにして返す。</summary>
        public static bool TryGet(string key, IntPtr excelHwnd, out Rect physical)
        {
            physical = Rect.Empty;
            if (string.IsNullOrEmpty(key)) return false;
            Entry e;
            lock (Gate)
            {
                EnsureLoaded();
                if (!_map.TryGetValue(key, out e) || e == null) return false;
            }
            var frame = Frame(e.Basis, excelHwnd);
            if (!frame.HasValue) return false;
            var f = frame.Value;
            physical = new Rect(f.X + e.Left * f.Width, f.Y + e.Top * f.Height, e.Width * f.Width, e.Height * f.Height);
            return physical.Width >= 4 && physical.Height >= 4;
        }

        public static void Set(string key, Rect physical, string basis, IntPtr excelHwnd)
        {
            if (string.IsNullOrEmpty(key) || physical.Width < 4 || physical.Height < 4) return;
            var frame = Frame(basis, excelHwnd);
            if (!frame.HasValue) return;
            var f = frame.Value;
            lock (Gate)
            {
                EnsureLoaded();
                _map[key] = new Entry
                {
                    Basis = basis,
                    Left = (physical.X - f.X) / f.Width,
                    Top = (physical.Y - f.Y) / f.Height,
                    Width = physical.Width / f.Width,
                    Height = physical.Height / f.Height
                };
                SaveFile();
            }
        }

        public static void Remove(string key)
        {
            if (string.IsNullOrEmpty(key)) return;
            lock (Gate)
            {
                EnsureLoaded();
                if (_map.Remove(key)) SaveFile();
            }
        }

        public static string HintKey(string hint) => "hint:" + hint;
        public static string QuizKey(string keyword, string part) => "quiz:" + (keyword ?? "").Trim() + ":" + part;
        public const string KeywordAnchorKey = "tutorial:keyword";
        public const string ReviewAnchorKey = "tutorial:review";
        public const string NextAnchorKey = "tutorial:next";
        public const string CorrectDialogKey = "tutorial:correctDialog";
    }
}
