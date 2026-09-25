using System;
using System.IO;
using System.Windows;
using Newtonsoft.Json;

namespace MOSExcelMogiApp.Vocabulary
{
    /// <summary>
    /// クリック校正した穴を Excel ウィンドウ相対比率で永続化する。
    /// 物理ピクセルは保存せず、表示のたびに現在のウィンドウ矩形から再計算する。
    /// </summary>
    public static class VocabularyHighlightCalibration
    {
        const string FileName = "VocabHighlightCalibration.json";

        public sealed class Store
        {
            [JsonProperty("locked")]
            public bool Locked { get; set; }

            [JsonProperty("table")]
            public HoleRatio Table { get; set; }

            [JsonProperty("chart")]
            public HoleRatio Chart { get; set; }
        }

        public sealed class HoleRatio
        {
            [JsonProperty("left")]
            public double Left { get; set; }

            [JsonProperty("top")]
            public double Top { get; set; }

            [JsonProperty("width")]
            public double Width { get; set; }

            [JsonProperty("height")]
            public double Height { get; set; }

            public bool IsValid =>
                Width >= 0.02 && Height >= 0.02
                && Left >= -0.05 && Top >= -0.05
                && Left + Width <= 1.05 && Top + Height <= 1.05;
        }

        public static string AppDataPath =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MOSapp",
                FileName);

        public static Store Load()
        {
            foreach (var path in CandidatePaths())
            {
                try
                {
                    if (!File.Exists(path)) continue;
                    string json = File.ReadAllText(path);
                    var store = JsonConvert.DeserializeObject<Store>(json);
                    if (store != null) return store;
                }
                catch { }
            }
            return null;
        }

        public static void Save(Store store)
        {
            if (store == null) return;
            string json = JsonConvert.SerializeObject(store, Formatting.Indented);

            foreach (var path in CandidateWritePaths())
            {
                try
                {
                    string dir = Path.GetDirectoryName(path);
                    if (!string.IsNullOrEmpty(dir))
                        Directory.CreateDirectory(dir);
                    File.WriteAllText(path, json);
                }
                catch { }
            }
        }

        /// <summary>壊れた校正（仮ウィンドウサイズで保存されたもの等）を無効化。</summary>
        public static void InvalidateAll()
        {
            var empty = new Store { Locked = false, Table = null, Chart = null };
            Save(empty);
            try
            {
                if (File.Exists(AppDataPath))
                    File.WriteAllText(AppDataPath, JsonConvert.SerializeObject(empty, Formatting.Indented));
            }
            catch { }
        }

        public static HoleRatio ToRatio(Rect physicalHole, Rect windowPhysical)
        {
            double ww = Math.Max(1, windowPhysical.Width);
            double wh = Math.Max(1, windowPhysical.Height);
            return new HoleRatio
            {
                Left = (physicalHole.X - windowPhysical.X) / ww,
                Top = (physicalHole.Y - windowPhysical.Y) / wh,
                Width = physicalHole.Width / ww,
                Height = physicalHole.Height / wh,
            };
        }

        public static Rect? ToPhysical(HoleRatio ratio, Rect windowPhysical)
        {
            if (ratio == null || !ratio.IsValid) return null;
            if (windowPhysical.Width < 50 || windowPhysical.Height < 50) return null;
            return new Rect(
                windowPhysical.X + ratio.Left * windowPhysical.Width,
                windowPhysical.Y + ratio.Top * windowPhysical.Height,
                Math.Max(24, ratio.Width * windowPhysical.Width),
                Math.Max(24, ratio.Height * windowPhysical.Height));
        }

        /// <summary>hwnd が無効なときは null（1920x1080 の仮値は使わない）。</summary>
        public static Rect? TryGetExcelWindowPhysical(IntPtr excelHwnd)
        {
            if (excelHwnd == IntPtr.Zero) return null;
            try
            {
                if (!GetWindowRect(excelHwnd, out NativeRect wr)) return null;
                int w = wr.Right - wr.Left;
                int h = wr.Bottom - wr.Top;
                if (w < 50 || h < 50) return null;
                return new Rect(wr.Left, wr.Top, w, h);
            }
            catch
            {
                return null;
            }
        }

        public static Rect? ResolveHole(HoleRatio ratio, IntPtr excelHwnd)
        {
            var win = TryGetExcelWindowPhysical(excelHwnd);
            if (!win.HasValue) return null;
            return ToPhysical(ratio, win.Value);
        }

        static string[] CandidatePaths()
        {
            // AppData を優先（ユーザー校正）。参照 JSON はフォールバック。
            var list = new System.Collections.Generic.List<string> { AppDataPath };
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                list.Add(Path.Combine(baseDir, "References", "JSON", FileName));
                list.Add(Path.Combine(baseDir, FileName));
            }
            catch { }
            return list.ToArray();
        }

        static string[] CandidateWritePaths()
        {
            var list = new System.Collections.Generic.List<string> { AppDataPath };
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                list.Add(Path.Combine(baseDir, "References", "JSON", FileName));
            }
            catch { }
            return list.ToArray();
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out NativeRect lpRect);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        struct NativeRect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
    }
}
