using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Libraries
{
    /// <summary>
    /// PowerPoint VSTO アドインで生成されたログファイルを読み込むユーティリティ。
    /// 採点で 10-4 グレースケール等のログを参照する。
    /// </summary>
    public static class PPLogReader
    {
        /// <summary>
        /// ログファイルのパスを取得（%TEMP%\mos_ppt_log.txt）
        /// </summary>
        public static string GetLogFilePath()
        {
            return Path.Combine(Path.GetTempPath(), "mos_ppt_log.txt");
        }

        /// <summary>
        /// VSTO アドインのログファイルをクリアする。リセット時に呼び出す。
        /// </summary>
        public static void ClearLog()
        {
            try
            {
                string path = GetLogFilePath();
                if (File.Exists(path))
                    File.Delete(path);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] Error clearing log: " + ex.Message);
            }
        }

        /// <summary>
        /// ログに [Task10-4] Grayscale が 1 行でも含まれるか（10-4 グレースケール操作の記録あり）
        /// </summary>
        public static bool HasTask10_4GrayscaleExecuted()
        {
            return HasLogLineContaining("[Task10-4] Grayscale");
        }

        /// <summary>ログに 5-1 の印刷記録（[Task5-1] Print）が含まれるか。</summary>
        public static bool HasTask5_1PrintExecuted()
        {
            return HasLogLineContaining("[Task5-1] Print");
        }

        /// <summary>ログに 11-7 の印刷記録（[Task11-7] Print）が含まれるか。</summary>
        public static bool HasTask11_7PrintExecuted()
        {
            return HasLogLineContaining("[Task11-7] Print");
        }

        /// <summary>ログに 8-4 オーディオ設定記録（[Task8-4] Audio）が含まれるか。</summary>
        public static bool HasTask8_4AudioExecuted()
        {
            return HasLogLineContaining("[Task8-4] Audio");
        }

        /// <summary>ログに 7-2 スライド再利用記録（[Task7-2] ReuseSlides）が含まれるか。</summary>
        public static bool HasTask7_2ReuseSlidesExecuted()
        {
            return HasLogLineContaining("[Task7-2] ReuseSlides");
        }

        /// <summary>ログに 7-3 アウトラインから挿入記録（[Task7-3] InsertFromOutline）が含まれるか。</summary>
        public static bool HasTask7_3InsertFromOutlineExecuted()
        {
            return HasLogLineContaining("[Task7-3] InsertFromOutline");
        }

        /// <summary>ログに 10-1 ドキュメント検査記録（[Task10-1] DocumentInspector）が含まれるか。</summary>
        public static bool HasTask10_1DocumentInspectorExecuted()
        {
            return HasLogLineContaining("[Task10-1] DocumentInspector");
        }

        /// <summary>ログに 10-7 レイアウト複製記録（[Task10-7] LayoutDuplicate）が含まれるか。</summary>
        public static bool HasTask10_7LayoutDuplicateExecuted()
        {
            return HasLogLineContaining("[Task10-7] LayoutDuplicate");
        }

        private static bool HasLogLineContaining(string marker)
        {
            string path = GetLogFilePath();
            if (!File.Exists(path))
                return false;
            try
            {
                return File.ReadAllLines(path).Any(line =>
                    line != null && line.IndexOf(marker, StringComparison.OrdinalIgnoreCase) >= 0);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] Error reading log: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// ログファイルからすべての行を読み込む（デバッグ・表示用）
        /// </summary>
        public static List<string> ReadAllLines()
        {
            var result = new List<string>();
            string path = GetLogFilePath();
            if (!File.Exists(path))
                return result;
            try
            {
                result.AddRange(File.ReadAllLines(path));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PPLogReader] Error reading log: " + ex.Message);
            }
            return result;
        }
    }
}
