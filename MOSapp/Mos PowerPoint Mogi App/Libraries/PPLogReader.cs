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
            string path = GetLogFilePath();
            if (!File.Exists(path))
                return false;
            try
            {
                return File.ReadAllLines(path).Any(line =>
                    line != null && line.IndexOf("[Task10-4] Grayscale", StringComparison.OrdinalIgnoreCase) >= 0);
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
