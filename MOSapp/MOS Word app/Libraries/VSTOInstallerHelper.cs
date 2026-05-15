using System;
using System.IO;
using Microsoft.Win32;

namespace Libraries
{
    /// <summary>
    /// VSTOアドインのインストール状態を確認・管理するヘルパークラス
    /// </summary>
    public static class VSTOInstallerHelper
    {
        private const string AddInName = "New_MOSWordVSTOAddIn";
        private const string RegistryKeyPath = @"Software\Microsoft\Office\Word\Addins\" + AddInName;
        private const string LoadBehaviorValueName = "LoadBehavior";
        private const int LoadBehaviorEnabled = 3; // 起動時に読み込む

        /// <summary>
        /// VSTOアドインがインストールされているかチェック
        /// </summary>
        /// <returns>インストールされている場合true</returns>
        public static bool IsInstalled()
        {
            try
            {
                // レジストリを確認
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                {
                    if (key == null)
                        return false;

                    object loadBehavior = key.GetValue(LoadBehaviorValueName);
                    if (loadBehavior == null)
                        return false;

                    int loadBehaviorValue = Convert.ToInt32(loadBehavior);
                    return loadBehaviorValue == LoadBehaviorEnabled;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[VSTOInstallerHelper] Error checking registry: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// VSTOアドインのインストールパスを取得
        /// </summary>
        /// <returns>インストールパス（見つからない場合はnull）</returns>
        public static string GetInstallPath()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
                {
                    if (key == null)
                        return null;

                    object manifestPath = key.GetValue("Manifest");
                    if (manifestPath == null)
                        return null;

                    string manifestPathStr = manifestPath.ToString();
                    // .vstoファイルのパスを取得
                    if (manifestPathStr.EndsWith(".vsto", StringComparison.OrdinalIgnoreCase))
                    {
                        return manifestPathStr;
                    }

                    return null;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[VSTOInstallerHelper] Error getting install path: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// VSTOアドインのビルド出力パスを取得
        /// 実行アセンブリのディレクトリから親へ辿り、New_MOSWordVSTOAddIn を含むパスを検索する。
        /// </summary>
        /// <returns>ビルド出力パス（.vstoファイルのパス）</returns>
        public static string GetBuildOutputPath()
        {
            string vstoFileName = AddInName + ".vsto";
            string[] relativeSuffixes = new[]
            {
                Path.Combine("New_MOSWordVSTOAddIn", "New_MOSWordVSTOAddIn", "bin", "Debug", vstoFileName),
                Path.Combine("New_MOSWordVSTOAddIn", "New_MOSWordVSTOAddIn", "bin", "Release", vstoFileName),
            };

            // 1) BaseDirectory 直下および相対パス
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            foreach (string suffix in relativeSuffixes)
            {
                string vstoPath = Path.Combine(baseDir, suffix);
                if (File.Exists(vstoPath)) return vstoPath;
            }

            // 2) BaseDirectory の親を複数段さかのぼって検索
            string searchDir = Path.GetFullPath(baseDir);
            for (int i = 0; i < 8; i++)
            {
                string parent = Path.GetDirectoryName(searchDir);
                if (string.IsNullOrEmpty(parent) || parent == searchDir) break;
                searchDir = parent;
                foreach (string suffix in relativeSuffixes)
                {
                    string vstoPath = Path.Combine(searchDir, suffix);
                    if (File.Exists(vstoPath)) return vstoPath;
                }
                // New_MOSWordVSTOAddIn フォルダを含むディレクトリを探す
                string addInFolder = Path.Combine(searchDir, "New_MOSWordVSTOAddIn");
                if (Directory.Exists(addInFolder))
                {
                    foreach (string suffix in relativeSuffixes)
                    {
                        string vstoPath = Path.Combine(searchDir, suffix);
                        if (File.Exists(vstoPath)) return vstoPath;
                    }
                }
            }

            // 3) 実行中アセンブリの Location から同様に親を辿る
            string asmDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            if (!string.IsNullOrEmpty(asmDir))
            {
                searchDir = Path.GetFullPath(asmDir);
                for (int i = 0; i < 8; i++)
                {
                    foreach (string suffix in relativeSuffixes)
                    {
                        string vstoPath = Path.Combine(searchDir, suffix);
                        if (File.Exists(vstoPath)) return vstoPath;
                    }
                    string parent = Path.GetDirectoryName(searchDir);
                    if (string.IsNullOrEmpty(parent) || parent == searchDir) break;
                    searchDir = parent;
                }
            }

            // 4) カレントディレクトリ
            string currentDir = Directory.GetCurrentDirectory();
            foreach (string suffix in relativeSuffixes)
            {
                string vstoPath = Path.Combine(currentDir, suffix);
                if (File.Exists(vstoPath)) return vstoPath;
            }

            return null;
        }

        /// <summary>
        /// VSTOアドインのビルド出力ファイルが存在するかチェック
        /// </summary>
        /// <returns>ファイルが存在する場合true</returns>
        public static bool BuildOutputExists()
        {
            string vstoPath = GetBuildOutputPath();
            if (string.IsNullOrEmpty(vstoPath))
                return false;

            // .vstoファイルと.dllファイルの両方が存在するか確認
            string dllPath = Path.Combine(Path.GetDirectoryName(vstoPath), $"{AddInName}.dll");
            return File.Exists(vstoPath) && File.Exists(dllPath);
        }

        /// <summary>
        /// VSTOアドインのインストール状態を詳細にチェック
        /// </summary>
        /// <returns>インストール状態の詳細情報</returns>
        public static VSTOInstallStatus GetInstallStatus()
        {
            var status = new VSTOInstallStatus
            {
                IsInstalled = IsInstalled(),
                InstallPath = GetInstallPath(),
                BuildOutputPath = GetBuildOutputPath(),
                BuildOutputExists = BuildOutputExists()
            };

            return status;
        }

        /// <summary>
        /// VSTOアドインのインストール状態を表すクラス
        /// </summary>
        public class VSTOInstallStatus
        {
            public bool IsInstalled { get; set; }
            public string InstallPath { get; set; }
            public string BuildOutputPath { get; set; }
            public bool BuildOutputExists { get; set; }

            public string GetInstallationMessage()
            {
                if (IsInstalled)
                {
                    return "VSTOアドインはインストールされています。";
                }

                if (!BuildOutputExists)
                {
                    return "VSTOアドインがビルドされていません。まずプロジェクトをビルドしてください。";
                }

                return $"VSTOアドインがインストールされていません。\n\n" +
                       $"インストール方法:\n" +
                       $"1. Wordを終了してください\n" +
                       $"2. 以下のファイルをダブルクリックしてインストールしてください:\n" +
                       $"   {BuildOutputPath ?? "（パスが見つかりません）"}";
            }
        }
    }
}






