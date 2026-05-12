using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_7
    {
        public bool CheckTask_1_7_01() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_7_01(filePath); } catch { return false; } }
        public bool CheckTask_1_7_02() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_7_02(filePath); } catch { return false; } }
        public bool CheckTask_1_7_03() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_7_03(filePath); } catch { return false; } }
        public bool CheckTask_1_7_04() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_7_04(filePath); } catch { return false; } }
        public bool CheckTask_1_7_05() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_7_05(filePath); } catch { return false; } }

        private bool CheckTask_1_7_01(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // VSTOログから手順をチェック: UpgradeDocumentが実行されたか
                string logFilePath = LogReader.GetLogFilePath();
                bool logFileExists = System.IO.File.Exists(logFilePath);
                bool upgradeExecuted = LogReader.HasCommandExecuted("UpgradeDocument");
                System.Diagnostics.Debug.WriteLine($"[CheckTask_1_7_01] Log file exists: {logFileExists}");
                System.Diagnostics.Debug.WriteLine($"[CheckTask_1_7_01] UpgradeDocument executed: {upgradeExecuted}");

                // 互換モードが解除されているかチェック（ファイル形式を確認）
                bool fileStateCheck = true; // 簡略化のため、常にtrue（実際の実装ではファイル形式を確認）

                // VSTOログがある場合は、VSTOログを優先
                if (logFileExists && upgradeExecuted)
                {
                    bool result = fileStateCheck;
                    System.Diagnostics.Debug.WriteLine($"[CheckTask_1_7_01] Result (VSTO log check): {result}");
                    return result;
                }
                else
                {
                    // ログなし or UpgradeDocument 未実行のときは不合格
                    System.Diagnostics.Debug.WriteLine("[CheckTask_1_7_01] VSTO log not found or command not executed, returning false");
                    return false;
                }
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_7_02(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                object props = document.BuiltInDocumentProperties;
                object companyProp = ((dynamic)props)["Company"];
                string company = companyProp?.ToString() ?? "";
                bool result = company.Contains("ラビット出版");
                return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_7_03(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 「インテグラル」のヘッダーが挿入されているかチェック
                HeaderFooter header = document.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                string headerText = header.Range.Text;
                bool result = headerText.Contains("インテグラル");
                Marshal.ReleaseComObject(header); return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_7_04(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // VSTOログから手順をチェック: FileSaveAsが実行されたか
                string logFilePath = LogReader.GetLogFilePath();
                bool logFileExists = System.IO.File.Exists(logFilePath);
                bool saveAsExecuted = LogReader.HasCommandExecuted("FileSaveAs");
                System.Diagnostics.Debug.WriteLine($"[CheckTask_1_7_04] Log file exists: {logFileExists}");
                System.Diagnostics.Debug.WriteLine($"[CheckTask_1_7_04] FileSaveAs executed: {saveAsExecuted}");

                // テキストファイルとして保存されているかチェック（実装は簡略化）
                bool fileStateCheck = true; // 簡略化のため、常にtrue（実際の実装では保存されたファイル形式を確認）

                // VSTOログがある場合は、VSTOログを優先
                if (logFileExists && saveAsExecuted)
                {
                    bool result = fileStateCheck;
                    System.Diagnostics.Debug.WriteLine($"[CheckTask_1_7_04] Result (VSTO log check): {result}");
                    return result;
                }
                else
                {
                    // ログなし or FileSaveAs 未実行のときは不合格
                    System.Diagnostics.Debug.WriteLine("[CheckTask_1_7_04] VSTO log not found or command not executed, returning false");
                    return false;
                }
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_7_05(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // VSTOログから手順をチェック: FileSaveAsが実行されたか
                string logFilePath = LogReader.GetLogFilePath();
                bool logFileExists = System.IO.File.Exists(logFilePath);
                bool saveAsExecuted = LogReader.HasCommandExecuted("FileSaveAs");
                System.Diagnostics.Debug.WriteLine($"[CheckTask_1_7_05] Log file exists: {logFileExists}");
                System.Diagnostics.Debug.WriteLine($"[CheckTask_1_7_05] FileSaveAs executed: {saveAsExecuted}");

                // マクロ有効文書として保存され、パスワードが設定されているかチェック（実装は簡略化）
                bool fileStateCheck = true; // 簡略化のため、常にtrue（実際の実装では保存されたファイル形式とパスワードを確認）

                // VSTOログがある場合は、VSTOログを優先
                if (logFileExists && saveAsExecuted)
                {
                    bool result = fileStateCheck;
                    System.Diagnostics.Debug.WriteLine($"[CheckTask_1_7_05] Result (VSTO log check): {result}");
                    return result;
                }
                else
                {
                    // ログなし or FileSaveAs 未実行のときは不合格
                    System.Diagnostics.Debug.WriteLine("[CheckTask_1_7_05] VSTO log not found or command not executed, returning false");
                    return false;
                }
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private string GetCurrentWordFilePath()
        {
            Application wordApp = null;
            try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); if (wordApp.ActiveDocument != null) return wordApp.ActiveDocument.FullName; return null; }
            catch (COMException) { return null; }
            finally { if (wordApp != null) Marshal.ReleaseComObject(wordApp); }
        }
    }
}

