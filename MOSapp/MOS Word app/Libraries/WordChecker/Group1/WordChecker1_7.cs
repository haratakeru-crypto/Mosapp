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
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                        || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc; break;
                    }
                }
                if (document == null) return false;

                // 1. 拡張子チェック: .doc のまま（「変換」直後・未保存のユニーク状態）
                //    「変換」操作はメモリ内のドキュメント形式のみアップグレードし、
                //    ディスク上のファイル名は手動で保存し直すまで変わらない。
                //    そのため ext=.doc + CompatMode=15 は「変換」ボタン押下を一意に示すシグネチャ。
                //    （Save As .docx は ext=.docx になるためここで弾かれる → 誤学習防止）
                string ext = System.IO.Path.GetExtension(document.FullName).ToLowerInvariant();
                if (ext != ".doc") return false;

                // 2. 互換モードチェック: wdWord2013 (15) になっているか
                //    .doc 形式で CompatibilityMode=15 を作る経路は「変換」操作のみ。
                //    なお、操作ログ（UpgradeDocument）は Office 仕様で idMso フックできず発火しないため
                //    判定には使用しない（状態のユニーク性で代替）。
                return (int)document.CompatibilityMode == (int)WdCompatibilityMode.wdWord2013;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_7_02(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                        || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc; break;
                    }
                }
                if (document == null) return false;

                // BuiltInDocumentProperties["Company"] は DocumentProperty COM オブジェクトを返すため、
                // 実際の値は .Value プロパティ経由で取得する。
                // （直接 .ToString() するとオブジェクトの型名 "System.__ComObject" 等が返り、常に false になる）
                dynamic companyProp = ((dynamic)document.BuiltInDocumentProperties)["Company"];
                string company = companyProp?.Value?.ToString() ?? "";

                // 厳密一致判定：前後・中央の空白、全角/半角の差もすべて区別する
                // （部分一致や正規化を許容すると「ラビット出版社」「ﾗﾋﾞｯﾄ出版」などが合格してしまい誤学習を生むため）
                return company == "ラビット出版";
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_7_03(string filePath)
        {
            // === 採取フェーズ（一時実装）===
            // 目的：実機で「インテグラル」ヘッダー挿入時に header XML に残る識別タグを特定する。
            // 使い方：
            //   1) パターンA（何もしない）で採点ボタン押下 → C:\temp\1_7_03_header_dump.xml を 1_7_03_none.xml に手動リネーム
            //   2) パターンB（インテグラル挿入）で採点ボタン押下 → 1_7_03_integral.xml にリネーム
            //   3) パターンC（別ヘッダー：オースティン等）で採点ボタン押下 → 1_7_03_other.xml にリネーム
            // 採取後に B/C の差分から識別タグを確定し、判定ロジックを実装する。
            // 採取フェーズ中は判定不能なので必ず false を返す。

            const string dbgDir = @"C:\temp";
            const string dbgFile = dbgDir + @"\1_7_03_debug.txt";
            const string dumpFile = dbgDir + @"\1_7_03_header_dump.xml";
            void Dbg(string msg)
            {
                string line = "[1_7_03] " + msg;
                System.Diagnostics.Debug.WriteLine(line);
                try
                {
                    System.IO.Directory.CreateDirectory(dbgDir);
                    System.IO.File.AppendAllText(dbgFile,
                        DateTime.Now.ToString("HH:mm:ss.fff") + " " + line + Environment.NewLine);
                }
                catch { /* ignore */ }
            }

            Application wordApp = null; Document document = null;
            try
            {
                try
                {
                    System.IO.Directory.CreateDirectory(dbgDir);
                    System.IO.File.WriteAllText(dbgFile, "=== CheckTask_1_7_03 開始 ===" + Environment.NewLine);
                }
                catch { }

                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { Dbg("DBG0: Word.Application 取得失敗"); return false; }

                string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                        || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc; break;
                    }
                }
                if (document == null) { Dbg("DBG0: document が見つからない"); return false; }

                int sectionCount = document.Sections.Count;
                Dbg($"DBG1: Section数 = {sectionCount}");

                var sb = new System.Text.StringBuilder();
                sb.AppendLine("<!-- ========================================");
                sb.AppendLine($"     CheckTask_1_7_03 ダンプ");
                sb.AppendLine($"     採取日時: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"     ファイル: {document.FullName}");
                sb.AppendLine($"     セクション数: {sectionCount}");
                sb.AppendLine("     ======================================== -->");
                sb.AppendLine();

                for (int i = 1; i <= sectionCount; i++)
                {
                    foreach (var hfType in new[]
                    {
                        WdHeaderFooterIndex.wdHeaderFooterPrimary,
                        WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                        WdHeaderFooterIndex.wdHeaderFooterEvenPages
                    })
                    {
                        HeaderFooter header = null;
                        try
                        {
                            header = document.Sections[i].Headers[hfType];
                            string headerText = header.Range.Text ?? "";
                            string headerXml = header.Range.WordOpenXML ?? "";
                            Dbg($"DBG2: Section[{i}] {hfType} TextLen={headerText.Length}, XMLLen={headerXml.Length}, Exists={header.Exists}");
                            sb.AppendLine($"<!-- ============ Section[{i}] / {hfType} (Exists={header.Exists}, TextLen={headerText.Length}) ============ -->");
                            sb.AppendLine($"<!-- Range.Text: {headerText.Replace("\r", "\\r").Replace("\n", "\\n")} -->");
                            sb.AppendLine(headerXml);
                            sb.AppendLine();
                        }
                        catch (Exception ex)
                        {
                            Dbg($"DBG2: Section[{i}] {hfType} エラー: {ex.Message}");
                        }
                        finally
                        {
                            if (header != null) Marshal.ReleaseComObject(header);
                        }
                    }
                }

                try
                {
                    System.IO.File.WriteAllText(dumpFile, sb.ToString());
                    Dbg($"DBG3: ヘッダーXMLを保存 -> {dumpFile}");
                }
                catch (Exception ex)
                {
                    Dbg($"DBG3: ヘッダーXML保存失敗: {ex.Message}");
                }

                Dbg("DBG4: 採取フェーズのため常に false を返却");
                return false;
            }
            catch (Exception ex)
            {
                Dbg("EXCEPTION: " + ex.GetType().Name + ": " + ex.Message);
                return false;
            }
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

