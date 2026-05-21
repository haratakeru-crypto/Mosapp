using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
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
                bool stateOk = ext == ".doc"
                    && (int)document.CompatibilityMode == (int)WdCompatibilityMode.wdWord2013;

                // 2. 後続タスク（別名保存で .txt 等）のあと ext が変わり stateOk が false になり得る。
                //    5-1 と同様「現在が合格」または「過去に互換解除したログがある」なら合格。
                //    UpgradeDocument は Ribbon の idMso では発火しないが、ThisAddIn ポーリングで .doc 上の非15→15 遷移時に記録する。
                bool logOk = LogReader.HasTaskEvidence(7, 1, "UpgradeDocument");
                return stateOk || logOk;
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
            Application wordApp = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                string fileName = Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                        || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        break;
                    }
                }
                if (document == null) return false;

                // SaveCopyAs + System.IO.Packaging でのパッケージ読取は、一部環境で参照が解決せずビルド不能となるため採用しない。
                // 全セクション×各ヘッダー種別の WordOpenXML のみで判定する（論点1: いずれかのヘッダーで一致すれば○）。
                // 7-1（互換15）かつ拡張子が .doc のとき（変換後～別名保存前）、tblGrid twip 等が変わり得るため補助指紋を OR する。
                // 7-4 後などヘッダーが読めない／状態が変わったあとも、VSTO ポーリングで記録した IntegralHeader ログがあれば ○（7-1 と同型）。
                return TryIntegralFromLiveHeaderWordOpenXml(document)
                    || LogReader.HasTaskEvidence(7, 3, "IntegralHeader");
            }
            catch { return false; }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
            }
        }

        private bool CheckTask_1_7_04(string filePath)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath)) return false;

                // 元 .docx と同一フォルダに「朗読会.txt」が存在するか（課題: 書式なしテキストで保存）
                string dir = Path.GetDirectoryName(filePath);
                if (string.IsNullOrEmpty(dir)) return false;
                string targetTxt = Path.Combine(dir, "朗読会.txt");
                if (!File.Exists(targetTxt)) return false;

                // 保存時の操作ログが存在するか
                if (!LogReader.HasTaskEvidence(7, 4, "FileSaveAsTxt")) return false;

                return true;
            }
            catch { return false; }
        }

        private bool CheckTask_1_7_05(string filePath)
        {
            Application wordApp = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                string fileName = Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                {
                    if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase)
                        || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        document = doc;
                        break;
                    }
                }
                if (document == null) return false;

                // 7-4 と同フォルダ・同ベース名（朗読会）の .docm（課題: マクロ有効＋読み取りパスワード）
                string dir = Path.GetDirectoryName(filePath);
                if (string.IsNullOrEmpty(dir)) return false;
                string targetDocm = Path.Combine(dir, "朗読会.docm");
                if (!File.Exists(targetDocm)) return false;
                if (!AppearsEncryptedByReadPassword(targetDocm)) return false;
                if (!LogReader.HasTaskEvidence(7, 5, "FileSaveAsDocm")) return false;

                // パスワードが「abc」に正しく設定されているかを一時コピーファイルで安全にサイレント検証
                WdAlertLevel originalAlertLevel = wordApp.DisplayAlerts;
                Document tempDoc = null;
                bool passwordOk = false;
                string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docm");
                try
                {
                    // 画面上のWordと干渉しないよう、ファイルを一時フォルダに安全に複製
                    using (var fsIn = new FileStream(targetDocm, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var fsOut = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                    {
                        fsIn.CopyTo(fsOut);
                    }

                    wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                    
                    // 非表示かつ読み取り専用で、複製した一時ファイルをパスワード "abc" でオープン試行
                    tempDoc = wordApp.Documents.Open(
                        FileName: tempPath,
                        ConfirmConversions: false,
                        ReadOnly: true,
                        AddToRecentFiles: false,
                        PasswordDocument: "abc",
                        Visible: false
                    );
                    
                    passwordOk = true;
                }
                catch (COMException)
                {
                    // パスワードが違う、または未設定の場合は例外が発生するため不合格
                    passwordOk = false;
                }
                finally
                {
                    if (tempDoc != null)
                    {
                        try { tempDoc.Close(SaveChanges: false); } catch { }
                        Marshal.ReleaseComObject(tempDoc);
                    }
                    // 警告レベルを復元
                    try { wordApp.DisplayAlerts = originalAlertLevel; } catch { }
                    // 一時ファイルの確実な削除
                    try
                    {
                        if (File.Exists(tempPath))
                        {
                            File.Delete(tempPath);
                        }
                    }
                    catch { }
                }

                return passwordOk;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>
        /// 読み取りパスワード付き保存の目安: 暗号化 OOXML は ZIP ではなく CFB（先頭 D0 CF 11 E0）。
        /// パスワード文字列「abc」の一致までは検証しない。
        /// </summary>
        private static bool AppearsEncryptedByReadPassword(string path)
        {
            byte[] header = new byte[4];
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                if (fs.Read(header, 0, 4) < 4) return false;
            }

            // 暗号化 Office 文書（Compound File Binary）
            if (header[0] == 0xD0 && header[1] == 0xCF && header[2] == 0x11 && header[3] == 0xE0)
                return true;

            // パスワードなし .docm は ZIP（PK..）
            if (header[0] == 0x50 && header[1] == 0x4B)
                return false;

            return false;
        }

        private string GetCurrentWordFilePath()
        {
            Application wordApp = null;
            try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); if (wordApp.ActiveDocument != null) return wordApp.ActiveDocument.FullName; return null; }
            catch (COMException) { return null; }
            finally { if (wordApp != null) Marshal.ReleaseComObject(wordApp); }
        }

        private static bool TryIntegralFromLiveHeaderWordOpenXml(Document document)
        {
            bool usePostCompatFingerprint = false;
            try
            {
                string ext = Path.GetExtension(document.FullName).ToLowerInvariant();
                usePostCompatFingerprint = ext == ".doc"
                    && (int)document.CompatibilityMode == (int)WdCompatibilityMode.wdWord2013;
            }
            catch { /* ignore */ }

            try
            {
                int sectionCount = document.Sections.Count;
                for (int i = 1; i <= sectionCount; i++)
                {
                    foreach (WdHeaderFooterIndex hfType in new[]
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
                            if (!header.Exists) continue;
                            string headerXml = header.Range.WordOpenXML ?? "";
                            if (HeaderXmlLooksLikeIntegral(headerXml, usePostCompatFingerprint))
                                return true;
                        }
                        finally
                        {
                            if (header != null) Marshal.ReleaseComObject(header);
                        }
                    }
                }
            }
            catch { /* ignore */ }
            return false;
        }

        private static bool HeaderXmlLooksLikeIntegral(string xml, bool usePostCompatFingerprint)
        {
            if (string.IsNullOrEmpty(xml)) return false;
            if (ContainsIntegralBuildingBlockMetadata(xml)) return true;
            if (HasIntegralStructureFingerprint(xml)) return true;
            if (usePostCompatFingerprint && HasIntegralStructureFingerprintAfterCompat(xml)) return true;
            return false;
        }

        private static bool ContainsIntegralBuildingBlockMetadata(string xml)
        {
            if (xml.IndexOf("Integral", StringComparison.OrdinalIgnoreCase) < 0) return false;
            if (Regex.IsMatch(xml, @"w:val\s*=\s*""Integral""", RegexOptions.IgnoreCase)) return true;
            if (xml.IndexOf("docPart", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (xml.IndexOf("w:sdt", StringComparison.Ordinal) >= 0) return true;
            return false;
        }

        private static bool HasIntegralStructureFingerprint(string xml)
        {
            if (xml.IndexOf("fill=\"E97132\"", StringComparison.Ordinal) < 0) return false;
            if (xml.IndexOf("w:w=\"1782\"", StringComparison.Ordinal) < 0) return false;
            if (xml.IndexOf("w:w=\"7286\"", StringComparison.Ordinal) < 0) return false;
            return true;
        }

        private static bool HasIntegralStructureFingerprintAfterCompat(string xml)
        {
            if (xml.IndexOf("<w:tbl", StringComparison.OrdinalIgnoreCase) < 0) return false;
            bool accent2Marker =
                xml.IndexOf("fill=\"E97132\"", StringComparison.Ordinal) >= 0
                || xml.IndexOf("w:themeFill=\"accent2\"", StringComparison.OrdinalIgnoreCase) >= 0
                || xml.IndexOf("w:themeColor=\"accent2\"", StringComparison.OrdinalIgnoreCase) >= 0;
            if (!accent2Marker) return false;
            int gridColCount = 0;
            for (int i = 0; ;)
            {
                int p = xml.IndexOf("<w:gridCol", i, StringComparison.Ordinal);
                if (p < 0) break;
                gridColCount++;
                i = p + 10;
            }
            return gridColCount >= 2;
        }
    }
}

