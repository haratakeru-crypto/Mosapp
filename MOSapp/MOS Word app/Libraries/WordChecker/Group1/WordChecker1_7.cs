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

        private const string P7CompanyTarget = "ラビット出版";
        private const string P7RdTxtFileName = "朗読会.txt";
        private const string P7RdDocmFileName = "朗読会.docm";
        private const string P7ReadPassword = "abc";

        private bool CheckTask_1_7_02(string filePath)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath)) return false;
                if (!LogReader.HasTaskEvidence(7, 2, "SetDocumentCompany"))
                    return false;

                // 主判定: VSTO が Project7 上で Company が目標値へ遷移したとき記録した SetDocumentCompany のみ（7-3/7-4/7-5 と同型の証跡中心）。
                // 7-5 パスワード誤り等で docm/txt から Company が取れなくても 7-2 を巻き添えにしない。
                LogCompanyStateDiagnosticsIfNeeded(filePath);
                return true;
            }
            catch { return false; }
        }

        /// <summary>
        /// 7-2 補助: 成果物から Company を読めるか監視用。採点結果には影響しない（非ブロッキング）。
        /// </summary>
        private void LogCompanyStateDiagnosticsIfNeeded(string filePath)
        {
            try
            {
                if (!TryGetCompanyForTask7_02(filePath, out string company))
                {
                    System.Diagnostics.Debug.WriteLine(
                        "[WordChecker1_7] 7-2: SetDocumentCompany あり。状態読取不可（採点は○のまま）");
                    return;
                }

                if (company != P7CompanyTarget)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordChecker1_7] 7-2: SetDocumentCompany あり。状態 Company=\"{company}\"（目標と不一致・採点は○のまま）");
                }
            }
            catch { }
        }

        /// <summary>
        /// 7-2: 採点用 Company 取得（補助・デバッグ用）。docm → txt → 開いている Project7。
        /// </summary>
        private bool TryGetCompanyForTask7_02(string filePath, out string company)
        {
            company = "";
            Application wordApp = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { return false; }

                string dir = Path.GetDirectoryName(filePath);
                string docmPath = string.IsNullOrEmpty(dir) ? null : Path.Combine(dir, P7RdDocmFileName);
                string txtPath = string.IsNullOrEmpty(dir) ? null : Path.Combine(dir, P7RdTxtFileName);
                string project7Path = ResolveProject7PathForTask702(filePath, dir);

                string derivedMismatch = null;

                // 1) 朗読会.docm（7-5 後・一括採点。読み取りパスワード abc）
                if (!string.IsNullOrEmpty(docmPath) && TryGetCompanyFromDocm(wordApp, docmPath, out string fromDocm))
                {
                    if (fromDocm == P7CompanyTarget)
                    {
                        company = fromDocm;
                        return true;
                    }
                    derivedMismatch = fromDocm;
                }

                // 2) 朗読会.txt（7-4 後〜7-5 前）
                if (!string.IsNullOrEmpty(txtPath) && TryGetCompanyFromTxt(wordApp, txtPath, filePath, out string fromTxt))
                {
                    if (fromTxt == P7CompanyTarget)
                    {
                        company = fromTxt;
                        return true;
                    }
                    if (derivedMismatch == null)
                        derivedMismatch = fromTxt;
                }

                // 3) 開いている Project7.doc（7-2 復習で修正した値。朗読会.* が誤値のときのフォールバック）
                if (!string.IsNullOrEmpty(project7Path)
                    && TryGetCompanyFromOpenDocuments(wordApp, project7Path, "Project7.doc", out string fromProject7))
                {
                    company = fromProject7;
                    return true;
                }

                if (derivedMismatch != null)
                {
                    company = derivedMismatch;
                    return true;
                }

                return false;
            }
            catch { return false; }
            finally
            {
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        private static string ResolveProject7PathForTask702(string filePath, string dir)
        {
            if (!string.IsNullOrEmpty(filePath))
            {
                string name = Path.GetFileName(filePath);
                if (name != null && System.Text.RegularExpressions.Regex.IsMatch(
                        Path.GetFileNameWithoutExtension(name), @"^project\s*7$",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    return filePath;
            }
            if (string.IsNullOrEmpty(dir)) return null;
            string p7 = Path.Combine(dir, "Project7.doc");
            return File.Exists(p7) ? p7 : filePath;
        }

        private static bool TryGetCompanyFromDocm(Application wordApp, string docmPath, out string company)
        {
            company = "";
            if (TryGetCompanyFromOpenDocuments(wordApp, docmPath, P7RdDocmFileName, out company))
                return true;
            if (File.Exists(docmPath) && TryReadCompanyFromDocmOnDisk(wordApp, docmPath, out company))
                return true;
            return false;
        }

        private static bool TryGetCompanyFromTxt(Application wordApp, string txtPath, string activeFilePath, out string company)
        {
            company = "";
            if (TryGetCompanyFromOpenDocuments(wordApp, txtPath, P7RdTxtFileName, out company))
                return true;

            if (!string.IsNullOrEmpty(activeFilePath)
                && string.Equals(Path.GetFileName(activeFilePath), P7RdTxtFileName, StringComparison.OrdinalIgnoreCase)
                && TryGetCompanyFromOpenDocuments(wordApp, activeFilePath, P7RdTxtFileName, out company))
                return true;

            if (File.Exists(txtPath) && TryReadCompanyFromTxtOnDisk(wordApp, txtPath, out company))
                return true;
            return false;
        }

        private static bool TryReadCompanyFromTxtOnDisk(Application wordApp, string txtPath, out string company)
        {
            company = "";
            Document opened = null;
            WdAlertLevel originalAlerts = wordApp.DisplayAlerts;
            try
            {
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                opened = wordApp.Documents.Open(
                    FileName: txtPath,
                    ConfirmConversions: false,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false);
                return TryGetCompanyFromDocument(opened, out company);
            }
            catch { return false; }
            finally
            {
                wordApp.DisplayAlerts = originalAlerts;
                if (opened != null)
                {
                    try { opened.Close(WdSaveOptions.wdDoNotSaveChanges); }
                    catch { }
                    Marshal.ReleaseComObject(opened);
                }
            }
        }

        private static bool TryGetCompanyFromOpenDocuments(Application wordApp, string fullPath, string fileName, out string company)
        {
            company = "";
            Document matched = null;
            try
            {
                foreach (Document doc in wordApp.Documents)
                {
                    try
                    {
                        if (doc.FullName.Equals(fullPath, StringComparison.OrdinalIgnoreCase)
                            || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                        {
                            matched = doc;
                            break;
                        }
                    }
                    catch { }
                }
                if (matched == null) return false;
                return TryGetCompanyFromDocument(matched, out company);
            }
            finally
            {
                if (matched != null) Marshal.ReleaseComObject(matched);
            }
        }

        /// <summary>7-2: ディスク上の朗読会.docm をサイレントオープン。7-5 の読み取りパスワード付きは一時コピー＋ abc で開く（7-5 チェッカーと同型）。</summary>
        private static bool TryReadCompanyFromDocmOnDisk(Application wordApp, string docmPath, out string company)
        {
            company = "";
            bool encrypted = AppearsEncryptedByReadPassword(docmPath);
            string pathToOpen = docmPath;
            string tempPath = null;

            if (encrypted)
            {
                tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docm");
                try
                {
                    using (var fsIn = new FileStream(docmPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var fsOut = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                    {
                        fsIn.CopyTo(fsOut);
                    }
                    pathToOpen = tempPath;
                }
                catch { return false; }
            }

            Document opened = null;
            WdAlertLevel originalAlerts = wordApp.DisplayAlerts;
            try
            {
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                if (encrypted)
                {
                    opened = wordApp.Documents.Open(
                        FileName: pathToOpen,
                        ConfirmConversions: false,
                        ReadOnly: true,
                        AddToRecentFiles: false,
                        PasswordDocument: P7ReadPassword,
                        Visible: false);
                }
                else
                {
                    try
                    {
                        opened = wordApp.Documents.Open(
                            FileName: pathToOpen,
                            ConfirmConversions: false,
                            ReadOnly: true,
                            AddToRecentFiles: false,
                            Visible: false);
                    }
                    catch (COMException)
                    {
                        opened = wordApp.Documents.Open(
                            FileName: pathToOpen,
                            ConfirmConversions: false,
                            ReadOnly: true,
                            AddToRecentFiles: false,
                            PasswordDocument: P7ReadPassword,
                            Visible: false);
                    }
                }

                return TryGetCompanyFromDocument(opened, out company);
            }
            catch { return false; }
            finally
            {
                wordApp.DisplayAlerts = originalAlerts;
                if (opened != null)
                {
                    try { opened.Close(WdSaveOptions.wdDoNotSaveChanges); }
                    catch { }
                    Marshal.ReleaseComObject(opened);
                }
                if (tempPath != null)
                {
                    try
                    {
                        if (File.Exists(tempPath))
                            File.Delete(tempPath);
                    }
                    catch { }
                }
            }
        }

        private static bool TryGetCompanyFromDocument(Document document, out string company)
        {
            company = "";
            try
            {
                dynamic companyProp = ((dynamic)document.BuiltInDocumentProperties)["Company"];
                company = companyProp?.Value?.ToString() ?? "";
                return true;
            }
            catch { return false; }
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
                // 拡張子 .doc のヘッダーは互換11/15で XML が異なる（Reference/XML: ED7D31 直書き vs themeFill accent2）。tbl+gridCol 補助指紋を OR する。
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
                usePostCompatFingerprint = ext == ".doc";
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

        /// <summary>インテグラル帯のオレンジ（accent2）。互換11 は fill="ED7D31" 直書き、互換15 は themeFill 等。</summary>
        private static bool HasIntegralAccent2ColorMarker(string xml)
        {
            if (string.IsNullOrEmpty(xml)) return false;
            return xml.IndexOf("fill=\"E97132\"", StringComparison.Ordinal) >= 0
                || xml.IndexOf("fill=\"ED7D31\"", StringComparison.Ordinal) >= 0
                || xml.IndexOf("w:themeFill=\"accent2\"", StringComparison.OrdinalIgnoreCase) >= 0
                || xml.IndexOf("w:themeColor=\"accent2\"", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool HasIntegralStructureFingerprint(string xml)
        {
            if (!HasIntegralAccent2ColorMarker(xml)) return false;
            if (xml.IndexOf("w:w=\"1782\"", StringComparison.Ordinal) < 0) return false;
            if (xml.IndexOf("w:w=\"7286\"", StringComparison.Ordinal) < 0) return false;
            return true;
        }

        private static bool HasIntegralStructureFingerprintAfterCompat(string xml)
        {
            if (xml.IndexOf("<w:tbl", StringComparison.OrdinalIgnoreCase) < 0) return false;
            if (!HasIntegralAccent2ColorMarker(xml)) return false;
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

