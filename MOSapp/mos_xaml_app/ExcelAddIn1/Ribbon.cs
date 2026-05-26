using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace ExcelAddIn1
{
    /// <summary>
    /// ログ関連の操作を提供するリボン拡張。
    /// </summary>
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        private Microsoft.Office.Tools.CustomTaskPane _logPane;

        // Backstage（ファイルタブ）開閉時の文書プロパティ差分検知用（開いた瞬間のタスク文脈を閉じたときのログに使う）
        private bool _backstageCaptureActive;
        private string _backstageWorkbookKey;
        private Dictionary<string, string> _backstagePropertySnapshot;
        private int _backstageProjectId = -1;
        private int _backstageTaskId = -1;
        private int _backstageAttemptNo = 1;

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelAddIn1.Ribbon.xml");
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            System.Diagnostics.Debug.WriteLine("[ExcelAddIn1.Ribbon] Ribbon_Load completed");
        }

        /// <summary>Backstage 表示開始。追跡対象の文書プロパティの基準とタスク文脈を記録。</summary>
        /// <remarks>
        /// Office の backstage の onShow/onHide は <see cref="IRibbonControl"/> ではなく <c>object contextObject</c> 1 個（公式リファレンス）。
        /// </remarks>
        public void OnBackstageShow(object contextObject)
        {
            try
            {
                Excel.Workbook wb = Globals.ThisAddIn.Application?.ActiveWorkbook;
                if (wb == null)
                {
                    ClearBackstageCapture();
                    return;
                }

                Logger.GetCurrentTaskContext(out int p, out int t, out int a);
                _backstageProjectId = p;
                _backstageTaskId = t;
                _backstageAttemptNo = a;
                _backstageWorkbookKey = GetWorkbookKey(wb);
                _backstagePropertySnapshot = CaptureWorkbookProperties(wb);
                _backstageCaptureActive = true;
                System.Diagnostics.Debug.WriteLine($"[ExcelAddIn1.Ribbon] OnBackstageShow wb={_backstageWorkbookKey} task={p}-{t}-{a}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelAddIn1.Ribbon] OnBackstageShow: " + ex.Message);
                ClearBackstageCapture();
            }
        }

        /// <summary>Backstage 終了時に追跡プロパティの差分があれば SetWorkbookProperty を記録（開いたときのタスク文脈で）。詳細は変更キーをカンマ区切り。</summary>
        public void OnBackstageHide(object contextObject)
        {
            try
            {
                if (!_backstageCaptureActive)
                    return;

                Excel.Workbook wb = Globals.ThisAddIn.Application?.ActiveWorkbook;
                if (wb == null)
                {
                    ClearBackstageCapture();
                    return;
                }

                string keyNow = GetWorkbookKey(wb);
                if (!string.Equals(keyNow, _backstageWorkbookKey, StringComparison.Ordinal))
                {
                    System.Diagnostics.Debug.WriteLine($"[ExcelAddIn1.Ribbon] OnBackstageHide workbook key mismatch (before={_backstageWorkbookKey}, now={keyNow})");
                    ClearBackstageCapture();
                    return;
                }

                if (_backstagePropertySnapshot == null)
                {
                    ClearBackstageCapture();
                    return;
                }

                Dictionary<string, string> now = CaptureWorkbookProperties(wb);
                List<string> changed = FindChangedPropertyKeys(_backstagePropertySnapshot, now);
                if (changed.Count > 0 && _backstageProjectId > 0 && _backstageTaskId > 0)
                {
                    string wbName = SafeWorkbookName(wb);
                    string detail = string.Join(",", changed);
                    Logger.RunWithTaskContext(_backstageProjectId, _backstageTaskId, _backstageAttemptNo, () =>
                    {
                        Logger.LogOperation("SetWorkbookProperty", $"{wbName}!{detail}");
                    });
                    System.Diagnostics.Debug.WriteLine("[ExcelAddIn1.Ribbon] OnBackstageHide logged SetWorkbookProperty: " + detail);
                }

                ClearBackstageCapture();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelAddIn1.Ribbon] OnBackstageHide: " + ex.Message);
                ClearBackstageCapture();
            }
        }

        private void ClearBackstageCapture()
        {
            _backstageCaptureActive = false;
            _backstageWorkbookKey = null;
            _backstagePropertySnapshot = null;
            _backstageProjectId = -1;
            _backstageTaskId = -1;
            _backstageAttemptNo = 1;
        }

        /// <summary>Backstage の「情報」で編集されやすい組み込みプロパティと、カスタム &quot;Tags&quot;。</summary>
        /// <remarks>Excel の <see cref="Excel.Workbook.BuiltinDocumentProperties"/> は英語名で参照する。</remarks>
        private static readonly string[] BackstageTrackedPropertyKeys = new[]
        {
            "Title",
            "Subject",
            "Author",
            "Keywords",
            "Comments",
            "Category",
            "Manager",
            "Company",
            "Hyperlink Base",
            "Tags", // カスタム「Tags」の別名（下で CustomDocumentProperties から取得）
        };

        private static Dictionary<string, string> CaptureWorkbookProperties(Excel.Workbook wb)
        {
            var d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string key in BackstageTrackedPropertyKeys)
            {
                if (string.Equals(key, "Tags", StringComparison.OrdinalIgnoreCase))
                    d[key] = TryGetCustomDocumentPropertyText(wb, "Tags");
                else
                    d[key] = GetBuiltinDocumentPropertyText(wb, key);
            }
            return d;
        }

        private static List<string> FindChangedPropertyKeys(
            IReadOnlyDictionary<string, string> before,
            IReadOnlyDictionary<string, string> after)
        {
            var changed = new List<string>();
            foreach (string key in BackstageTrackedPropertyKeys)
            {
                before.TryGetValue(key, out string ov);
                after.TryGetValue(key, out string nv);
                if (!string.Equals(ov ?? "", nv ?? "", StringComparison.Ordinal))
                    changed.Add(key);
            }
            return changed;
        }

        private static string GetBuiltinDocumentPropertyText(Excel.Workbook wb, string propertyName)
        {
            try
            {
                var props = wb.BuiltinDocumentProperties as DocumentProperties;
                if (props == null) return "";
                DocumentProperty p = props[propertyName];
                return SafeObjectToText(p?.Value);
            }
            catch
            {
                return "";
            }
        }

        private static string TryGetCustomDocumentPropertyText(Excel.Workbook wb, string propertyName)
        {
            try
            {
                var custom = wb.CustomDocumentProperties as DocumentProperties;
                if (custom == null) return "";
                for (int i = 1; i <= custom.Count; i++)
                {
                    try
                    {
                        DocumentProperty p = custom[i];
                        if (p != null && string.Equals(p.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                            return SafeObjectToText(p.Value);
                    }
                    catch { }
                }
            }
            catch { }
            return "";
        }

        private static string SafeObjectToText(object value)
        {
            try
            {
                if (value == null) return "";
                if (value is Array arr)
                {
                    var items = new System.Collections.Generic.List<string>();
                    foreach (object item in arr)
                        items.Add(item?.ToString() ?? "");
                    return "[" + string.Join(",", items) + "]";
                }
                return Convert.ToString(value) ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static string GetWorkbookKey(Excel.Workbook wb)
        {
            try
            {
                string p = wb.FullName;
                if (!string.IsNullOrEmpty(p)) return p;
            }
            catch { }
            try { return wb.Name ?? "?"; }
            catch { return "?"; }
        }

        private static string SafeWorkbookName(Excel.Workbook wb)
        {
            try { return wb?.Name ?? "(unknown)"; }
            catch { return "(unknown)"; }
        }

        public void ShowLogOnAction(IRibbonControl control)
        {
            try
            {
                if (_logPane != null)
                {
                    _logPane.Visible = true;
                    var ctrl = _logPane.Control as LogPaneUserControl;
                    if (ctrl != null) ctrl.RefreshLog();
                    return;
                }

                var logControl = new LogPaneUserControl();
                _logPane = Globals.ThisAddIn.CustomTaskPanes.Add(logControl, "ログ");
                _logPane.Visible = true;
                logControl.RefreshLog();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelAddIn1.Ribbon] ShowLogOnAction: " + ex.Message);
                MessageBox.Show("ログの表示に失敗しました: " + ex.Message, "ExcelAddIn1", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void ClearLogOnAction(IRibbonControl control)
        {
            Logger.ClearLog();
            MessageBox.Show("ログをクリアしました。", "ExcelAddIn1", MessageBoxButtons.OK, MessageBoxIcon.Information);
            RefreshLogPaneIfVisible();
        }

        public void CopyLogPathOnAction(IRibbonControl control)
        {
            try
            {
                string path = Logger.GetLogFilePath();
                Clipboard.SetText(path);
                MessageBox.Show("ログパスをクリップボードにコピーしました。\n" + path, "ExcelAddIn1", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelAddIn1.Ribbon] CopyLogPathOnAction: " + ex.Message);
                MessageBox.Show("コピーに失敗しました: " + ex.Message, "ExcelAddIn1", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static string GetResourceText(string resourceName)
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var stream = assembly.GetManifestResourceStream(resourceName);
                if (stream != null)
                {
                    using (var reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }

                string filePath = Path.Combine(Path.GetDirectoryName(assembly.Location) ?? string.Empty, "Ribbon.xml");
                if (File.Exists(filePath))
                {
                    return File.ReadAllText(filePath);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExcelAddIn1.Ribbon] GetResourceText: " + ex.Message);
            }

            return GetFallbackRibbonXml();
        }

        private void RefreshLogPaneIfVisible()
        {
            if (_logPane == null || !_logPane.Visible) return;
            var ctrl = _logPane.Control as LogPaneUserControl;
            if (ctrl != null) ctrl.RefreshLog();
        }

        private static string GetFallbackRibbonXml()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""Ribbon_Load"">
  <ribbon>
    <tabs>
      <tab id=""MosLogTab"" label=""ログ"" insertAfterMso=""TabHelp"">
        <group id=""MosLogGroup"" label=""ログ"">
          <button id=""ShowLog"" label=""ログを表示"" onAction=""ShowLogOnAction"" />
          <button id=""ClearLog"" label=""ログをクリア"" onAction=""ClearLogOnAction"" />
          <button id=""CopyLogPath"" label=""ログパスをコピー"" onAction=""CopyLogPathOnAction"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
  <backstage onShow=""OnBackstageShow"" onHide=""OnBackstageHide"" />
</customUI>";
        }
    }
}
