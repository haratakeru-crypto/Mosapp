using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
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

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelAddIn1.Ribbon.xml");
        }

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            System.Diagnostics.Debug.WriteLine("[ExcelAddIn1.Ribbon] Ribbon_Load completed");
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
</customUI>";
        }
    }
}
