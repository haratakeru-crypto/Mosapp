using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;

namespace PowerPointAddIn1
{
    /// <summary>
    /// デバックタブのリボン拡張。ログ表示・クリア・パスコピーを提供する。
    /// </summary>
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        private IRibbonUI _ribbon;
        private Microsoft.Office.Tools.CustomTaskPane _logPane;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            System.Diagnostics.Debug.WriteLine("[Ribbon] GetCustomUI called, ribbonID=" + (ribbonID ?? "(null)"));
            string xml = GetResourceText("PowerPointAddIn1.Ribbon.xml");
            System.Diagnostics.Debug.WriteLine("[Ribbon] GetCustomUI returning " + (xml?.Length ?? 0) + " chars");
            return xml;
        }

        #endregion

        #region リボンコールバック

        public void Ribbon_Load(IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
            System.Diagnostics.Debug.WriteLine("[Ribbon] Ribbon_Load completed - ログタブ有効");
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
                System.Diagnostics.Debug.WriteLine("[Ribbon] ShowLogOnAction: " + ex.Message);
                MessageBox.Show("ログの表示に失敗しました: " + ex.Message, "PowerPointAddIn1", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void ClearLogOnAction(IRibbonControl control)
        {
            try
            {
                Logger.ClearLog();
                if (_logPane != null && _logPane.Visible)
                {
                    var ctrl = _logPane.Control as LogPaneUserControl;
                    if (ctrl != null) ctrl.RefreshLog();
                }
                MessageBox.Show("ログをクリアしました。", "PowerPointAddIn1", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[Ribbon] ClearLogOnAction: " + ex.Message);
                MessageBox.Show("ログのクリアに失敗しました: " + ex.Message, "PowerPointAddIn1", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public void CopyLogPathOnAction(IRibbonControl control)
        {
            try
            {
                string path = Logger.GetLogFilePath();
                Clipboard.SetText(path);
                MessageBox.Show("ログパスをクリップボードにコピーしました。\n" + path, "PowerPointAddIn1", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[Ribbon] CopyLogPathOnAction: " + ex.Message);
                MessageBox.Show("コピーに失敗しました: " + ex.Message, "PowerPointAddIn1", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// 組み込みコマンド実行時に呼ばれる。操作をログに記録し、既定の動作はそのまま実行させる。
        /// </summary>
        // 通常のボタン用
        public void Ribbon_OnCommand(IRibbonControl control)
        {
            RecordCommand("RibbonCommand", control, null);
        }
        private void RecordCommand(string tag, IRibbonControl control, bool? pressed)
        {
            try
            {
                string id = control?.Id ?? "(null)";
                string detail = "Id=" + id + (pressed.HasValue ? " Pressed=" + pressed.Value : "");
                
                
                int? p = ThisAddIn.CurrentTaskProjectId >= 0 ? (int?)ThisAddIn.CurrentTaskProjectId : null;
                int? t = ThisAddIn.CurrentTaskTaskId >= 0 ? (int?)ThisAddIn.CurrentTaskTaskId : null;
                
                // ログファイルにも記録
                Logger.LogOperation(tag, detail, p, t);
            }
            catch
            {
            }
        }

        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var stream = assembly.GetManifestResourceStream(resourceName);

                if (stream == null)
                {
                    string filePath = Path.Combine(
                        Path.GetDirectoryName(assembly.Location) ?? "",
                        "Ribbon.xml");
                    if (File.Exists(filePath))
                    {
                        System.Diagnostics.Debug.WriteLine("[Ribbon] Ribbon XML from file: " + filePath);
                        return File.ReadAllText(filePath);
                    }
                    System.Diagnostics.Debug.WriteLine("[Ribbon] Ribbon XML from fallback (GetRibbonXmlContent)");
                    return GetRibbonXmlContent();
                }

                System.Diagnostics.Debug.WriteLine("[Ribbon] Ribbon XML from embedded resource: " + resourceName);
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[Ribbon] Error loading Ribbon XML: " + ex.Message);
                return GetRibbonXmlContent();
            }
        }

        private static string GetRibbonXmlContent()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""Ribbon_Load"">
  <commands>
    <command idMso=""Cut"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""Paste"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""SlideNew"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""SlideDuplicate"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""LayoutGallery"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""AlignLeft"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""AlignCenter"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""AlignRight"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""AlignTop"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""AlignMiddle"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""AlignBottom"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""BringForward"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""SendBackward"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""BringToFront"" onAction=""Ribbon_OnCommand"" />
    <command idMso=""SendToBack"" onAction=""Ribbon_OnCommand"" />
  </commands>
  <ribbon>
    <tabs>
      <tab id=""DebugTab"" label=""ログ"" insertAfterMso=""Help"">
        <group id=""DebugGroup"" label=""ログ"">
          <button id=""ShowLog"" label=""ログを表示"" onAction=""ShowLogOnAction"" />
          <button id=""ClearLog"" label=""ログをクリア"" onAction=""ClearLogOnAction"" />
          <button id=""CopyLogPath"" label=""ログパスをコピー"" onAction=""CopyLogPathOnAction"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        #endregion
    }
}
