using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;

namespace New_MOSWordVSTOAddIn
{
    /// <summary>
    /// Ribbonコマンドのイベントハンドラ。Office からコールバックを受けるため COM 公開する。
    /// </summary>
    [ComVisible(true)]
    public partial class Ribbon : Microsoft.Office.Core.IRibbonExtensibility
    {
        private Microsoft.Office.Core.IRibbonUI ribbon;

        public Ribbon()
        {
            System.Diagnostics.Debug.WriteLine("[Ribbon] Constructor called");
        }

        #region IRibbonExtensibility メンバー

        public string GetCustomUI(string ribbonID)
        {
            System.Diagnostics.Debug.WriteLine($"[Ribbon] GetCustomUI called, ribbonID={ribbonID}");
            string xml = GetResourceText("New_MOSWordVSTOAddIn.Ribbon.xml");
            System.Diagnostics.Debug.WriteLine($"[Ribbon] GetCustomUI returning {xml?.Length ?? 0} chars");
            return xml;
        }

        #endregion

        #region リボンのコールバック

        public void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            System.Diagnostics.Debug.WriteLine("[Ribbon] Ribbon_Load completed - MOSデバッグタブとコマンドフックが有効です");
        }

        /// <summary>
        /// 通常のコマンドのonActionイベントハンドラ。
        /// &lt;commands&gt; でフックした Cut / Paste / UpgradeDocument / FileSaveAs などは、この署名（control, ref cancelDefault）が必須。
        /// 署名が一致しないと「CommandOnAction の署名が一致しません」エラーになる。
        /// </summary>
        public void CommandOnAction(Microsoft.Office.Core.IRibbonControl control, ref bool cancelDefault)
        {
            try
            {
                string commandId = control.Id;
                // 6-1: チェッカーが期待する ID に統一（ConvertTextToTable → TableConvertTextToTable）
                if (string.Equals(commandId, "ConvertTextToTable", StringComparison.OrdinalIgnoreCase))
                {
                    Logger.LogCommand("TableConvertTextToTable");
                }
                // 3-2: SectionBreakInsert を InsertSectionBreakNextPage としてログ（チェッカーが参照する ID）
                else if (string.Equals(commandId, "SectionBreakInsert", StringComparison.OrdinalIgnoreCase))
                {
                    Logger.LogCommand("InsertSectionBreakNextPage");
                }
                // 3-4: ColumnsLeft/Right は Word の idMso が無いため ColumnsDialog をフックし、チェッカー互換 ID で記録
                else if (string.Equals(commandId, "ColumnsDialog", StringComparison.OrdinalIgnoreCase))
                {
                    Logger.LogCommand("ColumnsLeft");
                }
                else if (string.Equals(commandId, "PageBorders", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(commandId, "PageBorderOptionsDialog", StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(commandId, "PageBorderAndShadingDialog", StringComparison.OrdinalIgnoreCase))
                {
                    // 4-6: ページ罫線系。Word の Ribbon.xml では PageBorderAndShadingDialog のみ有効（PageBorders は不明 ID）
                    Logger.LogCommand("PageBorders");
                    Globals.ThisAddIn?.RegisterRibbonLoggedPageBorders();
                }
                else
                {
                    if (string.Equals(commandId, "PageOrientationPortraitLandscape", StringComparison.OrdinalIgnoreCase))
                        Globals.ThisAddIn?.RegisterRibbonLoggedPageOrientation();
                    Logger.LogCommand(commandId);
                }

                // 既定の動作（Cut / Paste / SaveAs など）をキャンセルせずに実行させる
                cancelDefault = false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Ribbon] Error in CommandOnAction: {ex.Message}");
                // エラー時も既定動作はブロックしない
                cancelDefault = false;
            }
        }

        /// <summary>
        /// 目次ギャラリー（TableOfContentsGallery）のonActionイベントハンドラ
        /// 引数から選択されたアイテムIDを取得
        /// </summary>
        public void TableOfContentsGalleryOnAction(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            try
            {
                // 選択されたアイテムIDをログに記録
                // TocAutomatic2が選択された場合のみログに記録
                if (selectedId == "TocAutomatic2")
                {
                    Logger.LogCommand(selectedId);
                }
                else
                {
                    // 他の目次タイプが選択された場合も記録（デバッグ用）
                    Logger.LogCommand($"TableOfContentsGallery-{selectedId}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Ribbon] Error in TableOfContentsGalleryOnAction: {ex.Message}");
            }
        }

        /// <summary>
        /// 開発用タブ「MOSデバッグ」の表示制御。常に表示（Release でもログ確認可能にする）。
        /// </summary>
        public bool GetDebugTabVisible(Microsoft.Office.Core.IRibbonControl control)
        {
            return true;
        }

        /// <summary>
        /// デバッグタブの「ログパス表示」ボタン押下時。現在のログファイルパスをメッセージボックスで表示。
        /// </summary>
        public void ShowLogPathButtonOnAction(Microsoft.Office.Core.IRibbonControl control)
        {
            try
            {
                string path = Logger.GetLogFilePath();
                System.Windows.Forms.MessageBox.Show(path, "ログファイルパス", System.Windows.Forms.MessageBoxButtons.OK);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "エラー", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        /// <summary>
        /// デバッグタブの「ログ内容表示」ボタン押下時。ログファイルの内容をフォームで表示する。
        /// </summary>
        public void ShowLogContentButtonOnAction(Microsoft.Office.Core.IRibbonControl control)
        {
            try
            {
                string path = Logger.GetLogFilePath();
                if (!System.IO.File.Exists(path))
                {
                    System.Windows.Forms.MessageBox.Show("ログファイルがありません。", "ログ内容", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }
                string content = System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8);
                using (var form = new System.Windows.Forms.Form())
                {
                    form.Text = "ログ内容";
                    form.Size = new System.Drawing.Size(600, 400);
                    form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                    var textBox = new System.Windows.Forms.TextBox
                    {
                        Multiline = true,
                        ReadOnly = true,
                        ScrollBars = System.Windows.Forms.ScrollBars.Both,
                        Dock = System.Windows.Forms.DockStyle.Fill,
                        Font = new System.Drawing.Font(System.Drawing.FontFamily.GenericMonospace, 9f),
                        Text = content
                    };
                    form.Controls.Add(textBox);
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "エラー", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region ヘルパー

        private static string GetResourceText(string resourceName)
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                System.IO.Stream stream = assembly.GetManifestResourceStream(resourceName);

                // 論理名はビルド環境により変わることがあるため、Ribbon.xml を含むリソースを検索
                if (stream == null)
                {
                    string[] names = assembly.GetManifestResourceNames();
                    foreach (string name in names)
                    {
                        if (name.EndsWith("Ribbon.xml", StringComparison.OrdinalIgnoreCase))
                        {
                            stream = assembly.GetManifestResourceStream(name);
                            if (stream != null)
                            {
                                System.Diagnostics.Debug.WriteLine($"[Ribbon] Loaded XML from resource: {name}");
                                break;
                            }
                        }
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[Ribbon] Loaded XML from resource: {resourceName}");
                }

                if (stream != null)
                {
                    using (var reader = new System.IO.StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }

                // 埋め込みリソースが見つからない場合、DLL 同梱のファイルから読み込む
                string asmDir = System.IO.Path.GetDirectoryName(assembly.Location);
                if (!string.IsNullOrEmpty(asmDir))
                {
                    string filePath = System.IO.Path.Combine(asmDir, "Ribbon.xml");
                    if (System.IO.File.Exists(filePath))
                    {
                        System.Diagnostics.Debug.WriteLine($"[Ribbon] Loaded XML from file: {filePath}");
                        return System.IO.File.ReadAllText(filePath);
                    }
                }

                // フォールバック: ハードコード XML（タブ・コマンドフック同一）
                System.Diagnostics.Debug.WriteLine("[Ribbon] Using fallback GetRibbonXmlContent()");
                return GetRibbonXmlContent();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Ribbon] Error loading Ribbon XML: {ex.Message}");
                return GetRibbonXmlContent();
            }
        }

        private static string GetRibbonXmlContent()
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""Ribbon_Load"">
  <commands>
    <command idMso=""Cut"" onAction=""CommandOnAction"" />
    <command idMso=""Copy"" onAction=""CommandOnAction"" />
    <command idMso=""Paste"" onAction=""CommandOnAction"" />
    <command idMso=""FontColorMoreColorsDialog"" onAction=""CommandOnAction"" />
    <command idMso=""PageOrientationPortraitLandscape"" onAction=""CommandOnAction"" />
    <command idMso=""ColumnsDialog"" onAction=""CommandOnAction"" />
    <command idMso=""SectionBreakInsert"" onAction=""CommandOnAction"" />
    <command idMso=""PageBorderAndShadingDialog"" onAction=""CommandOnAction"" />
    <command idMso=""ConvertTextToTable"" onAction=""CommandOnAction"" />
  </commands>
  <ribbon>
    <tabs>
      <tab id=""tabMOSDebug"" label=""MOSデバッグ"" insertAfterMso=""Help"">
        <group id=""grpLog"" label=""ログ"">
          <button id=""btnShowLogPath"" label=""ログパス表示"" onAction=""ShowLogPathButtonOnAction"" />
          <button id=""btnShowLogContent"" label=""ログ内容表示"" onAction=""ShowLogContentButtonOnAction"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        #endregion
    }
}




