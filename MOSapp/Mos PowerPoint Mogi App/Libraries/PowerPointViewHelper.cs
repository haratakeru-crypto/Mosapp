using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Libraries.Group1;

namespace Libraries
{
    /// <summary>
    /// PowerPoint ウィンドウ表示の共通操作。
    /// </summary>
    public static class PowerPointViewHelper
    {
        private const string ShowNotesMso = "ShowNotes";

        /// <summary>
        /// 通常表示のまま左のスライドサムネイルは表示し、下部のノートペインのみ非表示にする。
        /// </summary>
        public static void HideNotesPane(Application app)
        {
            if (app == null) return;

            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                DocumentWindow window = null;
                try
                {
                    window = app.ActiveWindow;
                    if (window == null) return;

                    // 左サムネイル付きの通常表示を確実にする（ppViewSlide 単体だとサムネイルが消える）
                    window.ViewType = PpViewType.ppViewSlide;
                    window.ViewType = PpViewType.ppViewNormal;

                    // スライド領域を最大化して下部ノートペインを実質非表示にする
                    TryCollapseNotesBySplit(window);

                    // ステータスバーの「ノート」がオンならオフにする
                    TryHideNotesToggle(app);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[PowerPointViewHelper] HideNotesPane: " + ex.Message);
                }
                finally
                {
                    if (window != null)
                    {
                        try { Marshal.ReleaseComObject(window); } catch { }
                    }
                }
            }
        }

        private static void TryCollapseNotesBySplit(DocumentWindow window)
        {
            if (window == null) return;
            try
            {
                window.SplitVertical = 100;
            }
            catch
            {
                try { window.SplitVertical = 99; } catch { }
            }
        }

        private static void TryHideNotesToggle(Application app)
        {
            if (app == null) return;
            try
            {
                CommandBars commandBars = app.CommandBars;
                if (commandBars == null) return;
                try
                {
                    if (commandBars.GetPressedMso(ShowNotesMso))
                        commandBars.ExecuteMso(ShowNotesMso);
                }
                finally
                {
                    try { Marshal.ReleaseComObject(commandBars); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[PowerPointViewHelper] TryHideNotesToggle: " + ex.Message);
            }
        }
    }
}
