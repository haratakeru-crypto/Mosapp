using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using Libraries.Group1;
using PowerPointApp = Microsoft.Office.Interop.PowerPoint.Application;
using PowerPointPresentation = Microsoft.Office.Interop.PowerPoint.Presentation;

namespace Libraries
{
    /// <summary>
    /// PowerPoint 取得とアプリ終了時のクリーンアップ。
    /// </summary>
    public static class PowerPointApplicationManager
    {
        private const int PowerPointQuitWaitMs = 5000;

        /// <summary>
        /// 模試アプリ終了時: プレゼンを保存して閉じ、PowerPoint を終了し、残プロセスがあれば待機／強制終了する。
        /// </summary>
        public static void ClosePowerPointApplicationForAppExit()
        {
            lock (PowerPointCheckerCommon.PowerPointComInteropSync)
            {
                PowerPointApp pptApp = null;
                try
                {
                    try
                    {
                        pptApp = (PowerPointApp)Marshal.GetActiveObject("PowerPoint.Application");
                    }
                    catch
                    {
                        return;
                    }

                    if (pptApp == null)
                        return;

                    try { pptApp.DisplayAlerts = Microsoft.Office.Interop.PowerPoint.PpAlertLevel.ppAlertsNone; }
                    catch { /* ignore */ }

                    // プレゼンを保存して閉じる
                    const int maxAttempts = 25;
                    int prevCount = GetPresentationCountSafe(pptApp);
                    for (int attempt = 0; attempt < maxAttempts && GetPresentationCountSafe(pptApp) > 0; attempt++)
                    {
                        PowerPointPresentation openPres = null;
                        try
                        {
                            openPres = pptApp.Presentations[1];
                            try { openPres.Save(); } catch { /* ignore */ }
                            try { openPres.Close(); } catch { /* ignore */ }

                            int newCount = GetPresentationCountSafe(pptApp);
                            if (newCount >= prevCount)
                                break;
                            prevCount = newCount;
                        }
                        catch (COMException)
                        {
                            break;
                        }
                        catch
                        {
                            break;
                        }
                        finally
                        {
                            if (openPres != null)
                            {
                                try { Marshal.ReleaseComObject(openPres); } catch { /* ignore */ }
                            }
                        }
                    }

                    try { pptApp.Quit(); } catch { /* ignore */ }
                    try { Marshal.ReleaseComObject(pptApp); } catch { /* ignore */ }
                    pptApp = null;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        "[ClosePowerPointApplicationForAppExit] " + ex.Message);
                }
                finally
                {
                    if (pptApp != null)
                    {
                        try { Marshal.ReleaseComObject(pptApp); } catch { /* ignore */ }
                    }
                }
            }

            WaitForPowerPointExitOrKill();
        }

        private static int GetPresentationCountSafe(PowerPointApp pptApp)
        {
            try
            {
                return pptApp?.Presentations?.Count ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        private static void WaitForPowerPointExitOrKill()
        {
            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < PowerPointQuitWaitMs)
            {
                if (Process.GetProcessesByName("POWERPNT").Length == 0)
                    return;
                Thread.Sleep(200);
            }

            foreach (Process p in Process.GetProcessesByName("POWERPNT"))
            {
                try
                {
                    if (!p.HasExited)
                        p.Kill();
                }
                catch { /* ignore */ }
                finally
                {
                    try { p.Dispose(); } catch { /* ignore */ }
                }
            }
        }
    }
}
