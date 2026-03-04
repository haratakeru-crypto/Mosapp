using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        private Timer _grayscalePollTimer;
        private bool _lastBlackAndWhite;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[PowerPointAddIn1] Add-in started. Log file: " + Logger.GetLogFilePath());

            _lastBlackAndWhite = false;
            _grayscalePollTimer = new Timer();
            _grayscalePollTimer.Interval = 500;
            _grayscalePollTimer.Tick += GrayscalePollTimer_Tick;
            _grayscalePollTimer.Start();
        }

        private void GrayscalePollTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (Application == null) return;
                dynamic window = Application.ActiveWindow;
                if (window == null) return;

                bool current = false;
                try
                {
                    // COM は MsoTriState を整数で返すため、enum 比較ではなく数値で判定する
                    current = (Convert.ToInt32(window.BlackAndWhite) == (int)Office.MsoTriState.msoTrue);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[GrayscalePoll] BlackAndWhite get failed: " + ex.Message);
                    return;
                }

                if (current && !_lastBlackAndWhite)
                {
                    Logger.LogTask10_4Grayscale();
                }
                _lastBlackAndWhite = current;
            }
            catch
            {
                // アドインが落ちないように握りつぶす
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (_grayscalePollTimer != null)
            {
                _grayscalePollTimer.Stop();
                _grayscalePollTimer.Dispose();
                _grayscalePollTimer = null;
            }
            System.Diagnostics.Debug.WriteLine("[PowerPointAddIn1] Add-in shutdown");
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            System.Diagnostics.Debug.WriteLine("[ThisAddIn] CreateRibbonExtensibilityObject called");
            return new Ribbon();
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
