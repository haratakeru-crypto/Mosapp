using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public partial class ThisAddIn
    {
        private static readonly string CurrentTaskFilePath = Path.Combine(Path.GetTempPath(), "mos_ppt_current_task.txt");

        private Timer _grayscalePollTimer;
        private bool _lastBlackAndWhite;
        private Timer _audio8_4PollTimer;
        private bool _task8_4Logged;
        private Timer _layout10_7PollTimer;
        private bool _task10_7Logged;
        private Timer _printOptionsPollTimer;
        private string _lastPrintPresFullName;
        private int _lastPrintOutputType = -1;
        private int _lastPrintCopies = -1;
        private int _lastPrintCollate = -1;
        private bool _printOptionsInitialized;
        private bool _task5_1PrintLogged;
        private bool _task11_7PrintLogged;

        private Timer _taskFilePollTimer;
        private int _currentTaskProjectId = -1;
        private int _currentTaskTaskId = -1;

        private Timer _shapePositionPollTimer;
        private Dictionary<string, Tuple<float, float, float, float>> _shapePositionSnapshot = new Dictionary<string, Tuple<float, float, float, float>>();
        private const float PositionTolerancePt = 0.5f;

        /// <summary>現在タスク（VSTO が読み取った ProjectId, TaskId）。ログ記録時に使用。</summary>
        internal static int CurrentTaskProjectId { get; private set; } = -1;
        /// <summary>現在タスクの TaskId。</summary>
        internal static int CurrentTaskTaskId { get; private set; } = -1;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[PowerPointAddIn1] Add-in started. Log file: " + Logger.GetLogFilePath());

            _lastBlackAndWhite = false;
            _grayscalePollTimer = new Timer();
            _grayscalePollTimer.Interval = 500;
            _grayscalePollTimer.Tick += GrayscalePollTimer_Tick;
            _grayscalePollTimer.Start();

            _task5_1PrintLogged = false;
            _task11_7PrintLogged = false;
            _printOptionsPollTimer = new Timer();
            _printOptionsPollTimer.Interval = 2000;
            _printOptionsPollTimer.Tick += PrintOptionsPollTimer_Tick;
            _printOptionsPollTimer.Start();

            _task8_4Logged = false;
            _audio8_4PollTimer = new Timer();
            _audio8_4PollTimer.Interval = 1000;
            _audio8_4PollTimer.Tick += Audio8_4PollTimer_Tick;
            _audio8_4PollTimer.Start();

            _task10_7Logged = false;
            _layout10_7PollTimer = new Timer();
            _layout10_7PollTimer.Interval = 1500;
            _layout10_7PollTimer.Tick += Layout10_7PollTimer_Tick;
            _layout10_7PollTimer.Start();

            _taskFilePollTimer = new Timer();
            _taskFilePollTimer.Interval = 500;
            _taskFilePollTimer.Tick += TaskFilePollTimer_Tick;
            _taskFilePollTimer.Start();

            _shapePositionPollTimer = new Timer();
            _shapePositionPollTimer.Interval = 1500;
            _shapePositionPollTimer.Tick += ShapePositionPollTimer_Tick;
            _shapePositionPollTimer.Start();
        }

        private void TaskFilePollTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (!File.Exists(CurrentTaskFilePath))
                    return;
                string line = null;
                try
                {
                    line = File.ReadAllText(CurrentTaskFilePath).Trim();
                    if (string.IsNullOrEmpty(line)) return;
                }
                catch { return; }
                var parts = line.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 2) return;
                if (!int.TryParse(parts[0].Trim(), out int projectId) || !int.TryParse(parts[1].Trim(), out int taskId))
                    return;
                if (projectId == _currentTaskProjectId && taskId == _currentTaskTaskId)
                    return;
                _currentTaskProjectId = projectId;
                _currentTaskTaskId = taskId;
                CurrentTaskProjectId = projectId;
                CurrentTaskTaskId = taskId;
                Logger.LogTaskStart(projectId, taskId);
                TakeShapePositionSnapshot();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[TaskFilePoll] " + ex.Message);
            }
        }

        private void TakeShapePositionSnapshot()
        {
            _shapePositionSnapshot.Clear();
            try
            {
                if (Application == null || Application.Presentations == null) return;
                PowerPoint.Presentation pres = null;
                try
                {
                    pres = Application.ActivePresentation;
                    if (pres == null) return;
                    PowerPoint.Slides slides = null;
                    try
                    {
                        slides = pres.Slides;
                        if (slides == null) return;
                        for (int si = 1; si <= slides.Count; si++)
                        {
                            PowerPoint.Slide slide = null;
                            try
                            {
                                slide = slides[si];
                                if (slide == null) continue;
                                PowerPoint.Shapes shapes = null;
                                try
                                {
                                    shapes = slide.Shapes;
                                    if (shapes == null) continue;
                                    for (int shi = 1; shi <= shapes.Count; shi++)
                                    {
                                        PowerPoint.Shape sh = null;
                                        try
                                        {
                                            sh = shapes[shi];
                                            if (sh == null) continue;
                                            if (!IsImageOrPlaceholder(sh)) continue;
                                            float left = (float)sh.Left;
                                            float top = (float)sh.Top;
                                            float w = (float)sh.Width;
                                            float h = (float)sh.Height;
                                            string key = si + "_" + sh.Id;
                                            _shapePositionSnapshot[key] = Tuple.Create(left, top, w, h);
                                        }
                                        finally { if (sh != null) try { Marshal.ReleaseComObject(sh); } catch { } }
                                    }
                                }
                                finally { if (shapes != null) try { Marshal.ReleaseComObject(shapes); } catch { } }
                            }
                            finally { if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { } }
                        }
                    }
                    finally { if (slides != null) try { Marshal.ReleaseComObject(slides); } catch { } }
                }
                finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ShapeSnapshot] " + ex.Message);
            }
        }

        private static bool IsImageOrPlaceholder(PowerPoint.Shape sh)
        {
            try
            {
                int t = (int)sh.Type;
                if (t == (int)Office.MsoShapeType.msoPicture) return true;
                if (t == (int)Office.MsoShapeType.msoPlaceholder) return true;
                try
                {
                    var pf = sh.PlaceholderFormat;
                    if (pf != null) { Marshal.ReleaseComObject(pf); return true; }
                }
                catch { }
                return false;
            }
            catch { return false; }
        }

        private void ShapePositionPollTimer_Tick(object sender, EventArgs e)
        {
            if (_currentTaskProjectId < 0 || _currentTaskTaskId < 0) return;
            try
            {
                if (Application == null || Application.Presentations == null) return;
                PowerPoint.Presentation pres = null;
                try
                {
                    pres = Application.ActivePresentation;
                    if (pres == null) return;
                    PowerPoint.Slides slides = null;
                    try
                    {
                        slides = pres.Slides;
                        if (slides == null) return;
                        for (int si = 1; si <= slides.Count; si++)
                        {
                            PowerPoint.Slide slide = null;
                            try
                            {
                                slide = slides[si];
                                if (slide == null) continue;
                                PowerPoint.Shapes shapes = null;
                                try
                                {
                                    shapes = slide.Shapes;
                                    if (shapes == null) continue;
                                    for (int shi = 1; shi <= shapes.Count; shi++)
                                    {
                                        PowerPoint.Shape sh = null;
                                        try
                                        {
                                            sh = shapes[shi];
                                            if (sh == null) continue;
                                            if (!IsImageOrPlaceholder(sh)) continue;
                                            string key = si + "_" + sh.Id;
                                            float left = (float)sh.Left;
                                            float top = (float)sh.Top;
                                            float w = (float)sh.Width;
                                            float h = (float)sh.Height;
                                            if (_shapePositionSnapshot.TryGetValue(key, out var old))
                                            {
                                                if (Math.Abs(old.Item1 - left) > PositionTolerancePt || Math.Abs(old.Item2 - top) > PositionTolerancePt ||
                                                    Math.Abs(old.Item3 - w) > PositionTolerancePt || Math.Abs(old.Item4 - h) > PositionTolerancePt)
                                                {
                                                    string detail = $"Slide={si} ShapeId={sh.Id} Left={old.Item1:F1}->{left:F1} Top={old.Item2:F1}->{top:F1} Width={old.Item3:F1} Height={old.Item4:F1}";
                                                    Logger.LogOperation("ShapePositionChange", detail, _currentTaskProjectId, _currentTaskTaskId);
                                                    _shapePositionSnapshot[key] = Tuple.Create(left, top, w, h);
                                                }
                                            }
                                            else
                                            {
                                                _shapePositionSnapshot[key] = Tuple.Create(left, top, w, h);
                                            }
                                        }
                                        finally { if (sh != null) try { Marshal.ReleaseComObject(sh); } catch { } }
                                    }
                                }
                                finally { if (shapes != null) try { Marshal.ReleaseComObject(shapes); } catch { } }
                            }
                            finally { if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { } }
                        }
                    }
                    finally { if (slides != null) try { Marshal.ReleaseComObject(slides); } catch { } }
                }
                finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ShapePositionPoll] " + ex.Message);
            }
        }

        private void Layout10_7PollTimer_Tick(object sender, EventArgs e)
        {
            if (_task10_7Logged) return;
            try
            {
                if (Application == null || Application.Presentations == null) return;
                PowerPoint.Presentation pres = null;
                try
                {
                    pres = Application.ActivePresentation;
                    if (pres == null) return;
                    PowerPoint.Master master = null;
                    try
                    {
                        master = pres.SlideMaster;
                        if (master == null) return;
                        PowerPoint.CustomLayouts layouts = null;
                        try
                        {
                            layouts = master.CustomLayouts;
                            if (layouts == null) return;
                            for (int i = 1; i <= layouts.Count; i++)
                            {
                                PowerPoint.CustomLayout cl = null;
                                try
                                {
                                    cl = layouts[i];
                                    if (cl == null) continue;
                                    string name = null;
                                    try { name = cl.Name ?? ""; } catch { continue; }
                                    if (name.IndexOf("画像付きスライド", StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        Logger.LogTask10_7LayoutDuplicate();
                                        _task10_7Logged = true;
                                        return;
                                    }
                                }
                                finally { if (cl != null) try { Marshal.ReleaseComObject(cl); } catch { } }
                            }
                        }
                        finally { if (layouts != null) try { Marshal.ReleaseComObject(layouts); } catch { } }
                    }
                    finally { if (master != null) try { Marshal.ReleaseComObject(master); } catch { } }
                }
                finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
            }
            catch { }
        }

        private void PrintOptionsPollTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (Application == null || Application.Presentations == null) return;
                PowerPoint.Presentation pres = null;
                try
                {
                    pres = Application.ActivePresentation;
                    if (pres == null) return;
                    string fullName = null;
                    try { fullName = pres.FullName ?? ""; } catch { return; }
                    if (string.IsNullOrEmpty(fullName)) fullName = pres.Name ?? "";

                    if (_lastPrintPresFullName != null && fullName != _lastPrintPresFullName)
                    {
                        _lastPrintPresFullName = null;
                        _printOptionsInitialized = false;
                        _lastPrintOutputType = -1;
                        _lastPrintCopies = -1;
                        _lastPrintCollate = -1;
                    }

                    PowerPoint.PrintOptions po = null;
                    try
                    {
                        po = pres.PrintOptions;
                        if (po == null) return;
                        int outputType = (int)po.OutputType;
                        int copies = po.NumberOfCopies;
                        int collateInt = Convert.ToInt32(po.Collate);
                        bool collate = (collateInt == (int)Office.MsoTriState.msoTrue);

                        if (!_printOptionsInitialized)
                        {
                            _lastPrintPresFullName = fullName;
                            _lastPrintOutputType = outputType;
                            _lastPrintCopies = copies;
                            _lastPrintCollate = collateInt;
                            _printOptionsInitialized = true;
                            return;
                        }

                        bool changed = (_lastPrintOutputType != outputType || _lastPrintCopies != copies || _lastPrintCollate != collateInt);
                        _lastPrintOutputType = outputType;
                        _lastPrintCopies = copies;
                        _lastPrintCollate = collateInt;

                        if (changed)
                        {
                            if (!_task5_1PrintLogged &&
                                outputType == (int)PowerPoint.PpPrintOutputType.ppPrintOutputThreeSlideHandouts &&
                                copies == 4 && collate)
                            {
                                Logger.LogTask5_1Print();
                                _task5_1PrintLogged = true;
                            }
                            if (!_task11_7PrintLogged &&
                                outputType == (int)PowerPoint.PpPrintOutputType.ppPrintOutputNotesPages &&
                                copies == 3 && collate)
                            {
                                Logger.LogTask11_7Print();
                                _task11_7PrintLogged = true;
                            }
                        }
                    }
                    finally { if (po != null) try { Marshal.ReleaseComObject(po); } catch { } }
                }
                finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
            }
            catch { }
        }

        private void Audio8_4PollTimer_Tick(object sender, EventArgs e)
        {
            if (_task8_4Logged) return;
            try
            {
                if (Application == null || Application.Presentations == null) return;
                PowerPoint.Presentation pres = null;
                try
                {
                    pres = Application.ActivePresentation;
                    if (pres == null) return;
                    PowerPoint.Slide slide = null;
                    try
                    {
                        PowerPoint.Slides slides = pres.Slides;
                        if (slides == null || slides.Count < 1) return;
                        slide = slides[1];
                        if (slide == null) return;
                        PowerPoint.Shapes shapes = slide.Shapes;
                        if (shapes == null) return;
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            PowerPoint.Shape sh = null;
                            try
                            {
                                sh = shapes[i];
                                try
                                {
                                    if (sh.MediaType != PowerPoint.PpMediaType.ppMediaTypeSound) continue;
                                }
                                catch { continue; }
                                PowerPoint.MediaFormat mf = null;
                                try
                                {
                                    mf = sh.MediaFormat;
                                    if (mf == null) continue;
                                    float fadeIn = (float)mf.FadeInDuration;
                                    if (Math.Abs(fadeIn - 4000f) < 500f)
                                    {
                                        Logger.LogTask8_4Audio();
                                        _task8_4Logged = true;
                                        return;
                                    }
                                }
                                finally { if (mf != null) try { Marshal.ReleaseComObject(mf); } catch { } }
                            }
                            finally { if (sh != null) try { Marshal.ReleaseComObject(sh); } catch { } }
                        }
                    }
                    finally { if (slide != null) try { Marshal.ReleaseComObject(slide); } catch { } }
                }
                finally { if (pres != null) try { Marshal.ReleaseComObject(pres); } catch { } }
            }
            catch { }
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
            if (_shapePositionPollTimer != null)
            {
                _shapePositionPollTimer.Stop();
                _shapePositionPollTimer.Dispose();
                _shapePositionPollTimer = null;
            }
            if (_taskFilePollTimer != null)
            {
                _taskFilePollTimer.Stop();
                _taskFilePollTimer.Dispose();
                _taskFilePollTimer = null;
            }
            if (_printOptionsPollTimer != null)
            {
                _printOptionsPollTimer.Stop();
                _printOptionsPollTimer.Dispose();
                _printOptionsPollTimer = null;
            }
            if (_audio8_4PollTimer != null)
            {
                _audio8_4PollTimer.Stop();
                _audio8_4PollTimer.Dispose();
                _audio8_4PollTimer = null;
            }
            if (_layout10_7PollTimer != null)
            {
                _layout10_7PollTimer.Stop();
                _layout10_7PollTimer.Dispose();
                _layout10_7PollTimer = null;
            }
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
