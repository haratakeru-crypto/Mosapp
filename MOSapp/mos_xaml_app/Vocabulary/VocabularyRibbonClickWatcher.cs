using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Threading;

namespace MOSExcelMogiApp.Vocabulary
{
    /// <summary>
    /// 画面全体の左クリック位置を受け取る（WH_MOUSE_LL）。
    /// イベントを出さないリボンのボタン（ギャラリー等）の押下判定に使う。
    /// </summary>
    public sealed class VocabularyRibbonClickWatcher : IDisposable
    {
        delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        struct POINT { public int X; public int Y; }

        [StructLayout(LayoutKind.Sequential)]
        struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool GetPhysicalCursorPos(out POINT lpPoint);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        static extern IntPtr GetModuleHandle(string lpModuleName);

        const int WH_MOUSE_LL = 14;
        const int WM_LBUTTONDOWN = 0x0201;

        readonly Dispatcher _dispatcher;
        readonly LowLevelMouseProc _proc;
        IntPtr _hook = IntPtr.Zero;

        /// <summary>物理スクリーン座標の左クリック。UI スレッドで呼ばれる。</summary>
        public event Action<Point> LeftClick;

        public VocabularyRibbonClickWatcher(Dispatcher dispatcher)
        {
            _dispatcher = dispatcher;
            _proc = HookProc;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public POINT pt;
        }

        [DllImport("user32.dll")]
        static extern int GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

        [DllImport("user32.dll")]
        static extern bool PostThreadMessage(uint threadId, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        static extern uint GetCurrentThreadId();

        const uint WM_QUIT = 0x0012;

        Thread _thread;
        uint _threadId;

        /// <summary>
        /// フックは専用スレッドで受ける。UI スレッドが UIA 等で止まっても、全体のマウスが重くならない。
        /// </summary>
        public void Start()
        {
            if (_thread != null) return;
            var ready = new ManualResetEventSlim(false);
            _thread = new Thread(() =>
            {
                _threadId = GetCurrentThreadId();
                try
                {
                    using (var module = Process.GetCurrentProcess().MainModule)
                        _hook = SetWindowsHookEx(WH_MOUSE_LL, _proc, GetModuleHandle(module.ModuleName), 0);
                }
                catch
                {
                    _hook = IntPtr.Zero;
                }
                ready.Set();
                if (_hook == IntPtr.Zero) return;

                while (GetMessage(out MSG msg, IntPtr.Zero, 0, 0) > 0) { }

                try { UnhookWindowsHookEx(_hook); } catch { }
                _hook = IntPtr.Zero;
            })
            {
                IsBackground = true,
                Name = "VocabularyRibbonClickWatcher"
            };
            _thread.Start();
            ready.Wait(1000);
        }

        public void Stop()
        {
            var thread = _thread;
            if (thread == null) return;
            _thread = null;
            try
            {
                if (_threadId != 0)
                    PostThreadMessage(_threadId, WM_QUIT, IntPtr.Zero, IntPtr.Zero);
            }
            catch { }
            _threadId = 0;
        }

        IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_LBUTTONDOWN)
            {
                try
                {
                    var info = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                    var pt = new Point(info.pt.X, info.pt.Y);
                    if (GetPhysicalCursorPos(out POINT phys))
                        pt = new Point(phys.X, phys.Y);
                    _dispatcher.BeginInvoke(new Action(() => LeftClick?.Invoke(pt)));
                }
                catch { }
            }
            return CallNextHookEx(_hook, nCode, wParam, lParam);
        }

        public void Dispose()
        {
            Stop();
        }
    }
}
