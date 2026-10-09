using System;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace MosPracticeSessionWatch
{
    static class Program
    {
        const string MutexName = "Local\\MosPracticeSessionWatch";
        const int HeartbeatIntervalMs = 120000;

        static readonly HttpClient Http = CreateClient();
        static System.Windows.Forms.Timer _heartbeatTimer;

        static string DataDirectory =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MOSapp");

        static string ExamineePath => Path.Combine(DataDirectory, "examinee.json");
        static string ClientIdPath => Path.Combine(DataDirectory, "client-id.txt");
        static string BaseUrlPath => Path.Combine(DataDirectory, "presence-baseurl.txt");
        static string KeyPath => Path.Combine(DataDirectory, "presence-key.txt");

        [STAThread]
        static void Main(string[] args)
        {
            // MOS 本体が COM で固まっても / 即終了しても Excel を落とすワンショット。
            // 例: MosPracticeSessionWatch.exe kill-excel 12345 638000000000000000 3000
            if (args != null && args.Length >= 1
                && string.Equals(args[0], "kill-excel", StringComparison.OrdinalIgnoreCase))
            {
                RunExcelKillWatch(args);
                return;
            }

            bool created;
            var mutex = new Mutex(true, MutexName, out created);
            if (!created) return;

            try
            {
                SystemEvents.SessionEnding += OnSessionEnding;
                SystemEvents.SessionSwitch += OnSessionSwitch;
                StartHeartbeatTimer();
                Application.Run(new ApplicationContext());
            }
            finally
            {
                SystemEvents.SessionEnding -= OnSessionEnding;
                SystemEvents.SessionSwitch -= OnSessionSwitch;
                mutex.ReleaseMutex();
                mutex.Dispose();
            }
        }

        /// <summary>
        /// kill-excel &lt;pid&gt; &lt;startTimeTicks|0&gt; &lt;delayMs&gt;
        /// 待機後、同一 Excel PID（起動時刻一致）が残っていれば Kill して終了する。
        /// 診断ログや COM は使わない。
        /// </summary>
        static void RunExcelKillWatch(string[] args)
        {
            try
            {
                if (args.Length < 4)
                    return;

                int pid;
                long startTicks;
                int delayMs;
                if (!int.TryParse(args[1], out pid) || pid <= 0)
                    return;
                if (!long.TryParse(args[2], out startTicks))
                    return;
                if (!int.TryParse(args[3], out delayMs) || delayMs < 0)
                    delayMs = 3000;
                if (delayMs > 60000)
                    delayMs = 60000;

                Thread.Sleep(delayMs);

                if (!IsMatchingExcelAlive(pid, startTicks))
                    return;

                try
                {
                    using (var p = Process.GetProcessById(pid))
                    {
                        if (p.HasExited)
                            return;
                        if (!string.Equals(p.ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase))
                            return;
                        p.Kill();
                    }
                }
                catch (ArgumentException)
                {
                    return;
                }
                catch
                {
                    return;
                }

                var sw = Stopwatch.StartNew();
                while (sw.ElapsedMilliseconds < 2000)
                {
                    if (!IsMatchingExcelAlive(pid, startTicks))
                        return;
                    Thread.Sleep(100);
                }
            }
            catch
            {
                /* ワンショット。失敗しても本体には影響させない */
            }
        }

        static bool IsMatchingExcelAlive(int pid, long startTicks)
        {
            try
            {
                using (var p = Process.GetProcessById(pid))
                {
                    if (p.HasExited)
                        return false;
                    if (!string.Equals(p.ProcessName, "EXCEL", StringComparison.OrdinalIgnoreCase))
                        return false;
                    if (startTicks > 0)
                    {
                        try
                        {
                            long actual = p.StartTime.ToUniversalTime().Ticks;
                            // 2秒相当のずれを許容
                            if (Math.Abs(actual - startTicks) > TimeSpan.FromSeconds(2).Ticks)
                                return false;
                        }
                        catch
                        {
                            return false;
                        }
                    }
                    return true;
                }
            }
            catch (ArgumentException)
            {
                return false;
            }
            catch
            {
                return false;
            }
        }

        static HttpClient CreateClient()
        {
            try
            {
                System.Net.ServicePointManager.SecurityProtocol |= System.Net.SecurityProtocolType.Tls12;
            }
            catch
            {
            }
            var client = new HttpClient();
            client.Timeout = TimeSpan.FromSeconds(8);
            return client;
        }

        static void StartHeartbeatTimer()
        {
            _heartbeatTimer = new System.Windows.Forms.Timer { Interval = HeartbeatIntervalMs };
            _heartbeatTimer.Tick += (s, e) =>
            {
                if (File.Exists(ExamineePath))
                    PostPresence("heartbeat");
            };
            _heartbeatTimer.Start();
        }

        static void OnSessionEnding(object sender, SessionEndingEventArgs e)
        {
            DeleteExamineeAndExit();
        }

        static void OnSessionSwitch(object sender, SessionSwitchEventArgs e)
        {
            if (e.Reason == SessionSwitchReason.SessionLogoff
                || e.Reason == SessionSwitchReason.ConsoleDisconnect)
            {
                DeleteExamineeAndExit();
            }
        }

        static void DeleteExamineeAndExit()
        {
            try
            {
                PostPresence("unregister");
            }
            catch
            {
            }

            try
            {
                if (File.Exists(ExamineePath))
                    File.Delete(ExamineePath);
            }
            catch
            {
            }

            try
            {
                Application.Exit();
            }
            catch
            {
            }
        }

        static void PostPresence(string action)
        {
            try
            {
                if (!File.Exists(ClientIdPath) || !File.Exists(BaseUrlPath) || !File.Exists(KeyPath))
                    return;

                string clientId = (File.ReadAllText(ClientIdPath) ?? "").Trim();
                string baseUrl = (File.ReadAllText(BaseUrlPath) ?? "").Trim().TrimEnd('/');
                string key = (File.ReadAllText(KeyPath) ?? "").Trim();
                if (clientId.Length < 8 || string.IsNullOrEmpty(baseUrl) || string.IsNullOrEmpty(key))
                    return;

                string json = "{\"clientId\":\"" + EscapeJson(clientId) + "\",\"action\":\"" + EscapeJson(action) + "\"}";
                using (var request = new HttpRequestMessage(HttpMethod.Post, baseUrl + "/api/mos-practice/presence"))
                {
                    request.Headers.TryAddWithoutValidation("X-MOS-Practice-Key", key);
                    request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                    using (var response = Http.SendAsync(request).GetAwaiter().GetResult())
                    {
                    }
                }
            }
            catch
            {
            }
        }

        static string EscapeJson(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value.Replace("\\", "\\\\").Replace("\"", "\\\"");
        }
    }
}
