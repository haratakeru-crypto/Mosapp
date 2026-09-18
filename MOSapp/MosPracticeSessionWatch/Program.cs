using System;
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
        static void Main()
        {
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
