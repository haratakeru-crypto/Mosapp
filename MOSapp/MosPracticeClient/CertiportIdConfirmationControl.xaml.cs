using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Web.WebView2.Core;
using Newtonsoft.Json;

namespace MosPracticeClient
{
    public partial class CertiportIdConfirmationControl : UserControl
    {
        public event EventHandler Verified;

        const string ExtractCertificateNameScript =
            @"(function(){
  function vis(el){
    if(!el) return '';
    return String(el.innerText||el.textContent||'').replace(/\s+/g,' ').trim();
  }
  function isRomaji(t){ return /^[A-Z][A-Z\s\-'.]+$/.test(t); }
  function isUi(t){
    if(!t || t.length>40) return true;
    var skip=['お知らせ','表示モード','合格した試験','印刷','登録情報','ログアウト','デジタル認定証','Digital','姓名の順','本人用','ようこそ'];
    for(var i=0;i<skip.length;i++){ if(t.indexOf(skip[i])>=0) return true; }
    return false;
  }
  function looksLikeName(t){
    if(isUi(t)||isRomaji(t)) return false;
    if(!/[\u3040-\u30ff\u4e00-\u9fff]/.test(t)) return false;
    return t.length>=2;
  }
  if(!document.body) return '';
  var nodes=document.body.getElementsByTagName('*');
  var found=false, names=[];
  for(var i=0;i<nodes.length;i++){
    var el=nodes[i];
    if(el.children && el.children.length>6) continue;
    var t=vis(el);
    if(!t) continue;
    if(!found && (t.indexOf('デジタル認定証')>=0 || /Digital\s*Certificate/i.test(t))){
      found=true;
      continue;
    }
    if(found && looksLikeName(t)){
      names.push(t);
      if(names.length>=8) break;
    }
  }
  for(var n=0;n<names.length;n++){
    if(/^[\u4e00-\u9fff々〆ヵヶ]{1,10}\s+[\u4e00-\u9fff々〆ヵヶ]{1,10}$/.test(names[n])) return names[n];
  }
  return names.length?names[0]:'';
})()";

        const string FillLoginScriptTemplate =
            @"(function(email,password){
  var inputs=document.querySelectorAll('input');
  var emailEl=null, passEl=null;
  for(var i=0;i<inputs.length;i++){
    var el=inputs[i];
    var t=String(el.type||'').toLowerCase();
    var n=String(el.name||el.id||el.getAttribute('placeholder')||'').toLowerCase();
    if(!emailEl && (t==='email' || n.indexOf('email')>=0 || n.indexOf('user')>=0 || t==='text')) emailEl=el;
    if(!passEl && t==='password') passEl=el;
  }
  function setVal(el,val){
    if(!el) return;
    el.focus();
    el.value=val;
    el.dispatchEvent(new Event('input',{bubbles:true}));
    el.dispatchEvent(new Event('change',{bubbles:true}));
  }
  setVal(emailEl,email);
  setVal(passEl,password);
  var clickables=document.querySelectorAll('button,input[type=submit],a');
  for(var j=0;j<clickables.length;j++){
    var label=String(clickables[j].innerText||clickables[j].value||'').replace(/\s+/g,' ').trim();
    if(/ログイン|log in|sign in/i.test(label)){
      clickables[j].click();
      return 'clicked';
    }
  }
  if(emailEl && passEl) return 'filled';
  return 'missing';
})";

        const string ClickLogoutScript =
            @"(function(){
  var els=document.querySelectorAll('a,button,input,span');
  for(var i=0;i<els.length;i++){
    var label=String(els[i].innerText||els[i].value||'').replace(/\s+/g,' ').trim();
    if(/ログアウト|log off|sign out|logout/i.test(label)){
      els[i].click();
      return 'clicked';
    }
  }
  return 'missing';
})()";

        bool _busy;
        bool _webViewReady;
        bool _fillLoginOnNextLoad;
        string _pendingEmail = "";
        string _capturedCertiportName = "";
        string _mosName = "";

        public CertiportIdConfirmationControl()
        {
            InitializeComponent();
            Loaded += CertiportIdConfirmationControl_Loaded;
        }

        void CertiportIdConfirmationControl_Loaded(object sender, RoutedEventArgs e)
        {
            Reload();
        }

        public void Reload()
        {
            var profile = ExamineeStore.Load();
            EmailBox.Text = profile?.CertiportLoginEmail ?? "";
            PasswordBox.Clear();
            ComparePanel.Visibility = Visibility.Collapsed;
            ChoicePanel.Visibility = Visibility.Collapsed;
            ManualNamePanel.Visibility = Visibility.Collapsed;
            _capturedCertiportName = "";
            _mosName = profile?.PersonName ?? "";
            UpdatePrerequisite(profile);
            if (ExamineeStore.IsCertiportVerified)
            {
                StatusText.Foreground = new SolidColorBrush(Color.FromRgb(0x16, 0x65, 0x34));
                StatusText.Text = "確認済みです。デジタル認定証: "
                    + (profile.CertiportDisplayName ?? "")
                    + " / 正しい名前: "
                    + (profile.CertiportCorrectName ?? "");
            }
            else
            {
                StatusText.Foreground = new SolidColorBrush(Color.FromRgb(0xB9, 0x1C, 0x1C));
                StatusText.Text = "";
            }
        }

        void UpdatePrerequisite(ExamineeProfile profile)
        {
            bool registered = ExamineeStore.IsRegistered;
            LoginButton.IsEnabled = registered && !_busy;
            EmailBox.IsEnabled = registered;
            PasswordBox.IsEnabled = registered;
            PrerequisiteText.Text = registered
                ? "大学名タブのお名前: " + (profile?.PersonName ?? "")
                : "先に「大学名」タブで大学名とお名前を登録してください。";
            PrerequisiteText.Foreground = registered
                ? new SolidColorBrush(Color.FromRgb(0x4B, 0x55, 0x63))
                : new SolidColorBrush(Color.FromRgb(0xB4, 0x53, 0x09));
        }

        void SignupButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Process.Start(new ProcessStartInfo(CertiportPortal.SignupUrl) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                SetError("ブラウザを開けませんでした。\n" + ex.Message);
            }
        }

        async void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            if (_busy) return;
            if (!ExamineeStore.IsRegistered)
            {
                SetError("先に「大学名」タブで大学名とお名前を登録してください。");
                return;
            }

            string email = (EmailBox.Text ?? "").Trim();
            string password = PasswordBox.SecurePassword != null
                ? new NetworkCredential("", PasswordBox.SecurePassword).Password
                : "";
            if (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(password))
            {
                SetError("Certiport ID とパスワードを入力してください。");
                return;
            }

            _busy = true;
            LoginButton.IsEnabled = false;
            ComparePanel.Visibility = Visibility.Collapsed;
            SetBusy("Certiport にログインしています…");
            _pendingEmail = email;
            _fillLoginOnNextLoad = true;
            _capturedCertiportName = "";

            try
            {
                await EnsureWebViewAsync();
                CertiportWebView.CoreWebView2.Navigate(CertiportPortal.LoginUrl);
            }
            catch (Exception ex)
            {
                _busy = false;
                UpdatePrerequisite(ExamineeStore.Load());
                SetError("Certiport 画面を開けませんでした。WebView2 Runtime がインストールされているか確認してください。\n" + ex.Message);
            }
        }

        async Task EnsureWebViewAsync()
        {
            if (_webViewReady && CertiportWebView.CoreWebView2 != null) return;
            string folder = Path.Combine(ExamineeStore.DataDirectory, "webview-certiport");
            Directory.CreateDirectory(folder);
            var env = await CoreWebView2Environment.CreateAsync(null, folder);
            await CertiportWebView.EnsureCoreWebView2Async(env);
            CertiportWebView.NavigationCompleted -= CertiportWebView_NavigationCompleted;
            CertiportWebView.NavigationCompleted += CertiportWebView_NavigationCompleted;
            _webViewReady = true;
        }

        async void CertiportWebView_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (!e.IsSuccess) return;
            string url = "";
            try { url = CertiportWebView.Source != null ? CertiportWebView.Source.AbsoluteUri : ""; }
            catch { }

            if (_fillLoginOnNextLoad && IsLoginUrl(url))
            {
                _fillLoginOnNextLoad = false;
                await FillLoginFormAsync();
                await Task.Delay(8000);
                if (_busy)
                {
                    string still = "";
                    try { still = CertiportWebView.Source != null ? CertiportWebView.Source.AbsoluteUri : ""; }
                    catch { }
                    if (IsLoginUrl(still))
                    {
                        FinishBusyError("ログインできませんでした。ID・パスワードを確認してください。");
                    }
                }
                return;
            }

            if (_busy && !IsLoginUrl(url))
            {
                await Task.Delay(1500);
                await CaptureCertificateNameAsync();
            }
        }

        async Task FillLoginFormAsync()
        {
            string emailJson = JsonConvert.SerializeObject(_pendingEmail);
            string password = new NetworkCredential("", PasswordBox.SecurePassword).Password;
            string passwordJson = JsonConvert.SerializeObject(password);
            string script = FillLoginScriptTemplate + "(" + emailJson + "," + passwordJson + ");";
            try
            {
                await CertiportWebView.ExecuteScriptAsync(script);
            }
            catch (Exception ex)
            {
                FinishBusyError("ログイン画面の入力に失敗しました。\n" + ex.Message);
            }
        }

        async Task CaptureCertificateNameAsync()
        {
            string name = "";
            for (int i = 0; i < 4 && string.IsNullOrEmpty(name); i++)
            {
                name = await ReadCertificateNameAsync();
                if (string.IsNullOrEmpty(name))
                    await Task.Delay(1200);
            }

            var profile = ExamineeStore.Load();
            _mosName = profile?.PersonName ?? "";
            _capturedCertiportName = CertiportPortal.NormalizePersonName(name);
            ShowCompare(_capturedCertiportName, string.IsNullOrEmpty(_capturedCertiportName));
        }

        async Task<string> ReadCertificateNameAsync()
        {
            try
            {
                string json = await CertiportWebView.ExecuteScriptAsync(ExtractCertificateNameScript);
                string value = JsonConvert.DeserializeObject<string>(json) ?? "";
                return CertiportPortal.NormalizePersonName(value);
            }
            catch
            {
                return "";
            }
        }

        void ShowCompare(string certiportName, bool needManual)
        {
            _busy = false;
            UpdatePrerequisite(ExamineeStore.Load());
            ComparePanel.Visibility = Visibility.Visible;
            MosNameText.Text = "MOSアプリのお名前: " + _mosName;
            CertiportNameText.Text = "デジタル認定証の氏名: "
                + (string.IsNullOrEmpty(certiportName) ? "（取得できませんでした）" : certiportName);
            ManualNamePanel.Visibility = needManual ? Visibility.Visible : Visibility.Collapsed;
            ManualNameBox.Text = certiportName ?? "";

            bool match = !needManual && CertiportPortal.NamesMatch(_mosName, certiportName);
            ChoicePanel.Visibility = match || needManual ? Visibility.Collapsed : Visibility.Visible;
            ChooseMosRadio.IsChecked = true;
            ChooseCertiportRadio.IsChecked = false;
            MatchResultText.Foreground = match
                ? new SolidColorBrush(Color.FromRgb(0x16, 0x65, 0x34))
                : new SolidColorBrush(Color.FromRgb(0xB4, 0x53, 0x09));
            MatchResultText.Text = needManual
                ? "認定証の氏名を入力してから一時保存してください。"
                : (match
                    ? "名前は一致しています。この内容で一時保存できます。"
                    : "名前が一致しません。正しい方を選んでください。");
            StatusText.Foreground = new SolidColorBrush(Color.FromRgb(0x16, 0x65, 0x34));
            StatusText.Text = "ログインできました。名前を確認してください。";
        }

        async void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            if (!ExamineeStore.IsRegistered)
            {
                SetError("先に「大学名」タブで大学名とお名前を登録してください。");
                return;
            }

            string certiportName = _capturedCertiportName;
            if (ManualNamePanel.Visibility == Visibility.Visible)
            {
                certiportName = CertiportPortal.NormalizePersonName(ManualNameBox.Text);
                _capturedCertiportName = certiportName;
            }
            if (string.IsNullOrEmpty(certiportName))
            {
                SetError("デジタル認定証の氏名を確認できません。画面の漢字氏名を入力してください。");
                return;
            }

            bool namesMatch = CertiportPortal.NamesMatch(_mosName, certiportName);
            string correctName = _mosName;
            if (!namesMatch)
            {
                if (ChoicePanel.Visibility != Visibility.Visible)
                {
                    CertiportNameText.Text = "デジタル認定証の氏名: " + certiportName;
                    ChoicePanel.Visibility = Visibility.Visible;
                    ChooseMosRadio.IsChecked = true;
                    MatchResultText.Foreground = new SolidColorBrush(Color.FromRgb(0xB4, 0x53, 0x09));
                    MatchResultText.Text = "名前が一致しません。正しい方を選んでから、もう一度一時保存してください。";
                    return;
                }
                if (ChooseCertiportRadio.IsChecked == true)
                    correctName = certiportName;
                else if (ChooseMosRadio.IsChecked == true)
                    correctName = _mosName;
                else
                {
                    SetError("正しい名前を選んでください。");
                    return;
                }
            }

            var existing = ExamineeStore.Load() ?? new ExamineeProfile();
            existing.CertiportLoginEmail = (EmailBox.Text ?? "").Trim();
            existing.CertiportDisplayName = certiportName;
            existing.CertiportCorrectName = correctName;
            existing.CertiportNameMatchesMos = namesMatch;
            existing.CertiportVerifiedAt = CertiportPortal.FormatVerifiedAt(DateTime.Now);
            ExamineeStore.Save(existing);

            SetBusy("Rabbit の一時保存へ送っています…");
            var send = await CertiportLinkClient.SubmitAsync(existing);
            if (!send.Success)
            {
                SetError("名前はPCに保存しましたが、Rabbit への送信に失敗しました。\n" + (send.Error ?? ""));
                Verified?.Invoke(this, EventArgs.Empty);
                await SignOutQuietAsync();
                return;
            }

            StatusText.Foreground = new SolidColorBrush(Color.FromRgb(0x16, 0x65, 0x34));
            StatusText.Text = "一時保存しました。Certiport からサインアウトします。";
            Verified?.Invoke(this, EventArgs.Empty);
            PasswordBox.Clear();
            await SignOutQuietAsync();
            StatusText.Foreground = new SolidColorBrush(Color.FromRgb(0x16, 0x65, 0x34));
            StatusText.Text = "確認済みです。デジタル認定証: " + certiportName + " / 正しい名前: " + correctName;
        }

        async Task SignOutQuietAsync()
        {
            try
            {
                if (CertiportWebView.CoreWebView2 != null)
                {
                    await CertiportWebView.ExecuteScriptAsync(ClickLogoutScript);
                    await Task.Delay(800);
                    CertiportWebView.CoreWebView2.CookieManager.DeleteAllCookies();
                    CertiportWebView.CoreWebView2.Navigate("about:blank");
                }
            }
            catch
            {
            }
        }

        static bool IsLoginUrl(string url)
        {
            return !string.IsNullOrEmpty(url)
                && url.IndexOf("/portal/login", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        void SetError(string message)
        {
            StatusText.Foreground = new SolidColorBrush(Color.FromRgb(0xB9, 0x1C, 0x1C));
            StatusText.Text = message ?? "";
        }

        void SetBusy(string message)
        {
            StatusText.Foreground = new SolidColorBrush(Color.FromRgb(0x1F, 0x29, 0x37));
            StatusText.Text = message ?? "";
        }

        void FinishBusyError(string message)
        {
            _busy = false;
            _fillLoginOnNextLoad = false;
            UpdatePrerequisite(ExamineeStore.Load());
            SetError(message);
        }
    }
}
