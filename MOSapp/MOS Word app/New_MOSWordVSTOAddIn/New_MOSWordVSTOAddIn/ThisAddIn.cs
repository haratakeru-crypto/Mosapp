using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Libraries;
using Libraries.Group1;
using Word = Microsoft.Office.Interop.Word;

namespace New_MOSWordVSTOAddIn
{
    public partial class ThisAddIn
    {
        /// <summary>Word の文書編集ペインのウィンドウクラス。リボン上の入力欄は通常これを祖先に持たない。</summary>
        private const string WordDocumentPaneClassName = "_WwG";

        private WordDestructiveMonitor _destructiveMonitor;
        private Timer _showAllFastTimer;
        private Timer _showAllPollTimer;
        private bool? _lastShowAllState;
        private const int ShowAllFastPollIntervalMs = 300;
        private static readonly string EvidenceFlushFilePath = Path.Combine(Path.GetTempPath(), "mos_word_flush_evidence.txt");
        private static readonly string CloseNavigationFilePath = Path.Combine(Path.GetTempPath(), "mos_word_close_navigation.txt");

        /// <summary>ナビゲーションウィンドウが開いているか（CommandBars["Navigation"] で同期）。</summary>
        private bool _navigationPaneOpen;
        private int _lastColumnBreakCount;
        private string _lastTask1_2_03ColorFingerprint;

        /// <summary>3-1: 先頭セクションが「やや狭い」余白プリセット相当か。</summary>
        private bool? _lastMarginsModerate;

        /// <summary>全セクションの向きを連結したフィンガープリント（先頭セクションのみでは 3-3 とチェッカーが不一致になるため）。</summary>
        private string _lastOrientationFingerprint;
        private string _lastPageBorderFingerprint;
        private bool? _lastHeading1LineSimple;
        private bool? _lastLineStylish;
        private string _lastStyleSetLineFingerprint;
        private bool _styleSetLineSimpleLoggedThisDoc;
        /// <summary>4-5: 下書き1 透かしの有無（社外秘・至急は含めない）</summary>
        private bool? _lastDraft1WatermarkFound;
        /// <summary>透かし指紋（スナップショット比較・4-1 等の [Op] 補完用）</summary>
        private string _lastWatermarkFingerprint;
        private int _suppressWatermarkPollLogs;
        private int _lastLaptopWrapType = -1;

        /// <summary>リボンが既にログした直後のポーリング二重記録を抑止する（約2ティック）。</summary>
        private int _suppressOrientationPollLogs;
        private int _suppressPageBorderPollLogs;
        private int _suppressStyleSetPollLogs;
        private int _suppressResolvePollLogs;
        private int _suppressDeletePollLogs;
        private int _heavyCheckTickCounter;

        /// <summary>4-3: 未解決エコ件数。1→0 で ReviewResolveComment を補完ログ。</summary>
        private int _lastUnresolvedEcoCommentCount = -1;

        /// <summary>4-3: エコ関連コメント総数（解決済み含む）。0 へ減ったら削除ログ候補。</summary>
        private int _lastEcoRelatedCommentCount = -1;

        /// <summary>4-3: Comment.Index → Done の前回値。</summary>
        private readonly Dictionary<int, bool> _ecoCommentDoneByIndex = new Dictionary<int, bool>();

        /// <summary>4-3: 当該試行で ReviewResolveComment を記録済みか。</summary>
        private bool _ecoResolveLoggedForAttempt;

        private bool? _lastAnyEcoPhraseInBalloons;

        private string _ecoBaselineDocumentKey;

        /// <summary>7-1: 前ティックの ActiveDocument.FullName（文書切替でベースライン再取得するため）</summary>
        private string _p7LastCompatDocFullName;

        /// <summary>7-1: 同一文書での前回 CompatibilityMode。未設定は -1。</summary>
        private int _p7LastCompatMode = -1;

        /// <summary>7-2: 前ティックの Project7 文書 FullName。</summary>
        private string _p7LastCompanyDocFullName;

        /// <summary>7-2: 同一 Project7 文書で前ティック時点の Company が目標値だったか。</summary>
        private bool _p7LastCompanyMatched;

        private const string P7CompanyTarget = "ラビット出版";

        private string _p7IntegralTrackedFullName;
        private bool _p7IntegralLastDetected;

        private const string P7RdTxtFileName = "朗読会.txt";
        private const string P7RdDocmFileName = "朗読会.docm";

        /// <summary>7-4/7-5: 前ティックで追跡していた ActiveDocument.FullName</summary>
        private string _p7FileSaveAsTrackedFullName;

        /// <summary>7-4: 前ティックで ActiveDocument が朗読会.txt だったか</summary>
        private bool _p7FileSaveAsLastTxt;

        /// <summary>7-5: 前ティックで ActiveDocument が朗読会.docm だったか</summary>
        private bool _p7FileSaveAsLastDocm;

        internal void RegisterRibbonLoggedPageOrientation()
        {
            _suppressOrientationPollLogs = 2;
        }

        internal void RegisterRibbonLoggedPageBorders()
        {
            _suppressPageBorderPollLogs = 2;
        }

        internal void RegisterRibbonLoggedStyleSetLineSimple()
        {
            _suppressStyleSetPollLogs = 2;
        }

        internal void RegisterRibbonLoggedReviewDeleteComment()
        {
            _suppressResolvePollLogs = 3;
            _suppressDeletePollLogs = 2;
        }

        internal void RegisterRibbonLoggedWatermark()
        {
            _suppressWatermarkPollLogs = 2;
        }

        /// <summary>
        /// タスク切替時: 透かしポーリングのベースラインを現文書に合わせる（[Op] / Executed は出さない）。
        /// 4-5 の下書き1が 4-6 表示後に「変化」と誤検知されるのを防ぐ。
        /// </summary>
        internal void SyncWatermarkPollingBaselineOnTaskSwitch(int projectId)
        {
            try
            {
                Word.Document doc = TryGetProjectDocument(projectId);
                ApplyWatermarkPollingBaseline(doc);
                _suppressWatermarkPollLogs = 2;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] SyncWatermarkPollingBaselineOnTaskSwitch: " + ex.Message);
            }
        }

        /// <summary>4-3 開始時: 未解決エコ件数のベースラインを現文書から再取得する。</summary>
        internal void SyncEcoCommentBaselineOnTaskSwitch(int projectId, int taskId)
        {
            if (projectId != 4 || taskId != 3)
                return;
            _lastUnresolvedEcoCommentCount = -1;
            _lastEcoRelatedCommentCount = -1;
            _ecoBaselineDocumentKey = null;
            _ecoCommentDoneByIndex.Clear();
            _ecoResolveLoggedForAttempt = false;
            try
            {
                Word.Document doc = TryGetProjectDocument(projectId);
                if (doc != null)
                    SeedEcoCommentBaseline(doc);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] SyncEcoCommentBaselineOnTaskSwitch: " + ex.Message);
            }
        }

        /// <summary>
        /// 4-5 離脱時の保険: 文書に下書き1透かしが残っていれば 4-5 の証跡を補記する。
        /// 重いポーリング前に 4-6/4-7 へ遷移した場合でも、4-5 の第2段採点（証跡必須）を満たせるようにする。
        /// </summary>
        internal void EnsureTask45WatermarkEvidenceBeforeLeave(int previousProjectId, int previousTaskId, int nextProjectId, int nextTaskId)
        {
            if (previousProjectId != 4 || previousTaskId != 5)
                return;
            if (nextProjectId == 4 && nextTaskId == 5)
                return;

            try
            {
                Word.Document doc = TryGetProjectDocument(4);
                if (doc == null)
                    return;
                string norm = WordWatermarkInspection.NormalizeXml(doc.WordOpenXML);
                if (WordWatermarkInspection.IsSample2Watermark(norm))
                    WordEvidenceHelper.LogCommandWithEvidence("Watermark");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] EnsureTask45WatermarkEvidenceBeforeLeave: " + ex.Message);
            }
        }

        private void ApplyWatermarkPollingBaseline(Word.Document doc)
        {
            if (doc == null)
            {
                _lastWatermarkFingerprint = "None";
                _lastDraft1WatermarkFound = false;
                return;
            }

            try
            {
                string norm = WordWatermarkInspection.NormalizeXml(doc.WordOpenXML);
                _lastWatermarkFingerprint = WordWatermarkInspection.GetWatermarkFingerprint(norm);
                _lastDraft1WatermarkFound = string.Equals(_lastWatermarkFingerprint, "Draft1Diagonal", StringComparison.Ordinal);
            }
            catch
            {
                _lastWatermarkFingerprint = "None";
                _lastDraft1WatermarkFound = false;
            }
        }

        internal Word.Document TryGetProjectDocument(int projectId)
        {
            try
            {
                var app = Application;
                if (app == null)
                    return null;
                for (int i = app.Documents.Count; i >= 1; i--)
                {
                    Word.Document doc = app.Documents[i];
                    string name = Path.GetFileName(doc.FullName ?? "");
                    if (name.StartsWith("Project" + projectId, StringComparison.OrdinalIgnoreCase)
                        || name.StartsWith("project" + projectId, StringComparison.OrdinalIgnoreCase))
                        return doc;
                }

                return app.ActiveDocument;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>4-5 正答証跡と採点ゲート②用 [Op] Watermark（指紋変化・リボン透かし）。</summary>
        private static void LogWatermarkForDestructiveGate(string fingerprint)
        {
            Logger.LogOperation("Watermark", fingerprint ?? "");
            if (string.Equals(fingerprint, "Sample2", StringComparison.Ordinal)
                || string.Equals(fingerprint, "Draft1Diagonal", StringComparison.Ordinal))
                WordEvidenceHelper.LogCommandWithEvidence("Watermark");
        }

        /// <summary>4-6 正答証跡と採点ゲート②用 [Op] PageBorders。</summary>
        private static void LogPageBordersForDestructiveGate()
        {
            WordEvidenceHelper.LogCommandWithEvidence("PageBorders");
            Logger.LogOperation("PageBorders", "");
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[New_MOSWordVSTOAddIn] Add-in started");
            System.Diagnostics.Debug.WriteLine($"[New_MOSWordVSTOAddIn] Log file: {Logger.GetLogFilePath()}");

            Logger.WriteVstoHeartbeat();

            this.Application.DocumentChange += Application_DocumentChange;
            this.Application.DocumentBeforeSave += Application_DocumentBeforeSave;
            RefreshBaselineFromActiveDocumentNoLog();
            StartShowAllPolling();
            _destructiveMonitor = new WordDestructiveMonitor(this);
            _destructiveMonitor.Start();
        }

        /// <summary>
        /// カスタムリボン（編集記号ログ等）を返す。テンプレートの Ribbon デザイナは使わず自前 Ribbon を使用。
        /// </summary>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            System.Diagnostics.Debug.WriteLine("[ThisAddIn] CreateRibbonExtensibilityObject called");
            try
            {
                var ribbon = new Ribbon();
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] Ribbon instance created");
                return ribbon;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ThisAddIn] CreateRibbonExtensibilityObject error: {ex.Message}\r\n{ex.StackTrace}");
                throw;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                this.Application.DocumentChange -= Application_DocumentChange;
                this.Application.DocumentBeforeSave -= Application_DocumentBeforeSave;
            }
            catch
            {
                // アンインストール時など Application が無い場合
            }

            _destructiveMonitor?.Stop();
            _destructiveMonitor = null;
            _showAllFastTimer?.Stop();
            _showAllFastTimer?.Dispose();
            _showAllFastTimer = null;
            _showAllPollTimer?.Stop();
            _showAllPollTimer?.Dispose();
            _showAllPollTimer = null;
        }

        private void Application_DocumentChange()
        {
            RefreshBaselineFromActiveDocumentNoLog();
        }

        /// <summary>
        /// 7-4/7-5: 上書き保存時など、保存前から対象ファイル名のとき専用ログを付与する。
        /// 初回の「名前を付けて保存」は ShowAllPoll（毎ティック）で保存後の FullName を検知する。
        /// </summary>
        private void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        {
            try
            {
                if (Doc == null) return;
                string fullName;
                try { fullName = Doc.FullName; }
                catch { return; }
                if (string.IsNullOrEmpty(fullName)) return;

                LogFileSaveAsCommandForPath(fullName);
            }
            catch { /* ignore */ }
        }

        /// <summary>
        /// 文書切替・リセット後の誤検知を防ぐため、現在の文書状態をベースラインにする（ログは出さない）。
        /// ホストアプリが LogReader.ClearLog しても VSTO 側のメモリは維持されるため、
        /// 文書が開き直されたタイミングで必ず整合させる。
        /// </summary>
        private void RefreshBaselineFromActiveDocumentNoLog()
        {
            try
            {
                var app = this.Application;
                if (app == null || app.Documents.Count == 0)
                    return;

                PreserveSelectionDuring(app, () => RefreshBaselineFromActiveDocumentNoLogCore(app));
            }
            catch
            {
                // COM 初期化中などは無視
            }
        }

        private void RefreshBaselineFromActiveDocumentNoLogCore(Word.Application app)
        {
            // ActiveDocument は Word が管理する参照のため ReleaseComObject しない
            Word.Document doc = app.ActiveDocument;

                if (app.ActiveWindow?.View != null)
                    _lastShowAllState = app.ActiveWindow.View.ShowAll;

                // 4-3: コメントペイン操作中も未解決件数の遷移を追跡（軽量のため defer より前に実行）
                UpdateEcoCommentBaselineAndMaybeLog(doc, CountUnresolvedEcoComments(doc));

                // ナビペイン・検索ダイアログ・コメントペイン操作中は侵入的 COM をスキップ
                if (ShouldDeferIntrusiveDocumentCom(app))
                    return;

                _lastColumnBreakCount = CountColumnBreaks(doc);
                _lastMarginsModerate = IsMarginsModeratePreset(doc);
                _lastOrientationFingerprint = GetAllSectionsOrientationFingerprint(doc);
                _lastPageBorderFingerprint = WordWatermarkInspection.GetPageBorderFingerprint(doc);
                _lastHeading1LineSimple = WordWatermarkInspection.IsDocumentStyleSetLineSimple(doc);
                _lastLineStylish = WordWatermarkInspection.IsDocumentStyleSetLineStylish(doc);
                _lastStyleSetLineFingerprint = WordWatermarkInspection.GetStyleSetLineFingerprint(doc);
                _styleSetLineSimpleLoggedThisDoc = false;
                _lastTask1_2_03ColorFingerprint = GetTask1_2_03ColorFingerprint(doc);

                ApplyWatermarkPollingBaseline(doc);

                // 7-1: 互換モードのベースライン（ポーリングで 非15→15 の遷移を検知するため）
                try
                {
                    _p7LastCompatDocFullName = doc.FullName;
                    _p7LastCompatMode = (int)doc.CompatibilityMode;
                }
                catch
                {
                    _p7LastCompatDocFullName = null;
                    _p7LastCompatMode = -1;
                }

                // 7-2: Project7 上の Company ベースライン
                try
                {
                    string fn = doc.FullName;
                    if (IsProject7DocumentPath(fn))
                    {
                        _p7LastCompanyDocFullName = fn;
                        _p7LastCompanyMatched = string.Equals(TryGetDocumentCompany(doc), P7CompanyTarget, StringComparison.Ordinal);
                    }
                }
                catch
                {
                    _p7LastCompanyDocFullName = null;
                    _p7LastCompanyMatched = false;
                }
        }

        /// <summary>4-3: 解決（Done）・削除（件数0）を検知して証跡を記録。</summary>
        private void UpdateEcoCommentBaselineAndMaybeLog(Word.Document doc, int newUnresolvedEcoCount)
        {
            try
            {
                if (doc != null)
                {
                    string key = null;
                    try { key = doc.FullName; } catch { }
                    if (!string.IsNullOrEmpty(key) &&
                        !string.Equals(_ecoBaselineDocumentKey, key, StringComparison.OrdinalIgnoreCase))
                    {
                        _ecoBaselineDocumentKey = key;
                        _lastUnresolvedEcoCommentCount = -1;
                        _lastEcoRelatedCommentCount = -1;
                        _lastAnyEcoPhraseInBalloons = null;
                        _ecoCommentDoneByIndex.Clear();
                        _ecoResolveLoggedForAttempt = false;
                    }
                }
            }
            catch { }

            bool resolveLoggedThisTick = DetectEcoCommentDoneTransitions(doc);

            int ecoRelatedCount = CountEcoRelatedComments(doc);
            bool anyPhrase = ecoRelatedCount > 0;

            // 未解決→0 は「解決」または「削除」で起きる。コメントがまだ残っているときだけ解決とみなす。
            bool willLogUnresolved = _lastUnresolvedEcoCommentCount >= 0 && _lastUnresolvedEcoCommentCount > 0 && newUnresolvedEcoCount == 0;
            int prevUnresolved = _lastUnresolvedEcoCommentCount;
            int prevEcoRelated = _lastEcoRelatedCommentCount;

            if (willLogUnresolved && doc != null && ecoRelatedCount > 0)
                TryLogReviewResolveComment();

            // 削除ログ: 解決済みのあと削除された場合のみ（削除のみでは prevUnresolved>0 のまま消える）
            bool deleteOnlySuspect = prevUnresolved > 0 && ecoRelatedCount == 0 && !resolveLoggedThisTick;
            if (prevEcoRelated > 0 && ecoRelatedCount == 0 && _ecoResolveLoggedForAttempt && !deleteOnlySuspect)
            {
                if (_suppressDeletePollLogs > 0)
                    _suppressDeletePollLogs--;
                else
                    WordEvidenceHelper.LogCommandWithEvidence("ReviewDeleteComment");
            }

            _lastUnresolvedEcoCommentCount = newUnresolvedEcoCount;
            _lastEcoRelatedCommentCount = ecoRelatedCount;
            _lastAnyEcoPhraseInBalloons = anyPhrase;
        }

        private void SeedEcoCommentBaseline(Word.Document doc)
        {
            _ecoCommentDoneByIndex.Clear();
            foreach (Word.Comment c in doc.Comments)
            {
                try
                {
                    if (!CommentRelatesToEcoPhrasePoll(c, doc))
                        continue;
                    bool done = false;
                    try { done = c.Done; } catch { }
                    _ecoCommentDoneByIndex[c.Index] = done;
                }
                finally
                {
                    Marshal.ReleaseComObject(c);
                }
            }
            int unresolved = CountUnresolvedEcoComments(doc);
            int related = CountEcoRelatedComments(doc);
            _lastUnresolvedEcoCommentCount = unresolved;
            _lastEcoRelatedCommentCount = related;
            _lastAnyEcoPhraseInBalloons = related > 0;
        }

        private bool DetectEcoCommentDoneTransitions(Word.Document doc)
        {
            if (doc?.Comments == null)
                return false;
            bool loggedResolve = false;
            var seen = new HashSet<int>();
            foreach (Word.Comment c in doc.Comments)
            {
                try
                {
                    if (!CommentRelatesToEcoPhrasePoll(c, doc))
                        continue;
                    int idx = c.Index;
                    seen.Add(idx);
                    bool done = false;
                    try { done = c.Done; } catch { }
                    if (_ecoCommentDoneByIndex.TryGetValue(idx, out bool wasDone) && !wasDone && done)
                        loggedResolve |= TryLogReviewResolveComment();
                    _ecoCommentDoneByIndex[idx] = done;
                }
                finally
                {
                    Marshal.ReleaseComObject(c);
                }
            }
            var stale = _ecoCommentDoneByIndex.Keys.Where(k => !seen.Contains(k)).ToList();
            foreach (int k in stale)
                _ecoCommentDoneByIndex.Remove(k);
            return loggedResolve;
        }

        private bool TryLogReviewResolveComment()
        {
            if (_suppressResolvePollLogs > 0)
            {
                _suppressResolvePollLogs--;
                return false;
            }
            WordEvidenceHelper.LogCommandWithEvidence("ReviewResolveComment");
            _ecoResolveLoggedForAttempt = true;
            return true;
        }

        /// <summary>
        /// 1-1 編集記号の表示/非表示は Word の idMso でフックできないため、
        /// View.ShowAll の状態をポーリングし、変化時に ShowAll をログに記録する。
        /// ShowAll は軽量のため専用の高速タイマー（300ms）を使う。
        /// 向き・ページ罫線・スタイルセット（線シンプル相当）も idMso が発火しない経路があるため別タイマーで差分検知する。
        /// </summary>
        private void StartShowAllPolling()
        {
            _showAllFastTimer = new Timer { Interval = ShowAllFastPollIntervalMs };
            _showAllFastTimer.Tick += ShowAllFastPoll_Tick;
            _showAllFastTimer.Start();

            // 重い COM 差分検知は従来どおり控えめな間隔
            _showAllPollTimer = new Timer { Interval = 1200 };
            _showAllPollTimer.Tick += ShowAllPoll_Tick;
            _showAllPollTimer.Start();
        }

        private void ShowAllFastPoll_Tick(object sender, EventArgs e)
        {
            Logger.WriteVstoHeartbeat();
            ProcessEvidenceFlushRequest();
            ProcessCloseNavigationRequest();
            try
            {
                var app = this.Application;
                if (app?.ActiveWindow?.View == null || app.Documents.Count == 0)
                    return;
                SyncNavigationPaneOpenState(app);
                // View.ShowAll の読み取りのみ。Selection / ScreenUpdating は触らない（Copilot 等の UI 点滅を抑える）。
                UpdateShowAllPolling(app);
                if (IsProject4CommentTaskActive())
                {
                    try
                    {
                        Word.Document doc = app.ActiveDocument;
                        if (doc != null)
                            UpdateEcoCommentBaselineAndMaybeLog(doc, CountUnresolvedEcoComments(doc));
                    }
                    catch { }
                }
            }
            catch { /* ignore */ }
        }

        /// <summary>採点アプリが mos_word_flush_evidence.txt を置いたとき、即座に ShowAll 状態を同期してログに反映する。</summary>
        private void ProcessEvidenceFlushRequest()
        {
            if (!File.Exists(EvidenceFlushFilePath))
                return;

            try
            {
                UpdateShowAllPolling(this.Application);
            }
            catch { /* ignore */ }
            finally
            {
                try { File.Delete(EvidenceFlushFilePath); } catch { /* ignore */ }
            }
        }

        /// <summary>採点アプリが mos_word_close_navigation.txt を置いたとき、ナビゲーションウィンドウが開いていれば閉じる。</summary>
        private void ProcessCloseNavigationRequest()
        {
            if (!File.Exists(CloseNavigationFilePath))
                return;

            try
            {
                CloseNavigationPaneIfOpen();
            }
            catch { /* ignore */ }
            finally
            {
                try { File.Delete(CloseNavigationFilePath); } catch { /* ignore */ }
            }
        }

        private bool TryIsNavigationPaneVisible(Word.Application app, out bool documentMapOpen)
        {
            documentMapOpen = false;
            if (app?.CommandBars == null)
                return false;

            try
            {
                var nav = app.CommandBars["Navigation"];
                if (nav != null && nav.Visible)
                    return true;
            }
            catch { /* ignore */ }

            try
            {
                if (app.ActiveWindow != null && app.ActiveWindow.DocumentMap)
                {
                    documentMapOpen = true;
                    return true;
                }
            }
            catch { /* ignore */ }

            return false;
        }

        private void SyncNavigationPaneOpenState(Word.Application app)
        {
            _navigationPaneOpen = TryIsNavigationPaneVisible(app, out _);
        }

        private void CloseNavigationPaneIfOpen()
        {
            var app = Application;
            if (app == null)
                return;

            if (!TryIsNavigationPaneVisible(app, out _))
                return;

            try
            {
                app.CommandBars["Navigation"].Visible = false;
            }
            catch { /* ignore */ }

            try
            {
                if (app.ActiveWindow != null)
                    app.ActiveWindow.DocumentMap = false;
            }
            catch { /* ignore */ }

            _navigationPaneOpen = false;
        }

        /// <summary>
        /// 1-1: 採点は <see cref="Word.View.ShowAll"/> のみ参照するため、Options.ShowAll との OR は使わない
        /// （OR だと Options が true の環境で View のトグルを検知できない）。
        /// リボン操作時はフォーカスが編集ペイン外になるため、呼び出し元はフォーカス判定より前にすること。
        /// </summary>
        private void UpdateShowAllPolling(Word.Application app)
        {
            if (app?.ActiveWindow?.View == null)
                return;

            bool viewShowAll = app.ActiveWindow.View.ShowAll;
            if (_lastShowAllState.HasValue && _lastShowAllState.Value != viewShowAll)
                WordEvidenceHelper.LogCommandWithEvidence("ShowAll");
            _lastShowAllState = viewShowAll;
        }

        private void ShowAllPoll_Tick(object sender, EventArgs e)
        {
            try
            {
                var app = this.Application;
                if (app?.ActiveWindow?.View == null || app.Documents.Count == 0)
                    return;

                PreserveSelectionDuring(app, () => ShowAllPoll_TickCore(app));
            }
            catch
            {
                // ドキュメント未表示などで COM エラーになることがあるため無視
            }
        }

        private void ShowAllPoll_TickCore(Word.Application app)
        {
            // ActiveDocument は Word が管理する参照のため ReleaseComObject しない
            Word.Document doc = app.ActiveDocument;

            // 4-3: コメントペイン操作中も未解決件数の遷移を追跡
            UpdateEcoCommentBaselineAndMaybeLog(doc, CountUnresolvedEcoComments(doc));

            // ナビペイン・Ctrl+F 検索・コメントペイン操作中は heavy Find 等の侵入的 COM を止める
            if (ShouldDeferIntrusiveDocumentCom(app))
                return;

                // 7-4/7-5: FullName のみの軽量検知（毎ティック≈1.2秒）。重いポーリング（約6秒）だと次プロジェクト押下前に取りこぼす。
                try
                {
                    UpdateFileSaveAsPolling(doc);
                }
                catch { }

                // 7-2: 「ファイルの情報」で会社を設定する操作は編集ペイン外のため、フォーカス判定より前にポーリングする。
                try
                {
                    UpdateP7CompanyPolling(doc);
                }
                catch { }

                bool marginsModerate = IsMarginsModeratePreset(doc);
                if (_lastMarginsModerate.HasValue && marginsModerate && !_lastMarginsModerate.Value)
                    WordEvidenceHelper.LogCommandWithEvidence("PageMarginsModerate");
                _lastMarginsModerate = marginsModerate;

                string orientFp = GetAllSectionsOrientationFingerprint(doc);
                if (_lastOrientationFingerprint != null && orientFp != _lastOrientationFingerprint)
                {
                    if (_suppressOrientationPollLogs > 0)
                        _suppressOrientationPollLogs--;
                    else
                        WordEvidenceHelper.LogCommandWithEvidence("PageOrientationPortraitLandscape");
                }
                _lastOrientationFingerprint = orientFp;

                // 7-1: Ribbon の UpgradeDocument は発火しないため、.doc で CompatibilityMode が非2013→2013 へ遷移したときだけログ（5-1 の「状態 OR ログ」と同型）
                try
                {
                    string fullName = doc.FullName;
                    string ext = Path.GetExtension(fullName).ToLowerInvariant();
                    int compat = (int)doc.CompatibilityMode;
                    const int wdWord2013 = (int)Word.WdCompatibilityMode.wdWord2013;

                    if (!string.Equals(fullName, _p7LastCompatDocFullName, StringComparison.OrdinalIgnoreCase))
                    {
                        _p7LastCompatDocFullName = fullName;
                        _p7LastCompatMode = compat;
                    }
                    else
                    {
                        if (ext == ".doc" && _p7LastCompatMode >= 0 && _p7LastCompatMode != wdWord2013 && compat == wdWord2013)
                            WordEvidenceHelper.LogCommandWithEvidence("UpgradeDocument");
                        _p7LastCompatMode = compat;
                    }
                }
                catch { }

                // 重い判定（文書全体テキスト化・セクション走査・スタイル解析）は毎回実行しない。
                // リボン入力欄（フォントサイズ等）でのフォーカス喪失を避けるため、約6秒ごとに間引く。
                _heavyCheckTickCounter++;
                bool runHeavyChecks = (_heavyCheckTickCounter % 5) == 0;
                if (runHeavyChecks)
                {
                    string borderFp = WordWatermarkInspection.GetPageBorderFingerprint(doc);
                    if (_lastPageBorderFingerprint != null && borderFp != _lastPageBorderFingerprint)
                    {
                        if (_suppressPageBorderPollLogs > 0)
                            _suppressPageBorderPollLogs--;
                        else
                            LogPageBordersForDestructiveGate();
                    }
                    _lastPageBorderFingerprint = borderFp;

                    bool lineSimple = WordWatermarkInspection.IsDocumentStyleSetLineSimple(doc);
                    string styleSetFp = WordWatermarkInspection.GetStyleSetLineFingerprint(doc);
                    if (_lastHeading1LineSimple.HasValue && lineSimple && !_lastHeading1LineSimple.Value)
                    {
                        if (_suppressStyleSetPollLogs > 0)
                            _suppressStyleSetPollLogs--;
                        else
                        {
                            WordEvidenceHelper.LogCommandWithEvidence("StyleSetLineSimple");
                            _styleSetLineSimpleLoggedThisDoc = true;
                            _lastStyleSetLineFingerprint = styleSetFp;
                        }
                    }
                    else if (_lastHeading1LineSimple.HasValue && _lastHeading1LineSimple.Value && !lineSimple)
                    {
                        // シンプル適用後に別スタイルセットへ変更した場合（スタイリッシュ等）のフォールバック
                        if (_suppressStyleSetPollLogs > 0)
                            _suppressStyleSetPollLogs--;
                        else
                            WordEvidenceHelper.LogCommandWithEvidence("StyleSetLineStylish");
                    }
                    else if (_styleSetLineSimpleLoggedThisDoc && lineSimple
                        && !string.IsNullOrEmpty(_lastStyleSetLineFingerprint)
                        && !string.IsNullOrEmpty(styleSetFp)
                        && !string.Equals(styleSetFp, _lastStyleSetLineFingerprint, StringComparison.Ordinal))
                    {
                        // COM 上はシンプルのまま見えても OpenXML スタイル定義が変わった場合
                        if (_suppressStyleSetPollLogs > 0)
                            _suppressStyleSetPollLogs--;
                        else
                            WordEvidenceHelper.LogCommandWithEvidence("StyleSetLineStylish");
                    }
                    _lastHeading1LineSimple = lineSimple;
                    if (!string.IsNullOrEmpty(styleSetFp))
                        _lastStyleSetLineFingerprint = styleSetFp;

                    bool lineStylish = WordWatermarkInspection.IsDocumentStyleSetLineStylish(doc);
                    if (_lastLineStylish.HasValue && lineStylish && !_lastLineStylish.Value)
                    {
                        if (_suppressStyleSetPollLogs > 0)
                            _suppressStyleSetPollLogs--;
                        else
                            WordEvidenceHelper.LogCommandWithEvidence("StyleSetLineStylish");
                    }
                    _lastLineStylish = lineStylish;

                    // 2-3: リボンの色ギャラリーを直接フックできないため、対象文字の色状態変化だけで補完ログを出す
                    string colorFp = GetTask1_2_03ColorFingerprint(doc);
                    if (!string.IsNullOrEmpty(colorFp))
                    {
                        if (_lastTask1_2_03ColorFingerprint != null &&
                            colorFp != _lastTask1_2_03ColorFingerprint)
                        {
                            Logger.LogCommand("FontColorPicker");
                        }
                        _lastTask1_2_03ColorFingerprint = colorFp;
                    }

                    int columnBreakCount = CountColumnBreaks(doc);
                    if (columnBreakCount > _lastColumnBreakCount)
                    {
                        Logger.LogCommand("ColumnBreak");
                    }
                    _lastColumnBreakCount = columnBreakCount;

                    // 4-1 等: 透かし指紋の変化で [Op] Watermark（4-5 は Sample2 / Draft1Diagonal で Executed も）
                    try
                    {
                        string norm = WordWatermarkInspection.NormalizeXml(doc.WordOpenXML);
                        string wmFp = WordWatermarkInspection.GetWatermarkFingerprint(norm);
                        if (_lastWatermarkFingerprint != null
                            && !string.Equals(wmFp, _lastWatermarkFingerprint, StringComparison.Ordinal))
                        {
                            // 4-5 正答（Sample2）はタスク切替直後の suppress でも必ず証跡を残す
                            bool isSample2Change = string.Equals(wmFp, "Sample2", StringComparison.Ordinal);
                            if (isSample2Change)
                                LogWatermarkForDestructiveGate(wmFp);
                            else if (_suppressWatermarkPollLogs > 0)
                                _suppressWatermarkPollLogs--;
                            else
                                LogWatermarkForDestructiveGate(wmFp);
                        }
                        _lastWatermarkFingerprint = wmFp;
                        _lastDraft1WatermarkFound = string.Equals(wmFp, "Draft1Diagonal", StringComparison.Ordinal);
                    }
                    catch { }

                    // 5-1, 5-2: 画像レイアウトの検知（5月21日...段落付近）
                    try
                    {
                        Word.Range searchRange = doc.Content.Duplicate;
                        Word.Find find = searchRange.Find;
                        find.ClearFormatting();
                        find.Text = "5月21日より5日間の";
                        find.Format = false;
                        find.Replacement.Text = "";
                        find.Wrap = Word.WdFindWrap.wdFindStop;
                        if (find.Execute())
                        {
                            Word.Range paraRange = searchRange.Paragraphs[1].Range;
                            int paraStart = paraRange.Start;
                            int paraEnd = paraRange.End;

                            int currentWrapType = -1; // -1: なし, 0: 行内, 1: 四角形, 2: 狭く, 3: 上下, 4: その他

                            // 行内画像チェック
                            if (paraRange.InlineShapes.Count > 0)
                            {
                                currentWrapType = 0; // Inline
                            }
                            else
                            {
                                // 浮動画像（Shape）チェック
                                foreach (Word.Shape sh in doc.Shapes)
                                {
                                    try
                                    {
                                        int anchor = sh.Anchor != null ? sh.Anchor.Start : -1;
                                        if (anchor >= paraStart && anchor <= paraEnd)
                                        {
                                            if (sh.WrapFormat.Type == Word.WdWrapType.wdWrapSquare) currentWrapType = 1;
                                            else if (sh.WrapFormat.Type == Word.WdWrapType.wdWrapTight) currentWrapType = 2;
                                            else if (sh.WrapFormat.Type == Word.WdWrapType.wdWrapTopBottom) currentWrapType = 3;
                                            else currentWrapType = 4;
                                            break;
                                        }
                                    }
                                    finally { Marshal.ReleaseComObject(sh); }
                                }
                            }

                            if (_lastLaptopWrapType != currentWrapType)
                            {
                                if (currentWrapType == 3) WordEvidenceHelper.LogCommandWithEvidence("WrapTopBottom");
                                else if (currentWrapType == 2) WordEvidenceHelper.LogCommandWithEvidence("WrapTight");
                                else if (currentWrapType == 1) WordEvidenceHelper.LogCommandWithEvidence("WrapSquare");
                            }
                            _lastLaptopWrapType = currentWrapType;
                            
                            Marshal.ReleaseComObject(paraRange);
                        }
                        Marshal.ReleaseComObject(find);
                        Marshal.ReleaseComObject(searchRange);
                    }
                    catch { }

                    // 7-3: インテグラル相当ヘッダーが false→true に遷移したとき IntegralHeader をログ（7-4 後の再採点用）。判定は WordChecker1_7 と同一。
                    try
                    {
                        bool nowIntegral = EvaluateIntegralHeaderPresenceForPolling(doc);
                        string iFull;
                        try { iFull = doc.FullName; } catch { iFull = null; }
                        if (!string.IsNullOrEmpty(iFull))
                        {
                            if (!string.Equals(iFull, _p7IntegralTrackedFullName, StringComparison.OrdinalIgnoreCase))
                            {
                                _p7IntegralTrackedFullName = iFull;
                                _p7IntegralLastDetected = nowIntegral;
                            }
                            else
                            {
                                if (!_p7IntegralLastDetected && nowIntegral)
                                    WordEvidenceHelper.LogCommandWithEvidence("IntegralHeader");
                                _p7IntegralLastDetected = nowIntegral;
                            }
                        }
                    }
                    catch { }
                }
        }

        /// <summary>7-2: Project7.doc 上で Company が目標値へ遷移したとき SetDocumentCompany をログする。</summary>
        private void UpdateP7CompanyPolling(Word.Document doc)
        {
            if (doc == null) return;
            string fullName;
            try { fullName = doc.FullName; }
            catch { return; }
            if (!IsProject7DocumentPath(fullName)) return;

            string company = TryGetDocumentCompany(doc) ?? "";
            bool nowMatched = string.Equals(company, P7CompanyTarget, StringComparison.Ordinal);

            if (!string.Equals(fullName, _p7LastCompanyDocFullName, StringComparison.OrdinalIgnoreCase))
            {
                _p7LastCompanyDocFullName = fullName;
                _p7LastCompanyMatched = nowMatched;
            }
            else
            {
                if (!_p7LastCompanyMatched && nowMatched)
                    WordEvidenceHelper.LogCommandWithEvidence("SetDocumentCompany");
                _p7LastCompanyMatched = nowMatched;
            }
        }

        private static bool IsProject7DocumentPath(string fullName)
        {
            if (string.IsNullOrEmpty(fullName)) return false;
            string name = Path.GetFileNameWithoutExtension(fullName);
            return System.Text.RegularExpressions.Regex.IsMatch(name, @"^project\s*7$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        private static string TryGetDocumentCompany(Word.Document doc)
        {
            try
            {
                dynamic companyProp = ((dynamic)doc.BuiltInDocumentProperties)["Company"];
                return companyProp?.Value?.ToString() ?? "";
            }
            catch { return null; }
        }

        private static bool IsRdTxtSavePath(string fullName)
        {
            if (string.IsNullOrEmpty(fullName)) return false;
            return string.Equals(Path.GetFileName(fullName), P7RdTxtFileName, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRdDocmSavePath(string fullName)
        {
            if (string.IsNullOrEmpty(fullName)) return false;
            return string.Equals(Path.GetFileName(fullName), P7RdDocmFileName, StringComparison.OrdinalIgnoreCase);
        }

        private static void LogFileSaveAsCommandForPath(string fullName)
        {
            if (IsRdTxtSavePath(fullName))
                WordEvidenceHelper.LogCommandWithEvidence("FileSaveAsTxt");
            else if (IsRdDocmSavePath(fullName))
                WordEvidenceHelper.LogCommandWithEvidence("FileSaveAsDocm");
        }

        /// <summary>
        /// 7-4/7-5: ActiveDocument が朗読会.txt / 朗読会.docm へ遷移したときそれぞれ専用ログを付与する（毎ティックで呼ぶ）。
        /// </summary>
        private void UpdateFileSaveAsPolling(Word.Document doc)
        {
            if (doc == null) return;
            string fullName;
            try { fullName = doc.FullName; }
            catch { return; }
            if (string.IsNullOrEmpty(fullName)) return;

            bool nowTxt = IsRdTxtSavePath(fullName);
            bool nowDocm = IsRdDocmSavePath(fullName);

            if (!string.Equals(fullName, _p7FileSaveAsTrackedFullName, StringComparison.OrdinalIgnoreCase))
            {
                bool wasTxt = false;
                bool wasDocm = false;
                if (!string.IsNullOrEmpty(_p7FileSaveAsTrackedFullName))
                {
                    wasTxt = IsRdTxtSavePath(_p7FileSaveAsTrackedFullName);
                    wasDocm = IsRdDocmSavePath(_p7FileSaveAsTrackedFullName);
                }

                if (!string.IsNullOrEmpty(_p7FileSaveAsTrackedFullName) && !wasTxt && nowTxt)
                    WordEvidenceHelper.LogCommandWithEvidence("FileSaveAsTxt");
                if (!string.IsNullOrEmpty(_p7FileSaveAsTrackedFullName) && !wasDocm && nowDocm)
                    WordEvidenceHelper.LogCommandWithEvidence("FileSaveAsDocm");

                _p7FileSaveAsTrackedFullName = fullName;
                _p7FileSaveAsLastTxt = nowTxt;
                _p7FileSaveAsLastDocm = nowDocm;
            }
            else
            {
                if (!_p7FileSaveAsLastTxt && nowTxt)
                    WordEvidenceHelper.LogCommandWithEvidence("FileSaveAsTxt");
                if (!_p7FileSaveAsLastDocm && nowDocm)
                    WordEvidenceHelper.LogCommandWithEvidence("FileSaveAsDocm");
                _p7FileSaveAsLastTxt = nowTxt;
                _p7FileSaveAsLastDocm = nowDocm;
            }
        }


        /// <summary>
        /// ポーリング中の COM アクセスで選択網掛けが消えないよう、選択位置を保存してから ScreenUpdating を止めて実行する。
        /// </summary>
        private static void PreserveSelectionDuring(Word.Application app, Action action)
        {
            if (action == null)
                return;
            if (app == null)
            {
                action();
                return;
            }

            int selStart = -1;
            int selEnd = -1;
            bool screenUpdatingWasEnabled = true;
            try
            {
                try
                {
                    Word.Selection sel = app.Selection;
                    if (sel != null)
                    {
                        selStart = sel.Start;
                        selEnd = sel.End;
                    }
                }
                catch { /* ignore */ }

                try
                {
                    screenUpdatingWasEnabled = app.ScreenUpdating;
                    app.ScreenUpdating = false;
                }
                catch { /* ignore */ }

                action();
            }
            finally
            {
                int afterStart = -1;
                int afterEnd = -1;
                try
                {
                    Word.Selection afterSel = app.Selection;
                    if (afterSel != null)
                    {
                        afterStart = afterSel.Start;
                        afterEnd = afterSel.End;
                    }
                }
                catch { /* ignore */ }

                bool selectionChangedByPoll = afterStart >= 0 && (afterStart != selStart || afterEnd != selEnd);
                bool userExpandedSelection = selStart >= 0 && selStart == selEnd && afterEnd > afterStart;
                try
                {
                    if (selStart >= 0 && selEnd >= 0
                        && selectionChangedByPoll
                        && !userExpandedSelection)
                    {
                        app.Selection.SetRange(selStart, selEnd);
                    }
                }
                catch { /* ignore */ }

                try
                {
                    app.ScreenUpdating = screenUpdatingWasEnabled;
                }
                catch { /* ignore */ }
            }
        }

        /// <summary>
        /// ナビペイン・検索ダイアログ・コメントペイン等、編集ペイン外操作中は侵入的な文書 COM を遅延する。
        /// </summary>
        private static bool ShouldDeferIntrusiveDocumentCom(Word.Application app)
        {
            return ShouldSkipDocumentComBecauseFocusNotInEditingPane();
        }

        private static bool IsProject4CommentTaskActive()
        {
            try
            {
                string path = Path.Combine(Path.GetTempPath(), "mos_word_current_task.txt");
                if (!File.Exists(path))
                    return false;
                var parts = File.ReadAllText(path).Trim().Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 2)
                    return false;
                if (!int.TryParse(parts[0].Trim(), out int projectId) || !int.TryParse(parts[1].Trim(), out int taskId))
                    return false;
                return projectId == 4 && taskId >= 1 && taskId <= 3;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// キーボードフォーカスが文書編集ペイン（_WwG）上にないとき true。
        /// リボンの数値欄・代替テキスト等の入力中は COM を触らない。
        /// GetFocus が取れない場合は false（従来どおり実行）。
        /// </summary>
        private static bool ShouldSkipDocumentComBecauseFocusNotInEditingPane()
        {
            IntPtr focus = NativeMethods.GetFocus();
            if (focus == IntPtr.Zero)
                return false;
            return !HasWindowOrAncestorClass(focus, WordDocumentPaneClassName);
        }

        private static bool HasWindowOrAncestorClass(IntPtr start, string className)
        {
            IntPtr h = start;
            for (int i = 0; i < 40 && h != IntPtr.Zero; i++)
            {
                if (string.Equals(GetWindowClassName(h), className, StringComparison.Ordinal))
                    return true;
                h = NativeMethods.GetParent(h);
            }
            return false;
        }

        private static string GetWindowClassName(IntPtr hWnd)
        {
            var sb = new StringBuilder(256);
            if (NativeMethods.GetClassName(hWnd, sb, sb.Capacity) == 0)
                return string.Empty;
            return sb.ToString();
        }

        private static class NativeMethods
        {
            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            internal static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

            [DllImport("user32.dll")]
            internal static extern IntPtr GetParent(IntPtr hWnd);

            [DllImport("user32.dll")]
            internal static extern IntPtr GetFocus();
        }

        private static int CountColumnBreaks(Word.Document doc)
        {
            Word.Range content = null;
            try
            {
                if (doc?.Content == null)
                    return 0;
                content = doc.Content.Duplicate;
                string text = content.Text ?? string.Empty;
                return text.Count(c => c == (char)14);
            }
            catch
            {
                return 0;
            }
            finally
            {
                if (content != null)
                {
                    try { Marshal.ReleaseComObject(content); } catch { }
                }
            }
        }

        private static string _ecoAnchorDocKeyPoll;
        private static int _ecoAnchorStartPoll = -1;
        private static int _ecoAnchorEndPoll = -1;

        private static bool TryGetEcoPhraseAnchorPoll(Word.Document doc, out int start, out int end)
        {
            start = end = -1;
            if (doc == null)
                return false;
            string key = "";
            try { key = doc.FullName ?? ""; } catch { }
            if (!string.IsNullOrEmpty(key) && key == _ecoAnchorDocKeyPoll && _ecoAnchorStartPoll >= 0)
            {
                start = _ecoAnchorStartPoll;
                end = _ecoAnchorEndPoll;
                return true;
            }
            Word.Range searchRange = null;
            Word.Find find = null;
            try
            {
                if (doc.Content == null)
                    return false;
                searchRange = doc.Content.Duplicate;
                find = searchRange.Find;
                find.ClearFormatting();
                find.Text = "エコと節約";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.Replacement.Text = "";
                if (!find.Execute())
                    return false;
                start = searchRange.Start;
                end = searchRange.End;
                _ecoAnchorDocKeyPoll = key;
                _ecoAnchorStartPoll = start;
                _ecoAnchorEndPoll = end;
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (find != null)
                {
                    try { Marshal.ReleaseComObject(find); } catch { }
                }
                if (searchRange != null)
                {
                    try { Marshal.ReleaseComObject(searchRange); } catch { }
                }
            }
        }

        private static bool CommentAnchorOverlapsEcoPoll(Word.Comment c, Word.Document doc)
        {
            if (!TryGetEcoPhraseAnchorPoll(doc, out int ecoStart, out int ecoEnd))
                return false;
            Word.Range scope = null;
            try
            {
                scope = c.Scope;
                if (scope == null)
                    return false;
                return scope.Start <= ecoEnd && scope.End >= ecoStart;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (scope != null)
                {
                    try { Marshal.ReleaseComObject(scope); } catch { }
                }
            }
        }

        /// <summary>エコ関連コメント総数（解決済み含む）。</summary>
        private static int CountEcoRelatedComments(Word.Document doc)
        {
            if (doc?.Comments == null || doc.Comments.Count == 0)
                return 0;
            Word.Comments comments = doc.Comments;
            int count = 0;
            try
            {
                foreach (Word.Comment c in comments)
                {
                    try
                    {
                        if (CommentRelatesToEcoPhrasePoll(c, doc))
                            count++;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(c);
                    }
                }
                return count;
            }
            finally
            {
                Marshal.ReleaseComObject(comments);
            }
        }

        /// <summary>吹き出し（返信含む）に「エコと節約」があり、かつ未解決（Done でない）コメントの件数。</summary>
        private static int CountUnresolvedEcoComments(Word.Document doc)
        {
            if (doc?.Comments == null || doc.Comments.Count == 0)
                return 0;
            Word.Comments comments = doc.Comments;
            int count = 0;
            try
            {
                foreach (Word.Comment c in comments)
                {
                    try
                    {
                        if (!CommentRelatesToEcoPhrasePoll(c, doc))
                            continue;
                        bool done = false;
                        try { done = c.Done; } catch { }
                        if (!done)
                            count++;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(c);
                    }
                }
                return count;
            }
            finally
            {
                Marshal.ReleaseComObject(comments);
            }
        }

        /// <summary>採点 CheckTask_1_4_03 と同様、本文に「エコと節約」があるか。コメント吹き出しが Range で読めない環境でも削除ログ補助に使う。</summary>
        private static bool DocumentBodyContainsEcoPhraseForTask4_3(Word.Document doc)
        {
            Word.Range searchRange = null;
            Word.Find find = null;
            try
            {
                if (doc?.Content == null)
                    return false;
                searchRange = doc.Content.Duplicate;
                find = searchRange.Find;
                find.ClearFormatting();
                find.Text = "エコと節約";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.Replacement.Text = "";
                return find.Execute();
            }
            catch
            {
                return false;
            }
            finally
            {
                if (find != null)
                {
                    try { Marshal.ReleaseComObject(find); } catch { }
                }
                if (searchRange != null)
                {
                    try { Marshal.ReleaseComObject(searchRange); } catch { }
                }
            }
        }

        /// <summary>「エコと節約」コメントが解決済み（Done）で文書に残っているか。削除との区別に使う。</summary>
        private static bool DocumentHasResolvedEcoComment(Word.Document doc)
        {
            if (doc?.Comments == null || doc.Comments.Count == 0)
                return false;
            Word.Comments comments = doc.Comments;
            try
            {
                foreach (Word.Comment c in comments)
                {
                    try
                    {
                        if (!CommentRelatesToEcoPhrasePoll(c, doc))
                            continue;
                        bool done = false;
                        try { done = c.Done; } catch { }
                        if (done)
                            return true;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(c);
                    }
                }
                return false;
            }
            finally
            {
                Marshal.ReleaseComObject(comments);
            }
        }

        /// <summary>いずれかのコメント吹き出し（返信含む）に「エコと節約」があるか。採点側の DocumentHasAnyEcoCommentBalloon と同趣旨。</summary>
        private static bool DocumentHasAnyEcoCommentBalloonPoll(Word.Document doc)
        {
            if (doc?.Comments == null || doc.Comments.Count == 0)
                return false;
            Word.Comments comments = doc.Comments;
            try
            {
                foreach (Word.Comment c in comments)
                {
                    try
                    {
                        if (CommentRelatesToEcoPhrasePoll(c, doc))
                            return true;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(c);
                    }
                }
                return false;
            }
            finally
            {
                Marshal.ReleaseComObject(comments);
            }
        }

        /// <summary>WordChecker1_4 の NormalizeCommentBody と同じ（エコ検出の食い違いを防ぐ）。</summary>
        private static string NormalizeCommentBodyForEcoPoll(string s)
        {
            if (string.IsNullOrEmpty(s))
                return "";
            string t = s.Trim();
            t = t.Replace("\r\n", "").Replace("\r", "").Replace("\n", "").Replace("\u3000", " ");
            while (t.IndexOf("  ", StringComparison.Ordinal) >= 0)
                t = t.Replace("  ", " ");
            return t.Trim();
        }

        /// <summary>WordChecker1_4 の NormalizedContainsEco と同じ。</summary>
        private static bool NormalizedContainsEcoPoll(string normalized)
        {
            const string needle = "エコと節約";
            if (string.IsNullOrEmpty(normalized))
                return false;
            if (normalized.Contains(needle))
                return true;
            try
            {
                return normalized.Normalize(NormalizationForm.FormKC).Contains(needle);
            }
            catch
            {
                return false;
            }
        }

        private static bool CommentRelatesToEcoPhrasePoll(Word.Comment c, Word.Document doc)
        {
            if (CommentHasEcoTextInBalloon(c))
                return true;
            string scope = "";
            try { scope = c.Scope?.Text ?? ""; } catch { }
            if (NormalizedContainsEcoPoll(NormalizeCommentBodyForEcoPoll(scope)))
                return true;
            return CommentAnchorOverlapsEcoPoll(c, doc);
        }

        private static bool CommentHasEcoTextInBalloon(Word.Comment c)
        {
            string raw = "";
            try { raw = c.Range?.Text ?? ""; } catch { }
            if (NormalizedContainsEcoPoll(NormalizeCommentBodyForEcoPoll(raw)))
                return true;
            try
            {
                if (c.Replies == null || c.Replies.Count == 0)
                    return false;
                int n = c.Replies.Count;
                for (int i = 1; i <= n; i++)
                {
                    Word.Comment r = null;
                    try
                    {
                        r = c.Replies[i];
                        string rt = r.Range?.Text ?? "";
                        if (NormalizedContainsEcoPoll(NormalizeCommentBodyForEcoPoll(rt)))
                            return true;
                    }
                    finally
                    {
                        if (r != null)
                            Marshal.ReleaseComObject(r);
                    }
                }
            }
            catch { }
            return false;
        }

        /// <summary>2-3 の採点対象テキストの色状態をフィンガープリント化する。</summary>
        private static string GetTask1_2_03ColorFingerprint(Word.Document doc)
        {
            Word.Range searchRange = null;
            Word.Find find = null;
            Word.Range found = null;
            Word.Font font = null;
            try
            {
                searchRange = doc.Content.Duplicate;
                find = searchRange.Find;
                find.ClearFormatting();
                find.Text = "朗読を楽しみましょう！";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                bool ok = find.Execute();
                if (!ok)
                    return string.Empty;

                found = searchRange.Duplicate;
                font = found.Characters[1].Font;

                int theme = -1;
                float shade = 0f;
                int rgb = 0;
                try { theme = (int)font.TextColor.ObjectThemeColor; } catch { }
                try
                {
                    object tintVal = font.TextColor.TintAndShade;
                    if (tintVal is float f) shade = f;
                    else if (tintVal is double d) shade = (float)d;
                    else if (tintVal != null) shade = Convert.ToSingle(tintVal);
                }
                catch { }
                try { rgb = (int)font.TextColor.RGB; } catch { }

                return string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    "{0}|{1}|{2}",
                    theme, shade, rgb);
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                if (font != null) Marshal.ReleaseComObject(font);
                if (found != null) Marshal.ReleaseComObject(found);
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
            }
        }

        /// <summary>2-3 の期待色（アクセント1 + 25%暗く + 青系）に一致するかを判定する。</summary>
        private static bool IsTask1_2_03ExpectedColorFingerprint(string fp)
        {
            if (string.IsNullOrEmpty(fp))
                return false;
            string[] parts = fp.Split('|');
            if (parts.Length != 3)
                return false;
            if (!int.TryParse(parts[0], out int theme))
                return false;
            if (!float.TryParse(parts[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out float shade))
                return false;
            if (!int.TryParse(parts[2], out int rgb))
                return false;

            bool isAccent1 = theme == (int)Word.WdThemeColorIndex.wdThemeColorAccent1;
            if (shade <= -0.35f)
                return false;
            bool isDarker25 = shade >= -0.31f && shade <= -0.19f;

            int r = rgb & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;
            bool isBlue = b >= r && b >= g && b > 0;

            if (isDarker25)
                return isAccent1 && isBlue;

            if (!isAccent1 || !isBlue)
                return false;

            // TintAndShade が 0 の環境: 解決 RGB がアクセント1基準より約20〜30%暗いか（50%は拒否）
            float textLum = (0.299f * r + 0.587f * g + 0.114f * b) / 255f;
            const float accent1BaseLum = 0.45f;
            float ratio = textLum / accent1BaseLum;
            if (ratio < 0.60f)
                return false;
            return ratio >= 0.70f && ratio <= 0.80f;
        }

        /// <summary>3-1: 先頭セクションの余白が「やや狭い」プリセット相当か（WordChecker1_3 と同じ許容誤差）。</summary>
        private static bool IsMarginsModeratePreset(Word.Document doc)
        {
            Word.Section section = null;
            Word.PageSetup ps = null;
            try
            {
                section = doc.Sections[1];
                ps = section.PageSetup;
                float top = ps.TopMargin;
                float bottom = ps.BottomMargin;
                float left = ps.LeftMargin;
                float right = ps.RightMargin;

                bool IsApprox(float value, float target) => Math.Abs(value - target) <= 1.5f;

                return IsApprox(top, 72.0f) &&
                       IsApprox(bottom, 72.0f) &&
                       IsApprox(left, 54.0f) &&
                       IsApprox(right, 54.0f);
            }
            catch
            {
                return false;
            }
            finally
            {
                if (ps != null) Marshal.ReleaseComObject(ps);
                if (section != null) Marshal.ReleaseComObject(section);
            }
        }

        /// <summary>文書内の全セクションの印刷の向きを連結した文字列（いずれかのセクションの向き変更で変化する）。</summary>
        private static string GetAllSectionsOrientationFingerprint(Word.Document doc)
        {
            var sb = new StringBuilder();
            foreach (Word.Section sec in doc.Sections)
            {
                try
                {
                    sb.Append((int)sec.PageSetup.Orientation).Append(',');
                }
                finally
                {
                    Marshal.ReleaseComObject(sec);
                }
            }
            return sb.ToString();
        }

        private static bool TryGetOptionsShowAll(Word.Application app)
        {
            try
            {
                if (app?.Options == null)
                    return false;
                var prop = app.Options.GetType().GetProperty("ShowAll");
                if (prop == null)
                    return false;
                object value = prop.GetValue(app.Options, null);
                return value is bool b && b;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>7-3: インテグラル相当ヘッダー（WordChecker1_7 と同ロジック。VSTO ポーリング用）。</summary>
        private static bool EvaluateIntegralHeaderPresenceForPolling(Word.Document document)
        {
            if (document == null) return false;
            bool usePostCompatFingerprint = false;
            try
            {
                string ext = Path.GetExtension(document.FullName).ToLowerInvariant();
                usePostCompatFingerprint = ext == ".doc";
            }
            catch { /* ignore */ }

            try
            {
                int sectionCount = document.Sections.Count;
                for (int i = 1; i <= sectionCount; i++)
                {
                    foreach (Word.WdHeaderFooterIndex hfType in new[]
                    {
                        Word.WdHeaderFooterIndex.wdHeaderFooterPrimary,
                        Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                        Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages
                    })
                    {
                        Word.HeaderFooter header = null;
                        try
                        {
                            header = document.Sections[i].Headers[hfType];
                            if (!header.Exists) continue;
                            string headerXml = header.Range.WordOpenXML ?? "";
                            if (HeaderXmlLooksLikeIntegralForPolling(headerXml, usePostCompatFingerprint))
                                return true;
                        }
                        finally
                        {
                            if (header != null) Marshal.ReleaseComObject(header);
                        }
                    }
                }
            }
            catch { /* ignore */ }
            return false;
        }

        private static bool HeaderXmlLooksLikeIntegralForPolling(string xml, bool usePostCompatFingerprint)
        {
            if (string.IsNullOrEmpty(xml)) return false;
            if (ContainsIntegralBuildingBlockMetadataForPolling(xml)) return true;
            if (HasIntegralStructureFingerprintForPolling(xml)) return true;
            if (usePostCompatFingerprint && HasIntegralStructureFingerprintAfterCompatForPolling(xml)) return true;
            return false;
        }

        private static bool ContainsIntegralBuildingBlockMetadataForPolling(string xml)
        {
            if (xml.IndexOf("Integral", StringComparison.OrdinalIgnoreCase) < 0) return false;
            if (Regex.IsMatch(xml, @"w:val\s*=\s*""Integral""", RegexOptions.IgnoreCase)) return true;
            if (xml.IndexOf("docPart", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (xml.IndexOf("w:sdt", StringComparison.Ordinal) >= 0) return true;
            return false;
        }

        private static bool HasIntegralAccent2ColorMarkerForPolling(string xml)
        {
            if (string.IsNullOrEmpty(xml)) return false;
            return xml.IndexOf("fill=\"E97132\"", StringComparison.Ordinal) >= 0
                || xml.IndexOf("fill=\"ED7D31\"", StringComparison.Ordinal) >= 0
                || xml.IndexOf("w:themeFill=\"accent2\"", StringComparison.OrdinalIgnoreCase) >= 0
                || xml.IndexOf("w:themeColor=\"accent2\"", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool HasIntegralStructureFingerprintForPolling(string xml)
        {
            if (!HasIntegralAccent2ColorMarkerForPolling(xml)) return false;
            if (xml.IndexOf("w:w=\"1782\"", StringComparison.Ordinal) < 0) return false;
            if (xml.IndexOf("w:w=\"7286\"", StringComparison.Ordinal) < 0) return false;
            return true;
        }

        private static bool HasIntegralStructureFingerprintAfterCompatForPolling(string xml)
        {
            if (xml.IndexOf("<w:tbl", StringComparison.OrdinalIgnoreCase) < 0) return false;
            if (!HasIntegralAccent2ColorMarkerForPolling(xml)) return false;
            int gridColCount = 0;
            for (int i = 0; ;)
            {
                int p = xml.IndexOf("<w:gridCol", i, StringComparison.Ordinal);
                if (p < 0) break;
                gridColCount++;
                i = p + 10;
            }
            return gridColCount >= 2;
        }

        /// <summary>mos_word_current_task.txt を監視し、スナップショット取得・タスク切替時の差分記録を行う。</summary>
        private sealed class WordDestructiveMonitor
        {
            private static readonly string CurrentTaskFilePath = Path.Combine(Path.GetTempPath(), "mos_word_current_task.txt");
            private static readonly string SnapshotFilePath = Path.Combine(Path.GetTempPath(), "mos_word_snapshot.txt");
            private static readonly string DestructiveLogPath = Path.Combine(Path.GetTempPath(), "mos_word_destructive_errors.log");

            private readonly ThisAddIn _addIn;
            private Timer _pollTimer;
            private int _projectId = -1;
            private int _taskId = -1;
            private int _attemptNo;
            private int _exemptFlags;

            public WordDestructiveMonitor(ThisAddIn addIn)
            {
                _addIn = addIn;
            }

            public void Start()
            {
                _pollTimer = new Timer { Interval = 500 };
                _pollTimer.Tick += PollTimer_Tick;
                _pollTimer.Start();
            }

            public void Stop()
            {
                _pollTimer?.Stop();
                _pollTimer?.Dispose();
                _pollTimer = null;
            }

            private void PollTimer_Tick(object sender, EventArgs e)
            {
                try
                {
                    if (!File.Exists(CurrentTaskFilePath))
                    {
                        _projectId = -1;
                        _taskId = -1;
                        return;
                    }

                    string line = File.ReadAllText(CurrentTaskFilePath).Trim();
                    if (string.IsNullOrEmpty(line))
                        return;

                    var parts = line.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length < 2)
                        return;
                    if (!int.TryParse(parts[0].Trim(), out int projectId) || !int.TryParse(parts[1].Trim(), out int taskId))
                        return;

                    int exemptFlags = 0;
                    if (parts.Length >= 3)
                        int.TryParse(parts[2].Trim(), out exemptFlags);
                    int attemptNo = 0;
                    if (parts.Length >= 4)
                        int.TryParse(parts[3].Trim(), out attemptNo);

                    bool forceSnapshot = !File.Exists(SnapshotFilePath);
                    bool taskChanged = projectId != _projectId || taskId != _taskId || attemptNo != _attemptNo;

                    if (!taskChanged && !forceSnapshot)
                        return;

                    if (taskChanged)
                        _addIn.EnsureTask45WatermarkEvidenceBeforeLeave(_projectId, _taskId, projectId, taskId);

                    if (_projectId >= 0 && _taskId >= 0 && !forceSnapshot && projectId == _projectId)
                        CompareAndLogDestructive(_projectId, _taskId, _attemptNo, _exemptFlags);

                    _projectId = projectId;
                    _taskId = taskId;
                    _attemptNo = attemptNo;
                    _exemptFlags = exemptFlags;

                    Logger.SetCurrentTaskContext(projectId, taskId, attemptNo);
                    _addIn.SyncWatermarkPollingBaselineOnTaskSwitch(projectId);
                    _addIn.SyncEcoCommentBaselineOnTaskSwitch(projectId, taskId);
                    TakeSnapshot(projectId, taskId, attemptNo);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[WordDestructiveMonitor] " + ex.Message);
                }
            }

            private void TakeSnapshot(int projectId, int taskId, int attemptNo)
            {
                Word.Document doc = _addIn.TryGetProjectDocument(projectId);
                if (doc == null)
                    return;

                try
                {
                    var snap = Capture(doc, projectId, taskId, attemptNo);
                    SaveSnapshot(snap);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[WordDestructiveMonitor] TakeSnapshot: " + ex.Message);
                }
            }

            private void CompareAndLogDestructive(int projectId, int taskId, int attemptNo, int exemptFlagsInt)
            {
                var baseline = LoadSnapshot();
                if (baseline == null || baseline.ProjectId != projectId || baseline.TaskId != taskId || baseline.AttemptNo != attemptNo)
                    return;

                Word.Document doc = _addIn.TryGetProjectDocument(projectId);
                if (doc == null)
                    return;

                var current = Capture(doc, projectId, taskId, attemptNo);
                var errors = Compare(baseline, current, exemptFlagsInt);
                if (errors.Count == 0)
                    return;

                try
                {
                    string key = $"{projectId},{taskId},{attemptNo}:";
                    string body = string.Join(" | ", errors);
                    File.AppendAllText(DestructiveLogPath, key + body + Environment.NewLine, Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[WordDestructiveMonitor] log: " + ex.Message);
                }
            }

            private static List<string> Compare(SnapshotData baseline, SnapshotData current, int exemptFlagsInt)
            {
                var errors = new List<string>();
                if (!HasFlag(exemptFlagsInt, 1) && current.Sections != baseline.Sections)
                    errors.Add($"SectionsCount changed {baseline.Sections}->{current.Sections}");
                if (!HasFlag(exemptFlagsInt, 2) && current.BodyTextLength != baseline.BodyTextLength)
                    errors.Add($"BodyTextLength changed {baseline.BodyTextLength}->{current.BodyTextLength}");
                if (!HasFlag(exemptFlagsInt, 4) && current.InlineShapes != baseline.InlineShapes)
                    errors.Add($"InlineShapesCount changed {baseline.InlineShapes}->{current.InlineShapes}");
                if (!HasFlag(exemptFlagsInt, 8) && current.FloatingShapes != baseline.FloatingShapes)
                    errors.Add($"FloatingShapesCount changed {baseline.FloatingShapes}->{current.FloatingShapes}");
                if (!HasFlag(exemptFlagsInt, 16) && current.Tables != baseline.Tables)
                    errors.Add($"TablesCount changed {baseline.Tables}->{current.Tables}");
                if (!HasFlag(exemptFlagsInt, 32) && current.Comments != baseline.Comments)
                    errors.Add($"CommentsCount changed {baseline.Comments}->{current.Comments}");
                if (!HasFlag(exemptFlagsInt, 64) && !string.Equals(baseline.HeaderPrimaryFp ?? "", current.HeaderPrimaryFp ?? "", StringComparison.Ordinal))
                    errors.Add("HeaderFooterFingerprint changed");
                if (!HasFlag(exemptFlagsInt, 1024) && baseline.CompatibilityMode >= 0 && current.CompatibilityMode >= 0
                    && baseline.CompatibilityMode != current.CompatibilityMode)
                    errors.Add($"CompatibilityMode changed {baseline.CompatibilityMode}->{current.CompatibilityMode}");
                if (!HasFlag(exemptFlagsInt, 128) && !string.Equals(baseline.PageBorderFingerprint ?? "", current.PageBorderFingerprint ?? "", StringComparison.Ordinal))
                    errors.Add($"PageBorderFingerprint changed {baseline.PageBorderFingerprint}->{current.PageBorderFingerprint}");
                if (!HasFlag(exemptFlagsInt, 256) && !string.Equals(baseline.WatermarkFingerprint ?? "None", current.WatermarkFingerprint ?? "None", StringComparison.Ordinal))
                    errors.Add($"WatermarkFingerprint changed {baseline.WatermarkFingerprint}->{current.WatermarkFingerprint}");
                if (baseline.ProjectId == 8 && baseline.TaskId == 3
                    && current.FootnoteReferenceCount != baseline.FootnoteReferenceCount)
                    errors.Add($"FootnoteReferenceCount changed {baseline.FootnoteReferenceCount}->{current.FootnoteReferenceCount}");
                return errors;
            }

            private static bool HasFlag(int flags, int bit) => (flags & bit) != 0;

            private static SnapshotData Capture(Word.Document doc, int projectId, int taskId, int attemptNo)
            {
                var d = new SnapshotData
                {
                    ProjectId = projectId,
                    TaskId = taskId,
                    AttemptNo = attemptNo,
                    FullName = doc.FullName ?? ""
                };
                try { d.Sections = doc.Sections.Count; } catch { }
                try { d.BodyTextLength = doc.Content.Text.Length; } catch { }
                try { d.InlineShapes = doc.InlineShapes.Count; } catch { }
                try { d.FloatingShapes = doc.Shapes.Count; } catch { }
                try { d.Comments = doc.Comments.Count; } catch { }
                try { d.Tables = doc.Tables.Count; } catch { }
                try { d.CompatibilityMode = (int)doc.CompatibilityMode; } catch { d.CompatibilityMode = -1; }
                d.HeaderPrimaryFp = GetHeaderFp(doc);
                try
                {
                    string norm = WordWatermarkInspection.NormalizeXml(doc.WordOpenXML);
                    d.WatermarkFingerprint = WordWatermarkInspection.GetWatermarkFingerprint(norm);
                }
                catch
                {
                    d.WatermarkFingerprint = "None";
                }
                d.PageBorderFingerprint = WordWatermarkInspection.GetPageBorderFingerprint(doc);
                try
                {
                    d.FootnoteReferenceCount = WordFindHelper.CountFootnoteReferencesInXml(doc.WordOpenXML);
                }
                catch
                {
                    d.FootnoteReferenceCount = 0;
                }
                return d;
            }

            private static string GetHeaderFp(Word.Document doc)
            {
                try
                {
                    var hdr = doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    string xml = hdr.WordOpenXML ?? "";
                    if (xml.IndexOf("ED7D31", StringComparison.OrdinalIgnoreCase) >= 0) return "ED7D31";
                    if (xml.IndexOf("E97132", StringComparison.OrdinalIgnoreCase) >= 0) return "E97132";
                    if (xml.IndexOf("accent2", StringComparison.OrdinalIgnoreCase) >= 0) return "accent2";
                }
                catch { }
                return "";
            }

            private static void SaveSnapshot(SnapshotData d)
            {
                var sb = new StringBuilder();
                sb.AppendLine("# WordSnapshot v2");
                sb.AppendLine($"ProjectId={d.ProjectId}");
                sb.AppendLine($"TaskId={d.TaskId}");
                sb.AppendLine($"AttemptNo={d.AttemptNo}");
                sb.AppendLine($"FullName={d.FullName}");
                sb.AppendLine($"Sections={d.Sections}");
                sb.AppendLine($"BodyTextLength={d.BodyTextLength}");
                sb.AppendLine($"InlineShapes={d.InlineShapes}");
                sb.AppendLine($"FloatingShapes={d.FloatingShapes}");
                sb.AppendLine($"Tables={d.Tables}");
                sb.AppendLine($"Comments={d.Comments}");
                sb.AppendLine($"HeaderPrimaryFp={d.HeaderPrimaryFp}");
                sb.AppendLine($"CompatibilityMode={d.CompatibilityMode}");
                sb.AppendLine($"WatermarkFingerprint={d.WatermarkFingerprint ?? "None"}");
                sb.AppendLine($"PageBorderFingerprint={d.PageBorderFingerprint ?? ""}");
                sb.AppendLine($"FootnoteReferenceCount={d.FootnoteReferenceCount}");
                File.WriteAllText(SnapshotFilePath, sb.ToString(), Encoding.UTF8);
            }

            private static SnapshotData LoadSnapshot()
            {
                if (!File.Exists(SnapshotFilePath))
                    return null;
                var d = new SnapshotData();
                foreach (string line in File.ReadAllLines(SnapshotFilePath, Encoding.UTF8))
                {
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                        continue;
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string key = line.Substring(0, eq).Trim();
                    string val = line.Substring(eq + 1).Trim();
                    switch (key)
                    {
                        case "ProjectId": int.TryParse(val, out int p); d.ProjectId = p; break;
                        case "TaskId": int.TryParse(val, out int t); d.TaskId = t; break;
                        case "AttemptNo": int.TryParse(val, out int a); d.AttemptNo = a; break;
                        case "Sections": int.TryParse(val, out int s); d.Sections = s; break;
                        case "BodyTextLength": int.TryParse(val, out int bl); d.BodyTextLength = bl; break;
                        case "InlineShapes": int.TryParse(val, out int ins); d.InlineShapes = ins; break;
                        case "FloatingShapes": int.TryParse(val, out int fs); d.FloatingShapes = fs; break;
                        case "Tables": int.TryParse(val, out int tb); d.Tables = tb; break;
                        case "Comments": int.TryParse(val, out int cm); d.Comments = cm; break;
                        case "HeaderPrimaryFp": d.HeaderPrimaryFp = val; break;
                        case "CompatibilityMode": int.TryParse(val, out int c); d.CompatibilityMode = c; break;
                        case "WatermarkFingerprint": d.WatermarkFingerprint = val; break;
                        case "PageBorderFingerprint": d.PageBorderFingerprint = val; break;
                        case "FootnoteReferenceCount": int.TryParse(val, out int fn); d.FootnoteReferenceCount = fn; break;
                    }
                }
                if (string.IsNullOrEmpty(d.WatermarkFingerprint))
                    d.WatermarkFingerprint = "None";
                return d;
            }

            private sealed class SnapshotData
            {
                public int ProjectId;
                public int TaskId;
                public int AttemptNo;
                public string FullName;
                public int Sections;
                public int BodyTextLength;
                public int InlineShapes;
                public int FloatingShapes;
                public int Tables;
                public int Comments;
                public string HeaderPrimaryFp;
                public int CompatibilityMode = -1;
                public string WatermarkFingerprint = "None";
                public string PageBorderFingerprint = "";
                public int FootnoteReferenceCount;
            }
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
