using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace New_MOSWordVSTOAddIn
{
    public partial class ThisAddIn
    {
        /// <summary>Word の文書編集ペインのウィンドウクラス。リボン上の入力欄は通常これを祖先に持たない。</summary>
        private const string WordDocumentPaneClassName = "_WwG";

        private Timer _showAllPollTimer;
        private bool? _lastShowAllState;
        private int _lastColumnBreakCount;
        private string _lastTask1_2_03ColorFingerprint;

        /// <summary>全セクションの向きを連結したフィンガープリント（先頭セクションのみでは 3-3 とチェッカーが不一致になるため）。</summary>
        private string _lastOrientationFingerprint;
        private string _lastPageBorderFingerprint;
        private bool? _lastHeading1LineSimple;
        private bool? _lastWatermarkFound;
        private int _lastLaptopWrapType = -1;

        /// <summary>リボンが既にログした直後のポーリング二重記録を抑止する（約2ティック）。</summary>
        private int _suppressOrientationPollLogs;
        private int _suppressPageBorderPollLogs;
        private int _suppressStyleSetPollLogs;
        private int _heavyCheckTickCounter;

        /// <summary>4-3: 吹き出しに「エコと節約」があり未解決のコメント数。ポーリングで 1→0 になったときログする。</summary>
        private int _lastUnresolvedEcoCommentCount = -1;

        /// <summary>4-3: 吹き出しに「エコと節約」が含まれるコメントが1件でもあるか（解決済み含む）。解決済みのみ残っている教材で削除したときの検知用。</summary>
        private bool? _lastAnyEcoPhraseInBalloons;

        private string _ecoBaselineDocumentKey;

        /// <summary>4-3: 前ティックの Comments.Count。吹き出し本文が Range で読めない環境でも削除を検知する。</summary>
        private int _lastCommentsCountForEco = -1;

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

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("[New_MOSWordVSTOAddIn] Add-in started");
            System.Diagnostics.Debug.WriteLine($"[New_MOSWordVSTOAddIn] Log file: {Logger.GetLogFilePath()}");

            this.Application.DocumentChange += Application_DocumentChange;
            RefreshBaselineFromActiveDocumentNoLog();
            StartShowAllPolling();
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
            }
            catch
            {
                // アンインストール時など Application が無い場合
            }

            _showAllPollTimer?.Stop();
            _showAllPollTimer?.Dispose();
            _showAllPollTimer = null;
        }

        private void Application_DocumentChange()
        {
            RefreshBaselineFromActiveDocumentNoLog();
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

                // ActiveDocument は Word が管理する参照のため ReleaseComObject しない
                Word.Document doc = app.ActiveDocument;

                // 4-3: コメントペイン等（_WwG 外フォーカス）でも未解決エコ件数だけは追跡する。全文書 COM をスキップする前に実行。
                UpdateEcoCommentBaselineAndMaybeLog(doc, CountUnresolvedEcoComments(doc));

                // リボン（フォントサイズ欄等）・代替テキスト等の入力中は COM を触らない
                if (ShouldSkipDocumentComBecauseFocusNotInEditingPane())
                    return;

                if (app.ActiveWindow?.View != null)
                {
                    bool viewShowAll = app.ActiveWindow.View.ShowAll;
                    bool optionsShowAll = TryGetOptionsShowAll(app);
                    _lastShowAllState = viewShowAll || optionsShowAll;
                }

                _lastColumnBreakCount = CountColumnBreaks(doc);
                _lastOrientationFingerprint = GetAllSectionsOrientationFingerprint(doc);
                _lastPageBorderFingerprint = GetPageBorderFingerprint(doc);
                _lastHeading1LineSimple = IsHeading1LineSimplePattern(doc);
                _lastTask1_2_03ColorFingerprint = GetTask1_2_03ColorFingerprint(doc);

                // 4-5: ベースライン取得
                try {
                    string xml = doc.WordOpenXML;
                    _lastWatermarkFound = !string.IsNullOrEmpty(xml) && xml.Contains("下書き");
                } catch { _lastWatermarkFound = false; }
            }
            catch
            {
                // COM 初期化中などは無視
            }
        }

        /// <summary>4-3: 未解決エコ件数の変化を追跡。1→0 のとき解決/削除をログ。加えて「エコ」吹き出しが true→false（解決済みスレッドの削除など）でも ReviewDeleteComment をログする。</summary>
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
                        _lastAnyEcoPhraseInBalloons = null;
                        _lastCommentsCountForEco = -1;
                    }
                }
            }
            catch { }

            bool anyPhrase = false;
            try
            {
                if (doc != null)
                    anyPhrase = DocumentHasAnyEcoCommentBalloonPoll(doc);
            }
            catch { }

            bool willLogUnresolved = _lastUnresolvedEcoCommentCount >= 0 && _lastUnresolvedEcoCommentCount > 0 && newUnresolvedEcoCount == 0;
            bool phraseDisappeared = _lastAnyEcoPhraseInBalloons == true && !anyPhrase && doc != null;

            int cc = -1;
            try
            {
                if (doc != null && doc.Comments != null)
                    cc = doc.Comments.Count;
            }
            catch { cc = -1; }

            bool commentCountDropped = _lastCommentsCountForEco >= 0 && cc >= 0 && cc < _lastCommentsCountForEco;

            if (willLogUnresolved && doc != null)
            {
                bool stillHasEcoBalloon = anyPhrase;
                string commandId = stillHasEcoBalloon ? "ReviewResolveComment" : "ReviewDeleteComment";
                Logger.LogCommand(commandId);
            }
            else if (phraseDisappeared)
            {
                Logger.LogCommand("ReviewDeleteComment");
            }
            else if (commentCountDropped && doc != null && DocumentBodyContainsEcoPhraseForTask4_3(doc))
            {
                Logger.LogCommand("ReviewDeleteComment");
            }

            _lastUnresolvedEcoCommentCount = newUnresolvedEcoCount;
            _lastAnyEcoPhraseInBalloons = anyPhrase;
            if (cc >= 0)
                _lastCommentsCountForEco = cc;
        }

        /// <summary>
        /// 1-1 編集記号の表示/非表示は Word の idMso でフックできないため、
        /// View.ShowAll の状態をポーリングし、変化時に ShowAll をログに記録する。
        /// 向き・ページ罫線・スタイルセット（線シンプル相当）も idMso が発火しない経路があるため同タイマーで差分検知する。
        /// </summary>
        private void StartShowAllPolling()
        {
            // Word UI（リボンのテキスト入力等）への干渉を減らすため、ポーリング間隔を控えめにする
            _showAllPollTimer = new Timer { Interval = 1200 };
            _showAllPollTimer.Tick += ShowAllPoll_Tick;
            _showAllPollTimer.Start();
        }

        private void ShowAllPoll_Tick(object sender, EventArgs e)
        {
            try
            {
                var app = this.Application;
                if (app?.ActiveWindow?.View == null || app.Documents.Count == 0)
                    return;

                // ActiveDocument は Word が管理する参照のため ReleaseComObject しない
                Word.Document doc = app.ActiveDocument;

                // 4-3: コメントペイン・リボン等でも未解決エコ件数だけは追跡（ShowAll 等の重い COM より前に実行）
                UpdateEcoCommentBaselineAndMaybeLog(doc, CountUnresolvedEcoComments(doc));

                if (ShouldSkipDocumentComBecauseFocusNotInEditingPane())
                    return;

                bool viewShowAll = app.ActiveWindow.View.ShowAll;
                bool optionsShowAll = TryGetOptionsShowAll(app);

                bool currentShowAll = viewShowAll || optionsShowAll;
                if (_lastShowAllState.HasValue && _lastShowAllState.Value != currentShowAll)
                {
                    Logger.LogCommand("ShowAll");
                }
                _lastShowAllState = currentShowAll;

                string orientFp = GetAllSectionsOrientationFingerprint(doc);
                if (_lastOrientationFingerprint != null && orientFp != _lastOrientationFingerprint)
                {
                    if (_suppressOrientationPollLogs > 0)
                        _suppressOrientationPollLogs--;
                    else
                        Logger.LogCommand("PageOrientationPortraitLandscape");
                }
                _lastOrientationFingerprint = orientFp;

                // 重い判定（文書全体テキスト化・セクション走査・スタイル解析）は毎回実行しない。
                // リボン入力欄（フォントサイズ等）でのフォーカス喪失を避けるため、約6秒ごとに間引く。
                _heavyCheckTickCounter++;
                bool runHeavyChecks = (_heavyCheckTickCounter % 5) == 0;
                if (runHeavyChecks)
                {
                    string borderFp = GetPageBorderFingerprint(doc);
                    if (_lastPageBorderFingerprint != null && borderFp != _lastPageBorderFingerprint)
                    {
                        if (_suppressPageBorderPollLogs > 0)
                            _suppressPageBorderPollLogs--;
                        else
                            Logger.LogCommand("PageBorders");
                    }
                    _lastPageBorderFingerprint = borderFp;

                    bool lineSimple = IsHeading1LineSimplePattern(doc);
                    if (_lastHeading1LineSimple.HasValue && lineSimple && !_lastHeading1LineSimple.Value)
                    {
                        if (_suppressStyleSetPollLogs > 0)
                            _suppressStyleSetPollLogs--;
                        else
                            Logger.LogCommand("StyleSetLineSimple");
                    }
                    _lastHeading1LineSimple = lineSimple;

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

                    // 4-5: 透かしの検知（WordOpenXML を使用）
                    try
                    {
                        string xml = doc.WordOpenXML;
                        bool hasWatermark = !string.IsNullOrEmpty(xml) && xml.Contains("下書き");
                        if (hasWatermark && (!_lastWatermarkFound.HasValue || !_lastWatermarkFound.Value))
                        {
                            Logger.LogCommand("Watermark");
                        }
                        _lastWatermarkFound = hasWatermark;
                    }
                    catch { }

                    // 5-1, 5-2: 画像レイアウトの検知（5月21日...段落付近）
                    try
                    {
                        Word.Range searchRange = doc.Content;
                        Word.Find find = searchRange.Find;
                        find.ClearFormatting();
                        find.Text = "5月21日より5日間の";
                        if (find.Execute())
                        {
                            Word.Range paraRange = searchRange.Paragraphs[1].Range;
                            int paraStart = paraRange.Start;
                            int paraEnd = paraRange.End;

                            int currentWrapType = -1; // -1: なし, 0: 行内, 1: 四角形など

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
                                            else currentWrapType = 2; // その他
                                            break;
                                        }
                                    }
                                    finally { Marshal.ReleaseComObject(sh); }
                                }
                            }

                            if (_lastLaptopWrapType != currentWrapType)
                            {
                                if (currentWrapType == 0) Logger.LogCommand("WrapInline");
                                else if (currentWrapType == 1) Logger.LogCommand("WrapSquare");
                            }
                            _lastLaptopWrapType = currentWrapType;
                            
                            Marshal.ReleaseComObject(paraRange);
                        }
                        Marshal.ReleaseComObject(find);
                        Marshal.ReleaseComObject(searchRange);
                    }
                    catch { }
                }
            }
            catch
            {
                // ドキュメント未表示などで COM エラーになることがあるため無視
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
            string text = doc.Content?.Text ?? string.Empty;
            return text.Count(c => c == (char)14);
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
                        if (!CommentHasEcoTextInBalloon(c))
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
                searchRange = doc.Content;
                find = searchRange.Find;
                find.ClearFormatting();
                find.Text = "エコと節約";
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
                        if (CommentHasEcoTextInBalloon(c))
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
                searchRange = doc.Content;
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
            bool isDarker25 = shade >= -0.31f && shade <= -0.19f;

            int r = rgb & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;
            bool isBlue = b >= r && b >= g && b > 0;

            return isAccent1 && isDarker25 && isBlue;
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

        private static string GetPageBorderFingerprint(Word.Document doc)
        {
            Word.Section sec = null;
            Word.Borders borders = null;
            Word.Border top = null;
            Word.Border bottom = null;
            try
            {
                sec = doc.Sections[1];
                borders = sec.Borders;
                top = borders[Word.WdBorderType.wdBorderTop];
                bottom = borders[Word.WdBorderType.wdBorderBottom];
                return string.Format(System.Globalization.CultureInfo.InvariantCulture,
                    "{0},{1},{2},{3}",
                    (int)top.LineStyle, (int)top.LineWidth, (int)bottom.LineStyle, (int)bottom.LineWidth);
            }
            catch
            {
                return string.Empty;
            }
            finally
            {
                if (bottom != null)
                    Marshal.ReleaseComObject(bottom);
                if (top != null)
                    Marshal.ReleaseComObject(top);
                if (borders != null)
                    Marshal.ReleaseComObject(borders);
                if (sec != null)
                    Marshal.ReleaseComObject(sec);
            }
        }

        /// <summary>
        /// WordChecker1_4 の 4-4（線・シンプル）と同条件。
        /// </summary>
        private static bool IsHeading1LineSimplePattern(Word.Document document)
        {
            Word.Style headingStyle = null;
            Word.Borders borders = null;
            Word.Border bottomBorder = null;
            try
            {
                headingStyle = GetHeading1Style(document);

                if (headingStyle == null)
                    return false;

                borders = headingStyle.ParagraphFormat.Borders;
                bottomBorder = borders[Word.WdBorderType.wdBorderBottom];
                return bottomBorder != null
                       && (Word.WdLineStyle)bottomBorder.LineStyle == Word.WdLineStyle.wdLineStyleSingle
                       && (Word.WdLineWidth)bottomBorder.LineWidth == Word.WdLineWidth.wdLineWidth050pt;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (bottomBorder != null)
                    Marshal.ReleaseComObject(bottomBorder);
                if (borders != null)
                    Marshal.ReleaseComObject(borders);
                if (headingStyle != null)
                    Marshal.ReleaseComObject(headingStyle);
            }
        }

        /// <summary>
        /// 表示言語によるスタイル名差異を吸収して Heading 1 を取得する。
        /// </summary>
        private static Word.Style GetHeading1Style(Word.Document document)
        {
            try
            {
                return document.Styles[Word.WdBuiltinStyle.wdStyleHeading1];
            }
            catch
            {
                // 旧ロジック互換の名前フォールバック
                try { return document.Styles["見出し 1"]; }
                catch
                {
                    try { return document.Styles["見出し1"]; }
                    catch
                    {
                        try { return document.Styles["Heading 1"]; }
                        catch { return null; }
                    }
                }
            }
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
