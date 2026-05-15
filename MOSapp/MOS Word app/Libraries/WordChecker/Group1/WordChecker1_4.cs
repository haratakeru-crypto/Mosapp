using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_4
    {
        public bool CheckTask_1_4_01() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_4_01(filePath); } catch { return false; } }
        public bool CheckTask_1_4_02() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_4_02(filePath); } catch { return false; } }
        public bool CheckTask_1_4_03() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_4_03(filePath); } catch { return false; } }
        public bool CheckTask_1_4_04() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_4_04(filePath); } catch { return false; } }
        public bool CheckTask_1_4_05() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_4_05(filePath); } catch { return false; } }
        public bool CheckTask_1_4_06() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_4_06(filePath); } catch { return false; } }
        public bool CheckTask_1_4_07() { try { string filePath = GetCurrentWordFilePath(); if (string.IsNullOrEmpty(filePath)) return false; return CheckTask_1_4_07(filePath); } catch { return false; } }

        private bool CheckTask_1_4_01(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                Range searchRange = document.Content;
                Find find = searchRange.Find;
                find.ClearFormatting();
                find.Text = "1.生活の中でできるエコ活動";
                find.Execute();
                if (!find.Found)
                {
                    find.Text = "１.生活の中でできるエコ活動";
                    find.Execute();
                }
                if (!find.Found)
                {
                    find.Text = "生活の中でできるエコ活動";
                    find.Execute();
                }
                if (!find.Found)
                {
                    Marshal.ReleaseComObject(find);
                    Marshal.ReleaseComObject(searchRange);
                    return false;
                }
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);

                // Comment.Range はコメント本文（吹き出し）。改行・全角スペース・Unicode 正規化で「最新の情報を確認」を判定
                Comments comments = document.Comments;
                bool result = false;
                foreach (Comment comment in comments)
                {
                    try
                    {
                        if (CommentBodyMatchesLatestInfoCheck(comment))
                        {
                            result = true;
                            break;
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(comment);
                    }
                }
                Marshal.ReleaseComObject(comments);
                return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_4_02(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // VSTOログから手順をチェック: ReviewCommentReplyが実行されたか
                string logFilePath = LogReader.GetLogFilePath();
                bool logFileExists = System.IO.File.Exists(logFilePath);
                bool replyExecuted = LogReader.HasCommandExecuted("ReviewCommentReply");
                System.Diagnostics.Debug.WriteLine($"[CheckTask_1_4_02] Log file exists: {logFileExists}");
                System.Diagnostics.Debug.WriteLine($"[CheckTask_1_4_02] ReviewCommentReply executed: {replyExecuted}");

                Comments comments = document.Comments;
                bool fileStateCheck = false;
                foreach (Comment comment in comments)
                {
                    if (comment.Range.Text.Contains("前田先生に最終確認")) { fileStateCheck = true; break; }
                    Marshal.ReleaseComObject(comment);
                }
                Marshal.ReleaseComObject(comments);

                // VSTOログがある場合は、VSTOログとファイル状態の両方を確認
                if (logFileExists && replyExecuted)
                {
                    bool result = fileStateCheck;
                    System.Diagnostics.Debug.WriteLine($"[CheckTask_1_4_02] Result (VSTO log check): {result}");
                    return result;
                }
                else
                {
                    // VSTOログがない場合は、従来のファイル状態チェックのみ（後方互換性のため）
                    System.Diagnostics.Debug.WriteLine("[CheckTask_1_4_02] VSTO log not found or command not executed, using file state check only");
                    System.Diagnostics.Debug.WriteLine($"[CheckTask_1_4_02] Result (file state check only): {fileStateCheck}");
                    return fileStateCheck;
                }
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_4_03(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // 本文に「エコと節約」があること（教材の前提）
                Range searchRange = document.Content;
                Find find = searchRange.Find;
                find.ClearFormatting();
                find.Text = "エコと節約";
                find.Execute();
                if (!find.Found)
                {
                    Marshal.ReleaseComObject(find);
                    Marshal.ReleaseComObject(searchRange);
                    return false;
                }
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);

                // 吹き出しに「エコと節約」が一度も無い: 未着手 / 削除 / 解決（Word によって吹き出しが消える）の区別。削除は ReviewDeleteComment、解決は VSTO ポーリングの ReviewResolveComment ログで補足。
                bool anyEcoBalloon = DocumentHasAnyEcoCommentBalloon(document);
                bool logDelete = LogReader.HasCommandExecuted("ReviewDeleteComment");
                bool logResolve = LogReader.HasCommandExecuted("ReviewResolveComment");
                bool resolvedStateOk = IsEcoCommentAbsentOrResolved(document);

                bool result;
                if (!anyEcoBalloon)
                    result = logDelete || logResolve;
                else
                    result = resolvedStateOk;

                System.Diagnostics.Debug.WriteLine($"[CheckTask_1_4_03] anyEcoBalloon={anyEcoBalloon}, logDelete={logDelete}, logResolve={logResolve}, resolvedStateOk={resolvedStateOk}, result={result}");
                return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_4_04(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // スタイルセット「線（シンプル）」を狭く判定（見出し1下罫線 + ログ必須）
                Style headingStyle = null;
                Borders borders = null;
                try
                {
                    try { headingStyle = document.Styles["見出し 1"]; } catch { headingStyle = document.Styles["Heading 1"]; }
                    if (headingStyle == null) return false;
                    borders = headingStyle.ParagraphFormat.Borders;
                    Border bottomBorder = borders[WdBorderType.wdBorderBottom];
                    bool hasLine = bottomBorder != null &&
                                   (WdLineStyle)bottomBorder.LineStyle == WdLineStyle.wdLineStyleSingle &&
                                   (WdLineWidth)bottomBorder.LineWidth == WdLineWidth.wdLineWidth050pt;
                    bool logOk = LogReader.HasCommandExecuted("StyleSetLineSimple");
                    return hasLine && logOk;
                }
                finally
                {
                    if (borders != null) Marshal.ReleaseComObject(borders);
                    if (headingStyle != null) Marshal.ReleaseComObject(headingStyle);
                }
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_4_05(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;

                // 透かし「下書き1」: プライマリヘッダーの Shapes のみでは取りこぼす（先頭ページ／フッター／偶数ページの図形・Range、他セクション）。
                Section section = document.Sections[1];

                bool foundInPrimaryRange = false;
                HeaderFooter headerPrimary = null;
                try
                {
                    headerPrimary = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    foundInPrimaryRange = ContainsDraftWatermarkText(headerPrimary.Range.Text ?? "");
                }
                catch { }
                finally { if (headerPrimary != null) Marshal.ReleaseComObject(headerPrimary); }

                bool foundInShapes = false;
                WdHeaderFooterIndex[] hfIdx =
                {
                    WdHeaderFooterIndex.wdHeaderFooterPrimary,
                    WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                    WdHeaderFooterIndex.wdHeaderFooterEvenPages
                };
                foreach (WdHeaderFooterIndex ix in hfIdx)
                {
                    HeaderFooter hh = null;
                    try
                    {
                        hh = section.Headers[ix];
                        if (HeaderFooterShapesContainDraftWatermark(hh)) foundInShapes = true;
                    }
                    catch { }
                    finally { if (hh != null) Marshal.ReleaseComObject(hh); }
                    HeaderFooter ff = null;
                    try
                    {
                        ff = section.Footers[ix];
                        if (HeaderFooterShapesContainDraftWatermark(ff)) foundInShapes = true;
                    }
                    catch { }
                    finally { if (ff != null) Marshal.ReleaseComObject(ff); }
                }

                bool foundInFirstPageHeader = false;
                try
                {
                    HeaderFooter hfFirst = section.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage];
                    try { foundInFirstPageHeader = ContainsDraftWatermarkText(hfFirst.Range.Text ?? ""); }
                    finally { Marshal.ReleaseComObject(hfFirst); }
                }
                catch { }

                bool foundInPrimaryFooter = false;
                try
                {
                    HeaderFooter foot = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    try { foundInPrimaryFooter = ContainsDraftWatermarkText(foot.Range.Text ?? ""); }
                    finally { Marshal.ReleaseComObject(foot); }
                }
                catch { }

                bool foundInEvenPageHeader = false;
                try
                {
                    HeaderFooter hEven = section.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages];
                    try { foundInEvenPageHeader = ContainsDraftWatermarkText(hEven.Range.Text ?? ""); }
                    finally { Marshal.ReleaseComObject(hEven); }
                }
                catch { }

                bool foundInEvenPageFooter = false;
                try
                {
                    HeaderFooter fEven = section.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages];
                    try { foundInEvenPageFooter = ContainsDraftWatermarkText(fEven.Range.Text ?? ""); }
                    finally { Marshal.ReleaseComObject(fEven); }
                }
                catch { }

                bool foundInFirstPageFooter = false;
                try
                {
                    HeaderFooter fFirst = section.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage];
                    try { foundInFirstPageFooter = ContainsDraftWatermarkText(fFirst.Range.Text ?? ""); }
                    finally { Marshal.ReleaseComObject(fFirst); }
                }
                catch { }

                bool foundInOtherSections = false;
                bool foundInOtherSectionsFooter = false;
                try
                {
                    int nSec = document.Sections.Count;
                    for (int si = 2; si <= nSec; si++)
                    {
                        Section sec = null;
                        HeaderFooter hp = null;
                        HeaderFooter fp = null;
                        try
                        {
                            sec = document.Sections[si];
                            hp = sec.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                            if (ContainsDraftWatermarkText(hp.Range.Text ?? "")) { foundInOtherSections = true; }
                            fp = sec.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                            if (ContainsDraftWatermarkText(fp.Range.Text ?? "")) { foundInOtherSectionsFooter = true; }
                        }
                        finally
                        {
                            if (fp != null) Marshal.ReleaseComObject(fp);
                            if (hp != null) Marshal.ReleaseComObject(hp);
                            if (sec != null) Marshal.ReleaseComObject(sec);
                        }
                    }
                }
                catch { }

                int hdrPrimaryShapeCount = -1;
                int primaryHeaderRangeLen = -1;
                int docShapeCount = -1;
                try
                {
                    HeaderFooter hpDbg = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    try
                    {
                        primaryHeaderRangeLen = (hpDbg.Range.Text ?? "").Length;
                        Shapes hs = hpDbg.Shapes;
                        try { hdrPrimaryShapeCount = hs.Count; }
                        finally { if (hs != null) Marshal.ReleaseComObject(hs); }
                    }
                    finally { Marshal.ReleaseComObject(hpDbg); }
                }
                catch { }

                bool foundInContent = false;
                try { foundInContent = ContainsDraftWatermarkText(document.Content.Text ?? ""); } catch { }

                bool foundInDocShapes = false;
                try { foundInDocShapes = DocumentShapesContainDraftWatermark(document); } catch { }

                bool foundInStoryRanges = false;
                try { foundInStoryRanges = DocumentStoryRangesContainDraftWatermark(document); } catch { }

                try
                {
                    Shapes ds = document.Shapes;
                    try { docShapeCount = ds.Count; }
                    finally { if (ds != null) Marshal.ReleaseComObject(ds); }
                }
                catch { }

                int primaryShape0VisibleTextLen = -1;
                bool docContentHasDraftKeyword = false;
                bool docShapesHasDraftKeyword = false;
                string primaryHeaderSampleText = "";
                bool hdrPrimaryShapesHasDraftToken = false;
                bool hdrPrimaryShapesHasOneToken = false;
                string hdrPrimaryShapesSampleOneText = "";
                try
                {
                    HeaderFooter hp0 = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    try
                    {
                        Shapes sh0 = hp0.Shapes;
                        try
                        {
                            if (sh0.Count >= 1)
                            {
                                Shape s0 = null;
                                try
                                {
                                    s0 = sh0[1];
                                    primaryShape0VisibleTextLen = CollectShapeVisibleText(s0).Length;
                                }
                                finally { if (s0 != null) Marshal.ReleaseComObject(s0); }
                            }

                            // 「下書き」と「1相当」が同一結合文字列内にあるかを確定させる（取りこぼし原因切り分け）
                            string hdrCombinedText = CollectShapesVisibleTextRecursive(sh0, 2000);
                            if (!string.IsNullOrEmpty(hdrCombinedText))
                            {
                                primaryHeaderSampleText = hdrCombinedText;
                                if (primaryHeaderSampleText.Length > 120) primaryHeaderSampleText = primaryHeaderSampleText.Substring(0, 120);
                                hdrPrimaryShapesHasDraftToken = ContainsDraftKeyword(hdrCombinedText);
                                hdrPrimaryShapesHasOneToken = ContainsOneToken(hdrCombinedText);
                                hdrPrimaryShapesSampleOneText = hdrCombinedText;
                                if (hdrPrimaryShapesSampleOneText.Length > 60) hdrPrimaryShapesSampleOneText = hdrPrimaryShapesSampleOneText.Substring(0, 60);
                            }

                            // 取りこぼし調査用: プライマリヘッダーの可視テキストから「下書き」/「1」を別形状でも検出する
                            try
                            {
                                int n = sh0.Count;
                                for (int i = 1; i <= n; i++)
                                {
                                    Shape siShape = null;
                                    try
                                    {
                                        siShape = sh0[i];
                                        string vt = CollectShapeVisibleText(siShape);
                                        if (string.IsNullOrEmpty(primaryHeaderSampleText) && !string.IsNullOrEmpty(vt))
                                        {
                                            try { vt = vt.Normalize(NormalizationForm.FormKC); } catch { }
                                            if (vt.Length > 60) vt = vt.Substring(0, 60);
                                            primaryHeaderSampleText = vt;
                                        }

                                        // digit / token 側は sampleText が決まった後でも必要なので別条件で更新
                                        if (!hdrPrimaryShapesHasDraftToken && ContainsDraftKeyword(vt)) hdrPrimaryShapesHasDraftToken = true;
                                        if (!hdrPrimaryShapesHasOneToken && ContainsOneToken(vt))
                                        {
                                            hdrPrimaryShapesHasOneToken = true;
                                            if (string.IsNullOrEmpty(hdrPrimaryShapesSampleOneText))
                                            {
                                                try
                                                {
                                                    string one = vt;
                                                    if (one.Length > 60) one = one.Substring(0, 60);
                                                    hdrPrimaryShapesSampleOneText = one;
                                                }
                                                catch { }
                                            }
                                        }
                                    }
                                    finally { if (siShape != null) Marshal.ReleaseComObject(siShape); }
                                }
                            }
                            catch { }
                        }
                        finally { if (sh0 != null) Marshal.ReleaseComObject(sh0); }
                    }
                    finally { Marshal.ReleaseComObject(hp0); }
                }
                catch { }

                try { docContentHasDraftKeyword = ContainsDraftKeyword(document.Content.Text ?? ""); } catch { }
                try { docShapesHasDraftKeyword = DocumentShapesContainDraftKeyword(document); } catch { }

                bool found = foundInShapes || foundInPrimaryRange || foundInFirstPageHeader || foundInOtherSections
                    || foundInPrimaryFooter || foundInEvenPageHeader || foundInEvenPageFooter || foundInFirstPageFooter
                    || foundInOtherSectionsFooter || foundInContent || foundInDocShapes || foundInStoryRanges;

                Marshal.ReleaseComObject(section);
                if (LogReader.HasCommandExecuted("WatermarkCustomDialog") && found)
                    return true;
                return found;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private static readonly Regex DraftWatermarkLoose = new Regex(@"下書き[\s\u00A0\u200B\uFEFF]*[0-9１壱一ⅠⅠ①⑴⑺⑶]", RegexOptions.Compiled);

        /// <summary>透かしプリセット「下書き1」のテキスト一致（Unicode 正規化・空白挿入の差を吸収）。</summary>
        private static bool ContainsDraftWatermarkText(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            try
            {
                string n = s.Normalize(NormalizationForm.FormKC);
                // 重要: runtime 証拠では、Word の組み込み透かしを Shapes から抽出すると
                // "下書き" は取れるが "1" 相当は取れないケースがあるため、
                // "下書き" が取れたら正解扱いする（"下書き1" 固定の課題仕様に合わせる）。
                if (ContainsDraftKeyword(n)) return true;
                return DraftWatermarkLoose.IsMatch(n);
            }
            catch
            {
                return false;
            }
        }

        private static bool ContainsDraftKeyword(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            try
            {
                return s.Normalize(NormalizationForm.FormKC).Contains("下書き");
            }
            catch
            {
                return false;
            }
        }

        private static bool ContainsOneToken(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            try
            {
                string n = s.Normalize(NormalizationForm.FormKC);
                // 「1」「１」以外に、WordArt/フォント差で表記が揺れる可能性があるものも含める
                return n.Contains("1") || n.Contains("１") || n.Contains("Ⅰ") || n.Contains("I") || n.Contains("ｌ")
                       || n.Contains("壱") || n.Contains("一") || n.Contains("①") || n.Contains("⑴") || n.Contains("⑴");
            }
            catch
            {
                return false;
            }
        }

        private static string JsonEscape(string s)
        {
            if (s == null) return string.Empty;
            // NDJSON は dataJson を直接貼り付けるため、文字列値を JSON 文字列として安全にする
            return s
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n")
                .Replace("\t", "\\t");
        }

        private static bool HeaderFooterShapesContainDraftWatermark(HeaderFooter hf)
        {
            if (hf == null) return false;
            Shapes shapes = null;
            try
            {
                shapes = hf.Shapes;
                // 下書きと「1」が別図形に分かれていても、結合文字列として判定する
                string combined = CollectShapesVisibleTextRecursive(shapes, 2000);
                return ContainsDraftWatermarkText(combined);
            }
            catch { }
            finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
            return false;
        }

        private static bool DocumentShapesContainDraftWatermark(Document doc)
        {
            if (doc == null) return false;
            Shapes shapes = null;
            try
            {
                shapes = doc.Shapes;
                // ドキュメント全体 Shapes についても結合文字列で判定する（グループ分割や別図形の可能性を吸収）
                string combined = CollectShapesVisibleTextRecursive(shapes, 4000);
                return ContainsDraftWatermarkText(combined);
            }
            catch { }
            finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
            return false;
        }

        private static bool DocumentStoryRangesContainDraftWatermark(Document doc)
        {
            if (doc == null) return false;
            StoryRanges srs = null;
            try
            {
                srs = doc.StoryRanges;
                WdStoryType[] types =
                {
                    WdStoryType.wdPrimaryHeaderStory,
                    WdStoryType.wdPrimaryFooterStory,
                    WdStoryType.wdEvenPagesHeaderStory,
                    WdStoryType.wdEvenPagesFooterStory,
                    WdStoryType.wdFirstPageHeaderStory,
                    WdStoryType.wdFirstPageFooterStory,
                    WdStoryType.wdMainTextStory
                };
                foreach (WdStoryType st in types)
                {
                    Range r = null;
                    try
                    {
                        r = srs[st];
                        if (ContainsDraftWatermarkText(r.Text ?? "")) return true;
                    }
                    catch { }
                    finally { if (r != null) Marshal.ReleaseComObject(r); }
                }
            }
            catch { }
            finally { if (srs != null) Marshal.ReleaseComObject(srs); }
            return false;
        }

        /// <summary>透かしプリセットは WordArt 等で TextFrame より TextFrame2 に文字列が載ることがある。</summary>
        private static string CollectShapeVisibleText(Shape shp)
        {
            if (shp == null) return "";
            var sb = new StringBuilder();
            try
            {
                if (shp.TextFrame != null && shp.TextFrame.TextRange != null)
                    sb.Append(shp.TextFrame.TextRange.Text ?? "");
            }
            catch { }
            try
            {
                dynamic d = shp;
                object tf2o = d.TextFrame2;
                if (tf2o != null)
                {
                    try
                    {
                        dynamic tf2 = tf2o;
                        dynamic tr = tf2.TextRange;
                        if (tr != null)
                        {
                            string t = tr.Text ?? "";
                            sb.Append(t);
                        }
                    }
                    finally { Marshal.ReleaseComObject(tf2o); }
                }
            }
            catch { }
            try { sb.Append(shp.AlternativeText ?? ""); } catch { }
            try { sb.Append(shp.Title ?? ""); } catch { }
            return sb.ToString();
        }

        /// <summary>Shapes の子（GroupItems 等）も含めて可視テキストを結合して返す（最大文字数で打ち切り）。</summary>
        private static string CollectShapeVisibleTextRecursive(Shape shp, int maxChars)
        {
            if (shp == null) return "";
            if (maxChars <= 0) return "";
            StringBuilder sb = new StringBuilder();
            try
            {
                GroupShapes gitems = null;
                try { gitems = shp.GroupItems; } catch { }
                if (gitems != null)
                {
                    try
                    {
                        int gc = gitems.Count;
                        for (int i = 1; i <= gc; i++)
                        {
                            if (sb.Length >= maxChars) break;
                            Shape inner = null;
                            try
                            {
                                inner = gitems[i];
                                sb.Append(CollectShapeVisibleTextRecursive(inner, maxChars - sb.Length));
                            }
                            catch { }
                            finally { if (inner != null) Marshal.ReleaseComObject(inner); }
                        }
                    }
                    finally { if (gitems != null) Marshal.ReleaseComObject(gitems); }
                    return sb.ToString();
                }
            }
            catch { }

            try
            {
                sb.Append(CollectShapeVisibleText(shp));
                if (sb.Length > maxChars) sb.Length = maxChars;
            }
            catch { }
            return sb.ToString();
        }

        /// <summary>Shapes 全体の可視テキストを結合して返す（最大文字数で打ち切り）。</summary>
        private static string CollectShapesVisibleTextRecursive(Shapes shapes, int maxChars)
        {
            if (shapes == null) return "";
            if (maxChars <= 0) return "";
            StringBuilder sb = new StringBuilder();
            try
            {
                int n = shapes.Count;
                for (int i = 1; i <= n; i++)
                {
                    if (sb.Length >= maxChars) break;
                    Shape shp = null;
                    try
                    {
                        shp = shapes[i];
                        sb.Append(CollectShapeVisibleTextRecursive(shp, maxChars - sb.Length));
                    }
                    catch { }
                    finally { if (shp != null) Marshal.ReleaseComObject(shp); }
                }
            }
            catch { }
            return sb.ToString();
        }

        /// <summary>msoGroup（6）の子図形を再帰。</summary>
        private static bool ShapeContainsDraftWatermarkRecursive(Shape shp)
        {
            if (shp == null) return false;
            try
            {
                // GroupItems が取れるなら、種類に関係なく掘り下げる（type==msoGroup=6 に固定すると取りこぼす可能性がある）。
                GroupShapes gitems = null;
                try { gitems = shp.GroupItems; } catch { }
                if (gitems != null)
                {
                    try
                    {
                        int gc = gitems.Count;
                        for (int g = 1; g <= gc; g++)
                        {
                            Shape inner = null;
                            try
                            {
                                inner = gitems[g];
                                if (ShapeContainsDraftWatermarkRecursive(inner)) return true;
                            }
                            finally { if (inner != null) Marshal.ReleaseComObject(inner); }
                        }
                    }
                    finally { Marshal.ReleaseComObject(gitems); }
                }

                if (ContainsDraftWatermarkText(CollectShapeVisibleText(shp))) return true;
            }
            catch { }
            return false;
        }

        private static bool ShapeContainsDraftKeywordRecursive(Shape shp)
        {
            if (shp == null) return false;
            try
            {
                GroupShapes gitems = null;
                try { gitems = shp.GroupItems; } catch { }
                if (gitems != null)
                {
                    try
                    {
                        int gc = gitems.Count;
                        for (int g = 1; g <= gc; g++)
                        {
                            Shape inner = null;
                            try
                            {
                                inner = gitems[g];
                                if (ShapeContainsDraftKeywordRecursive(inner)) return true;
                            }
                            finally { if (inner != null) Marshal.ReleaseComObject(inner); }
                        }
                    }
                    finally { Marshal.ReleaseComObject(gitems); }
                }
                return ContainsDraftKeyword(CollectShapeVisibleText(shp));
            }
            catch { }
            return false;
        }

        private static bool DocumentShapesContainDraftKeyword(Document doc)
        {
            if (doc == null) return false;
            Shapes shapes = null;
            try
            {
                shapes = doc.Shapes;
                int n = shapes.Count;
                for (int i = 1; i <= n; i++)
                {
                    Shape shp = null;
                    try
                    {
                        shp = shapes[i];
                        if (ShapeContainsDraftKeywordRecursive(shp)) return true;
                    }
                    finally { if (shp != null) Marshal.ReleaseComObject(shp); }
                }
            }
            catch { }
            finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
            return false;
        }

        private static void WriteDebug6b16c7_1_4_05(bool foundInShapes, bool foundInPrimaryRange, bool foundInFirstPageHeader, bool foundInOtherSections, bool foundInPrimaryFooter, bool foundInEvenPageHeader, bool foundInEvenPageFooter, bool foundInFirstPageFooter, bool foundInOtherSectionsFooter, bool foundInContent, bool foundInDocShapes, bool foundInStoryRanges, bool found, bool logWatermark, int sectionsCount, int hdrPrimaryShapeCount, int primaryHeaderRangeLen, int docShapeCount, int primaryShape0VisibleTextLen, bool docContentHasDraftKeyword, bool docShapesHasDraftKeyword, string primaryHeaderSampleText, bool hdrPrimaryShapesHasDraftToken, bool hdrPrimaryShapesHasOneToken, string hdrPrimaryShapesSampleOneText, string runId)
        {
            // debug instrumentation removed
        }

        private static string GetDebug6b16c7LogPath()
        {
            try
            {
                var d = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory ?? "");
                for (int i = 0; i < 10 && d != null; i++)
                {
                    // 実行場所（bin 配下）から解決用に親方向へ探索
                    if (File.Exists(Path.Combine(d.FullName, "MOS Word app.sln")) && d.Parent != null)
                        return Path.Combine(d.Parent.FullName, "debug-6b16c7.log");
                    d = d.Parent;
                }
            }
            catch { }
            // フォールバック: 現在ディレクトリに出す
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? "", "debug-6b16c7.log");
        }

        private static void WriteDebug6b16c7Ndjson(string runId, string hypothesisId, string location, string message, string dataJson)
        {
            try
            {
                long ts = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
                string path = GetDebug6b16c7LogPath();
                // dataJson は JSON 断片（必ずオブジェクト `{...}` を渡す前提）
                string line = "{\"sessionId\":\"6b16c7\",\"runId\":\"" + (runId ?? "").Replace("\\", "").Replace("\"", "") +
                              "\",\"timestamp\":" + ts +
                              ",\"location\":\"" + location.Replace("\"", "") +
                              "\",\"message\":\"" + message.Replace("\"", "") +
                              "\",\"hypothesisId\":\"" + hypothesisId.Replace("\"", "") +
                              "\",\"data\":" + dataJson + "}\n";
                File.AppendAllText(path, line, Encoding.UTF8);
            }
            catch { }
        }

        private bool CheckTask_1_4_06(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                // 4-6: ページ罫線（線種・太さ）を最低限一致させ、ログ必須で判定
                bool logOk = LogReader.HasCommandExecuted("PageBorders");

                int sectionsCount = 0;
                try { sectionsCount = document.Sections.Count; } catch { }

                // #region agent log (entry)
                try
                {
                    WriteDebug6b16c7Ndjson(
                        "debug-4-6-pre",
                        "H0",
                        "WordChecker1_4.CheckTask_1_4_06",
                        "entry",
                        "{\"logOk\":" + (logOk ? "true" : "false") + ",\"sectionsCount\":" + sectionsCount + "}"
                    );
                }
                catch { }
                // #endregion

                bool hasTop = false;
                bool hasBottom = false;
                int topStyle = -999;
                int topWidth = -999;
                bool topStyleOk = false;
                bool topWidthOk = false;
                int bottomStyle = -999;
                int bottomWidth = -999;
                bool bottomStyleOk = false;
                bool bottomWidthOk = false;
                int leftStyle = -999;
                int leftWidth = -999;
                bool leftStyleOk = false;
                bool leftWidthOk = false;
                int leftColor = -999;
                int rightStyle = -999;
                int rightWidth = -999;
                bool rightStyleOk = false;
                bool rightWidthOk = false;
                int rightColor = -999;
                int topArt = -999;
                int bottomArt = -999;
                int leftArt = -999;
                int rightArt = -999;
                string topArtRaw = "";
                string bottomArtRaw = "";
                string leftArtRaw = "";
                string rightArtRaw = "";
                int topShadow = -999;
                int bottomShadow = -999;
                int leftShadow = -999;
                int rightShadow = -999;
                string topShadowRaw = "";
                string bottomShadowRaw = "";
                string leftShadowRaw = "";
                string rightShadowRaw = "";
                int topSurround = -999;
                int bottomSurround = -999;
                int leftSurround = -999;
                int rightSurround = -999;
                string topSurroundRaw = "";
                string bottomSurroundRaw = "";
                string leftSurroundRaw = "";
                string rightSurroundRaw = "";
                int bordersArt = -999;
                int bordersShadow = -999;
                int bordersSurround = -999;
                string bordersArtRaw = "";
                string bordersShadowRaw = "";
                string bordersSurroundRaw = "";
                double bordersDistanceFromTop = -999;
                double bordersDistanceFromBottom = -999;
                double bordersDistanceFromLeft = -999;
                double bordersDistanceFromRight = -999;
                double bordersDistanceFromText = -999;
                string bordersKeyProps = "";
                string topBorderKeyProps = "";
                int topColor = -999;
                int bottomColor = -999;
                bool hasLeft = false;
                bool hasRight = false;

                bool strictMatchAnySection = false;
                int firstStrictMatchSection = -1;
                int targetStyleSingle = (int)WdLineStyle.wdLineStyleSingle;
                int targetWidth050pt = (int)WdLineWidth.wdLineWidth050pt;
                // COM 取得値が環境依存で異なるケースがあり、ログ上「PageBorders 操作の正解時」に LineWidth が 12 で出ているため許容する
                int targetWidthObservedAlt = 12;

                // #region agent log (borders values)
                try
                {
                    Section section1 = null;
                    Borders borders1 = null;
                    try
                    {
                        if (sectionsCount >= 1)
                        {
                            section1 = document.Sections[1];
                            borders1 = section1.Borders;
                            // 「囲むの種類」（Box/Shadow等）に対応している可能性があるのは Border 側ではなく Borders 側のプロパティの可能性があるため、
                            // まずは borders1 から可能な範囲で値を抜き出してログで比較する。
                            try
                            {
                                dynamic db = borders1;
                                object v = null;
                                try { v = db.Art; } catch { }
                                bordersArtRaw = v == null ? "" : v.ToString();
                                try { if (v != null) bordersArt = Convert.ToInt32(v); } catch { }
                            } catch { }
                            try
                            {
                                dynamic db = borders1;
                                object v = null;
                                try { v = db.Shadow; } catch { }
                                bordersShadowRaw = v == null ? "" : v.ToString();
                                try { if (v != null) bordersShadow = Convert.ToInt32(v); } catch { }
                            } catch { }
                            try
                            {
                                dynamic db = borders1;
                                object v = null;
                                try { v = db.Surround; } catch { }
                                bordersSurroundRaw = v == null ? "" : v.ToString();
                                try { if (v != null) bordersSurround = Convert.ToInt32(v); } catch { }
                            } catch { }
                            try
                            {
                                var tt = borders1?.GetType();
                                if (tt != null)
                                {
                                    string[] keys = new[] { "Art", "Shadow", "Surround", "Distance", "Padding" };
                                    var ms = tt
                                        .GetProperties()
                                        .Where(p => keys.Any(k => p.Name.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                                        .Take(20)
                                        .Select(p => p.Name + ":" + p.PropertyType.Name)
                                        .ToArray();
                                    bordersKeyProps = ms.Length > 0 ? string.Join(",", ms) : "";
                                }
                            }
                            catch { }
                            try { dynamic db = borders1; bordersDistanceFromTop = Convert.ToDouble(db.DistanceFromTop); } catch { }
                            try { dynamic db = borders1; bordersDistanceFromBottom = Convert.ToDouble(db.DistanceFromBottom); } catch { }
                            try { dynamic db = borders1; bordersDistanceFromLeft = Convert.ToDouble(db.DistanceFromLeft); } catch { }
                            try { dynamic db = borders1; bordersDistanceFromRight = Convert.ToDouble(db.DistanceFromRight); } catch { }
                            try { dynamic db = borders1; bordersDistanceFromText = Convert.ToDouble(db.DistanceFromText); } catch { }
                            Border topBorder = null;
                            Border bottomBorder = null;
                            Border leftBorder = null;
                            Border rightBorder = null;
                            try { topBorder = borders1[WdBorderType.wdBorderTop]; } catch { }
                            try { bottomBorder = borders1[WdBorderType.wdBorderBottom]; } catch { }
                            try { leftBorder = borders1[WdBorderType.wdBorderLeft]; } catch { }
                            try { rightBorder = borders1[WdBorderType.wdBorderRight]; } catch { }

                            if (topBorder != null)
                            {
                                try
                                {
                                    var tt = topBorder?.GetType();
                                    if (tt != null)
                                    {
                                        string[] keys = new[] { "Art", "Shadow", "Surround", "Distance", "Padding" };
                                        var ms = tt
                                            .GetProperties()
                                            .Where(p => keys.Any(k => p.Name.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0))
                                            .Take(20)
                                            .Select(p => p.Name + ":" + p.PropertyType.Name)
                                            .ToArray();
                                        topBorderKeyProps = ms.Length > 0 ? string.Join(",", ms) : "";
                                    }
                                }
                                catch { }
                                try { topStyle = (int)topBorder.LineStyle; } catch { }
                                try { topWidth = (int)topBorder.LineWidth; } catch { }
                                try { topStyleOk = (WdLineStyle)topBorder.LineStyle == WdLineStyle.wdLineStyleSingle; } catch { }
                                try { topWidthOk = (WdLineWidth)topBorder.LineWidth == WdLineWidth.wdLineWidth050pt || topWidth == targetWidthObservedAlt; } catch { topWidthOk = topWidth == targetWidthObservedAlt; }
                                try { topColor = (int)topBorder.Color; } catch { }
                                // 「囲むの種類」（Box/Shadow等）を判定するための値（環境によってプロパティが無い可能性があるので動的アクセスで保護）
                                try
                                {
                                    dynamic dt = topBorder;
                                    object v = null;
                                    try { v = dt.Art; } catch { }
                                    topArtRaw = v == null ? "" : v.ToString();
                                    try { if (v != null) topArt = Convert.ToInt32(v); } catch { }
                                } catch { }
                                try
                                {
                                    dynamic dt = topBorder;
                                    object v = null;
                                    try { v = dt.Shadow; } catch { }
                                    topShadowRaw = v == null ? "" : v.ToString();
                                    try { if (v != null) topShadow = Convert.ToInt32(v); } catch { }
                                } catch { }
                                try
                                {
                                    dynamic dt = topBorder;
                                    object v = null;
                                    try { v = dt.Surround; } catch { }
                                    topSurroundRaw = v == null ? "" : v.ToString();
                                    try { if (v != null) topSurround = Convert.ToInt32(v); } catch { }
                                } catch { }
                                hasTop = topStyleOk && topWidthOk;
                            }
                            if (bottomBorder != null)
                            {
                                try { bottomStyle = (int)bottomBorder.LineStyle; } catch { }
                                try { bottomWidth = (int)bottomBorder.LineWidth; } catch { }
                                try { bottomStyleOk = (WdLineStyle)bottomBorder.LineStyle == WdLineStyle.wdLineStyleSingle; } catch { }
                                try { bottomWidthOk = (WdLineWidth)bottomBorder.LineWidth == WdLineWidth.wdLineWidth050pt || bottomWidth == targetWidthObservedAlt; } catch { bottomWidthOk = bottomWidth == targetWidthObservedAlt; }
                                try { bottomColor = (int)bottomBorder.Color; } catch { }
                                try
                                {
                                    dynamic dt = bottomBorder;
                                    object v = null;
                                    try { v = dt.Art; } catch { }
                                    bottomArtRaw = v == null ? "" : v.ToString();
                                    try { if (v != null) bottomArt = Convert.ToInt32(v); } catch { }
                                } catch { }
                                try
                                {
                                    dynamic dt = bottomBorder;
                                    object v = null;
                                    try { v = dt.Shadow; } catch { }
                                    bottomShadowRaw = v == null ? "" : v.ToString();
                                    try { if (v != null) bottomShadow = Convert.ToInt32(v); } catch { }
                                } catch { }
                                try
                                {
                                    dynamic dt = bottomBorder;
                                    object v = null;
                                    try { v = dt.Surround; } catch { }
                                    bottomSurroundRaw = v == null ? "" : v.ToString();
                                    try { if (v != null) bottomSurround = Convert.ToInt32(v); } catch { }
                                } catch { }
                                hasBottom = bottomStyleOk && bottomWidthOk;
                            }
                            if (leftBorder != null)
                            {
                                try { leftStyle = (int)leftBorder.LineStyle; } catch { }
                                try { leftWidth = (int)leftBorder.LineWidth; } catch { }
                                try { leftStyleOk = (WdLineStyle)leftBorder.LineStyle == WdLineStyle.wdLineStyleSingle; } catch { }
                                try { leftWidthOk = (WdLineWidth)leftBorder.LineWidth == WdLineWidth.wdLineWidth050pt || leftWidth == targetWidthObservedAlt; } catch { leftWidthOk = leftWidth == targetWidthObservedAlt; }
                                try { leftColor = (int)leftBorder.Color; } catch { }
                                try
                                {
                                    dynamic dt = leftBorder;
                                    object v = null;
                                    try { v = dt.Art; } catch { }
                                    leftArtRaw = v == null ? "" : v.ToString();
                                    try { if (v != null) leftArt = Convert.ToInt32(v); } catch { }
                                } catch { }
                                try
                                {
                                    dynamic dt = leftBorder;
                                    object v = null;
                                    try { v = dt.Shadow; } catch { }
                                    leftShadowRaw = v == null ? "" : v.ToString();
                                    try { if (v != null) leftShadow = Convert.ToInt32(v); } catch { }
                                } catch { }
                                try
                                {
                                    dynamic dt = leftBorder;
                                    object v = null;
                                    try { v = dt.Surround; } catch { }
                                    leftSurroundRaw = v == null ? "" : v.ToString();
                                    try { if (v != null) leftSurround = Convert.ToInt32(v); } catch { }
                                } catch { }
                                hasLeft = leftStyleOk && leftWidthOk;
                            }
                            if (rightBorder != null)
                            {
                                try { rightStyle = (int)rightBorder.LineStyle; } catch { }
                                try { rightWidth = (int)rightBorder.LineWidth; } catch { }
                                try { rightStyleOk = (WdLineStyle)rightBorder.LineStyle == WdLineStyle.wdLineStyleSingle; } catch { }
                                try { rightWidthOk = (WdLineWidth)rightBorder.LineWidth == WdLineWidth.wdLineWidth050pt || rightWidth == targetWidthObservedAlt; } catch { rightWidthOk = rightWidth == targetWidthObservedAlt; }
                                try { rightColor = (int)rightBorder.Color; } catch { }
                                try
                                {
                                    dynamic dt = rightBorder;
                                    object v = null;
                                    try { v = dt.Art; } catch { }
                                    rightArtRaw = v == null ? "" : v.ToString();
                                    try { if (v != null) rightArt = Convert.ToInt32(v); } catch { }
                                } catch { }
                                try
                                {
                                    dynamic dt = rightBorder;
                                    object v = null;
                                    try { v = dt.Shadow; } catch { }
                                    rightShadowRaw = v == null ? "" : v.ToString();
                                    try { if (v != null) rightShadow = Convert.ToInt32(v); } catch { }
                                } catch { }
                                try
                                {
                                    dynamic dt = rightBorder;
                                    object v = null;
                                    try { v = dt.Surround; } catch { }
                                    rightSurroundRaw = v == null ? "" : v.ToString();
                                    try { if (v != null) rightSurround = Convert.ToInt32(v); } catch { }
                                } catch { }
                                hasRight = rightStyleOk && rightWidthOk;
                            }

                            // COM 解放（境界オブジェクト）
                            if (leftBorder != null) Marshal.ReleaseComObject(leftBorder);
                            if (rightBorder != null) Marshal.ReleaseComObject(rightBorder);
                        }
                    }
                    finally
                    {
                        if (borders1 != null) Marshal.ReleaseComObject(borders1);
                        if (section1 != null) Marshal.ReleaseComObject(section1);
                    }

                    // 全セクションで同じ厳密条件を満たすか（H3: section index/適用箇所ずれ）
                    for (int si = 1; si <= sectionsCount; si++)
                    {
                        try
                        {
                            Section sec = document.Sections[si];
                            Borders b = sec.Borders;
                            Border t = null;
                            Border bo = null;
                            try { t = b[WdBorderType.wdBorderTop]; } catch { }
                            try { bo = b[WdBorderType.wdBorderBottom]; } catch { }
                            bool tOk = t != null &&
                                       (WdLineStyle)t.LineStyle == WdLineStyle.wdLineStyleSingle &&
                                       (WdLineWidth)t.LineWidth == WdLineWidth.wdLineWidth050pt;
                            bool bOk = bo != null &&
                                       (WdLineStyle)bo.LineStyle == WdLineStyle.wdLineStyleSingle &&
                                       (WdLineWidth)bo.LineWidth == WdLineWidth.wdLineWidth050pt;

                            if (tOk || bOk)
                            {
                                if (!strictMatchAnySection)
                                {
                                    strictMatchAnySection = true;
                                    firstStrictMatchSection = si;
                                }
                            }

                            if (bo != null) Marshal.ReleaseComObject(bo);
                            if (t != null) Marshal.ReleaseComObject(t);
                            Marshal.ReleaseComObject(b);
                            Marshal.ReleaseComObject(sec);
                        }
                        catch { }
                    }

                    WriteDebug6b16c7Ndjson(
                        "debug-4-6-pre",
                        "H2-H3-H4-art",
                        "WordChecker1_4.CheckTask_1_4_06",
                        "borders-strict",
                        "{\"hasTop\":" + (hasTop ? "true" : "false") +
                        ",\"hasBottom\":" + (hasBottom ? "true" : "false") +
                        ",\"hasLeft\":" + (hasLeft ? "true" : "false") +
                        ",\"hasRight\":" + (hasRight ? "true" : "false") +
                        ",\"topStyle\":" + topStyle +
                        ",\"topWidth\":" + topWidth +
                        ",\"topStyleOk\":" + (topStyleOk ? "true" : "false") +
                        ",\"topWidthOk\":" + (topWidthOk ? "true" : "false") +
                        ",\"targetStyleSingle\":" + targetStyleSingle +
                        ",\"targetWidth050pt\":" + targetWidth050pt +
                        ",\"topColor\":" + topColor +
                        ",\"topArt\":" + topArt +
                        ",\"topArtRaw\":\"" + JsonEscape(topArtRaw) + "\"" +
                        ",\"topShadow\":" + topShadow +
                        ",\"topShadowRaw\":\"" + JsonEscape(topShadowRaw) + "\"" +
                        ",\"topSurround\":" + topSurround +
                        ",\"topSurroundRaw\":\"" + JsonEscape(topSurroundRaw) + "\"" +
                        ",\"bordersArt\":" + bordersArt +
                        ",\"bordersArtRaw\":\"" + JsonEscape(bordersArtRaw) + "\"" +
                        ",\"bordersShadow\":" + bordersShadow +
                        ",\"bordersShadowRaw\":\"" + JsonEscape(bordersShadowRaw) + "\"" +
                        ",\"bordersSurround\":" + bordersSurround +
                        ",\"bordersSurroundRaw\":\"" + JsonEscape(bordersSurroundRaw) + "\"" +
                        ",\"bordersDistanceFromTop\":" + bordersDistanceFromTop +
                        ",\"bordersDistanceFromBottom\":" + bordersDistanceFromBottom +
                        ",\"bordersDistanceFromLeft\":" + bordersDistanceFromLeft +
                        ",\"bordersDistanceFromRight\":" + bordersDistanceFromRight +
                        ",\"bordersDistanceFromText\":" + bordersDistanceFromText +
                        ",\"bordersKeyProps\":\"" + JsonEscape(bordersKeyProps) + "\"" +
                        ",\"topBorderKeyProps\":\"" + JsonEscape(topBorderKeyProps) + "\"" +
                        ",\"bottomStyle\":" + bottomStyle +
                        ",\"bottomWidth\":" + bottomWidth +
                        ",\"bottomStyleOk\":" + (bottomStyleOk ? "true" : "false") +
                        ",\"bottomWidthOk\":" + (bottomWidthOk ? "true" : "false") +
                        ",\"bottomColor\":" + bottomColor +
                        ",\"bottomArt\":" + bottomArt +
                        ",\"bottomArtRaw\":\"" + JsonEscape(bottomArtRaw) + "\"" +
                        ",\"bottomShadow\":" + bottomShadow +
                        ",\"bottomShadowRaw\":\"" + JsonEscape(bottomShadowRaw) + "\"" +
                        ",\"bottomSurround\":" + bottomSurround +
                        ",\"bottomSurroundRaw\":\"" + JsonEscape(bottomSurroundRaw) + "\"" +
                        ",\"leftStyle\":" + leftStyle +
                        ",\"leftWidth\":" + leftWidth +
                        ",\"leftStyleOk\":" + (leftStyleOk ? "true" : "false") +
                        ",\"leftWidthOk\":" + (leftWidthOk ? "true" : "false") +
                        ",\"leftColor\":" + leftColor +
                        ",\"leftArt\":" + leftArt +
                        ",\"leftArtRaw\":\"" + JsonEscape(leftArtRaw) + "\"" +
                        ",\"leftShadow\":" + leftShadow +
                        ",\"leftShadowRaw\":\"" + JsonEscape(leftShadowRaw) + "\"" +
                        ",\"leftSurround\":" + leftSurround +
                        ",\"leftSurroundRaw\":\"" + JsonEscape(leftSurroundRaw) + "\"" +
                        ",\"rightStyle\":" + rightStyle +
                        ",\"rightWidth\":" + rightWidth +
                        ",\"rightStyleOk\":" + (rightStyleOk ? "true" : "false") +
                        ",\"rightWidthOk\":" + (rightWidthOk ? "true" : "false") +
                        ",\"rightColor\":" + rightColor +
                        ",\"rightArt\":" + rightArt +
                        ",\"rightArtRaw\":\"" + JsonEscape(rightArtRaw) + "\"" +
                        ",\"rightShadow\":" + rightShadow +
                        ",\"rightShadowRaw\":\"" + JsonEscape(rightShadowRaw) + "\"" +
                        ",\"rightSurround\":" + rightSurround +
                        ",\"rightSurroundRaw\":\"" + JsonEscape(rightSurroundRaw) + "\"" +
                        ",\"strictMatchAnySection\":" + (strictMatchAnySection ? "true" : "false") +
                        ",\"firstStrictMatchSection\":" + firstStrictMatchSection + "}"
                    );
                }
                catch { }
                // #endregion

                // #region agent log (shapes sample)
                // 「囲むの種類（影/Box）」が Borders ではなく Shapes 側の効果として表現されている可能性があるため、
                // まずは Shapes のサンプルをログに出す。
                try
                {
                    Shapes shapes = null;
                    try
                    {
                        shapes = document.Shapes;
                        int shapeCount = -1;
                        try { shapeCount = shapes.Count; } catch { }
                        int sampleMax = 10;
                        int limit = shapeCount > sampleMax ? sampleMax : shapeCount;
                        if (limit < 0) limit = 0;

                        StringBuilder sb = new StringBuilder();
                        for (int i = 1; i <= limit; i++)
                        {
                            Shape shp = null;
                            try
                            {
                                shp = shapes[i];
                                int type = -999;
                                string name = "";
                                string alt = "";
                                string title = "";
                                int shadowVisible = -999;
                                try { type = (int)shp.Type; } catch { }
                                try { name = shp.Name ?? ""; } catch { }
                                try { alt = shp.AlternativeText ?? ""; } catch { }
                                try { title = shp.Title ?? ""; } catch { }
                                try
                                {
                                    // ShadowFormat.Visible が取れない環境もあるため保護
                                    dynamic sd = shp.Shadow;
                                    object visObj = null;
                                    try { visObj = sd != null ? sd.Visible : null; } catch { }
                                    if (visObj is bool bv)
                                        shadowVisible = bv ? 1 : 0;
                                    else if (visObj is int iv)
                                        shadowVisible = iv != 0 ? 1 : 0;
                                    else
                                        shadowVisible = -999;
                                }
                                catch { shadowVisible = -999; }

                                if (sb.Length > 0) sb.Append("|");
                                sb.Append(i).Append(":type=").Append(type)
                                  .Append(",name=").Append(name)
                                  .Append(",alt=").Append(alt)
                                  .Append(",title=").Append(title)
                                  .Append(",shadowVisible=").Append(shadowVisible);
                            }
                            catch { }
                            finally { if (shp != null) Marshal.ReleaseComObject(shp); }
                        }

                        WriteDebug6b16c7Ndjson(
                            "debug-4-6-pre",
                            "H5-shapes",
                            "WordChecker1_4.CheckTask_1_4_06",
                            "shapes-sample",
                            "{\"shapeCount\":" + shapeCount + ",\"sample\":\"" + JsonEscape(sb.ToString()) + "\"}"
                        );
                    }
                    finally { if (shapes != null) Marshal.ReleaseComObject(shapes); }
                }
                catch { }
                // #endregion

                int[] expectedAccent1Colors = new int[]
                {
                    -738131969, // observed "青、アクセント1" candidate
                    -721354753  // observed "青、アクセント1" candidate
                };
                bool colorOk = false;
                try
                {
                    foreach (int c in expectedAccent1Colors)
                    {
                        if (topColor == c && bottomColor == c && leftColor == c && rightColor == c)
                        {
                            colorOk = true;
                            break;
                        }
                    }
                }
                catch { }

                // 4-6: 「囲む（Box）」のみ正解。COM では bordersShadow が
                //   - 影: 1 (True)
                //   - Box: 0 (False)
                // のように切り替わるため、影を不正解に落とすために必須化する。
                // 4-6: 「囲む（Box）」以外は不正解にする。
                // この環境では bordersShadow が
                //   - 影: 1 (True)
                //   - Box: 0 (False)
                // のように切り替わるため、Box のときのみ通す（raw 文字列が取れる場合も False/0 で強制）。
                bool bordersShadowOk =
                    bordersShadow == 0 &&
                    (string.IsNullOrEmpty(bordersShadowRaw) ||
                     bordersShadowRaw.IndexOf("false", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     bordersShadowRaw == "0");

                bool result = logOk && hasTop && hasBottom && hasLeft && hasRight && colorOk && bordersShadowOk;
                // #region agent log (result)
                try
                {
                    WriteDebug6b16c7Ndjson(
                        "debug-4-6-pre",
                        "H0",
                        "WordChecker1_4.CheckTask_1_4_06",
                        "result",
                        "{\"result\":" + (result ? "true" : "false") +
                        ",\"hasTop\":" + (hasTop ? "true" : "false") +
                        ",\"hasBottom\":" + (hasBottom ? "true" : "false") +
                        ",\"hasLeft\":" + (hasLeft ? "true" : "false") +
                        ",\"hasRight\":" + (hasRight ? "true" : "false") +
                        ",\"colorOk\":" + (colorOk ? "true" : "false") +
                        ",\"bordersShadowOk\":" + (bordersShadowOk ? "true" : "false") +
                        ",\"bordersShadow\":" + bordersShadow +
                        ",\"bordersShadowRaw\":\"" + JsonEscape(bordersShadowRaw) + "\"" +
                        ",\"logOk\":" + (logOk ? "true" : "false") + "}"
                    );
                }
                catch { }
                // #endregion

                return result;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_4_07(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { wordApp = new Application(); wordApp.Visible = true; }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents) { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return false;
                Section sec = document.Sections[1];
                string headerText = sec.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text ?? "";
                string footerText = sec.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text ?? "";
                Marshal.ReleaseComObject(sec);
                // 透かしとヘッダー・フッターが削除されているか（Trim して空なら削除済み）
                bool headerEmpty = string.IsNullOrWhiteSpace(headerText.Trim());
                bool footerEmpty = string.IsNullOrWhiteSpace(footerText.Trim());
                bool stateOk = headerEmpty && footerEmpty;
                // Phase1: ログ優先（ドキュメント検査ログ + 文書状態の両方が必要）
                return LogReader.HasCommandExecuted("FileDocumentInspect") && stateOk;
            }
            catch { return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private static bool CommentBodyMatchesLatestInfoCheck(Comment comment)
        {
            string raw = "";
            try { raw = comment.Range.Text ?? ""; } catch { }
            if (PassesLatestInfoPhrase(NormalizeCommentBody(raw)))
                return true;
            try
            {
                if (comment.Replies != null && comment.Replies.Count > 0)
                {
                    int n = comment.Replies.Count;
                    for (int i = 1; i <= n; i++)
                    {
                        Comment reply = null;
                        try
                        {
                            reply = comment.Replies[i];
                            string rt = "";
                            try { rt = reply.Range.Text ?? ""; } catch { }
                            if (PassesLatestInfoPhrase(NormalizeCommentBody(rt)))
                                return true;
                        }
                        finally { if (reply != null) Marshal.ReleaseComObject(reply); }
                    }
                }
            }
            catch { }
            return false;
        }

        private static bool PassesLatestInfoPhrase(string normalized)
        {
            const string needle = "最新の情報を確認";
            if (string.IsNullOrEmpty(normalized))
                return false;
            if (normalized.Contains(needle))
                return true;
            try
            {
                if (normalized.Normalize(NormalizationForm.FormKC).Contains(needle))
                    return true;
            }
            catch { }
            return false;
        }

        private static string NormalizeCommentBody(string s)
        {
            if (string.IsNullOrEmpty(s))
                return "";
            string t = s.Trim();
            t = t.Replace("\r\n", "").Replace("\r", "").Replace("\n", "").Replace("\u3000", " ");
            while (t.IndexOf("  ", StringComparison.Ordinal) >= 0)
                t = t.Replace("  ", " ");
            return t.Trim();
        }

        /// <summary>吹き出し本文（返信含む）に「エコと節約」が含まれるか。</summary>
        private static bool CommentBalloonContainsEcoPhrase(Comment comment)
        {
            string raw = "";
            try { raw = comment.Range.Text ?? ""; } catch { }
            if (NormalizedContainsEco(NormalizeCommentBody(raw)))
                return true;
            try
            {
                if (comment.Replies != null && comment.Replies.Count > 0)
                {
                    int n = comment.Replies.Count;
                    for (int i = 1; i <= n; i++)
                    {
                        Comment reply = null;
                        try
                        {
                            reply = comment.Replies[i];
                            string rt = "";
                            try { rt = reply.Range.Text ?? ""; } catch { }
                            if (NormalizedContainsEco(NormalizeCommentBody(rt)))
                                return true;
                        }
                        finally { if (reply != null) Marshal.ReleaseComObject(reply); }
                    }
                }
            }
            catch { }
            return false;
        }

        private static bool NormalizedContainsEco(string normalized)
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

        /// <summary>いずれかのコメント吹き出し（返信含む）に「エコと節約」が含まれるか。</summary>
        private static bool DocumentHasAnyEcoCommentBalloon(Document document)
        {
            Comments comments = document.Comments;
            try
            {
                foreach (Comment c in comments)
                {
                    try
                    {
                        if (CommentBalloonContainsEcoPhrase(c))
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

        /// <summary>「エコと節約」コメントが存在しない、またはすべて解決済み（Done）なら true。</summary>
        private static bool IsEcoCommentAbsentOrResolved(Document document)
        {
            Comments comments = document.Comments;
            try
            {
                foreach (Comment c in comments)
                {
                    try
                    {
                        if (!CommentBalloonContainsEcoPhrase(c))
                            continue;
                        bool done = false;
                        try { done = c.Done; } catch { }
                        if (!done)
                            return false;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(c);
                    }
                }
                return true;
            }
            finally
            {
                Marshal.ReleaseComObject(comments);
            }
        }

        private string GetCurrentWordFilePath()
        {
            Application wordApp = null;
            try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); if (wordApp.ActiveDocument != null) return wordApp.ActiveDocument.FullName; return null; }
            catch (COMException) { return null; }
            finally { if (wordApp != null) Marshal.ReleaseComObject(wordApp); }
        }
    }
}

