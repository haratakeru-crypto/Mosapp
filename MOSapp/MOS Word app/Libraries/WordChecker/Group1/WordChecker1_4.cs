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
            System.Diagnostics.Debug.WriteLine(">>> [CheckTask_1_4_01] 開始");
            Application wordApp = null;
            Documents docs = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { wordApp = new Application(); wordApp.Visible = true; }

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_01] 文書取得中...");
                string fileName = System.IO.Path.GetFileName(filePath);
                docs = wordApp.Documents;
                int docCount = 0;
                try { docCount = docs.Count; } catch { }

                for (int i = 1; i <= docCount; i++)
                {
                    Document doc = null;
                    try
                    {
                        doc = docs[i];
                        if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                            doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                        {
                            document = doc;
                            break;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (doc != null && doc != document) Marshal.ReleaseComObject(doc);
                    }
                }

                if (document == null) return false;

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_01] ロジック実行開始");
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
                    System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_01] 指定のテキストが見つかりません。");
                    return false;
                }
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_01] コメント取得開始...");
                Comments comments = document.Comments;
                bool result = false;
                try
                {
                    int count = comments.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        Comment comment = null;
                        try
                        {
                            comment = comments[i];
                            if (CommentBodyMatchesLatestInfoCheck(comment))
                            {
                                result = true;
                                break;
                            }
                        }
                        finally
                        {
                            if (comment != null) Marshal.ReleaseComObject(comment);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(comments);
                }
                System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_01] 終了。結果={result}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_01] 例外発生: {ex.Message}");
                return false;
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
                if (docs != null) Marshal.ReleaseComObject(docs);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        private bool CheckTask_1_4_02(string filePath)
        {
            System.Diagnostics.Debug.WriteLine(">>> [CheckTask_1_4_02] 開始");
            Application wordApp = null;
            Documents docs = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { wordApp = new Application(); wordApp.Visible = true; }

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_02] 文書取得中...");
                string fileName = System.IO.Path.GetFileName(filePath);
                docs = wordApp.Documents;
                int docCount = 0;
                try { docCount = docs.Count; } catch { }

                for (int i = 1; i <= docCount; i++)
                {
                    Document doc = null;
                    try
                    {
                        doc = docs[i];
                        if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                            doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                        {
                            document = doc;
                            break;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (doc != null && doc != document) Marshal.ReleaseComObject(doc);
                    }
                }

                if (document == null)
                {
                    System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_02] 文書が見つかりません。");
                    return false;
                }

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_02] ロジック実行開始");
                // コメント返信は Word の有効な idMso が環境により異なり Ribbon フック不可のため、ログは使わず WordOpenXML のみで判定する。

                // 【最重要】COMの Comments コレクションはモダンコメント環境で不安定なため使用しない。
                // 文書全体の WordOpenXML を取得し、XML内のテキストノードから判定する。
                bool fileStateCheck = false;
                try
                {
                    System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_02] WordOpenXML取得中...");
                    string xml = document.WordOpenXML;
                    if (!string.IsNullOrEmpty(xml))
                    {
                        if (xml.Contains("前田先生に最終確認"))
                        {
                            fileStateCheck = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"    [CheckTask_1_4_02] XML解析エラー: {ex.Message}");
                }

                bool result = fileStateCheck;
                System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_02] 終了。結果={result}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_02] 例外発生: {ex.Message}");
                return false;
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
                if (docs != null) Marshal.ReleaseComObject(docs);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        private bool CheckTask_1_4_03(string filePath)
        {
            System.Diagnostics.Debug.WriteLine(">>> [CheckTask_1_4_03] 開始");
            Application wordApp = null;
            Documents docs = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { wordApp = new Application(); wordApp.Visible = true; }

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_03] 文書取得中...");
                string fileName = System.IO.Path.GetFileName(filePath);
                docs = wordApp.Documents;
                int docCount = 0;
                try { docCount = docs.Count; } catch { }

                for (int i = 1; i <= docCount; i++)
                {
                    Document doc = null;
                    try
                    {
                        doc = docs[i];
                        if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                            doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                        {
                            document = doc;
                            break;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (doc != null && doc != document) Marshal.ReleaseComObject(doc);
                    }
                }

                if (document == null) return false;

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_03] ロジック実行開始");
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
                    System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_03] 'エコと節約' が見つかりません。");
                    return false;
                }
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);

                // 吹き出しに「エコと節約」が一度も無い: 未着手 / 削除 / 解決（Word によって吹き出しが消える）の区別。削除は ReviewDeleteComment、解決は VSTO ポーリングの ReviewResolveComment ログで補足。
                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_03] コメント状態チェック中...");
                bool anyEcoBalloon = DocumentHasAnyEcoCommentBalloon(document);
                // 初期状態と完了後が見分けづらいため、個別リセット後の旧ログ誤判定を避け証跡のみ参照
                bool logDelete = LogReader.HasTaskEvidence(4, 3, "ReviewDeleteComment");
                bool logResolve = LogReader.HasTaskEvidence(4, 3, "ReviewResolveComment");
                bool resolvedStateOk = IsEcoCommentAbsentOrResolved(document);

                bool result;
                if (!anyEcoBalloon)
                    result = logDelete || logResolve;
                else
                    result = resolvedStateOk;

                System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_03] 終了。結果={result}");
                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_03] 例外発生: {ex.Message}");
                return false;
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
                if (docs != null) Marshal.ReleaseComObject(docs);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        private bool CheckTask_1_4_04(string filePath)
        {
            System.Diagnostics.Debug.WriteLine(">>> [CheckTask_1_4_04] 開始");
            Application wordApp = null;
            Documents docs = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { wordApp = new Application(); wordApp.Visible = true; }

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_04] 文書取得中...");
                string fileName = System.IO.Path.GetFileName(filePath);
                docs = wordApp.Documents;
                int docCount = 0;
                try { docCount = docs.Count; } catch { }

                for (int i = 1; i <= docCount; i++)
                {
                    Document doc = null;
                    try
                    {
                        doc = docs[i];
                        if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                            doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                        {
                            document = doc;
                            break;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (doc != null && doc != document) Marshal.ReleaseComObject(doc);
                    }
                }

                if (document == null)
                {
                    System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_04] 文書が見つかりません。");
                    return false;
                }

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_04] ロジック実行開始");
                // スタイルセット「線（シンプル）」を狭く判定（見出し1下罫線 + ログ必須）
                Style headingStyle = null;
                Borders borders = null;
                try
                {
                    try { headingStyle = document.Styles["見出し 1"]; } catch { headingStyle = document.Styles["Heading 1"]; }
                    if (headingStyle == null)
                    {
                        System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_04] '見出し 1' スタイルが見つかりません。");
                        return false;
                    }
                    borders = headingStyle.ParagraphFormat.Borders;
                    Border bottomBorder = null;
                    try { bottomBorder = borders[WdBorderType.wdBorderBottom]; } catch { }
                    
                    int lineStyle = -999;
                    int lineWidth = -999;
                    if (bottomBorder != null)
                    {
                        try { lineStyle = (int)bottomBorder.LineStyle; } catch { }
                        try { lineWidth = (int)bottomBorder.LineWidth; } catch { }
                    }

                    bool hasLine = bottomBorder != null &&
                                   lineStyle == (int)WdLineStyle.wdLineStyleSingle &&
                                   lineWidth == (int)WdLineWidth.wdLineWidth050pt;
                    bool logOk = LogReader.HasTaskEvidence(4, 4, "StyleSetLineSimple");

                    bool result = hasLine && logOk;
                    System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_04] 終了。結果={result} (hasLine={hasLine}, logOk={logOk}, style={lineStyle}, width={lineWidth})");
                    return result;
                }
                finally
                {
                    if (borders != null) Marshal.ReleaseComObject(borders);
                    if (headingStyle != null) Marshal.ReleaseComObject(headingStyle);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_04] 例外発生: {ex.Message}");
                return false;
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
                if (docs != null) Marshal.ReleaseComObject(docs);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        private bool CheckTask_1_4_05(string filePath)
        {
            System.Diagnostics.Debug.WriteLine(">>> [CheckTask_1_4_05] 開始");
            Application wordApp = null;
            Documents docs = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { wordApp = new Application(); wordApp.Visible = true; }

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_05] 文書取得中...");
                string fileName = System.IO.Path.GetFileName(filePath);
                docs = wordApp.Documents;
                int docCount = 0;
                try { docCount = docs.Count; } catch { }

                for (int i = 1; i <= docCount; i++)
                {
                    Document doc = null;
                    try
                    {
                        doc = docs[i];
                        if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) ||
                            doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase))
                        {
                            document = doc;
                            break;
                        }
                    }
                    catch { }
                    finally
                    {
                        if (doc != null && doc != document) Marshal.ReleaseComObject(doc);
                    }
                }

                if (document == null)
                {
                    System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_05] 文書が見つかりません。");
                    return false;
                }

                bool found = false;
                try
                {
                    string xml = document.WordOpenXML;
                    if (!string.IsNullOrEmpty(xml))
                    {
                        if (xml.Normalize(NormalizationForm.FormKC).Contains("下書き"))
                        {
                            found = true;
                        }
                    }
                }
                catch { }

                bool logWatermark = LogReader.HasAnyTaskEvidence(4, 5,
                    "Watermark", "WatermarkMenu", "GalleryWatermark", "WatermarkCustomDialog");

                return found || logWatermark;






            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_05] 例外発生: {ex.Message}");
                return false;
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
                if (docs != null) Marshal.ReleaseComObject(docs);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        private bool CheckTask_1_4_06(string filePath)
        {
            System.Diagnostics.Debug.WriteLine(">>> [CheckTask_1_4_06] 開始");
            Application wordApp = null;
            Documents docs = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { wordApp = new Application(); wordApp.Visible = true; }
                string fileName = System.IO.Path.GetFileName(filePath);
                docs = wordApp.Documents;
                int docCount = 0; try { docCount = docs.Count; } catch { }
                for (int i = 1; i <= docCount; i++)
                {
                    Document doc = null;
                    try { doc = docs[i]; if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                    catch { } finally { if (doc != null && doc != document) Marshal.ReleaseComObject(doc); }
                }
                if (document == null) return false;

                bool logOk = LogReader.HasTaskEvidence(4, 6, "PageBorders");
                bool hasTop = false, hasBottom = false, hasLeft = false, hasRight = false;
                int topColor = -999, bottomColor = -999, leftColor = -999, rightColor = -999, bordersShadow = -999;

                if (document.Sections.Count >= 1)
                {
                    Section section1 = document.Sections[1];
                    Borders borders1 = section1.Borders;
                    try { dynamic db = borders1; bordersShadow = Convert.ToInt32(db.Shadow); } catch { }
                    Border top = borders1[WdBorderType.wdBorderTop];
                    Border bottom = borders1[WdBorderType.wdBorderBottom];
                    Border left = borders1[WdBorderType.wdBorderLeft];
                    Border right = borders1[WdBorderType.wdBorderRight];

                    hasTop = top != null && (WdLineStyle)top.LineStyle == WdLineStyle.wdLineStyleSingle && (WdLineWidth)top.LineWidth == WdLineWidth.wdLineWidth150pt;
                    hasBottom = bottom != null && (WdLineStyle)bottom.LineStyle == WdLineStyle.wdLineStyleSingle && (WdLineWidth)bottom.LineWidth == WdLineWidth.wdLineWidth150pt;
                    hasLeft = left != null && (WdLineStyle)left.LineStyle == WdLineStyle.wdLineStyleSingle && (WdLineWidth)left.LineWidth == WdLineWidth.wdLineWidth150pt;
                    hasRight = right != null && (WdLineStyle)right.LineStyle == WdLineStyle.wdLineStyleSingle && (WdLineWidth)right.LineWidth == WdLineWidth.wdLineWidth150pt;

                    try { topColor = (int)top.Color; } catch { }
                    try { bottomColor = (int)bottom.Color; } catch { }
                    try { leftColor = (int)left.Color; } catch { }
                    try { rightColor = (int)right.Color; } catch { }

                    Marshal.ReleaseComObject(right); Marshal.ReleaseComObject(left); Marshal.ReleaseComObject(bottom); Marshal.ReleaseComObject(top); Marshal.ReleaseComObject(borders1); Marshal.ReleaseComObject(section1);
                }

                int[] expectedAccent1Colors = new int[] { -738131969, -721354753 };
                bool colorOk = false;
                foreach (int c in expectedAccent1Colors) { if (topColor == c && bottomColor == c && leftColor == c && rightColor == c) { colorOk = true; break; } }

                bool result = logOk && hasTop && hasBottom && hasLeft && hasRight && colorOk && bordersShadow == 0;
                System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_06] 終了。結果={result}");
                return result;
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_06] 例外: {ex.Message}"); return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); if (docs != null) Marshal.ReleaseComObject(docs); if (wordApp != null) Marshal.ReleaseComObject(wordApp); }
        }

        private bool CheckTask_1_4_07(string filePath)
        {
            System.Diagnostics.Debug.WriteLine(">>> [CheckTask_1_4_07] 開始");
            Application wordApp = null;
            Documents docs = null;
            Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); }
                catch { wordApp = new Application(); wordApp.Visible = true; }
                string fileName = System.IO.Path.GetFileName(filePath);
                docs = wordApp.Documents;
                int docCount = 0; try { docCount = docs.Count; } catch { }
                for (int i = 1; i <= docCount; i++)
                {
                    Document doc = null;
                    try { doc = docs[i]; if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                    catch { } finally { if (doc != null && doc != document) Marshal.ReleaseComObject(doc); }
                }
                if (document == null) return false;

                string headerText = "";
                string footerText = "";
                try
                {
                    Section sec = document.Sections[1];
                    headerText = sec.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text ?? "";
                    footerText = sec.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text ?? "";
                    Marshal.ReleaseComObject(sec);
                }
                catch { }

                string cleanH = headerText.Replace("\r", "").Replace("\f", "").Trim();
                string cleanF = footerText.Replace("\r", "").Replace("\f", "").Trim();
                bool stateOk = string.IsNullOrEmpty(cleanH) && string.IsNullOrEmpty(cleanF);

                // 文書検査はバージョンにより idMso が無効・非公開のため Ribbon フック不可。ヘッダー/フッターが空であることのみで判定する。
                return stateOk;
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_07] 例外: {ex.Message}"); return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); if (docs != null) Marshal.ReleaseComObject(docs); if (wordApp != null) Marshal.ReleaseComObject(wordApp); }
        }

        private static readonly Regex DraftWatermarkLoose = new Regex(@"下書き[\s\u00A0\u200B\uFEFF]*[0-9１壱一ⅠⅠ①⑴⑺⑶]", RegexOptions.Compiled);

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

        private static string JsonEscape(string s)
        {
            if (s == null) return string.Empty;
            return s
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n")
                .Replace("\t", "\\t");
        }

        private static string GetDebug6b16c7LogPath()
        {
            try
            {
                var d = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory ?? "");
                for (int i = 0; i < 10 && d != null; i++)
                {
                    if (File.Exists(Path.Combine(d.FullName, "MOS Word app.sln")) && d.Parent != null)
                        return Path.Combine(d.Parent.FullName, "debug-6b16c7.log");
                    d = d.Parent;
                }
            }
            catch { }
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? "", "debug-6b16c7.log");
        }

        private static void WriteDebug6b16c7Ndjson(string runId, string hypothesisId, string location, string message, string dataJson)
        {
            try
            {
                long ts = (long)(DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)).TotalMilliseconds;
                string path = GetDebug6b16c7LogPath();
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


        /// <summary>コメントまたはその返信に指定したテキストが含まれるか判定する</summary>
        private static bool CommentBalloonContainsText(Comment comment, string searchText)
        {
            if (comment == null) return false;
            string raw = "";
            try { raw = comment.Range.Text ?? ""; } catch { }
            if (raw.Contains(searchText)) return true;

            Comments replies = null;
            try
            {
                replies = comment.Replies;
                if (replies != null)
                {
                    int n = replies.Count;
                    for (int i = 1; i <= n; i++)
                    {
                        Comment reply = null;
                        try
                        {
                            reply = replies[i];
                            string rt = "";
                            try { rt = reply.Range.Text ?? ""; } catch { }
                            if (rt.Contains(searchText)) return true;
                        }
                        finally { if (reply != null) Marshal.ReleaseComObject(reply); }
                    }
                }
            }
            catch { }
            finally { if (replies != null) Marshal.ReleaseComObject(replies); }
            return false;
        }

        private static bool CommentBodyMatchesLatestInfoCheck(Comment comment)
        {
            string raw = "";
            try { raw = comment.Range.Text ?? ""; } catch { }
            if (PassesLatestInfoPhrase(NormalizeCommentBody(raw)))
                return true;
            Comments replies = null;
            try
            {
                replies = comment.Replies;
                if (replies != null)
                {
                    int n = replies.Count;
                    for (int i = 1; i <= n; i++)
                    {
                        Comment reply = null;
                        try
                        {
                            reply = replies[i];
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
            finally { if (replies != null) Marshal.ReleaseComObject(replies); }
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
            Comments replies = null;
            try
            {
                replies = comment.Replies;
                if (replies != null)
                {
                    int n = replies.Count;
                    for (int i = 1; i <= n; i++)
                    {
                        Comment reply = null;
                        try
                        {
                            reply = replies[i];
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
            finally { if (replies != null) Marshal.ReleaseComObject(replies); }
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
            System.Diagnostics.Debug.WriteLine("    [DocumentHasAnyEcoCommentBalloon] 開始...");
            Comments comments = null;
            try
            {
                comments = document.Comments;
                if (comments == null) return false;
                int count = comments.Count;
                System.Diagnostics.Debug.WriteLine($"    [DocumentHasAnyEcoCommentBalloon] コメント数: {count}");
                for (int i = 1; i <= count; i++)
                {
                    Comment c = null;
                    try
                    {
                        c = comments[i];
                        if (CommentBalloonContainsEcoPhrase(c))
                        {
                            System.Diagnostics.Debug.WriteLine($"    [DocumentHasAnyEcoCommentBalloon] エコフレーズ発見 (index: {i})");
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"    [DocumentHasAnyEcoCommentBalloon] 個別コメント処理エラー (index: {i}): {ex.Message}");
                    }
                    finally
                    {
                        if (c != null) Marshal.ReleaseComObject(c);
                    }
                }
                return false;
            }
            finally
            {
                if (comments != null) Marshal.ReleaseComObject(comments);
            }
        }

        /// <summary>「エコと節約」コメントが存在しない、またはすべて解決済み（Done）なら true。</summary>
        private static bool IsEcoCommentAbsentOrResolved(Document document)
        {
            System.Diagnostics.Debug.WriteLine("    [IsEcoCommentAbsentOrResolved] 開始...");
            Comments comments = null;
            try
            {
                comments = document.Comments;
                if (comments == null) return true;
                int count = comments.Count;
                System.Diagnostics.Debug.WriteLine($"    [IsEcoCommentAbsentOrResolved] コメント数: {count}");
                for (int i = 1; i <= count; i++)
                {
                    Comment c = null;
                    try
                    {
                        c = comments[i];
                        if (!CommentBalloonContainsEcoPhrase(c))
                            continue;
                        bool done = false;
                        try { done = c.Done; } catch { }
                        if (!done)
                        {
                            System.Diagnostics.Debug.WriteLine($"    [IsEcoCommentAbsentOrResolved] 未解決コメント発見 (index: {i})");
                            return false;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"    [IsEcoCommentAbsentOrResolved] 個別コメント処理エラー (index: {i}): {ex.Message}");
                    }
                    finally
                    {
                        if (c != null) Marshal.ReleaseComObject(c);
                    }
                }
                return true;
            }
            finally
            {
                if (comments != null) Marshal.ReleaseComObject(comments);
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

