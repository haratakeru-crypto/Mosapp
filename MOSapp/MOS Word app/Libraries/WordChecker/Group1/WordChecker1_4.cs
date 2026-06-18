using System;
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
                Range searchRange = WordFindHelper.DuplicateContent(document);
                Find find = searchRange.Find;
                WordFindHelper.ConfigureSafeFind(find, "1.生活の中でできるエコ活動");
                find.Execute();
                if (!find.Found)
                {
                    WordFindHelper.ConfigureSafeFind(find, "１.生活の中でできるエコ活動");
                    find.Execute();
                }
                if (!find.Found)
                {
                    WordFindHelper.ConfigureSafeFind(find, "生活の中でできるエコ活動");
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
                bool xmlContainsPhrase = false;
                bool replyOnSonoTa = false;
                try
                {
                    System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_02] WordOpenXML取得中...");
                    string xml = document.WordOpenXML;
                    if (!string.IsNullOrEmpty(xml))
                    {
                        xmlContainsPhrase = xml.Contains("前田先生に最終確認");
                        if (xmlContainsPhrase)
                            fileStateCheck = true;
                        replyOnSonoTa = DocumentHasMaedaReplyOnSonoTaComment(document);
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
                Range searchRange = WordFindHelper.DuplicateContent(document);
                Find find = searchRange.Find;
                WordFindHelper.ConfigureSafeFind(find, "エコと節約");
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
                bool logResolve = LogReader.HasTaskEvidence(4, 3, "ReviewResolveComment");
                bool logDelete = LogReader.HasTaskEvidence(4, 3, "ReviewDeleteComment");
                bool anyEcoComment = DocumentHasAnyEcoCommentBalloon(document);
                bool result = logResolve && logDelete && !anyEcoComment;

                System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_03] 終了。結果={result} (logResolve={logResolve}, logDelete={logDelete}, anyEcoComment={anyEcoComment})");
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
                bool lineSimple = WordWatermarkInspection.IsDocumentStyleSetLineSimple(document);
                bool lineStylish = WordWatermarkInspection.IsDocumentStyleSetLineStylish(document);
                bool logOk = LogReader.HasTaskEvidence(4, 4, "StyleSetLineSimple");

                bool result = lineSimple && !lineStylish && logOk;
                System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_04] 終了。結果={result} (lineSimple={lineSimple}, lineStylish={lineStylish}, logOk={logOk})");
                return result;
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

                string normalizedXml = "";
                try
                {
                    string xml = document.WordOpenXML;
                    normalizedXml = WordWatermarkInspection.NormalizeXml(xml);
                }
                catch { }

                // 1) 現在の文書: 透かし「サンプル２」（ヘッダー内「サンプル」）のみ ○
                if (WordWatermarkInspection.HasForbiddenWatermark(normalizedXml))
                    return false;
                if (WordWatermarkInspection.IsSample2Watermark(normalizedXml))
                    return true;
                if (WordWatermarkInspection.IsDraft1Watermark(normalizedXml))
                    return false;
                if (WordWatermarkInspection.IsDraft2HorizontalWatermark(normalizedXml))
                    return false;

                // 2) 4-7 後: 透かしなし + 4-5 で透かし挿入証跡 + 4-7 相当（ヘッダー/フッター空）
                bool cleared = WordWatermarkInspection.IsWatermarkClearedForTask47FollowUp(normalizedXml);
                bool evidenceWatermark = LogReader.HasTaskEvidence(4, 5, "Watermark");
                bool task47State = IsPrimaryHeaderFooterEmpty(document);
                return cleared && evidenceWatermark && task47State;




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
                int bordersShadow = -999;

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

                    Marshal.ReleaseComObject(right); Marshal.ReleaseComObject(left); Marshal.ReleaseComObject(bottom); Marshal.ReleaseComObject(top); Marshal.ReleaseComObject(borders1); Marshal.ReleaseComObject(section1);
                }

                bool colorOk = WordWatermarkInspection.ArePageBordersPureAccent3(document);

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

                // 文書検査はバージョンにより idMso が無効・非公開のため Ribbon フック不可。ヘッダー/フッターが空であることのみで判定する。
                return IsPrimaryHeaderFooterEmpty(document);
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_07] 例外: {ex.Message}"); return false; }
            finally { if (document != null) Marshal.ReleaseComObject(document); if (docs != null) Marshal.ReleaseComObject(docs); if (wordApp != null) Marshal.ReleaseComObject(wordApp); }
        }

        /// <summary>OpenXML で向きが取れない場合の補完: ヘッダー内「下書き」シェイプの Rotation で斜め/横を判定。</summary>
        private static bool TryIsDraft1WatermarkViaHeaderShapes(Document document)
        {
            if (document == null)
                return false;

            Section sec = null;
            HeaderFooter header = null;
            Shapes shapes = null;
            try
            {
                if (document.Sections.Count < 1)
                    return false;
                sec = document.Sections[1];
                header = sec.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary];
                shapes = header.Shapes;
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    Shape shape = null;
                    try
                    {
                        shape = shapes[i];
                        string text = "";
                        try
                        {
                            if (shape.TextFrame != null)
                                text = shape.TextFrame.TextRange?.Text ?? "";
                        }
                        catch { }

                        if (string.IsNullOrEmpty(text) || text.IndexOf("下書き", StringComparison.Ordinal) < 0)
                            continue;

                        double rotation = 0;
                        try { rotation = shape.Rotation; } catch { }

                        if (IsDiagonalShapeRotationDegrees(rotation))
                            return true;
                    }
                    finally
                    {
                        if (shape != null)
                            Marshal.ReleaseComObject(shape);
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (shapes != null) Marshal.ReleaseComObject(shapes);
                if (header != null) Marshal.ReleaseComObject(header);
                if (sec != null) Marshal.ReleaseComObject(sec);
            }
        }

        private static bool IsDiagonalShapeRotationDegrees(double rotation)
        {
            rotation = NormalizeShapeRotationDegrees(rotation);
            if (rotation < 2)
                return false;
            double[] diagonalAngles = { 45, 135, 225, 315 };
            foreach (double target in diagonalAngles)
            {
                double diff = Math.Abs(rotation - target);
                if (diff <= 10 || Math.Abs(diff - 360) <= 10)
                    return true;
            }
            return false;
        }

        private static double NormalizeShapeRotationDegrees(double rotation)
        {
            rotation %= 360;
            if (rotation < 0)
                rotation += 360;
            return rotation;
        }

        /// <summary>4-7 および 4-5 の 2 段目: 先頭セクションの primary ヘッダー/フッターが空か。</summary>
        private static bool IsPrimaryHeaderFooterEmpty(Document document)
        {
            if (document == null)
                return false;
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
            return string.IsNullOrEmpty(cleanH) && string.IsNullOrEmpty(cleanF);
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

        private static bool CommentScopeLooksLikeEcoHeading(string scopeText)
        {
            string n = NormalizeCommentBody(scopeText);
            if (string.IsNullOrEmpty(n))
                return false;
            return n.Contains("1.生活の中でできるエコ活動")
                || n.Contains("１.生活の中でできるエコ活動")
                || n.Contains("生活の中でできるエコ活動");
        }

        private static bool CommentScopeLooksLikeSonoTaHeading(string scopeText)
        {
            string n = NormalizeCommentBody(scopeText);
            return !string.IsNullOrEmpty(n) && n.Contains("その他");
        }

        private static bool DocumentHasMaedaReplyOnSonoTaComment(Document document)
        {
            Comments comments = null;
            try
            {
                comments = document.Comments;
                int count = comments.Count;
                for (int i = 1; i <= count; i++)
                {
                    Comment comment = null;
                    try
                    {
                        comment = comments[i];
                        Range scope = null;
                        string scopeText = "";
                        try
                        {
                            scope = comment.Scope;
                            scopeText = scope?.Text ?? "";
                        }
                        finally { if (scope != null) Marshal.ReleaseComObject(scope); }
                        if (!CommentScopeLooksLikeSonoTaHeading(scopeText))
                            continue;
                        Comments replies = null;
                        try
                        {
                            replies = comment.Replies;
                            if (replies == null)
                                continue;
                            int rc = replies.Count;
                            for (int r = 1; r <= rc; r++)
                            {
                                Comment reply = null;
                                try
                                {
                                    reply = replies[r];
                                    string rt = "";
                                    try { rt = reply.Range?.Text ?? ""; } catch { }
                                    if (NormalizeCommentBody(rt).Contains("前田先生に最終確認"))
                                        return true;
                                }
                                finally { if (reply != null) Marshal.ReleaseComObject(reply); }
                            }
                        }
                        finally { if (replies != null) Marshal.ReleaseComObject(replies); }
                    }
                    finally { if (comment != null) Marshal.ReleaseComObject(comment); }
                }
            }
            catch { }
            finally { if (comments != null) Marshal.ReleaseComObject(comments); }
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

        /// <summary>吹き出し本文（返信含む）、Scope、または「エコと節約」アンカー重なりで関連判定。</summary>
        private static bool CommentRelatesToEcoPhrase(Comment comment, Document document)
        {
            if (CommentBalloonContainsEcoPhrase(comment))
                return true;
            string scope = "";
            try { scope = comment.Scope?.Text ?? ""; } catch { }
            if (NormalizedContainsEco(NormalizeCommentBody(scope)))
                return true;
            return CommentAnchorOverlapsEco(comment, document);
        }

        private static string _ecoAnchorDocKey;
        private static int _ecoAnchorStart = -1;
        private static int _ecoAnchorEnd = -1;

        private static bool TryGetEcoPhraseAnchor(Document document, out int start, out int end)
        {
            start = end = -1;
            if (document == null)
                return false;
            string key = "";
            try { key = document.FullName ?? ""; } catch { }
            if (!string.IsNullOrEmpty(key) && key == _ecoAnchorDocKey && _ecoAnchorStart >= 0)
            {
                start = _ecoAnchorStart;
                end = _ecoAnchorEnd;
                return true;
            }
            Range searchRange = null;
            Find find = null;
            try
            {
                searchRange = WordFindHelper.DuplicateContent(document);
                find = searchRange.Find;
                WordFindHelper.ConfigureSafeFind(find, "エコと節約");
                if (!find.Execute())
                    return false;
                start = searchRange.Start;
                end = searchRange.End;
                _ecoAnchorDocKey = key;
                _ecoAnchorStart = start;
                _ecoAnchorEnd = end;
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (find != null) Marshal.ReleaseComObject(find);
                if (searchRange != null) Marshal.ReleaseComObject(searchRange);
            }
        }

        private static bool CommentAnchorOverlapsEco(Comment comment, Document document)
        {
            if (!TryGetEcoPhraseAnchor(document, out int ecoStart, out int ecoEnd))
                return false;
            Range scope = null;
            try
            {
                scope = comment.Scope;
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
                if (scope != null) Marshal.ReleaseComObject(scope);
            }
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
                        if (CommentRelatesToEcoPhrase(c, document))
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

        /// <summary>「エコと節約」コメントが解決済み（Done）で残っているか。</summary>
        private static bool DocumentHasResolvedEcoComment(Document document)
        {
            Comments comments = null;
            try
            {
                comments = document?.Comments;
                if (comments == null || comments.Count == 0)
                    return false;
                int count = comments.Count;
                for (int i = 1; i <= count; i++)
                {
                    Comment c = null;
                    try
                    {
                        c = comments[i];
                        if (!CommentRelatesToEcoPhrase(c, document))
                            continue;
                        bool done = false;
                        try { done = c.Done; } catch { }
                        if (done)
                            return true;
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
                        if (!CommentRelatesToEcoPhrase(c, document))
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

