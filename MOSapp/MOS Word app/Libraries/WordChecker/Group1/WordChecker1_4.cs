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
        public bool CheckTask_1_4_01() => RunCheck(CheckTask_1_4_01_Impl);
        public bool CheckTask_1_4_02() => RunCheck(CheckTask_1_4_02_Impl);
        public bool CheckTask_1_4_03() => RunCheck(CheckTask_1_4_03_Impl);
        public bool CheckTask_1_4_04() => RunCheck(CheckTask_1_4_04_Impl);
        public bool CheckTask_1_4_05() => RunCheck(CheckTask_1_4_05_Impl);
        public bool CheckTask_1_4_06() => RunCheck(CheckTask_1_4_06_Impl);
        public bool CheckTask_1_4_07() => RunCheck(CheckTask_1_4_07_Impl);

        private bool RunCheck(Func<string, bool> task)
        {
            try
            {
                string filePath = GetCurrentWordFilePath();
                if (string.IsNullOrEmpty(filePath))
                    return Miss(WordScoreExplanation.UnavailableText);
                return task(filePath);
            }
            catch
            {
                return Miss(WordScoreExplanation.UnavailableText);
            }
        }

        private bool CheckTask_1_4_01_Impl(string filePath)
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

                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

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
                    return Miss("見出し「1.生活の中でできるエコ活動」がありません。");
                }
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_01] コメント取得開始...");
                Comments comments = document.Comments;
                bool result = false;
                string otherBodyOnHeading = null;
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
                            if (otherBodyOnHeading == null && CommentScopeLooksLikeEcoHeading(SafeCommentScopeText(comment)))
                            {
                                string body = StripWordCommentControlChars(SafeCommentRangeText(comment));
                                if (!string.IsNullOrEmpty(body) && !IsExactCommentPhrase(body, "最新の情報を確認"))
                                    otherBodyOnHeading = body;
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
                if (result)
                    return true;
                if (!string.IsNullOrEmpty(otherBodyOnHeading))
                    return Miss($"見出し「1.生活の中でできるエコ活動」のコメントが「{Quote(otherBodyOnHeading)}」になっています。");
                return Miss("見出し「1.生活の中でできるエコ活動」に「最新の情報を確認」のコメントがありません。");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_01] 例外発生: {ex.Message}");
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
                if (docs != null) Marshal.ReleaseComObject(docs);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        private bool CheckTask_1_4_02_Impl(string filePath)
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
                    return Miss(WordScoreExplanation.UnavailableText);
                }

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_02] ロジック実行開始");
                // 合否は指定文言との完全一致（余分な空白は×）。Word 由来の段落記号のみ除去。
                // COM を優先し、取れない環境向けに OpenXML の <w:t> テキストも同じ厳密照合する。
                const string expectedReply = "前田先生に最終確認";
                bool fileStateCheck = DocumentHasExactReplyOnSonoTaComment(document, expectedReply);
                if (!fileStateCheck)
                {
                    try
                    {
                        System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_02] WordOpenXML取得中...");
                        string xml = WordOpenXmlSession.Get(document);
                        fileStateCheck = XmlHasExactWtText(xml, expectedReply);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"    [CheckTask_1_4_02] XML解析エラー: {ex.Message}");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_02] 終了。結果={fileStateCheck}");
                if (fileStateCheck)
                    return true;

                string otherReply = TryGetFirstReplyBodyOnSonoTaComment(document);
                if (!string.IsNullOrEmpty(otherReply))
                    return Miss($"見出し「その他」のコメントの返信が「{Quote(otherReply)}」になっています。");
                return Miss("見出し「その他」のコメントに「前田先生に最終確認」の返信がありません。");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_02] 例外発生: {ex.Message}");
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
                if (docs != null) Marshal.ReleaseComObject(docs);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        private bool CheckTask_1_4_03_Impl(string filePath)
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

                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

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
                    return Miss("本文の「エコと節約」が見つかりません。");
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
                if (result)
                    return true;
                if (!logResolve)
                    WordScoreExplanation.Note("「エコと節約」のコメントを解決した記録がありません。");
                if (!logDelete)
                    WordScoreExplanation.Note("「エコと節約」のコメントを削除した記録がありません。");
                if (anyEcoComment)
                    WordScoreExplanation.Note("「エコと節約」に関するコメントが残っています。");
                if (logResolve || logDelete || anyEcoComment)
                    return false;
                return Miss("「エコと節約」のコメントについて、解決と削除が完了していません。");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_03] 例外発生: {ex.Message}");
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
                if (docs != null) Marshal.ReleaseComObject(docs);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        private bool CheckTask_1_4_04_Impl(string filePath)
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
                    return Miss(WordScoreExplanation.UnavailableText);
                }

                System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_04] ロジック実行開始");
                WordWatermarkInspection.StyleSetLineResult styleSet = WordWatermarkInspection.EvaluateStyleSetLine(document);
                bool lineSimple = styleSet.IsSimple;
                bool logSimple = LogReader.HasTaskEvidence(4, 4, "StyleSetLineSimple");
                // 過去の StyleSetLineStylish は見ない。別セットのままなら lineSimple が false で ×。
                bool result = lineSimple && logSimple;
                System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_04] 終了。結果={result} (lineSimple={lineSimple}, logSimple={logSimple})");
                if (result)
                    return true;
                if (styleSet.IsStylish)
                    return Miss("スタイルセットが「線（スタイリッシュ）」になっています。");
                if (!lineSimple)
                    return Miss("スタイルセットが「線（シンプル）」になっていません。");
                return Miss("スタイルセット「線（シンプル）」について、必要な操作の記録がありません。");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_04] 例外発生: {ex.Message}");
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
                if (docs != null) Marshal.ReleaseComObject(docs);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        private bool CheckTask_1_4_05_Impl(string filePath)
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
                    return Miss(WordScoreExplanation.UnavailableText);
                }

                string normalizedXml = "";
                try
                {
                    string xml = WordOpenXmlSession.Get(document);
                    normalizedXml = WordWatermarkInspection.NormalizeXml(xml);
                }
                catch { }

                // 1) 現在の文書: 透かし「サンプル２」（ヘッダー内「サンプル」）のみ ○
                if (WordWatermarkInspection.HasForbiddenWatermark(normalizedXml))
                {
                    string forbidden = FindForbiddenWatermarkKeyword(normalizedXml);
                    if (!string.IsNullOrEmpty(forbidden))
                        return Miss($"透かしが「{forbidden}」になっています。");
                    return Miss("指定以外の透かしが入っています。");
                }
                if (WordWatermarkInspection.IsSample2Watermark(normalizedXml))
                    return true;
                if (WordWatermarkInspection.IsDraft1Watermark(normalizedXml))
                    return Miss("透かしが「下書き1」になっています。");
                if (WordWatermarkInspection.IsDraft2HorizontalWatermark(normalizedXml))
                    return Miss("透かしが「下書き2」になっています。");

                // 2) 4-7 後: 透かしなし + 4-5 で透かし挿入証跡 + 4-7 相当（ヘッダー/フッター空）
                bool cleared = WordWatermarkInspection.IsWatermarkClearedForTask47FollowUp(normalizedXml);
                bool evidenceWatermark = LogReader.HasTaskEvidence(4, 5, "Watermark");
                bool task47State = IsPrimaryHeaderFooterEmpty(document);
                if (cleared && evidenceWatermark && task47State)
                    return true;
                return Miss("透かし「サンプル２」がありません。");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_05] 例外発生: {ex.Message}");
                return Miss(WordScoreExplanation.UnavailableText);
            }
            finally
            {
                if (document != null) Marshal.ReleaseComObject(document);
                if (docs != null) Marshal.ReleaseComObject(docs);
                if (wordApp != null) Marshal.ReleaseComObject(wordApp);
            }
        }

        private bool CheckTask_1_4_06_Impl(string filePath)
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
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                bool logOk = LogReader.HasTaskEvidence(4, 6, "PageBorders");
                bool hasTop = false, hasBottom = false, hasLeft = false, hasRight = false;
                int bordersShadow = -999;
                string wrongWidthLabel = null;
                bool anySidePresent = false;

                if (document.Sections.Count >= 1)
                {
                    Section section1 = document.Sections[1];
                    Borders borders1 = section1.Borders;
                    try { dynamic db = borders1; bordersShadow = Convert.ToInt32(db.Shadow); } catch { }
                    Border top = borders1[WdBorderType.wdBorderTop];
                    Border bottom = borders1[WdBorderType.wdBorderBottom];
                    Border left = borders1[WdBorderType.wdBorderLeft];
                    Border right = borders1[WdBorderType.wdBorderRight];

                    hasTop = IsPageBorderSide150Single(top);
                    hasBottom = IsPageBorderSide150Single(bottom);
                    hasLeft = IsPageBorderSide150Single(left);
                    hasRight = IsPageBorderSide150Single(right);

                    if (!hasTop) wrongWidthLabel = DescribePageBorderSideMismatch(top) ?? wrongWidthLabel;
                    if (!hasBottom) wrongWidthLabel = DescribePageBorderSideMismatch(bottom) ?? wrongWidthLabel;
                    if (!hasLeft) wrongWidthLabel = DescribePageBorderSideMismatch(left) ?? wrongWidthLabel;
                    if (!hasRight) wrongWidthLabel = DescribePageBorderSideMismatch(right) ?? wrongWidthLabel;
                    anySidePresent =
                        IsPageBorderSideVisible(top) || IsPageBorderSideVisible(bottom)
                        || IsPageBorderSideVisible(left) || IsPageBorderSideVisible(right);

                    Marshal.ReleaseComObject(right); Marshal.ReleaseComObject(left); Marshal.ReleaseComObject(bottom); Marshal.ReleaseComObject(top); Marshal.ReleaseComObject(borders1); Marshal.ReleaseComObject(section1);
                }

                bool colorOk = WordWatermarkInspection.ArePageBordersPureAccent3(document);

                bool result = logOk && hasTop && hasBottom && hasLeft && hasRight && colorOk && bordersShadow == 0;
                System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_06] 終了。結果={result}");
                if (result)
                    return true;

                bool sidesOk = hasTop && hasBottom && hasLeft && hasRight;
                bool noted = false;
                if (!anySidePresent)
                {
                    WordScoreExplanation.Note("ページ罫線が設定されていません。");
                    noted = true;
                }
                else if (!sidesOk && !string.IsNullOrEmpty(wrongWidthLabel))
                {
                    WordScoreExplanation.Note($"ページ罫線の線の太さが「{wrongWidthLabel}」になっています。");
                    noted = true;
                }
                else if (!sidesOk)
                {
                    WordScoreExplanation.Note("上下左右のページ罫線が「1.5pt」の実線になっていません。");
                    noted = true;
                }
                if (anySidePresent && sidesOk && !colorOk)
                {
                    WordScoreExplanation.Note("ページ罫線の色が「オリーブ、アクセント3」になっていません。");
                    noted = true;
                }
                if (anySidePresent && bordersShadow != 0 && bordersShadow != -999)
                {
                    WordScoreExplanation.Note("ページ罫線に影が付いています。");
                    noted = true;
                }
                if (!logOk)
                {
                    WordScoreExplanation.Note("ページ罫線について、必要な操作の記録がありません。");
                    noted = true;
                }
                if (noted)
                    return false;
                return Miss("ページ罫線が指定どおりになっていません。");
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_06] 例外: {ex.Message}"); return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); if (docs != null) Marshal.ReleaseComObject(docs); if (wordApp != null) Marshal.ReleaseComObject(wordApp); }
        }

        private bool CheckTask_1_4_07_Impl(string filePath)
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
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                // 文書検査はバージョンにより idMso が無効・非公開のため Ribbon フック不可。ヘッダー/フッターが空であることのみで判定する。
                if (IsPrimaryHeaderFooterEmpty(document))
                    return true;

                string normalizedXml = "";
                try
                {
                    normalizedXml = WordWatermarkInspection.NormalizeXml(WordOpenXmlSession.Get(document));
                }
                catch { }

                bool hasWatermark =
                    WordWatermarkInspection.IsSample2Watermark(normalizedXml)
                    || WordWatermarkInspection.IsDraft1Watermark(normalizedXml)
                    || WordWatermarkInspection.IsDraft2HorizontalWatermark(normalizedXml)
                    || WordWatermarkInspection.HasForbiddenWatermark(normalizedXml)
                    || (!string.IsNullOrEmpty(normalizedXml)
                        && normalizedXml.IndexOf("下書き", StringComparison.Ordinal) >= 0);

                if (hasWatermark)
                    return Miss("透かし文字またはヘッダー・フッターが残っています。");
                return Miss("ヘッダーまたはフッターが残っています。");
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"!!! [CheckTask_1_4_07] 例外: {ex.Message}"); return Miss(WordScoreExplanation.UnavailableText); }
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

        private static bool DocumentHasExactReplyOnSonoTaComment(Document document, string expectedPhrase)
        {
            Comments comments = null;
            try
            {
                comments = document?.Comments;
                if (comments == null)
                    return false;
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
                                    if (IsExactCommentPhrase(rt, expectedPhrase))
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

        /// <summary>
        /// 見出し「その他」コメントへの返信のうち、正解文でない最初の本文を返す。
        /// 返信が無い／COM で取れない／すべて正解文のときは null。
        /// </summary>
        private static string TryGetFirstReplyBodyOnSonoTaComment(Document document)
        {
            const string expectedReply = "前田先生に最終確認";
            Comments comments = null;
            try
            {
                comments = document?.Comments;
                if (comments == null)
                    return null;
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
                                    string body = StripWordCommentControlChars(rt);
                                    if (string.IsNullOrEmpty(body))
                                        continue;
                                    if (IsExactCommentPhrase(body, expectedReply))
                                        continue;
                                    return body;
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
            return null;
        }

        private static bool CommentBodyMatchesLatestInfoCheck(Comment comment)
        {
            const string expected = "最新の情報を確認";
            string raw = "";
            try { raw = comment.Range.Text ?? ""; } catch { }
            if (IsExactCommentPhrase(raw, expected))
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
                            if (IsExactCommentPhrase(rt, expected))
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

        /// <summary>
        /// 指定文言との完全一致。Word 由来の段落記号などは除去するが、空白は残す（余分な空白は×）。
        /// </summary>
        private static bool IsExactCommentPhrase(string raw, string expectedPhrase)
        {
            if (string.IsNullOrEmpty(expectedPhrase))
                return false;
            string actual = StripWordCommentControlChars(raw ?? "");
            if (actual == expectedPhrase)
                return true;
            try
            {
                return actual.Normalize(NormalizationForm.FormKC) == expectedPhrase.Normalize(NormalizationForm.FormKC);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>OpenXML の各 &lt;w:t&gt; テキストが指定文言と完全一致するか（空白付きは不一致）。</summary>
        private static bool XmlHasExactWtText(string xml, string expectedPhrase)
        {
            if (string.IsNullOrEmpty(xml) || string.IsNullOrEmpty(expectedPhrase))
                return false;
            MatchCollection matches = Regex.Matches(
                xml,
                @"<w:t(?:\s[^>]*)?>([^<]*)</w:t>",
                RegexOptions.CultureInvariant);
            foreach (Match match in matches)
            {
                if (!match.Success || match.Groups.Count < 2)
                    continue;
                if (IsExactCommentPhrase(match.Groups[1].Value, expectedPhrase))
                    return true;
            }
            return false;
        }

        /// <summary>Word がコメント本文に付けやすい制御文字・段落記号だけ除去。空白は残す。</summary>
        private static string StripWordCommentControlChars(string s)
        {
            if (string.IsNullOrEmpty(s))
                return "";
            var sb = new StringBuilder(s.Length);
            foreach (char c in s)
            {
                if (c == '\r' || c == '\n' || c == '\a' || c == '\v' || c == '\f' || c == '\u0007')
                    continue;
                sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>スコープ判定など寛容照合用。4-1/4-2 の合否には使わない。</summary>
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

        private static bool IsPageBorderSide150Single(Border border)
        {
            if (border == null)
                return false;
            try
            {
                return (WdLineStyle)border.LineStyle == WdLineStyle.wdLineStyleSingle
                    && (WdLineWidth)border.LineWidth == WdLineWidth.wdLineWidth150pt;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsPageBorderSideVisible(Border border)
        {
            if (border == null)
                return false;
            try
            {
                return (WdLineStyle)border.LineStyle != WdLineStyle.wdLineStyleNone;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>1.5pt 実線でない側について、太さが取れればその表示名を返す。</summary>
        private static string DescribePageBorderSideMismatch(Border border)
        {
            if (border == null)
                return null;
            try
            {
                WdLineStyle style = (WdLineStyle)border.LineStyle;
                if (style == WdLineStyle.wdLineStyleNone)
                    return null;
                if (style != WdLineStyle.wdLineStyleSingle)
                    return null;
                WdLineWidth width = (WdLineWidth)border.LineWidth;
                if (width == WdLineWidth.wdLineWidth150pt)
                    return null;
                return FormatLineWidth(width);
            }
            catch
            {
                return null;
            }
        }

        private static string FormatLineWidth(WdLineWidth width)
        {
            switch (width)
            {
                case WdLineWidth.wdLineWidth025pt: return "0.25pt";
                case WdLineWidth.wdLineWidth050pt: return "0.5pt";
                case WdLineWidth.wdLineWidth075pt: return "0.75pt";
                case WdLineWidth.wdLineWidth100pt: return "1pt";
                case WdLineWidth.wdLineWidth150pt: return "1.5pt";
                case WdLineWidth.wdLineWidth225pt: return "2.25pt";
                case WdLineWidth.wdLineWidth300pt: return "3pt";
                case WdLineWidth.wdLineWidth450pt: return "4.5pt";
                case WdLineWidth.wdLineWidth600pt: return "6pt";
                default: return null;
            }
        }

        private static string FindForbiddenWatermarkKeyword(string normalizedXml)
        {
            if (string.IsNullOrEmpty(normalizedXml))
                return null;
            string[] keywords = { "社外秘", "至急", "案内", "転送厳禁", "CONFIDENTIAL", "URGENT", "DO NOT COPY" };
            foreach (string keyword in keywords)
            {
                if (normalizedXml.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    return keyword;
            }
            return null;
        }

        private static string SafeCommentScopeText(Comment comment)
        {
            try
            {
                Range scope = comment.Scope;
                try { return scope?.Text ?? ""; }
                finally { if (scope != null) Marshal.ReleaseComObject(scope); }
            }
            catch
            {
                return "";
            }
        }

        private static string SafeCommentRangeText(Comment comment)
        {
            try { return comment.Range?.Text ?? ""; }
            catch { return ""; }
        }

        private static string Quote(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "（空）";
            string text = value.Replace("\r", "").Replace("\n", " ");
            const int maxLen = 40;
            if (text.Length <= maxLen)
                return text;
            return text.Substring(0, maxLen) + "…";
        }

        private static bool Miss(string reason)
        {
            WordScoreExplanation.Note(reason);
            return false;
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

