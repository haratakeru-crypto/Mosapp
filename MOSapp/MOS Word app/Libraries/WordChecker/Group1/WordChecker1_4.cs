using System;
using System.Collections.Generic;
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
                catch { return Miss(WordScoreExplanation.UnavailableText); }

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

                // モダンコメントでも安定するよう、合否は OpenXML のみ（COM Comments は使わない）。
                // 合否: comments* の本文完全一致 + document.xml の commentRange が指定見出し文言上。
                const string expected = "最新の情報を確認";
                const string anchorPhrase = "生活の中でできるエコ活動";
                string xml = "";
                try
                {
                    System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_01] WordOpenXML取得中...");
                    xml = WordOpenXmlSession.Get(document);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"    [CheckTask_1_4_01] XML取得エラー: {ex.Message}");
                }

                string failReason;
                bool result = TryPassCommentExactOnAnchor(
                    xml,
                    expected,
                    anchorPhrase,
                    requireReply: false,
                    out failReason);
                System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_01] 終了。結果={result}");
                if (result)
                    return true;
                return Miss(failReason
                    ?? "見出し「1.生活の中でできるエコ活動」に「最新の情報を確認」のコメントがありません。");
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
                catch { return Miss(WordScoreExplanation.UnavailableText); }

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

                // 合否: comments* の本文完全一致 + 返信(paraIdParent) + アンカーが「その他」。COM Comments は使わない。
                const string expectedReply = "前田先生に最終確認";
                const string anchorPhrase = "その他";
                string xml = "";
                try
                {
                    System.Diagnostics.Debug.WriteLine("    [CheckTask_1_4_02] WordOpenXML取得中...");
                    xml = WordOpenXmlSession.Get(document);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"    [CheckTask_1_4_02] XML解析エラー: {ex.Message}");
                }

                string failReason;
                bool fileStateCheck = TryPassCommentExactOnAnchor(
                    xml,
                    expectedReply,
                    anchorPhrase,
                    requireReply: true,
                    out failReason);
                System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_02] 終了。結果={fileStateCheck}");
                if (fileStateCheck)
                    return true;
                return Miss(failReason
                    ?? "見出し「その他」のコメントに「前田先生に最終確認」の返信がありません。");
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
                catch { return Miss(WordScoreExplanation.UnavailableText); }

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
                catch { return Miss(WordScoreExplanation.UnavailableText); }

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
                catch { return Miss(WordScoreExplanation.UnavailableText); }

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

                // 1) 現在の文書: 透かし「サンプル２」（「サンプル」＋横書き）のみ ○
                if (WordWatermarkInspection.IsSample2Watermark(normalizedXml))
                    return true;

                // 2) 4-7 後: 透かしなし + 4-5 で透かし挿入証跡 + 4-7 相当（H/F・透かしクリア）
                bool cleared = WordWatermarkInspection.IsWatermarkClearedForTask47FollowUp(normalizedXml);
                bool evidenceWatermark = LogReader.HasTaskEvidence(4, 5, "Watermark");
                bool task47State = IsTask47HeaderFooterWatermarkCleared(document, normalizedXml);
                if (cleared && evidenceWatermark && task47State)
                    return true;

                return Miss(BuildTask45FailReason(normalizedXml));
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
                catch { return Miss(WordScoreExplanation.UnavailableText); }
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
                bool topVisible = false, bottomVisible = false, leftVisible = false, rightVisible = false;

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

                    topVisible = IsPageBorderSideVisible(top);
                    bottomVisible = IsPageBorderSideVisible(bottom);
                    leftVisible = IsPageBorderSideVisible(left);
                    rightVisible = IsPageBorderSideVisible(right);

                    if (!hasTop) wrongWidthLabel = DescribePageBorderSideMismatch(top) ?? wrongWidthLabel;
                    if (!hasBottom) wrongWidthLabel = DescribePageBorderSideMismatch(bottom) ?? wrongWidthLabel;
                    if (!hasLeft) wrongWidthLabel = DescribePageBorderSideMismatch(left) ?? wrongWidthLabel;
                    if (!hasRight) wrongWidthLabel = DescribePageBorderSideMismatch(right) ?? wrongWidthLabel;

                    Marshal.ReleaseComObject(right); Marshal.ReleaseComObject(left); Marshal.ReleaseComObject(bottom); Marshal.ReleaseComObject(top); Marshal.ReleaseComObject(borders1); Marshal.ReleaseComObject(section1);
                }

                bool colorOk = WordWatermarkInspection.ArePageBordersPureAccent3(document);

                bool result = logOk && hasTop && hasBottom && hasLeft && hasRight && colorOk && bordersShadow == 0;
                System.Diagnostics.Debug.WriteLine($"<<< [CheckTask_1_4_06] 終了。結果={result}");
                if (result)
                    return true;

                bool anySidePresent = topVisible || bottomVisible || leftVisible || rightVisible;
                bool allSidesPresent = topVisible && bottomVisible && leftVisible && rightVisible;
                bool sidesOk = hasTop && hasBottom && hasLeft && hasRight;
                bool shadowOk = bordersShadow == 0;
                bool visualOk = sidesOk && colorOk && shadowOk;
                bool noted = false;

                if (!anySidePresent)
                {
                    // 未設定時は色・太さ・影・辺・操作記録は出さない
                    WordScoreExplanation.Note("ページ罫線が設定されていません。");
                    noted = true;
                }
                else
                {
                    // 辺・太さ・色・影は、ずれているものだけ複数行で出す
                    if (!allSidesPresent)
                    {
                        WordScoreExplanation.Note("上下左右のページ罫線が揃っていません。");
                        noted = true;
                    }
                    if (!string.IsNullOrEmpty(wrongWidthLabel))
                    {
                        WordScoreExplanation.Note($"ページ罫線の線の太さが「{wrongWidthLabel}」になっています。");
                        noted = true;
                    }
                    else if (allSidesPresent && !sidesOk)
                    {
                        WordScoreExplanation.Note("上下左右のページ罫線が「1.5pt」の実線になっていません。");
                        noted = true;
                    }
                    if (!colorOk)
                    {
                        WordScoreExplanation.Note("ページ罫線の色が「オリーブ、アクセント3」になっていません。");
                        noted = true;
                    }
                    if (bordersShadow != 0 && bordersShadow != -999)
                    {
                        WordScoreExplanation.Note("ページ罫線に影が付いています。");
                        noted = true;
                    }
                }

                // 操作記録は「見た目はすべて正解なのに ×」のときだけ出す
                if (visualOk && !logOk)
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
                catch { return Miss(WordScoreExplanation.UnavailableText); }
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

                // 文書検査は Ribbon フック不可。合否は OpenXML（ヘッダー/フッター本文 + 透かし）を主とし、
                // XML が取れないときだけ COM の primary H/F 空判定へフォールバックする。
                if (IsTask47HeaderFooterWatermarkCleared(document))
                    return true;

                return Miss("ヘッダー・フッターまたは透かし文字が残っています。");
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

        /// <summary>
        /// 4-7 / 4-5 第2段: ヘッダー・フッター本文と透かしが文書検査相当で消えているか。
        /// OpenXML を優先（COM Range.Text だけだと透かしシェイプ残留を見逃し得る）。
        /// </summary>
        private static bool IsTask47HeaderFooterWatermarkCleared(Document document, string normalizedXmlHint = null)
        {
            string xml = "";
            try { xml = WordOpenXmlSession.Get(document); } catch { }

            string normalized = !string.IsNullOrEmpty(normalizedXmlHint)
                ? normalizedXmlHint
                : WordWatermarkInspection.NormalizeXml(xml);

            if (!string.IsNullOrEmpty(xml) || !string.IsNullOrEmpty(normalized))
            {
                string xmlForParts = !string.IsNullOrEmpty(xml) ? xml : normalized;
                bool hfEmpty = AreHeaderFooterXmlPartsTextEmpty(xmlForParts);
                bool watermarkCleared = WordWatermarkInspection.IsWatermarkClearedForTask47FollowUp(
                    string.IsNullOrEmpty(normalized) ? WordWatermarkInspection.NormalizeXml(xmlForParts) : normalized);
                return hfEmpty && watermarkCleared;
            }

            // WordOpenXML が取れない環境のみ COM にフォールバック
            return IsPrimaryHeaderFooterEmptyCom(document);
        }

        /// <summary>Flat OPC / 正規化 XML 内の w:hdr・w:ftr から w:t を連結し、実質空か。</summary>
        private static bool AreHeaderFooterXmlPartsTextEmpty(string wordOpenXml)
        {
            if (string.IsNullOrEmpty(wordOpenXml))
                return false;

            foreach (Match block in Regex.Matches(
                wordOpenXml,
                @"<w:(?:hdr|ftr)\b[^>]*>.*?</w:(?:hdr|ftr)>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                string text = StripWordCommentControlChars(ExtractConcatenatedWtText(block.Value))
                    .Replace("\r", "")
                    .Replace("\f", "")
                    .Replace("\a", "")
                    .Trim();
                if (!string.IsNullOrEmpty(text))
                    return false;
            }

            return true;
        }

        /// <summary>COM: 先頭セクション primary ヘッダー/フッターの Range.Text が空か（フォールバック専用）。</summary>
        private static bool IsPrimaryHeaderFooterEmptyCom(Document document)
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

        /// <summary>
        /// comments.xml の本文完全一致 + document.xml の commentRange アンカー。
        /// requireReply 時は commentsExtended の paraIdParent（返信）も必須。
        /// </summary>
        private static bool TryPassCommentExactOnAnchor(
            string wordOpenXml,
            string expectedBody,
            string anchorPhrase,
            bool requireReply,
            out string failReason)
        {
            failReason = null;
            if (string.IsNullOrEmpty(wordOpenXml)
                || string.IsNullOrEmpty(expectedBody)
                || string.IsNullOrEmpty(anchorPhrase))
            {
                failReason = requireReply
                    ? $"見出し「{anchorPhrase}」のコメントに「{expectedBody}」の返信がありません。"
                    : $"見出し「1.{anchorPhrase}」に「{expectedBody}」のコメントがありません。";
                return false;
            }

            var comments = ParseCommentEntries(wordOpenXml);
            var rangeById = BuildCommentRangeTexts(wordOpenXml);
            var parentParaByChild = BuildCommentParaIdParentMap(wordOpenXml);
            var commentByParaId = new Dictionary<string, CommentXmlEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (var c in comments)
            {
                if (!string.IsNullOrEmpty(c.ParaId) && !commentByParaId.ContainsKey(c.ParaId))
                    commentByParaId[c.ParaId] = c;
            }

            string wrongBodyOnAnchor = null;

            foreach (var c in comments)
            {
                bool anchoredHere = CommentAnchorsOnPhrase(c, rangeById, commentByParaId, parentParaByChild, anchorPhrase);
                if (!anchoredHere)
                    continue;

                bool isReply = !string.IsNullOrEmpty(c.ParaId)
                    && parentParaByChild.ContainsKey(c.ParaId);
                if (requireReply && !isReply)
                    continue;
                if (!requireReply && isReply)
                    continue;

                if (IsExactCommentPhrase(c.Body, expectedBody))
                    return true;

                string cleaned = StripWordCommentControlChars(c.Body ?? "");
                if (!string.IsNullOrEmpty(cleaned))
                    wrongBodyOnAnchor = cleaned;
            }

            // 正しい見出し上に別本文がある場合のみ具体理由。位置違い・欠損は「ありません」。
            if (!string.IsNullOrEmpty(wrongBodyOnAnchor))
            {
                failReason = requireReply
                    ? $"見出し「{anchorPhrase}」のコメントへの返信が「{wrongBodyOnAnchor}」になっています。"
                    : $"見出し「1.{anchorPhrase}」のコメントが「{wrongBodyOnAnchor}」になっています。";
            }
            else
            {
                failReason = requireReply
                    ? $"見出し「{anchorPhrase}」のコメントに「{expectedBody}」の返信がありません。"
                    : $"見出し「1.{anchorPhrase}」に「{expectedBody}」のコメントがありません。";
            }
            return false;
        }

        private static bool CommentAnchorsOnPhrase(
            CommentXmlEntry comment,
            Dictionary<string, string> rangeById,
            Dictionary<string, CommentXmlEntry> commentByParaId,
            Dictionary<string, string> parentParaByChild,
            string anchorPhrase)
        {
            if (comment == null || string.IsNullOrEmpty(anchorPhrase))
                return false;

            if (rangeById != null
                && !string.IsNullOrEmpty(comment.Id)
                && rangeById.TryGetValue(comment.Id, out string ownRange)
                && !string.IsNullOrEmpty(ownRange)
                && ownRange.IndexOf(anchorPhrase, StringComparison.Ordinal) >= 0)
            {
                return true;
            }

            // 返信は親コメントの range で位置判定（返信側に range が無い場合もある）
            if (!string.IsNullOrEmpty(comment.ParaId)
                && parentParaByChild != null
                && parentParaByChild.TryGetValue(comment.ParaId, out string parentPara)
                && !string.IsNullOrEmpty(parentPara)
                && commentByParaId != null
                && commentByParaId.TryGetValue(parentPara, out CommentXmlEntry parent)
                && parent != null
                && rangeById != null
                && !string.IsNullOrEmpty(parent.Id)
                && rangeById.TryGetValue(parent.Id, out string parentRange)
                && !string.IsNullOrEmpty(parentRange)
                && parentRange.IndexOf(anchorPhrase, StringComparison.Ordinal) >= 0)
            {
                return true;
            }

            return false;
        }

        private sealed class CommentXmlEntry
        {
            public string Id;
            public string Body;
            public string ParaId;
        }

        private static List<CommentXmlEntry> ParseCommentEntries(string wordOpenXml)
        {
            var list = new List<CommentXmlEntry>();
            if (string.IsNullOrEmpty(wordOpenXml))
                return list;

            foreach (string partXml in EnumerateCommentsPackageParts(wordOpenXml))
            {
                // 本文は comments.xml の w:comment のみ（Extended 等は除外）
                if (partXml.IndexOf("pkg:name=\"/word/comments.xml\"", StringComparison.OrdinalIgnoreCase) < 0
                    && partXml.IndexOf("pkg:name='/word/comments.xml'", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                foreach (Match comment in Regex.Matches(
                    partXml,
                    @"<w:comment\b([^>]*)>(.*?)</w:comment>",
                    RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                {
                    string attrs = comment.Groups[1].Value;
                    string inner = comment.Groups[2].Value;
                    string id = GetXmlAttr(attrs, "w:id");
                    if (string.IsNullOrEmpty(id))
                        continue;

                    string paraId = "";
                    Match para = Regex.Match(
                        inner,
                        @"<w:p\b[^>]*\bw14:paraId\s*=\s*""([^""]+)""",
                        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                    if (para.Success)
                        paraId = para.Groups[1].Value;

                    list.Add(new CommentXmlEntry
                    {
                        Id = id,
                        Body = ExtractConcatenatedWtText(inner),
                        ParaId = paraId
                    });
                }
            }

            return list;
        }

        private static Dictionary<string, string> BuildCommentRangeTexts(string wordOpenXml)
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            string docXml = GetPackagePartXml(wordOpenXml, "/word/document.xml");
            if (string.IsNullOrEmpty(docXml))
                return map;

            // 返信コメントは親 range の内側に Start が重なることがあるため、
            // Matches の非重複走査ではなく Start ごとに対応 End を探す。
            foreach (Match start in Regex.Matches(
                docXml,
                @"<w:commentRangeStart\b[^>]*\bw:id\s*=\s*""(\d+)""[^/]*/>",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                string id = start.Groups[1].Value;
                if (map.ContainsKey(id))
                    continue;

                string fromStart = docXml.Substring(start.Index);
                Match end = Regex.Match(
                    fromStart,
                    @"<w:commentRangeEnd\b[^>]*\bw:id\s*=\s*""" + Regex.Escape(id) + @"""[^/]*/>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                if (!end.Success)
                    continue;

                string fragment = fromStart.Substring(0, end.Index + end.Length);
                map[id] = ExtractConcatenatedWtText(fragment);
            }

            return map;
        }

        private static Dictionary<string, string> BuildCommentParaIdParentMap(string wordOpenXml)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(wordOpenXml))
                return map;

            foreach (string partXml in EnumerateCommentsPackageParts(wordOpenXml))
            {
                if (partXml.IndexOf("/word/commentsExtended", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                foreach (Match m in Regex.Matches(
                    partXml,
                    @"<w15:commentEx\b([^>/]*)/?>",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
                {
                    string attrs = m.Groups[1].Value;
                    string paraId = GetXmlAttr(attrs, "w15:paraId");
                    string parent = GetXmlAttr(attrs, "w15:paraIdParent");
                    if (!string.IsNullOrEmpty(paraId) && !string.IsNullOrEmpty(parent))
                        map[paraId] = parent;
                }
            }

            return map;
        }

        private static string GetPackagePartXml(string wordOpenXml, string partName)
        {
            if (string.IsNullOrEmpty(wordOpenXml) || string.IsNullOrEmpty(partName))
                return "";
            string escaped = Regex.Escape(partName);
            Match part = Regex.Match(
                wordOpenXml,
                @"<pkg:part\b[^>]*\bpkg:name\s*=\s*""" + escaped + @"""[^>]*>.*?</pkg:part>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return part.Success ? part.Value : "";
        }

        private static string GetXmlAttr(string attrs, string name)
        {
            if (string.IsNullOrEmpty(attrs) || string.IsNullOrEmpty(name))
                return "";
            Match m = Regex.Match(
                attrs,
                Regex.Escape(name) + @"\s*=\s*""([^""]*)""",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return m.Success ? m.Groups[1].Value : "";
        }

        private static IEnumerable<string> EnumerateCommentsPackageParts(string wordOpenXml)
        {
            if (string.IsNullOrEmpty(wordOpenXml))
                yield break;

            foreach (Match part in Regex.Matches(
                wordOpenXml,
                @"<pkg:part\b[^>]*\bpkg:name\s*=\s*""(/word/comments[^""]*)""[^>]*>.*?</pkg:part>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                if (part.Success)
                    yield return part.Value;
            }
        }

        private static string ExtractConcatenatedWtText(string xmlFragment)
        {
            if (string.IsNullOrEmpty(xmlFragment))
                return "";
            var sb = new StringBuilder();
            foreach (Match wt in Regex.Matches(
                xmlFragment,
                @"<w:t(?:\s[^>]*)?>([^<]*)</w:t>",
                RegexOptions.CultureInvariant))
            {
                if (wt.Success && wt.Groups.Count >= 2)
                    sb.Append(wt.Groups[1].Value);
            }
            return sb.ToString();
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

        /// <summary>
        /// 4-5 ×理由。文字・向きが取れたら分岐し、取れなければ汎用文。
        /// 向きは「斜め／横」で伝え、ギャラリーの 1・2 表記は使わない。
        /// </summary>
        private static string BuildTask45FailReason(string normalizedXml)
        {
            WordWatermarkInspection.WatermarkStudentView view =
                WordWatermarkInspection.AnalyzeForStudent(normalizedXml);

            bool textIsSample = IsSampleDisplayText(view.DisplayText);
            bool orientationWrong = view.IsHorizontal == false;
            bool orientationOk = view.IsHorizontal == true;
            bool hasText = !string.IsNullOrEmpty(view.DisplayText);

            // 正解文字（サンプル）＋斜め
            if (textIsSample && orientationWrong)
                return "透かしの文字の向きが斜めになっています。";

            // 下書きは従来どおりギャラリー名で出す
            if (string.Equals(view.DisplayText, "下書き", StringComparison.Ordinal))
            {
                if (orientationWrong)
                    return "透かしが「下書き1」になっています。";
                if (orientationOk)
                    return "透かしが「下書き2」になっています。";
                return "透かしが「下書き」になっています。";
            }

            if (hasText && !textIsSample)
            {
                if (orientationWrong)
                    return $"文字が「{view.DisplayText}」になっており、向きも斜めになっています。";
                return $"文字が「{view.DisplayText}」になっています。";
            }

            return "透かし「サンプル２」がありません。";
        }

        private static bool IsSampleDisplayText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;
            return string.Equals(text, "サンプル", StringComparison.Ordinal)
                || string.Equals(text, "SAMPLE", StringComparison.OrdinalIgnoreCase);
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

