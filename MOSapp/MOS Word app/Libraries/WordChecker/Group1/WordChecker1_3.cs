using System;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Libraries;

namespace Libraries.Group1
{
    public class WordChecker1_3
    {
        public bool CheckTask_1_3_01() => RunCheck(CheckTask_1_3_01_Impl);
        public bool CheckTask_1_3_02() => RunCheck(CheckTask_1_3_02_Impl);
        public bool CheckTask_1_3_03() => RunCheck(CheckTask_1_3_03_Impl);
        public bool CheckTask_1_3_04() => RunCheck(CheckTask_1_3_04_Impl);
        public bool CheckTask_1_3_05() => RunCheck(CheckTask_1_3_05_Impl);
        public bool CheckTask_1_3_06() => RunCheck(CheckTask_1_3_06_Impl);

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

        private bool CheckTask_1_3_01_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);

                // 文書の余白が「やや狭い」プリセット相当になっているかを判定する
                Section section = document.Sections[1];
                PageSetup ps = section.PageSetup;
                float top = ps.TopMargin;
                float bottom = ps.BottomMargin;
                float left = ps.LeftMargin;
                float right = ps.RightMargin;
                string presetName = IdentifyCommonMarginPreset(top, bottom, left, right);

                Marshal.ReleaseComObject(ps);
                Marshal.ReleaseComObject(section);

                // 文書状態が本番試験と同等であれば正解とする（ログは補助情報として扱い、必須にはしない）
                if (presetName == "やや狭い")
                    return true;

                // 余白プリセットを使った操作は証跡で記録（個別リセット後の旧全体ログ誤判定を防ぐ）
                bool logOk = LogReader.HasTaskEvidence(3, 1, "PageMarginsModerate");
                if (logOk)
                    return true;

                if (!string.IsNullOrEmpty(presetName))
                    return Miss($"文書の余白が「{presetName}」になっています。");
                return Miss("文書の余白が「やや狭い」になっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>
        /// 日本語版の一般的な余白プリセット名を返す。一致しなければ null。
        /// やや狭い=上下72pt/左右54pt、狭い=全辺36pt、広い=上下72pt/左右144pt、
        /// 標準=上約35.01mm・下左右30mm（約99.26/85.04pt）。
        /// </summary>
        private static string IdentifyCommonMarginPreset(float top, float bottom, float left, float right)
        {
            bool IsApprox(float value, float target) => Math.Abs(value - target) <= 1.5f;
            bool Match(float t, float b, float l, float r) =>
                IsApprox(top, t) && IsApprox(bottom, b) && IsApprox(left, l) && IsApprox(right, r);

            if (Match(72.0f, 72.0f, 54.0f, 54.0f))
                return "やや狭い";
            if (Match(36.0f, 36.0f, 36.0f, 36.0f))
                return "狭い";
            if (Match(72.0f, 72.0f, 144.0f, 144.0f))
                return "広い";
            // 日本語版「標準」: 上 35.01mm / 下・左・右 30mm
            if (Match(99.26f, 85.04f, 85.04f, 85.04f))
                return "標準";
            return null;
        }

        private bool CheckTask_1_3_02_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);
                Range searchRange = WordFindHelper.DuplicateContent(document); Find find = searchRange.Find; WordFindHelper.ConfigureSafeFind(find, "参考文献一覧"); find.Execute();
                if (!find.Found)
                {
                    Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                    return Miss("見出し「参考文献一覧」がありません。");
                }
                Range foundRange = searchRange;
                // 教材どおり「参考文献一覧」はセクション2の本文にあること（区切りだけ挿入して別セクションに見出しが無い誤正解を防ぐ）
                bool headingInSection2 = false;
                if (document.Sections.Count >= 2)
                {
                    Section section2 = document.Sections[2];
                    try
                    {
                        int s2Start = section2.Range.Start;
                        int s2End = section2.Range.End;
                        headingInSection2 = foundRange.Start >= s2Start && foundRange.Start < s2End;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(section2);
                    }
                }
                Paragraph headingParagraph = foundRange.Paragraphs[1];
                int headingStart = headingParagraph.Range.Start;
                int headingEnd = headingParagraph.Range.End;
                int currentSectionIndex = foundRange.Sections[1].Index;
                bool fileStateOk = false;
                int boundaryPos = -1;
                bool nextPageStart = false;
                int absDiff = -1;
                int secLastIdx = -1;
                int secAtEndIdx = -1;
                int secBeforeHeadingIdx = -1;
                int secAtHeadingIdx = -1;
                // 段落最終文字と段落直後位置でセクションが変わる＝見出し直後に区切り（Word の境界位置と headingEnd の数文字ズレに強い）
                if (headingEnd > 1)
                {
                    Section secLast = GetSectionContainingPosition(document, headingEnd - 1);
                    Section secAtEnd = GetSectionContainingPosition(document, headingEnd);
                    try
                    {
                        if (secLast != null && secAtEnd != null)
                        {
                            secLastIdx = secLast.Index;
                            secAtEndIdx = secAtEnd.Index;
                            if (secAtEnd.Index > secLast.Index)
                            {
                                PageSetup ps = secAtEnd.PageSetup;
                                try
                                {
                                    nextPageStart = ps.SectionStart == WdSectionStart.wdSectionNewPage;
                                    boundaryPos = secAtEnd.Range.Start;
                                    absDiff = Math.Abs(boundaryPos - headingEnd);
                                    fileStateOk = nextPageStart;
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(ps);
                                }
                            }
                        }
                    }
                    finally
                    {
                        if (secLast != null) Marshal.ReleaseComObject(secLast);
                        if (secAtEnd != null) Marshal.ReleaseComObject(secAtEnd);
                    }
                }
                // 区切りが「見出しの次のページから開始」の直前にあり、見出し本文が次セクション先頭から始まるケース（secLast/End が同一セクションのまま）
                if (!fileStateOk && headingStart > 1)
                {
                    Section secBefore = GetSectionContainingPosition(document, headingStart - 1);
                    Section secAtStart = GetSectionContainingPosition(document, headingStart);
                    try
                    {
                        if (secBefore != null && secAtStart != null)
                        {
                            secBeforeHeadingIdx = secBefore.Index;
                            secAtHeadingIdx = secAtStart.Index;
                            if (secAtStart.Index > secBefore.Index)
                            {
                                PageSetup ps = secAtStart.PageSetup;
                                try
                                {
                                    nextPageStart = ps.SectionStart == WdSectionStart.wdSectionNewPage;
                                    boundaryPos = secAtStart.Range.Start;
                                    absDiff = Math.Abs(boundaryPos - headingStart);
                                    fileStateOk = nextPageStart;
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(ps);
                                }
                            }
                        }
                    }
                    finally
                    {
                        if (secBefore != null) Marshal.ReleaseComObject(secBefore);
                        if (secAtStart != null) Marshal.ReleaseComObject(secAtStart);
                    }
                }
                if (!fileStateOk && currentSectionIndex < document.Sections.Count)
                {
                    Section nextSec = document.Sections[currentSectionIndex + 1];
                    try
                    {
                        nextPageStart = nextSec.PageSetup.SectionStart == WdSectionStart.wdSectionNewPage;
                        boundaryPos = nextSec.Range.Start;
                        absDiff = Math.Abs(boundaryPos - headingEnd);
                        // 段落境界方式で取れない環境向け。区切り文字の差で ±2 だと誤不合格になりやすい
                        bool boundaryAtHeadingTail = absDiff <= 12;
                        fileStateOk = nextPageStart && boundaryAtHeadingTail;
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(nextSec);
                    }
                }
                Marshal.ReleaseComObject(headingParagraph);
                Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                // 文書状態が「次のページから開始」のセクション区切りで見出し直後に境界があれば正解
                // 失敗理由は典型ミス別に分ける（改ページのみ／区切り位置違い／連続など）
                if (document.Sections.Count < 2)
                {
                    WordScoreExplanation.Note("見出し「参考文献一覧」の直前にセクション区切りがありません。");
                    return Miss("セクション区切りは「次のページから開始」にする必要があります。");
                }
                if (!headingInSection2)
                    return Miss("見出し「参考文献一覧」が、セクション区切りの次のセクション先頭になっていません。");
                if (!fileStateOk)
                    return Miss("セクション区切りが「次のページから開始」になっていません。");
                return true;
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_3_03_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);
                Range searchRange = WordFindHelper.DuplicateContent(document); Find find = searchRange.Find; WordFindHelper.ConfigureSafeFind(find, "参考文献一覧"); find.Execute();
                if (!find.Found)
                {
                    Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                    return Miss("見出し「参考文献一覧」がありません。");
                }
                Range foundRange = searchRange;
                Paragraph headingPara = foundRange.Paragraphs[1];
                int headingEnd = headingPara.Range.End;
                Marshal.ReleaseComObject(headingPara);
                int anchorPos = foundRange.Start;
                Section section = GetSectionContainingPosition(document, anchorPos);
                if (section == null)
                {
                    Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                    return Miss(WordScoreExplanation.UnavailableText);
                }
                bool landscape = false;
                int secIndex = 0;
                try
                {
                    secIndex = section.Index;
                    landscape = section.PageSetup.Orientation == WdOrientation.wdOrientLandscape;
                    // 見出しが前セクション末尾にあり、横向きが「次のセクション」だけに付くとアンカー位置の PageSetup は縦のまま
                    if (!landscape && secIndex < document.Sections.Count)
                    {
                        Section nextSec = document.Sections[secIndex + 1];
                        try
                        {
                            int nextStart = nextSec.Range.Start;
                            bool nextLandscape = nextSec.PageSetup.Orientation == WdOrientation.wdOrientLandscape;
                            if (nextLandscape && Math.Abs(nextStart - headingEnd) <= 12)
                                landscape = true;
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(nextSec);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(section);
                }
                Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                bool logOk = LogReader.HasTaskEvidence(3, 3, "PageOrientationPortraitLandscape");
                if (!landscape)
                    return Miss("見出し「参考文献一覧」で始まるセクションが横向きになっていません。");
                if (!logOk)
                    return Miss("印刷の向きについて、必要な操作の記録がありません。");
                return true;
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_3_04_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);
                Range searchRange = WordFindHelper.DuplicateContent(document); Find find = searchRange.Find; WordFindHelper.ConfigureSafeFind(find, "参考文献一覧"); find.Execute();
                if (!find.Found)
                {
                    Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                    return Miss("見出し「参考文献一覧」がありません。");
                }
                Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                // セクション2のみ「2段組」（等幅）を判定。他セクションの段組や「狭くした2段組」（不等幅）は不正解
                if (document.Sections.Count < 2)
                    return Miss("見出し「参考文献一覧」で始まるセクションが2段組みになっていません。");
                bool stateOk = false;
                int columnCount = 0;
                bool evenlySpaced = false;
                bool equalWidths = false;
                Section section2 = document.Sections[2];
                try
                {
                    TextColumns cols = section2.PageSetup.TextColumns;
                    columnCount = cols.Count;
                    if (cols.Count == 2)
                    {
                        TextColumn c1 = cols[1];
                        TextColumn c2 = cols[2];
                        try
                        {
                            evenlySpaced = cols.EvenlySpaced != 0;
                            float w1 = (float)c1.Width;
                            float w2 = (float)c2.Width;
                            equalWidths = Math.Abs(w1 - w2) <= 2.0f;
                            stateOk = evenlySpaced && equalWidths;
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(c1);
                            Marshal.ReleaseComObject(c2);
                        }
                    }
                    Marshal.ReleaseComObject(cols);
                }
                finally
                {
                    Marshal.ReleaseComObject(section2);
                }
                if (stateOk)
                    return true;
                if (columnCount != 2)
                    return Miss("見出し「参考文献一覧」で始まるセクションが2段組みになっていません。");
                if (!evenlySpaced || !equalWidths)
                    return Miss("見出し「参考文献一覧」で始まるセクションが、幅の等しい2段組みになっていません。");
                return Miss("見出し「参考文献一覧」で始まるセクションが2段組みになっていません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private bool CheckTask_1_3_05_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);
                Range searchRange = WordFindHelper.DuplicateContent(document);
                Find find = searchRange.Find;
                WordFindHelper.ConfigureSafeFind(find, "茨城県天心記念五浦美術館");
                find.Execute();
                if (!find.Found)
                {
                    find.Text = "●茨城県天心記念五浦美術館";
                    find.Execute();
                }
                if (!find.Found)
                {
                    find.Text = "tenshin.museum.ibk.ed.jp";
                    find.Execute();
                }
                if (!find.Found)
                {
                    Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                    return Miss("「●茨城県天心記念五浦美術館『…』」の段落がありません。");
                }
                Paragraph para = searchRange.Paragraphs[1];
                Range paraRange = para.Range;
                int paraStart = paraRange.Start;
                Marshal.ReleaseComObject(paraRange);
                Marshal.ReleaseComObject(para);
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);
                // 該当段落の先頭の直前に列区切りのみ正解。セクション内の別位置に区切りがあっても不正解
                if (IsColumnBreakImmediatelyBeforeParagraph(document, paraStart))
                    return true;
                return Miss("「●茨城県天心記念五浦美術館『…』」の先頭に段区切りがありません。");
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        private static bool IsParagraphEndChar(char c)
        {
            return c == '\r' || c == '\n' || c == '\v' || c == (char)7 || c == '\u000b';
        }

        /// <summary>該当段落の先頭（paraStart）の直前に列区切り（Chr(14)）があるか。段落記号(\r等)のみが列区切りと先頭の間に挟まる場合も正解（ログで paraStart-2 が本文のケースに対応）。</summary>
        private static bool IsColumnBreakImmediatelyBeforeParagraph(Document document, int paraStart)
        {
            if (paraStart < 1) return false;
            const char colBreak = (char)14;
            int maxScan = Math.Min(paraStart, 8192);
            try
            {
                // 列区切りを段落の直前に挿入すると、Word では段落の先頭 1 文字が Chr(14) になることがある。Range(scanStart, paraStart) は paraStart を含まないため、ここで拾う。
                Range rFirst = document.Range(paraStart, paraStart + 1);
                try
                {
                    string tf = rFirst.Text ?? "";
                    if (tf.Length > 0 && tf[0] == colBreak)
                        return true;
                }
                finally
                {
                    Marshal.ReleaseComObject(rFirst);
                }
                Range r1 = document.Range(paraStart - 1, paraStart);
                try
                {
                    string t = r1.Text ?? "";
                    if (t.Length > 0 && t[t.Length - 1] == colBreak)
                        return true;
                }
                finally
                {
                    Marshal.ReleaseComObject(r1);
                }
                int scanStart = Math.Max(0, paraStart - maxScan);
                Range rScan = document.Range(scanStart, paraStart);
                try
                {
                    string s = rScan.Text ?? "";
                    if (s.Length < 2)
                        return false;
                    int last = s.Length - 1;
                    if (!IsParagraphEndChar(s[last]))
                        return false;
                    for (int i = last - 1; i >= 0; i--)
                    {
                        if (s[i] != colBreak)
                            continue;
                        bool onlyParaMarksAfter = true;
                        for (int j = i + 1; j <= last; j++)
                        {
                            if (!IsParagraphEndChar(s[j]))
                            {
                                onlyParaMarksAfter = false;
                                break;
                            }
                        }
                        if (onlyParaMarksAfter)
                            return true;
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(rScan);
                }
            }
            catch { }
            return false;
        }

        private bool CheckTask_1_3_06_Impl(string filePath)
        {
            Application wordApp = null; Document document = null;
            try
            {
                try { wordApp = (Application)Marshal.GetActiveObject("Word.Application"); } catch { return Miss(WordScoreExplanation.UnavailableText); }
                document = null; string fileName = System.IO.Path.GetFileName(filePath);
                foreach (Document doc in wordApp.Documents)
                { if (doc.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase) || doc.Name.Equals(fileName, StringComparison.OrdinalIgnoreCase)) { document = doc; break; } }
                if (document == null) return Miss(WordScoreExplanation.UnavailableText);
                Range searchRange = WordFindHelper.DuplicateContent(document); Find find = searchRange.Find; WordFindHelper.ConfigureSafeFind(find, "参考文献一覧"); find.Execute();
                if (!find.Found)
                {
                    Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                    return Miss("見出し「参考文献一覧」がありません。");
                }
                Range foundRange = searchRange;
                Section section = GetSectionContainingPosition(document, foundRange.Start);
                if (section == null)
                {
                    Marshal.ReleaseComObject(find); Marshal.ReleaseComObject(searchRange);
                    return Miss(WordScoreExplanation.UnavailableText);
                }
                Paragraph headingPara = foundRange.Paragraphs[1];
                int headingEnd = headingPara.Range.End;
                Marshal.ReleaseComObject(headingPara);
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);

                // task5 と同じ検索で「段区切り」の直後の段落先頭＝1段目の終わり（その手前の段落までが 3-6 の対象）
                searchRange = WordFindHelper.DuplicateContent(document);
                find = searchRange.Find;
                WordFindHelper.ConfigureSafeFind(find, "茨城県天心記念五浦美術館");
                find.Execute();
                if (!find.Found)
                {
                    find.Text = "●茨城県天心記念五浦美術館";
                    find.Execute();
                }
                if (!find.Found)
                {
                    find.Text = "tenshin.museum.ibk.ed.jp";
                    find.Execute();
                }
                if (!find.Found)
                {
                    Marshal.ReleaseComObject(find);
                    Marshal.ReleaseComObject(searchRange);
                    Marshal.ReleaseComObject(section);
                    return Miss("「●茨城県天心記念五浦美術館『…』」の段落がありません。");
                }
                Paragraph task5Para = searchRange.Paragraphs[1];
                int firstColumnEndStart = task5Para.Range.Start;
                Marshal.ReleaseComObject(task5Para);
                Marshal.ReleaseComObject(find);
                Marshal.ReleaseComObject(searchRange);

                if (firstColumnEndStart <= headingEnd)
                {
                    Marshal.ReleaseComObject(section);
                    return Miss("1段目の範囲が確認できません。");
                }

                // 1段目のみ: 見出し直後〜 task5 段落の直前までの全段落が行間 1.3（部分選択のみ変更は不正解）
                bool sawFirstColumnParagraph = false;
                bool allParagraphs13 = true;
                string firstWrongSpacingLabel = null;
                try
                {
                    int n = section.Range.Paragraphs.Count;
                    for (int i = 1; i <= n; i++)
                    {
                        Paragraph p = section.Range.Paragraphs[i];
                        try
                        {
                            if (p.Range.End <= headingEnd)
                                continue;
                            if (p.Range.Start < headingEnd)
                                continue;
                            if (p.Range.Start >= firstColumnEndStart)
                                break;
                            sawFirstColumnParagraph = true;
                            ParagraphFormat pf = p.Range.ParagraphFormat;
                            float fontSize = 0f;
                            try { fontSize = (float)p.Range.Font.Size; } catch { }
                            if (fontSize <= 0f)
                            {
                                Range oneChar = null;
                                try
                                {
                                    oneChar = document.Range(p.Range.Start, Math.Min(p.Range.Start + 1, p.Range.End));
                                    fontSize = (float)oneChar.Font.Size;
                                }
                                catch { }
                                finally { if (oneChar != null) Marshal.ReleaseComObject(oneChar); }
                            }
                            if (!IsLineSpacingAbout13Times(pf, fontSize))
                            {
                                allParagraphs13 = false;
                                if (firstWrongSpacingLabel == null)
                                    firstWrongSpacingLabel = DescribeLineSpacing(pf, fontSize);
                            }
                            Marshal.ReleaseComObject(pf);
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(p);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(section);
                }
                if (!sawFirstColumnParagraph)
                    return Miss("1段目の段落がありません。");
                if (!allParagraphs13)
                {
                    if (!string.IsNullOrEmpty(firstWrongSpacingLabel))
                        return Miss($"見出し「参考文献一覧」で始まるセクションの1段目の行間が「{firstWrongSpacingLabel}」になっています。");
                    return Miss("見出し「参考文献一覧」で始まるセクションの1段目の行間がすべて「1.3」になっていません。");
                }
                return true;
            }
            catch { return Miss(WordScoreExplanation.UnavailableText); }
            finally { if (document != null) Marshal.ReleaseComObject(document); }
        }

        /// <summary>行間が 1.3 倍として設定されているか（倍数・固定・以上）。1.2 や 1.5 などは不正解。</summary>
        private static bool IsLineSpacingAbout13Times(ParagraphFormat pf, float fontSize)
        {
            WdLineSpacing rule = pf.LineSpacingRule;
            float spacing = 0f;
            try { spacing = (float)pf.LineSpacing; } catch { return false; }

            const float target = 1.3f;
            const float eps = 0.05f;

            if (rule == WdLineSpacing.wdLineSpaceMultiple)
            {
                // 行の倍数（1.3）そのまま
                if (spacing >= 1.0f && spacing <= 3.0f)
                    return Math.Abs(spacing - target) <= eps;
                // 環境によって 1 行を 12 単位で表す（1.3 行 ≈ 15.6）。
                // ここに入ったら /12 のみで判定する（/fontSize へ落とすと 1.1 行などが誤〇になる）。
                if (spacing >= 10f && spacing <= 48f)
                {
                    float m12 = spacing / 12f;
                    return m12 >= 1.0f && m12 <= 2.0f && Math.Abs(m12 - target) <= eps;
                }
                return false;
            }
            if (rule == WdLineSpacing.wdLineSpaceExactly || rule == WdLineSpacing.wdLineSpaceAtLeast)
            {
                if (fontSize > 0f)
                {
                    float multiple = spacing / fontSize;
                    return Math.Abs(multiple - target) <= eps;
                }
                return false;
            }
            return false;
        }

        /// <summary>学生向けに行間の現状を短いラベルにする（倍数なら「1.1」、固定なら「固定値」など）。</summary>
        private static string DescribeLineSpacing(ParagraphFormat pf, float fontSize)
        {
            WdLineSpacing rule;
            float spacing;
            try
            {
                rule = pf.LineSpacingRule;
                spacing = (float)pf.LineSpacing;
            }
            catch
            {
                return null;
            }

            if (rule == WdLineSpacing.wdLineSpaceMultiple)
            {
                float multiple;
                if (spacing >= 1.0f && spacing <= 3.0f)
                    multiple = spacing;
                else if (spacing >= 10f && spacing <= 48f)
                    multiple = spacing / 12f;
                else
                    return "倍数（不明）";
                return FormatLineSpacingMultiple(multiple);
            }
            if (rule == WdLineSpacing.wdLineSpaceSingle)
                return "1行";
            if (rule == WdLineSpacing.wdLineSpace1pt5)
                return "1.5行";
            if (rule == WdLineSpacing.wdLineSpaceDouble)
                return "2行";
            if (rule == WdLineSpacing.wdLineSpaceExactly)
                return "固定値";
            if (rule == WdLineSpacing.wdLineSpaceAtLeast)
                return "最小値";
            return null;
        }

        private static string FormatLineSpacingMultiple(float multiple)
        {
            // 1.1 / 1.15 / 1.3 など、見た目の桁を抑える
            float rounded = (float)Math.Round(multiple, 2);
            if (Math.Abs(rounded * 10f - Math.Round(rounded * 10f)) < 0.001f)
                return rounded.ToString("0.0", CultureInfo.InvariantCulture);
            return rounded.ToString("0.##", CultureInfo.InvariantCulture);
        }

        /// <summary>指定文字位置を含むセクションを返す。呼び出し元で Section を Release すること。</summary>
        private static Section GetSectionContainingPosition(Document document, int position)
        {
            try
            {
                int n = document.Sections.Count;
                for (int i = 1; i <= n; i++)
                {
                    Section s = document.Sections[i];
                    if (position >= s.Range.Start && position < s.Range.End)
                        return s;
                    Marshal.ReleaseComObject(s);
                }
            }
            catch { }
            return null;
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
