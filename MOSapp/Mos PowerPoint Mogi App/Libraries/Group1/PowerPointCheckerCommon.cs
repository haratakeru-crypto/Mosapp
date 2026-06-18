using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace Libraries.Group1
{
    /// <summary>
    /// PowerPoint 採点チェッカー共通ヘルパー。
    /// </summary>
    public static class PowerPointCheckerCommon
    {
        /// <summary>
        /// PowerPoint COM（開閉・ActivePresentation・採点）を直列化する。採点スレッドと UI スレッドの競合で無効な Presentation を触るのを防ぐ。
        /// </summary>
        public static readonly object PowerPointComInteropSync = new object();

        /// <summary>
        /// 現在アクティブな PowerPoint プレゼンテーションを取得する。
        /// </summary>
        /// <returns>アクティブプレゼンテーション。取得失敗時は null。呼び出し元で Marshal.ReleaseComObject すること。</returns>
        public static Presentation GetActivePresentation()
        {
            lock (PowerPointComInteropSync)
            {
                try
                {
                    var app = (Application)Marshal.GetActiveObject("PowerPoint.Application");
                    if (app == null) return null;
                    try
                    {
                        var pres = app.ActivePresentation;
                        return pres;
                    }
                    catch
                    {
                        return null;
                    }
                }
                catch (COMException)
                {
                    return null;
                }
            }
        }

        /// <summary>
        /// 指定したスライド番号（1 始まり）のスライドを取得する。
        /// </summary>
        /// <param name="pres">プレゼンテーション。</param>
        /// <param name="slideNumber">スライド番号（1 始まり）。</param>
        /// <returns>該当スライド。範囲外または取得失敗時は null。呼び出し元で Marshal.ReleaseComObject すること。</returns>
        public static Slide GetSlideByNumber(Presentation pres, int slideNumber)
        {
            if (pres == null || slideNumber < 1)
                return null;
            try
            {
                Slides slides = pres.Slides;
                if (slides == null)
                    return null;
                try
                {
                    if (slideNumber > slides.Count)
                        return null;
                    return slides[slideNumber];
                }
                finally
                {
                    if (slides != null)
                    {
                        try { Marshal.ReleaseComObject(slides); } catch { }
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 作業中 pptx（TabN\ProjectM.pptx）から、対応する初期ファイル（TabN\Initial\projectM.pptx）のパスを解決する。
        /// </summary>
        public static bool TryResolveInitialPptxPath(string activePptxPath, out string initialPptxPath)
        {
            initialPptxPath = null;
            if (string.IsNullOrWhiteSpace(activePptxPath) || !File.Exists(activePptxPath))
                return false;

            string fileName = Path.GetFileNameWithoutExtension(activePptxPath);
            if (string.IsNullOrEmpty(fileName))
                return false;

            int projectId = 0;
            if (fileName.StartsWith("Project", StringComparison.OrdinalIgnoreCase))
                int.TryParse(fileName.Substring("Project".Length), out projectId);
            if (projectId < 1)
                return false;

            string tabFolder = Path.GetDirectoryName(activePptxPath);
            if (string.IsNullOrEmpty(tabFolder))
                return false;

            string ext = Path.GetExtension(activePptxPath);
            if (string.IsNullOrEmpty(ext))
                ext = ".pptx";

            initialPptxPath = Path.Combine(tabFolder, "Initial", $"project{projectId}{ext}");
            return File.Exists(initialPptxPath);
        }

        /// <summary>
        /// 指定したスライド内で、特定のテキストを含む図形を探す（1 階層のみ）。
        /// </summary>
        /// <param name="slide">スライド。</param>
        /// <param name="searchText">検索するテキスト。</param>
        /// <returns>該当図形。見つからない場合は null。呼び出し元で Marshal.ReleaseComObject すること。</returns>
        public static PptShape FindShapeWithText(Slide slide, string searchText)
        {
            if (slide == null || string.IsNullOrEmpty(searchText))
                return null;

            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null)
                    return null;

                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        if (sh.HasTextFrame == MsoTriState.msoTrue)
                        {
                            string text = null;
                            try
                            {
                                var pptTf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                                text = pptTf.TextRange.Text;
                            }
                            catch { }
                            if (text != null && text.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                PptShape result = sh;
                                sh = null;
                                return result;
                            }
                        }
                    }
                    finally
                    {
                        if (sh != null)
                        {
                            try { Marshal.ReleaseComObject(sh); } catch { }
                        }
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (shapes != null)
                {
                    try { Marshal.ReleaseComObject(shapes); } catch { }
                }
            }
        }

        /// <summary>
        /// スライドのタイトル（またはスライド内のテキスト）に指定文字列を含むスライドを取得する。
        /// </summary>
        /// <param name="pres">プレゼンテーション。</param>
        /// <param name="titlePart">検索する文字列（部分一致、大文字小文字区別なし）。</param>
        /// <returns>該当スライド。見つからなければ null。呼び出し元で Marshal.ReleaseComObject すること。</returns>
        public static Slide GetSlideByTitle(Presentation pres, string titlePart)
        {
            if (pres == null || string.IsNullOrEmpty(titlePart))
                return null;
            try
            {
                Slides slides = pres.Slides;
                if (slides == null) return null;
                try
                {
                    int count = slides.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        Slide slide = null;
                        try
                        {
                            slide = slides[i];
                            PptShapes shapes = null;
                            try
                            {
                                shapes = slide.Shapes;
                                if (shapes == null) continue;
                                int sc = shapes.Count;
                                for (int j = 1; j <= sc; j++)
                                {
                                    PptShape sh = null;
                                    try
                                    {
                                        sh = shapes[j];
                                        if (sh.HasTextFrame != MsoTriState.msoTrue) continue;
                                        string text = null;
                                        try
                                        {
                                            var pptTf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                                            text = pptTf.TextRange.Text ?? "";
                                        }
                                        catch { continue; }
                                        if (text.IndexOf(titlePart, StringComparison.OrdinalIgnoreCase) >= 0)
                                        {
                                            Slide result = slide;
                                            slide = null;
                                            return result;
                                        }
                                    }
                                    finally
                                    {
                                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                                    }
                                }
                            }
                            finally
                            {
                                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
                            }
                        }
                        finally
                        {
                            if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                        }
                    }
                    return null;
                }
                finally
                {
                    if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// タイトル系プレースホルダーの文字が <paramref name="requiredTitle"/> と一致するスライドを返す。
        /// 本文やセクションズーム上の「1.機能の概要」など、部分文字列だけ一致するスライドは除外する。
        /// </summary>
        public static Slide GetSlideByTitlePlaceholderExact(Presentation pres, string requiredTitle)
        {
            if (pres == null || string.IsNullOrEmpty(requiredTitle))
                return null;
            string requiredNorm = NormalizeSlideTitleText(requiredTitle);
            if (string.IsNullOrEmpty(requiredNorm))
                return null;

            try
            {
                Slides slides = pres.Slides;
                if (slides == null) return null;
                try
                {
                    int count = slides.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        Slide slide = null;
                        try
                        {
                            slide = slides[i];
                            if (SlideHasTitlePlaceholderExact(slide, requiredNorm))
                            {
                                Slide result = slide;
                                slide = null;
                                return result;
                            }
                        }
                        finally
                        {
                            if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                        }
                    }
                    return null;
                }
                finally
                {
                    if (slides != null) { try { Marshal.ReleaseComObject(slides); } catch { } }
                }
            }
            catch
            {
                return null;
            }
        }

        private static bool SlideHasTitlePlaceholderExact(Slide slide, string requiredNorm)
        {
            if (slide == null) return false;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return false;
                int sc = shapes.Count;
                for (int j = 1; j <= sc; j++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[j];
                        if (sh.HasTextFrame != MsoTriState.msoTrue) continue;
                        if (!IsTitleLikePlaceholderShape(sh)) continue;

                        string text = null;
                        try
                        {
                            var pptTf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                            text = pptTf.TextRange?.Text ?? "";
                        }
                        catch { continue; }

                        if (string.Equals(NormalizeSlideTitleText(text), requiredNorm, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                return false;
            }
            catch { return false; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private static bool IsTitleLikePlaceholderShape(PptShape sh)
        {
            try
            {
                if ((int)sh.Type != (int)MsoShapeType.msoPlaceholder) return false;
                var pf = sh.PlaceholderFormat;
                if (pf == null) return false;
                try
                {
                    var ppt = (PpPlaceholderType)pf.Type;
                    return ppt == PpPlaceholderType.ppPlaceholderTitle
                        || ppt == PpPlaceholderType.ppPlaceholderCenterTitle
                        || ppt == PpPlaceholderType.ppPlaceholderVerticalTitle
                        || ppt == PpPlaceholderType.ppPlaceholderSubtitle;
                }
                finally
                {
                    try { Marshal.ReleaseComObject(pf); } catch { }
                }
            }
            catch { return false; }
        }

        private static string NormalizeSlideTitleText(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder();
            foreach (char c in s.Trim().Replace("\r", "").Replace("\n", ""))
            {
                if (c == ' ' || c == '\u3000' || c == '\t') continue;
                sb.Append(NormalizeWidthInsensitiveChar(c));
            }
            return sb.ToString();
        }

        private static char NormalizeWidthInsensitiveChar(char c)
        {
            if (c >= '\uFF10' && c <= '\uFF19')
                return (char)('0' + (c - '\uFF10'));
            if (c == '\uFF0E') return '.';
            return c;
        }

        /// <summary>
        /// スライド内の 3D モデル図形を検索する。
        /// </summary>
        public static PptShape Find3DModelShape(Slide slide)
        {
            if (slide == null) return null;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return null;
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        try
                        {
                            var st = (MsoShapeType)sh.Type;
                            if ((int)st == 31 || (int)st == 30) { PptShape r = sh; sh = null; return r; }
                        }
                        catch { }
                        try
                        {
                            if (sh.Name != null && sh.Name.IndexOf("3D", StringComparison.OrdinalIgnoreCase) >= 0) { PptShape r = sh; sh = null; return r; }
                        }
                        catch { }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                return null;
            }
            catch { return null; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        /// <summary>
        /// スライド内で名前が指定文字列を含む 3D モデル図形を検索する。
        /// </summary>
        public static PptShape Find3DModelShapeByName(Slide slide, string namePart)
        {
            if (slide == null || string.IsNullOrEmpty(namePart)) return null;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return null;
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        try
                        {
                            var st = (MsoShapeType)sh.Type;
                            if ((int)st != 31 && (int)st != 30) continue;
                        }
                        catch { continue; }
                        try
                        {
                            if (sh.Name != null && sh.Name.IndexOf(namePart, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                PptShape r = sh;
                                sh = null;
                                return r;
                            }
                        }
                        catch { }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                return null;
            }
            catch { return null; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        /// <summary>
        /// スライド内の SmartArt 図形を検索する。
        /// </summary>
        public static PptShape FindSmartArtShape(Slide slide)
        {
            if (slide == null) return null;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return null;
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        if (sh.HasSmartArt == MsoTriState.msoTrue) { PptShape r = sh; sh = null; return r; }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                return null;
            }
            catch { return null; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        /// <summary>
        /// 図形が画像であるか判定する。
        /// </summary>
        public static bool IsPictureShape(PptShape sh)
        {
            try
            {
                int t = (int)sh.Type;
                return t == (int)MsoShapeType.msoPicture || t == 11;
            }
            catch { return false; }
        }

        /// <summary>
        /// スライド内の最初のビデオ（動画）図形を検索する。
        /// </summary>
        /// <param name="slide">スライド。</param>
        /// <returns>該当図形。見つからない場合は null。呼び出し元で Marshal.ReleaseComObject すること。</returns>
        public static PptShape FindFirstVideoShape(Slide slide)
        {
            if (slide == null) return null;
            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                if (shapes == null) return null;
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        // 可能なら MediaType を使う（取得が例外になる環境がある）
                        try
                        {
                            if (sh.MediaType == PpMediaType.ppMediaTypeMovie)
                            {
                                PptShape r = sh;
                                sh = null;
                                return r;
                            }
                        }
                        catch
                        {
                            // ignore
                        }

                        // MediaType が取得できない/不安定な環境向けのフォールバック:
                        // msoMedia で MediaFormat が取得できるものを動画候補として扱う（8-1～8-3 は動画タスク）。
                        try
                        {
                            if (sh.Type == MsoShapeType.msoMedia)
                            {
                                var mf = sh.MediaFormat;
                                if (mf != null)
                                {
                                    try { Marshal.ReleaseComObject(mf); } catch { }
                                    PptShape r = sh;
                                    sh = null;
                                    return r;
                                }
                            }
                        }
                        catch
                        {
                            // ignore
                        }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                return null;
            }
            catch { return null; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        /// <summary>
        /// スライド内（グループ内を含む）で指定テキストを含む図形を探す。
        /// </summary>
        public static PptShape FindShapeWithTextDeep(Slide slide, string searchText)
        {
            if (slide == null || string.IsNullOrEmpty(searchText))
                return null;

            PptShapes shapes = null;
            try
            {
                shapes = slide.Shapes;
                return FindShapeWithTextInShapeCollection(shapes, searchText, searchGroups: true);
            }
            catch { return null; }
            finally
            {
                if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } }
            }
        }

        private static PptShape FindShapeWithTextInShapeCollection(PptShapes shapes, string searchText, bool searchGroups)
        {
            if (shapes == null || string.IsNullOrEmpty(searchText))
                return null;

            try
            {
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = shapes[i];
                        PptShape found = TryFindShapeWithTextOnShape(sh, searchText, searchGroups);
                        if (found != null)
                        {
                            if (ReferenceEquals(found, sh))
                                sh = null;
                            return found;
                        }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                return null;
            }
            catch { return null; }
        }

        private static PptShape TryFindShapeWithTextOnShape(PptShape sh, string searchText, bool searchGroups)
        {
            if (sh == null) return null;

            if (sh.HasTextFrame == MsoTriState.msoTrue)
            {
                string text = null;
                try
                {
                    var pptTf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                    text = pptTf.TextRange?.Text;
                }
                catch { }

                if (TryFindTextSpan(text, searchText, out _, out _))
                {
                    return sh;
                }
            }

            if (!searchGroups || sh.Type != MsoShapeType.msoGroup)
                return null;

            Microsoft.Office.Interop.PowerPoint.GroupShapes group = null;
            try
            {
                group = sh.GroupItems;
                return FindShapeWithTextInGroupShapes(group, searchText, searchGroups: true);
            }
            catch { return null; }
            finally
            {
                if (group != null) { try { Marshal.ReleaseComObject(group); } catch { } }
            }
        }

        private static PptShape FindShapeWithTextInGroupShapes(Microsoft.Office.Interop.PowerPoint.GroupShapes group, string searchText, bool searchGroups)
        {
            if (group == null || string.IsNullOrEmpty(searchText))
                return null;

            try
            {
                int count = group.Count;
                for (int i = 1; i <= count; i++)
                {
                    PptShape sh = null;
                    try
                    {
                        sh = group[i];
                        PptShape found = TryFindShapeWithTextOnShape(sh, searchText, searchGroups);
                        if (found != null)
                        {
                            if (ReferenceEquals(found, sh))
                                sh = null;
                            return found;
                        }
                    }
                    finally
                    {
                        if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                    }
                }
                return null;
            }
            catch { return null; }
        }

        /// <summary>
        /// 図形テキスト内の検索語位置を返す（! / ！ のゆらぎに対応）。
        /// </summary>
        public static bool TryFindTextSpan(string fullText, string searchText, out int start, out int length)
        {
            start = 0;
            length = 0;
            if (string.IsNullOrEmpty(fullText) || string.IsNullOrEmpty(searchText))
                return false;

            int idx = fullText.IndexOf(searchText, StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
            {
                start = idx + 1;
                length = searchText.Length;
                return true;
            }

            string alt = searchText.Replace('!', '！');
            if (!string.Equals(alt, searchText, StringComparison.Ordinal))
            {
                idx = fullText.IndexOf(alt, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                {
                    start = idx + 1;
                    length = alt.Length;
                    return true;
                }
            }

            alt = searchText.Replace('！', '!');
            if (!string.Equals(alt, searchText, StringComparison.Ordinal))
            {
                idx = fullText.IndexOf(alt, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                {
                    start = idx + 1;
                    length = alt.Length;
                    return true;
                }
            }

            string core = searchText.TrimEnd('!', '！');
            if (core.Length > 0 && core.Length < searchText.Length)
            {
                idx = fullText.IndexOf(core, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                {
                    start = idx + 1;
                    length = core.Length;
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 図形内の指定テキスト部分がテーマ色フォントかどうか（部分書式・RGB 解決にも対応）。
        /// </summary>
        public static bool IsShapeSubstringFontThemeColor(PptShape sh, Presentation pres, string searchText, MsoThemeColorIndex themeColor)
        {
            if (sh == null || pres == null || string.IsNullOrEmpty(searchText))
                return false;
            if (sh.HasTextFrame != MsoTriState.msoTrue)
                return false;

            Microsoft.Office.Interop.PowerPoint.TextFrame tf = null;
            try
            {
                tf = (Microsoft.Office.Interop.PowerPoint.TextFrame)sh.TextFrame;
                if (tf == null) return false;
                TextRange tr = null;
                try
                {
                    tr = tf.TextRange;
                    if (tr == null) return false;
                    string fullText = null;
                    try { fullText = tr.Text; } catch { }
                    if (!TryFindTextSpan(fullText, searchText, out int start, out int length))
                        return false;

                    TextRange sub = null;
                    try
                    {
                        sub = tr.Characters(start, length);
                        return sub != null && IsTextRangeFontThemeColor(sub, pres, themeColor);
                    }
                    finally
                    {
                        if (sub != null) { try { Marshal.ReleaseComObject(sub); } catch { } }
                    }
                }
                finally
                {
                    if (tr != null) { try { Marshal.ReleaseComObject(tr); } catch { } }
                }
            }
            finally
            {
                if (tf != null) { try { Marshal.ReleaseComObject(tf); } catch { } }
            }
        }

        private static bool IsTextRangeFontThemeColor(TextRange tr, Presentation pres, MsoThemeColorIndex themeColor)
        {
            if (tr == null) return false;

            Font font = null;
            try
            {
                font = tr.Font;
                if (font == null) return false;
                Microsoft.Office.Interop.PowerPoint.ColorFormat cf = null;
                try
                {
                    cf = font.Color;
                    if (IsFontColorThemeColor(cf, pres, themeColor))
                        return true;
                }
                finally
                {
                    if (cf != null) { try { Marshal.ReleaseComObject(cf); } catch { } }
                }
            }
            finally
            {
                if (font != null) { try { Marshal.ReleaseComObject(font); } catch { } }
            }

            string text = null;
            try { text = tr.Text ?? ""; } catch { }
            int len = text.Length;
            if (len <= 1) return false;

            for (int i = 1; i <= len; i++)
            {
                TextRange ch = null;
                Font chFont = null;
                Microsoft.Office.Interop.PowerPoint.ColorFormat chCf = null;
                try
                {
                    ch = tr.Characters(i, 1);
                    if (ch == null) return false;
                    chFont = ch.Font;
                    if (chFont == null) return false;
                    chCf = chFont.Color;
                    if (!IsFontColorThemeColor(chCf, pres, themeColor))
                        return false;
                }
                finally
                {
                    if (chCf != null) { try { Marshal.ReleaseComObject(chCf); } catch { } }
                    if (chFont != null) { try { Marshal.ReleaseComObject(chFont); } catch { } }
                    if (ch != null) { try { Marshal.ReleaseComObject(ch); } catch { } }
                }
            }
            return true;
        }

        private static bool IsFontColorThemeColor(Microsoft.Office.Interop.PowerPoint.ColorFormat cf, Presentation pres, MsoThemeColorIndex themeColor)
        {
            if (cf == null) return false;
            try
            {
                return cf.ObjectThemeColor == themeColor;
            }
            catch { return false; }
        }
    }
}
