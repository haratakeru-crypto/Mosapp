using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
    }
}
