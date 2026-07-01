using System;
using System.IO;
using System.IO.Packaging;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;
using PptShapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using Libraries;

namespace Libraries.Group1
{
    public class PowerPointChecker1_8
    {
        /// <summary>8-1: スライド「スクールの様子」に動画が挿入されているか。</summary>
        public bool CheckTask_1_8_01()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByTitle(pres, "スクールの様子");
                    if (slide == null) return false;
                    PptShape videoShape = null;
                    try
                    {
                        // 結果表示時の採点は「開き直し直後」に走るため、メディア読み込みが間に合わないことがある。
                        // 短時間だけリトライして安定化させる。
                        for (int attempt = 1; attempt <= 5 && videoShape == null; attempt++)
                        {
                            videoShape = PowerPointCheckerCommon.FindFirstVideoShape(slide);
                            if (videoShape == null)
                            {
                                Thread.Sleep(400);
                            }
                        }
                        if (videoShape == null)
                        {
                            return false;
                        }
                        return true;
                    }
                    finally
                    {
                        if (videoShape != null) { try { Marshal.ReleaseComObject(videoShape); } catch { } }
                    }
                }
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>8-2: スライド5にビデオが挿入されているか。</summary>
        public bool CheckTask_1_8_02()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 5);
                    if (slide == null) return false;
                    PptShape videoShape = null;
                    try
                    {
                        for (int attempt = 1; attempt <= 5 && videoShape == null; attempt++)
                        {
                            videoShape = PowerPointCheckerCommon.FindFirstVideoShape(slide);
                            if (videoShape == null)
                            {
                                Thread.Sleep(400);
                            }
                        }
                        if (videoShape == null)
                        {
                            return false;
                        }
                        return true;
                    }
                    finally
                    {
                        if (videoShape != null) { try { Marshal.ReleaseComObject(videoShape); } catch { } }
                    }
                }
                finally
                {
                    if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } }
                }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>8-3: スライド5のビデオを開始00:05・終了00:10にトリム。</summary>
        public bool CheckTask_1_8_03()
        {
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                Slide slide = null;
                try
                {
                    slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 5);
                    if (slide == null) return false;
                    PptShape videoShape = null;
                    try
                    {
                        for (int attempt = 1; attempt <= 5 && videoShape == null; attempt++)
                        {
                            videoShape = PowerPointCheckerCommon.FindFirstVideoShape(slide);
                            if (videoShape == null)
                            {
                                Thread.Sleep(400);
                            }
                        }
                        if (videoShape == null) return false;
                        MediaFormat mf = null;
                        try
                        {
                            mf = videoShape.MediaFormat;
                            if (mf == null) return false;
                            try
                            {
                                float startPt = (float)mf.StartPoint;
                                float endPt = (float)mf.EndPoint;
                                bool ok = Math.Abs(startPt - 5000f) < 500f && Math.Abs(endPt - 10000f) < 500f;
                                return ok;
                            }
                            catch { return false; }
                        }
                        finally { if (mf != null) { try { Marshal.ReleaseComObject(mf); } catch { } } }
                    }
                    finally { if (videoShape != null) { try { Marshal.ReleaseComObject(videoShape); } catch { } } }
                }
                finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>8-4: スライド1のオーディオをスライド切り替えでも1回再生・フェードイン4秒・繰り返し。COM の FadeInDuration または VSTO ログで判定。</summary>
        public bool CheckTask_1_8_04()
        {
            // 過去ログが残っていると誤って true になり得るため、TaskStart 8-4 区間内のみを見る。
            if (PPLogReader.HasMarkerWithinTask(8, 4, "[Task8-4] Audio"))
            {
                return true;
            }
            Presentation pres = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;
                // 結果表示時の採点は開き直し直後に走るため、メディアが遅延ロード中で COM 参照が不安定なことがある。
                // 短時間だけリトライしつつ、Shape 単位で例外を握りつぶして走査を継続する。
                for (int attempt = 1; attempt <= 5; attempt++)
                {
                    Slide slide = null;
                    try
                    {
                        slide = PowerPointCheckerCommon.GetSlideByNumber(pres, 1);
                        if (slide != null)
                        {
                            PptShapes shapes = null;
                            try
                            {
                                shapes = slide.Shapes;
                                if (shapes != null)
                                {
                                    int shapeCount = -1;
                                    try { shapeCount = shapes.Count; } catch { }

                                    int limit = shapeCount > 0 ? shapeCount : 0;
                                    for (int i = 1; i <= limit; i++)
                                    {
                                        PptShape sh = null;
                                        try
                                        {
                                            try
                                            {
                                                sh = shapes[i];
                                            }
                                            catch (Exception ex)
                                            {
                                                continue;
                                            }

                                            // 誤ヒット防止: メディア図形以外は対象外
                                            try
                                            {
                                                if (sh.Type != MsoShapeType.msoMedia)
                                                {
                                                    continue;
                                                }
                                            }
                                            catch
                                            {
                                                // Type 取得に失敗する場合は安全側でスキップ
                                                continue;
                                            }

                                            // 可能なら音声だけに絞る（MediaType が取得できない環境ではフォールバックで続行）
                                            try
                                            {
                                                if (sh.MediaType != PpMediaType.ppMediaTypeSound)
                                                {
                                                    continue;
                                                }
                                            }
                                            catch
                                            {
                                                // fallback to MediaFormat
                                            }

                                            MediaFormat mf = null;
                                            try
                                            {
                                                try
                                                {
                                                    mf = sh.MediaFormat;
                                                }
                                                catch (Exception ex)
                                                {
                                                    continue;
                                                }
                                                if (mf == null)
                                                {
                                                    continue;
                                                }

                                                try
                                                {
                                                    // PlayAcrossSlides/RewindAfterPlaying は PIA で未定義。VSTO で検証予定。
                                                    float fadeIn = (float)mf.FadeInDuration;
                                                    bool ok = Math.Abs(fadeIn - 4000f) < 500f;
                                                    if (ok) return true;
                                                }
                                                catch (Exception ex)
                                                {
                                                    continue;
                                                }
                                            }
                                            finally
                                            {
                                                if (mf != null) { try { Marshal.ReleaseComObject(mf); } catch { } }
                                            }
                                        }
                                        catch { }
                                        finally
                                        {
                                            if (sh != null) { try { Marshal.ReleaseComObject(sh); } catch { } }
                                        }
                                    }
                                }
                            }
                            finally { if (shapes != null) { try { Marshal.ReleaseComObject(shapes); } catch { } } }
                        }
                    }
                    catch { }
                    finally { if (slide != null) { try { Marshal.ReleaseComObject(slide); } catch { } } }

                    Thread.Sleep(400);
                }

                return false;
            }
            catch { return false; }
            finally { if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } } }
        }

        /// <summary>8-5: プレゼンテーションを読み取り専用に設定。pptx内XMLのreadOnlyRecommendedタグを直接解析して判定。</summary>
        public bool CheckTask_1_8_05()
        {
            Presentation pres = null;
            string tempPath = null;
            try
            {
                pres = PowerPointCheckerCommon.GetActivePresentation();
                if (pres == null) return false;

                // 現在のメモリ状態（未保存の設定変更も含む）を一時ファイルに書き出す
                tempPath = Path.Combine(Path.GetTempPath(), "mos_8_5_check_" + Guid.NewGuid().ToString("N") + ".pptx");
                try
                {
                    pres.SaveCopyAs(tempPath);
                }
                catch { return false; }

                // pptx（OPC/ZIP）内の全XMLパートを検索して設定を確認
                try
                {
                    using (var package = Package.Open(tempPath, FileMode.Open, FileAccess.Read))
                    {
                        foreach (var part in package.GetParts())
                        {
                            // XML関連のパートをすべてチェック（presentation.xml, presProps.xml, app.xml等）
                            if (part.Uri.OriginalString.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                            {
                                string xml;
                                using (var reader = new StreamReader(part.GetStream()))
                                    xml = reader.ReadToEnd();

                                // readOnlyRecommended が 1 または true であれば合格
                                if (Regex.IsMatch(xml, @"readOnlyRecommended[^>]*val\s*=\s*""(?:1|true)""", RegexOptions.IgnoreCase))
                                {
                                    return true;
                                }
                            }
                        }
                    }
                }
                catch { }

                // バックアップとしてCOMプロパティ（すでに読取専用として開かれている場合）も確認
                try
                {
                    if (pres.ReadOnly == MsoTriState.msoTrue) return true;
                }
                catch { }

                return false;
            }
            catch { return false; }
            finally
            {
                if (pres != null) { try { Marshal.ReleaseComObject(pres); } catch { } }
                if (tempPath != null && File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { }
                }
            }
        }
    }
}
