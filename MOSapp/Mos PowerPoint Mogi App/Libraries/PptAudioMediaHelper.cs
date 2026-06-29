using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PptShape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace Libraries.Group1
{
    /// <summary>P9-3 等のオーディオ再生設定判定（Checker / VSTO 共用）。</summary>
    public static class PptAudioMediaHelper
    {
        public const float P9_3ExpectedFadeOutMs = 3000f;
        public const float P9_3FadeOutToleranceMs = 500f;
        public const int P9_3PlayAcrossStopAfterSlidesLegacy = 999;

        public static bool IsFadeOutDurationAbout3Seconds(float fadeOutRaw)
        {
            if (Math.Abs(fadeOutRaw - P9_3ExpectedFadeOutMs) < P9_3FadeOutToleranceMs)
                return true;
            // UI で「3」秒と ms 未換算の値が返る環境向け
            if (Math.Abs(fadeOutRaw - 3f) < 0.5f)
                return true;
            return false;
        }

        public static bool TryIsSoundShape(PptShape sh)
        {
            if (sh == null) return false;
            try
            {
                if (sh.Type != MsoShapeType.msoMedia)
                    return false;
            }
            catch { return false; }

            try
            {
                return sh.MediaType == PpMediaType.ppMediaTypeSound;
            }
            catch
            {
                return true;
            }
        }

        public static bool IsAudioPlayAcrossSlides(PptShape sh, Presentation pres)
        {
            if (sh == null) return false;

            MediaFormat mf = null;
            try
            {
                try { mf = sh.MediaFormat; }
                catch { return false; }
                if (mf != null)
                {
                    try
                    {
                        PropertyInfo prop = mf.GetType().GetProperty("PlayAcrossSlides");
                        if (prop != null && prop.PropertyType == typeof(bool))
                        {
                            if ((bool)prop.GetValue(mf))
                                return true;
                        }
                    }
                    catch { }
                }
            }
            finally
            {
                if (mf != null) { try { Marshal.ReleaseComObject(mf); } catch { } }
            }

            AnimationSettings anim = null;
            PlaySettings ps = null;
            try
            {
                try { anim = sh.AnimationSettings; }
                catch { return false; }
                if (anim == null) return false;

                try { ps = anim.PlaySettings; }
                catch { return false; }
                if (ps == null) return false;

                int stopAfter;
                try { stopAfter = ps.StopAfterSlides; }
                catch { return false; }

                if (stopAfter >= P9_3PlayAcrossStopAfterSlidesLegacy)
                    return true;

                if (pres != null)
                {
                    try
                    {
                        int slideCount = pres.Slides.Count;
                        if (slideCount > 1 && stopAfter >= slideCount)
                            return true;
                    }
                    catch { }
                }

                return stopAfter > 1;
            }
            finally
            {
                if (ps != null) { try { Marshal.ReleaseComObject(ps); } catch { } }
                if (anim != null) { try { Marshal.ReleaseComObject(anim); } catch { } }
            }
        }

        public static bool IsAudioFadeOutAbout3Seconds(PptShape sh)
        {
            if (sh == null) return false;
            MediaFormat mf = null;
            try
            {
                try { mf = sh.MediaFormat; }
                catch { return false; }
                if (mf == null) return false;

                float fadeOut;
                try { fadeOut = (float)mf.FadeOutDuration; }
                catch { return false; }

                return IsFadeOutDurationAbout3Seconds(fadeOut);
            }
            finally
            {
                if (mf != null) { try { Marshal.ReleaseComObject(mf); } catch { } }
            }
        }

        public static bool ShapeMatchesP9_3AudioSettings(PptShape sh, Presentation pres)
        {
            return TryIsSoundShape(sh)
                && IsAudioPlayAcrossSlides(sh, pres)
                && IsAudioFadeOutAbout3Seconds(sh);
        }
    }
}
