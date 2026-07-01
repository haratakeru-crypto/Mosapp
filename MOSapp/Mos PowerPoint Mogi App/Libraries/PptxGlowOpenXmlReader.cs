using System;
using System.IO;
using System.IO.Packaging;
using System.Text.RegularExpressions;

namespace Libraries
{
    /// <summary>pptx 内の光彩（18pt・アクセント6）を OpenXML で検出する。</summary>
    public static class PptxGlowOpenXmlReader
    {
        public static bool ContainsGlow18ptAccent6OnSlide(string pptxPath, int slideNumber)
        {
            if (string.IsNullOrEmpty(pptxPath) || !File.Exists(pptxPath) || slideNumber < 1)
                return false;

            try
            {
                using (var package = Package.Open(pptxPath, FileMode.Open, FileAccess.Read))
                {
                    string slidePartPath = $"/ppt/slides/slide{slideNumber}.xml";
                    var slideUri = new Uri(slidePartPath, UriKind.Relative);
                    if (!package.PartExists(slideUri))
                        return false;

                    string xml;
                    using (var reader = new StreamReader(package.GetPart(slideUri).GetStream()))
                        xml = reader.ReadToEnd();

                    return ContainsGlow18ptAccent6OpenXml(xml);
                }
            }
            catch
            {
                return false;
            }
        }

        public static bool ContainsGlow18ptAccent6OpenXml(string xml)
        {
            if (string.IsNullOrEmpty(xml))
                return false;

            // 18pt = 12700 EMU/pt → 228600 EMU
            const int targetRad = 228600;
            const int radTolerance = 25000;

            foreach (Match match in Regex.Matches(xml, @"<a:glow\b[^>]*\brad=""(\d+)""[^>]*>([\s\S]*?)</a:glow>", RegexOptions.IgnoreCase))
            {
                if (!int.TryParse(match.Groups[1].Value, out int rad))
                    continue;
                if (Math.Abs(rad - targetRad) > radTolerance)
                    continue;

                string inner = match.Groups[2].Value;
                if (Regex.IsMatch(inner, @"<a:schemeClr\b[^>]*\bval=""accent6""", RegexOptions.IgnoreCase))
                    return true;
            }

            return false;
        }
    }
}
