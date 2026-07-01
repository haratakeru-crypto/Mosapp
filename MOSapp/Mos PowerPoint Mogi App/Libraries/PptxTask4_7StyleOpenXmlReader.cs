using System;
using System.IO;
using System.IO.Packaging;
using System.Text.RegularExpressions;

namespace Libraries
{
    /// <summary>P4-7: スライド上テキスト図形の塗りつぶし（アクセント1・白+基本色60％）と枠線（濃い青）を OpenXML で検証する。</summary>
    public static class PptxTask4_7StyleOpenXmlReader
    {
        private const int LineWeightTargetEmu = 9525; // 0.75pt
        private const int LineWeightToleranceEmu = 1200;

        public static bool TryValidateOnSlide(string pptxPath, int slideNumber)
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

                    return ContainsValidTask4_7TextShape(xml);
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool ContainsValidTask4_7TextShape(string slideXml)
        {
            if (string.IsNullOrEmpty(slideXml))
                return false;

            foreach (Match spMatch in Regex.Matches(slideXml, @"<p:sp\b[\s\S]*?</p:sp>", RegexOptions.IgnoreCase))
            {
                string spXml = spMatch.Value;
                if (!Regex.IsMatch(spXml, @"<p:txBody\b", RegexOptions.IgnoreCase))
                    continue;

                string text = ExtractPlainText(spXml);
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                string spPr = ExtractTagInner(spXml, "p:spPr");
                if (string.IsNullOrEmpty(spPr))
                    continue;

                string fillScheme = ExtractSchemeClrBlock(spPr, "a:solidFill");
                string lineBlock = ExtractLineBlock(spPr);
                if (string.IsNullOrEmpty(fillScheme) || string.IsNullOrEmpty(lineBlock))
                    continue;

                if (!IsAccent1FillWhiteBlend60OpenXml(fillScheme))
                    continue;

                if (!IsTask4_7DarkBlueLineOpenXml(lineBlock))
                    continue;

                if (!IsLineWeight075PtOpenXml(lineBlock))
                    continue;

                return true;
            }

            return false;
        }

        private static string ExtractPlainText(string spXml)
        {
            var sb = new System.Text.StringBuilder();
            foreach (Match m in Regex.Matches(spXml, @"<a:t[^>]*>([\s\S]*?)</a:t>", RegexOptions.IgnoreCase))
                sb.Append(System.Net.WebUtility.HtmlDecode(m.Groups[1].Value));
            return sb.ToString();
        }

        private static string ExtractTagInner(string xml, string tagName)
        {
            var match = Regex.Match(xml, $@"<{tagName}\b[^>]*>([\s\S]*?)</{tagName}>", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value : null;
        }

        private static string ExtractLineBlock(string spPrXml)
        {
            var match = Regex.Match(spPrXml, @"<a:ln\b[\s\S]*?</a:ln>", RegexOptions.IgnoreCase);
            return match.Success ? match.Value : null;
        }

        private static string ExtractSchemeClrBlock(string outerXml, string containerTag)
        {
            string container = ExtractTagInner(outerXml, containerTag);
            if (string.IsNullOrEmpty(container))
                return null;

            var match = Regex.Match(container, @"<a:schemeClr\b[\s\S]*?(?:/>|>[\s\S]*?</a:schemeClr>)", RegexOptions.IgnoreCase);
            return match.Success ? match.Value : null;
        }

        internal static bool IsAccent1FillWhiteBlend60OpenXml(string schemeClrXml)
        {
            if (string.IsNullOrEmpty(schemeClrXml))
                return false;
            if (!Regex.IsMatch(schemeClrXml, @"\bval=""accent1""", RegexOptions.IgnoreCase))
                return false;

            if (Regex.IsMatch(schemeClrXml, @"<a:shade\b", RegexOptions.IgnoreCase))
                return false;

            var tint = Regex.Match(schemeClrXml, @"<a:tint\b[^>]*\bval=""(\d+)""", RegexOptions.IgnoreCase);
            if (tint.Success && int.TryParse(tint.Groups[1].Value, out int tintVal))
                return tintVal >= 55000 && tintVal <= 65000;

            var lumOff = Regex.Match(schemeClrXml, @"<a:lumOff\b[^>]*\bval=""(\d+)""", RegexOptions.IgnoreCase);
            if (lumOff.Success && int.TryParse(lumOff.Groups[1].Value, out int lumOffVal)
                && lumOffVal >= 55000 && lumOffVal <= 65000)
                return true;

            var lumMod = Regex.Match(schemeClrXml, @"<a:lumMod\b[^>]*\bval=""(\d+)""", RegexOptions.IgnoreCase);
            if (lumMod.Success && int.TryParse(lumMod.Groups[1].Value, out int lumModVal)
                && lumModVal >= 38000 && lumModVal <= 45000
                && lumOff.Success && int.TryParse(lumOff.Groups[1].Value, out int lumOffVal2)
                && lumOffVal2 >= 55000 && lumOffVal2 <= 65000)
                return true;

            return false;
        }

        private static bool IsTask4_7DarkBlueLineOpenXml(string lineBlock)
        {
            if (string.IsNullOrEmpty(lineBlock))
                return false;

            var srgb = Regex.Match(lineBlock, @"<a:srgbClr\b[^>]*\bval=""([0-9A-Fa-f]{6})""", RegexOptions.IgnoreCase);
            if (!srgb.Success)
                return false;

            return TryParseTask4_7LineSrgb(srgb.Groups[1].Value);
        }

        private static bool TryParseTask4_7LineSrgb(string hex)
        {
            if (string.IsNullOrEmpty(hex) || hex.Length != 6)
                return false;
            if (!int.TryParse(hex.Substring(0, 2), System.Globalization.NumberStyles.HexNumber, null, out int r))
                return false;
            if (!int.TryParse(hex.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, null, out int g))
                return false;
            if (!int.TryParse(hex.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, null, out int b))
                return false;
            return Task4_7LineColorRules.IsDarkBlueRgb(r, g, b);
        }

        private static bool IsLineWeight075PtOpenXml(string lineXml)
        {
            var match = Regex.Match(lineXml, @"\bw=""(\d+)""", RegexOptions.IgnoreCase);
            if (!match.Success || !int.TryParse(match.Groups[1].Value, out int emu))
                return false;
            return Math.Abs(emu - LineWeightTargetEmu) <= LineWeightToleranceEmu;
        }
    }

    /// <summary>P4-7 枠線色判定（COM / OpenXML 共通）。</summary>
    internal static class Task4_7LineColorRules
    {
        private const int RedMax = 20;
        private const int GreenMin = 18;
        private const int GreenMax = 48;
        private const int BlueMin = 72;
        private const int BlueMax = 120;
        private const double LuminanceMin = 0.08;
        private const double LuminanceMax = 0.22;

        public static bool IsDarkBlueRgb(int r, int g, int b)
        {
            double lum = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0;
            if (lum < LuminanceMin || lum > LuminanceMax)
                return false;
            if (r > RedMax)
                return false;
            if (g < GreenMin || g > GreenMax)
                return false;
            if (b < BlueMin || b > BlueMax)
                return false;
            return b > g && g >= r;
        }
    }
}
