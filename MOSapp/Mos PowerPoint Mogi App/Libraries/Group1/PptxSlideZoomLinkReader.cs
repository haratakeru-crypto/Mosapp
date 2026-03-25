using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Libraries.Group1
{
    /// <summary>
    /// pptx（Open XML）からスライドズームのリンク先スライドを読み取り、タイトル文字列で検証する。
    /// COM の SlideZoom プロパティが使えない環境向け。
    /// </summary>
    public static class PptxSlideZoomLinkReader
    {
        private const string RNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        /// <summary>
        /// プレゼンテーションの <paramref name="presentationSlideNumber1Based"/> 枚目のスライド上のスライドズームが、
        /// 指定2タイトルをリンク先として持つか検証する。
        /// </summary>
        public static bool TryValidateSlideSlideZoomTargetTitles(
            string pptxFilePath,
            int presentationSlideNumber1Based,
            string requiredTitle1,
            string requiredTitle2,
            out string errorMessage)
        {
            errorMessage = null;
            if (string.IsNullOrEmpty(pptxFilePath) || !File.Exists(pptxFilePath))
            {
                errorMessage = "pptx ファイルが見つかりません";
                return false;
            }

            if (presentationSlideNumber1Based < 1)
            {
                errorMessage = "スライド番号が不正です";
                return false;
            }

            string n1 = NormalizeTitle(requiredTitle1);
            string n2 = NormalizeTitle(requiredTitle2);

            try
            {
                using (var fs = new FileStream(pptxFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    if (zip.GetEntry("ppt/presentation.xml") == null)
                    {
                        errorMessage = "ppt/presentation.xml が存在しません";
                        return false;
                    }

                    var orderedSlidePartPaths = BuildOrderedSlidePartPaths(zip);
                    if (orderedSlidePartPaths == null || orderedSlidePartPaths.Count < presentationSlideNumber1Based)
                    {
                        errorMessage = "プレゼンテーションのスライド一覧を解釈できません";
                        return false;
                    }

                    string firstSlidePart = orderedSlidePartPaths[presentationSlideNumber1Based - 1];

                    string relsPath = "ppt/slides/_rels/" + Path.GetFileName(firstSlidePart) + ".rels";
                    var relsEntry = zip.GetEntry(relsPath);
                    if (relsEntry == null)
                    {
                        errorMessage = "スライドの .rels が見つかりません: " + relsPath;
                        return false;
                    }

                    XmlDocument relsDoc = new XmlDocument();
                    using (var s = relsEntry.Open())
                        relsDoc.Load(s);

                    var ridToTarget = BuildRidToTargetMap(relsDoc);

                    var slideToSlideTargets = CollectSlidePartTargetsFromRelsLoose(relsDoc, firstSlidePart);

                    if (slideToSlideTargets.Count != 2)
                    {
                        string slideXmlPath = firstSlidePart.Replace('\\', '/');
                        var entrySlide = zip.GetEntry(slideXmlPath);
                        if (entrySlide != null)
                        {
                            XmlDocument slideXmlDoc = new XmlDocument();
                            using (var st = entrySlide.Open())
                                slideXmlDoc.Load(st);
                            var fromIds = CollectSlidePartsFromSlideXmlGraphicData(slideXmlDoc, ridToTarget, firstSlidePart);
                            if (fromIds.Count == 2)
                                slideToSlideTargets = fromIds;
                        }
                    }

                    if (slideToSlideTargets.Count != 2)
                    {
                        errorMessage = "スライドズームのリンク先スライドを2件特定できませんでした（.rels と slide XML の両方を試行、実際: " + slideToSlideTargets.Count + "）";
                        return false;
                    }

                    var titles = new List<string>();
                    foreach (string slidePart in slideToSlideTargets)
                    {
                        var entry = zip.GetEntry(slidePart.Replace('\\', '/'));
                        if (entry == null)
                        {
                            errorMessage = "リンク先スライド部分が見つかりません: " + slidePart;
                            return false;
                        }
                        XmlDocument slideDoc = new XmlDocument();
                        using (var st = entry.Open())
                            slideDoc.Load(st);
                        string title = ExtractSlideTitleFromSlidePart(slideDoc);
                        titles.Add(NormalizeTitle(title));
                    }

                    if (TitlesMatchPair(titles[0], titles[1], n1, n2))
                        return true;

                    errorMessage = "リンク先2スライドのタイトルが問題文と一致しません";
                    return false;
                }
            }
            catch (Exception ex)
            {
                errorMessage = "pptx 解析エラー: " + ex.Message;
                return false;
            }
        }

        private static List<string> BuildOrderedSlidePartPaths(ZipArchive zip)
        {
            var presEntry = zip.GetEntry("ppt/presentation.xml");
            var presRelsEntry = zip.GetEntry("ppt/_rels/presentation.xml.rels");
            if (presEntry == null || presRelsEntry == null)
                return null;

            var ridToTarget = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            XmlDocument relsDoc = new XmlDocument();
            using (var s = presRelsEntry.Open())
                relsDoc.Load(s);
            foreach (XmlNode relNode in relsDoc.SelectNodes("//*[local-name()='Relationship']"))
            {
                var rel = relNode as XmlElement;
                if (rel == null) continue;
                string id = rel.GetAttribute("Id");
                string target = rel.GetAttribute("Target");
                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                    ridToTarget[id] = target.Replace('\\', '/');
            }

            XmlDocument presDoc = new XmlDocument();
            using (var s = presEntry.Open())
                presDoc.Load(s);

            var list = new List<string>();
            foreach (XmlNode sldNode in presDoc.SelectNodes("//*[local-name()='sldId']"))
            {
                var sldId = sldNode as XmlElement;
                if (sldId == null) continue;
                string rid = sldId.GetAttribute("id", RNamespace);
                if (string.IsNullOrEmpty(rid))
                    rid = sldId.GetAttribute("r:id");
                if (string.IsNullOrEmpty(rid) || !ridToTarget.TryGetValue(rid, out string target))
                    continue;
                string path = NormalizeZipTargetToSlidesPath(target);
                if (path != null && path.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase))
                    list.Add(path.Replace('\\', '/'));
            }
            return list;
        }

        private static Dictionary<string, string> BuildRidToTargetMap(XmlDocument relsDoc)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (XmlNode relNode in relsDoc.SelectNodes("//*[local-name()='Relationship']"))
            {
                var rel = relNode as XmlElement;
                if (rel == null) continue;
                string id = rel.GetAttribute("Id");
                string target = rel.GetAttribute("Target");
                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                    map[id] = target.Replace('\\', '/');
            }
            return map;
        }

        /// <summary>
        /// .rels を Type に依存せず、Target が ppt/slides/slideN.xml になる参照を集める（Microsoft 拡張の relationship 型も対象）。
        /// </summary>
        private static List<string> CollectSlidePartTargetsFromRelsLoose(XmlDocument relsDoc, string firstSlidePart)
        {
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string firstNorm = firstSlidePart.Replace('\\', '/');
            foreach (XmlNode relNode in relsDoc.SelectNodes("//*[local-name()='Relationship']"))
            {
                var rel = relNode as XmlElement;
                if (rel == null) continue;
                string type = rel.GetAttribute("Type") ?? "";
                string target = rel.GetAttribute("Target");
                if (string.IsNullOrEmpty(target)) continue;
                if (IsExcludedRelationshipForSlideZoom(type, target)) continue;
                string normalized = NormalizeZipTargetToSlidesPath(target);
                if (normalized == null || !normalized.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase))
                    continue;
                normalized = normalized.Replace('\\', '/');
                if (string.Equals(normalized, firstNorm, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (seen.Add(normalized))
                    result.Add(normalized);
            }
            return result;
        }

        private static bool IsExcludedRelationshipForSlideZoom(string type, string target)
        {
            if (string.IsNullOrEmpty(target)) return true;
            if (target.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                target.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                return true;
            if (target.IndexOf("slideLayouts", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (target.IndexOf("notesSlide", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (target.IndexOf("notesSlides", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (type.IndexOf("slideLayout", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (type.IndexOf("notesSlide", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (type.IndexOf("hyperlink", StringComparison.OrdinalIgnoreCase) >= 0 &&
                target.IndexOf("slides/slide", StringComparison.OrdinalIgnoreCase) < 0)
                return true;
            return false;
        }

        /// <summary>
        /// graphicData の uri に slide が含まれるブロックの <c r:id> / p:sld 等からリンク先スライド部分を解決する。
        /// </summary>
        private static List<string> CollectSlidePartsFromSlideXmlGraphicData(
            XmlDocument slideXmlDoc,
            Dictionary<string, string> ridToTarget,
            string firstSlidePart)
        {
            var result = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string firstNorm = firstSlidePart.Replace('\\', '/');
            foreach (XmlNode gdNode in slideXmlDoc.SelectNodes("//*[local-name()='graphicData']"))
            {
                var gd = gdNode as XmlElement;
                if (gd == null) continue;
                string uri = gd.GetAttribute("uri");
                if (string.IsNullOrEmpty(uri) || uri.IndexOf("slide", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                TryResolveRidFromElement(gd, ridToTarget, firstNorm, result, seen);
                foreach (XmlNode n in gd.SelectNodes(".//*"))
                {
                    var el = n as XmlElement;
                    if (el == null) continue;
                    TryResolveRidFromElement(el, ridToTarget, firstNorm, result, seen);
                }
            }
            return result;
        }

        private static void TryResolveRidFromElement(
            XmlElement el,
            Dictionary<string, string> ridToTarget,
            string firstSlidePartNorm,
            List<string> result,
            HashSet<string> seen)
        {
            string rid = GetRelationshipIdFromElement(el);
            if (string.IsNullOrEmpty(rid) || !ridToTarget.TryGetValue(rid, out string target))
                return;
            if (IsExcludedRelationshipForSlideZoom("", target)) return;
            string normalized = NormalizeZipTargetToSlidesPath(target);
            if (normalized == null || !normalized.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase))
                return;
            normalized = normalized.Replace('\\', '/');
            if (string.Equals(normalized, firstSlidePartNorm, StringComparison.OrdinalIgnoreCase))
                return;
            if (seen.Add(normalized))
                result.Add(normalized);
        }

        private static string GetRelationshipIdFromElement(XmlElement el)
        {
            string rid = el.GetAttribute("id", RNamespace);
            if (!string.IsNullOrEmpty(rid)) return rid;
            foreach (XmlAttribute a in el.Attributes)
            {
                if (a.LocalName == "id" && a.NamespaceURI == RNamespace)
                    return a.Value ?? "";
            }
            return "";
        }

        /// <summary>
        /// presentation.xml.rels の Target（例: slides/slide1.xml または ../slides/slide1.xml）を zip 内パスにする。
        /// </summary>
        private static string NormalizeZipTargetToSlidesPath(string target)
        {
            if (string.IsNullOrEmpty(target)) return null;
            string t = target.Replace('\\', '/').Trim();
            if (t.StartsWith("/ppt/", StringComparison.OrdinalIgnoreCase))
                return t.TrimStart('/').Replace('\\', '/');
            if (t.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase))
                return t;
            if (t.StartsWith("../slides/", StringComparison.OrdinalIgnoreCase))
                return "ppt/slides/" + t.Substring("../slides/".Length);
            if (t.StartsWith("../../slides/", StringComparison.OrdinalIgnoreCase))
                return "ppt/slides/" + t.Substring("../../slides/".Length);
            if (t.StartsWith("slides/", StringComparison.OrdinalIgnoreCase))
                return "ppt/" + t;
            if (Regex.IsMatch(t, @"^slide\d+\.xml$", RegexOptions.IgnoreCase))
                return "ppt/slides/" + t;
            return t;
        }

        private static string ExtractSlideTitleFromSlidePart(XmlDocument slideDoc)
        {
            foreach (XmlNode spNode in slideDoc.SelectNodes("//*[local-name()='sp']"))
            {
                var sp = spNode as XmlElement;
                if (sp == null) continue;
                foreach (XmlNode phNode in sp.SelectNodes(".//*[local-name()='ph']"))
                {
                    var ph = phNode as XmlElement;
                    if (ph == null) continue;
                    string typ = ph.GetAttribute("type");
                    if (!string.Equals(typ, "title", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(typ, "ctrTitle", StringComparison.OrdinalIgnoreCase))
                        continue;
                    var sb = new StringBuilder();
                    foreach (XmlNode teNode in sp.SelectNodes(".//*[local-name()='t']"))
                    {
                        string fragment = teNode.InnerText ?? "";
                        sb.Append(fragment);
                    }
                    string result = sb.ToString().Trim();
                    if (result.Length > 0)
                        return result;
                }
            }
            return "";
        }

        private static string NormalizeTitle(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Trim().Replace("\r", "").Replace("\n", "");
        }

        private static bool TitlesMatchPair(string a, string b, string req1, string req2)
        {
            if (string.Equals(a, b, StringComparison.Ordinal))
                return false;
            bool has1a = string.Equals(a, req1, StringComparison.Ordinal);
            bool has2a = string.Equals(a, req2, StringComparison.Ordinal);
            bool has1b = string.Equals(b, req1, StringComparison.Ordinal);
            bool has2b = string.Equals(b, req2, StringComparison.Ordinal);
            return (has1a && has2b) || (has2a && has1b);
        }
    }
}
