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
        /// 指定タイトルをリンク先としてすべて持つか検証する。
        /// </summary>
        public static bool TryValidateSlideSlideZoomTargetTitles(
            string pptxFilePath,
            int presentationSlideNumber1Based,
            IReadOnlyList<string> requiredTitles,
            out string errorMessage)
        {
            errorMessage = null;
            if (requiredTitles == null || requiredTitles.Count < 1)
            {
                errorMessage = "必須タイトルが指定されていません";
                return false;
            }

            if (requiredTitles.Count == 2)
            {
                return TryValidateSlideSlideZoomTargetTitles(
                    pptxFilePath,
                    presentationSlideNumber1Based,
                    requiredTitles[0],
                    requiredTitles[1],
                    out errorMessage);
            }

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

            var normalizedRequired = new List<string>();
            foreach (string title in requiredTitles)
                normalizedRequired.Add(NormalizeTitle(title));

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

                    if (slideToSlideTargets.Count != requiredTitles.Count)
                    {
                        string slideXmlPath = firstSlidePart.Replace('\\', '/');
                        var entrySlide = zip.GetEntry(slideXmlPath);
                        if (entrySlide != null)
                        {
                            XmlDocument slideXmlDoc = new XmlDocument();
                            using (var st = entrySlide.Open())
                                slideXmlDoc.Load(st);
                            var fromIds = CollectSlidePartsFromSlideXmlGraphicData(slideXmlDoc, ridToTarget, firstSlidePart);
                            if (fromIds.Count == requiredTitles.Count)
                                slideToSlideTargets = fromIds;
                        }
                    }

                    if (slideToSlideTargets.Count != requiredTitles.Count)
                    {
                        errorMessage = "スライドズームのリンク先スライドを"
                            + requiredTitles.Count
                            + "件特定できませんでした（実際: "
                            + slideToSlideTargets.Count
                            + "）";
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
                        titles.Add(NormalizeTitle(ExtractSlideTitleFromSlidePart(slideDoc)));
                    }

                    if (!TitlesMatchRequiredSetFuzzy(titles, normalizedRequired))
                    {
                        errorMessage = "リンク先スライドのタイトルが問題文と一致しません（実際: "
                            + string.Join(" | ", titles)
                            + "）";
                        return false;
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                errorMessage = "pptx 解析エラー: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// セクションズームが、対応するラベル文字の真下にあり、リンク先セクション名も一致するか検証する（P3-7）。
        /// </summary>
        public static bool TryValidateSectionZoomPlacedUnderLabels(
            string pptxFilePath,
            int presentationSlideNumber1Based,
            string labelText1,
            string expectedSectionName1,
            string labelText2,
            string expectedSectionName2,
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

            try
            {
                using (var fs = new FileStream(pptxFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    var sectionGuidToName = BuildSectionGuidToNameMap(zip);
                    if (sectionGuidToName == null || sectionGuidToName.Count == 0)
                    {
                        errorMessage = "セクション定義を解釈できません";
                        return false;
                    }

                    var orderedSlidePartPaths = BuildOrderedSlidePartPaths(zip);
                    if (orderedSlidePartPaths == null || orderedSlidePartPaths.Count < presentationSlideNumber1Based)
                    {
                        errorMessage = "プレゼンテーションのスライド一覧を解釈できません";
                        return false;
                    }

                    string slidePart = orderedSlidePartPaths[presentationSlideNumber1Based - 1];
                    var slideEntry = zip.GetEntry(slidePart.Replace('\\', '/'));
                    if (slideEntry == null)
                    {
                        errorMessage = "スライド部分が見つかりません: " + slidePart;
                        return false;
                    }

                    XmlDocument slideDoc = new XmlDocument();
                    using (var st = slideEntry.Open())
                        slideDoc.Load(st);

                    var labelPairs = new[]
                    {
                        Tuple.Create(labelText1, expectedSectionName1),
                        Tuple.Create(labelText2, expectedSectionName2)
                    };

                    var labels = CollectMatchingTextLabelsOnSlide(slideDoc, labelPairs);
                    var zooms = CollectSectionZoomsWithBoundsFromSlideXml(slideDoc, sectionGuidToName);

                    if (zooms.Count != 2)
                    {
                        errorMessage = "セクションズームを2件特定できませんでした（実際: " + zooms.Count + "）";
                        return false;
                    }

                    var usedLabelIndices = new HashSet<int>();
                    var usedZoomIndices = new HashSet<int>();

                    foreach (var pair in labelPairs)
                    {
                        string requiredLabel = pair.Item1;
                        string requiredSection = pair.Item2;
                        string requiredLabelNorm = NormalizeTitleForFuzzyMatch(requiredLabel);
                        string requiredSectionNorm = NormalizeTitleForFuzzyMatch(requiredSection);

                        int labelIndex = -1;
                        for (int i = 0; i < labels.Count; i++)
                        {
                            if (usedLabelIndices.Contains(i)) continue;
                            if (TitleFuzzyEquals(labels[i].TextNorm, requiredLabelNorm))
                            {
                                labelIndex = i;
                                break;
                            }
                        }

                        if (labelIndex < 0)
                        {
                            errorMessage = "ラベル文字が見つかりません: " + requiredLabel;
                            return false;
                        }
                        usedLabelIndices.Add(labelIndex);

                        TextLabelOnSlide label = labels[labelIndex];
                        int zoomIndex = -1;
                        for (int i = 0; i < zooms.Count; i++)
                        {
                            if (usedZoomIndices.Contains(i)) continue;
                            if (!TitleFuzzyEquals(zooms[i].SectionNameNorm, requiredSectionNorm)) continue;
                            if (!IsSectionZoomBelowLabel(zooms[i], label)) continue;
                            zoomIndex = i;
                            break;
                        }

                        if (zoomIndex < 0)
                        {
                            errorMessage = "「" + requiredLabel + "」の下にセクション「" + requiredSection + "」へのズームがありません";
                            return false;
                        }
                        usedZoomIndices.Add(zoomIndex);
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                errorMessage = "pptx 解析エラー: " + ex.Message;
                return false;
            }
        }

        private const double EmuPerPoint = 12700.0;
        private const double SectionZoomBelowLabelTolerancePt = 5.0;
        private const double SectionZoomHorizontalAlignMarginPt = 40.0;

        private sealed class SlideBoundsPt
        {
            public double Left;
            public double Top;
            public double Width;
            public double Height;
            public double Right => Left + Width;
            public double Bottom => Top + Height;
            public double CenterX => Left + Width / 2.0;
        }

        private sealed class TextLabelOnSlide
        {
            public string TextRaw;
            public string TextNorm;
            public SlideBoundsPt Bounds;
        }

        private sealed class SectionZoomOnSlide
        {
            public string SectionNameRaw;
            public string SectionNameNorm;
            public SlideBoundsPt Bounds;
        }

        private static List<TextLabelOnSlide> CollectMatchingTextLabelsOnSlide(
            XmlDocument slideDoc,
            Tuple<string, string>[] requiredPairs)
        {
            var requiredNorms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in requiredPairs)
            {
                string norm = NormalizeTitleForFuzzyMatch(pair.Item1);
                if (!string.IsNullOrEmpty(norm))
                    requiredNorms.Add(norm);
            }

            var results = new List<TextLabelOnSlide>();
            foreach (XmlNode spNode in slideDoc.SelectNodes("//*[local-name()='sp']"))
            {
                var sp = spNode as XmlElement;
                if (sp == null) continue;

                string textRaw = ExtractAllTextFromShapeXml(sp);
                if (string.IsNullOrWhiteSpace(textRaw)) continue;

                string textNorm = NormalizeTitleForFuzzyMatch(textRaw);
                bool matchesRequired = false;
                foreach (string req in requiredNorms)
                {
                    if (TitleFuzzyEquals(textNorm, req))
                    {
                        matchesRequired = true;
                        break;
                    }
                }
                if (!matchesRequired) continue;

                if (!TryParseShapeBoundsPt(sp, out SlideBoundsPt bounds)) continue;

                results.Add(new TextLabelOnSlide
                {
                    TextRaw = textRaw.Trim(),
                    TextNorm = textNorm,
                    Bounds = bounds
                });
            }
            return results;
        }

        private static List<SectionZoomOnSlide> CollectSectionZoomsWithBoundsFromSlideXml(
            XmlDocument slideDoc,
            Dictionary<string, string> sectionGuidToName)
        {
            var results = new List<SectionZoomOnSlide>();

            foreach (XmlNode gfNode in slideDoc.SelectNodes("//*[local-name()='graphicFrame']"))
            {
                var graphicFrame = gfNode as XmlElement;
                if (graphicFrame == null) continue;

                var gd = graphicFrame.SelectSingleNode(".//*[local-name()='graphicData']") as XmlElement;
                if (gd == null) continue;

                string uri = GetXmlAttributeByLocalName(gd, "uri");
                if (string.IsNullOrEmpty(uri) || uri.IndexOf("sectionzoom", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                string sectionId = null;
                foreach (XmlNode n in gd.SelectNodes(".//*"))
                {
                    var el = n as XmlElement;
                    if (el == null) continue;
                    sectionId = GetXmlAttributeByLocalName(el, "sectionId");
                    if (string.IsNullOrEmpty(sectionId))
                        sectionId = GetXmlAttributeByLocalName(el, "sectionGuid");
                    if (!string.IsNullOrEmpty(sectionId)) break;
                }
                if (string.IsNullOrEmpty(sectionId)) continue;
                if (!TryLookupSectionName(sectionGuidToName, sectionId, out string sectionName)) continue;
                if (!TryParseShapeBoundsPt(graphicFrame, out SlideBoundsPt bounds)) continue;

                results.Add(new SectionZoomOnSlide
                {
                    SectionNameRaw = sectionName,
                    SectionNameNorm = NormalizeTitleForFuzzyMatch(sectionName),
                    Bounds = bounds
                });
            }

            if (results.Count == 0)
            {
                foreach (XmlNode gdNode in slideDoc.SelectNodes("//*[local-name()='graphicData']"))
                {
                    var gd = gdNode as XmlElement;
                    if (gd == null) continue;
                    string uri = GetXmlAttributeByLocalName(gd, "uri");
                    if (string.IsNullOrEmpty(uri) || uri.IndexOf("sectionzoom", StringComparison.OrdinalIgnoreCase) < 0)
                        continue;

                    string sectionId = null;
                    foreach (XmlNode n in gd.SelectNodes(".//*"))
                    {
                        var el = n as XmlElement;
                        if (el == null) continue;
                        sectionId = GetXmlAttributeByLocalName(el, "sectionId");
                        if (string.IsNullOrEmpty(sectionId))
                            sectionId = GetXmlAttributeByLocalName(el, "sectionGuid");
                        if (!string.IsNullOrEmpty(sectionId)) break;
                    }
                    if (string.IsNullOrEmpty(sectionId)) continue;
                    if (!TryLookupSectionName(sectionGuidToName, sectionId, out string sectionName)) continue;

                    var host = FindAncestorByLocalName(gd, "graphicFrame") as XmlElement;
                    if (host == null) continue;
                    if (!TryParseShapeBoundsPt(host, out SlideBoundsPt bounds)) continue;

                    results.Add(new SectionZoomOnSlide
                    {
                        SectionNameRaw = sectionName,
                        SectionNameNorm = NormalizeTitleForFuzzyMatch(sectionName),
                        Bounds = bounds
                    });
                }
            }

            return results;
        }

        private static XmlNode FindAncestorByLocalName(XmlNode node, string localName)
        {
            for (XmlNode p = node?.ParentNode; p != null; p = p.ParentNode)
            {
                if (p is XmlElement el && el.LocalName == localName)
                    return el;
            }
            return null;
        }

        private static string ExtractAllTextFromShapeXml(XmlElement shape)
        {
            var sb = new StringBuilder();
            foreach (XmlNode tNode in shape.SelectNodes(".//*[local-name()='t']"))
            {
                if (tNode != null)
                    sb.Append(tNode.InnerText);
            }
            return sb.ToString();
        }

        private static bool TryParseShapeBoundsPt(XmlElement container, out SlideBoundsPt bounds)
        {
            bounds = new SlideBoundsPt();
            if (container == null) return false;

            XmlElement xfrm = null;
            if (string.Equals(container.LocalName, "graphicFrame", StringComparison.OrdinalIgnoreCase))
                xfrm = container.SelectSingleNode("*[local-name()='xfrm']") as XmlElement;
            if (xfrm == null)
                xfrm = container.SelectSingleNode(".//*[local-name()='spPr']/*[local-name()='xfrm']") as XmlElement;
            if (xfrm == null)
                xfrm = container.SelectSingleNode(".//*[local-name()='xfrm']") as XmlElement;
            if (xfrm == null) return false;

            XmlElement off = xfrm.SelectSingleNode("*[local-name()='off']") as XmlElement;
            XmlElement ext = xfrm.SelectSingleNode("*[local-name()='ext']") as XmlElement;
            if (off == null || ext == null) return false;

            if (!long.TryParse(GetXmlAttributeByLocalName(off, "x"), out long xEmu)) return false;
            if (!long.TryParse(GetXmlAttributeByLocalName(off, "y"), out long yEmu)) return false;
            if (!long.TryParse(GetXmlAttributeByLocalName(ext, "cx"), out long cxEmu)) return false;
            if (!long.TryParse(GetXmlAttributeByLocalName(ext, "cy"), out long cyEmu)) return false;

            bounds.Left = xEmu / EmuPerPoint;
            bounds.Top = yEmu / EmuPerPoint;
            bounds.Width = cxEmu / EmuPerPoint;
            bounds.Height = cyEmu / EmuPerPoint;
            return bounds.Width > 0 && bounds.Height > 0;
        }

        private static bool IsSectionZoomBelowLabel(SectionZoomOnSlide zoom, TextLabelOnSlide label)
        {
            if (zoom?.Bounds == null || label?.Bounds == null) return false;

            if (zoom.Bounds.Top < label.Bounds.Bottom - SectionZoomBelowLabelTolerancePt)
                return false;

            double margin = SectionZoomHorizontalAlignMarginPt;
            return zoom.Bounds.CenterX >= label.Bounds.Left - margin
                && zoom.Bounds.CenterX <= label.Bounds.Right + margin;
        }

        private static Dictionary<string, string> BuildSectionGuidToNameMap(ZipArchive zip)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var presEntry = zip.GetEntry("ppt/presentation.xml");
            if (presEntry == null) return map;

            XmlDocument presDoc = new XmlDocument();
            using (var s = presEntry.Open())
                presDoc.Load(s);

            foreach (XmlNode sectionNode in presDoc.SelectNodes("//*[local-name()='section']"))
            {
                var section = sectionNode as XmlElement;
                if (section == null) continue;
                string guid = section.GetAttribute("id");
                if (string.IsNullOrEmpty(guid))
                {
                    foreach (XmlAttribute attr in section.Attributes)
                    {
                        if (attr.LocalName == "id")
                        {
                            guid = attr.Value ?? "";
                            break;
                        }
                    }
                }
                if (string.IsNullOrEmpty(guid)) continue;

                string name = section.GetAttribute("name");
                if (string.IsNullOrEmpty(name))
                {
                    foreach (XmlAttribute attr in section.Attributes)
                    {
                        if (attr.LocalName == "name")
                        {
                            name = attr.Value ?? "";
                            break;
                        }
                    }
                }
                if (!string.IsNullOrEmpty(name))
                    map[NormalizeSectionGuidKey(guid)] = name;
            }
            return map;
        }

        private static string NormalizeSectionGuidKey(string guid)
        {
            if (string.IsNullOrEmpty(guid)) return "";
            return guid.Trim().Trim('{', '}');
        }

        private static bool TryLookupSectionName(Dictionary<string, string> sectionGuidToName, string sectionId, out string sectionName)
        {
            sectionName = null;
            if (string.IsNullOrEmpty(sectionId) || sectionGuidToName == null) return false;

            if (sectionGuidToName.TryGetValue(sectionId, out sectionName)) return true;

            string norm = NormalizeSectionGuidKey(sectionId);
            if (sectionGuidToName.TryGetValue(norm, out sectionName)) return true;

            foreach (var kvp in sectionGuidToName)
            {
                if (string.Equals(NormalizeSectionGuidKey(kvp.Key), norm, StringComparison.OrdinalIgnoreCase))
                {
                    sectionName = kvp.Value;
                    return true;
                }
            }
            return false;
        }

        private static string GetXmlAttributeByLocalName(XmlElement el, string localName)
        {
            if (el == null || string.IsNullOrEmpty(localName)) return "";
            string value = el.GetAttribute(localName);
            if (!string.IsNullOrEmpty(value)) return value;
            foreach (XmlAttribute attr in el.Attributes)
            {
                if (attr.LocalName == localName)
                    return attr.Value ?? "";
            }
            return "";
        }

        private static bool TitlesMatchRequiredSetFuzzy(List<string> actualTitles, List<string> requiredNorm)
        {
            if (actualTitles == null || requiredNorm == null || actualTitles.Count != requiredNorm.Count)
                return false;

            var remaining = new List<string>(requiredNorm);
            foreach (string actual in actualTitles)
            {
                int matchIndex = -1;
                for (int i = 0; i < remaining.Count; i++)
                {
                    if (TitleFuzzyEquals(actual, remaining[i]))
                    {
                        matchIndex = i;
                        break;
                    }
                }
                if (matchIndex < 0)
                    return false;
                remaining.RemoveAt(matchIndex);
            }
            return remaining.Count == 0;
        }

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
                    string result = ExtractShapeText(sp);
                    if (result.Length > 0)
                        return result;
                }
            }

            // リンク先スライドの見出しが本文プレースホルダーのみの場合のフォールバック
            foreach (XmlNode spNode in slideDoc.SelectNodes("//*[local-name()='sp']"))
            {
                var sp = spNode as XmlElement;
                if (sp == null) continue;
                string result = ExtractShapeText(sp);
                if (result.Length > 0)
                    return result;
            }
            return "";
        }

        private static string ExtractShapeText(XmlElement sp)
        {
            var sb = new StringBuilder();
            foreach (XmlNode teNode in sp.SelectNodes(".//*[local-name()='t']"))
            {
                string fragment = teNode.InnerText ?? "";
                sb.Append(fragment);
            }
            return sb.ToString().Trim();
        }

        private static string NormalizeTitle(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return NormalizeTitleForFuzzyMatch(s);
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

        /// <summary>
        /// サマリーズームスライド上のズームリンク先を Open XML で検証する。
        /// 指定2タイトルへのリンクがちょうど2件あり、禁止スライド番号へのリンクがないことを確認する。
        /// </summary>
        public static bool TryValidateSummaryZoomTargetTitles(
            string pptxFilePath,
            int summaryZoomSlideNumber1Based,
            string requiredTitle1,
            string requiredTitle2,
            IReadOnlyList<int> forbiddenTargetSlideIndices1Based,
            out string errorMessage)
        {
            errorMessage = null;
            if (string.IsNullOrEmpty(pptxFilePath) || !File.Exists(pptxFilePath))
            {
                errorMessage = "pptx ファイルが見つかりません";
                return false;
            }

            if (summaryZoomSlideNumber1Based < 1)
            {
                errorMessage = "サマリーズームのスライド番号が不正です";
                return false;
            }

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
                    if (orderedSlidePartPaths == null || orderedSlidePartPaths.Count < summaryZoomSlideNumber1Based)
                    {
                        errorMessage = "プレゼンテーションのスライド一覧を解釈できません";
                        return false;
                    }

                    string summarySlidePart = orderedSlidePartPaths[summaryZoomSlideNumber1Based - 1];
                    string relsPath = "ppt/slides/_rels/" + Path.GetFileName(summarySlidePart) + ".rels";
                    var relsEntry = zip.GetEntry(relsPath);
                    if (relsEntry == null)
                    {
                        errorMessage = "サマリーズームスライドの .rels が見つかりません: " + relsPath;
                        return false;
                    }

                    XmlDocument relsDoc = new XmlDocument();
                    using (var s = relsEntry.Open())
                        relsDoc.Load(s);

                    var ridToTarget = BuildRidToTargetMap(relsDoc);
                    var slideToSlideTargets = CollectSlidePartTargetsFromRelsLoose(relsDoc, summarySlidePart);

                    if (slideToSlideTargets.Count != 2)
                    {
                        string slideXmlPath = summarySlidePart.Replace('\\', '/');
                        var entrySlide = zip.GetEntry(slideXmlPath);
                        if (entrySlide != null)
                        {
                            XmlDocument slideXmlDoc = new XmlDocument();
                            using (var st = entrySlide.Open())
                                slideXmlDoc.Load(st);
                            var fromIds = CollectSlidePartsFromSlideXmlGraphicData(slideXmlDoc, ridToTarget, summarySlidePart);
                            if (fromIds.Count == 2)
                                slideToSlideTargets = fromIds;
                        }
                    }

                    if (slideToSlideTargets.Count != 2)
                    {
                        errorMessage = "サマリーズームのリンク先スライドを2件特定できませんでした（実際: " + slideToSlideTargets.Count + "）";
                        return false;
                    }

                    var forbidden = new HashSet<int>();
                    if (forbiddenTargetSlideIndices1Based != null)
                    {
                        foreach (int idx in forbiddenTargetSlideIndices1Based)
                        {
                            if (idx >= 1) forbidden.Add(idx);
                        }
                    }

                    var linkedTitles = new List<string>();
                    foreach (string slidePart in slideToSlideTargets)
                    {
                        int slideIndex = GetSlidePartIndex1Based(orderedSlidePartPaths, slidePart);
                        if (slideIndex < 1)
                        {
                            errorMessage = "リンク先スライドの位置を特定できません: " + slidePart;
                            return false;
                        }

                        if (forbidden.Contains(slideIndex))
                        {
                            errorMessage = "禁止されたスライド（スライド " + slideIndex + "）へのリンクが含まれています";
                            return false;
                        }

                        var entry = zip.GetEntry(slidePart.Replace('\\', '/'));
                        if (entry == null)
                        {
                            errorMessage = "リンク先スライド部分が見つかりません: " + slidePart;
                            return false;
                        }

                        XmlDocument slideDoc = new XmlDocument();
                        using (var st = entry.Open())
                            slideDoc.Load(st);
                        linkedTitles.Add(ExtractSlideTitleFromSlidePart(slideDoc));
                    }

                    if (!TitlesMatchRequiredPairFuzzy(linkedTitles, requiredTitle1, requiredTitle2))
                    {
                        errorMessage = "リンク先2スライドのタイトルが問題文と一致しません";
                        return false;
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                errorMessage = "pptx 解析エラー: " + ex.Message;
                return false;
            }
        }

        private static int GetSlidePartIndex1Based(List<string> orderedSlidePartPaths, string slidePart)
        {
            if (orderedSlidePartPaths == null || string.IsNullOrEmpty(slidePart))
                return -1;
            string norm = slidePart.Replace('\\', '/');
            for (int i = 0; i < orderedSlidePartPaths.Count; i++)
            {
                if (string.Equals(orderedSlidePartPaths[i].Replace('\\', '/'), norm, StringComparison.OrdinalIgnoreCase))
                    return i + 1;
            }
            return -1;
        }

        private static string NormalizeTitleForFuzzyMatch(string s)
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

        /// <summary>全角数字・句読点などを半角に揃える（教材スライドの「１.教育理念」等）。</summary>
        private static char NormalizeWidthInsensitiveChar(char c)
        {
            if (c >= '\uFF10' && c <= '\uFF19')
                return (char)('0' + (c - '\uFF10'));
            if (c == '\uFF0E')
                return '.';
            if (c == '\uFF01')
                return '!';
            if (c == '\uFF1F')
                return '?';
            return c;
        }

        private static bool TitlesMatchRequiredPairFuzzy(List<string> actualTitles, string req1, string req2)
        {
            if (actualTitles == null || actualTitles.Count != 2)
                return false;

            string n1 = NormalizeTitleForFuzzyMatch(req1);
            string n2 = NormalizeTitleForFuzzyMatch(req2);
            string a = NormalizeTitleForFuzzyMatch(actualTitles[0]);
            string b = NormalizeTitleForFuzzyMatch(actualTitles[1]);

            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b) || string.Equals(a, b, StringComparison.OrdinalIgnoreCase))
                return false;

            return (TitleFuzzyEquals(a, n1) && TitleFuzzyEquals(b, n2))
                || (TitleFuzzyEquals(a, n2) && TitleFuzzyEquals(b, n1));
        }

        private static bool TitleFuzzyEquals(string actualNorm, string requiredNorm)
        {
            if (string.IsNullOrEmpty(actualNorm) || string.IsNullOrEmpty(requiredNorm))
                return false;
            if (string.Equals(actualNorm, requiredNorm, StringComparison.OrdinalIgnoreCase))
                return true;
            return actualNorm.IndexOf(requiredNorm, StringComparison.OrdinalIgnoreCase) >= 0
                || requiredNorm.IndexOf(actualNorm, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
