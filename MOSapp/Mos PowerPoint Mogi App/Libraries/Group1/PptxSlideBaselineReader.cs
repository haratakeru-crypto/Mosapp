using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace Libraries.Group1
{
    /// <summary>
    /// pptx（Open XML）からプレゼンテーション順のスライド ID を読み取る。
    /// 初期ファイルと現在ファイルのスライド構成比較に使用する。
    /// </summary>
    public static class PptxSlideBaselineReader
    {
        private const string RNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        /// <summary>
        /// プレゼンテーション順（1 枚目から）のスライド ID を読み取る。
        /// COM の <see cref="Microsoft.Office.Interop.PowerPoint.Slide.SlideID"/> と一致する。
        /// </summary>
        public static bool TryReadOrderedSlideIds(string pptxFilePath, out List<int> slideIds)
        {
            slideIds = null;
            if (string.IsNullOrWhiteSpace(pptxFilePath) || !File.Exists(pptxFilePath))
                return false;

            try
            {
                using (var fs = new FileStream(pptxFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    var presEntry = zip.GetEntry("ppt/presentation.xml");
                    var presRelsEntry = zip.GetEntry("ppt/_rels/presentation.xml.rels");
                    if (presEntry == null || presRelsEntry == null)
                        return false;

                    var ridToTarget = BuildRidToTargetMap(presRelsEntry);

                    XmlDocument presDoc = new XmlDocument();
                    using (var s = presEntry.Open())
                        presDoc.Load(s);

                    // presentation.xml 内に sldIdLst が複数ある pptx があるため、
                    // 先頭の sldIdLst 直下の sldId のみを採用する（//*[local-name()='sldId'] だと二重計上になる）。
                    XmlNodeList sldIdLstNodes = presDoc.SelectNodes("//*[local-name()='sldIdLst']");
                    if (sldIdLstNodes == null || sldIdLstNodes.Count == 0)
                        return false;

                    var ids = new List<int>();
                    var seenSlideParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (XmlNode sldNode in sldIdLstNodes[0].ChildNodes)
                    {
                        var el = sldNode as XmlElement;
                        if (el == null || !string.Equals(el.LocalName, "sldId", StringComparison.OrdinalIgnoreCase))
                            continue;

                        string rid = el.GetAttribute("id", RNamespace);
                        if (string.IsNullOrEmpty(rid))
                            rid = el.GetAttribute("r:id");
                        if (!string.IsNullOrEmpty(rid) && ridToTarget.TryGetValue(rid, out string target))
                        {
                            string slidePart = NormalizeSlidePartPath(target);
                            if (slidePart == null)
                                continue;
                            if (!seenSlideParts.Add(slidePart))
                                continue;
                        }

                        string idAttr = el.GetAttribute("id");
                        if (string.IsNullOrEmpty(idAttr))
                            continue;
                        if (uint.TryParse(idAttr, out uint slideId) && slideId > 0)
                            ids.Add((int)slideId);
                    }

                    if (ids.Count == 0)
                        return false;

                    slideIds = ids;
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private static Dictionary<string, string> BuildRidToTargetMap(ZipArchiveEntry presRelsEntry)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
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
                    map[id] = target.Replace('\\', '/');
            }
            return map;
        }

        private static string NormalizeSlidePartPath(string target)
        {
            if (string.IsNullOrEmpty(target))
                return null;
            string path = target.Replace('\\', '/');
            if (path.StartsWith("/"))
                path = path.TrimStart('/');
            if (!path.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase))
                path = "ppt/" + path.TrimStart('/');
            if (!path.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase))
                return null;
            return path;
        }
    }
}
