using System;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace Libraries.Group1
{
    /// <summary>7-2: Project7 の会社名「ラビット出版」（全角）判定。OpenXML の生値で半角カナを拒否。</summary>
    public static class WordP7CompanyValidation
    {
        public const string TargetCompany = "ラビット出版";

        /// <summary>COM + OpenXML で 7-2 正答の会社名か（半角カナ入力は OpenXML 上で ✖）。</summary>
        public static bool IsProject7CompanyValid(Document document)
        {
            if (document == null)
                return false;

            string xmlCompany = null;
            try
            {
                string xml = document.WordOpenXML;
                if (!string.IsNullOrEmpty(xml))
                    TryExtractCompanyFromOpenXml(xml, out xmlCompany);
            }
            catch { }

            if (!string.IsNullOrEmpty(xmlCompany))
            {
                if (ContainsHalfWidthKatakana(xmlCompany))
                    return false;
                return string.Equals(xmlCompany.Trim(), TargetCompany, StringComparison.Ordinal);
            }

            string comCompany = TryGetCompanyFromCom(document);
            if (string.IsNullOrEmpty(comCompany))
                return false;
            if (ContainsHalfWidthKatakana(comCompany))
                return false;
            return string.Equals(comCompany.Trim(), TargetCompany, StringComparison.Ordinal);
        }

        public static bool TryExtractCompanyFromOpenXml(string xml, out string company)
        {
            company = null;
            if (string.IsNullOrEmpty(xml))
                return false;

            Match m = Regex.Match(
                xml,
                @"<(?:cp:)?company\b[^>]*>([\s\S]*?)</(?:cp:)?company>",
                RegexOptions.IgnoreCase);
            if (!m.Success)
                return false;

            company = DecodeXmlText(m.Groups[1].Value);
            return true;
        }

        public static bool ContainsHalfWidthKatakana(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;
            foreach (char c in text)
            {
                if (c >= '\uFF66' && c <= '\uFF9F')
                    return true;
            }
            return false;
        }

        public static string TryGetCompanyFromCom(Document document)
        {
            if (document == null)
                return null;
            try
            {
                dynamic companyProp = ((dynamic)document.BuiltInDocumentProperties)["Company"];
                return companyProp?.Value?.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static string DecodeXmlText(string raw)
        {
            if (string.IsNullOrEmpty(raw))
                return "";
            return System.Net.WebUtility.HtmlDecode(raw.Trim());
        }
    }
}
