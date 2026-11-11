using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace MosPracticeClient
{
    public static class CertiportPortal
    {
        public const string LoginUrl = "https://app.certiport.com/portal/login";
        public const string SignupUrl = "https://app.certiport.com/portal/signup?defaultlang=jpn";

        public static string NormalizePersonName(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return "";
            string text = value.Trim().Replace("　", " ");
            text = Regex.Replace(text, @"\s+", " ");
            try
            {
                text = text.Normalize(NormalizationForm.FormKC);
            }
            catch
            {
            }
            return text;
        }

        public static bool NamesMatch(string left, string right)
        {
            string a = NormalizePersonName(left).Replace(" ", "");
            string b = NormalizePersonName(right).Replace(" ", "");
            return a.Length > 0 && string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        public static string FormatVerifiedAt(DateTime value)
        {
            return value.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
        }
    }
}
