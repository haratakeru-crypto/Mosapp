using System;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace MosPracticeClient
{
    public static class QrUrlBuilder
    {
        public static string BuildCanonical(string university, string name, string classroom, string wrong, string at, string subject)
        {
            return "university=" + (university ?? "")
                + "\nname=" + (name ?? "")
                + "\nclassroom=" + (classroom ?? "")
                + "\nwrong=" + (wrong ?? "")
                + "\nat=" + (at ?? "")
                + "\nsubject=" + (subject ?? "");
        }

        public static string SignHex(string canonical, string key)
        {
            using (var hmac = new HMACSHA256(Encoding.UTF8.GetBytes(key ?? "")))
            {
                byte[] hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(canonical ?? ""));
                var sb = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash)
                    sb.Append(b.ToString("x2", CultureInfo.InvariantCulture));
                return sb.ToString();
            }
        }

        public static string BuildSubmitPageUrl(
            string baseUrl,
            string ingestKey,
            string university,
            string name,
            string classroom,
            int wrongCount,
            DateTime scoredAt,
            string subject,
            string qrPathId = null)
        {
            string at = scoredAt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
            string wrong = wrongCount.ToString(CultureInfo.InvariantCulture);
            string classroomValue = classroom ?? "";
            string canonical = BuildCanonical(university, name, classroomValue, wrong, at, subject);
            string sig = SignHex(canonical, ingestKey);

            var q = new StringBuilder();
            q.Append("university=").Append(Encode(university));
            q.Append("&name=").Append(Encode(name));
            if (!string.IsNullOrEmpty(classroomValue))
                q.Append("&classroom=").Append(Encode(classroomValue));
            q.Append("&wrong=").Append(Encode(wrong));
            q.Append("&at=").Append(Encode(at));
            if (!string.IsNullOrEmpty(subject))
                q.Append("&subject=").Append(Encode(subject));
            q.Append("&sig=").Append(Encode(sig));

            string root = (baseUrl ?? "").Trim().TrimEnd('/');
            return root + MosPracticePublicQrPath.BuildPagePath(qrPathId) + "?" + q;
        }

        static string Encode(string value)
        {
            return Uri.EscapeDataString(value ?? "");
        }
    }
}
