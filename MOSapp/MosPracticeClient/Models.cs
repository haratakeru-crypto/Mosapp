using System.Collections.Generic;

namespace MosPracticeClient
{
    public sealed class ExamineeProfile
    {
        public string UniversityName { get; set; }
        public string PersonName { get; set; }
        public string ClassroomName { get; set; }
        public List<string> SubmittedSubjects { get; set; } = new List<string>();
        /// <summary>Certiport ログインID（メール）。パスワードは保存しない。</summary>
        public string CertiportLoginEmail { get; set; }
        /// <summary>デジタル認定証に表示されていた氏名。</summary>
        public string CertiportDisplayName { get; set; }
        /// <summary>受験者が正しいとした氏名。一致時は MOS お名前と同じ。</summary>
        public string CertiportCorrectName { get; set; }
        public bool CertiportNameMatchesMos { get; set; }
        public string CertiportVerifiedAt { get; set; }
    }

    public sealed class UniversityLookup
    {
        public string Name { get; set; }
        public List<string> Classrooms { get; set; } = new List<string>();
    }

    public sealed class LookupsPayload
    {
        public List<UniversityLookup> Universities { get; set; } = new List<UniversityLookup>();
    }

    public sealed class PracticeSubmitSettings
    {
        public string BaseUrl { get; set; }
        public string IngestKey { get; set; }
        public string QrPathId { get; set; }
    }
}
