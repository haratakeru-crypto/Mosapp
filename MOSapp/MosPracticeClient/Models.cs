using System.Collections.Generic;

namespace MosPracticeClient
{
    public sealed class ExamineeProfile
    {
        public string UniversityName { get; set; }
        public string PersonName { get; set; }
        public string ClassroomName { get; set; }
        public List<string> SubmittedSubjects { get; set; } = new List<string>();
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
