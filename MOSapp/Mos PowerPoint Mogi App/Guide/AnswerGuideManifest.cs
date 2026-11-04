using System.Collections.Generic;
using Newtonsoft.Json;

namespace MOS_PowerPoint_app.Guide
{
    /// <summary>GuideSteps の manifest.json モデル。</summary>
    public sealed class AnswerGuideManifest
    {
        [JsonProperty("groupId")]
        public int GroupId { get; set; }

        [JsonProperty("projectId")]
        public int ProjectId { get; set; }

        [JsonProperty("taskId")]
        public int TaskId { get; set; }

        [JsonProperty("steps")]
        public List<AnswerGuideStep> Steps { get; set; } = new List<AnswerGuideStep>();
    }

    public sealed class AnswerGuideStep
    {
        [JsonProperty("image")]
        public string Image { get; set; }

        [JsonProperty("caption")]
        public string Caption { get; set; }
    }

    /// <summary>読み込み済みのタスク手順（画像フルパス付き）。</summary>
    public sealed class AnswerGuideTask
    {
        public int TaskId { get; set; }
        public List<AnswerGuideLoadedStep> Steps { get; set; } = new List<AnswerGuideLoadedStep>();
    }

    public sealed class AnswerGuideLoadedStep
    {
        public string ImagePath { get; set; }
        public string Caption { get; set; }
    }
}
