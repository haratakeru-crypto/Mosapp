using System;
using System.Globalization;

namespace Libraries
{
    /// <summary>
    /// 教材 / 類題を区別する操作ログのタスク文脈。
    /// VariantSetNo=0 は教材。1以上は類題セット番号。
    /// Word の current_task は Excel と異なり 3 フィールド目が免除フラグ。
    /// </summary>
    public struct WordTaskLogContext : IEquatable<WordTaskLogContext>
    {
        public int ProjectId;
        public int TaskId;
        public int AttemptNo;
        public int VariantSetNo;
        public int ExemptFlags;

        public bool IsValid => ProjectId > 0 && TaskId > 0;
        public bool IsVariant => VariantSetNo > 0;

        public static WordTaskLogContext Create(
            int projectId,
            int taskId,
            int attemptNo = 0,
            int variantSetNo = 0,
            int exemptFlags = 0)
        {
            return new WordTaskLogContext
            {
                ProjectId = projectId,
                TaskId = taskId,
                AttemptNo = attemptNo < 0 ? 0 : attemptNo,
                VariantSetNo = variantSetNo < 0 ? 0 : variantSetNo,
                ExemptFlags = exemptFlags < 0 ? 0 : exemptFlags
            };
        }

        /// <summary>TaskStart / [Op] 用。教材は p-t-a、類題は Vn-p-t-a。</summary>
        public string FormatTaskToken()
        {
            if (VariantSetNo > 0)
                return "V" + VariantSetNo.ToString(CultureInfo.InvariantCulture)
                    + "-" + ProjectId.ToString(CultureInfo.InvariantCulture)
                    + "-" + TaskId.ToString(CultureInfo.InvariantCulture)
                    + "-" + AttemptNo.ToString(CultureInfo.InvariantCulture);
            return ProjectId.ToString(CultureInfo.InvariantCulture)
                + "-" + TaskId.ToString(CultureInfo.InvariantCulture)
                + "-" + AttemptNo.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>mos_word_current_task.txt。教材 p,t,flags,a / 類題 p,t,flags,a,Vn。</summary>
        public string FormatCurrentTaskLine()
        {
            string line = ProjectId.ToString(CultureInfo.InvariantCulture)
                + "," + TaskId.ToString(CultureInfo.InvariantCulture)
                + "," + ExemptFlags.ToString(CultureInfo.InvariantCulture)
                + "," + AttemptNo.ToString(CultureInfo.InvariantCulture);
            if (VariantSetNo > 0)
                line += ",V" + VariantSetNo.ToString(CultureInfo.InvariantCulture);
            return line;
        }

        public string FormatEvidenceProjectToken()
        {
            if (VariantSetNo > 0)
                return "ProjectV" + VariantSetNo.ToString(CultureInfo.InvariantCulture)
                    + "-" + ProjectId.ToString(CultureInfo.InvariantCulture);
            return "Project" + ProjectId.ToString(CultureInfo.InvariantCulture);
        }

        public string FormatEvidenceTaskToken()
        {
            if (VariantSetNo > 0)
                return "TaskV" + VariantSetNo.ToString(CultureInfo.InvariantCulture)
                    + "-" + ProjectId.ToString(CultureInfo.InvariantCulture)
                    + "-" + TaskId.ToString(CultureInfo.InvariantCulture);
            return "Task" + ProjectId.ToString(CultureInfo.InvariantCulture)
                + "-" + TaskId.ToString(CultureInfo.InvariantCulture);
        }

        public string FormatScoringLine(string timestamp, string commandId)
        {
            return "[" + timestamp + "] [" + FormatEvidenceProjectToken() + "] ["
                + FormatEvidenceTaskToken() + "] [" + commandId + "] Executed";
        }

        public string FormatDestructivePrefix()
        {
            if (VariantSetNo > 0)
                return "V" + VariantSetNo.ToString(CultureInfo.InvariantCulture)
                    + "," + ProjectId.ToString(CultureInfo.InvariantCulture)
                    + "," + TaskId.ToString(CultureInfo.InvariantCulture)
                    + "," + AttemptNo.ToString(CultureInfo.InvariantCulture)
                    + ":";
            return ProjectId.ToString(CultureInfo.InvariantCulture)
                + "," + TaskId.ToString(CultureInfo.InvariantCulture)
                + "," + AttemptNo.ToString(CultureInfo.InvariantCulture)
                + ":";
        }

        public string FormatExplicitTaskPrefix()
        {
            return "[Task " + FormatTaskToken() + "]";
        }

        public bool MatchesProjectScope(int projectId, int variantSetNo)
        {
            return ProjectId == projectId && VariantSetNo == (variantSetNo < 0 ? 0 : variantSetNo);
        }

        public bool Equals(WordTaskLogContext other)
        {
            return ProjectId == other.ProjectId
                && TaskId == other.TaskId
                && AttemptNo == other.AttemptNo
                && VariantSetNo == other.VariantSetNo
                && ExemptFlags == other.ExemptFlags;
        }

        public override bool Equals(object obj)
        {
            return obj is WordTaskLogContext other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = ProjectId;
                hash = (hash * 397) ^ TaskId;
                hash = (hash * 397) ^ AttemptNo;
                hash = (hash * 397) ^ VariantSetNo;
                hash = (hash * 397) ^ ExemptFlags;
                return hash;
            }
        }
    }

    public static class WordTaskLogFormat
    {
        public static bool TryParseTaskToken(string token, out WordTaskLogContext context)
        {
            context = default(WordTaskLogContext);
            if (string.IsNullOrWhiteSpace(token))
                return false;

            string[] parts = token.Trim().Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
                return false;

            int offset = 0;
            int variantSetNo = 0;
            if (parts[0].Length >= 2 && (parts[0][0] == 'V' || parts[0][0] == 'v'))
            {
                if (parts.Length < 3)
                    return false;
                if (!int.TryParse(parts[0].Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out variantSetNo)
                    || variantSetNo < 1)
                    return false;
                offset = 1;
            }

            if (!int.TryParse(parts[offset], NumberStyles.Integer, CultureInfo.InvariantCulture, out int projectId)
                || projectId <= 0)
                return false;
            if (!int.TryParse(parts[offset + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int taskId)
                || taskId <= 0)
                return false;

            int attemptNo = 0;
            if (parts.Length > offset + 2)
            {
                if (!int.TryParse(parts[offset + 2], NumberStyles.Integer, CultureInfo.InvariantCulture, out attemptNo)
                    || attemptNo < 0)
                    attemptNo = 0;
            }

            context = WordTaskLogContext.Create(projectId, taskId, attemptNo, variantSetNo);
            return true;
        }

        public static bool TryParseCurrentTaskLine(string line, out WordTaskLogContext context)
        {
            context = default(WordTaskLogContext);
            if (string.IsNullOrWhiteSpace(line))
                return false;

            string[] parts = line.Trim().Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
                return false;

            if (!int.TryParse(parts[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int projectId)
                || projectId <= 0)
                return false;
            if (!int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int taskId)
                || taskId <= 0)
                return false;

            int exemptFlags = 0;
            if (parts.Length >= 3)
                int.TryParse(parts[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out exemptFlags);

            int attemptNo = 0;
            if (parts.Length >= 4)
                int.TryParse(parts[3].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out attemptNo);

            int variantSetNo = 0;
            if (parts.Length >= 5)
            {
                string marker = parts[4].Trim();
                if (marker.Length < 2 || (marker[0] != 'V' && marker[0] != 'v'))
                    return false;
                if (!int.TryParse(marker.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out variantSetNo)
                    || variantSetNo < 1)
                    return false;
            }

            context = WordTaskLogContext.Create(projectId, taskId, attemptNo, variantSetNo, exemptFlags);
            return true;
        }

        public static bool TryParseTaskStartLine(string line, out WordTaskLogContext context)
        {
            context = default(WordTaskLogContext);
            if (string.IsNullOrWhiteSpace(line))
                return false;

            int idx = line.IndexOf("[TaskStart]", StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
                return false;

            string token = line.Substring(idx + 11).Trim();
            int space = token.IndexOf(' ');
            if (space > 0)
                token = token.Substring(0, space);
            return TryParseTaskToken(token, out context);
        }

        public static bool TryParseExplicitTaskPrefix(string line, out WordTaskLogContext context)
        {
            context = default(WordTaskLogContext);
            if (string.IsNullOrWhiteSpace(line))
                return false;

            int idx = line.IndexOf("[Task ", StringComparison.OrdinalIgnoreCase);
            if (idx < 0)
                return false;

            int start = idx + 6;
            int end = line.IndexOf(']', start);
            if (end <= start)
                return false;

            return TryParseTaskToken(line.Substring(start, end - start), out context);
        }

        /// <summary>mos_word_destructive_errors.log の行頭キー（教材 7,3,0: / 類題 V1,7,3,0:）。</summary>
        public static bool TryParseDestructiveLinePrefix(string line, out WordTaskLogContext context)
        {
            context = default(WordTaskLogContext);
            if (string.IsNullOrWhiteSpace(line))
                return false;

            int colon = line.IndexOf(':');
            if (colon <= 0)
                return false;

            string key = line.Substring(0, colon);
            string[] parts = key.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 3)
                return false;

            int offset = 0;
            int variantSetNo = 0;
            string first = parts[0].Trim();
            if (first.Length >= 2 && (first[0] == 'V' || first[0] == 'v'))
            {
                if (parts.Length < 4)
                    return false;
                if (!int.TryParse(first.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out variantSetNo)
                    || variantSetNo < 1)
                    return false;
                offset = 1;
            }

            if (!int.TryParse(parts[offset].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int projectId)
                || projectId <= 0)
                return false;
            if (!int.TryParse(parts[offset + 1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int taskId)
                || taskId <= 0)
                return false;

            int attemptNo = 0;
            if (parts.Length > offset + 2)
            {
                if (!int.TryParse(parts[offset + 2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out attemptNo)
                    || attemptNo < 0)
                    attemptNo = 0;
            }

            context = WordTaskLogContext.Create(projectId, taskId, attemptNo, variantSetNo);
            return true;
        }
    }
}
