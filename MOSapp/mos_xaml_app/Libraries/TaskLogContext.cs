using System;
using System.Globalization;

namespace Libraries
{
    /// <summary>
    /// 教材 / 類題を区別する操作ログのタスク文脈。
    /// VariantSetNo=0 は教材。1以上は類題セット番号。
    /// </summary>
    public struct TaskLogContext : IEquatable<TaskLogContext>
    {
        public int ProjectId;
        public int TaskId;
        public int AttemptNo;
        public int VariantSetNo;

        public bool IsValid => ProjectId > 0 && TaskId > 0;
        public bool IsVariant => VariantSetNo > 0;

        public static TaskLogContext Create(int projectId, int taskId, int attemptNo = 1, int variantSetNo = 0)
        {
            return new TaskLogContext
            {
                ProjectId = projectId,
                TaskId = taskId,
                AttemptNo = attemptNo < 1 ? 1 : attemptNo,
                VariantSetNo = variantSetNo < 0 ? 0 : variantSetNo
            };
        }

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

        public string FormatCurrentTaskLine()
        {
            if (VariantSetNo > 0)
                return ProjectId.ToString(CultureInfo.InvariantCulture)
                    + "," + TaskId.ToString(CultureInfo.InvariantCulture)
                    + "," + AttemptNo.ToString(CultureInfo.InvariantCulture)
                    + ",V" + VariantSetNo.ToString(CultureInfo.InvariantCulture);
            return ProjectId.ToString(CultureInfo.InvariantCulture)
                + "," + TaskId.ToString(CultureInfo.InvariantCulture)
                + "," + AttemptNo.ToString(CultureInfo.InvariantCulture);
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
            return ProjectId == projectId && VariantSetNo == variantSetNo;
        }

        public bool Equals(TaskLogContext other)
        {
            return ProjectId == other.ProjectId
                && TaskId == other.TaskId
                && AttemptNo == other.AttemptNo
                && VariantSetNo == other.VariantSetNo;
        }

        public override bool Equals(object obj)
        {
            return obj is TaskLogContext other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = ProjectId;
                hash = (hash * 397) ^ TaskId;
                hash = (hash * 397) ^ AttemptNo;
                hash = (hash * 397) ^ VariantSetNo;
                return hash;
            }
        }
    }

    public static class TaskLogFormat
    {
        public static bool TryParseTaskToken(string token, out TaskLogContext context)
        {
            context = default(TaskLogContext);
            if (string.IsNullOrWhiteSpace(token))
                return false;

            string[] parts = token.Trim().Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 3)
                return false;

            int offset = 0;
            int variantSetNo = 0;
            if (parts[0].Length >= 2 && (parts[0][0] == 'V' || parts[0][0] == 'v'))
            {
                if (parts.Length < 4)
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
            int attemptNo = 1;
            if (parts.Length > offset + 2)
            {
                if (!int.TryParse(parts[offset + 2], NumberStyles.Integer, CultureInfo.InvariantCulture, out attemptNo)
                    || attemptNo < 1)
                    attemptNo = 1;
            }

            context = TaskLogContext.Create(projectId, taskId, attemptNo, variantSetNo);
            return true;
        }

        public static bool TryParseCurrentTaskLine(string line, out TaskLogContext context)
        {
            context = default(TaskLogContext);
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

            int attemptNo = 1;
            if (parts.Length >= 3
                && int.TryParse(parts[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedAttempt)
                && parsedAttempt >= 1)
            {
                attemptNo = parsedAttempt;
            }

            int variantSetNo = 0;
            if (parts.Length >= 4)
            {
                string marker = parts[3].Trim();
                if (marker.Length < 2 || (marker[0] != 'V' && marker[0] != 'v'))
                    return false;
                if (!int.TryParse(marker.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out variantSetNo)
                    || variantSetNo < 1)
                    return false;
            }

            context = TaskLogContext.Create(projectId, taskId, attemptNo, variantSetNo);
            return true;
        }

        public static bool TryParseTaskStartLine(string line, out TaskLogContext context)
        {
            context = default(TaskLogContext);
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

        public static bool TryParseExplicitTaskPrefix(string line, out TaskLogContext context)
        {
            context = default(TaskLogContext);
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

        public static bool TryParseDestructivePrefix(string line, out TaskLogContext context)
        {
            context = default(TaskLogContext);
            if (string.IsNullOrWhiteSpace(line))
                return false;

            int colon = line.IndexOf(':');
            string prefix = colon >= 0 ? line.Substring(0, colon) : line;
            string[] parts = prefix.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 3)
            {
                if (!int.TryParse(parts[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int projectId)
                    || projectId <= 0)
                    return false;
                if (!int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int taskId)
                    || taskId <= 0)
                    return false;
                if (!int.TryParse(parts[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int attemptNo)
                    || attemptNo < 1)
                    attemptNo = 1;
                context = TaskLogContext.Create(projectId, taskId, attemptNo, 0);
                return true;
            }

            if (parts.Length == 4)
            {
                string marker = parts[0].Trim();
                if (marker.Length < 2 || (marker[0] != 'V' && marker[0] != 'v'))
                    return false;
                if (!int.TryParse(marker.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out int variantSetNo)
                    || variantSetNo < 1)
                    return false;
                if (!int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int projectId)
                    || projectId <= 0)
                    return false;
                if (!int.TryParse(parts[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int taskId)
                    || taskId <= 0)
                    return false;
                if (!int.TryParse(parts[3].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int attemptNo)
                    || attemptNo < 1)
                    attemptNo = 1;
                context = TaskLogContext.Create(projectId, taskId, attemptNo, variantSetNo);
                return true;
            }

            return false;
        }
    }
}
