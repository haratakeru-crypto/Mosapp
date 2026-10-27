using System;
using System.Globalization;

namespace ExcelAddIn1
{
    internal struct TaskLogContext : IEquatable<TaskLogContext>
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

        public string FormatExplicitTaskPrefix()
        {
            return "[Task " + FormatTaskToken() + "]";
        }

        public bool Equals(TaskLogContext other)
        {
            return ProjectId == other.ProjectId
                && TaskId == other.TaskId
                && AttemptNo == other.AttemptNo
                && VariantSetNo == other.VariantSetNo;
        }
    }

    internal static class TaskLogFormat
    {
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
    }
}
