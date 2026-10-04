using System;
using System.Collections.Generic;
using System.Globalization;

namespace AxialSqlTools.JobQuickView
{
    internal sealed class JobQuickViewSnapshot
    {
        public Guid JobId { get; internal set; }
        public string Name { get; internal set; }
        public string Owner { get; internal set; }
        public string Category { get; internal set; }
        public string Description { get; internal set; }
        public DateTime? CreatedAt { get; internal set; }
        public DateTime? ModifiedAt { get; internal set; }
        public bool IsEnabled { get; internal set; }
        public int StartStepId { get; internal set; }
        public bool? IsRunning { get; internal set; }
        public string ExecutionStatus { get; internal set; }
        public DateTime? RunningSince { get; internal set; }
        public DateTime? NextRun { get; internal set; }
        public DateTime? LastRunStarted { get; internal set; }
        public TimeSpan? LastRunDuration { get; internal set; }
        public string LastRunOutcome { get; internal set; }
        public string LastRunMessage { get; internal set; }
        public string ActivityWarning { get; internal set; }
        public bool SchedulesLoaded { get; internal set; } = true;
        public bool HistoryLoaded { get; internal set; } = true;
        public IReadOnlyList<JobQuickViewStep> Steps { get; internal set; } = new JobQuickViewStep[0];
        public IReadOnlyList<JobQuickViewSchedule> Schedules { get; internal set; } = new JobQuickViewSchedule[0];

        public const string NextRunNote = "Times use the SQL Server's local time. Next run is an estimate; SQL Server Agent's schedule cache can lag by up to 20 minutes.";
    }

    internal sealed class JobQuickViewStep
    {
        public int StepId { get; internal set; }
        public Guid StepUid { get; internal set; }
        public string Name { get; internal set; }
        public string Subsystem { get; internal set; }
        public string DatabaseName { get; internal set; }
        public string DatabaseUserName { get; internal set; }
        public int ProxyId { get; internal set; }
        public string Command { get; internal set; }
        public string SuccessAction { get; internal set; }
        public string FailureAction { get; internal set; }
        public bool IsSql => string.Equals(Subsystem, "TSQL", StringComparison.OrdinalIgnoreCase);
    }

    internal sealed class JobQuickViewSchedule
    {
        public int ScheduleId { get; internal set; }
        public string Name { get; internal set; }
        public string Description { get; internal set; }
        public bool IsEnabled { get; internal set; }
        public DateTime? NextRun { get; internal set; }
    }

    internal sealed class JobCommandConflictException : InvalidOperationException
    {
        public JobCommandConflictException(string message, Exception innerException)
            : base(message, innerException) { }
    }

    internal static class JobQuickViewValue
    {
        // Agent dates are server-local values, not UTC. Zero means no execution/date.
        public static DateTime? DateTimeFromAgent(int date, int time)
        {
            if (date <= 0 || time < 0)
                return null;

            DateTime parsed;
            return DateTime.TryParseExact(
                date.ToString("D8", CultureInfo.InvariantCulture) + time.ToString("D6", CultureInfo.InvariantCulture),
                "yyyyMMddHHmmss", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed)
                ? (DateTime?)DateTime.SpecifyKind(parsed, DateTimeKind.Unspecified)
                : null;
        }

        // HHmmss is a duration, so the hour part may exceed both 23 and 99.
        public static TimeSpan? DurationFromAgent(int duration)
        {
            if (duration < 0 || duration % 100 >= 60 || duration / 100 % 100 >= 60)
                return null;
            return TimeSpan.FromSeconds((long)(duration / 10000) * 3600
                + duration / 100 % 100 * 60 + duration % 100);
        }

        public static string Outcome(int outcome)
        {
            switch (outcome)
            {
                case 0: return "Failed";
                case 1: return "Succeeded";
                case 2: return "Retry";
                case 3: return "Canceled";
                case 4: return "In progress";
                default: return "Unknown";
            }
        }

        public static bool? IsRunning(int status)
        {
            switch (status)
            {
                case 1:
                case 2:
                case 3:
                case 7: return true;
                case 4:
                case 5: return false;
                default: return null;
            }
        }

        public static string ExecutionStatus(int status)
        {
            switch (status)
            {
                case 1: return "Running";
                case 2: return "Waiting for thread";
                case 3: return "Between retries";
                case 4: return "Idle";
                case 5: return "Suspended";
                case 7: return "Completing";
                default: return "Unknown";
            }
        }

        public static string StepAction(int action, int target)
        {
            switch (action)
            {
                case 1: return "Quit with success";
                case 2: return "Quit with failure";
                case 3: return "Go to next step";
                case 4: return "Go to step " + target.ToString(CultureInfo.InvariantCulture);
                default: return "Unknown";
            }
        }
    }
}
