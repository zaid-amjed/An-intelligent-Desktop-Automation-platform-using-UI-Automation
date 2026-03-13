using System;

namespace AutoDataEntryProject.Models
{
    public class AutomationProgress
    {
        public int CurrentCount { get; set; }
        public int TotalCount { get; set; }
        public string Status { get; set; }
        public TimeSpan ElapsedTime { get; set; }
        public double ProgressPercentage => TotalCount > 0 ? (CurrentCount * 100.0 / TotalCount) : 0;
        public bool IsCompleted => CurrentCount >= TotalCount;

        public AutomationProgress()
        {
            Status = string.Empty;
            ElapsedTime = TimeSpan.Zero;
        }

        public string GetFormattedStatus()
        {
            return $"[{DateTime.Now:HH:mm:ss}] {Status} | {CurrentCount}/{TotalCount} ({ProgressPercentage:F1}%) | الوقت: {ElapsedTime:mm\\:ss}";
        }
    }
}
