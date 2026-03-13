using System;
using System.Diagnostics;

namespace AutoDataEntryProject.Utilities
{
    /// <summary>
    /// مراقب الأداء لقياس الوقت المستغرق
    /// </summary>
    public class PerformanceMonitor
    {
        private Stopwatch _stopwatch;
        private DateTime _startTime;

        public PerformanceMonitor()
        {
            _stopwatch = new Stopwatch();
        }

        public void Start()
        {
            _startTime = DateTime.Now;
            _stopwatch.Start();
        }

        public void Stop()
        {
            _stopwatch.Stop();
        }

        public void Reset()
        {
            _stopwatch.Reset();
            _startTime = DateTime.Now;
        }

        public TimeSpan LogTime()
        {
            return _stopwatch.Elapsed;
        }

        public TimeSpan GetElapsed()
        {
            return _stopwatch.Elapsed;
        }

        public string GetFormattedElapsed()
        {
            var elapsed = _stopwatch.Elapsed;
            if (elapsed.TotalHours >= 1)
                return $"{elapsed:hh\\:mm\\:ss}";
            else
                return $"{elapsed:mm\\:ss}";
        }
    }
}
