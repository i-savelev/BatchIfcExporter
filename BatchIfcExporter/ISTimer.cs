using System;
using System.Diagnostics;

namespace ISTools
{
    public static class ISTimer
    {
        public static Stopwatch Stopwatch {  get; set; }
        static ISTimer()
        {
           
        }

        public static void Start()
        {
            Stopwatch = Stopwatch.StartNew();
        }

        public static string Stop()
        {
            Stopwatch.Stop();
            string time = $"Время выполнения: {Math.Round(Stopwatch.Elapsed.TotalMinutes, 1)} мин. = {Math.Round(Stopwatch.Elapsed.TotalSeconds, 2)} сек.";
            Stopwatch = null;
            return time;
        }
    }
}
