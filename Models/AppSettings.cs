using System;

namespace GSheetAutoConverter.Models
{
    public class AppSettings
    {
        public string GSheetFilePath { get; set; } = string.Empty;
        public string OutputXlsxPath { get; set; } = string.Empty;

        // "Interval" or "ScheduledTime"
        public string SyncMode { get; set; } = "Interval";
        public int SyncIntervalMinutes { get; set; } = 5;
        public string ScheduledTimes { get; set; } = "09:00, 14:00, 18:00";

        public bool AutoStartWithWindows { get; set; } = false;
        public bool MinimizeToTrayOnClose { get; set; } = true;
        public string GoogleAuthCookie { get; set; } = string.Empty;
    }
}
