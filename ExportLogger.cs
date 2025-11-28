using System;
using System.IO;
using System.Text;

namespace ChartToSVG
{
    public static class ExportLogger
    {
        private static string GetLogPath()
        {
            try
            {
                // User's documents/chart2svg folder
                var logFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "chart2svg"
                );
                if (!Directory.Exists(logFolder))
                {
                    Directory.CreateDirectory(logFolder);
                }

                return Path.Combine(logFolder, "ChartToSVG.log");
            }
            catch
            {
                // Last resort fallback
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "ChartToSVG.log"
                );
            }
        }

        public static void Log(string message)
        {
            try
            {
                string logPath = GetLogPath();
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string logEntry = $"[{timestamp}] {message}\n";

                // Append to log file (or create if doesn't exist)
                File.AppendAllText(logPath, logEntry);
            }
            catch
            {
                // Silent fail - don't crash the add-in if logging fails
            }
        }

        public static void StartNewLog(string operation)
        {
            try
            {
                string logPath = GetLogPath();
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                // Overwrite existing log file
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("=".PadRight(70, '='));
                sb.AppendLine($"ChartToSVG Export Log");
                sb.AppendLine($"Operation: {operation}");
                sb.AppendLine($"Started: {timestamp}");
                sb.AppendLine("=".PadRight(70, '='));
                sb.AppendLine();

                File.WriteAllText(logPath, sb.ToString());
            }
            catch
            {
                // Silent fail
            }
        }

        public static void LogSuccess(string filePath)
        {
            Log($"✓ SUCCESS: Exported to {filePath}");
            Log($"  File size: {new FileInfo(filePath).Length:N0} bytes");
        }

        public static void LogError(Exception ex)
        {
            Log($"✗ ERROR: {ex.GetType().Name}");
            Log($"  Message: {ex.Message}");
            if (ex.InnerException != null)
            {
                Log($"  Inner: {ex.InnerException.Message}");
            }
        }
    }
}
