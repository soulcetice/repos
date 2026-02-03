using System;
using System.IO;
using System.Windows.Forms;

namespace NCM_Downloader.Services
{
    public interface ILoggerService
    {
        void Log(string message);
    }

    public class FileLoggerService : ILoggerService
    {
        private readonly string _logFilePath;

        public FileLoggerService(string logFileName)
        {
            _logFilePath = Path.Combine(Application.StartupPath, logFileName);
        }

        public void Log(string message)
        {
            try
            {
                using (var fileWriter = new StreamWriter(_logFilePath, true))
                {
                    DateTime date = DateTime.UtcNow;
                    string logEntry = $"{date:yyyy/MM/dd HH:mm:ss:fff} UTC: {message}";
                    fileWriter.WriteLine(logEntry);
                }
            }
            catch (Exception)
            {
                // Fallback or ignore if logging fails (e.g. file locked)
            }
        }
    }
}
