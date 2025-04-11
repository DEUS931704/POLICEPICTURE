using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Diagnostics;

namespace POLICEPICTURE
{
    /// <summary>
    /// 日誌記錄類 - 提供統一的日誌記錄功能
    /// </summary>
    public static class Logger
    {
        // 日誌檔案路徑
        private static readonly string LogFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "POLICEPICTURE",
            "logs",
            $"log_{DateTime.Now:yyyyMMdd}.txt");

        // 日誌鎖對象 - 用於確保多線程寫入安全
        private static readonly object LogLock = new object();

        // 是否初始化
        private static bool _initialized = false;

        /// <summary>
        /// 日誌級別
        /// </summary>
        public enum LogLevel
        {
            Info,
            Warning,
            Error,
            Debug
        }

        /// <summary>
        /// 初始化日誌系統
        /// </summary>
        public static void Initialize()
        {
            if (_initialized) return;

            try
            {
                // 確保日誌目錄存在
                string logDir = Path.GetDirectoryName(LogFilePath);
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }

                // 寫入日誌頭
                lock (LogLock)
                {
                    using (StreamWriter writer = new StreamWriter(LogFilePath, true, Encoding.UTF8))
                    {
                        writer.WriteLine("==============================================");
                        writer.WriteLine($"應用程式啟動 - {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                        writer.WriteLine("==============================================");
                    }
                }

                _initialized = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"初始化日誌系統時發生錯誤: {ex.Message}");
                // 無法初始化日誌系統，僅寫入調試輸出
            }
        }

        /// <summary>
        /// 記錄日誌
        /// </summary>
        /// <param name="message">日誌消息</param>
        /// <param name="level">日誌級別</param>
        public static void Log(string message, LogLevel level = LogLevel.Info)
        {
            try
            {
                // 確保已初始化
                if (!_initialized)
                {
                    Initialize();
                }

                // 組織日誌信息
                StringBuilder logEntry = new StringBuilder();
                logEntry.Append($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] ");
                logEntry.Append($"[{level}] ");
                logEntry.Append($"[Thread: {Thread.CurrentThread.ManagedThreadId}] ");
                logEntry.Append(message);

                // 寫入日誌
                lock (LogLock)
                {
                    using (StreamWriter writer = new StreamWriter(LogFilePath, true, Encoding.UTF8))
                    {
                        writer.WriteLine(logEntry.ToString());
                    }
                }

                // 同時輸出到調試窗口
                Debug.WriteLine(logEntry.ToString());
            }
            catch (Exception ex)
            {
                // 記錄日誌本身發生錯誤，僅輸出到調試窗口
                Debug.WriteLine($"記錄日誌時發生錯誤: {ex.Message}");
                Debug.WriteLine($"原始日誌: {message}");
            }
        }

        /// <summary>
        /// 清理舊日誌文件
        /// </summary>
        /// <param name="daysToKeep">保留天數</param>
        public static void CleanupOldLogs(int daysToKeep = 30)
        {
            try
            {
                string logDir = Path.GetDirectoryName(LogFilePath);
                if (!Directory.Exists(logDir))
                {
                    return;
                }

                DateTime cutoffDate = DateTime.Now.AddDays(-daysToKeep);

                foreach (string file in Directory.GetFiles(logDir, "log_*.txt"))
                {
                    try
                    {
                        FileInfo fi = new FileInfo(file);
                        if (fi.CreationTime < cutoffDate)
                        {
                            fi.Delete();
                            Debug.WriteLine($"已刪除舊日誌文件: {fi.Name}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"刪除舊日誌時發生錯誤: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"清理舊日誌時發生錯誤: {ex.Message}");
            }
        }
    }
}