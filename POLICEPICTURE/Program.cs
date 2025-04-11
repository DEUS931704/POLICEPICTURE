using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Globalization;

namespace POLICEPICTURE
{
    internal static class Program
    {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            try
            {
                // 設置默認區域和文化
                Thread.CurrentThread.CurrentUICulture = new CultureInfo("zh-TW");
                Thread.CurrentThread.CurrentCulture = new CultureInfo("zh-TW");

                // 初始化日誌系統
                Logger.Initialize();
                Logger.Log("應用程序啟動");

                // 啟用視覺樣式
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                // 設置全局異常處理
                Application.ThreadException += Application_ThreadException;
                AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

                // 運行主窗體
                Application.Run(new Form1());
            }
            catch (Exception ex)
            {
                // 記錄致命錯誤
                try
                {
                    Logger.Log($"應用程序發生嚴重錯誤: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.Error);
                }
                catch
                {
                    // 如果日誌記錄失敗，顯示消息框
                }

                MessageBox.Show($"應用程序發生嚴重錯誤，即將關閉:\n{ex.Message}", "嚴重錯誤",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 處理UI線程異常
        /// </summary>
        private static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            Logger.Log($"UI線程異常: {e.Exception.Message}\n{e.Exception.StackTrace}", Logger.LogLevel.Error);
            MessageBox.Show($"應用程序發生錯誤:\n{e.Exception.Message}", "錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        /// <summary>
        /// 處理未捕獲的異常
        /// </summary>
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var ex = e.ExceptionObject as Exception;
            Logger.Log($"未處理的異常: {ex?.Message ?? "未知錯誤"}\n{ex?.StackTrace ?? "無堆疊信息"}", Logger.LogLevel.Error);

            MessageBox.Show($"應用程序發生嚴重錯誤，即將關閉:\n{ex?.Message ?? "未知錯誤"}", "嚴重錯誤",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}