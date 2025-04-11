using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace POLICEPICTURE
{
    /// <summary>
    /// 使用者設定類，負責存儲和載入使用者偏好設定
    /// </summary>
    [Serializable]
    public class UserSettings
    {
        /// <summary>
        /// 當前設定版本 - 用於未來版本兼容性
        /// </summary>
        public string Version { get; set; } = "1.0.1";

        /// <summary>
        /// 設定檔案路徑
        /// </summary>
        private static readonly string SettingsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "POLICEPICTURE",
            "settings.xml");

        /// <summary>
        /// 最近使用的單位名稱
        /// </summary>
        public string LastUnit { get; set; } = string.Empty;

        /// <summary>
        /// 最近使用的攝影人姓名
        /// </summary>
        public string LastPhotographer { get; set; } = string.Empty;

        /// <summary>
        /// 範本檔案路徑
        /// </summary>
        public string TemplatePath { get; set; } = string.Empty;

        /// <summary>
        /// 最後使用的儲存目錄
        /// </summary>
        public string LastSaveDirectory { get; set; } = string.Empty;

        /// <summary>
        /// 最近的文件列表
        /// </summary>
        public List<string> RecentFiles { get; set; } = new List<string>();

        /// <summary>
        /// 最近文件最大保存數量
        /// </summary>
        private const int MAX_RECENT_FILES = 10;

        /// <summary>
        /// 從檔案載入使用者設定，如果檔案不存在則建立新的設定
        /// </summary>
        /// <returns>使用者設定對象</returns>
        public static UserSettings Load()
        {
            try
            {
                // 確保目錄存在
                string directory = Path.GetDirectoryName(SettingsFilePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // 檢查設定檔是否存在
                if (!File.Exists(SettingsFilePath))
                {
                    return new UserSettings();
                }

                // 從XML檔案中載入設定
                XmlSerializer serializer = new XmlSerializer(typeof(UserSettings));

                using (FileStream fs = new FileStream(SettingsFilePath, FileMode.Open))
                {
                    UserSettings settings = (UserSettings)serializer.Deserialize(fs);

                    // 檢查版本兼容性
                    if (string.IsNullOrEmpty(settings.Version))
                    {
                        // 舊版本設定沒有版本號，添加當前版本
                        settings.Version = "1.0.1";
                    }

                    // 可以在這裡添加版本特定的遷移代碼

                    return settings;
                }
            }
            catch (Exception ex)
            {
                // 記錄錯誤
                Logger.Log($"載入設定時發生錯誤: {ex.Message}", Logger.LogLevel.Error);

                // 建立新的設定檔案
                UserSettings newSettings = new UserSettings();

                // 嘗試保存新的設定以修復錯誤
                try
                {
                    newSettings.Save();
                }
                catch
                {
                    // 忽略保存異常
                }

                return newSettings;
            }
        }

        /// <summary>
        /// 儲存使用者設定到檔案
        /// </summary>
        /// <returns>是否成功保存</returns>
        public bool Save()
        {
            try
            {
                // 確保目錄存在
                string directory = Path.GetDirectoryName(SettingsFilePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // 確保設定版本正確
                if (string.IsNullOrEmpty(this.Version))
                {
                    this.Version = "1.0.1";
                }

                // 序列化設定到XML檔案
                XmlSerializer serializer = new XmlSerializer(typeof(UserSettings));

                using (FileStream fs = new FileStream(SettingsFilePath, FileMode.Create))
                {
                    serializer.Serialize(fs, this);
                }

                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"儲存設定時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                System.Windows.Forms.MessageBox.Show($"儲存設定時發生錯誤: {ex.Message}", "錯誤",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// 添加最近文件
        /// </summary>
        /// <param name="filePath">文件路徑</param>
        public void AddRecentFile(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return;

            // 如果已存在，先移除
            RecentFiles.Remove(filePath);

            // 添加到列表開頭
            RecentFiles.Insert(0, filePath);

            // 如果超過最大數量，移除最後一個
            while (RecentFiles.Count > MAX_RECENT_FILES)
            {
                RecentFiles.RemoveAt(RecentFiles.Count - 1);
            }

            // 記錄
            Logger.Log($"添加到最近文件: {filePath}", Logger.LogLevel.Debug);
        }

        /// <summary>
        /// 移除最近文件
        /// </summary>
        /// <param name="filePath">文件路徑</param>
        public void RemoveRecentFile(string filePath)
        {
            if (RecentFiles.Remove(filePath))
            {
                Logger.Log($"從最近文件中移除: {filePath}", Logger.LogLevel.Debug);
            }
        }

        /// <summary>
        /// 清除最近文件列表
        /// </summary>
        public void ClearRecentFiles()
        {
            RecentFiles.Clear();
            Logger.Log("清除最近文件列表", Logger.LogLevel.Info);
        }
    }
}