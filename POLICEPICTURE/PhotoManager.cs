using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using POLICEPICTURE;

namespace POLICEPICTURE
{
    /// <summary>
    /// 照片管理類 - 使用單例模式取代靜態類
    /// </summary>
    public class PhotoManager
    {
        // 單例實例
        private static PhotoManager _instance;

        // 照片列表
        private List<PhotoItem> _photos = new List<PhotoItem>();

        // 最大照片數量限制
        public const int MAX_PHOTOS = 20; // 修改為更大的值以支持多張照片和分頁

        // 照片變更事件 - 添加事件通知機制
        public event EventHandler<PhotoCollectionChangedEventArgs> PhotosChanged;

        /// <summary>
        /// 照片列表變更事件參數
        /// </summary>
        public class PhotoCollectionChangedEventArgs : EventArgs
        {
            public enum ChangeType { Add, Remove, Clear, Reorder }
            public ChangeType Type { get; set; }
            public int Index { get; set; }
            public PhotoItem Photo { get; set; }
        }

        /// <summary>
        /// 私有構造函數（單例模式）
        /// </summary>
        private PhotoManager() { }

        /// <summary>
        /// 獲取PhotoManager實例
        /// </summary>
        public static PhotoManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new PhotoManager();
                }
                return _instance;
            }
        }

        /// <summary>
        /// 添加照片
        /// </summary>
        /// <param name="path">照片路徑</param>
        /// <param name="description">照片描述</param>
        /// <returns>是否添加成功</returns>
        public bool AddPhoto(string path, string description = "")
        {
            // 檢查是否超過最大照片數量
            if (_photos.Count >= MAX_PHOTOS)
            {
                Logger.Log($"照片數量已達上限 ({MAX_PHOTOS} 張)", Logger.LogLevel.Warning);
                return false;
            }

            try
            {
                // 驗證是否為有效圖片
                using (var img = Image.FromFile(path))
                {
                    // 創建照片對象
                    PhotoItem newPhoto = PhotoItem.FromFile(path);

                    if (newPhoto == null)
                    {
                        Logger.Log($"無法從文件創建照片對象: {path}", Logger.LogLevel.Error);
                        return false;
                    }

                    // 如果提供了描述，則設置描述
                    if (!string.IsNullOrEmpty(description))
                    {
                        newPhoto.Description = description;
                    }

                    _photos.Add(newPhoto);

                    // 觸發照片變更事件
                    OnPhotosChanged(PhotoCollectionChangedEventArgs.ChangeType.Add, _photos.Count - 1, newPhoto);
                    Logger.Log($"已添加照片: {path}", Logger.LogLevel.Info);
                    return true;
                }
            }
            catch (Exception ex)
            {
                // 記錄異常並返回失敗
                Logger.Log($"無法添加照片 {path}: {ex.Message}", Logger.LogLevel.Error);
                return false;
            }
        }

        /// <summary>
        /// 批量添加照片
        /// </summary>
        /// <param name="paths">照片路徑列表</param>
        /// <returns>成功添加的照片數量</returns>
        public int AddPhotos(IEnumerable<string> paths)
        {
            int successCount = 0;

            foreach (string path in paths)
            {
                // 檢查是否超過最大照片數量
                if (_photos.Count >= MAX_PHOTOS)
                {
                    Logger.Log($"照片數量已達上限 ({MAX_PHOTOS} 張)，停止添加", Logger.LogLevel.Warning);
                    break;
                }

                if (AddPhoto(path))
                {
                    successCount++;
                }
            }

            return successCount;
        }

        /// <summary>
        /// 獲取所有照片
        /// </summary>
        /// <returns>照片列表（只讀）</returns>
        public IReadOnlyList<PhotoItem> GetAllPhotos()
        {
            return _photos.AsReadOnly();
        }

        /// <summary>
        /// 獲取指定索引的照片
        /// </summary>
        /// <param name="index">索引</param>
        /// <returns>照片項目或null</returns>
        public PhotoItem GetPhoto(int index)
        {
            if (index >= 0 && index < _photos.Count)
            {
                return _photos[index];
            }
            return null;
        }

        /// <summary>
        /// 更新照片描述
        /// </summary>
        /// <param name="index">索引</param>
        /// <param name="description">新描述</param>
        /// <returns>是否更新成功</returns>
        public bool UpdatePhotoDescription(int index, string description)
        {
            if (index >= 0 && index < _photos.Count)
            {
                // 記錄舊描述，用於記錄
                string oldDescription = _photos[index].Description;

                // 設置新描述
                _photos[index].Description = description;

                // 觸發照片變更事件
                OnPhotosChanged(PhotoCollectionChangedEventArgs.ChangeType.Add, index, _photos[index]);

                Logger.Log($"已更新照片描述 (索引: {index}): '{oldDescription}' -> '{description}'", Logger.LogLevel.Debug);
                return true;
            }

            Logger.Log($"更新照片描述失敗，索引無效: {index}", Logger.LogLevel.Warning);
            return false;
        }

        /// <summary>
        /// 移除指定索引的照片
        /// </summary>
        /// <param name="index">索引</param>
        /// <returns>是否移除成功</returns>
        public bool RemovePhoto(int index)
        {
            if (index >= 0 && index < _photos.Count)
            {
                PhotoItem removedPhoto = _photos[index];
                _photos.RemoveAt(index);

                // 觸發照片變更事件
                OnPhotosChanged(PhotoCollectionChangedEventArgs.ChangeType.Remove, index, removedPhoto);

                Logger.Log($"已移除照片 (索引: {index}, 路徑: {removedPhoto.FilePath})", Logger.LogLevel.Info);
                return true;
            }

            Logger.Log($"移除照片失敗，索引無效: {index}", Logger.LogLevel.Warning);
            return false;
        }

        /// <summary>
        /// 清除所有照片
        /// </summary>
        public void ClearPhotos()
        {
            int count = _photos.Count;
            _photos.Clear();

            // 觸發照片變更事件
            OnPhotosChanged(PhotoCollectionChangedEventArgs.ChangeType.Clear, -1, null);

            Logger.Log($"已清除所有照片 (共 {count} 張)", Logger.LogLevel.Info);
        }

        /// <summary>
        /// 重新排序照片
        /// </summary>
        /// <param name="oldIndex">原始索引</param>
        /// <param name="newIndex">新索引</param>
        /// <returns>是否成功重新排序</returns>
        public bool ReorderPhotos(int oldIndex, int newIndex)
        {
            // 驗證索引
            if (oldIndex < 0 || oldIndex >= _photos.Count ||
                newIndex < 0 || newIndex >= _photos.Count)
            {
                Logger.Log($"重新排序照片失敗，索引無效: oldIndex={oldIndex}, newIndex={newIndex}", Logger.LogLevel.Warning);
                return false;
            }

            // 不需要移動的情况
            if (oldIndex == newIndex)
                return true;

            // 獲取要移動的照片
            var photoToMove = _photos[oldIndex];

            // 從舊位置移除
            _photos.RemoveAt(oldIndex);

            // 插入到新位置
            _photos.Insert(newIndex, photoToMove);

            // 觸發照片變更事件
            OnPhotosChanged(PhotoCollectionChangedEventArgs.ChangeType.Reorder, -1, null);

            Logger.Log($"已重新排序照片: 將索引 {oldIndex} 移動到 {newIndex}", Logger.LogLevel.Debug);
            return true;
        }

        /// <summary>
        /// 按拍攝日期排序照片
        /// </summary>
        /// <param name="ascending">是否升序排序</param>
        public void SortPhotosByDate(bool ascending = true)
        {
            if (_photos.Count <= 1)
                return;

            if (ascending)
            {
                _photos.Sort((a, b) =>
                    a.CaptureTime.GetValueOrDefault().CompareTo(b.CaptureTime.GetValueOrDefault()));
                Logger.Log("已按拍攝日期升序排序照片", Logger.LogLevel.Info);
            }
            else
            {
                _photos.Sort((a, b) =>
                    b.CaptureTime.GetValueOrDefault().CompareTo(a.CaptureTime.GetValueOrDefault()));
                Logger.Log("已按拍攝日期降序排序照片", Logger.LogLevel.Info);
            }

            // 觸發照片變更事件
            OnPhotosChanged(PhotoCollectionChangedEventArgs.ChangeType.Reorder, -1, null);
        }

        /// <summary>
        /// 按文件名排序照片
        /// </summary>
        /// <param name="ascending">是否升序排序</param>
        public void SortPhotosByFileName(bool ascending = true)
        {
            if (_photos.Count <= 1)
                return;

            if (ascending)
            {
                _photos.Sort((a, b) =>
                    Path.GetFileName(a.FilePath).CompareTo(Path.GetFileName(b.FilePath)));
                Logger.Log("已按文件名升序排序照片", Logger.LogLevel.Info);
            }
            else
            {
                _photos.Sort((a, b) =>
                    Path.GetFileName(b.FilePath).CompareTo(Path.GetFileName(a.FilePath)));
                Logger.Log("已按文件名降序排序照片", Logger.LogLevel.Info);
            }

            // 觸發照片變更事件
            OnPhotosChanged(PhotoCollectionChangedEventArgs.ChangeType.Reorder, -1, null);
        }

        /// <summary>
        /// 批量更新照片描述
        /// </summary>
        /// <param name="descriptionTemplate">描述模板</param>
        /// <returns>是否成功更新所有照片描述</returns>
        public bool BatchUpdateDescriptions(string descriptionTemplate)
        {
            if (string.IsNullOrEmpty(descriptionTemplate))
            {
                Logger.Log("批量更新描述失敗，模板為空", Logger.LogLevel.Warning);
                return false;
            }

            if (_photos.Count == 0)
            {
                Logger.Log("批量更新描述失敗，沒有照片", Logger.LogLevel.Warning);
                return false;
            }

            try
            {
                StringBuilder log = new StringBuilder();
                log.AppendLine($"批量更新照片描述，模板: '{descriptionTemplate}'");

                for (int i = 0; i < _photos.Count; i++)
                {
                    string oldDescription = _photos[i].Description;

                    // 替換模板中的特殊標記
                    string newDescription = descriptionTemplate
                        .Replace("{INDEX}", (i + 1).ToString())
                        .Replace("{DATE}", _photos[i].CaptureTime?.ToString("yyyy/MM/dd") ?? "")
                        .Replace("{TIME}", _photos[i].CaptureTime?.ToString("HH:mm:ss") ?? "")
                        .Replace("{FILENAME}", Path.GetFileName(_photos[i].FilePath));

                    // 更新描述
                    _photos[i].Description = newDescription;

                    log.AppendLine($"  照片 {i + 1}: '{oldDescription}' -> '{newDescription}'");
                }

                // 觸發照片變更事件
                OnPhotosChanged(PhotoCollectionChangedEventArgs.ChangeType.Add, -1, null);

                Logger.Log(log.ToString(), Logger.LogLevel.Debug);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"批量更新照片描述時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                return false;
            }
        }

        /// <summary>
        /// 匯出照片列表到文字檔
        /// </summary>
        /// <param name="filePath">輸出文件路徑</param>
        /// <param name="includeExif">是否包含EXIF資訊</param>
        /// <returns>是否成功匯出</returns>
        public bool ExportPhotoList(string filePath, bool includeExif = false)
        {
            if (_photos.Count == 0)
            {
                Logger.Log("匯出照片列表失敗，沒有照片", Logger.LogLevel.Warning);
                return false;
            }

            try
            {
                using (var writer = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    writer.WriteLine("警察照片證據清單");
                    writer.WriteLine($"生成時間: {DateTime.Now:yyyy/MM/dd HH:mm:ss}");
                    writer.WriteLine($"照片數量: {_photos.Count}");
                    writer.WriteLine(new string('-', 50));
                    writer.WriteLine();

                    for (int i = 0; i < _photos.Count; i++)
                    {
                        var photo = _photos[i];

                        writer.WriteLine($"照片 {i + 1}:");
                        writer.WriteLine($"文件名: {Path.GetFileName(photo.FilePath)}");
                        writer.WriteLine($"拍攝時間: {photo.GetFormattedTime()}");
                        writer.WriteLine($"尺寸: {photo.Width}x{photo.Height} 像素");
                        writer.WriteLine($"檔案大小: {photo.GetFormattedFileSize()}");

                        if (!string.IsNullOrEmpty(photo.Description))
                        {
                            writer.WriteLine($"描述: {photo.Description}");
                        }

                        // 如果需要匯出EXIF資訊
                        if (includeExif)
                        {
                            var exifData = photo.GetExifData();
                            if (exifData.Count > 0)
                            {
                                writer.WriteLine("EXIF信息:");
                                foreach (var item in exifData)
                                {
                                    writer.WriteLine($"  {item.Key}: {item.Value}");
                                }
                            }
                        }

                        writer.WriteLine(new string('-', 30));
                    }
                }

                Logger.Log($"已成功匯出照片列表到: {filePath}", Logger.LogLevel.Info);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"匯出照片列表時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                return false;
            }
        }

        /// <summary>
        /// 照片數量
        /// </summary>
        public int Count => _photos.Count;

        /// <summary>
        /// 觸發照片變更事件
        /// </summary>
        /// <param name="type">變更類型</param>
        /// <param name="index">變更索引</param>
        /// <param name="photo">相關照片</param>
        private void OnPhotosChanged(PhotoCollectionChangedEventArgs.ChangeType type, int index, PhotoItem photo)
        {
            PhotosChanged?.Invoke(this, new PhotoCollectionChangedEventArgs
            {
                Type = type,
                Index = index,
                Photo = photo
            });
        }
    }
}