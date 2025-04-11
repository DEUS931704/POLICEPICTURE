using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using System.Text;

namespace POLICEPICTURE
{
    /// <summary>
    /// 照片項目類 - 統一使用PhotoItem取代兩個不同的照片類
    /// </summary>
    public class PhotoItem
    {
        /// <summary>
        /// 照片檔案路徑
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

        /// <summary>
        /// 照片說明文字
        /// </summary>
        public string Description { get; set; } = string.Empty;

        /// <summary>
        /// 拍攝時間
        /// </summary>
        public DateTime? CaptureTime { get; set; } = null;

        /// <summary>
        /// 照片寬度 (延遲載入)
        /// </summary>
        private int? _width = null;
        public int Width
        {
            get
            {
                if (!_width.HasValue)
                {
                    LoadImageDimensions();
                }
                return _width ?? 0;
            }
        }

        /// <summary>
        /// 照片高度 (延遲載入)
        /// </summary>
        private int? _height = null;
        public int Height
        {
            get
            {
                if (!_height.HasValue)
                {
                    LoadImageDimensions();
                }
                return _height ?? 0;
            }
        }

        /// <summary>
        /// 照片EXIF數據 (延遲載入)
        /// </summary>
        private Dictionary<string, string> _exifData = null;

        /// <summary>
        /// 照片檔案大小 (延遲載入)
        /// </summary>
        private long? _fileSize = null;
        public long FileSize
        {
            get
            {
                if (!_fileSize.HasValue && !string.IsNullOrEmpty(FilePath) && File.Exists(FilePath))
                {
                    try
                    {
                        _fileSize = new FileInfo(FilePath).Length;
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"獲取文件大小時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                        _fileSize = 0;
                    }
                }
                return _fileSize ?? 0;
            }
        }

        /// <summary>
        /// 取得格式化後的拍攝時間字串
        /// </summary>
        /// <returns>格式化的時間字串</returns>
        public string GetFormattedTime()
        {
            if (CaptureTime.HasValue)
            {
                return CaptureTime.Value.ToString("yyyy年MM月dd日 HH:mm");
            }
            return string.Empty;
        }

        /// <summary>
        /// 取得格式化後的檔案大小字串
        /// </summary>
        /// <returns>格式化的檔案大小字串</returns>
        public string GetFormattedFileSize()
        {
            long size = FileSize;
            string[] units = { "B", "KB", "MB", "GB" };
            int unitIndex = 0;
            double adjustedSize = size;

            while (adjustedSize >= 1024 && unitIndex < units.Length - 1)
            {
                adjustedSize /= 1024;
                unitIndex++;
            }

            return $"{adjustedSize:0.##} {units[unitIndex]}";
        }

        /// <summary>
        /// 從文件路徑創建照片數據
        /// </summary>
        /// <param name="filePath">文件路徑</param>
        /// <returns>照片數據對象</returns>
        public static PhotoItem FromFile(string filePath)
        {
            try
            {
                // 檢查文件是否存在
                if (!File.Exists(filePath))
                {
                    Logger.Log($"無法找到文件: {filePath}", Logger.LogLevel.Error);
                    return null;
                }

                var photo = new PhotoItem
                {
                    FilePath = filePath,
                    Description = Path.GetFileName(filePath)
                };

                // 嘗試從照片EXIF信息獲取拍攝時間，如果失敗則使用文件創建時間
                try
                {
                    var exifData = photo.GetExifData();
                    if (exifData.ContainsKey("拍攝時間") && !string.IsNullOrEmpty(exifData["拍攝時間"]))
                    {
                        // 嘗試解析EXIF中的時間格式
                        if (DateTime.TryParse(exifData["拍攝時間"], out DateTime captureTime))
                        {
                            photo.CaptureTime = captureTime;
                        }
                        else
                        {
                            photo.CaptureTime = File.GetCreationTime(filePath);
                        }
                    }
                    else
                    {
                        photo.CaptureTime = File.GetCreationTime(filePath);
                    }
                }
                catch
                {
                    // 如果讀取EXIF失敗，使用文件創建時間
                    photo.CaptureTime = File.GetCreationTime(filePath);
                }

                return photo;
            }
            catch (Exception ex)
            {
                Logger.Log($"創建照片對象時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                return null;
            }
        }

        /// <summary>
        /// 獲取照片的縮圖
        /// </summary>
        /// <param name="width">目標寬度</param>
        /// <param name="height">目標高度</param>
        /// <returns>縮圖對象</returns>
        public Image GetThumbnail(int width, int height)
        {
            try
            {
                // 檢查文件是否存在
                if (!File.Exists(FilePath))
                {
                    Logger.Log($"無法找到照片文件: {FilePath}", Logger.LogLevel.Error);
                    return null;
                }

                using (var original = Image.FromFile(FilePath))
                {
                    // 計算縮放比例
                    float ratio = Math.Min((float)width / original.Width, (float)height / original.Height);
                    int newWidth = (int)(original.Width * ratio);
                    int newHeight = (int)(original.Height * ratio);

                    // 創建縮圖
                    var thumbnail = new Bitmap(newWidth, newHeight);
                    using (var graphics = Graphics.FromImage(thumbnail))
                    {
                        graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        graphics.DrawImage(original, 0, 0, newWidth, newHeight);
                    }

                    return thumbnail;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"創建縮圖時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                return null;
            }
        }

        /// <summary>
        /// 加載圖像尺寸信息
        /// </summary>
        private void LoadImageDimensions()
        {
            try
            {
                if (!string.IsNullOrEmpty(FilePath) && File.Exists(FilePath))
                {
                    // 使用無需鎖定文件的方式讀取尺寸
                    using (var stream = new FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        using (var image = Image.FromStream(stream))
                        {
                            _width = image.Width;
                            _height = image.Height;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"讀取圖片尺寸時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                _width = 0;
                _height = 0;
            }
        }

        /// <summary>
        /// 獲取照片的EXIF數據
        /// </summary>
        /// <returns>包含EXIF信息的字典</returns>
        public Dictionary<string, string> GetExifData()
        {
            // 如果已經讀取過，直接返回緩存的數據
            if (_exifData != null)
                return _exifData;

            _exifData = new Dictionary<string, string>();

            try
            {
                if (string.IsNullOrEmpty(FilePath) || !File.Exists(FilePath))
                    return _exifData;

                using (var stream = new FileStream(FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    using (var image = Image.FromStream(stream))
                    {
                        // 檢查圖片是否有屬性項
                        if (image.PropertyIdList.Length > 0)
                        {
                            // 讀取所有屬性項
                            foreach (var propId in image.PropertyIdList)
                            {
                                try
                                {
                                    var propItem = image.GetPropertyItem(propId);
                                    string value = "";

                                    // 根據數據類型解析屬性值
                                    switch (propItem.Type)
                                    {
                                        case 1: // Byte
                                            if (propItem.Value.Length > 0)
                                                value = propItem.Value[0].ToString();
                                            break;
                                        case 2: // ASCII
                                            value = Encoding.ASCII.GetString(propItem.Value).TrimEnd('\0');
                                            break;
                                        case 3: // Short
                                            if (propItem.Value.Length >= 2)
                                                value = BitConverter.ToUInt16(propItem.Value, 0).ToString();
                                            break;
                                        case 4: // Long
                                            if (propItem.Value.Length >= 4)
                                                value = BitConverter.ToUInt32(propItem.Value, 0).ToString();
                                            break;
                                        case 5: // Rational
                                            if (propItem.Value.Length >= 8)
                                            {
                                                uint numerator = BitConverter.ToUInt32(propItem.Value, 0);
                                                uint denominator = BitConverter.ToUInt32(propItem.Value, 4);

                                                if (denominator != 0)
                                                    value = $"{numerator}/{denominator}";
                                                else
                                                    value = numerator.ToString();
                                            }
                                            break;
                                        default:
                                            value = $"[未知類型: {propItem.Type}]";
                                            break;
                                    }

                                    // 將屬性ID轉換為有意義的名稱
                                    string propName = GetExifPropertyName(propId);
                                    _exifData[propName] = value;
                                }
                                catch
                                {
                                    // 忽略無法解析的屬性
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"讀取EXIF數據時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
            }

            return _exifData;
        }

        /// <summary>
        /// 將EXIF屬性ID轉換為有意義的名稱
        /// </summary>
        /// <param name="propId">屬性ID</param>
        /// <returns>屬性名稱</returns>
        private string GetExifPropertyName(int propId)
        {
            switch (propId)
            {
                case 0x010F: return "製造商";
                case 0x0110: return "型號";
                case 0x0132: return "拍攝時間";
                case 0x010E: return "圖像描述";
                case 0x013B: return "攝影師";
                case 0x8298: return "版權";
                case 0x8769: return "EXIF數據";
                case 0x9003: return "原始時間";
                case 0x9004: return "數字化時間";
                case 0x829A: return "曝光時間";
                case 0x829D: return "光圈值";
                case 0x8822: return "曝光模式";
                case 0x8827: return "ISO速度";
                case 0x9207: return "測光模式";
                case 0x9209: return "閃光燈";
                case 0x920A: return "焦距";
                case 0xA001: return "色彩空間";
                case 0xA002: return "圖像寬度";
                case 0xA003: return "圖像高度";
                case 0xA405: return "焦距35mm";
                case 0xA406: return "場景類型";
                case 0xA408: return "對比度";
                case 0xA409: return "飽和度";
                case 0xA40A: return "銳度";
                // 添加更多EXIF屬性
                default: return $"屬性-{propId:X4}";
            }
        }

        /// <summary>
        /// 獲取照片的格式化信息摘要
        /// </summary>
        /// <returns>格式化的信息摘要</returns>
        public string GetInfoSummary()
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine($"文件名: {Path.GetFileName(FilePath)}");
            sb.AppendLine($"拍攝時間: {GetFormattedTime()}");
            sb.AppendLine($"尺寸: {Width}x{Height} 像素");
            sb.AppendLine($"檔案大小: {GetFormattedFileSize()}");

            // 添加描述(如果有)
            if (!string.IsNullOrEmpty(Description))
            {
                sb.AppendLine($"描述: {Description}");
            }

            // 添加一些重要的EXIF信息(如果有)
            var exifData = GetExifData();
            if (exifData.Count > 0)
            {
                sb.AppendLine("EXIF信息:");

                string[] importantProps = { "製造商", "型號", "攝影師", "曝光時間", "光圈值", "ISO速度", "焦距" };
                foreach (var prop in importantProps)
                {
                    if (exifData.ContainsKey(prop) && !string.IsNullOrEmpty(exifData[prop]))
                    {
                        sb.AppendLine($"  {prop}: {exifData[prop]}");
                    }
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// 裁剪照片
        /// </summary>
        /// <param name="outputPath">輸出路徑</param>
        /// <param name="x">起始X坐標</param>
        /// <param name="y">起始Y坐標</param>
        /// <param name="width">裁剪寬度</param>
        /// <param name="height">裁剪高度</param>
        /// <returns>是否成功裁剪</returns>
        public bool CropImage(string outputPath, int x, int y, int width, int height)
        {
            try
            {
                if (!File.Exists(FilePath))
                {
                    Logger.Log($"找不到原始照片文件: {FilePath}", Logger.LogLevel.Error);
                    return false;
                }

                // 檢查裁剪參數
                if (x < 0 || y < 0 || width <= 0 || height <= 0)
                {
                    Logger.Log("裁剪參數無效", Logger.LogLevel.Error);
                    return false;
                }

                using (var original = Image.FromFile(FilePath))
                {
                    // 檢查裁剪範圍是否有效
                    if (x + width > original.Width || y + height > original.Height)
                    {
                        Logger.Log("裁剪範圍超出原始圖片範圍", Logger.LogLevel.Error);
                        return false;
                    }

                    // 創建裁剪矩形
                    Rectangle cropRect = new Rectangle(x, y, width, height);

                    // 創建新的裁剪後的圖片
                    using (Bitmap target = new Bitmap(width, height))
                    {
                        using (Graphics g = Graphics.FromImage(target))
                        {
                            // 設置高質量繪圖
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                            g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

                            // 繪製裁剪後的圖片
                            g.DrawImage(original, new Rectangle(0, 0, width, height), cropRect, GraphicsUnit.Pixel);
                        }

                        // 保存裁剪後的圖片
                        target.Save(outputPath, ImageFormat.Jpeg);
                    }
                }

                Logger.Log($"成功裁剪照片並保存到: {outputPath}", Logger.LogLevel.Info);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Log($"裁剪照片時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                return false;
            }
        }
    }
}