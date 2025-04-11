using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Xml.Linq;
using Xceed.Words.NET;
using System.Drawing;
using System.Windows.Forms;
using System.Threading.Tasks;
using POLICEPICTURE;

namespace POLICEPICTURE
{
    /// <summary>
    /// 進度報告委托 - 用於報告文檔生成進度
    /// </summary>
    /// <param name="progress">進度百分比 (0-100)</param>
    /// <param name="message">進度消息</param>
    public delegate void ProgressReportHandler(int progress, string message);

    /// <summary>
    /// 文檔相關處理類
    /// </summary>
    public class DocHelper
    {
        /// <summary>
        /// 生成文檔 - 支持進度報告
        /// </summary>
        /// <param name="templatePath">模板文件路徑</param>
        /// <param name="outputPath">輸出文件路徑</param>
        /// <param name="unit">單位</param>
        /// <param name="caseDesc">案由</param>
        /// <param name="time">時間</param>
        /// <param name="location">地點</param>
        /// <param name="photographer">攝影人</param>
        /// <param name="progressReport">進度報告回調</param>
        /// <returns>是否成功生成文檔</returns>
        public static async Task<bool> GenerateDocumentAsync(
            string templatePath,
            string outputPath,
            string unit,
            string caseDesc,
            string time,
            string location,
            string photographer,
            ProgressReportHandler progressReport = null)
        {
            // 使用Task.Run在後台線程執行耗時操作
            return await Task.Run(() =>
            {
                try
                {
                    // 報告進度 - 10%
                    progressReport?.Invoke(10, "準備生成文檔...");

                    // 驗證輸入參數
                    if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
                    {
                        throw new FileNotFoundException("找不到範本檔案", templatePath);
                    }

                    if (string.IsNullOrWhiteSpace(outputPath))
                    {
                        throw new ArgumentNullException(nameof(outputPath), "輸出路徑不能為空");
                    }

                    // 確保輸出目錄存在
                    string outputDir = Path.GetDirectoryName(outputPath);
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // 報告進度 - 20%
                    progressReport?.Invoke(20, "複製範本檔案...");

                    // 複製範本到輸出路徑
                    File.Copy(templatePath, outputPath, true);

                    // 報告進度 - 30%
                    progressReport?.Invoke(30, "處理文檔內容...");

                    // 開啟文件
                    using (DocX document = DocX.Load(outputPath))
                    {
                        // 使用字符串替換來填充文檔模板
                        document.ReplaceText(new Xceed.Document.NET.StringReplaceTextOptions
                        {
                            SearchValue = "%%UNIT%%",
                            NewValue = unit ?? string.Empty
                        });

                        document.ReplaceText(new Xceed.Document.NET.StringReplaceTextOptions
                        {
                            SearchValue = "%%CASE%%",
                            NewValue = caseDesc ?? string.Empty
                        });

                        document.ReplaceText(new Xceed.Document.NET.StringReplaceTextOptions
                        {
                            SearchValue = "%%TIME%%",
                            NewValue = time ?? string.Empty
                        });

                        document.ReplaceText(new Xceed.Document.NET.StringReplaceTextOptions
                        {
                            SearchValue = "%%ADDRESS%%",
                            NewValue = location ?? string.Empty
                        });

                        document.ReplaceText(new Xceed.Document.NET.StringReplaceTextOptions
                        {
                            SearchValue = "%%NAME%%",
                            NewValue = photographer ?? string.Empty
                        });

                        // 報告進度 - 50%
                        progressReport?.Invoke(50, "處理照片...");

                        // 獲取照片列表
                        var photos = PhotoManager.Instance.GetAllPhotos();

                        // 照片處理進度基準
                        int baseProgress = 50;
                        int progressPerPhoto = photos.Count > 0 ? 40 / photos.Count : 0;

                        if (photos.Count > 0)
                        {
                            // 使用新的照片處理方法
                            ProcessPhotosInTemplate(document, photos, (i) =>
                            {
                                progressReport?.Invoke(baseProgress + i * progressPerPhoto,
                                    $"處理照片 {i + 1}/{photos.Count}...");
                            });
                        }

                        // 報告進度 - 90%
                        progressReport?.Invoke(90, "儲存文件...");

                        // 儲存文件
                        document.Save();
                    }

                    // 報告進度 - 100%
                    progressReport?.Invoke(100, "文件生成完成");

                    return true;
                }
                catch (Exception ex)
                {
                    // 記錄詳細錯誤信息
                    Logger.Log($"生成文件時發生錯誤: {ex.Message}\n堆疊追蹤: {ex.StackTrace}", Logger.LogLevel.Error);

                    // 確保錯誤信息傳回給調用者
                    MessageBox.Show($"生成文件時發生錯誤: {ex.Message}\n\n詳細信息已記錄", "錯誤",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return false;
                }
            });
        }

        /// <summary>
        /// 在模板中處理照片，實現自動分頁和表格複製
        /// </summary>
        /// <param name="document">Word文檔對象</param>
        /// <param name="photos">照片列表</param>
        /// <param name="progressCallback">進度回調</param>
        private static void ProcessPhotosInTemplate(DocX document, IReadOnlyList<PhotoItem> photos, Action<int> progressCallback)
        {
            // 尋找所有含有 %%PICTURE%% 的段落及其位置
            var pictureTokens = FindAllPictureTokens(document);
            Logger.Log($"在模板中找到 {pictureTokens.Count} 個圖片標記", Logger.LogLevel.Debug);

            // 獲取模板中表格的信息
            var allTables = document.Tables;
            Logger.Log($"文檔中共有 {allTables.Count} 個表格", Logger.LogLevel.Debug);

            // 如果照片數量超過模板中的標記數量，需要複製表格和添加新頁
            if (photos.Count > pictureTokens.Count && pictureTokens.Count > 0)
            {
                // 尋找包含第一個 %%PICTURE%% 的表格
                var firstPictureTable = FindTableContainingToken(document, "%%PICTURE%%");

                if (firstPictureTable != null)
                {
                    Logger.Log("找到包含圖片標記的表格，準備進行複製", Logger.LogLevel.Info);

                    // 計算需要添加幾個額外的表格(向上取整)
                    int remainingPhotos = photos.Count - pictureTokens.Count;
                    int photosPerPage = pictureTokens.Count;
                    int additionalPagesNeeded = (int)Math.Ceiling((double)remainingPhotos / photosPerPage);

                    Logger.Log($"需要添加 {additionalPagesNeeded} 個額外頁面", Logger.LogLevel.Info);

                    // 逐個添加新頁面和表格
                    for (int i = 0; i < additionalPagesNeeded; i++)
                    {
                        // 插入分頁符
                        var pageBreakPara = document.InsertParagraph();
                        pageBreakPara.InsertPageBreakAfterSelf();
                        Logger.Log($"已插入第 {i + 1} 個分頁符", Logger.LogLevel.Debug);

                        // 複製表格
                        // 注意：DocX中表格複製功能有限，這裡使用一個幫助方法
                        var clonedTable = CloneTable(document, firstPictureTable);

                        // 添加 null 檢查
                        if (clonedTable != null)
                        {
                            document.InsertTable(clonedTable);
                            Logger.Log($"已複製並插入第 {i + 1} 個表格", Logger.LogLevel.Debug);

                            // 更新標記集合，加入新表格中的標記
                            var newTokens = FindAllPictureTokensInTable(clonedTable);
                            pictureTokens.AddRange(newTokens);
                            Logger.Log($"在新表格中找到 {newTokens.Count} 個圖片標記", Logger.LogLevel.Debug);
                        }
                        else
                        {
                            Logger.Log($"無法複製表格 {i + 1}，跳過此頁", Logger.LogLevel.Error);
                        }
                    }
                }
                else
                {
                    Logger.Log("無法找到包含圖片標記的表格，無法自動分頁", Logger.LogLevel.Warning);
                }
            }

            // 現在處理所有照片
            for (int i = 0; i < Math.Min(photos.Count, pictureTokens.Count); i++)
            {
                var photo = photos[i];
                var token = pictureTokens[i];

                try
                {
                    Logger.Log($"開始處理第 {i + 1} 張照片: {photo.FilePath}", Logger.LogLevel.Debug);

                    // 載入照片並設定合理大小
                    using (var img = Image.FromFile(photo.FilePath))
                    {
                        // 首先找到包含圖片的表格單元格以獲取可用空間
                        var para = token.Item1;
                        var cell = FindParentCell(para);

                        // 設定最大尺寸（根據單元格尺寸或默認值）
                        float maxWidth = 400; // 默認最大寬度
                        float maxHeight = 300; // 默認最大高度

                        // 如果能獲取到單元格，則使用單元格的寬度和高度作為限制
                        if (cell != null)
                        {
                            try
                            {
                                // 嘗試獲取單元格的寬度作為限制 (DocX可能不支持直接獲取尺寸)
                                // 預設單元格寬度為頁面寬度的80%
                                maxWidth = 400; // 保守估計的單元格寬度
                                maxHeight = 300; // 保守估計的單元格高度
                            }
                            catch (Exception ex)
                            {
                                Logger.Log($"獲取單元格尺寸時出錯，使用默認尺寸: {ex.Message}", Logger.LogLevel.Warning);
                            }
                        }

                        // 確保尺寸不會太小
                        maxWidth = Math.Max(maxWidth, 200);
                        maxHeight = Math.Max(maxHeight, 150);

                        // 計算縮放比例，確保圖片完全在表格內
                        float ratio = Math.Min(maxWidth / img.Width, maxHeight / img.Height);

                        // 如果圖片實際上比最大尺寸小，不需要放大
                        if (ratio > 1)
                            ratio = 1;

                        int newWidth = (int)(img.Width * ratio);
                        int newHeight = (int)(img.Height * ratio);

                        Logger.Log($"照片原始尺寸: {img.Width}x{img.Height}, 調整後: {newWidth}x{newHeight}", Logger.LogLevel.Debug);

                        // 添加圖片到文檔
                        var pic = document.AddImage(photo.FilePath);
                        var image = pic.CreatePicture(newWidth, newHeight);

                        // 在添加圖片前，添加照片描述(如果有)
                        if (!string.IsNullOrEmpty(photo.Description))
                        {
                            // 在段落上方插入描述
                            var descPara = para.InsertParagraphBeforeSelf(photo.Description);
                            descPara.Bold(); // 使描述加粗
                            Logger.Log($"已添加照片描述: {photo.Description}", Logger.LogLevel.Debug);
                        }

                        // 替換標記為空，然後添加圖片
                        string originalText = para.Text;
                        string newText = originalText.Replace("%%PICTURE%%", "");

                        // 如果替換後段落為空，直接在其位置插入圖片
                        if (string.IsNullOrWhiteSpace(newText))
                        {
                            para.AppendPicture(image);
                        }
                        else
                        {
                            // 替換文本並在適當位置插入圖片
                            int picPos = originalText.IndexOf("%%PICTURE%%");
                            if (picPos >= 0)
                            {
                                para.RemoveText(picPos, "%%PICTURE%%".Length);
                                para.InsertPicture(image, picPos);
                            }
                        }

                        Logger.Log($"已成功處理照片 {i + 1}/{photos.Count}", Logger.LogLevel.Debug);
                    }

                    // 報告進度
                    progressCallback(i);
                }
                catch (Exception ex)
                {
                    // 記錄錯誤並繼續處理其他照片
                    Logger.Log($"處理照片 {i + 1} 時發生錯誤: {ex.Message}", Logger.LogLevel.Error);

                    // 在文檔中添加錯誤訊息
                    var errorPara = document.InsertParagraph();
                    errorPara.Append($"照片 {i + 1} 處理錯誤: {ex.Message}").Bold();
                }
            }
        }

        /// <summary>
        /// 查找包含段落的單元格
        /// </summary>
        /// <param name="paragraph">段落對象</param>
        /// <returns>包含該段落的單元格，如果找不到則返回null</returns>
        private static Xceed.Document.NET.Cell FindParentCell(Xceed.Document.NET.Paragraph paragraph)
        {
            try
            {
                // 嘗試通過反射獲取單元格
                var cellProperty = paragraph.GetType().GetProperty("Cell", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                if (cellProperty != null)
                {
                    return cellProperty.GetValue(paragraph) as Xceed.Document.NET.Cell;
                }

                // 如果反射方式失敗，嘗試通過其他手段確定
                // 注意：這是一個粗略的估計，DocX可能不提供直接從段落到單元格的映射
                return null;
            }
            catch (Exception ex)
            {
                Logger.Log($"查找段落所在單元格時出錯: {ex.Message}", Logger.LogLevel.Warning);
                return null;
            }
        }        /// <summary>
                 /// 找到文檔中所有含有 %%PICTURE%% 的段落
                 /// </summary>
                 /// <param name="document">Word文檔對象</param>
                 /// <returns>包含段落對象和標記位置的列表</returns>
        private static List<Tuple<Xceed.Document.NET.Paragraph, int>> FindAllPictureTokens(DocX document)
        {
            var result = new List<Tuple<Xceed.Document.NET.Paragraph, int>>();

            // 遍歷所有段落
            foreach (var para in document.Paragraphs)
            {
                if (para.Text.Contains("%%PICTURE%%"))
                {
                    // 儲存段落和標記在段落中的位置
                    result.Add(new Tuple<Xceed.Document.NET.Paragraph, int>(para, para.Text.IndexOf("%%PICTURE%%")));
                    Logger.Log($"在段落中找到圖片標記，文本: {para.Text}", Logger.LogLevel.Debug);
                }
            }

            // 遍歷所有表格中的段落
            foreach (var table in document.Tables)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.Cells)
                    {
                        foreach (var para in cell.Paragraphs)
                        {
                            if (para.Text.Contains("%%PICTURE%%"))
                            {
                                result.Add(new Tuple<Xceed.Document.NET.Paragraph, int>(para, para.Text.IndexOf("%%PICTURE%%")));
                                Logger.Log($"在表格單元格中找到圖片標記，文本: {para.Text}", Logger.LogLevel.Debug);
                            }
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 找到包含指定標記的表格
        /// </summary>
        /// <param name="document">Word文檔對象</param>
        /// <param name="token">要查找的標記</param>
        /// <returns>包含標記的表格對象，如果未找到則返回null</returns>
        private static Xceed.Document.NET.Table FindTableContainingToken(DocX document, string token)
        {
            foreach (var table in document.Tables)
            {
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.Cells)
                    {
                        foreach (var para in cell.Paragraphs)
                        {
                            if (para.Text.Contains(token))
                            {
                                Logger.Log($"找到包含標記的表格，行數: {table.Rows.Count}", Logger.LogLevel.Debug);
                                return table;
                            }
                        }
                    }
                }
            }

            Logger.Log("未找到包含指定標記的表格", Logger.LogLevel.Warning);
            return null;
        }

        /// <summary>
        /// 找到表格中所有含有 %%PICTURE%% 的段落
        /// </summary>
        /// <param name="table">表格對象</param>
        /// <returns>包含段落對象和標記位置的列表</returns>
        private static List<Tuple<Xceed.Document.NET.Paragraph, int>> FindAllPictureTokensInTable(Xceed.Document.NET.Table table)
        {
            var result = new List<Tuple<Xceed.Document.NET.Paragraph, int>>();

            foreach (var row in table.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    foreach (var para in cell.Paragraphs)
                    {
                        if (para.Text.Contains("%%PICTURE%%"))
                        {
                            result.Add(new Tuple<Xceed.Document.NET.Paragraph, int>(para, para.Text.IndexOf("%%PICTURE%%")));
                            Logger.Log($"在表格段落中找到圖片標記，文本: {para.Text}", Logger.LogLevel.Debug);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 複製表格(DocX不直接支持表格克隆，需自行實現)
        /// </summary>
        /// <param name="document">Word文檔對象</param>
        /// <param name="sourceTable">源表格</param>
        /// <returns>克隆的表格</returns>
        private static Xceed.Document.NET.Table CloneTable(DocX document, Xceed.Document.NET.Table sourceTable)
        {
            try
            {
                // 創建新表格，行數和列數與源表格相同
                int rowCount = sourceTable.Rows.Count;
                int columnCount = sourceTable.ColumnCount;

                Logger.Log($"創建新表格，行數: {rowCount}, 列數: {columnCount}", Logger.LogLevel.Debug);

                // 安全檢查
                if (rowCount <= 0 || columnCount <= 0)
                {
                    Logger.Log("無法創建表格: 行數或列數無效", Logger.LogLevel.Error);
                    return null;
                }

                // 先檢查是否有合併單元格
                bool hasMergedCells = CheckForMergedCells(sourceTable);

                // 創建基本表格
                var newTable = document.AddTable(rowCount, columnCount);

                // 設定表格寬度與原表格相同
                try
                {
                    // 嘗試獲取並設置表格寬度
                    var widthProp = sourceTable.GetType().GetProperty("Width");
                    if (widthProp != null)
                    {
                        var sourceWidth = widthProp.GetValue(sourceTable);
                        var targetWidthProp = newTable.GetType().GetProperty("Width");
                        if (targetWidthProp != null && sourceWidth != null)
                        {
                            targetWidthProp.SetValue(newTable, sourceWidth);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"設定表格寬度時出錯: {ex.Message}", Logger.LogLevel.Warning);
                }

                // 如果有合併單元格，採用另一種方法進行複製
                if (hasMergedCells)
                {
                    Logger.Log("檢測到合併單元格，使用替代方法複製表格", Logger.LogLevel.Info);

                    // 使用替代方法 - 僅複製文本並保留基本結構
                    for (int i = 0; i < Math.Min(rowCount, sourceTable.Rows.Count); i++)
                    {
                        for (int j = 0; j < Math.Min(columnCount, sourceTable.Rows[i].Cells.Count); j++)
                        {
                            try
                            {
                                // 獲取源單元格和目標單元格
                                var sourceCell = sourceTable.Rows[i].Cells[j];
                                var targetCell = newTable.Rows[i].Cells[j];

                                // 複製單元格內容
                                if (sourceCell.Paragraphs.Count > 0)
                                {
                                    targetCell.InsertParagraph();
                                    foreach (var para in sourceCell.Paragraphs)
                                    {
                                        // 保留 %%PICTURE%% 標記和其他文本
                                        targetCell.InsertParagraph(para.Text);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.Log($"複製單元格 [{i},{j}] 時出錯: {ex.Message}", Logger.LogLevel.Warning);
                            }
                        }
                    }

                    return newTable;
                }

                // 標準複製方法(無合併單元格)
                for (int i = 0; i < rowCount; i++)
                {
                    // 確保源表格有足夠的行
                    if (i >= sourceTable.Rows.Count)
                    {
                        Logger.Log($"源表格行數不足，跳過行 {i}", Logger.LogLevel.Warning);
                        continue;
                    }

                    for (int j = 0; j < columnCount; j++)
                    {
                        // 確保源表格和目標表格有足夠的列
                        if (j >= sourceTable.Rows[i].Cells.Count || j >= newTable.Rows[i].Cells.Count)
                        {
                            Logger.Log($"單元格數不足，跳過單元格 [{i},{j}]", Logger.LogLevel.Warning);
                            continue;
                        }

                        // 獲取源單元格和目標單元格
                        var sourceCell = sourceTable.Rows[i].Cells[j];
                        var targetCell = newTable.Rows[i].Cells[j];

                        // 清空目標單元格的默認段落
                        if (targetCell.Paragraphs.Count > 0)
                        {
                            targetCell.Paragraphs[0].RemoveText(0, targetCell.Paragraphs[0].Text.Length);
                        }

                        // 複製單元格內容
                        foreach (var sourcePara in sourceCell.Paragraphs)
                        {
                            // 創建新段落並複製文本
                            var targetPara = targetCell.InsertParagraph();
                            targetPara.Append(sourcePara.Text);
                        }
                    }
                }

                // 嘗試複製表格設置(如邊框、合併單元格等)
                try
                {
                    // 設置表格的基本屬性(如果DocX支持)
                    newTable.Design = sourceTable.Design;

                    // 嘗試複製表格樣式
                    var designProp = sourceTable.GetType().GetProperty("Design");
                    if (designProp != null)
                    {
                        var design = designProp.GetValue(sourceTable);
                        designProp.SetValue(newTable, design);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"複製表格屬性時出錯: {ex.Message}", Logger.LogLevel.Warning);
                }

                return newTable;
            }
            catch (Exception ex)
            {
                Logger.Log($"複製表格時發生嚴重錯誤: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.Error);
                return null;
            }
        }

        /// <summary>
        /// 檢查表格是否有合併單元格
        /// </summary>
        /// <param name="table">要檢查的表格</param>
        /// <returns>是否有合併單元格</returns>
        private static bool CheckForMergedCells(Xceed.Document.NET.Table table)
        {
            try
            {
                // 檢查是否有跨行或跨列的單元格
                foreach (var row in table.Rows)
                {
                    foreach (var cell in row.Cells)
                    {
                        // 嘗試通過反射獲取單元格的GridSpan和RowSpan屬性
                        var gridSpanProp = cell.GetType().GetProperty("GridSpan");
                        var rowSpanProp = cell.GetType().GetProperty("RowSpan");

                        if (gridSpanProp != null)
                        {
                            var gridSpan = gridSpanProp.GetValue(cell);
                            if (gridSpan != null && Convert.ToInt32(gridSpan) > 1)
                                return true;
                        }

                        if (rowSpanProp != null)
                        {
                            var rowSpan = rowSpanProp.GetValue(cell);
                            if (rowSpan != null && Convert.ToInt32(rowSpan) > 1)
                                return true;
                        }

                        // 另一種檢查方法 - 檢查單元格計數異常
                        if (row.Cells.Count != table.ColumnCount)
                            return true;
                    }
                }

                // 如果表格的行數與列數乘積不等於單元格總數，可能有合併單元格
                int totalCells = 0;
                foreach (var row in table.Rows)
                {
                    totalCells += row.Cells.Count;
                }

                if (totalCells != table.Rows.Count * table.ColumnCount)
                    return true;

                return false;
            }
            catch (Exception ex)
            {
                Logger.Log($"檢查合併單元格時出錯: {ex.Message}", Logger.LogLevel.Warning);
                // 如果檢查出錯，假設有合併單元格，採取更安全的處理方式
                return true;
            }
        }
        /// <summary>
        /// 檢查模板文件是否存在
        /// </summary>
        /// <param name="templatePath">模板路徑</param>
        /// <param name="alternativePaths">備選路徑列表</param>
        /// <returns>有效的模板路徑或null</returns>
        public static string FindValidTemplatePath(string templatePath, params string[] alternativePaths)
        {
            // 首先檢查主要路徑
            if (!string.IsNullOrEmpty(templatePath) && File.Exists(templatePath))
            {
                return templatePath;
            }

            // 檢查備選路徑
            foreach (var path in alternativePaths)
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    return path;
                }
            }

            // 找不到有效路徑
            return null;
        }

        /// <summary>
        /// 安全地檢查文件是否可訪問
        /// </summary>
        /// <param name="filePath">文件路徑</param>
        /// <returns>文件是否可訪問</returns>
        public static bool IsFileAccessible(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;

            try
            {
                // 嘗試以讀取模式打開文件
                using (var fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"檢查文件訪問權限時發生錯誤: {ex.Message}", Logger.LogLevel.Warning);
                return false;
            }
        }

        /// <summary>
        /// 生成安全的文件名
        /// </summary>
        /// <param name="baseFileName">基本文件名</param>
        /// <returns>處理後的安全文件名</returns>
        public static string GenerateSafeFileName(string baseFileName)
        {
            if (string.IsNullOrEmpty(baseFileName))
                return $"Document_{DateTime.Now:yyyyMMdd_HHmmss}";

            // 移除不合法的字符
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string safeName = baseFileName;

            foreach (char c in invalidChars)
            {
                safeName = safeName.Replace(c, '_');
            }

            // 確保文件名不過長
            int maxLength = 100; // 適當的最大長度
            if (safeName.Length > maxLength)
            {
                safeName = safeName.Substring(0, maxLength);
            }

            return safeName;
        }
    }
}
