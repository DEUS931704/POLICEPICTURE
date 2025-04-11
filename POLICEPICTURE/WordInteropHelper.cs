using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace POLICEPICTURE
{
    /// <summary>
    /// 使用 Office Interop 操作 Word 文檔的高級幫助類
    /// </summary>
    public class WordInteropHelper
    {
        /// <summary>
        /// 從模板生成文檔並處理照片
        /// </summary>
        public static async System.Threading.Tasks.Task<bool> GenerateDocumentAsync(
            string templatePath,
            string outputPath,
            string unit,
            string caseDesc,
            string time,
            string location,
            string photographer,
            IReadOnlyList<PhotoItem> photos,
            ProgressReportHandler progressReport = null)
        {
            // 使用 Task.Run 在後台線程執行耗時操作
            return await System.Threading.Tasks.Task.Run(() => 
            {
                // Word 應用實例、文檔和範圍變數
                Microsoft.Office.Interop.Word.Application wordApp = null;
                Microsoft.Office.Interop.Word.Document doc = null;
                object missing = System.Reflection.Missing.Value;

                try
                {
                    // 報告進度 - 10%
                    progressReport?.Invoke(10, "準備生成文檔...");

                    // 驗證模板路徑
                    if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
                    {
                        throw new FileNotFoundException("找不到範本檔案", templatePath);
                    }

                    // 確保輸出目錄存在
                    string outputDir = Path.GetDirectoryName(outputPath);
                    if (!Directory.Exists(outputDir))
                    {
                        Directory.CreateDirectory(outputDir);
                    }

                    // 報告進度 - 20%
                    progressReport?.Invoke(20, "初始化 Word 應用...");

                    // 創建 Word 應用實例
                    wordApp = new Microsoft.Office.Interop.Word.Application();
                    wordApp.Visible = false; // 隱藏 Word 應用

                    // 報告進度 - 25%
                    progressReport?.Invoke(25, "加載範本文檔...");

                    // 打開模板文件
                    doc = wordApp.Documents.Open(
                        templatePath,
                        ReadOnly: false,
                        Visible: false);

                    // 報告進度 - 30%
                    progressReport?.Invoke(30, "填充文檔內容...");

                    // 替換文檔中的佔位符
                    ReplaceTextInDocument(doc, "%%UNIT%%", unit ?? string.Empty);
                    ReplaceTextInDocument(doc, "%%CASE%%", caseDesc ?? string.Empty);
                    ReplaceTextInDocument(doc, "%%TIME%%", time ?? string.Empty);
                    ReplaceTextInDocument(doc, "%%ADDRESS%%", location ?? string.Empty);
                    ReplaceTextInDocument(doc, "%%NAME%%", photographer ?? string.Empty);

                    // 報告進度 - 40%
                    progressReport?.Invoke(40, "查找照片標記...");

                    // 查找文檔中的所有照片標記
                    List<Range> pictureRanges = FindAllPictureMarkers(doc);
                    Logger.Log($"在文檔中找到 {pictureRanges.Count} 個圖片標記", Logger.LogLevel.Info);

                    // 處理照片數量超過標記數量的情況
                    if (photos.Count > pictureRanges.Count && pictureRanges.Count > 0)
                    {
                        // 報告進度 - 45%
                        progressReport?.Invoke(45, "處理額外的照片頁...");

                        // 計算需要添加的額外頁數
                        int neededPages = (int)Math.Ceiling((double)(photos.Count - pictureRanges.Count) / 2);

                        // 找到第一個包含圖片標記的表格
                        Table firstTable = FindTableContainingMarker(doc, "%%PICTURE%%");

                        if (firstTable != null)
                        {
                            // 獲取最後一頁的位置
                            int lastPageNumber = doc.ComputeStatistics(WdStatistic.wdStatisticPages);

                            // 將游標移動到文檔末尾
                            doc.Content.Collapse(WdCollapseDirection.wdCollapseEnd);

                            // 添加換頁符和表格副本
                            for (int i = 0; i < neededPages; i++)
                            {
                                // 添加換頁符
                                doc.Content.InsertBreak(WdBreakType.wdPageBreak);

                                // 複製表格
                                firstTable.Range.Copy();
                                doc.Content.Paste();

                                // 尋找新添加的表格中的圖片標記
                                List<Range> newMarkers = FindAllPictureMarkers(doc, lastPageNumber + i + 1);
                                pictureRanges.AddRange(newMarkers);

                                Logger.Log($"已添加第 {i + 1} 個額外頁面，找到 {newMarkers.Count} 個新圖片標記", Logger.LogLevel.Info);
                            }
                        }
                        else
                        {
                            Logger.Log("無法找到包含圖片標記的表格", Logger.LogLevel.Warning);
                        }
                    }

                    // 報告進度 - 50%
                    progressReport?.Invoke(50, "處理照片...");

                    // 處理照片
                    for (int i = 0; i < Math.Min(photos.Count, pictureRanges.Count); i++)
                    {
                        // 計算進度百分比
                        int photoProgress = 50 + (40 * (i + 1)) / Math.Min(photos.Count, pictureRanges.Count);
                        progressReport?.Invoke(photoProgress, $"處理照片 {i + 1}/{Math.Min(photos.Count, pictureRanges.Count)}...");

                        // 處理照片
                        ProcessPhoto(doc, pictureRanges[i], photos[i]);
                    }

                    // 報告進度 - 90%
                    progressReport?.Invoke(90, "保存文檔...");

                    // 保存文檔
                    doc.SaveAs2(outputPath);

                    // 報告進度 - 100%
                    progressReport?.Invoke(100, "文件生成完成");

                    return true;
                }
                catch (Exception ex)
                {
                    Logger.Log($"使用 Word Interop 生成文件時發生錯誤: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.Error);
                    MessageBox.Show($"生成文件時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                finally
                {
                    // 清理資源
                    if (doc != null)
                    {
                        doc.Close(SaveChanges: false);
                        Marshal.ReleaseComObject(doc);
                    }

                    if (wordApp != null)
                    {
                        wordApp.Quit();
                        Marshal.ReleaseComObject(wordApp);
                    }

                    // 強制垃圾回收，確保 COM 對象被釋放
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            });
        }

        /// <summary>
        /// 在文檔中替換指定文本
        /// </summary>
        private static void ReplaceTextInDocument(Document doc, string findText, string replaceText)
        {
            var findObject = doc.Application.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = findText;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = replaceText;

            object replaceAll = WdReplace.wdReplaceAll;
            findObject.Execute(
                FindText: findText,
                MatchCase: false,
                MatchWholeWord: false,
                MatchWildcards: false,
                MatchSoundsLike: false,
                MatchAllWordForms: false,
                Forward: true,
                Wrap: WdFindWrap.wdFindContinue,
                Format: false,
                ReplaceWith: replaceText,
                Replace: replaceAll);
        }

        /// <summary>
        /// 查找文檔中所有含有 %%PICTURE%% 標記的範圍
        /// </summary>
        private static List<Range> FindAllPictureMarkers(Document doc, int pageNumber = 0)
        {
            List<Range> results = new List<Range>();

            // 選擇要搜尋的範圍
            Range searchRange;
            if (pageNumber > 0)
            {
                // 如果指定了頁碼，則只搜索該頁
                searchRange = doc.GoTo(
                    What: WdGoToItem.wdGoToPage,
                    Which: WdGoToDirection.wdGoToAbsolute,
                    Count: pageNumber).Duplicate;

                // 將範圍擴展到整頁
                if (pageNumber < doc.ComputeStatistics(WdStatistic.wdStatisticPages))
                {
                    Range nextPage = doc.GoTo(
                        What: WdGoToItem.wdGoToPage,
                        Which: WdGoToDirection.wdGoToAbsolute,
                        Count: pageNumber + 1);
                    searchRange.End = nextPage.Start - 1;
                }
            }
            else
            {
                // 否則搜索整個文檔
                searchRange = doc.Content;
            }

            // 搜索標記
            Range currentRange = searchRange.Duplicate;
            currentRange.Find.ClearFormatting();
            currentRange.Find.Text = "%%PICTURE%%";

            // 循環查找所有匹配項
            while (currentRange.Find.Execute())
            {
                // 將找到的範圍添加到結果列表
                results.Add(currentRange.Duplicate);

                // 移動到下一個搜索位置
                currentRange.Start = currentRange.End;
                currentRange.End = searchRange.End;
            }

            return results;
        }

        /// <summary>
        /// 查找包含指定標記的表格
        /// </summary>
        private static Table FindTableContainingMarker(Document doc, string marker)
        {
            foreach (Table table in doc.Tables)
            {
                if (table.Range.Text.Contains(marker))
                {
                    return table;
                }
            }
            return null;
        }

        /// <summary>
        /// 處理單張照片
        /// </summary>
        private static void ProcessPhoto(Document doc, Range markerRange, PhotoItem photo)
        {
            try
            {
                // 確保照片文件存在
                if (!File.Exists(photo.FilePath))
                {
                    Logger.Log($"照片文件不存在: {photo.FilePath}", Logger.LogLevel.Error);
                    return;
                }

                // 添加照片描述
                if (!string.IsNullOrEmpty(photo.Description))
                {
                    // 在標記前插入描述
                    Range descriptionRange = markerRange.Duplicate;
                    descriptionRange.Collapse(WdCollapseDirection.wdCollapseStart);
                    descriptionRange.InsertBefore(photo.Description + "\r\n");
                    descriptionRange.Bold = 1; // 加粗描述
                }

                // 獲取表格單元格的大小以限制圖片尺寸
                Cell cell = GetContainingCell(markerRange);
                float maxWidth = 400; // 默認最大寬度
                float maxHeight = 300; // 默認最大高度

                if (cell != null)
                {
                    try
                    {
                        // 嘗試獲取單元格寬度和高度
                        float cellWidth = (float)cell.Width;
                        if (cellWidth > 0)
                        {
                            maxWidth = cellWidth * 0.8f; // 使用單元格寬度的 80%
                        }

                        // 單元格高度在 Word 中可能難以獲取
                        // 使用保守估計的高度
                        maxHeight = maxWidth * 0.75f;
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"獲取單元格尺寸時出錯: {ex.Message}", Logger.LogLevel.Warning);
                    }
                }

                // 確保尺寸不會太小
                maxWidth = Math.Max(maxWidth, 150);
                maxHeight = Math.Max(maxHeight, 120);

                // 計算等比例縮放後的尺寸
                using (System.Drawing.Image img = System.Drawing.Image.FromFile(photo.FilePath))
                {
                    float ratio = Math.Min(maxWidth / img.Width, maxHeight / img.Height);
                    ratio = Math.Min(ratio, 1.0f); // 不放大圖片
                    float width = img.Width * ratio;
                    float height = img.Height * ratio;

                    // 記錄圖片處理信息
                    Logger.Log($"處理照片: {Path.GetFileName(photo.FilePath)}, 原始尺寸: {img.Width}x{img.Height}, 縮放比例: {ratio}, 目標尺寸: {width}x{height}", Logger.LogLevel.Info);

                    // 插入圖片
                    markerRange.Text = ""; // 清除標記
                    InlineShape shape = markerRange.InlineShapes.AddPicture(
                        FileName: photo.FilePath,
                        LinkToFile: false,
                        SaveWithDocument: true);

                    // 設置圖片尺寸
                    shape.Width = (float)width;
                    shape.Height = (float)height;
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"處理照片時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                // 在錯誤位置插入錯誤信息
                markerRange.Text = $"[照片錯誤: {ex.Message}]";
                markerRange.Bold = 1;
                markerRange.Font.Color = WdColor.wdColorRed;
            }
        }

        /// <summary>
        /// 獲取包含指定範圍的表格單元格
        /// </summary>
        private static Cell GetContainingCell(Range range)
        {
            try
            {
                // 嘗試獲取範圍所在的表格和單元格
                if (range.Tables.Count > 0)
                {
                    Table table = range.Tables[1];
                    foreach (Row row in table.Rows)
                    {
                        foreach (Cell cell in row.Cells)
                        {
                            if (range.Start >= cell.Range.Start && range.End <= cell.Range.End)
                            {
                                return cell;
                            }
                        }
                    }
                }

                // 如果上面的方法失敗，嘗試通過其他方式獲取
                range.Select();
                return range.Application.Selection.Cells.Count > 0
                    ? range.Application.Selection.Cells[1]
                    : null;
            }
            catch (Exception ex)
            {
                Logger.Log($"獲取包含單元格時發生錯誤: {ex.Message}", Logger.LogLevel.Warning);
                return null;
            }
        }
    }
}