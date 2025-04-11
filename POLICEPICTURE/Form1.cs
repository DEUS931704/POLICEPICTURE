using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using POLICEPICTURE.Properties;

namespace POLICEPICTURE
{
    // 這是 Form1.cs 中與照片管理相關的部分，包含擴展的照片處理功能
    public partial class Form1 : Form
    {
        // 以下是照片管理相關的方法

        // 新增這些成員變數
        private const string APP_VERSION = "1.0.0"; // 應用程式版本常數
        private UserSettings settings; // 使用者設定
        private ErrorProvider errorProvider; // 錯誤提供者
        private ProgressForm progressForm; // 進度表單


        public Form1()
        {
            InitializeComponent();

            // 初始化變數
            errorProvider = new ErrorProvider(this);
            settings = UserSettings.Load();

            // 設置列表視圖
            SetupListView();

            // 訂閱照片管理器的事件
            PhotoManager.Instance.PhotosChanged += PhotoManager_PhotosChanged;

            // 從設定中填充表單
            if (!string.IsNullOrEmpty(settings.LastUnit))
                txtUnit.Text = settings.LastUnit;

            if (!string.IsNullOrEmpty(settings.LastPhotographer))
                txtPhotographer.Text = settings.LastPhotographer;

            // 更新最近檔案選單
            UpdateRecentFilesMenu();

            // 更新狀態列
            UpdateStatusBar("應用程式就緒");
        }

        // 新增以下更新最近檔案選單的方法
        private void UpdateRecentFilesMenu()
        {
            // 清空現有選單項
            menuRecentFiles.DropDownItems.Clear();

            // 如果沒有最近文件，顯示無項目訊息
            if (settings.RecentFiles.Count == 0)
            {
                var noItemsMenuItem = new ToolStripMenuItem("無項目");
                noItemsMenuItem.Enabled = false;
                menuRecentFiles.DropDownItems.Add(noItemsMenuItem);
                return;
            }

            // 添加每個最近文件
            foreach (var filePath in settings.RecentFiles)
            {
                if (File.Exists(filePath))
                {
                    var menuItem = new ToolStripMenuItem(Path.GetFileName(filePath));
                    menuItem.ToolTipText = filePath;
                    menuItem.Tag = filePath;
                    menuItem.Click += RecentFileMenuItem_Click;
                    menuRecentFiles.DropDownItems.Add(menuItem);
                }
            }

            // 添加分隔線和清除選單
            menuRecentFiles.DropDownItems.Add(new ToolStripSeparator());
            var clearMenuItem = new ToolStripMenuItem("清除列表");
            clearMenuItem.Click += ClearRecentFiles_Click;
            menuRecentFiles.DropDownItems.Add(clearMenuItem);
        }

        // 最近文件選單項點擊事件
        private void RecentFileMenuItem_Click(object sender, EventArgs e)
        {
            var menuItem = sender as ToolStripMenuItem;
            if (menuItem != null && menuItem.Tag is string filePath)
            {
                // 在這裡添加打開最近文件的代碼
                MessageBox.Show($"開啟文件: {filePath}");
            }
        }

        // 清除最近文件列表選單項點擊事件
        private void ClearRecentFiles_Click(object sender, EventArgs e)
        {
            settings.ClearRecentFiles();
            settings.Save();
            UpdateRecentFilesMenu();
        }

        /// <summary>
        /// 設置列表視圖
        /// </summary>
        private void SetupListView()
        {
            // 啟用整行選擇和網格線
            lvPhotos.FullRowSelect = true;
            lvPhotos.GridLines = true;

            // 啟用虛擬模式以提高性能
            lvPhotos.VirtualMode = true;
            lvPhotos.RetrieveVirtualItem += LvPhotos_RetrieveVirtualItem;

            // 啟用拖放功能，用於照片排序
            lvPhotos.AllowDrop = true;
            lvPhotos.ItemDrag += LvPhotos_ItemDrag;
            lvPhotos.DragEnter += LvPhotos_DragEnter;
            lvPhotos.DragDrop += LvPhotos_DragDrop;

            // 添加右鍵菜單
            lvPhotos.ContextMenuStrip = CreatePhotoContextMenu();
        }

        /// <summary>
        /// 創建照片列表的右鍵菜單
        /// </summary>
        private ContextMenuStrip CreatePhotoContextMenu()
        {
            ContextMenuStrip menu = new ContextMenuStrip();

            // 添加照片
            ToolStripMenuItem addItem = new ToolStripMenuItem("添加照片");
            addItem.Image = SystemIcons.Application.ToBitmap(); // 或使用自定義圖標
            addItem.Click += (s, e) => btnAddPhoto_Click(s, e);
            menu.Items.Add(addItem);

            // 移除照片
            ToolStripMenuItem removeItem = new ToolStripMenuItem("移除選中照片");
            removeItem.Click += (s, e) => btnRemovePhoto_Click(s, e);
            menu.Items.Add(removeItem);

            // 分隔線
            menu.Items.Add(new ToolStripSeparator());

            // 排序子菜單
            ToolStripMenuItem sortMenu = new ToolStripMenuItem("排序方式");

            ToolStripMenuItem sortByDateAsc = new ToolStripMenuItem("按日期升序");
            sortByDateAsc.Click += (s, e) => {
                PhotoManager.Instance.SortPhotosByDate(true);
                UpdatePhotoListView();
            };
            sortMenu.DropDownItems.Add(sortByDateAsc);

            ToolStripMenuItem sortByDateDesc = new ToolStripMenuItem("按日期降序");
            sortByDateDesc.Click += (s, e) => {
                PhotoManager.Instance.SortPhotosByDate(false);
                UpdatePhotoListView();
            };
            sortMenu.DropDownItems.Add(sortByDateDesc);

            ToolStripMenuItem sortByNameAsc = new ToolStripMenuItem("按文件名升序");
            sortByNameAsc.Click += (s, e) => {
                PhotoManager.Instance.SortPhotosByFileName(true);
                UpdatePhotoListView();
            };
            sortMenu.DropDownItems.Add(sortByNameAsc);

            ToolStripMenuItem sortByNameDesc = new ToolStripMenuItem("按文件名降序");
            sortByNameDesc.Click += (s, e) => {
                PhotoManager.Instance.SortPhotosByFileName(false);
                UpdatePhotoListView();
            };
            sortMenu.DropDownItems.Add(sortByNameDesc);

            menu.Items.Add(sortMenu);

            // 批量描述
            ToolStripMenuItem batchDescItem = new ToolStripMenuItem("批量設定描述");
            batchDescItem.Click += (s, e) => ShowBatchDescriptionDialog();
            menu.Items.Add(batchDescItem);

            // 顯示EXIF資訊
            ToolStripMenuItem showExifItem = new ToolStripMenuItem("顯示詳細資訊");
            showExifItem.Click += (s, e) => ShowPhotoDetailDialog();
            menu.Items.Add(showExifItem);

            // 匯出照片清單
            ToolStripMenuItem exportItem = new ToolStripMenuItem("匯出照片清單");
            exportItem.Click += (s, e) => ExportPhotoList();
            menu.Items.Add(exportItem);

            // 開啟菜單前檢查項目啟用狀態
            menu.Opening += (s, e) => {
                bool hasPhotos = PhotoManager.Instance.Count > 0;
                bool hasSelection = lvPhotos.SelectedIndices.Count > 0;

                removeItem.Enabled = hasSelection;
                sortMenu.Enabled = hasPhotos && PhotoManager.Instance.Count > 1;
                batchDescItem.Enabled = hasPhotos;
                showExifItem.Enabled = hasSelection;
                exportItem.Enabled = hasPhotos;
            };

            return menu;
        }

        /// <summary>
        /// 虛擬項目檢索事件
        /// </summary>
        private void LvPhotos_RetrieveVirtualItem(object sender, RetrieveVirtualItemEventArgs e)
        {
            var photos = PhotoManager.Instance.GetAllPhotos();
            if (e.ItemIndex < photos.Count)
            {
                var photo = photos[e.ItemIndex];
                var item = new ListViewItem((e.ItemIndex + 1).ToString()); // 編號列

                // 添加子項
                item.SubItems.Add(Path.GetFileName(photo.FilePath));
                item.SubItems.Add(photo.CaptureTime?.ToString("yyyy/MM/dd HH:mm:ss") ?? "未知");
                item.SubItems.Add($"{photo.Width}x{photo.Height}");
                item.SubItems.Add(photo.Description);

                e.Item = item;
            }
        }

        /// <summary>
        /// 照片列表項目拖動開始事件
        /// </summary>
        private void LvPhotos_ItemDrag(object sender, ItemDragEventArgs e)
        {
            // 只允許拖動列表項目
            if (e.Item is ListViewItem)
            {
                DoDragDrop(e.Item, DragDropEffects.Move);
                Logger.Log("開始拖動照片項目", Logger.LogLevel.Debug);
            }
        }

        /// <summary>
        /// 照片列表拖放進入事件
        /// </summary>
        private void LvPhotos_DragEnter(object sender, DragEventArgs e)
        {
            // 檢查拖放的數據類型
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                e.Effect = DragDropEffects.Move;
            }
            else if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                // 也接受文件拖放以添加照片
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                bool hasValidImages = files.Any(file => IsImageFile(file));

                if (hasValidImages)
                {
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        /// <summary>
        /// 照片列表拖放事件
        /// </summary>
        private void LvPhotos_DragDrop(object sender, DragEventArgs e)
        {
            // 處理項目重新排序
            if (e.Data.GetDataPresent(typeof(ListViewItem)))
            {
                // 獲取拖動項的索引
                int dragIndex = lvPhotos.SelectedIndices[0];

                // 獲取拖放位置的項
                Point targetPoint = lvPhotos.PointToClient(new Point(e.X, e.Y));
                ListViewItem targetItem = lvPhotos.GetItemAt(targetPoint.X, targetPoint.Y);
                int targetIndex = targetItem != null ? targetItem.Index : lvPhotos.Items.Count - 1;

                // 調用 PhotoManager 重新排序照片
                if (PhotoManager.Instance.ReorderPhotos(dragIndex, targetIndex))
                {
                    // 更新列表視圖
                    UpdatePhotoListView();

                    // 選中移動後的項目
                    lvPhotos.Items[targetIndex].Selected = true;
                    lvPhotos.Items[targetIndex].Focused = true;

                    Logger.Log($"已將照片從位置 {dragIndex} 移動到位置 {targetIndex}", Logger.LogLevel.Info);
                }
            }
            // 處理文件拖放以添加照片
            else if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);

                // 過濾出圖片文件
                var imageFiles = files.Where(file => IsImageFile(file)).ToList();

                if (imageFiles.Count > 0)
                {
                    // 添加照片
                    int addedCount = PhotoManager.Instance.AddPhotos(imageFiles);

                    if (addedCount > 0)
                    {
                        UpdatePhotoListView();
                        MessageBox.Show($"已成功添加 {addedCount} 張照片", "添加成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Logger.Log($"通過拖放添加了 {addedCount} 張照片", Logger.LogLevel.Info);
                    }
                    else
                    {
                        MessageBox.Show("無法添加照片，請檢查文件格式或照片數量限制", "添加失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        /// <summary>
        /// 判斷文件是否為圖片
        /// </summary>
        private bool IsImageFile(string filePath)
        {
            // 檢查文件擴展名
            string ext = Path.GetExtension(filePath).ToLower();
            string[] validExtensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif" };

            return validExtensions.Contains(ext);
        }

        /// <summary>
        /// 照片管理器變更事件處理
        /// </summary>
        private void PhotoManager_PhotosChanged(object sender, PhotoManager.PhotoCollectionChangedEventArgs e)
        {
            // 確保在UI線程中執行
            if (InvokeRequired)
            {
                Invoke(new Action(() => PhotoManager_PhotosChanged(sender, e)));
                return;
            }

            // 根據變更類型更新UI
            switch (e.Type)
            {
                case PhotoManager.PhotoCollectionChangedEventArgs.ChangeType.Add:
                case PhotoManager.PhotoCollectionChangedEventArgs.ChangeType.Remove:
                case PhotoManager.PhotoCollectionChangedEventArgs.ChangeType.Reorder:
                    // 更新列表視圖
                    UpdatePhotoListView();
                    break;
                case PhotoManager.PhotoCollectionChangedEventArgs.ChangeType.Clear:
                    // 清空列表視圖
                    lvPhotos.VirtualListSize = 0;
                    lvPhotos.Refresh();
                    // 清除預覽
                    if (pbPhotoPreview.Image != null)
                    {
                        pbPhotoPreview.Image.Dispose();
                        pbPhotoPreview.Image = null;
                    }
                    txtPhotoDescription.Text = string.Empty;
                    txtPhotoDescription.Enabled = false;
                    break;
            }
        }

        /// <summary>
        /// 更新照片列表視圖
        /// </summary>
        private void UpdatePhotoListView()
        {
            // 更新虛擬列表大小
            lvPhotos.VirtualListSize = PhotoManager.Instance.Count;
            lvPhotos.Refresh();

            // 更新狀態欄顯示照片數量
            UpdateStatusBar($"共有 {PhotoManager.Instance.Count} 張照片");
        }

        /// <summary>
        /// 照片描述變更事件
        /// </summary>
        private void TxtPhotoDescription_TextChanged(object sender, EventArgs e)
        {
            // 如果有選取的照片，更新其描述
            if (lvPhotos.SelectedIndices.Count > 0)
            {
                int index = lvPhotos.SelectedIndices[0];
                if (index >= 0 && index < PhotoManager.Instance.Count)
                {
                    PhotoManager.Instance.UpdatePhotoDescription(index, txtPhotoDescription.Text);
                }
            }
        }

        /// <summary>
        /// 照片列表選擇變更事件
        /// </summary>
        private void lvPhotos_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 當清單選擇改變時更新右側顯示和描述文字框
            if (lvPhotos.SelectedIndices.Count > 0)
            {
                int index = lvPhotos.SelectedIndices[0];
                if (index >= 0 && index < PhotoManager.Instance.Count)
                {
                    var photo = PhotoManager.Instance.GetPhoto(index);
                    if (photo != null)
                    {
                        // 啟用文字框和顯示照片描述
                        txtPhotoDescription.Enabled = true;
                        txtPhotoDescription.Text = photo.Description;

                        // 顯示照片預覽
                        try
                        {
                            if (File.Exists(photo.FilePath))
                            {
                                // 釋放之前的圖像
                                if (pbPhotoPreview.Image != null)
                                {
                                    pbPhotoPreview.Image.Dispose();
                                    pbPhotoPreview.Image = null;
                                }

                                // 使用新的GetThumbnail方法創建適合顯示的縮圖
                                int previewWidth = pbPhotoPreview.Width;
                                int previewHeight = pbPhotoPreview.Height;

                                // 獲取縮圖
                                Image thumbnail = photo.GetThumbnail(previewWidth, previewHeight);

                                if (thumbnail != null)
                                {
                                    pbPhotoPreview.Image = thumbnail;

                                    // 顯示尺寸信息
                                    string sizeInfo = $"尺寸: {photo.Width}x{photo.Height} 像素 | 大小: {photo.GetFormattedFileSize()}";
                                    lblPhotoInfo.Text = sizeInfo;
                                }
                                else
                                {
                                    // 如果縮圖創建失敗，使用傳統方法
                                    using (var stream = new FileStream(photo.FilePath, FileMode.Open, FileAccess.Read))
                                    {
                                        pbPhotoPreview.Image = Image.FromStream(stream);
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            pbPhotoPreview.Image = null;
                            Logger.Log($"載入照片預覽時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                            MessageBox.Show($"載入照片時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            else
            {
                // 如果沒有選擇，停用文字框和清除預覽
                txtPhotoDescription.Enabled = false;
                txtPhotoDescription.Text = string.Empty;
                lblPhotoInfo.Text = "";

                // 釋放之前的圖像
                if (pbPhotoPreview.Image != null)
                {
                    pbPhotoPreview.Image.Dispose();
                    pbPhotoPreview.Image = null;
                }
            }
        }

        /// <summary>
        /// 生成文件點擊事件
        /// </summary>
        private void btnGenerate_Click(object sender, EventArgs e)
        {
            // 驗證表單並顯示錯誤
            if (!ValidateFormData(true))
                return;

            // 使用改進的SaveDocument方法
            if (!SaveDocument())
            {
                Logger.Log("生成文件失敗", Logger.LogLevel.Warning);
            }
        }

        /// <summary>
        /// 新建文件點擊事件
        /// </summary>
        private void MenuFileNew_Click(object sender, EventArgs e)
        {
            // 詢問用戶是否要儲存當前工作
            if (IsFormDirty())
            {
                DialogResult result = MessageBox.Show("您有未儲存的工作，是否儲存？", "確認",
                    MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                if (result == DialogResult.Cancel)
                    return;

                if (result == DialogResult.Yes)
                {
                    // 如果用戶取消了儲存，則也取消新建
                    if (!SaveDocument())
                        return;
                }
            }

            // 清除所有欄位和照片
            txtUnit.Text = string.Empty;
            txtCase.Text = string.Empty;
            dtpDateTime.Value = DateTime.Now;
            txtLocation.Text = string.Empty;
            txtPhotographer.Text = settings.LastPhotographer; // 保留攝影人

            // 清除照片
            PhotoManager.Instance.ClearPhotos();

            UpdateStatusBar("已建立新文件");
            Logger.Log("已建立新文件");
        }

        /// <summary>
        /// 文件退出點擊事件
        /// </summary>
        private void MenuFileExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /// <summary>
        /// 設定模板點擊事件
        /// </summary>
        private void MenuSettingsTemplate_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Word 文件 (*.docx)|*.docx";
                dlg.Title = "選擇文件範本";

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    settings.TemplatePath = dlg.FileName;
                    settings.Save();
                    MessageBox.Show($"已設定範本檔案為: {dlg.FileName}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Logger.Log($"設定範本檔案: {dlg.FileName}");
                }
            }
        }

        /// <summary>
        /// 關於點擊事件
        /// </summary>
        private void MenuHelpAbout_Click(object sender, EventArgs e)
        {
            MessageBox.Show($"警察照片證據生成器 v{APP_VERSION}\n\n用於生成包含照片的證據文件。",
                "關於", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 預覽按鈕點擊事件
        /// </summary>
        private void btnPreview_Click(object sender, EventArgs e)
        {
            // 檢查表單數據
            if (!ValidateFormData())
            {
                MessageBox.Show("請先填寫必要的資訊再預覽文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            MessageBox.Show("預覽功能正在開發中...", "功能未完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Logger.Log("用戶嘗試使用未完成的預覽功能");
        }

        /// <summary>
        /// 檢查表單是否有未保存的內容
        /// </summary>
        private bool IsFormDirty()
        {
            // 如果有輸入內容或照片，則認為有未儲存的工作
            return !string.IsNullOrWhiteSpace(txtUnit.Text) ||
                   !string.IsNullOrWhiteSpace(txtCase.Text) ||
                   !string.IsNullOrWhiteSpace(txtLocation.Text) ||
                   PhotoManager.Instance.Count > 0;
        }

        /// <summary>
        /// 驗證表單數據
        /// </summary>
        /// <param name="showErrors">是否顯示錯誤訊息</param>
        /// <returns>是否驗證通過</returns>
        private bool ValidateFormData(bool showErrors = false)
        {
            bool isValid = true;
            string errorMessage = "";

            // 檢查必填字段
            if (string.IsNullOrWhiteSpace(txtUnit.Text))
            {
                isValid = false;
                errorMessage += "• 請填寫單位欄位\n";
                if (showErrors) errorProvider.SetError(txtUnit, "單位是必填欄位");
            }

            if (string.IsNullOrWhiteSpace(txtCase.Text))
            {
                isValid = false;
                errorMessage += "• 請填寫案由欄位\n";
                if (showErrors) errorProvider.SetError(txtCase, "案由是必填欄位");
            }

            // 檢查是否有照片
            if (PhotoManager.Instance.Count == 0)
            {
                if (showErrors)
                {
                    DialogResult result = MessageBox.Show("尚未添加任何照片，確定要繼續嗎？", "確認",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.No)
                    {
                        return false;
                    }
                }
            }

            // 顯示驗證錯誤
            if (!isValid && showErrors)
            {
                MessageBox.Show("請修正以下問題:\n" + errorMessage, "驗證錯誤", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return isValid;
        }

        // 修改 SaveDocument 方法，使用 WordInteropHelper 替代 DocHelper
        private bool SaveDocument()
        {
            try
            {
                // 收集表單數據
                string unit = txtUnit.Text.Trim();
                string caseDescription = txtCase.Text.Trim();
                string time = dtpDateTime.Text.Trim();
                string address = txtLocation.Text.Trim();
                string name = txtPhotographer.Text.Trim();

                // 儲存設定
                settings.LastUnit = unit;
                settings.LastPhotographer = name;
                settings.Save();

                // 生成新檔案名稱
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string defaultFileName = $"警察證據照片_{timestamp}.docx";
                string initialDir = string.IsNullOrEmpty(settings.LastSaveDirectory)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    : settings.LastSaveDirectory;

                // 使用SaveFileDialog讓用戶選擇儲存位置
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Word文件 (*.docx)|*.docx";
                    saveFileDialog.Title = "儲存證據文件";
                    saveFileDialog.FileName = defaultFileName;
                    saveFileDialog.InitialDirectory = initialDir;

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string saveFilePath = saveFileDialog.FileName;

                        // 更新最後儲存目錄
                        settings.LastSaveDirectory = Path.GetDirectoryName(saveFilePath);
                        settings.Save();

                        // 查找有效的模板路徑
                        string appPath = Application.StartupPath;
                        string projectPath = Directory.GetParent(appPath)?.Parent?.FullName ?? appPath;

                        string templatePath = DocHelper.FindValidTemplatePath(
                            settings.TemplatePath,
                            Path.Combine(appPath, "template.docx"),
                            Path.Combine(projectPath, "template.docx")
                        );

                        // 檢查範本是否存在
                        if (string.IsNullOrEmpty(templatePath))
                        {
                            Logger.Log("找不到範本檔案", Logger.LogLevel.Error);
                            MessageBox.Show($"找不到範本檔案！\n\n已嘗試尋找：\n1. {settings.TemplatePath}\n2. {Path.Combine(appPath, "template.docx")}\n3. {Path.Combine(projectPath, "template.docx")}\n\n請確認template.docx的位置或使用設定選單來指定範本。",
                                "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return false;
                        }

                        // 顯示進度表單
                        progressForm = new ProgressForm();
                        progressForm.Show(this);

                        // 獲取照片列表
                        var photos = PhotoManager.Instance.GetAllPhotos();

                        // 使用新的 WordInteropHelper 生成文檔（非同步）
                        Task.Run(async () =>
                        {
                            bool success = await WordInteropHelper.GenerateDocumentAsync(
                                templatePath,
                                saveFilePath,
                                unit,
                                caseDescription,
                                time,
                                address,
                                name,
                                photos,
                                ProgressReportCallback);

                            // 回到UI線程處理結果
                            this.Invoke(new Action(() =>
                            {
                                // 關閉進度表單
                                progressForm.Close();
                                progressForm = null;

                                if (success)
                                {
                                    // 添加到最近文件列表
                                    settings.AddRecentFile(saveFilePath);
                                    settings.Save();
                                    UpdateRecentFilesMenu();

                                    UpdateStatusBar("文件已成功生成");
                                    MessageBox.Show($"文件已成功生成！\n儲存路徑: {saveFilePath}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    Logger.Log($"文件已成功生成: {saveFilePath}");

                                    // 詢問是否開啟已儲存的文件
                                    if (MessageBox.Show("是否立即開啟文件？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                    {
                                        try
                                        {
                                            System.Diagnostics.Process.Start(saveFilePath);
                                        }
                                        catch (Exception ex)
                                        {
                                            Logger.Log($"無法開啟已生成的文件: {ex.Message}", Logger.LogLevel.Error);
                                            MessageBox.Show($"無法開啟文件: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                }
                                else
                                {
                                    UpdateStatusBar("生成文件失敗");
                                }
                            }));
                        });

                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"生成文件時發生錯誤: {ex.Message}\n{ex.StackTrace}", Logger.LogLevel.Error);
                MessageBox.Show($"生成文件時發生錯誤: {ex.Message}\n\n錯誤類型: {ex.GetType().FullName}",
                    "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                UpdateStatusBar("生成文件時發生錯誤");
            }

            return false;
        }

        /// <summary>
        /// 進度報告回調
        /// </summary>
        private void ProgressReportCallback(int progress, string message)
        {
            if (progressForm != null && !progressForm.IsDisposed)
            {
                this.Invoke(new Action(() =>
                {
                    progressForm.UpdateProgress(progress, message);
                }));
            }
        }

        /// <summary>
        /// 添加照片按鈕點擊事件
        /// </summary>
        private void btnAddPhoto_Click(object sender, EventArgs e)
        {
            if (PhotoManager.Instance.Count >= PhotoManager.MAX_PHOTOS)
            {
                MessageBox.Show($"最多只能添加{PhotoManager.MAX_PHOTOS}張照片！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "圖片檔案|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.tif|所有檔案|*.*";
                openFileDialog.Title = "選擇照片";
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    int addedCount = 0;
                    int failedCount = 0;

                    foreach (string file in openFileDialog.FileNames)
                    {
                        if (PhotoManager.Instance.Count >= PhotoManager.MAX_PHOTOS) break;

                        // 使用PhotoManager添加照片
                        if (PhotoManager.Instance.AddPhoto(file))
                        {
                            addedCount++;
                        }
                        else
                        {
                            failedCount++;
                        }
                    }

                    // 提供添加結果反饋
                    if (addedCount > 0)
                    {
                        string message = $"已添加 {addedCount} 張照片，目前共有 {PhotoManager.Instance.Count} 張照片";
                        if (failedCount > 0)
                        {
                            message += $"\n有 {failedCount} 張照片無法添加";
                        }
                        MessageBox.Show(message, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        Logger.Log(message);
                    }
                    else if (failedCount > 0)
                    {
                        MessageBox.Show($"所選的 {failedCount} 張照片均無法添加", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        /// <summary>
        /// 移除照片按鈕點擊事件
        /// </summary>
        private void btnRemovePhoto_Click(object sender, EventArgs e)
        {
            if (lvPhotos.SelectedIndices.Count > 0)
            {
                int index = lvPhotos.SelectedIndices[0];

                // 使用PhotoManager移除照片
                if (PhotoManager.Instance.RemovePhoto(index))
                {
                    MessageBox.Show("照片已移除", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Logger.Log($"已移除索引為 {index} 的照片");
                }
            }
            else
            {
                MessageBox.Show("請先選擇要移除的照片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// 顯示批量設定描述對話框
        /// </summary>
        private void ShowBatchDescriptionDialog()
        {
            if (PhotoManager.Instance.Count == 0)
            {
                MessageBox.Show("沒有照片可以設定描述", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 創建一個簡單的輸入對話框
            using (var form = new Form())
            {
                form.Text = "批量設定照片描述";
                form.Size = new Size(500, 240);
                form.StartPosition = FormStartPosition.CenterParent;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                var lblInfo = new Label
                {
                    Text = "您可以使用以下標記創建描述模板：",
                    Location = new Point(10, 15),
                    Size = new Size(460, 20)
                };

                var lblTags = new Label
                {
                    Text = "{INDEX} - 照片序號\n{DATE} - 拍攝日期\n{TIME} - 拍攝時間\n{FILENAME} - 檔案名稱",
                    Location = new Point(20, 35),
                    Size = new Size(460, 60)
                };

                var lblTemplate = new Label
                {
                    Text = "描述模板：",
                    Location = new Point(10, 95),
                    Size = new Size(100, 20)
                };

                var textBox = new TextBox
                {
                    Text = "照片 {INDEX} - 拍攝於 {DATE}",
                    Location = new Point(20, 115),
                    Size = new Size(450, 20),
                    Multiline = true,
                    Height = 40
                };

                var okButton = new Button
                {
                    Text = "確定",
                    DialogResult = DialogResult.OK,
                    Location = new Point(300, 170),
                    Width = 80
                };

                var cancelButton = new Button
                {
                    Text = "取消",
                    DialogResult = DialogResult.Cancel,
                    Location = new Point(390, 170),
                    Width = 80
                };

                form.Controls.Add(lblInfo);
                form.Controls.Add(lblTags);
                form.Controls.Add(lblTemplate);
                form.Controls.Add(textBox);
                form.Controls.Add(okButton);
                form.Controls.Add(cancelButton);

                form.AcceptButton = okButton;
                form.CancelButton = cancelButton;

                // 顯示對話框
                if (form.ShowDialog() == DialogResult.OK)
                {
                    // 批量更新描述
                    if (PhotoManager.Instance.BatchUpdateDescriptions(textBox.Text))
                    {
                        MessageBox.Show("已更新所有照片描述", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdatePhotoListView();
                    }
                    else
                    {
                        MessageBox.Show("更新照片描述失敗", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// 顯示照片詳細資訊對話框
        /// </summary>
        private void ShowPhotoDetailDialog()
        {
            if (lvPhotos.SelectedIndices.Count == 0)
            {
                MessageBox.Show("請先選擇一張照片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            int index = lvPhotos.SelectedIndices[0];
            var photo = PhotoManager.Instance.GetPhoto(index);

            if (photo == null)
                return;

            // 創建詳細資訊對話框
            using (var form = new Form())
            {
                form.Text = $"照片詳細資訊 - {Path.GetFileName(photo.FilePath)}";
                form.Size = new Size(600, 500);
                form.StartPosition = FormStartPosition.CenterParent;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                // 照片預覽
                var preview = new PictureBox
                {
                    Location = new Point(10, 10),
                    Size = new Size(200, 200),
                    SizeMode = PictureBoxSizeMode.Zoom,
                    BorderStyle = BorderStyle.FixedSingle
                };

                try
                {
                    preview.Image = photo.GetThumbnail(200, 200);
                }
                catch
                {
                    // 忽略預覽錯誤
                }

                // 基本資訊
                var lblBasicInfo = new Label
                {
                    Text = "基本資訊：",
                    Location = new Point(220, 10),
                    Size = new Size(100, 20),
                    Font = new Font(this.Font, FontStyle.Bold)
                };

                var txtBasicInfo = new TextBox
                {
                    Location = new Point(220, 30),
                    Size = new Size(350, 180),
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    Text = $"文件名稱: {Path.GetFileName(photo.FilePath)}\r\n" +
                           $"文件路徑: {photo.FilePath}\r\n" +
                           $"拍攝時間: {photo.GetFormattedTime()}\r\n" +
                           $"圖片尺寸: {photo.Width}x{photo.Height} 像素\r\n" +
                           $"檔案大小: {photo.GetFormattedFileSize()}\r\n" +
                           $"描述: {photo.Description}"
                };

                // EXIF資訊
                var lblExif = new Label
                {
                    Text = "EXIF資訊：",
                    Location = new Point(10, 220),
                    Size = new Size(100, 20),
                    Font = new Font(this.Font, FontStyle.Bold)
                };

                var txtExif = new TextBox
                {
                    Location = new Point(10, 240),
                    Size = new Size(560, 170),
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Both,
                    WordWrap = false
                };

                // 獲取EXIF數據
                var exifData = photo.GetExifData();
                StringBuilder sb = new StringBuilder();

                if (exifData.Count > 0)
                {
                    foreach (var pair in exifData)
                    {
                        sb.AppendLine($"{pair.Key}: {pair.Value}");
                    }
                    txtExif.Text = sb.ToString();
                }
                else
                {
                    txtExif.Text = "沒有發現EXIF資訊";
                }

                // 關閉按鈕
                var closeButton = new Button
                {
                    Text = "關閉",
                    DialogResult = DialogResult.Cancel,
                    Location = new Point(490, 420),
                    Size = new Size(80, 30)
                };

                form.Controls.Add(preview);
                form.Controls.Add(lblBasicInfo);
                form.Controls.Add(txtBasicInfo);
                form.Controls.Add(lblExif);
                form.Controls.Add(txtExif);
                form.Controls.Add(closeButton);

                form.AcceptButton = closeButton;
                form.CancelButton = closeButton;

                form.ShowDialog();

                // 釋放資源
                if (preview.Image != null)
                {
                    preview.Image.Dispose();
                }
            }
        }

        /// <summary>
        /// 更新狀態列
        /// </summary>
        private void UpdateStatusBar(string message)
        {
            // 記錄狀態到日誌
            Logger.Log($"狀態: {message}", Logger.LogLevel.Debug);
            statusLabel.Text = message;
        }

        /// <summary>
        /// 匯出照片清單
        /// </summary>
        private void ExportPhotoList()
        {
            if (PhotoManager.Instance.Count == 0)
            {
                MessageBox.Show("沒有照片可供匯出", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "文字檔案 (*.txt)|*.txt|所有檔案 (*.*)|*.*";
                dlg.Title = "匯出照片清單";
                dlg.FileName = $"照片清單_{DateTime.Now:yyyyMMdd}.txt";

                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    // 詢問是否包含EXIF資訊
                    DialogResult result = MessageBox.Show("是否包含詳細的EXIF資訊？", "確認",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    bool includeExif = (result == DialogResult.Yes);

                    // 匯出照片列表
                    if (PhotoManager.Instance.ExportPhotoList(dlg.FileName, includeExif))
                    {
                        MessageBox.Show($"照片清單已成功匯出到:\n{dlg.FileName}", "匯出成功",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // 詢問是否開啟檔案
                        result = MessageBox.Show("是否立即開啟匯出的檔案？", "確認",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            try
                            {
                                System.Diagnostics.Process.Start(dlg.FileName);
                            }
                            catch (Exception ex)
                            {
                                Logger.Log($"開啟檔案時發生錯誤: {ex.Message}", Logger.LogLevel.Error);
                                MessageBox.Show($"無法開啟檔案: {ex.Message}", "錯誤",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("匯出照片清單失敗", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }
}