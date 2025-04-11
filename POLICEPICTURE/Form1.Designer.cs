namespace POLICEPICTURE
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.Label lblPhotoInfo;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.txtUnit = new System.Windows.Forms.TextBox();
            this.lblUnit = new System.Windows.Forms.Label();
            this.lblCase = new System.Windows.Forms.Label();
            this.txtCase = new System.Windows.Forms.TextBox();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.menuFile = new System.Windows.Forms.ToolStripMenuItem();
            this.menuFileNew = new System.Windows.Forms.ToolStripMenuItem();
            this.menuFileOpen = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.menuRecentFiles = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.menuFileExit = new System.Windows.Forms.ToolStripMenuItem();
            this.menuSettings = new System.Windows.Forms.ToolStripMenuItem();
            this.menuSettingsTemplate = new System.Windows.Forms.ToolStripMenuItem();
            this.menuHelp = new System.Windows.Forms.ToolStripMenuItem();
            this.menuHelpAbout = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.statusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabPageBasicInfo = new System.Windows.Forms.TabPage();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblDateTime = new System.Windows.Forms.Label();
            this.dtpDateTime = new System.Windows.Forms.DateTimePicker();
            this.lblLocation = new System.Windows.Forms.Label();
            this.txtLocation = new System.Windows.Forms.TextBox();
            this.lblPhotographer = new System.Windows.Forms.Label();
            this.txtPhotographer = new System.Windows.Forms.TextBox();
            this.tabPagePhotos = new System.Windows.Forms.TabPage();
            this.splitContainer = new System.Windows.Forms.SplitContainer();
            this.lvPhotos = new System.Windows.Forms.ListView();
            this.colFilename = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colDate = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtPhotoDescription = new System.Windows.Forms.TextBox();
            this.lblPhotoDescription = new System.Windows.Forms.Label();
            this.btnRemovePhoto = new System.Windows.Forms.Button();
            this.btnAddPhoto = new System.Windows.Forms.Button();
            this.pbPhotoPreview = new System.Windows.Forms.PictureBox();
            this.lblPreview = new System.Windows.Forms.Label();
            this.btnPreview = new System.Windows.Forms.Button();
            this.menuStrip.SuspendLayout();
            this.statusStrip.SuspendLayout();
            this.tabControl.SuspendLayout();
            this.tabPageBasicInfo.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPagePhotos.SuspendLayout();
            this.lblPhotoInfo = new System.Windows.Forms.Label();
            this.lblPhotoInfo.AutoSize = true;
            this.lblPhotoInfo.Location = new System.Drawing.Point(3, 370);
            this.lblPhotoInfo.Name = "lblPhotoInfo";
            this.lblPhotoInfo.Size = new System.Drawing.Size(65, 12);
            this.lblPhotoInfo.TabIndex = 2;
            this.lblPhotoInfo.Text = "";

            this.splitContainer.Panel2.Controls.Add(this.lblPhotoInfo);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).BeginInit();
            this.splitContainer.Panel1.SuspendLayout();
            this.splitContainer.Panel2.SuspendLayout();
            this.splitContainer.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbPhotoPreview)).BeginInit();
            this.SuspendLayout();
            // 
            // txtUnit
            // 
            this.txtUnit.Location = new System.Drawing.Point(96, 40);
            this.txtUnit.Name = "txtUnit";
            this.txtUnit.Size = new System.Drawing.Size(285, 22);
            this.txtUnit.TabIndex = 0;
            // 
            // lblUnit
            // 
            this.lblUnit.AutoSize = true;
            this.lblUnit.Location = new System.Drawing.Point(29, 45);
            this.lblUnit.Name = "lblUnit";
            this.lblUnit.Size = new System.Drawing.Size(29, 12);
            this.lblUnit.TabIndex = 1;
            this.lblUnit.Text = "單位";
            // 
            // lblCase
            // 
            this.lblCase.AutoSize = true;
            this.lblCase.Location = new System.Drawing.Point(29, 87);
            this.lblCase.Name = "lblCase";
            this.lblCase.Size = new System.Drawing.Size(29, 12);
            this.lblCase.TabIndex = 2;
            this.lblCase.Text = "案由";
            // 
            // txtCase
            // 
            this.txtCase.Location = new System.Drawing.Point(96, 82);
            this.txtCase.Name = "txtCase";
            this.txtCase.Size = new System.Drawing.Size(285, 22);
            this.txtCase.TabIndex = 3;
            // 
            // btnGenerate
            // 
            this.btnGenerate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnGenerate.Location = new System.Drawing.Point(374, 470);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(120, 38);
            this.btnGenerate.TabIndex = 4;
            this.btnGenerate.Text = "生成文件";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // menuStrip
            // 
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuFile,
            this.menuSettings,
            this.menuHelp});
            this.menuStrip.Location = new System.Drawing.Point(0, 0);
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Size = new System.Drawing.Size(534, 24);
            this.menuStrip.TabIndex = 5;
            this.menuStrip.Text = "menuStrip1";
            // 
            // menuFile
            // 
            this.menuFile.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuFileNew,
            this.menuFileOpen,
            this.toolStripSeparator1,
            this.menuRecentFiles,
            this.toolStripSeparator2,
            this.menuFileExit});
            this.menuFile.Name = "menuFile";
            this.menuFile.Size = new System.Drawing.Size(43, 20);
            this.menuFile.Text = "檔案";
            // 
            // menuFileNew
            // 
            this.menuFileNew.Name = "menuFileNew";
            this.menuFileNew.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N)));
            this.menuFileNew.Size = new System.Drawing.Size(180, 22);
            this.menuFileNew.Text = "新增";
            this.menuFileNew.Click += new System.EventHandler(this.MenuFileNew_Click);
            // 
            // menuFileOpen
            // 
            this.menuFileOpen.Name = "menuFileOpen";
            this.menuFileOpen.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.menuFileOpen.Size = new System.Drawing.Size(180, 22);
            this.menuFileOpen.Text = "開啟舊檔";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(177, 6);
            // 
            // menuRecentFiles
            // 
            this.menuRecentFiles.Name = "menuRecentFiles";
            this.menuRecentFiles.Size = new System.Drawing.Size(180, 22);
            this.menuRecentFiles.Text = "最近的檔案";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(177, 6);
            // 
            // menuFileExit
            // 
            this.menuFileExit.Name = "menuFileExit";
            this.menuFileExit.Size = new System.Drawing.Size(180, 22);
            this.menuFileExit.Text = "結束";
            this.menuFileExit.Click += new System.EventHandler(this.MenuFileExit_Click);
            // 
            // menuSettings
            // 
            this.menuSettings.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuSettingsTemplate});
            this.menuSettings.Name = "menuSettings";
            this.menuSettings.Size = new System.Drawing.Size(43, 20);
            this.menuSettings.Text = "設定";
            // 
            // menuSettingsTemplate
            // 
            this.menuSettingsTemplate.Name = "menuSettingsTemplate";
            this.menuSettingsTemplate.Size = new System.Drawing.Size(180, 22);
            this.menuSettingsTemplate.Text = "範本設定";
            this.menuSettingsTemplate.Click += new System.EventHandler(this.MenuSettingsTemplate_Click);
            // 
            // menuHelp
            // 
            this.menuHelp.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.menuHelpAbout});
            this.menuHelp.Name = "menuHelp";
            this.menuHelp.Size = new System.Drawing.Size(43, 20);
            this.menuHelp.Text = "說明";
            // 
            // menuHelpAbout
            // 
            this.menuHelpAbout.Name = "menuHelpAbout";
            this.menuHelpAbout.Size = new System.Drawing.Size(180, 22);
            this.menuHelpAbout.Text = "關於";
            this.menuHelpAbout.Click += new System.EventHandler(this.MenuHelpAbout_Click);
            // 
            // statusStrip
            // 
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.statusLabel});
            this.statusStrip.Location = new System.Drawing.Point(0, 517);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(534, 22);
            this.statusStrip.TabIndex = 6;
            this.statusStrip.Text = "statusStrip1";
            // 
            // statusLabel
            // 
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(31, 17);
            this.statusLabel.Text = "就緒";
            // 
            // tabControl
            // 
            this.tabControl.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl.Controls.Add(this.tabPageBasicInfo);
            this.tabControl.Controls.Add(this.tabPagePhotos);
            this.tabControl.Location = new System.Drawing.Point(12, 35);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(510, 429);
            this.tabControl.TabIndex = 7;
            // 
            // tabPageBasicInfo
            // 
            this.tabPageBasicInfo.Controls.Add(this.groupBox1);
            this.tabPageBasicInfo.Controls.Add(this.lblUnit);
            this.tabPageBasicInfo.Controls.Add(this.txtUnit);
            this.tabPageBasicInfo.Controls.Add(this.lblCase);
            this.tabPageBasicInfo.Controls.Add(this.txtCase);
            this.tabPageBasicInfo.Location = new System.Drawing.Point(4, 22);
            this.tabPageBasicInfo.Name = "tabPageBasicInfo";
            this.tabPageBasicInfo.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageBasicInfo.Size = new System.Drawing.Size(502, 403);
            this.tabPageBasicInfo.TabIndex = 0;
            this.tabPageBasicInfo.Text = "基本資訊";
            this.tabPageBasicInfo.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.lblDateTime);
            this.groupBox1.Controls.Add(this.dtpDateTime);
            this.groupBox1.Controls.Add(this.lblLocation);
            this.groupBox1.Controls.Add(this.txtLocation);
            this.groupBox1.Controls.Add(this.lblPhotographer);
            this.groupBox1.Controls.Add(this.txtPhotographer);
            this.groupBox1.Location = new System.Drawing.Point(18, 136);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(466, 169);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "攝影資訊";
            // 
            // lblDateTime
            // 
            this.lblDateTime.AutoSize = true;
            this.lblDateTime.Location = new System.Drawing.Point(19, 35);
            this.lblDateTime.Name = "lblDateTime";
            this.lblDateTime.Size = new System.Drawing.Size(53, 12);
            this.lblDateTime.TabIndex = 5;
            this.lblDateTime.Text = "攝影時間";
            // 
            // dtpDateTime
            // 
            this.dtpDateTime.CustomFormat = "yyyy年MM月dd日 HH:mm";
            this.dtpDateTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpDateTime.Location = new System.Drawing.Point(78, 30);
            this.dtpDateTime.Name = "dtpDateTime";
            this.dtpDateTime.Size = new System.Drawing.Size(285, 22);
            this.dtpDateTime.TabIndex = 10;
            // 
            // lblLocation
            // 
            this.lblLocation.AutoSize = true;
            this.lblLocation.Location = new System.Drawing.Point(19, 77);
            this.lblLocation.Name = "lblLocation";
            this.lblLocation.Size = new System.Drawing.Size(53, 12);
            this.lblLocation.TabIndex = 6;
            this.lblLocation.Text = "攝影地點";
            // 
            // txtLocation
            // 
            this.txtLocation.Location = new System.Drawing.Point(78, 72);
            this.txtLocation.Name = "txtLocation";
            this.txtLocation.Size = new System.Drawing.Size(285, 22);
            this.txtLocation.TabIndex = 9;
            // 
            // lblPhotographer
            // 
            this.lblPhotographer.AutoSize = true;
            this.lblPhotographer.Location = new System.Drawing.Point(19, 119);
            this.lblPhotographer.Name = "lblPhotographer";
            this.lblPhotographer.Size = new System.Drawing.Size(41, 12);
            this.lblPhotographer.TabIndex = 7;
            this.lblPhotographer.Text = "攝影人";
            // 
            // txtPhotographer
            // 
            this.txtPhotographer.Location = new System.Drawing.Point(78, 114);
            this.txtPhotographer.Name = "txtPhotographer";
            this.txtPhotographer.Size = new System.Drawing.Size(285, 22);
            this.txtPhotographer.TabIndex = 8;
            // 
            // tabPagePhotos
            // 
            this.tabPagePhotos.Controls.Add(this.splitContainer);
            this.tabPagePhotos.Location = new System.Drawing.Point(4, 22);
            this.tabPagePhotos.Name = "tabPagePhotos";
            this.tabPagePhotos.Padding = new System.Windows.Forms.Padding(3);
            this.tabPagePhotos.Size = new System.Drawing.Size(502, 403);
            this.tabPagePhotos.TabIndex = 1;
            this.tabPagePhotos.Text = "照片管理";
            this.tabPagePhotos.UseVisualStyleBackColor = true;
            // 
            // splitContainer
            // 
            this.splitContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer.Location = new System.Drawing.Point(3, 3);
            this.splitContainer.Name = "splitContainer";
            // 
            // splitContainer.Panel1
            // 
            this.splitContainer.Panel1.Controls.Add(this.lvPhotos);
            this.splitContainer.Panel1.Controls.Add(this.panel1);
            // 
            // splitContainer.Panel2
            // 
            this.splitContainer.Panel2.Controls.Add(this.pbPhotoPreview);
            this.splitContainer.Panel2.Controls.Add(this.lblPreview);
            this.splitContainer.Size = new System.Drawing.Size(496, 397);
            this.splitContainer.SplitterDistance = 245;
            this.splitContainer.TabIndex = 0;
            // 
            // lvPhotos
            // 
            this.lvPhotos.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colFilename,
            this.colDate});
            this.lvPhotos.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lvPhotos.HideSelection = false;
            this.lvPhotos.Location = new System.Drawing.Point(0, 0);
            this.lvPhotos.Name = "lvPhotos";
            this.lvPhotos.Size = new System.Drawing.Size(245, 297);
            this.lvPhotos.TabIndex = 0;
            this.lvPhotos.UseCompatibleStateImageBehavior = false;
            this.lvPhotos.View = System.Windows.Forms.View.Details;
            this.lvPhotos.SelectedIndexChanged += new System.EventHandler(this.lvPhotos_SelectedIndexChanged);
            // 
            // colFilename
            // 
            this.colFilename.Text = "檔案名稱";
            this.colFilename.Width = 120;
            // 
            // colDate
            // 
            this.colDate.Text = "日期";
            this.colDate.Width = 120;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.txtPhotoDescription);
            this.panel1.Controls.Add(this.lblPhotoDescription);
            this.panel1.Controls.Add(this.btnRemovePhoto);
            this.panel1.Controls.Add(this.btnAddPhoto);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 297);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(245, 100);
            this.panel1.TabIndex = 1;
            // 
            // txtPhotoDescription
            // 
            this.txtPhotoDescription.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPhotoDescription.Enabled = false;
            this.txtPhotoDescription.Location = new System.Drawing.Point(65, 15);
            this.txtPhotoDescription.Multiline = true;
            this.txtPhotoDescription.Name = "txtPhotoDescription";
            this.txtPhotoDescription.Size = new System.Drawing.Size(167, 40);
            this.txtPhotoDescription.TabIndex = 3;
            this.txtPhotoDescription.TextChanged += new System.EventHandler(this.TxtPhotoDescription_TextChanged);
            // 
            // lblPhotoDescription
            // 
            this.lblPhotoDescription.AutoSize = true;
            this.lblPhotoDescription.Location = new System.Drawing.Point(6, 18);
            this.lblPhotoDescription.Name = "lblPhotoDescription";
            this.lblPhotoDescription.Size = new System.Drawing.Size(29, 12);
            this.lblPhotoDescription.TabIndex = 2;
            this.lblPhotoDescription.Text = "說明";
            // 
            // btnRemovePhoto
            // 
            this.btnRemovePhoto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnRemovePhoto.Location = new System.Drawing.Point(150, 65);
            this.btnRemovePhoto.Name = "btnRemovePhoto";
            this.btnRemovePhoto.Size = new System.Drawing.Size(82, 26);
            this.btnRemovePhoto.TabIndex = 1;
            this.btnRemovePhoto.Text = "移除照片";
            this.btnRemovePhoto.UseVisualStyleBackColor = true;
            this.btnRemovePhoto.Click += new System.EventHandler(this.btnRemovePhoto_Click);
            // 
            // btnAddPhoto
            // 
            this.btnAddPhoto.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAddPhoto.Location = new System.Drawing.Point(12, 65);
            this.btnAddPhoto.Name = "btnAddPhoto";
            this.btnAddPhoto.Size = new System.Drawing.Size(82, 26);
            this.btnAddPhoto.TabIndex = 0;
            this.btnAddPhoto.Text = "添加照片";
            this.btnAddPhoto.UseVisualStyleBackColor = true;
            this.btnAddPhoto.Click += new System.EventHandler(this.btnAddPhoto_Click);
            // 
            // pbPhotoPreview
            // 
            this.pbPhotoPreview.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pbPhotoPreview.BackColor = System.Drawing.Color.White;
            this.pbPhotoPreview.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pbPhotoPreview.Location = new System.Drawing.Point(3, 25);
            this.pbPhotoPreview.Name = "pbPhotoPreview";
            this.pbPhotoPreview.Size = new System.Drawing.Size(241, 372);
            this.pbPhotoPreview.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pbPhotoPreview.TabIndex = 1;
            this.pbPhotoPreview.TabStop = false;
            // 
            // lblPreview
            // 
            this.lblPreview.AutoSize = true;
            this.lblPreview.Location = new System.Drawing.Point(3, 10);
            this.lblPreview.Name = "lblPreview";
            this.lblPreview.Size = new System.Drawing.Size(65, 12);
            this.lblPreview.TabIndex = 0;
            this.lblPreview.Text = "照片預覽：";
            // 
            // btnPreview
            // 
            this.btnPreview.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnPreview.Location = new System.Drawing.Point(40, 470);
            this.btnPreview.Name = "btnPreview";
            this.btnPreview.Size = new System.Drawing.Size(120, 38);
            this.btnPreview.TabIndex = 8;
            this.btnPreview.Text = "預覽文件";
            this.btnPreview.UseVisualStyleBackColor = true;
            this.btnPreview.Click += new System.EventHandler(this.btnPreview_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(534, 539);
            this.Controls.Add(this.btnPreview);
            this.Controls.Add(this.tabControl);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.menuStrip);
            this.Controls.Add(this.btnGenerate);
            this.MainMenuStrip = this.menuStrip;
            this.MinimumSize = new System.Drawing.Size(550, 450);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "警察照片證據生成器";
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.tabControl.ResumeLayout(false);
            this.tabPageBasicInfo.ResumeLayout(false);
            this.tabPageBasicInfo.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPagePhotos.ResumeLayout(false);
            this.splitContainer.Panel1.ResumeLayout(false);
            this.splitContainer.Panel2.ResumeLayout(false);
            this.splitContainer.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer)).EndInit();
            this.splitContainer.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbPhotoPreview)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtUnit;
        private System.Windows.Forms.Label lblUnit;
        private System.Windows.Forms.Label lblCase;
        private System.Windows.Forms.TextBox txtCase;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStripMenuItem menuFile;
        private System.Windows.Forms.ToolStripMenuItem menuFileNew;
        private System.Windows.Forms.ToolStripMenuItem menuFileOpen;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem menuRecentFiles;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem menuFileExit;
        private System.Windows.Forms.ToolStripMenuItem menuSettings;
        private System.Windows.Forms.ToolStripMenuItem menuSettingsTemplate;
        private System.Windows.Forms.ToolStripMenuItem menuHelp;
        private System.Windows.Forms.ToolStripMenuItem menuHelpAbout;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel statusLabel;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabPageBasicInfo;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label lblDateTime;
        private System.Windows.Forms.DateTimePicker dtpDateTime;
        private System.Windows.Forms.Label lblLocation;
        private System.Windows.Forms.TextBox txtLocation;
        private System.Windows.Forms.Label lblPhotographer;
        private System.Windows.Forms.TextBox txtPhotographer;
        private System.Windows.Forms.TabPage tabPagePhotos;
        private System.Windows.Forms.SplitContainer splitContainer;
        private System.Windows.Forms.ListView lvPhotos;
        private System.Windows.Forms.ColumnHeader colFilename;
        private System.Windows.Forms.ColumnHeader colDate;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txtPhotoDescription;
        private System.Windows.Forms.Label lblPhotoDescription;
        private System.Windows.Forms.Button btnRemovePhoto;
        private System.Windows.Forms.Button btnAddPhoto;
        private System.Windows.Forms.PictureBox pbPhotoPreview;
        private System.Windows.Forms.Label lblPreview;
        private System.Windows.Forms.Button btnPreview;
    }
}