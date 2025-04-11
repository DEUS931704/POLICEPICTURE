using System;
using System.Windows.Forms;

namespace POLICEPICTURE
{
    /// <summary>
    /// 進度表單類 - 用於顯示進度
    /// </summary>
    public class ProgressForm : Form
    {
        private ProgressBar progressBar;
        private Label lblStatus;
        private Label lblPercentage;

        public ProgressForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.progressBar = new ProgressBar();
            this.lblStatus = new Label();
            this.lblPercentage = new Label();
            this.SuspendLayout();
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 40);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(360, 23);
            this.progressBar.TabIndex = 0;
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(12, 15);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(110, 12);
            this.lblStatus.TabIndex = 1;
            this.lblStatus.Text = "正在處理文件...";
            // 
            // lblPercentage
            // 
            this.lblPercentage.AutoSize = true;
            this.lblPercentage.Location = new System.Drawing.Point(334, 15);
            this.lblPercentage.Name = "lblPercentage";
            this.lblPercentage.Size = new System.Drawing.Size(21, 12);
            this.lblPercentage.TabIndex = 2;
            this.lblPercentage.Text = "0%";
            this.lblPercentage.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // ProgressForm
            // 
            this.ClientSize = new System.Drawing.Size(384, 81);
            this.Controls.Add(this.lblPercentage);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.progressBar);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProgressForm";
            this.StartPosition = FormStartPosition.CenterParent;
            this.Text = "處理進度";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        /// <summary>
        /// 更新進度值和消息
        /// </summary>
        /// <param name="percentage">進度百分比 (0-100)</param>
        /// <param name="status">進度消息</param>
        public void UpdateProgress(int percentage, string status)
        {
            if (percentage < 0) percentage = 0;
            if (percentage > 100) percentage = 100;

            progressBar.Value = percentage;
            lblStatus.Text = status;
            lblPercentage.Text = $"{percentage}%";

            // 強制更新UI
            Application.DoEvents();
        }
    }
}