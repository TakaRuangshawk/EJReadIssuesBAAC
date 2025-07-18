namespace EJReadIssuesBAAC
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private System.Windows.Forms.DateTimePicker datePicker;
        private System.Windows.Forms.ComboBox cmbTerminalId;
        private System.Windows.Forms.Button btnDownload;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnDownloadAll;
        private System.Windows.Forms.Button btnExportEJAndCsv;
        private System.Windows.Forms.TextBox txtFolderPath;
        private System.Windows.Forms.Button btnSelectFolder;
        private void InitializeComponent()
        {
            datePicker = new DateTimePicker();
            cmbTerminalId = new ComboBox();
            btnDownload = new Button();
            lblStatus = new Label();
            btnDownloadAll = new Button();
            btnExportEJAndCsv = new Button();
            SuspendLayout();
            // 
            // datePicker
            // 
            datePicker.Location = new Point(30, 30);
            datePicker.Name = "datePicker";
            datePicker.Size = new Size(200, 23);
            datePicker.TabIndex = 2;
            datePicker.ValueChanged += datePicker_ValueChanged;
            // 
            // cmbTerminalId
            // 
            cmbTerminalId = new ComboBox();
            cmbTerminalId.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmbTerminalId.AutoCompleteSource = AutoCompleteSource.ListItems;
            cmbTerminalId.Location = new Point(30, 70);
            cmbTerminalId.Name = "cmbTerminalId";
            cmbTerminalId.Size = new Size(410, 23);
            cmbTerminalId.TabIndex = 3;
            cmbTerminalId.TextChanged += cmbTerminalId_TextChanged;
            this.Controls.Add(cmbTerminalId);
            // 
            // btnDownload
            // 
            btnDownload.Location = new Point(30, 110);
            btnDownload.Name = "btnDownload";
            btnDownload.Size = new Size(200, 30);
            btnDownload.TabIndex = 4;
            btnDownload.Text = "Download";
            btnDownload.UseVisualStyleBackColor = true;
            btnDownload.Click += btnDownload_Click;
            // 
            // lblStatus
            // 
            lblStatus.Location = new Point(30, 160);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(500, 30);
            lblStatus.TabIndex = 5;
            // 
            // btnDownloadAll
            // 
            btnDownloadAll.Location = new Point(240, 110);
            btnDownloadAll.Name = "btnDownloadAll";
            btnDownloadAll.Size = new Size(200, 30);
            btnDownloadAll.TabIndex = 0;
            btnDownloadAll.Text = "Download All";
            btnDownloadAll.UseVisualStyleBackColor = true;
            btnDownloadAll.Click += btnDownloadAll_Click;
            // 
            // btnExportEJAndCsv
            // 
            btnExportEJAndCsv.Location = new Point(30, 160);
            btnExportEJAndCsv.Name = "btnExportEJAndCsv";
            btnExportEJAndCsv.Size = new Size(200, 30);
            btnExportEJAndCsv.TabIndex = 1;
            btnExportEJAndCsv.Text = "Export Report";
            btnExportEJAndCsv.UseVisualStyleBackColor = true;
            btnExportEJAndCsv.Click += btnExportEJAndCsv_Click;
            // txtFolderPath
            txtFolderPath = new TextBox();
            txtFolderPath.Location = new Point(30, 200); // ปรับตำแหน่งตามที่ต้องการ
            txtFolderPath.Size = new Size(370, 23);
            txtFolderPath.ReadOnly = true;
            Controls.Add(txtFolderPath);
           
            // btnSelectFolder
            btnSelectFolder = new Button();
            btnSelectFolder.Text = "เลือกโฟลเดอร์...";
            btnSelectFolder.Location = new Point(410, 200); // ข้างกล่องข้อความ
            btnSelectFolder.Size = new Size(120, 23);
            btnSelectFolder.Click += btnSelectFolder_Click;
            Controls.Add(btnSelectFolder);
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(600, 276);
            Controls.Add(btnDownloadAll);
            Controls.Add(btnExportEJAndCsv);
            Controls.Add(datePicker);
            Controls.Add(cmbTerminalId);
            Controls.Add(btnDownload);
            Controls.Add(lblStatus);
            Name = "Form1";
            Text = "EJ File Downloader";
            ResumeLayout(false);
        }



        #endregion
    }
}
