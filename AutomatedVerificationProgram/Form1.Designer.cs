namespace AutomatedVerificationProgram
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.pnlMainContent = new System.Windows.Forms.Panel();
            this.pnlSideMenu = new System.Windows.Forms.Panel();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.btnAutomatedBatchVerificationPage = new System.Windows.Forms.Button();
            this.btnManualSingleVerificationPage = new System.Windows.Forms.Button();
            this.btnManualBatchVerificationPage = new System.Windows.Forms.Button();
            this.btnSystemSettingsPage = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.pnlSideMenu.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // pnlMainContent
            // 
            this.pnlMainContent.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pnlMainContent.BackColor = System.Drawing.SystemColors.Control;
            this.pnlMainContent.Location = new System.Drawing.Point(187, 20);
            this.pnlMainContent.Name = "pnlMainContent";
            this.pnlMainContent.Size = new System.Drawing.Size(590, 410);
            this.pnlMainContent.TabIndex = 1;
            // 
            // pnlSideMenu
            // 
            this.pnlSideMenu.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.pnlSideMenu.BackColor = System.Drawing.SystemColors.Control;
            this.pnlSideMenu.Controls.Add(this.tableLayoutPanel1);
            this.pnlSideMenu.Location = new System.Drawing.Point(20, 20);
            this.pnlSideMenu.Name = "pnlSideMenu";
            this.pnlSideMenu.Size = new System.Drawing.Size(161, 410);
            this.pnlSideMenu.TabIndex = 0;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.btnAutomatedBatchVerificationPage, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.btnManualSingleVerificationPage, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.btnManualBatchVerificationPage, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.btnSystemSettingsPage, 0, 3);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(161, 252);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // btnAutomatedBatchVerificationPage
            // 
            this.btnAutomatedBatchVerificationPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnAutomatedBatchVerificationPage.Location = new System.Drawing.Point(3, 3);
            this.btnAutomatedBatchVerificationPage.Name = "btnAutomatedBatchVerificationPage";
            this.btnAutomatedBatchVerificationPage.Size = new System.Drawing.Size(155, 57);
            this.btnAutomatedBatchVerificationPage.TabIndex = 0;
            this.btnAutomatedBatchVerificationPage.Text = "自動批次審查";
            this.btnAutomatedBatchVerificationPage.UseVisualStyleBackColor = true;
            this.btnAutomatedBatchVerificationPage.Click += new System.EventHandler(this.btnAutomatedBatchVerificationPage_Click);
            // 
            // btnManualSingleVerificationPage
            // 
            this.btnManualSingleVerificationPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnManualSingleVerificationPage.Location = new System.Drawing.Point(3, 66);
            this.btnManualSingleVerificationPage.Name = "btnManualSingleVerificationPage";
            this.btnManualSingleVerificationPage.Size = new System.Drawing.Size(155, 57);
            this.btnManualSingleVerificationPage.TabIndex = 1;
            this.btnManualSingleVerificationPage.Text = "手動單筆審查";
            this.btnManualSingleVerificationPage.UseVisualStyleBackColor = true;
            this.btnManualSingleVerificationPage.Click += new System.EventHandler(this.btnManualSingleVerificationPage_Click);
            // 
            // btnManualBatchVerificationPage
            // 
            this.btnManualBatchVerificationPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnManualBatchVerificationPage.Location = new System.Drawing.Point(3, 129);
            this.btnManualBatchVerificationPage.Name = "btnManualBatchVerificationPage";
            this.btnManualBatchVerificationPage.Size = new System.Drawing.Size(155, 57);
            this.btnManualBatchVerificationPage.TabIndex = 2;
            this.btnManualBatchVerificationPage.Text = "手動批次審查";
            this.btnManualBatchVerificationPage.UseVisualStyleBackColor = true;
            this.btnManualBatchVerificationPage.Click += new System.EventHandler(this.btnManualBatchVerificationPage_Click);
            // 
            // btnSystemSettingsPage
            // 
            this.btnSystemSettingsPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnSystemSettingsPage.Location = new System.Drawing.Point(3, 192);
            this.btnSystemSettingsPage.Name = "btnSystemSettingsPage";
            this.btnSystemSettingsPage.Size = new System.Drawing.Size(155, 57);
            this.btnSystemSettingsPage.TabIndex = 3;
            this.btnSystemSettingsPage.Text = "參數設置";
            this.btnSystemSettingsPage.UseVisualStyleBackColor = true;
            this.btnSystemSettingsPage.Click += new System.EventHandler(this.btnSystemSettingsPage_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.pnlMainContent);
            this.Controls.Add(this.pnlSideMenu);
            this.Name = "Form1";
            this.Padding = new System.Windows.Forms.Padding(20);
            this.Text = "Form1";
            this.pnlSideMenu.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Panel pnlMainContent;
        private System.Windows.Forms.Panel pnlSideMenu;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Button btnAutomatedBatchVerificationPage;
        private System.Windows.Forms.Button btnManualSingleVerificationPage;
        private System.Windows.Forms.Button btnManualBatchVerificationPage;
        private System.Windows.Forms.Button btnSystemSettingsPage;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}

