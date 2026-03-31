namespace AutomatedVerificationProgram
{
    partial class ManualBatchVerificationPage
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

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnStartBatchVerification = new System.Windows.Forms.Button();
            this.cbMetalNonmetal = new System.Windows.Forms.ComboBox();
            this.cbVerificationScenario = new System.Windows.Forms.ComboBox();
            this.lbMetalNonmetal = new System.Windows.Forms.Label();
            this.lbVerificationScenario = new System.Windows.Forms.Label();
            this.btnOpenFolder = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tbVerificationResult = new System.Windows.Forms.RichTextBox();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnStartBatchVerification
            // 
            this.btnStartBatchVerification.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnStartBatchVerification.Location = new System.Drawing.Point(406, 337);
            this.btnStartBatchVerification.Name = "btnStartBatchVerification";
            this.btnStartBatchVerification.Size = new System.Drawing.Size(175, 60);
            this.btnStartBatchVerification.TabIndex = 11;
            this.btnStartBatchVerification.Text = "啟動審查";
            this.btnStartBatchVerification.UseVisualStyleBackColor = true;
            this.btnStartBatchVerification.Click += new System.EventHandler(this.btnStartBatchVerification_Click);
            // 
            // cbMetalNonmetal
            // 
            this.cbMetalNonmetal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cbMetalNonmetal.FormattingEnabled = true;
            this.cbMetalNonmetal.Items.AddRange(new object[] {
            "金屬",
            "非金屬"});
            this.cbMetalNonmetal.Location = new System.Drawing.Point(406, 140);
            this.cbMetalNonmetal.Name = "cbMetalNonmetal";
            this.cbMetalNonmetal.Size = new System.Drawing.Size(165, 20);
            this.cbMetalNonmetal.TabIndex = 10;
            this.cbMetalNonmetal.SelectedIndexChanged += new System.EventHandler(this.cbMetalNonmetal_SelectedIndexChanged);
            // 
            // cbVerificationScenario
            // 
            this.cbVerificationScenario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cbVerificationScenario.FormattingEnabled = true;
            this.cbVerificationScenario.Items.AddRange(new object[] {
            "Ass\'y 直間材-金屬(1a)",
            "Ass\'y 直間材-非金屬(1b)",
            "SiP直間材(2)",
            "Ass\'y/SiP 包材-塑膠(3a)",
            "Ass\'y/SiP 包材-非塑膠(3b)"});
            this.cbVerificationScenario.Location = new System.Drawing.Point(406, 92);
            this.cbVerificationScenario.Name = "cbVerificationScenario";
            this.cbVerificationScenario.Size = new System.Drawing.Size(165, 20);
            this.cbVerificationScenario.TabIndex = 9;
            this.cbVerificationScenario.SelectedIndexChanged += new System.EventHandler(this.cbVerificationScenario_SelectedIndexChanged);
            // 
            // lbMetalNonmetal
            // 
            this.lbMetalNonmetal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbMetalNonmetal.AutoSize = true;
            this.lbMetalNonmetal.Location = new System.Drawing.Point(404, 125);
            this.lbMetalNonmetal.Name = "lbMetalNonmetal";
            this.lbMetalNonmetal.Size = new System.Drawing.Size(68, 12);
            this.lbMetalNonmetal.TabIndex = 8;
            this.lbMetalNonmetal.Text = "金屬/非金屬";
            // 
            // lbVerificationScenario
            // 
            this.lbVerificationScenario.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lbVerificationScenario.AutoSize = true;
            this.lbVerificationScenario.Location = new System.Drawing.Point(404, 77);
            this.lbVerificationScenario.Name = "lbVerificationScenario";
            this.lbVerificationScenario.Size = new System.Drawing.Size(53, 12);
            this.lbVerificationScenario.TabIndex = 7;
            this.lbVerificationScenario.Text = "審核情境";
            // 
            // btnOpenFolder
            // 
            this.btnOpenFolder.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFolder.Location = new System.Drawing.Point(406, 14);
            this.btnOpenFolder.Name = "btnOpenFolder";
            this.btnOpenFolder.Size = new System.Drawing.Size(175, 60);
            this.btnOpenFolder.TabIndex = 12;
            this.btnOpenFolder.Text = "選擇資料夾";
            this.btnOpenFolder.UseVisualStyleBackColor = true;
            this.btnOpenFolder.Click += new System.EventHandler(this.btnOpenFolder_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.Controls.Add(this.tbVerificationResult);
            this.panel1.Controls.Add(this.btnStartBatchVerification);
            this.panel1.Controls.Add(this.btnOpenFolder);
            this.panel1.Controls.Add(this.cbMetalNonmetal);
            this.panel1.Controls.Add(this.cbVerificationScenario);
            this.panel1.Controls.Add(this.lbMetalNonmetal);
            this.panel1.Controls.Add(this.lbVerificationScenario);
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(590, 410);
            this.panel1.TabIndex = 13;
            // 
            // tbVerificationResult
            // 
            this.tbVerificationResult.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbVerificationResult.Location = new System.Drawing.Point(14, 14);
            this.tbVerificationResult.Name = "tbVerificationResult";
            this.tbVerificationResult.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this.tbVerificationResult.Size = new System.Drawing.Size(375, 383);
            this.tbVerificationResult.TabIndex = 13;
            this.tbVerificationResult.Text = "";
            // 
            // ManualBatchVerificationPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Name = "ManualBatchVerificationPage";
            this.Size = new System.Drawing.Size(590, 410);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnStartBatchVerification;
        private System.Windows.Forms.ComboBox cbMetalNonmetal;
        private System.Windows.Forms.ComboBox cbVerificationScenario;
        private System.Windows.Forms.Label lbMetalNonmetal;
        private System.Windows.Forms.Label lbVerificationScenario;
        private System.Windows.Forms.Button btnOpenFolder;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.RichTextBox tbVerificationResult;
    }
}
