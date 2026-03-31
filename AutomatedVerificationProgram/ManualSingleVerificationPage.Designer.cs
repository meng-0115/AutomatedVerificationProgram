namespace AutomatedVerificationProgram
{
    partial class ManualSingleVerificationPage
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
            this.lbVerificationScenario = new System.Windows.Forms.Label();
            this.lbMetalNonmetal = new System.Windows.Forms.Label();
            this.cbVerificationScenario = new System.Windows.Forms.ComboBox();
            this.cbMetalNonmetal = new System.Windows.Forms.ComboBox();
            this.btnStartSingeVerification = new System.Windows.Forms.Button();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.tbVerificationResult = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // lbVerificationScenario
            // 
            this.lbVerificationScenario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbVerificationScenario.AutoSize = true;
            this.lbVerificationScenario.Location = new System.Drawing.Point(404, 77);
            this.lbVerificationScenario.Name = "lbVerificationScenario";
            this.lbVerificationScenario.Size = new System.Drawing.Size(53, 12);
            this.lbVerificationScenario.TabIndex = 1;
            this.lbVerificationScenario.Text = "審核情境";
            // 
            // lbMetalNonmetal
            // 
            this.lbMetalNonmetal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbMetalNonmetal.AutoSize = true;
            this.lbMetalNonmetal.Location = new System.Drawing.Point(404, 125);
            this.lbMetalNonmetal.Name = "lbMetalNonmetal";
            this.lbMetalNonmetal.Size = new System.Drawing.Size(68, 12);
            this.lbMetalNonmetal.TabIndex = 2;
            this.lbMetalNonmetal.Text = "金屬/非金屬";
            // 
            // cbVerificationScenario
            // 
            this.cbVerificationScenario.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cbVerificationScenario.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbVerificationScenario.FormattingEnabled = true;
            this.cbVerificationScenario.IntegralHeight = false;
            this.cbVerificationScenario.Items.AddRange(new object[] {
            "Ass\'y 直間材-金屬(1a)",
            "Ass\'y 直間材-非金屬(1b)",
            "SiP直間材(2)",
            "Ass\'y/SiP 包材-塑膠(3a)",
            "Ass\'y/SiP 包材-非塑膠(3b)"});
            this.cbVerificationScenario.Location = new System.Drawing.Point(406, 92);
            this.cbVerificationScenario.Name = "cbVerificationScenario";
            this.cbVerificationScenario.Size = new System.Drawing.Size(165, 20);
            this.cbVerificationScenario.TabIndex = 3;
            this.cbVerificationScenario.SelectedIndexChanged += new System.EventHandler(this.cbVerificationScenario_SelectedIndexChanged);
            // 
            // cbMetalNonmetal
            // 
            this.cbMetalNonmetal.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cbMetalNonmetal.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbMetalNonmetal.FormattingEnabled = true;
            this.cbMetalNonmetal.Items.AddRange(new object[] {
            "金屬",
            "非金屬"});
            this.cbMetalNonmetal.Location = new System.Drawing.Point(406, 140);
            this.cbMetalNonmetal.Name = "cbMetalNonmetal";
            this.cbMetalNonmetal.Size = new System.Drawing.Size(165, 20);
            this.cbMetalNonmetal.TabIndex = 4;
            this.cbMetalNonmetal.SelectedIndexChanged += new System.EventHandler(this.cbMetalNonmetal_SelectedIndexChanged);
            // 
            // btnStartSingeVerification
            // 
            this.btnStartSingeVerification.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnStartSingeVerification.Location = new System.Drawing.Point(406, 337);
            this.btnStartSingeVerification.Name = "btnStartSingeVerification";
            this.btnStartSingeVerification.Size = new System.Drawing.Size(175, 60);
            this.btnStartSingeVerification.TabIndex = 5;
            this.btnStartSingeVerification.Text = "啟動審查";
            this.btnStartSingeVerification.UseVisualStyleBackColor = true;
            this.btnStartSingeVerification.Click += new System.EventHandler(this.btnStartSingeVerification_Click);
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpenFile.Location = new System.Drawing.Point(406, 14);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(175, 60);
            this.btnOpenFile.TabIndex = 6;
            this.btnOpenFile.Text = "開啟檔案";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog1";
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
            this.tbVerificationResult.TabIndex = 7;
            this.tbVerificationResult.Text = "";
            // 
            // ManualSingleVerificationPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.tbVerificationResult);
            this.Controls.Add(this.btnOpenFile);
            this.Controls.Add(this.btnStartSingeVerification);
            this.Controls.Add(this.cbMetalNonmetal);
            this.Controls.Add(this.cbVerificationScenario);
            this.Controls.Add(this.lbMetalNonmetal);
            this.Controls.Add(this.lbVerificationScenario);
            this.Name = "ManualSingleVerificationPage";
            this.Size = new System.Drawing.Size(590, 410);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label lbVerificationScenario;
        private System.Windows.Forms.Label lbMetalNonmetal;
        private System.Windows.Forms.ComboBox cbVerificationScenario;
        private System.Windows.Forms.ComboBox cbMetalNonmetal;
        private System.Windows.Forms.Button btnStartSingeVerification;
        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.RichTextBox tbVerificationResult;
    }
}
