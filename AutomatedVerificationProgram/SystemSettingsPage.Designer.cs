namespace AutomatedVerificationProgram
{
    partial class SystemSettingsPage
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
            this.btnOutdatedCoordinatesSetting = new System.Windows.Forms.Button();
            this.btnmcdCoordinatesSetting = new System.Windows.Forms.Button();
            this.btnOcrSetting = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnOutdatedCoordinatesSetting
            // 
            this.btnOutdatedCoordinatesSetting.Location = new System.Drawing.Point(14, 24);
            this.btnOutdatedCoordinatesSetting.Name = "btnOutdatedCoordinatesSetting";
            this.btnOutdatedCoordinatesSetting.Size = new System.Drawing.Size(154, 47);
            this.btnOutdatedCoordinatesSetting.TabIndex = 5;
            this.btnOutdatedCoordinatesSetting.Text = "過期報告座標設定";
            this.btnOutdatedCoordinatesSetting.UseVisualStyleBackColor = true;
            this.btnOutdatedCoordinatesSetting.Click += new System.EventHandler(this.btnOutdatedCoordinatesSetting_Click);
            // 
            // btnmcdCoordinatesSetting
            // 
            this.btnmcdCoordinatesSetting.Location = new System.Drawing.Point(14, 96);
            this.btnmcdCoordinatesSetting.Name = "btnmcdCoordinatesSetting";
            this.btnmcdCoordinatesSetting.Size = new System.Drawing.Size(154, 47);
            this.btnmcdCoordinatesSetting.TabIndex = 6;
            this.btnmcdCoordinatesSetting.Text = "過期報告座標設定";
            this.btnmcdCoordinatesSetting.UseVisualStyleBackColor = true;
            this.btnmcdCoordinatesSetting.Click += new System.EventHandler(this.btnmcdCoordinatesSetting_Click);
            // 
            // btnOcrSetting
            // 
            this.btnOcrSetting.Location = new System.Drawing.Point(14, 159);
            this.btnOcrSetting.Name = "btnOcrSetting";
            this.btnOcrSetting.Size = new System.Drawing.Size(154, 47);
            this.btnOcrSetting.TabIndex = 7;
            this.btnOcrSetting.Text = "OCR 四角落座標設定";
            this.btnOcrSetting.UseVisualStyleBackColor = true;
            this.btnOcrSetting.Click += new System.EventHandler(this.btnOcrSetting_Click);
            // 
            // SystemSettingsPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnOcrSetting);
            this.Controls.Add(this.btnmcdCoordinatesSetting);
            this.Controls.Add(this.btnOutdatedCoordinatesSetting);
            this.Name = "SystemSettingsPage";
            this.Size = new System.Drawing.Size(590, 410);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnOutdatedCoordinatesSetting;
        private System.Windows.Forms.Button btnmcdCoordinatesSetting;
        private System.Windows.Forms.Button btnOcrSetting;
    }
}
