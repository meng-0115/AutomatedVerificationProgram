namespace AutomatedVerificationProgram
{
    partial class OcrSettingPage
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.dGVOcrCoordinatesShow = new System.Windows.Forms.DataGridView();
            this.btnSaveSetting = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dGVOcrCoordinatesShow)).BeginInit();
            this.SuspendLayout();
            // 
            // dGVOcrCoordinatesShow
            // 
            this.dGVOcrCoordinatesShow.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dGVOcrCoordinatesShow.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dGVOcrCoordinatesShow.Location = new System.Drawing.Point(12, 12);
            this.dGVOcrCoordinatesShow.Name = "dGVOcrCoordinatesShow";
            this.dGVOcrCoordinatesShow.RowTemplate.Height = 24;
            this.dGVOcrCoordinatesShow.Size = new System.Drawing.Size(565, 279);
            this.dGVOcrCoordinatesShow.TabIndex = 1;
            // 
            // btnSaveSetting
            // 
            this.btnSaveSetting.Location = new System.Drawing.Point(12, 301);
            this.btnSaveSetting.Name = "btnSaveSetting";
            this.btnSaveSetting.Size = new System.Drawing.Size(103, 43);
            this.btnSaveSetting.TabIndex = 7;
            this.btnSaveSetting.Text = "儲存設定";
            this.btnSaveSetting.UseVisualStyleBackColor = true;
            this.btnSaveSetting.Click += new System.EventHandler(this.btnSaveSetting_Click);
            // 
            // OcrSettingPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(589, 356);
            this.Controls.Add(this.btnSaveSetting);
            this.Controls.Add(this.dGVOcrCoordinatesShow);
            this.Name = "OcrSettingPage";
            this.Text = "OcrSettingPage";
            ((System.ComponentModel.ISupportInitialize)(this.dGVOcrCoordinatesShow)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dGVOcrCoordinatesShow;
        private System.Windows.Forms.Button btnSaveSetting;
    }
}