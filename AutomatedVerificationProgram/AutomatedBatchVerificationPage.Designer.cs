namespace AutomatedVerificationProgram
{
    partial class AutomatedBatchVerificationPage
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
            this.lbVerificationMode = new System.Windows.Forms.Label();
            this.lbVerificationCycleTime = new System.Windows.Forms.Label();
            this.lbMin = new System.Windows.Forms.Label();
            this.numUpDownMin = new System.Windows.Forms.NumericUpDown();
            this.btnAutomatedBatchVerificationStart = new System.Windows.Forms.Button();
            this.pnlVerificationModeRadioButn = new System.Windows.Forms.Panel();
            this.rBtnMcdOnly = new System.Windows.Forms.RadioButton();
            this.rBtnOutdatedOnly = new System.Windows.Forms.RadioButton();
            this.rBtmMcdFirst = new System.Windows.Forms.RadioButton();
            this.rBtnOutdatedFrist = new System.Windows.Forms.RadioButton();
            this.lbOngoing = new System.Windows.Forms.Label();
            this.lbHomogeneousMaterialHumanReview = new System.Windows.Forms.Label();
            this.lbMcdHumanReview = new System.Windows.Forms.Label();
            this.tbOngoingResult = new System.Windows.Forms.TextBox();
            this.btnHomogeneousCustomVerificationSequence = new System.Windows.Forms.Button();
            this.btnMcdCustomVerificationSequence = new System.Windows.Forms.Button();
            this.flplHomogeneousMaterialHumanReview = new System.Windows.Forms.FlowLayoutPanel();
            this.flplMcdHumanReview = new System.Windows.Forms.FlowLayoutPanel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lbOngingItem = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.numUpDownMin)).BeginInit();
            this.pnlVerificationModeRadioButn.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbVerificationMode
            // 
            this.lbVerificationMode.AutoSize = true;
            this.lbVerificationMode.Location = new System.Drawing.Point(25, 23);
            this.lbVerificationMode.Name = "lbVerificationMode";
            this.lbVerificationMode.Size = new System.Drawing.Size(53, 12);
            this.lbVerificationMode.TabIndex = 0;
            this.lbVerificationMode.Text = "審查模式";
            // 
            // lbVerificationCycleTime
            // 
            this.lbVerificationCycleTime.AutoSize = true;
            this.lbVerificationCycleTime.Location = new System.Drawing.Point(25, 61);
            this.lbVerificationCycleTime.Name = "lbVerificationCycleTime";
            this.lbVerificationCycleTime.Size = new System.Drawing.Size(53, 12);
            this.lbVerificationCycleTime.TabIndex = 1;
            this.lbVerificationCycleTime.Text = "審核間隔";
            // 
            // lbMin
            // 
            this.lbMin.AutoSize = true;
            this.lbMin.Location = new System.Drawing.Point(145, 61);
            this.lbMin.Name = "lbMin";
            this.lbMin.Size = new System.Drawing.Size(17, 12);
            this.lbMin.TabIndex = 2;
            this.lbMin.Text = "分";
            // 
            // numUpDownMin
            // 
            this.numUpDownMin.Location = new System.Drawing.Point(84, 55);
            this.numUpDownMin.Name = "numUpDownMin";
            this.numUpDownMin.Size = new System.Drawing.Size(55, 22);
            this.numUpDownMin.TabIndex = 3;
            // 
            // btnAutomatedBatchVerificationStart
            // 
            this.btnAutomatedBatchVerificationStart.Location = new System.Drawing.Point(27, 89);
            this.btnAutomatedBatchVerificationStart.Name = "btnAutomatedBatchVerificationStart";
            this.btnAutomatedBatchVerificationStart.Size = new System.Drawing.Size(160, 25);
            this.btnAutomatedBatchVerificationStart.TabIndex = 4;
            this.btnAutomatedBatchVerificationStart.Text = "啟動審查";
            this.btnAutomatedBatchVerificationStart.UseVisualStyleBackColor = true;
            this.btnAutomatedBatchVerificationStart.Click += new System.EventHandler(this.btnAutomatedBatchVerificationStart_Click);
            // 
            // pnlVerificationModeRadioButn
            // 
            this.pnlVerificationModeRadioButn.Controls.Add(this.rBtnMcdOnly);
            this.pnlVerificationModeRadioButn.Controls.Add(this.rBtnOutdatedOnly);
            this.pnlVerificationModeRadioButn.Controls.Add(this.rBtmMcdFirst);
            this.pnlVerificationModeRadioButn.Controls.Add(this.rBtnOutdatedFrist);
            this.pnlVerificationModeRadioButn.Location = new System.Drawing.Point(84, 0);
            this.pnlVerificationModeRadioButn.Name = "pnlVerificationModeRadioButn";
            this.pnlVerificationModeRadioButn.Size = new System.Drawing.Size(371, 49);
            this.pnlVerificationModeRadioButn.TabIndex = 6;
            // 
            // rBtnMcdOnly
            // 
            this.rBtnMcdOnly.AutoSize = true;
            this.rBtnMcdOnly.Location = new System.Drawing.Point(301, 23);
            this.rBtnMcdOnly.Name = "rBtnMcdOnly";
            this.rBtnMcdOnly.Size = new System.Drawing.Size(61, 16);
            this.rBtnMcdOnly.TabIndex = 3;
            this.rBtnMcdOnly.TabStop = true;
            this.rBtnMcdOnly.Text = "僅MCD";
            this.rBtnMcdOnly.UseVisualStyleBackColor = true;
            this.rBtnMcdOnly.CheckedChanged += new System.EventHandler(this.rBtnMcdOnly_CheckedChanged);
            // 
            // rBtnOutdatedOnly
            // 
            this.rBtnOutdatedOnly.AutoSize = true;
            this.rBtnOutdatedOnly.Location = new System.Drawing.Point(207, 23);
            this.rBtnOutdatedOnly.Name = "rBtnOutdatedOnly";
            this.rBtnOutdatedOnly.Size = new System.Drawing.Size(83, 16);
            this.rBtnOutdatedOnly.TabIndex = 2;
            this.rBtnOutdatedOnly.TabStop = true;
            this.rBtnOutdatedOnly.Text = "僅過期報告";
            this.rBtnOutdatedOnly.UseVisualStyleBackColor = true;
            this.rBtnOutdatedOnly.CheckedChanged += new System.EventHandler(this.rBtnOutdatedOnly_CheckedChanged);
            // 
            // rBtmMcdFirst
            // 
            this.rBtmMcdFirst.AutoSize = true;
            this.rBtmMcdFirst.Location = new System.Drawing.Point(123, 23);
            this.rBtmMcdFirst.Name = "rBtmMcdFirst";
            this.rBtmMcdFirst.Size = new System.Drawing.Size(73, 16);
            this.rBtmMcdFirst.TabIndex = 1;
            this.rBtmMcdFirst.TabStop = true;
            this.rBtmMcdFirst.Text = "MCD優先";
            this.rBtmMcdFirst.UseVisualStyleBackColor = true;
            this.rBtmMcdFirst.CheckedChanged += new System.EventHandler(this.rBtmMcdFirst_CheckedChanged);
            // 
            // rBtnOutdatedFrist
            // 
            this.rBtnOutdatedFrist.AutoSize = true;
            this.rBtnOutdatedFrist.Location = new System.Drawing.Point(17, 23);
            this.rBtnOutdatedFrist.Name = "rBtnOutdatedFrist";
            this.rBtnOutdatedFrist.Size = new System.Drawing.Size(95, 16);
            this.rBtnOutdatedFrist.TabIndex = 0;
            this.rBtnOutdatedFrist.TabStop = true;
            this.rBtnOutdatedFrist.Text = "過期報告優先";
            this.rBtnOutdatedFrist.UseVisualStyleBackColor = true;
            this.rBtnOutdatedFrist.CheckedChanged += new System.EventHandler(this.rBtnOutdatedFrist_CheckedChanged);
            // 
            // lbOngoing
            // 
            this.lbOngoing.AutoSize = true;
            this.lbOngoing.Location = new System.Drawing.Point(26, 120);
            this.lbOngoing.Name = "lbOngoing";
            this.lbOngoing.Size = new System.Drawing.Size(52, 12);
            this.lbOngoing.TabIndex = 7;
            this.lbOngoing.Text = "Ongoing: ";
            // 
            // lbHomogeneousMaterialHumanReview
            // 
            this.lbHomogeneousMaterialHumanReview.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lbHomogeneousMaterialHumanReview.AutoSize = true;
            this.lbHomogeneousMaterialHumanReview.Location = new System.Drawing.Point(327, 120);
            this.lbHomogeneousMaterialHumanReview.Name = "lbHomogeneousMaterialHumanReview";
            this.lbHomogeneousMaterialHumanReview.Size = new System.Drawing.Size(65, 12);
            this.lbHomogeneousMaterialHumanReview.TabIndex = 8;
            this.lbHomogeneousMaterialHumanReview.Text = "均質轉人審";
            // 
            // lbMcdHumanReview
            // 
            this.lbMcdHumanReview.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.lbMcdHumanReview.AutoSize = true;
            this.lbMcdHumanReview.Location = new System.Drawing.Point(325, 258);
            this.lbMcdHumanReview.Name = "lbMcdHumanReview";
            this.lbMcdHumanReview.Size = new System.Drawing.Size(67, 12);
            this.lbMcdHumanReview.TabIndex = 9;
            this.lbMcdHumanReview.Text = "MCD轉人審";
            // 
            // tbOngoingResult
            // 
            this.tbOngoingResult.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbOngoingResult.Location = new System.Drawing.Point(27, 137);
            this.tbOngoingResult.Multiline = true;
            this.tbOngoingResult.Name = "tbOngoingResult";
            this.tbOngoingResult.Size = new System.Drawing.Size(253, 253);
            this.tbOngoingResult.TabIndex = 10;
            // 
            // btnHomogeneousCustomVerificationSequence
            // 
            this.btnHomogeneousCustomVerificationSequence.Location = new System.Drawing.Point(168, 55);
            this.btnHomogeneousCustomVerificationSequence.Name = "btnHomogeneousCustomVerificationSequence";
            this.btnHomogeneousCustomVerificationSequence.Size = new System.Drawing.Size(160, 25);
            this.btnHomogeneousCustomVerificationSequence.TabIndex = 13;
            this.btnHomogeneousCustomVerificationSequence.Text = "自訂審核順序: 均質 ";
            this.btnHomogeneousCustomVerificationSequence.UseVisualStyleBackColor = true;
            this.btnHomogeneousCustomVerificationSequence.Click += new System.EventHandler(this.btnHomogeneousCustomVerificationSequence_Click);
            // 
            // btnMcdCustomVerificationSequence
            // 
            this.btnMcdCustomVerificationSequence.Location = new System.Drawing.Point(334, 55);
            this.btnMcdCustomVerificationSequence.Name = "btnMcdCustomVerificationSequence";
            this.btnMcdCustomVerificationSequence.Size = new System.Drawing.Size(160, 25);
            this.btnMcdCustomVerificationSequence.TabIndex = 14;
            this.btnMcdCustomVerificationSequence.Text = "自訂審核順序: MCD ";
            this.btnMcdCustomVerificationSequence.UseVisualStyleBackColor = true;
            this.btnMcdCustomVerificationSequence.Click += new System.EventHandler(this.btnMcdCustomVerificationSequence_Click);
            // 
            // flplHomogeneousMaterialHumanReview
            // 
            this.flplHomogeneousMaterialHumanReview.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.flplHomogeneousMaterialHumanReview.BackColor = System.Drawing.SystemColors.Window;
            this.flplHomogeneousMaterialHumanReview.Location = new System.Drawing.Point(327, 137);
            this.flplHomogeneousMaterialHumanReview.Name = "flplHomogeneousMaterialHumanReview";
            this.flplHomogeneousMaterialHumanReview.Size = new System.Drawing.Size(245, 115);
            this.flplHomogeneousMaterialHumanReview.TabIndex = 15;
            // 
            // flplMcdHumanReview
            // 
            this.flplMcdHumanReview.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.flplMcdHumanReview.BackColor = System.Drawing.SystemColors.Window;
            this.flplMcdHumanReview.Location = new System.Drawing.Point(327, 275);
            this.flplMcdHumanReview.Name = "flplMcdHumanReview";
            this.flplMcdHumanReview.Size = new System.Drawing.Size(245, 115);
            this.flplMcdHumanReview.TabIndex = 16;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.Control;
            this.panel1.Controls.Add(this.lbOngingItem);
            this.panel1.Controls.Add(this.lbVerificationMode);
            this.panel1.Controls.Add(this.pnlVerificationModeRadioButn);
            this.panel1.Controls.Add(this.lbHomogeneousMaterialHumanReview);
            this.panel1.Controls.Add(this.lbMcdHumanReview);
            this.panel1.Controls.Add(this.lbOngoing);
            this.panel1.Controls.Add(this.tbOngoingResult);
            this.panel1.Controls.Add(this.btnAutomatedBatchVerificationStart);
            this.panel1.Controls.Add(this.numUpDownMin);
            this.panel1.Controls.Add(this.btnHomogeneousCustomVerificationSequence);
            this.panel1.Controls.Add(this.btnMcdCustomVerificationSequence);
            this.panel1.Controls.Add(this.lbVerificationCycleTime);
            this.panel1.Controls.Add(this.lbMin);
            this.panel1.Controls.Add(this.flplHomogeneousMaterialHumanReview);
            this.panel1.Controls.Add(this.flplMcdHumanReview);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(590, 410);
            this.panel1.TabIndex = 17;
            // 
            // lbOngingItem
            // 
            this.lbOngingItem.AutoSize = true;
            this.lbOngingItem.Location = new System.Drawing.Point(74, 120);
            this.lbOngingItem.Name = "lbOngingItem";
            this.lbOngingItem.Size = new System.Drawing.Size(0, 12);
            this.lbOngingItem.TabIndex = 17;
            this.lbOngingItem.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // AutomatedBatchVerificationPage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panel1);
            this.Name = "AutomatedBatchVerificationPage";
            this.Size = new System.Drawing.Size(590, 410);
            this.Load += new System.EventHandler(this.AutomatedBatchVerificationPage_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numUpDownMin)).EndInit();
            this.pnlVerificationModeRadioButn.ResumeLayout(false);
            this.pnlVerificationModeRadioButn.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lbVerificationMode;
        private System.Windows.Forms.Label lbVerificationCycleTime;
        private System.Windows.Forms.Label lbMin;
        private System.Windows.Forms.NumericUpDown numUpDownMin;
        private System.Windows.Forms.Button btnAutomatedBatchVerificationStart;
        private System.Windows.Forms.Panel pnlVerificationModeRadioButn;
        private System.Windows.Forms.RadioButton rBtnMcdOnly;
        private System.Windows.Forms.RadioButton rBtnOutdatedOnly;
        private System.Windows.Forms.RadioButton rBtmMcdFirst;
        private System.Windows.Forms.RadioButton rBtnOutdatedFrist;
        private System.Windows.Forms.Label lbOngoing;
        private System.Windows.Forms.Label lbHomogeneousMaterialHumanReview;
        private System.Windows.Forms.Label lbMcdHumanReview;
        private System.Windows.Forms.TextBox tbOngoingResult;
        private System.Windows.Forms.Button btnHomogeneousCustomVerificationSequence;
        private System.Windows.Forms.Button btnMcdCustomVerificationSequence;
        private System.Windows.Forms.FlowLayoutPanel flplHomogeneousMaterialHumanReview;
        private System.Windows.Forms.FlowLayoutPanel flplMcdHumanReview;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lbOngingItem;
    }
}
