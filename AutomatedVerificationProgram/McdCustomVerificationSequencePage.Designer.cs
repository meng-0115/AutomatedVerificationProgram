namespace AutomatedVerificationProgram
{
    partial class McdCustomVerificationSequencePage
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
            this.btnImportMcdCustomVerificationSequence = new System.Windows.Forms.Button();
            this.btnDelAllMcdCustomVerificationSequence = new System.Windows.Forms.Button();
            this.btnDelMcdCustomVerificationSequence = new System.Windows.Forms.Button();
            this.btnAddMcdCustomVerificationSequence = new System.Windows.Forms.Button();
            this.dgvMcdCustomVerificationSequence = new System.Windows.Forms.DataGridView();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMcdCustomVerificationSequence)).BeginInit();
            this.SuspendLayout();
            // 
            // btnImportMcdCustomVerificationSequence
            // 
            this.btnImportMcdCustomVerificationSequence.Location = new System.Drawing.Point(620, 367);
            this.btnImportMcdCustomVerificationSequence.Name = "btnImportMcdCustomVerificationSequence";
            this.btnImportMcdCustomVerificationSequence.Size = new System.Drawing.Size(156, 53);
            this.btnImportMcdCustomVerificationSequence.TabIndex = 9;
            this.btnImportMcdCustomVerificationSequence.Text = "匯入Excel";
            this.btnImportMcdCustomVerificationSequence.UseVisualStyleBackColor = true;
            this.btnImportMcdCustomVerificationSequence.Click += new System.EventHandler(this.btnImportMcdCustomVerificationSequence_Click);
            // 
            // btnDelAllMcdCustomVerificationSequence
            // 
            this.btnDelAllMcdCustomVerificationSequence.Location = new System.Drawing.Point(421, 367);
            this.btnDelAllMcdCustomVerificationSequence.Name = "btnDelAllMcdCustomVerificationSequence";
            this.btnDelAllMcdCustomVerificationSequence.Size = new System.Drawing.Size(156, 53);
            this.btnDelAllMcdCustomVerificationSequence.TabIndex = 8;
            this.btnDelAllMcdCustomVerificationSequence.Text = "全部清除";
            this.btnDelAllMcdCustomVerificationSequence.UseVisualStyleBackColor = true;
            this.btnDelAllMcdCustomVerificationSequence.Click += new System.EventHandler(this.btnDelAllMcdCustomVerificationSequence_Click);
            // 
            // btnDelMcdCustomVerificationSequence
            // 
            this.btnDelMcdCustomVerificationSequence.Location = new System.Drawing.Point(222, 367);
            this.btnDelMcdCustomVerificationSequence.Name = "btnDelMcdCustomVerificationSequence";
            this.btnDelMcdCustomVerificationSequence.Size = new System.Drawing.Size(156, 53);
            this.btnDelMcdCustomVerificationSequence.TabIndex = 7;
            this.btnDelMcdCustomVerificationSequence.Text = "刪除優先均勻材料編號";
            this.btnDelMcdCustomVerificationSequence.UseVisualStyleBackColor = true;
            this.btnDelMcdCustomVerificationSequence.Click += new System.EventHandler(this.btnDelMcdCustomVerificationSequence_Click);
            // 
            // btnAddMcdCustomVerificationSequence
            // 
            this.btnAddMcdCustomVerificationSequence.Location = new System.Drawing.Point(23, 367);
            this.btnAddMcdCustomVerificationSequence.Name = "btnAddMcdCustomVerificationSequence";
            this.btnAddMcdCustomVerificationSequence.Size = new System.Drawing.Size(156, 53);
            this.btnAddMcdCustomVerificationSequence.TabIndex = 6;
            this.btnAddMcdCustomVerificationSequence.Text = "新增優先均勻材料編號";
            this.btnAddMcdCustomVerificationSequence.UseVisualStyleBackColor = true;
            this.btnAddMcdCustomVerificationSequence.Click += new System.EventHandler(this.btnAddMcdCustomVerificationSequence_Click);
            // 
            // dgvMcdCustomVerificationSequence
            // 
            this.dgvMcdCustomVerificationSequence.AllowUserToAddRows = false;
            this.dgvMcdCustomVerificationSequence.AllowUserToDeleteRows = false;
            this.dgvMcdCustomVerificationSequence.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvMcdCustomVerificationSequence.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvMcdCustomVerificationSequence.Location = new System.Drawing.Point(20, 31);
            this.dgvMcdCustomVerificationSequence.Name = "dgvMcdCustomVerificationSequence";
            this.dgvMcdCustomVerificationSequence.RowTemplate.Height = 24;
            this.dgvMcdCustomVerificationSequence.Size = new System.Drawing.Size(760, 308);
            this.dgvMcdCustomVerificationSequence.TabIndex = 5;
            // 
            // openFileDialog
            // 
            this.openFileDialog.FileName = "openFileDialog1";
            // 
            // McdCustomVerificationSequencePage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnImportMcdCustomVerificationSequence);
            this.Controls.Add(this.btnDelAllMcdCustomVerificationSequence);
            this.Controls.Add(this.btnDelMcdCustomVerificationSequence);
            this.Controls.Add(this.btnAddMcdCustomVerificationSequence);
            this.Controls.Add(this.dgvMcdCustomVerificationSequence);
            this.Name = "McdCustomVerificationSequencePage";
            this.Padding = new System.Windows.Forms.Padding(20);
            this.Text = "McdCustomVerificationSequencePage";
            ((System.ComponentModel.ISupportInitialize)(this.dgvMcdCustomVerificationSequence)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnImportMcdCustomVerificationSequence;
        private System.Windows.Forms.Button btnDelAllMcdCustomVerificationSequence;
        private System.Windows.Forms.Button btnDelMcdCustomVerificationSequence;
        private System.Windows.Forms.Button btnAddMcdCustomVerificationSequence;
        private System.Windows.Forms.DataGridView dgvMcdCustomVerificationSequence;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
    }
}