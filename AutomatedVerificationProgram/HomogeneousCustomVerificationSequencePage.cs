using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.Json; 
using System.IO;
using System.Text.Encodings.Web;
using Spire.Xls;
using Spire.Xls.Core;

namespace AutomatedVerificationProgram
{
    public partial class HomogeneousCustomVerificationSequencePage : Form
    {
        public class HomogeneousCustomVerificationSequenceRowData
        {
            public string supplierId { get; set; }
            public string homogeneousNumber { get; set; }
        }
        private BindingList<HomogeneousCustomVerificationSequenceRowData> homogeneousCustomVerificationSequenceList = new BindingList<HomogeneousCustomVerificationSequenceRowData>();
        public HomogeneousCustomVerificationSequencePage()
        {
            InitializeComponent();
            dgvHomogeneousCustomVerificationSequence.DataSource = homogeneousCustomVerificationSequenceList;
        }

        private void btnAddHomogeneousCustomVerificationSequence_Click(object sender, EventArgs e)
        {
            homogeneousCustomVerificationSequenceList.Add(new HomogeneousCustomVerificationSequenceRowData
            {
                supplierId = "",
                homogeneousNumber = ""
            });
        }

        private void btnDelHomogeneousCustomVerificationSequence_Click(object sender, EventArgs e)
        {
            if (dgvHomogeneousCustomVerificationSequence.CurrentRow != null)
            {
                // 取得目前選取的物件
                var selectedItem = (HomogeneousCustomVerificationSequenceRowData)dgvHomogeneousCustomVerificationSequence.CurrentRow.DataBoundItem;
                homogeneousCustomVerificationSequenceList.Remove(selectedItem);
            }
            else
            {
                MessageBox.Show("請選取要刪除的序列資料。", "提示");
            }
        }

        private void btnDelAllHomogeneousCustomVerificationSequence_Click(object sender, EventArgs e)
        {
            if (homogeneousCustomVerificationSequenceList.Count > 0)
            {
                var result = MessageBox.Show("確定要清空所有「均質自訂驗證序列」嗎？", "確認操作",
                                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    homogeneousCustomVerificationSequenceList.Clear();
                }
            }
        }
        
        private void btnImportHomogeneousCustomVerificationSequence_Click(object sender, EventArgs e)
        {
            string selecteSequenceExceldPath = "";
            openFileDialog.Filter = "Excel 檔案 (*.xlsx;*.xls)|*.xlsx;*.xls|所有檔案 (*.*)|*.*";
            openFileDialog.Title = "匯入優先順序Excel";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selecteSequenceExceldPath = openFileDialog.FileName;
            }
            Workbook selecteSequenceExcel = new Workbook();
            selecteSequenceExcel.LoadFromFile(selecteSequenceExceldPath);
            Worksheet sheet = selecteSequenceExcel.Worksheets[0];
            homogeneousCustomVerificationSequenceList.Clear();
            for (int i = 1; i < sheet.Rows.Length; i++)
            {
                string excelSupplierId = sheet.Rows[i].Columns[0].Text;
                string excelHomogeneousNumber = sheet.Rows[i].Columns[1].Text;
                var rowData = new HomogeneousCustomVerificationSequenceRowData
                {
                    supplierId = excelSupplierId,
                    homogeneousNumber = excelHomogeneousNumber
                };
                homogeneousCustomVerificationSequenceList.Add(rowData);
            }
            dgvHomogeneousCustomVerificationSequence.ResetBindings();
        }
    }
}
