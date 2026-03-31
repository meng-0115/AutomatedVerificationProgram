using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static AutomatedVerificationProgram.HomogeneousCustomVerificationSequencePage;
using Spire.Xls;

namespace AutomatedVerificationProgram
{
    public partial class McdCustomVerificationSequencePage : Form
    {
        public class McdCustomVerificationSequenceRowData
        {
            public string supplierId { get; set; }
            public string itemNumber { get; set; }
        }
        private BindingList<McdCustomVerificationSequenceRowData> mcdCustomVerificationSequenceList = new BindingList<McdCustomVerificationSequenceRowData>();
        public McdCustomVerificationSequencePage()
        {
            InitializeComponent();
            dgvMcdCustomVerificationSequence.DataSource = mcdCustomVerificationSequenceList;
        }

        private void btnAddMcdCustomVerificationSequence_Click(object sender, EventArgs e)
        {
            mcdCustomVerificationSequenceList.Add(new McdCustomVerificationSequenceRowData
            {
                supplierId = "",
                itemNumber = ""
            });
        }

        private void btnDelMcdCustomVerificationSequence_Click(object sender, EventArgs e)
        {
            if (dgvMcdCustomVerificationSequence.CurrentRow != null)
            {
                // 取得目前選取的物件
                var selectedItem = (McdCustomVerificationSequenceRowData)dgvMcdCustomVerificationSequence.CurrentRow.DataBoundItem;
                mcdCustomVerificationSequenceList.Remove(selectedItem);
            }
            else
            {
                MessageBox.Show("請選取要刪除的序列資料。", "提示");
            }
        }

        private void btnDelAllMcdCustomVerificationSequence_Click(object sender, EventArgs e)
        {
            if (mcdCustomVerificationSequenceList.Count > 0)
            {
                var result = MessageBox.Show("確定要清空所有「MCD自訂驗證序列」嗎？", "確認操作",
                                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    mcdCustomVerificationSequenceList.Clear();
                }
            }
        }

        private void btnImportMcdCustomVerificationSequence_Click(object sender, EventArgs e)
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
            mcdCustomVerificationSequenceList.Clear();
            for (int i = 1; i < sheet.Rows.Length; i++)
            {
                string excelSupplierId = sheet.Rows[i].Columns[0].Text;
                string excelItemNumber = sheet.Rows[i].Columns[1].Text;
                var rowData = new McdCustomVerificationSequenceRowData
                {
                    supplierId = excelSupplierId,
                    itemNumber = excelItemNumber
                };
                mcdCustomVerificationSequenceList.Add(rowData);
            }
            mcdCustomVerificationSequenceList.ResetBindings();
        }
    }
}
