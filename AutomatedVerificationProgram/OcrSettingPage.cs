using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using static AutomatedVerificationProgram.OutdatedCoordinatesSetting;

namespace AutomatedVerificationProgram
{
    public partial class OcrSettingPage : Form
    {
        private void SetupDataGridView()
        {
            dGVOcrCoordinatesShow.Columns.Clear();
            dGVOcrCoordinatesShow.AutoGenerateColumns = false;
            dGVOcrCoordinatesShow.Columns.Add(new DataGridViewTextBoxColumn { Name = "X1", HeaderText = "X1 座標"});
            dGVOcrCoordinatesShow.Columns.Add(new DataGridViewTextBoxColumn { Name = "X2", HeaderText = "X2 座標"});
            dGVOcrCoordinatesShow.Columns.Add(new DataGridViewTextBoxColumn { Name = "Y1", HeaderText = "Y1 座標"});
            dGVOcrCoordinatesShow.Columns.Add(new DataGridViewTextBoxColumn { Name = "Y2", HeaderText = "Y2 座標"});


        }
        public class OcrParameterClass
        {
            public int X1 { get; set; }
            public int X2 { get; set; }
            public int Y1 { get; set; }
            public int Y2 { get; set; }
        }
        public OcrSettingPage()
        {
            InitializeComponent();
            SetupDataGridView();
            string json = File.ReadAllText(@"Coordinates/ocrCoordinatesSetting.json");
            List<OcrParameterClass> OcrParameters = JsonConvert.DeserializeObject<List<OcrParameterClass>>(json);
            dGVOcrCoordinatesShow.Rows.Clear();
            foreach (var OcrParameter in OcrParameters)
            {
                dGVOcrCoordinatesShow.Rows.Add(OcrParameter.X1, OcrParameter.X2, OcrParameter.Y1, OcrParameter.Y2);
            }
        }

        private void btnSaveSetting_Click(object sender, EventArgs e)
        {
            DialogResult saveResponceResult = MessageBox.Show("確定要儲存資料嗎？", "確認", MessageBoxButtons.YesNoCancel);
            // 1. 轉成整數存檔 (Yes=6, No=7, Cancel=2)
            int resultCode = (int)saveResponceResult;
            // 2. 轉成字串存檔
            string resultString = saveResponceResult.ToString(); // "Yes" 或 "No"
            if (saveResponceResult == DialogResult.Yes)
            {
                List<OcrParameterClass> OcrParameters = new List<OcrParameterClass>();
                foreach (DataGridViewRow row in dGVOcrCoordinatesShow.Rows)
                {
                    if (row.IsNewRow) continue;
                    OcrParameters.Add(new OcrParameterClass
                    {
                        X1 = Convert.ToInt32(row.Cells["X1"].Value),
                        X2 = Convert.ToInt32(row.Cells["X2"].Value),
                        Y1 = Convert.ToInt32(row.Cells["Y1"].Value),
                        Y2 = Convert.ToInt32(row.Cells["Y2"].Value)
                    });
                }
                string json = JsonConvert.SerializeObject(OcrParameters, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(@"Coordinates/ocrCoordinatesSetting.json", json);
            }
        }
    }
}
