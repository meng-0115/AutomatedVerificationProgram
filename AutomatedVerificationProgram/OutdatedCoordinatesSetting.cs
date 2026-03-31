using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using System.IO;

namespace AutomatedVerificationProgram
{
    public partial class OutdatedCoordinatesSetting : Form
    {
        public class ClickStep
        {
            public string ActionType { get; set; }
            public string X { get; set; }
            public string Y { get; set; }
            public string Param { get; set; }
        }
        int indexOfParagraph = -1; // 0: 第一段, 1: 第二段, 2: 第三段, 3: 第四段, 4: 第五段
        string[] paragraphFilesPaths = new string[]
        {
            @"Coordinates/outdatedReportFirstParagraphCoordinatesSetting.json",
            @"Coordinates/outdatedReportSecondParagraphCoordinatesSetting.json",
            @"Coordinates/outdatedReportThirdParagraphCoordinatesSetting.json",
            @"Coordinates/outdatedReportForthParagraphCoordinatesSetting.json",
            @"Coordinates/outdatedReportFivethParagraphCoordinatesSetting.json",
            @"Coordinates/outdatedReportSixthParagraphCoordinatesSetting.json",
            @"Coordinates/outdatedReportSeventhParagraphCoordinatesSetting.json",
            @"Coordinates/outdatedReportEighthParagraphCoordinatesSetting.json"
        };
        public OutdatedCoordinatesSetting()
        {
            InitializeComponent();
            SetupDataGridView();
            btnSaveSetting.Enabled = false; // 預設先禁用，等使用者選擇段落後才啟用
        }

        private void SetupDataGridView()
        {
            dGVCoordinatesShow.Columns.Clear();
            dGVCoordinatesShow.AutoGenerateColumns = false;

            // 1. 類型 (ComboBoxColumn)
            DataGridViewComboBoxColumn typeCol = new DataGridViewComboBoxColumn();
            typeCol.Name = "ActionType";
            typeCol.HeaderText = "動作類型";
            typeCol.Items.AddRange("左鍵單下點擊", "右鍵單下點擊", "左鍵雙下點擊", "右鍵雙下點擊", "等待時間", "貼上Excel文字", "檔案偏移");

            dGVCoordinatesShow.Columns.Add(typeCol);

            // 2. X 座標 (TextBoxColumn)
            dGVCoordinatesShow.Columns.Add(new DataGridViewTextBoxColumn { Name = "X", HeaderText = "X 座標", Width = 70 });

            // 3. Y 座標 (TextBoxColumn)
            dGVCoordinatesShow.Columns.Add(new DataGridViewTextBoxColumn { Name = "Y", HeaderText = "Y 座標", Width = 70 });

            // 4. 參數 (TextBoxColumn)
            dGVCoordinatesShow.Columns.Add(new DataGridViewTextBoxColumn { Name = "Param", HeaderText = "參數 (ms/偏移值)", Width = 100 });

        }
        private void btnFirstParagraphSetting_Click(object sender, EventArgs e)
        {
            indexOfParagraph = 0;
            btnSaveSetting.Enabled = true;
            string json = File.ReadAllText(paragraphFilesPaths[indexOfParagraph]);
            List<ClickStep> steps = JsonConvert.DeserializeObject<List<ClickStep>>(json);
            dGVCoordinatesShow.Rows.Clear();
            foreach (var step in steps)
            {
                dGVCoordinatesShow.Rows.Add(step.ActionType, step.X, step.Y, step.Param);
            }
           
        }

        private void btnSecondParagraphSetting_Click(object sender, EventArgs e)
        {
            indexOfParagraph = 1;
            btnSaveSetting.Enabled = true;
            string json = File.ReadAllText(paragraphFilesPaths[indexOfParagraph]);
            List<ClickStep> steps = JsonConvert.DeserializeObject<List<ClickStep>>(json);
            dGVCoordinatesShow.Rows.Clear();
            foreach (var step in steps)
            {
                dGVCoordinatesShow.Rows.Add(step.ActionType, step.X, step.Y, step.Param);
            }
        }

        private void btnThirdParagraphSetting_Click(object sender, EventArgs e)
        {
            indexOfParagraph = 2;
            btnSaveSetting.Enabled = true;
            string json = File.ReadAllText(paragraphFilesPaths[indexOfParagraph]);
            List<ClickStep> steps = JsonConvert.DeserializeObject<List<ClickStep>>(json);
            dGVCoordinatesShow.Rows.Clear();
            foreach (var step in steps)
            {
                dGVCoordinatesShow.Rows.Add(step.ActionType, step.X, step.Y, step.Param);
            }
        }

        private void btnForthParagraphSetting_Click(object sender, EventArgs e)
        {
            indexOfParagraph = 3;
            btnSaveSetting.Enabled = true;
            string json = File.ReadAllText(paragraphFilesPaths[indexOfParagraph]);
            List<ClickStep> steps = JsonConvert.DeserializeObject<List<ClickStep>>(json);
            dGVCoordinatesShow.Rows.Clear();
            foreach (var step in steps)
            {
                dGVCoordinatesShow.Rows.Add(step.ActionType, step.X, step.Y, step.Param);
            }
        }

        private void btnFivethParagraphSetting_Click(object sender, EventArgs e)
        {
            indexOfParagraph = 4;
            btnSaveSetting.Enabled = true;
            string json = File.ReadAllText(paragraphFilesPaths[indexOfParagraph]);
            List<ClickStep> steps = JsonConvert.DeserializeObject<List<ClickStep>>(json);
            dGVCoordinatesShow.Rows.Clear();
            foreach (var step in steps)
            {
                dGVCoordinatesShow.Rows.Add(step.ActionType, step.X, step.Y, step.Param);
            }
        }

        private void btnSixthParagraphSetting_Click(object sender, EventArgs e)
        {
            indexOfParagraph = 5;
            btnSaveSetting.Enabled = true;
            string json = File.ReadAllText(paragraphFilesPaths[indexOfParagraph]);
            List<ClickStep> steps = JsonConvert.DeserializeObject<List<ClickStep>>(json);
            dGVCoordinatesShow.Rows.Clear();
            foreach (var step in steps)
            {
                dGVCoordinatesShow.Rows.Add(step.ActionType, step.X, step.Y, step.Param);
            }
        }

        private void btnSeventhParagraphSetting_Click(object sender, EventArgs e)
        {
            indexOfParagraph = 6;
            btnSaveSetting.Enabled = true;
            string json = File.ReadAllText(paragraphFilesPaths[indexOfParagraph]);
            List<ClickStep> steps = JsonConvert.DeserializeObject<List<ClickStep>>(json);
            dGVCoordinatesShow.Rows.Clear();
            foreach (var step in steps)
            {
                dGVCoordinatesShow.Rows.Add(step.ActionType, step.X, step.Y, step.Param);
            }
        }
        private void btnEightthParagraphSetting_Click(object sender, EventArgs e)
        {
            indexOfParagraph = 7;
            btnSaveSetting.Enabled = true;
            string json = File.ReadAllText(paragraphFilesPaths[indexOfParagraph]);
            List<ClickStep> steps = JsonConvert.DeserializeObject<List<ClickStep>>(json);
            dGVCoordinatesShow.Rows.Clear();
            foreach (var step in steps)
            {
                dGVCoordinatesShow.Rows.Add(step.ActionType, step.X, step.Y, step.Param);
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
                List<ClickStep> saveSteps = new List<ClickStep>();
                foreach (DataGridViewRow row in dGVCoordinatesShow.Rows)
                {
                    if (row.IsNewRow) continue;
                    saveSteps.Add(new ClickStep
                    {
                        ActionType = row.Cells["ActionType"].Value?.ToString(),
                        X = row.Cells["X"].Value?.ToString(),
                        Y = row.Cells["Y"].Value?.ToString(),
                        Param = row.Cells["Param"].Value?.ToString()
                    });
                }
                string json = JsonConvert.SerializeObject(saveSteps, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(paragraphFilesPaths[indexOfParagraph], json);
            }
        }

        
    }
}
