using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;
using System.Runtime.InteropServices;
using System.IO;

using System.Windows.Forms;

namespace AutomatedVerificationProgram
{
    public partial class OutdatedReportCoordinatesSettingForm : Form
    {
        public class OutdatedReportClickStepList
        {
            public string ActionType { get; set; }
            public string X { get; set; }
            public string Y { get; set; }
            public string Param { get; set; }
            public string Memo { get; set; }
        }
        public OutdatedReportCoordinatesSettingForm()
        {
            InitializeComponent();
            OutdatedReportCoordinatesStepsSetupDataGridView();
            string json = File.ReadAllText("outdatedReportCoordinatesSteps.json");
            List<OutdatedReportClickStepList> outdatedReportCoordinatesSteps = JsonConvert.DeserializeObject<List<OutdatedReportClickStepList>>(json);

            outdatedReportCoordinatesDataGridView.Rows.Clear(); // 清除當前內容
            foreach (var step in outdatedReportCoordinatesSteps)
            {
                outdatedReportCoordinatesDataGridView.Rows.Add(step.ActionType, step.X, step.Y, step.Param, step.Memo);
            }
        }
        private void OutdatedReportCoordinatesStepsSetupDataGridView()
        {
            // 1. 基本設定與清除
            outdatedReportCoordinatesDataGridView.Columns.Clear();
            outdatedReportCoordinatesDataGridView.AutoGenerateColumns = false;

            
            // 3. 關鍵：讓所有欄位自動平分剩餘空間
            outdatedReportCoordinatesDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            // --- 建立欄位 ---

            // 動作類型 (ComboBox)
            DataGridViewComboBoxColumn col1ActionType = new DataGridViewComboBoxColumn();
            col1ActionType.Name = "ActionType";
            col1ActionType.HeaderText = "動作類型";
            col1ActionType.Items.AddRange("左鍵點擊1下", "右鍵點擊1下", "左鍵點擊2下", "等待時間");
            outdatedReportCoordinatesDataGridView.Columns.Add(col1ActionType);

            // X 座標
            outdatedReportCoordinatesDataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "X",
                HeaderText = "X 座標"
            });

            // Y 座標
            outdatedReportCoordinatesDataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Y",
                HeaderText = "Y 座標"
            });

            // 參數 (ms)
            outdatedReportCoordinatesDataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Param",
                HeaderText = "參數 (ms)"
            });
        }
    }
    }
