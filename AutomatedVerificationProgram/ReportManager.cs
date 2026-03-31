using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static AutomatedVerificationProgram.VerificationReport;

namespace AutomatedVerificationProgram
{
    
    public class ReportParaments
    {
        public string reportNunber { get; set; }
        public string sampleName { get; set; }
        public string styleItemNo { get; set; }
        public string finalExtractedDate { get; set; }
        public string company { get; set; }
        public string testParts { get; set; }
    }
    public class RowData
    {
        public string TestItem { get; set; }  // 測試項目
        public string Method { get; set; }    // 測試方法
        public string Unit { get; set; }      // 單位
        public string MDL { get; set; }       // MDL
        public string Result { get; set; }    // 檢測結果
    }
    public class ReportRowDataManager
    {
        

        public static readonly List<string> targetsOfVerification = new List<string>
        {
            "Lead",
            "Cadmium",
            "Mercury",
            "Hexavalent Chromium",
            #region PBBs
            "Monobromobiphenyl",
            "Dibromobiphenyl",
            "Tribromobiphenyl",
            "Tetrabromobiphenyl",
            "Pentabromobiphenyl",
            "Hexabromobiphenyl",
            "Heptabromobiphenyl",
            "Octabromobiphenyl",
            "Nonabromobiphenyl",
            "Decabromobiphenyl",
            #endregion
            #region PBDEs
            "Monobromodiphenyl ether",
            "Dibromodiphenyl ether",
            "Tribromodiphenyl ether",
            "Tetrabromodiphenyl ether",
            "Pentabromodiphenyl ether",
            "Hexabromodiphenyl ether",
            "Heptabromodiphenyl ether",
            "Octabromodiphenyl ether",
            "Nonabromodiphenyl ether",
            "Decabromodiphenyl ether",
            #endregion
            "(DIBP)",
            "(DEHP)",
            "(BBP)",
            "(DBP)",
            "Chlorine",
            "Bromine",
            "Fluorine",
            "(PFOS",
            "(PFOA",
            "Antimony",
            "Beryllium",
            "PVC",
            "PFAS"
        };
        // 儲存 36 筆資料的清單
        public List<RowData> reportRowData { get; set; }

        // --- 建構子 (Constructor) ---
        // 當你 new ReportManager() 時，會自動執行這裡
        public ReportRowDataManager()
        {
            // 取得標準的 37 個項目名稱
            var targets = targetsOfVerification;

            // 初始化 37 筆物件容器，並給予預設值
            this.reportRowData = targets.Select(t => new RowData
            {
                TestItem = t,
                Method = "---",
                Unit = "---",
                MDL = "---",
                Result = "---"
            }).ToList();
        }

        public void showDebug()
        {
            foreach (var item in reportRowData)
            {
                // 這樣會直接印在 Visual Studio 下方的 Output 視窗
                System.Diagnostics.Debug.WriteLine($"Item: {item.TestItem}, Method: {item.Method}, Unit: {item.Unit}, MDL: {item.MDL}, Result: {item.Result}");
            }
        }

        public void dispose()
        {
            if (this.reportRowData != null)
            {
                this.reportRowData.Clear(); // 清空列表中的所有 RowData 物件
                this.reportRowData = null;  // 切斷引用，讓 GC 更快回收
            }
            GC.SuppressFinalize(this);
        }
    }
}
