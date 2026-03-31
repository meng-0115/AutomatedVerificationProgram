using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace AutomatedVerificationProgram
{
    public class TwSgsReport
    {
        public class twSgsParaments
        {
            public string sampleName { get; set; }
            public string styleItemNo { get; set; }
            public string finalExtractedDate { get; set; }
            public string company { get; set; }
            public Dictionary<string, string> testParts { get; set; }
        }

        public static string finalExtractedDate = "";
        public static string testPeriodResult = "";
        public static string sampleName = "";
        public static string styleItemNo = "";
        public static string company = "";
        public static string testParts = "";
        public async Task detectionReportParament(string selectedPath,  ReportParaments reportParaments)
        {
            //傳遞給API
            string exePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tw-sgs-info-extrator-exe.exe");
            string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "twSgsParaments/twSgsReportParament.json"); // 固定輸出路徑
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                // 確保將路徑傳給外部 EXE
                Arguments = $"\"{selectedPath}\" \"{outputPath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };

            try
            {
                using (Process process = new Process { StartInfo = startInfo })
                {
                    process.Start();

                    // 異步等待 EXE 執行完畢，UI 不會卡死
                    await Task.Run(() => process.WaitForExit());

                    if (process.ExitCode == 0)
                    {
                        // 讀取並賦值
                        string jsonContent = File.ReadAllText("twSgsParaments/twSgsReportParament.json");
                        var data = JsonConvert.DeserializeObject<twSgsParaments>(jsonContent);
                        // 基本欄位賦值
                        TwSgsReport.finalExtractedDate = data.finalExtractedDate ?? "";
                        reportParaments.finalExtractedDate = TwSgsReport.finalExtractedDate;
                        TwSgsReport.sampleName = data.sampleName ?? "";
                        reportParaments.sampleName = TwSgsReport.sampleName;
                        TwSgsReport.styleItemNo = data.styleItemNo ?? "";
                        reportParaments.styleItemNo = TwSgsReport.styleItemNo;
                        TwSgsReport.company = data.company ?? "";
                        reportParaments.company = TwSgsReport.company;

                        // 處理 testParts：將 Dictionary 轉為單一字串顯示
                        if (data.testParts != null && data.testParts.Count > 0)
                        {
                            // 將所有項次串接成：No.1: 敘述; No.2: 敘述
                            TwSgsReport.testParts = string.Join("; ", data.testParts.Select(kv => $"{kv.Key}: {kv.Value}"));
                            reportParaments.testParts = TwSgsReport.testParts;
                        }
                        else
                        {
                            TwSgsReport.testParts = "";
                        }

                        //MessageBox.Show("轉換並讀取完成！");

                    }
                    else if (process.ExitCode == 1)
                    {
                        string error = await process.StandardError.ReadToEndAsync();
                        MessageBox.Show($"轉換失敗: {error}");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"執行發生錯誤: {ex.Message}");
            }
            
        }

    }
}