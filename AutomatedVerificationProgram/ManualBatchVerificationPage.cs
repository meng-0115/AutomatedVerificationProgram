using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace AutomatedVerificationProgram
{
    public partial class ManualBatchVerificationPage : UserControl
    {
        int verificationScenarioIndex = -1;
        int metalNonmetalIndex = -1;
        string rootPath = "";
        public static string[] pdfReportFilePage2StringSplitArray;
        public ManualBatchVerificationPage()
        {
            InitializeComponent();
            btnStartBatchVerification.Enabled = false;
        }

        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            folderBrowserDialog.Description = "請選擇批次審核文件夾";
            folderBrowserDialog.ShowNewFolderButton = false;
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                rootPath = folderBrowserDialog.SelectedPath;
            }
        }

        private async void btnStartBatchVerification_Click(object sender, EventArgs e)
        {
            btnStartBatchVerification.Enabled = false;
            string[] subDirectories = Directory.GetDirectories(rootPath, "*", SearchOption.AllDirectories);

            foreach (string dirPath in subDirectories)
            {
                List<List<string>> list2Excel = new List<List<string>>();
                ReportRowDataManager combineAllReportRowData = new ReportRowDataManager();             
                string folderName = Path.GetFileName(dirPath);
                string[] pdfFiles = Directory.GetFiles(dirPath, "*.pdf");
                uint pdfCount = 1;
                foreach (string pdfPath in pdfFiles)
                {
                    ReportRowDataManager reportRowData = new ReportRowDataManager();
                    ReportParaments reportParaments = new ReportParaments();
                    #region 判斷哪間檢測單位
                    int testingOrganizationNum = -1;
                    ProcessFirstPageOfReport processFirstPageOfReport = new ProcessFirstPageOfReport();
                    testingOrganizationNum = processFirstPageOfReport.processFirstPageOfReport(pdfPath, reportParaments);
                    #endregion
                    if (testingOrganizationNum == 0)
                    {
                        #region TW SGS報告參數提取
                        TwSgsReport twSgsReport = new TwSgsReport();
                        await twSgsReport.detectionReportParament(pdfPath, reportParaments);
                        #endregion

                        #region  TW SGS報告 表格提取程式
                        //傳遞給API
                        string exePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PDF2XML.exe");
                        string outputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "xml"); // 固定輸出路徑為跟目錄xml資料夾
                        ProcessStartInfo startInfo = new ProcessStartInfo
                        {
                            FileName = exePath,
                            // 確保將路徑傳給外部 EXE
                            Arguments = $"\"{pdfPath}\" \"{outputPath}\"",
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            RedirectStandardOutput = true,
                            RedirectStandardError = true
                        };
                        #endregion

                        try
                        {
                            using (Process process = new Process { StartInfo = startInfo })
                            {

                                process.Start();
                                await Task.Run(() => process.WaitForExit());
                                int pdf2XmlExitCode = process.ExitCode;

                                if (process.ExitCode == 0 || process.ExitCode == 1)
                                {
                                    #region 讀取xml並賦值給 ReportManager
                                    string xmlPath = "xml/pdf2XmlResult.xml";

                                    //載入並檢查 XML 先6再5再4、3欄位的 TableGroup
                                    XDocument pdf2XmlResult = XDocument.Load(xmlPath);
                                    #endregion

                                    #region  TW SGS報告 將XML表格資料賦值給 ReportRowDataManager
                                    var table3Group = pdf2XmlResult.Descendants("TableGroup_3Cols").FirstOrDefault();
                                    var table4Group = pdf2XmlResult.Descendants("TableGroup_4Cols").FirstOrDefault();
                                    var table5Group = pdf2XmlResult.Descendants("TableGroup_5Cols").FirstOrDefault();
                                    var table6Group = pdf2XmlResult.Descendants("TableGroup_6Cols").FirstOrDefault();

                                    //遍歷 XML 中的每一列 Row
                                    if (table6Group != null)
                                    {
                                        var xmlRows = table6Group.Elements("Row");
                                        foreach (var row in xmlRows)
                                        {
                                            string xmlValue = row.Element("col_0")?.Value?.Trim() ?? "";
                                            var matchedItem = reportRowData.reportRowData.FirstOrDefault(d =>
                                                xmlValue.IndexOf(d.TestItem, StringComparison.OrdinalIgnoreCase) >= 0);
                                            var matchedForAllReportRowDatas = combineAllReportRowData.reportRowData.FirstOrDefault(d =>
                                               xmlValue.IndexOf(d.TestItem, StringComparison.OrdinalIgnoreCase) >= 0);

                                            if (matchedItem != null)
                                            {
                                                // 5. 更新該筆物件的值，TestItem 維持原先建構子給的標準名稱
                                                matchedItem.Method = row.Element("col_1")?.Value?.Trim() ?? "---";
                                                matchedItem.Unit = row.Element("col_2")?.Value?.Trim() ?? "---";
                                                matchedItem.MDL = row.Element("col_3")?.Value?.Trim() ?? "---";
                                                matchedItem.Result = row.Element("col_4")?.Value?.Trim() ?? "---";
                                            }
                                            if (matchedForAllReportRowDatas != null)
                                            {
                                                // 5. 更新該筆物件的值，TestItem 維持原先建構子給的標準名稱
                                                matchedForAllReportRowDatas.Method = row.Element("col_1")?.Value?.Trim() ?? "---";
                                                matchedForAllReportRowDatas.Unit = row.Element("col_2")?.Value?.Trim() ?? "---";
                                                matchedForAllReportRowDatas.MDL = row.Element("col_3")?.Value?.Trim() ?? "---";
                                                matchedForAllReportRowDatas.Result = row.Element("col_4")?.Value?.Trim() ?? "---";
                                            }
                                        }
                                        //reportRowData.showDebug();
                                    }
                                    else if (table5Group != null)
                                    {
                                        var xmlRows = table5Group.Elements("Row");
                                        foreach (var row in xmlRows)
                                        {
                                            string xmlValue = row.Element("col_0")?.Value?.Trim() ?? "";
                                            // 4. 比對：在 manager 已經建立好的 36 筆資料中找匹配項
                                            // 這裡 matchedItem 是對 manager.reportRowData 裡面某個物件的引用
                                            var matchedItem = reportRowData.reportRowData.FirstOrDefault(d =>
                                                xmlValue.IndexOf(d.TestItem, StringComparison.OrdinalIgnoreCase) >= 0);
                                            var matchedForAllReportRowDatas = combineAllReportRowData.reportRowData.FirstOrDefault(d =>
                                                xmlValue.IndexOf(d.TestItem, StringComparison.OrdinalIgnoreCase) >= 0);

                                            if (matchedItem != null)
                                            {
                                                // 5. 更新該筆物件的值，TestItem 維持原先建構子給的標準名稱
                                                matchedItem.Method = row.Element("col_1")?.Value?.Trim() ?? "---";
                                                matchedItem.Unit = row.Element("col_2")?.Value?.Trim() ?? "---";
                                                matchedItem.MDL = row.Element("col_3")?.Value?.Trim() ?? "---";
                                                matchedItem.Result = row.Element("col_4")?.Value?.Trim() ?? "---";
                                            }

                                            if (matchedForAllReportRowDatas != null)
                                            {
                                                // 5. 更新該筆物件的值，TestItem 維持原先建構子給的標準名稱
                                                matchedForAllReportRowDatas.Method = row.Element("col_1")?.Value?.Trim() ?? "---";
                                                matchedForAllReportRowDatas.Unit = row.Element("col_2")?.Value?.Trim() ?? "---";
                                                matchedForAllReportRowDatas.MDL = row.Element("col_3")?.Value?.Trim() ?? "---";
                                                matchedForAllReportRowDatas.Result = row.Element("col_4")?.Value?.Trim() ?? "---";
                                            }
                                        }
                                    }
                                    else if (table4Group != null && table3Group != null)
                                    {
                                        string pdfReportFilePage2String = "";
                                        string testMethod = "";
                                        string testItemPFAS = "PFAS";
                                        string testItemPFASResult = "";

                                        using (StreamReader reader = new StreamReader("pdfReportFilePage2.txt", Encoding.UTF8))
                                        {
                                            pdfReportFilePage2String = reader.ReadToEnd();
                                        }
                                        pdfReportFilePage2StringSplitArray = Regex.Split(pdfReportFilePage2String, "\r\n");
                                        for (int i = 0; i < pdfReportFilePage2StringSplitArray.Length; i++)
                                        {
                                            string rawLine = pdfReportFilePage2StringSplitArray[i];
                                            if (string.IsNullOrWhiteSpace(rawLine)) continue;

                                            string cleanLine = rawLine.Replace(" ", "");

                                            if (cleanLine.Contains("EN17681-1") && cleanLine.Contains("EN17681-2"))
                                            {
                                                testMethod = "EN17681-1 & EN17681-2";
                                                var matchedItem = reportRowData.reportRowData.FirstOrDefault(d => d.TestItem.Equals("PFAS", StringComparison.OrdinalIgnoreCase));
                                                var matchedForAllReportRowDatas = combineAllReportRowData.reportRowData.FirstOrDefault(d => d.TestItem.Equals("PFAS", StringComparison.OrdinalIgnoreCase));
                                                matchedItem.Method = testMethod;
                                                matchedForAllReportRowDatas.Method = testMethod;
                                                break;
                                            }
                                            else if (cleanLine.Contains("EN17681-2"))
                                            {
                                                testMethod = "EN17681-2";
                                                var matchedItem = reportRowData.reportRowData.FirstOrDefault(d => d.TestItem.Equals("PFAS", StringComparison.OrdinalIgnoreCase));
                                                var matchedForAllReportRowDatas = combineAllReportRowData.reportRowData.FirstOrDefault(d => d.TestItem.Equals("PFAS", StringComparison.OrdinalIgnoreCase));
                                                matchedItem.Method = testMethod;
                                                matchedForAllReportRowDatas.Method = testMethod;
                                                break;
                                            }
                                            else if (cleanLine.Contains("EN17681-1"))
                                            {
                                                testMethod = "EN17681-1";
                                                var matchedItem = reportRowData.reportRowData.FirstOrDefault(d => d.TestItem.Equals("PFAS", StringComparison.OrdinalIgnoreCase));
                                                var matchedForAllReportRowDatas = combineAllReportRowData.reportRowData.FirstOrDefault(d => d.TestItem.Equals("PFAS", StringComparison.OrdinalIgnoreCase));
                                                matchedItem.Method = testMethod;
                                                matchedForAllReportRowDatas.Method = testMethod;
                                                break;
                                            }
                                            else
                                            {
                                                testMethod = "---";
                                                var matchedItem = reportRowData.reportRowData.FirstOrDefault(d => d.TestItem.Equals("PFAS", StringComparison.OrdinalIgnoreCase));
                                                var matchedForAllReportRowDatas = combineAllReportRowData.reportRowData.FirstOrDefault(d => d.TestItem.Equals("PFAS", StringComparison.OrdinalIgnoreCase));
                                                matchedItem.Method = testMethod;
                                                matchedForAllReportRowDatas.Method = testMethod;
                                            }
                                        }

                                        var xmlRows_3 = table3Group.Elements("Row");
                                        var xmlRows_4 = table4Group.Elements("Row");
                                        foreach (var row in xmlRows_3)
                                        {
                                            string xmlValue = row.Element("col_0")?.Value?.Trim() ?? "";
                                            var matchedItem = reportRowData.reportRowData.FirstOrDefault(d =>
                                                xmlValue.IndexOf(d.TestItem, StringComparison.OrdinalIgnoreCase) >= 0);
                                            var matchedForAllReportRowDatas = combineAllReportRowData.reportRowData.FirstOrDefault(d =>
                                                xmlValue.IndexOf(d.TestItem, StringComparison.OrdinalIgnoreCase) >= 0);

                                            if (matchedItem != null)
                                            {
                                                matchedItem.Result = row.Element("col_2")?.Value?.Trim() ?? "---";
                                                testItemPFASResult = row.Element("col_2")?.Value?.Trim() ?? "---";
                                            }
                                            if (matchedForAllReportRowDatas != null)
                                            {
                                                matchedForAllReportRowDatas.Result = row.Element("col_2")?.Value?.Trim() ?? "---";
                                                testItemPFASResult = row.Element("col_2")?.Value?.Trim() ?? "---";
                                            }
                                        }
                                        foreach (var row in xmlRows_4)
                                        {
                                            string xmlValue = row.Element("col_1")?.Value?.Trim() ?? "";
                                            var matchedItem = reportRowData.reportRowData.FirstOrDefault(d => d.TestItem.Equals("(PFOS)", StringComparison.OrdinalIgnoreCase));
                                            var matchedForAllReportRowDatas = combineAllReportRowData.reportRowData.FirstOrDefault(d => d.TestItem.Equals("(PFOS)", StringComparison.OrdinalIgnoreCase));
                                            if (matchedItem != null)
                                            {
                                                matchedItem.Result = testItemPFASResult;
                                                matchedItem.Method = testMethod;
                                                matchedItem.MDL = row.Element("col_3")?.Value?.Trim() ?? "---";
                                            }
                                            if (matchedForAllReportRowDatas != null)
                                            {
                                                matchedForAllReportRowDatas.Result = testItemPFASResult;
                                                matchedForAllReportRowDatas.Method = testMethod;
                                                matchedForAllReportRowDatas.MDL = row.Element("col_3")?.Value?.Trim() ?? "---";
                                            }
                                            matchedItem = reportRowData.reportRowData.FirstOrDefault(d => d.TestItem.Equals("(PFOA)", StringComparison.OrdinalIgnoreCase));
                                            matchedForAllReportRowDatas = combineAllReportRowData.reportRowData.FirstOrDefault(d => d.TestItem.Equals("(PFOA)", StringComparison.OrdinalIgnoreCase));
                                            if (matchedItem != null)
                                            {
                                                matchedItem.Result = testItemPFASResult;
                                                matchedItem.Method = testMethod;
                                                matchedItem.MDL = row.Element("col_3")?.Value?.Trim() ?? "---";
                                            }
                                            if (matchedForAllReportRowDatas != null)
                                            {
                                                matchedForAllReportRowDatas.Result = testItemPFASResult;
                                                matchedForAllReportRowDatas.Method = testMethod;
                                                matchedForAllReportRowDatas.MDL = row.Element("col_3")?.Value?.Trim() ?? "---";
                                            }
                                        }
                                        reportRowData.showDebug();
                                    }
                                    else
                                    {
                                        throw new Exception("No suitable TableGroup found in XML!");
                                    }



                                    #endregion

                                    #region 報告審核
                                    VerificationReport verificationReport = new VerificationReport();
                                    List<string> singleResult = new List<string>();
                                    singleResult = verificationReport.verificationReport(verificationScenarioIndex, metalNonmetalIndex, reportRowData, reportParaments);
                                    list2Excel.Add(singleResult);
                                    #endregion

                                }
                                else
                                {
                                    string error = await process.StandardError.ReadToEndAsync();
                                    MessageBox.Show($"pdf 轉換失敗: {error}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"執行pdf表格提取發生錯誤: {ex.Message}");
                        }


                        #region 測試用參數顯示
                        /*
                        MessageBox.Show(
                             "公司名稱:" + TwSgsReport.company +
                             "\r\n樣品名稱:" + TwSgsReport.sampleName +
                             "\r\n樣品型號:" + TwSgsReport.styleItemNo +
                             "\r\n測試期間:" + TwSgsReport.finalExtractedDate +
                             "\r\n測試結果備註:" + TwSgsReport.testPeriodResult +
                             "\r\n測試部位敘述:" + TwSgsReport.testParts
                         );
                        */
                        #endregion

                    }

                    else if (testingOrganizationNum == 1)
                    {
                    }
                    else if (testingOrganizationNum == 2)
                    {
                    }
                    else if (testingOrganizationNum == 3)
                    {
                    }
                    else if (testingOrganizationNum == 4)
                    {
                    }
                    else if (testingOrganizationNum == 5)
                    {
                    }
                    else if (testingOrganizationNum == 6)
                    {


                    }
                    else
                    {
                        MessageBox.Show("無法識別的檢測單位，請確認報告格式是否正確。");
                    }                  

                }
                combineAllReportRowData.showDebug();
                VerificationReport verificationAllReport = new VerificationReport();
                List<string> batchResult = new List<string>();
                batchResult = verificationAllReport.verificationReport(verificationScenarioIndex, metalNonmetalIndex, combineAllReportRowData);
                /*For testing
                foreach (string result in batchResult)
                {
                    Console.WriteLine(result);
                }*/
                ProcessAndDisplayResults(batchResult);
                list2Excel.Add(batchResult);

                // 1. 準備資料與路徑 (避免非法字元)
                string excelDateString = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                

                // 取得程式執行的目錄，並建立 batchResult 資料夾
                string folderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "batchResult");
                string excelFileName = Path.GetFileName(dirPath);
                if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);

                // 最終完整的檔案路徑
                string fullPath = Path.Combine(folderPath, $"{excelFileName} - {excelDateString}.xlsx");

                // 2. 開始建立 Excel 檔案
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(fullPath, SpreadsheetDocumentType.Workbook))
                {
                    // 建立活頁簿結構
                    WorkbookPart workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

                    for (int i = 0; i < list2Excel.Count; i++)
                    {
                        uint sheetId = (uint)(i + 1);
                        string sheetName;
                        if (i == list2Excel.Count - 1)
                        {
                            sheetName = "batchResult";
                        }
                        else
                        {
                            sheetName = $"pdf{(i + 1):D2}";
                        }
                        AddWorksheet(workbookPart, sheets, sheetId, sheetName, list2Excel[i]);
                    }

                    
                    workbookPart.Workbook.Save();
                }
                btnStartBatchVerification.Enabled = true;
            }


        }
            
        


        private void cbVerificationScenario_SelectedIndexChanged(object sender, EventArgs e)
        {
            verificationScenarioIndex = cbVerificationScenario.SelectedIndex;
            if (verificationScenarioIndex != -1 && metalNonmetalIndex != -1)
            {
                btnStartBatchVerification.Enabled = true;
            }
            else
            {
                btnStartBatchVerification.Enabled = false;
            }
        }

        private void cbMetalNonmetal_SelectedIndexChanged(object sender, EventArgs e)
        {
            metalNonmetalIndex = cbMetalNonmetal.SelectedIndex;
            if (verificationScenarioIndex != -1 && metalNonmetalIndex != -1)
            {
                btnStartBatchVerification.Enabled = true;
            }
            else
            {
                btnStartBatchVerification.Enabled = false;
            }
        }

        public void ProcessAndDisplayResults(List<string> resultList)
        {
            tbVerificationResult.Clear();
            tbVerificationResult.Font = new System.Drawing.Font("Microsoft JhengHei", 10);

            foreach (string line in resultList)
            {

                Match match = Regex.Match(line.Trim(), @"^([^:]+):(.*)$");
                if (match.Success)
                {
                    string elementName = match.Groups[1].Value;
                    string content = match.Groups[2].Value;

                    System.Drawing.Color contentColor = System.Drawing.Color.Black;
                    if (content.Contains("未檢測") ||
                        content.Contains("超標") ||
                        content.Contains("未經矽品認可") ||
                        content.Contains("超出矽品認可"))
                    {
                        contentColor = System.Drawing.Color.Red;
                    }

                    AppendTextToRichTextBox(elementName + ":", System.Drawing.Color.Blue);
                    AppendTextToRichTextBox(content + Environment.NewLine, contentColor);
                }
                else
                {
                    AppendTextToRichTextBox(line + Environment.NewLine, System.Drawing.Color.Black);
                }
            }
        }

        // 輔助方法：處理 RichTextBox 著色
        private void AppendTextToRichTextBox(string text, System.Drawing.Color color)
        {
            tbVerificationResult.SelectionStart = tbVerificationResult.TextLength;
            tbVerificationResult.SelectionLength = 0;
            tbVerificationResult.SelectionColor = color;
            tbVerificationResult.AppendText(text);
            tbVerificationResult.SelectionColor = tbVerificationResult.ForeColor; // 恢復預設
        }

        private void AddWorksheet(WorkbookPart workbookPart, Sheets sheets, uint sheetId, string sheetName, List<string> data)
        {
            // 建立新分頁的 Part
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            // 將分頁註冊到 Sheets 清單中
            Sheet sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(newWorksheetPart),
                SheetId = sheetId,
                Name = sheetName
            };
            sheets.Append(sheet);

            // 寫入資料
            SheetData sheetData = newWorksheetPart.Worksheet.GetFirstChild<SheetData>();
            foreach (string text in data)
            {
                Row row = new Row();
                row.Append(new Cell { DataType = CellValues.String, CellValue = new CellValue(text) });
                sheetData.Append(row);
            }
        }

    }
}
