using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static AutomatedVerificationProgram.OutdatedCoordinatesSetting;
using Spire.Xls;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System.Threading;
using DocumentFormat.OpenXml.Spreadsheet;
using Spire.Pdf.Graphics;
using System.Diagnostics;
using System.Xml.Linq;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Ink;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office.Word;

namespace AutomatedVerificationProgram
{
    public  partial class AutomatedBatchVerificationPage : UserControl
    {
        #region 自動化座標參數儲存區
        public class ClickStepCoordinates
        {
            public string ActionType { get; set; }
            public string X { get; set; }
            public string Y { get; set; }
            public string Param { get; set; }
        }
        #region 過期報告驗證的五段點擊座標參數
        List<ClickStepCoordinates> outdated1stClickSteps;
        List<ClickStepCoordinates> outdated2ndClickSteps;
        List<ClickStepCoordinates> outdated3rdClickSteps;
        List<ClickStepCoordinates> outdated4thClickSteps;
        List<ClickStepCoordinates> outdated5thClickSteps;
        List<ClickStepCoordinates> outdated6thClickSteps;
        List<ClickStepCoordinates> outdated7thClickSteps;
        List<ClickStepCoordinates> outdated8thClickSteps;
        #endregion
        #region mcd報告驗證滑鼠座標
        List<ClickStepCoordinates> mcd1stClickSteps;
        List<ClickStepCoordinates> mcd2ndClickSteps;
        List<ClickStepCoordinates> mcd3thClickSteps;
        List<ClickStepCoordinates> mcd4thClickSteps;
        List<ClickStepCoordinates> mcd5thClickSteps;
        List<ClickStepCoordinates> mcd6thClickSteps;
        List<ClickStepCoordinates> mcd7thClickSteps;
        #endregion


        #endregion

        #region OCR截圖參數
        public class OcrParameterClass
        {
            public int X1 { get; set; }
            public int X2 { get; set; }
            public int Y1 { get; set; }
            public int Y2 { get; set; }
        }
        List<OcrParameterClass> OcrParameters;
        int[] ocrParametersArray = new int[4]; //0:X1 1:X2 2:Y1 3:Y2
        #endregion


        #region 下載區參數
        public class FileTimeInfo
        {
            public string FileName;  //檔名
            public DateTime FileCreateTime; //建立時間
        }
        static FileTimeInfo GetLatestFileTimeInfo(string dir, string ext)
        {
            List<FileTimeInfo> list = new List<FileTimeInfo>();
            DirectoryInfo d = new DirectoryInfo(dir);
            foreach (FileInfo file in d.GetFiles())
            {
                if (file.Extension.ToUpper() == ext.ToUpper())
                {
                    list.Add(new FileTimeInfo()
                    {
                        FileName = file.FullName,
                        FileCreateTime = file.CreationTime
                    });
                }
            }
            var f = from x in list
                    orderby x.FileCreateTime
                    select x;
            return f.LastOrDefault();
        }
        string downloads_path = Environment.GetEnvironmentVariable("USERPROFILE") + @"\" + "Downloads";
        #endregion

        #region 轉人審區塊
        string converToHumanSurvey = ""; //已轉人工審查
        #endregion
        int indexOfVerificationMode = 0; // 0: 過期報告優先, 1: MCD優先, 2: 僅過期報告, 3: 僅MCD       
        public static string[] pdfReportFilePage2StringSplitArray;

        List<string> reasonsOfRehection  = new List<string>();

        public AutomatedBatchVerificationPage()
        {
            InitializeComponent();
            string[] paragraphFilesPaths = new string[]
            {
                @"Coordinates/outdatedReportFirstParagraphCoordinatesSetting.json",
                @"Coordinates/outdatedReportSecondParagraphCoordinatesSetting.json",
                @"Coordinates/outdatedReportThirdParagraphCoordinatesSetting.json",
                @"Coordinates/outdatedReportForthParagraphCoordinatesSetting.json",
                @"Coordinates/outdatedReportFivethParagraphCoordinatesSetting.json",
                @"Coordinates/outdatedReportSixthParagraphCoordinatesSetting.json",
                @"Coordinates/outdatedReportSeventhParagraphCoordinatesSetting.json",
                @"Coordinates/outdatedReportEighthParagraphCoordinatesSetting.json",
                @"Coordinates/mcd1stParagraphCoordinatesSetting.json",
                @"Coordinates/mcd2ndParagraphCoordinatesSetting.json",
                @"Coordinates/mcd3thParagraphCoordinatesSetting.json",
                @"Coordinates/mcd4thParagraphCoordinatesSetting.json",
                @"Coordinates/mcd5thParagraphCoordinatesSetting.json",
                @"Coordinates/mcd6thParagraphCoordinatesSetting.json",
                @"Coordinates/mcd7thParagraphCoordinatesSetting.json"
            };
          
            outdated1stClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[0]));
            outdated2ndClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[1]));
            outdated3rdClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[2]));
            outdated4thClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[3]));
            outdated5thClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[4]));
            outdated6thClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[5]));
            outdated7thClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[6]));
            
            mcd1stClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[7]));
            mcd2ndClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[8]));
            mcd3thClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[9]));
            mcd4thClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[10]));
            mcd5thClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[11]));
            mcd6thClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[12]));
            mcd7thClickSteps = JsonConvert.DeserializeObject<List<ClickStepCoordinates>>(File.ReadAllText(paragraphFilesPaths[13]));
            OcrParameters = JsonConvert.DeserializeObject<List<OcrParameterClass>>(File.ReadAllText(@"Coordinates/ocrCoordinatesSetting.json"));
            ocrParametersArray[0] = OcrParameters[0].X1;
            ocrParametersArray[1] = OcrParameters[0].X2;
            ocrParametersArray[2] = OcrParameters[0].Y1;
            ocrParametersArray[3] = OcrParameters[0].Y1;
        }

        private void AutomatedBatchVerificationPage_Load(object sender, EventArgs e)
        {
            btnAutomatedBatchVerificationStart.Enabled = false; //需先選擇驗證模式
            btnHomogeneousCustomVerificationSequence.Enabled = false; //do it in the future
            btnMcdCustomVerificationSequence.Enabled = false; //do it in the future           
        }

        private async void btnAutomatedBatchVerificationStart_Click(object sender, EventArgs e)
        {
            // 過期報告優先
            if (indexOfVerificationMode == 0)
            {
                
                
            }
            // MCD優先
            else if (indexOfVerificationMode == 1)
            {
               
               
            }
            // 僅過期報告
            else if (indexOfVerificationMode == 2)
            {
                #region 過期報告_log 初始化
                int reportLogIndex = 0; // 報告紀錄的行數索引，從1開始（第一行是標題）
                // 1. 準備路徑與資料夾
                string folderPath = "./過期報告_log";
                if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);
                string nowDayTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fullPath = Path.Combine(folderPath, $"{nowDayTime}.xlsx");
                // 2. 【迴圈外】建立檔案基礎結構
                SpreadsheetDocument outDatedReportLog = SpreadsheetDocument.Create(fullPath, SpreadsheetDocumentType.Workbook);
                WorkbookPart workbookPart = outDatedReportLog.AddWorkbookPart();
                workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                // 建立一個分頁與其資料容器 (SheetData)
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData(); // 這是我們要在迴圈內持續操作的對象
                worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                Sheet sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "LogReport"
                };
                sheets.Append(sheet);
                // --- 標題初始化 (寫入第一列) --
                string[] headers = { "審核時間", "供應商代碼", "供應商名稱", 
                                     "均勻材質編號", "審核結果", "退件原因", 
                                     "Pb","Cd","Hg","Cr(VI)","PBB","PBDE","DBP","BBP","DIBP",
                                     "DEHP","Sb","PFOS","PFOA","PVC","Pb","報告號碼","報告日期","檢測方法正確",
                };

                Row headerRow = new Row() { RowIndex = (uint)reportLogIndex };
                for (int i = 0; i < headers.Length; i++)
                {
                    Cell cell = new Cell()
                    {
                        CellReference = GetColumnName(i) + reportLogIndex,
                        DataType = CellValues.String,
                        CellValue = new CellValue(headers[i])
                    };
                    headerRow.AppendChild(cell);
                }
                sheetData.AppendChild(headerRow);

                #endregion

                List<List<string>> humanSurveyReport = new List<List<string>>(); //轉人工審查的報告
                int numberOfHumanSurvey = 0;    
                int currentVertificationNumber = 0;              
                int numberOfBlackList = 0;
                int pbContain = 0;//含鉛材
                int halogenContain = 0;//含鹵材                

                #region 報告更新特定供應商/vendor_black list.xlsx
                string outdatedVendorBlackListPath = "./報告更新特定供應商/vendor_black list.xlsx";
                Spire.Xls.Workbook outdatedVendorBlackListExcel = new Spire.Xls.Workbook();
                outdatedVendorBlackListExcel.LoadFromFile(outdatedVendorBlackListPath);
                Spire.Xls.Worksheet outdatedVendorBlackListExcelSheet = outdatedVendorBlackListExcel.Worksheets[0];
                numberOfBlackList = outdatedVendorBlackListExcelSheet.Rows.Length - 1;
                string[] outdatedVendorBlackListVendorCode = new string[outdatedVendorBlackListExcelSheet.Rows.Length];
                string[] outdatedVendorBlackListItems = new string[outdatedVendorBlackListExcelSheet.Rows.Length];
                for (int i = 1; i < outdatedVendorBlackListExcelSheet.Rows.Length; i++)
                {
                    outdatedVendorBlackListVendorCode[i] = outdatedVendorBlackListExcelSheet.Rows[i].Columns[0].Text.ToUpper();

                    if (outdatedVendorBlackListExcelSheet.Rows[i].Columns[1].Text != null)
                    {
                        outdatedVendorBlackListItems[i] = outdatedVendorBlackListExcelSheet.Rows[i].Columns[1].Text.ToUpper();
                    }
                }
                #endregion

               

                while(true)
                {
                    #region paragraph1
                    foreach (var step in outdated1stClickSteps)
                    {
                        string action = step.ActionType ?? string.Empty;
                        int x = int.TryParse(step.X, out int ix) ? ix : 0;
                        int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                        int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                        if (action == "等待時間")
                        {
                            ClickFnctionClass.MouseWin32.Click(action, param);
                        }
                        else
                        {
                            ClickFnctionClass.MouseWin32.Click(action, x, y);
                        }
                    }
                    #endregion
                    #region 處理第一份excel
                    FileTimeInfo excelFilePath_1 = GetLatestFileTimeInfo(downloads_path, ".xls"); //獲取下載的最新Excel GetLatestFileTimeInfo(downloads_path, ".xls")
                    Spire.Xls.Workbook excel_1 = new Spire.Xls.Workbook();  //第一個Excel 物件叫做excel_1
                    excel_1.LoadFromFile(excelFilePath_1.FileName);
                    Spire.Xls.Worksheet excel_1_sheet = excel_1.Worksheets[0]; //第一頁工作表 物件叫做excel_1_sheet
                    int numberOfVendor = excel_1_sheet.Rows.Length - 1;
                    if(numberOfVendor == 0)
                    {
                        break;
                    }
                    tbOngoingResult.Text += "供應商數量：" + numberOfVendor.ToString() + "\r\n";
                    int[] listOfNeedToVerifiction = new int[numberOfVendor];
                    for (int i = 0; i < listOfNeedToVerifiction.Length; i++)
                    {
                        listOfNeedToVerifiction[i] = -1;
                    }
                    string[,] waitForVertify = new string[numberOfVendor,2];

                    for (int i = 1; i < numberOfVendor; i++)
                    { 
                        waitForVertify[i,0] = excel_1_sheet.Rows[i].Columns[1].Text; //供應商代號
                        waitForVertify[i, 1] = excel_1_sheet.Rows[i].Columns[3].Text; //均勻材質編號
                    }
                    #endregion

                    #region 先判斷是否為黑名單
                    for (int i = 0; i < numberOfVendor; i++)
                    {
                        string currentVendor = waitForVertify[i, 0]; // 供應商代號
                        string currentMaterial = waitForVertify[i, 1]; // 均勻材質編號
                        for (int j = 0; j < numberOfBlackList; j++)
                        {
                            if ((currentVendor == outdatedVendorBlackListVendorCode[j]) && (currentMaterial == outdatedVendorBlackListItems[j])) //若供應商代號在黑名單且無特定材質則全轉人工審查
                            {
                                humanSurveyReport.Add(new List<string> { waitForVertify[i, 0], waitForVertify[i, 1] }); //紀錄轉人工審查的報告
                                numberOfHumanSurvey += 1; //紀錄人工審查報告數量
                                listOfNeedToVerifiction[currentVertificationNumber] = 1;
                                break;
                            }
                        }
                        currentVertificationNumber++;
                    }
                    currentVertificationNumber = 0;
                    #endregion

                    #region 判斷是否還有需機器審查的報告，若無則結束驗證
                    for (int i = humanSurveyReport.Count - 1; i >= 0; i--)
                    {
                        bool isExist = false;
                        string currentVendorId = humanSurveyReport[i][0]; // 取得清單中的廠商 ID
                        string currentItemNumber = humanSurveyReport[i][0]; // 取得清單中的料號

                        for (int j = 0; j < numberOfVendor; j++)
                        {
                            if (currentVendorId == waitForVertify[j, 0] && currentItemNumber == waitForVertify[j, 1])
                            {
                                listOfNeedToVerifiction[j] = 1; // 標記該供應商報告為需人工審查
                                isExist = true;
                                break;
                            }
                        }
                        if (!isExist)
                        {
                            humanSurveyReport.RemoveAt(i);
                            numberOfHumanSurvey--;
                        }
                    }
                    if (numberOfHumanSurvey == numberOfVendor || numberOfVendor == 0) //若剩下全轉人工審查
                    {
                        break;
                        //MessageBox.Show("本次驗證已完成，全部供應商報告已轉人工審查。", "驗證完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    #endregion

                    for(int i = 0; i < listOfNeedToVerifiction.Length; i++)
                    {
                        if (listOfNeedToVerifiction[currentVertificationNumber] == 1)
                        {
                            currentVertificationNumber++;
                        }
                        else if(listOfNeedToVerifiction[currentVertificationNumber] == -1)
                        {
                            break;
                        }
                    }

                    lbOngingItem.Text = $"{waitForVertify[currentVertificationNumber, 0]} : {waitForVertify[currentVertificationNumber, 1]}";
                    //輸入供應商代號+均勻材質編號 尋找唯一報告
                    #region paragraph2
                    int pastedTimes = 0;
                    foreach (var step in outdated2ndClickSteps)
                    {
                        string action = step.ActionType ?? string.Empty;
                        int x = int.TryParse(step.X, out int ix) ? ix : 0;
                        int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                        int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                        if (action == "等待時間")
                        {
                            ClickFnctionClass.MouseWin32.Click(action, param);
                        }
                        else if (action == "貼上文字")
                        {
                            Clipboard.SetText(waitForVertify[currentVertificationNumber, pastedTimes]);
                            ClickFnctionClass.MouseWin32.Click(action);
                            pastedTimes++;
                        }
                        else if(action == "全選文字")
                        {
                            ClickFnctionClass.MouseWin32.Click(action);
                        }
                        else
                        {
                            ClickFnctionClass.MouseWin32.Click(action, x, y);
                        }
                    }
                    #endregion

                   

                    #region paragraph3
                    foreach (var step in outdated3rdClickSteps)
                    {
                        string action = step.ActionType ?? string.Empty;
                        int x = int.TryParse(step.X, out int ix) ? ix : 0;
                        int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                        int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                        if (action == "等待時間")
                        {
                            ClickFnctionClass.MouseWin32.Click(action, param);
                        }
                        else
                        {
                            ClickFnctionClass.MouseWin32.Click(action, x, y);
                        }
                    }
                    #endregion

                    #region 開啟第二 份Excel
                    //判斷XX-情境1-------------------------------------------------------------
                    FileTimeInfo excelFilePath_2 = GetLatestFileTimeInfo(downloads_path, ".xls"); //抓取最新下載的excel  workbook物件為第二份Excel
                    Spire.Xls.Workbook excel_2 = new Spire.Xls.Workbook();
                    excel_2.LoadFromFile(excelFilePath_2.FileName);
                    Spire.Xls.Worksheet excel_2_sheet = excel_2.Worksheets[0];  //第二份Excel的第一張工作表
                    string materialNumber = excel_2_sheet.Rows[1].Columns[0].Text;
                    string materialDescription = excel_2_sheet.Rows[1].Columns[1].Text;
                    string homogeneousType = "";
                    #endregion
                    #region 初判情境
                    //1-a 1-b 2 3-a 3-b 0~4
                    int[] tentativeScenariosResult = new int[5];
                    for(int i = 0; i < tentativeScenariosResult.Length; i++)
                    {
                        tentativeScenariosResult[i] = -1;
                    }
                    // 1. 判斷 A 欄是否只有 「5」 或者 「44-」 開頭
                    // 正則解釋：^ 代表開頭，(5|44-) 代表 5 或 44-
                    if (Regex.IsMatch(materialNumber, @"^(5|44-)"))
                    {
                        tentativeScenariosResult[3] = 1;
                        tentativeScenariosResult[4] = 1;
                    }
                    // 2. 判斷 A 欄是否只有 「XX-」 開頭 (X = 數字)
                    // 正則解釋：^\d{2}- 代表開頭為兩位數字加上橫槓
                    else if (Regex.IsMatch(materialNumber, @"^\d{2}-"))
                    {
                        tentativeScenariosResult[2] = 1;
                    }
                    // 3. 如果以上皆非，則為 1a or 1b
                    else
                    {
                        tentativeScenariosResult[0] = 1;
                        tentativeScenariosResult[1] = 1;
                    }
                    #endregion

                    #region 含鉛材含鹵材
                    // 4. 判斷 B 欄是否包含關鍵字 (含鉛材)
                    if (materialDescription.Contains("_HS-PB") || materialDescription.Contains("-HS-PB"))
                    {
                        pbContain = 1; //含鉛材
                    }
                    // 5. 判斷 B 欄是否包含關鍵字 (含鹵材)
                    if (materialDescription.Contains("HS-NON-HF"))
                    {
                        halogenContain = 1;//含鹵材
                    }
                    #endregion
                    int[] scenarioArray = new int[2];
                    for(int i = 0; i < scenarioArray.Length; i++)
                    {
                        scenarioArray[i] = -1;
                    }

                    homogeneousType = Ocr.OcrString(ocrParametersArray[0], ocrParametersArray[1], ocrParametersArray[2], ocrParametersArray[3]);
                    scenarioArray = Ocr.FindScenario(tentativeScenariosResult, ocrParametersArray[0], ocrParametersArray[1], ocrParametersArray[2], ocrParametersArray[3]);

                    if (scenarioArray[0] == -1 || scenarioArray[1] == -1)
                    {
                        humanSurveyReport.Add(new List<string> { waitForVertify[currentVertificationNumber, 0], waitForVertify[currentVertificationNumber, 1] }); //紀錄轉人工審查的報告
                        continue;
                    }

                    #region paragraph4 下載第三份Excel並點擊相關座標
                    foreach (var step in outdated4thClickSteps)
                    {
                        string action = step.ActionType ?? string.Empty;
                        int x = int.TryParse(step.X, out int ix) ? ix : 0;
                        int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                        int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                        if (action == "等待時間")
                        {
                            ClickFnctionClass.MouseWin32.Click(action, param);
                        }
                        else
                        {
                            ClickFnctionClass.MouseWin32.Click(action, x, y);
                        }
                    }
                    #endregion

                    #region 第三份Excel 開啟  物件名稱：workbook_test
                    FileTimeInfo excelFilePath_3 = GetLatestFileTimeInfo(downloads_path, ".xls"); //第三份Excel workbook_test
                    Spire.Xls.Workbook excel_3 = new Spire.Xls.Workbook();
                    excel_3.LoadFromFile(excelFilePath_3.FileName);
                    Spire.Xls.Worksheet excel_3_sheet = excel_3.Worksheets[0];  //第三份Excel的第一張工作表
                    #endregion

                    int reportCount = 0;
                    string[] excel3DocumentType = new string[excel_3_sheet.Rows.Length];
                    string[] excel3ApprovalStatus = new string[excel_3_sheet.Rows.Length];
                    string[] excel3MdsTime = new string[excel_3_sheet.Rows.Length];

                    for (int i = 1; i < excel_3_sheet.Rows.Length; i++)
                    {
                        excel3DocumentType[i] = excel_3_sheet.Rows[i].Columns[2].Text;
                        excel3ApprovalStatus[i] = excel_3_sheet.Rows[i].Columns[1].Text;
                        excel3MdsTime[i] = excel_3_sheet.Rows[i].Columns[7].Text;
                        if (excel_3_sheet.Rows[i].Columns[2].Text == "Test report")
                        {
                            reportCount++;
                        }
                    }
                    
                    string[] excel3ReportNumber = new string[reportCount];
                    string[] excel3Date = new string[reportCount];
                    string[] excel3ThirdParty = new string[reportCount];
                    string[] excel3TestItem = new string[reportCount];

                    for(int i = 0; i < reportCount; i++)
                    {                       
                        excel3ReportNumber[i] = excel_3_sheet.Rows[i + 1].Columns[3].Text;
                        excel3Date[i] = excel_3_sheet.Rows[i + 1].Columns[4].Text;
                        excel3ThirdParty[i] = excel_3_sheet.Rows[i + 1].Columns[5].Text;
                        excel3TestItem[i] = excel_3_sheet.Rows[i + 1].Columns[6].Text;
                    }
                    ReportRowDataManager combineAllReportRowData = new ReportRowDataManager(); //儲存一個均勻材質所有報告的值
                    int bias = 0;
                    for (int i = 0; i < reportCount; i++)
                    {
                        #region paragraph5 下載pdf  
                        foreach (var step in outdated5thClickSteps)
                        {
                            string action = step.ActionType ?? string.Empty;
                            int x = int.TryParse(step.X, out int ix) ? ix : 0;
                            int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                            x += i * 20;
                            int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                            if (action == "等待時間")
                            {
                                ClickFnctionClass.MouseWin32.Click(action, param);
                            }
                            else if(action == "檔案偏移")
                            {
                                bias = param;
                            }
                            else
                            {
                                ClickFnctionClass.MouseWin32.Click(action, x , y + i * bias);//間隔沒設定
                            }
                        }
                        #endregion
                        FileTimeInfo pdfPath = GetLatestFileTimeInfo(downloads_path, ".pdf");
                        ReportRowDataManager reportRowData = new ReportRowDataManager();
                        ReportParaments reportParaments = new ReportParaments();
                        #region 判斷哪間檢測單位
                        int testingOrganizationNum = -1;
                        ProcessFirstPageOfReport processFirstPageOfReport = new ProcessFirstPageOfReport();
                        testingOrganizationNum = processFirstPageOfReport.processFirstPageOfReport(pdfPath.FileName, reportParaments);
                        #endregion
                        #region 判斷報告是否有數位簽章以及解析PDF2XML狀態 0 有 1 無 2 解析失敗
                        int pdf2XmlExitCode = -1;
                        #endregion
                        if (testingOrganizationNum == 0)
                        {
                            #region TW SGS報告參數提取
                            TwSgsReport twSgsReport = new TwSgsReport();
                            await twSgsReport.detectionReportParament(pdfPath.FileName, reportParaments);
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
                                    pdf2XmlExitCode = process.ExitCode; // 0有數位簽章 1無數位簽章 2解析失敗

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
                                            for (int j = 0; j < pdfReportFilePage2StringSplitArray.Length; j++)
                                            {
                                                string rawLine = pdfReportFilePage2StringSplitArray[j];
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

                                    }
                                    else
                                    {
                                        string error = await process.StandardError.ReadToEndAsync();
                                    }

                                }
                            }
                            catch (Exception ex)
                            {
                                
                            }

                        }
                        else if(testingOrganizationNum == 1)
                        {
                            //未來擴充其他檢測單位的報告處理程式
                        }
                        // To be continued 其他檢測單位的報告處理程式還沒好



                        if(pdf2XmlExitCode == 1)
                        {
                            reasonsOfRehection.Add( ReasonsOfRejectedDocument.GetRejectionReason("R-10"));
                        }
                        if (excel3ReportNumber[i] != reportParaments.reportNunber)
                        {
                            reasonsOfRehection.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-01"));
                        }

                        if(excel3Date[i] != reportParaments.finalExtractedDate)
                        {
                            reasonsOfRehection.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-02"));
                        }

                        if(testingOrganizationNum == 0) // 0表示TW SGS
                        {
                            if (excel3ThirdParty[i] != "SGS")
                            {
                                reasonsOfRehection.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-03"));
                            }
                        }

                        //To be continued 方法還沒想到留給你做
                        /*
                        else if (!testItemCheck)
                        {
                        }
                        */
                        if (homogeneousType != "Copper foil" || homogeneousType != "electrode" ||
                           homogeneousType != "Heat sink" || homogeneousType != "Lead Frame" ||
                           homogeneousType != "Plating/Coatings layer" || homogeneousType != "Ceramic/Glass" ||
                           homogeneousType != "Die/ Wafer" || homogeneousType != "Passives/Capacitor/Resistor/Connector" ||
                           homogeneousType != "Resistive layer" || homogeneousType != "Ceramic Substrate"
                         )
                        {
                            for (int j = 0; j < excel3DocumentType.Length; j++)
                            {
                                if (excel3DocumentType[i].Contains("MDS/SDS"))
                                {
                                    reasonsOfRehection.Add("MDS/SDS 不須上傳");
                                }
                            }

                        }
                        else
                        {
                            bool whetherMdsSdsExist = false;

                            for (int j = 0; j < excel3DocumentType.Length; j++)
                            {
                                if (excel3DocumentType[i].Contains("MDS/SDS"))
                                {
                                    whetherMdsSdsExist = true;
                                    if (excel3ApprovalStatus[j] == "False")
                                    {
                                        DateTime excelTime = DateTime.Parse(excel3MdsTime[i]);
                                        if (excelTime.AddYears(3) < DateTime.Now)
                                        {
                                            reasonsOfRehection.Add("MDS/SDS 有時間問題");
                                        }
                                    }
                                    reasonsOfRehection.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-02"));
                                }
                            }
                            if (!whetherMdsSdsExist)
                            {
                                reasonsOfRehection.Add("缺少MDS/SDS");
                            }
                        }
                    }
                    //全部報告數值驗證
                    VerificationReport verificationAllReport = new VerificationReport();
                    reasonsOfRehection.AddRange(verificationAllReport.verificationReport(scenarioArray[0], scenarioArray[1], combineAllReportRowData, null, homogeneousType, 0, true, waitForVertify[currentVertificationNumber, 0], waitForVertify[currentVertificationNumber, 1], pbContain, halogenContain));

                    #region paragraph6
                    foreach (var step in outdated6thClickSteps)
                    {
                        string action = step.ActionType ?? string.Empty;
                        int x = int.TryParse(step.X, out int ix) ? ix : 0;
                        int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                        int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                        if (action == "等待時間")
                        {
                            ClickFnctionClass.MouseWin32.Click(action, param);
                        }
                        else
                        {
                            ClickFnctionClass.MouseWin32.Click(action, x, y);
                        }
                    }
                    #endregion
                    #region 第四份Excel 開啟
                    FileTimeInfo excelFilePath_4 = GetLatestFileTimeInfo(downloads_path, ".xls"); 
                    Spire.Xls.Workbook excel_4 = new Spire.Xls.Workbook();
                    excel_4.LoadFromFile(excelFilePath_4.FileName);
                    Spire.Xls.Worksheet excel_4_sheet = excel_4.Worksheets[0];
                    #endregion

                    for (int i = 1; i < excel_3_sheet.Rows.Length; i++)
                    {
                        excel3DocumentType[i] = excel_3_sheet.Rows[i].Columns[2].Text;
                        excel3ApprovalStatus[i] = excel_3_sheet.Rows[i].Columns[1].Text;
                        excel3MdsTime[i] = excel_3_sheet.Rows[i].Columns[7].Text;
                        if (excel_3_sheet.Rows[i].Columns[2].Text == "Test report")
                        {
                            reportCount++;
                        }
                    }

                    for(int i = 1; i<excel_4_sheet.Rows.Length; i++)
                    {
                        var row = combineAllReportRowData.reportRowData.FirstOrDefault(r => r.TestItem.Contains(excel_4_sheet.Rows[i].Columns[1].Text));
                        if(row != null)
                        {
                            if (row.Result == "---")
                            {
                                if (!(excel_4_sheet.Rows[i].Columns[6].Text == "N.D" && (excel_4_sheet.Rows[i].Columns[7].Text == "N.D" || excel_4_sheet.Rows[i].Columns[7].Text == "0") && excel_4_sheet.Rows[i].Columns[9].Text == "" && excel_4_sheet.Rows[i].Columns[10].Text == "" && excel_4_sheet.Rows[i].Columns[11].Text != ""))
                                {
                                    reasonsOfRehection.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-02", excel_4_sheet.Rows[i].Columns[2].Text));
                                }

                            }
                            else if (row.Result == "N.D")
                            {

                                if (!(excel_4_sheet.Rows[i].Columns[6].Text == "N.D" && excel_4_sheet.Rows[i].Columns[7].Text == row.MDL && excel_4_sheet.Rows[i].Columns[9].Text == "" && excel_4_sheet.Rows[i].Columns[10].Text == "" ))
                                {
                                    reasonsOfRehection.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-02", excel_4_sheet.Rows[i].Columns[2].Text));
                                }
                            }
                            else
                            {
                                if(!(excel_4_sheet.Rows[i].Columns[6].Text ==row.Result && excel_4_sheet.Rows[i].Columns[7].Text == "" && excel_4_sheet.Rows[i].Columns[9].Text != "" && excel_4_sheet.Rows[i].Columns[10].Text != ""))
                                {
                                    reasonsOfRehection.Add(ReasonsOfRejectedDocument.GetRejectionReason("R-02", excel_4_sheet.Rows[i].Columns[2].Text));
                                }
                            }
                        }
                    }

                    bool flagPassOrFail = false;
                    string rejectionReasonString = "";
                    foreach (var reasons in reasonsOfRehection)
                    {
                        rejectionReasonString += reasons + "\r\n";
                    }


                    if (reasonsOfRehection.Count==0)
                    {
                        flagPassOrFail = true;
                        #region paragraph7
                        foreach (var step in outdated7thClickSteps)
                        {
                            string action = step.ActionType ?? string.Empty;
                            int x = int.TryParse(step.X, out int ix) ? ix : 0;
                            int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                            int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                            if (action == "等待時間")
                            {
                                ClickFnctionClass.MouseWin32.Click(action, param);
                            }
                            else if (action == "貼上文字")
                            {
                                Clipboard.SetText("PASS");
                                ClickFnctionClass.MouseWin32.Click(action);
                            }
                            else if (action == "全選文字")
                            {
                                ClickFnctionClass.MouseWin32.Click(action);
                            }
                            else
                            {
                                ClickFnctionClass.MouseWin32.Click(action, x, y);
                            }
                        }
                        #endregion

                    }

                    else
                    {
                        flagPassOrFail = false;
                        #region paragraph8
                        foreach (var step in outdated8thClickSteps)
                        {
                            string action = step.ActionType ?? string.Empty;
                            int x = int.TryParse(step.X, out int ix) ? ix : 0;
                            int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                            int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                            if (action == "等待時間")
                            {
                                ClickFnctionClass.MouseWin32.Click(action, param);
                            }
                            else if (action == "貼上文字")
                            {
                                Clipboard.SetText(rejectionReasonString);
                                ClickFnctionClass.MouseWin32.Click(action);
                            }
                            else if (action == "全選文字")
                            {
                                ClickFnctionClass.MouseWin32.Click(action);
                            }
                            else
                            {
                                ClickFnctionClass.MouseWin32.Click(action, x, y);
                            }
                        }
                        #endregion                       
                    }
                    #region 紀錄LOG
                    Row rowLog = new Row() { RowIndex = (uint)reportLogIndex };
                    Cell cellA = new Cell()
                    {
                        CellReference = GetColumnName(0),
                        DataType = CellValues.String,
                        CellValue = new CellValue(DateTime.Now)
                    };
                    rowLog.AppendChild(cellA);

                    Cell cellB = new Cell()
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(waitForVertify[currentVertificationNumber, 0])
                    };
                    rowLog.AppendChild(cellB);

                    //要上面找供應商名稱 留給你
                    Cell cellC = new Cell()
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue("")
                    };
                    rowLog.AppendChild(cellB);

                    Cell cellD = new Cell()
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(waitForVertify[currentVertificationNumber, 1])
                    };
                    rowLog.AppendChild(cellD);

                    Cell cellE = new Cell()
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(flagPassOrFail ? "PASS" : "FAIL")
                    };
                    rowLog.AppendChild(cellE);

                    Cell cellF = new Cell()
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(rejectionReasonString)
                    };
                    rowLog.AppendChild(cellF);


                    for (int i = 0; i < 15; i++)
                    {
                        var rowOfRawData = combineAllReportRowData.reportRowData.FirstOrDefault(r => r.TestItem.Contains(excel_4_sheet.Rows[i].Columns[1].Text));

                        Cell cell = new Cell()
                        {
                            CellReference = GetColumnName(i + 6),
                            DataType = CellValues.String,
                            CellValue = new CellValue(rowOfRawData.Result) // 假設資料有 Name 屬性
                        };
                        rowLog.AppendChild(cell);
                    }
                    sheetData.AppendChild(rowLog);

                    //報告號碼、報告日期要去單篇pdf判斷取用 留給你除錯

                    reportLogIndex++;
                    #endregion
                }
                #region 關閉並儲存log
                workbookPart.Workbook.Save();
                outDatedReportLog.Dispose(); // 釋放資源並關閉檔案
                Console.WriteLine("儲存log");
                #endregion
            }



            // 僅MCD
            else if (indexOfVerificationMode == 3)
            {
                #region MCD_log 初始化
                /*int reportLogIndex = 0; // 報告紀錄的行數索引，從1開始（第一行是標題）
                // 1. 準備路徑與資料夾
                string folderPath = "./MCD_log";
                if (!Directory.Exists(folderPath)) Directory.CreateDirectory(folderPath);
                string nowDayTime = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fullPath = Path.Combine(folderPath, $"{nowDayTime}.xlsx");
                // 2. 【迴圈外】建立檔案基礎結構
                SpreadsheetDocument mcdReportLog = SpreadsheetDocument.Create(fullPath, SpreadsheetDocumentType.Workbook);
                WorkbookPart workbookPart = mcdReportLog.AddWorkbookPart();
                workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                // 建立一個分頁與其資料容器 (SheetData)
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData(); // 這是我們要在迴圈內持續操作的對象
                worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(sheetData);

                Sheet sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "McdLogReport"
                };
                sheets.Append(sheet);
                // --- 標題初始化 (寫入第一列) --
                string[] headers = { "審核時間", "供應商代碼", "料號"};

                Row headerRow = new Row() { RowIndex = (uint)reportLogIndex };
                for (int i = 0; i < headers.Length; i++)
                {
                    Cell cell = new Cell()
                    {
                        CellReference = GetColumnName(i) + reportLogIndex,
                        DataType = CellValues.String,
                        CellValue = new CellValue(headers[i])
                    };
                    headerRow.AppendChild(cell);
                }
                sheetData.AppendChild(headerRow);
                */
                #endregion

                List<List<string>> humanSurveyReport = new List<List<string>>(); //轉人工審查的報告
                int numberOfHumanSurvey = 0;
                int currentVertificationNumber = 0;
                int numberOfBlackList = 0;
                #region 報告更新特定供應商/vendor_black list.xlsx
                string mcdVendorBlackListPath = "./MCD特定供應商/vendor_black list.xlsx";
                Spire.Xls.Workbook mcdVendorBlackListExcel = new Spire.Xls.Workbook();
                mcdVendorBlackListExcel.LoadFromFile(mcdVendorBlackListPath);
                Spire.Xls.Worksheet mcdVendorBlackListExcelSheet = mcdVendorBlackListExcel.Worksheets[0];
                numberOfBlackList = mcdVendorBlackListExcelSheet.Rows.Length - 1;
                string[] mcdVendorBlackListVendorCode = new string[mcdVendorBlackListExcelSheet.Rows.Length];
                string[] mcdVendorBlackListItems = new string[mcdVendorBlackListExcelSheet.Rows.Length];
                for (int i = 1; i < mcdVendorBlackListExcelSheet.Rows.Length; i++)
                {
                    mcdVendorBlackListVendorCode[i] = mcdVendorBlackListExcelSheet.Rows[i].Columns[0].Text.ToUpper();

                    if (mcdVendorBlackListExcelSheet.Rows[i].Columns[1].Text != null)
                    {
                        mcdVendorBlackListItems[i] = mcdVendorBlackListExcelSheet.Rows[i].Columns[1].Text.ToUpper();
                    }
                }
                #endregion                

                while(true)
                {
                    int pbContain = 0;//含鉛材
                    int halogenContain = 0;//含鹵材
                    
                    #region paragraph1
                    foreach (var step in mcd1stClickSteps)
                    {
                        string action = step.ActionType ?? string.Empty;
                        int x = int.TryParse(step.X, out int ix) ? ix : 0;
                        int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                        int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                        if (action == "等待時間")
                        {
                            ClickFnctionClass.MouseWin32.Click(action, param);
                        }
                        else
                        {
                            ClickFnctionClass.MouseWin32.Click(action, x, y);
                        }
                    }
                    #endregion
                    #region 處理第一份excel
                    FileTimeInfo mcdExcelFilePath_1 = GetLatestFileTimeInfo(downloads_path, ".xls"); //獲取下載的最新Excel GetLatestFileTimeInfo(downloads_path, ".xls")
                    Spire.Xls.Workbook mcdExcel_1 = new Spire.Xls.Workbook();  //第一個Excel 物件叫做excel_1
                    mcdExcel_1.LoadFromFile(mcdExcelFilePath_1.FileName);
                    Spire.Xls.Worksheet mcdExcel_1_sheet = mcdExcel_1.Worksheets[0]; //第一頁工作表 物件叫做excel_1_sheet
                    int numberOfVendor = mcdExcel_1_sheet.Rows.Length - 1;
                    if (numberOfVendor == 0)
                    {
                        break;
                    }
                    tbOngoingResult.Text += "供應商數量：" + numberOfVendor.ToString() + "\r\n";
                    int[] listOfNeedToVerifiction = new int[numberOfVendor];
                    for (int i = 0; i < listOfNeedToVerifiction.Length; i++)
                    {
                        listOfNeedToVerifiction[i] = -1;
                    }
                    string[,] waitForVertify = new string[numberOfVendor, 3]; //第一份excel A B C 欄

                    for (int i = 1; i < numberOfVendor; i++)
                    {
                        waitForVertify[i, 0] = mcdExcel_1_sheet.Rows[i].Columns[1].Text; //料號
                        waitForVertify[i, 1] = mcdExcel_1_sheet.Rows[i].Columns[2].Text; //規格描述
                        waitForVertify[i, 2] = mcdExcel_1_sheet.Rows[i].Columns[3].Text; //供應商代碼
                    }
                    #endregion

                   

                    #region 先判斷是否為黑名單
                    for (int i = 0; i < numberOfVendor; i++)
                    {
                        string currentVendor = waitForVertify[i, 0]; // 料號
                        string currentMaterial = waitForVertify[i, 1]; //供應商代碼
                        for (int j = 0; j < numberOfBlackList; j++)
                        {
                            if ((currentVendor == mcdVendorBlackListVendorCode[j]) && (currentMaterial == mcdVendorBlackListItems[j])) //若供應商代號在黑名單且無特定材質則全轉人工審查
                            {
                                humanSurveyReport.Add(new List<string> { waitForVertify[i, 0], waitForVertify[i, 1] }); //紀錄轉人工審查的報告
                                numberOfHumanSurvey += 1; //紀錄人工審查報告數量
                                listOfNeedToVerifiction[currentVertificationNumber] = 1;
                                break;
                            }
                        }
                        currentVertificationNumber++;
                    }
                    currentVertificationNumber = 0;
                    #endregion

                    #region 判斷是否還有需機器審查的報告，若無則結束驗證
                    for (int i = humanSurveyReport.Count - 1; i >= 0; i--)
                    {
                        bool isExist = false;
                        string currentVendorId = humanSurveyReport[i][0]; // 取得清單中的廠商 ID
                        string currentItemNumber = humanSurveyReport[i][0]; // 取得清單中的料號

                        for (int j = 0; j < numberOfVendor; j++)
                        {
                            if (currentVendorId == waitForVertify[j, 0] && currentItemNumber == waitForVertify[j, 1])
                            {
                                listOfNeedToVerifiction[j] = 1; // 標記該供應商報告為需人工審查
                                isExist = true;
                                break;
                            }
                        }
                        if (!isExist)
                        {
                            humanSurveyReport.RemoveAt(i);
                            numberOfHumanSurvey--;
                        }
                    }
                    if (numberOfHumanSurvey == numberOfVendor || numberOfVendor == 0) //若剩下全轉人工審查
                    {
                        break;
                        //MessageBox.Show("本次驗證已完成，全部供應商報告已轉人工審查。", "驗證完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    #endregion

                    for (int i = 0; i < listOfNeedToVerifiction.Length; i++)
                    {
                        if (listOfNeedToVerifiction[currentVertificationNumber] == 1)
                        {
                            currentVertificationNumber++;
                        }
                        else if (listOfNeedToVerifiction[currentVertificationNumber] == -1)
                        {
                            break;
                        }
                    }

                    lbOngingItem.Text = $"{waitForVertify[currentVertificationNumber, 0]} : {waitForVertify[currentVertificationNumber, 1]}";
                    //輸入供應商代號+均勻材質編號 尋找唯一報告
                    #region paragraph2
                    int pastedTimes = 0;
                    foreach (var step in mcd2ndClickSteps)
                    {
                        string action = step.ActionType ?? string.Empty;
                        int x = int.TryParse(step.X, out int ix) ? ix : 0;
                        int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                        int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                        if (action == "等待時間")
                        {
                            ClickFnctionClass.MouseWin32.Click(action, param);
                        }
                        else if (action == "貼上文字")
                        {
                            Clipboard.SetText(waitForVertify[currentVertificationNumber, pastedTimes]);
                            ClickFnctionClass.MouseWin32.Click(action);
                            pastedTimes++;
                        }
                        else if (action == "全選文字")
                        {
                            ClickFnctionClass.MouseWin32.Click(action);
                        }
                        else
                        {
                            ClickFnctionClass.MouseWin32.Click(action, x, y);
                        }
                    }
                    #endregion
                    #region paragraph3
                    foreach (var step in mcd3thClickSteps)
                    {
                        string action = step.ActionType ?? string.Empty;
                        int x = int.TryParse(step.X, out int ix) ? ix : 0;
                        int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                        int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                        if (action == "等待時間")
                        {
                            ClickFnctionClass.MouseWin32.Click(action, param);
                        }
                        else
                        {
                            ClickFnctionClass.MouseWin32.Click(action, x, y);
                        }
                    }
                    #endregion
                    #region 開啟第二 份Excel
                    //判斷XX-情境1-------------------------------------------------------------
                    FileTimeInfo mcdExcelFilePath_2 = GetLatestFileTimeInfo(downloads_path, ".xls"); //抓取最新下載的excel  workbook物件為第二份Excel
                    Spire.Xls.Workbook mcdExcel_2 = new Spire.Xls.Workbook();
                    mcdExcel_2.LoadFromFile(mcdExcelFilePath_2.FileName);
                    Spire.Xls.Worksheet mcdExcel_2_sheet = mcdExcel_2.Worksheets[0];  //第二份Excel的第一張工作表
                    string[] homogeneousNumber = new string[mcdExcel_2_sheet.Rows.Length-1]; //均質編號
                    string[] approvalStatus = new string[mcdExcel_2_sheet.Rows.Length - 1]; //核准狀態
                    
                    for (int i = 1; i< mcdExcel_2_sheet.Rows.Length; i++)
                    {
                        homogeneousNumber[i] = mcdExcel_2_sheet.Rows[i].Columns[1].Text;
                        approvalStatus[i] = mcdExcel_2_sheet.Rows[i].Columns[3].Text;
                    }
                    #endregion
                    #region 初判情境
                    //1-a 1-b 2 3-a 3-b 0~4
                    int[] tentativeScenariosResult = new int[5];
                    for (int i = 0; i < tentativeScenariosResult.Length; i++)
                    {
                        tentativeScenariosResult[i] = -1;
                    }
                    // 1. 判斷 A 欄是否只有 「5」 或者 「44-」 開頭
                    // 正則解釋：^ 代表開頭，(5|44-) 代表 5 或 44-
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 0], @"^(5|44-)"))
                    {
                        tentativeScenariosResult[3] = 1;
                        tentativeScenariosResult[4] = 1;
                    }
                    // 2. 判斷 A 欄是否只有 「XX-」 開頭 (X = 數字)
                    // 正則解釋：^\d{2}- 代表開頭為兩位數字加上橫槓
                    else if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 0], @"^\d{2}-"))
                    {
                        tentativeScenariosResult[2] = 1;
                    }
                    // 3. 如果以上皆非，則為 1a or 1b
                    else
                    {
                        tentativeScenariosResult[0] = 1;
                        tentativeScenariosResult[1] = 1;
                    }
                    #endregion

                    #region 含鉛材含鹵材
                    // 4. 判斷 B 欄是否包含關鍵字 (含鉛材)
                    if (waitForVertify[currentVertificationNumber, 1].Contains("_HS-PB") || waitForVertify[currentVertificationNumber, 1].Contains("-HS-PB"))
                    {
                        pbContain = 1; //含鉛材
                    }
                    // 5. 判斷 B 欄是否包含關鍵字 (含鹵材)
                    if (waitForVertify[currentVertificationNumber, 1].Contains("HS-NON-HF"))
                    {
                        halogenContain = 1;//含鹵材
                    }
                    #endregion
                    List<string> compositionOfHomogeneous = new List<string>();
                    //載入list中與供應商代號對應的excel-------------------------------------------------------------------------------
                    compositionOfHomogeneous = FindListName.findCompositionOfHomogeneous(waitForVertify[currentVertificationNumber, 0], waitForVertify[currentVertificationNumber, 1]);
                    if (compositionOfHomogeneous == null)
                    {
                        humanSurveyReport.Add(new List<string> { waitForVertify[currentVertificationNumber, 0], waitForVertify[currentVertificationNumber, 1] }); //紀錄轉人工審查的報告
                    }

                    //找需要審核的均質
                    List<string> needToVertifyHomogeneousItems = new List<string>();
                    bool flagNeedToReject = false;
                    foreach (var compositionOfHomogeneousItem in compositionOfHomogeneous)
                    {
                        bool whetherFind = false;
                        for (int i = 1; i < homogeneousNumber.Length; i++)
                        {
                            if (compositionOfHomogeneousItem == homogeneousNumber[i])
                            {
                                whetherFind = true;
                                if (approvalStatus[i] != "Approved")
                                {
                                    needToVertifyHomogeneousItems.Add(homogeneousNumber[i]);
                                    break;
                                }
                            }
                        }                            
                        if(!whetherFind)
                        {
                            flagNeedToReject = true;
                        }
                    }
                    if(needToVertifyHomogeneousItems.Count!= homogeneousNumber.Length)
                    {
                        flagNeedToReject = true;
                    }

                    if(flagNeedToReject)
                    {
                        reasonsOfRehection.Add(ReasonsOfRejectedDocument.GetRejectionReason("M-01"));
                    }
                    #region 有待檢項目跑過期報告審查流程 這可能會出事
                    if (needToVertifyHomogeneousItems.Count > 0)
                    {
                        //先有個跳轉
                        #region paragraph4
                        foreach (var step in mcd4thClickSteps)
                        {
                            string action = step.ActionType ?? string.Empty;
                            int x = int.TryParse(step.X, out int ix) ? ix : 0;
                            int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                            int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                            if (action == "等待時間")
                            {
                                ClickFnctionClass.MouseWin32.Click(action, param);
                            }
                            else
                            {
                                ClickFnctionClass.MouseWin32.Click(action, x, y);
                            }
                        }
                        #endregion
                        //複製貼上過期報告審查流程
                    }
                    #endregion

                    #region paragraph5
                    foreach (var step in mcd5thClickSteps)
                    {
                        string action = step.ActionType ?? string.Empty;
                        int x = int.TryParse(step.X, out int ix) ? ix : 0;
                        int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                        int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                        if (action == "等待時間")
                        {
                            ClickFnctionClass.MouseWin32.Click(action, param);
                        }
                        else
                        {
                            ClickFnctionClass.MouseWin32.Click(action, x, y);
                        }
                    }
                    #endregion

                    #region 下載excel3

                    FileTimeInfo mcdExcelFilePath_3 = GetLatestFileTimeInfo(downloads_path, ".xls"); //獲取下載的最新Excel GetLatestFileTimeInfo(downloads_path, ".xls")
                    Spire.Xls.Workbook mcdExcel_3 = new Spire.Xls.Workbook();  //第一個Excel 物件叫做excel_1
                    mcdExcel_3.LoadFromFile(mcdExcelFilePath_3.FileName);
                    Spire.Xls.Worksheet mcdExcel_3_sheet = mcdExcel_3.Worksheets[0]; //第一頁工作表 物件叫做excel_1_sheet

                    bool sdsMdsMust = false;//是否需要SDS/MDS
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 0], @"^(1B|01-)")) sdsMdsMust = true;
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 1], @"^Lead frame_")) sdsMdsMust = true;
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 1], @"CAPACITOR")) sdsMdsMust = true;
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 1], @"^H/S_")) sdsMdsMust = true;
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 1], @"RESISTOR")) sdsMdsMust = true;
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 1], @"^Inductor_")) sdsMdsMust = true;
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 1], @"^Filter_")) sdsMdsMust = true;
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 1], @"^IC_")) sdsMdsMust = true;
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 1], @"^CRYSTAL_")) sdsMdsMust = true;
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 1], @"^Chip bead_")) sdsMdsMust = true;
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 1], @"CONNECTOR_")) sdsMdsMust = true;
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 1], @"^Transformer_")) sdsMdsMust = true;
                    if (Regex.IsMatch(waitForVertify[currentVertificationNumber, 1], @"^Switch_")) sdsMdsMust = true;

                    for (int i = 0; i < mcdExcel_3_sheet.Rows.Length; i++)
                    {
                        if(mcdExcel_3_sheet.Rows[i].Columns[1].Text == "" && mcdExcel_3_sheet.Rows[i].Columns[5].Text == "MDS/SDS")
                        {
                            DateTime excelTime = DateTime.Parse(mcdExcel_3_sheet.Rows[i].Columns[1].Text);
                            if (!(excelTime.AddYears(3) > DateTime.Now))
                            {
                                reasonsOfRehection.Add(ReasonsOfRejectedDocument.GetRejectionReason("M-03"));
                            }
                        }
                        if(!(mcdExcel_3_sheet.Rows[i].Columns[1].Text == "" && mcdExcel_3_sheet.Rows[i].Columns[5].Text == "MDS/SDS"))
                        {
                            if (sdsMdsMust)
                            {
                                reasonsOfRehection.Add(ReasonsOfRejectedDocument.GetRejectionReason("M-02"));
                            }
                        }
                    }
                    for (int i = 1; i < numberOfVendor; i++)
                    {
                        waitForVertify[i, 0] = mcdExcel_1_sheet.Rows[i].Columns[1].Text; //料號
                        waitForVertify[i, 1] = mcdExcel_1_sheet.Rows[i].Columns[2].Text; //規格描述
                        waitForVertify[i, 2] = mcdExcel_1_sheet.Rows[i].Columns[3].Text; //供應商代碼
                    }
                    #endregion

                    if(reasonsOfRehection.Count==0)
                    {
                        #region paragraph6
                        foreach (var step in mcd6thClickSteps)
                        {
                            string action = step.ActionType ?? string.Empty;
                            int x = int.TryParse(step.X, out int ix) ? ix : 0;
                            int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                            int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                            if (action == "等待時間")
                            {
                                ClickFnctionClass.MouseWin32.Click(action, param);
                            }
                            else if (action == "貼上文字")
                            {
                                Clipboard.SetText("PASS");
                                ClickFnctionClass.MouseWin32.Click(action);
                            }
                            else if (action == "全選文字")
                            {
                                ClickFnctionClass.MouseWin32.Click(action);
                            }
                            else
                            {
                                ClickFnctionClass.MouseWin32.Click(action, x, y);
                            }
                        }
                        #endregion
                    }
                    else
                    {
                        string rejectionReasonString = "";
                        foreach (var reasons in reasonsOfRehection)
                        {
                            rejectionReasonString += reasons + "\r\n";
                        }
                        #region paragraph7
                        foreach (var step in mcd7thClickSteps)
                        {
                            string action = step.ActionType ?? string.Empty;
                            int x = int.TryParse(step.X, out int ix) ? ix : 0;
                            int y = int.TryParse(step.Y, out int iy) ? iy : 0;
                            int param = int.TryParse(step.Param, out int ip) ? ip : 0;
                            if (action == "等待時間")
                            {
                                ClickFnctionClass.MouseWin32.Click(action, param);
                            }
                            else if (action == "貼上文字")
                            {
                                Clipboard.SetText(rejectionReasonString);
                                ClickFnctionClass.MouseWin32.Click(action);
                            }
                            else if (action == "全選文字")
                            {
                                ClickFnctionClass.MouseWin32.Click(action);
                            }
                            else
                            {
                                ClickFnctionClass.MouseWin32.Click(action, x, y);
                            }
                        }
                        #endregion
                    }
                    //while 尾
                }

            }
        }

        private void btnHomogeneousCustomVerificationSequence_Click(object sender, EventArgs e)
        {
            HomogeneousCustomVerificationSequencePage homogeneousCustomVerificationSequencePage = new HomogeneousCustomVerificationSequencePage();
            homogeneousCustomVerificationSequencePage.ShowDialog();
        }

        private void btnMcdCustomVerificationSequence_Click(object sender, EventArgs e)
        {
            McdCustomVerificationSequencePage mcdCustomVerificationSequencePage = new McdCustomVerificationSequencePage();
            mcdCustomVerificationSequencePage.ShowDialog();
        }


        #region 驗證模式選項的RadioButton CheckedChanged事件
        private void rBtnOutdatedFrist_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtnOutdatedFrist.Checked)
            {
                indexOfVerificationMode = 0;
                btnAutomatedBatchVerificationStart.Enabled = true;
            }
        }

        private void rBtmMcdFirst_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtmMcdFirst.Checked)
            {
                indexOfVerificationMode = 1;
                btnAutomatedBatchVerificationStart.Enabled = true;
            }
        }

        private void rBtnOutdatedOnly_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtnOutdatedOnly.Checked)
            {
                indexOfVerificationMode = 2;
                btnAutomatedBatchVerificationStart.Enabled = true;
            }
        }

        private void rBtnMcdOnly_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtnMcdOnly.Checked)
            {
                indexOfVerificationMode = 3;
                btnAutomatedBatchVerificationStart.Enabled = true;
            }
        }

        #endregion

        static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex + 1;
            string columnName = String.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }
            return columnName;
        }
    }
}
