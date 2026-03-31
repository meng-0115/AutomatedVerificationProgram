using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomatedVerificationProgram
{
    internal class DetectionReportNo
    {
        public string detectionReportNo(string[] pdfReportFilePage1StringSplitArray)
        {
            string foundReportNo = "";

            try
            {
                for (int i = 0; i < pdfReportFilePage1StringSplitArray.Length; i++)
                {
                    string rawLine = pdfReportFilePage1StringSplitArray[i];
                    if (string.IsNullOrWhiteSpace(rawLine)) continue;

                    // 轉大寫比對，但「不」預先移除空白，否則編號會跟日期黏死
                    string upperLine = rawLine.ToUpper();

                    // 1. 定位起始點 (依據舊邏輯中的所有關鍵字)
                    int startIndex = -1;
                    int offset = 0;

                    if (upperLine.Contains("(NO.):")) { startIndex = upperLine.IndexOf("(NO.):"); offset = 6; }
                    else if (upperLine.Contains("NO.:")) { startIndex = upperLine.IndexOf("NO.:"); offset = 4; }
                    else if (upperLine.Contains("NO.) :")) { startIndex = upperLine.IndexOf("NO.) :"); offset = 7; }
                    else if (upperLine.Contains("NO. :")) { startIndex = upperLine.IndexOf("NO. :"); offset = 5; }
                    else if (upperLine.Contains("REPORT NO.")) { startIndex = upperLine.IndexOf("NO."); offset = 3; }
                    else if (upperLine.Contains("NO.")) { startIndex = upperLine.IndexOf("NO."); offset = 3; }
                    else if (upperLine.Contains(":TWNC")) { startIndex = upperLine.IndexOf(":TWNC"); offset = 1; }

                    // 處理起始點
                    if (startIndex != -1)
                    {
                        // 擷取關鍵字之後的剩餘字串
                        string afterKey = rawLine.Substring(startIndex + offset).Trim();

                        // 2. 處理結束點：遇到空格、全形空格、或特定中文字就切斷
                        // 這樣可以分離出 ETR25C07357 與後面的 日期(Date)
                        char[] splitChars = new char[] { ' ', '　', '\t', '(', '（' };
                        string[] parts = afterKey.Split(splitChars, StringSplitOptions.RemoveEmptyEntries);

                        if (parts.Length > 0)
                        {
                            foundReportNo = parts[0];

                            // 3. 進階清理：移除可能殘留的干擾字眼
                            // 有些 OCR 會把編號跟 "DATE" 辨識在同一個單字裡
                            if (foundReportNo.ToUpper().Contains("DATE"))
                            {
                                int dateIdx = foundReportNo.ToUpper().IndexOf("DATE");
                                foundReportNo = foundReportNo.Substring(0, dateIdx);
                            }

                            if (foundReportNo.Contains("頁數"))
                            {
                                int pageIdx = foundReportNo.IndexOf("頁數");
                                foundReportNo = foundReportNo.Substring(0, pageIdx);
                            }

                            // 移除頭尾標點符號
                            foundReportNo = foundReportNo.Trim(':', ')', ']', ' ', '　');
                        }
                    }

                    // 4. 特別處理 Intertek KR (RT...R 格式) - 如果一般邏輯沒抓到再用 Regex
                    if (string.IsNullOrEmpty(foundReportNo) && System.Text.RegularExpressions.Regex.IsMatch(upperLine, @"RT\d{2}R"))
                    {
                        var match = System.Text.RegularExpressions.Regex.Match(upperLine, @"RT\d{2}R[A-Z0-9]+");
                        if (match.Success) foundReportNo = match.Value;
                    }

                    // 5. 驗證排除：如果是 JOBNO 則重來
                    if (upperLine.Contains("JOBNO"))
                    {
                        foundReportNo = "";
                        continue;
                    }

                    if (!string.IsNullOrEmpty(foundReportNo)) break;
                }
            }
            catch (Exception ex)
            {
                // 建議保留 Log 紀錄以便除錯
            }

            return foundReportNo;
        }
    }
}
