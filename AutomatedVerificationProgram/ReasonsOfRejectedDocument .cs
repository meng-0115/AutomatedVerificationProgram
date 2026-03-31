using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomatedVerificationProgram
{
    internal class ReasonsOfRejectedDocument
    {
        private static readonly Dictionary<string, string> ReasonData = new Dictionary<string, string>
        {
            { "R-01", "均質類型異常：請挑選此材料適用的均質類型 Abnormal homogeneity type: Please select the applicable type for this material." },
            { "R-02", "資訊不一致：報告日期應為測試期間最後一日。 Information mismatch: Testing date shall be the last day of testing period." },
            { "R-03", "資訊不一致：檢測單位與您所提交於GPM的資訊不符。 Information mismatch: Testing laboratory does not match your submitted information on GPM." },
            { "R-04", "資訊不一致：檢測物質與您所提交於GPM的資訊不符。 Information mismatch: Testing substances does not match your submitted information on GPM." },
            { "R-05", "報告已過期：有效期須在一年內。 Report expired: Validity must be within 1 year." },
            // R-06, R-07, R-08, R-09, R-10 包含可變參數 {0}
            { "R-06", "檢測項目缺項：報告未檢測 [{0}]。 Missing test item: The report does not include testing for [{0}]." },
            { "R-07", "檢測方法錯誤：[{0}] 未採用指定之檢測方法。 Incorrect test method: [{0}] was not tested using specified method." },
            { "R-08", "檢測精度不足：[{0}] MDL 未達其他3rd party 能力水準 Insufficient precision: [{0}] MDL does not meet 3rd party capability levels." },
            { "R-09", "物質含量超標：[{0}] 檢測值超過恕限值。 Exceeds limit: The test result for [{0}] exceeds the allowed limit." },
            { "R-10", "缺少數位簽章：於報告 [{0}] 查無數位簽章，請聯繫您的供應商或檢測單位 Missing electronic signature: Can't find electronic signature in test report [{0}]. Please contact your supplier or 3rd party lab." },
            { "R-11", "報告不可重複上傳 Duplicate report is not allowed." },
            { "R-12", "缺少必要文件：未上傳SDS/MDS。 Missing required document: SDS/MDS not uploaded." },
            { "R-13", "SDS/MDS已過期：有效期須在三年內。 SDS/MDS expired: Validity must be within 3 years." },
            { "R-16", "資訊不一致：請檢查報告是否有測試{0}。如無此測項，GPM Test Data請申報ND、MDL請寫0 or ND Information mismatch: Please check whether {0} is tested in the report. If not tested, enter Test Data = ND and MDL = 0/ND in GPM." },
            { "R-17", "資訊不一致：{0}的報告數值與您所提交於GPM的資訊不符 Information mismatch: Data mismatch: {0} report value does not match GPM submission" },
            { "M-01", "資訊不一致：請確認所引用的均質材料是否適用於此料號。\nPlease verify that the referenced homogeneous material applies to this part number." },
            { "M-02", "資訊空缺：請上傳MDS\nPlease submit MDS." },
            { "M-03", "資訊空缺：請在MDS下方的表格，填入此料號所使用均質材料的Material, Manufacturer, Type\nMissing information: Please fill in Material, Manufacturer, and Type in the MDS table below." },
            { "M-04", "資訊不一致：請確認MDS所申報的料號是否正確\nInformation mismatch: Please verify if the part number declared in MDS is correct." }
        };

        /// <summary>
        /// 取得退件原因，支援填入可變參數
        /// </summary>
        /// <param name="code">原因編號</param>
        /// <param name="param">要填入 [測項] 的文字 (選填)</param>
        public static string GetRejectionReason(string rejectionReason, string param = "")
        {
            if (string.IsNullOrEmpty(rejectionReason)) return "請輸入編號\nPlease enter a code";

            string key = rejectionReason.Trim().ToUpper();
            if (ReasonData.TryGetValue(key, out string template))
            {
                // 如果有提供參數，且模板中含有 {0}，則進行替換
                if (!string.IsNullOrEmpty(param) && template.Contains("{0}"))
                {
                    return string.Format(template, param);
                }
                return template;
            }
            return "找不到對應編號\nCode not found";
        }
    }
}
