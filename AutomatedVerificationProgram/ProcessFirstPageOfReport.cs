using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Spire.Pdf.Graphics;
using Spire.Pdf;
using System.Windows.Forms;

namespace AutomatedVerificationProgram
{
    internal class ProcessFirstPageOfReport
    {
        public static string[] pdfReportFilePage1StringSplitArray;
        public int processFirstPageOfReport(string pdfReportFilePath, ReportParaments reportParaments)
        {
            PdfDocument pdfReportFile = new PdfDocument();
            pdfReportFile.LoadFromFile(pdfReportFilePath);
            PdfDocument pdfReportFilePage1 = new PdfDocument();
            PdfDocument pdfReportFilePage2 = new PdfDocument();
            PdfPageBase page01;
            PdfPageBase page02;

            //將檔案第一頁複製到pdfReportFilePage1.pdf
            page01 = pdfReportFilePage1.Pages.Add(pdfReportFile.Pages[0].Size, new PdfMargins(0));
            page02 = pdfReportFilePage2.Pages.Add(pdfReportFile.Pages[1].Size, new PdfMargins(0));

            pdfReportFile.Pages[0].CreateTemplate().Draw(page01, new PointF(0, 0));
            pdfReportFile.Pages[1].CreateTemplate().Draw(page02, new PointF(0, 0));
            pdfReportFilePage1.SaveToFile("pdfReportFilePage1.pdf", Spire.Pdf.FileFormat.PDF);
            pdfReportFilePage2.SaveToFile("pdfReportFilePage2.pdf", Spire.Pdf.FileFormat.PDF);
            StringBuilder pdfReportFilePage1StringBulider = new StringBuilder();
            pdfReportFilePage1StringBulider.Append(page01.ExtractText());
            pdfReportFilePage1.Close();
            String pdfReportFilePage1FileName = "pdfReportFilePage1.txt";
            File.WriteAllText(pdfReportFilePage1FileName, pdfReportFilePage1StringBulider.ToString());

            StringBuilder pdfReportFilePage2StringBulider = new StringBuilder();
            pdfReportFilePage2StringBulider.Append(page02.ExtractText());
            pdfReportFilePage2.Close();
            String pdfReportFilePage2FileName = "pdfReportFilePage2.txt";
            File.WriteAllText(pdfReportFilePage2FileName, pdfReportFilePage2StringBulider.ToString());

            pdfReportFilePage1StringBulider = null;
            pdfReportFilePage2StringBulider = null;
            string pdfReportFilePage1String = "";
            using (StreamReader reader = new StreamReader(pdfReportFilePage1FileName, Encoding.UTF8))
            {
                pdfReportFilePage1String = reader.ReadToEnd();
            }
            pdfReportFilePage1StringSplitArray = Regex.Split(pdfReportFilePage1String, "\r\n");
            DetectionReportNo detectionReportNo = new DetectionReportNo();
            reportParaments.reportNunber = detectionReportNo.detectionReportNo(pdfReportFilePage1StringSplitArray);
            ReportTestingOrganizationDetction reportTestingOrganizationDetction = new ReportTestingOrganizationDetction();
            int testingOrganizationNum = -1;
            testingOrganizationNum = reportTestingOrganizationDetction.detectFun(reportParaments.reportNunber);
            return testingOrganizationNum;

        }
    }
}
