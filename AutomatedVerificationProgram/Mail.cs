using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using Microsoft.Office.Interop.Outlook;
using Exception = System.Exception;


namespace AutomatedVerificationProgram
{
    internal class Mail
    {
        public static void Outlook_to_F(string name_no_sp, string number_no_sp, string ff_result)
        {
            try
            {
                // 建立Outlook Application物件
                Microsoft.Office.Interop.Outlook.Application outlookApp = new Microsoft.Office.Interop.Outlook.Application();

                // 建立信件物件
                MailItem mail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

                // 設定信件的收件人、主旨、內容等資訊
                mail.To = "GRP-GPM@spil.com.tw";
                mail.Subject = "供應商代號: " + name_no_sp + " 均勻材質編號: " + number_no_sp + " 檢測結果: " + ff_result;
                mail.Body = ("供應商代號: " + name_no_sp + "\r\n"
                           + "均勻材質編號: " + number_no_sp + "\r\n"
                           + "檢測結果: " + ff_result + "\r\n");

                // 寄出信件
                mail.Send();
            }
            catch (Exception ex)
            {
                Console.WriteLine("郵件發送失敗: " + ex.Message);
            }
        }
    }
}
