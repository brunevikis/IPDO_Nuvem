using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;


namespace IPDO_Compass
{
    public static class Tools
    {

        public static string GetMonthNameAbrev(int month)
        {

            switch (month)
            {
                case 1: return "JAN";
                case 2: return "FEV";
                case 3: return "MAR";
                case 4: return "ABR";
                case 5: return "MAI";
                case 6: return "JUN";
                case 7: return "JUL";
                case 8: return "AGO";
                case 9: return "SET";
                case 10: return "OUT";
                case 11: return "NOV";
                case 12: return "DEZ";

                default:
                    return null;
            }
        }

        public static string GetMonthNumAbrev(int month)
        {

            switch (month)
            {
                case 1: return "01_jan";
                case 2: return "02_fev";
                case 3: return "03_mar";
                case 4: return "04_abr";
                case 5: return "05_mai";
                case 6: return "06_jun";
                case 7: return "07_jul";
                case 8: return "08_ago";
                case 9: return "09_set";
                case 10: return "10_out";
                case 11: return "11_nov";
                case 12: return "12_dez";

                default:
                    return null;
            }
        }

        public static string GetMonthNum(int month)
        {

            switch (month)
            {
                case 1: return "01_janeiro";
                case 2: return "02_fevereiro";
                case 3: return "03_março";
                case 4: return "04_abril";
                case 5: return "05_maio";
                case 6: return "06_junho";
                case 7: return "07_julho";
                case 8: return "08_agosto";
                case 9: return "09_setembro";
                case 10: return "10_outubro";
                case 11: return "11_novembro";
                case 12: return "12_dezembro";

                default:
                    return null;
            }
        }

        public static string GetMonthName(int month)
        {

            switch (month)
            {
                case 1: return "Janeiro";
                case 2: return "Fevereiro";
                case 3: return "Marco";
                case 4: return "Abril";
                case 5: return "Maio";
                case 6: return "Junho";
                case 7: return "Julho";
                case 8: return "Agosto";
                case 9: return "Setembro";
                case 10: return "Outubro";
                case 11: return "Novembro";
                case 12: return "Dezembro";

                default:
                    return null;
            }
        }

        public static async Task SendMail(string attach, string body, string subject, string receiversGroup)
        {
            System.Net.Mail.SmtpClient cli = new System.Net.Mail.SmtpClient();

            cli.Host = "smtp.gmail.com";
            cli.Port = 587;
            cli.Credentials = new System.Net.NetworkCredential("cpas.robot@gmail.com", "cp@s9876");

            cli.EnableSsl = true;


            var msg = new System.Net.Mail.MailMessage()
            {
                Subject = subject,
            };


            if (attach.Contains(";"))
                foreach (var att in attach.Split(';'))
                    if (File.Exists(att))
                        msg.Attachments.Add(new System.Net.Mail.Attachment(att));

            msg.Body = body;

            msg.Sender = msg.From = new System.Net.Mail.MailAddress("cpas.robot@gmail.com");

            var receivers = ConfigurationManager.AppSettings[receiversGroup];

            foreach (var receiver in receivers.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (!string.IsNullOrWhiteSpace(receiver.Trim()))
                    msg.To.Add(new System.Net.Mail.MailAddress(receiver.Trim()));
            }

            if (body.Contains("html"))
                msg.IsBodyHtml = true;

            int trials = 3;
            int sleepTime = 1000 * 60;
            int trial = 0;
            while (trial++ < trials)
            {
                try
                {
                    await cli.SendMailAsync(msg);
                    break;
                }
                catch (Exception e)
                {
                    System.Threading.Thread.Sleep(sleepTime);
                }
            }

        }



     
    }

}

