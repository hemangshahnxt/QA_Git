using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Security.Cryptography;

namespace SeleniumAutomation
{
    public class Verification
    {
        //private static IWebDriver driver;

       public static void Sleep()
        {
            System.Threading.Thread.Sleep(3000);
        }

        //Encrypting a password method
        public static string Encrypt(string value)
        {
            using(MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider())
            {
                UTF8Encoding uTF8 = new UTF8Encoding();
                byte[] data = md5.ComputeHash(uTF8.GetBytes(value));
                return Convert.ToBase64String(data);
            }
        }
        //Creating Random String for username
        public static Random random = new Random((int)DateTime.Now.Ticks);
        public static string RandomString(int size)
        {
            StringBuilder builder = new StringBuilder();
            char ch;
            for (int i = 0; i < size; i++)
            {
                ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));
                builder.Append(ch);
            }

            return builder.ToString();


            //// get 1st random string 
            //string Rand1 = RandomString(4);

            //// get 2nd random string 
            //string Rand2 = RandomString(6);

            //// creat full rand string
            //string docNum = Rand1 + "-" + Rand2;
        }

        public static string GenerateRandomString()
        {
            return "SeleniumTest" + RandomString(8);
        }

       public static void Email_send()
        {
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp-mail.outlook.com");
            mail.From = new MailAddress("h.shah@nextech.com");
            mail.To.Add("devsquadb@nextech.com");
            mail.Subject = "MPV Smoke Test";
            mail.Body = "mail with attachment";

            System.Net.Mail.Attachment attachment;
            attachment = new System.Net.Mail.Attachment("C:\\Selenium Result\\Test Results.txt");
            mail.Attachments.Add(attachment);

            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("h.shah@nextech.com", "");
            SmtpServer.EnableSsl = true;

            SmtpServer.Send(mail);
        }

    }
   
}
