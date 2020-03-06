using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.ServiceProcess;
using System.Text;
using System.Timers;

namespace Claysys.Altra.ProcessListener
{
    public partial class Service1 : ServiceBase
    {
        Timer timer = new Timer();
        private int timePeriod = 0;
        private string processName;
        private string StatusACKMailID;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Claysys Altra process listener service is started at " + DateTime.Now);
            ReadCustomConfig();
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = timePeriod == 0 ? 1000 : timePeriod; //number in milisecinds  
            timer.Enabled = true;
        }


        private void ReadCustomConfig()
        {
            timePeriod = Convert.ToInt32(ConfigurationManager.AppSettings["TimePeriod"]);
            processName = Convert.ToString(ConfigurationManager.AppSettings["ProcessName"]);
            StatusACKMailID = Convert.ToString(ConfigurationManager.AppSettings["StatusACKMailID"]);
        }

        protected override void OnStop()
        {
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            WriteToFile("Claysys Altra process listener service is recall at " + DateTime.Now);
            Process[] pname = Process.GetProcessesByName(processName);
            if (pname.Length == 0)
            {
                WriteToFile(DateTime.Now + " : The process is not running : ");
                try
                {
                    SendEmail(StatusACKMailID, "", "", "Altra process listener service status", "The process is not running");
                    WriteToFile(DateTime.Now + " : Send mail successfully ");
                }
                catch (Exception ex)
                {
                    WriteToFile(DateTime.Now + " : Send mail not successfully , " + ex);

                    throw ex;
                }
            }
            else
            {
                WriteToFile(DateTime.Now + " : The process is  running : ");
                //try
                //{
                //    SendEmail(StatusACKMailID, "", "", "Altra process listener service status", "The process is  running");
                //    WriteToFile(DateTime.Now + " : Send mail successfully ");
                //}
                //catch (Exception ex)
                //{
                //    WriteToFile(DateTime.Now + " : Send mail not successfully , " + ex);

                //    throw ex;
                //}
            }
        }

        public static void SendEmail(String ToEmail, string cc, string bcc, String Subj, string Message)
        {
            //Reading sender Email credential from web.config file  

            string HostAdd = ConfigurationManager.AppSettings["Host"].ToString();
            string FromEmailid = ConfigurationManager.AppSettings["FromMail"].ToString();
            string Pass = ConfigurationManager.AppSettings["Password"].ToString();

            //creating the object of MailMessage  
            MailMessage mailMessage = new MailMessage();
            mailMessage.From = new MailAddress(FromEmailid); //From Email Id  
            mailMessage.Subject = Subj; //Subject of Email  
            mailMessage.Body = Message; //body or message of Email  
            mailMessage.IsBodyHtml = true;

            string[] ToMuliId = ToEmail.Split(',');
            foreach (string ToEMailId in ToMuliId)
            {
                mailMessage.To.Add(new MailAddress(ToEMailId)); //adding multiple TO Email Id  
            }


            if (!string.IsNullOrEmpty(cc))
            {
                string[] CCId = cc.Split(',');

                foreach (string CCEmail in CCId)
                {

                    mailMessage.CC.Add(new MailAddress(CCEmail)); //Adding Multiple CC email Id  
                }
            }

            if (!string.IsNullOrEmpty(bcc))
            {
                string[] bccid = bcc.Split(',');
                foreach (string bccEmailId in bccid)
                {
                    mailMessage.Bcc.Add(new MailAddress(bccEmailId)); //Adding Multiple BCC email Id  
                }
            }
            SmtpClient smtp = new SmtpClient();  // creating object of smptpclient  
            smtp.Host = HostAdd;              //host of emailaddress for example smtp.gmail.com etc  

            //network and security related credentials  

            smtp.EnableSsl = false;
            NetworkCredential NetworkCred = new NetworkCredential();
            NetworkCred.UserName = mailMessage.From.Address;
            NetworkCred.Password = Pass;
            smtp.UseDefaultCredentials = true;
            smtp.Credentials = NetworkCred;
            smtp.Port = 587;
            smtp.Send(mailMessage); //sending Email  
        }

        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }
    }
}
