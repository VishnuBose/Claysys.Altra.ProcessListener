using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
using System.Xml;

namespace Claysys.Altra.ProcessListener
{
    class Config
    {

        public string TimePeriod { get; set; }
        public string ProcessName { get; set; }
        public string StatusACKMailID { get; set; }
        public string Host { get; set; }
        public string FromMail { get; set; }
        public string Password { get; set; }
        public string ServerName { get; set; }

        public Config()
        {
            XmlDocument doc = new XmlDocument();
            string path = AppDomain.CurrentDomain.BaseDirectory;
            doc.Load(Path.Combine(path, "ProcessConfig.xml"));
            XmlNodeList nodelist = doc.SelectNodes("/configuration");
            foreach (XmlNode node in nodelist)
            {
                TimePeriod = node.SelectSingleNode("TimePeriod").InnerText;
                ProcessName = node.SelectSingleNode("ProcessName").InnerText;
                StatusACKMailID = node.SelectSingleNode("StatusACKMailID").InnerText;
                Host = node.SelectSingleNode("Host").InnerText;
                FromMail = node.SelectSingleNode("FromMail").InnerText;
                Password = node.SelectSingleNode("Password").InnerText;
                ServerName = node.SelectSingleNode("ServerName").InnerText;
            }
        }
    }
}
