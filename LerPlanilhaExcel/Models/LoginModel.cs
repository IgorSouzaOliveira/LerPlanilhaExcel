using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace LerPlanilhaExcel.Models
{
    class LoginModel
    {
        public String CompanyDB { get; set; }
        public String Password { get; set; }
        public String UserName { get; set; }
        public String Language { get; set; }
        public LoginModel()
        {
            ReadConnection();
        }
        public void ReadConnection()
        {
            try
            {
                var directory = AppDomain.CurrentDomain.BaseDirectory;
                var path = Path.Combine(directory, @"Config\\Connection.xml");

                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(path);

                XmlNodeList xnNode = xmlDocument.GetElementsByTagName("ConnectionData");

                foreach (XmlNode xn in xnNode)
                {
                    CompanyDB = xn["DBCompany"].InnerText;
                    UserName = xn["SAPUserName"].InnerText;
                    Password = xn["SAPPassword"].InnerText;
                    Language = xn["Language"].InnerText;                   

                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);                
            }
        }
    }
}
