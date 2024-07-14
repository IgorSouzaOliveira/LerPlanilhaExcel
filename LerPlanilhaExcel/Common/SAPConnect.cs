using LerPlanilhaExcel.Models;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace LerPlanilhaExcel.Common
{
    class SAPConnect
    {

        private static string B1Session;
        private static String _slAddress;
        private static String _slServer;

        public static void ReadAddress()
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
                    _slAddress = xn["AddressLayer"].InnerText;
                    _slServer = xn["ServerLayer"].InnerText;          

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static string SAPLogin()
        {
            try
            {
                LoginModel login = new LoginModel();
                ReadAddress();
                var client = new RestClient(_slAddress);
                var request = new RestRequest("/Login", Method.POST);

                var body = Newtonsoft.Json.JsonConvert.SerializeObject(login);
                request.AddHeader("Content-Type", "application/json");
                request.AddParameter("application/json", body, ParameterType.RequestBody);

                ServicePointManager.ServerCertificateValidationCallback += new System.Net.Security.RemoteCertificateValidationCallback(ValidateServerCertificate);

                IRestResponse rest = client.Execute(request);

                B1Session = rest.Cookies.FirstOrDefault().Value;
                DateTime today = DateTime.Now;

                if (rest.StatusCode == HttpStatusCode.OK)
                {
                    Console.WriteLine($"Conexão realizada com sucesso na base de dados: {login.CompanyDB}.\n");
                    return B1Session;
                }
                else
                {
                    Console.WriteLine($"Erro - {rest.StatusDescription.ToString()}");
                    return null;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }           

        }
        public static void UpdateBP(BusinessPartner oBP, int line)
        {
            try
            {
                var client = new RestClient(_slAddress);                
                var request = new RestRequest($"/BusinessPartners('{oBP.CardCode}')", Method.PATCH);
                request.AddHeader("Content-Type", "application/json");
                
                var body = Newtonsoft.Json.JsonConvert.SerializeObject(oBP);                
                request.AddParameter("application/json", body, ParameterType.RequestBody);                

                CookieContainer cookiecon = new CookieContainer();                
                cookiecon.Add(new Cookie("B1SESSION", B1Session, "/b1s/v1", _slServer)); 
                client.CookieContainer = cookiecon;
                
                ServicePointManager.ServerCertificateValidationCallback += new System.Net.Security.RemoteCertificateValidationCallback(ValidateServerCertificate);               

                IRestResponse response = client.Execute(request);

                switch (response.StatusCode)
                {
                    case HttpStatusCode.NoContent:
                        Console.WriteLine($"#Service Layer - #{line}º Parceiro de negocio: {oBP.CardCode}, atualizado com sucesso.");
                        break;

                    default:
                        dynamic ret = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(response.Content);
                        DateTime today = DateTime.Today;
                        Console.WriteLine($"{today} - Erro: {ret.error.message.value}");
                        break;
                }

            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }
        public static bool ValidateServerCertificate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors)
        {
            return true;
        }
    }
}
