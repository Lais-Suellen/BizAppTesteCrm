using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using System.Net;
using Microsoft.Xrm.Tooling.Connector;
using System.Configuration;
using ClosedXML.Excel;




namespace BizAppTesteCrm
{
    class Program
    {
        private static CrmServiceClient crmServiceClientDestino;

        static void Main(string[] args)
        {
            Console.WriteLine("Validating Connection...");
            IOrganizationService service = getCRMService();
            Console.WriteLine("Connection Successful!");

            var xls = new XLWorkbook(@"C:\BizAppCrm\ClientesPotenciais.xlsx");
            var planilha = xls.Worksheets.First(w => w.Name == "Planilha1");
            var totalLinhas = planilha.Rows().Count();

            // primeira linha do cabecalho
            for (int l = 2; l <= totalLinhas; l++)
            {
                var subject = planilha.Cell($"A{l}").Value.ToString();
                var firstname = planilha.Cell($"B{l}").Value.ToString();
                var lastname = planilha.Cell($"C{l}").Value.ToString();
                var telephone1 = planilha.Cell($"D{l}").Value.ToString();
                var mobilephone = planilha.Cell($"E{l}").Value.ToString();
                var emailaddress1 = planilha.Cell($"F{l}").Value.ToString();
                var companyname = planilha.Cell($"G{l}").Value.ToString();


                Console.WriteLine($"{subject} - {firstname} - {lastname} - {telephone1} - {mobilephone} - {emailaddress1} - {companyname}");

                Entity lead = new Entity("lead");

                lead["subject"] = subject;
                lead["firstname"] = firstname;
                lead["lastname"] = lastname;
                lead["telephone1"] = telephone1;
                lead["mobilephone"] = mobilephone;
                lead["emailaddress1"] = emailaddress1;
                lead["companyname"] = companyname;

                
                service.Create(lead);
                Console.WriteLine($"Registro { lead.Attributes["firstname"] } criado.");


            }

            Console.WriteLine($"Importação finalizada... ");
            Console.ReadKey();

        }

        public static IOrganizationService getCRMService()
                {
                try
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    string connectionString = ConfigurationManager.ConnectionStrings["connectionStringCRM"].ConnectionString;
                    crmServiceClientDestino = new CrmServiceClient(connectionString);

                    if (crmServiceClientDestino != null)
                    {
                        return crmServiceClientDestino;
                    }
                    else
                    {
                        Console.WriteLine("Connection failed...");
                        throw new Exception(crmServiceClientDestino.LastCrmError);
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error - " + ex.ToString());
                    throw ex;
                }
         }
     }
    
}
