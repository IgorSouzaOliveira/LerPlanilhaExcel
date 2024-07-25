using ClosedXML.Excel;
using LerPlanilhaExcel.Common;
using LerPlanilhaExcel.ExcelModels;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace LerPlanilhaExcel.Models
{
    public class ExcelModel
    {
        public static void ReadExcel()
        {
            try
            {

                Console.WriteLine($"{DateTime.Now} - Iniciando processo de atualização.\n");
                LogCreate.Log($"{DateTime.Now} - Iniciando processo de atualização.\n");

                BusinessPartner oBP = new BusinessPartner();

                var directy = AppDomain.CurrentDomain.BaseDirectory;
                var path = Path.Combine(directy, @"Arquivo\\BusinessPartners.xlsx");

                var xls = new XLWorkbook(path);
                var planilha = xls.Worksheets.First(w => w.Name == "Planilha1");
                var totalLinhas = planilha.Rows().Count();

                Console.WriteLine($"{DateTime.Now} - Quantidade de cadastro(s): {totalLinhas - 1}.\n");
                LogCreate.Log($"{DateTime.Now} - Quantidade de cadastro(s): {totalLinhas - 1}.\n");

                for (int i = 2; i <= totalLinhas; i++)
                {

                    oBP.CardCode = planilha.Cell($"A{i}").Value.ToString();
                    oBP.U_LGO_Rota = planilha.Cell($"B{i}").Value.ToString();
                    oBP.PayTermsGrpCode = planilha.Cell($"C{i}").Value.ToString();
                    oBP.PriceListNum = planilha.Cell($"D{i}").Value.ToString();
                    oBP.SalesPersonCode = (int)planilha.Cell($"E{i}").Value;
                    
                    int line = i - 2 ;
                    SAPConnect.UpdateBP(oBP, line);

                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{DateTime.Now} - {ex.Message}");
                LogCreate.Log($"{DateTime.Now} - {ex.Message}");
            }
            finally
            {
                Console.WriteLine($"{DateTime.Now} - Processo Finalizado.");
                LogCreate.Log($"{DateTime.Now} - Processo Finalizado.");
            }
        }
    }
}
