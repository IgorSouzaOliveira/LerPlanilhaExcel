using ClosedXML.Excel;
using LerPlanilhaExcel.Common;
using LerPlanilhaExcel.ExcelModels;
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
                    oBP.EmailAddress = planilha.Cell($"B{i}").Value.ToString();
                    oBP.Phone1 = planilha.Cell($"C{i}").Value.ToString();
                    oBP.Phone2 = planilha.Cell($"D{i}").Value.ToString();
                    int line = i - 2 ;
                    SAPConnect.UpdateBP(oBP, line);

                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"{DateTime.Now} - {ex.Message}");
                LogCreate.Log($"{DateTime.Now} - {ex.Message}");
            }
        }
    }
}
