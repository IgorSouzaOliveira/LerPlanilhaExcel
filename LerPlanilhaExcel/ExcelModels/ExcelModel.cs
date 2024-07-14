using ClosedXML.Excel;
using LerPlanilhaExcel.Common;
using System;
using System.Collections.Generic;
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

                var xls = new XLWorkbook(@"C:\Users\SAPB1DEV\Desktop\_source\BusinessPartner.xlsx");
                var planilha = xls.Worksheets.First(w => w.Name == "Plan1");
                var totalLinhas = planilha.Rows().Count();

                Console.WriteLine($"Quantidade: {totalLinhas - 1}\n");

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
                Console.WriteLine(ex.Message);
            }
        }
    }
}
