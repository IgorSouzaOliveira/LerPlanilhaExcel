using LerPlanilhaExcel.Common;
using LerPlanilhaExcel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace LerPlanilhaExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            SAPConnect.SAPLogin();

            ExcelModel.ReadExcel();
            Thread.Sleep(5000);

        }
    }
}
