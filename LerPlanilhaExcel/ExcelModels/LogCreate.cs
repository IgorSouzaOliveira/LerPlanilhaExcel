using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LerPlanilhaExcel.ExcelModels
{
    internal class LogCreate
    {
        static String directory = AppDomain.CurrentDomain.BaseDirectory;
        static String pathDirectory = "LOG";
        static String path = Path.Combine(directory, $@"LOG\LOG_{DateTime.Today.ToString("ddMMyyyy")}.txt");

        public static void Log(string swLog)
        {
            try
            {
                if (!File.Exists(pathDirectory))
                {
                    Directory.CreateDirectory(pathDirectory);
                }

                using (var file = File.AppendText(path))
                {
                    file.WriteLine(swLog);

                }
            }
            catch (IOException ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
