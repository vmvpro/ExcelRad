using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

[assembly: InternalsVisibleToAttribute("ExcelReadTests")]
namespace ExcelRead
{
    public class Program
    {
        private static void Main()
        {
            //string path = @"\\erp\TEMP\App\Остатки\ЛИиДБ\";
            string path = Directory.GetCurrentDirectory().Trim();

            string fileName = @"
            011 - 23010 - 1010134 - vmv
            ".Trim();

            FormingRows.Run(path, fileName);
            
        }
       
    }

    
}
