using ExcelRead;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace ExcelReadTests
{
    public static class DataTests
    {

        //FileInfo file = new FileInfo().fu

        static string currentDirectory = Directory.GetCurrentDirectory();
        static string fullPath = @"..\..\DataTests\001 - 1301 - 1222333 - LastName - test_01.xlsx";

        public static string GetFullPathExcelFile
        {
            get { return Path.GetFullPath(Path.Combine(currentDirectory, fullPath)); }
        }

        public static DataTable mockLoadDataTableForExcel()
        {

            oExcel excel = new oExcel(GetFullPathExcelFile);


            return
                   excel.ImportDataForExcel();
            //Functions.ImportDataForExcel();
        }



        public static DataTable mockLoadData()
        {
            DataTable dt = new DataTable();

            DataColumn dc1 = new DataColumn("ceh", DbType.String.GetType());
            DataColumn dc2 = new DataColumn("n_kdk", DbType.String.GetType());
            DataColumn dc3 = new DataColumn("kmat", DbType.String.GetType());
            DataColumn dc4 = new DataColumn("naim", DbType.String.GetType());
            DataColumn dc5 = new DataColumn("size_type", DbType.String.GetType());
            DataColumn dc6 = new DataColumn("ei", DbType.Int32.GetType());
            DataColumn dc7 = new DataColumn("price", DbType.Decimal.GetType());
            DataColumn dc8 = new DataColumn("count", DbType.Int32.GetType());
            DataColumn dc9 = new DataColumn("sum", DbType.Decimal.GetType());

            dt.Columns.AddRange(new DataColumn[] { dc1, dc2, dc3, dc4, dc5, dc6, dc7, dc8, dc9 });

            return null;
        }

        public static Dictionary<string, int> DictionaryGroupBy()
        {
            Dictionary<string, int> expectedDictionaryGroupBy =
                new Dictionary<string, int>
                {
                    {"1100109",3 },
                    {"1100184",2 },
                    {"1100277",2 },
                    {"1100337",1 },
                    {"1100418",1 },
                    {"1100603",1 },
                    {"1000004",1 },
                    {"80030480", 2 },
                    {"111112345678900", 1},
                    {"2345678900",5 }
                    
                    
                };

            return expectedDictionaryGroupBy;
        }
    }
}
