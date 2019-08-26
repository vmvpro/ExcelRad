using ExcelReadC;
using System.Data;
using System.IO;
using System.Text;

namespace ExcelReadCTests
{
    public static class DataTests
    {

        public static DataTable mockLoadDataTableForExcel()
        {
            StringBuilder pathFullName = new StringBuilder();

            string currentDirectory = Directory.GetCurrentDirectory();
            string pathFile = @"..\..\DataTests\001 - 1301 - 1222333 - LastName - test_01.xlsx";

            return 
                Functions.ImportDataForExcel(Path.GetFullPath(Path.Combine(currentDirectory, pathFile)));
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
    }
}
