using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelReadC
{
    public static class Functions
    {

        #region "    Методы преобразования полей таблицы Excel    "
        
        public static bool Flag(object[] arrayColumn)
        {
            bool result2 = false;
            for (int i = 0; i < arrayColumn.Length - 1; i++)
            {
                if (arrayColumn[i].ToString() == "")
                {
                    result2 = true;
                }
                else
                    return false;
            }

            return result2;
        }

        public static int FuncEI(string ei)
        {
            try
            {
                if (ei == "")
                    return Convert.ToInt32("796");
                else
                {
                    return ConvertEI(ei);
                }
            }
            catch (Exception)
            {
                return 796;
            }
        }

        public static int ConvertEI(string ei)
        {
            switch (ei)
            {
                case "1": return 839;
                case "2": return 796;
                case "3": return 166;
                case "4": return 163;
                case "5": return 6;
                case "6": return 761;
                case "7": return 168;
                case "8": return 798;
                case "9": return 797;
                case "10": return 112;
                case "11": return 736;
                case "796": return 796;
            }

            return 796;
        }

        public static decimal FuncPrice(string price_)
        {
            decimal price = 0;
            try
            {
                price = Convert.ToDecimal(price_);
            }
            catch (Exception)
            {
                price = 0;
            }

            return price;
        }

        public static decimal FuncSum(string sum_)
        {
            decimal sum = 0;
            try
            {
                sum = Convert.ToDecimal(sum_);
            }
            catch (Exception)
            {
                sum = 0;
            }

            return sum;
        }

        public static decimal FuncCount(string count_)
        {
            decimal count = 0;
            try
            {
                count = Convert.ToDecimal(count_);
            }
            catch (Exception)
            {
                count = 0;
            }

            return count;
        }

        #endregion

        public static Dictionary<string, int> DictionaryLoadedDataExcelResource(DataTable dt)
        {
            return null;
        }

        public static DataTable ImportDataForExcel(string path, string fileName)
        {
            DataTable dt = new DataTable("SheetExcel");

            string connectionString;
            OleDbConnection connection;

            //'Для Excel 12.0 
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + path + fileName + "; Extended Properties=\"Excel 12.0 Xml;HDR=Yes\";";
            connection = new OleDbConnection(connectionString);
            connection.Open();

            OleDbCommand command = connection.CreateCommand();

            command.CommandText = "Select * From [sheet$A0:I15000] "; 

            var da = new OleDbDataAdapter(command);
            dt = new DataTable();

            da.Fill(dt);

            return dt;
        }


        public static string ConvertKmat(string kmat_old, string ceh, List<string> DoubleKmat)
        {
            string kmat = "";
            string ceh_convert = "";
            int count_kmat_old = 0;

            string old_kmat_str = "";
            try
            {
                string old_kmat_convert = kmat_old.Replace(" ", "").Replace(",", "").Replace("-", "").Replace(".", "").Replace("+", "");    //00123456
                old_kmat_str = Convert.ToInt32(old_kmat_convert).ToString();   // 8
            }
            catch (Exception)
            {
                string old_kmat_convert = kmat_old.Replace(" ", "").Replace(",", "").Replace("-", "").Replace(".", "").Replace("+", "");    //00123456
                old_kmat_str = old_kmat_convert;
            }

            if (!DoubleKmat.Contains(kmat_old) || kmat_old == "")
            {

            }

            if (ceh.Count() < 6 && old_kmat_str.Count() <= 7)
                ceh_convert = ceh;
            else if (ceh.Count() > 4)
                ceh_convert = ceh.ToString().Substring(0, 1) + ceh.ToString().Substring(2, 3);
            else
                ceh_convert = ceh;

            int len = ceh.Count();

            if (kmat_old == "" || DoubleKmat.Contains(kmat_old))
            {

                string str_counter = "";//counter.ToString();
                int len_counter = str_counter.Length;
                int len_ceh = ceh.Length;

                kmat = "920" + "vmv" + ceh + new String('0', 4 - str_counter.Length) + str_counter;

                return kmat;
            }

            int len_kmat_old = old_kmat_str.Count();
            if (len_kmat_old >= 12 && !DoubleKmat.Contains(kmat_old))
            {
                kmat = "920" + old_kmat_str.Substring(len_kmat_old - 12, 12);
            }
            else if (old_kmat_str.Count() == 11)
            {
                kmat = "920" + "0" + old_kmat_str;
            }
            else if (old_kmat_str.Count() == 10)
            {
                kmat = "920" + "00" + old_kmat_str;
            }
            else if (old_kmat_str.Count() == 9)
            {
                kmat = "920" + ceh.Substring(len - 3, 3) + old_kmat_str;
            }
            else if (old_kmat_str.Count() == 8)
            {
                kmat = "920" + ceh_convert + old_kmat_str;
            }
            else
            {
                count_kmat_old = 12 - ceh_convert.ToString().Count() - old_kmat_str.Count();
                kmat = "920" + ceh_convert.ToString() + new String('0', count_kmat_old) + old_kmat_str;   // 3 + 4 + 1 + 7
            }

            return kmat;
        }

        public static string ConvertOldResource(string kmat_old)
        {
            int result;
            char[] chars = new char[] { ' ', ',', '-', '.', '+' };

            var convertOldResource = String.Join("", kmat_old.Split(chars));

            if (Int32.TryParse(convertOldResource, out result))
                return result.ToString();
            
            return convertOldResource;
        }

        public static string ConvertKmatTest(string kmat_old, string ceh, List<string> DoubleKmat)
        {
            string kmat = "";
            string ceh_convert = "";
            int count_kmat_old = 0;

            string old_kmat_str = "";   
            try
            {
                string old_kmat_convert = kmat_old.Replace(" ", "").Replace(",", "").Replace("-", "").Replace(".", "").Replace("+", "");    //00123456
                old_kmat_str = Convert.ToInt32(old_kmat_convert).ToString();   // 8
            }
            catch (Exception)
            {
                string old_kmat_convert = kmat_old.Replace(" ", "").Replace(",", "").Replace("-", "").Replace(".", "").Replace("+", "");    //00123456
                old_kmat_str = old_kmat_convert;
            }

            if (!DoubleKmat.Contains(kmat_old) || kmat_old == "")
            {

            }

            if (ceh.Count() < 6 && old_kmat_str.Count() <= 7)
                ceh_convert = ceh;
            else if (ceh.Count() > 4)
                ceh_convert = ceh.ToString().Substring(0, 1) + ceh.ToString().Substring(2, 3);
            else
                ceh_convert = ceh;

            int len = ceh.Count();

            //--------------------------------------------------------

            if (kmat_old == "" || DoubleKmat.Contains(kmat_old))
            {
                //return CreateNewKmat(ceh, counter);
                return CreateNewKmat(ceh, 0);
            }

            int len_kmat_old = old_kmat_str.Count();
            if (len_kmat_old >= 12 && !DoubleKmat.Contains(kmat_old))
            {
                kmat = "920" + old_kmat_str.Substring(len_kmat_old - 12, 12);
            }
            else if (old_kmat_str.Count() == 11)
            {
                kmat = "920" + "0" + old_kmat_str;
            }
            else if (old_kmat_str.Count() == 10)
            {
                kmat = "920" + "00" + old_kmat_str;
            }
            else if (old_kmat_str.Count() == 9)
            {
                kmat = "920" + ceh.Substring(len - 3, 3) + old_kmat_str;
            }
            else if (old_kmat_str.Count() == 8)
            {
                kmat = "920" + ceh_convert + old_kmat_str;
            }
            else
            {
                count_kmat_old = 12 - ceh_convert.ToString().Count() - old_kmat_str.Count();
                kmat = "920" + ceh_convert.ToString() + new String('0', count_kmat_old) + old_kmat_str;   // 3 + 4 + 1 + 7
            }

            return kmat;
        }

        public static string CreateNewKmat(string ceh, int counter)
        {
            string str_counter = counter.ToString();
            int len_counter = str_counter.Length;
            int len_ceh = ceh.Length;

            string kmat = "920" + "vmv" + ceh + new String('0', 4 - str_counter.Length) + str_counter;

            return kmat;


        }
    }
}
