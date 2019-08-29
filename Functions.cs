using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;

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
                    result2 = true;
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

        // 1
        public static DataTable ImportDataForExcel(string pathFullName)
        {
            DataTable dt = new DataTable("Sheet");

            string connectionString;
            OleDbConnection connection;

            //'Для Excel 12.0 
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + pathFullName + "; Extended Properties=\"Excel 12.0 Xml;HDR=Yes\";";
            connection = new OleDbConnection(connectionString);
            connection.Open();

            OleDbCommand command = connection.CreateCommand();

            command.CommandText = "Select * From [sheet$A0:I15000] ";

            OleDbDataAdapter da = new OleDbDataAdapter(command);
            dt = new DataTable();

            da.Fill(dt);

            return dt;
        }

        // 2.1
        public static string ConvertOldResource(string kmat_old)
        {
            int result;
            char[] chars = new char[] { ' ', ',', '-', '.', '+' };

            string convertOldResource = String.Join("", kmat_old.Split(chars));

            if (Int32.TryParse(convertOldResource, out result))
                return result.ToString();

            return convertOldResource;
        }

        // 2.1
        public static List<string> ListFieldKmatForExcel(DataTable dtExcel, string fieldName)
        {
            List<string> list = new List<string>();


            foreach (DataRow row in dtExcel.Rows)
            {
                string rowString = row[fieldName].ToString().Trim();

                string cellCeh = row["ceh"].ToString().Trim();

                if (cellCeh == "") break; 

                if (fieldName == "kmat")
                {
                    int lenSymbols = rowString.Length;

                    if (lenSymbols >= 11 && lenSymbols <= 15)
                    {
                        int diff = lenSymbols % 10;

                        rowString = rowString.Substring(diff, lenSymbols - diff);
                    }

                    list.Add(ConvertOldResource(rowString));
                }
                else
                    list.Add(ConvertOldResource(rowString));
            }

            return list;
        }

        // 3
        public static Dictionary<string, int> DictionaryResourceAndCount(List<string> listField)
        {
            Dictionary<string, int> dic = new Dictionary<string, int>();

            IEnumerable<IGrouping<string, string>> listGroupBy = listField.GroupBy(x => x);

            foreach (IGrouping<string, string> item in listGroupBy)
                dic.Add(item.Key, item.Count());

            return dic;
        }

        // 4
        public static List<string> ListUniqueFieldResource(Dictionary<string, int> dicResourcesAndCount)
        {
            // Для того, чтобы предусмотреть не повторающие значения в списке
            HashSet<string> listUniqu = new HashSet<string>();

            foreach (KeyValuePair<string, int> row in dicResourcesAndCount)
                for (int i = 1; i < row.Value + 1; i++)
                    if (row.Value > 1)
                        listUniqu.Add(i + row.Key);
                    else
                        listUniqu.Add(row.Key);

            return listUniqu.ToList();
        }

        //-------------------------------------------------------------------

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

        public static string ConvertKmatTest(string kmat_old, string ceh)
        {
            string kmat = "";
            string ceh_convert = "";
            int count_kmat_old = 0;

            string old_kmat_str = "";

            old_kmat_str = Functions.ConvertOldResource(kmat_old);

            // -----   Convert Ceh   ---------

            ceh_convert = ConvertCeh(ceh, old_kmat_str);

            int len_ceh = ceh_convert.Length;

            //--------------------------------------------------------

            int len_kmat_old = old_kmat_str.Count();

            #region "   CreateNewResource   "

            if (kmat_old == "")
            {
                //return CreateNewKmat(ceh, counter);
                return CreateNewResource(ceh, 0);
            }

            #endregion

            #region "    RenameOldResourceInNew    "


            if (len_kmat_old >= 12)
            {
                kmat = "920" + old_kmat_str.Substring(len_kmat_old - 12);
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
                kmat = "920" + ceh.Substring(len_ceh - 3) + old_kmat_str;
            }
            else if (old_kmat_str.Count() == 8)
            {
                kmat = "920" + ceh.Substring(len_ceh - 4) + old_kmat_str;
            }
            else
            {
                count_kmat_old = 12 - ceh_convert.ToString().Count() - old_kmat_str.Count();
                kmat = "920" + ceh_convert.ToString() + new String('0', count_kmat_old) + old_kmat_str;   // 3 + 4 + 1 + 7
            }

            return kmat;

            #endregion
        }

        public static string ConvertCeh(string ceh, string old_kmat_str)
        {
            string ceh_convert = "";

            if (ceh.Count() < 6 && old_kmat_str.Count() <= 7)
            {
                ceh_convert = ceh;
            }

            else if (ceh.Count() > 5)
            {
                ceh_convert = ceh.ToString().Substring(ceh.Length - 5); //4
            }
            else
                ceh_convert = ceh;


            return ceh_convert;


        }

        public static string CreateNewResource(string ceh, int counter)
        {
            string str_counter = counter.ToString();
            int len_counter = str_counter.Length;
            int len_ceh = ceh.Length;

            string kmat = "920" + "vmv" + ceh + new String('0', 4 - str_counter.Length) + str_counter;

            return kmat;

        }

        public static string RenameOldResourceInNew(string old_kmat)
        {
            int maxLenSymbolsOldResource = 12;

            if (old_kmat.Length >= maxLenSymbolsOldResource)
            {
                return old_kmat.Substring(old_kmat.Length - maxLenSymbolsOldResource);
            }
            else
            {
                int countSymbols = maxLenSymbolsOldResource - old_kmat.Length;
                return new String('0', countSymbols) + old_kmat;
            }

        }
    }
}
