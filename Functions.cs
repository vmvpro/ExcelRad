using System;
using System.Linq;
using System.Runtime.CompilerServices;

//[assembly: InternalsVisibleToAttribute("ExcelReadTests")]
namespace ExcelRead
{
    
    public static class Functions
    {

        #region "    Методы преобразования полей таблицы Excel    "
        
        /// <summary>
        /// Метод преобразования едениц измерения
        /// </summary>
        /// <param name="ei"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Расшифровка кодирвки едениц измерения из Серийки в ИТ
        /// </summary>
        /// <param name="ei"></param>
        /// <returns></returns>
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
                string priceString = price_.Replace('.', ',');
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
                string sumString = sum_.Replace('.', ',');
                sum = Convert.ToDecimal(sumString);
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
                count = Convert.ToInt32(count_);
            }
            catch (Exception)
            {
                count = 0;
            }

            return count;
        }

        #endregion

        //-------------------------------------------------------------------

        /// <summary>
        /// Конвертирование ресурса в 15-значный код
        /// </summary>
        /// <param name="kmat_old">Старый ресурс (Excel)</param>
        /// <param name="ceh">Склад (1301, 23008)</param>
        /// <returns></returns>
        internal static string RenameOldResourceInNew(string kmat_old, string ceh)
        {
            #region "    ConvertCeh    "

            string ceh_convert = "";
            ceh_convert = ConvertCeh(ceh, kmat_old);

            #endregion

            #region "    RenameOldResourceInNew    "

            string kmat = "";
            return kmat = _renameOldResourceInNew("920", kmat_old, ceh_convert);

            #endregion
        }

        internal static string ConvertCeh(string ceh, string old_kmat_str)
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

        internal static string CreateNewResource(string ceh, int counter)
        {
            string str_counter = counter.ToString();
            int len_counter = str_counter.Length;
            int len_ceh = ceh.Length;

            string kmat = "920" + "vmv" + ceh + new String('0', 4 - str_counter.Length) + str_counter;

            return kmat;

        }

        private static string _renameOldResourceInNew(string groupResource, string old_kmat_str, string ceh_convert)
        {
            //int maxLenSymbolsOldResource = 12;



            //if (old_kmat_str.Length == 15)
            //{
            //    return old_kmat_str;
            //}

            //if (old_kmat_str.Length >= maxLenSymbolsOldResource)
            //{
            //    return old_kmat_str.Substring(old_kmat_str.Length - maxLenSymbolsOldResource);
            //}
            //else
            //{
            //    int countSymbols = maxLenSymbolsOldResource - old_kmat_str.Length;
            //    return new String('0', countSymbols) + old_kmat_str;
            //}

            //-----------------------

            string kmat = null;
            int len_ceh = ceh_convert.Length;

            int len_kmat_old = old_kmat_str.Count();

            if (old_kmat_str.Length == 15)
            {
                return old_kmat_str;
            }
            else if (len_kmat_old >= 12)
            {
                kmat = groupResource + old_kmat_str.Substring(len_kmat_old - 12);
            }
            else if (old_kmat_str.Count() == 11)
            {
                kmat = groupResource + "0" + old_kmat_str;
            }
            else if (old_kmat_str.Count() == 10)
            {
                kmat = groupResource + "00" + old_kmat_str;
            }
            else if (old_kmat_str.Count() == 9)
            {
                kmat = groupResource + ceh_convert.Substring(len_ceh - 3) + old_kmat_str;
            }
            else if (old_kmat_str.Count() == 8)
            {
                kmat = groupResource + ceh_convert.Substring(len_ceh - 4) + old_kmat_str;
            }
            else
            {
                int count_kmat_old = 12 - ceh_convert.ToString().Count() - old_kmat_str.Count();
                kmat = "920" + ceh_convert.ToString() + new String('0', count_kmat_old) + old_kmat_str;   // 3 + 4 + 1 + 7
            }

            return kmat;

        }
    }
}
