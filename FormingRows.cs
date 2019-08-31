using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelRead
{
    public class FormingRows
    {
        public static void Run(string path, string fileName)
        {
            string stringFormat = "{0}\t{1}      {2}     \t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}";

            Dictionary<string, string> dmsRows = new Dictionary<string, string>();

            DateTime DT = new DateTime(2017, 11, 1);

            //------------------------------------


            string[] stringSeparator = new string[] { " - " };

            string[] result = fileName.Split(stringSeparator, StringSplitOptions.RemoveEmptyEntries);

            string ceh_ = result[1];
            string n_kdk_file = result[2];

            // ------------------------------------

            string pathFullName = Path.Combine(path, fileName.Trim());

            oExcel excel = new oExcel(pathFullName);

            DataTable dtExcel = excel.ImportDataForExcel();
            List<string> listFieldResource = excel.ListField("kmat");

            // Отсортированный список с уникальными значениями ресурсов
            List<string> listResourceUnique = excel.ListUniqueField();

            string name = dtExcel.Rows[0]["n_kdk"].ToString();

            List<string> KsmTable = new List<string>();
            string n_kdk = "";

            for (int i = 0; i < listFieldResource.Count; i++)
            {

                // Форматиррование старого ресурса с листа Excel
                // из 001-123456 преобразовывает => 1123456
                string oldKmatRename = listFieldResource[i];

                // Ищет первое вхождение в списке заканчивающее на преобразованный выше ресурс
                // Идея след.:
                // в списке одинаковый код отличающейся только в начале порядковым номером
                // 1555777999
                // 2555777999
                // oldKmatRename возвращает 555777999, в этом случае находится первое попавшее 
                // значение => 1555777999
                string kmat_ = listResourceUnique.Find(x => x.EndsWith(oldKmatRename));

                // потом удаляем со списка 1555777999, и след. раз функция возвратит значение 2555777999
                // так происходит избавление от дублей
                listResourceUnique.Remove(kmat_);

                string ceh_s = dtExcel.Rows[i]["ceh"].ToString();
                string kmat_old = dtExcel.Rows[i]["kmat"].ToString().Trim();

                // конвертирование в новый ресурс 920[ceh_s][kmat_]
                string kmat = Functions.RenameOldResourceInNew(kmat_, ceh_s);

                int ceh = Convert.ToInt32(ceh_s);

                n_kdk = dtExcel.Rows[i]["n_kdk"].ToString();
                string naim = dtExcel.Rows[i]["naim"].ToString();
                string size_type = dtExcel.Rows[i]["size_type"].ToString();

                // Преобразование едениц имерения
                // Нужно смотреть по листу Excel, какая кодировка введена
                // Если кодировка 'Серийки', то необходимо использовать Functions.FuncEI
                // Если кодировка ИТ, так и оставлять
                //ei = Functions.FuncEI(dtExcel.Rows[i]["ei"].ToString());
                int ei = Convert.ToInt32(dtExcel.Rows[i]["ei"].ToString());

                decimal price = Functions.FuncPrice(dtExcel.Rows[i]["price"].ToString());
                decimal count = Functions.FuncCount(dtExcel.Rows[i]["count"].ToString());
                decimal sum = Functions.FuncSum(dtExcel.Rows[i]["sum"].ToString());

                // Прототип таблицы KSM
                if (!KsmTable.Contains(kmat_))
                    KsmTable.Add(kmat_);

                // Прототип строк DMS
                dmsRows.Add(kmat, naim);

                Console.WriteLine(String.Format(stringFormat, 
                    i + 1, ceh, kmat, kmat_old.PadRight(15), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));

            }

            Console.WriteLine("-------------------------------------------------");

            Console.WriteLine();

            decimal sumCount = dtExcel.Select("n_kdk = '" + n_kdk.ToString() + "'")
                .Sum(x => Convert.ToDecimal(x["count"]));

            Console.WriteLine("Сумма количество = " + sumCount);
            
            Console.ReadLine();

        }
    }
}
