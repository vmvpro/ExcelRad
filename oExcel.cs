using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;

namespace ExcelRead
{

    public class oExcel
    {
        string pathFullName;
        DataTable dtExcel;

        List<string> listField;
        Dictionary<string, int> groupByFieldAndCount;

        public oExcel(string pathFullName)
        {
            this.pathFullName = pathFullName;

            dtExcel = new DataTable();

            listField = new List<string>();
        }

        // 1
        /// <summary>
        /// Импортирование таблицы Excel 
        /// </summary>
        /// <param name="sheetName">Имя листа</param>
        /// <returns></returns>
        public DataTable ImportDataForExcel(string sheetName = "Sheet")
        {
            dtExcel = new DataTable(sheetName);

            string connectionString;

            //'Для Excel 12.0 
            connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0; 
                Data Source = " + pathFullName + "; " +
                "Extended Properties=\"Excel 12.0 Xml;HDR=Yes\";";

            using (var connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                OleDbCommand command = connection.CreateCommand();

                command.CommandText = "Select * From [" + sheetName + "$A0:I15000] ";

                OleDbDataAdapter da = new OleDbDataAdapter(command);

                da.Fill(dtExcel);
            }

            return dtExcel;
        }

        // 2.1
        /// <summary>
        /// Удаление специальных символов в строке (' ', ',', '-', '.', '+')
        /// </summary>
        /// <param name="valueString"></param>
        /// <returns></returns>
        public string DeleteSpecialCharacters(string valueString)
        {
            int result;
            char[] chars = new char[] { ' ', ',', '-', '.', '+' };

            string stringValue = String.Join("", valueString.Split(chars));

            if (Int32.TryParse(stringValue, out result))
                return result.ToString();

            return stringValue;
        }

        // 2.1
        /// <summary>
        /// Сформировать список по полю таблицы 
        /// </summary>
        /// <param name="fieldName">Имя поля на листе</param>
        /// <returns></returns>
        public List<string> ListField(string fieldName = "kmat")
        {
            listField = new List<string>();

            foreach (DataRow row in dtExcel.Rows)
            {
                string rowString = row[fieldName].ToString().Trim();

                string cellCeh = row["ceh"].ToString().Trim();

                if (cellCeh == "") break;

                if (fieldName == "kmat")
                {
                    int lenSymbols = rowString.Length;

                    if (lenSymbols >= 11 && lenSymbols < 15)
                    {
                        int diff = lenSymbols % 10;

                        rowString = rowString.Substring(diff);
                    }

                    listField.Add(DeleteSpecialCharacters(rowString));
                }
                else
                    listField.Add(DeleteSpecialCharacters(rowString));
            }

            return listField;
        }

        // 3
        /// <summary>
        /// Группировка: имя поля и количество повторений
        /// </summary>
        private void _groupByFieldAndCount()
        {
            groupByFieldAndCount = new Dictionary<string, int>();

            IEnumerable<IGrouping<string, string>> listGroupBy = listField.GroupBy(x => x);

            foreach (IGrouping<string, string> item in listGroupBy)
                groupByFieldAndCount.Add(item.Key, item.Count());

        }

        // 4
        /// <summary>
        /// Формирование списка по полю уникальных значений
        /// </summary>
        /// <returns></returns>
        public List<string> ListUniqueField()
        {
            _groupByFieldAndCount();

            // Для того, чтобы предусмотреть не повторающие значения в списке
            HashSet<string> listUniqu = new HashSet<string>();

            foreach (KeyValuePair<string, int> row in groupByFieldAndCount)
                for (int i = 1; i < row.Value + 1; i++)
                    if (row.Value > 1)
                        listUniqu.Add(i + row.Key);
                    else
                        listUniqu.Add(row.Key);

            return listUniqu.ToList();
        }
    }
}
