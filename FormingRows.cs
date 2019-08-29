using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace ExcelReadC
{
    public class FormingRows
    {
        public static void Main2(string path, string fileName)
        {
            string stringFormat = "{0}\t{1}      {2}     \t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}";
            
            Dictionary<string, string> KSM = new Dictionary<string, string>();

            DateTime DT = new DateTime(2017, 11, 1);

            int counter = 0;
            
            //------------------------------------

            int len_ = 15%10;
            int max_ = 12;

            string extension = "xlsx";

            string[] stringSeparator = new string[] { " - " };
            string[] result;

            result = fileName.Split(stringSeparator, StringSplitOptions.RemoveEmptyEntries);

            string ceh_ = result[1].Substring(0, 1) + result[1].Substring(2, 3);
            string n_kdk_file = result[2];

            string kmat_s2 = "1234567890123";
            string kmat_replace = kmat_s2.Replace("-", "");
            int len_kmat_s2 = kmat_replace.Count();


            string pathFullName = Path.Combine(path, fileName.Trim());

            DataTable dtExcel = Functions.ImportDataForExcel(pathFullName);
            List<string> listFieldResource = Functions.ListFieldKmatForExcel(dtExcel, "kmat");

            var dictionary = Functions.DictionaryResourceAndCount(listFieldResource);

            List<string> listResourceUnique = Functions.ListUniqueFieldResource(dictionary);

            string name = dtExcel.Rows[0]["n_kdk"].ToString();

            

            RowTXT rowTXT = new RowTXT();
            rowTXT.CloneTable(dtExcel);

            List<string> DoubleKmat = new List<string>();
            string n_kdk = "";

            try
            {
                int k = 0;
                //foreach (DataRow row in dtExcel.Rows)
                for (int i = 0; i < dtExcel.Rows.Count; i++)
                //    Console.WriteLine(String.Format("{0}", row.Field<string>("ceh").PadRight(11)));D:\Doc\Work\MS Visual Studio\1_MyApplication\ExcelReadC\ExcelReadC (Git)\FormingRows.cs
                {

                    k++;
                    try
                    {
                        string ceh_s = dtExcel.Rows[i]["ceh"].ToString();
                        string kmat_old = dtExcel.Rows[i]["kmat"].ToString().Trim();
                        string kmat_ = listResourceUnique[i]; //dtExcel.Rows[i]["kmat"].ToString().Trim();
                        //string kmat = Functions.ConvertKmat(kmat_old, ceh_s, DoubleKmat);
                        string kmat = Functions.ConvertKmatTest(kmat_, ceh_s);

                        bool flag1 = false;
                        object[] arrayColumn = dtExcel.Rows[i].ItemArray;

                        if (Functions.Flag(arrayColumn)) break;

                        int ceh = Convert.ToInt32(ceh_s);

                        n_kdk = dtExcel.Rows[i]["n_kdk"].ToString();
                        string naim = dtExcel.Rows[i]["naim"].ToString();
                        string size_type = dtExcel.Rows[i]["size_type"].ToString();
                        int ei = Functions.FuncEI(dtExcel.Rows[i]["ei"].ToString());
                        decimal price = Functions.FuncPrice(dtExcel.Rows[i]["price"].ToString());
                        decimal count = Functions.FuncCount(dtExcel.Rows[i]["count"].ToString());
                        decimal sum = Functions.FuncSum(dtExcel.Rows[i]["sum"].ToString());

                        try
                        {
                            if (!DoubleKmat.Contains(kmat_) | kmat_ == "")
                            {
                                DoubleKmat.Add(kmat_);
                                
                                //KSM.Add(rowKmat["kmat"].ToString(), rowKmat["kmat"].ToString());
                                //Console.WriteLine(String.Format("{0}\t{1}\t{2}\t  {3}\t{4}\t{5}\t{6}\t{7}", k, ceh, kmat, kmat_old, ei, price, count, sum));

                                //int kk = 0;
                                //var rows = dt.Select("kmat = '" + kmat_old + "' ");

                                DataRow[] rows1 = null;
                                try
                                {
                                    rows1 = dtExcel.Select("kmat = '" + kmat_old + "' and naim = '" + naim + "'");
                                }
                                catch (Exception)
                                {
                                    rows1 = dtExcel.Select("kmat = " + kmat_old + " and naim = '" + naim + "'");
                                }


                                int flag = 0;

                                if (rows1.Count() > 1)
                                {
                                    flag = 1;
                                    for (int iRow1 = 0; iRow1 < rows1.Count(); iRow1++)
                                    {
                                        DataRow r1 = rows1[iRow1];

                                        naim = r1["naim"].ToString();
                                        size_type = r1["size_type"].ToString();
                                        ei = Functions.FuncEI(r1["ei"].ToString());
                                        price = Functions.FuncPrice(r1["price"].ToString());
                                        count = Functions.FuncCount(r1["count"].ToString());
                                        sum = Functions.FuncSum(r1["sum"].ToString());

                                        if (iRow1 == 0)
                                        {
                                            string ss1 = rows1[0].ItemArray[2].ToString();
                                            string ss2 = rows1[0].ItemArray[3].ToString();

                                            string sss = rows1.ToString();
                                            //if (!KsmTable.IsRecord(kmat))
                                            //{
                                            //InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);
                                            //InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                            KSM.Add(kmat, ss2);
                                            Console.WriteLine(String.Format(stringFormat, k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                            //}
                                        }
                                        else
                                        {
                                            //InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                            Console.WriteLine(String.Format(stringFormat, k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                        }

                                    }
                                }

                                //var rows2 = dt.Select("kmat ='" + kmat_old + "'");

                                DataRow[] rows2 = null;
                                try
                                {
                                    rows2 = dtExcel.Select("kmat = '" + kmat_old + "'");
                                }
                                catch (Exception)
                                {
                                    rows2 = dtExcel.Select("kmat = " + kmat_old);
                                }

                                if (flag == 0)
                                {
                                    if (rows2.Count() > 1)
                                    {
                                        for (int iRow2 = 0; iRow2 < rows2.Count(); iRow2++)
                                        {
                                            DataRow r2 = rows2[iRow2];

                                            naim = r2["naim"].ToString();
                                            size_type = r2["size_type"].ToString();
                                            ei = Functions.FuncEI(r2["ei"].ToString());
                                            price = Functions.FuncPrice(r2["price"].ToString());
                                            count = Functions.FuncCount(r2["count"].ToString());
                                            sum = Functions.FuncSum(r2["sum"].ToString());

                                            if (iRow2 == 0)
                                            {
                                                //naim = r2["naim"].ToString();
                                                //size_type = r2["size_type"].ToString();

                                                //if (!KsmTable.IsRecord(kmat))
                                                //{
                                                //InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);
                                                //InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                                KSM.Add(kmat, naim);
                                                Console.WriteLine(String.Format(stringFormat, k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                                //}
                                            }
                                            else
                                            {
                                                counter++;
                                                kmat = Functions.ConvertKmatTest("", ceh_s);

                                                //if (!KsmTable.IsRecord(kmat))
                                                //{
                                                //InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);
                                                KSM.Add(kmat, naim);
                                                //}
                                                //InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                                Console.WriteLine(String.Format(stringFormat, k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                            }

                                        }
                                    }
                                    else
                                    {

                                        //string ss1 = rows2[0].ItemArray[2].ToString();
                                        //string ss2 = rows2[0].ItemArray[3].ToString();
                                        //if (!KsmTable.IsRecord(kmat))
                                        //{
                                        //InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);

                                        if (kmat_old == "")
                                        {
                                            counter++;
                                            kmat = Functions.ConvertKmatTest("", ceh_s);
                                        }

                                        KSM.Add(kmat, naim);
                                        //}

                                        //InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                                        Console.WriteLine(String.Format(stringFormat, k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                                    }
                                }
                            }
                            else
                            {
                                //KSM.Add(kmat, naim);
                                Console.WriteLine(String.Format(stringFormat, k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));
                            }

                        }
                        catch (Exception ex)
                        {
                            //if (!KsmTable.IsRecord(kmat))
                            //{
                            //InsertKmat(kmat, kmat_old, naim, size_type, Convert.ToInt32(ei), fileName, BS);

                            //---------------------------------------------------------
                            if (kmat_old == "")
                            {
                                counter++;
                                kmat = Functions.ConvertKmatTest("", ceh_s);
                            }

                            KSM.Add(kmat, naim);
                            //}

                            //InsertBalanceDMS(undoc, ndm_s, ceh, n_kdk, kmat, Convert.ToInt32(ei), count, price, sum, BS);
                            Console.WriteLine(String.Format(stringFormat, k, ceh, kmat, kmat_old.PadRight(12), naim.PadRight(30), size_type.PadRight(20), ei, price, count, sum));

                            //---------------------------------------------------------

                            //throw new Exception(ex.Message);
                            //DoubleKmat.Add(kmat_old);

                            //var rowsKmat = dt.AsEnumerable().Where(x => x["kmat"].ToString() == kmat_old);

                            //foreach (var rowKmat in rowsKmat)
                            //    rowTXT.Add(rowKmat, dt.Rows.IndexOf(rowKmat) + 2);

                        }

                    }
                    catch (Exception ex)
                    {
                        throw new Exception(ex.Message);

                        bool flag = false;
                        object[] arrayColumn = dtExcel.Rows[i].ItemArray;

                        for (int i2 = 0; i2 < arrayColumn.Length - 1; i2++)
                            flag = arrayColumn[i2].ToString() == "" ? true : false;

                        if (flag) break;

                        rowTXT.Add(dtExcel.Rows[i], k + 1);

                    }

                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);

            }

            Console.WriteLine("-------------------------------------------------");

            //foreach (var ss in KSM)
            //Console.WriteLine(ss.Key); 

            Console.WriteLine();

            decimal sumCount = dtExcel.Select("n_kdk = '" + n_kdk.ToString() + "'").Sum(x => Convert.ToDecimal(x["count"]));
            //sumCountString = sumCount.ToString();

            Console.WriteLine("Сумма количество = " + sumCount);
            //rowTXT.WriteTXT(path, fileName, n_kdk_file);

            Console.ReadLine();

        }
    }
}
