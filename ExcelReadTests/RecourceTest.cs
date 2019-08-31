using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.CompilerServices;
using ExcelRead;
using Microsoft.VisualStudio.TestTools.UnitTesting;

//[assembly: InternalsVisibleToAttribute("ExcelRead")] 
//[assembly: IgnoresAccessChecksTo("ExcelRead")]
namespace ExcelReadTests
{
    
    [TestClass]
    public class RecourceTest
    {
        public TestContext TestContext { get; set; }

        int counterListFieldKmatForExcel_Test = 0;

        static DataTable dtExcel;

        static oExcel excel;

        static List<string> listField = new List<string>();
        static Dictionary<string, int> dictionaryGroupBy = new Dictionary<string, int>();
        static List<string> listUnique = new List<string>();

        //private static List<string> DoubleKmat;

        [ClassInitialize]
        public static void Initialization(TestContext context)
        {
            excel = new oExcel(DataTests.GetFullPathExcelFile);
            dtExcel = excel.ImportDataForExcel();

            listField = excel.ListField("kmat"); // Functions.ListFieldKmatForExcel(dtExcel, "kmat");
            dictionaryGroupBy = excel._groupByFieldAndCount("kmat"); //DictionaryResourceAndCount(listField);

            listUnique = excel.ListUniqueField("kmat"); //new List<string>(); // Functions.ListUniqueFieldResource(dictionaryGroupBy);

            // = DoubleKmat.ToArray();
        }

        [TestMethod]
        public void ConvertOldResource_Test()
        {
            // arrange
            string expendetOldResource = "1123456";

            //act   
            //string actualOldResource = Functions.DeleteSpecialCharacters("001-123,45 6");
            string actualOldResource = "";

            // assert
            Assert.AreEqual<string>(expendetOldResource, actualOldResource);

        }

        [TestMethod]
        public void ConvertOldResource_CollectionsTest()
        {
            // arrange
            List<string> expendetOldRecourceList =
                new List<string>(new[] { "112345", "223", "0023vmv", "555111", "1000004" });

            //act   

            List<string> OldRecourceList =
                new List<string>(new[] { "11-2345", "00223", "0023vmv", "555.11,1", "001-000004" });

            List<string> actualList = new List<string>();

            foreach (string oldRecource in OldRecourceList)
                actualList.Add("Functions.DeleteSpecialCharacters(oldRecource)");
                //actualList.Add(Functions.DeleteSpecialCharacters(oldRecource));

            // assert
            CollectionAssert.AreEqual(expendetOldRecourceList, actualList);

        }

        [TestMethod]
        [DataSource(
        "System.Data.OleDb",
        @"Provider=Microsoft.ACE.OLEDB.12.0;
            Data Source=..\..\DataTests\001 - 1301 - 1222333 - LastName - test_01.xlsx;
            Persist Security Info=False;
            Extended Properties='Excel 12.0 Xml; HDR=YES'",
        "Sheet$",
        DataAccessMethod.Sequential)]
        public void ConvertOldResource_DataContextTest()
        {
            // arrange
            List<string> expendetOldRecourceList =
                new List<string>(new[] { "112345", "223", "0023vmv", "555111" });

            //act   

            List<string> OldRecourceList =
                new List<string>(new[] { "11-2345", "00223", "0023vmv", "555.11,1" });

            List<string> actualList = new List<string>();

            foreach (string oldRecource in OldRecourceList)
                actualList.Add("Functions.DeleteSpecialCharacters(oldRecource)");
            // assert

            CollectionAssert.AreEqual(expendetOldRecourceList, actualList);

        }

        //-------------------------------------------------------

        // 1
        static string dataProvider = "System.Data.OleDb";
        static string connectionStr = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
            @"Data Source = " + Path.Combine(Directory.GetCurrentDirectory(), "") +
            @";Extended Properties = Excel 12.0 Xml; HDR=YES;";

        [TestMethod]
        [DataSource(
        "System.Data.OleDb",
        @"Provider=Microsoft.ACE.OLEDB.12.0;
            Data Source=..\..\DataTests\001 - 1301 - 1222333 - LastName - test_01.xlsx;
            Persist Security Info=False;
            Extended Properties='Excel 12.0 Xml; HDR=YES'",
        "Sheet$",
        DataAccessMethod.Sequential)]
        public void ImportDataForExcel_Test()
        {
            // arrange
            string expendetValidKmat = Convert.ToString(TestContext.DataRow["valid_kmat"]).Replace(" ", "");

            //act
            string kmat = Convert.ToString(TestContext.DataRow["kmat"]);
            string actual = "Functions.DeleteSpecialCharacters(kmat)";

            // assert
            Assert.AreEqual(expendetValidKmat, actual);

        }



        int counter = 0;
        // 2.2
        [TestMethod]
        public void ListFieldKmatForExcel_Test()
        {


            // arrange
            List<string> expectedListField = new List<string>(); // Functions.ListFieldKmatForExcel(dtExcel, "valid_kmat");

            // act
            List<string> actualListField = new List<string>(); // Functions.ListFieldKmatForExcel(dtExcel, "kmat");


            for (int i = 0; i <= expectedListField.Count - 1; i++)
                Assert.AreEqual(expectedListField[i], actualListField[i], "{0} - {1}: строка = {2}",
                    expectedListField[i], actualListField[i], i);

            // assert
            //CollectionAssert.AreEqual(expectedListField, actualListField, "{0} - {1}", 
            //    expectedListField[counter], actualListField[counter]);

            //counter++;
        }

        [TestMethod]
        public void DictionaryResourceAndCount_Test_CountElements()
        {
            // arrange
            Dictionary<string, int> expectedDictionaryGroupBy =
                DataTests.DictionaryGroupBy();

            // act
            List<string> expectedListField = excel.ListField("valid_kmat");
            Dictionary<string, int> actualDictionary = excel._groupByFieldAndCount("kmat");

            Assert.AreEqual(expectedDictionaryGroupBy.Count, actualDictionary.Count);

        }

        [TestMethod]
        public void DictionaryResourceAndCount_Test_CollectionEquals()
        {
            // arrange
            Dictionary<string, int> expectedDictionaryGroupBy =
                DataTests.DictionaryGroupBy();

            // act
            List<string> actualListField = excel.ListField("valid_kmat");
            Dictionary<string, int> actualDictionary = excel._groupByFieldAndCount("kmat"); // new Dictionary<string, int>(); // Functions.DictionaryResourceAndCount(actualListField);

            //Assert

            CollectionAssert.AreEqual(expectedDictionaryGroupBy, actualDictionary);

        }

        [TestMethod]
        //[DataSource(
        //"System.Data.OleDb",
        //@"Provider=Microsoft.ACE.OLEDB.12.0;
        //    Data Source=..\..\DataTests\001 - 1301 - 1222333 - LastName - test_01.xlsx;
        //    Persist Security Info=False;
        //    Extended Properties='Excel 12.0 Xml; HDR=YES'",
        //"Sheet$",
        //DataAccessMethod.Sequential)]
        public void ListUniqueFieldResource_Test()
        {
            HashSet<string> hashSetUnique = new HashSet<string>();

            // arrange
            var expendetListUnique = excel.ListField("kmat_double");

            foreach (var row in expendetListUnique)
                hashSetUnique.Add(row);

            //act
            var actualListUnique = excel.ListUniqueField("kmat") ;

            CollectionAssert.AreEqual(expendetListUnique, actualListUnique);

            //Assert.IsTrue(hashSetUnique.SetEquals(actualListUnique));

        }

        [TestMethod]
        [DataSource(
        "System.Data.OleDb",
        @"Provider=Microsoft.ACE.OLEDB.12.0;
            Data Source=..\..\DataTests\ConvertKmat_Test.xlsx;
            Persist Security Info=False;
            Extended Properties='Excel 12.0 Xml; HDR=YES'",
        "Sheet$",
        DataAccessMethod.Sequential)]
        public void RenameOldResourceInNew_Test()
        {
            // arrange

            //string expendetValidKmat = TestContext.DataRow["valid_kmat"];

            //act
            //string kmat = Convert.ToString(TestContext.DataRow["kmat"]);
            //string actual = Functions.ConvertOldResource(kmat);

            // assert
            //Assert.AreEqual(expendetValidKmat, actual);
        }

        [TestMethod]
        [DataSource(
        "System.Data.OleDb",
        @"Provider=Microsoft.ACE.OLEDB.12.0;
            Data Source=..\..\DataTests\RenameOldKmatInNew_Test.xlsx;
            Persist Security Info=False;
            Extended Properties='Excel 12.0 Xml; HDR=YES'",
        "Sheet$",
        DataAccessMethod.Sequential)]
        public void ConvertKmatList_Test()
        {

            HashSet<string> hashSetUnique = new HashSet<string>();

            // arrange

            string expectedKmat = TestContext.DataRow["kmat"].ToString();

            //act
            string kmat_old = TestContext.DataRow["kmat_old"].ToString();
            string ceh = TestContext.DataRow["ceh"].ToString();

            string actualKmat = ""; // Functions.RenameOldResourceInNew(kmat_old, ceh);

            // assert
            Assert.AreEqual(expectedKmat, actualKmat);

        }

        [TestMethod]
        public void ConvertCeh_Kmat_8_Symbols_Test()
        {

            List<string> expectedList = new List<string>() { "1301", "13002", "30003" };

            List<string> experimentlList = new List<string>() { "1301", "13002", "130003" };

            List<string> actualList = new List<string>();

            foreach (var ceh in experimentlList)
                actualList.Add(Functions.ConvertCeh(ceh, "12345678"));
                //actualList.Add("ConvertCeh");
                
            
            // arrange

            //string expectedKmat = TestContext.DataRow["kmat"].ToString();

            //act
            //string kmat_old = TestContext.DataRow["kmat_old"].ToString();
            //string ceh = TestContext.DataRow["ceh"].ToString();

            //string actualKmat = Functions.ConvertKmatTest(kmat_old, ceh);

            // assert
            //Assert.AreEqual(expectedKmat, actualKmat);

            CollectionAssert.AreEqual(expectedList, actualList);

        }

        [TestMethod]
        public void ConvertCeh_Kmat_6_Symbols_Test()
        {

            List<string> expectedList = new List<string>() { "1301", "13002", "30003" };

            List<string> experimentlList = new List<string>() { "1301", "13002", "130003" };

            List<string> actualList = new List<string>();

            foreach (var ceh in experimentlList)
                actualList.Add(Functions.ConvertCeh(ceh, "12345678"));
                //actualList.Add(Functions.ConvertCeh(ceh, "12345678"));

            // arrange

            //string expectedKmat = TestContext.DataRow["kmat"].ToString();

            //act
            //string kmat_old = TestContext.DataRow["kmat_old"].ToString();
            //string ceh = TestContext.DataRow["ceh"].ToString();

            //string actualKmat = Functions.ConvertKmatTest(kmat_old, ceh);

            // assert
            //Assert.AreEqual(expectedKmat, actualKmat);

            CollectionAssert.AreEqual(expectedList, actualList);

        }

        //[TestMethod]
        //public void ConvertAllResourceInExcelAndUniqueList()
        //{
        //    List<string> listKmat = Functions.ListFieldKmatForExcel(dtExcel, "kmat");
        //    List<string> listCeh = Functions.ListFieldKmatForExcel(dtExcel, "ceh");

        //    List<string> listConvertKmat = new List<string>();

        //    for(int i = 0; i < listKmat.Count; i++)
        //        listConvertKmat.Add(Functions.RenameOldResourceInNew(listUnique[i], listCeh[i]));



        //    //CollectionAssert.AreEqual();
        //    //listUnique
        //}


    }
}
