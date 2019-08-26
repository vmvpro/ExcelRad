using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelReadC;
using System.Collections.Generic;
using System.Data;

namespace ExcelReadC_UnitTest
{
    [TestClass]
    public class UnitTest1
    {

        public TestContext TestContext { get; set; }
        static List<string> DoubleKmat; 

        [ClassInitialize]
        public static void Initialization(TestContext context)
        {
            DoubleKmat = new List<string>();
        }

        [TestMethod]
        public void ConvertKmat_OneMethodTest()
        {
            // arrange
            string expendet = "920" + "540" + "000012345";

            //act
            string kmat_old = "12345";
            string ceh = "540";
            List<string> DoubleKmat_ = new List<string>() { };


            string kmat = Functions.ConvertKmat(kmat_old, ceh, DoubleKmat_);

            // assert
            Assert.AreEqual(expendet, kmat);

        }

        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.XML",
        "testDataKmat.xml",
        "Resource",
        DataAccessMethod.Sequential)]
        [TestMethod]
        public void ConvertKmat_AllMethodTest()
        {
            // arrange
            string expendet = Convert.ToString(TestContext.DataRow["KmatNew"]).Replace(" ", "");

            //act
            string kmat_old = Convert.ToString(TestContext.DataRow["KmatOld"]);
            string ceh = Convert.ToString(TestContext.DataRow["Ceh"]);

            string kmat = Functions.ConvertKmat(kmat_old, ceh, DoubleKmat);
            DoubleKmat.Add(kmat_old);


            

            // assert
            //Assert.AreEqual(expendet, kmat);
            StringAssert.Equals("1", "2");

        }

        


        [DataSource("Microsoft.VisualStudio.TestTools.DataSource.XML",
        "testDataKmat.xml",
        "Resource",
        DataAccessMethod.Sequential)]
        [TestMethod]
        public void ConvertKmatTest_AllMethodTest_()
        {
            // arrange
            string expendet = Convert.ToString(TestContext.DataRow["KmatNew"]).Replace(" ", "");

            //act
            string kmat_old = Convert.ToString(TestContext.DataRow["KmatOld"]);
            string ceh = Convert.ToString(TestContext.DataRow["Ceh"]);

            string kmat = Functions.ConvertKmat(kmat_old, ceh, DoubleKmat);
            DoubleKmat.Add(kmat_old);




            // assert
            Assert.AreEqual(expendet, kmat);

        }

    }
}
