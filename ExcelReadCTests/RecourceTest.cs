using System.Collections.Generic;
using System.Data;
using ExcelReadC;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReadCTests
{
    [TestClass]
    public class RecourceTest
    {
        public static TestContext TestContext { get; set; }

        private static List<string> DoubleKmat;

        [ClassInitialize]
        public static void Initialization(TestContext context)
        {
            DoubleKmat = new List<string>();
            
            // = DoubleKmat.ToArray();
        }

        [TestMethod]
        public void ConvertOldResource_Test()
        {
            // arrange
            string expendetOldResource = "1123456";

            //act   
            string actualOldResource = Functions.ConvertOldResource("001-123,45 6");

            // assert
            Assert.AreEqual<string>(expendetOldResource, actualOldResource);

        }

        [TestMethod]
        public void ConvertOldResource_CollectionsTest()
        {
            // arrange
            List<string> expendetOldRecourceList =
                new List<string>(new[] { "112345", "223", "0023vmv", "555111" });

            //act   

            List<string> OldRecourceList =
                new List<string>(new[] { "11-2345", "00223", "0023vmv", "555.11,1" });

            List<string> actualList = new List<string>();

            foreach (string oldRecource in OldRecourceList)
                actualList.Add(Functions.ConvertOldResource(oldRecource));
            // assert
            CollectionAssert.AreEqual(expendetOldRecourceList, actualList);

        }

        [TestMethod]
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
                actualList.Add(Functions.ConvertOldResource(oldRecource));
            // assert
            CollectionAssert.AreEqual(expendetOldRecourceList, actualList);

        }
    }
}
