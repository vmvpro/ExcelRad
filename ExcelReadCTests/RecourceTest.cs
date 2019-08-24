using System;
using System.Collections.Generic;
using ExcelReadC;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelReadCTests
{
    [TestClass]
    public class RecourceTest
    {
        public TestContext TestContext { get; set; }
        static List<string> DoubleKmat;

        [ClassInitialize]
        public static void Initialization(TestContext context)
        {
            DoubleKmat = new List<string>();
        }

        [TestMethod]
        public void ConvertOldResource_Test()
        {
            // arrange
            string expendetOldResource = "1123456";

            //act   
            string actualOldResource = Functions.ConvertOldResource("001-123456");

            // assert
            Assert.AreEqual<string>(expendetOldResource, actualOldResource);
            
        }
    }
}
