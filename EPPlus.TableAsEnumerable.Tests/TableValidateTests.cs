using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using System.Reflection;
using EPPlus.Extensions;
using System.Linq;

namespace TESTS
{
    [TestClass]
    public class TableValidateTests
    {
        private TestContext testContextInstance;
        private static ExcelPackage excelPackage;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        /// <summary>
        /// Initializes EPPLus excelPackage with the embedded content
        /// </summary>
        [ClassInitialize()]
        public static void MyClassInitialize(TestContext testContext)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "TESTS.Resources.testsheets.xlsx";

            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                excelPackage = new ExcelPackage(stream);
            }
        }

        /// <summary>
        /// Frees up excelPackage
        /// </summary>
        [ClassCleanup()]
        public static void MyClassCleanup()
        {
            excelPackage.Dispose();
        }

        enum Manufacturers { Opel = 1, Ford, Mercedes};
        class WrongCars
        {
            [ExcelTableColumn(ColumnName = "License plate")]
            public string licensePlate { get; set; }

            [ExcelTableColumn]
            public Manufacturers manufacturer { get; set; }

            [ExcelTableColumn(ColumnName = "Manufacturing date")]
            public DateTime manufacturingDate { get; set; }

            [ExcelTableColumn(ColumnName = "Is ready for traffic?")]
            public bool ready { get; set; }
        }

        [TestMethod]
        public void Test_TableValidation()
        {
            var table = excelPackage.GetTable("TEST3");

            Assert.IsNotNull(table, "We have TEST3 table");

            var validation = table.Validate<WrongCars>().ToList();

            Assert.IsNotNull(validation, "we have errors here");
            Assert.AreEqual(2, validation.Count, "We have 2 errors");
            Assert.IsTrue(validation.Exists(x => x.cellAddress.Address.Equals("C6", StringComparison.InvariantCultureIgnoreCase)), "Toyota is not in the enumeration");
            Assert.IsTrue(validation.Exists(x => x.cellAddress.Address.Equals("D7", StringComparison.InvariantCultureIgnoreCase)), "Date is null");
        }
    }
}
