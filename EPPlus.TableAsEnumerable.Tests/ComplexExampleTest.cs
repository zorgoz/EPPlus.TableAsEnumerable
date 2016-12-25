using EPPlus.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;

namespace TESTS
{
    /// <summary>
    /// Summary description for ComplexExampleTest
    /// </summary>
    [TestClass]
    public class ComplexExampleTest
    {
        public static ExcelPackage excelPackage;

        public ComplexExampleTest()
        {
            //
            // NOOP
            //
        }

        private TestContext testContextInstance;

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

        [TestMethod]
        public void TestComplexFixtures()
        {
            Assert.IsNotNull(excelPackage, "Excel package is null");

            // TEST3

            var workSheet = excelPackage.Workbook.Worksheets["TEST3"];
            Assert.IsNotNull(workSheet, "Worksheet TEST3 missing");

            var table = workSheet.Tables["TEST3"];
            Assert.IsNotNull(table, "Table TEST3 missing");

            Assert.IsTrue(table.Address.Columns == 6, "Table3 is not as expected");

            Assert.IsTrue(table.Address.Rows == 5 + (table.ShowTotal ? 1 : 0) + (table.ShowHeader ? 1 : 0), "Table3 has missing rows");
        }

        enum Manufacturers { Opel = 1, Ford, Toyota };
        class Cars
        {
            [ExcelTableColumn(ColumnIndex = 1)]
            public string licensePlate { get; set; }

            [ExcelTableColumn]
            public Manufacturers manufacturer { get; set; }

            [ExcelTableColumn(ColumnName = "Manufacturing date")]
            public DateTime? manufacturingDate { get; set; }

            [ExcelTableColumn]
            public int price { get; set; }

            [ExcelTableColumn]
            public Color color { get; set; }

            [ExcelTableColumn(ColumnName = "Is ready for traffic?")]
            public bool ready { get; set; }

            public override string ToString()
            {
                return $"{(color.ToString())} {(manufacturer.ToString())} {(manufacturingDate?.ToShortDateString())}";
            }
        }

        [TestMethod]
        public void TestComplexExample()
        {
            var table = excelPackage.GetTable("TEST3");

            IEnumerable<Cars> enumerable = table.AsEnumerable<Cars>();
            IList<Cars> list = null;

            Assert.IsNotNull(enumerable);
            list = enumerable.ToList();

            Assert.IsTrue(list.Count() == 5, "We have 5 rows");
            Assert.IsTrue(list.Count(x => string.IsNullOrWhiteSpace(x.licensePlate)) == 1, "There is one without license plate");
            Assert.IsTrue(list.All(x => x.manufacturer > 0), "All should have manufacturers");
            Assert.IsNull(list.Last().manufacturingDate, "The last one's manufacturing date is unknown");
            Assert.IsTrue(list.Count(x => x.manufacturingDate == null) == 1, "Only one manufacturig date is unknown");
            Assert.AreSame(list.Single( x => x.licensePlate == null) , list.Single(x => !x.ready), "The one without the license plate is not ready");
            Assert.IsTrue(list.Max(x => x.price) == 12000, "Highest price is 12000");
            Assert.AreEqual(new DateTime(2015, 3, 10), list.Max(x => x.manufacturingDate), "Oldest was manufactured on 2015.03.10");
        }

    }
}
