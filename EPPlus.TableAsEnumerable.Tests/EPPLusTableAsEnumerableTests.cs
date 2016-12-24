using EPPlus.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace TESTS
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class EPPlusTableAsEnumerableTests
    {
        public static ExcelPackage excelPackage;

        public EPPlusTableAsEnumerableTests()
        {
            // NOOP
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
        
        #region Additional test attributes
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        /// <summary>
        /// Test existence of test objects in the embedded workbook
        /// </summary>
        [TestMethod]
        public void WarmUp()
        {
            Assert.IsNotNull(excelPackage, "Excel package is null");

            // TEST1

            var workSheet = excelPackage.Workbook.Worksheets["TEST1"];
            Assert.IsNotNull(workSheet, "Worksheet TEST1 missing");

            var table = workSheet.Tables["TEST1"];
            Assert.IsNotNull(table, "Table TEST1 missing");

            Assert.IsTrue(table.Address.Columns == 5, "Table1 is not as expected");

            Assert.IsTrue(table.Address.Rows > 2, "Table1 has missing rows");

            // TEST2

            workSheet = excelPackage.Workbook.Worksheets["TEST2"];
            Assert.IsNotNull(workSheet, "Worksheet TEST2 missing");

            table = workSheet.Tables["TEST2"];
            Assert.IsNotNull(table, "Table TEST2 missing");

            Assert.IsTrue(table.Address.Columns == 2, "Table2 is not as expected");

            Assert.IsTrue(table.Address.Rows > 2, "Table2 has missing rows");
        }

        #region Test for default mapping and mapping my name and index

        class DefaultMap
        {
            [ExcelTableColumn]
            public string name { get; set; }

            [ExcelTableColumn]
            public string gender { get; set; }
        }

        [TestMethod]
        public void Test_MapByDefault()
        {
            var table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            
            IEnumerable<DefaultMap> enumerable = table.AsEnumerable<DefaultMap>();
            IList<DefaultMap> list = null;

            Assert.IsNotNull(enumerable);
            try
            {
                list = enumerable.ToList();
                Assert.IsTrue(list.Count == 5, "We have expected 5 elements");
                Assert.IsTrue(list.First().name == "John", "We have expected John to be first");
                Assert.IsTrue(list.First().gender == "MALE", "We have expected a male to be first");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        class NamedMap
        {
            [ExcelTableColumn(ColumnName = "Name")]
            public string name { get; set; }

            [ExcelTableColumn(ColumnName = "Gender")]
            public string gender { get; set; }
        }

        [TestMethod]
        public void Test_MapByName()
        {
            var table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];


            IEnumerable<NamedMap> enumerable = table.AsEnumerable<NamedMap>();
            IList<NamedMap> list = null;

            Assert.IsNotNull(enumerable);
            try
            {
                list = enumerable.ToList();
                Assert.IsTrue(list.Count == 5, "We have expected 5 elements");
                Assert.IsTrue(list.First().name == "John", "We have expected John to be first");
                Assert.IsTrue(list.First().gender == "MALE", "We have expected a male to be first");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        class IndexMap
        {
            [ExcelTableColumn(ColumnIndex = 1)]
            public string name { get; set; }

            [ExcelTableColumn(ColumnIndex = 3)]
            public string gender { get; set; }
        }

        [TestMethod]
        public void Test_MapByIndex()
        {
            var table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];


            IEnumerable<IndexMap> enumerable = table.AsEnumerable<IndexMap>();
            IList<IndexMap> list = null;

            Assert.IsNotNull(enumerable);
            try
            {
                list = enumerable.ToList();
                Assert.IsTrue(list.Count == 5, "We have expected 5 elements");
                Assert.IsTrue(list.First().name == "John", "We have expected John to be first");
                Assert.IsTrue(list.First().gender == "MALE", "We have expected a male to be first");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }
        #endregion

        #region Type cast mapping tests
        enum gender { MALE = 1, FEMALE = 2}
        class EnumStringMap
        {
            [ExcelTableColumn(ColumnName = "Name")]
            public string name { get; set; }

            [ExcelTableColumn(ColumnName = "Gender")]
            public gender gender { get; set; }
        }

        [TestMethod]
        public void Test_MapEnumString()
        {
            var table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];


            IEnumerable<EnumStringMap> enumerable = table.AsEnumerable<EnumStringMap>();
            IList<EnumStringMap> list = null;

            Assert.IsNotNull(enumerable);
            try
            {
                list = enumerable.ToList();
                Assert.IsTrue(list.Count == 5, "We have expected 5 elements");
                Assert.IsTrue(list.Count(x => x.gender == gender.MALE) == 3, "We have expected 3 males");
                Assert.IsTrue(list.Count(x => x.gender == gender.FEMALE) == 2, "We have expected 2 females");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        enum Class : byte { Ten = 10, Nine = 9}
        class EnumByteMap
        {
            [ExcelTableColumn]
            public string name { get; set; }

            [ExcelTableColumn]
            public Class @class { get; set; }
        }

        [TestMethod]
        public void Test_MapEnumNumeric()
        {
            var table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            IEnumerable<EnumByteMap> enumerable = table.AsEnumerable<EnumByteMap>();
            IList<EnumByteMap> list = null;

            Assert.IsNotNull(enumerable);
            try
            {
                list = enumerable.ToList();
                Assert.IsTrue(list.Count == 5, "We have expected 5 elements");
                Assert.IsTrue(list.Count(x => x.@class == Class.Ten) == 2, "We have expected 2 in 10th class");
                Assert.IsTrue(list.Count(x => x.@class == Class.Nine) == 3, "We have expected 3 in 9th class");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        class MultiMap
        {
            [ExcelTableColumn]
            public string name { get; set; }

            [ExcelTableColumn(ColumnName = "Class")]
            public Class @class { get; set; }

            [ExcelTableColumn(ColumnName = "Class")]
            public int classAsInt { get; set; }
        }

        /// <summary>
        /// Test cases when a column is mapped to multiple properties (with different type)
        /// </summary>
        [TestMethod]
        public void Test_MultiMap()
        {
            var table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            IEnumerable<MultiMap> enumerable = table.AsEnumerable<MultiMap>();
            IList<MultiMap> list = null;

            Assert.IsNotNull(enumerable);
            try
            {
                list = enumerable.ToList();
                Assert.IsTrue(list.Count == 5, "We have expected 5 elements");

                var m = list.First(x => x.@class == Class.Ten);
                Assert.IsTrue((int)m.@class == m.classAsInt, "Ten sould be 10");

                var n = list.First(x => x.@class == Class.Nine);
                Assert.IsTrue((int)n.@class == n.classAsInt, "Nine sould be 9");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        class DateMap
        {
            [ExcelTableColumn]
            public string name { get; set; }

            [ExcelTableColumn]
            public gender gender { get; set; }

            [ExcelTableColumn(ColumnName = "Birth date")]
            public DateTime birthdate { get; set; }
        }

        [TestMethod]
        public void Test_DateMap()
        {
            var table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            IEnumerable<DateMap> enumerable = table.AsEnumerable<DateMap>();
            IList<DateMap> list = null;

            Assert.IsNotNull(enumerable);
            try
            {
                list = enumerable.ToList();
                Assert.IsTrue(list.Count == 5, "We have expected 5 elements");

                var a = list.FirstOrDefault(x => x.name == "Adam");
                Assert.AreEqual(new DateTime(1981, 4, 2), a.birthdate, "Adam' birthday is 1981.04.02");

                Assert.AreEqual(new DateTime(1979, 12, 1), list.Min(x => x.birthdate), "Oldest one was born on 1979.12.01");
            }
            catch (Exception ex)
            {
                Assert.Fail(ex.Message);
            }
        }

        #endregion

        #region Failure tests
        class EnumFailMap
        {
            [ExcelTableColumn]
            public string name { get; set; }

            [ExcelTableColumn(ColumnName = "Gender")]
            public Class gender { get; set; }
        }

        [TestMethod]
        public void Test_MapFail()
        {
            var table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            IEnumerable<EnumFailMap> enumerable = table.AsEnumerable<EnumFailMap>();
            IList<EnumFailMap> list = null;

            Assert.IsNotNull(enumerable);
            try
            {
                list = enumerable.ToList();
            }
            catch(ExcelTableConvertException ex)
            {
                Assert.IsTrue(ex.args.cellValue.ToString() == "MALE");
                Assert.IsTrue(ex.args.expectedType == typeof(Class));
                Assert.AreEqual("gender", ex.args.propertyName, true);
                Assert.AreEqual("gender", ex.args.columnName, true);
                return;
            }

            Assert.Fail("We should get an ExcelTableConvertException");
        }

        [TestMethod]
        public void Test_MapSilentFail()
        {
            var table = excelPackage.Workbook.Worksheets["TEST1"].Tables["TEST1"];

            IEnumerable<EnumFailMap> enumerable = table.AsEnumerable<EnumFailMap>(true);
            IList<EnumFailMap> list = null;

            Assert.IsNotNull(enumerable);

            list = enumerable.ToList();
            Assert.IsNotNull(list, "We should get the list");
            Assert.IsTrue(list.All(x => !string.IsNullOrWhiteSpace(x.name)), "All names should be there");
            Assert.IsTrue(list.All(x => x.gender == 0 ), "All genders should be 0");
        }

        #endregion

        #region Testing nullable

        class CarNullable
        {
            [ExcelTableColumn(ColumnName = "Car name")]
            public string name { get; set; }

            [ExcelTableColumn]
            public int? price { get; set; }
        }

        [TestMethod]
        public void Test_Nullable()
        {
            var table = excelPackage.Workbook.Worksheets["TEST2"].Tables["TEST2"];

            IEnumerable<CarNullable> enumerable = table.AsEnumerable<CarNullable>(true);
            IList<CarNullable> list = null;
            
            list = enumerable.ToList();
            Assert.IsTrue(list.Count(x => !x.price.HasValue) == 2, "Should have two ");
        }

        #endregion
    }
}
