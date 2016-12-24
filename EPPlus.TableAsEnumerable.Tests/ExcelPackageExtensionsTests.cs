using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using System.Reflection;
using EPPlus.Extensions;
using System.Linq;
using System.Collections.Generic;

namespace TESTS
{
    [TestClass]
    public class ExcelPackageExtensionsTests
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


        [TestMethod]
        public void Test_TableNameExtensions()
        {
            var tables = excelPackage.GetTables();

            Assert.IsNotNull(tables, "We have 3 tables");
            Assert.AreEqual(3, tables.Count(), "We have 3 tables");

            Assert.IsTrue(excelPackage.HasTable("TEST2"), "We have TEST2 table");
            Assert.IsTrue(excelPackage.HasTable("test2"), "Table names are case insensitive");

            Assert.AreSame(
                excelPackage.Workbook.Worksheets["TEST2"].Tables["TEST2"]
                , excelPackage.GetTable("TEST2")
                , "We are accessing the same objects");

            Assert.IsFalse(excelPackage.HasTable("NOTABLE"), "We don't have NOTABLE table");
            Assert.IsNull(excelPackage.GetTable("NOTABLE"), "We don't have NOTABLE table");
        }
    }
}
