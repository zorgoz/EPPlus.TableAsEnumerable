using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPPLus.Extensions;

namespace TESTS
{
    [TestClass]
    public class ExcelColumnAttributeTests
    {
        [TestMethod]
        public void Tets_MappingIndexBase()
        {
            try
            {
                var a = new ExcelTableColumnAttribute();
                a.ColumnIndex = 0;

                Assert.Fail("Should get an exception");
            }
            catch (ArgumentException)
            {
                return;
            }

            Assert.Fail("Should get an ArgumentException");
        }

        [TestMethod]
        public void Tets_MappingName()
        {
            try
            {
                var a = new ExcelTableColumnAttribute();
                a.ColumnName = "   ";

                Assert.Fail("Should get an exception");
            }
            catch (ArgumentException)
            {
                return;
            }

            Assert.Fail("Should get an ArgumentException");
        }

        [TestMethod]
        public void Tets_MappingUnambiguity()
        {
            try
            {
                var a = new ExcelTableColumnAttribute();
                a.ColumnIndex = 100;
                a.ColumnName = "TEST";
            Assert.Fail("Should get an exception");
            }
            catch (ArgumentException)
            {
                return;
            }

            Assert.Fail("Should get an ArgumentException");
        }
    }
}
