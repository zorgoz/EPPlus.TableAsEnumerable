using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EPPLus.Extensions;

namespace TESTS
{
    [TestClass]
    public class TypeExtensionTests
    {
        [TestMethod]
        public void Test_IsNullable()
        {
            Assert.IsTrue(typeof(int?).IsNullable());

            Assert.IsFalse(typeof(long).IsNullable());
        }

        [TestMethod]
        public void Test_IsNumeric()
        {
            Assert.IsTrue(typeof(int).IsNumeric());

            Assert.IsFalse(typeof(string).IsNumeric());

            Assert.IsFalse(typeof(Exception).IsNumeric());
        }
    }
}
