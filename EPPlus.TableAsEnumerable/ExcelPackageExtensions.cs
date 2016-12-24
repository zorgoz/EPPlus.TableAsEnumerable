/* 
 * Copyright (c) 2016 Zoltán Zörgő
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
 * to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
 * and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS 
 * IN THE SOFTWARE.
 */

using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Extensions
{
    /// <summary>
    /// Class holds extensions on ExcelPackage object
    /// </summary>
    public static class ExcelPackageExtensions
    {
        /// <summary>
        /// Method returns all table names in the opened worksheet
        /// </summary>
        /// <remarks>Excel is ensuring the uniqueness of table names</remarks>
        /// <param name="excel">Extended ExcelPackage object</param>
        /// <returns>Enumeration of ExcelTables</returns>
        public static IEnumerable<ExcelTable> GetTables(this ExcelPackage excel)
        {
            foreach (var ws in excel.Workbook.Worksheets)
            {
                foreach (var t in ws.Tables)
                    yield return t;
            }
        }

        /// <summary>
        /// Method returns concrete ExcelTable by it's name 
        /// </summary>
        /// <param name="excel">Extended ExcelPackage object</param>
        /// <param name="name">Name of the table</param>
        /// <returns>ExcelTable object if found, null inf not</returns>
        public static ExcelTable GetTable(this ExcelPackage excel, string name)
        {
            return excel.GetTables().FirstOrDefault(t => t.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
        }

        /// <summary>
        /// Method checks for table in the ExcelPackage
        /// </summary>
        /// <param name="excel">Extended ExcelPackage object</param>
        /// <param name="name">Name of the table</param>
        /// <returns>Result of search as bool</returns>
        public static bool HasTable(this ExcelPackage excel, string name)
        {
            return excel.GetTables().Any(t => t.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase));
        }
    }
}
