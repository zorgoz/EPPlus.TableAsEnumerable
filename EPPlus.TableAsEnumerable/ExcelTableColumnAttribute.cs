/* 
 * Copyright (c) 2016 zoltan.zorgo@gmail.com
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

using System;

namespace EPPlus.Extensions
{
    /// <summary>
    /// Attribute used to map property to Excel table column
    /// </summary>
    /// <remarks>Can not map by both name and index. It will map to the property name if none is specified</remarks>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExcelTableColumnAttribute : Attribute
    {
        private string columnName = null;

        /// <summary>
        /// Set this property to map by name
        /// </summary>
        public string ColumnName
        {
            get
            {
                return columnName;
            }
            set
            {
                if (columnIndex > 0) throw new ArgumentException("Can not set both Column Name and Column Index!");
                if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Column Name can't be empty!");

                columnName = value;
            }
        }

        private int columnIndex = 0;

        /// <summary>
        /// Use this property to map by 1-based index
        /// </summary>
        public int ColumnIndex
        {
            get
            {
                return columnIndex;
            }
            set
            {
                if (columnName != null) throw new ArgumentException("Can not set both Column Name and Column Index!");
                if (value <= 0) throw new ArgumentException("Column Index can't be zero or negative!");

                columnIndex = value;
            }
        }
    }
}