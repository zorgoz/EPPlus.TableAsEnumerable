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

using OfficeOpenXml.Table;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace EPPlus.Extensions
{
    /// <summary>
    /// Class holding extensien methods implemented
    /// </summary>
    public static class EPPlusAsEnumerableExtension
    {
        #region Helpers
        /// <summary>
        /// Helper extension method determining if a type is nullable
        /// </summary>
        /// <param name="type">Type to test</param>
        /// <returns>True if type is nullable</returns>
        internal static bool IsNullable(this Type type)
        {
            return (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>));
        }

        /// <summary>
        /// Helper extension method to test if a type is numeric or not
        /// </summary>
        /// <param name="type">Type to test</param>
        /// <returns>True if type is numeric</returns>
        internal static bool IsNumeric(this Type type)
        {
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return true;
                default:
                    return false;
            }
        }
        #endregion

        /// <summary>
        /// Generic extension method yielding objects of specified type from table.
        /// </summary>
        /// <remarks>Exceptions are not cathed. It works on all or nothing basis. 
        /// Only primitives and enums are supported as property.
        /// Currently supports only tables with header.</remarks>
        /// <typeparam name="T">Type to map to. Type should be a class and should have parameterless constructor.</typeparam>
        /// <param name="table">Table object to fetch</param>
        /// <param name="skipCastErrors">Determines how the method should handle exceptions when casting cell value to property type. 
        /// If this is true, invlaid casts are silently skipped, otherwise any error will cause method to fail with exception.</param>
        /// <returns>An enumerable of the generating type</returns>
        public static IEnumerable<T> AsEnumerable<T>(this ExcelTable table, bool skipCastErrors = false) where T : class, new()
        {
            IList mapping = PrepareMappings<T>(table);

            // Parse table
            for (int row = table.Address.Start.Row + (table.ShowHeader ? 1 : 0); 
                     row <= table.Address.End.Row - (table.ShowTotal ? 1 : 0); 
                     row++)
            {
                T item = (T)Activator.CreateInstance(typeof(T));

                foreach (KeyValuePair<int, PropertyInfo> map in mapping)
                {
                    object cell = table.WorkSheet.Cells[row, map.Key + table.Address.Start.Column].Value;

                    if (cell == null) continue;

                    var property = map.Value;

                    Type type = property.PropertyType;

                    // If type is nullable, get base type instead
                    if (property.PropertyType.IsNullable())
                    {
                        type = type.GetGenericArguments()[0];
                    }

                    try
                    {
                        TrySetProperty(item, type, property, cell);
                    }
                    catch (Exception ex)
                    {
                        if (!skipCastErrors)
                            throw new ExcelTableConvertException(
                                "Cell casting error occures",
                                ex,
                                new ExcelTableConvertExceptionArgs
                                {
                                    columnName = table.Columns[map.Key].Name,
                                    expectedType = type,
                                    propertyName = property.Name,
                                    cellValue = cell
                                }
                                );
                    }
                }

                yield return item;
            }
        }

        /// <summary>
        /// Method prepares mapping using the type and the attributes decorating its properties
        /// </summary>
        /// <typeparam name="T">Type to parse</typeparam>
        /// <param name="table">Table to get columns from</param>
        /// <returns>A list of mappings from column index to property</returns>
        private static IList PrepareMappings<T>(ExcelTable table)
        {
            IList mapping = new List<KeyValuePair<int, PropertyInfo>>();

            var propInfo = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public);

            // Build property-table column mapping
            foreach (var property in propInfo)
            {
                var mappingAttribute = (ExcelTableColumnAttribute)property.GetCustomAttributes(typeof(ExcelTableColumnAttribute), true).First();
                if (mappingAttribute != null)
                {
                    int col = -1;
                    // There is no case when both column name and index is specified since this is excluded by the attribute
                    // Neither index, nor column name is specified, use property name
                    if (mappingAttribute.ColumnIndex == 0 && string.IsNullOrWhiteSpace(mappingAttribute.ColumnName))
                    {
                        col = table.Columns[property.Name].Position;
                    }

                    // Column index was specified
                    if (mappingAttribute.ColumnIndex > 0)
                    {
                        col = table.Columns[mappingAttribute.ColumnIndex - 1].Position;
                    }

                    // Column name was specified
                    if (!string.IsNullOrWhiteSpace(mappingAttribute.ColumnName))
                    {
                        col = table.Columns[mappingAttribute.ColumnName].Position;
                    }

                    if (col == -1)
                    {
                        throw new ArgumentException("Sould never get here, but I can not identify column.");
                    }

                    mapping.Add(new KeyValuePair<int, PropertyInfo>(col, property));
                }
            }

            return mapping;
        }

        private static void TrySetProperty(object item, Type type, PropertyInfo property, object cell)
        {
            var itemType = item.GetType();

            if (type == typeof(string))
            {
                itemType.InvokeMember(
                property.Name,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                null,
                item,
                new object[] { cell.ToString() });

                return;
            }

            if (type == typeof(DateTime))
            {
                DateTime d = DateTime.Parse(cell.ToString());

                itemType.InvokeMember(
                property.Name,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                null,
                item,
                new object[] { d });

                return;
            }

            if (type == typeof(bool))
            {
                itemType.InvokeMember(
                property.Name,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                null,
                item,
                new object[] { cell });

                return;
            }

            if (type.IsEnum)
            {
                if (cell.GetType() == typeof(string)) // Support Enum conversion from string...
                {
                    itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { Enum.Parse(type, cell.ToString(), true) });
                }
                else // ...and numeric cell value
                {
                    var underType = type.GetEnumUnderlyingType();

                    itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { Enum.ToObject(type, Convert.ChangeType(cell, underType)) });
                }

                return;
            }

            if (type.IsNumeric())
            {
                itemType.InvokeMember(
                    property.Name,
                    BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                    null,
                    item,
                    new object[] { Convert.ChangeType(cell, type) });
            }
        }
    }
}