/* 
 * Copyright (c) 2025 Zoltán Zörgő
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
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using ClosedXML.Excel;

namespace ClosedXml.Extensions;

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
    /// Method validates the excel table against the generating type. While AsEnumerable skips null cells, validation winn not.
    /// </summary>
    /// <typeparam name="T">Generating class type</typeparam>
    /// <param name="table">Extended object</param>
    /// <returns>An enumerable of <see cref="ExcelTableConvertExceptionArgs"/> containing </returns>
    public static IEnumerable<ExcelTableConvertExceptionArgs> Validate<T>(this IXLTable table) where T : class
    {
        var mapping = PrepareMappings<T>(table);
        var result = new LinkedList<ExcelTableConvertExceptionArgs>();

        T item = (T)Activator.CreateInstance(typeof(T));

        var from = table.DataRange.FirstRow().RowNumber();
        var to = table.DataRange.LastRow().RowNumber();

        // Parse table
        for (int row = from; row <= to; row++)
        {
            foreach (var map in mapping)
            {
                var cell = table.Cell(row - 1, map.Key);

                PropertyInfo property = map.Value.propertyInfo;

                try
                {
                    TrySetProperty(item, property, map.Value.required, cell.Value);
                }
                catch
                {
                    result.AddLast( 
                            new ExcelTableConvertExceptionArgs
                            {
                                columnName = table.HeadersRow().Cell(map.Key).GetText(),
                                expectedType = property.PropertyType,
                                propertyName = property.Name,
                                cellValue = cell.IsEmpty() ? null : cell.Value.GetText(),
                                cellAddress = cell.Address
                            });
                }
            }
        }

        return result;
    }

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
    public static IEnumerable<T> AsEnumerable<T>(this IXLTable table, bool skipCastErrors = false) where T : class
    {
        var mapping = PrepareMappings<T>(table);

        var from = table.DataRange.FirstRow().RowNumber();
        var to = table.DataRange.LastRow().RowNumber();

        // Parse table
        for (int row = from; row <= to; row++)
        {
            T item = (T)Activator.CreateInstance(typeof(T));

            foreach (var map in mapping)
            {
                var cell = table.Cell(row - 1, map.Key);

                PropertyInfo property = map.Value.propertyInfo;

                try
                {
                    TrySetProperty(item, property, map.Value.required, cell.Value);
                }
                catch (Exception ex)
                {
                    if (!skipCastErrors)
                        throw new ExcelTableConvertException(
                            "Cell casting error occures",
                            ex,
                            new ExcelTableConvertExceptionArgs
                            {
                                columnName = table.HeadersRow().Cell(map.Key).GetText(),
                                expectedType = property.PropertyType,
                                propertyName = property.Name,
                                cellValue = cell.IsEmpty() ? null : cell.Value.GetText(),
                                cellAddress = cell.Address
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
    private static Dictionary<int, (PropertyInfo propertyInfo, bool required)> PrepareMappings<T>(IXLTable table)
    {
        var mapping = new Dictionary<int, (PropertyInfo, bool)>();

        var propInfo = typeof(T).GetProperties(BindingFlags.Instance | BindingFlags.Public);

        // Build property-table column mapping
        foreach (var property in propInfo)
        {
            var mappingAttribute = (ExcelTableColumnAttribute)property.GetCustomAttributes(typeof(ExcelTableColumnAttribute), true).FirstOrDefault();
            if (mappingAttribute != null)
            {
                int idx = -1;
                // There is no case when both column name and index is specified since this is excluded by the attribute
                // Neither index, nor column name is specified, use property name
                if (mappingAttribute.ColumnIndex == 0 && string.IsNullOrWhiteSpace(mappingAttribute.ColumnName))
                {
                    idx = table.Field(property.Name).Index;
                } else
                // Column index was specified
                if (mappingAttribute.ColumnIndex > 0)
                {
                    idx = table.Field(mappingAttribute.ColumnIndex - 1).Index;
                } else
                // Column name was specified
                if (!string.IsNullOrWhiteSpace(mappingAttribute.ColumnName))
                {
                    idx = table.Field(mappingAttribute.ColumnName).Index;
                }

                if (idx == -1)
                {
                    throw new ArgumentException("Sould never get here, but I can not identify column.");
                }

                var required = property.GetCustomAttributes().Any(x => x.GetType().FullName == "System.Runtime.CompilerServices.RequiredMemberAttribute");

                mapping[idx + 1] = (property, required);
            }
        }

        return mapping;
    }

    /// <summary>
    /// Method tries to set property of item9
    /// </summary>
    /// <param name="item">target object</param>
    /// <param name="property">property to be set</param>
    /// <param name="required">is property required</param>
    /// <param name="cell">cell value</param>
    private static void TrySetProperty(object item, PropertyInfo property, bool required, XLCellValue cell)
    {
        Type type = property.PropertyType;
        Type itemType = item.GetType();

        // If type is nullable, get base type instead
        if (property.PropertyType.IsNullable())
        {
            if (cell.IsBlank) return; // If it is nullable, and we have null we should not waste time
             
            type = type.GetGenericArguments()[0];
        }

        if (required && cell.IsBlank)
        {
            throw new MissingMemberException();
        }

        if (type == typeof(string))
        {
            itemType.InvokeMember(
            property.Name,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
            null,
            item,
            [cell.IsBlank ? null : cell.GetText()]);

            return;
        }

        if (type == typeof(DateTime))
        {
            DateTime d = cell.GetDateTime();

            itemType.InvokeMember(
            property.Name,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
            null,
            item,
            [d]);

            return;
        }

        if (type == typeof(bool))
        {
            itemType.InvokeMember(
            property.Name,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
            null,
            item,
            [cell.GetBoolean()]);

            return;
        }

        if (type.IsEnum)
        {
            if (cell.IsText) // Support Enum conversion from string...
            {
                itemType.InvokeMember(
                property.Name,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                null,
                item,
                [Enum.Parse(type, cell.GetText(), true)]);
            }
            else // ...and numeric cell value
            {
                var underType = type.GetEnumUnderlyingType();

                itemType.InvokeMember(
                property.Name,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty,
                null,
                item,
                [Enum.ToObject(type, Convert.ChangeType(cell.GetNumber(), underType))]);
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
                [Convert.ChangeType(cell.GetNumber(), type)]);
        }
    }
}