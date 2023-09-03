/*
FPA License - EDUCATIONAL & PERSONAL

Copyright (©) 2023 Nguyen Truong Thinh

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without limitation in the rights to use, copy, modify, merge,
publish, and/ or distribute copies of the Software in an educational or
personal context, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

Permission is granted to sell and/ or distribute copies of the Software in
a commercial context, subject to the following conditions:
Substantial changes: adding, removing, or modifying large parts, shall be
developed in the Software. Reorganizing logic in the software does not warrant
a substantial change.

This project and source code may use libraries or frameworks that are
released under various Open-Source licenses. Use of those libraries and
frameworks are governed by their own individual licenses.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
 */

using ReportLibrary.ReportColumn;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ReportLibrary.BusinessRules
{
    /// <summary>
    /// A <c>BaseGenerator</c> reusable base class for generating reports.
    /// </summary>
    /// <typeparam name="D">A generic data type</typeparam>
    public abstract class BaseGenerator<D>
    {
        /// <summary>
        /// The <c>GetReflectedResult</c> tuple that pulls the value out of an object property.
        /// </summary>
        /// <param name="item">A generic data object item</param>
        /// <param name="property">A PropertyInfo instance</param>
        /// <returns>A tuple (object, Type)</returns>
        (object, Type) GetReflectedResult(D item, PropertyInfo property)
        {
            object result = property.GetValue(item);
            Type type = property.PropertyType;

            return (result, type);
        }
        /// <summary>
        /// The <c>GetColumnDetails</c> dictionary that collects the data with LINQ to populate column metadata.
        /// </summary>
        /// <param name="items">The data object list</param>
        /// <returns>The data object list was collected</returns>
        Dictionary<string, ReportColumnDetail> GetColumnDetails(List<D> items)
        {
            D item = items.First();
            Type type = item.GetType();
            PropertyInfo[] properties = type.GetProperties();

            return (
                from pro in properties
                let attribute = pro.GetCustomAttribute<ReportColumnAttribute>()
                where attribute != null
                select new ReportColumnDetail
                {
                    Name = pro.Name,
                    Attribute = attribute,
                    PropertyInfo = pro
                }).ToDictionary(key => key.Name, val => val);
        }
        /// <summary>
        /// The <c>Generate</c> method that builds a complete report.
        /// </summary>
        /// <param name="items">The generic data items list</param>
        public string Generate(List<D> items)
        {
            StringBuilder report = GetTitle();

            Dictionary<string, ReportColumnDetail> columnDetails = GetColumnDetails(items);

            report.Append(GetHeaders(columnDetails));
            report.Append(GetRows(items, columnDetails));

            return report.ToString();
        }
        /// <summary>
        /// The <c>GetTitle</c> abstract method that gets title/ writing the title
        /// </summary>
        protected abstract StringBuilder GetTitle();
        /// <summary>
        /// The <c>GetHeaders</c> abstract method that gets header/ writing the header
        /// </summary>
        protected abstract StringBuilder GetHeaders(Dictionary<string, ReportColumnDetail> details);
        /// <summary>
        /// The <c>GetRows</c> abstract method gets rows of data.
        /// </summary>
        protected abstract StringBuilder GetRows(List<D> items, Dictionary<string, ReportColumnDetail> details);
        /// <summary>
        /// The <c>GetColumns</c> abstract method that retrieve and format property data.
        /// </summary>
        protected List<string> GetColumns(IEnumerable<ReportColumnDetail> details, D item)
        {
            var columns = new List<string>();

            foreach (var column in details)
            {
                PropertyInfo member = column.PropertyInfo;

                string formatter = string
                    .IsNullOrWhiteSpace(column.Attribute.Format) ? "{0}" : column.Attribute.Format;

                (object result, Type colType) = GetReflectedResult(item, member);

                switch (colType.Name)
                {
                    case "Decimal":
                        columns.Add(string.Format(formatter, (decimal)result)); break;
                    case "Double":
                        columns.Add(string.Format(formatter, (double)result)); break;
                    case "Int32":
                        columns.Add(string.Format(formatter, (int)result)); break;
                    case "String":
                        columns.Add(string.Format(formatter, (string)result)); break;
                    default: break;
                }
            }

            return columns;
        }
    }
}
