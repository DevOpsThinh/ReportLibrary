﻿/*
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

using ReportLibrary.OfficeApps;
using System;
using System.Collections.Generic;

namespace ReportLibrary.BusinessRules
{
    /// <summary>
    /// A <c>ReportBusiness</c> class, that representing for report business rules.
    /// It contains methods for generating a report.
    /// </summary>
    public class ReportBusiness<D>
    {
        /// <summary>
        /// The <c>GenerateReport</c> method that generating a report
        /// </summary>
        public string GenerateReport(List<D> items, ReportTypeEnumeration reportType) 
        {
            BaseGenerator<D> generator = CreateGenerator(reportType);

            string report = generator.Generate(items);

            return report;
        }
        /// <summary>
        /// Figure out which report format to generate.
        /// </summary>
        /// <param name="reportType">A report type enumeration case</param>
        /// <returns></returns>
        BaseGenerator<D> CreateGenerator(ReportTypeEnumeration reportType)
        {
            Type gType;

            switch (reportType)
            {
                case ReportTypeEnumeration.ExcelTyped:
                    gType = typeof(ExcelTypedGenerator<>); 
                    break;
                case ReportTypeEnumeration.ExcelDynamic:
                    gType = typeof(ExcelDynamicGenerator<>); 
                    break;

                default: throw new ArgumentException($"Unexpected Report type case: '{reportType}'");
            }

            Type dataType = typeof(D);
            Type genericType = gType.MakeGenericType(dataType); ;
            object generator = Activator.CreateInstance(genericType);

            return (BaseGenerator<D>)generator;
        }
    }
}
