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

using Microsoft.Office.Interop.Excel;
using ReportLibrary.BusinessRules;
using ReportLibrary.ReportColumn;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReportLibrary.OfficeApps
{
    /// <summary>
    /// The <c>ExcelDynamicGenerator</c> class that uses dynamic code to generate an Excel report.
    /// It inherits from <c>BaseGenerator</c> abstract base class.
    /// </summary>
    public class ExcelDynamicGenerator<D>: BaseGenerator<D>
    {
        Application appExcel;
        dynamic workbook;
        Worksheet worksheet;

        public ExcelDynamicGenerator()
        {
            appExcel = new Application();
            appExcel.Visible = true;

            workbook = appExcel.Workbooks.Add();
            worksheet = workbook.ActiveSheet;
        }
        /// <summary>
        /// The <c>GetHeaders</c> override method that gets header data.
        /// </summary>
        protected override StringBuilder GetHeaders(Dictionary<string, ReportColumnDetail> details)
        {
            ReportColumnDetail[] values = details.Values.ToArray();

            for (int i = 0; i < values.Length; i++)
            {
                ReportColumnDetail detail = values[i];
                worksheet.Cells[3, i + 1] = detail.Attribute.Name;
            }

            return new StringBuilder("Đề mục đã thêm...\n");
        }
        /// <summary>
        /// The <c>GetRows</c> override method that combines and formats all rows of data.
        /// </summary>
        protected override StringBuilder GetRows(List<D> items, Dictionary<string, ReportColumnDetail> details)
        {
            const int INITROW = 4;

            int rows = items.Count;
            int cols = details.Count;

            var data = new string[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                List<string> columns = GetColumns(details.Values, items[i]);

                for (int j = 0; j < cols; j++)
                {
                    data[i, j] = columns[j];
                }
            }

            int firstColumn = 'A';
            int lastExcelColumn = firstColumn + cols - 1;
            int lastExcelRow = INITROW + rows - 1;

            string EndRangeColumn = ((char)lastExcelColumn).ToString();
            string EndRangeRow = lastExcelRow.ToString();
            string EndRange = EndRangeColumn + EndRangeRow;
            string BeginRange = "A" + INITROW.ToString();

            var dataRange = worksheet.get_Range(BeginRange, EndRange);
            dataRange.Value2 = data;

            var gui = Guid.NewGuid();

            workbook.SaveAs($"BaoCaoXe_{gui}.xlsx", XlSaveAsAccessMode.xlShared);

            return new StringBuilder("Dữ liệu đã thêm...\n" + "tệp Excel đã tạo có tên BaoCaoXe.xlsx");
        }
        /// <summary>
        /// The <c>GetTitle</c> override method that gets title data
        /// </summary>
        protected override StringBuilder GetTitle()
        {
            worksheet.Cells[1, 1] = "Báo Cáo";

            return new StringBuilder("Tựa đề đã thêm...\n");
        }
    }
}
