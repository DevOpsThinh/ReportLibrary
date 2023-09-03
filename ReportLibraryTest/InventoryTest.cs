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

using Microsoft.VisualStudio.TestTools.UnitTesting;
using ReportLibrary.BusinessRules;
using System.Collections.Generic;

namespace ReportLibraryTest
{
    [TestClass]
    public class InventoryTest
    {

        [TestMethod]
        public void TestMethod1()
        {
            var inventory = new List<Inventory>
            {
                new Inventory
                {
                    PartNumber = "1",
                    Description = "Part #1: Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Et netus et malesuada fames ac turpis. Tortor posuere ac ut consequat semper viverra. Eget egestas purus viverra accumsan. Hac habitasse platea dictumst vestibulum rhoncus. Elementum tempus egestas sed sed risus pretium quam vulputate. Donec massa sapien faucibus et molestie. Urna neque viverra justo nec ultrices dui sapien eget mi.",
                    Count = 3,
                    ItemUnit = "kg",
                    ItemPrice = 15.95m,
                    ItemNote = "",
                    IsOutStock = "0"
                },
                new Inventory
                {
                    PartNumber = "2",
                    Description = "Part #2: Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Et netus et malesuada fames ac turpis. Tortor posuere ac ut consequat semper viverra. Eget egestas purus viverra accumsan. Hac habitasse platea dictumst vestibulum rhoncus. Elementum tempus egestas sed sed risus pretium quam vulputate. Donec massa sapien faucibus et molestie. Urna neque viverra justo nec ultrices dui sapien eget mi.",
                    Count = 1,
                    ItemUnit = "liter",
                    ItemPrice = 8.95m,
                    ItemNote = "Go to Cart",
                    IsOutStock = "0"
                },
                new Inventory
                {
                    PartNumber = "3",
                    Description = "Part #3: Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Et netus et malesuada fames ac turpis. Tortor posuere ac ut consequat semper viverra. Eget egestas purus viverra accumsan. Hac habitasse platea dictumst vestibulum rhoncus. Elementum tempus egestas sed sed risus pretium quam vulputate. Donec massa sapien faucibus et molestie. Urna neque viverra justo nec ultrices dui sapien eget mi.",
                    ItemUnit = "piece",
                    Count = 2,
                    ItemPrice = 5.95m,
                    ItemNote = "Out of stock",
                    IsOutStock = "1"
                }
            };

            var expectedResult = "Added title...\nAdded header...\nAdded data...\nExcel file created at ExcelReport.xlsx";

            string report = new ReportBusiness<Inventory>()
                .GenerateReport(inventory, ReportTypeEnumeration.ExcelDynamic);

            Assert.AreEqual(expectedResult, report);
        }
    }
}
