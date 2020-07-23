//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="ExcelAssignmentTest.cs" company="Chuck Hill">
// Copyright (c) 2020 Chuck Hill.
//
// This library is free software; you can redistribute it and/or
// modify it under the terms of the GNU Lesser General Public License
// as published by the Free Software Foundation; either version 2.1
// of the License, or (at your option) any later version.
//
// This library is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU Lesser General Public License for more details.
//
// The GNU Lesser General Public License can be viewed at
// http://www.opensource.org/licenses/lgpl-license.php. If
// you unfamiliar with this license or have questions about
// it, here is an http://www.gnu.org/licenses/gpl-faq.html.
//
// All code and executables are provided "as is" with no warranty
// either express or implied. The author accepts no liability for
// any damage or loss of business that this product may cause.
// </copyright>
// <author>Chuck Hill</author>
//--------------------------------------------------------------------------

using System;
using System.CodeDom;
using System.Globalization;
using System.IO;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;

namespace CsvExcelExportImport.UnitTests
{
    [TestFixture]
    public class ExcelAssignmentTest
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test, Order(1)]
        public void TestSerialize()
        {
            using (var stream = new MemoryStream())
            {
                var excel = new ExcelAssignment();
                var wb = TestData.CreateWorkbookProperties();
                excel.Serialize(stream, TestData.AssignmentData, "Worksheet Tab Name", wb);
                stream.Flush();
                stream.Position = 0;

                // The size changes all the time!
                // Assert.AreEqual(9149, stream.Position, "Serialized byte count are equal.");
            }
        }

        [Test, Order(2)]
        public void TestDeserialize()
        {
            using (var stream = new MemoryStream())
            {
                var excel = new ExcelAssignment();
                var wb = TestData.CreateWorkbookProperties();
                excel.Serialize(stream, TestData.AssignmentData, "Worksheet Tab Name", wb);
                stream.Flush();
                stream.Position = 0;

                var wb2 = new WorkbookProperties();
                string[,] result = excel.Deserialize(stream, wb2);
                stream.Position = 0;
                Assert.IsTrue(result == null, "Serialize/Deserialize Assignment round trip");

                ChangeExcelValues(stream);
                wb2 = new WorkbookProperties();
                result = excel.Deserialize(stream, wb2);
                Assert.IsTrue(result[2, 3] == "O", "Removed Assignment");
                Assert.IsTrue(result[2, 4] == "X", "Added Assignment");
            }
        }

        private static void ChangeExcelValues(Stream stream)
        {
            using (var pkg = new ExcelPackage(stream))
            {
                ExcelWorkbook wb = pkg.Workbook;
                var ws = wb.Worksheets.FirstOrDefault();
                ws.Cells[3, 4].Value = string.Empty;
                ws.Cells[3, 5].Value = "x";
                stream.Position = 0;
                pkg.SaveAs(stream);
            }

            stream.Position = 0;
        }
    }
}