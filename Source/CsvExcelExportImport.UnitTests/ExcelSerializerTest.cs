//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="ExcelSerializerTest.cs" company="Chuck Hill">
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
using System.Globalization;
using System.IO;
using System.Linq;
using NUnit.Framework;

namespace CsvExcelExportImport.UnitTests
{
    [TestFixture]
    public class ExcelSerializerTest
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
                var excel = new ExcelSerializer();

                var wb = TestData.CreateWorkbookProperties();

                excel.Serialize(stream, TestData.ArrayOfModels, wb);
                stream.Flush();

                // The size changes all the time!
                // Assert.AreEqual(9149, stream.Position, "Serialized byte count are equal.");
            }
        }

        [Test, Order(2)]
        public void TestDeserialize()
        {
            using (var stream = new MemoryStream())
            {
                var excel = new ExcelSerializer();

                var wb = TestData.CreateWorkbookProperties();

                excel.Serialize(stream, TestData.ArrayOfModels, wb);
                stream.Flush();

                stream.Position = 0;

                var wb2 = new WorkbookProperties();

                var results = excel.Deserialize(stream, wb2);

                var page1 = TestData.ArrayOfModels[0].OfType<DataModel>();
                var page2 = TestData.ArrayOfModels[1].OfType<DataModel2>();
                var result1 = results[0].OfType<DataModel>();
                var result2 = results[1].OfType<DataModel2>();

                Assert.IsTrue(page1.SequenceEqual(result1), "Serialize/Deserialize round trip");
                Assert.IsTrue(page2.SequenceEqual(result2), "Serialize/Deserialize round trip");
            }
        }
    }
}