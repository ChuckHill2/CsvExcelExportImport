//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="CsvSerializerTest.cs" company="Chuck Hill">
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
    public class CsvSerializerTest
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test, Order(1)]
        public void TestSerialize()
        {
            var csv = new CsvSerializer(CultureInfo.InvariantCulture);
            var stream = new MemoryStream();

            var sw = new StreamWriter(stream);
            csv.Serialize(sw, TestData.ArrayOfModels);

            sw.Flush();
            stream.Position = 0;

            var str = new StreamReader(stream).ReadToEnd();

            Assert.AreEqual(TestData.Csv, str);
        }

        [Test, Order(2)]
        public void TestDeserialize()
        {
            var csv = new CsvSerializer(CultureInfo.InvariantCulture);

            var sr = new StringReader(TestData.Csv);

            var result = csv.Deserialize(sr, new Type[] { typeof(DataModel), typeof(DataModel2) });

            Assert.IsTrue(TestData.ArrayOfModels[0].OfType<DataModel>().SequenceEqual(result[0].OfType<DataModel>()));
            Assert.IsTrue(TestData.ArrayOfModels[1].OfType<DataModel2>().SequenceEqual(result[1].OfType<DataModel2>()));
        }
    }
}