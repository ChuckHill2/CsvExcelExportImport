//-------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="Program.cs" company="Chuck Hill">
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
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using CsvExcelExportImport.UnitTests;

namespace CsvExcelExportImport.TestConsole
{
    class Program
    {
        public const int Million = 1048576;

        static void Main()
        {
            TestWriteExcel();
            Console.Write("\nPress any key> ");
            Console.Read();
        }

#region 'Test' Functions

        public static void TestFormatting()
        {
            var data = DataModelColumnFormatting.GenerateData(1000);

            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\TestFormatting.xlsx";
            using (var stream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                var excel = new ExcelSerializer();
                var wb = TestData.CreateWorkbookProperties();
                excel.Serialize(stream, new IEnumerable[]{ data }, wb);
            }
        }

        private static void CreateData()
        {
            TestData.CreateData();
        }

        private static void TestReadRawJson()
        {
            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\FakeData(Not Indented).json";
            using (var sw = new StreamReader(path))
            {
                var page1 = Cast.JsonToModels<DataModel>(sw);
                var page2= Cast.JsonToModels<DataModel2>(sw);

                var page1a = page1.ToArray();
                var page2a = page2.ToArray();

                bool b1 = TestData.ArrayOfModels[0].OfType<DataModel>().ToArray().SequenceEqual(page1a);
                bool b2 = TestData.ArrayOfModels[1].OfType<DataModel2>().ToArray().SequenceEqual(page2a);
            }

            path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\FakeData(Indented).json";
            using (var sw = new StreamReader(path))
            {
                var page1 = Cast.JsonToModels<DataModel>(sw);
                var page2 = Cast.JsonToModels<DataModel2>(sw);

                var page1a = page1.ToArray();
                var page2a = page2.ToArray();

                bool b1 = TestData.ArrayOfModels[0].OfType<DataModel>().ToArray().SequenceEqual(page1a);
                bool b2 = TestData.ArrayOfModels[1].OfType<DataModel2>().ToArray().SequenceEqual(page2a);
            }
        }

        private static void TestWriteExcel(IEnumerable dm1 = null)
        {
            var wb = TestData.CreateWorkbookProperties();
            wb.Culture = CultureInfo.GetCultureInfo("de-DE");
            if (dm1 == null) dm1 = DataModel.GenerateData(100, true).OrderBy(m=>m.MyInt).SplitBy(m => m.MyInt).ToArray();
            var dm2 = DataModel.GenerateData(100, true).Select(m => new DataModel2(m));

            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\TestWriteExcel.xlsx";
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete))
            {
                var xls = new ExcelSerializer();
                xls.Serialize(fs, new IEnumerable[] { /*dm2,*/ dm1 }, wb);

                fs.Flush(true);
            }
        }

        private static IList<IList> TestReadExcel()
        {
            var wbProps = new WorkbookProperties();
            var path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\TestWriteExcel.xlsx";
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete))
            {
                var xls = new ExcelSerializer();
                var result = xls.Deserialize(fs, wbProps);

                var test = result[0].Cast<DataModel>().ToArray();

                return result;
            }
        }

        private static void TestWriteReadExcel()
        {
            var dm1 = DataModel.GenerateData(100, true);
            TestWriteExcel(dm1);
            var dm2 = TestReadExcel()[0].Cast<DataModel>().ToArray();
            //Console.WriteLine($"Serialized workbook properties {(dm1.SequenceEqual(dm2) ? "==" : "!=")} Deserialized workbook properties");
            Console.WriteLine($"Serialized data {(dm1.SequenceEqual(dm2) ? "==" : "!=")} Deserialized data");
        }

        private static void TestCurrency()
        {
            Console.OutputEncoding = Encoding.UTF8;
            foreach (var ci in CultureInfo.GetCultures(CultureTypes.AllCultures))
            {
                //System.Threading.Thread.CurrentThread.CurrentUICulture = ci; //ignored
                //System.Threading.Thread.CurrentThread.CurrentCulture = ci; //used for numeric formtting
                Console.WriteLine($"{ci.EnglishName} ({ci.ToString()}, {ci.LCID.ToString("X")}): {(12345.6789).ToString("c", ci)} {(-12345.6789).ToString("c", ci)}");
            }
        }

        private static void TestWriteReadCsv()
        {
            var path = Path.ChangeExtension(Assembly.GetExecutingAssembly().Location, ".csv");
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete))
            {
                var data1 = DataModel.GenerateData(1000, true);

                var csv = new CsvSerializer();
                csv.Serialize(new StreamWriter(fs), data1);

                fs.Flush(true);
                fs.Position = 0;

                var data2 = csv.Deserialize<DataModel>(new StreamReader(fs)).ToArray();

                Console.WriteLine($"Serialized data {(data1.SequenceEqual(data2) ? "==" : "!=")} Deserialized data");
            }
        }

        private static void TestWriteReadMultiCsv()
        {
            var path = Path.ChangeExtension(Assembly.GetExecutingAssembly().Location, ".csv");
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite | FileShare.Delete))
            {
                var dms1 = new DataModel[][]
                {
                    DataModel.GenerateData(1000, true).ToArray(), 
                    DataModel.GenerateData(500, true).ToArray()
                };

                var csv = new CsvSerializer();
                csv.Serialize(new StreamWriter(fs), dms1);
                fs.Flush(true);
                fs.Position = 0;

                var dms2 = csv.Deserialize(new StreamReader(fs), new Type[]{ typeof(DataModel), typeof(DataModel) });
                var dms2a = new DataModel[][]
                {
                    dms2[0].Cast<DataModel>().ToArray(),
                    dms2[1].Cast<DataModel>().ToArray()
                };

                Console.WriteLine($"Serialized data {(dms1[0].SequenceEqual(dms2[0].Cast<DataModel>()) ? "==" : "!=")} Deserialized data");
                Console.WriteLine($"Serialized data {(dms1[1].SequenceEqual(dms2[1].Cast<DataModel>()) ? "==" : "!=")} Deserialized data");

                Console.WriteLine($"Serialized data {(dms1[0].SequenceEqual(dms2a[0]) ? "==" : "!=")} Deserialized data");
                Console.WriteLine($"Serialized data {(dms1[1].SequenceEqual(dms2a[1]) ? "==" : "!=")} Deserialized data");
            }
        }

        private static void TestLocalization(CultureInfo ci)
        {
            string key = "MyDateTime";
            string propName = "MyPropertyName";
            string value;

            value = LocalizedStrings.GetString(key, propName, ci);
            Console.WriteLine($"HappyPath Lookup {key} = {value}");
            var lookupValue = value;

            value = LocalizedStrings.GetString(key+"BAD", propName, ci);
            Console.WriteLine($"Bad Key Lookup {key + "BAD"} = {value}");

            value = LocalizedStrings.GetString(key.ToLower(), propName, ci);
            Console.WriteLine($"Case-Insensitive Key Lookup {key.ToLower()} = {value}");

            value = LocalizedStrings.GetString(null, propName, ci);
            Console.WriteLine($"Null or empty key Lookup null = {value}");

            value = LocalizedStrings.GetString(key + "BAD", key, ci);
            Console.WriteLine($"Bad key, Lookup by propertyName = {value}");

            value = LocalizedStrings.ReverseLookup(lookupValue, ci);
            Console.WriteLine($"Reverse lookup by value \"{lookupValue}\" = {value}");

            value = LocalizedStrings.ReverseLookup(lookupValue.ToLower(), ci);
            Console.WriteLine($"Reverse lookup by value (case-insensitive) \"{lookupValue.ToLower()}\" = {value}");

            key = "Dict1";
            var dict = new Dictionary<string, string>() { { key, "This is dictionary value 1" }, { "Dict2", "This is dictionary value 2" } };
            LocalizedStrings.AddCustomResource(ci, dict);

            value = LocalizedStrings.GetString(key, ci);
            Console.WriteLine($"Custom lookup by value \"{key}\" = {value}");

            value = LocalizedStrings.GetString(key.ToLower(), ci);
            Console.WriteLine($"Custom lookup by value \"{key.ToLower()}\" = {value}");

            lookupValue = value;
            value = LocalizedStrings.ReverseLookup(lookupValue.ToLower(), ci);
            Console.WriteLine($"Reverse lookup by value (case-insensitive) \"{lookupValue.ToLower()}\" = {value}");
        }

        private static void TestTransform()
        {
            //string[,] s = new string[,]
            //{
            //    { "MyInt","MyDouble","MyChar","MyString","MyDateTime","MyDateTimeOffset","MyDate","MyTime","MyGuid","MyTimeSpan","MyIPAddress","MyVersion","MyEnum","MyBool" },
            //    { "1602974185","-28.550273309718","6","Lorem ipsum dolor sit amet","1953-01-17 19:03:06.315","","1953-01-17","0001-01-01 19:03:06.315","65a64d9a-eab4-40aa-86fc-64daf9c2325e","(02.19:03:06.315)","7.19.3.6","19.3.6.315","Seven","False" },
            //    { "1341696704","48.8613054616662","4","consectetur adipiscing elit","1956-05-09 09:42:37.029","","1956-05-09","0001-01-01 09:42:37.029","2a12f7a6-0f16-4d58-8526-cf3097ff818e","(02.09:42:37.029)","3.9.42.37","9.42.37.29","Three","False" },
            //    { "936017156","47.6244580455238","S","deserunt mollit anim id est laborum.","1974-11-01 11:12:25.062","","1974-11-01","0001-01-01 11:12:25.062","ecdb86c9-b603-4b81-997f-14117989b64b","(02.11:12:25.062)","7.11.12.25","11.12.25.62","Three","True" },
            //    { "183225232","3.54967473705749","4","cupidatat non proident","2014-01-05 17:54:52.604","","2014-01-05","0001-01-01 17:54:52.604","debfe8d1-27b7-47a3-b39a-8a0537a647e5","(02.17:54:52.604)","7.17.54.52","17.54.52.604","Six","False" },
            //    { "866943026","-30.6075286029873","1","fugiat nulla pariatur.","1998-04-14 00:47:08.688","1998-04-14 00:47:08.688","1998-04-14","0001-01-01 00:47:08.688","094c292e-1158-44ec-9625-96fde2bdafa1","(02.00:47:08.688)","1.0.47.8","0.47.8.688","Zero","True" },
            //    { "2053767806","-1.42321598316693","1","quis nostrud exercitation","2004-07-02 22:14:58.652","","2004-07-02","0001-01-01 22:14:58.652","bdd36a55-1f7e-481a-bc73-21ec2a554d71","(02.22:14:58.652)","4.22.14.58","22.14.58.652","Two","True" },
            //    { "867346412","-34.378672383902","S","velit esse cillum dolore eu","1998-05-22 05:27:50.735","","1998-05-22","0001-01-01 05:27:50.735","c65665cc-6354-47fa-bdbf-4fb6ab64625f","(00.05:27:50.735)","7.5.27.50","5.27.50.735","Four","True" },
            //    { "94995610","-18.7237879115733","4","cupidatat non proident","1967-01-15 05:56:18.49","1967-01-15 05:56:18.49","1967-01-15","0001-01-01 05:56:18.49","0aaf8ea6-b2fa-47da-a1b4-a02211896372","(00.05:56:18.490)","4.5.56.18","5.56.18.490","Three","True" },
            //    { "2022868820","30.2752689832241","1","sunt in culpa qui officia","1968-08-13 10:37:13.718","","1968-08-13","0001-01-01 10:37:13.718","ba5c61d5-fc93-45cc-80ce-534750ef06f1","(01.10:37:13.718)","9.10.37.13","10.37.13.718","Zero","" },
            //    { "17477104","-26.7293549965738","3","deserunt mollit anim id est laborum.","1961-04-06 21:06:09.147","1961-04-06 21:06:09.147","1961-04-06","0001-01-01 21:06:09.147","fc0146ef-124f-4f89-b01e-e7eaaea4d173","(01.21:06:09.147)","3.21.6.9","21.6.9.147","One","True" }
            //};
            string[,] s = new string[,]
            {
                { "MyInt","MyDouble","MyChar","MyString","MyDateTime","MyDateTimeOffset","MyDate","MyTime","MyGuid","MyTimeSpan","MyIPAddress","MyVersion","MyEnum","MyBool" },
                { "1602974185","-28.550273309718","6","Lorem ipsum dolor sit amet","1953-01-17 19:03:06.315","","1953-01-17","0001-01-01 19:03:06.315","65a64d9a-eab4-40aa-86fc-64daf9c2325e","(02.19:03:06.315)","7.19.3.6","19.3.6.315","Seven","False" },
                { "1602974185","48.8613054616662","4","consectetur adipiscing elit","1956-05-09 09:42:37.029","","1956-05-09","0001-01-01 09:42:37.029","2a12f7a6-0f16-4d58-8526-cf3097ff818e","(02.09:42:37.029)","3.9.42.37","9.42.37.29","Three","False" },
                { "936017156","47.6244580455238","S","deserunt mollit anim id est laborum.","1974-11-01 11:12:25.062","","1974-11-01","0001-01-01 11:12:25.062","ecdb86c9-b603-4b81-997f-14117989b64b","(02.11:12:25.062)","7.11.12.25","11.12.25.62","Three","True" },
                { "936017156","3.54967473705749","4","cupidatat non proident","2014-01-05 17:54:52.604","","2014-01-05","0001-01-01 17:54:52.604","debfe8d1-27b7-47a3-b39a-8a0537a647e5","(02.17:54:52.604)","7.17.54.52","17.54.52.604","Six","False" },
                { "936017156","-30.6075286029873","1","fugiat nulla pariatur.","1998-04-14 00:47:08.688","1998-04-14 00:47:08.688","1998-04-14","0001-01-01 00:47:08.688","094c292e-1158-44ec-9625-96fde2bdafa1","(02.00:47:08.688)","1.0.47.8","0.47.8.688","Zero","True" },
                { "936017156","-1.42321598316693","1","quis nostrud exercitation","2004-07-02 22:14:58.652","","2004-07-02","0001-01-01 22:14:58.652","bdd36a55-1f7e-481a-bc73-21ec2a554d71","(02.22:14:58.652)","4.22.14.58","22.14.58.652","Two","True" },
                { "867346412","-34.378672383902","S","velit esse cillum dolore eu","1998-05-22 05:27:50.735","","1998-05-22","0001-01-01 05:27:50.735","c65665cc-6354-47fa-bdbf-4fb6ab64625f","(00.05:27:50.735)","7.5.27.50","5.27.50.735","Four","True" },
                { "867346412","-18.7237879115733","4","cupidatat non proident","1967-01-15 05:56:18.49","1967-01-15 05:56:18.49","1967-01-15","0001-01-01 05:56:18.49","0aaf8ea6-b2fa-47da-a1b4-a02211896372","(00.05:56:18.490)","4.5.56.18","5.56.18.490","Three","True" },
                { "867346412","30.2752689832241","1","sunt in culpa qui officia","1968-08-13 10:37:13.718","","1968-08-13","0001-01-01 10:37:13.718","ba5c61d5-fc93-45cc-80ce-534750ef06f1","(01.10:37:13.718)","9.10.37.13","10.37.13.718","Zero","" },
                { "17477104","-26.7293549965738","3","deserunt mollit anim id est laborum.","1961-04-06 21:06:09.147","1961-04-06 21:06:09.147","1961-04-06","0001-01-01 21:06:09.147","fc0146ef-124f-4f89-b01e-e7eaaea4d173","(01.21:06:09.147)","3.21.6.9","21.6.9.147","One","True" }
            };

            TestWriteExcel(s.ToModels<DataModel>(true).OrderBy(m=>m.MyInt).SplitBy(m=>m.MyInt));
        }
#endregion
    }
}
