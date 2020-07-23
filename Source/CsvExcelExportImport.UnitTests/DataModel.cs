﻿//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="DataModel.cs" company="Chuck Hill">
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
using System.Collections.Generic;
using System.ComponentModel;

namespace CsvExcelExportImport.UnitTests
{
    /// <summary>
    /// Test data model for validating Serialization and Deserialization for both CSV and Excel.
    /// </summary>
    [XlWorkheetTab("TestDataModel")]
    public class DataModel : IEquatable<DataModel>, IEqualityComparer<DataModel>
    {
        public enum Numbers
        {
            Zero, 
            [XlEnumName("NumOne")] One, 
            Two,
            [Description("NumThree")] Three, 
            Four,
            [XlEnumName("NumFive")] Five, 
            Six,
            [Description("NumSeven")] Seven, 
            Eight,
            [XlEnumName("NumNine")] Nine, 
            Ten
        }

        #region Test Properties
        public int MyInt { get; set; }

        [XlColumn("YOURDOUBLE", Format = "f3", Frozen = true)]
        public double MyDouble { get; set; }

        [XlColumn("YourDecimal", Format = "f3")]
        public decimal MyDecimal { get; set; }

        [XlColumn("YourChar", HasFilter = true, Justification = Justified.Center)]
        public char MyChar { get; set; }

        public string MyString { get; set; }

        public DateTime MyDateTime { get; set; }

        public DateTimeOffset? MyDateTimeOffset { get; set; }

        [XlColumn("YourDate", Format = "d")]
        public DateTime MyDate { get; set; }

        [XlColumn("YourTime", Format = "t")]
        public DateTime MyTime { get; set; }

        public Guid MyGuid { get; set; }

        public TimeSpan MyTimeSpan { get; set; }

        public Version MyVersion { get; set; }

        [XlColumn(TranslateData = true)]
        public Numbers MyEnum { get; set; }

        public int FieldIgnored;  // public and private fields are ignored upon export

        public int PropIgnored { get; private set; } // read-only properties are ignored upon export

        [XlIgnore]
        public int PropIgnored2 { get; set; } // Explicitly ignored properties are ignored upon export

        [XlColumn(TranslateData = true)]
        public bool? MyBool { get; set; }
        #endregion // Test Properties

        #region Fake Test Data
        private static readonly string[] Phrases = new string[]
        {
            "Lorem ipsum dolor sit amet",
            "consectetur adipiscing elit",
            "sed do eiusmod tempor",
            "incididunt ut labore et",
            "dolore magna aliqua.",
            "Ut enim ad minim veniam",
            "quis nostrud exercitation",
            "ullamco laboris nisi ut",
            "aliquip ex ea commodo consequat.",
            "Duis aute irure dolor in",
            "reprehenderit in voluptate",
            "velit esse cillum dolore eu",
            "fugiat nulla pariatur.",
            "Excepteur sint occaecat",
            "cupidatat non proident",
            "sunt in culpa qui officia",
            "deserunt mollit anim id est laborum.",
            null
        };

        private static readonly Guid[] Guids = new Guid[]
        {
            new Guid("2d68dacc-34b1-4277-ac31-d6384055ff88"),
            new Guid("ba232d38-caca-4340-940a-ed57b8f1aee7"),
            new Guid("11537a9a-6396-4927-8e8f-34289a8a827e"),
            new Guid("ce064b69-ca35-4a7d-9299-d339512c90e8"),
            new Guid("46c6394d-f0e1-433c-9387-1bf15773f825"),
            new Guid("6c2204d6-3b36-486d-a3e9-1639179adc54"),
            new Guid("16a101dd-936e-48e7-bfd5-469dca1c57ba"),
            new Guid("37b01808-42fd-4083-9b4e-17b7620e49c2"),
            new Guid("e95a293f-5cee-4e44-8434-056b2fdf8f64"),
            new Guid("9062594f-b9af-4e43-ae8a-3ee7babebfbd"),
        };
        #endregion // Test Data

        public DataModel() { } // for deserialization

        /// <summary>
        /// Generate an array of reproducible pseudo-random data for testing.
        /// </summary>
        /// <param name="count">Number of items in array</param>
        /// <param name="highPrecision">True for enhanced precision using milliseconds in datetimes</param>
        /// <returns>Array of pseudo-random data</returns>
        public static IEnumerable<DataModel> GenerateData(int count, bool highPrecision = false)
        {
            var results = new DataModel[count];
            Random rand = new Random(1);

            for (int i = 0; i < count; i++)
            {
                int ms = highPrecision ? rand.Next(0, 1000) : 0;
                // Precompute to make date & time numerical values the same in a given row.
                var dto = new DateTimeOffset(rand.Next(1950, 2021), rand.Next(1, 13), rand.Next(1, 29), rand.Next(0, 24), rand.Next(0, 60), rand.Next(0, 60), ms, TimeSpan.FromHours(rand.Next(0, 15)));
                var dt = dto.DateTime;

                yield return new DataModel()
                {
                    MyInt = rand.Next(-10, 10),
                    MyDouble = (rand.NextDouble() * 100) - 50.0,
                    MyDecimal = (decimal)((rand.NextDouble() * 100) - 50.0),
                    MyChar = "123456SD"[rand.Next(0, 8)],
                    MyString = Phrases[rand.Next(0, Phrases.Length - 1)],
                    MyDateTimeOffset = rand.Next(0, 2) == 0 ? dto : (DateTimeOffset?)null,
                    MyDateTime = dt,
                    MyDate = dt.Date,
                    MyTime = new DateTime(1900, 1, 1, dt.Hour, dt.Minute, dt.Second, dt.Millisecond),
                    MyGuid = Guids[rand.Next(0, 10)],
                    MyTimeSpan = new TimeSpan(rand.Next(0, 4), dt.Hour, dt.Minute, dt.Second, dt.Millisecond),
                    MyVersion = new Version(dt.Hour, dt.Minute, dt.Second, dt.Millisecond),
                    MyEnum = (Numbers)(int)rand.Next(0, 11),
                    FieldIgnored = rand.Next(1, 1000),
                    PropIgnored = rand.Next(1, 1000),
                    PropIgnored2 = rand.Next(1, 1000),
                    MyBool = rand.Next(0, 3) == 0 ? (bool?)null : rand.Next(0, 2) == 0 ? true : false,
                };
            }

            yield break;
        }

        public bool Equals(DataModel o) => this.Equals(this, o);
        public override bool Equals(object obj) => this.Equals((DataModel)obj);
        public override int GetHashCode() => this.GetHashCode(this);
        public int GetHashCode(DataModel obj) => throw new NotImplementedException();
        public override string ToString() => $"{MyInt}, {MyDouble}, {MyDecimal}, {MyChar}, {(MyBool.HasValue ? MyBool.Value.ToString() : "(null)")}";

        public bool Equals(DataModel x, DataModel y)
        {
            if (x == null && y == null) return true;
            if (x == null || y == null) return false;

            var v00 = x.MyInt == y.MyInt;
            var v01 = Math.Round(x.MyDouble, 9).Equals(Math.Round(y.MyDouble, 9));
            var v02 = x.MyDecimal == y.MyDecimal;
            var v03 = x.MyChar == y.MyChar;
            var v04 = x.MyString == y.MyString;
            var v05 = x.MyDateTime == y.MyDateTime;
            var v06 = x.MyDateTimeOffset?.Ticks == y.MyDateTimeOffset?.Ticks;
            var v07 = x.MyDate == y.MyDate;
            var v08 = x.MyTime == y.MyTime;
            var v09 = x.MyGuid == y.MyGuid;
            var v10 = x.MyTimeSpan == y.MyTimeSpan;
            var v11 = x.MyVersion == y.MyVersion;
            var v12 = x.MyEnum == y.MyEnum;
            var v13 = x.FieldIgnored != y.FieldIgnored;
            var v14 = x.PropIgnored != y.PropIgnored;
            var v15 = x.PropIgnored2 != y.PropIgnored2;

            bool vx = v00 && v01 && v02 && v03 && v04 && v05 && v06 && v07 && v08 && v09 && v10 && v11 && v12 && v13 && v14 && v15;
            return vx;
        }
    }
}
