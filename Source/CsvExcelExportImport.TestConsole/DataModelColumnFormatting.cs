//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="DataModelColumnFormatting.cs" company="Chuck Hill">
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
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace CsvExcelExportImport.TestConsole
{
    /// <summary>
    /// Test data model for validating Serialization and Deserialization for both CSV and Excel.
    /// </summary>
    public class DataModelColumnFormatting
    {
        #region Test Properties

        [XlColumn]
        public double Double { get; set; }

        [XlColumn(Format = "f")]
        public double? DoubleF { get; set; }

        [XlColumn(Format = "f0")]
        public double? DoubleF0 { get; set; }

        [XlColumn(Format = "f3")]
        public double? DoubleF3 { get; set; }

        [XlColumn(Format = "f3, mg")]
        public double? DoubleF3_mg { get; set; }

        [XlColumn(Format = "f3,mg")]
        public double? DoubleF3mg { get; set; }

        [XlColumn(Hidden = true)]
        public string DoseUnits { get; set; }

        [XlColumn(Format = "f3,-1")]
        public double? DoubleF3DoseUnits { get; set; }

        [XlColumn(Format = "n")]
        public double? DoubleN { get; set; }

        [XlColumn(Format = "n0")]
        public double? DoubleN0 { get; set; }

        [XlColumn(Format = "n3")]
        public double? DoubleN3 { get; set; }

        [XlColumn(Format = "n, -6")]
        public double? DoubleNDoseUnits { get; set; }

        [XlColumn(Format = "c")]
        public double DoubleC { get; set; }

        [XlColumn(Format = "c0")]
        public double? DoubleC0 { get; set; }

        [XlColumn(Format = "c3")]
        public double? DoubleC3 { get; set; }

        [XlColumn(Format = "c, /yr")]
        public double? DoubleC_yr { get; set; }

        [XlColumn(Format = "c,1")]
        public double? DoubleCUnits { get; set; }

        [XlColumn(Hidden = true)]
        public string MoneyUnits { get; set; }

        [XlColumn]
        public DateTime DateTime { get; set; }

        [XlColumn(Format = "d")]
        public DateTime DateTimed { get; set; }

        [XlColumn(Format = "D")]
        public DateTime DateTimeD { get; set; }

        [XlColumn(Format = "t")]
        public DateTime DateTimet { get; set; }

        [XlColumn(Format = "T")]
        public DateTime DateTimeT { get; set; }

        [XlColumn(Format = "f")]
        public DateTime DateTimef { get; set; }

        [XlColumn(Format = "F")]
        public DateTime DateTimeF { get; set; }

        [XlColumn(Format = "g")]
        public DateTime DateTimeg { get; set; }

        [XlColumn(Format = "G")]
        public DateTime DateTimeG { get; set; }

        [XlColumn(Format = "M")]
        public DateTime DateTimeM { get; set; }

        [XlColumn(Format = "r")]
        public DateTime DateTimer { get; set; }

        [XlColumn(Format = "s")]
        public DateTime DateTimes { get; set; }

        [XlColumn(Format = "u")]
        public DateTime DateTimeu { get; set; }

        [XlColumn(Format = "U")]
        public DateTime DateTimeU { get; set; }

        [XlColumn(Format = "x")]
        public DateTime DateTimex { get; set; }

        [XlColumn(Format = "X")]
        public DateTime DateTimeX { get; set; }

        [XlColumn(Format = "yyyy-MMM-dd HH:mm")]
        public DateTime DateTimeCustom { get; set; }
        #endregion // Test Properties

        #region Test Data

        private static readonly string[] DoseUnitValues = new[]
        {
            " mg",
            " \x00b5g",
            " ml",
            " \x00b5l",
            " l",
            " fl oz",
            " pt",
            " qt",
            " gal",
            " tsp"
        };

        private static readonly string[] MoneyUnitValues = new[]
        {
            "/year",
            "/month",
            "/day",
            "/hour",
            "/minute",
            " hourly",
            " daily",
            " monthly",
            " quarterly",
            " semi-annually",
            " annually",
            " bi-annually"
        };

        #endregion

        public DataModelColumnFormatting() { } // for deserialization

        /// <summary>
        /// Generate an array of reproducible pseudo-random data for testing.
        /// </summary>
        /// <param name="count">Number of items in array</param>
        /// <param name="highPrecision">True for enhanced precision using milliseconds in datetimes</param>
        /// <returns>Array of pseudo-random data</returns>
        public static DataModelColumnFormatting[] GenerateData(int count)
        {
            var results = new DataModelColumnFormatting[count];
            Random rand = new Random(1);

            var props = typeof(DataModelColumnFormatting).GetProperties();

            for (int i = 0; i < count; i++)
            {
                // Precompute to make date & time numerical values the same in a given row.
                var dt = new DateTime(rand.Next(1950, 2021), rand.Next(1, 13), rand.Next(1, 29), rand.Next(0, 24), rand.Next(0, 60), rand.Next(0, 60), rand.Next(0, 1000));
                var d = rand.NextDouble() * 10000;

                results[i] = new DataModelColumnFormatting();
                foreach(var p in props)
                {
                    if (p.Name == "DoseUnits")
                    {
                        p.SetValue(results[i], rand.Next(0, 3) == 0 ? null : DoseUnitValues[rand.Next(0, DoseUnitValues.Length)]);
                        continue;
                    }
                    if (p.Name == "MoneyUnits")
                    {
                        p.SetValue(results[i], rand.Next(0, 3) == 0 ? null : MoneyUnitValues[rand.Next(0, MoneyUnitValues.Length)]);
                        continue;
                    }

                    if (p.PropertyType == typeof(DateTime))
                    {
                        p.SetValue(results[i], dt);
                        continue;
                    }

                    if (p.PropertyType == typeof(double))
                    {
                        p.SetValue(results[i], d);
                        continue;
                    }

                    if (p.PropertyType == typeof(double?))
                    {
                        p.SetValue(results[i], rand.Next(0, 3) == 0 ? (double?)null : d);
                        continue;
                    }
                }
            }

            return results;
        }

        public override string ToString() => $"{Double},{DateTime}";
    }
}
