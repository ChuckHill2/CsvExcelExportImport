// --------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="WorkbookProperties.cs" company="Chuck Hill">
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
// --------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;

namespace CsvExcelExportImport
{
    /// <summary>
    /// All the Excel workbook properties to write or read. String values must be already
    /// localized/translated.Upon serialization, these values are entered into the workbook.
    /// Upon deserialization, this object is populated by the contents of the workbook.
    /// All properties are optional and (null) may be passed for this WorkbookProperties 
    /// object. If (null) upon serialization, all default values will be used. If (null) 
    /// upon deserialization, no workbook properties will be returned.
    /// </summary>
    public class WorkbookProperties
    {
        private Dictionary<string, string> _extraProperties;
        private List<string> _worksheetHeading;
        private RegionInfoOverride _regionInfo;

        /// <summary>
        /// Excel Green. The default worksheet theme color. Very close to Color.SeaGreen.
        /// </summary>
        public static readonly Color ExcelGreen = Color.FromArgb(33, 115, 70);

        #region Standard Excel Document Properties
        /// <summary>
        /// Standard Excel document metadata property.
        /// Optional title of this workbook.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Standard Excel document metadata property.
        /// Optional subject this workbook is about.
        /// </summary>
        public string Subject { get; set; }

        /// <summary>
        /// Standard Excel document metadata property.
        /// Optional name of user that requested this document.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Standard Excel document metadata property.
        /// </summary>
        public string Manager { get; set; }

        /// <summary>
        /// Standard Excel document metadata property.
        /// Optional name of the company that owns this data.
        /// If undefined, defaults to "".
        /// This field is also used by in the print page footer as "[Company] Confidential"
        /// </summary>
        public string Company { get; set; }

        /// <summary>
        /// Standard Excel document metadata property.
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// Standard Excel document metadata property.
        /// Optional comma-delimited list of keywords that may be used in a search for this document.
        /// </summary>
        public string Keywords { get; set; }

        /// <summary>
        /// Standard Excel document metadata property.
        /// Optional comments for this workbook.
        /// </summary>
        public string Comments { get; set; }

        /// <summary>
        /// Standard Excel document metadata property.
        /// May be used to define the source of this spreadsheet data.
        /// </summary>
        public Uri HyperlinkBase { get; set; }

        /// <summary>
        /// Standard Excel document metadata property.
        /// </summary>
        public string Status { get; set; }
        #endregion Standard Excel Document Properties

        /// <summary>
        /// Various shades of this color are used for the worksheet tab and column headings.
        /// If undefined, uses Excel Green. For uncolored worksheets, use Color.White.
        /// </summary>
        public Color? ThemeColor { get; set; }

        /// <summary>
        /// Optional culture to use for translation and formatting. If undefined, 
        /// serialization uses the CurrentUICulture of the thread that is executing 
        /// ExcelSerializer.Serialize(). This property is visible in Excel advanced 
        /// workbook properties popup dialog, custom tab. Upon deserialization, this 
        /// property holds the culture that this workbook was originally serialized 
        /// into. Note: If the Invariant culture is used, translation is disabled and
        /// formatting defaults to en-US.
        /// </summary>
        public CultureInfo Culture { get; set; }

        /// <summary>
        /// Optional additional caller-defined properties. All are visible in
        /// the Excel advanced workbook properties popup dialog, custom tab.
        /// Values will be quietly truncated by Excel if they exceed 255 characters.
        /// Any string key/value pairs may be used.
        /// </summary>
        public Dictionary<string, string> ExtraProperties
        {
            get
            {
                if (this._extraProperties == null)
                    this._extraProperties = new Dictionary<string, string>(0, StringComparer.CurrentCultureIgnoreCase);

                return this._extraProperties;
            }
        }

        /// <summary>
        /// In the first row of the worksheet, an optional worksheet header that spans all the 
        /// columns, may be supplied. The header may contain multiple lines (not worksheet rows). The 
        /// height of this row will be automatically adjusted accordingly. The index of each string in 
        /// this list represents the index in the collection of enumerable arrays passed to Serialize() 
        /// where each enumerable array represents a single worksheet. If the header index is is out of 
        /// range or the index refers to a null string, there will be no worksheet header. These
        /// headings must be pre-localized.
        /// </summary>
        public List<string> WorksheetHeading
        {
            get
            {
                if (this._worksheetHeading == null)
                    this._worksheetHeading = new List<string>();

                return this._worksheetHeading;
            }
        }

        /// <summary>
        /// Set worksheet heading justification for all worksheets.
        /// Choices are: Left, Center, Right, and Auto, where Auto is the default.
        /// </summary>
        public Justified WorksheetHeadingJustification { get; set; }

        /// <summary>
        /// Custom DateTime column formatting. Values use the .NET datetime format specifiers.
        /// See http://blog.stevex.net/string-formatting-in-csharp/ <br/>
        /// Month names and days of week are localized. This is normally used for the custom
        /// date and time patterns.
        /// </summary>
        public RegionInfoOverride CustomRegionInfo
        {
            get
            {
                if (_regionInfo == null)
                {
                    _regionInfo = new RegionInfoOverride();
                }

                return _regionInfo;
            }
        }

        /// <summary>
        /// Custom column formatting information. Overrides the current culture regional information. 
        /// Month names, days of week, and AM/PM are automatically localized.
        /// If undefined, the current culture regional information is used.
        /// </summary>
        public class RegionInfoOverride : object
        {
            /// <summary>
            /// Custom value override for formatting dates. Used in the 'D', 'F' and 'f' 
            /// format specifiers. If undefined, defaults to the the 
            /// value used within the culture defined during serialization.
            /// </summary>
            public string LongDatePattern { get; set; }

            /// <summary>
            /// Custom value override for formatting dates. Used in the 'T', 'F', and 'G' 
            /// format specifiers. If undefined, defaults to the the 
            /// value used within the culture defined during serialization.
            /// </summary>
            public string LongTimePattern { get; set; }

            /// <summary>
            /// Custom value override for formatting dates. Used in the 'd', 'G', 'g' 
            /// format specifiers. If undefined, defaults to the the 
            /// value used within the culture defined during serialization.<br />
            /// Warning: If the date separator is a backslash/escape character ('\'),
            /// the pattern must include 2 consecutive backslash characters. e.g.
            /// to escape the escape character.
            /// </summary>
            public string ShortDatePattern { get; set; }

            /// <summary>
            /// Custom value override for formatting dates. Used in the 't', 'f', 'g' 
            /// format specifiers. If undefined, defaults to the the 
            /// value used within the culture defined during serialization.
            /// </summary>
            public string ShortTimePattern { get; set; }

            /// <summary>
            /// Custom value override for formatting numbers. Used in the 'n', 'f'
            /// format specifiers. If undefined, defaults
            /// to the value used within the culture defined during serialization.
            /// </summary>
            public string NumberGroupSeparator { get; set; }

            /// <summary>
            /// Custom value override for formatting numbers. Used in the 'n', 'f' 
            /// format specifiers. If undefined, defaults
            /// to the value used within the culture defined during serialization.
            /// </summary>
            public int[] NumberGroupSizes { get; set; }

            /// <summary>
            /// Custom value override for formatting numbers. Used in the 'n', 'f' 
            /// format specifiers. If undefined, defaults
            /// to the value used within the culture defined during serialization.
            /// </summary>
            public string NumberDecimalSeparator { get; set; }

            /// <summary>
            /// Custom value override for formatting numbers. Used in the 'n', 'f' 
            /// format specifiers. If undefined, defaults
            /// to the value used within the culture defined during serialization.
            /// </summary>
            public int NumberNegativePattern { get; set; }

            /// <summary>
            /// Custom value override for formatting numbers. Used in the 'n', 'f' 
            /// format specifiers. If undefined, defaults
            /// to the value used within the culture defined during serialization.
            /// </summary>
            public int NumberDecimalDigits { get; set; }

            /// <summary>
            /// Custom value override for formatting currency. Used in the 'c' 
            /// format specifier. If undefined, defaults 
            /// to the the value used within the culture defined during serialization.
            /// </summary>
            public string CurrencySymbol { get; set; }

            /// <summary>
            /// Custom value override for formatting currency. Used in the 'c' 
            /// format specifier. If undefined, defaults 
            /// to the the value used within the culture defined during serialization.
            /// </summary>
            public string CurrencyGroupSeparator { get; set; }

            /// <summary>
            /// Custom value override for formatting currency. Used in the 'c' 
            /// format specifier. If undefined, defaults 
            /// to the the value used within the culture defined during serialization.
            /// </summary>
            public int[] CurrencyGroupSizes { get; set; }

            /// <summary>
            /// Custom value override for formatting currency. Used in the 'c' 
            /// format specifier. If undefined, defaults 
            /// to the the value used within the culture defined during serialization.
            /// </summary>
            public string CurrencyDecimalSeparator { get; set; }

            /// <summary>
            /// Custom value override for formatting currency. Used in the 'c' 
            /// format specifier. If undefined, defaults 
            /// to the the value used within the culture defined during serialization.
            /// </summary>
            public int CurrencyDecimalDigits { get; set; }

            /// <summary>
            /// Custom value override for formatting currency.  Used in the 'c' 
            /// format specifier. If undefined, defaults 
            /// to the the value used within the culture defined during serialization.
            /// </summary>
            public int CurrencyPositivePattern { get; set; }

            /// <summary>
            /// Custom value override for formatting currency. Used in the 'c' 
            /// format specifier. If undefined, defaults 
            /// to the the value used within the culture defined during serialization.
            /// </summary>
            public int CurrencyNegativePattern { get; set; }

            /// <summary>
            /// If some or all the settings are already in a class, just bulk copy all the setting properties into this object by
            /// matching name and type. Mismatched properties are ignored. Any additional properties may be added separately or 
            /// left to the default culture settings.  
            /// </summary>
            /// <typeparam name="T">Any class object.</typeparam>
            /// <param name="src">Any class containing the matching name and property types.</param>
            public void SetAllValues<T>(T src) where T : class
            {
                var srcProps = src.GetType().GetProperties();
                var dstProps = this.GetType().GetProperties();
                foreach (var srcP in srcProps)
                {
                    var dstP = dstProps.FirstOrDefault(p => p.Name.Equals(srcP.Name) && p.PropertyType == srcP.PropertyType);
                    if (dstP == null) continue;
                    dstP.SetValue(this, srcP.GetValue(src));
                }
            }

            internal RegionInfoOverride()
            {
                // Initialize as undefined.
                NumberNegativePattern = -1;
                NumberDecimalDigits = -1;
                CurrencyDecimalDigits = -1;
                CurrencyPositivePattern = -1;
                CurrencyNegativePattern = -1;
            }

            internal RegionInfoOverride Clone()
            {
                // We cannot reference this object, so we just duplicate its content.
                return (RegionInfoOverride)MemberwiseClone();
            }

            internal void InitializeUndefined(CultureInfo ci)
            {
                // Populate with the current culture defaults.
                if (ci == null) ci = CultureInfo.CurrentCulture;

                var dti = ci.DateTimeFormat;
                var nf = ci.NumberFormat;

                if (LongDatePattern == null)          LongDatePattern = dti.LongDatePattern.Replace("dddd, ", string.Empty);
                if (LongTimePattern == null)          LongTimePattern = dti.LongTimePattern;
                if (ShortDatePattern == null)         ShortDatePattern = dti.ShortDatePattern;
                if (ShortTimePattern == null)         ShortTimePattern = dti.ShortTimePattern;
                if (NumberGroupSeparator == null)     NumberGroupSeparator = nf.NumberGroupSeparator;
                if (NumberGroupSizes == null)         NumberGroupSizes = nf.NumberGroupSizes;
                if (NumberDecimalSeparator == null)   NumberDecimalSeparator = nf.NumberDecimalSeparator;
                if (NumberNegativePattern == -1)      NumberNegativePattern = nf.NumberNegativePattern;
                if (NumberDecimalDigits == -1)        NumberDecimalDigits = nf.NumberDecimalDigits;
                if (CurrencySymbol == null)           CurrencySymbol = nf.CurrencySymbol;
                if (CurrencyGroupSeparator == null)   CurrencyGroupSeparator = nf.CurrencyGroupSeparator;
                if (CurrencyGroupSizes == null)       CurrencyGroupSizes = nf.CurrencyGroupSizes;
                if (CurrencyDecimalSeparator == null) CurrencyDecimalSeparator = nf.CurrencyDecimalSeparator;
                if (CurrencyDecimalDigits == -1)      CurrencyDecimalDigits = nf.CurrencyDecimalDigits;
                if (CurrencyPositivePattern == -1)    CurrencyPositivePattern = nf.CurrencyPositivePattern;
                if (CurrencyNegativePattern == -1)    CurrencyNegativePattern = nf.CurrencyNegativePattern;
            }

            internal string Serialize()
            {
                // Cannot use XML or JSON serialization because the resulting string would exceed 255 characters.
                // The maximum Excel custom property size is 255 characters. Excess characters are quietly truncated.
                // This formatted string should be well under 128 characters. 

                return $"{LongDatePattern}|" +
                       $"{LongTimePattern}|" +
                       $"{ShortDatePattern}|" +
                       $"{ShortTimePattern}|" +
                       $"{NumberGroupSeparator}|" +
                       $"{string.Join(",", NumberGroupSizes ?? new int[0])}|" +
                       $"{NumberDecimalSeparator}|" +
                       $"{NumberNegativePattern}|" +
                       $"{NumberDecimalDigits}|" +
                       $"{CurrencySymbol}|" +
                       $"{CurrencyGroupSeparator}|" +
                       $"{string.Join(",", CurrencyGroupSizes ?? new int[0])}|" +
                       $"{CurrencyDecimalSeparator}|" +
                       $"{CurrencyDecimalDigits}|" +
                       $"{CurrencyPositivePattern}|" +
                       $"{CurrencyNegativePattern}";
            }

            internal void Deserialize(string s, CultureInfo ci)
            {
                if (string.IsNullOrWhiteSpace(s))
                {
                    InitializeUndefined(ci);
                    return;
                }

                try
                {
                    var e = s.Split('|');

                    this.LongDatePattern = e[0];
                    this.LongTimePattern = e[1];
                    this.ShortDatePattern = e[2];
                    this.ShortTimePattern = e[3];
                    this.NumberGroupSeparator = e[4];
                    this.NumberGroupSizes = e[5].Split(',').Select(int.Parse).ToArray();
                    this.NumberDecimalSeparator = e[6];
                    this.NumberNegativePattern = int.Parse(e[7]);
                    this.NumberDecimalDigits = int.Parse(e[8]);
                    this.CurrencySymbol = e[9];
                    this.CurrencyGroupSeparator = e[10];
                    this.CurrencyGroupSizes = e[11].Split(',').Select(int.Parse).ToArray();
                    this.CurrencyDecimalSeparator = e[12];
                    this.CurrencyDecimalDigits = int.Parse(e[13]);
                    this.CurrencyPositivePattern = int.Parse(e[14]);
                    this.CurrencyNegativePattern = int.Parse(e[15]);
                }
                catch (Exception)
                {
                    // Something bad happened, partial assignment, so we just populate the rest with the defaults. 
                    InitializeUndefined(ci);
                }
            }
        }
    }
}
