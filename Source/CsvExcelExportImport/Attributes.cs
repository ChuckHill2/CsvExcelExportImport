//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="Attributes.cs" company="Chuck Hill">
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
using System.Runtime.CompilerServices;

namespace CsvExcelExportImport
{
    /// <summary>
    /// XlColumnAttribute Justification property possible values.
    /// </summary>
    public enum Justified
    {
        /// <summary>
        /// Automatic justification by column type (e.g. strings to the left and numbers and dates to the right). This is the default.
        /// </summary>
        Auto,

        /// <summary>
        /// Justifies all data in column to the left side of column.
        /// </summary>
        Left,

        /// <summary>
        /// Centers all data in column.
        /// </summary>
        Center,

        /// <summary>
        /// Justifies all data in column to the right side of column.
        /// </summary>
        Right
    }

    /// <summary>
    /// Exported Excel worksheet tab translation key. Not used in CSV. If undefined, 
    /// class name is the translation key. If the translation key doesn't exist in 
    /// the resource file or culture is the invariant culture, all translations are 
    /// disabled and the class name is used as the worksheet tab name.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct, AllowMultiple = false)]
    public class XlWorkheetTabAttribute : Attribute
    {
        private readonly string _id;

        /// <summary>
        /// Initializes a new instance of the XlWorkheetTabAttribute class with the resource ID key.
        /// </summary>
        /// <param name="id">Resource translation key.</param>
        public XlWorkheetTabAttribute(string id)
        {
            this._id = id;
        }

        /// <summary>
        /// Exported Excel worksheet tab translation key.
        /// </summary>
        internal string Id { get => _id; }
    }

    /// <summary>
    /// Exported Excel/CSV enum translation key. If undefined, the enum value name is 
    /// the translation key. If the translation key doesn't exist in the resource file 
    /// or culture is the invariant culture, all translations are disabled. In addition, 
    /// column enums are only translated when XlColumnAttribute.TranslateData == true.
    /// System.ComponentModel.DescriptionAttribute may be used interchangeably with XlEnumNameAttribute.
    /// </summary>
    /// <remarks>
    /// Example:
    /// <code>
    ///   enum MyEnum {
    ///     [XlEnumName("TranslationKey1")] MyValue1,
    ///     [Description("TranslationKey2")] MyValue2
    ///   }
    /// </code>
    /// </remarks>
    [AttributeUsage(AttributeTargets.Field, AllowMultiple = false)]
    public class XlEnumNameAttribute : Attribute
    {
        private string _id;

        /// <summary>
        /// Initializes a new instance of the XlEnumNameAttribute class with the resource ID key.
        /// </summary>
        /// <param name="id">Resource translation key.</param>
        public XlEnumNameAttribute(string id)
        {
            this._id = id;
        }

        /// <summary>
        /// Exported Excel/CSV enum translation key.
        /// </summary>
        internal string Id { get => _id; }
    }

    /// <summary>
    /// Ignore flag to NOT serialize/deserialize this property. Actually, any
    /// attribute containing "ignore" substring will do the same thing.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class XlIgnoreAttribute : Attribute
    {
    }

    /// <summary>
    /// Column descriptor containing the column attributes used for serialization/deserialization of Excel and CSV.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class XlColumnAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the XlColumnAttribute class.
        /// </summary>
        public XlColumnAttribute()
        {
        }

        /// <summary>
        /// Initializes a new instance of the XlColumnAttribute class with the resource ID key.
        /// </summary>
        /// <param name="id">
        /// Column heading localized string resource ID. If the specified culture is the 
        /// invariant culture this ID is ignored and the property name is used as the column 
        /// heading. Otherwise, if <i>Id</i> is undefined or does not refer to a localized 
        /// string then if the property name does not refer to a localized string then the 
        /// property name is used as the column heading name.
        /// </param>
        public XlColumnAttribute(string id)
        {
            this.Id = id;
        }

        /// <summary>
        /// Column heading localized string resource ID. If the specified culture is the 
        /// invariant culture this ID is ignored and the property name is used as the column 
        /// heading. Otherwise, if <i>Id</i> is undefined or does not refer to a localized 
        /// string then if the property name does not refer to a localized string then the 
        /// property name is used as the column heading name.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Formats the column using the specified format and the formatting
        /// conventions of the current culture. The format values are the same 
        /// as used in <i>value.ToString(format)</i> but does not really convert the 
        /// value to a string, just sets the format of the column.
        /// Invalid or missing codes default to Excel default format.<br />
        /// See http://www.cheat-sheets.org/saved-copy/msnet-formatting-strings.pdf <br />
        /// See http://blog.stevex.net/string-formatting-in-csharp/
        /// </summary>
        /// <remarks>
        /// Valid DateTime Format Codes (en-US examples, case-sensitive):
        /// <code>
        ///   d = ShortDatePattern¹ (ex. 2/4/2020) 
        ///   D = LongDatePattern¹ (ex. Tuesday, February 4, 2020)
        ///   t = ShortTimePattern¹ (ex. 1:34 PM)
        ///   T = LongTimePattern¹ (ex. 1:34:26 PM)
        ///   f = LongDatePattern¹ + ShortTimePattern¹ aka FullShortDateTimePattern (ex. Tuesday, February 04, 2020 1:34 PM)
        ///   F = LongDatePattern¹ + LongTimePattern¹ aka FullLongDateTimePattern (ex. Tuesday, February 04, 2020 1:34:26 PM)
        ///   g = ShortDatePattern¹ + ShortTimePattern¹ aka GeneralShortTimePattern (ex. 2/4/2020 1:34 PM)
        ///   G = ShortDatePattern¹ + LongTimePattern¹ aka GeneralLongTimePattern (ex. 2/4/2020 01:34:26 PM)
        ///   M = MonthDayPattern (ex. February 04)
        ///   o = Round-trip date/time pattern² (ex. 2020-02-04T13:34:26.775)
        ///   O = Round-trip date/time pattern² (ex. 2020-02-04T13:34:26.775)
        ///   r = RFC1123Pattern³ (ex. 04 Feb 2020 13:34:26 GMT)
        ///   R = RFC1123Pattern³ (ex. 04 Feb 2020 13:34:26 GMT)
        ///   s = SortableDateTimePattern (ex. 2020-02-04T13:34:26)
        ///   S = Invalid. Defaults to "g" format.
        ///   u = UniversalSortableDateTimePattern (ex. 2020-02-04 13:34Z)
        ///   U = UniversalSortableDateTimePattern (ex. 2020-02-04 13:34:26Z)
        ///   x = CustomShortDateTimePattern⁴ (ex. 02/04/2020 01:34 PM)
        ///   X = CustomLongDateTimePattern⁴ (ex. 02/04/2020 01:34:26 PM)
        ///   y = YearMonthPattern (ex. Feb 2020)
        ///   Y = YearMonthPattern (ex. Feb 2020)
        /// </code>
        /// ¹May be overridden by custom format info.<br />
        /// ²This pattern cannot be accurately represented in Excel as it does not understand timezone offsets. This is not a localized pattern.<br />
        /// ³The values are not automatically converted into GMT. It is up to the caller to convert them first.<br />
        /// ⁴Localized pattern contains leading zeros so all values in the column may align.<br />
        /// <br />
        /// Valid Number Format Codes (case-insensitive):
        /// <code>
        ///   f = Fixed point
        ///   n = Number with commas for thousands
        ///   c = Currency
        /// </code>
        /// The numeric format styling code may contain 2 comma delimited fields.<br />
        /// The first field contains the code char + n where the n represents the number of digits after the decimal (aka precision). If the precision is not defined, it defaults to 2.<br />
        /// The optional 2nd field contains the units string(including leading and trailing whitespace). However if the 2nd field contains an integer, it is used as a relative column index to the column containing the units.<br />
        /// Examples: format="f3, mg" will be displayed as "5.320 mg". When precision is undefined, the precision==2. To To force zero digits after the decimal, use "f0".<br />
        /// <br />
        /// Valid TimeSpan Format Codes (case-sensitive):
        /// <code>
        ///   g = low precision e.g. d.hh:mm (default)
        ///   G = high precision e.g. d.hh:mm.ss.fff
        /// </code>
        /// </remarks>
        public string Format { get; set; }

        /// <summary>
        /// All columns preceding this column are fixed and do not move when scrolling 
        /// left and right. They are always visible on the page while scrolling.
        /// </summary>
        public bool Frozen { get; set; }

        /// <summary>
        /// Excel column has a grouped value filter dropdown.
        /// </summary>
        public bool HasFilter { get; set; }

        /// <summary>
        /// Set column justification.
        /// Choices are: Left, Center, Right, and Auto, where Auto is the default.
        /// </summary>
        public Justified Justification { get; set; }

        /// <summary>
        /// Column is hidden from view.
        /// </summary>
        public bool Hidden { get; set; }

        /// <summary>
        /// Translate string data in this column by using the actual value as the ID/Key.
        /// However if the specified culture is the invariant culture, all translations
        /// are disabled for the entire Excel document.
        /// </summary>
        public bool TranslateData { get; set; }

        /// <summary>
        /// Gets or sets fixed column width. Units are in characters. A 'character' is defined as
        /// the width of a numeric character of the Excel default font (Calibri,11). Autoformat
        /// adds 0.29 characters for margins. When this is set, AutoWrap is enabled.
        /// </summary>
        public int MaxWidth { get; set; }
    }
}
