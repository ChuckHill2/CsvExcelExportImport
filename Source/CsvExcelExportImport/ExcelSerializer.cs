//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="ExcelSerializer.cs" company="Chuck Hill">
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
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;

namespace CsvExcelExportImport
{
    /// <summary>
    /// Serializing an enumerable list of objects to a caller-defined Excel stream.
    /// Uses the current UI culture for formatting.
    /// </summary>
    public sealed class ExcelSerializer
    {
        private const int MaxRowsPerSheet = ExcelPackage.MaxRows;  // maximum records/sheet before rollover to a new sheet. This may be reduced for performance reasons.
        private static readonly Guid ExcelIdentifier = new Guid("5CC2DB65-6E9A-4690-A3C3-C98E59F79467");
        private Color TopGradientColor;     // Colors initialized in SetWorkbookProperties()
        private Color BottomGradientColor;
        private Color SolidFillColor;
        private Color DividerColor;
        private string TRUE;  // Boolean localized Yes/No
        private string FALSE;
        private WorkbookProperties.RegionInfoOverride CustomRegionInfo = null; // Initialized by SetWorkbookProperties()

        /// <summary>
        /// Initialize instance of ExcelSerializer
        /// </summary>
        public ExcelSerializer()
        {
        }

        /// <summary>
        /// Serialize multiple arrays of classes into a XLSX Excel workbook stream.
        /// </summary>
        /// <param name="stream">Open stream to write to. May be a file or in-memory stream.</param>
        /// <param name="lists">
        ///   Object array of enumerable sequences. One for each worksheet. If a sequence record is null, a divider line is written.
        ///   examples:
        ///   <code>
        ///     IList&lt;IEnumerable&gt; wkSheets = new IEnumerable[] { list0.OrderBy(), list1 };
        ///     IList&lt;IEnumerable&gt; wkSheets = new List&lt;IEnumerable&gt;() { list0.OrderBy(), list1 };
        ///   </code>
        /// </param>
        /// <param name="wbProps">Optional Workbook-level Excel properties to set. May be null if the defaults are OK.</param>
        public void Serialize(Stream stream, IEnumerable<IEnumerable> lists, WorkbookProperties wbProps)
        {
            // We use the current thread UI culture for localization and restore it upon exit.
            CultureInfo originalUICulture = System.Threading.Thread.CurrentThread.CurrentUICulture;  // for language
            CultureInfo originalCulture = System.Threading.Thread.CurrentThread.CurrentCulture;      // for region

            try
            {
                using (ExcelPackage pkg = new ExcelPackage(stream))
                {
                    ExcelWorkbook wb = pkg.Workbook;
                    var xlprops = ExcelCommon.SetWorkbookProperties(wb, ExcelIdentifier, wbProps);

                    this.TRUE = LocalizedStrings.GetString("True");
                    this.FALSE = LocalizedStrings.GetString("False");
                    this.CustomRegionInfo = xlprops.CustomRegionInfo;
                    this.TopGradientColor = xlprops.Medium;
                    this.BottomGradientColor = xlprops.Light;
                    this.SolidFillColor = xlprops.Light;
                    this.DividerColor = xlprops.Dark;
                    var worksheetHeadingJustification = wbProps?.WorksheetHeadingJustification ?? Justified.Auto;

                    int i = 0;
                    foreach (var list in lists)
                    {
                        AddWorkSheet(wb, new SerializerProperties(list, ExcelPropertyUpdater), wbProps?.WorksheetHeading?.ElementAtOrDefault(i), worksheetHeadingJustification);
                        i++;
                    }

                    pkg.Save();
                }
            }
            finally
            {
                System.Threading.Thread.CurrentThread.CurrentUICulture = originalUICulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = originalCulture;
            }
        }

        /// <summary>
        /// Read an open stream containing an excel workbook that was originally written by this matching serializer.
        /// This deserializes the stream into an array of classes. Note: The CultureInfo and 
        /// worksheet classes do not need to be specified as they are already encoded into the previously 
        /// serialized workbook.
        /// </summary>
        /// <param name="stream">Open stream to read from. May be a file or in-memory stream.</param>
        /// <param name="wbProps">Optional Workbook-level Excel properties to get. May be null if not needed.</param>
        /// <returns>An array of dissimilar arrays. One for each worksheet. ex: result[0].Cast&lt;DataModel&gt;().ToArray();</returns>
        public IList<IList> Deserialize(Stream stream, WorkbookProperties wbProps)
        {
            using (var pkg = new ExcelPackage(stream))
            {
                var wb = pkg.Workbook;
                if (wbProps == null) wbProps = new WorkbookProperties();
                ExcelCommon.GetWorkbookProperties(wb, ExcelIdentifier, wbProps);

                this.CustomRegionInfo = wbProps.CustomRegionInfo.Clone();
                this.TRUE = LocalizedStrings.GetString("True", wbProps.Culture);
                this.FALSE = LocalizedStrings.GetString("False", wbProps.Culture);

                var sheets = new List<IList>(wb.Worksheets.Count);
                foreach (ExcelWorksheet ws in wb.Worksheets)
                {
                    Type wsType = GetWorkSheetClassType(wbProps.ExtraProperties, ws.Name);
                    if (wsType == null) continue;
                    var sp = new SerializerProperties(wsType, ExcelPropertyUpdater, wbProps.Culture);

                    int headerRow = ws.Cells[1, 1].Merge == true ? 2 : 1;
                    wbProps.WorksheetHeading.Add(headerRow == 1 ? null : ws.Cells[1, 1].Value.ToString());
                    wbProps.WorksheetHeadingJustification = headerRow == 1 ? Justified.Auto : JustificationMap(ws.Cells[1, 1].Style.HorizontalAlignment);

                    var headers = new List<string>(ws.Dimension.Columns);
                    for (int i = 1; i <= ws.Dimension.Columns; i++)
                    {
                        headers.Add(ws.Cells[headerRow, i].Value?.ToString());
                    }

                    var properties = sp.ReOrderPropertiesByHeaders(headers);
                    if (properties.Count == 0) continue; // no column headers match the headers in sb.Properties!

                    var rows = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(wsType), ws.Dimension.Rows);

                    for (int r = headerRow + 1; r <= ws.Dimension.Rows; r++)
                    {
                        var obj = Activator.CreateInstance(wsType);
                        rows.Add(obj);
                        for (int c = 0; c < properties.Count; c++)
                        {
                            var p = properties[c];
                            p.SetValue(obj, ws.Cells[r, c + 1].Value);
                        }
                    }

                    sheets.Add(rows);
                }

                return sheets;
            }
        }

        #region Serialize Excel private methods
        /// <summary>
        /// Add new worksheet to Excel workbook
        /// </summary>
        /// <param name="wb">Workbook object to add new worksheet to.</param>
        /// <param name="sp">Serializer properties.</param>
        /// <param name="worksheetHeading">Value to add as a header in row 0, spanning the entire set of columns.</param>
        /// <param name="worksheetHeadingJustification">Justification of the worksheet heading</param>
        private void AddWorkSheet(ExcelWorkbook wb, SerializerProperties sp, string worksheetHeading, Justified worksheetHeadingJustification)
        {
            var wsName = GetUniqueWorksheetName(wb, sp.WorksheetTabName);
            var ws = wb.Worksheets.Add(wsName);
            SetWorkSheetClassType(wb, wsName, sp.EnumeratedType);
            ExcelRange range; // We use this a lot.

            int headerRow = string.IsNullOrWhiteSpace(worksheetHeading) ? 1 : 2;  // Excel indices start at 1

            // Add columns and styling
            for (int i = 0; i < sp.Properties.Count; i++)
            {
                var p = sp.Properties[i];

                ws.Cells[headerRow, i + 1].Value = p.Header;  // Add localized header name

                if (p.Style != null) // the custom PropertyUpdater() stores the excel style here
                {
                    ws.Column(i + 1).Style.Numberformat.Format = p.Style;
                }

                ws.Column(i + 1).Style.HorizontalAlignment = JustificationMap(p.Justification);
                ws.Column(i + 1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Column(i + 1).Hidden = p.Hidden;

                // Enums and booleans are restricted to a dropdown choice
                if (p.CellType.IsEnum)
                {
                    range = GetColumnRange(ws, i + 1);
                    range.Style.Locked = false;
                    var val = range.DataValidation.AddListDataValidation();
                    val.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                    val.AllowBlank = !p.IsNullable;
                    val.ShowErrorMessage = true;
                    val.ErrorTitle = LocalizedStrings.GetString("Enum.ErrorTitle", "An invalid value was entered");
                    val.Error = LocalizedStrings.GetString("Enum.ErrorMsg", "Select a value from the list");
                    // val.ShowDropdown = false; // disable in-cell dropdown doesn't exist. See HideCellValidationDropdowns()
                    if (p.IsNullable) val.Formula.Values.Add("\xA0");
                    var k = 0;
                    foreach (Enum e in Enum.GetValues(p.CellType))  // Supports max 255 items in dropdown
                    {
                        val.Formula.Values.Add(p.TranslateData ? SerializerProperties.LocalizedEnumName(e) : e.ToString());
                        k++;
                        if (k > 254) break;
                    }
                }
                else if (p.CellType == typeof(bool))
                {
                    range = GetColumnRange(ws, i + 1);
                    range.Style.Locked = false;
                    var val = range.DataValidation.AddListDataValidation();
                    val.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                    val.AllowBlank = !p.IsNullable;
                    val.ShowErrorMessage = true;
                    val.ErrorTitle = LocalizedStrings.GetString("Enum.ErrorTitle", "An invalid value was entered");
                    val.Error = LocalizedStrings.GetString("Enum.ErrorMsg", "Select a value from the list");
                    // val.ShowDropdown = false; // disable in-cell dropdown doesn't exist. See HideCellValidationDropdowns()
                    if (p.IsNullable) val.Formula.Values.Add("\xA0");
                    val.Formula.Values.Add(p.TranslateData ? this.TRUE : "True");
                    val.Formula.Values.Add(p.TranslateData ? this.FALSE : "False");
                }

                if (p.HasFilter)
                {
                    range = ws.Cells[headerRow, i + 1, ExcelPackage.MaxRows, i + 1];
                    range.AutoFilter = true;
                }

                if (p.MaxWidth > 0)
                {
                    ws.Column(i + 1).Style.WrapText = true;
                    ws.Column(i + 1).Width = p.MaxWidth;
                }
            }

            // If it exists, add/format worksheet heading
            if (!string.IsNullOrWhiteSpace(worksheetHeading))
            {
                range = ws.Cells[1, 1, 1, sp.Properties.Count];
                range.Merge = true;
                range.Style.Numberformat.Format = "@";
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                range.Style.Font.Bold = true;
                range.Style.Font.Color.SetColor(Color.Black);
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(SolidFillColor);
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                range.Style.WrapText = true;
                range.Style.HorizontalAlignment = JustificationMap(worksheetHeadingJustification);
                ws.Cells[1, 1].Value = worksheetHeading;
                ws.Row(1).Height = (worksheetHeading.Count(m => m == '\n') + 1.25) * ws.Row(headerRow).Height;
            }

            // Style the column headings.
            range = ws.Cells[headerRow, 1, headerRow, sp.Properties.Count];
            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Font.Bold = true;
            range.Style.Font.Color.SetColor(Color.Black);
            range.Style.Numberformat.Format = "@";
            range.Style.Fill.Gradient.Type = ExcelFillGradientType.Linear;
            range.Style.Fill.Gradient.Degree = 90;
            range.Style.Fill.Gradient.Color1.SetColor(TopGradientColor);
            range.Style.Fill.Gradient.Color2.SetColor(BottomGradientColor);

            // Add data.
            int r = headerRow;
            while (sp.ItemEnumerator.MoveNext())
            {
                if (sp.ItemEnumerator.Current == null)
                {
                    // Add row delimiter.
                    range = ws.Cells[r, 1, r, sp.Properties.Count];
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                    range.Style.Border.Bottom.Color.SetColor(DividerColor);
                    continue;
                }

                // Populate current row
                for (int c = 0; c < sp.Properties.Count; c++)
                {
                    var p = sp.Properties[c];
                    ws.Cells[r + 1, c + 1].Value = p.GetValue(sp.ItemEnumerator.Current);

                    // Update cell style.
                    if (p.RelUnitsIndex != 0)
                    {
                        var units = sp.Properties[c + p.RelUnitsIndex].GetValue(sp.ItemEnumerator.Current);
                        if (!string.IsNullOrWhiteSpace(units as string))
                            ws.Cells[r + 1, c + 1].Style.Numberformat.Format = GetNumberStyle(p.Format, (string)units);
                    }
                }

                r++;

                if (r >= MaxRowsPerSheet)
                {
                    // Continue in another worksheet.
                    AddWorkSheet(wb, sp, worksheetHeading, worksheetHeadingJustification);
                    break;
                }
            }

            ExcelCommon.SetPrintProperties(ws);

            var sRow = string.IsNullOrWhiteSpace(worksheetHeading) ? 2 : 3;
            var sCol = sp.Properties.FindLastIndex(m => m.Frozen) + 1;
            if (sCol == 0) sCol = 1;  // If frozen flag is on the first column.
            ws.View.FreezePanes(sRow, sCol); // (2,1) == first row/column NOT frozen! Weird logic.

            ExcelCommon.AutoFitColumns(ws, headerRow, true);

            ExcelCommon.DisableCellWarnings(ws);
        }

        /// <summary>
        /// Excel worksheet names cannot be larger than 31 chars and must be unique in the workbook.
        /// </summary>
        /// <param name="wb">Workbook object to search.</param>
        /// <param name="sugestedName">Suggested name for new worksheet.</param>
        /// <returns>Validated new worksheet name</returns>
        private static string GetUniqueWorksheetName(ExcelWorkbook wb, string sugestedName)
        {
            if (sugestedName.Length > 31) sugestedName = sugestedName.Substring(0, 31); 
            var wsName = sugestedName;
            int dupeNumber = 1;
            while (wb.Worksheets.FirstOrDefault(m => m.Name == wsName) != null)
            {
                var dn = dupeNumber.ToString();
                var extLen = dn.Length + 2;
                if (sugestedName.Length + extLen > 31) sugestedName = sugestedName.Substring(0, 31 - extLen);
                wsName = $"{sugestedName}({dn})";
                dupeNumber++;
            }
            return wsName;
        }

        /// <summary>
        /// Get column styling and adjust excel-specific getter/setter handlers for a given property/column.
        /// </summary>
        /// <param name="p">Property to fixup</param>
        private void ExcelPropertyUpdater(PropertyAttribute p)
        {
            string style = null;
            Func<object, object> oldGetValue = p.GetValue;    // Note: GetValue retrieves the value as-is as part of PropertyAttribute instantiation.
            Action<object, object> oldSetValue = p.SetValue;  // Note: SetValue already calls Cast.To(p.PropertyName, value) as part of PropertyAttribute instantiation. So we don't have to do it here again.

            switch (p.CellType.Name)
            {
                case "Char":
                case "String":
                    style = "@"; // "TEXT" format
                    break;

                case "Byte":
                case "SByte":
                case "Int16":
                case "Int32":
                case "Int64":
                case "UInt16":
                case "UInt32":
                case "UInt64":
                    style = GetNumberStyle(p.Format);
                    break;

                case "Boolean":
                    if (p.TranslateData) style = "@"; // "TEXT" format
                    break;

                case "DateTime":
                    style = GetDateTimeStyle(p.Format);
                    break;

                case "DateTimeOffset":
                    style = GetDateTimeStyle(p.Format);
                    p.GetValue = obj => ((DateTimeOffset?)oldGetValue(obj))?.CastTo<DateTime>();
                    break;

                case "TimeSpan":
                    style = GetTimeSpanStyle(p.Format);
                    break;

                case "Single":
                case "Double":
                case "Decimal":
                    style = GetNumberStyle(p.Format);
                    break;

                default:
                    if (p.CellType.IsEnum) style = "@"; // "TEXT" format
                    break;
            }

            p.Style = style;
        }

        private Dictionary<string, Lazy<string>> _dateFormats; // load on demand. Used exclusively by GetDateTimeStyle()

        /// <summary>
        /// Convert C# ToString format codes into their localized Excel equivalents.
        /// Invalid codes default to Excel auto-format.
        /// </summary>
        /// <param name="format">C#-style DateTime format codes and formats</param>
        /// <returns>Excel format for the specified C# format identifier or null if not found.</returns>
        private string GetDateTimeStyle(string format)
        {
            const string AltStyle = ";[Blue]@";  // Used when not a date.

            // Convert localized .NET format styles to equivalent Excel format styles.
            // It's all load-on-demand since not all formats are used.
            if (_dateFormats == null) 
            {
                var ci = CultureInfo.CurrentUICulture.Name == string.Empty ? CultureInfo.GetCultureInfo("en-US") : CultureInfo.CurrentUICulture;
                var dti = CultureInfo.CurrentCulture.DateTimeFormat;  // CurrentCulture for region

                // GeneralShortTimePattern and GeneralLongTimePattern are PRIVATE! Go Figure...
                var GeneralShortTimePattern = new Lazy<string>(() => (string)typeof(DateTimeFormatInfo).GetProperty("GeneralShortTimePattern", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(dti));
                var GeneralLongTimePattern = new Lazy<string>(() => (string)typeof(DateTimeFormatInfo).GetProperty("GeneralLongTimePattern", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(dti));

                // EPPlus spazzes during column autoformat over CultureInfo names so we use old legacy 
                // LCID's wherever possible. LCID==4096 means LCID is undefined for this culture.
                var lang = $"[$-{(ci.LCID == 4096 ? ci.Name : ci.LCID.ToString("X"))}]";

                // Compose our own localized format that includes millisecond precision.
                var longXDateTimePattern = new Lazy<string>(() =>
                {
                    try
                    {
                        var sb = new StringBuilder();
                        var ms = Regex.Matches(dti.ShortDatePattern.ToLowerInvariant(), $@"(?<S>[mdy]+)(?<D>{Regex.Escape(dti.DateSeparator)}|\s+)?");
                        foreach (Match m in ms)
                        {
                            var c = m.Groups["S"].Value[0];
                            sb.Append(new string(c, c == 'y' ? 4 : 2));
                            sb.Append(m.Groups["D"].Value);
                        }

                        sb.Append(" ");
                        ms = Regex.Matches(dti.LongTimePattern.ToLowerInvariant(), $@"(?<S>[hmst]+)(?<D>{Regex.Escape(dti.TimeSeparator)}|\s+)?");
                        var delimiter = string.Empty; 
                        foreach (Match m in ms)
                        {
                            sb.Append(delimiter);
                            var c = m.Groups["S"].Value[0];
                            if (c == 't') sb.Append("AM/PM");
                            else sb.Append(new string(c, 2));
                            if (c == 's') sb.Append(".000");
                            delimiter = m.Groups["D"].Value;
                        }

                        return sb.ToString();
                    }
                    catch (Exception)
                    {
                        return lang + GeneralLongTimePattern.Value.Replace("tt", "AM/PM");
                    }
                });
                var shortXDateTimePattern = new Lazy<string>(() =>
                {
                    try
                    {
                        var sb = new StringBuilder();
                        var ms = Regex.Matches(dti.ShortDatePattern.ToLowerInvariant(), $@"(?<S>[mdy]+)(?<D>{Regex.Escape(dti.DateSeparator)}|\s+)?");
                        foreach (Match m in ms)
                        {
                            var c = m.Groups["S"].Value[0];
                            sb.Append(new string(c, c == 'y' ? 4 : 2));
                            sb.Append(m.Groups["D"].Value);
                        }

                        sb.Append(" ");
                        ms = Regex.Matches(dti.ShortTimePattern.ToLowerInvariant(), $@"(?<S>[hmst]+)(?<D>{Regex.Escape(dti.TimeSeparator)}|\s+)?");
                        var delimiter = string.Empty;
                        foreach (Match m in ms)
                        {
                            sb.Append(delimiter);
                            var c = m.Groups["S"].Value[0];
                            if (c == 't') sb.Append("AM/PM");
                            else if (c != 's') sb.Append(new string(c, 2));
                            delimiter = m.Groups["D"].Value;
                        }

                        return sb.ToString();
                    }
                    catch (Exception)
                    {
                        return lang + GeneralShortTimePattern.Value.Replace("tt", "AM/PM");
                    }
                });

                // Datetime months, am/pm, and days of week strings are localized to
                // current culture. Non-dates are displayed in alternate color. 
                _dateFormats = new Dictionary<string, Lazy<string>>()
                {
                    // Date patterns may contain short or long month names
                    // Time patterns tt or AM/PM need to be localized.
                    { "d", new Lazy<string>(() => string.Concat(lang, CustomRegionInfo.ShortDatePattern.Replace("'", string.Empty), AltStyle)) },
                    { "D", new Lazy<string>(() => string.Concat(lang, CustomRegionInfo.LongDatePattern.Replace("'", string.Empty), AltStyle)) },
                    { "t", new Lazy<string>(() => string.Concat(lang, CustomRegionInfo.ShortTimePattern.Replace("'", string.Empty).Replace("tt", "AM/PM"), AltStyle)) },
                    { "T", new Lazy<string>(() => string.Concat(lang, CustomRegionInfo.LongTimePattern.Replace("'", string.Empty).Replace("tt", "AM/PM"), AltStyle)) },
                    { "g", new Lazy<string>(() => string.Concat(lang, CustomRegionInfo.ShortDatePattern.Replace("'", string.Empty), " ", CustomRegionInfo.ShortTimePattern.Replace("'", string.Empty).Replace("tt", "AM/PM"), AltStyle)) },
                    { "G", new Lazy<string>(() => string.Concat(lang, CustomRegionInfo.ShortDatePattern.Replace("'", string.Empty), " ", CustomRegionInfo.LongTimePattern.Replace("'", string.Empty).Replace("tt", "AM/PM"), AltStyle)) },
                    { "f", new Lazy<string>(() => string.Concat(lang, CustomRegionInfo.LongDatePattern.Replace("'", string.Empty), " ", CustomRegionInfo.ShortTimePattern.Replace("'", string.Empty).Replace("tt", "AM/PM"), AltStyle)) },
                    { "F", new Lazy<string>(() => string.Concat(lang, CustomRegionInfo.LongDatePattern.Replace("'", string.Empty), " ", CustomRegionInfo.LongTimePattern.Replace("'", string.Empty).Replace("tt", "AM/PM"), AltStyle)) },

                    { "m", new Lazy<string>(() => string.Concat(lang, dti.MonthDayPattern.Replace("'", string.Empty), AltStyle)) },
                    { "M", new Lazy<string>(() => string.Concat(lang, dti.MonthDayPattern.Replace("'", string.Empty), AltStyle)) },
                    { "o", new Lazy<string>(() => string.Concat("yyyy-mm-ddThh:mm:ss.000", AltStyle)) },
                    { "O", new Lazy<string>(() => string.Concat("yyyy-mm-ddThh:mm:ss.000", AltStyle)) },
                    { "r", new Lazy<string>(() => string.Concat(lang, dti.RFC1123Pattern.Replace("'", string.Empty).Replace("GMT", "\"GMT\""), AltStyle)) },
                    { "R", new Lazy<string>(() => string.Concat(lang, dti.RFC1123Pattern.Replace("'", string.Empty), AltStyle)) },
                    { "s", new Lazy<string>(() => string.Concat(lang, dti.SortableDateTimePattern.Replace("'", string.Empty), AltStyle)) },
                    // "S" is an illegal .NET format type so we don't support it either.
                    { "u", new Lazy<string>(() => string.Concat(lang, dti.UniversalSortableDateTimePattern.Replace("'", string.Empty), AltStyle)) },
                    { "U", new Lazy<string>(() => string.Concat(lang, dti.FullDateTimePattern.Replace("'", string.Empty).Replace("tt", "AM/PM"), AltStyle)) },
                    { "y", new Lazy<string>(() => string.Concat(lang, dti.YearMonthPattern.Replace("'", string.Empty), AltStyle)) },
                    { "Y", new Lazy<string>(() => string.Concat(lang, dti.YearMonthPattern.Replace("'", string.Empty), AltStyle)) },
                    // Our own custom localized formats composed from built-in ShortDatePattern and ShortTimePattern
                    { "x", new Lazy<string>(() => string.Concat(lang, shortXDateTimePattern.Value, AltStyle)) },
                    { "X", new Lazy<string>(() => string.Concat(lang, longXDateTimePattern.Value, AltStyle)) },
                };
            }

            if (format == null) return _dateFormats["g"].Value;

            if (_dateFormats.TryGetValue(format, out var style)) return style.Value;

            if (format.Any(c => c == 'y' || c == 'Y' || c == 'm' || c == 'M' || c == 'd' || c == 'D' || c == 'h' || c == 'H'))
            {
                var ci = CultureInfo.CurrentUICulture.Name == string.Empty ? CultureInfo.GetCultureInfo("en-US") : CultureInfo.CurrentUICulture;
                var lang = $"[$-{(ci.LCID == 4096 ? ci.Name : ci.LCID.ToString("X"))}]";
                return string.Concat(lang, format.Replace("tt", "AM/PM").Replace('f', '0'), AltStyle);
            }

            return _dateFormats["g"].Value;
        }

        /// <summary>
        /// Convert C# number styles to localized Excel format styles. The underlying data is not modified.
        /// Culture RegionInfo properties may be overridden by the custom regioninfo properties.
        /// </summary>
        /// <param name="format">
        /// Known format types are f, n, and c (case-insensitive).
        /// Zero or more digits that follow represent the precision count after the decimal.
        /// A second optional comma-delimited field represents the units string. Leading and trailing whitespace is included.
        /// example: format="f3, mg" will be displayed as "5.320 mg",
        /// When precision is undefined, the precision==2. To To force zero digits after the decimal, use "f0".
        /// </param>
        /// <param name="dynamicUnits">Dynamic value to use for units</param>
        /// <returns>Excel style or null if unformatted.</returns>
        private string GetNumberStyle(string format, string dynamicUnits = null)
        {
            if (string.IsNullOrWhiteSpace(format)) return null;

            CultureInfo ci = CultureInfo.CurrentUICulture.Name == string.Empty ? CultureInfo.GetCultureInfo("en-US") : CultureInfo.CurrentUICulture;

            // var nf = ci.NumberFormat;
            var nf = CustomRegionInfo;

            if (format[0] == 'f' || format[0] == 'F')
            {
                format = format.Substring(1);
                var e = format.Split(new[] { ',' });
                int i = 0;
                string units = dynamicUnits ?? string.Empty;
                int decimalDigits = 2;
                if (e.Length > i && int.TryParse(e[i++], out var digits)) decimalDigits = digits;
                if (e.Length > i) units = e[i];
                if (decimalDigits < 0) decimalDigits = 0;
                units = string.IsNullOrEmpty(units) ? string.Empty : string.Concat("\"", units, "\"");

                var precision = decimalDigits == 0 ? string.Empty : nf.NumberDecimalSeparator + new string('0', decimalDigits);

                return string.Concat("0", precision, units, ";[Blue]@");
            }

            if (format[0] == 'n' || format[0] == 'N')
            {
                format = format.Substring(1);
                var e = format.Split(new[] { ',' });
                int i = 0;
                string units = dynamicUnits ?? string.Empty;
                int decimalDigits = 2;
                if (e.Length > i && int.TryParse(e[i++], out var digits)) decimalDigits = digits;
                if (e.Length > i) units = e[i];
                if (decimalDigits < 0) decimalDigits = 0;
                units = string.IsNullOrEmpty(units) ? string.Empty : string.Concat("\"", units, "\"");

                string[] negativePattern =
                {
                    "({0})",
                    "-{0}",
                    "- {0}",
                    "{0}-",
                    "{0} -",
                };

                string[] positivePattern =
                {
                    "{0}_)",
                    "{0}",
                    "{0}",
                    "{0}_)",
                    "{0}__)",
                };

                var precision = decimalDigits == 0 ? string.Empty : nf.NumberDecimalSeparator + new string('0', decimalDigits);
                var num = $"#{nf.NumberGroupSeparator}##0{precision}";

                return string.Concat(
                    string.Format(positivePattern[nf.NumberNegativePattern], num), 
                    units, 
                    ";", 
                    string.Format(negativePattern[nf.NumberNegativePattern], num), 
                    units);
            }

            if (format[0] == 'c' || format[0] == 'C')
            {
                format = format.Substring(1);
                var e = format.Split(new[] { ',' });
                int i = 0;
                string units = dynamicUnits ?? string.Empty;
                int decimalDigits = nf.CurrencyDecimalDigits;
                if (e.Length > i && int.TryParse(e[i++], out var digits)) decimalDigits = digits;
                if (e.Length > i) units = e[i];
                if (decimalDigits < 0) decimalDigits = 0;
                units = string.IsNullOrEmpty(units) ? string.Empty : string.Concat("\"", units, "\"");

                string[] positivePattern =
                {
                    "{0}{1}",
                    "{1}{0}",
                    "{0} {1}",
                    "{1} {0}",
                };

                string[] negativePattern =
                {
                    "({0}{1})",
                    "-{0}{1}",
                    "{0}-{1}",
                    "{0}{1}-",
                    "({0}{1})",
                    "-{1}{0}",
                    "{1}-{0}",
                    "{1}{0}-",
                    "-{1} {0}",
                    "-{0} {1}",
                    "{1} {0}-",
                    "{0} {1}-",
                    "{0} -{1}",
                    "{1}- {0}",
                    "({0} {1})",
                    "({1} {0})"
                };

                var precision = decimalDigits == 0 ? string.Empty : nf.CurrencyDecimalSeparator + new string('0', decimalDigits);
                var num = $"#{nf.CurrencyGroupSeparator}##0{precision}";

                // EPPlus spazzes during column autoformat over CultureInfo names so we use old legacy 
                // LCID's wherever possible. LCID==4096 means LCID is undefined for this culture.
                var sym = $"[${nf.CurrencySymbol}-{(ci.LCID == 4096 ? ci.Name : ci.LCID.ToString("X"))}]";

                return string.Concat(
                    string.Format(positivePattern[nf.CurrencyPositivePattern], sym, num), 
                    units,
                    ";[Red]",
                    string.Format(negativePattern[nf.CurrencyNegativePattern], sym, num), 
                    units);
            }

            return null;
        }

        /// <summary>
        /// Get Excel TimeSpan style
        /// </summary>
        /// <param name="format">Valid formats are:
        ///   <code>
        ///   g = low precision e.g. d.hh:mm (default)
        ///   G = high precision e.g. d.hh:mm.ss.fff
        ///   </code>
        /// </param>
        /// <returns>Excel format for the specified format identifier.</returns>
        private static string GetTimeSpanStyle(string format)
        {
            // Closest built-in styles
            // 21=="h:mm:ss", 45=="mm:ss", 46=="[h]:mm:ss", 47=="mm:ss.0"

            if (format == "G") return "[<=1]h:mm:ss;d.hh:mm:ss.000;[Blue]@";
            return "[<=1]h:mm:ss;d.hh:mm;[Blue]@";
        }

        /// <summary>
        /// Select range of entire columns.
        /// </summary>
        /// <param name="ws">Worksheet to create range from</param>
        /// <param name="startCol">Starting 1-based index in column range</param>
        /// <param name="endCol">Optional ending 1-based index in column range. If undefined then equals starting index.</param>
        /// <returns>Excel range object.</returns>
        private static ExcelRange GetColumnRange(ExcelWorksheet ws, int startCol, int endCol = 0)
        {
            if (endCol == 0) endCol = startCol;
            return ws.Cells[string.Concat(ExcelCellAddress.GetColumnLetter(startCol), ":", ExcelCellAddress.GetColumnLetter(endCol))];
        }
        #endregion

        #region Matched Getters/Setters
        private static void SetWorkSheetClassType(ExcelWorkbook wb, string worksheetName, Type t)
        {
            var s = t.CastTo<string>();
            if (s.Length < 255) wb.Properties.SetCustomPropertyValue(worksheetName, s);
            else
            {
                // Just in case the full type name exceeds the Excel custom property size.
                var s0 = s;
                int i = 0;
                while (s0.Length > 255)
                {
                    var maxlen = s0.Length > 255 ? 255 : s0.Length;
                    var value = s0.Substring(0, maxlen);
                    s0 = s0.Substring(maxlen);
                    wb.Properties.SetCustomPropertyValue($"{worksheetName}+{i++}", value);
                }

                if (s0.Length > 0)
                    wb.Properties.SetCustomPropertyValue($"{worksheetName}+{i++}", s0);
            }
        }

        private static Type GetWorkSheetClassType(Dictionary<string, string> extraProperties, string worksheetName)
        {
            if (extraProperties.TryGetValue(worksheetName, out var typeName))
            {
                return Type.GetType(typeName, false);
            }

            // Need to reassemble full assembly-qualified type name from truncated pieces.
            var sb = new StringBuilder();
            int i = 0;
            var key = $"{worksheetName}+{i++}";
            while (extraProperties.TryGetValue(key, out typeName))
            {
                extraProperties.Remove(key);
                sb.Append(typeName);
                key = $"{worksheetName}+{i++}";
            }

            if (sb.Length == 0) return null;

            extraProperties.Add(worksheetName, sb.ToString());
            return Type.GetType(sb.ToString(), false);
        }

        private static ExcelHorizontalAlignment JustificationMap(Justified justification)
        {
            switch (justification)
            {
                case Justified.Auto: return ExcelHorizontalAlignment.General;
                case Justified.Left: return ExcelHorizontalAlignment.Left;
                case Justified.Center: return ExcelHorizontalAlignment.Center;
                case Justified.Right: return ExcelHorizontalAlignment.Right;
                default: return ExcelHorizontalAlignment.General;
            }
        }

        private static Justified JustificationMap(ExcelHorizontalAlignment justification)
        {
            switch (justification)
            {
                case ExcelHorizontalAlignment.General: return Justified.Auto;
                case ExcelHorizontalAlignment.Left: return Justified.Left;
                case ExcelHorizontalAlignment.Center: return Justified.Center;
                case ExcelHorizontalAlignment.Right: return Justified.Right;
                default: return Justified.Auto;
            }
        }
        #endregion
    }
}