﻿// <auto-generated/>
// --------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="ExcelAssignment.cs" company="Chuck Hill">
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
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Style;

namespace CsvExcelExportImport
{
    /// <summary>
    /// Specialized import/export of 2-dimensional array of a formatted many-to-many assignment spreadsheet.
    /// </summary>
    public class ExcelAssignment
    {
        private static readonly Guid ExcelIdentifier = new Guid("3202CA33-8220-4D12-9F3D-CBC7CF7531DB");

        /// <summary>
        /// Initialize instance of ExcelAssignment object.
        /// </summary>
        public ExcelAssignment()
        {
        }

        /// <summary>
        /// Write a formatted many-to-many assignment spreadsheet to an open stream.
        /// </summary>
        /// <param name="stream">The open stream to write the excel workbook to</param>
        /// <param name="table">2-dimensional string table to write out. Header names must be pre-localized.</param>
        /// <param name="worksheetTabName">The localized worksheet tab name</param>
        /// <param name="wbProps">The workbook properties object to write to the excel workbook</param>
        /// <remarks>
        /// Table format:
        /// <code>
        ///   ┌──────┬──────┬──────┬──────┬──────┬──────┬──────┬──────┐
        ///   │ RHN1 │ RHN2 │ RHNn │ AHN1 │ AHN2 │ AHN2 │ AHN3 │ AHNn │
        ///   │ null │ null │ null │ key1 │ key2 │ key2 │ key3 │ keyN │
        ///   ├──────┼──────┼──────╆━━━━━━┿━━━━━━┿━━━━━━┿━━━━━━┿━━━━━━┥
        ///   │ RH1  │ RH2  │ RH3  ┃  X   │      │      │  X   │  X   │
        ///   │ RH1  │ RH2  │ RH3  ┃  X   │      │      │  X   │  X   │
        ///   │ RH1  │ RH2  │ RH3  ┃  X   │      │      │  X   │  X   │
        ///   │ ...  │ ...  │ ...  ┃ ...  │ ...  │ ...  │ ...  │ ...  │
        ///   └──────┴──────┴──────┸──────┴──────┴──────┴──────┴──────┘
        ///   where:
        ///     RHN  = row header column names
        ///     AHN  = assignment column names
        ///     null = null or empty as field has no meaning for row header columns
        ///            The last empty cell marks the beginning of the assignment columns.
        ///     key  = keys used upon import of each column (row is hidden in excel)
        ///     RH   = row header value
        ///     Assignment values:
        ///      'X' = row value is currently assigned in the database
        ///      null or empty = row value is not assigned
        /// </code>
        /// </remarks>
        public void Serialize(Stream stream, string[,] table, string worksheetTabName, WorkbookProperties wbProps)
        {
            if (table == null) throw new ArgumentNullException(nameof(table), "Source data must not be null.");
            Serialize(stream, ToEnumerable(table), worksheetTabName, wbProps);
        }

        /// <summary>
        ///   Write a formatted many-to-many assignment spreadsheet to an open stream.
        /// </summary>
        /// <param name="stream">The open stream to write the excel workbook to</param>
        /// <param name="table">
        ///   Enumerable array of string[]. Table rows are NOT random access. 
        ///   Forward read ONCE only. All string[] rows must be of the same 
        ///   length. Header names must be pre-localized. Specifically 
        ///   designed for end-to-end streaming.
        /// </param>
        /// <param name="worksheetTabName">The localized worksheet tab name</param>
        /// <param name="wbProps">The workbook properties object to write to the excel workbook</param>
        /// <remarks>
        /// Table format:
        /// <code>
        ///   ┌──────┬──────┬──────┬──────┬──────┬──────┬──────┬──────┐
        ///   │ RHN1 │ RHN2 │ RHNn │ AHN1 │ AHN2 │ AHN2 │ AHN3 │ AHNn │
        ///   │ null │ null │ null │ key1 │ key2 │ key2 │ key3 │ keyN │
        ///   ├──────┼──────┼──────╆━━━━━━┿━━━━━━┿━━━━━━┿━━━━━━┿━━━━━━┥
        ///   │ RH1  │ RH2  │ RH3  ┃  X   │      │      │  X   │  X   │
        ///   │ RH1  │ RH2  │ RH3  ┃  X   │      │      │  X   │  X   │
        ///   │ RH1  │ RH2  │ RH3  ┃  X   │      │      │  X   │  X   │
        ///   └──────┴──────┴──────┸──────┴──────┴──────┴──────┴──────┘
        ///   where:
        ///     RHN  = row header column names
        ///     AHN  = assignment column names
        ///     null = null or empty as field has no meaning for row header columns.
        ///            The last empty cell marks the beginning of the assignment columns.
        ///     key  = keys used upon import of each column (row is hidden in excel)
        ///     RH   = row header value
        ///     Assignment values:
        ///      'X' = row value is currently assigned in the database
        ///      null or empty = row value is not assigned
        /// </code>
        /// </remarks>
        public void Serialize(Stream stream, IEnumerable<string[]> table, string worksheetTabName, WorkbookProperties wbProps)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream), "Output stream must not be null.");
            if (table == null) throw new ArgumentNullException(nameof(table), "Source data must not be null.");
            if (string.IsNullOrWhiteSpace(worksheetTabName)) throw new ArgumentNullException(nameof(worksheetTabName), "The worksheet tab name must not be empty.");

            // We use the current thread UI culture for localization and restore it upon exit.
            CultureInfo originalUICulture = System.Threading.Thread.CurrentThread.CurrentUICulture;  // for language
            CultureInfo originalCulture = System.Threading.Thread.CurrentThread.CurrentCulture;      // for region

            try
            {
                using (var pkg = new ExcelPackage(stream))
                {
                    ExcelWorkbook wb = pkg.Workbook;
                    var xlprops = ExcelCommon.SetWorkbookProperties(wb, ExcelIdentifier, wbProps);
                    var ws = wb.Worksheets.Add(worksheetTabName);

                    int colCount = 99999;
                    int assignmentColIndex = 0;

                    // Set Header and Data values
                    // Table is enumerable array of string[]. Thus table rows are NOT random access. Forward read ONCE only.
                    int r = 0;
                    foreach (var row in table)
                    {
                        if (r == 0) // header row
                        {
                            colCount = row.Length;
                            for (int c = 0; c < colCount; c++)
                            {
                                ws.Cells[r + 1, c + 1].Value = row[c];
                            }

                            r++;
                            continue;
                        }

                        if (row.Length != colCount) throw new InvalidDataException("Column count mismatch.");

                        if (r == 1) // key row
                        {
                            for (int c = 0; c < colCount; c++)
                            {
                                if (string.IsNullOrWhiteSpace(row[c])) assignmentColIndex = c + 1; // find index of first assignment column.
                                ws.Cells[r + 1, c + 1].Value = row[c];
                            }

                            if (assignmentColIndex < 1) throw new ArgumentNullException(nameof(table), "There is no row header column.");
                            r++;
                            continue;
                        }

                        for (int c = 0; c < colCount; c++)
                        {
                            ws.Cells[r + 1, c + 1].Value = row[c].AppendSp();
                        }

                        r++;
                    }

                    var rowCount = r;
                    if (rowCount < 3) throw new ArgumentNullException(nameof(table), "Source data must not be empty."); // headerRow + keyRow + users count

                    // Add Checksum column
                    ws.Cells[1, colCount + 1].Value = "CheckSum";
                    for (r = 2; r < rowCount; r++)
                    {
                        var range = ws.Cells[r + 1, assignmentColIndex + 1, r + 1, colCount].Value as object[,];
                        ws.Cells[r + 1, colCount + 1].Value = EncodeChecksum(range);
                    }

                    ws.Cells.Style.Numberformat.Format = "@"; // All cells have the TEXT format.

                    // Hidden Key Row Formatting
                    using (var range = ws.Cells[2, 1, 2, colCount + 1])
                    {
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(xlprops.Light);
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        range.Style.TextRotation = 90;
                        ws.Row(2).Hidden = true; // Hide 2nd row. This contains the guid keys
                    }

                    // Visible header row formatting
                    using (var range = ws.Cells[1, 1, 1, colCount + 1])
                    {
                        range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        range.Style.Font.Bold = true;
                        range.Style.Fill.Gradient.Type = ExcelFillGradientType.Linear;
                        range.Style.Fill.Gradient.Degree = 90;
                        range.Style.Fill.Gradient.Color1.SetColor(xlprops.Medium); // TopGradientColor
                        range.Style.Fill.Gradient.Color2.SetColor(xlprops.Light);  // BottomGradientColor
                        range.Style.TextRotation = 90;
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Bottom;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }

                    // Reset visible row header, header formatting
                    using (var range = ws.Cells[1, 1, 1, assignmentColIndex])
                    {
                        range.Style.TextRotation = 0;
                        range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    }

                    ws.Protection.IsProtected = true;
                    ws.Protection.AllowSelectLockedCells = false;
                    using (var range = ws.Cells[3, assignmentColIndex + 1, rowCount, colCount])
                    {
                        range.Style.Locked = false;
                        range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        var val = range.DataValidation.AddListDataValidation();
                        val.ErrorStyle = ExcelDataValidationWarningStyle.stop;
                        val.AllowBlank = true;
                        val.ShowErrorMessage = true;
                        // val.ShowDropdown = false; // disable in-cell dropdown...Arrgh! Does't exist. See XML fixups below...
                        val.ErrorTitle = LocalizedStrings.GetString("AssignmentExcel_PopupErrorTitle", "Cell Assignment", wbProps.Culture);
                        val.Error = LocalizedStrings.GetString("AssignmentExcel_PopupErrorMessage", "Must enter 'X' to assign, or set to empty to unassign.", wbProps.Culture);
                        val.Formula.Values.Add(string.Empty);
                        val.Formula.Values.Add("X");
                        val.Formula.Values.Add("x");

                        var cf = range.ConditionalFormatting.AddEqual();
                        cf.Formula = "\"X\"";
                        cf.Style.Border.Right.Style = cf.Style.Border.Left.Style = cf.Style.Border.Top.Style = cf.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        cf.Style.Border.Right.Color.Color = cf.Style.Border.Left.Color.Color = cf.Style.Border.Top.Color.Color = cf.Style.Border.Bottom.Color.Color = Color.FromArgb(83, 141, 213);
                        cf.Style.Fill.PatternType = ExcelFillStyle.Solid; // ExcelFillStyle.Gradient does not exist! Too complicated to hack it with XML.
                        cf.Style.Fill.BackgroundColor.Color = Color.FromArgb(221, 231, 242);
                        cf.Style.Fill.PatternColor.Color = Color.FromArgb(150, 180, 216);
                        cf.Style.Font.Color.Color = Color.Brown;
                    }

                    ws.View.FreezePanes(3, assignmentColIndex + 1); // 2,4 refers to the first upper-left cell that is NOT frozen
                    ws.Column(colCount + 1).Hidden = true; // Hide last col. This contains the 'checksum' flags

                    ExcelCommon.SetPrintProperties(ws, wbProps.Culture);
                    ExcelCommon.DisableCellWarnings(ws);
                    ExcelCommon.HideCellValidationDropdowns(ws);
                    ExcelCommon.AutoFitColumns(ws, 3, false);

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
        ///   Read an open stream containing an excel workbook that was originally written by this matching serializer.
        /// </summary>
        /// <param name="stream">The open stream to read the excel workbook from</param>
        /// <param name="wbProps">The workbook properties object</param>
        /// <param name="disableChangeSync">
        ///   Disable change synchronization based upon export state 'checksum'.
        ///   Default = false. If true, the 2-dimensional table/matrix is always
        ///   returned as the modification state will be handled by the caller.<br />
        ///   Remarks:<br />
        ///   When false, this utility works fine if the user Exports-Modifies-
        ///   Imports just once. However if the intent is to create a backup and 
        ///   at some future date restore everything back to this state, then 
        ///   disableChangeSync must be set to True.<br />
        ///   Compared to the serialization example, the following is an example where
        ///   all AH1 are unassigned and AH2 are assigned when disableChangeSync==true.
        ///   <code>
        ///   ┌──────┬──────┬──────┬──────┬──────┬──────┬──────┬──────┐
        ///   │ RHN1 │ RHN2 │ RHNn │ AHN1 │ AHN2 │ AHN2 │ AHN3 │ AHNn │
        ///   │ null │ null │ null │ key1 │ key2 │ key2 │ key3 │ keyN │
        ///   ├──────┼──────┼──────╆━━━━━━┿━━━━━━┿━━━━━━┿━━━━━━┿━━━━━━┥
        ///   │ RH1  │ RH2  │ RH3  ┃      │      │  X   │  X   │  X   │
        ///   │ RH1  │ RH2  │ RH3  ┃      │      │  X   │  X   │  X   │
        ///   │ RH1  │ RH2  │ RH3  ┃      │      │  X   │  X   │  X   │
        ///   └──────┴──────┴──────┸──────┴──────┴──────┴──────┴──────┘
        ///   </code>
        /// </param>
        /// <returns>An updated 2-dimensional table/matrix that was read by the writer or null if workbook not modified.</returns>
        /// <remarks>
        /// Compared to the serialization example, the following is an example where all AH1 are
        /// unassigned and AH2 are assigned. Note that unchanged assignment values are cleared.<br />
        /// Table format when disableChangeSync==false (the default):
        /// <code>
        ///   ┌──────┬──────┬──────┬──────┬──────┬──────┬──────┬──────┐
        ///   │ RHN1 │ RHN2 │ RHNn │ AHN1 │ AHN2 │ AHN2 │ AHN3 │ AHNn │
        ///   │ null │ null │ null │ key1 │ key2 │ key2 │ key3 │ keyN │
        ///   ├──────┼──────┼──────╆━━━━━━┿━━━━━━┿━━━━━━┿━━━━━━┿━━━━━━┥
        ///   │ RH1  │ RH2  │ RH3  ┃  O   │      │  X   │      │      │
        ///   │ RH1  │ RH2  │ RH3  ┃  O   │      │  X   │      │      │
        ///   │ RH1  │ RH2  │ RH3  ┃  O   │      │  X   │      │      │
        ///   └──────┴──────┴──────┸──────┴──────┴──────┴──────┴──────┘
        ///   where:
        ///     RHN  = row header column names
        ///     AHN  = assignment column names
        ///     null = null or empty as field has no meaning for row header columns
        ///            The last empty cell marks the beginning of the assignment columns.
        ///     key  = keys used upon import of each column (row is hidden in excel)
        ///     RH   = row header value
        ///     Assignment values:
        ///      'X' = unassigned row value is to be assigned.
        ///      'O' = existing value is to be removed.
        ///      null or empty = row value is unchanged.
        /// </code>
        /// </remarks>
        public string[,] Deserialize(Stream stream, WorkbookProperties wbProps, bool disableChangeSync = false)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream), "Input stream must not be null.");

            using (var pkg = new ExcelPackage(stream))
            {
                var wb = pkg.Workbook;

                ExcelCommon.GetWorkbookProperties(wb, ExcelIdentifier, wbProps);
                var ws = pkg.Workbook.Worksheets.FirstOrDefault();

                // Not reliable.
                // var wsRowsLength = ws.Dimension.End.Row;
                // var wsColsLength = ws.Dimension.End.Column;

                var wsRowsLength = 0;
                for (; ws.Cells[wsRowsLength + 1, 1].Value != null; wsRowsLength++);
                var wsColsLength = 0;
                for (; ws.Cells[1, wsColsLength + 1].Value != null; wsColsLength++);

                int rowsLength = wsRowsLength;
                int colsLength = wsColsLength - 1; // exclude checksum in last column

                // Get number of rowheader columns by counting the sequential empty values in the key row (e.g. table[1,c])
                int assignmentColIndex = -1;
                for (int c = 0; c < colsLength; c++)
                {
                    if (!string.IsNullOrEmpty(ws.Cells[2, c + 1].Value as string))
                    {
                        assignmentColIndex = c;
                        break;
                    }
                }

                var table = new string[rowsLength, colsLength];

                bool modified = false;
                for (int r = 0; r < rowsLength; r++)
                {
                    for (int c = 0; c < colsLength; c++)
                    {
                        // Remove trailing space from numeric strings. Non-breaking space was added to keep Excel from automatically convert to a number.
                        table[r, c] = (ws.Cells[r + 1, c + 1].Value as string).TrimSp();
                    }

                    if (r > 1)
                    {
                        // Let caller handle any change synchronization.
                        if (disableChangeSync)
                        {
                            FixAssignments(table, r, assignmentColIndex);
                            modified = true;
                        }
                        else
                        {
                            if (UpdateAssignments(ws.Cells[r + 1, colsLength + 1].Value as string, table, r, assignmentColIndex)) modified = true;
                        }
                    }
                }

                if (!modified) return null;
                return table;
            }
        }

        #region -------------------- Private Methods --------------------
        /// <summary>
        /// Make sure assignment cells contain ONLY upper-case 'X' or null.
        /// </summary>
        /// <param name="table">2-D table containing the assignments</param>
        /// <param name="row">Row to update</param>
        /// <param name="startCol">Offset to the beginning of the assignment columns</param>
        private static void FixAssignments(string[,] table, int row, int startCol)
        {
            int colLength = table.GetLength(1);
            int assignmentsLength = colLength - startCol;

            for (int i = 0; i < assignmentsLength; i++)
            {
                if (table[row, i + startCol] == "x" || table[row, i + startCol] == "X")
                    table[row, i + startCol] = "X"; // force upper-case
                else table[row, i + startCol] = null;
            }
        }

        /// <summary>
        /// Use checksum to determine if previously assigned value is to be de-assigned OR if previously assigned value wasn't changed.
        /// </summary>
        /// <param name="checksum">Encoded base64 string containing bit flags of the original unmodified state.</param>
        /// <param name="table">2-D table containing the assignments</param>
        /// <param name="row">Row to update</param>
        /// <param name="startCol">Offset to the beginning of the assignment columns</param>
        /// <returns>True if any assignments modified</returns>
        private static bool UpdateAssignments(string checksum, string[,] table, int row, int startCol)
        {
            if (string.IsNullOrWhiteSpace(checksum)) return false;
            int colLength = table.GetLength(1);
            int assignmentsLength = colLength - startCol;
            bool modified = false;

            var cs = DecodeChecksum(checksum); // This can only contain true and null. never false.
            if (cs.Length != assignmentsLength) throw new TargetParameterCountException("Checksum length does not match the assignments length.");

            for (int i = 0; i < cs.Length; i++)
            {
                if (table[row, i + startCol] == "x")  // force upper-case
                    table[row, i + startCol] = "X";

                if (cs[i] == "X" && string.IsNullOrEmpty(table[row, i + startCol]))  // Remove
                    table[row, i + startCol] = "O";

                if (cs[i] == "X" && table[row, i + startCol] == "X")  // No Change
                    table[row, i + startCol] = null;

                if (table[row, i + startCol] == "X" || table[row, i + startCol] == "O")
                    modified = true;
            }

            return modified;
        }

        private static string EncodeChecksum(object[,] values)
        {
            var bitArray = new BitArray(values.GetLength(1));
            for (int i = 0; i < bitArray.Length; i++)
            {
                if ((values[0, i] as string) == "X") bitArray[i] = true;
            }

            var bytes = new byte[(bitArray.Length / 8) + (bitArray.Length % 8 > 0 ? 1 : 0)];
            bitArray.CopyTo(bytes, 0);

            return string.Concat("[", bitArray.Length, "]", Convert.ToBase64String(bytes));
        }

        private static string[] DecodeChecksum(string bitEncodedStr)
        {
            var lengthEndIndex = bitEncodedStr.IndexOf(']');
            var bitCount = Convert.ToInt16(bitEncodedStr.Substring(1, lengthEndIndex - 1));
            var uuStr = bitEncodedStr.Substring(lengthEndIndex + 1, bitEncodedStr.Length - (lengthEndIndex + 1));

            var bitArray = new BitArray(Convert.FromBase64String(uuStr));
            var assignments = new string[bitCount];

            for (var i = 0; i < bitCount; i++)
            {
                if (bitArray[i]) assignments[i] = "X";
            }

            return assignments;
        }

        private static IEnumerable<T[]> ToEnumerable<T>(T[,] table)
        {
            int rowCount = table.GetLength(0);
            int colCount = table.GetLength(1);

            for (int r = 0; r < rowCount; r++)
            {
                var row = new T[colCount];
                for (int c = 0; c < colCount; c++)
                {
                    row[c] = table[r, c];
                }

                yield return row;
            }

            yield break;
        }
        #endregion
    }
}
