﻿//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="CsvWriter.cs" company="Chuck Hill">
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
// --------------------------------------------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

namespace CsvExcelExportImport
{
    /// <summary>
    ///     Robust CSV writer that can write to a file, or any open stream.
    /// </summary>
    internal sealed class CsvWriter : IDisposable
    {
        #region -------------------- Constants and Fields --------------------
        private static readonly CultureInfo InvariantCulture = CultureInfo.InvariantCulture;
        private readonly string numberDecimalSeparator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        private readonly StringBuilder fieldBuilder = new StringBuilder();
        private TextWriter writer;
        private bool itemIsNull;
        private bool closeOnDispose = false;
        #endregion

        #region -------------------- Constructors and Destructors --------------------
        /// <summary>
        /// Initializes a new instance of the <see cref="CsvWriter"/> class.
        /// Because the offset pointer is maintained, the open stream object may continue to be appended to.
        /// </summary>
        /// <param name="stream">The open TextWriter stream to sequentially write text to.</param>
        public CsvWriter(TextWriter stream)
        {
            RecordCount = 0;
            FieldCount = 0;
            FieldIndex = 0;

            this.writer = stream;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CsvWriter"/> class.
        /// Further data cannot be appended to this stream.
        /// </summary>
        /// <param name="stream">The open binary stream to write text to.</param>
        public CsvWriter(Stream stream) : this(new StreamWriter(stream, Encoding.UTF8, 4096, true))
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CsvWriter"/> class for the specified file.
        /// </summary>
        /// <param name="filename">The full new or existing file path to write to.</param>
        public CsvWriter(string filename)
            : this(File.Open(filename, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
        {
            closeOnDispose = true;
        }
        #endregion

        #region -------------------- Public Properties --------------------
        /// <summary>
        /// Current column written, so far.
        /// </summary>
        public int FieldIndex { get; private set; }

        /// <summary>
        /// The maximum number fields in any record, so far.
        /// </summary>
        public int FieldCount { get; private set; }

        /// <summary>
        /// The number of lines written, so far.
        /// </summary>
        public int RecordCount { get; private set; }

        /// <summary>
        /// Check if this is closed for writing.
        /// </summary>
        public bool Disposed
        {
            get
            {
                return this.writer == null;
            }
        }
        #endregion

        #region -------------------- Public Methods --------------------
        /// <summary>
        /// Converts an enumerable list of objects into a single CSV record
        ///     that is properly quoted and escaped. Do not use this for writing
        ///     multiple records. Open a CsvWriter instance instead.
        /// </summary>
        /// <param name="list">Enumerable list of objects to convert</param>
        /// <returns>Formatted CSV string</returns>
        public static string Join(IEnumerable list)
        {
            using (var ms = new MemoryStream())
            {
                using (var csv = new CsvWriter(ms))
                {
                    foreach (object o in list)
                    {
                        csv.WriteField(o);
                    }

                    ms.Position = 0;
                    var sr = new StreamReader(ms);
                    return sr.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Write single field item to CSV. Field item is allowed to contain newlines.
        /// All data is formatted using the invariant culture, so Excel can import with no issues.
        /// </summary>
        /// <param name="item">Item to write</param>
        public void WriteField(object item)
        {
            if (this.writer == null)
            {
                return;
            }

            if (FieldIndex > 0)
            {
                this.writer.Write(',');
            }

            FieldIndex++;

            if (item == null || (item is string && string.IsNullOrWhiteSpace((string)item)))
            {
                itemIsNull = true;
                return;
            }

            itemIsNull = false;

            if (item is Guid || item is sbyte || item is byte || item is short || item is ushort || item is int || item is uint || item is long || item is ulong || item is double || item is float)
            {
                this.writer.Write(((IFormattable)item).ToString(null, InvariantCulture)); // doesn't need floating-pt trailing zero truncation or illegal char quoting
            }
            else if (item is bool)
            {
                this.writer.Write(((bool)item).ToString(InvariantCulture));
            }
            else if (item is decimal)
            {
                this.writer.Write(((decimal)item).ToString("0.#################################", InvariantCulture));  // 'Normalize', strip trailing zeros
            }
            else if (item is DateTime)
            {
                // Hardcode to convert into invariant (aka SortableDateTime) datetime string format so this can
                // be parsed by Excel from within ANY culture. For brevity, we don't include any trailing zero fields.
                var dt = (DateTime)item;
                this.writer.Write(dt.ToString(GetDateTimeFormat(dt), InvariantCulture));
            }
            else if (item is DateTimeOffset)
            {
                var dt = ((DateTimeOffset)item).CastTo<DateTime>();  // Excel does not understand DateTimeOffset.
                this.writer.Write(dt.ToString(GetDateTimeFormat(dt), InvariantCulture));
            }
            else if (item is TimeSpan)
            {
                // this.writer.Write(((TimeSpan)item).ToString(null, InvariantCulture)); // default. Looks nice in CSV but Excel converts incorrectly.
                // this.writer.Write(string.Format("1900-01-{0:00} {1:00}:{2:00}:{3:00}", ts.Days + 1, ts.Hours, ts.Minutes, ts.Seconds)); // doesn't work! too many days! ds.Days (aka day of month) cannot be zero!
                // this.writer.Write(((TimeSpan)item).TotalDays.ToString(null, InvariantCulture)); // Converts well to Excel format where double=days.fractionalday but requires excel style="[<=1]h:mm:ss;d.hh:mm:ss". Native excel sorting and filtering works well!

                // The following is excel-unconvertable and will be maintained as a string but still may be string sortable. NO filtering!
                TimeSpan ts = (TimeSpan)item;
                var negative = string.Empty;
                if (ts.Ticks < 0)
                {
                    negative = "-";
                    ts = new TimeSpan(-ts.Ticks);
                }
                this.writer.Write(string.Format(GetTimeSpanFormat(ts), negative, ts.Days, ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds));
            }
            else
            {
                this.writer.Write(MaybeQuoteCsvField(item.ToString().TrimSp()));
            }
        }

        /// <summary>
        ///     Write EOL to CSV (aka "\r\n")
        /// </summary>
        public void WriteEOL()
        {
            if (this.writer == null)
            {
                return;
            }

            // We must quote the last zero-length item because IF this is the last item in a record, we cannot have a 
            // trailing comma because any CSV reader will think that the current record is continuing on the next line!
            if (itemIsNull)
            {
                this.writer.Write("\"\"");
                itemIsNull = false;
            }

            this.writer.Write(Environment.NewLine);
            RecordCount++;
            if (FieldIndex > FieldCount)
            {
                FieldCount = FieldIndex;
            }

            FieldIndex = 0;
        }

        /// <summary>
        ///     Dispose CsvWriter.
        ///     CsvWriter properties are still available.
        /// </summary>
        public void Dispose()
        {
            if (this.writer == null)
            {
                return;
            }

            // We must quote the last zero-length item because IF this is the last item in a record, we cannot have a 
            // trailing comma because any CSV reader will think that the current record is continuing on the next line!
            if (itemIsNull)
            {
                this.writer.Write("\"\"");
                itemIsNull = false;
            }

            RecordCount++;
            if (FieldIndex > FieldCount)
            {
                FieldCount = FieldIndex;
            }

            if (closeOnDispose)
            {
                this.writer.Dispose();
            }

            this.writer = null;
        }
        #endregion

        #region -------------------- Private Methods --------------------
        /// <summary>
        /// Trim leading and trailing whitespace.
        ///     Quote fields that contain ',' and '\n'.
        ///     Escape embedded quote char '"'.
        ///     Strip '\r' chars in multi-line fields so older versions of Excel can read them properly.
        ///     Trim trailing zeros from floating-point numbers converted to strings.
        /// </summary>
        /// <param name="field">String field to format.</param>
        /// <returns>Formatted CSV field</returns>
        private string MaybeQuoteCsvField(string field)
        {
            if (field.Length == 0)
            {
                return field;
            }

            // If it is a string-ized floating point number, trim the trailing zeros
            int index1 = field.LastIndexOf(numberDecimalSeparator, StringComparison.InvariantCulture);
            int index2 = field.LastIndexOf(".", StringComparison.InvariantCulture);

            if ((index1 != -1 || index2 != -1) && decimal.TryParse(field, out decimal value))
            {
                return value.ToString("0.#################################", InvariantCulture);
            }

            this.fieldBuilder.Length = 0;
            bool needsQuote = false;
            foreach (char c in field)
            {
                if (this.fieldBuilder.Length == 0 && c == ' ')
                {
                    continue;
                }

                if (c == '\r')
                {
                    continue;
                }

                if (c == ',')
                {
                    needsQuote = true;
                }

                if (c == '\n')
                {
                    needsQuote = true;
                }

                if (c == '"')
                {
                    this.fieldBuilder.Append('"');
                }

                this.fieldBuilder.Append(c);
            }

            while (this.fieldBuilder.Length > 0 && this.fieldBuilder[this.fieldBuilder.Length - 1] == ' ')
            {
                this.fieldBuilder.Length--;
            }

            if (needsQuote)
            {
                this.fieldBuilder.Insert(0, '"');
                this.fieldBuilder.Append('"');
            }

            return this.fieldBuilder.ToString();
        }

        private string GetTimeSpanFormat(TimeSpan ts)
        {
            if (ts.Seconds == 0 && ts.Milliseconds == 0) return "({0}{1:00}.{2:00}:{3:00})";
            if (ts.Milliseconds == 0) return "({0}{1:00}.{2:00}:{3:00}:{4:00})";
            return "({0}{1:00}.{2:00}:{3:00}:{4:00}.{5:000})";
        }

        private string GetDateTimeFormat(DateTime dt)
        {
            // Can't use the 'T' between date and time because Excel does not know how to convert it to a 
            // proper datetime object. Excel also does not understand '.fff' in a DateTime. It treats it as a TimeSpan!
            if (dt.Hour == 0 && dt.Minute == 0 && dt.Second == 0 && dt.Millisecond == 0) return "yyyy'-'MM'-'dd";
            if (dt.Second == 0 && dt.Millisecond == 0) return "yyyy'-'MM'-'dd' 'HH':'mm";
            if (dt.Millisecond == 0) return "yyyy'-'MM'-'dd' 'HH':'mm':'ss";
            if (dt.Millisecond % 100 == 0) return "yyyy'-'MM'-'dd' 'HH':'mm':'ss.f";
            if (dt.Millisecond % 10 == 0) return "yyyy'-'MM'-'dd' 'HH':'mm':'ss.ff";
            return "yyyy'-'MM'-'dd' 'HH':'mm':'ss.fff";
        }
        #endregion
    }
}