//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="CsvSerializer.cs" company="Chuck Hill">
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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace CsvExcelExportImport
{
    /// <summary>
    /// Serialize/Deserialize enumerable objects to/from CSV format. Note: Values are NOT localized
    /// into any specified culture(though headings are), so CSV document may be seamlessly read into
    /// Excel as a native Excel values and not just as plain text (or worse). The data is enumerable
    /// to allow late conversion upon read and create the smallest memory footprint. The resulting CSV
    /// document is also serialized into the smallest file size and still be readable is standard CSV.
    /// NOTE: When a JSON string[,] is instead serialized as CSV it is at a minimum 20% smaller than JSON!
    /// </summary>
    public sealed class CsvSerializer
    {
        private readonly CultureInfo Culture;

        /// <summary>
        /// Initializes a new instance of the CsvSerializer class that can serialize objects 
        /// of the specified type into CSV documents, and deserialize CSV documents back into 
        /// objects of the specified type. This is extremely lightweight and fast with a very
        /// small memory footprint.
        /// </summary>
        /// <param name="ci">
        ///   Culture to use when serializing/deserializing headers. If undefined, uses current UI culture.
        ///   Invariant culture is used for the actual data in order to load properly into Excel.
        /// </param>
        public CsvSerializer(CultureInfo ci = null)
        {
            Culture = ci ?? System.Threading.Thread.CurrentThread.CurrentUICulture;
        }

        /// <summary>
        /// Serializes a single enumerable object into a single CSV document into the 
        /// specified Stream. There are NO limits to the number of records written to the 
        /// stream. There is no caching or buffering. Individual values are formatted and 
        /// written to the stream immediately.
        /// </summary>
        /// <param name="textwriter">
        ///   Open stream to write to. Note: stream is not closed and stream pointer is not
        ///   reset to beginning in order to potentially perform further processing.
        /// </param>
        /// <param name="items">Enumerable list of items to write.</param>
        public void Serialize(TextWriter textwriter, IEnumerable items)
        {
            var sp = new SerializerProperties(items, null, Culture);

            using (var writer = new CsvWriter(textwriter))
            {
                // Write header record
                foreach (var p in sp.Properties)
                {
                    writer.WriteField(p.Header);
                }

                writer.WriteEOL();

                // Write records
                while (sp.ItemEnumerator.MoveNext())
                {
                    foreach (var p in sp.Properties)
                    {
                        writer.WriteField(p.GetValue(sp.ItemEnumerator.Current));
                    }

                    writer.WriteEOL();
                }
            }
        }

        /// <summary>
        /// Serializes multiple enumerable objects into multiple CSV documents into the 
        /// specified Stream. There are NO limits to the number of records written to the 
        /// stream. There is no caching or buffering. Individual values are formatted and 
        /// written to the stream immediately.
        /// </summary>
        /// <param name="textWriter">
        ///   Open stream to write to. Note: stream is not closed and stream pointer is not
        ///   reset to beginning in order to potentially perform further processing.
        /// </param>
        /// <param name="multiItems">Enumerable list of items to write.</param>
        public void Serialize(TextWriter textWriter, IEnumerable<IEnumerable> multiItems)
        {
            bool isFirst = true;
            foreach (var items in multiItems)
            {
                if (!isFirst) textWriter.Write(CsvReader.EndOfCSV.ToString() + Environment.NewLine);
                Serialize(textWriter, items);
                isFirst = false;
            }
        }

        /// <summary>
        /// Deserialize a single page in a CSV document contained in the specified Stream.
        /// </summary>
        /// <typeparam name="T">Enumerable class to serialize into CSV</typeparam>
        /// <param name="textReader">
        ///   The Stream that contains the CSV document to deserialize. Warning: Do not close
        ///   the stream until AFTER the output has been entirely evaluated/used/enumerated.
        /// </param>
        /// <returns>
        ///   Returns an Enumerable list of objects of type T.  The enumerable list can only
        ///   be enumerated just ONCE AND must be enumerated BEFORE the stream is closed.
        ///   Further re-enumerating will return ZERO items. If one wishes to actually save
        ///   these values for re-use elsewhere then the enumerable object must be converted
        ///   to an array or List. e.g. enumerableList.ToArray() or enumerableList.ToList().
        /// </returns>
        public IEnumerable<T> Deserialize<T>(TextReader textReader) where T : class, new()
        {
            return Deserialize(textReader, typeof(T)).Cast<T>();
        }

        /// <summary>
        /// Deserializes multiple pages in a CSV document contained in the specified Stream.
        /// </summary>
        /// <param name="textReader">
        ///   The Stream that contains one or more CSV pages to deserialize. Warning: Do not close
        ///   the stream until AFTER the output has been entirely evaluated/used/enumerated.
        /// </param>
        /// <param name="types">Array of class types for each CSV document in the stream.</param>
        /// <returns>
        ///   Returns a sequential list of Enumerable objects for each specified class type.  The 
        ///   enumerable lists can only be enumerated just ONCE AND must be enumerated BEFORE the 
        ///   stream is closed. Further re-enumerating will return ZERO items. If one wishes to 
        ///   actually save these values for re-use elsewhere then the enumerable object must be 
        ///   converted to an array or List. e.g. enumerableList.ToArray() or enumerableList.ToList().
        /// </returns>
        public IList<IEnumerable> Deserialize(TextReader textReader, IList<Type> types)
        {
            var list = new IEnumerable[types.Count];
            for (int i = 0; i < types.Count; i++)
            {
                list[i] = Deserialize(textReader, types[i]);
            }

            return list;
        }

        /// <summary>
        /// Get enumerable text objects beginning at the current position in the stream.
        /// </summary>
        /// <param name="textReader">Stream to read</param>
        /// <param name="t">Type of class objects to read into.</param>
        /// <returns>
        ///   Enumerable list of class object. DO NOT close stream until after the
        ///   enumerable list has been evaluated.
        /// </returns>
        private IEnumerable Deserialize(TextReader textReader, Type t)
        {
            var sp = new SerializerProperties(t, null, Culture);
            using (var reader = new CsvReader(textReader))
            {
                var list = (IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(t));
                var headers = reader.ReadRecord();
                var properties = sp.ReOrderPropertiesByHeaders(headers);
                while (!reader.EndOfFile)
                {
                    var index = 0;
                    var nuClass = Activator.CreateInstance(t);
                    foreach (var field in reader.ReadField())
                    {
                        var pa = properties[index];
                        // CSV cannot detect difference between "" and null, so we opt for null to support nullable types.
                        pa.SetValue(nuClass, field.Length == 0 ? null : field);
                        index++;
                    }

                    if (index > 0)
                    {
                        yield return nuClass;
                    }
                }

                yield break;
            }
        }

        /// <summary>
        /// Detect if this is the first time anything has been written to the TextWriter stream
        /// </summary>
        /// <param name="tw">TextWriter stream to check.</param>
        /// <returns>True if TestWriter is at the beginning of the stream</returns>
        private bool IsBeginningOfStream(TextWriter tw)
        {
            Type t = tw.GetType();
            int? pos;

            var fullName = t.FullName;
            if (fullName.Equals("System.CodeDom.Compiler.IndentedTextWriter") ||
                fullName.Equals("System.Web.HttpWriter") ||
                fullName.Equals("System.Web.UI.HtmlTextWriter"))
                throw new NotSupportedException($"TextWriter {fullName} is not supported.");

            // The TextWriter stream position is not entirely accurate but it is good enough for detecting if the stream has not been written to yet.
            pos = t.GetField("charPos", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.GetValue(tw) as int?;
            if (pos == null)
                pos = (t.GetField("_sb", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.GetValue(tw) as StringBuilder)?.Length;
            else pos = 0;

            return pos.Value == 0;
        }
    }
}