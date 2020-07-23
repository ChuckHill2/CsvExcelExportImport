//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="SerializerProperties.cs" company="Chuck Hill">
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
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

namespace CsvExcelExportImport
{
    /// <summary>
    /// Base class/common functionality for all Serializer classes
    /// </summary>
    internal class SerializerProperties
    {
        /// <summary>
        /// Get serialization/deserialization properties for the specified type. This is a common class for any serializer.
        /// </summary>
        /// <param name="enumeratedType">The class type to deserialize into.</param>
        /// <param name="propertyUpdater">Optional method to custom modify the properties as required for the requisite serializer/deserializer.</param>
        /// <param name="ci">Culture to use for header translation</param>
        public SerializerProperties(Type enumeratedType, Action<PropertyAttribute> propertyUpdater = null, CultureInfo ci = null)
        {
            if (enumeratedType == null) throw new ArgumentNullException("Enumerated type is unknown.");

            ci = ci ?? System.Threading.Thread.CurrentThread.CurrentUICulture;

            if (!enumeratedType.IsClass || enumeratedType.GetConstructor(Type.EmptyTypes) == null)
                throw new NotSupportedException($"Type {enumeratedType.FullName} not supported. Must be a class and have a parameterless constructor.");

            this.ItemEnumerator = null;
            this.EnumeratedType = enumeratedType;
            this.WorksheetTabName = GetWorkSheetTabName(this.EnumeratedType, ci);
            this.Properties = PropertyAttribute.GetProperties(this.EnumeratedType, ci);

            if (this.Properties.Count == 0) 
                throw new ArgumentOutOfRangeException($"Type {this.EnumeratedType.FullName} has no valid properties to serialize.");

            if (propertyUpdater != null)
            {
                Properties.ForEach(propertyUpdater);
            }
        }

        /// <summary>
        /// Get serialization/deserialization properties for the type used within the enumerable item list.
        /// </summary>
        /// <param name="items">The list of items to enumerate over and serialize.</param>
        /// <param name="propertyUpdater">Optional method to custom modify the properties as required for the requisite serializer/deserializer.</param>
        /// <param name="ci">Culture to use for header translation</param>
        public SerializerProperties(IEnumerable items, Action<PropertyAttribute> propertyUpdater = null, CultureInfo ci = null) : this(GetElementType(items), propertyUpdater, ci)
        {
            this.ItemEnumerator = items.GetEnumerator();
        }

        /// <summary>
        /// Gets a list of the column info
        /// </summary>
        public List<PropertyAttribute> Properties { get; private set; }

        /// <summary>
        /// Gets the Excel worksheet tab name.
        /// </summary>
        public string WorksheetTabName { get; private set; }

        /// <summary>
        /// Gets the data class type
        /// </summary>
        public Type EnumeratedType { get; private set; }

        /// <summary>
        /// Enumerator for list of items. This is defined here because data items may need to continue in another worksheet, file, etc
        /// </summary>
        public IEnumerator ItemEnumerator { get; private set; }

        /// <summary>
        /// Given a sequence of table headers, return a matching sequence of Properties
        /// </summary>
        /// <param name="headers">Column headers to arrange properties by</param>
        /// <returns>List of properties that match the specified headers</returns>
        public List<PropertyAttribute> ReOrderPropertiesByHeaders(IEnumerable<string> headers)
        {
            var list = new List<PropertyAttribute>();

            // this won't work if there too many, too few, duplicate, or missing headers or properties
            // list = DataClassProperties.OrderBy(o => Array.IndexOf(headers, o.Header)).ToList();

            int colIndex = 0;
            int dummyCount = 0;
            foreach (var hdr in headers)
            {
                var p = Properties.FirstOrDefault(m => m.Header.Equals(hdr, StringComparison.CurrentCultureIgnoreCase));
                if (p == null) { p = PropertyAttribute.Dummy(hdr); dummyCount++; }
                list.Add(p);
                colIndex++;
            }

            if (list.Count == 0 || list.Count == dummyCount)
                throw new ArrayTypeMismatchException("There are no matching headers in property list.");

            return list;
        }

        /// <summary>
        /// Get localized enum name by value.
        /// Using XlEnumNameAttribute or System.ComponentModel.DescriptionAttribute.
        /// Example: 
        /// enum MyEnum {
        ///     [XlEnumName("TranslationKey1")] MyValue1,
        ///     [Description("TranslationKey2")] MyValue2
        ///  }
        /// </summary>
        /// <type name="value">Enum value</type>
        /// <type name="ci">Optional Culture</type>
        /// <returns>Localized enum string name.</returns>
        /// <remarks>
        /// Use Enum.GetValues(t) to create a dictionary or list from this method.
        /// </remarks>
        public static string LocalizedEnumName(Enum value, CultureInfo ci = null)
        {
            return LocalizedStrings.GetString(value.GetType().GetField(value.ToString()).CustomAttributes.FirstOrDefault(m => m.AttributeType == typeof(XlEnumNameAttribute) || m.AttributeType == typeof(DescriptionAttribute))?.ConstructorArguments?[0].Value.ToString(), value.ToString(), ci);
        }

        private string GetWorkSheetTabName(Type dataClassType, CultureInfo ci)
        {
            XlWorkheetTabAttribute a = dataClassType.GetCustomAttribute<XlWorkheetTabAttribute>(true);
            return LocalizedStrings.GetString(a?.Id, dataClassType.Name, ci);
        }

        private static Type GetElementType(IEnumerable enumerable)
        {
            Type[] interfaces = enumerable.GetType().GetInterfaces();
            Type elementType = (from i in interfaces
                where i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>)
                select i.GetGenericArguments()[0]).FirstOrDefault();

            // Peek at the first element in the list if we couldn't determine the element type.
            if (elementType == null || elementType == typeof(object))
            {
                throw new InvalidDataException($"Cannot determine underlying type of enumerable object.");
                // First element will be lost if element is returned via 'yield return'.
                // object firstElement = enumerable.Cast<object>().FirstOrDefault();
                // if (firstElement != null) elementType = firstElement.GetType();
            }

            return elementType;
        }
    }
}
