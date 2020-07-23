//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="PropertyAttribute.cs" company="Chuck Hill">
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
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace CsvExcelExportImport
{
    /// <summary>
    /// Actual model property/column meta-properties
    /// </summary>
    internal class PropertyAttribute
    {
        /// <summary>
        /// Gets the column property name
        /// </summary>
        public string Name { get; private set; }

        #region XlColumnAttribute Values
        /// <summary>
        /// Gets the column header
        /// </summary>
        public string Header { get; private set; }

        /// <summary>
        /// Column order. If undefined, the column order is ordered alphabetically by 
        /// translated header name. Columns with order defined take precedence. All 
        /// unordered columns follow in alphabetical order by translated header name.
        /// </summary>
        public int Order { get; private set; }

        /// <summary>
        /// Formats the column using the specified format and the formatting
        /// conventions of the current culture. The format values are the same 
        /// as used in value.ToString(format) but does not really convert the 
        /// value to a string, just sets the format for the column.
        /// Invalid codes default to Excel auto-format.
        /// </summary>
        public string Format { get; private set; }

        /// <summary>
        /// Relative column index (to this column index) that contains the units styling for this value.
        /// Zero is defined as no relative units styling.
        /// </summary>
        public int RelUnitsIndex { get; set; }

        /// <summary>
        /// All columns preceding this column are fixed and do not move when scrolling 
        /// left and right. They are always visible on the page while scrolling.
        /// </summary>
        public bool Frozen { get; private set; }

        /// <summary>
        /// Excel column has a filter dropdown.
        /// </summary>
        public bool HasFilter { get; private set; }

        /// <summary>
        /// Set column justification.
        /// Choices are: Left, Center, Right, and Auto, where Auto is the default.
        /// </summary>
        public Justified Justification { get; private set; }

        /// <summary>
        /// Column is hidden from view.
        /// </summary>
        public bool Hidden { get; set; }

        /// <summary>
        /// Translate string data in this column by using the value as the ID/Key.
        /// </summary>
        public bool TranslateData { get; private set; }

        /// <summary>
        /// Gets or sets fixed column width. Units are in characters. A 'character' is defined as
        /// the width of a numeric character of the Excel default font (Calibri,11). Autoformat
        /// adds 0.29 characters for margins. When this is set, AutoWrap is enabled.
        /// </summary>
        public int MaxWidth { get; set; }
        #endregion

        /// <summary>
        /// Gets a delegate function for retrieving the value
        /// </summary>
        public Func<object, object> GetValue { get; set; }

        /// <summary>
        /// Gets a delegate function for retrieving the value
        /// </summary>
        public Action<object, object> SetValue { get; set; }

        /// <summary>
        /// Gets or sets the column type. If Nullable type, it is unwrapped.
        /// </summary>
        public Type CellType { get; set; }

        /// <summary>
        /// Gets the original class property type.
        /// </summary>
        public Type PropertyType { get; private set; }
        
        /// <summary>
        /// Flag if this property is Nullable.
        /// </summary>
        public bool IsNullable { get; private set; }

        /// <summary>
        /// Gets or sets the Excel column style
        /// </summary>
        public string Style { get; set; }

        /// <summary>
        /// Gets or sets the custom property data used by caller for whatever.
        /// </summary>
        public object UserData { get; set; }

        /// <summary>
        /// Instantiates an empty PropertyAttribute. Used within this class only.
        /// See: Dummy(...) and GetProperties(...)
        /// </summary>
        private PropertyAttribute()
        {
        }

        /// <summary>
        /// Create new Property Attribute object with all necessary values pre-computed for fast usability.
        /// </summary>
        /// <param name="p">Property info to use to set values.</param>
        /// <param name="ci">Culture to use for localization of header</param>
        private PropertyAttribute(PropertyInfo p, CultureInfo ci)
        {
            Name = p.Name;
            PropertyType = p.PropertyType;
            CellType = p.PropertyType.IsGenericType ? p.PropertyType.GenericTypeArguments[0] : p.PropertyType;
            IsNullable = p.PropertyType.IsGenericType;
            GetValue = (o) => p.GetValue(o);
            SetValue = (o, v) => p.SetValue(o, Cast.To(p.PropertyType, v));

            if (CellType == typeof(bool))
            {
                GetValue = (o) => p.GetValue(o)?.ToString();  // Excel uses TRUE/FALSE. We like True/False
            }
            else if (CellType == typeof(string))
            {
                GetValue = (o) => p.GetValue(o)?.ToString().AppendSp();
                SetValue = (o, v) => p.SetValue(o, v.ToString().TrimSp());
            }
            else if (CellType == typeof(char))  // EPPlus assumes primitive types are always numbers!
            {
                GetValue = (o) => p.GetValue(o)?.ToString().AppendSp();
                SetValue = (o, v) => p.SetValue(o, v?.ToString()?[0]);
            }

            XlColumnAttribute a = p.GetCustomAttribute<XlColumnAttribute>(true);
            if (a == null)
            {
                Header = LocalizedStrings.GetString(p.Name, null, ci);
            }
            else
            {
                if (ci.Name != string.Empty && a.TranslateData)
                {
                    if (CellType == typeof(string) || CellType == typeof(bool))
                    {
                        GetValue = (o) =>
                        {
                            var v = p.GetValue(o);
                            if (v == null) return null;
                            return LocalizedStrings.GetString(v.ToString(), ci);
                        };
                        SetValue = (o, v) =>
                        {
                            if (!string.IsNullOrWhiteSpace(v as string)) p.SetValue(o, Cast.To(p.PropertyType, LocalizedStrings.ReverseLookup(v.ToString(), ci)));
                        };
                    }
                    else if (CellType.IsEnum)
                    {
                        // Note: ExcelSerializer restricts values to a limited set of choices.
                        var vals = Enum.GetValues(CellType);
                        var elookup = new Dictionary<Enum, string>(vals.Length);
                        var erevlookup = new Dictionary<string, Enum>(vals.Length);
                        foreach (Enum e in vals)
                        {
                            var s = SerializerProperties.LocalizedEnumName(e, ci);
                            elookup.Add(e, s);
                            erevlookup.Add(s, e);
                        }

                        GetValue = (o) =>
                        {
                            var v = p.GetValue(o);
                            if (v == null) return null;
                            return elookup[(Enum)v];
                        };
                        SetValue = (o, v) =>
                        {
                            if (!string.IsNullOrWhiteSpace(v as string)) p.SetValue(o, erevlookup[(string)v]);
                        };
                    }
                }

                Header = LocalizedStrings.GetString(a.Id, p.Name, ci);
                Format = a.Format;
                Frozen = a.Frozen;
                HasFilter = a.HasFilter;
                Justification = a.Justification;
                Hidden = a.Hidden;
                TranslateData = a.TranslateData;
                MaxWidth = a.MaxWidth;

                // Format styling may contain 2 comma delimited fields. The first field contains code char + int
                // where the int represents the number of digits after the decimal (aka precision). The optional 2nd
                // field contains the units string(including leading and trailing whitespace). However if the 2nd
                // field contains an integer, it is used as a relative column index to the column value containing the units.
                if (!string.IsNullOrWhiteSpace(Format) && (Format[0] == 'f' || Format[0] == 'F' || 
                                                           Format[0] == 'n' || Format[0] == 'N' || 
                                                           Format[0] == 'c' || Format[0] == 'C'))
                {
                    var e = Format.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                    if (e.Length > 1 && int.TryParse(e[1].Trim(), out var index))
                    {
                        RelUnitsIndex = index;
                        Format = e[0];
                    }
                }
            }
        }

        /// <summary>
        /// Used for write when there are more columns in in the source
        /// CSV or Excel than in the the destination class data model.
        /// </summary>
        /// <param name="headerName">Name of column header in the source CSV or Excel.</param>
        /// <returns>Dummy property attribute.</returns>
        public static PropertyAttribute Dummy(string headerName)
        {
            return new PropertyAttribute()
            {
                Name = "Missing",
                Header = headerName,
                CellType = typeof(DBNull),
                GetValue = obj => null,
                SetValue = (obj, val) => { }
            };
        }

        /// <summary>
        /// Get a list of all the valid/sorted properties of specified type to write. Invalid/ignored properties
        /// consist of private, static or write-only properties, and properties marked with XlIgnore attribute.
        /// All fields are ignored. Does not recurse into child classes or 
        /// enumerable objects. Uses public field 'ColumnProperties' for user-defined column attributes.
        /// If there is no ColumnProperty for the given reflected class property, the column property 
        /// information is extracted the [Display] or [DisplayName] attributes, if they exist.
        /// </summary>
        /// <param name="t">Type of data class to reflect</param>
        /// <param name="ci">Culture to use for header translation</param>
        /// <returns>List properties. Count may be zero, but never null.</returns>
        public static List<PropertyAttribute> GetProperties(Type t, CultureInfo ci)
        {
            // https://stackoverflow.com/questions/9062235/get-properties-in-order-of-declaration-using-reflection/17998371
            var list = t.GetProperties()
                .Where(p => p.CanRead && p.CanWrite &&
                            !p.GetMethod.IsPrivate && !p.SetMethod.IsPrivate &&
                            !p.PropertyType.IsArray &&
                            ((p.PropertyType.IsGenericType && p.PropertyType.GetGenericTypeDefinition() == typeof(System.Nullable<>)) ||
                             !p.PropertyType.IsGenericType) &&
                            !p.CustomAttributes.Any(m => m.AttributeType.Name.Contains("Ignore")))
                .OrderBy(x => x.MetadataToken)
                .Select(p => new PropertyAttribute(p, ci))
                .ToList();

            // Validate RelUnitsIndex
            for (int c = 0; c < list.Count; c++)
            {
                var p = list[c];
                if (p.RelUnitsIndex == 0) continue;
                var unitsIndex = c + p.RelUnitsIndex;
                if (unitsIndex < 0 || unitsIndex >= list.Count) { p.RelUnitsIndex = 0; continue; }
                if (list[unitsIndex].CellType != typeof(string)) { p.RelUnitsIndex = 0; continue; }
            }

            return list;
        }

        /// <inheritdoc />
        public override string ToString() => $"\"{Name}\",\"{Header}\",{Order},{CellType.Name + (IsNullable ? "?" : string.Empty)}";
    }
}
