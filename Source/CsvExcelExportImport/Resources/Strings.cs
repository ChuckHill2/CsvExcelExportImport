// <auto-generated/>
// --------------------------------------------------------------------------
// <copyright file="AssemblyInfo.cs" company="Chuck Hill">
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
namespace CsvExcelExportImport.Resources
{
    /// <summary>
    /// Resource class for looking up localized strings.
    /// </summary>
    public class Strings
    {
        private static System.Resources.ResourceManager resourceMan;

        [global::System.Diagnostics.CodeAnalysis.SuppressMessageAttribute("Microsoft.Performance", "CA1811:AvoidUncalledPrivateCode", Justification = "Not Called")]
        internal Strings()
        {
        }

        /// <summary>
        /// Gets the cached ResourceManager instance for resources within this assembly. Name/Id's are case-insensitive.
        /// </summary>
        [global::System.ComponentModel.EditorBrowsableAttribute(global::System.ComponentModel.EditorBrowsableState.Advanced)]
        public static System.Resources.ResourceManager ResourceManager
        {
            get
            {
                if (resourceMan == null)
                {
                    resourceMan = new CsvExcelExportImport.Resources.ResourceManager(typeof(Strings)) { IgnoreCase = true };
                }

                return resourceMan;
            }
        }

        /// <summary>
        /// Retrieves the resource set for a particular culture.
        /// </summary>
        /// <param name="culture">The culture whose resources are to be retrieved.</param>
        /// <param name="createIfNotExists">true to load the resource set, if it has not been loaded yet; otherwise, false.</param>
        /// <param name="tryParents">true to use resource fallback to load an appropriate resource if the resource set cannot be found; false to bypass the resource fallback process.</param>
        /// <returns>The resource set for the specified culture.</returns>
        public static global::System.Resources.ResourceSet GetResourceSet(System.Globalization.CultureInfo culture, bool createIfNotExists, bool tryParents)
        {
            return ResourceManager.GetResourceSet(culture, createIfNotExists, tryParents);
        }

        /// <summary>
        /// Returns the value of the specified string resource.
        /// </summary>
        /// <param name="name">The name/id of the resource to retrieve. It is case-insensitive.</param>
        /// <returns>The value of the resource localized for the caller's current UI culture, or null if name cannot be found in a resource set.</returns>
        public static string GetString(string name)
        {
            return ResourceManager.GetString(name);
        }

        /// <summary>
        /// Returns the value of the string resource localized for the specified culture.
        /// </summary>
        /// <param name="name">The name/id of the resource to retrieve. It is case-insensitive.</param>
        /// <param name="culture">An object that represents the culture for which the resource is localized.</param>
        /// <returns>The value of the resource localized for the specified culture, or null if name cannot be found in a resource set.</returns>
        public static string GetString(string name, System.Globalization.CultureInfo culture)
        {
            return ResourceManager.GetString(name, culture);
        }
    }
}