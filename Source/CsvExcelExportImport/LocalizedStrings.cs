//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="LocalizedStrings.cs" company="Chuck Hill">
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
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Threading;

namespace CsvExcelExportImport
{
    /// <summary>
    /// Get localized string by key or get key by localized name.
    /// Automatically finds all resources to retrieve data from.
    /// </summary>
    public static class LocalizedStrings
    {
        private static readonly CustomResourceManager MyResMan = new CustomResourceManager();
        private static ResourceManager[] _resourceManagers;

        private static ResourceManager[] ResourceManagers
        {
            get
            {
                // Lazy instantiation. Get string resources only after most,
                // if not all assemblies of interest have been loaded.
                if (_resourceManagers == null) _resourceManagers = GetStringsResourceManagers();
                return _resourceManagers;
            }
        }

        /// <summary>
        /// Returns the value of the string resource localized for the specific culture.
        /// If not found, return the ResourceKey.
        /// If CurrentUICulture is the Invariant culture, no translation is performed.
        /// All lookups are case-insensitive.
        /// </summary>
        /// <param name="resKey">Resource key</param>
        /// <returns>Localized string or resource key if all else fails</returns>
        public static string GetString(string resKey)
        {
            return GetString(resKey, null, null);
        }

        /// <summary>
        /// Returns the value of the string resource localized for the specific culture.
        /// If not found, tries to lookup localized string resource value by PropertyName (aka 'default').
        /// If not found, return the ResourceKey.
        /// If PropertyName undefined, return the ResourceKey.
        /// If ResourceKey undefined, return "(null)".
        /// If CurrentUICulture is the Invariant culture, no translation is performed.
        /// All lookups are case-insensitive.
        /// </summary>
        /// <param name="resKey">Resource key</param>
        /// <param name="defalt">Default literal value or PropertyName</param>
        /// <returns>Localized string or resource key if all else fails</returns>
        public static string GetString(string resKey, string defalt)
        {
            return GetString(resKey, defalt, null);
        }

        /// <summary>
        /// Returns the value of the string resource localized for the specific culture.
        /// If not found, return the ResourceKey.
        /// If ResourceKey undefined, return "(null)".
        /// If CurrentUICulture is the Invariant culture, no translation is performed.
        /// All lookups are case-insensitive.
        /// </summary>
        /// <param name="resKey">Resource key</param>
        /// <param name="ci">CultureInfo to translate string into. null == CurrentUICulture.</param>
        /// <returns>Localized string or resource key if all else fails</returns>
        public static string GetString(string resKey, CultureInfo ci)
        {
            return GetString(resKey, null, ci);
        }

        /// <summary>
        /// Returns the value of the string resource localized for the specific culture.
        /// If not found, tries to lookup localized string resource value by PropertyName (aka 'default').
        /// If not found, return the PropertyName.
        /// If PropertyName undefined, return the ResourceKey.
        /// If ResourceKey undefined, return "(null)".
        /// If CurrentUICulture is the Invariant culture, no translation is performed.
        /// All lookups are case-insensitive.
        /// </summary>
        /// <param name="resKey">Resource key</param>
        /// <param name="defalt">Default literal value or PropertyName</param>
        /// <param name="ci">CultureInfo to translate string into. null == CurrentUICulture.</param>
        /// <returns>Localized string or resource key if all else fails</returns>
        public static string GetString(string resKey, string defalt, CultureInfo ci)
        {
            // Do not translate if this is the Invariant culture.
            if ((ci ?? System.Threading.Thread.CurrentThread.CurrentUICulture).Name != string.Empty)
            {
                if (!string.IsNullOrWhiteSpace(resKey))
                {
                    foreach (var rm in ResourceManagers)
                    {
                        string value = rm.GetString(resKey, ci);
                        if (!string.IsNullOrWhiteSpace(value)) return value;
                    }
                }

                if (!string.IsNullOrWhiteSpace(defalt))
                {
                    foreach (var rm in ResourceManagers)
                    {
                        string value = rm.GetString(defalt, ci);
                        if (!string.IsNullOrWhiteSpace(value)) return value;
                    }
                }
            }

            return defalt ?? resKey ?? "(null)";
        }

        /// <summary>
        /// Get resource string ID by value (case-insensitive).
        /// </summary>
        /// <param name="value">String literal search for.</param>
        /// <param name="ci">Optional CultureInfo to search. default == CurrentUICulture.</param>
        /// <returns>Resource string ID or null if not found.</returns>
        public static string ReverseLookup(string value, CultureInfo ci = null)
        {
            if (ci == null) ci = Thread.CurrentThread.CurrentUICulture;
            if (ci.Name == string.Empty || string.IsNullOrWhiteSpace(value)) return value; // Is Invariant culture. No translation.

            Func<string, string, bool> equals = (s1, s2) => ci.CompareInfo.Compare(s1, s2, CompareOptions.IgnoreCase) == 0;
            // var comparer = ci.CompareInfo.GetStringComparer(CompareOptions.IgnoreCase);

            foreach (var rm in ResourceManagers)
            {
                // var entry = rm.GetResourceSet(ci, true, true).OfType<DictionaryEntry>().FirstOrDefault(e => comparer.Equals(e.Value.ToString(), value));
                var entry = rm.GetResourceSet(ci, true, true).OfType<DictionaryEntry>().FirstOrDefault(e => equals(e.Value.ToString(), value));
                if (entry.Key == null) continue;
                return entry.Key.ToString();
            }

            return value;
        }

        /// <summary>
        /// Dynamically add custom string resource lookup dictionary during runtime. 
        /// This is the last string resource in which to find a match. Assembly 
        /// resources take precedence. Duplicate calls for the same culture will 
        /// merge strings with pre-existing ones.
        /// </summary>
        /// <param name="ci">Required CultureInfo to associate with the lookup dictionary.</param>
        /// <param name="stringLookup">String lookup dictionary</param>
        public static void AddCustomResource(CultureInfo ci, Dictionary<string, string> stringLookup)
        {
            MyResMan.Add(ci.Name, stringLookup);
        }

        /// <summary>
        /// Remove a custom resource for a specific culture.
        /// </summary>
        /// <param name="ci">CultureInfo to remove.</param>
        public static void RemoveCustomResource(CultureInfo ci)
        {
            MyResMan.Remove(ci.Name);
        }

        /// <summary>
        /// Clear/purge all custom resources.
        /// </summary>
        public static void ClearCustomResources()
        {
            MyResMan.Clear();
        }

        /// <summary>
        /// Force reloading of all assembly resource managers. Only necessary if an
        /// assembly containing requisite strings was not loaded at the time of first
        /// use of GetString(). Does not effect custom resource strings. To clear custom
        /// resources use ClearCustomResources().
        /// </summary>
        public static void Refresh()
        {
            _resourceManagers = null;
        }

        /// <summary>
        /// Get Strings ResourceManager object from ALL currently loaded assemblies
        /// that contain the class Strings.ResourceManager public static property.
        /// </summary>
        /// <returns>Array of ResourceManager objects. Will always contain </returns>
        private static ResourceManager[] GetStringsResourceManagers()
        {
            List<ResourceManager> resManagers = new List<ResourceManager>();
            foreach (var asm in AppDomain.CurrentDomain.GetAssemblies())
            {
                try
                {
                    if (asm.IsDynamic) continue; // assemblies built on-the-fly do not have resources.

                    Type resType = asm.ExportedTypes.FirstOrDefault(t => t.IsClass && t.Name.Equals("Strings", StringComparison.Ordinal) && t.GetProperty("ResourceManager", BindingFlags.Public | BindingFlags.Static) != null);
                    if (resType != null)
                    {
                        // We don't use 'new ResourceManager(resType)' because it creates a copy. Best to use the original.
                        var resman = resType.GetProperty("ResourceManager", BindingFlags.Public | BindingFlags.Static).GetValue(null) as ResourceManager;
                        resman.IgnoreCase = true;  // we need to support case-insensitive keys
                        resManagers.Add(resman);
                    }
                }
                catch
                {
                    // If a dependent assembly is not found (because they are only loaded upon demand), it will throw an error.
                    // This occurs during enumeration of asm.ExportedTypes.
                }
            }

            // Add custom Resource manager so users can specify custom language-specific string resources at runtime.
            resManagers.Add(MyResMan);

            return resManagers.ToArray();
        }

        /// <summary>
        /// A custom resource manager to handle user-defined string resources during runtime.
        /// It is a ResourceManager type to access it identically to system resource managers.
        /// It is also an IDictionary to add and remove culture/string lookups.
        /// </summary>
        private class CustomResourceManager : ResourceManager, IDictionary<string, Dictionary<string, string>>
        {
            private Dictionary<string, Dictionary<string, string>> _resources = null;

            private Dictionary<string, Dictionary<string, string>> Resources
            {
                get
                {
                    if (_resources == null) _resources = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);
                    return _resources;
                }
            }

            public CustomResourceManager() { }
            private CustomResourceManager(Type t) { }

            public override string GetString(string name) => this.GetString(name, null);

            public override string GetString(string name, CultureInfo culture)
            {
                if (Resources.TryGetValue(culture?.Name ?? CultureInfo.CurrentUICulture.Name, out var res))
                {
                    if (res.TryGetValue(name, out var result))
                        return result;
                }

                return null;
            }

            public override ResourceSet GetResourceSet(CultureInfo culture, bool createIfNotExists, bool tryParents) => new ResourceSet(new ResReader(Resources, culture));

            #region Unused ResourceManager Members
            public override object GetObject(string name) => throw new NotSupportedException();
            public override object GetObject(string name, CultureInfo culture) => throw new NotSupportedException();
            public override string BaseName => string.Empty;
            public override bool IgnoreCase { get => true; set { } }
            public override Type ResourceSetType => typeof(Dictionary<string, string>);
            protected override ResourceSet InternalGetResourceSet(CultureInfo culture, bool createIfNotExists, bool tryParents) => throw new NotSupportedException();
            protected override string GetResourceFileName(CultureInfo culture) => throw new NotSupportedException();

            public override void ReleaseAllResources()
            {
                if (_resources != null)
                {
                    foreach (var kv in _resources)
                    {
                        kv.Value.Clear();
                    }

                    _resources.Clear();
                    _resources = null;
                }
            }
            #endregion

            #region Dictionary Members
            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
            public IEnumerator<KeyValuePair<string, Dictionary<string, string>>> GetEnumerator() => Resources.GetEnumerator();
            public void Add(KeyValuePair<string, Dictionary<string, string>> item) => this.Add(item.Key, item.Value);
            public void Clear() => ReleaseAllResources();
            public bool Contains(KeyValuePair<string, Dictionary<string, string>> item) => Resources.Contains(item);
            public void CopyTo(KeyValuePair<string, Dictionary<string, string>>[] array, int arrayIndex) => throw new NotImplementedException();
            public bool Remove(KeyValuePair<string, Dictionary<string, string>> item) => Resources.Remove(item.Key);
            public int Count => Resources.Count;
            public bool IsReadOnly => false;
            public bool ContainsKey(string key) => Resources.ContainsKey(key);
            public bool Remove(string key) => Resources.Remove(key);
            public bool TryGetValue(string key, out Dictionary<string, string> value) => Resources.TryGetValue(key, out value);
            public ICollection<string> Keys => Resources.Keys;
            public ICollection<Dictionary<string, string>> Values => Resources.Values;

            public void Add(string key, Dictionary<string, string> value)
            {
                if (value == null || value.Count == 0) return;
                key = CultureInfo.GetCultureInfo(key).ToString();  // verify and normalize cultureInfo string

                if (!Resources.TryGetValue(key, out var existingValue) || existingValue == null)
                {
                    var newValue = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (var kv in value)
                    {
                        newValue[kv.Key] = kv.Value;
                    }

                    Resources.Add(key, newValue);
                    return;
                }

                foreach (var kv in value)
                {
                    existingValue[kv.Key] = kv.Value;
                }
            }

            public Dictionary<string, string> this[string key]
            {
                get => Resources[key];
                set => this.Add(key, value);
            }
            #endregion

            private class ResReader : IResourceReader
            {
                private readonly Dictionary<string, string> res;

                public ResReader(Dictionary<string, Dictionary<string, string>> allResources, CultureInfo culture)
                {
                    allResources.TryGetValue(culture.Name, out res);
                    if (res == null) res = new Dictionary<string, string>();
                }

                public void Close() { }
                public void Dispose() { }
                public IDictionaryEnumerator GetEnumerator() => res.GetEnumerator();
                IEnumerator IEnumerable.GetEnumerator() => res.GetEnumerator();
            }
        }
    }
}
