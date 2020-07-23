//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="LocalizationTest.cs" company="Chuck Hill">
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

using System.Collections.Generic;
using System.Globalization;
using NUnit.Framework;

namespace CsvExcelExportImport.UnitTests
{
    [TestFixture]
    public class LocalizationTest
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void TestLocalization()
        {
            CultureInfo ci = CultureInfo.GetCultureInfo("sq-AL");
            string key = "TestLocalization_Key";
            var expectedValue = "†èš‡ Łòçå∫ìžå‡ìòñ Kèÿ Ω8ß"; // Test Localization Key
            string propNameAsKey = "MyPropertyName";
            string propNameAsKeyTranslated = "πÿ Pròþèr‡ÿ Ñåmè ƒF"; // My Property Name
            string literalPropName = "Literal Value"; // not in string reources
            string value;

            value = LocalizedStrings.GetString(key, literalPropName, ci);
            Assert.AreEqual(expectedValue, value, "HappyPath Lookup");
            var lookupValue = value;

            value = LocalizedStrings.GetString(key + "BAD", literalPropName, ci);
            Assert.AreEqual(literalPropName, value, "Bad Key Lookup");

            value = LocalizedStrings.GetString(key.ToLower(), literalPropName, ci);
            Assert.AreEqual(expectedValue, value, "Case-Insensitive Key Lookup");

            value = LocalizedStrings.GetString(null, literalPropName, ci);
            Assert.AreEqual(literalPropName, value, "Null or empty key Lookup");

            value = LocalizedStrings.GetString(key + "BAD", propNameAsKey, ci);
            Assert.AreEqual(propNameAsKeyTranslated, value, "Missing key, Lookup by propertyName");

            value = LocalizedStrings.ReverseLookup(lookupValue, ci);
            Assert.AreEqual(key, value, "Reverse lookup by value");

            value = LocalizedStrings.ReverseLookup(lookupValue.ToLower(), ci);
            Assert.AreEqual(key, value, "Reverse lookup by value (case-insensitive)");

            key = "Dict1";
            expectedValue = "This is dictionary value 1";
            var dict = new Dictionary<string, string>() { { key, "This is dictionary value 1" }, { "Dict2", "This is dictionary value 2" } };
            LocalizedStrings.AddCustomResource(ci, dict);

            value = LocalizedStrings.GetString(key, ci);
            Assert.AreEqual(expectedValue, value, "Custom lookup by value");

            value = LocalizedStrings.GetString(key.ToLower(), ci);
            Assert.AreEqual(expectedValue, value, "Custom Case-Insensitive Key Lookup");
            lookupValue = value;

            value = LocalizedStrings.ReverseLookup(lookupValue.ToLower(), ci);
            Assert.AreEqual(key, value, "Custom Reverse lookup by value (case-insensitive)");
        }

        [Test]
        public void TestInvariantLocalization()
        {
            CultureInfo ci = CultureInfo.InvariantCulture;
            string key = "TestLocalization_Key";
            var expectedValue = "TestLocalization_Key"; // Test Localization Key
            string propNameAsKey = "MyPropertyName";
            string propNameAsKeyTranslated = "MyPropertyName"; // My Property Name
            string literalPropName = "Literal Value"; // not in string reources
            string value;

            value = LocalizedStrings.GetString(key, literalPropName, ci);
            Assert.AreEqual(literalPropName, value, "HappyPath Lookup");
            var lookupValue = value;

            value = LocalizedStrings.GetString(key + "BAD", literalPropName, ci);
            Assert.AreEqual(literalPropName, value, "Bad Key Lookup");

            value = LocalizedStrings.GetString(key.ToLower(), literalPropName, ci);
            Assert.AreEqual(literalPropName, value, "Case-Insensitive Key Lookup");

            value = LocalizedStrings.GetString(null, literalPropName, ci);
            Assert.AreEqual(literalPropName, value, "Null or empty key Lookup");

            value = LocalizedStrings.GetString(key + "BAD", propNameAsKey, ci);
            Assert.AreEqual(propNameAsKeyTranslated, value, "Missing key, Lookup by propertyName");

            value = LocalizedStrings.ReverseLookup(lookupValue, ci);
            Assert.AreEqual(lookupValue, value, "Reverse lookup by value");

            value = LocalizedStrings.ReverseLookup(lookupValue.ToLower(), ci);
            Assert.AreEqual(lookupValue.ToLower(), value, "Reverse lookup by value (case-insensitive)");

            key = "Dict1";
            expectedValue = "Dict1";
            var dict = new Dictionary<string, string>() { { key, "This is dictionary value 1" }, { "Dict2", "This is dictionary value 2" } };
            LocalizedStrings.AddCustomResource(ci, dict);

            value = LocalizedStrings.GetString(key, ci);
            Assert.AreEqual(expectedValue, value, "Custom lookup by value");

            value = LocalizedStrings.GetString(key.ToLower(), ci);
            Assert.AreEqual(expectedValue.ToLower(), value, "Custom Case-Insensitive Key Lookup");
            lookupValue = value;

            value = LocalizedStrings.ReverseLookup(lookupValue.ToLower(), ci);
            Assert.AreEqual(lookupValue.ToLower(), value, "Custom Reverse lookup by value (case-insensitive)");
        }
    }
}