//--------------------------------------------------------------------------
// <summary>
// Simple import/export of array of classes to Excel or CSV.
// </summary>
// <copyright file="ExcelCommon.cs" company="Chuck Hill">
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
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml;

namespace CsvExcelExportImport
{
    internal static class ExcelCommon
    {
        /// <summary>
        /// Values needed by the caller of SetWorkbookProperties()
        /// </summary>
        public class ExcelProps
        {
            public WorkbookProperties.RegionInfoOverride CustomRegionInfo { get; set; }
            public Color Dark { get; set; }
            public Color Medium { get; set; }
            public Color Light { get; set; }
        }

        public static ExcelProps SetWorkbookProperties(ExcelWorkbook wb, Guid excelIdentifier, WorkbookProperties wbProps = null)
        {
            if (wbProps?.Culture != null)
            {
                System.Threading.Thread.CurrentThread.CurrentUICulture = wbProps.Culture;
                System.Threading.Thread.CurrentThread.CurrentCulture = wbProps.Culture.Name == string.Empty ? CultureInfo.GetCultureInfo("en-US") : wbProps.Culture;
            }

            var customRegionInfo = wbProps?.CustomRegionInfo.Clone() ?? new WorkbookProperties.RegionInfoOverride();
            customRegionInfo.InitializeUndefined(CultureInfo.CurrentCulture); // CurrentUICulture for language, CurrentCulture for region

            wb.Properties.SetCustomPropertyValue("RegionInfoOverride", customRegionInfo.Serialize());
            wb.Properties.SetCustomPropertyValue("ExcelIdentifier", excelIdentifier.ToString());

            wb.Properties.Created = DateTime.Now;
            wb.Properties.SetCustomPropertyValue("CultureName", System.Globalization.CultureInfo.CurrentUICulture.ToString()); // Used for deserialization.
            wb.Properties.SetCustomPropertyValue("Version", Assembly.GetCallingAssembly().GetName().Version.ToString());
            wb.Properties.Company = Assembly.GetCallingAssembly().GetCustomAttribute<AssemblyCompanyAttribute>()?.Company ?? string.Empty;

            Color themeColor = wbProps?.ThemeColor ?? WorkbookProperties.ExcelGreen;
            Color dark, medium, light;
            if (themeColor == Color.White || themeColor == Color.Black)
            {
                dark = Color.Gray;
                medium = Color.White;
                light = Color.White;
            }
            else if (themeColor == Color.Transparent)
            {
                dark = themeColor;
                medium = themeColor;
                light = themeColor;
            }
            else
            {
                dark = SetLuminance(themeColor, 85);
                medium = SetLuminance(themeColor, 180);
                light = SetLuminance(themeColor, 235);
            }

            wb.Properties.SetCustomPropertyValue("ThemeColor", themeColor.Name);

            var outProps = new ExcelProps()
            {
                CustomRegionInfo = customRegionInfo,
                Dark = dark,
                Medium = medium,
                Light = light
            };

            if (wbProps == null) return outProps;  // No custom work properties

            if (wbProps.Title != null) wb.Properties.Title = wbProps.Title;
            if (wbProps.Subject != null) wb.Properties.Subject = wbProps.Subject;
            if (wbProps.Author != null) wb.Properties.Author = wbProps.Author;
            if (wbProps.Manager != null) wb.Properties.Manager = wbProps.Manager;
            if (wbProps.Company != null) wb.Properties.Company = wbProps.Company;
            if (wbProps.Category != null) wb.Properties.Category = wbProps.Category;
            if (wbProps.Keywords != null) wb.Properties.Keywords = wbProps.Keywords; // comma-delimited keywords
            if (wbProps.Comments != null) wb.Properties.Comments = wbProps.Comments;
            if (wbProps.HyperlinkBase != null) wb.Properties.HyperlinkBase = wbProps.HyperlinkBase;
            if (wbProps.Status != null) wb.Properties.Status = wbProps.Status;
            // wb.Properties.Application; -- do not modify. This is always set to "Microsoft Excel"
            // wb.Properties.AppVersion; -- do not modify. This is the Excel version!
            // wb.Properties.LastModifiedBy --set upon save
            // wb.Properties.Modified; --set upon save

            foreach (var kv in wbProps.ExtraProperties)
            {
                if (string.IsNullOrWhiteSpace(kv.Value)) continue;
                if (kv.Key.Equals("Created", StringComparison.OrdinalIgnoreCase)) continue; // reserved
                if (kv.Key.Equals("LastModifiedBy", StringComparison.OrdinalIgnoreCase)) continue;
                if (kv.Key.Equals("Modified", StringComparison.OrdinalIgnoreCase)) continue;
                wb.Properties.SetCustomPropertyValue(kv.Key, kv.Value);
            }

            return outProps;
        }

        public static void GetWorkbookProperties(ExcelWorkbook wb, Guid excelIdentifier, WorkbookProperties wbProps)
        {
            if (wb.Properties.GetCustomPropertyValue("ExcelIdentifier")?.ToString() != excelIdentifier.ToString())
            {
                throw new FormatException("Not a valid Excel workbook generated by this library.");
            }

            if (wbProps == null) throw new ArgumentNullException(nameof(wbProps));

            wbProps.CustomRegionInfo.Deserialize(wb.Properties.GetCustomPropertyValue("RegionInfoOverride")?.ToString(), wbProps.Culture);

            wbProps.ExtraProperties.Clear();
            foreach (var key in GetCustomWorkbookPropertyNames(wb))
            {
                if (key.Equals("CultureName", StringComparison.OrdinalIgnoreCase)) continue; // reserved
                if (key.Equals("RegionInfoOverride", StringComparison.OrdinalIgnoreCase)) continue;
                if (key.Equals("ThemeColor", StringComparison.OrdinalIgnoreCase)) continue;
                if (key.Equals("ExcelIdentifier", StringComparison.OrdinalIgnoreCase)) continue;
                wbProps.ExtraProperties.Add(key, wb.Properties.GetCustomPropertyValue(key)?.ToString());
            }

            wbProps.ExtraProperties.Add("Created", wb.Properties.Created.ToString("yyyy-MM-dd HH:mm:ss"));
            wbProps.ExtraProperties.Add("LastModifiedBy", wb.Properties.LastModifiedBy ?? string.Empty);
            wbProps.ExtraProperties.Add("Modified", wb.Properties.Modified.ToString("yyyy-MM-dd HH:mm:ss"));

            wbProps.Culture = CultureInfo.GetCultureInfo(wb.Properties.GetCustomPropertyValue("CultureName") as string ?? CultureInfo.CurrentUICulture.Name);
            var tcolor = wb.Properties.GetCustomPropertyValue("ThemeColor")?.ToString();
            wbProps.ThemeColor = tcolor == null ? (Color?)null : int.TryParse(tcolor, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var icolor) ? Color.FromArgb(icolor) : Color.FromName(tcolor);

            wbProps.Title = wb.Properties.Title;
            wbProps.Subject = wb.Properties.Subject;
            wbProps.Author = wb.Properties.Author;
            wbProps.Manager = wb.Properties.Manager;
            wbProps.Company = wb.Properties.Company;
            wbProps.Category = wb.Properties.Category;
            wbProps.Keywords = wb.Properties.Keywords;
            wbProps.Comments = wb.Properties.Comments;
            wbProps.HyperlinkBase = wb.Properties.HyperlinkBase;
            wbProps.Status = wb.Properties.Status;
        }

        /// <summary>
        /// Set Worksheet print properties
        /// </summary>
        /// <param name="ws">Worksheet object whose print properties to set.</param>
        /// <param name="ci">Culture to use. If undefined, uses current UI culture.</param>
        public static void SetPrintProperties(ExcelWorksheet ws, CultureInfo ci = null)
        {
            ci = ci ?? System.Threading.Thread.CurrentThread.CurrentUICulture;

            ws.HeaderFooter.differentOddEven = false;
            ws.HeaderFooter.OddHeader.CenteredText = "&16&\"Arial,Bold\" " + ws.Name;
            ws.HeaderFooter.OddFooter.LeftAlignedText = "&10&\"Arial,Regular\" " + ws.Workbook.Properties.Company + " " + LocalizedStrings.GetString("Confidential", ci);
            ws.HeaderFooter.OddFooter.CenteredText = "&10&\"Arial,Regular\" " + ws.Workbook.Properties.Created.ToString("d");
            ws.HeaderFooter.OddFooter.RightAlignedText = "&10&\"Arial,Regular\" " + string.Format(LocalizedStrings.GetString("Page {0} of {1}", ci), ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
            ws.PrinterSettings.RepeatRows = ws.Cells["1:1"];
            ws.PrinterSettings.ShowGridLines = true;
            ws.PrinterSettings.FitToPage = true;
            ws.PrinterSettings.FitToWidth = 1;
            ws.PrinterSettings.FitToHeight = 32767;
            ws.PrinterSettings.Orientation = eOrientation.Landscape;
            ws.PrinterSettings.BottomMargin = 0.5m;
            ws.PrinterSettings.TopMargin = 0.5m;
            ws.PrinterSettings.LeftMargin = 0.5m;
            ws.PrinterSettings.RightMargin = 0.5m;
            ws.PrinterSettings.HorizontalCentered = true;
        }

        /// <summary>
        /// Disable cell error checking for text cells containing numbers.
        /// Do not call more than once per worksheet.
        /// </summary>
        /// <param name="ws">Worksheet to operate upon</param>
        public static void DisableCellWarnings(ExcelWorksheet ws)
        {
            var xdoc = ws.WorksheetXml;
            var ns = xdoc.ChildNodes[1].NamespaceURI;

            // Many of the text identifiers are numerical in nature. Excel does not like numerical text data. It thinks 
            // it needs to be converted and puts a tiny green triangle in the corner of the offending cells. When one 
            // hovers over it, the tooltip says 'The number in this cell is formatted as text...'. Very annoying.
            // NPPlus does not support disabling the cell warning, so we have to modify the low-level worksheet xml document.
            // See Hint: http://stackoverflow.com/questions/11858109/using-epplus-excel-how-to-ignore-excel-error-checking-or-remove-green-tag-on-t

            var node = xdoc.ChildNodes[1].AppendChild(xdoc.CreateElement("ignoredErrors", ns)).AppendChild(xdoc.CreateElement("ignoredError", ns));
            node.Attributes.Append(xdoc.CreateAttribute("sqref")).Value = "A:" + ExcelCellAddress.GetColumnLetter(ws.Dimension.End.Column);
            node.Attributes.Append(xdoc.CreateAttribute("numberStoredAsText")).Value = "1";
        }

        /// <summary>
        /// Hide dropdown button on all cell multi-select choice lists.
        /// </summary>
        /// <param name="ws">Worksheet to operate upon</param>
        public static void HideCellValidationDropdowns(ExcelWorksheet ws)
        {
            var xdoc = ws.WorksheetXml;
            var ns = xdoc.ChildNodes[1].NamespaceURI;
            var nsmgr = new XmlNamespaceManager(xdoc.NameTable);
            nsmgr.AddNamespace("ns", ns);

            foreach (XmlElement n in xdoc.SelectNodes("//ns:dataValidations/ns:dataValidation[@type='list']", nsmgr)) // there may be more than one.
            {
                (n.Attributes["showDropDown"] ?? n.Attributes.Append(xdoc.CreateAttribute("showDropDown"))).Value = "1";
            }
        }

        /// <summary>
        /// Auto-fit all columns in worksheet with optional header autowrap on words.
        /// </summary>
        /// <param name="ws">Worksheet to column autofit.</param>
        /// <param name="headerRowIndex">Row index of header row.</param>
        /// <param name="autowrapHeaders">True to auto wrap header strings on words.</param>
        public static void AutoFitColumns(ExcelWorksheet ws, int headerRowIndex, bool autowrapHeaders)
        {
            var maxRows = ws.Dimension.Rows > 1000 ? 1000 : ws.Dimension.Rows;  // for performance, just autosize on the first 1000 rows

            // https://stackoverflow.com/questions/51432238/hide-column-in-epplus-not-working
            var hiddenCols = Enumerable.Range(1, ws.Dimension.Columns).Select(i => ws.Column(i).Hidden ? ws.Column(i) : null).Where(m => m != null).ToArray();

            if (!autowrapHeaders)
            {
                ws.Row(headerRowIndex).Style.WrapText = false;
                ws.Cells[headerRowIndex, 1, maxRows, ws.Dimension.Columns - 1].AutoFitColumns(3, 200);
                Array.ForEach(hiddenCols, c => c.Hidden = true); // Restore hidden flag
                return;
            }

            ws.Row(headerRowIndex).Style.WrapText = true; // flags that will not use headers in computing width

            Font nfont = null;
            Bitmap b = null;
            Graphics g = null;
            try
            {
                // Code extracted from EPPlus
                var nf = ws.Cells[headerRowIndex, 1].Style.Font;  // All cells in the header have the same style.
                var fs = FontStyle.Regular;
                if (nf.Bold) fs |= FontStyle.Bold;
                if (nf.UnderLine) fs |= FontStyle.Underline;
                if (nf.Italic) fs |= FontStyle.Italic;
                if (nf.Strike) fs |= FontStyle.Strikeout;
                nfont = new Font(nf.Name, nf.Size, fs);

                b = new Bitmap(1, 1);
                g = Graphics.FromImage(b);
                g.PageUnit = GraphicsUnit.Pixel;
                double normalSize = decimal.ToDouble(ws.Workbook.MaxFontWidth); // ==7
                // double normalSize = g.MeasureString("_8_", nfont).Width - g.MeasureString("__", nfont).Width; // ==7.6555976867675781

                for (int i = 1; i <= ws.Dimension.Columns; i++)
                {
                    var range = ws.Cells[headerRowIndex, i, maxRows, i];
                    var hdr = ws.Cells[headerRowIndex, i].Value.ToString();

                    // split words on ANY unicode whitespace boundary and get size of widest word.

                    var minWidth = 3.0; // Single char cell is a perfect square.
                    foreach (var word in Regex.Split(hdr, @"\p{Z}").Where(w => !string.IsNullOrWhiteSpace(w)))
                    {
                        var size = g.MeasureString(word, nfont, 10000, StringFormat.GenericDefault);
                        var mw = (size.Width + 5) / normalSize;
                        if (mw > minWidth) minWidth = mw;
                    }

                    bool isWrapped = false;
                    if (ws.Column(i).Style.WrapText)
                    {
                        isWrapped = true;
                        ws.Column(i).Style.WrapText = false;
                    }

                    var maxWidth = ws.DefaultColWidth == ws.Column(i).Width ? 240.0 : ws.Column(i).Width;
                    range.AutoFitColumns(minWidth, maxWidth);

                    if (isWrapped) ws.Column(i).Style.WrapText = true;
                }
            }
            finally
            {
                nfont?.Dispose();
                b?.Dispose();
                g?.Dispose();
            }

            Array.ForEach(hiddenCols, c => c.Hidden = true); // Restore hidden flag
        }

        private static string[] GetCustomWorkbookPropertyNames(ExcelWorkbook wb)
        {
            var ns = wb.Properties.GetType().GetProperty("NameSpaceManager", BindingFlags.Instance | BindingFlags.NonPublic)?.GetValue(wb.Properties) as XmlNamespaceManager;
            var pathlist = wb.Properties.CustomPropertiesXml.SelectNodes("ctp:Properties/ctp:property/@name", ns);
            if (pathlist == null) return new string[0];
            var names = new string[pathlist.Count];
            int i = 0;
            foreach (XmlAttribute attr in pathlist)
            {
                names[i++] = attr.Value;
            }

            return names;
        }

        /// <summary>
        /// Append zero-width character. Excel string columns that have numeric values, cause Excel 
        /// to create a warning icon that text cell contains a number. We append a zero-width 
        /// space to force Excel to treat number as string. However, string.Trim() does not 
        /// recognize zero-width spaces so we have to remove it ourselves with TrimSp().
        /// </summary>
        /// <param name="s">String to check</param>
        /// <returns>String with appended space</returns>
        /// <remarks>
        /// Normally DisableCellWarnings() will work until one touches a cell containing
        /// a numeric string then the cell warning are re-enabled for that cell.
        /// </remarks>
        public static string AppendSp(this string s)
        {
            // See: https://en.wikipedia.org/wiki/Whitespace_character
            return IsNumeric(s) ? s + "\x200B" : s;
        }

        /// <summary>
        /// Trim zero-width character. Excel string columns that have numeric values, cause Excel 
        /// to create a warning icon that text cell contains a number. We append a zero-width 
        /// space to force Excel to treat number as string. However, string.Trim() does not 
        /// recognize zero-width spaces so we have to remove it ourselves with TrimSp().
        /// </summary>
        /// <param name="s">String to check</param>
        /// <returns>String with appended space</returns>
        public static string TrimSp(this string s)
        {
            if (s == null) return s;
            if (s.Length == 0) return s;
            return s[s.Length - 1] == '\x200B' ? s.Substring(0, s.Length - 1) : s;
        }

        private static bool IsNumeric(string s)
        {
            if (s == null) return false;
            if (s.Length == 0) return false;
            if (s.Length == 1 && s[0] == '-') return false;
            var i = 0;
            var hasDot = false;
            foreach (var c in s)
            {
                if (i++ == 0 && c == '-') continue;
                if ((c == '.' || c == ',') && !hasDot)
                {
                    hasDot = true;
                    continue;
                }

                if (c < '0' || c > '9') return false;
            }

            return true;
        }

        #region Color Theme Luminance Scaling for Worksheet Color Shades
        /// <summary>
        /// Change Luminance/Lightness/Shade of specified color. Alpha transparency is left unchanged.
        /// </summary>
        /// <param name="c">Source color value</param>
        /// <param name="newLuminance">New luminance value on the Excel scale of 0 to 255</param>
        /// <returns>New color shade</returns>
        private static Color SetLuminance(Color c, int newLuminance)
        {
            if (newLuminance < 0 || newLuminance > 255) new ArgumentOutOfRangeException(nameof(newLuminance), "Excel luminance must be between 0 and 255.");
            var hsl = GetHSLFromRGB(c.R, c.G, c.B);
            var rgb = GetRGBFromHSL(hsl[0], hsl[1], newLuminance);
            return Color.FromArgb(c.A, rgb[0], rgb[1], rgb[2]);
        }

        /// <summary>
        /// Get HSL values in the context of Excel numbering system (e.g. all values range from 0 to 255).
        /// See: https://exceloffthegrid.com/convert-color-codes/
        /// </summary>
        /// <param name="red">Red RGB value between 0 and 255</param>
        /// <param name="green">Green RGB value between 0 and 255</param>
        /// <param name="blue">Blue RGB value between 0 and 255</param>
        /// <returns>Int array of HSL values in Excel context</returns>
        private static int[] GetHSLFromRGB(int red, int green, int blue)
        {
            var r = red / 255.0;
            var g = green / 255.0;
            var b = blue / 255.0;
            var minRGB = Min(r, g, b);
            var maxRGB = Max(r, g, b);
            double H, S, L;

            L = (minRGB + maxRGB) / 2;
            if (minRGB == maxRGB) S = 0;
            else if (L < 0.5) S = (maxRGB - minRGB) / (maxRGB + minRGB);
            else S = (maxRGB - minRGB) / (2 - maxRGB - minRGB);

            if (S == 0) H = 0;
            else if (r > Max(g, b))
                H = (g - b) / (maxRGB - minRGB);
            else if (g > Max(r, b))
                H = 2 + ((b - r) / (maxRGB - minRGB));
            else
                H = 4 + ((r - g) / (maxRGB - minRGB));

            H = H * 60;
            if (H < 0) H = H + 360;

            // Excel scales the HSL double (0-360, 0-1, 0-1) values to integers (0-255, 0-255, 0-255)!
            return new[] { (int)Math.Round(H / 360 * 255), (int)Math.Round(S * 255), (int)Math.Round(L * 255) };
        }

        /// <summary>
        /// Get RGB values from the context of the Excel HLS numbering system (e.g. all values range from 0 to 255).
        /// See: https://exceloffthegrid.com/convert-color-codes/
        /// </summary>
        /// <param name="hue">Hue HSL value between 0 and 255</param>
        /// <param name="saturation">Saturation HSL value between 0 and 255</param>
        /// <param name="luminance">Luminance HSL value between 0 and 255</param>
        /// <returns>Int array of RGB values</returns>
        private static int[] GetRGBFromHSL(int hue, int saturation, int luminance)
        {
            double R, G, B;
            double temp1, temp2;
            double tempR, tempG, tempB;
            double H = hue / 255.0 * 360.0;
            double S = saturation / 255.0;
            double L = luminance / 255.0;

            if (saturation == 0) return new[] { (int)Math.Round(L * 255), (int)Math.Round(L * 255), (int)Math.Round(L * 255) };

            if (L < 0.5) temp1 = L * (1 + S);
            else temp1 = L + S - (L * S);

            temp2 = (2 * L) - temp1;
            H = H / 360;
            tempR = H + 0.333333;
            tempG = H;
            tempB = H - 0.333333;
            if (tempR < 0) tempR = tempR + 1;
            if (tempR > 1) tempR = tempR - 1;
            if (tempG < 0) tempG = tempG + 1;
            if (tempG > 1) tempG = tempG - 1;
            if (tempB < 0) tempB = tempB + 1;
            if (tempB > 1) tempB = tempB - 1;

            if (6 * tempR < 1) R = temp2 + ((temp1 - temp2) * 6 * tempR);
            else if (2 * tempR < 1) R = temp1;
            else if (3 * tempR < 2) R = temp2 + ((temp1 - temp2) * (0.666666 - tempR) * 6);
            else R = temp2;

            if (6 * tempG < 1) G = temp2 + ((temp1 - temp2) * 6 * tempG);
            else if (2 * tempG < 1) G = temp1;
            else if (3 * tempG < 2) G = temp2 + ((temp1 - temp2) * (0.666666 - tempG) * 6);
            else G = temp2;

            if (6 * tempB < 1) B = temp2 + ((temp1 - temp2) * 6 * tempB);
            else if (2 * tempB < 1) B = temp1;
            else if (3 * tempB < 2) B = temp2 + ((temp1 - temp2) * (0.666666 - tempB) * 6);
            else B = temp2;

            return new[] { (int)Math.Round(R * 255), (int)Math.Round(G * 255), (int)Math.Round(B * 255) };
        }

        private static T Min<T>(params T[] vals)
        {
            T v = vals[0];
            for (int i = 1; i < vals.Length; i++)
            {
                if (Comparer<T>.Default.Compare(vals[i], v) < 0) v = vals[i];
            }

            return v;
        }

        private static T Max<T>(params T[] vals)
        {
            T v = vals[0];
            for (int i = 1; i < vals.Length; i++)
            {
                if (Comparer<T>.Default.Compare(vals[i], v) > 0) v = vals[i];
            }

            return v;
        }
        #endregion // Color Luminance Scaling
    }
}
