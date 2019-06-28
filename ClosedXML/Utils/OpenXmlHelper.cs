// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using Drawing = System.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace ClosedXML.Utils
{
    internal static class OpenXmlHelper
    {
        #region Public Methods

        /// <summary>
        /// Convert color in ClosedXML representation to specified OpenXML type.
        /// </summary>
        /// <typeparam name="T">The descendant of <see cref="ColorType"/>.</typeparam>
        /// <param name="openXMLColor">The existing instance of ColorType.</param>
        /// <param name="xlColor">Color in ClosedXML format.</param>
        /// <param name="isDifferential">Flag specifiying that the color should be saved in
        /// differential format (affects the transparent color processing).</param>
        /// <returns>The original color in OpenXML format.</returns>
        public static T FromClosedXMLColor<T>(this ColorType openXMLColor, XLColor xlColor, bool isDifferential = false)
            where T : ColorType
        {
            FillFromClosedXMLColor(openXMLColor, xlColor, isDifferential);
            return (T)openXMLColor;
        }

        /// <summary>
        /// Convert color in ClosedXML representation to specified OpenXML type.
        /// </summary>
        /// <typeparam name="T">The descendant of <see cref="X14.ColorType"/>.</typeparam>
        /// <param name="openXMLColor">The existing instance of ColorType.</param>
        /// <param name="xlColor">Color in ClosedXML format.</param>
        /// <param name="isDifferential">Flag specifiying that the color should be saved in
        /// differential format (affects the transparent color processing).</param>
        /// <returns>The original color in OpenXML format.</returns>
        public static T FromClosedXMLColor<T>(this X14.ColorType openXMLColor, XLColor xlColor, bool isDifferential = false)
            where T : X14.ColorType
        {
            FillFromClosedXMLColor(openXMLColor, xlColor, isDifferential);
            return (T)openXMLColor;
        }

        public static BooleanValue GetBooleanValue(bool value, bool? defaultValue = null)
        {
            return (defaultValue.HasValue && value == defaultValue.Value) ? null : new BooleanValue(value);
        }

        public static bool GetBooleanValueAsBool(BooleanValue value, bool defaultValue)
        {
            return (value?.HasValue ?? false) ? value.Value : defaultValue;
        }

        /// <summary>
        /// Convert color in OpenXML representation to ClosedXML type.
        /// </summary>
        /// <param name="openXMLColor">Color in OpenXML format.</param>
        /// <param name="colorCache">The dictionary containing parsed colors to optimize performance.</param>
        /// <returns>The color in ClosedXML format.</returns>
        public static XLColor ToClosedXMLColor(this ColorType openXMLColor, IDictionary<string, Drawing.Color> colorCache = null)
        {
            return ConvertToClosedXMLColor(openXMLColor, colorCache);
        }

        /// <summary>
        /// Convert color in OpenXML representation to ClosedXML type.
        /// </summary>
        /// <param name="openXMLColor">Color in OpenXML format.</param>
        /// <param name="colorCache">The dictionary containing parsed colors to optimize performance.</param>
        /// <returns>The color in ClosedXML format.</returns>
        public static XLColor ToClosedXMLColor(this X14.ColorType openXMLColor, IDictionary<string, Drawing.Color> colorCache = null)
        {
            return ConvertToClosedXMLColor(openXMLColor, colorCache);
        }

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Here we perform the actual convertion from OpenXML color to ClosedXML color.
        /// </summary>
        /// <param name="openXMLColor">OpenXML color. Must be either <see cref="ColorType"/> or <see cref="X14.ColorType"/>. 
        /// Since these types do not implement a common interface we use dynamic.</param>
        /// <param name="colorCache">The dictionary containing parsed colors to optimize performance.</param>
        /// <returns>The color in ClosedXML format.</returns>
        private static XLColor ConvertToClosedXMLColor(dynamic openXMLColor, IDictionary<string, Drawing.Color> colorCache )
        {
            XLColor retVal = null;
            if (openXMLColor != null)
            {
                if (openXMLColor.Rgb != null)
                {
                    String htmlColor = "#" + openXMLColor.Rgb.Value;
                    Drawing.Color thisColor;
                    if (colorCache?.ContainsKey(htmlColor) ?? false)
                    {
                        thisColor = colorCache[htmlColor];
                    }
                    else
                    {
                        thisColor = ColorStringParser.ParseFromHtml(htmlColor);
                        colorCache?.Add(htmlColor, thisColor);
                    }

                    retVal = XLColor.FromColor(thisColor);
                }
                else if (openXMLColor.Indexed != null && openXMLColor.Indexed <= 64)
                    retVal = XLColor.FromIndex((Int32)openXMLColor.Indexed.Value);
                else if (openXMLColor.Theme != null)
                {
                    retVal = openXMLColor.Tint != null
                        ? XLColor.FromTheme((XLThemeColor)openXMLColor.Theme.Value, openXMLColor.Tint.Value)
                        : XLColor.FromTheme((XLThemeColor)openXMLColor.Theme.Value);
                }
            }
            return retVal ?? XLColor.NoColor;
        }

        /// <summary>
        /// Initialize properties of the existing instance of the color in OpenXML format basing on properties of the color 
        /// in ClosedXML format.
        /// </summary>
        /// <param name="openXMLColor">OpenXML color. Must be either <see cref="ColorType"/> or <see cref="X14.ColorType"/>. 
        /// Since these types do not implement a common interface we use dynamic.</param>
        /// <param name="xlColor">Color in ClosedXML format.</param>
        /// <param name="isDifferential">Flag specifiying that the color should be saved in
        /// differential format (affects the transparent color processing).</param>
        private static void FillFromClosedXMLColor(dynamic openXMLColor, XLColor xlColor, bool isDifferential)
        {
            if (openXMLColor == null)
                throw new ArgumentNullException(nameof(openXMLColor));

            if (xlColor == null)
                throw new ArgumentNullException(nameof(xlColor));

            switch (xlColor.ColorType)
            {
                case XLColorType.Color:
                    openXMLColor.Rgb = xlColor.Color.ToHex();
                    break;

                case XLColorType.Indexed:
                    // 64 is 'transparent' and should be ignored for differential formats
                    if (!isDifferential || xlColor.Indexed != 64)
                        openXMLColor.Indexed = (UInt32)xlColor.Indexed;
                    break;

                case XLColorType.Theme:
                    openXMLColor.Theme = (UInt32)xlColor.ThemeColor;

                    if (xlColor.ThemeTint != 0)
                        openXMLColor.Tint = xlColor.ThemeTint;
                    break;
            }
        }

        #endregion Private Methods
    }
}
