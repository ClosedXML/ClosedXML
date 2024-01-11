// Keep this file CodeMaid organised and cleaned
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
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
        /// <param name="isDifferential">Flag specifying that the color should be saved in
        /// differential format (affects the transparent color processing).</param>
        /// <returns>The original color in OpenXML format.</returns>
        public static T FromClosedXMLColor<T>(this ColorType openXMLColor, XLColor xlColor, bool isDifferential = false)
            where T : ColorType
        {
            var adapter = new ColorTypeAdapter(openXMLColor);
            FillFromClosedXMLColor(adapter, xlColor, isDifferential);
            return (T)adapter.ColorType;
        }

        /// <summary>
        /// Convert color in ClosedXML representation to specified OpenXML type.
        /// </summary>
        /// <typeparam name="T">The descendant of <see cref="X14.ColorType"/>.</typeparam>
        /// <param name="openXMLColor">The existing instance of ColorType.</param>
        /// <param name="xlColor">Color in ClosedXML format.</param>
        /// <param name="isDifferential">Flag specifying that the color should be saved in
        /// differential format (affects the transparent color processing).</param>
        /// <returns>The original color in OpenXML format.</returns>
        public static T FromClosedXMLColor<T>(this X14.ColorType openXMLColor, XLColor xlColor, bool isDifferential = false)
            where T : X14.ColorType
        {
            var adapter = new X14ColorTypeAdapter(openXMLColor);
            FillFromClosedXMLColor(adapter, xlColor, isDifferential);
            return (T)adapter.ColorType;
        }

        public static BooleanValue? GetBooleanValue(bool value, bool? defaultValue = null)
        {
            return (defaultValue.HasValue && value == defaultValue.Value) ? null : new BooleanValue(value);
        }

        public static bool GetBooleanValueAsBool(BooleanValue? value, bool defaultValue)
        {
            return (value?.HasValue ?? false) ? value.Value : defaultValue;
        }

        /// <summary>
        /// Convert color in OpenXML representation to ClosedXML type.
        /// </summary>
        /// <param name="openXMLColor">Color in OpenXML format.</param>
        /// <returns>The color in ClosedXML format.</returns>
        public static XLColor ToClosedXMLColor(this ColorType openXMLColor)
        {
            return ConvertToClosedXMLColor(new ColorTypeAdapter(openXMLColor));
        }

        /// <summary>
        /// Convert color in OpenXML representation to ClosedXML type.
        /// </summary>
        /// <param name="openXMLColor">Color in OpenXML format.</param>
        /// <returns>The color in ClosedXML format.</returns>
        public static XLColor ToClosedXMLColor(this X14.ColorType openXMLColor)
        {
            return ConvertToClosedXMLColor(new X14ColorTypeAdapter(openXMLColor));
        }

#nullable disable

        internal static void LoadNumberFormat(NumberingFormat nfSource, IXLNumberFormat nf)
        {
            if (nfSource == null) return;

            if (nfSource.NumberFormatId != null && nfSource.NumberFormatId.Value < XLConstants.NumberOfBuiltInStyles)
                nf.NumberFormatId = (Int32)nfSource.NumberFormatId.Value;
            else if (nfSource.FormatCode != null)
                nf.Format = nfSource.FormatCode.Value;
        }

        internal static void LoadBorder(Border borderSource, IXLBorder border)
        {
            if (borderSource == null) return;

            LoadBorderValues(borderSource.DiagonalBorder, border.SetDiagonalBorder, border.SetDiagonalBorderColor);

            if (borderSource.DiagonalUp != null)
                border.DiagonalUp = borderSource.DiagonalUp.Value;
            if (borderSource.DiagonalDown != null)
                border.DiagonalDown = borderSource.DiagonalDown.Value;

            LoadBorderValues(borderSource.LeftBorder, border.SetLeftBorder, border.SetLeftBorderColor);
            LoadBorderValues(borderSource.RightBorder, border.SetRightBorder, border.SetRightBorderColor);
            LoadBorderValues(borderSource.TopBorder, border.SetTopBorder, border.SetTopBorderColor);
            LoadBorderValues(borderSource.BottomBorder, border.SetBottomBorder, border.SetBottomBorderColor);
        }

        private static void LoadBorderValues(BorderPropertiesType source, Func<XLBorderStyleValues, IXLStyle> setBorder, Func<XLColor, IXLStyle> setColor)
        {
            if (source != null)
            {
                if (source.Style != null)
                    setBorder(source.Style.Value.ToClosedXml());
                if (source.Color != null)
                    setColor(source.Color.ToClosedXMLColor());
            }
        }

        // Differential fills store the patterns differently than other fills
        // Actually differential fills make more sense. bg is bg and fg is fg
        // 'Other' fills store the bg color in the fg field when pattern type is solid
        internal static void LoadFill(Fill openXMLFill, IXLFill closedXMLFill, Boolean differentialFillFormat)
        {
            if (openXMLFill == null || openXMLFill.PatternFill == null) return;

            if (openXMLFill.PatternFill.PatternType != null)
                closedXMLFill.PatternType = openXMLFill.PatternFill.PatternType.Value.ToClosedXml();
            else
                closedXMLFill.PatternType = XLFillPatternValues.Solid;

            switch (closedXMLFill.PatternType)
            {
                case XLFillPatternValues.None:
                    break;

                case XLFillPatternValues.Solid:
                    if (differentialFillFormat)
                    {
                        if (openXMLFill.PatternFill.BackgroundColor != null)
                            closedXMLFill.BackgroundColor = openXMLFill.PatternFill.BackgroundColor.ToClosedXMLColor();
                        else
                            closedXMLFill.BackgroundColor = XLColor.FromIndex(64);
                    }
                    else
                    {
                        // yes, source is foreground!
                        if (openXMLFill.PatternFill.ForegroundColor != null)
                            closedXMLFill.BackgroundColor = openXMLFill.PatternFill.ForegroundColor.ToClosedXMLColor();
                        else
                            closedXMLFill.BackgroundColor = XLColor.FromIndex(64);
                    }
                    break;

                default:
                    if (openXMLFill.PatternFill.ForegroundColor != null)
                        closedXMLFill.PatternColor = openXMLFill.PatternFill.ForegroundColor.ToClosedXMLColor();

                    if (openXMLFill.PatternFill.BackgroundColor != null)
                        closedXMLFill.BackgroundColor = openXMLFill.PatternFill.BackgroundColor.ToClosedXMLColor();
                    else
                        closedXMLFill.BackgroundColor = XLColor.FromIndex(64);
                    break;
            }
        }

        internal static void LoadFont(OpenXmlElement fontSource, IXLFontBase fontBase)
        {
            if (fontSource == null) return;

            fontBase.Bold = GetBoolean(fontSource.Elements<Bold>().FirstOrDefault());
            var fontColor = fontSource.Elements<DocumentFormat.OpenXml.Spreadsheet.Color>().FirstOrDefault();
            if (fontColor != null)
                fontBase.FontColor = fontColor.ToClosedXMLColor();

            var fontFamilyNumbering =
                fontSource.Elements<DocumentFormat.OpenXml.Spreadsheet.FontFamily>().FirstOrDefault();
            if (fontFamilyNumbering != null && fontFamilyNumbering.Val != null)
                fontBase.FontFamilyNumbering =
                    (XLFontFamilyNumberingValues)Int32.Parse(fontFamilyNumbering.Val.ToString());
            var runFont = fontSource.Elements<RunFont>().FirstOrDefault();
            if (runFont != null)
            {
                if (runFont.Val != null)
                    fontBase.FontName = runFont.Val;
            }
            var fontSize = fontSource.Elements<FontSize>().FirstOrDefault();
            if (fontSize != null)
            {
                if ((fontSize).Val != null)
                    fontBase.FontSize = (fontSize).Val;
            }

            fontBase.Italic = GetBoolean(fontSource.Elements<Italic>().FirstOrDefault());
            fontBase.Shadow = GetBoolean(fontSource.Elements<Shadow>().FirstOrDefault());
            fontBase.Strikethrough = GetBoolean(fontSource.Elements<Strike>().FirstOrDefault());

            var underline = fontSource.Elements<Underline>().FirstOrDefault();
            if (underline != null)
            {
                fontBase.Underline = underline.Val != null ? underline.Val.Value.ToClosedXml() : XLFontUnderlineValues.Single;
            }

            var verticalTextAlignment = fontSource.Elements<VerticalTextAlignment>().FirstOrDefault();
            if (verticalTextAlignment is not null)
            {
                fontBase.VerticalAlignment = verticalTextAlignment.Val is not null ? verticalTextAlignment.Val.Value.ToClosedXml() : XLFontVerticalTextAlignmentValues.Baseline;
            }

            var fontScheme = fontSource.Elements<FontScheme>().FirstOrDefault();
            if (fontScheme is not null)
            {
                fontBase.FontScheme = fontScheme.Val is not null ? fontScheme.Val.Value.ToClosedXml() : XLFontScheme.None;
            }
        }

        internal static Boolean GetBoolean(BooleanPropertyType property)
        {
            if (property != null)
            {
                if (property.Val != null)
                    return property.Val;
                return true;
            }

            return false;
        }

#nullable enable

        #endregion Public Methods

        #region Private Methods

        /// <summary>
        /// Here we perform the actual conversion from OpenXML color to ClosedXML color.
        /// </summary>
        /// <param name="openXMLColor">OpenXML color. Must be either <see cref="ColorType"/> or <see cref="X14.ColorType"/>.
        /// Since these types do not implement a common interface we use dynamic.</param>
        /// <returns>The color in ClosedXML format.</returns>
        private static XLColor ConvertToClosedXMLColor(IColorTypeAdapter openXMLColor)
        {
            XLColor? retVal = null;
            if (openXMLColor.Rgb?.Value is not null)
            {
                var thisColor = ColorStringParser.ParseFromArgb(openXMLColor.Rgb.Value.AsSpan());
                retVal = XLColor.FromColor(thisColor);
            }
            else if (openXMLColor.Indexed is not null && openXMLColor.Indexed <= 64)
                retVal = XLColor.FromIndex((Int32)openXMLColor.Indexed.Value);
            else if (openXMLColor.Theme is not null)
            {
                retVal = openXMLColor.Tint is not null
                    ? XLColor.FromTheme((XLThemeColor)openXMLColor.Theme.Value, openXMLColor.Tint.Value)
                    : XLColor.FromTheme((XLThemeColor)openXMLColor.Theme.Value);
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
        /// <param name="isDifferential">Flag specifying that the color should be saved in
        /// differential format (affects the transparent color processing).</param>
        private static void FillFromClosedXMLColor(IColorTypeAdapter openXMLColor, XLColor xlColor, bool isDifferential)
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

        internal static int GetClosedXmlTextRotation(Alignment alignment)
        {
            if (alignment.TextRotation is null)
                return 0;

            var textRotation = (int)alignment.TextRotation.Value;
            return textRotation switch
            {
                255 => 255,
                > 90 => 90 - textRotation,
                _ => textRotation
            };
        }

        #endregion Private Methods
    }
}
