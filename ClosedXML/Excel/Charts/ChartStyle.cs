using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace ClosedXML.Excel.Charts
{
    public class ChartStyle
    {
        public static TextProperties SetTextProperties(OpenXmlCompositeElement axisElement, XLChart chartData)
        {
            int rotation = 5400000;
            var txPr = new TextProperties();
            if (chartData.ChartType == XLChartType.Radar || chartData.Rotated)
                rotation = 0;
            var bodyProp = new BodyProperties() { Rotation = axisElement is CategoryAxis ? rotation : 0, UseParagraphSpacing = true, VerticalOverflow = TextVerticalOverflowValues.Ellipsis, Vertical = TextVerticalValues.Horizontal, Wrap = TextWrappingValues.Square, Anchor = TextAnchoringTypeValues.Center, AnchorCenter = true };
            var spAutoFit = new ShapeAutoFit();
            bodyProp.Append(spAutoFit);
            var listSyle = new ListStyle();
            txPr.Append(bodyProp); txPr.Append(listSyle);

            var p = new Paragraph();
            var pPr = new ParagraphProperties();
            var defRProp = new DefaultRunProperties() { Bold = false, Italic = false, Underline = TextUnderlineValues.None, Strike = TextStrikeValues.NoStrike, Kerning = 1200, Spacing = 0, Baseline = 0 };
            var solidFill = new SolidFill();
            var schemeColor = new SchemeColor() { Val = SchemeColorValues.Text1 };
            var lumMod = new LuminanceModulation() { Val = 65000 };
            var lumOff = new LuminanceOffset() { Val = 35000 };
            schemeColor.Append(lumMod); schemeColor.Append(lumOff);
            solidFill.Append(schemeColor);
            defRProp.Append(solidFill);

            var latin2 = new LatinFont() { Typeface = "+mn-lt" };
            var ea2 = new EastAsianFont() { Typeface = "+mn-ea" };
            var cs2 = new ComplexScriptFont() { Typeface = "+mn-cs" };
            defRProp.Append(latin2); defRProp.Append(ea2); defRProp.Append(cs2);

            pPr.Append(defRProp);
            p.Append(pPr);

            txPr.Append(p);

            return txPr;
        }

        public static TextProperties SetTextProperties()
        {
            var txPr = new TextProperties();
            var bodyProp = new BodyProperties() { Rotation = 0, UseParagraphSpacing = true, VerticalOverflow = TextVerticalOverflowValues.Ellipsis, Vertical = TextVerticalValues.Horizontal, Wrap = TextWrappingValues.Square, Anchor = TextAnchoringTypeValues.Center, AnchorCenter = true };
            var spAutoFit = new ShapeAutoFit();
            bodyProp.Append(spAutoFit);
            var listSyle = new ListStyle();
            txPr.Append(bodyProp); txPr.Append(listSyle);

            var p = new Paragraph();
            var pPr = new ParagraphProperties();
            var defRProp = new DefaultRunProperties() { Bold = false, Italic = false, Underline = TextUnderlineValues.None, Strike = TextStrikeValues.NoStrike, Kerning = 1200, Spacing = 0, Baseline = 0/*, FontSize = 1400*/ };
            var solidFill = new SolidFill();
            var schemeColor = new SchemeColor() { Val = SchemeColorValues.Text1 };
            var lumMod = new LuminanceModulation() { Val = 65000 };
            var lumOff = new LuminanceOffset() { Val = 35000 };
            schemeColor.Append(lumMod); schemeColor.Append(lumOff);
            solidFill.Append(schemeColor);
            defRProp.Append(solidFill);

            var latin2 = new LatinFont() { Typeface = "+mn-lt" };
            var ea2 = new EastAsianFont() { Typeface = "+mn-ea" };
            var cs2 = new ComplexScriptFont() { Typeface = "+mn-cs" };
            defRProp.Append(latin2); defRProp.Append(ea2); defRProp.Append(cs2);

            pPr.Append(defRProp);
            p.Append(pPr);

            txPr.Append(p);

            return txPr;
        }

        public static ChartShapeProperties SetShapeProperties()
        {
            var spPr = new ChartShapeProperties();
            var ln = new DocumentFormat.OpenXml.Drawing.Outline() { Width = 9525, CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Flat, CompoundLineType = DocumentFormat.OpenXml.Drawing.CompoundLineValues.Single, Alignment = DocumentFormat.OpenXml.Drawing.PenAlignmentValues.Center };
            var solidFill = new DocumentFormat.OpenXml.Drawing.SolidFill();
            var schemeClr = new DocumentFormat.OpenXml.Drawing.SchemeColor() { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.Text1 };
            var lumMod = new DocumentFormat.OpenXml.Drawing.LuminanceModulation() { Val = 15000 };
            var lumOff = new DocumentFormat.OpenXml.Drawing.LuminanceOffset() { Val = 85000 };
            schemeClr.Append(lumMod); schemeClr.Append(lumOff);
            solidFill.Append(schemeClr);
            ln.Append(solidFill);
            var effectLst = new DocumentFormat.OpenXml.Drawing.EffectList();
            spPr.Append(ln); spPr.Append(effectLst);

            return spPr;
        }

        public static ChartShapeProperties GetNoFillProperties(bool outline)
        {
            ChartShapeProperties chartShapeProperties = new ChartShapeProperties();
            if (outline)
            {
                var ol = new Outline();
                ol.Append(new NoFill());
                chartShapeProperties.Append(ol);
            }
            else
            {
                chartShapeProperties.Append(new NoFill());
            }
            return chartShapeProperties;
        }

        public static Marker AddMarker(ChartSeriesData seriesData)
        {
            Marker marker = new Marker();

            Symbol symbol = new Symbol { Val = MarkerStyleValues.Square };
            marker.Append(symbol);

            Size size = new Size() { Val = 5 };
            marker.Append(size);

            return marker;
        }

        public static Marker AddNoMarker()
        {
            Marker marker = new Marker { Symbol = new Symbol { Val = MarkerStyleValues.None } };
            return marker;
        }
    }
}
