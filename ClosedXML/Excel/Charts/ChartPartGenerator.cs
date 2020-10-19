using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using System;

namespace ClosedXML.Excel.Charts
{
    public static class ChartPartGenerator
    {
        private static void SetSafeLabelPosition(ChartSeriesData seriesData, DataLabels label)
        {
            var pos = seriesData.Label.Position;
            if (seriesData.SeriesType == ChartSeriesType.Area)
                return;

            if (seriesData.SeriesType == ChartSeriesType.Bar || seriesData.SeriesType == ChartSeriesType.Column || seriesData.SeriesType == ChartSeriesType.Column100Percent || seriesData.SeriesType == ChartSeriesType.Pie)
                if (pos == DataLabelPositionValues.Center || pos == DataLabelPositionValues.InsideEnd || pos == DataLabelPositionValues.OutsideEnd || pos == DataLabelPositionValues.InsideBase)
                    label.Append(new DataLabelPosition { Val = pos });

            if (seriesData.SeriesType == ChartSeriesType.Line || seriesData.SeriesType == ChartSeriesType.Scatter)
                if (pos == DataLabelPositionValues.Center || pos == DataLabelPositionValues.Left || pos == DataLabelPositionValues.Right || pos == DataLabelPositionValues.Top || pos == DataLabelPositionValues.Bottom)
                    label.Append(new DataLabelPosition { Val = pos });
        }

        public static DataLabels GenerateReferencedDataLables(ChartSeriesData seriesData)
        {
            if (!seriesData.Label.IsEnabled)
                return null;

            switch (seriesData.Label.LabelType)
            {
                case LabelType.CategoryValueLabel:
                    DataLabels categoryLabels = new DataLabels();
                    categoryLabels.Append(new NumberingFormat { FormatCode = "General", SourceLinked = false });
                    if (seriesData.Label.Position != DataLabelPositionValues.BestFit)
                        SetSafeLabelPosition(seriesData, categoryLabels);
                    categoryLabels.Append(new ShowLegendKey { Val = false });
                    categoryLabels.Append(new ShowValue { Val = false });
                    categoryLabels.Append(new ShowCategoryName { Val = true });
                    categoryLabels.Append(new ShowSeriesName { Val = false });
                    categoryLabels.Append(new ShowPercent { Val = false });
                    categoryLabels.Append(new ShowBubbleSize { Val = false });
                    categoryLabels.Append(new ShowLeaderLines { Val = false });
                    categoryLabels.Append(ChartStyle.SetTextProperties());
                    return categoryLabels;
                case LabelType.PercentValuesOutsideEnd:
                    DataLabels dataLabels = new DataLabels();
                    dataLabels.Append(new NumberingFormat { FormatCode = "General", SourceLinked = false });
                    if (seriesData.Label.Position != DataLabelPositionValues.BestFit)
                    {
                        seriesData.Label.Position = DataLabelPositionValues.OutsideEnd;
                        SetSafeLabelPosition(seriesData, dataLabels);
                    }
                    dataLabels.Append(new ShowLegendKey { Val = false });
                    dataLabels.Append(new ShowValue { Val = true });
                    dataLabels.Append(new ShowCategoryName { Val = false });
                    dataLabels.Append(new ShowSeriesName { Val = false });
                    dataLabels.Append(new ShowPercent { Val = false });
                    dataLabels.Append(new ShowBubbleSize { Val = false });
                    dataLabels.Append(new ShowLeaderLines { Val = false });
                    dataLabels.Append(ChartStyle.SetTextProperties());
                    return dataLabels;
                case LabelType.SingleElementLabels:
                    return GenerateDataLabels(seriesData);
                case LabelType.RegularLabel:
                    DataLabels regularLabels = new DataLabels();
                    regularLabels.Append(new NumberingFormat { FormatCode = "General", SourceLinked = false });
                    if (seriesData.Label.Position != DataLabelPositionValues.BestFit)
                        SetSafeLabelPosition(seriesData, regularLabels);
                    regularLabels.Append(new ShowLegendKey { Val = false });
                    regularLabels.Append(new ShowValue { Val = true });
                    regularLabels.Append(new ShowCategoryName { Val = false });
                    regularLabels.Append(new ShowSeriesName { Val = false });
                    regularLabels.Append(new ShowPercent { Val = false });
                    regularLabels.Append(new ShowBubbleSize { Val = false });
                    regularLabels.Append(new ShowLeaderLines { Val = false });
                    regularLabels.Append(ChartStyle.SetTextProperties());
                    return regularLabels;
                default:
                    return null;
            }
        }

        public static DataLabels GenerateDataLabels(ChartSeriesData seriesData)
        {
            if (!seriesData.Label.IsEnabled)
                return null;

            DataLabels dataLabels = new DataLabels();

            for (int i = 0; i < seriesData.Names.Length; i++)
            {
                DataLabel dataLabel = new DataLabel();
                var index = new Index() { Val = (uint)i };
                var tx = new ChartText();
                var rich = new RichText();
                rich.Append(new BodyProperties());
                rich.Append(new ListStyle());

                var p = new Paragraph();
                var r = new Run();
                r.Append(new RunProperties() { Language = "de-DE" });
                r.Append(new Text(seriesData.Names[i]));
                p.Append(r);

                rich.Append(p);
                tx.Append(rich);
                dataLabel.Append(index);
                dataLabel.Append(tx);

                if (seriesData.Label.Position != DataLabelPositionValues.BestFit)
                {
                    dataLabels.Append(new DataLabelPosition { Val = seriesData.Label.Position });
                }
                dataLabel.Append(new ShowLegendKey { Val = false });
                dataLabel.Append(new ShowValue { Val = false });
                dataLabel.Append(new ShowCategoryName { Val = false });
                dataLabel.Append(new ShowSeriesName { Val = false });
                dataLabel.Append(new ShowPercent { Val = false });
                dataLabel.Append(new ShowBubbleSize { Val = false });

                dataLabels.Append(dataLabel);
            }

            dataLabels.Append(ChartStyle.SetTextProperties());
            if (seriesData.Label.Position != DataLabelPositionValues.BestFit)
            {
                dataLabels.Append(new DataLabelPosition { Val = seriesData.Label.Position });
            }
            dataLabels.Append(new ShowLegendKey { Val = false });
            dataLabels.Append(new ShowValue { Val = false });
            dataLabels.Append(new ShowCategoryName { Val = false });
            dataLabels.Append(new ShowSeriesName { Val = false });
            dataLabels.Append(new ShowPercent { Val = false });
            dataLabels.Append(new ShowBubbleSize { Val = false });
            dataLabels.Append(new ShowLeaderLines { Val = false });

            return dataLabels;

        }

        public static RichText GenerateRichTextSingleElement(String text, String language = null, String hexColor = null)
        {
            if (language == null)
                language = "de-DE";

            RichText richText = new RichText();
            BodyProperties bodyProperties = new BodyProperties();
            ListStyle listStyle = new ListStyle();

            Paragraph paragraph = new Paragraph();

            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(new DefaultRunProperties());

            Run run = new Run();

            RunProperties runProperties = new RunProperties { Language = language, Dirty = false };
            runProperties.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(String.Empty, @"smtClean", String.Empty, @"0"));

            if (hexColor != null)
            {
                SolidFill solidFill = new SolidFill();
                RgbColorModelHex rgbColorModelHex = new RgbColorModelHex { Val = hexColor };

                solidFill.Append(rgbColorModelHex);
                runProperties.Append(solidFill);
            }

            Text textObject = new Text { Text = text };

            run.Append(runProperties);
            run.Append(textObject);
            EndParagraphRunProperties endParagraphRunProperties = new EndParagraphRunProperties { Language = language, Dirty = false };

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);
            paragraph.Append(endParagraphRunProperties);

            richText.Append(bodyProperties);
            richText.Append(listStyle);
            richText.Append(paragraph);

            return richText;
        }

        public static SeriesText GenerateReferencedSeriesText(ChartSeriesData seriesData)
        {
            SeriesText seriesText = new SeriesText();

            StringReference seriesStringReference = new StringReference();
            Formula seriesFormula = new Formula { Text = seriesData.SeriesName.Reference };

            StringCache seriesStringCache = new StringCache();
            PointCount seriesPointCount = new PointCount { Val = 1U };

            StringPoint seriesStringPoint = new StringPoint { Index = 0U };
            NumericValue seriesNumericValue = new NumericValue { Text = seriesData.SeriesName.Value };

            seriesStringPoint.Append(seriesNumericValue);

            seriesStringCache.Append(seriesPointCount);
            seriesStringCache.Append(seriesStringPoint);

            seriesStringReference.Append(seriesFormula);
            seriesStringReference.Append(seriesStringCache);

            seriesText.Append(seriesStringReference);

            return seriesText;
        }

        public static SeriesText GenerateSeriesText(ChartSeriesData seriesData)
        {
            // Series text
            SeriesText seriesText = new SeriesText();
            NumericValue numericValue = new NumericValue() { Text = seriesData.SeriesName.Value };
            seriesText.Append(numericValue);

            return seriesText;
        }

        public static CategoryAxisData GenerateReferencedCategoryAxisData(ChartSeriesData seriesData)
        {
            CategoryAxisData categoryAxisData = new CategoryAxisData();

            StringReference stringReference = new StringReference();

            Formula formula = new Formula { Text = seriesData.Category.Reference };
            stringReference.Append(formula);

            StringCache stringCache = new StringCache();
            PointCount pointCount = new PointCount { Val = (UInt32)seriesData.Category.Values.Length };
            stringCache.Append(pointCount);

            UInt32 i = 0;
            foreach (var categoryText in seriesData.Category.Values)
            {
                StringPoint stringPoint = new StringPoint { Index = i };
                NumericValue numericValue = new NumericValue { Text = categoryText };

                stringPoint.Append(numericValue);
                stringCache.Append(stringPoint);

                i++;
            }

            stringReference.Append(stringCache);
            categoryAxisData.Append(stringReference);

            return categoryAxisData;
        }

        public static CategoryAxisData GenerateCategoryAxisData(ChartSeriesData seriesData)
        {
            CategoryAxisData categoryAxisData = new CategoryAxisData();
            categoryAxisData.Append(GetStringLiteral(seriesData));

            return categoryAxisData;
        }

        public static StringLiteral GetStringLiteral(ChartSeriesData seriesData)
        {
            StringLiteral stringLiteral = new StringLiteral();
            PointCount pointCount = new PointCount() { Val = (UInt32)seriesData.Category.Values.Length };
            stringLiteral.Append(pointCount);

            UInt32 i = 0;
            foreach (var categoryText in seriesData.Category.Values)
            {
                StringPoint stringPoint = new StringPoint { Index = i };
                NumericValue numericValue = new NumericValue { Text = categoryText };

                stringPoint.Append(numericValue);
                stringLiteral.Append(stringPoint);

                i++;
            }
            return stringLiteral;
        }

        public static Values GenerateValues(ChartSeriesData seriesData)
        {
            Values values = new Values();

            values.Append(GetNumberLiteral(seriesData.Values.Values));

            return values;
        }

        public static Values GenerateReferencedValues(ChartSeriesData seriesData)
        {
            Values values = new Values();

            values.Append(GetNumberReference(seriesData.Values.Reference, seriesData.Values.Values));

            return values;
        }

        public static XValues GenerateXValues(ChartSeriesData seriesData)
        {
            XValues values = new XValues();
            values.Append(GetStringLiteral(seriesData));

            return values;
        }

        public static YValues GenerateYValues(ChartSeriesData seriesData)
        {
            YValues values = new YValues();
            values.Append(GetNumberLiteral(seriesData.Values.Values));

            return values;
        }

        public static XValues GenerateReferencedXValues(ChartSeriesData seriesData)
        {
            var values = new XValues();

            NumberReference numberReference = GenerateNumberReference(seriesData);
            values.Append(numberReference);

            return values;
        }

        public static YValues GenerateReferencedYValues(ChartSeriesData seriesData)
        {
            var values = new YValues();

            NumberReference numberReference = GenerateNumberReference(seriesData);
            values.Append(numberReference);

            return values;
        }

        public static NumberReference GenerateNumberReference(ChartSeriesData seriesData)
        {
            NumberReference numberReference = new NumberReference();
            Formula formula = new Formula { Text = seriesData.Values.Reference };
            numberReference.Append(formula);

            NumberingCache numberingCache = new NumberingCache();
            FormatCode formatCode = new FormatCode { Text = "General" };
            numberingCache.Append(formatCode);

            PointCount pointCount = new PointCount { Val = (UInt32)seriesData.Values.Values.Length };
            numberingCache.Append(pointCount);

            UInt32 i = 0;
            foreach (var valueText in seriesData.Values.Values)
            {
                NumericPoint numericPoint = new NumericPoint { Index = i };
                NumericValue numericValue = new NumericValue { Text = valueText };

                numericPoint.Append(numericValue);
                numberingCache.Append(numericPoint);

                i++;
            }

            numberReference.Append(numberingCache);
            return numberReference;
        }

        private static NumberReference GetNumberReference(String reference, String[] values)
        {
            NumberReference numberReference = new NumberReference();
            Formula formula = new Formula { Text = reference };
            numberReference.Append(formula);

            NumberingCache numberingCache = new NumberingCache();
            FormatCode formatCode = new FormatCode { Text = "General" };
            numberingCache.Append(formatCode);

            PointCount pointCount = new PointCount { Val = (UInt32)values.Length };
            numberingCache.Append(pointCount);

            UInt32 i = 0;
            foreach (var valueText in values)
            {
                NumericPoint numericPoint = new NumericPoint { Index = i };
                NumericValue numericValue = new NumericValue { Text = valueText };

                numericPoint.Append(numericValue);
                numberingCache.Append(numericPoint);

                i++;
            }

            numberReference.Append(numberingCache);
            return numberReference;
        }

        private static NumberLiteral GetNumberLiteral(String[] values)
        {
            NumberLiteral numberLiteral = new NumberLiteral();

            FormatCode formatCode = new FormatCode("General");
            numberLiteral.Append(formatCode);

            PointCount pointCount = new PointCount { Val = (UInt32)values.Length };
            numberLiteral.Append(pointCount);

            UInt32 i = 0;
            foreach (var valueText in values)
            {
                NumericPoint numericPoint = new NumericPoint { Index = i };
                NumericValue numericValue = new NumericValue { Text = valueText };

                numericPoint.Append(numericValue);
                numberLiteral.Append(numericPoint);

                i++;
            }
            return numberLiteral;
        }
    }
}
