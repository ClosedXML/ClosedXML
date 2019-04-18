using System.IO;
using System.Linq;
using ClosedXML.Excel;
using NUnit.Framework;

namespace ClosedXML_Tests
{
    [TestFixture]
    public class PivotTableStylesTests
    {
        [Test]
        public void PivotTableStyleFormatsTest()
        {
            using (var ms = new MemoryStream())
            {
                using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\PivotTables\PivotTables.xlsx")))
                using (var wbSource = new XLWorkbook(stream))
                using (var wbDestination = new XLWorkbook())
                {
                    var ws = wbSource.Worksheet("PastrySalesData");
                    wbDestination.AddWorksheet(ws);
                    ws = wbDestination.Worksheet("PastrySalesData");

                    var table = ws.Table("PastrySalesData");
                    var ptSheet = wbDestination.Worksheets.Add("PivotTableStyleFormats");
                    var pt = ptSheet.PivotTables.Add("pvtStyleFormats", ptSheet.Cell(1, 1), table);
                    pt.Layout = XLPivotLayout.Tabular;

                    pt.SetSubtotals(XLPivotSubtotals.AtBottom);

                    var monthPivotField = pt.ColumnLabels.Add("Month");

                    var namePivotField = pt.RowLabels.Add("Name")
                        .SetSubtotalCaption("Test caption")
                        .SetCustomName("Test name")
                        .AddSubtotal(XLSubtotalFunction.Sum);

                    ptSheet.SetTabActive();

                    var numberOfOrdersPivotValue = pt.Values.Add("NumberOfOrders")
                        .SetSummaryFormula(XLPivotSummary.Sum);

                    var qualityPivotValue = pt.Values.Add("Quality").SetSummaryFormula(XLPivotSummary.Sum);

                    pt.StyleFormats.RowGrandTotalFormats.ForElement(XLPivotStyleFormatElement.All).Style.Font.FontColor = XLColor.VenetianRed;

                    namePivotField.StyleFormats.Subtotal.AddValuesFormat().Style.Fill.BackgroundColor = XLColor.Blue;
                    monthPivotField.StyleFormats.Label.Style.Fill.BackgroundColor = XLColor.Amber;
                    monthPivotField.StyleFormats.Header.Style.Font.FontColor = XLColor.Yellow;
                    namePivotField.StyleFormats.AddValuesFormat()
                        .AndWith(monthPivotField, v => v.ToString() == "May")
                        .ForValueField(numberOfOrdersPivotValue)
                        .Style.Font.FontColor = XLColor.Green;

                    wbDestination.SaveAs(ms);
                }

                ms.Seek(0, SeekOrigin.Begin);

                using (var wb = new XLWorkbook(ms))
                {
                    var ws = wb.Worksheet("PivotTableStyleFormats");
                    var pt = ws.PivotTable("pvtStyleFormats").CastTo<XLPivotTable>();

                    Assert.AreEqual(0, pt.StyleFormats.ColumnGrandTotalFormats.Count());

                    Assert.NotNull(pt.StyleFormats.RowGrandTotalFormats);
                    Assert.AreEqual(1, pt.StyleFormats.RowGrandTotalFormats.Count());
                    Assert.AreEqual(XLPivotStyleFormatElement.All, pt.StyleFormats.RowGrandTotalFormats.First().AppliesTo);
                    Assert.AreEqual(XLColor.VenetianRed, pt.StyleFormats.RowGrandTotalFormats.ForElement(XLPivotStyleFormatElement.All).Style.Font.FontColor);

                    var namePivotField = pt.RowLabels.Get("Name");
                    var monthPivotField = pt.ColumnLabels.Get("Month");
                    var numberOfOrdersPivotValue = pt.Values.Get("NumberOfOrders");

                    Assert.AreEqual(XLStyle.Default, namePivotField.StyleFormats.Label.Style);
                    Assert.AreEqual(XLColor.Blue, namePivotField.StyleFormats.Subtotal.DataValuesFormats.First().Style.Fill.BackgroundColor);

                    Assert.AreEqual(XLColor.Amber, monthPivotField.StyleFormats.Label.Style.Fill.BackgroundColor);
                    Assert.AreEqual(XLColor.Yellow, monthPivotField.StyleFormats.Header.Style.Font.FontColor);

                    var nameDataValuesFormat = namePivotField.StyleFormats.DataValuesFormats.First() as XLPivotValueStyleFormat;
                    Assert.AreEqual(2, nameDataValuesFormat.FieldReferences.Count());

                    Assert.AreEqual(monthPivotField, nameDataValuesFormat.FieldReferences.First().CastTo<PivotLabelFieldReference>().PivotField);

                    Assert.AreEqual(numberOfOrdersPivotValue.CustomName, nameDataValuesFormat.FieldReferences.Last().CastTo<PivotValueFieldReference>().Value);

                    wb.Save();
                }
            }
        }

        [Test]
        public void DataOutlineStylesTest()
        {
            TestHelper.CreateAndCompare(() =>
            {
                var wb = new XLWorkbook();

                var sheet = wb.Worksheets.Add("PastrySalesData");
                var table = sheet.Cell(1, 1).InsertTable(Pastry.WithMonthSet, "PastrySalesData", true);

                for (var i = 1; i <= 5; i++)
                {
                    // Add a new sheet for our pivot table
                    var ptSheet = wb.Worksheets.Add("pvt" + i);

                    // Create the pivot table, using the data from the "PastrySalesData" table
                    var pt = ptSheet.PivotTables.Add("pvt" + i, ptSheet.Cell(1, 1), table);
                    pt.ClassicPivotTableLayout = true;

                    IXLPivotField nameField = null;
                    if (i == 1 || i == 4 || i == 5)
                        nameField = pt.ColumnLabels.Add("Name");
                    else if (i == 2 || i == 3)
                    {
                        pt.RowLabels.Add("Code");
                        nameField = pt.RowLabels.Add("Name");
                    }

                    IXLPivotField secondField = null;
                    if (i == 1)
                        secondField = pt.RowLabels.Add("Month");
                    else if (i == 4 || i == 3)
                        secondField = pt.ColumnLabels.Add("Month");
                    else if (i == 2 || i == 5)
                        secondField = pt.RowLabels.Add("BakeDate");

                    // The values in our table will come from the "NumberOfOrders" and "Quality" fields
                    // The default calculation setting is a total of each row/column
                    var ordersField = pt.Values.Add("NumberOfOrders");
                    var qtyField = pt.Values.Add("Quality");
                    qtyField.SummaryFormula = XLPivotSummary.Sum;

                    // formatting of labels
                    nameField.StyleFormats.Label.Style.Fill.BackgroundColor = XLColor.YellowMunsell;

                    // formatting of value fields
                    var ordersFormat = secondField.StyleFormats.AddValuesFormat().ForValueField(ordersField);
                    ordersFormat.Style.Fill.BackgroundColor = XLColor.Apricot;
                    var qtyFormat = secondField.StyleFormats.AddValuesFormat().ForValueField(qtyField);
                    qtyFormat.Style.Fill.BackgroundColor = XLColor.ArmyGreen;

                    // formatting of subtotal
                    nameField.StyleFormats.Subtotal.Label
                        .Style.Fill.BackgroundColor = XLColor.Aqua;
                    nameField.StyleFormats.Subtotal.AddValuesFormat()
                        .Style.Fill.BackgroundColor = XLColor.Aqua;
                    nameField.AddSubtotal(XLSubtotalFunction.Average);

                    ptSheet.Columns().AdjustToContents();
                }
                return wb;
            }, @"Examples\PivotTables\pivotstyles.xlsx");
        }
    }
}
