using System.Collections.Generic;
using ClosedXML.Excel;
using ClosedXML.Tests.Utils;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles
{
    public class BorderTests
    {
        private const int None = 0x0;
        private const int Left = 0x1;
        private const int Right = 0x2;
        private const int Top = 0x4;
        private const int Bottom = 0x8;
        private const int All = Left | Top | Right | Bottom;

        [Test]
        public void InsideBorders_preserves_outside_borders()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var halfAccent1Color = XLColor.FromTheme(XLThemeColor.Accent1, 0.5);
            ws.Cells("B2:C2").Style
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(halfAccent1Color);

            // Check pre-conditions
            AssertCellBorder(ws, "B2", Left | Right | Top | Bottom, XLBorderStyleValues.Thin, halfAccent1Color);
            AssertCellBorder(ws, "C2", Left | Right | Top | Bottom, XLBorderStyleValues.Thin, halfAccent1Color);

            ws.Range("B2:C2").Style.Border.SetInsideBorder(XLBorderStyleValues.None);

            AssertCellBorder(ws, "B2", Left | Top | Bottom, XLBorderStyleValues.Thin);
            AssertCellBorder(ws, "C2", Right| Top | Bottom, XLBorderStyleValues.Thin);
            Assert.AreEqual(XLThemeColor.Accent1, ws.Cell("B2").Style.Border.LeftBorderColor.ThemeColor);
            Assert.AreEqual(XLThemeColor.Accent1, ws.Cell("C2").Style.Border.RightBorderColor.ThemeColor);
        }

        [Test]
        public void InsideBorder_sets_border_that_is_not_on_edge()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:D4").Style
                .Border.SetInsideBorderColor(XLColor.Red)
                .Border.SetInsideBorder(XLBorderStyleValues.Thick);

            AssertCellBorder(ws, "B2", Bottom | Right);
            AssertCellBorder(ws, "C2", Left | Bottom | Right);
            AssertCellBorder(ws, "D2", Bottom | Left);
            AssertCellBorder(ws, "B3", Top | Right | Bottom);
            AssertCellBorder(ws, "C3", Left | Top | Right | Bottom);
            AssertCellBorder(ws, "D3", Top | Left | Bottom);
            AssertCellBorder(ws, "B4", Top | Right);
            AssertCellBorder(ws, "C4", Left | Top | Right);
            AssertCellBorder(ws, "D4", Left | Top);
        }

        [Test]
        public void OutsideBorder_sets_only_borders_on_edge()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Range("B2:D4").Style
                .Border.SetOutsideBorderColor(XLColor.Red)
                .Border.SetOutsideBorder(XLBorderStyleValues.Thick);

            AssertCellBorder(ws, "B2", Top | Left);
            AssertCellBorder(ws, "C2", Top);
            AssertCellBorder(ws, "D2", Top | Right);
            AssertCellBorder(ws, "B3", Left);
            AssertCellBorder(ws, "C3", None);
            AssertCellBorder(ws, "D3", Right);
            AssertCellBorder(ws, "B4", Bottom | Left);
            AssertCellBorder(ws, "C4", Bottom);
            AssertCellBorder(ws, "D4", Bottom | Right);
        }

        [Test]
        public void OutsideBorder_used_for_cells_will_make_outside_border_around_each_cell()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cells("B2:C3").Style
                .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XLColor.Red);

            AssertCellBorder(ws, "B2", All, XLBorderStyleValues.Thick, XLColor.Red);
            AssertCellBorder(ws, "C2", All, XLBorderStyleValues.Thick, XLColor.Red);
            AssertCellBorder(ws, "B3", All, XLBorderStyleValues.Thick, XLColor.Red);
            AssertCellBorder(ws, "C3", All, XLBorderStyleValues.Thick, XLColor.Red);
        }

        [Test]
        public void InsideBorder_used_for_cells_is_ignored_because_individual_cells_have_no_inside()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cells("B2:C3").Style
                .Border.SetInsideBorder(XLBorderStyleValues.Thick)
                .Border.SetInsideBorderColor(XLColor.Red);

            AssertCellBorder(ws, "B2", None, XLBorderStyleValues.None);
            AssertCellBorder(ws, "C2", None, XLBorderStyleValues.None);
            AssertCellBorder(ws, "B3", None, XLBorderStyleValues.None);
            AssertCellBorder(ws, "C3", None, XLBorderStyleValues.None);
        }

        [Test, Ignore("Performance reasons")] // TODO Styles: Enable after switch
        public void OutsideBorder_for_columns()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // Materialize a cell with style to test interactions between column and materialized cell
            ws.Cell("B2").Style
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(XLColor.Blue);

            // Set border for a whole column
            ws.Column("B").Style
                .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XLColor.Red);

            var b2 = ws.Cell("B2").Style.Border;
            Assert.AreEqual(XLBorderStyleValues.Thin, b2.TopBorder);
            Assert.AreEqual(XLColor.Blue, b2.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, b2.BottomBorder);
            Assert.AreEqual(XLColor.Blue, b2.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thick, b2.LeftBorder);
            Assert.AreEqual(XLColor.Red, b2.LeftBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thick, b2.RightBorder);
            Assert.AreEqual(XLColor.Red, b2.RightBorderColor);
            AssertCellBorder(ws, "B1", Left | Right, XLBorderStyleValues.Thick, XLColor.Red);
            AssertCellBorder(ws, "B3", Left | Right, XLBorderStyleValues.Thick, XLColor.Red);
        }

        [Test, Ignore("Performance reasons")] // TODO Styles: Enable after switch
        public void InsideBorder_for_columns()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // Materialize a cell with style to test interactions between column and materialized cell
            ws.Cell("B2").Style
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(XLColor.Blue);

            // Set border for a whole column
            ws.Column("B").Style
                .Border.SetInsideBorder(XLBorderStyleValues.Thick)
                .Border.SetInsideBorderColor(XLColor.Red);

            var b2 = ws.Cell("B2").Style.Border;
            Assert.AreEqual(XLBorderStyleValues.Thick, b2.TopBorder);
            Assert.AreEqual(XLColor.Red, b2.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thick, b2.BottomBorder);
            Assert.AreEqual(XLColor.Red, b2.BottomBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, b2.LeftBorder);
            Assert.AreEqual(XLColor.Blue, b2.LeftBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, b2.RightBorder);
            Assert.AreEqual(XLColor.Blue, b2.RightBorderColor);
            AssertCellBorder(ws, "B1", Top | Bottom, XLBorderStyleValues.Thick, XLColor.Red);
            AssertCellBorder(ws, "B3", Top | Bottom, XLBorderStyleValues.Thick, XLColor.Red);
        }

        [Test]
        public void OutsideBorder_for_rows()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // Materialize a cell with style to test interactions between row and materialized cell
            ws.Cell("B2").Style
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(XLColor.Blue);

            // Set border for a whole row
            ws.Row(2).Style
                .Border.SetOutsideBorder(XLBorderStyleValues.Thick)
                .Border.SetOutsideBorderColor(XLColor.Red);

            var b2 = ws.Cell("B2").Style.Border;
            Assert.AreEqual(XLBorderStyleValues.Thin, b2.LeftBorder);
            Assert.AreEqual(XLColor.Blue, b2.LeftBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, b2.RightBorder);
            Assert.AreEqual(XLColor.Blue, b2.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thick, b2.TopBorder);
            Assert.AreEqual(XLColor.Red, b2.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thick, b2.BottomBorder);
            Assert.AreEqual(XLColor.Red, b2.BottomBorderColor);

            // TODO Styles: Enable after switch, repository makes a mess with equality
            // AssertCellBorder(ws, "A2", Top | Bottom, XLBorderStyleValues.Thick, XLColor.Red);
            // AssertCellBorder(ws, "C2", Top | Bottom, XLBorderStyleValues.Thick, XLColor.Red);
        }

        [Test]
        public void InsideBorder_for_rows()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            // Materialize a cell with style to test interactions between row and materialized cell
            ws.Cell("B2").Style
                .Border.SetOutsideBorder(XLBorderStyleValues.Thin)
                .Border.SetOutsideBorderColor(XLColor.Blue);

            // Set border for a whole row
            ws.Row(2).Style
                .Border.SetInsideBorder(XLBorderStyleValues.Thick)
                .Border.SetInsideBorderColor(XLColor.Red);

            var b2 = ws.Cell("B2").Style.Border;
            Assert.AreEqual(XLBorderStyleValues.Thick, b2.LeftBorder);
            Assert.AreEqual(XLColor.Red, b2.LeftBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thick, b2.RightBorder);
            Assert.AreEqual(XLColor.Red, b2.RightBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, b2.TopBorder);
            Assert.AreEqual(XLColor.Blue, b2.TopBorderColor);
            Assert.AreEqual(XLBorderStyleValues.Thin, b2.BottomBorder);
            Assert.AreEqual(XLColor.Blue, b2.BottomBorderColor);

            // TODO Styles: Enable after switch, repository makes a mess with equality
            // AssertCellBorder(ws, "A2", Left | Right, XLBorderStyleValues.Thick, XLColor.Red);
            // AssertCellBorder(ws, "C2", Left | Right, XLBorderStyleValues.Thick, XLColor.Red);
        }

        [Test]
        [TestCaseSource(nameof(BorderApiSetters))]
        public void Border_property_can_be_individually_set(FormatTestCase<IXLBorder> testCase)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var cell = ws.Cell("B2");

            foreach (var testValue in testCase.Values)
            {
                testCase.SetPropertyValue(cell.Style.Border, testValue);
                var setValue = testCase.GetPropertyValue(cell.Style.Border);
                Assert.AreEqual(testValue, setValue);
                cell = cell.CellRight();
            }
        }

        private static void AssertCellBorder(IXLWorksheet ws, string cell, int sides, XLBorderStyleValues style = XLBorderStyleValues.Thick, XLColor color = null)
        {
            var border = ws.Cell(cell).Style.Border;
            Assert.AreEqual((sides & Left) != 0 ? style : XLBorderStyleValues.None, border.LeftBorder);
            Assert.AreEqual((sides & Right) != 0 ? style : XLBorderStyleValues.None, border.RightBorder);
            Assert.AreEqual((sides & Top) != 0 ? style : XLBorderStyleValues.None, border.TopBorder);
            Assert.AreEqual((sides & Bottom) != 0 ? style : XLBorderStyleValues.None, border.BottomBorder);

            if (color is not null)
            {
                Assert.AreEqual((sides & Left) != 0 ? color : XLColor.Auto, border.LeftBorderColor);
                Assert.AreEqual((sides & Right) != 0 ? color : XLColor.Auto, border.RightBorderColor);
                Assert.AreEqual((sides & Top) != 0 ? color : XLColor.Auto, border.TopBorderColor);
                Assert.AreEqual((sides & Bottom) != 0 ? color : XLColor.Auto, border.BottomBorderColor);
            }
        }

        private static IEnumerable<object> BorderApiSetters()
        {
            var styleValues = EnumPolyfill.GetValues<XLBorderStyleValues>();
            var colors = new[] { XLColor.Red, XLColor.Black, XLColor.Auto };

            // Outside border style - check that once set, all outer borders are set to the style
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.LeftBorder, (border, value) => border.OutsideBorder = value, styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.RightBorder, (border, value) => border.OutsideBorder = value, styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.TopBorder, (border, value) => border.OutsideBorder = value, styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.BottomBorder, (border, value) => border.OutsideBorder = value, styleValues);

            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.LeftBorder, (border, value) => border.SetOutsideBorder(value), styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.RightBorder, (border, value) => border.SetOutsideBorder(value), styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.TopBorder, (border, value) => border.SetOutsideBorder(value), styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.BottomBorder, (border, value) => border.SetOutsideBorder(value), styleValues);

            // Outside border color - check that once set, all outer borders are set to the color
            // Because of hash key adn repos, we must first set style other than none
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.LeftBorderColor, (border, value) => border.SetOutsideBorder(XLBorderStyleValues.Thin).Border.OutsideBorderColor = value, colors);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.RightBorderColor, (border, value) => border.SetOutsideBorder(XLBorderStyleValues.Thin).Border.OutsideBorderColor = value, colors);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.TopBorderColor, (border, value) => border.SetOutsideBorder(XLBorderStyleValues.Thin).Border.OutsideBorderColor = value, colors);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.BottomBorderColor, (border, value) => border.SetOutsideBorder(XLBorderStyleValues.Thin).Border.OutsideBorderColor = value, colors);

            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.LeftBorderColor, (border, value) => border.SetOutsideBorder(XLBorderStyleValues.Thin).Border.SetOutsideBorderColor(value), colors);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.RightBorderColor, (border, value) => border.SetOutsideBorder(XLBorderStyleValues.Thin).Border.SetOutsideBorderColor(value), colors);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.TopBorderColor, (border, value) => border.SetOutsideBorder(XLBorderStyleValues.Thin).Border.SetOutsideBorderColor(value), colors);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.BottomBorderColor, (border, value) => border.SetOutsideBorder(XLBorderStyleValues.Thin).Border.SetOutsideBorderColor(value), colors);

            // Left border
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.LeftBorder, (border, value) => border.LeftBorder = value, styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.LeftBorder, (border, value) => border.SetLeftBorder(value), styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.LeftBorderColor, (border, value) => border.SetLeftBorder(XLBorderStyleValues.Thin).Border.LeftBorderColor = value, colors);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.LeftBorderColor, (border, value) => border.SetLeftBorder(XLBorderStyleValues.Thin).Border.SetLeftBorderColor(value), colors);

            // Right border
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.RightBorder, (border, value) => border.RightBorder = value, styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.RightBorder, (border, value) => border.SetRightBorder(value), styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.RightBorderColor, (border, value) => border.SetRightBorder(XLBorderStyleValues.Thin).Border.RightBorderColor = value, colors);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.RightBorderColor, (border, value) => border.SetRightBorder(XLBorderStyleValues.Thin).Border.SetRightBorderColor(value), colors);

            // Top border
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.TopBorder, (border, value) => border.TopBorder = value, styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.TopBorder, (border, value) => border.SetTopBorder(value), styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.TopBorderColor, (border, value) => border.SetTopBorder(XLBorderStyleValues.Thin).Border.TopBorderColor = value, colors);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.TopBorderColor, (border, value) => border.SetTopBorder(XLBorderStyleValues.Thin).Border.SetTopBorderColor(value), colors);

            // Bottom border
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.BottomBorder, (border, value) => border.BottomBorder = value, styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.BottomBorder, (border, value) => border.SetBottomBorder(value), styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.BottomBorderColor, (border, value) => border.SetBottomBorder(XLBorderStyleValues.Thin).Border.BottomBorderColor = value, colors);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.BottomBorderColor, (border, value) => border.SetBottomBorder(XLBorderStyleValues.Thin).Border.SetBottomBorderColor(value), colors);

            // Diagonal up
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.DiagonalUp, (border, value) => border.DiagonalUp = value, true, false);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.DiagonalUp, (border, value) => border.SetDiagonalUp(value), true, false);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.DiagonalUp, (border, _) => border.SetDiagonalUp(), true);

            // Diagonal down
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.DiagonalDown, (border, value) => border.DiagonalDown = value, true, false);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.DiagonalDown, (border, value) => border.SetDiagonalDown(value), true, false);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.DiagonalDown, (border, _) => border.SetDiagonalDown(), true);

            // Diagonal border
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.DiagonalBorder, (border, value) => border.DiagonalBorder = value, styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.DiagonalBorder, (border, value) => border.SetDiagonalBorder(value), styleValues);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.DiagonalBorderColor, (border, value) => border.SetDiagonalBorder(XLBorderStyleValues.Thin).Border.DiagonalBorderColor = value, colors);
            yield return FormatTestCase<IXLBorder>.ForBorder(border => border.DiagonalBorderColor, (border, value) => border.SetDiagonalBorder(XLBorderStyleValues.Thin).Border.SetDiagonalBorderColor(value), colors);
        }
    }
}
