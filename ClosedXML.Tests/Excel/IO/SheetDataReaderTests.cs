using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.Excel.IO;
using ClosedXML.IO;
using NUnit.Framework;
using StringItem = ClosedXML.Excel.CalcEngine.OneOf<string, ClosedXML.Excel.XLImmutableRichText>;

namespace ClosedXML.Tests.Excel.IO;

[TestFixture]
internal class SheetDataReaderTests
{
    private const string Ns = OpenXmlConst.Main2006SsNs;

    [Test]
    public void Empty_sheetData_is_ok()
    {
        AssertSheetData(
            "<sheetData/>",
            (_, ws) => Assert.True(ws.Internals.CellsCollection.IsEmpty));
    }

    [Test]
    public void Tracks_row_number_when_r_is_omitted()
    {
        AssertSheetData(
            """
            <sheetData>
              <row>
                <c r="A1"><v>1</v></c>
              </row>
              <row>
                <c r="A2"><v>2</v></c>
              </row>
              <row r="5">
                <c r="A5"><v>5</v></c>
              </row>
              <row>
                <c r="A6"><v>6</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual(1, ws.Cell("A1").GetDouble());
                Assert.AreEqual(2, ws.Cell("A2").GetDouble());
                Assert.AreEqual(5, ws.Cell("A5").GetDouble());
                Assert.AreEqual(6, ws.Cell("A6").GetDouble());
            });
    }

    [Test]
    public void Tracks_column_when_cell_r_is_omitted()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c><v>1</v></c>
                <c><v>2</v></c>
                <c r="D1"><v>4</v></c>
                <c><v>5</v></c>
              </row>
              <row r="2">
                <c><v>10</v></c>
                <c><v>11</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual(1, ws.Cell("A1").GetDouble());
                Assert.AreEqual(2, ws.Cell("B1").GetDouble());
                Assert.AreEqual(4, ws.Cell("D1").GetDouble());
                Assert.AreEqual(5, ws.Cell("E1").GetDouble());
                Assert.AreEqual(10, ws.Cell("A2").GetDouble());
                Assert.AreEqual(11, ws.Cell("B2").GetDouble());
            });
    }

    [Test]
    public void Default_cell_type_is_number()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1"><v>1.5</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual(XLDataType.Number, ws.Cell("A1").DataType);
                Assert.AreEqual(1.5, ws.Cell("A1").GetDouble());
            });
    }

    [Test]
    public void Number_without_value_stays_blank()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="n"/>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.True(ws.Cell("A1").Value.IsBlank));
    }

    [Test]
    public void Unparsable_number_stays_blank()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="n"><v>not-a-number</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.True(ws.Cell("A1").Value.IsBlank));
    }

    [TestCase("1", true)]
    [TestCase("TRUE", true)]
    [TestCase("true", true)]
    [TestCase("0", false)]
    [TestCase("FALSE", false)]
    public void Reads_boolean(string xmlValue, bool expected)
    {
        AssertSheetData(
            $"""
             <sheetData>
               <row r="1">
                 <c r="A1" t="b"><v>{xmlValue}</v></c>
               </row>
             </sheetData>
             """,
            (ws, _) => Assert.AreEqual(expected, ws.Cell("A1").GetBoolean()));
    }

    [Test]
    public void Reads_error()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="e"><v>#DIV/0!</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.AreEqual(XLError.DivisionByZero, ws.Cell("A1").GetError()));
    }

    [Test]
    public void Shared_string_uses_sst()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="s"><v>1</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.AreEqual("world", ws.Cell("A1").GetString()),
            sst: ["hello", "world"]);
    }

    [Test]
    public void Shared_string_out_of_range_is_empty()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="s"><v>9</v></c>
                <c r="B1" t="s"><v>nope</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual(string.Empty, ws.Cell("A1").GetString());
                Assert.AreEqual(string.Empty, ws.Cell("B1").GetString());
            },
            sst: ["only"]);
    }

    [Test]
    public void Shared_string_without_value_is_empty_string()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="s"/>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.AreEqual(string.Empty, ws.Cell("A1").GetString()));
    }

    [Test]
    public void Inline_string()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="inlineStr">
                  <is>
                    <t>Hello_x0009_!</t>
                  </is>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual("Hello\t!", ws.Cell("A1").GetString());
                Assert.False(ws.Cell("A1").ShareString);
            });
    }

    [Test]
    public void Inline_rich_text()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="inlineStr">
                  <is>
                    <r>
                      <t>Hi</t>
                    </r>
                  </is>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.True(ws.Cell("A1").HasRichText);
                Assert.AreEqual("Hi", ws.Cell("A1").GetRichText().Text);
                Assert.False(ws.Cell("A1").ShareString);
            });
    }

    [Test]
    public void Formula_that_results_in_string_type_has_cell_type_formula_string()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="str">
                  <f>T("x")</f>
                  <v>x</v>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual("T(\"x\")", ws.Cell("A1").FormulaA1);
                Assert.AreEqual("x", ws.Cell("A1").GetString());
                Assert.False(ws.Cell("A1").ShareString);
            });
    }

    [TestCase("2006-02-03T09:30:00.000", "2006-02-03T09:30:00")]
    [TestCase("2006-02-03T09:30", "2006-02-03T09:30:00")]
    [TestCase("2006-02-03", "2006-02-03T00:00:00")]
    [TestCase(" 2006-02-03 ", "2006-02-03T00:00:00")]
    public void Reads_iso_date(string xmlValue, string expectedIso)
    {
        AssertSheetData(
            $"""
             <sheetData>
               <row r="1">
                 <c r="A1" t="d"><v>{xmlValue}</v></c>
               </row>
             </sheetData>
             """,
            (ws, _) => Assert.AreEqual(DateTime.Parse(expectedIso), ws.Cell("A1").GetDateTime()));
    }

    [Test]
    public void Unknown_cell_type_throws()
    {
        Assert.Throws<FormatException>(() => AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="nope"><v>1</v></c>
              </row>
            </sheetData>
            """,
            (_, _) => { }));
    }

    [Test]
    public void Normal_formula_with_cached_value()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1">
                  <f>1+2</f>
                  <v>3</v>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual("1+2", ws.Cell("A1").FormulaA1);
                Assert.AreEqual(3.0, ws.Cell("A1").CachedValue);
                Assert.False(ws.Cell("A1").ShareString);
                Assert.False(ws.Cell("A1").NeedsRecalculation);
            });
    }

    [Test]
    public void Formula_without_value_is_dirty()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1">
                  <f>1+2</f>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual("1+2", ws.Cell("A1").FormulaA1);
                Assert.True(ws.Cell("A1").NeedsRecalculation);
            });
    }

    [Test]
    public void Array_formula_applies_to_range()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1">
                  <f t="array" ref="A1:B1">1+2</f>
                  <v>3</v>
                </c>
                <c r="B1">
                  <v>3</v>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.True(ws.Cell("A1").HasArrayFormula);
                Assert.True(ws.Cell("B1").HasArrayFormula);
                Assert.AreEqual("1+2", ws.Cell("A1").FormulaA1);
                Assert.AreEqual("1+2", ws.Cell("B1").FormulaA1);
                Assert.AreEqual("A1:B1", ws.Cell("A1").FormulaReference.ToStringRelative());
                Assert.AreEqual("A1:B1", ws.Cell("B1").FormulaReference.ToStringRelative());
            });
    }

    [Test]
    public void Shared_formula_translates_relative_references()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="B1">
                  <f t="shared" si="0">A1+1</f>
                  <v>2</v>
                </c>
                <c r="C1">
                  <f t="shared" si="0"/>
                  <v>3</v>
                </c>
              </row>
            </sheetData>
            """,
            (ws, _) =>
            {
                Assert.AreEqual("A1+1", ws.Cell("B1").FormulaA1);
                Assert.AreEqual("B1+1", ws.Cell("C1").FormulaA1);
                Assert.False(ws.Cell("B1").ShareString);
                Assert.False(ws.Cell("C1").ShareString);
            });
    }

    [Test]
    public void Data_table_formula_1d_row()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="2">
                <c r="B2">
                  <f t="dataTable" ref="B2:D2" dt2D="0" dtr="1" r1="A1" del1="1"/>
                </c>
              </row>
            </sheetData>
            """,
            (_, ws) =>
            {
                var formula = ws.Cell("B2")!.Formula;
                Assert.AreEqual(FormulaType.DataTable, formula.Type);
                Assert.True(formula.IsRowDataTable);
                Assert.False(formula.Is2DDataTable);
                Assert.True(formula.Input1Deleted);
                Assert.AreEqual(new Point(1, 1), formula.Input1);
                Assert.AreEqual("B2:D2", formula.Range.ToString());
            });
    }

    [Test]
    public void Data_table_formula_2d()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="2">
                <c r="B2">
                  <f t="dataTable" ref="B2:C3" dt2D="1" r1="A1" r2="A2" del2="1"/>
                </c>
              </row>
            </sheetData>
            """,
            (_, ws) =>
            {
                var formula = ws.Cell("B2")!.Formula;
                Assert.AreEqual(FormulaType.DataTable, formula.Type);
                Assert.True(formula.Is2DDataTable);
                Assert.False(formula.Input1Deleted);
                Assert.True(formula.Input2Deleted);
                Assert.AreEqual(new Point(1, 1), formula.Input1);
                Assert.AreEqual(new Point(2, 1), formula.Input2);
            });
    }

    [Test]
    public void Row_properties()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1" ht="20" hidden="1" collapsed="1" outlineLevel="2" ph="1" customFormat="1" s="0"
                   x14ac:dyDescent="0.25">
                <c r="A1"><v>1</v></c>
              </row>
            </sheetData>
            """,
            (_, ws) =>
            {
                var row = ws.Row(1);
                Assert.AreEqual(20, row.Height);
                Assert.True(row.IsHidden);
                Assert.True(row.Collapsed);
                Assert.AreEqual(2, row.OutlineLevel);
                Assert.True(row.ShowPhonetic);
                Assert.AreEqual(0.25, row.DyDescent);
                Assert.AreSame(ws.Workbook.Styles.DefaultCellFormat, row.FormatValue);
            });
    }

    [Test]
    public void Row_without_ht_uses_sheet_row_height()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1"><v>1</v></c>
              </row>
            </sheetData>
            """,
            (ws, _) => Assert.AreEqual(ws.RowHeight, ws.Row(1).Height));
    }

    [Test]
    public void Cell_style_phonetic_and_meta_indexes()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1" s="0" ph="1" cm="3" vm="4"><v>1</v></c>
              </row>
            </sheetData>
            """,
            (_, ws) =>
            {
                var cell = ws.Cell("A1")!;
                Assert.AreSame(ws.Workbook.Styles.DefaultCellFormat, cell.FormatValue);
                Assert.True(cell.ShowPhonetic);
                Assert.AreEqual(3u, cell.CellMetaIndex);
                Assert.AreEqual(4u, cell.ValueMetaIndex);
            });
    }

    [Test]
    public void Skips_extLst_on_row_and_cell()
    {
        AssertSheetData(
            """
            <sheetData>
              <row r="1">
                <c r="A1">
                  <v>1</v>
                  <extLst>
                    <ext uri="{12345678-1234-1234-1234-1234567890AB}"/>
                  </extLst>
                </c>
                <c r="B1"><v>2</v></c>
                <extLst>
                  <ext uri="{12345678-1234-1234-1234-1234567890AB}"/>
                </extLst>
              </row>
              <row r="2">
                <c r="A2"><v>3</v></c>
              </row>
            </sheetData>
            """,
            (_, _) => { });
    }

    [Test]
    public void Number_with_date_format_is_datetime()
    {
        using var wb = new XLWorkbook();
        var dateTimeFormat = wb.Styles.NumberFormats[(int)XLPredefinedFormat.DateTime.DayMonthYear4WithSlashes];
        var ws = wb.AddWorksheet();
        var dateFormat = wb.Styles.GetRegisteredCellFormat(
            wb.Styles.DefaultCellFormat,
            f => f with { NumberFormat = dateTimeFormat });
        var xfId = wb.Styles.CellFormats[dateFormat];

        LoadSheetData(
            wb,
            ws,
            $"""
             <sheetData>
               <row r="1">
                 <c r="A1" s="{xfId}" t="n"><v>44927</v></c>
               </row>
             </sheetData>
             """);

        Assert.AreEqual(XLDataType.DateTime, ws.Cell("A1").DataType);
        Assert.AreEqual(DateTime.FromOADate(44927), ws.Cell("A1").GetDateTime());
    }

    [Test]
    public void Number_with_time_format_is_timespan()
    {
        using var wb = new XLWorkbook();
        var timeSpanFormat = wb.Styles.NumberFormats[(int)XLPredefinedFormat.DateTime.Hour24MinutesSeconds];
        var ws = wb.AddWorksheet();
        var timeFormat = wb.Styles.GetRegisteredCellFormat(
            wb.Styles.DefaultCellFormat,
            f => f with { NumberFormat = timeSpanFormat });
        var xfId = wb.Styles.CellFormats[timeFormat];

        LoadSheetData(
            wb,
            ws,
            $"""
             <sheetData>
               <row r="1">
                 <c r="A1" s="{xfId}" t="n"><v>0.5</v></c>
               </row>
             </sheetData>
             """);

        Assert.AreEqual(XLDataType.TimeSpan, ws.Cell("A1").DataType);
        Assert.AreEqual(TimeSpan.FromDays(0.5), ws.Cell("A1").GetTimeSpan());
    }

    [Test]
    public void Date1904_shifts_datetime()
    {
        using var wb = new XLWorkbook();
        wb.Use1904DateSystem = true;
        var ws = wb.AddWorksheet();
        LoadSheetData(
            wb,
            ws,
            """
            <sheetData>
              <row r="1">
                <c r="A1" t="d"><v>2006-01-01</v></c>
              </row>
            </sheetData>
            """);

        Assert.AreEqual(new DateTime(2006, 1, 1).AddDays(1462), ws.Cell("A1").GetDateTime());
    }

    private static void AssertSheetData(
        string sheetDataXml,
        Action<IXLWorksheet, XLWorksheet> assert,
        IReadOnlyList<StringItem> sst = null)
    {
        using var wb = new XLWorkbook();
        var ws = (XLWorksheet)wb.AddWorksheet();
        LoadSheetData(wb, ws, sheetDataXml, sst);
        assert(ws, ws);
    }

    private static void LoadSheetData(
        XLWorkbook wb,
        IXLWorksheet ws,
        string sheetDataXml,
        IReadOnlyList<StringItem> sst = null)
    {
        using var xmlReader = OpenWorksheet(sheetDataXml);
        var parsed = new SheetDataReader(xmlReader, (XLWorksheet)ws, wb.Styles, [.. sst ?? []])
            .ParseCtSheetData("sheetData", Ns);
        Assert.True(parsed.IsSuccess);
    }

    private static XmlTreeReader OpenWorksheet(string innerXml)
    {
        var xml = $"""
                   <worksheet xmlns="{Ns}" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
                     {innerXml}
                   </worksheet>
                   """;
        var stream = new MemoryStream(XLHelper.NoBomUTF8.GetBytes(xml));
        var xmlReader = new XmlTreeReader(stream, XmlToEnumMapper.Instance, true);
        xmlReader.Open("worksheet", Ns);
        return xmlReader;
    }
}
