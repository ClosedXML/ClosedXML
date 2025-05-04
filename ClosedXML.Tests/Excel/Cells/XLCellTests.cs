using ClosedXML.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;

namespace ClosedXML.Tests
{
    [TestFixture]
    public class XLCellTests
    {
        [SuppressMessage("ReSharper", "RedundantCast")]
        private static readonly object[] AllNumberTypes =
        {
            (sbyte)1,
            (byte)2,
            (short)3,
            (ushort)4,
            (int)5,
            (uint)6,
            (long)7,
            (ulong)8,
            (float)9.5f,
            (double)10.75,
            (decimal)11.875m
        };

        [Test]
        public void CellsUsed()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Cell(1, 1);
            ws.Cell(2, 2);
            int count = ws.Range("A1:B2").CellsUsed().Count();
            Assert.That(count, Is.EqualTo(0));
        }

        [Test]
        public void CellsUsedIncludeStyles1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Row(3).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Column(3).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell(2, 2).Value = "ASDF";
            var range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
            Assert.That(range, Is.EqualTo("B2:C3"));
        }

        [Test]
        public void CellsUsedIncludeStyles2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Row(2).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Column(2).Style.Fill.BackgroundColor = XLColor.Red;
            ws.Cell(3, 3).Value = "ASDF";
            var range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
            Assert.That(range, Is.EqualTo("B2:C3"));
        }

        [Test]
        public void CellsUsedIncludeStyles3()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var range = ws.RangeUsed(XLCellsUsedOptions.All);
            Assert.That(range, Is.Null);
        }

        [Test]
        public void CellUsedIncludesSparklines()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.Range("C3:E4").Value = 1;
            ws.SparklineGroups.Add("B2", "C3:E3");
            ws.SparklineGroups.Add("F5", "C4:E4");

            var range = ws.RangeUsed(XLCellsUsedOptions.All).RangeAddress.ToString();
            Assert.That(range, Is.EqualTo("B2:F5"));
        }

        [Test]
        public void GetValue_Nullable()
        {
            var cell = new XLWorkbook().AddWorksheet().FirstCell();

            Assert.That(cell.Clear().GetValue<double?>(), Is.Null);
            Assert.Multiple(() =>
            {
                Assert.That(cell.SetValue(1.5).GetValue<double?>(), Is.EqualTo(1.5));
                Assert.That(cell.SetValue(2).GetValue<int?>(), Is.EqualTo(2));
            });
            Assert.That(cell.SetValue(Blank.Value).GetValue<double?>(), Is.Null);
            Assert.Throws<InvalidCastException>(() => cell.SetValue("text").GetValue<double?>());
        }

        [Test]
        public void InsertData1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRange range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" });
            Assert.That(range.ToString(), Is.EqualTo("Sheet1!B2:B4"));
        }

        [Test]
        public void InsertData_DoesntTransposeDataOnFalseFlag()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRange range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" }, false);
            Assert.That(range.ToString(), Is.EqualTo("Sheet1!B2:B4"));
        }

        [Test]
        public void InsertData_TransposesDataOnTrueFlag()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLRange range = ws.Cell(2, 2).InsertData(new[] { "a", "b", "c" }, true);
            Assert.That(range.ToString(), Is.EqualTo("Sheet1!B2:D2"));
        }

        [Test]
        public void InsertData_DifferentTypes()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            object[] values = { "Text", 45, DateTime.Today, true, "More text" };

            ws.FirstCell().InsertData(values);

            Assert.Multiple(() =>
            {
                Assert.That(ws.FirstCell().GetString(), Is.EqualTo("Text"));
                Assert.That(ws.Cell("A2").GetDouble(), Is.EqualTo(45));
                Assert.That(ws.Cell("A3").GetDateTime(), Is.EqualTo(DateTime.Today));
                Assert.That(ws.Cell("A4").GetBoolean(), Is.True);
                Assert.That(ws.Cell("A5").GetString(), Is.EqualTo("More text"));
                Assert.That(ws.Cell("A6").IsEmpty(), Is.True);
            });
        }

        [Test]
        public void InsertData_with_Guids()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.FirstCell().InsertData(Enumerable.Range(1, 20).Select(i => new { Guid = Guid.NewGuid() }));

            Assert.Multiple(() =>
            {
                Assert.That(ws.FirstCell().DataType, Is.EqualTo(XLDataType.Text));
                Assert.That(ws.FirstCell().GetText().Length, Is.EqualTo(Guid.NewGuid().ToString().Length));
            });
        }

        [Test]
        public void InsertData_with_Nulls()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");

            var table = new DataTable();
            table.TableName = "Patients";
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            table.Rows.Add(25, "Indocin", "David", new DateTime(2000, 1, 1));
            table.Rows.Add(50, "Enebrel", "Sam", new DateTime(2000, 1, 2));
            table.Rows.Add(10, "Hydralazine", "Christoff", new DateTime(2000, 1, 3));
            table.Rows.Add(21, "Combivent", DBNull.Value, new DateTime(2000, 1, 4));
            table.Rows.Add(100, "Dilantin", "Melanie", DBNull.Value);

            ws.FirstCell().InsertData(table);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").Value, Is.EqualTo(25));
                Assert.That(ws.Cell("C4").Value, Is.EqualTo(""));
                Assert.That(ws.Cell("D5").Value, Is.EqualTo(""));
            });
        }

        [Test]
        public void InsertData_with_Nulls_IEnumerable()
        {
            var ws = new XLWorkbook().Worksheets.Add("Sheet1");

            var dateTimeList = new List<DateTime?>()
            {
                new DateTime(2000, 1, 1),
                new DateTime(2000, 1, 2),
                new DateTime(2000, 1, 3),
                new DateTime(2000, 1, 4),
                null
            };

            ws.FirstCell().InsertData(dateTimeList);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").GetDateTime(), Is.EqualTo(new DateTime(2000, 1, 1)));
                Assert.That(ws.Cell("A5").Value, Is.EqualTo(Blank.Value));
            });
        }

        [Test]
        public void InsertData_AllNumberTypes_AreInsertedAsNumbers()
        {
            var ws = new XLWorkbook().Worksheets.Add();

            ws.FirstCell().InsertData(AllNumberTypes);

            for (var row = 1; row <= AllNumberTypes.Length; ++row)
            {
                var expectedValue = Convert.ChangeType(AllNumberTypes[row - 1], typeof(double));
                var actualValue = ws.Cell(row, 1).Value;
                Assert.That(actualValue, Is.EqualTo(expectedValue));
            }
        }

        [Test]
        public void InsertTable_AllNumberTypes_AreInsertedAsNumbers()
        {
            var ws = new XLWorkbook().Worksheets.Add();

            var table = new DataTable("Numbers");
            foreach (var number in AllNumberTypes)
            {
                var numberType = number.GetType();
                table.Columns.Add(numberType.Name, numberType);
            }

            table.Rows.Add(AllNumberTypes);

            ws.FirstCell().InsertTable(table);

            for (var column = 1; column <= AllNumberTypes.Length; ++column)
            {
                var expectedValue = Convert.ChangeType(AllNumberTypes[column - 1], typeof(double));
                var actualValue = ws.Cell(2, column).Value;
                Assert.That(actualValue, Is.EqualTo(expectedValue));
            }
        }

        [Test]
        public void IsEmpty1()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            bool actual = cell.IsEmpty();
            bool expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            bool actual = cell.IsEmpty(XLCellsUsedOptions.All);
            bool expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty3()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            bool actual = cell.IsEmpty();
            bool expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty4()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            bool actual = cell.IsEmpty(XLCellsUsedOptions.AllContents);
            bool expected = true;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty5()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Style.Fill.BackgroundColor = XLColor.Red;
            bool actual = cell.IsEmpty(XLCellsUsedOptions.All);
            bool expected = false;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void IsEmpty6()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Value = "X";
            bool actual = cell.IsEmpty();
            bool expected = false;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void NaN_is_not_a_number()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1");
            cell.Value = "NaN";

            Assert.That(cell.DataType, Is.Not.EqualTo(XLDataType.Number));
        }

        [Test]
        public void Nan_is_not_a_number()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1");
            cell.Value = "Nan";

            Assert.That(cell.DataType, Is.Not.EqualTo(XLDataType.Number));
        }

        [Test]
        public void TryGetValue_Boolean_Bad()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue("ABC");
            bool success = cell.TryGetValue(out bool outValue);
            Assert.That(success, Is.False);
        }

        [Test]
        public void TryGetValue_Boolean_False()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue(false);
            bool success = cell.TryGetValue(out bool outValue);
            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.False);
            });
        }

        [Test]
        public void TryGetValue_Boolean_FalseText()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue("False");
            var success = cell.TryGetValue(out bool outValue);
            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.False);
            });
        }

        [Test]
        public void TryGetValue_Boolean_True()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue(true);
            bool success = cell.TryGetValue(out bool outValue);
            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.True);
            });
        }

        [Test]
        public void TryGetValue_Boolean_TrueText()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue("True");
            var success = cell.TryGetValue(out bool outValue);
            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.True);
            });
        }

        [Test]
        public void TryGetValue_DateTime_Good2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            bool success = ws.Cell("A1").SetFormulaA1("=TODAY() + 10").TryGetValue(out DateTime outValue);
            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.EqualTo(DateTime.Today.AddDays(10)));
            });
        }

        [Test]
        public void TryGetValue_DateTime_BadButFormulaGood()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            bool success = ws.Cell("A1").SetFormulaA1("=\"44\"&\"020\"").TryGetValue(out DateTime outValue);
            Assert.That(success, Is.False);

            ws.Cell("B1").SetFormulaA1("=A1+1");

            success = ws.Cell("B1").TryGetValue(out outValue);
            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.EqualTo(new DateTime(2020, 07, 09)));
            });
        }

        [Test]
        public void TryGetValue_DateTime_BadString()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var date = "ABC";
            bool success = ws.Cell("A1").SetValue(date).TryGetValue(out DateTime outValue);
            Assert.That(success, Is.False);
        }

        [Test]
        public void TryGetValue_DateTime_SerialDateTimeOutsideRange()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var serialDateTimeOutsideRange = 5545454;
            ws.FirstCell().SetValue(serialDateTimeOutsideRange);
            bool success = ws.FirstCell().TryGetValue(out DateTime _);
            Assert.That(success, Is.False);
        }

        [Test]
        public void TryGetValue_Enum_Good()
        {
            var ws = new XLWorkbook().AddWorksheet();
            Assert.Multiple(() =>
            {
                Assert.That(ws.FirstCell().SetValue(nameof(NumberStyles.AllowCurrencySymbol)).TryGetValue(out NumberStyles value), Is.True);
                Assert.That(value, Is.EqualTo(NumberStyles.AllowCurrencySymbol));

                // Nullable alternative
                Assert.That(ws.FirstCell().SetValue(nameof(NumberStyles.AllowCurrencySymbol)).TryGetValue(out NumberStyles? value2), Is.True);
                Assert.That(value2, Is.EqualTo(NumberStyles.AllowCurrencySymbol));
            });
        }

        [Test]
        public void TryGetValue_Enum_BadString()
        {
            var ws = new XLWorkbook().AddWorksheet();
            Assert.Multiple(() =>
            {
                Assert.That(ws.FirstCell().SetValue("ABC").TryGetValue(out NumberStyles value), Is.False);
                Assert.That(ws.FirstCell().SetValue("ABC").TryGetValue(out NumberStyles? value2), Is.False);
            });
        }

        [Test]
        public void TryGetValue_TimeSpan_BadString()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            string timeSpan = "ABC";
            bool success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
            Assert.That(success, Is.False);
        }

        [Test]
        public void TryGetValue_TimeSpan_Good()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var timeSpan = new TimeSpan(1, 1, 1);
            bool success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.EqualTo(timeSpan));
            });
        }

        [Test]
        public void TryGetValue_TimeSpan_Good2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            bool success = ws.Cell("A1").SetValue(0.0034722222222222199).TryGetValue(out TimeSpan outValue);
            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.EqualTo(TimeSpan.FromMinutes(5)));
            });
        }

        [Test]
        public void TryGetValue_TimeSpan_Good_Large()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            var timeSpan = TimeSpan.FromMilliseconds((double)int.MaxValue + 1);
            bool success = ws.Cell("A1").SetValue(timeSpan).TryGetValue(out TimeSpan outValue);
            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.EqualTo(timeSpan));
            });
        }

        [Test]
        [SetCulture("en-US")]
        public void TryGetValue_TimeSpan_Good_FromText()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            bool success = ws.Cell("A1").SetValue("300:14:50.453").TryGetValue(out TimeSpan outValue);
            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.EqualTo(new TimeSpan(12, 12, 14, 50, 453)));
            });
        }

        [Test]
        public void TryGetValue_sbyte_Bad2()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue("255");
            bool success = cell.TryGetValue(out sbyte outValue);
            Assert.That(success, Is.False);
        }

        [Test]
        public void TryGetValue_sbyte_Good()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell("A1").SetValue(5);
            bool success = cell.TryGetValue(out sbyte outValue);
            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.EqualTo(5));
            });
        }

        [Test]
        public void TryGetValue_Unicode_String()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");

            bool success;
            string outValue;

            success = ws.Cell("A1")
                .SetValue("Site_x0020_Column_x0020_Test")
                .TryGetValue(out outValue);
            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.EqualTo("Site Column Test"));
            });

            success = ws.Cell("A1")
                .SetValue("Site_x005F_x0020_Column_x005F_x0020_Test")
                .TryGetValue(out outValue);

            Assert.Multiple(() =>
            {
                Assert.That(success, Is.True);
                Assert.That(outValue, Is.EqualTo("Site_x005F_x0020_Column_x005F_x0020_Test"));
            });
        }

        [Test]
        public void TryGetValue_Nullable()
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();

            ws.Cell("A1").Clear();
            ws.Cell("A2").SetValue(1.5);
            ws.Cell("A3").SetValue(2.5.ToString(CultureInfo.CurrentCulture));
            ws.Cell("A4").SetValue("text");

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("A1").TryGetValue(out double? _), Is.True);
                Assert.That(ws.Cell("A2").TryGetValue(out double? _), Is.True);
                Assert.That(ws.Cell("A3").TryGetValue(out double? _), Is.True);
                Assert.That(ws.Cell("A4").TryGetValue(out double? _), Is.False);
            });
        }

        [Test]
        public void CopyRangeAtCellAddress()
        {
            var ws = new XLWorkbook().AddWorksheet("Sheet1");

            ws.Cell("A1").SetValue(2)
                .CellRight().SetValue(3)
                .CellRight().SetValue(5)
                .CellRight().SetValue(7);

            var range = ws.Range("1:1");

            ws.Cell("B2").CopyFrom(range);

            Assert.Multiple(() =>
            {
                Assert.That(ws.Cell("B2").Value, Is.EqualTo(2));
                Assert.That(ws.Cell("C2").Value, Is.EqualTo(3));
                Assert.That(ws.Cell("D2").Value, Is.EqualTo(5));
                Assert.That(ws.Cell("E2").Value, Is.EqualTo(7));
            });
        }

        [Test]
        public void ValueSetToEmptyString()
        {
            string expected = string.Empty;

            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            cell.Value = new DateTime(2000, 1, 2);
            cell.Value = string.Empty;
            Assert.Multiple(() =>
            {
                Assert.That(cell.GetText(), Is.EqualTo(expected));
                Assert.That(cell.Value, Is.EqualTo(expected));
            });

            cell.Value = new DateTime(2000, 1, 2);
            cell.SetValue(string.Empty);
            Assert.Multiple(() =>
            {
                Assert.That(cell.GetText(), Is.EqualTo(expected));
                Assert.That(cell.Value, Is.EqualTo(expected));
            });
        }

        [Test]
        public void ValueSetDateWithShortUserDateFormat()
        {
            // For this test to make sense, user's local date format should be dd/MM/yy (note without the 2 century digits)
            // What happened previously was that the century digits got lost in .ToString() conversion and wrong century was sometimes returned.
            var ci = new CultureInfo(CultureInfo.InvariantCulture.LCID);
            ci.DateTimeFormat.ShortDatePattern = "dd/MM/yy";
            Thread.CurrentThread.CurrentCulture = ci;
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            IXLCell cell = ws.Cell(1, 1);
            var expected = DateTime.Today.AddYears(20);
            cell.Value = expected;
            var actual = (DateTime)cell.Value;
            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void SetStringValueTooLong()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.FirstCell().Value = new DateTime(2018, 5, 15);

            ws.FirstCell().SetValue(new string('A', 32767));

            Assert.Throws<ArgumentOutOfRangeException>(() => ws.FirstCell().Value = new string('A', 32768));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.FirstCell().SetValue(new string('A', 32768)));
        }

        [Test]
        public void SetCellValueWipesFormulas()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.FirstCell().FormulaA1 = "=TODAY()";
            ws.FirstCell().Value = "hello world";
            Assert.That(ws.FirstCell().HasFormula, Is.False);

            ws.FirstCell().FormulaA1 = "=TODAY()";
            ws.FirstCell().SetValue("hello world");
            Assert.That(ws.FirstCell().HasFormula, Is.False);
        }

        [Test]
        public void CellValueLineWrapping()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.FirstCell().Value = "hello world";
            Assert.That(ws.FirstCell().Style.Alignment.WrapText, Is.False);

            ws.FirstCell().Value = "hello\r\nworld";
            Assert.That(ws.FirstCell().Style.Alignment.WrapText, Is.True);

            ws.FirstCell().Style.Alignment.WrapText = false;

            ws.FirstCell().SetValue("hello world");
            Assert.That(ws.FirstCell().Style.Alignment.WrapText, Is.False);

            ws.FirstCell().SetValue("hello\r\nworld");
            Assert.That(ws.FirstCell().Style.Alignment.WrapText, Is.True);
        }

        [Test]
        public void TestInvalidXmlCharacters()
        {
            byte[] data;

            using (var stream = new MemoryStream())
            {
                var wb = new XLWorkbook();
                wb.AddWorksheet("Sheet1").FirstCell().SetValue("\u0018");
                wb.SaveAs(stream);
                data = stream.ToArray();
            }

            using (var stream = new MemoryStream(data))
            {
                var wb = new XLWorkbook(stream);
                Assert.That(wb.Worksheets.First().FirstCell().Value, Is.EqualTo("\u0018"));
            }
        }

        [Test]
        public void CanClearDateTimeCellValue()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Sheet1");
                var c = ws.FirstCell();
                c.SetValue(new DateTime(2017, 10, 08));
                Assert.Multiple(() =>
                {
                    Assert.That(c.DataType, Is.EqualTo(XLDataType.DateTime));
                    Assert.That(c.Value, Is.EqualTo(new DateTime(2017, 10, 08)));
                });

                wb.SaveAs(ms);
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                var c = ws.FirstCell();
                Assert.Multiple(() =>
                {
                    Assert.That(c.DataType, Is.EqualTo(XLDataType.DateTime));
                    Assert.That(c.Value, Is.EqualTo(new DateTime(2017, 10, 08)));
                });

                c.Clear();
                wb.Save();
            }

            using (var wb = new XLWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                var c = ws.FirstCell();
                Assert.Multiple(() =>
                {
                    Assert.That(c.DataType, Is.EqualTo(XLDataType.Blank));
                    Assert.That(c.IsEmpty(), Is.True);
                });
            }
        }

        [Test]
        public void ClearCellRemovesSparkline()
        {
            IXLWorksheet ws = new XLWorkbook().Worksheets.Add("Sheet1");
            ws.SparklineGroups.Add("B1:B3", "C1:E3");

            ws.Cell("B1").Clear(XLClearOptions.All);
            ws.Cell("B2").Clear(XLClearOptions.Sparklines);

            Assert.Multiple(() =>
            {
                Assert.That(ws.SparklineGroups.Single().Count(), Is.EqualTo(1));
                Assert.That(ws.Cell("B1").HasSparkline, Is.False);
                Assert.That(ws.Cell("B2").HasSparkline, Is.False);
                Assert.That(ws.Cell("B3").HasSparkline, Is.True);
            });
        }

        [Test]
        public void CurrentRegion()
        {
            // Partially based on sample in https://github.com/ClosedXML/ClosedXML/issues/120
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.Cell("B1").SetValue("x")
                .CellBelow().SetValue("x")
                .CellBelow().SetValue("x");

            ws.Cell("C1").SetValue("x")
                .CellBelow().SetValue("x")
                .CellBelow().SetValue("x");

            //Deliberately D2
            ws.Cell("D2").SetValue("x")
                .CellBelow().SetValue("x");

            ws.Cell("G1").SetValue("x")
                .CellBelow() // skip a cell
                .CellBelow().SetValue("x")
                .CellBelow().SetValue("x");

            // Deliberately H2
            ws.Cell("H2").SetValue("x")
                .CellBelow().SetValue("x")
                .CellBelow().SetValue("x");

            // A diagonal
            ws.Cell("E8").SetValue("x")
                .CellBelow().CellRight().SetValue("x")
                .CellBelow().CellRight().SetValue("x")
                .CellBelow().CellRight().SetValue("x")
                .CellBelow().CellRight().SetValue("x");

            Assert.That(ws.Cell("A10").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("A10:A10"));
            Assert.That(ws.Cell("B5").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("B5:B5"));
            Assert.AreEqual("P1:P1", ws.Cell("P1").CurrentRegion.RangeAddress.ToString());

            Assert.AreEqual("B1:D3", ws.Cell("D3").CurrentRegion.RangeAddress.ToString());
            Assert.AreEqual("B1:D4", ws.Cell("D4").CurrentRegion.RangeAddress.ToString());
            Assert.AreEqual("B1:E4", ws.Cell("E4").CurrentRegion.RangeAddress.ToString());

            foreach (var c in ws.Range("B1:D3").Cells())
            {
                Assert.That(c.CurrentRegion.RangeAddress.ToString(), Is.EqualTo("B1:D3"));
            }

            foreach (var c in ws.Range("A1:A3").Cells())
            {
                Assert.That(c.CurrentRegion.RangeAddress.ToString(), Is.EqualTo("A1:D3"));
            }

            Assert.That(ws.Cell("A4").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("A1:D4"));

            foreach (var c in ws.Range("E1:E3").Cells())
            {
                Assert.That(c.CurrentRegion.RangeAddress.ToString(), Is.EqualTo("B1:E3"));
            }

            Assert.That(ws.Cell("E4").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("B1:E4"));

            //// SECOND REGION
            foreach (var c in ws.Range("F1:F4").Cells())
            {
                Assert.That(c.CurrentRegion.RangeAddress.ToString(), Is.EqualTo("F1:H4"));
            }

            Assert.That(ws.Cell("F5").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("F1:H5"));

            //// DIAGONAL
            Assert.That(ws.Cell("E8").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));
            Assert.That(ws.Cell("F9").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));
            Assert.That(ws.Cell("G10").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));
            Assert.That(ws.Cell("H11").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));
            Assert.That(ws.Cell("I12").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));

            Assert.That(ws.Cell("G9").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));
            Assert.That(ws.Cell("F10").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:I12"));

            Assert.That(ws.Cell("D7").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("D7:I12"));
            Assert.That(ws.Cell("J13").CurrentRegion.RangeAddress.ToString(), Is.EqualTo("E8:J13"));

            // Four corners of a sheet
            Assert.That(ws.Cell(1, 1).CurrentRegion.RangeAddress.ToString(), Is.EqualTo("A1:D3"));
            Assert.That(ws.Cell(1, XLHelper.MaxColumnNumber).CurrentRegion.RangeAddress.ToString(), Is.EqualTo("XFD1:XFD1"));
            Assert.That(ws.Cell(XLHelper.MaxRowNumber, XLHelper.MaxColumnNumber).CurrentRegion.RangeAddress.ToString(), Is.EqualTo("XFD1048576:XFD1048576"));
            Assert.That(ws.Cell(XLHelper.MaxRowNumber, 1).CurrentRegion.RangeAddress.ToString(), Is.EqualTo("A1048576:A1048576"));
        }

        // https://github.com/ClosedXML/ClosedXML/issues/630
        [Test]
        public void ConsiderEmptyValueAsNumericInSumFormula()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");

            ws.Cell("A1").SetValue("Empty");
            ws.Cell("A2").SetValue("Numeric");
            ws.Cell("A3").SetValue("Copy of numeric");

            ws.Cell("B2").SetFormulaA1("=B1");
            ws.Cell("B3").SetFormulaA1("=B2");

            ws.Cell("C2").SetFormulaA1("=SUM(C1)");
            ws.Cell("C3").SetFormulaA1("=C2");

            var b1 = ws.Cell("B1").Value;
            var b2 = ws.Cell("B2").Value;
            var b3 = ws.Cell("B3").Value;

            Assert.Multiple(() =>
            {
                Assert.That(b1, Is.EqualTo(Blank.Value));
                Assert.That(b2, Is.EqualTo(0));
                Assert.That(b3, Is.EqualTo(0));
            });

            var c1 = ws.Cell("C1").Value;
            var c2 = ws.Cell("C2").Value;
            var c3 = ws.Cell("C3").Value;

            Assert.Multiple(() =>
            {
                Assert.That(c1, Is.EqualTo(Blank.Value));
                Assert.That(c2, Is.EqualTo(0));
                Assert.That(c3, Is.EqualTo(0));
            });
        }

        [Test]
        public void SetFormulaA1AffectsR1C1()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.FormulaR1C1 = "R[1]C";

            cell.FormulaA1 = "B2";

            Assert.That(cell.FormulaR1C1, Is.EqualTo("R[1]C[1]"));
        }

        [Test]
        public void SetFormulaR1C1AffectsA1()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var cell = ws.Cell(1, 1);
            cell.FormulaA1 = "A2";

            cell.FormulaR1C1 = "R[1]C[1]";

            Assert.That(cell.FormulaA1, Is.EqualTo("B2"));
        }

        [TestCase(" = 1 + SUM({ 1; 7})  - A8  ", "1 + SUM({ 1; 7})  - A8")]
        public void FormulaA1_setter_trims_and_removes_equal_if_present(string formula, string expectedResult)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaA1 = formula;
            Assert.That(ws.Cell("A1").FormulaA1, Is.EqualTo(expectedResult));
        }

        [TestCase(" =  1 +   R[1]C[7]  ", "1 +   R[1]C[7]")]
        public void FormulaR1C1_setter_trims_and_removes_equal_if_present(string formula, string expectedResult)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").FormulaR1C1 = formula;
            Assert.That(ws.Cell("A1").FormulaR1C1, Is.EqualTo(expectedResult));
        }

        [Test]
        public void FormulaWithCircularReferenceFails()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var A1 = ws.Cell("A1");
            var A2 = ws.Cell("A2");
            A1.FormulaA1 = "A2 + 1";
            A2.FormulaA1 = "A1 + 1";

            Assert.Throws(
                Is.TypeOf<InvalidOperationException>().And.Message.Contains("cycle"),
                () => _ = A1.Value);
            Assert.Throws(
                Is.TypeOf<InvalidOperationException>().And.Message.Contains("cycle"),
                () => _ = A2.Value);
        }

        [Test]
        public void InvalidFormulaShiftProducesREF()
        {
            using var ms = new MemoryStream();
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                ws.Cell("A1").Value = 1;
                ws.Cell("B1").Value = 2;
                ws.Cell("B2").FormulaA1 = "=A1+B1";

                Assert.That(ws.Cell("B2").Value, Is.EqualTo(3));

                ws.Range("B2").CopyTo(ws.Range("A2"));
                var fA2 = ws.Cell("A2").FormulaA1;

                wb.SaveAs(ms);

                Assert.That(fA2, Is.EqualTo("#REF!+A1"));
            }

            using (var wb2 = new XLWorkbook(ms))
            {
                var fA2 = wb2.Worksheets.First().Cell("A2").FormulaA1;
                Assert.That(fA2, Is.EqualTo("#REF!+A1"));
            }
        }

        [Test]
        public void FormulaWithCircularReferenceFails2()
        {
            var cell = new XLWorkbook().Worksheets.Add("Sheet1").FirstCell();
            cell.FormulaA1 = "A1";
            Assert.Throws<InvalidOperationException>(() =>
            {
                var _ = cell.Value;
            });
        }

        [Test]
        public void TryGetValueFormula_EvaluationFail_ReturnFalse()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var A1 = ws.Cell("A1");
            var A2 = ws.Cell("A2");
            var A3 = ws.Cell("A3");
            A1.FormulaA1 = "A2 + 1";
            A2.FormulaA1 = "A1 + 1";

            Assert.Multiple(() =>
            {
                Assert.That(A1.TryGetValue(out string _), Is.False);
                Assert.That(A2.TryGetValue(out string _), Is.False);
                Assert.That(A3.TryGetValue(out string _), Is.True);
            });
        }

        [Test]
        public void ToStringNoFormatString()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var c = ws.FirstCell().CellBelow(2).CellRight(3);

            Assert.That(c.ToString(), Is.EqualTo("D3"));
        }

        [Test]
        [TestCase("D3", "A")]
        [TestCase("YEAR(DATE(2018, 1, 1))", "F")]
        [TestCase("YEAR(DATE(2018, 1, 1))", "f")]
        [TestCase("0000.00", "NF")]
        [TestCase("0000.00", "nf")]
        [TestCase("FFFF0000", "fg")]
        [TestCase("Color Theme: Accent5, Tint: 0", "BG")]
        [TestCase("2018.00", "v")]
        public void ToStringFormatString(string expected, string format)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var c = ws.FirstCell().CellBelow(2).CellRight(3);

            var formula = "YEAR(DATE(2018, 1, 1))";
            c.FormulaA1 = formula;

            var numberFormat = "0000.00";
            c.Style.NumberFormat.Format = numberFormat;

            c.Style.Font.FontColor = XLColor.Red;
            c.Style.Fill.BackgroundColor = XLColor.FromTheme(XLThemeColor.Accent5);

            Assert.That(c.ToString(format), Is.EqualTo(expected));

            Assert.Throws<FormatException>(() => c.ToString("dummy"));
        }

        [Test]
        public void ToStringInvalidFormat()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Sheet1");
            var c = ws.FirstCell();

            Assert.Throws<FormatException>(() => c.ToString("dummy"));
        }

        [Test]
        public void Property_Active_is_true_when_cell_has_same_address_as_active_cell_in_worksheet()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.That(ws.ActiveCell, Is.Null);
            Assert.False(ws.Cell(1, 1).Active);

            ws.ActiveCell = ws.Cell("C4");
            Assert.That(ws.Cell("C4").Active, Is.True);
            Assert.False(ws.Cell("C5").Active);

            ws.ActiveCell = null;
            Assert.False(ws.Cell("C4").Active);
        }

        [Test]
        public void Property_Active_deactivates_cell_only_when_the_cell_is_active()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            ws.ActiveCell = ws.Cell("A2");

            ws.Cell("B2").Active = false;
            Assert.That(ws.ActiveCell, Is.EqualTo(ws.Cell("A2")));

            ws.Cell("A2").Active = false;
            Assert.That(ws.ActiveCell, Is.Null);
        }

        [Test]
        public void Property_Active_sets_cell_as_active_cell_of_worksheet()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            Assert.That(ws.ActiveCell, Is.Null);

            ws.Cell("B2").Active = true;
            Assert.That(ws.ActiveCell, Is.EqualTo(ws.Cell("B2")));
        }

        [TestCase("PY(4)", "_xlfn._xlws.PY(4)")]
        [TestCase("5 + py(abs(4) )", "5 + _xlfn._xlws.PY(abs(4) )")]
        [TestCase("COT(COTH(A5 + 2 * SIN(B7)))", "_xlfn.COT(_xlfn.COTH(A5 + 2 * SIN(B7)))")]
        [TestCase("_xlfn.COT(_xlfn.COTH(A5 + 2 * SIN(B7)))", "_xlfn.COT(_xlfn.COTH(A5 + 2 * SIN(B7)))")]
        public void FormulaA1_adds_prefix_to_future_functions(string formula, string expected)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var cell = ws.Cell("A1");
            cell.FormulaA1 = formula;

            Assert.That(cell.FormulaA1, Is.EqualTo(expected));
        }

        [TestCase("PY(4)", "_xlfn._xlws.PY(4)")]
        [TestCase("5 + py(abs(4) )", "5 + _xlfn._xlws.PY(abs(4) )")]
        [TestCase("COT(COTH(R[3]C[5] + 2 * SIN(R[7]C[2])))", "_xlfn.COT(_xlfn.COTH(R[3]C[5] + 2 * SIN(R[7]C[2])))")]
        [TestCase("_xlfn.COT(_xlfn.COTH(R[3]C[5] + 2 * SIN(R[7]C[2])))", "_xlfn.COT(_xlfn.COTH(R[3]C[5] + 2 * SIN(R[7]C[2])))")]
        public void FormulaR1C1_adds_prefix_to_future_functions(string formula, string expected)
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var cell = ws.Cell("A1");
            cell.FormulaR1C1 = formula;

            Assert.That(cell.FormulaR1C1, Is.EqualTo(expected));
        }

        [Test]
        public void FormulaA1_adds_prefix_to_all_future_functions()
        {
            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet();
            var cell = ws.Cell("A1");
            foreach (var (simpleName, prefixedName) in XLConstants.FutureFunctionMap.Value)
            {
                cell.FormulaA1 = simpleName + "()";
                Assert.That(cell.FormulaA1, Is.EqualTo(prefixedName + "()"));
            }
        }
    }
}
