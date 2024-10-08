using NUnit.Framework;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using ClosedXML.Excel;

namespace ClosedXML.Tests.Excel.Cells
{
    [TestFixture]
    public class XLCellValueTests
    {
        [Test]
        public void Creation_Blank()
        {
            XLCellValue blank = Blank.Value;
            Assert.AreEqual(XLDataType.Blank, blank.Type);
            Assert.True(blank.IsBlank);
        }

        [Test]
        public void Creation_Boolean()
        {
            XLCellValue logical = true;
            Assert.AreEqual(XLDataType.Boolean, logical.Type);
            Assert.True(logical.GetBoolean());
            Assert.True(logical.IsBoolean);
        }

        [Test]
        public void Creation_Number()
        {
            XLCellValue number = 14.0;
            Assert.AreEqual(XLDataType.Number, number.Type);
            Assert.True(number.IsNumber);
            Assert.AreEqual(14.0, number.GetNumber());
        }

        [TestCase(Double.NaN)]
        [TestCase(Double.PositiveInfinity)]
        [TestCase(Double.NegativeInfinity)]
        public void Creation_Number_CantBeNonNumber(Double nonNumber)
        {
            Assert.Throws<ArgumentException>(() => _ = (XLCellValue)nonNumber);
        }

        // Decimal is not allowed as a member of an attribute, so TestCase can't be used.
        private static readonly object[] DecimalTestCases =
        {
            new object[] { 5.875m, 5.875d },
            new object[] { Decimal.MaxValue, 7.922816251426434E+28 },
            new object[] { 1.0E-28m, 1.0000000000000001E-28d }
        };

        [TestCaseSource(nameof(DecimalTestCases))]
        public void Creation_Decimal(Decimal decimalNumber, Double expectedNumber)
        {
            XLCellValue cellValue = decimalNumber;
            Assert.True(cellValue.IsNumber);
            Assert.AreEqual(expectedNumber, cellValue.GetNumber());
        }

        [Test]
        public void Creation_Text()
        {
            XLCellValue text = "Hello World";
            Assert.AreEqual(XLDataType.Text, text.Type);
            Assert.AreEqual("Hello World", text.GetText());
        }

        [Test]
        public void NullString_IsConvertedToBlank()
        {
            XLCellValue value = (string)null;
            Assert.IsTrue(value.IsBlank);
            Assert.IsFalse(value.IsText);
        }

        [Test]
        public void Creation_Text_HasLimitedLength()
        {
            var longText = new string('A', 32768);
            Assert.Throws<ArgumentOutOfRangeException>(() => _ = (XLCellValue)longText);
        }

        [Test]
        public void Creation_Error()
        {
            XLCellValue error = XLError.NumberInvalid;
            Assert.AreEqual(XLDataType.Error, error.Type);
            Assert.True(error.IsError);
            Assert.AreEqual(XLError.NumberInvalid, error.GetError());
        }

        [Test]
        public void Creation_DateTime()
        {
            XLCellValue dateTime = new DateTime(2021, 1, 1);
            Assert.AreEqual(XLDataType.DateTime, dateTime.Type);
            Assert.True(dateTime.IsDateTime);
            Assert.AreEqual(new DateTime(2021, 1, 1), dateTime.GetDateTime());
        }

        [Test]
        public void Creation_TimeSpan()
        {
            XLCellValue dateTime = new TimeSpan(10, 1, 2, 3, 456);
            Assert.AreEqual(XLDataType.TimeSpan, dateTime.Type);
            Assert.True(dateTime.IsTimeSpan);
            Assert.AreEqual(new TimeSpan(10, 1, 2, 3, 456), dateTime.GetTimeSpan());
        }

        [Test]
        public void Creation_FromObject()
        {
            Assert.AreEqual(XLDataType.Blank, XLCellValue.FromObject(null).Type);
            Assert.AreEqual(XLDataType.Blank, XLCellValue.FromObject(Blank.Value).Type);
            Assert.AreEqual(XLDataType.Boolean, XLCellValue.FromObject(true).Type);
            Assert.AreEqual(XLDataType.Text, XLCellValue.FromObject("Hello World").Type);
            Assert.AreEqual(XLDataType.Error, XLCellValue.FromObject(XLError.NumberInvalid).Type);
            Assert.AreEqual(XLDataType.DateTime, XLCellValue.FromObject(new DateTime(2021, 1, 1)).Type);
            Assert.AreEqual(XLDataType.TimeSpan, XLCellValue.FromObject(new TimeSpan(10, 1, 2, 3, 456)).Type);
            Assert.AreEqual(XLDataType.Number, XLCellValue.FromObject((sbyte)42).Type);
            Assert.AreEqual(XLDataType.Number, XLCellValue.FromObject((byte)42).Type);
            Assert.AreEqual(XLDataType.Number, XLCellValue.FromObject((short)42).Type);
            Assert.AreEqual(XLDataType.Number, XLCellValue.FromObject((ushort)42).Type);
            Assert.AreEqual(XLDataType.Number, XLCellValue.FromObject((int)42).Type);
            Assert.AreEqual(XLDataType.Number, XLCellValue.FromObject((uint)42).Type);
            Assert.AreEqual(XLDataType.Number, XLCellValue.FromObject((long)42).Type);
            Assert.AreEqual(XLDataType.Number, XLCellValue.FromObject((ulong)42).Type);
            Assert.AreEqual(XLDataType.Number, XLCellValue.FromObject((float)42).Type);
            Assert.AreEqual(XLDataType.Number, XLCellValue.FromObject((double)42).Type);
            Assert.AreEqual(XLDataType.Number, XLCellValue.FromObject((decimal)42).Type);
            Assert.AreEqual(XLDataType.Text, XLCellValue.FromObject(DayOfWeek.Sunday).Type);
        }

        [Test]
        public void NumberTypes_HaveUnambiguousConversion()
        {
            {
                sbyte sbyteNumber = 5;
                XLCellValue sbyteCellValue = sbyteNumber;
                Assert.IsTrue(sbyteCellValue.IsNumber);
                Assert.AreEqual(5d, sbyteCellValue.GetNumber());
            }
            {
                byte byteNumber = 6;
                XLCellValue byteCellValue = byteNumber;
                Assert.IsTrue(byteCellValue.IsNumber);
                Assert.AreEqual(6d, byteCellValue.GetNumber());
            }
            {
                short shortNumber = 7;
                XLCellValue shortCellValue = shortNumber;
                Assert.IsTrue(shortCellValue.IsNumber);
                Assert.AreEqual(7d, shortCellValue.GetNumber());
            }
            {
                ushort ushortNumber = 8;
                XLCellValue ushortCellValue = ushortNumber;
                Assert.IsTrue(ushortCellValue.IsNumber);
                Assert.AreEqual(8d, ushortCellValue.GetNumber());
            }
            {
                int intNumber = 9;
                XLCellValue intCellValue = intNumber;
                Assert.IsTrue(intCellValue.IsNumber);
                Assert.AreEqual(9d, intCellValue.GetNumber());
            }
            {
                uint uintNumber = 10;
                XLCellValue uintCellValue = uintNumber;
                Assert.IsTrue(uintCellValue.IsNumber);
                Assert.AreEqual(10d, uintCellValue.GetNumber());
            }
            {
                long longNumber = 11;
                XLCellValue longCellValue = longNumber;
                Assert.IsTrue(longCellValue.IsNumber);
                Assert.AreEqual(11d, longCellValue.GetNumber());
            }
            {
                ulong ulongNumber = 12;
                XLCellValue ulongCellValue = ulongNumber;
                Assert.IsTrue(ulongCellValue.IsNumber);
                Assert.AreEqual(12d, ulongCellValue.GetNumber());
            }
            {
                float floatNumber = 13.5f;
                XLCellValue floatCellValue = floatNumber;
                Assert.IsTrue(floatCellValue.IsNumber);
                Assert.AreEqual(13.5d, floatCellValue.GetNumber());
            }
            {
                double doubleNumber = 14.5;
                XLCellValue doubleCellValue = doubleNumber;
                Assert.IsTrue(doubleCellValue.IsNumber);
                Assert.AreEqual(14.5d, doubleCellValue.GetNumber());
            }
            {
                decimal decimalNumber = 15.75m;
                XLCellValue decimalCellValue = decimalNumber;
                Assert.IsTrue(decimalCellValue.IsNumber);
                Assert.AreEqual(15.75d, decimalCellValue.GetNumber());
            }
        }

        [Test]
        [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
        public void NullableNumber_WithNullValue_AreConvertedToBlank()
        {
            {
                sbyte? sbyteNull = null;
                XLCellValue sbyteCellValue = sbyteNull;
                Assert.IsFalse(sbyteCellValue.IsNumber);
                Assert.IsTrue(sbyteCellValue.IsBlank);
            }
            {
                byte? byteNull = null;
                XLCellValue byteCellValue = byteNull;
                Assert.IsFalse(byteCellValue.IsNumber);
                Assert.IsTrue(byteCellValue.IsBlank);
            }
            {
                short? shortNull = null;
                XLCellValue shortCellValue = shortNull;
                Assert.IsFalse(shortCellValue.IsNumber);
                Assert.IsTrue(shortCellValue.IsBlank);
            }
            {
                ushort? ushortNull = null;
                XLCellValue ushortCellValue = ushortNull;
                Assert.IsFalse(ushortCellValue.IsNumber);
                Assert.IsTrue(ushortCellValue.IsBlank);
            }
            {
                int? intNull = null;
                XLCellValue intCellValue = intNull;
                Assert.IsFalse(intCellValue.IsNumber);
                Assert.IsTrue(intCellValue.IsBlank);
            }
            {
                uint? uintNull = null;
                XLCellValue uintCellValue = uintNull;
                Assert.IsFalse(uintCellValue.IsNumber);
                Assert.IsTrue(uintCellValue.IsBlank);
            }
            {
                long? longNull = null;
                XLCellValue longCellValue = longNull;
                Assert.IsFalse(longCellValue.IsNumber);
                Assert.IsTrue(longCellValue.IsBlank);
            }
            {
                ulong? ulongNull = null;
                XLCellValue ulongCellValue = ulongNull;
                Assert.IsFalse(ulongCellValue.IsNumber);
                Assert.IsTrue(ulongCellValue.IsBlank);
            }
            {
                float? floatValue = null;
                XLCellValue floatCellValue = floatValue;
                Assert.IsFalse(floatCellValue.IsNumber);
                Assert.IsTrue(floatCellValue.IsBlank);
            }
            {
                double? doubleValue = null;
                XLCellValue doubleCellValue = doubleValue;
                Assert.IsFalse(doubleCellValue.IsNumber);
                Assert.IsTrue(doubleCellValue.IsBlank);
            }
            {
                decimal? decimalValue = null;
                XLCellValue decimalCellValue = decimalValue;
                Assert.IsFalse(decimalCellValue.IsNumber);
                Assert.IsTrue(decimalCellValue.IsBlank);
            }
        }

        [Test]
        public void NullableNumber_WithNumberValue_AreConvertedToNumber()
        {
            {
                sbyte? sbyteNumber = 5;
                XLCellValue sbyteCellValue = sbyteNumber;
                Assert.IsTrue(sbyteCellValue.IsNumber);
                Assert.AreEqual(5d, sbyteCellValue.GetNumber());
            }
            {
                byte? byteNumber = 6;
                XLCellValue byteCellValue = byteNumber;
                Assert.IsTrue(byteCellValue.IsNumber);
                Assert.AreEqual(6d, byteCellValue.GetNumber());
            }
            {
                short? shortNumber = 7;
                XLCellValue shortCellValue = shortNumber;
                Assert.IsTrue(shortCellValue.IsNumber);
                Assert.AreEqual(7d, shortCellValue.GetNumber());
            }
            {
                ushort? ushortNumber = 8;
                XLCellValue ushortCellValue = ushortNumber;
                Assert.IsTrue(ushortCellValue.IsNumber);
                Assert.AreEqual(8d, ushortCellValue.GetNumber());
            }
            {
                int? intNumber = 9;
                XLCellValue intCellValue = intNumber;
                Assert.IsTrue(intCellValue.IsNumber);
                Assert.AreEqual(9d, intCellValue.GetNumber());
            }
            {
                uint? uintNumber = 9;
                XLCellValue uintCellValue = uintNumber;
                Assert.IsTrue(uintCellValue.IsNumber);
                Assert.AreEqual(9d, uintCellValue.GetNumber());
            }
            {
                long? longNumber = 10;
                XLCellValue longCellValue = longNumber;
                Assert.IsTrue(longCellValue.IsNumber);
                Assert.AreEqual(10d, longCellValue.GetNumber());
            }
            {
                ulong? ulongNumber = 11;
                XLCellValue ulongCellValue = ulongNumber;
                Assert.IsTrue(ulongCellValue.IsNumber);
                Assert.AreEqual(11d, ulongCellValue.GetNumber());
            }
            {
                float? floatNumber = 12.875f;
                XLCellValue floatCellValue = floatNumber;
                Assert.IsTrue(floatCellValue.IsNumber);
                Assert.AreEqual(12.875d, floatCellValue.GetNumber());
            }
            {
                double? doubleNumber = 13.875d;
                XLCellValue doubleCellValue = doubleNumber;
                Assert.IsTrue(doubleCellValue.IsNumber);
                Assert.AreEqual(13.875d, doubleCellValue.GetNumber());
            }
            {
                decimal? decimalNumber = 14.875m;
                XLCellValue decimalCellValue = decimalNumber;
                Assert.IsTrue(decimalCellValue.IsNumber);
                Assert.AreEqual(14.875d, decimalCellValue.GetNumber());
            }
        }

        [Test]
        [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
        public void NullableDateTime_WithNullValue_IsConvertedToBlank()
        {
            DateTime? dateTimeNull = null;
            XLCellValue dateTimeCellValue = dateTimeNull;
            Assert.IsFalse(dateTimeCellValue.IsDateTime);
            Assert.IsTrue(dateTimeCellValue.IsBlank);
        }

        [Test]
        public void NullableDateTime_WithDateValue_IsConvertedToDateTime()
        {
            DateTime? dateTime = new DateTime(2020, 5, 14, 8, 14, 30);
            XLCellValue dateTimeCellValue = dateTime;
            Assert.IsTrue(dateTimeCellValue.IsDateTime);
            Assert.AreEqual(dateTime.Value, dateTimeCellValue.GetDateTime());
        }

        [Test]
        [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
        public void NullableTimeSpan_WithNullValue_IsConvertedToBlank()
        {
            TimeSpan? timeSpanNull = null;
            XLCellValue timeSpanCellValue = timeSpanNull;
            Assert.IsFalse(timeSpanCellValue.IsTimeSpan);
            Assert.IsTrue(timeSpanCellValue.IsBlank);
        }

        [Test]
        public void NullableTimeSpan_WithTimeSpanValue_IsConvertedToTimeSpan()
        {
            TimeSpan? timeSpan = new TimeSpan(48, 12, 45, 30);
            XLCellValue timeSpanCellValue = timeSpan;
            Assert.IsTrue(timeSpanCellValue.IsTimeSpan);
            Assert.AreEqual(timeSpan.Value, timeSpanCellValue.GetTimeSpan());
        }

        [Test]
        public void UnifiedNumber_IsFormOf_Number_DateTime_And_TimeSpan()
        {
            XLCellValue value = Blank.Value;
            Assert.False(value.IsUnifiedNumber);

            value = true;
            Assert.False(value.IsUnifiedNumber);

            value = 14;
            Assert.True(value.IsUnifiedNumber);
            Assert.AreEqual(14.0, value.GetUnifiedNumber());

            value = new DateTime(1900, 1, 1);
            Assert.True(value.IsUnifiedNumber);
            Assert.AreEqual(1.0, value.GetUnifiedNumber());

            value = new TimeSpan(2, 12, 0, 0);
            Assert.True(value.IsUnifiedNumber);
            Assert.AreEqual(2.5, value.GetUnifiedNumber());

            value = "Text";
            Assert.False(value.IsUnifiedNumber);

            value = XLError.CellReference;
            Assert.False(value.IsUnifiedNumber);
        }

        [TestCase("1900-01-01", 1)]
        [TestCase("1900-01-02", 2)]
        [TestCase("1900-02-01", 32)]
        [TestCase("1900-02-28", 59)] // Excel assumes 1900 was a leap year and 29.1.1900 existed
        [TestCase("1900-03-01", 61)]
        [TestCase("2017-01-01", 42736)]
        public void SerialDateTime(string dateString, double expectedSerial)
        {
            XLCellValue date = DateTime.Parse(dateString);
            Assert.AreEqual(expectedSerial, date.GetUnifiedNumber());
        }

        [Test]
        [SetCulture("cs-CZ")]
        public void ToString_RespectsCulture()
        {
            XLCellValue v = Blank.Value;
            Assert.AreEqual(String.Empty, v.ToString());

            v = true;
            Assert.AreEqual("TRUE", v.ToString());

            v = 25.4;
            Assert.AreEqual("25,4", v.ToString());

            v = "Hello";
            Assert.AreEqual("Hello", v.ToString());

            v = XLError.IncompatibleValue;
            Assert.AreEqual("#VALUE!", v.ToString());

            v = new DateTime(1900, 1, 2);
            Assert.AreEqual("02.01.1900 0:00:00", v.ToString());

            v = new DateTime(1900, 3, 1, 4, 10, 5);
            Assert.AreEqual("01.03.1900 4:10:05", v.ToString());

            v = new TimeSpan(4, 5, 6, 7, 82);
            Assert.AreEqual("101:06:07,082", v.ToString());
        }

        [Test]
        public void TryConvert_Blank()
        {
            XLCellValue value = Blank.Value;
            Assert.True(value.TryConvert(out Blank blank));
            Assert.AreEqual(Blank.Value, blank);

            value = String.Empty;
            Assert.True(value.TryConvert(out blank));
            Assert.AreEqual(Blank.Value, blank);
        }

        [Test]
        public void TryConvert_Boolean()
        {
            XLCellValue value = true;
            Assert.True(value.TryConvert(out Boolean boolean));
            Assert.True(boolean);

            value = "True";
            Assert.True(value.TryConvert(out boolean));
            Assert.True(boolean);

            value = "False";
            Assert.True(value.TryConvert(out boolean));
            Assert.False(boolean);

            value = 0;
            Assert.True(value.TryConvert(out boolean));
            Assert.False(boolean);

            value = 0.001;
            Assert.True(value.TryConvert(out boolean));
            Assert.True(boolean);
        }

        [Test]
        public void TryConvert_Number()
        {
            var c = CultureInfo.GetCultureInfo("cs-CZ");
            XLCellValue value = 5;
            Assert.True(value.TryConvert(out Double number, c));
            Assert.AreEqual(5.0, number);

            value = "1,5";
            Assert.True(value.TryConvert(out number, c));
            Assert.AreEqual(1.5, number);

            value = "1 1/4";
            Assert.True(value.TryConvert(out number, c));
            Assert.AreEqual(1.25, number);

            value = "3.1.1900";
            Assert.True(value.TryConvert(out number, c));
            Assert.AreEqual(3, number);

            value = true;
            Assert.True(value.TryConvert(out number, c));
            Assert.AreEqual(1.0, number);

            value = false;
            Assert.True(value.TryConvert(out number, c));
            Assert.AreEqual(0.0, number);

            value = new DateTime(2020, 4, 5, 10, 14, 5);
            Assert.True(value.TryConvert(out number, c));
            Assert.AreEqual(43926.42644675926, number);

            value = new TimeSpan(18, 0, 0);
            Assert.True(value.TryConvert(out number, c));
            Assert.AreEqual(0.75, number);
        }

        [Test]
        public void TryConvert_DateTime()
        {
            XLCellValue v = new DateTime(2020, 1, 1);
            Assert.True(v.TryConvert(out DateTime dt));
            Assert.AreEqual(new DateTime(2020, 1, 1), dt);

            var lastSerialDate = 2958465;
            v = lastSerialDate;
            Assert.True(v.TryConvert(out dt));
            Assert.AreEqual(new DateTime(9999, 12, 31), dt);

            v = lastSerialDate + 1;
            Assert.False(v.TryConvert(out dt));

            v = new TimeSpan(14, 0, 0, 0);
            Assert.True(v.TryConvert(out dt));
            Assert.AreEqual(new DateTime(1900, 1, 14), dt);
        }

        [Test]
        public void TryConvert_TimeSpan()
        {
            var c = CultureInfo.GetCultureInfo("cs-CZ");
            XLCellValue v = new TimeSpan(10, 15, 30);
            Assert.True(v.TryConvert(out TimeSpan ts, c));
            Assert.AreEqual(new TimeSpan(10, 15, 30), ts);

            v = "26:15:30,5";
            Assert.True(v.TryConvert(out ts, c));
            Assert.AreEqual(new TimeSpan(1, 2, 15, 30, 500), ts);

            v = 0.75;
            Assert.True(v.TryConvert(out ts, c));
            Assert.AreEqual(new TimeSpan(18, 0, 0), ts);
        }

        [TestCase(1)]
        [TestCase(10)] // microsecond
        [TestCase(3000000001)] // 5 min 1 tick
        public void TimeSpan_can_have_sub_millisecond_precision(long ticks)
        {
            var subMsTimeSpan = TimeSpan.FromTicks(ticks);
            XLCellValue value = subMsTimeSpan;
            Assert.AreEqual(subMsTimeSpan, value.GetTimeSpan());
        }

        [TestCase(1)]
        [TestCase(10)] // microsecond
        [TestCase(3000000001)] // 5 min 1 tick
        public void TimeSpan_with_sub_millisecond_precision_is_written_and_loaded_correctly(long ticks)
        {
            // NetFx converts double to string using G15. Core changed it to G17, but ClosedXML still use G15.
            var subMsTimeSpan = TimeSpan.FromTicks(ticks);
            TestHelper.CreateSaveLoadAssert(
                (_, ws) =>
                {
                    ws.Cell("A1").Value = subMsTimeSpan;
                },
                (_, ws) =>
                {
                    var cellValue = ws.Cell("A1").CachedValue;
                    Assert.AreEqual(subMsTimeSpan, cellValue.GetTimeSpan());
                });
        }

        [TestCase(long.MaxValue / (double)TimeSpan.TicksPerDay + 0.01)]
        [TestCase(long.MinValue / (double)TimeSpan.TicksPerDay - 0.01)]
        public void TimeSpan_throws_when_not_representable(double serialDateTime)
        {
            var value = XLCellValue.FromSerialTimeSpan(serialDateTime);
            var ex = Assert.Throws<OverflowException>(() => value.GetTimeSpan())!;
            Assert.AreEqual("The serial date time value is too large to be represented in a TimeSpan.", ex.Message);
        }
    }
}
