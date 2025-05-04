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
            Assert.Multiple(() =>
            {
                Assert.That(blank.Type, Is.EqualTo(XLDataType.Blank));
                Assert.That(blank.IsBlank, Is.True);
            });
        }

        [Test]
        public void Creation_Boolean()
        {
            XLCellValue logical = true;
            Assert.Multiple(() =>
            {
                Assert.That(logical.Type, Is.EqualTo(XLDataType.Boolean));
                Assert.That(logical.GetBoolean(), Is.True);
                Assert.That(logical.IsBoolean, Is.True);
            });
        }

        [Test]
        public void Creation_Number()
        {
            XLCellValue number = 14.0;
            Assert.Multiple(() =>
            {
                Assert.That(number.Type, Is.EqualTo(XLDataType.Number));
                Assert.That(number.IsNumber, Is.True);
                Assert.That(number.GetNumber(), Is.EqualTo(14.0));
            });
        }

        [TestCase(double.NaN)]
        [TestCase(double.PositiveInfinity)]
        [TestCase(double.NegativeInfinity)]
        public void Creation_Number_CantBeNonNumber(double nonNumber)
        {
            Assert.Throws<ArgumentException>(() => _ = (XLCellValue)nonNumber);
        }

        // Decimal is not allowed as a member of an attribute, so TestCase can't be used.
        private static readonly object[] DecimalTestCases =
        {
            new object[] { 5.875m, 5.875d },
            new object[] { decimal.MaxValue, 7.922816251426434E+28 },
            new object[] { 1.0E-28m, 1.0000000000000001E-28d }
        };

        [TestCaseSource(nameof(DecimalTestCases))]
        public void Creation_Decimal(decimal decimalNumber, double expectedNumber)
        {
            XLCellValue cellValue = decimalNumber;
            Assert.Multiple(() =>
            {
                Assert.That(cellValue.IsNumber, Is.True);
                Assert.That(cellValue.GetNumber(), Is.EqualTo(expectedNumber));
            });
        }

        [Test]
        public void Creation_Text()
        {
            XLCellValue text = "Hello World";
            Assert.Multiple(() =>
            {
                Assert.That(text.Type, Is.EqualTo(XLDataType.Text));
                Assert.That(text.GetText(), Is.EqualTo("Hello World"));
            });
        }

        [Test]
        public void NullString_IsConvertedToBlank()
        {
            XLCellValue value = (string)null;
            Assert.Multiple(() =>
            {
                Assert.That(value.IsBlank, Is.True);
                Assert.That(value.IsText, Is.False);
            });
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
            Assert.Multiple(() =>
            {
                Assert.That(error.Type, Is.EqualTo(XLDataType.Error));
                Assert.That(error.IsError, Is.True);
                Assert.That(error.GetError(), Is.EqualTo(XLError.NumberInvalid));
            });
        }

        [Test]
        public void Creation_DateTime()
        {
            XLCellValue dateTime = new DateTime(2021, 1, 1);
            Assert.Multiple(() =>
            {
                Assert.That(dateTime.Type, Is.EqualTo(XLDataType.DateTime));
                Assert.That(dateTime.IsDateTime, Is.True);
                Assert.That(dateTime.GetDateTime(), Is.EqualTo(new DateTime(2021, 1, 1)));
            });
        }

        [Test]
        public void Creation_TimeSpan()
        {
            XLCellValue dateTime = new TimeSpan(10, 1, 2, 3, 456);
            Assert.Multiple(() =>
            {
                Assert.That(dateTime.Type, Is.EqualTo(XLDataType.TimeSpan));
                Assert.That(dateTime.IsTimeSpan, Is.True);
                Assert.That(dateTime.GetTimeSpan(), Is.EqualTo(new TimeSpan(10, 1, 2, 3, 456)));
            });
        }

        [Test]
        public void Creation_FromObject()
        {
            Assert.Multiple(() =>
            {
                Assert.That(XLCellValue.FromObject(null).Type, Is.EqualTo(XLDataType.Blank));
                Assert.That(XLCellValue.FromObject(Blank.Value).Type, Is.EqualTo(XLDataType.Blank));
                Assert.That(XLCellValue.FromObject(true).Type, Is.EqualTo(XLDataType.Boolean));
                Assert.That(XLCellValue.FromObject("Hello World").Type, Is.EqualTo(XLDataType.Text));
                Assert.That(XLCellValue.FromObject(XLError.NumberInvalid).Type, Is.EqualTo(XLDataType.Error));
                Assert.That(XLCellValue.FromObject(new DateTime(2021, 1, 1)).Type, Is.EqualTo(XLDataType.DateTime));
                Assert.That(XLCellValue.FromObject(new TimeSpan(10, 1, 2, 3, 456)).Type, Is.EqualTo(XLDataType.TimeSpan));
                Assert.That(XLCellValue.FromObject((sbyte)42).Type, Is.EqualTo(XLDataType.Number));
                Assert.That(XLCellValue.FromObject((byte)42).Type, Is.EqualTo(XLDataType.Number));
                Assert.That(XLCellValue.FromObject((short)42).Type, Is.EqualTo(XLDataType.Number));
                Assert.That(XLCellValue.FromObject((ushort)42).Type, Is.EqualTo(XLDataType.Number));
                Assert.That(XLCellValue.FromObject((int)42).Type, Is.EqualTo(XLDataType.Number));
                Assert.That(XLCellValue.FromObject((uint)42).Type, Is.EqualTo(XLDataType.Number));
                Assert.That(XLCellValue.FromObject((long)42).Type, Is.EqualTo(XLDataType.Number));
                Assert.That(XLCellValue.FromObject((ulong)42).Type, Is.EqualTo(XLDataType.Number));
                Assert.That(XLCellValue.FromObject((float)42).Type, Is.EqualTo(XLDataType.Number));
                Assert.That(XLCellValue.FromObject((double)42).Type, Is.EqualTo(XLDataType.Number));
                Assert.That(XLCellValue.FromObject((decimal)42).Type, Is.EqualTo(XLDataType.Number));
                Assert.That(XLCellValue.FromObject(DayOfWeek.Sunday).Type, Is.EqualTo(XLDataType.Text));
            });
        }

        [Test]
        public void NumberTypes_HaveUnambiguousConversion()
        {
            {
                sbyte sbyteNumber = 5;
                XLCellValue sbyteCellValue = sbyteNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(sbyteCellValue.IsNumber, Is.True);
                    Assert.That(sbyteCellValue.GetNumber(), Is.EqualTo(5d));
                });
            }
            {
                byte byteNumber = 6;
                XLCellValue byteCellValue = byteNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(byteCellValue.IsNumber, Is.True);
                    Assert.That(byteCellValue.GetNumber(), Is.EqualTo(6d));
                });
            }
            {
                short shortNumber = 7;
                XLCellValue shortCellValue = shortNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(shortCellValue.IsNumber, Is.True);
                    Assert.That(shortCellValue.GetNumber(), Is.EqualTo(7d));
                });
            }
            {
                ushort ushortNumber = 8;
                XLCellValue ushortCellValue = ushortNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(ushortCellValue.IsNumber, Is.True);
                    Assert.That(ushortCellValue.GetNumber(), Is.EqualTo(8d));
                });
            }
            {
                int intNumber = 9;
                XLCellValue intCellValue = intNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(intCellValue.IsNumber, Is.True);
                    Assert.That(intCellValue.GetNumber(), Is.EqualTo(9d));
                });
            }
            {
                uint uintNumber = 10;
                XLCellValue uintCellValue = uintNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(uintCellValue.IsNumber, Is.True);
                    Assert.That(uintCellValue.GetNumber(), Is.EqualTo(10d));
                });
            }
            {
                long longNumber = 11;
                XLCellValue longCellValue = longNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(longCellValue.IsNumber, Is.True);
                    Assert.That(longCellValue.GetNumber(), Is.EqualTo(11d));
                });
            }
            {
                ulong ulongNumber = 12;
                XLCellValue ulongCellValue = ulongNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(ulongCellValue.IsNumber, Is.True);
                    Assert.That(ulongCellValue.GetNumber(), Is.EqualTo(12d));
                });
            }
            {
                float floatNumber = 13.5f;
                XLCellValue floatCellValue = floatNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(floatCellValue.IsNumber, Is.True);
                    Assert.That(floatCellValue.GetNumber(), Is.EqualTo(13.5d));
                });
            }
            {
                double doubleNumber = 14.5;
                XLCellValue doubleCellValue = doubleNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(doubleCellValue.IsNumber, Is.True);
                    Assert.That(doubleCellValue.GetNumber(), Is.EqualTo(14.5d));
                });
            }
            {
                decimal decimalNumber = 15.75m;
                XLCellValue decimalCellValue = decimalNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(decimalCellValue.IsNumber, Is.True);
                    Assert.That(decimalCellValue.GetNumber(), Is.EqualTo(15.75d));
                });
            }
        }

        [Test]
        [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
        public void NullableNumber_WithNullValue_AreConvertedToBlank()
        {
            {
                sbyte? sbyteNull = null;
                XLCellValue sbyteCellValue = sbyteNull;
                Assert.Multiple(() =>
                {
                    Assert.That(sbyteCellValue.IsNumber, Is.False);
                    Assert.That(sbyteCellValue.IsBlank, Is.True);
                });
            }
            {
                byte? byteNull = null;
                XLCellValue byteCellValue = byteNull;
                Assert.Multiple(() =>
                {
                    Assert.That(byteCellValue.IsNumber, Is.False);
                    Assert.That(byteCellValue.IsBlank, Is.True);
                });
            }
            {
                short? shortNull = null;
                XLCellValue shortCellValue = shortNull;
                Assert.Multiple(() =>
                {
                    Assert.That(shortCellValue.IsNumber, Is.False);
                    Assert.That(shortCellValue.IsBlank, Is.True);
                });
            }
            {
                ushort? ushortNull = null;
                XLCellValue ushortCellValue = ushortNull;
                Assert.Multiple(() =>
                {
                    Assert.That(ushortCellValue.IsNumber, Is.False);
                    Assert.That(ushortCellValue.IsBlank, Is.True);
                });
            }
            {
                int? intNull = null;
                XLCellValue intCellValue = intNull;
                Assert.Multiple(() =>
                {
                    Assert.That(intCellValue.IsNumber, Is.False);
                    Assert.That(intCellValue.IsBlank, Is.True);
                });
            }
            {
                uint? uintNull = null;
                XLCellValue uintCellValue = uintNull;
                Assert.Multiple(() =>
                {
                    Assert.That(uintCellValue.IsNumber, Is.False);
                    Assert.That(uintCellValue.IsBlank, Is.True);
                });
            }
            {
                long? longNull = null;
                XLCellValue longCellValue = longNull;
                Assert.Multiple(() =>
                {
                    Assert.That(longCellValue.IsNumber, Is.False);
                    Assert.That(longCellValue.IsBlank, Is.True);
                });
            }
            {
                ulong? ulongNull = null;
                XLCellValue ulongCellValue = ulongNull;
                Assert.Multiple(() =>
                {
                    Assert.That(ulongCellValue.IsNumber, Is.False);
                    Assert.That(ulongCellValue.IsBlank, Is.True);
                });
            }
            {
                float? floatValue = null;
                XLCellValue floatCellValue = floatValue;
                Assert.Multiple(() =>
                {
                    Assert.That(floatCellValue.IsNumber, Is.False);
                    Assert.That(floatCellValue.IsBlank, Is.True);
                });
            }
            {
                double? doubleValue = null;
                XLCellValue doubleCellValue = doubleValue;
                Assert.Multiple(() =>
                {
                    Assert.That(doubleCellValue.IsNumber, Is.False);
                    Assert.That(doubleCellValue.IsBlank, Is.True);
                });
            }
            {
                decimal? decimalValue = null;
                XLCellValue decimalCellValue = decimalValue;
                Assert.Multiple(() =>
                {
                    Assert.That(decimalCellValue.IsNumber, Is.False);
                    Assert.That(decimalCellValue.IsBlank, Is.True);
                });
            }
        }

        [Test]
        public void NullableNumber_WithNumberValue_AreConvertedToNumber()
        {
            {
                sbyte? sbyteNumber = 5;
                XLCellValue sbyteCellValue = sbyteNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(sbyteCellValue.IsNumber, Is.True);
                    Assert.That(sbyteCellValue.GetNumber(), Is.EqualTo(5d));
                });
            }
            {
                byte? byteNumber = 6;
                XLCellValue byteCellValue = byteNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(byteCellValue.IsNumber, Is.True);
                    Assert.That(byteCellValue.GetNumber(), Is.EqualTo(6d));
                });
            }
            {
                short? shortNumber = 7;
                XLCellValue shortCellValue = shortNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(shortCellValue.IsNumber, Is.True);
                    Assert.That(shortCellValue.GetNumber(), Is.EqualTo(7d));
                });
            }
            {
                ushort? ushortNumber = 8;
                XLCellValue ushortCellValue = ushortNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(ushortCellValue.IsNumber, Is.True);
                    Assert.That(ushortCellValue.GetNumber(), Is.EqualTo(8d));
                });
            }
            {
                int? intNumber = 9;
                XLCellValue intCellValue = intNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(intCellValue.IsNumber, Is.True);
                    Assert.That(intCellValue.GetNumber(), Is.EqualTo(9d));
                });
            }
            {
                uint? uintNumber = 9;
                XLCellValue uintCellValue = uintNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(uintCellValue.IsNumber, Is.True);
                    Assert.That(uintCellValue.GetNumber(), Is.EqualTo(9d));
                });
            }
            {
                long? longNumber = 10;
                XLCellValue longCellValue = longNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(longCellValue.IsNumber, Is.True);
                    Assert.That(longCellValue.GetNumber(), Is.EqualTo(10d));
                });
            }
            {
                ulong? ulongNumber = 11;
                XLCellValue ulongCellValue = ulongNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(ulongCellValue.IsNumber, Is.True);
                    Assert.That(ulongCellValue.GetNumber(), Is.EqualTo(11d));
                });
            }
            {
                float? floatNumber = 12.875f;
                XLCellValue floatCellValue = floatNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(floatCellValue.IsNumber, Is.True);
                    Assert.That(floatCellValue.GetNumber(), Is.EqualTo(12.875d));
                });
            }
            {
                double? doubleNumber = 13.875d;
                XLCellValue doubleCellValue = doubleNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(doubleCellValue.IsNumber, Is.True);
                    Assert.That(doubleCellValue.GetNumber(), Is.EqualTo(13.875d));
                });
            }
            {
                decimal? decimalNumber = 14.875m;
                XLCellValue decimalCellValue = decimalNumber;
                Assert.Multiple(() =>
                {
                    Assert.That(decimalCellValue.IsNumber, Is.True);
                    Assert.That(decimalCellValue.GetNumber(), Is.EqualTo(14.875d));
                });
            }
        }

        [Test]
        [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
        public void NullableDateTime_WithNullValue_IsConvertedToBlank()
        {
            DateTime? dateTimeNull = null;
            XLCellValue dateTimeCellValue = dateTimeNull;
            Assert.Multiple(() =>
            {
                Assert.That(dateTimeCellValue.IsDateTime, Is.False);
                Assert.That(dateTimeCellValue.IsBlank, Is.True);
            });
        }

        [Test]
        public void NullableDateTime_WithDateValue_IsConvertedToDateTime()
        {
            DateTime? dateTime = new DateTime(2020, 5, 14, 8, 14, 30);
            XLCellValue dateTimeCellValue = dateTime;
            Assert.Multiple(() =>
            {
                Assert.That(dateTimeCellValue.IsDateTime, Is.True);
                Assert.That(dateTimeCellValue.GetDateTime(), Is.EqualTo(dateTime.Value));
            });
        }

        [Test]
        [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
        public void NullableTimeSpan_WithNullValue_IsConvertedToBlank()
        {
            TimeSpan? timeSpanNull = null;
            XLCellValue timeSpanCellValue = timeSpanNull;
            Assert.Multiple(() =>
            {
                Assert.That(timeSpanCellValue.IsTimeSpan, Is.False);
                Assert.That(timeSpanCellValue.IsBlank, Is.True);
            });
        }

        [Test]
        public void NullableTimeSpan_WithTimeSpanValue_IsConvertedToTimeSpan()
        {
            TimeSpan? timeSpan = new TimeSpan(48, 12, 45, 30);
            XLCellValue timeSpanCellValue = timeSpan;
            Assert.Multiple(() =>
            {
                Assert.That(timeSpanCellValue.IsTimeSpan, Is.True);
                Assert.That(timeSpanCellValue.GetTimeSpan(), Is.EqualTo(timeSpan.Value));
            });
        }

        [Test]
        public void UnifiedNumber_IsFormOf_Number_DateTime_And_TimeSpan()
        {
            XLCellValue value = Blank.Value;
            Assert.False(value.IsUnifiedNumber);

            value = true;
            Assert.False(value.IsUnifiedNumber);

            value = 14;
            Assert.Multiple(() =>
            {
                Assert.That(value.IsUnifiedNumber, Is.True);
                Assert.That(value.GetUnifiedNumber(), Is.EqualTo(14.0));
            });

            value = new DateTime(1900, 1, 1);
            Assert.Multiple(() =>
            {
                Assert.That(value.IsUnifiedNumber, Is.True);
                Assert.That(value.GetUnifiedNumber(), Is.EqualTo(1.0));
            });

            value = new TimeSpan(2, 12, 0, 0);
            Assert.Multiple(() =>
            {
                Assert.That(value.IsUnifiedNumber, Is.True);
                Assert.That(value.GetUnifiedNumber(), Is.EqualTo(2.5));
            });

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
            Assert.That(date.GetUnifiedNumber(), Is.EqualTo(expectedSerial));
        }

        [Test]
        [SetCulture("cs-CZ")]
        public void ToString_RespectsCulture()
        {
            XLCellValue v = Blank.Value;
            Assert.That(v.ToString(), Is.Empty);

            v = true;
            Assert.That(v.ToString(), Is.EqualTo("TRUE"));

            v = 25.4;
            Assert.That(v.ToString(), Is.EqualTo("25,4"));

            v = "Hello";
            Assert.That(v.ToString(), Is.EqualTo("Hello"));

            v = XLError.IncompatibleValue;
            Assert.That(v.ToString(), Is.EqualTo("#VALUE!"));

            v = new DateTime(1900, 1, 2);
            Assert.That(v.ToString(), Is.EqualTo("02.01.1900 0:00:00"));

            v = new DateTime(1900, 3, 1, 4, 10, 5);
            Assert.That(v.ToString(), Is.EqualTo("01.03.1900 4:10:05"));

            v = new TimeSpan(4, 5, 6, 7, 82);
            Assert.That(v.ToString(), Is.EqualTo("101:06:07,082"));
        }

        [Test]
        public void TryConvert_Blank()
        {
            XLCellValue value = Blank.Value;
            Assert.That(value.TryConvert(out Blank blank), Is.True);
            Assert.That(blank, Is.EqualTo(Blank.Value));

            value = string.Empty;
            Assert.That(value.TryConvert(out blank), Is.True);
            Assert.That(blank, Is.EqualTo(Blank.Value));
        }

        [Test]
        public void TryConvert_Boolean()
        {
            XLCellValue value = true;
            Assert.That(value.TryConvert(out bool boolean), Is.True);
            Assert.That(boolean, Is.True);

            value = "True";
            Assert.That(value.TryConvert(out boolean), Is.True);
            Assert.That(boolean, Is.True);

            value = "False";
            Assert.That(value.TryConvert(out boolean), Is.True);
            Assert.False(boolean);

            value = 0;
            Assert.That(value.TryConvert(out boolean), Is.True);
            Assert.False(boolean);

            value = 0.001;
            Assert.That(value.TryConvert(out boolean), Is.True);
            Assert.That(boolean, Is.True);
        }

        [Test]
        public void TryConvert_Number()
        {
            var c = CultureInfo.GetCultureInfo("cs-CZ");
            XLCellValue value = 5;
            Assert.That(value.TryConvert(out double number, c), Is.True);
            Assert.That(number, Is.EqualTo(5.0));
            
            value = "1,5";
            Assert.Multiple(() =>
            {
                Assert.That(value.TryConvert(out number, c), Is.True);
                Assert.That(number, Is.EqualTo(1.5));
            });

            value = "1 1/4";
            Assert.Multiple(() =>
            {
                Assert.That(value.TryConvert(out number, c), Is.True);
                Assert.That(number, Is.EqualTo(1.25));
            });

            value = "3.1.1900";
            Assert.Multiple(() =>
            {
                Assert.That(value.TryConvert(out number, c), Is.True);
                Assert.That(number, Is.EqualTo(3));
            });

            value = true;
            Assert.Multiple(() =>
            {
                Assert.That(value.TryConvert(out number, c), Is.True);
                Assert.That(number, Is.EqualTo(1.0));
            });

            value = false;
            Assert.Multiple(() =>
            {
                Assert.That(value.TryConvert(out number, c), Is.True);
                Assert.That(number, Is.EqualTo(0.0));
            });

            value = new DateTime(2020, 4, 5, 10, 14, 5);
            Assert.Multiple(() =>
            {
                Assert.That(value.TryConvert(out number, c), Is.True);
                Assert.That(number, Is.EqualTo(43926.42644675926));
            });

            value = new TimeSpan(18, 0, 0);
            Assert.Multiple(() =>
            {
                Assert.That(value.TryConvert(out number, c), Is.True);
                Assert.That(number, Is.EqualTo(0.75));
            });
        }

        [Test]
        public void TryConvert_DateTime()
        {
            XLCellValue v = new DateTime(2020, 1, 1);
            Assert.That(v.TryConvert(out DateTime dt), Is.True);
            Assert.That(dt, Is.EqualTo(new DateTime(2020, 1, 1)));

            var lastSerialDate = 2958465;
            v = lastSerialDate;
            Assert.That(v.TryConvert(out dt), Is.True);
            Assert.That(dt, Is.EqualTo(new DateTime(9999, 12, 31)));

            v = lastSerialDate + 1;
            Assert.False(v.TryConvert(out dt));

            v = new TimeSpan(14, 0, 0, 0);
            Assert.That(v.TryConvert(out dt), Is.True);
            Assert.That(dt, Is.EqualTo(new DateTime(1900, 1, 14)));
        }

        [Test]
        public void TryConvert_TimeSpan()
        {
            var c = CultureInfo.GetCultureInfo("cs-CZ");
            XLCellValue v = new TimeSpan(10, 15, 30);
            Assert.That(v.TryConvert(out TimeSpan ts, c), Is.True);
            Assert.That(ts, Is.EqualTo(new TimeSpan(10, 15, 30)));

            v = "26:15:30,5";
            Assert.That(v.TryConvert(out ts, c), Is.True);
            Assert.That(ts, Is.EqualTo(new TimeSpan(1, 2, 15, 30, 500)));

            v = 0.75;
            Assert.That(v.TryConvert(out ts, c), Is.True);
            Assert.That(ts, Is.EqualTo(new TimeSpan(18, 0, 0)));
        }

        [TestCase(1)]
        [TestCase(10)] // microsecond
        [TestCase(3000000001)] // 5 min 1 tick
        public void TimeSpan_can_have_sub_millisecond_precision(long ticks)
        {
            var subMsTimeSpan = TimeSpan.FromTicks(ticks);
            XLCellValue value = subMsTimeSpan;
            Assert.That(value.GetTimeSpan(), Is.EqualTo(subMsTimeSpan));
        }

        [TestCase(1)]
        [TestCase(10)] // microsecond
        [TestCase(3000000001)] // 5 min 1 tick
        public void TimeSpan_with_sub_millisecond_precision_is_written_and_loaded_correctly(long ticks)
        {
            // NetFx converts double to string using G15. Core changed it to G17, but ClosedXML still use G15.
            var subMsTimeSpan = TimeSpan.FromTicks(ticks);
            TestHelper.CreateSaveLoadAssert(
                (_, ws) => { ws.Cell("A1").Value = subMsTimeSpan; },
                (_, ws) =>
                {
                    var cellValue = ws.Cell("A1").CachedValue;
                    Assert.That(cellValue.GetTimeSpan(), Is.EqualTo(subMsTimeSpan));
                });
        }

        [TestCase(long.MaxValue / (double)TimeSpan.TicksPerDay + 0.01)]
        [TestCase(long.MinValue / (double)TimeSpan.TicksPerDay - 0.01)]
        public void TimeSpan_throws_when_not_representable(double serialDateTime)
        {
            var value = XLCellValue.FromSerialTimeSpan(serialDateTime);
            var ex = Assert.Throws<OverflowException>(() => value.GetTimeSpan())!;
            Assert.That(ex.Message, Is.EqualTo("The serial date time value is too large to be represented in a TimeSpan."));
        }
    }
}