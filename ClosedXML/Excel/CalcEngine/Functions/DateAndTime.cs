#nullable disable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using static ClosedXML.Excel.CalcEngine.Functions.SignatureAdapter;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    internal static class DateAndTime
    {
        /// <summary>
        /// Serial date of 9999-12-31. Date is generally considered invalid, if above that or below 0.
        /// </summary>
        private const int Year10K = 2958465;

        public static void Register(FunctionRegistry ce)
        {
            ce.RegisterFunction("DATE", 3, Date); // Returns the serial number of a particular date
            ce.RegisterFunction("DATEDIF", 3, Datedif); // Calculates the number of days, months, or years between two dates
            ce.RegisterFunction("DATEVALUE", 1, Datevalue); // Converts a date in the form of text to a serial number
            ce.RegisterFunction("DAY", 1, Day); // Converts a serial number to a day of the month
            ce.RegisterFunction("DAYS", 2, Days); // Returns the number of days between two dates.
            ce.RegisterFunction("DAYS360", 2, 3, Days360); // Calculates the number of days between two dates based on a 360-day year
            ce.RegisterFunction("EDATE", 2, Edate); // Returns the serial number of the date that is the indicated number of months before or after the start date
            ce.RegisterFunction("EOMONTH", 2, Eomonth); // Returns the serial number of the last day of the month before or after a specified number of months
            ce.RegisterFunction("HOUR", 1, Hour); // Converts a serial number to an hour
            ce.RegisterFunction("ISOWEEKNUM", 1, IsoWeekNum); // Returns number of the ISO week number of the year for a given date.
            ce.RegisterFunction("MINUTE", 1, Minute); // Converts a serial number to a minute
            ce.RegisterFunction("MONTH", 1, Month); // Converts a serial number to a month
            ce.RegisterFunction("NETWORKDAYS", 2, 3, Networkdays, AllowRange.Only, 2); // Returns the number of whole workdays between two dates
            ce.RegisterFunction("NOW", 0, Now); // Returns the serial number of the current date and time
            ce.RegisterFunction("SECOND", 1, Second); // Converts a serial number to a second
            ce.RegisterFunction("TIME", 3, Time); // Returns the serial number of a particular time
            ce.RegisterFunction("TIMEVALUE", 1, Timevalue); // Converts a time in the form of text to a serial number
            ce.RegisterFunction("TODAY", 0, Today); // Returns the serial number of today's date
            ce.RegisterFunction("WEEKDAY", 1, 2, AdaptLastOptional(Weekday), FunctionFlags.Scalar, AllowRange.None); // Converts a serial number to a day of the week
            ce.RegisterFunction("WEEKNUM", 1, 2, Weeknum); // Converts a serial number to a number representing where the week falls numerically with a year
            ce.RegisterFunction("WORKDAY", 2, 3, Workday, AllowRange.Only, 2); // Returns the serial number of the date before or after a specified number of workdays
            ce.RegisterFunction("YEAR", 1, 1, Adapt(Year), FunctionFlags.Scalar); // Converts a serial number to a year
            ce.RegisterFunction("YEARFRAC", 2, 3, Yearfrac); // Returns the year fraction representing the number of whole days between start_date and end_date
        }

        /// <summary>
        /// Calculates number of business days, taking into account:
        ///  - weekends (Saturdays and Sundays)
        ///  - bank holidays in the middle of the week
        /// </summary>
        /// <param name="firstDay">First day in the time interval</param>
        /// <param name="lastDay">Last day in the time interval</param>
        /// <param name="bankHolidays">List of bank holidays excluding weekends</param>
        /// <returns>Number of business days during the 'span'</returns>
        private static int BusinessDaysUntil(DateTime firstDay, DateTime lastDay, IEnumerable<DateTime> bankHolidays)
        {
            firstDay = firstDay.Date;
            lastDay = lastDay.Date;
            if (firstDay > lastDay)
                return -BusinessDaysUntil(lastDay, firstDay, bankHolidays);

            TimeSpan span = lastDay - firstDay;
            int businessDays = span.Days + 1;
            int fullWeekCount = businessDays / 7;
            // find out if there are weekends during the time exceeding the full weeks
            if (businessDays > fullWeekCount * 7)
            {
                // we are here to find out if there is a 1-day or 2-days weekend
                // in the time interval remaining after subtracting the complete weeks
                var firstDayOfWeek = (int)firstDay.DayOfWeek;
                var lastDayOfWeek = (int)lastDay.DayOfWeek;
                if (lastDayOfWeek < firstDayOfWeek)
                    lastDayOfWeek += 7;
                if (firstDayOfWeek <= 6)
                {
                    if (lastDayOfWeek >= 7)// Both Saturday and Sunday are in the remaining time interval
                        businessDays -= 2;
                    else if (lastDayOfWeek >= 6)// Only Saturday is in the remaining time interval
                        businessDays -= 1;
                }
                else if (firstDayOfWeek <= 7 && lastDayOfWeek >= 7)// Only Sunday is in the remaining time interval
                    businessDays -= 1;
            }

            // subtract the weekends during the full weeks in the interval
            businessDays -= fullWeekCount + fullWeekCount;

            // subtract the number of bank holidays during the time interval
            foreach (var bh in bankHolidays)
            {
                if (firstDay <= bh && bh <= lastDay)
                    --businessDays;
            }

            return businessDays;
        }

        private static object Date(List<Expression> p)
        {
            var year = (int)p[0];
            var month = (int)p[1];
            var day = (int)p[2];

            // Excel allows months and days outside the normal range, and adjusts the date accordingly
            if (month > 12 || month < 1)
            {
                year += (int)Math.Floor((double)(month - 1d) / 12.0);
                month -= (int)Math.Floor((double)(month - 1d) / 12.0) * 12;
            }

            int daysAdjustment = 0;
            if (day > DateTime.DaysInMonth(year, month))
            {
                daysAdjustment = day - DateTime.DaysInMonth(year, month);
                day = DateTime.DaysInMonth(year, month);
            }
            else if (day < 1)
            {
                daysAdjustment = day - 1;
                day = 1;
            }

            return (int)Math.Floor(new DateTime(year, month, day).AddDays(daysAdjustment).ToOADate());
        }

        private static object Datedif(List<Expression> p)
        {
            DateTime startDate = p[0];
            DateTime endDate = p[1];
            string unit = p[2];

            if (startDate > endDate)
                return XLError.NumberInvalid;

            return (unit.ToUpper()) switch
            {
                "Y" => endDate.Year - startDate.Year - (new DateTime(startDate.Year, endDate.Month, endDate.Day) < startDate ? 1 : 0),
                "M" => Math.Truncate((endDate.Year - startDate.Year) * 12d + endDate.Month - startDate.Month - (endDate.Day < startDate.Day ? 1 : 0)),
                "D" => Math.Truncate(endDate.Date.Subtract(startDate.Date).TotalDays),

                // Microsoft discourages the use of the MD parameter
                // https://support.microsoft.com/en-us/office/datedif-function-25dba1a4-2812-480b-84dd-8b32a451b35c
                "MD" => (endDate.Day - startDate.Day + DateTime.DaysInMonth(startDate.Year, startDate.Month)) % DateTime.DaysInMonth(startDate.Year, startDate.Month),

                "YM" => (endDate.Month - startDate.Month + 12) % 12 - (endDate.Day < startDate.Day ? 1 : 0),
                "YD" => Math.Truncate(new DateTime(startDate.Year + (new DateTime(startDate.Year, endDate.Month, endDate.Day) < startDate ? 1 : 0), endDate.Month, endDate.Day).Subtract(startDate).TotalDays),
                _ => XLError.NumberInvalid,
            };
        }

        private static object Datevalue(List<Expression> p)
        {
            var date = (string)p[0];

            return (int)Math.Floor(DateTime.Parse(date).ToOADate());
        }

        private static object Day(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return date.Day;
        }

        private static object Days(List<Expression> p)
        {
            int end_date;
            var endDateValue = p[0].Evaluate();
            if (endDateValue is string)
                end_date = (int)Datevalue(new List<Expression>() { p[0] });
            else
                end_date = (int)p[0];

            int start_date;
            var startDateValue = p[1].Evaluate();
            if (startDateValue is string)
                start_date = (int)Datevalue(new List<Expression>() { p[1] });
            else
                start_date = (int)p[1];

            return end_date - start_date;
        }

        private static object Days360(List<Expression> p)
        {
            var date1 = (DateTime)p[0];
            var date2 = (DateTime)p[1];
            var isEuropean = p.Count == 3 ? p[2] : false;

            return Days360(date1, date2, isEuropean);
        }

        private static Int32 Days360(DateTime date1, DateTime date2, Boolean isEuropean)
        {
            var d1 = date1.Day;
            var m1 = date1.Month;
            var y1 = date1.Year;
            var d2 = date2.Day;
            var m2 = date2.Month;
            var y2 = date2.Year;

            if (isEuropean)
            {
                if (d1 == 31) d1 = 30;
                if (d2 == 31) d2 = 30;
            }
            else
            {
                if (d1 == 31) d1 = 30;
                if (d2 == 31 && d1 == 30) d2 = 30;
            }

            return 360 * (y2 - y1) + 30 * (m2 - m1) + (d2 - d1);
        }

        private static object Edate(List<Expression> p)
        {
            var date = (DateTime)p[0];
            var mod = (int)p[1];

            var retDate = date.AddMonths(mod);
            return retDate;
        }

        private static object Eomonth(List<Expression> p)
        {
            var start_date = (DateTime)p[0];
            var months = (int)p[1];

            var retDate = start_date.AddMonths(months);
            return new DateTime(retDate.Year, retDate.Month, DateTime.DaysInMonth(retDate.Year, retDate.Month));
        }

        private static Double GetYearAverage(DateTime date1, DateTime date2)
        {
            var daysInYears = new List<Int32>();
            for (int year = date1.Year; year <= date2.Year; year++)
                daysInYears.Add(DateTime.IsLeapYear(year) ? 366 : 365);
            return daysInYears.Average();
        }

        private static object Hour(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return date.Hour;
        }

        // http://stackoverflow.com/questions/11154673/get-the-correct-week-number-of-a-given-date
        private static object IsoWeekNum(List<Expression> p)
        {
            var date = (DateTime)p[0];

            // Seriously cheat.  If its Monday, Tuesday or Wednesday, then it'll
            // be the same week# as whatever Thursday, Friday or Saturday are,
            // and we always get those right
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(date);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday)
            {
                date = date.AddDays(3);
            }

            // Return the week of our adjusted day
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        private static object Minute(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return date.Minute;
        }

        private static object Month(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return date.Month;
        }

        private static object Networkdays(List<Expression> p)
        {
            var date1 = (DateTime)p[0];
            var date2 = (DateTime)p[1];
            var bankHolidays = new List<DateTime>();
            if (p.Count == 3)
            {
                var t = new Tally { p[2] };

                bankHolidays.AddRange(t.Select(XLHelper.GetDate));
            }

            return BusinessDaysUntil(date1, date2, bankHolidays);
        }

        private static object Now(List<Expression> p)
        {
            return DateTime.Now;
        }

        private static object Second(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return date.Second;
        }

        private static object Time(List<Expression> p)
        {
            var hour = (int)p[0];
            var minute = (int)p[1];
            var second = (int)p[2];

            return new TimeSpan(0, hour, minute, second);
        }

        private static object Timevalue(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return (DateTime.MinValue + date.TimeOfDay).ToOADate();
        }

        private static object Today(List<Expression> p)
        {
            return DateTime.Today;
        }

        private static ScalarValue Weekday(CalcContext ctx, ScalarValue date, ScalarValue flag)
        {
            if (!TryGetDate(ctx, date, out var serialDate, out var dateError))
                return dateError;

            var flagValue = 1d;
            if (!flag.IsBlank)
            {
                // Caller provided a value for optional parameter
                if (!flag.ToNumber(ctx.Culture).TryPickT0(out flagValue, out var flagError))
                    return flagError;
            }

            var result = Weekday(serialDate, (int)Math.Truncate(flagValue));

            if (!result.TryPickT0(out var weekday, out var weekdayError))
                return weekdayError;

            return weekday;
        }

        private static OneOf<int, XLError> Weekday(int serialDate, int startFlag)
        {
            // There are two offsets:
            // - what is the starting day
            // - how are days numbered (0-6, 1-7 ...)
            int? weekStartOffset = startFlag switch
            {
                1 => 0, // Sun
                2 => 6, // Mon
                3 => 6, // Mon
                11 => 6, // Mon
                12 => 5, // Tue
                13 => 4, // Wed
                14 => 3, // Thu
                15 => 2, // Fri
                16 => 1, // Sat
                17 => 0, // Sunday
                _ => null,
            };
            if (weekStartOffset is null)
                return XLError.NumberInvalid;

            var numberOffset = startFlag == 3 ? 0 : 1;

            // Because we don't go below 1900, there is no need to deal with UTC vs Gregorian calendar.
            // It is affected by 1900 bug, so no accurate weekdays before 1900-02-29. It was Wednesday BTW :)
            var weekday = (serialDate + 6 + weekStartOffset.Value) % 7 + numberOffset;
            return weekday;
        }

        private static object Weeknum(List<Expression> p)
        {
            var date = (DateTime)p[0];
            var retType = p.Count == 2 ? (int)p[1] : 1;

            DayOfWeek dayOfWeek = retType == 1 ? DayOfWeek.Sunday : DayOfWeek.Monday;
            var cal = new GregorianCalendar(GregorianCalendarTypes.Localized);
            var val = cal.GetWeekOfYear(date, CalendarWeekRule.FirstDay, dayOfWeek);

            return val;
        }

        private static object Workday(List<Expression> p)
        {
            var startDate = (DateTime)p[0];
            var daysRequired = (int)p[1];

            if (daysRequired == 0) return startDate;

            var bankHolidays = new List<DateTime>();
            if (p.Count == 3)
            {
                var t = new Tally { p[2] };

                bankHolidays.AddRange(t.Select(XLHelper.GetDate));
            }
            var testDate = startDate.AddDays(((daysRequired / 7) + 2) * 7 * Math.Sign(daysRequired));
            var return_date = Workday(startDate, testDate, daysRequired, bankHolidays);
            if (Math.Sign(daysRequired) == 1)
                return_date = return_date.NextWorkday(bankHolidays);
            else
                return_date = return_date.PreviousWorkDay(bankHolidays);

            return return_date;
        }

        private static DateTime Workday(DateTime startDate, DateTime testDate, int daysRequired, IReadOnlyCollection<DateTime> bankHolidays)
        {
            var businessDays = BusinessDaysUntil(startDate, testDate, bankHolidays);
            if (businessDays == daysRequired)
                return testDate;

            int days = businessDays > daysRequired ? -1 : 1;

            return Workday(startDate, testDate.AddDays(days), daysRequired, bankHolidays);
        }

        private static AnyValue Year(double serialDateTime)
        {
            serialDateTime = Math.Truncate(serialDateTime);
            if (serialDateTime < 0)
                return XLError.NumberInvalid;

            // Serial date time values from [0, 1) are from 1899-12-31,
            // but Excel represents them as 1900-01-00.
            if (serialDateTime < 1)
                return 1900;

            return serialDateTime.ToSerialDateTime().Year;
        }

        private static object Yearfrac(List<Expression> p)
        {
            var date1 = (DateTime)p[0];
            var date2 = (DateTime)p[1];
            var option = p.Count == 3 ? (int)p[2] : 0;

            if (option == 0)
                return Days360(date1, date2, false) / 360.0;
            if (option == 1)
                return Math.Floor((date2 - date1).TotalDays) / GetYearAverage(date1, date2);
            if (option == 2)
                return Math.Floor((date2 - date1).TotalDays) / 360.0;
            if (option == 3)
                return Math.Floor((date2 - date1).TotalDays) / 365.0;

            return Days360(date1, date2, true) / 360.0;
        }

        private static bool TryGetDate(CalcContext ctx, ScalarValue value, out int serialDate, out XLError error)
        {
            if (!value.ToNumber(ctx.Culture).TryPickT0(out var serialDateTime, out error))
            {
                serialDate = default;
                return false;
            }

            if (serialDateTime is < 0 or > Year10K)
            {
                serialDate = default;
                error = XLError.NumberInvalid;
                return false;
            }

            serialDate = (int)Math.Truncate(serialDateTime);
            error = default;
            return true;
        }
    }
}
