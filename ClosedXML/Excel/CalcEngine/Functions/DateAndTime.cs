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
            ce.RegisterFunction("DATE", 3, 3, Adapt(Date), FunctionFlags.Scalar); // Returns the serial number of a particular date
            ce.RegisterFunction("DATEDIF", 3, Datedif); // Calculates the number of days, months, or years between two dates
            ce.RegisterFunction("DATEVALUE", 1, Datevalue); // Converts a date in the form of text to a serial number
            ce.RegisterFunction("DAY", 1, 1, Adapt(Day), FunctionFlags.Scalar); // Converts a serial number to a day of the month
            ce.RegisterFunction("DAYS", 2, Days); // Returns the number of days between two dates.
            ce.RegisterFunction("DAYS360", 2, 3, AdaptLastOptional(Days360, false), FunctionFlags.Scalar); // Calculates the number of days between two dates based on a 360-day year
            ce.RegisterFunction("EDATE", 2, Edate); // Returns the serial number of the date that is the indicated number of months before or after the start date
            ce.RegisterFunction("EOMONTH", 2, Eomonth); // Returns the serial number of the last day of the month before or after a specified number of months
            ce.RegisterFunction("HOUR", 1, 1, Adapt(Hour), FunctionFlags.Scalar); // Converts a serial number to an hour
            ce.RegisterFunction("ISOWEEKNUM", 1, IsoWeekNum); // Returns number of the ISO week number of the year for a given date.
            ce.RegisterFunction("MINUTE", 1, 1, Adapt(Minute), FunctionFlags.Scalar); // Converts a serial number to a minute
            ce.RegisterFunction("MONTH", 1, 1, Adapt(Month), FunctionFlags.Scalar); // Converts a serial number to a month
            ce.RegisterFunction("NETWORKDAYS", 2, 3, AdaptLastOptional(NetWorkDays), FunctionFlags.Range, AllowRange.Only, 2); // Returns the number of whole workdays between two dates
            ce.RegisterFunction("NOW", 0, 0, Adapt(Now), FunctionFlags.Scalar | FunctionFlags.Volatile); // Returns the serial number of the current date and time
            ce.RegisterFunction("SECOND", 1, 1, Adapt(Second), FunctionFlags.Scalar); // Converts a serial number to a second
            ce.RegisterFunction("TIME", 3, 3, Adapt(Time), FunctionFlags.Scalar); // Returns the serial number of a particular time
            ce.RegisterFunction("TIMEVALUE", 1, Timevalue); // Converts a time in the form of text to a serial number
            ce.RegisterFunction("TODAY", 0, 0, Adapt(Today), FunctionFlags.Scalar | FunctionFlags.Volatile); // Returns the serial number of today's date
            ce.RegisterFunction("WEEKDAY", 1, 2, AdaptLastOptional(Weekday), FunctionFlags.Scalar, AllowRange.None); // Converts a serial number to a day of the week
            ce.RegisterFunction("WEEKNUM", 1, 2, Weeknum); // Converts a serial number to a number representing where the week falls numerically with a year
            ce.RegisterFunction("WORKDAY", 2, 3, AdaptLastOptional(Workday), FunctionFlags.Range, AllowRange.Only, 2); // Returns the serial number of the date before or after a specified number of workdays
            ce.RegisterFunction("YEAR", 1, 1, Adapt(Year), FunctionFlags.Scalar); // Converts a serial number to a year
            ce.RegisterFunction("YEARFRAC", 2, 3, Yearfrac); // Returns the year fraction representing the number of whole days between start_date and end_date
        }

        private static int BusinessDaysUntil(int firstDay, int lastDay, ICollection<int> distinctHolidays)
        {
            if (firstDay > lastDay)
                return -BusinessDaysUntil(lastDay, firstDay, distinctHolidays);

            var workDays = lastDay - firstDay + 1;
            var fullWeekCount = Math.DivRem(workDays, 7, out var remainingDays);

            // find out if there are weekends during the time exceeding the full weeks
            for (var day = lastDay - remainingDays + 1; day <= lastDay; ++day)
            {
                if (IsWeekend(day))
                    workDays--;
            }

            // subtract the weekends during the full weeks in the interval
            workDays -= fullWeekCount * 2;

            // subtract the number of bank holidays during the time interval
            foreach (var holidayDate in distinctHolidays)
            {
                if (firstDay <= holidayDate && holidayDate <= lastDay)
                {
                    if (!IsWeekend(holidayDate))
                        --workDays;
                }
            }

            return workDays;
        }

        private static ScalarValue Date(CalcContext ctx, double year, double month, double day)
        {
            // Unlike most functions, values are floored - not truncated.
            year = Math.Floor(year);
            month = Math.Floor(month);
            day = Math.Floor(day);

            if (month is < -short.MaxValue or >= short.MaxValue)
                return XLError.NumberInvalid;

            // Excel behaves out of spec. Spec says 0-99 are interpreted as year + 1900,
            // but reality is that anything below 1900 is interpreted as year + 1900.
            // That seems to be true for both 1900 and 1904 date systems.
            if (year < 1900)
                year += 1900;

            // Excel buggy implementation :) Should probably return error,
            // but silently changes the result instead.
            day = Math.Min(day, short.MaxValue);
            if (day < short.MinValue)
                day = short.MaxValue;

            if (year > 10000)
                year = 10000;

            // Excel allows months and days outside the normal range, and adjusts the date
            // accordingly.
            var yearAdjustment = Math.Floor((month - 1d) / 12.0);
            year += yearAdjustment;
            month -= yearAdjustment * 12;

            // Year 1 is earliest allowable in both date system. Also avoid the double
            // to int conversion problems when double is too small.
            if (year < 1)
                return XLError.NumberInvalid;

            var startOfMonth = new DateTime((int)year, (int)month, 1).ToSerialDateTime();
            var serialDate = startOfMonth + day - 1;
            if (serialDate < 0 || serialDate >= ctx.DateSystemUpperLimit)
                return XLError.NumberInvalid;

            return serialDate;
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

        private static ScalarValue Day(CalcContext ctx, double serialDateTime)
        {
            if (!TryGetDate(ctx, serialDateTime, out var serialDate))
                return XLError.NumberInvalid;

            return DateParts.From(ctx, serialDate).Day;
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

        private static ScalarValue Days360(CalcContext ctx, double startDateTime, double endDateTime, bool isEuropean)
        {
            if (!TryGetDate(ctx, startDateTime, out var startDate))
                return XLError.NumberInvalid;

            if (!TryGetDate(ctx, endDateTime, out var endDate))
                return XLError.NumberInvalid;

            return Days360(ctx, startDate, endDate, isEuropean);
        }

        private static int Days360(CalcContext ctx, int startSerialDate, int endSerialDate, bool isEuropean)
        {
            var startDate = DateParts.From(ctx, startSerialDate);
            var (startYear, startMonth, startDay) = startDate;
            var (endYear, endMonth, endDay) = DateParts.From(ctx, endSerialDate);

            if (isEuropean)
            {
                if (startDay == 31)
                    startDay = 30;

                if (endDay == 31)
                    endDay = 30;
            }
            else
            {
                // There are several descriptions of the US algorithm: spec, wikipedia, function help,
                // ODF. Out of these, only ODF is correct (rest is incomplete/has different results).
                if (startDate.IsLastDayOfMonth())
                    startDay = 30;

                if (endDay == 31 && startDay == 30)
                    endDay = 30;
            }

            return 360 * (endYear - startYear) + 30 * (endMonth - startMonth) + (endDay - startDay);
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

        private static ScalarValue Hour(CalcContext ctx, double serialTime)
        {
            return GetTimeComponent(ctx, serialTime, static d => d.Hour);
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

        private static bool IsWeekend(int date)
        {
            return WeekdayCalc(date) is 1 or 7;
        }

        private static ScalarValue Minute(CalcContext ctx, double serialTime)
        {
            return GetTimeComponent(ctx, serialTime, static d => d.Minute);
        }

        private static ScalarValue Month(CalcContext ctx, double serialDateTime)
        {
            if (!TryGetDate(ctx, serialDateTime, out var serialDate))
                return XLError.NumberInvalid;

            return DateParts.From(ctx, serialDate).Month;
        }

        private static ScalarValue NetWorkDays(CalcContext ctx, ScalarValue startDate, ScalarValue endDate, AnyValue holidays)
        {
            if (!TryGetDate(ctx, startDate, out var startSerialDate, out var startDateError))
                return startDateError;

            if (!TryGetDate(ctx, endDate, out var endSerialDate, out var endDateError))
                return endDateError;

            // Use set to skip duplicate values
            var allHolidays = new HashSet<int>();
            foreach (var holidayValue in ctx.GetNonBlankValues(holidays))
            {
                if (!TryGetDate(ctx, holidayValue, out var holidayDate, out var error))
                    return error;

                allHolidays.Add(holidayDate);
            }

            return BusinessDaysUntil(startSerialDate, endSerialDate, allHolidays);
        }

        private static ScalarValue Now()
        {
            return DateTime.Now.ToSerialDateTime();
        }

        private static ScalarValue Second(CalcContext ctx, double serialDate)
        {
            return GetTimeComponent(ctx, serialDate, static d => d.Second);
        }

        private static ScalarValue Time(CalcContext ctx, double hour, double minute, double second)
        {
            if (!TryGetComponent(hour, out var hourFloored))
                return XLError.NumberInvalid;

            if (!TryGetComponent(minute, out var minuteFloored))
                return XLError.NumberInvalid;

            if (!TryGetComponent(second, out var secondFloored))
                return XLError.NumberInvalid;

            var serialDate = new TimeSpan(hourFloored, minuteFloored, secondFloored).ToSerialDateTime();
            return serialDate % 1.0;

            static bool TryGetComponent(double value, out int truncated)
            {
                value = Math.Floor(value);
                if (value is < 0 or > 32767)
                {
                    truncated = default;
                    return false;
                }

                truncated = checked((int)value);
                return true;
            }
        }

        private static object Timevalue(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return (DateTime.MinValue + date.TimeOfDay).ToOADate();
        }

        private static ScalarValue Today()
        {
            return DateTime.Today.ToSerialDateTime();
        }

        private static ScalarValue Weekday(CalcContext ctx, ScalarValue date, ScalarValue flag)
        {
            if (!TryGetDate(ctx, date, out var serialDate, out var dateError, acceptLogical: true))
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
            var weekday = WeekdayCalc(serialDate, weekStartOffset.Value, numberOffset);
            return weekday;
        }

        /// <summary>
        /// Calculate week day. No checks. The default form is form 3 (week starts at Sun, range 1..7). 
        /// </summary>
        private static int WeekdayCalc(int serialDate, int weekStartOffset = 0, int numberOffset = 1)
        {
            return (serialDate + 6 + weekStartOffset) % 7 + numberOffset;
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

        private static ScalarValue Workday(CalcContext ctx, ScalarValue startDateScalar, ScalarValue dayOffsetValue, AnyValue holidays)
        {
            if (!TryGetDate(ctx, startDateScalar, out var startDate, out var startDateError))
                return startDateError;

            if (!dayOffsetValue.ToNumber(ctx.Culture).TryPickT0(out var dayOffsetDouble, out var dayOffsetError))
                return dayOffsetError;

            var dayOffset = (int)Math.Truncate(dayOffsetDouble);

            // When offset is zero, return the startDate, regardless if it is Saturday or Sunday.
            if (dayOffset == 0)
                return startDate;

            var cmp = dayOffset > 0 ? Comparer<int>.Default : Comparer<int>.Create(static (x, y) => y.CompareTo(x));
            var oneDay = dayOffset > 0 ? 1 : -1; // One day in a specified direction

            if (!GetHolidays(cmp).TryPickT0(out var orderedHolidays, out var holidaysError))
                return holidaysError;

            // The algorithm should count workdays for each segment between holiday days
            // and sum them up.

            // A date up to which we have counted workdays. It is inclusive, so if
            // the lastDateSoFar is a workday, it is already counted in workdaysSoFar.
            var lastDateSoFar = startDate;

            // Number of workdays that have already been processed from startDate up to
            // the lastDateSoFar (inclusive).
            var workdaysSoFar = 0;
            var startIsHoliday = orderedHolidays.Count > 0 && orderedHolidays[0] == startDate;
            for (var i = startIsHoliday ? 1 : 0; i < orderedHolidays.Count; ++i)
            {
                var holidayDate = orderedHolidays[i];

                // Because workdays up to and including lastDateSoFar has already been counted, we add + 1.
                // The holidayDate is not a Saturday or Sunday (it has been filtered out).
                // Because we know there is no holiday between lastDateSoFar (which might have been
                // a holiday or not) + 1 (by adding 1, we are sure it's not counted).
                // We are counting up to and including holidayDate. We can't use `holidayDate-1`, because
                // if two holidays were next to each other, the `holidayDate-1` might be *before* `lastDateSoFar+1`.
                // When days are same, BusinessDaysUntil returns 1 regardless of direction, so add a condition.
                var segmentWorkdays = lastDateSoFar + oneDay != holidayDate
                    ? BusinessDaysUntil(lastDateSoFar + oneDay, holidayDate, System.Array.Empty<int>())
                    : oneDay;

                if (cmp.Compare(workdaysSoFar + segmentWorkdays, dayOffset) > 0)
                {
                    // We know that the target day for desired dayOffset is in this segment.
                    // Possibly at the start, in the middle or at the end. Because of it,
                    // we know that there are no holidays from `lastDateSoFar..{resultDate}`.
                    break;
                }

                // The segment workdays include holidayDate as a workday, use -1 so it is not counted.
                workdaysSoFar += segmentWorkdays - oneDay;
                lastDateSoFar = holidayDate;
            }

            // At this point, we can just have to find the target date without any interference from holidays.
            var remainingWorkdays = dayOffset - workdaysSoFar;
            var weekCount = Math.DivRem(remainingWorkdays, 5, out var remaining);

            // When we start on Sunday and want 5 dayOffset, ensure that we end up on friday of same week, not Sunday.
            if (remaining == 0)
            {
                // We know that weekCount is at least 1, so decreasing one won't go negative.
                weekCount -= oneDay;
                remaining += oneDay * 5;
            }

            var workday = lastDateSoFar + weekCount * 7;
            while (remaining != 0)
            {
                do
                {
                    workday += oneDay;
                } while (IsWeekend(workday));
                remaining -= oneDay;
            }

            return workday;

            OneOf<List<int>, XLError> GetHolidays(IComparer<int> comparer)
            {
                // Use set to skip duplicate values
                var distinctHolidays = new HashSet<int>();
                foreach (var holidayValue in ctx.GetNonBlankValues(holidays))
                {
                    if (!TryGetDate(ctx, holidayValue, out var holidayDate, out var error))
                        return error;

                    if (comparer.Compare(holidayDate, startDate) < 0)
                        continue;

                    if (IsWeekend(holidayDate))
                        continue;

                    distinctHolidays.Add(holidayDate);
                }

                // Distinct, ordered holidays during a workweek
                var sortedHolidays = new List<int>(distinctHolidays);
                sortedHolidays.Sort(comparer);
                return sortedHolidays;
            }
        }

        private static ScalarValue Year(CalcContext ctx, double serialDateTime)
        {
            if (!TryGetDate(ctx, serialDateTime, out var serialDate))
                return XLError.NumberInvalid;

            return DateParts.From(ctx, serialDate).Year;
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

        private static bool TryGetDate(CalcContext ctx, ScalarValue value, out int serialDate, out XLError error, bool acceptLogical = false)
        {
            // For some reason, Excel dislikes logical for things that are "date" in functions.
            // Someone likely though it was a good idea 40yrs ago.
            if (value.IsLogical && !acceptLogical)
            {
                serialDate = default;
                error = XLError.IncompatibleValue;
                return false;
            }

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

        private static bool TryGetDate(CalcContext ctx, double serialDateTime, out int serialDate)
        {
            if (serialDateTime < 0 || serialDateTime >= ctx.DateSystemUpperLimit)
            {
                serialDate = default;
                return false;
            }

            serialDate = checked((int)Math.Truncate(serialDateTime));
            return true;
        }

        private static ScalarValue GetTimeComponent(CalcContext ctx, double serialTime, Func<DateTime, int> component)
        {
            if (serialTime < 0 || serialTime >= ctx.DateSystemUpperLimit)
                return XLError.NumberInvalid;

            return component(DateTime.FromOADate(serialTime));
        }

        /// <summary>
        /// A date type unconstrained by DateTime limitations (1900-01-00 or 1900-02-29).
        /// </summary>
        private readonly record struct DateParts(int Year, int Month, int Day)
        {
            private static readonly DateParts Epoch1900 = new(1900, 1, 0);

            private static readonly DateParts Feb29 = new(1900, 2, 29);

            internal static DateParts From(CalcContext ctx, int serialDate)
            {
                if (ctx.Use1904DateSystem)
                {
                    var date1904 = DateTime.FromOADate(serialDate + 1462);
                    return From(date1904);
                }

                // Return value for 1900-01-00
                if (serialDate == 0)
                    return Epoch1900;

                // Everyone loves 29th Feb 1900
                if (serialDate == 60)
                    return Feb29;

                // January and February 1900. Because of non-existent feb29, adjust by one day
                if (serialDate < 60)
                    serialDate++;

                return From(DateTime.FromOADate(serialDate));
            }

            private static DateParts From(DateTime date)
            {
                return new DateParts(date.Year, date.Month, date.Day);
            }

            internal bool IsLastDayOfMonth()
            {
                // 1900-02-29 is last day of a month per Excel, thus we have:
                // * return true for that date
                if (this == Feb29)
                    return true;

                // * can't return true for real end of month
                if (Year == 1900 && Month == 2 && Day == 28)
                    return false;

                return Day == DateTime.DaysInMonth(Year, Month);
            }
        }
    }
}
