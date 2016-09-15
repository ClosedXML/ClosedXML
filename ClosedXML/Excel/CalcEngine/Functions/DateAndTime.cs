﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace ClosedXML.Excel.CalcEngine.Functions
{
    internal static class DateAndTime
    {
        public static void Register(CalcEngine ce)
        {
            ce.RegisterFunction("DATE", 3, Date); // Returns the serial number of a particular date
            ce.RegisterFunction("DATEVALUE", 1, Datevalue); // Converts a date in the form of text to a serial number
            ce.RegisterFunction("DAY", 1, Day); // Converts a serial number to a day of the month
            ce.RegisterFunction("DAYS360", 2, 3, Days360); // Calculates the number of days between two dates based on a 360-day year
            ce.RegisterFunction("EDATE", 2, Edate); // Returns the serial number of the date that is the indicated number of months before or after the start date
            ce.RegisterFunction("EOMONTH", 2, Eomonth); // Returns the serial number of the last day of the month before or after a specified number of months
            ce.RegisterFunction("HOUR", 1, Hour); // Converts a serial number to an hour
            ce.RegisterFunction("MINUTE", 1, Minute); // Converts a serial number to a minute
            ce.RegisterFunction("MONTH", 1, Month); // Converts a serial number to a month
            ce.RegisterFunction("NETWORKDAYS", 2, 3, Networkdays); // Returns the number of whole workdays between two dates
            ce.RegisterFunction("NOW", 0, Now); // Returns the serial number of the current date and time
            ce.RegisterFunction("SECOND", 1, Second); // Converts a serial number to a second
            ce.RegisterFunction("TIME", 3, Time); // Returns the serial number of a particular time
            ce.RegisterFunction("TIMEVALUE", 1, Timevalue); // Converts a time in the form of text to a serial number
            ce.RegisterFunction("TODAY", 0, Today); // Returns the serial number of today's date
            ce.RegisterFunction("WEEKDAY", 1, 2, Weekday); // Converts a serial number to a day of the week
            ce.RegisterFunction("WEEKNUM", 1, 2, Weeknum); // Converts a serial number to a number representing where the week falls numerically with a year
            ce.RegisterFunction("WORKDAY", 2, 3, Workday); // Returns the serial number of the date before or after a specified number of workdays
            ce.RegisterFunction("YEAR", 1, Year); // Converts a serial number to a year
            ce.RegisterFunction("YEARFRAC", 2, 3, Yearfrac); // Returns the year fraction representing the number of whole days between start_date and end_date

        }

        private static object Date(List<Expression> p)
        {
            var year = (int) p[0];
            var month = (int) p[1];
            var day = (int) p[2];

            return (int) Math.Floor(new DateTime(year, month, day).ToOADate());
        }

        private static object Datevalue(List<Expression> p)
        {
            var date = (string) p[0];

            return (int) Math.Floor(DateTime.Parse(date).ToOADate());
        }

        private static object Day(List<Expression> p)
        {
            var date = (DateTime) p[0];

            return date.Day;
        }

        private static object Month(List<Expression> p)
        {
            var date = (DateTime) p[0];

            return date.Month;
        }

        private static object Year(List<Expression> p)
        {
            var date = (DateTime) p[0];

            return date.Year;
        }

        private static object Minute(List<Expression> p)
        {
            var date = (DateTime) p[0];

            return date.Minute;
        }

        private static object Hour(List<Expression> p)
        {
            var date = (DateTime) p[0];

            return date.Hour;
        }

        private static object Second(List<Expression> p)
        {
            var date = (DateTime) p[0];

            return date.Second;
        }

        private static object Now(List<Expression> p)
        {
            return DateTime.Now;
        }

        private static object Time(List<Expression> p)
        {
            var hour = (int) p[0];
            var minute = (int) p[1];
            var second = (int) p[2];

            return new TimeSpan(0, hour, minute, second);
        }

        private static object Timevalue(List<Expression> p)
        {
            var date = (DateTime) p[0];

            return (DateTime.MinValue + date.TimeOfDay).ToOADate();
        }

        private static object Today(List<Expression> p)
        {
            return DateTime.Now.Date;
        }

        private static object Days360(List<Expression> p)
        {
            var date1 = (DateTime) p[0];
            var date2 = (DateTime) p[1];
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
            var date = (DateTime)p[0];
            var mod = (int)p[1];

            var retDate = date.AddMonths(mod);
            return new DateTime(retDate.Year, retDate.Month, 1).AddMonths(1).AddDays(-1);
        }

        private static object Networkdays(List<Expression> p)
        {
            var date1 = (DateTime)p[0];
            var date2 = (DateTime)p[1];
            var bankHolidays = new List<DateTime>();
            if (p.Count == 3)
            {
                var t = new Tally {p[2]};

                bankHolidays.AddRange(t.Select(XLHelper.GetDate));
            }

            return BusinessDaysUntil(date1, date2, bankHolidays);
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
                throw new ArgumentException("Incorrect last day " + lastDay);

            TimeSpan span = lastDay - firstDay;
            int businessDays = span.Days + 1;
            int fullWeekCount = businessDays / 7;
            // find out if there are weekends during the time exceedng the full weeks
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

        private static object Weekday(List<Expression> p)
        {
            var dayOfWeek = (int)((DateTime)p[0]).DayOfWeek;
            var retType = p.Count == 2 ? (int)p[1] : 1;

            if (retType == 2) return dayOfWeek;
            if (retType == 1) return dayOfWeek + 1;

            return dayOfWeek - 1;
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
            if (daysRequired <  0) throw new ArgumentOutOfRangeException("DaysRequired must be >= 0.");

            var bankHolidays = new List<DateTime>();
            if (p.Count == 3)
            {
                var t = new Tally { p[2] };

                bankHolidays.AddRange(t.Select(XLHelper.GetDate));
            }
            var testDate = startDate.AddDays(((daysRequired / 7) + 2) * 7);
            return Workday(startDate, testDate, daysRequired, bankHolidays).NextWorkday(bankHolidays);
        }

        private static DateTime Workday(DateTime startDate, DateTime testDate, int daysRequired, IEnumerable<DateTime> bankHolidays)
        {
            
            var businessDays = BusinessDaysUntil(startDate, testDate, bankHolidays);
            if (businessDays == daysRequired)
                return testDate;

            int days = businessDays > daysRequired ? -1 : 1;

            return Workday(startDate, testDate.AddDays(days), daysRequired, bankHolidays);
        }

        private static object Yearfrac(List<Expression> p)
        {
            var date1 = (DateTime) p[0];
            var date2 = (DateTime) p[1];
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

        private static Double GetYearAverage(DateTime date1, DateTime date2)
        {
            var daysInYears = new List<Int32>();
            for (int year = date1.Year; year <= date2.Year; year++)
                daysInYears.Add(DateTime.IsLeapYear(year) ? 366 : 365);
            return daysInYears.Average();
        }
    }

}