using System;
using System.Collections.Generic;
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
            //ce.RegisterFunction("DAYS360", 1, Days360); // Calculates the number of days between two dates based on a 360-day year
            //ce.RegisterFunction("EDATE", 1, Edate); // Returns the serial number of the date that is the indicated number of months before or after the start date
            //ce.RegisterFunction("EOMONTH", 1, Eomonth); // Returns the serial number of the last day of the month before or after a specified number of months
            ce.RegisterFunction("HOUR", 1, Hour); // Converts a serial number to an hour
            ce.RegisterFunction("MINUTE", 1, Minute); // Converts a serial number to a minute
            ce.RegisterFunction("MONTH", 1, Month); // Converts a serial number to a month
            //ce.RegisterFunction("NETWORKDAYS", 1, Networkdays); // Returns the number of whole workdays between two dates
            ce.RegisterFunction("NOW", 0, Now); // Returns the serial number of the current date and time
            ce.RegisterFunction("SECOND", 1, Second); // Converts a serial number to a second
            ce.RegisterFunction("TIME", 3, Time); // Returns the serial number of a particular time
            ce.RegisterFunction("TIMEVALUE", 1, Timevalue); // Converts a time in the form of text to a serial number
            ce.RegisterFunction("TODAY", 0, Today); // Returns the serial number of today's date
            //ce.RegisterFunction("WEEKDAY", 1, Weekday); // Converts a serial number to a day of the week
            //ce.RegisterFunction("WEEKNUM", 1, Weeknum); // Converts a serial number to a number representing where the week falls numerically with a year
            //ce.RegisterFunction("WORKDAY", 1, Workday); // Returns the serial number of the date before or after a specified number of workdays
            ce.RegisterFunction("YEAR", 1, Year); // Converts a serial number to a year
            //ce.RegisterFunction("YEARFRAC", 1, Yearfrac); // Returns the year fraction representing the number of whole days between start_date and end_date

        }

        static object Date(List<Expression> p)
        {
            var year = (int)p[0];
            var month = (int)p[1];
            var day = (int)p[2];

            return (int)Math.Floor(new DateTime(year, month, day).ToOADate());
        }

        static object Datevalue(List<Expression> p)
        {
            var date = (string)p[0];

            return (int)Math.Floor(DateTime.Parse(date).ToOADate());
        }

        static object Day(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return date.Day;
        }

        static object Month(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return date.Month;
        }

        static object Year(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return date.Year;
        }

        static object Minute(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return date.Minute;
        }

        static object Hour(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return date.Hour;
        }

        static object Second(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return date.Second;
        }

        static object Now(List<Expression> p)
        {
            return DateTime.Now;
        }

        static object Time(List<Expression> p)
        {
            var hour = (int)p[0];
            var minute = (int)p[1];
            var second = (int)p[2];
            
            return new TimeSpan(0, hour, minute, second);
        }

        static object Timevalue(List<Expression> p)
        {
            var date = (DateTime)p[0];

            return (DateTime.MinValue + date.TimeOfDay).ToOADate();
        }

        static object Today(List<Expression> p)
        {
            return DateTime.Now.Date;
        }
    }
}
