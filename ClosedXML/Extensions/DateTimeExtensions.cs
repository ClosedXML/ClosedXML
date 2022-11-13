// Keep this file CodeMaid organised and cleaned
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClosedXML.Excel
{
    internal static class DateTimeExtensions
    {
        public static Double MaxOADate
        {
            get
            {
                return 2958465.99999999;
            }
        }

        public static bool IsWorkDay(this DateTime date, IEnumerable<DateTime> bankHolidays)
        {
            return date.DayOfWeek != DayOfWeek.Saturday
                   && date.DayOfWeek != DayOfWeek.Sunday
                   && !bankHolidays.Contains(date);
        }

        public static DateTime NextWorkday(this DateTime date, IEnumerable<DateTime> bankHolidays)
        {
            var nextDate = date.AddDays(1);
            while (!nextDate.IsWorkDay(bankHolidays))
                nextDate = nextDate.AddDays(1);

            return nextDate;
        }

        public static DateTime PreviousWorkDay(this DateTime date, IEnumerable<DateTime> bankHolidays)
        {
            var previousDate = date.AddDays(-1);
            while (!previousDate.IsWorkDay(bankHolidays))
                previousDate = previousDate.AddDays(-1);

            return previousDate;
        }

        public static double ToSerialDateTime(this DateTime dateTime)
        {
            // Excel says 1900 was a leap year  :( Replicate an incorrect behavior thanks
            // to Lotus 1-2-3 decision from 1983...
            var oDate = dateTime.ToOADate();
            const int nonExistent1900Feb29SerialDate = 60;
            return oDate <= nonExistent1900Feb29SerialDate ? oDate - 1 : oDate;
        }
    }
}
