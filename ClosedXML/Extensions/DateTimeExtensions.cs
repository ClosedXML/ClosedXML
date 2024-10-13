// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel;

internal static class DateTimeExtensions
{
    public static double ToSerialDateTime(this DateTime dateTime)
    {
        // Excel says 1900 was a leap year  :( Replicate an incorrect behavior thanks
        // to Lotus 1-2-3 decision from 1983...
        var oDate = dateTime.ToOADate();
        const int nonExistent1900Feb29SerialDate = 60;
        return oDate <= nonExistent1900Feb29SerialDate ? oDate - 1 : oDate;
    }
}
