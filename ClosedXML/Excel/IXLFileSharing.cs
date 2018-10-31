// Keep this file CodeMaid organised and cleaned
using System;

namespace ClosedXML.Excel
{
    public interface IXLFileSharing
    {
        //String AlgorithmName { get; set; }
        //Byte[] HashValue { get; set; }
        Boolean ReadOnlyRecommended { get; set; }

        //Byte[] ReservationPassword { get; set; }
        //Byte[] SaltValue { get; set; }
        //Int32 SpinCount { get; set; }
        String UserName { get; set; }
    }
}
