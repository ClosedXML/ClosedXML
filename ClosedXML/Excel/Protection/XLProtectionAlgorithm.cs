// Keep this file CodeMaid organised and cleaned
using System.ComponentModel;

namespace ClosedXML.Excel
{
    public static class XLProtectionAlgorithm
    {
        public const Algorithm DefaultProtectionAlgorithm = Algorithm.SimpleHash;

        public enum Algorithm
        {
            // The default hashing algorithm as described in http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/
            [Description("SimpleHash")]
            SimpleHash,

            //[Description("MD2")]
            //MD2,

            //[Description("MD4")]
            //MD4,

            //[Description("MD5")]
            //MD5,

            //[Description("RIPEMD-128")]
            //RIPEMD128,

            //[Description("RIPEMD-160")]
            //RIPEMD160,

            //[Description("SHA-1")]
            //SHA1,

            //[Description("SHA-256")]
            //SHA256,

            //[Description("SHA-384")]
            //SHA384,

            [Description("SHA-512")]
            SHA512,

            //[Description("WHIRLPOOL")]
            //WHIRLPOOL
        }
    }
}
