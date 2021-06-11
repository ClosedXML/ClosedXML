using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel.Charts
{
    public abstract class ReferenceBase
    {
        public int ReferenceColumn { get; set; }
        public int ReferenceRow { get; set; }

        public abstract String Reference { get; }

        public String ReferenceTable { get; set; }

        public String ReferenceDirect { get; set; }

        public String ReferenceColumnName
        {
            get { return GetColumnName(ReferenceColumn); }
        }

        protected ReferenceBase()
        {
            ReferenceColumn = 1;
            ReferenceRow = 1;
            ReferenceTable = @"Tabelle1";
        }

        protected static String GetColumnName(int index)
        {
            if (index < 1)
                throw new ArgumentOutOfRangeException(@"index");

            const int firstChar = 65;

            string result = string.Empty;
            int value = index;

            while (value > 0)
            {
                int remainder = (value - 1) % 26;
                result = (char)(firstChar + remainder) + result;
                value = (int)(Math.Floor((double)((value - remainder) / 26)));
            }
            return result;
        }
    }
}
