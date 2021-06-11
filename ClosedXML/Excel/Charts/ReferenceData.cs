using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel.Charts
{
    public class ReferenceData : ReferenceBase
    {
        public String[] Values { get; set; }

        public int ReferenceRowEnd
        {
            get { return ReferenceRow + Values.Length - 1; }
        }

        public override String Reference
        {
            get
            {
                if (ReferenceDirect == null)
                    return ReferenceTable + '!' + ReferenceColumnName + ReferenceRow + ':' + ReferenceColumnName + ReferenceRowEnd;
                return ReferenceDirect;
            }
        }

        public ReferenceData()
        {
            ReferenceRow = 2;
        }
    }
}
