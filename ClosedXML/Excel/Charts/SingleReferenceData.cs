using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXML.Excel.Charts
{
    public class SingleReferenceData : ReferenceBase
    {
        public override String Reference
        {
            get
            {
                if (ReferenceDirect == null)
                    return ReferenceTable + '!' + ReferenceColumnName + ReferenceRow;
                return ReferenceDirect;
            }
        }

        public String Value { get; set; }

        public SingleReferenceData(String value)
        {
            Value = value;
        }
    }
}
