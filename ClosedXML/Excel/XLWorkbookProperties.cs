using System;

namespace ClosedXML.Excel
{
    public class XLWorkbookProperties
    {
        public XLWorkbookProperties()
        { 
            Company = null;
            Manager = null;
        }
        public String Author { get; set; }
        public String Title { get; set; }
        public String Subject { get; set; }
        public String Category { get; set; }
        public String Keywords { get; set; }
        public String Comments { get; set; }
        public String Status { get; set; }
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        public String LastModifiedBy { get; set; }
        public String Company { get; set; }
        public String Manager { get; set; }
    }
}
