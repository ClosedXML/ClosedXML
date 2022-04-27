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
        public string Author { get; set; }
        public string Title { get; set; }
        public string Subject { get; set; }
        public string Category { get; set; }
        public string Keywords { get; set; }
        public string Comments { get; set; }
        public string Status { get; set; }
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        public string LastModifiedBy { get; set; }
        public string Company { get; set; }
        public string Manager { get; set; }
    }
}
