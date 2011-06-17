using System;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ClosedXML.Excel
{
    internal partial class XLCell
    {
        private String GetFieldName(Object[] customAttributes)
        {
            var displayAttributes = customAttributes.Where(a => a is DisplayAttribute).Select(a => (a as DisplayAttribute).Name);
            if (displayAttributes.Any())
                return displayAttributes.Single();
            else
                return null;
        }

    }
}
