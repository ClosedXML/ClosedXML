using System;

namespace ClosedXML.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    [Obsolete("Use XLColumnAttribute instead")]
    public class ColumnOrderAttribute : Attribute
    {
        // Deprecated, use XLColumnAttribute instead

        // This attribute should be kept for "a while" to allow existing users to find the new XLColumnAttribute
        // Should be removed before ClosedXML turns version 1.0
    }
}
