using System;

namespace ClosedXML.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class ColumnOrderAttribute : Attribute
    {
        public ColumnOrderAttribute(long order)
        {
            this.Order = order;
        }

        public long Order { get; private set; }
    }
}
