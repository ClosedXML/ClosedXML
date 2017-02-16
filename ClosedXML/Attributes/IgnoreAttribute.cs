using System;

namespace ClosedXML.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class IgnoreAttribute : Attribute
    {
        public IgnoreAttribute()
            : this(true)
        { }

        public IgnoreAttribute(bool ignore)
        {
            this.Ignore = ignore;
        }

        public bool Ignore { get; private set; }
    }
}
