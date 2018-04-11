using ClosedXML.Excel;
using System;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class XLColumnAttribute : Attribute
    {
        public String Header { get; set; }
        public Boolean Ignore { get; set; }
        public Int32 Order { get; set; }

        private static XLColumnAttribute GetXLColumnAttribute(MemberInfo mi)
        {
            if (!mi.HasAttribute<XLColumnAttribute>()) return null;
            return mi.GetAttributes<XLColumnAttribute>().First();
        }

        internal static String GetHeader(MemberInfo mi)
        {
            var attribute = GetXLColumnAttribute(mi);
            if (attribute == null) return null;
            return String.IsNullOrWhiteSpace(attribute.Header) ? null : attribute.Header;
        }

        internal static Int32 GetOrder(MemberInfo mi)
        {
            var attribute = GetXLColumnAttribute(mi);
            if (attribute == null) return Int32.MaxValue;
            return attribute.Order;
        }

        internal static Boolean IgnoreMember(MemberInfo mi)
        {
            var attribute = GetXLColumnAttribute(mi);
            if (attribute == null) return false;
            return attribute.Ignore;
        }
    }
}
