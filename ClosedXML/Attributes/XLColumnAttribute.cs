using System;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class XLColumnAttribute : Attribute
    {
        public string Header { get; set; }
        public bool Ignore { get; set; }
        public int Order { get; set; }

        private static XLColumnAttribute GetXLColumnAttribute(MemberInfo mi)
        {
            if (!mi.HasAttribute<XLColumnAttribute>())
            {
                return null;
            }

            return mi.GetAttributes<XLColumnAttribute>().First();
        }

        internal static string GetHeader(MemberInfo mi)
        {
            var attribute = GetXLColumnAttribute(mi);
            if (attribute == null)
            {
                return null;
            }

            return string.IsNullOrWhiteSpace(attribute.Header) ? null : attribute.Header;
        }

        internal static int GetOrder(MemberInfo mi)
        {
            var attribute = GetXLColumnAttribute(mi);
            if (attribute == null)
            {
                return int.MaxValue;
            }

            return attribute.Order;
        }

        internal static bool IgnoreMember(MemberInfo mi)
        {
            var attribute = GetXLColumnAttribute(mi);
            if (attribute == null)
            {
                return false;
            }

            return attribute.Ignore;
        }
    }
}