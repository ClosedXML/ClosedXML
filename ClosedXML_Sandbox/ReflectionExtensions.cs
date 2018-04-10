#if _NET40_
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace ClosedXML_Sandbox
{
    internal static class ReflectionExtensions
    {
        public static void SetValue(this PropertyInfo info, object obj, object value)
        {
            info.SetValue(obj, value, null);
        }
    }
}
#endif
