// Keep this file CodeMaid organised and cleaned
using System.Reflection;

namespace ClosedXML.Extensions
{
    internal static class ReflectionExtensions
    {
        public static bool IsStatic(this MemberInfo memberInfo) =>
            memberInfo switch
            {
                ConstructorInfo constructorInfo => constructorInfo.IsStatic,
                EventInfo eventInfo => eventInfo.GetAddMethod().IsStatic,
                FieldInfo fieldInfo => fieldInfo.IsStatic,
                MethodInfo methodInfo => methodInfo.IsStatic,
                PropertyInfo propertyInfo => propertyInfo.GetGetMethod().IsStatic,
                _ => false
            };
    }
}
