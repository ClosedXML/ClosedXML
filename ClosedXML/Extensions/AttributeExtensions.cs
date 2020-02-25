// Keep this file CodeMaid organised and cleaned
using System;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace ClosedXML
{
    public static class AttributeExtensions
    {
        public static TAttribute[] GetAttributes<TAttribute>(
            this MemberInfo member)
            where TAttribute : Attribute
        {
            var attributes = member.GetCustomAttributes(typeof(TAttribute), true);

            return (TAttribute[])attributes;
        }

        public static MethodInfo GetMethod<T>(this T instance, Expression<Func<T, object>> methodSelector)
        {
            return ((MethodCallExpression)methodSelector.Body).Method;
        }

        public static MethodInfo GetMethod<T>(this T instance, Expression<Action<T>> methodSelector)
        {
            return ((MethodCallExpression)methodSelector.Body).Method;
        }

        public static bool HasAttribute<TAttribute>(
            this MemberInfo member)
            where TAttribute : Attribute
        {
            return GetAttributes<TAttribute>(member).Any();
        }
    }
}
