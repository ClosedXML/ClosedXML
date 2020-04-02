// Keep this file CodeMaid organised and cleaned
using ClosedXML.Attributes;
using ClosedXML.Extensions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ClosedXML.Excel.InsertData
{
    internal class ObjectReader : IInsertDataReader
    {
        private const BindingFlags bindingFlags = BindingFlags.Public
                                                  | BindingFlags.Instance
                                                  | BindingFlags.Static;

        private readonly IEnumerable<object> _data;
        private readonly MemberInfo[] _members;
        private readonly bool[] _staticMembers;

        public ObjectReader(IEnumerable data)
        {
            _data = data?.Cast<object>() ?? throw new ArgumentNullException(nameof(data));

            var itemType = data.GetItemType();
            if (itemType.IsNullableType())
                itemType = itemType.GetUnderlyingType();

            _members = itemType.GetFields(bindingFlags).Cast<MemberInfo>()
                .Concat(itemType.GetProperties(bindingFlags))
                .Where(mi => !XLColumnAttribute.IgnoreMember(mi))
                .OrderBy(mi => XLColumnAttribute.GetOrder(mi))
                .ToArray();

            _staticMembers = _members.Select(ReflectionExtensions.IsStatic).ToArray();
        }

        public IEnumerable<IEnumerable<object>> GetData()
        {
            return _data.Select(GetItemData);
        }

        public int GetPropertiesCount()
        {
            return _members.Length;
        }

        public string GetPropertyName(int propertyIndex)
        {
            if (propertyIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(propertyIndex), "Property index must be non-negative");

            if (propertyIndex >= GetPropertiesCount())
                throw new ArgumentOutOfRangeException($"{propertyIndex} exceeds the number of the object properties");

            var memberInfo = _members[propertyIndex];
            var fieldName = XLColumnAttribute.GetHeader(memberInfo);
            if (String.IsNullOrWhiteSpace(fieldName))
                fieldName = memberInfo.Name;

            return fieldName;
        }

        public int GetRecordsCount()
        {
            return _data.Count();
        }

        private IEnumerable<object> GetItemData(object item)
        {
            for (int i = 0; i < _members.Length; i++)
            {
                if (item == null)
                {
                    yield return null;
                    continue;
                }

                var memberInfo = _members[i];
                switch (memberInfo)
                {
                    case PropertyInfo propertyInfo when _staticMembers[i]:
                        yield return propertyInfo.GetValue(null, null);
                        break;

                    case PropertyInfo propertyInfo:
                        yield return propertyInfo.GetValue(item, null);
                        break;

                    case FieldInfo fieldInfo when _staticMembers[i]:
                        yield return fieldInfo.GetValue(null);
                        break;

                    case FieldInfo fieldInfo:
                        yield return fieldInfo.GetValue(item);
                        break;
                }
            }
        }
    }
}
