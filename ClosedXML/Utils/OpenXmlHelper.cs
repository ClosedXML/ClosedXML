using DocumentFormat.OpenXml;

namespace ClosedXML.Utils
{
    internal static class OpenXmlHelper
    {
        public static BooleanValue GetBooleanValue(bool value, bool defaultValue)
        {
            return value == defaultValue ? null : new BooleanValue(value);
        }

        public static bool GetBooleanValueAsBool(BooleanValue value, bool defaultValue)
        {
            return value == null ? defaultValue : value.Value;
        }
    }
}