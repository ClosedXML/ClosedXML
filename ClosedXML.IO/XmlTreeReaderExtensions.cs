using System;

namespace ClosedXML.IO;

/// <summary>
/// Extension methods to make reading from <see cref="XmlTreeReader"/> simpler and keep the reader slim.
/// </summary>
public static class XmlTreeReaderExtensions
{
    public static bool GetBool(this XmlTreeReader reader, string attributeName)
    {
        return reader.GetOptionalBool(attributeName) ?? throw PartStructureException.MissingAttribute(attributeName, reader);
    }

    public static int GetInt(this XmlTreeReader reader, string attributeName)
    {
        return reader.GetOptionalInt(attributeName) ?? throw PartStructureException.MissingAttribute(attributeName, reader);
    }

    public static uint GetUint(this XmlTreeReader reader, string attributeName)
    {
        return reader.GetOptionalUint(attributeName) ?? throw PartStructureException.MissingAttribute(attributeName, reader);
    }

    public static int? GetOptionalUintAsInt(this XmlTreeReader reader, string attributeName)
    {
        return checked((int?)reader.GetOptionalUint(attributeName));
    }

    public static double GetDouble(this XmlTreeReader reader, string attributeName)
    {
        return reader.GetOptionalDouble(attributeName) ?? throw PartStructureException.MissingAttribute(attributeName, reader);
    }

    public static string GetString(this XmlTreeReader reader, string attributeName)
    {
        return reader.GetOptionalString(attributeName) ?? throw PartStructureException.MissingAttribute(attributeName, reader);
    }

    public static TEnum GetEnum<TEnum>(this XmlTreeReader reader, string attributeName)
        where TEnum : struct, Enum
    {
        return reader.GetOptionalEnum<TEnum>(attributeName) ?? throw PartStructureException.MissingAttribute(attributeName, reader);
    }

    public static TEnum GetOptionalEnum<TEnum>(this XmlTreeReader reader, string attributeName, TEnum defaultValue)
        where TEnum : struct, Enum
    {
        return reader.GetOptionalEnum<TEnum>(attributeName) ?? defaultValue;
    }
}
