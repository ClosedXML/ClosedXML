using System;
using System.Collections.Generic;

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

    public static uint GetUInt(this XmlTreeReader reader, string attributeName)
    {
        return reader.GetOptionalUInt(attributeName) ?? throw PartStructureException.MissingAttribute(attributeName, reader);
    }

    public static int? GetOptionalUintAsInt(this XmlTreeReader reader, string attributeName)
    {
        return checked((int?)reader.GetOptionalUInt(attributeName));
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

    public static string GetXString(this XmlTreeReader reader, string attributeName)
    {
        return GetOptionalXString(reader, attributeName) ?? throw PartStructureException.MissingAttribute(attributeName, reader);
    }

    public static string? GetOptionalXString(this XmlTreeReader reader, string attributeName)
    {
        var text = reader.GetOptionalString(attributeName);
        return XStringConvert.Decode(text);
    }

    /// <summary>
    /// Get an attribute with <c>ST_UnsignedIntHex</c> content.
    /// </summary>
    public static uint? GetOptionalUIntHex(this XmlTreeReader reader, string attributeName)
    {
        // XmlReader has ReadContentAsBinHex, but it also requires allocation, so we can just do it
        // as extension method without polluting reader.
        var hexString = reader.GetOptionalString(attributeName);
        if (hexString is null)
            return null;

        if (hexString.Length != 8 || !XStringConvert.TryGetHexValue(hexString.AsSpan(), out var number))
        {
            if (reader.StrictAttributeParsing)
                throw PartStructureException.InvalidAttributeFormat(attributeName, hexString, reader);

            return null;
        }

        return number;
    }

    /// <summary>
    /// Read <c>xsd:dateTime</c> attribute.
    /// </summary>
    public static DateTime GetDateTime(this XmlTreeReader reader, string attributeName)
    {
        return reader.GetOptionalDateTime(attributeName) ?? throw PartStructureException.MissingAttribute(attributeName, reader);
    }

    /// <summary>
    /// Get count for various collections in parts.
    /// </summary>
    public static int GetCount(this XmlTreeReader reader)
    {
        var count = reader.GetOptionalUInt("count") ?? 1;
        return checked((int)count);
    }

    /// <summary>
    /// Get a value mapped from a string attribute value.
    /// </summary>
    /// <typeparam name="TResult">Type of mapped value.</typeparam>
    /// <param name="reader">Reader to read the attribute.</param>
    /// <param name="attributeName">Name of the attribute.</param>
    /// <param name="map">The dictionary that contains mapped values.</param>
    /// <returns>Mapped value.</returns>
    /// <exception cref="PartStructureException">Attribute value isn't in the <paramref name="map"/>.</exception>
    public static TResult GetStringMappedValue<TResult>(this XmlTreeReader reader, string attributeName, IReadOnlyDictionary<string, TResult> map)
    {
        var attributeValue = reader.GetString(attributeName);
        if (!map.TryGetValue(attributeValue, out var result))
            throw PartStructureException.InvalidAttributeFormat(attributeName, attributeValue, reader);

        return result;
    }
}
