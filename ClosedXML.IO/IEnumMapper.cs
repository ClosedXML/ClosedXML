using System;

namespace ClosedXML.IO;

/// <summary>
/// A mapper from a text name in XML to actual enum value. Used by <see cref="XmlTreeReader"/>.
/// </summary>
public interface IEnumMapper
{
    /// <summary>
    /// Try to get a concrete enum value for a passed text.
    /// </summary>
    /// <typeparam name="TEnum">Type of expected enum.</typeparam>
    /// <param name="text">Text value of enum. Comparison is case-sensitive.</param>
    /// <param name="enumValue">Output enum value.</param>
    /// <returns><c>True</c> if enum was found for passed text value, false otherwise.</returns>
    bool TryGetEnum<TEnum>(string text, out TEnum enumValue)
        where TEnum : struct, Enum;

    /// <summary>
    /// Try to get a text for an enum value.
    /// </summary>
    /// <typeparam name="TEnum">Type of an enum.</typeparam>
    /// <param name="enumValue">Enum value.</param>
    /// <param name="text">Output text for the enum value.</param>
    /// <returns><c>True</c> if text was found for passed enum value, false otherwise.</returns>
    bool TryGetText<TEnum>(TEnum enumValue, out string text)
        where TEnum : struct, Enum;
}
