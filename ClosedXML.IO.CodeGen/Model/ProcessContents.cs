namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// How should XML processor process content of <c>any</c> element from <c>xsd</c>.
/// </summary>
public enum ProcessContents
{
    /// <summary>
    /// Default value, has same meaning as <see cref="Strict"/>.
    /// </summary>
    Default,

    /// <summary>
    /// All elements should be validated against schema. This should be used when content
    /// should only contain known schema.
    /// </summary>
    Strict,

    /// <summary>
    /// Validation should only be performed only on elements found in schema. If there is no
    /// schema, there is no error. That means some elements may belong to known schema, but
    /// unknown ones shouldn't cause errors.
    /// </summary>
    Lax
}
