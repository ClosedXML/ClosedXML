namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// Is attribute optional or does it always has to be specified?
/// </summary>
public enum AttributeUseType
{
    /// <summary>
    /// Default is <see cref="Optional"/>.
    /// </summary>
    Default,
    Optional,
    Required
}
