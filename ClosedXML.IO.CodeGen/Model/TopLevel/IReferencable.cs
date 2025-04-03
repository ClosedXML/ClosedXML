namespace ClosedXML.IO.CodeGen.Model.TopLevel;

/// <summary>
/// A marker interface for elements that can be referenced from other parts of XSD.
/// </summary>
internal interface IReferencable
{
    /// <summary>
    /// The name used by other to reference this element.
    /// </summary>
    string Name { get; }
}
