using ClosedXML.IO.CodeGen.Model.TopLevel;
using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// A representation of a one XSD file.
/// </summary>
public class Schema
{
    /// <summary>
    /// Imports in the file.
    /// </summary>
    public List<ImportElement> Imports = [];

    /// <summary>
    /// One of <see cref="AttributeGroupDefinition"/>, <see cref="ComplexType"/>, <see cref="ElementDefinition"/>.
    /// </summary>
    public List<object> Entries = [];
}
