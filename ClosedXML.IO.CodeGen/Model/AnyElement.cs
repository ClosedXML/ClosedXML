using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// <code><![CDATA[<xsd:any processContents="lax"/>]]></code>
/// </summary>
public class AnyElement : IElementGroup
{
    public List<IElementGroup> Children { get; } = [];

    public required ProcessContents ProcessContent { get; init; }
}
