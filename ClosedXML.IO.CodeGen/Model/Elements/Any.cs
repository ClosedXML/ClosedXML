using System.Collections.Generic;

namespace ClosedXML.IO.CodeGen.Model.Elements;

/// <summary>
/// <code><![CDATA[<xsd:any processContents="lax"/>]]></code>
/// </summary>
public class Any : IElementGroup
{
    public List<IElementGroup> Children { get; } = [];

    public required ProcessContents ProcessContent { get; init; }
}
