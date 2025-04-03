namespace ClosedXML.IO.CodeGen.Model;

/// <summary>
/// Representation of <c>import</c> element in <c>xsd</c> file.
/// <code><![CDATA[
///  <xsd:import namespace="http://purl.oclc.org/ooxml/officeDocument/relationships" schemaLocation="shared-relationshipReference.xsd"/>
/// ]]></code>
/// </summary>
public class ImportElement
{
    public required string Namespace { get; init; }

    public required string SchemaLocation { get; init; }
}
