using System.IO;
using System.Text;
using System.Xml;
using ClosedXML.Excel.IO;
using ClosedXML.IO;
using NUnit.Framework;

namespace ClosedXML.Tests.IO;

[TestFixture]
internal class XmlTreeReaderTests
{
    [Test]
    public void Can_transparently_processes_MCE()
    {
        const string xml = $"""
                            <font xmlns="{OpenXmlConst.Main2006SsNs}"
                                   xmlns:mc="{OpenXmlConst.MarkupCompatibilityNs}">
                              <mc:AlternateContent>
                                <mc:Choice xmlns:cs="http://example.com/custom" Requires="cs">
                                  <cs:bold weight="10"/>
                                </mc:Choice>
                                <mc:Fallback>
                                  <b/>
                                </mc:Fallback>
                              </mc:AlternateContent>
                            </font>
                            """;
        using var reader = CreateReader(xml);
        reader.Open("font", OpenXmlConst.Main2006SsNs);
        reader.Open("b", OpenXmlConst.Main2006SsNs);
        reader.Close("b", OpenXmlConst.Main2006SsNs);
        reader.Close("font", OpenXmlConst.Main2006SsNs);
    }

    [Test]
    public void GetContent_reads_xml_in_element()
    {
        const string xml =
            """
            <root>
              Hello <![CDATA[world]]>
              ! 
            </root>
            """;
        using var reader = CreateReader(xml);
        reader.Open("root", string.Empty);
        var content = reader.GetContent();
        reader.Close("root", string.Empty);

        Assert.That(content, Is.EqualTo("\n  Hello world\n  ! \n"));
    }

    [Test]
    public void LeaveOpen_true_does_not_close_wrapped_XmlReader()
    {
        using var xmlReader = CreateXmlReader("<root/>");
        using (new XmlTreeReader(xmlReader, XmlToEnumMapper.Instance, true, leaveOpen: true))
        {
        }

        Assert.That(xmlReader.ReadState, Is.Not.EqualTo(ReadState.Closed));
    }

    [Test]
    public void LeaveOpen_false_closes_wrapped_XmlReader()
    {
        var xmlReader = CreateXmlReader("<root/>");
        using (new XmlTreeReader(xmlReader, XmlToEnumMapper.Instance, true, leaveOpen: false))
        {
        }

        Assert.That(xmlReader.ReadState, Is.EqualTo(ReadState.Closed));
    }

    [Test]
    public void TryOpen_opens_current_element_of_already_positioned_XmlReader()
    {
        using var xmlReader = CreateXmlReader("""
                                              <root>
                                                <child/>
                                              </root>
                                              """);
        xmlReader.Read();
        xmlReader.Read();
        Assert.That(xmlReader.NodeType, Is.EqualTo(XmlNodeType.Element));
        Assert.That(xmlReader.LocalName, Is.EqualTo("child"));

        using var treeReader = new XmlTreeReader(xmlReader, XmlToEnumMapper.Instance);

        Assert.That(treeReader.TryOpen("root", string.Empty), Is.False);
        Assert.That(treeReader.TryOpen("child", string.Empty), Is.True);
    }

    [Test]
    public void TryOpen_opens_next_element_when_wrapped_XmlReader_is_not_on_element()
    {
        using var xmlReader = CreateXmlReader("""
                                              <root>
                                                text
                                                <child/>
                                              </root>
                                              """);
        xmlReader.Read();
        xmlReader.Read();
        Assert.That(xmlReader.NodeType, Is.Not.EqualTo(XmlNodeType.Element));
        Assert.That(xmlReader.NodeType, Is.EqualTo(XmlNodeType.Text));

        using var treeReader = new XmlTreeReader(xmlReader, XmlToEnumMapper.Instance);

        Assert.That(treeReader.TryOpen("root", string.Empty), Is.False);
        Assert.That(treeReader.TryOpen("child", string.Empty), Is.True);
    }

    private static XmlReader CreateXmlReader(string xml)
    {
        return XmlReader.Create(new StringReader(xml), new XmlReaderSettings
        {
            IgnoreWhitespace = true,
            IgnoreComments = true,
            DtdProcessing = DtdProcessing.Prohibit,
        });
    }

    private static XmlTreeReader CreateReader(string xml)
    {
        return new XmlTreeReader(new MemoryStream(Encoding.UTF8.GetBytes(xml)), XmlToEnumMapper.Instance, true);
    }
}
