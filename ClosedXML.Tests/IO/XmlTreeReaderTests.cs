using ClosedXML.Excel.IO;
using ClosedXML.IO;
using NUnit.Framework;
using System.IO;
using System.Text;

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

    private static XmlTreeReader CreateReader(string xml)
    {
        return new XmlTreeReader(new MemoryStream(Encoding.UTF8.GetBytes(xml)), XmlToEnumMapper.Instance, true);
    }
}
