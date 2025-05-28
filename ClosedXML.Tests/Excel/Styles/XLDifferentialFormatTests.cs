using ClosedXML.Excel;
using ClosedXML.Excel.Formatting;
using NUnit.Framework;

namespace ClosedXML.Tests.Excel.Styles;

/// <summary>
/// Tests of <see cref="IXLDifferentialFormat"/> API.
/// </summary>
[TestOf(typeof(XLDifferentialFormat))]
[TestOf(typeof(IXLDifferentialFormat))]
internal class XLDifferentialFormatTests
{
    [Test]
    public void SetBold_changes_font_format_to_bold_and_registers_new_format_to_styles()
    {
        var styles = new XLWorkbookStyles();
        var dxf = new TestDxfObject(styles);

        dxf.Format.Font.SetBold();

        Assert.IsTrue(dxf.FormatValue.Font?.Bold);
        Assert.AreSame(dxf.FormatValue, styles.DifferentialFormats[0]);
        Assert.AreSame(dxf.FormatValue.Font, styles.Fonts[0]);
    }

    private class TestDxfObject : IDxfContainer
    {
        private readonly XLWorkbookStyles _styles;

        public TestDxfObject(XLWorkbookStyles styles)
        {
            _styles = styles;
        }

        public XLDxfValue FormatValue { get; set; } = XLDxfValue.Empty;

        public IXLDifferentialFormat Format => new XLDifferentialFormat(this, _styles);
    }
}
