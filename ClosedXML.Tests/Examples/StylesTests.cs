using ClosedXML.Examples.Styles;
using NUnit.Framework;
using System.Runtime.InteropServices;

namespace ClosedXML.Tests.Examples
{
    [TestFixture]
    public class StylesTests
    {
        [Test]
        public void DefaultStyles()
        {
            TestHelper.RunTestExample<DefaultStyles>(@"Styles\DefaultStyles.xlsx");
        }

        [Test]
        public void PurpleWorksheet()
        {
            TestHelper.RunTestExample<PurpleWorksheet>(@"Styles\PurpleWorksheet.xlsx");
        }

        [Test]
        public void StyleAlignment()
        {
            TestHelper.RunTestExample<StyleAlignment>(@"Styles\StyleAlignment.xlsx");
        }

        [Test]
        public void StyleBorder()
        {
            TestHelper.RunTestExample<StyleBorder>(@"Styles\StyleBorder.xlsx");
        }

        [Test]
        public void StyleFill()
        {
            TestHelper.RunTestExample<StyleFill>(@"Styles\StyleFill.xlsx");
        }

        [Test]
        public void StyleFont()
        {
            TestHelper.RunTestExample<StyleFont>(@"Styles\StyleFont.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void StyleNumberFormat()
        {
            TestHelper.RunTestExample<StyleNumberFormat>(@"Styles\StyleNumberFormat.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void StyleIncludeQuotePrefix()
        {
            TestHelper.RunTestExample<StyleIncludeQuotePrefix>(@"Styles\StyleIncludeQuotePrefix.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void StyleRowsColumns()
        {
            TestHelper.RunTestExample<StyleRowsColumns>(@"Styles\StyleRowsColumns.xlsx");
        }

        [Test]
        public void StyleWorksheet()
        {
            TestHelper.RunTestExample<StyleWorksheet>(@"Styles\StyleWorksheet.xlsx");
        }

        [Test]
        public void UsingColors()
        {
            TestHelper.RunTestExample<UsingColors>(@"Styles\UsingColors.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }

        [Test]
        public void UsingPhonetics()
        {
            TestHelper.RunTestExample<UsingPhonetics>(@"Styles\UsingPhonetics.xlsx");
        }

        [Test]
        public void UsingRichText()
        {
            TestHelper.RunTestExample<UsingRichText>(@"Styles\UsingRichText.xlsx", ignoreColumnFormats: !RuntimeInformation.IsOSPlatform(OSPlatform.Windows));
        }
    }
}
