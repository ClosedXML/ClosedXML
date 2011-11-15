using ClosedXML_Examples.Styles;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ClosedXML_Tests.Examples
{
    [TestClass]
    public class StylesTests
    {
        [TestMethod]
        public void DefaultStyles()
        {
            TestHelper.RunTestExample<DefaultStyles>(@"Styles\DefaultStyles.xlsx");
        }
        [TestMethod]
        public void StyleAlignment()
        {
            TestHelper.RunTestExample<StyleAlignment>(@"Styles\StyleAlignment.xlsx");
        }
        [TestMethod]
        public void StyleBorder()
        {
            TestHelper.RunTestExample<StyleBorder>(@"Styles\StyleBorder.xlsx");
        }
        [TestMethod]
        public void StyleFill()
        {
            TestHelper.RunTestExample<StyleFill>(@"Styles\StyleFill.xlsx");
        }
        [TestMethod]
        public void StyleFont()
        {
            TestHelper.RunTestExample<StyleFont>(@"Styles\StyleFont.xlsx");
        }
        [TestMethod]
        public void StyleNumberFormat()
        {
            TestHelper.RunTestExample<StyleNumberFormat>(@"Styles\StyleNumberFormat.xlsx");
        }
        [TestMethod]
        public void StyleRowsColumns()
        {
            TestHelper.RunTestExample<StyleRowsColumns>(@"Styles\StyleRowsColumns.xlsx");
        }
        [TestMethod]
        public void StyleWorksheet()
        {
            TestHelper.RunTestExample<StyleWorksheet>(@"Styles\StyleWorksheet.xlsx");
        }
      
         [TestMethod]
        public void UsingRichText()
        {
            TestHelper.RunTestExample<UsingRichText>(@"Styles\UsingRichText.xlsx");
        }

         [TestMethod]
         public void PurpleWorksheet()
         {
             TestHelper.RunTestExample<PurpleWorksheet>(@"Styles\PurpleWorksheet.xlsx");
         }

         [TestMethod]
         public void UsingColors()
         {
             TestHelper.RunTestExample<UsingColors>(@"Styles\UsingColors.xlsx");
         }
    }
}