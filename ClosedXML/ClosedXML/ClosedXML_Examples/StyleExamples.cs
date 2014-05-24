using ClosedXML_Examples.Styles;

namespace ClosedXML_Examples
{
    public class StyleExamples
    {
        #region Variables

        // Public

        // Private


        #endregion

        #region Properties

        // Public

        // Private

        // Override


        #endregion

        #region Constructors

        // Public
        public StyleExamples()
        {

        }


        // Private


        #endregion

        #region Events

        // Public

        // Private

        // Override


        #endregion

        #region Methods

        // Public
        public void Create()
        {
            var path = @"C:\ClosedXML_Tests\Created";
            new StyleFont().Create(path + @"\styleFont.xlsx");
            new StyleFill().Create(path + @"\styleFill.xlsx");
            new StyleBorder().Create(path + @"\styleBorder.xlsx");
            new StyleAlignment().Create(path + @"\styleAlignment.xlsx");
            new StyleNumberFormat().Create(path + @"\styleNumberFormat.xlsx");
        }

        // Private

        // Override


        #endregion
    }
}
