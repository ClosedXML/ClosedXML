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
            new StyleFont().Create(@"D:\Excel Files\Created\styleFont.xlsx");
            new StyleFill().Create(@"D:\Excel Files\Created\styleFill.xlsx");
            new StyleBorder().Create(@"D:\Excel Files\Created\styleBorder.xlsx");
            new StyleAlignment().Create(@"D:\Excel Files\Created\styleAlignment.xlsx");
            new StyleNumberFormat().Create(@"D:\Excel Files\Created\styleNumberFormat.xlsx");
        }

        // Private

        // Override


        #endregion
    }
}
