using ClosedXML.Examples.Styles;
using System.IO;

namespace ClosedXML.Examples
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
            var path = Program.BaseCreatedDirectory;
            new StyleFont().Create(Path.Combine(path, "styleFont.xlsx"));
            new StyleFill().Create(Path.Combine(path, "styleFill.xlsx"));
            new StyleBorder().Create(Path.Combine(path, "styleBorder.xlsx"));
            new StyleAlignment().Create(Path.Combine(path, "styleAlignment.xlsx"));
            new StyleNumberFormat().Create(Path.Combine(path, "styleNumberFormat.xlsx"));
            new StyleIncludeQuotePrefix().Create(Path.Combine(path, "styleIncludeQuotePrefix.xlsx"));
        }

        // Private

        // Override


        #endregion
    }
}
