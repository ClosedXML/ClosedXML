using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
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
            new StyleFont().Create(@"c:\styleFont.xlsx");
            new StyleFill().Create(@"c:\styleFill.xlsx");
            new StyleBorder().Create(@"c:\styleBorder.xlsx");
            new StyleAlignment().Create(@"c:\styleAlignment.xlsx");
            new StyleNumberFormat().Create(@"c:\styleNumberFormat.xlsx");
        }

        // Private

        // Override


        #endregion
    }
}
