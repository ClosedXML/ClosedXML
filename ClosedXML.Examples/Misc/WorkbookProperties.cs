using System;
using ClosedXML.Excel;


namespace ClosedXML.Examples.Misc
{
    public class WorkbookProperties : IXLExample
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

        #region Events

        // Public

        // Private

        // Override


        #endregion

        #region Methods

        // Public
        public void Create(String filePath)
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Workbook Properties");

            wb.Properties.Author = "theAuthor";
            wb.Properties.Title = "theTitle";
            wb.Properties.Subject = "theSubject";
            wb.Properties.Category = "theCategory";
            wb.Properties.Keywords = "theKeywords";
            wb.Properties.Comments = "theComments";
            wb.Properties.Status = "theStatus";
            wb.Properties.LastModifiedBy = "theLastModifiedBy";
            wb.Properties.Company = "theCompany";
            wb.Properties.Manager = "theManager";

            // Creating/Using custom properties
            wb.CustomProperties.Add("theText", "XXX");
            wb.CustomProperties.Add("theDate", new DateTime(2011, 1, 1, 17, 0, 0, DateTimeKind.Utc)); // Use UTC to make sure test can be run in any time zone
            wb.CustomProperties.Add("theNumber", 123.456);
            wb.CustomProperties.Add("theBoolean", true);

            wb.SaveAs(filePath);
        }

        // Private

        // Override


        #endregion
    }
}
