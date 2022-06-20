using ClosedXML.Excel;
using static ClosedXML.Excel.XLProtectionAlgorithm;

namespace ClosedXML.Examples.Misc
{
    public class SheetProtection : IXLExample
    {
        public void Create(string filePath)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Protected No-Password");

            ws.Protect().AllowElement
            (
                // On this sheet we will only allow:
                XLSheetProtectionElements.FormatCells
                | XLSheetProtectionElements.InsertColumns
                | XLSheetProtectionElements.DeleteColumns
                | XLSheetProtectionElements.DeleteRows
                | XLSheetProtectionElements.EditScenarios
            );

            ws.Cell("A1").SetValue("Locked, No Hidden (Default):").Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.Cyan);
            ws.Cell("B1").Style
                .Border.SetOutsideBorder(XLBorderStyleValues.Medium);

            ws.Cell("A2").SetValue("Locked, Hidden:").Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.Cyan);
            ws.Cell("B2").Style
                .Protection.SetHidden()
                .Border.SetOutsideBorder(XLBorderStyleValues.Medium);

            ws.Cell("A3").SetValue("Not Locked, Hidden:").Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.Cyan);
            ws.Cell("B3").Style
                .Protection.SetLocked(false)
                .Protection.SetHidden()
                .Border.SetOutsideBorder(XLBorderStyleValues.Medium);

            ws.Cell("A4").SetValue("Not Locked, Not Hidden:").Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.Cyan);
            ws.Cell("B4").Style
                .Protection.SetLocked(false)
                .Border.SetOutsideBorder(XLBorderStyleValues.Medium);

            ws.Columns().AdjustToContents();

            // Protect a sheet with a password
            var protectedSheet = wb.Worksheets.Add("Protected Password = 123");
            var protection = protectedSheet.Protect("123", Algorithm.SimpleHash);
            protection.AllowElement
            (
                XLSheetProtectionElements.InsertRows
                | XLSheetProtectionElements.InsertColumns
            );

            wb.SaveAs(filePath);
        }
    }
}