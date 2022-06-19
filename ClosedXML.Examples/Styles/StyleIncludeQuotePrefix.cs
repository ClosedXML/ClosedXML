using ClosedXML.Excel;
using System.Linq;

namespace ClosedXML.Examples.Styles
{
    public class StyleIncludeQuotePrefix : IXLExample
    {
        public void Create(string filePath)
        {
            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Style IncludeQuotePrefix");

            var data = Enumerable.Range(1, 20)
                .Select(i =>
                new
                {
                    IntegerIndex = i,
                    StringIndex = i.ToString(),
                    PaddedString1000 = (i * 1000).ToString().PadLeft(8, '0'),
                    PrependedString1000 = "Str" + (i * 1000).ToString().PadLeft(8, '0')
                });

            ws.FirstCell().InsertData(data);

            // Columns B to D will be of type text
            // but column B will not have the leading quotation mark
            ws.Column("B").Style.IncludeQuotePrefix = false;

            // Columns C and D will have the leading quotation mark
            ws.Column("C").Style.IncludeQuotePrefix = true;
            ws.Column("D").Style.SetIncludeQuotePrefix();

            ws.Columns().AdjustToContents();

            workbook.SaveAs(filePath);
        }
    }
}