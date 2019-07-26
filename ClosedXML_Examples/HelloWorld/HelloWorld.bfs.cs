using System;
using ClosedXML.Excel;

namespace ClosedXML_Examples
{
    public class HelloWorld
    {
	    public void Create(String filePath)
	    {
		    IXLWorkbook wb = new XLWorkbook();
		    IXLWorksheet ws = wb.Worksheets.Add("Sample Sheet");

		    ws.Cell(2,3).Value = "Hello World!";
		    ws.Cell(4,2).Value = "Project:";
		    ws.Cell(4,4).Value = "ClosedXML Example";
		    ws.Cell(6,2).Value = "Author:";
		    ws.Cell(6,4).Value = "KnapSac";

		    wb.SaveAs(filePath);
	    }
    }
}
