Imports ClosedXML.Excel
Imports System.IO
Module Module1
    Sub Main1()
        Dim counter As Integer = 0

        Dim workbook As New XLWorkbook
        Dim worksheet = workbook.Worksheets.Add("Sample Sheet")

        'Row1
        worksheet.Cell(1, 1).Value = "Some Random Text"

        'Row2
        For counter = 0 To 6 Step 1
            worksheet.Cell(2, (counter * 2) + 2).Value = Now.AddDays(counter).ToString("yyyy-MM-dd")
        Next

        'Row3
        worksheet.Cell(3, 1).Value = "val1"
        worksheet.Cell(3, 2).Value = "val2"
        worksheet.Cell(3, 3).Value = "val3"

        'worksheet.PageSetup.PrintAreas.Clear()

        workbook.SaveAs("C:\Excel Files\ForTesting\Issue_5957_Saved.xlsx")
    End Sub
    Sub Main()
        Dim wb = New XLWorkbook()
        Dim ws = wb.Worksheets.Add("Sheet1")
        For Each ro In Enumerable.Range(1, 100)
            For Each co In Enumerable.Range(1, 10)
                ws.Cell(ro, co).Value = ws.Cell(ro, co).Address.ToString()
            Next
        Next
        ws.PageSetup.PagesWide = 1

        Dim ms As New MemoryStream
        wb.SaveAs(ms)
    End Sub

End Module
