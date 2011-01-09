
Imports ClosedXML
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
        Dim table = GetDataTable(17, 8280)
        DataSetToClosedXML1(table, "Center")
        'Console.ReadKey()
    End Sub

    Public Function GetDataTable(ByVal NumberOfColumns As Integer, ByVal NumberOfRows As Integer)
        Dim table = New DataTable()
        For co = 1 To NumberOfColumns
            Dim coName = "Column" & co
            Dim coType As Type
            Dim val = co Mod 5
            Select Case val
                Case 1
                    coType = GetType(String)
                Case 2
                    coType = GetType(Boolean)
                Case 3
                    coType = GetType(Date)
                Case 4
                    coType = GetType(Integer)
                Case Else
                    coType = GetType(TimeSpan)
            End Select

            table.Columns.Add(coName, coType)
        Next
        Dim baseDate = Date.Now
        Dim rnd = New Random()
        For ro = 1 To NumberOfRows
            Dim dr As DataRow = table.NewRow()
            For co = 1 To NumberOfColumns
                Dim coName As String = "Column" & co
                Dim coValue As Object
                Dim val = co Mod 5
                Select Case val
                    Case 1
                        coValue = Guid.NewGuid().ToString().Substring(1, 5)
                    Case 2
                        coValue = (ro Mod 2 = 0)
                    Case 3
                        coValue = DateTime.Now
                    Case 4
                        coValue = rnd.Next(1, 1000)
                    Case Else
                        coValue = (DateTime.Now - baseDate)
                End Select
                dr.Item(coName) = coValue
            Next
            table.Rows.Add(dr)
        Next

        Return table
    End Function

    Public Sub DataSetToClosedXML1(ByVal MyDataTable As DataTable, ByVal BodyAlignment As String)

        'based on ClosedXML.dll downloaded from this website (free license: Version 0.39.0 12/30/2010) (ASPNET 3.5, not 4.0) - add to References
        'http://closedxml.codeplex.com/

        'requires DocumentFormat.OpenXML.dll - add to References
        'DLL can be obtained by downloading the full OpenXML SDK 2.0 package OR just the assembly containing the DLL
        'http://www.microsoft.com/downloads/en/details.aspx?FamilyID=c6e744e5-36e9-45f5-8d8c-331df206e0d0&DisplayLang=en

        'OpenXML also requires Reference to the WindowsBase assembly (WindowsBase.dll) in order to use the System.IO.Packaging namespace.

        'inputs: dataset; filename; Tab name; BodyAlignment = None (default), Left, Center, Right; optional SaveToDisk = Yes
        'code sample: DataSetToClosedXML(MyDS, "TestFile", "Test", "Left")


        On Error GoTo ErrHandler


        Dim wb As ClosedXML.Excel.XLWorkbook = New ClosedXML.Excel.XLWorkbook

        Dim ws As ClosedXML.Excel.IXLWorksheet = wb.Worksheets.Add("Sheet1")

        Dim column As DataColumn
        Dim ColCount As Integer = MyDataTable.Columns.Count
        Dim RowCount As Integer = MyDataTable.Rows.Count
        Dim ColLetter As String

        For Each column In MyDataTable.Columns

            ws.Cell(1, MyDataTable.Columns.IndexOf(column) + 1).Value = column.ColumnName

        Next

        Dim contentRow As DataRow
        Dim r, c As Integer

        For r = 0 To MyDataTable.Rows.Count - 1

            contentRow = MyDataTable.Rows(r)

            For c = 0 To ColCount - 1

                'adjust for header in first row
                ws.Cell(r + 2, c + 1).Value = contentRow(c)

                'format for data type:

                Select Case MyDataTable.Columns(c).DataType.ToString
                    Case "System.Int16", "System.Int32", "System.Int64", "System.UInt16", "System.UInt32", "System.UInt64", "System.Byte", "System.SByte"
                        ws.Cell(r + 2, c + 1).Style.NumberFormat.NumberFormatId = 3
                    Case "System.Single", "System.Double", "System.Decimal"
                        ws.Cell(r + 2, c + 1).Style.NumberFormat.NumberFormatId = 0
                    Case "System.Boolean"
                        ws.Cell(r + 2, c + 1).Value = "'" & contentRow(c).ToString()
                    Case "System.DateTime"
                        ws.Cell(r + 2, c + 1).Style.NumberFormat.NumberFormatId = 14
                    Case "System.String", "System.Char", "System.TimeSpan"
                        ws.Cell(r + 2, c + 1).Value = "'" & contentRow(c).ToString()
                    Case "System.Byte[]"
                        ws.Cell(r + 2, c + 1).DataType = ClosedXML.Excel.XLCellValues.Text
                    Case Else
                        ws.Cell(r + 2, c + 1).DataType = ClosedXML.Excel.XLCellValues.Text
                End Select

            Next

        Next

        'header: set to Bold
        ws.Range(1, 1, 1, ColCount).Style.Font.Bold = True

        'header column alignment (always centered)
        ws.Range(1, 1, 1, ColCount).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center

        'body column alignment
        Select Case BodyAlignment
            Case "None"
                'do nothing (default)
            Case "Left"
                ws.Range(2, 1, RowCount + 1, ColCount).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left
            Case "Center"
                ws.Range(2, 1, RowCount + 1, ColCount).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center
            Case "Right"
                ws.Range(2, 1, RowCount + 1, ColCount).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Right
        End Select

        'auto-fit cols
        'ws.Columns(1, ColCount).AdjustToContents()

        'View: freeze pane - freezes top row (headers)
        ws.SheetView.FreezeRows(1)


        'save to disk
        Dim startTime = DateTime.Now
        wb.SaveAs("C:\Excel Files\ForTesting\Benchmark.xlsx")
        Dim endTime = DateTime.Now
        Console.WriteLine("Saved in {0} secs.", (endTime - startTime).TotalSeconds)

        Exit Sub

ErrHandler:
        'this is a library Sub that displays a javascript alert
        Console.WriteLine("Error: " & Err.Description)

    End Sub



End Module
