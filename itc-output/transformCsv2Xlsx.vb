Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports iqb.lib.openxml
Public Class transformCsv2Xlsx
    Private columDefs As Dictionary(Of String, String) = Nothing
    Private lineDataList As List(Of Dictionary(Of String, String))
    Public Sub New(sourceFilename As String)
        lineDataList = New List(Of Dictionary(Of String, String))
        Dim separator As String = ""","""
        For Each line As String In IO.File.ReadAllLines(sourceFilename)
            If columDefs Is Nothing Then
                columDefs = New Dictionary(Of String, String)
                Dim myCol As String = "A"
                Dim columnHeaders As String() = line.Substring(1, line.Length - 2).Split({separator}, StringSplitOptions.RemoveEmptyEntries)
                If columnHeaders.Count = 1 Then
                    separator = """;"""
                    columnHeaders = line.Substring(1, line.Length - 2).Split({separator}, StringSplitOptions.RemoveEmptyEntries)
                End If
                For Each col As String In columnHeaders
                    columDefs.Add(myCol, col)
                    myCol = xlsxFactory.GetNextColumn(myCol)
                Next
            Else
                Dim myCol As String = "A"
                Dim lineData As New Dictionary(Of String, String)
                For Each colValue As String In line.Substring(1, line.Length - 2).Split({separator}, StringSplitOptions.None)
                    If Not String.IsNullOrEmpty(colValue) Then
                        lineData.Add(myCol, colValue)
                    End If
                    myCol = xlsxFactory.GetNextColumn(myCol)
                Next
                lineDataList.Add(lineData)
            End If
        Next
        Debug.Print("ok")
    End Sub

    Public Sub ToXlsx(targetFilename As String)
        Dim TmpZielXLS As SpreadsheetDocument = SpreadsheetDocument.Create(targetFilename, SpreadsheetDocumentType.Workbook)
        Dim myWorkbookPart As WorkbookPart = TmpZielXLS.AddWorkbookPart()
        myWorkbookPart.Workbook = New Workbook()
        myWorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())
        TmpZielXLS.Close()

        Dim myTemplate As Byte() = IO.File.ReadAllBytes(targetFilename)
        Using MemStream As New IO.MemoryStream()
            MemStream.Write(myTemplate, 0, myTemplate.Length)
            Using ZielXLS As SpreadsheetDocument = SpreadsheetDocument.Open(MemStream, True)
                Dim myStyles As ExcelStyleDefs = xlsxFactory.AddIQBStandardStyles(ZielXLS.WorkbookPart)
                Dim Tabelle_A As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "Syscheck-Daten")

                Dim zeile As Integer = 1
                For Each colHeader As KeyValuePair(Of String, String) In columDefs
                    xlsxFactory.SetCellValueString(colHeader.Key, zeile, Tabelle_A, colHeader.Value, CellFormatting.RowHeader2, myStyles)
                Next
                For Each lineData As Dictionary(Of String, String) In lineDataList
                    zeile += 1
                    Dim myARow As New List(Of RowData)
                    For Each lineValue As KeyValuePair(Of String, String) In lineData
                        myARow.Add(New RowData With {.CellType = CellTypes.str, .Column = lineValue.Key, .Value = lineValue.Value})
                    Next
                    xlsxFactory.AppendRow(zeile, myARow, Tabelle_A)
                Next

                Tabelle_A.Worksheet.Save()
            End Using

            Using fs As New IO.FileStream(targetFilename, IO.FileMode.Create)
                MemStream.WriteTo(fs)
            End Using
        End Using
    End Sub
End Class
