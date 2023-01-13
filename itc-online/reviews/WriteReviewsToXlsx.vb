Imports Newtonsoft.Json
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports iqb.lib.openxml
Imports System.ComponentModel
Class WriteReviewsToXlsx
    Public Shared Sub Write(
                           workbookTemplate As Byte(),
                           worker As BackgroundWorker,
                           e As DoWorkEventArgs,
                           AllReviews As List(Of ReviewDTO),
                           targetXlsxFilename As String
                           )
        Dim allBooklets As New List(Of String)
        Dim allUnits As New List(Of String)
        For Each r As ReviewDTO In AllReviews
            If String.IsNullOrEmpty(r.unitname) Then
                If Not allBooklets.Contains(r.bookletname) Then allBooklets.Add(r.bookletname)
            Else
                If Not allUnits.Contains(r.unitname) Then allUnits.Add(r.unitname)
            End If
        Next

        Using MemStream As New IO.MemoryStream()
            MemStream.Write(workbookTemplate, 0, workbookTemplate.Length)
            Using ZielXLS As SpreadsheetDocument = SpreadsheetDocument.Open(MemStream, True)
                worker.ReportProgress(0.0#, "Schreibe Daten")
                Dim myStyles As ExcelStyleDefs = xlsxFactory.AddIQBStandardStyles(ZielXLS.WorkbookPart)
                '########################################################
                For Each b As String In From bName In allBooklets Order By bName
                    Dim targetTable As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "Booklet_" + b)
                    Dim myRow As Integer = 1
                    xlsxFactory.SetCellValueString("A", myRow, targetTable, "groupname", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("A", targetTable, 10)
                    xlsxFactory.SetCellValueString("B", myRow, targetTable, "loginname", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("B", targetTable, 10)
                    xlsxFactory.SetCellValueString("C", myRow, targetTable, "code", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("C", targetTable, 5)
                    xlsxFactory.SetCellValueString("D", myRow, targetTable, "priority", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("D", targetTable, 10)
                    xlsxFactory.SetCellValueString("E", myRow, targetTable, "category", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("E", targetTable, 10)
                    xlsxFactory.SetCellValueString("F", myRow, targetTable, "reviewtime", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("F", targetTable, 10)
                    xlsxFactory.SetCellValueString("G", myRow, targetTable, "entry", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("G", targetTable, 80)

                    For Each review As ReviewDTO In AllReviews
                        If String.IsNullOrEmpty(review.unitname) AndAlso review.bookletname = b Then
                            Dim RowData As New List(Of RowData)
                            RowData.Add(New RowData With {.Column = "A", .Value = review.groupname, .CellType = CellTypes.str})
                            RowData.Add(New RowData With {.Column = "B", .Value = review.loginname, .CellType = CellTypes.str})
                            RowData.Add(New RowData With {.Column = "C", .Value = review.code, .CellType = CellTypes.str})
                            RowData.Add(New RowData With {.Column = "D", .Value = review.priority, .CellType = CellTypes.str})
                            Dim categories As New List(Of String)
                            If review.categoryDesign Then categories.Add("D")
                            If review.categoryTech Then categories.Add("T")
                            If review.categoryContent Then categories.Add("C")
                            If categories.Count > 0 Then RowData.Add(New RowData With {.Column = "E", .Value = String.Join(" / ", categories), .CellType = CellTypes.str})
                            RowData.Add(New RowData With {.Column = "F", .Value = review.reviewTime, .CellType = CellTypes.str})
                            RowData.Add(New RowData With {.Column = "G", .Value = review.entry, .CellType = CellTypes.str})

                            myRow += 1
                            xlsxFactory.AppendRow(myRow, RowData, targetTable)
                        End If
                    Next
                Next

                '########################################################
                For Each u As String In From uName In allUnits Order By uName
                    Dim targetTable As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, u)
                    Dim myRow As Integer = 1
                    xlsxFactory.SetCellValueString("A", myRow, targetTable, "groupname", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("A", targetTable, 10)
                    xlsxFactory.SetCellValueString("B", myRow, targetTable, "loginname", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("B", targetTable, 10)
                    xlsxFactory.SetCellValueString("C", myRow, targetTable, "code", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("C", targetTable, 5)
                    xlsxFactory.SetCellValueString("D", myRow, targetTable, "priority", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("D", targetTable, 10)
                    xlsxFactory.SetCellValueString("E", myRow, targetTable, "category", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("E", targetTable, 10)
                    xlsxFactory.SetCellValueString("F", myRow, targetTable, "reviewtime", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("F", targetTable, 10)
                    xlsxFactory.SetCellValueString("G", myRow, targetTable, "entry", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("G", targetTable, 80)

                    For Each review As ReviewDTO In AllReviews
                        If review.unitname = u Then
                            Dim RowData As New List(Of RowData)
                            RowData.Add(New RowData With {.Column = "A", .Value = review.groupname, .CellType = CellTypes.str})
                            RowData.Add(New RowData With {.Column = "B", .Value = review.loginname, .CellType = CellTypes.str})
                            RowData.Add(New RowData With {.Column = "C", .Value = review.code, .CellType = CellTypes.str})
                            RowData.Add(New RowData With {.Column = "D", .Value = review.priority, .CellType = CellTypes.str})
                            Dim categories As New List(Of String)
                            If review.categoryDesign Then categories.Add("D")
                            If review.categoryTech Then categories.Add("T")
                            If review.categoryContent Then categories.Add("C")
                            If categories.Count > 0 Then RowData.Add(New RowData With {.Column = "E", .Value = String.Join(" / ", categories), .CellType = CellTypes.str})
                            RowData.Add(New RowData With {.Column = "F", .Value = review.reviewTime, .CellType = CellTypes.str})
                            RowData.Add(New RowData With {.Column = "G", .Value = review.entry, .CellType = CellTypes.str})

                            myRow += 1
                            xlsxFactory.AppendRow(myRow, RowData, targetTable)
                        End If
                    Next
                Next

            End Using
            worker.ReportProgress(100.0#, "Speichern Datei")
            Try
                Using fs As New IO.FileStream(targetXlsxFilename, IO.FileMode.Create)
                    MemStream.WriteTo(fs)
                End Using
            Catch ex As Exception
                worker.ReportProgress(100.0#, "e: Konnte Datei nicht schreiben: " + ex.Message)
            End Try
        End Using
    End Sub
End Class
