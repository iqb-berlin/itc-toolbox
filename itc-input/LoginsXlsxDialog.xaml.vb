Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports iqb.lib.openxml

Public Class LoginsXlsxDialog
    Private Const codeCharacters = "abcdefghprqstuvxyz"
    Private Const codeNumbers = "2345679"

#Region "Vorspann"
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        BtnClose.Visibility = Windows.Visibility.Collapsed

        Process1_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
        Process1_bw.RunWorkerAsync()
    End Sub

    Private WithEvents Process1_bw As ComponentModel.BackgroundWorker = Nothing

    Private Sub BtnCancel_Click() Handles BtnCancel.Click
        If Process1_bw IsNot Nothing AndAlso Process1_bw.IsBusy Then
            Process1_bw.CancelAsync()
            BtnCancel.IsEnabled = False
        Else
            DialogResult = False
        End If
    End Sub

    Private Sub BtnClose_Click() Handles BtnClose.Click
        DialogResult = False
    End Sub

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ComponentModel.ProgressChangedEventArgs) Handles Process1_bw.ProgressChanged
        MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub Process1_bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles Process1_bw.RunWorkerCompleted
        MBUC.AddMessage("beendet")
        BtnCancel.Visibility = Windows.Visibility.Collapsed

        If e.Cancelled Then MBUC.AddMessage("durch Nutzer abgebrochen.")

        BtnClose.Visibility = Windows.Visibility.Visible
    End Sub
#End Region

#Region "Process1_bw_DoWork"

    '######################################################################################
    '######################################################################################
    Private Sub Process1_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles Process1_bw.DoWork
        Dim myworker As ComponentModel.BackgroundWorker = sender

        myworker.ReportProgress(20.0#, "Schreibe Datei '" + IO.Path.GetFileName(My.Settings.lastfile_OutputTargetXlsx))
        Dim sourceFile As Byte() = Nothing
        Try
            Dim TmpZielXLS As SpreadsheetDocument = SpreadsheetDocument.Create(My.Settings.lastfile_OutputTargetXlsx, SpreadsheetDocumentType.Workbook)
            Dim myWorkbookPart As WorkbookPart = TmpZielXLS.AddWorkbookPart()
            myWorkbookPart.Workbook = New Workbook()
            myWorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())
            TmpZielXLS.Close()

            sourceFile = IO.File.ReadAllBytes(My.Settings.lastfile_OutputTargetXlsx)
        Catch ex As Exception
            myworker.ReportProgress(0.0#, "e: Konnte Datei '" + My.Settings.lastfile_OutputTargetXlsx + "' nicht schreiben (noch geöffnet?)" + vbNewLine + ex.Message)
        End Try

        If sourceFile IsNot Nothing Then
            Using MemStream As New IO.MemoryStream()
                MemStream.Write(sourceFile, 0, sourceFile.Length)
                Using sourceXLS As SpreadsheetDocument = SpreadsheetDocument.Open(MemStream, True)
                    Dim myStyles As ExcelStyleDefs = xlsxFactory.AddIQBStandardStyles(sourceXLS.WorkbookPart)
                    Dim TableResponses As WorksheetPart = xlsxFactory.InsertWorksheet(sourceXLS.WorkbookPart, "Logins")
                    xlsxFactory.SetCellValueString("A", 1, TableResponses, "2 Stellen", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("A", TableResponses, 10)
                    xlsxFactory.SetCellValueString("B", 1, TableResponses, "3 Stellen", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("B", TableResponses, 10)
                    xlsxFactory.SetCellValueString("C", 1, TableResponses, "4 Stellen", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("C", TableResponses, 10)
                    xlsxFactory.SetCellValueString("D", 1, TableResponses, "5 Stellen", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("D", TableResponses, 10)

                    Dim codes2 As List(Of String) = GetNewCodeList(2, 100)
                    Dim codes3 As List(Of String) = GetNewCodeList(3, 200)
                    Dim codes4 As List(Of String) = GetNewCodeList(4, 500)
                    Dim codes5 As List(Of String) = GetNewCodeList(5, 2000)


                    For zeile As Integer = 1 To 2000
                        Dim myRowData As New List(Of RowData)
                        If zeile <= codes2.Count Then myRowData.Add(New RowData With {.Column = "A", .Value = codes2.Item(zeile - 1), .CellType = CellTypes.str})
                        If zeile <= codes3.Count Then myRowData.Add(New RowData With {.Column = "B", .Value = codes3.Item(zeile - 1), .CellType = CellTypes.str})
                        If zeile <= codes4.Count Then myRowData.Add(New RowData With {.Column = "C", .Value = codes4.Item(zeile - 1), .CellType = CellTypes.str})
                        If zeile <= codes5.Count Then myRowData.Add(New RowData With {.Column = "D", .Value = codes5.Item(zeile - 1), .CellType = CellTypes.str})
                        xlsxFactory.AppendRow(zeile + 1, myRowData, TableResponses)
                    Next
                End Using
                Try
                    Using fs As New IO.FileStream(My.Settings.lastfile_OutputTargetXlsx, IO.FileMode.Create)
                        MemStream.WriteTo(fs)
                    End Using
                Catch ex As Exception
                    myworker.ReportProgress(100.0#, "e: Konnte Datei nicht schreiben: " + ex.Message)
                End Try
            End Using
        End If
    End Sub

#End Region


    Public Shared Function GetNewCodeList(codeLen As Integer, codeCount As Integer) As List(Of String)
        Dim codeList As New List(Of String)
        Randomize()
        For i As Integer = 1 To codeCount
            Dim newCode As String
            Do
                newCode = ""
                Dim isNumber As Boolean = False
                Do
                    newCode = newCode & IIf(isNumber, Mid(codeNumbers, Int(codeNumbers.Length * Rnd() + 1), 1), Mid(codeCharacters, Int(codeCharacters.Length * Rnd() + 1), 1))
                    isNumber = Not isNumber
                Loop Until newCode.Length = codeLen
            Loop While codeList.Contains(newCode)
            codeList.Add(newCode)
        Next
        Return codeList
    End Function
End Class

