Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging

Public Class ToXlsxDialog
    Public sqliteConnection As SQLiteConnector
#Region "Vorspann"
    Public Sub New(sqliteConnection As SQLiteConnector)
        InitializeComponent()
        Me.sqliteConnection = sqliteConnection
    End Sub

    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        BtnClose.Visibility = Windows.Visibility.Collapsed

        If String.IsNullOrEmpty(My.Settings.lastfile_OutputTargetXlsx) Then
            BtnContinue.Visibility = Windows.Visibility.Collapsed
            MBUC.AddMessage("Keine Zieldatei gewählt.")
        Else
            MBUC.AddMessage("Zieldatei: " + IO.Path.GetFileName(My.Settings.lastfile_OutputTargetXlsx))
            MBUC.AddMessage("Bitte Optionen wählen")
        End If
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
        DialogResult = True
    End Sub

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ComponentModel.ProgressChangedEventArgs) Handles Process1_bw.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
        If Not String.IsNullOrEmpty(e.UserState) Then MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub Process1_bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles Process1_bw.RunWorkerCompleted
        APBUC.Value = 0.0#
        MBUC.AddMessage("beendet")
        BtnCancel.Visibility = Windows.Visibility.Collapsed
        If e.Cancelled Then MBUC.AddMessage("durch Nutzer abgebrochen.")

        BtnClose.Visibility = Windows.Visibility.Visible
    End Sub

    Private Sub BtnContinue_Click() Handles BtnContinue.Click
        If ChBResonses.IsChecked Then
            DPParameters.IsEnabled = False
            BtnClose.Visibility = Windows.Visibility.Collapsed
            BtnContinue.Visibility = Windows.Visibility.Collapsed
            BtnCancel.Visibility = Windows.Visibility.Visible

            Process1_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
            Process1_bw.RunWorkerAsync()
        Else
            MBUC.AddMessage("Bitte Optionen wählen")
        End If
    End Sub

#End Region

    '######################################################################################
    '######################################################################################
    Private Sub Process1_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles Process1_bw.DoWork
        Dim myworker As ComponentModel.BackgroundWorker = sender
        Dim targetXlsxFilename As String = My.Settings.lastfile_OutputTargetXlsx
        Dim myTemplate As Byte() = Nothing
        Try
            Dim TmpZielXLS As SpreadsheetDocument = SpreadsheetDocument.Create(targetXlsxFilename, SpreadsheetDocumentType.Workbook)
            Dim myWorkbookPart As WorkbookPart = TmpZielXLS.AddWorkbookPart()
            myWorkbookPart.Workbook = New Workbook()
            myWorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())
            TmpZielXLS.Close()

            myTemplate = IO.File.ReadAllBytes(targetXlsxFilename)
        Catch ex As Exception
            myworker.ReportProgress(0.0#, "e: Konnte Datei '" + targetXlsxFilename + "' nicht schreiben (noch geöffnet?)" + vbNewLine + ex.Message)
        End Try

        If myTemplate IsNot Nothing Then
            If Not myworker.CancellationPending Then WriteOutputToXlsx.Write(myTemplate, myworker, e, New WriteXlsxConfig With {
                    .targetXlsxFilename = targetXlsxFilename,
                    .sourceDatabase = sqliteConnection,
                    .subformMode = SubformMode.None
                })
        End If
    End Sub
End Class
