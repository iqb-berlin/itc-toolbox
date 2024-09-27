Imports Newtonsoft.Json
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports iqb.lib.openxml
Imports System.ComponentModel

Public Class LoadDataFromTestcenterPage4ReviewsXlsx
    Private WithEvents myBackgroundWorker As BackgroundWorker = Nothing

    Private Sub Me_Loaded() Handles Me.Loaded
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        Me.BtnCancelClose.IsEnabled = False

        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_OutputTargetXlsx) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_OutputTargetXlsx)
        Dim filepicker As New Microsoft.Win32.SaveFileDialog With {.FileName = My.Settings.lastfile_OutputTargetXlsx, .Filter = "Excel-Dateien|*.xlsx",
                                                        .InitialDirectory = defaultDir, .DefaultExt = "xlsx", .Title = "Antworten Zieldatei wählen"}
        If filepicker.ShowDialog Then
            My.Settings.lastfile_OutputTargetXlsx = filepicker.FileName
            My.Settings.Save()

            myBackgroundWorker = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
            myBackgroundWorker.RunWorkerAsync()
        Else
            ParentDlg.DialogResult = False
        End If
    End Sub

    Private Sub myBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles myBackgroundWorker.DoWork
        Dim myBW As BackgroundWorker = sender
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
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
            myBW.ReportProgress(0.0#, "e: Konnte Datei '" + targetXlsxFilename + "' nicht schreiben (noch geöffnet?)" + vbNewLine + ex.Message)
        End Try

        If myTemplate IsNot Nothing Then
            myBW.ReportProgress(3.0#, "Lese Booklets")
            Dim reviewEntryCount As Integer = 0
            Dim allReviews As New List(Of ReviewDTO)
            Dim maxProgressValue As Integer = ParentDlg.selectedDataGroups.Count * 2
            Dim progressValue As Integer = 0
            For Each dataGroupId As String In ParentDlg.selectedDataGroups
                myBW.ReportProgress(progressValue * 100 / maxProgressValue, "Lese '" + dataGroupId + "': ")
                Dim reviewData As List(Of ReviewDTO) = globalOutputStore.itcConnection.getReviews(dataGroupId)
                If Not String.IsNullOrEmpty(globalOutputStore.itcConnection.lastErrorMsgText) Then
                    myBW.ReportProgress(progressValue * 100 / maxProgressValue, "e: Problem bei Logingruppe '" + dataGroupId + "': " +
                        globalOutputStore.itcConnection.lastErrorMsgText + " (Logs)")
                Else
                    reviewEntryCount += 1
                    allReviews.AddRange(reviewData)
                End If
            Next
            myBW.ReportProgress(0.0#, "Daten für " + reviewEntryCount.ToString("#,##0") + " Testgruppen gelesen.")
            myBW.ReportProgress(0.0#, allReviews.Count.ToString("#,##0") + " Review-Einträge gelesen.")

            If Not myBW.CancellationPending Then WriteReviewsToXlsx.Write(myTemplate, myBW, e, allReviews, targetXlsxFilename)
        End If
    End Sub

    Private Sub myBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles myBackgroundWorker.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
        If Not String.IsNullOrEmpty(e.UserState) Then Me.MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub myBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles myBackgroundWorker.RunWorkerCompleted
        Me.BtnCancelClose.IsEnabled = True
        Me.APBUC.UpdateProgressState(0.0#)
    End Sub

    Private Sub BtnCancelClose_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        ParentDlg.DialogResult = False
    End Sub

End Class
