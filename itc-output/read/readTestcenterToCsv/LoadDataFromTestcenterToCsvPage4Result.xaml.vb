Imports Newtonsoft.Json
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports iqb.lib.openxml
Imports System.ComponentModel

Public Class LoadDataFromTestcenterToCsvPage4Result
    Private WithEvents myBackgroundWorker As BackgroundWorker = Nothing

    Private Sub Me_Loaded() Handles Me.Loaded
        Dim ParentDlg As LoadDataFromTestcenterToCsvDialog = Me.Parent
        Me.BtnCancelClose.IsEnabled = False

        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_OutputTargetCsv) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_OutputTargetCsv)
        Dim filepicker As New Microsoft.Win32.SaveFileDialog With {.FileName = My.Settings.lastfile_OutputTargetCsv, .Filter = "CSV-Dateien|*.csv",
                                                        .InitialDirectory = defaultDir, .DefaultExt = "csv", .Title = "Antworten Zieldatei wählen"}
        If filepicker.ShowDialog Then
            My.Settings.lastfile_OutputTargetCsv = filepicker.FileName
            My.Settings.Save()

            myBackgroundWorker = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
            myBackgroundWorker.RunWorkerAsync()
        Else
            ParentDlg.DialogResult = False
        End If
    End Sub

    Private Sub myBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles myBackgroundWorker.DoWork
        Dim myBW As BackgroundWorker = sender
        Dim ParentDlg As LoadDataFromTestcenterToCsvDialog = Me.Parent
        Dim maxProgressValue As Integer = ParentDlg.selectedDataGroups.Count
        Dim progressValue As Integer = 0
        Dim fatalError As Boolean = False

        Dim TargetFile As New IO.FileStream(My.Settings.lastfile_OutputTargetCsv, IO.FileMode.Create)
        For Each dataGroupId As String In ParentDlg.selectedDataGroups
            If myBW.CancellationPending Then Exit For
            If myBW.CancellationPending Then
                e.Cancel = True
                Exit For
            End If
            If fatalError Then Exit For

            Try
                Using buffer As New IO.MemoryStream
                    myBW.ReportProgress(progressValue * 100 / maxProgressValue, dataGroupId)
                    If Not ParentDlg.itcConnection.getResponsesCsvFile(dataGroupId, buffer, myBW) Then
                        myBW.ReportProgress(progressValue * 100 / maxProgressValue, "- unbekannter Fehler")
                    End If
                    TargetFile.Write(buffer.ToArray, 0, buffer.Length)
                End Using
            Catch ex As Exception
                myBW.ReportProgress(progressValue * 100 / maxProgressValue, "e:Konnte Gruppe nicht lesen: " + ex.Message)
            End Try

            progressValue += 1
        Next
        TargetFile.Close()
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
        Dim ParentDlg As LoadDataFromTestcenterToCsvDialog = Me.Parent
        ParentDlg.DialogResult = False
    End Sub

End Class
