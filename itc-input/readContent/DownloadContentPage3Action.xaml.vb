Imports System.ComponentModel
Imports System.Security.Policy
Imports DocumentFormat.OpenXml.Office2016.Drawing.Command

Public Class DownloadContentPage3Action
    Private WithEvents myBackgroundWorker As BackgroundWorker = Nothing
    Private fullFileList As List(Of WorkspaceFileDTO)

    Private Sub Me_Loaded() Handles Me.Loaded
        Dim ParentDlg As DownloadContentDialog = Me.Parent
        fullFileList = ParentDlg.itcConnection.getFullFileList()
        If fullFileList.Count > 0 Then
            Me.MBUC.AddMessage(fullFileList.Count.ToString + " Dateien gefunden.")

            Dim folderpicker As New Forms.FolderBrowserDialog With {.Description = "Zielverzeichnis für Dateien",
                                                                .ShowNewFolderButton = True, .SelectedPath = My.Settings.lastfolder_ContentTargetFolder}
            If folderpicker.ShowDialog() AndAlso Not String.IsNullOrEmpty(folderpicker.SelectedPath) Then
                My.Settings.lastfolder_ContentTargetFolder = folderpicker.SelectedPath
                My.Settings.Save()

                myBackgroundWorker = New BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
                myBackgroundWorker.RunWorkerAsync()
            Else
                ParentDlg.DialogResult = False
            End If
        Else
            Me.MBUC.AddMessage("Keine Dateien gefunden.")
            BtnCancelClose.Content = "Schließen"
            BtnCancelClose.IsEnabled = True
        End If
    End Sub

    Private Sub myBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles myBackgroundWorker.DoWork
        Dim myBW As BackgroundWorker = sender
        Dim ParentDlg As DownloadContentDialog = Me.Parent

        Dim LogEntryCount As Long = 0
        Dim maxProgressValue As Integer = fullFileList.Count
        Dim progressValue As Integer = 0
        Dim fatalError As Boolean = False

        myBW.ReportProgress(3.0#, "Lese Dateien")
        For Each fileToDownload As WorkspaceFileDTO In fullFileList
            If myBW.CancellationPending Then Exit For
            If myBW.CancellationPending Then
                e.Cancel = True
                Exit For
            End If
            If fatalError Then Exit For

            Dim TargetFilename As String = My.Settings.lastfolder_ContentTargetFolder + IO.Path.DirectorySeparatorChar + fileToDownload.name

            Dim ProcessStatus As String = "new"
            Try
                Dim createOrOverwrite As Boolean = True
                Dim remoteDT As New DateTime(1970, 1, 1, 0, 0, 0, 0)
                remoteDT = remoteDT.AddSeconds(fileToDownload.modificationTime)
                If IO.File.Exists(TargetFilename) Then
                    Dim fi_pre As New IO.FileInfo(TargetFilename)
                    Dim localeDT As New DateTime
                    With fi_pre.LastWriteTime
                        localeDT = New DateTime(.Year, .Month, .Day, .Hour, .Minute, .Second)
                    End With
                    createOrOverwrite = remoteDT > localeDT
                    If createOrOverwrite Then
                        ProcessStatus = "update"
                    Else
                        ProcessStatus = "skip"
                    End If
                End If
                If createOrOverwrite Then
                    Dim TargetFile As New IO.FileStream(TargetFilename, IO.FileMode.Create)
                    myBW.ReportProgress(progressValue * 100 / maxProgressValue, IO.Path.GetFileName(TargetFile.Name) + " - " + ProcessStatus)
                    If Not ParentDlg.itcConnection.getFile(fileToDownload, TargetFile, myBW) Then
                        myBW.ReportProgress(progressValue * 100 / maxProgressValue, "- unbekannter Fehler")
                    End If
                    TargetFile.Close()
                    Dim fi As New IO.FileInfo(TargetFilename)
                    If myBW.CancellationPending Then
                        fi.Delete()
                    Else
                        fi.LastWriteTime = remoteDT
                    End If
                End If
            Catch ex As Exception
                myBW.ReportProgress(progressValue * 100 / maxProgressValue, "e:Konnte Datei nicht schreiben: " + fileToDownload.name + " (" + ex.Message + ")")
            End Try

            progressValue += 1
        Next
        If Not myBW.CancellationPending AndAlso Not fatalError Then
            myBW.ReportProgress(95.0#, "Schreibe Index")
            Dim filesWithDependency As List(Of WorkspaceFileDTO) = ParentDlg.itcConnection.getFileDependencies((From f As WorkspaceFileDTO In fullFileList
                                                                                                                Where f.type = "Unit" Select f.name).ToList)
            filesWithDependency.AddRange(From f As WorkspaceFileDTO In fullFileList Where f.type <> "Unit")
            IndexHtml.WriteTo(My.Settings.lastfolder_ContentTargetFolder + IO.Path.DirectorySeparatorChar + "index.html",
                              ParentDlg.itcConnection.url, ParentDlg.itcConnection.selectedWorkspaceName, filesWithDependency)
        End If
    End Sub

    Private Sub myBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles myBackgroundWorker.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
        If Not String.IsNullOrEmpty(e.UserState) Then Me.MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub myBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles myBackgroundWorker.RunWorkerCompleted
        APBUC.Value = 0.0#
        MBUC.AddMessage("beendet")
        BtnCancelClose.Content = "Schließen"
        BtnCancelClose.IsEnabled = True
        If e.Cancelled Then MBUC.AddMessage("durch Nutzer abgebrochen.")
        Me.APBUC.UpdateProgressState(0.0#)
    End Sub

    Private Sub BtnCancelClose_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        If myBackgroundWorker IsNot Nothing AndAlso myBackgroundWorker.IsBusy Then
            myBackgroundWorker.CancelAsync()
            BtnCancelClose.IsEnabled = False
            MBUC.AddMessage("Abbruch - bitte warten")
        Else
            Dim ParentDlg As DownloadContentDialog = Me.Parent
            ParentDlg.DialogResult = False
        End If
    End Sub
End Class
