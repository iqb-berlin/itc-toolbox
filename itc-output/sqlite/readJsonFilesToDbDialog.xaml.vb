Imports Newtonsoft.Json

Public Class readJsonFilesToDbDialog
    Private files As String() = Nothing
    Public SqliteDB As SQLiteConnector = Nothing

#Region "Vorspann"
    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        If SqliteDB Is Nothing Then
            APBUC.Value = 0.0#
            MBUC.AddMessage("Keine Datenbank zugewiesen")
            BtnCancel.Content = "Schließen"
        Else
            Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            If Not String.IsNullOrEmpty(My.Settings.lastfile_InputTargetJson) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_InputTargetJson)
            Dim filepicker As New Microsoft.Win32.OpenFileDialog With {.FileName = IO.Path.GetFileName(My.Settings.lastfile_InputTargetJson), .Filter = "JSON-Dateien|*.json",
                .InitialDirectory = defaultDir, .DefaultExt = "json", .Multiselect = True, .Title = "JSON Daten einlesen - Wähle Datei(en)"}
            If filepicker.ShowDialog Then
                My.Settings.lastfile_InputTargetJson = filepicker.FileName
                My.Settings.Save()

                files = filepicker.FileNames
                If files.Length > 0 Then Me.DialogResult = False

                Process1_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
                Process1_bw.RunWorkerAsync()
            End If
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

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ComponentModel.ProgressChangedEventArgs) Handles Process1_bw.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
        If Not String.IsNullOrEmpty(e.UserState) Then MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub Process1_bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles Process1_bw.RunWorkerCompleted
        APBUC.Value = 0.0#
        MBUC.AddMessage("beendet")
        BtnCancel.Content = "Schließen"
        If e.Cancelled Then MBUC.AddMessage("durch Nutzer abgebrochen.")
    End Sub
#End Region

    '######################################################################################
    '######################################################################################
    Private Sub Process1_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles Process1_bw.DoWork
        Dim myworker As ComponentModel.BackgroundWorker = sender
        Dim progressMax As Integer = files.Length
        Dim progressCount As Integer = 1
        Dim progressValue As Double
        For Each fn In files
            progressValue = progressCount * (100 / progressMax)
            progressCount += 1
            myworker.ReportProgress(progressValue, IO.Path.GetFileName(fn))
            Try
                Using file As New IO.StreamReader(fn)
                    Dim js As New JsonSerializer()
                    Dim groupData As List(Of Person) = js.Deserialize(file, GetType(List(Of Person)))
                    For Each p As Person In groupData
                        SqliteDB.addPerson(p)
                    Next
                End Using
            Catch ex As Exception
                myworker.ReportProgress(progressValue, "Fehler " + IO.Path.GetFileName(fn) + ": " + ex.Message)
            End Try
        Next
    End Sub
End Class
