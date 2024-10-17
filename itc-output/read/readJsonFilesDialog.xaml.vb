Public Class readJsonFilesDialog
    Private files As String() = Nothing
    Private Shared readResponses As Boolean = True
    Private Shared ignoreDisplayed As Boolean = True
    Private Shared ignoreNotReached As Boolean = True
    Private Shared readLogs As Boolean = False

#Region "Vorspann"
    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        ChBResponses.IsChecked = readResponses
        ChBResponsesIgnoreDisplayed.IsChecked = ignoreDisplayed
        ChBResponsesIgnoreNotReached.IsChecked = ignoreNotReached
        ChBLogs.IsChecked = readLogs
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

    Private Sub BtnContinue_Click() Handles BtnContinue.Click
        If Process1_bw Is Nothing Then
            Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            If Not String.IsNullOrEmpty(My.Settings.lastfile_InputTargetJson) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_InputTargetJson)
            Dim filepicker As New Microsoft.Win32.OpenFileDialog With {.FileName = IO.Path.GetFileName(My.Settings.lastfile_InputTargetJson), .Filter = "JSON-Dateien|*.json",
                .InitialDirectory = defaultDir, .DefaultExt = "json", .Multiselect = True, .Title = "JSON Daten einlesen - Wähle Datei(en)"}
            If filepicker.ShowDialog Then
                My.Settings.lastfile_InputTargetJson = filepicker.FileName
                My.Settings.Save()

                files = filepicker.FileNames
                If files.Length > 0 Then
                    BtnContinue.Visibility = Windows.Visibility.Collapsed
                    readResponses = ChBResponses.IsChecked
                    ignoreDisplayed = ChBResponsesIgnoreDisplayed.IsChecked
                    ignoreNotReached = ChBResponsesIgnoreNotReached.IsChecked
                    readLogs = ChBLogs.IsChecked

                    Process1_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
                    Process1_bw.RunWorkerAsync()
                End If
            End If
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
        If readResponses Then
            If readLogs Then
                JsonReadWrite.ReadFull(files, myworker)
            Else
                JsonReadWrite.ReadResponsesOnly(files, myworker, ignoreDisplayed, ignoreNotReached)
            End If
        ElseIf readLogs Then
            JsonReadWrite.ReadLogsOnly(files, myworker)
        Else
            MBUC.AddMessage("Weder Antworten noch Logs ausgewählt.")
        End If
    End Sub

    Private Sub UpdateTarget(sender As Object, e As RoutedEventArgs)
        If ChBResponses.IsChecked Then
            If ChBLogs.IsChecked Then
                TBTarget.Text = "Ziel: Volldaten-Store"
            Else
                TBTarget.Text = "Ziel: Nur-Antworten-Store"
            End If
        ElseIf ChBLogs.IsChecked Then
            TBTarget.Text = "Ziel: Nur-Logs-Store"
        Else
            TBTarget.Text = ""
        End If
    End Sub
End Class
