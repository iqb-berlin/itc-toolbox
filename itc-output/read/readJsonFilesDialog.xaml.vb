Public Class readJsonFilesDialog
    Private files As String()

#Region "Vorspann"
    Public Sub New(fileNames As String())
        InitializeComponent()
        files = fileNames
    End Sub

    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        BtnClose.Visibility = Windows.Visibility.Collapsed

        If files.Length > 0 Then
            BtnClose.Visibility = Windows.Visibility.Collapsed
            BtnCancel.Visibility = Windows.Visibility.Visible
            Process1_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
            Process1_bw.RunWorkerAsync()
        Else
            MBUC.AddMessage("Keine Dateien ausgewählt.")
        End If
    End Sub

    Private WithEvents Process1_bw As ComponentModel.BackgroundWorker = Nothing
    Private WithEvents Process2_bw As ComponentModel.BackgroundWorker = Nothing

    Private Sub BtnCancel_Click() Handles BtnCancel.Click
        If Process1_bw IsNot Nothing AndAlso Process1_bw.IsBusy Then
            Process1_bw.CancelAsync()
            BtnCancel.IsEnabled = False
        ElseIf Process2_bw IsNot Nothing AndAlso Process2_bw.IsBusy Then
            Process2_bw.CancelAsync()
            BtnCancel.IsEnabled = False
        Else
            DialogResult = False
        End If
    End Sub

    Private Sub BtnClose_Click() Handles BtnClose.Click
        DialogResult = True
    End Sub

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ComponentModel.ProgressChangedEventArgs) Handles Process1_bw.ProgressChanged, Process2_bw.ProgressChanged
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
#End Region

    '######################################################################################
    '######################################################################################
    Private Sub Process1_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles Process1_bw.DoWork
        Dim myworker As ComponentModel.BackgroundWorker = sender
        JsonReadWrite.Read(files, myworker)
    End Sub
End Class
