Imports System.ComponentModel

Public Class RegisterWebResultPage
    Private WithEvents myBackgroundWorker As BackgroundWorker = Nothing
    Private studyChanged As Boolean = False

    Private Sub Me_Loaded() Handles Me.Loaded
        Me.MBUC.AddMessage("Bitte warten!")
        Me.BtnCancelClose.IsEnabled = False

        myBackgroundWorker = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
        myBackgroundWorker.RunWorkerAsync()
    End Sub

    Private Sub myBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles myBackgroundWorker.DoWork
        Dim myBW As BackgroundWorker = sender
        Dim ParentDlg As RegisterWebDialog = Me.Parent
        Dim newStudyOnlineKey As String = System.Guid.NewGuid.ToString
        studyChanged = True
    End Sub

    Private Sub myBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles myBackgroundWorker.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
        If Not String.IsNullOrEmpty(e.UserState) Then Me.MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub myBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles myBackgroundWorker.RunWorkerCompleted
        Me.BtnCancelClose.IsEnabled = True
        Me.APBUC.UpdateProgressState(0.0#)
        Dim ParentDlg As RegisterWebDialog = Me.Parent
        For Each w As KeyValuePair(Of Integer, String) In ParentDlg.itcConnection.accessTo
            Me.MBUC.AddMessage(w.Value)
        Next
        Me.MBUC.AddMessage("i: Beendet.")
        BtnCancelClose.IsEnabled = True
    End Sub

    Private Sub BtnCancelClose_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim ParentDlg As RegisterWebDialog = Me.Parent
        ParentDlg.DialogResult = IIf(studyChanged, True, False)
    End Sub

End Class
