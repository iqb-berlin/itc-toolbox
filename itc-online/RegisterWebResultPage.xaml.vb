Imports System.ComponentModel

Public Class RegisterWebResultPage
    Private WithEvents myBackgroundWorker As BackgroundWorker = Nothing

    Private Sub Me_Loaded() Handles Me.Loaded
        Me.MBUC.AddMessage("Bitte warten!")
        Me.BtnCancelClose.IsEnabled = False

        myBackgroundWorker = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
        myBackgroundWorker.RunWorkerAsync()
    End Sub

    Private Sub myBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles myBackgroundWorker.DoWork
        Dim myBW As BackgroundWorker = sender
        Dim ParentDlg As RegisterWebDialog = Me.Parent
        Dim maxProgressValue As Integer = ParentDlg.itcConnection.accessTo.Count
        Dim progressValue As Integer = 0
        Dim wsIdList As New List(Of Integer)(ParentDlg.itcConnection.accessTo.Keys)
        For Each workspaceId As Integer In wsIdList
            progressValue += 1
            myBW.ReportProgress(progressValue * 100 / maxProgressValue)
            ParentDlg.itcConnection.GetWorkspaceName(workspaceId)
        Next
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
            Me.MBUC.AddMessage(w.Key.ToString + ": " + w.Value)
        Next
        Me.MBUC.AddMessage("i: Beendet.")
        BtnCancelClose.IsEnabled = True
    End Sub

    Private Sub BtnCancelClose_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim ParentDlg As RegisterWebDialog = Me.Parent
        ParentDlg.DialogResult = False
    End Sub

End Class
