Imports System.ComponentModel
Imports iqb.lib.components

Public Class LoadDataFromTestcenterPage1Credentials
    Private WithEvents myBackgroundWorker As BackgroundWorker = Nothing
    Private myConnection As ITCConnection = Nothing
    Private credentials As Net.NetworkCredential = Nothing
    Private url As String
    Private Sub Me_Loaded() Handles Me.Loaded
        TBUrl.Text = My.Settings.lastServerUrl
        CrUC.UserCredentials = New Net.NetworkCredential(My.Settings.lastlogin_name, "")
        APBUC.UpdateProgressState(0.0#)
        If globalOutputStore.itcConnection Is Nothing Then
            DPOldLogin.Visibility = Visibility.Collapsed
            TBUrl.Focus()
        Else
            LbLoginTip.Visibility = Visibility.Collapsed
            TBUrl.IsEnabled = False
            CrUC.IsEnabled = False
        End If
    End Sub

    Private Sub BtnCancel_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        ParentDlg.DialogResult = False
    End Sub

    Private Sub BtnContinue_Click(sender As Object, e As RoutedEventArgs)
        If globalOutputStore.itcConnection Is Nothing Then
            If String.IsNullOrEmpty(CrUC.UserCredentials.UserName) OrElse String.IsNullOrEmpty(CrUC.UserCredentials.Password) Then
                DialogFactory.MsgError(Me, Me.Title, "Bitte geben Sie Namen und Kennwort ein!")
            ElseIf String.IsNullOrEmpty(TBUrl.Text) Then
                DialogFactory.MsgError(Me, Me.Title, "Bitte geben Sie die Url eines Testcenters ein!")
            Else
                My.Settings.lastlogin_name = CrUC.UserCredentials.UserName
                My.Settings.lastServerUrl = TBUrl.Text
                My.Settings.Save()
                BtnCancel.IsEnabled = False
                BtnContinue.IsEnabled = False
                Me.url = TBUrl.Text
                If Not Me.url.StartsWith("http", StringComparison.CurrentCultureIgnoreCase) Then Me.url = "https://" + Me.url
                If Me.url.EndsWith("/") Then Me.url = Me.url.Substring(0, Me.url.Length - 1)
                If Not Me.url.EndsWith("/api", StringComparison.CurrentCultureIgnoreCase) Then Me.url = Me.url + "/api"
                Me.credentials = CrUC.UserCredentials
                myBackgroundWorker = New BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
                myBackgroundWorker.RunWorkerAsync()
            End If
        Else
            Me.NavigationService.Navigate(New LoadDataFromTestcenterPage2SelectWorkspace)
        End If
    End Sub

    Private Sub myBackgroundWorker_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs) Handles myBackgroundWorker.DoWork
        Dim myBW As BackgroundWorker = sender
        myConnection = New ITCConnection(Me.url, Me.credentials, myBW)
    End Sub

    Private Sub myBackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs) Handles myBackgroundWorker.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
    End Sub

    Private Sub myBackgroundWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs) Handles myBackgroundWorker.RunWorkerCompleted
        BtnCancel.IsEnabled = True
        BtnContinue.IsEnabled = True
        Me.APBUC.UpdateProgressState(0.0#)
        If String.IsNullOrEmpty(myConnection.lastErrorMsgText) Then
            globalOutputStore.itcConnection = myConnection
            Me.NavigationService.Navigate(New LoadDataFromTestcenterPage2SelectWorkspace)
        Else
            DialogFactory.MsgError(Me, Me.Title + " - Fehler", "Es ist ein Fehler aufgetreten beim Verbindungsversuch: " + myConnection.lastErrorMsgText)
        End If
    End Sub


    Private Sub BtnOldLogin_Click(sender As Object, e As RoutedEventArgs)
        DPOldLogin.Visibility = Visibility.Collapsed
        LbLoginTip.Visibility = Visibility.Visible
        TBUrl.IsEnabled = True
        CrUC.IsEnabled = True
        globalOutputStore.itcConnection = Nothing
    End Sub
End Class
