Imports iqb.lib.components

Public Class RegisterWebCredentialsPage
    Private Sub Me_Loaded() Handles Me.Loaded
        CrUC.UserCredentials = New Net.NetworkCredential(My.Settings.lastlogin_name, "")
    End Sub

    Private Sub BtnCancel_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim ParentDlg As RegisterWebDialog = Me.Parent
        ParentDlg.DialogResult = False
    End Sub

    Private Sub BtnContinue_Click(sender As Object, e As RoutedEventArgs)
        If String.IsNullOrEmpty(CrUC.UserCredentials.UserName) OrElse String.IsNullOrEmpty(CrUC.UserCredentials.Password) Then
            DialogFactory.MsgError(Me, Me.Title, "Bitte geben Sie Namen und Kennwort ein!")
        Else
            My.Settings.lastlogin_name = CrUC.UserCredentials.UserName
            My.Settings.Save()

            Dim ParentDlg As RegisterWebDialog = Me.Parent
            ParentDlg.itcConnection = New ITCConnection("https://www.iqb-testcenter.de/api", CrUC.UserCredentials)
            Me.NavigationService.Navigate(New RegisterWebResultPage)
        End If
    End Sub
End Class
