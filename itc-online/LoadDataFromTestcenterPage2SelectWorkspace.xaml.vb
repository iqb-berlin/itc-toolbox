Imports System.ComponentModel

Public Class LoadDataFromTestcenterPage2SelectWorkspace
    Private Sub Me_Loaded() Handles Me.Loaded
        LBIDList.ItemsSource = From ws As KeyValuePair(Of Integer, String) In globalOutputStore.itcConnection.accessTo Order By ws.Value
        If globalOutputStore.itcConnection.selectedWorkspace > 0 Then LBIDList.SelectedValue = globalOutputStore.itcConnection.selectedWorkspace
    End Sub

    Private Sub BtnCancel_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        ParentDlg.DialogResult = False
    End Sub

    Private Sub BtnContinue_Click(sender As Object, e As RoutedEventArgs)
        If LBIDList.SelectedItems.Count > 0 Then
            globalOutputStore.itcConnection.selectedWorkspace = LBIDList.SelectedValue
            Me.NavigationService.Navigate(New LoadDataFromTestcenterPage3SelectGroups)
        End If
    End Sub
End Class
