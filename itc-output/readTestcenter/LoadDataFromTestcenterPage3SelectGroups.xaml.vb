Imports System.ComponentModel

Public Class LoadDataFromTestcenterPage3SelectGroups
    Private dataGroupNames As List(Of String)
    Private Sub Me_Loaded() Handles Me.Loaded
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        If ParentDlg.readMode <> TestcenterReadMode.Responses Then Me.CBBigData.Visibility = Visibility.Collapsed
        Dim dataGroups As List(Of GroupDataDTO) = globalOutputStore.itcConnection.getDataGroups()
        dataGroupNames = (From ds As GroupDataDTO In dataGroups Order By ds.groupName Where ds.bookletsStarted > 0 Select ds.groupName).ToList
        ICDataGroups.ItemsSource = dataGroupNames.Select(Of XElement)(Function(name, index) New XElement(<g checked="true" number=<%= index %>><%= name %></g>))
    End Sub

    Private Sub BtnCancel_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        ParentDlg.DialogResult = False
    End Sub

    Private Sub BtnContinue_Click(sender As Object, e As RoutedEventArgs)
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        ParentDlg.selectedDataGroups = (From xe As XElement In ICDataGroups.Items Where xe.@checked = "true" Select xe.Value).ToList
        If ParentDlg.readMode = TestcenterReadMode.Responses Then
            ParentDlg.replaceBigdata = Not CBBigData.IsChecked
            Me.NavigationService.Navigate(New LoadDataFromTestcenterPage4Responses)
        Else
            Me.NavigationService.Navigate(New LoadDataFromTestcenterPage4ReviewsXlsx)
        End If
    End Sub

    Private Sub BtnToggleCheck_Click(sender As Object, e As RoutedEventArgs)
        If ICDataGroups.Items.Count > 0 Then
            Dim firstItem As XElement = ICDataGroups.Items.Item(0)
            Dim newValue As String = "true"
            If firstItem.@checked = "true" Then newValue = "false"
            ICDataGroups.ItemsSource = dataGroupNames.Select(Of XElement)(Function(name, index) New XElement(<g checked=<%= newValue %> number=<%= index %>><%= name %></g>))
        End If
    End Sub
End Class
