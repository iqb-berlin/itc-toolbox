Imports System.ComponentModel

Public Class LoadDataFromTestcenterPage3SelectGroups
    Private dataGroupsToSelect As Dictionary(Of String, String)
    Private Sub Me_Loaded() Handles Me.Loaded
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        Dim dataGroups As List(Of GroupDataDTO) = ParentDlg.itcConnection.getDataGroups()
        dataGroupsToSelect = (From ds As GroupDataDTO In dataGroups
                              Order By ds.groupName
                              Where ds.bookletsStarted > 0).ToDictionary(Function(a) a.groupName,
                                                                         Function(a) IIf(a.groupLabel = a.groupName, "", " - " + a.groupLabel).ToString)
        ICDataGroups.ItemsSource = dataGroupsToSelect.Select(Of XElement)(Function(g, index) New XElement(<g checked="True" number=<%= index %> label=<%= g.Value %>><%= g.Key %></g>))
    End Sub

    Private Sub BtnCancel_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        ParentDlg.DialogResult = False
    End Sub

    Private Sub BtnContinue_Click(sender As Object, e As RoutedEventArgs)
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        ParentDlg.selectedDataGroups = (From xe As XElement In ICDataGroups.Items Where xe.@checked = "True" Select xe.Value).ToList
        If ParentDlg.readMode = TestcenterReadMode.Responses Then
            Me.NavigationService.Navigate(New LoadDataFromTestcenterPage4Responses)
        Else
            Me.NavigationService.Navigate(New LoadDataFromTestcenterPage4ReviewsXlsx)
        End If
    End Sub

    Private Sub BtnToggleCheck_Click(sender As Object, e As RoutedEventArgs)
        If ICDataGroups.Items.Count > 0 Then
            Dim firstItem As XElement = ICDataGroups.Items.Item(0)
            Dim newValue As String = "True"
            If firstItem.@checked = "True" Then newValue = "False"
            ICDataGroups.ItemsSource = dataGroupsToSelect.Select(Of XElement)(Function(g, index) New XElement(<g checked=<%= newValue %> number=<%= index %> label=<%= g.Value %>><%= g.Key %></g>))
        End If
    End Sub
End Class
