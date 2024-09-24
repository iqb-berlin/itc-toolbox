Imports System.ComponentModel

Public Class LoadDataFromTestcenterPage3SelectGroups
    Private Sub Me_Loaded() Handles Me.Loaded
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        If ParentDlg.ResponsesOnly Then
            Me.BtnResponses.Content = "Weiter"
            Me.BtnReviews.Visibility = Visibility.Collapsed
        End If
        Dim dataGroups As List(Of GroupDataDTO) = ParentDlg.itcConnection.getDataGroups()
        ICDataGroups.ItemsSource = From ds As GroupDataDTO In dataGroups Order By ds.groupName
                                   Where ds.bookletsStarted > 0
                                   Let xGroup = New XElement(<g checked="true"><%= ds.groupName %></g>)
                                   Select xGroup
    End Sub

    Private Sub BtnCancel_Click(sender As System.Object, e As System.Windows.RoutedEventArgs)
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        ParentDlg.DialogResult = False
    End Sub

    Private Sub BtnResponses_Click(sender As Object, e As RoutedEventArgs)
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        ParentDlg.selectedDataGroups = (From xe As XElement In ICDataGroups.Items Where xe.@checked = "true" Select xe.Value).ToList
        Me.NavigationService.Navigate(New LoadDataFromTestcenterPage4Responses)
    End Sub

    Private Sub BtnReviews_Click(sender As Object, e As RoutedEventArgs)
        Dim ParentDlg As LoadDataFromTestcenterDialog = Me.Parent
        ParentDlg.selectedDataGroups = (From xe As XElement In ICDataGroups.Items Where xe.@checked = "true" Select xe.Value).ToList
        Me.NavigationService.Navigate(New LoadDataFromTestcenterPage4ReviewsXlsx)
    End Sub
End Class
