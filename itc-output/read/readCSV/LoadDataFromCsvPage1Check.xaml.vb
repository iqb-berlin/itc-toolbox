Public Class LoadDataFromCsvPage1Check
    Private Sub Me_Loaded() Handles Me.Loaded
        Dim SearchDir As New IO.DirectoryInfo(My.Settings.lastdir_OutputSource)
        Dim LogFileCount As Integer = 0
        Dim ResponsesFileCount As Integer = 0
        Dim parentDlg As OutputDialog = Me.Parent
        For Each fi As IO.FileInfo In SearchDir.GetFiles("*.csv", IO.SearchOption.AllDirectories)
            Try
                Dim readFile As System.IO.TextReader = New IO.StreamReader(fi.FullName)
                Dim line As String = readFile.ReadLine().Replace("""", "")
                If line = LogSymbols.LogFileFirstLine2024 OrElse line = LogSymbols.LogFileFirstLineLegacy Then
                    LogFileCount += 1
                ElseIf line = ResponseSymbols.ResponsesFileFirstLine2019 OrElse
                    line = ResponseSymbols.ResponsesFileFirstLineLegacy OrElse line = ResponseSymbols.ResponsesFileFirstLine2024 Then
                    ResponsesFileCount += 1
                Else
                    Me.MBUC.AddMessage("w: Datei nicht erkannt: " + fi.Name)
                End If
            Catch ex As Exception
                Me.MBUC.AddMessage("e: Fehler beim Lesen der Datei " + fi.Name + "; noch geöffnet?")
            End Try
        Next
        Dim msgLog As String = IIf(parentDlg.WriteToXls, "", LogFileCount.ToString + " Log-Dateien und ")
        Me.MBUC.AddMessage(msgLog + ResponsesFileCount.ToString + " Antwortdateien erkannt.")
    End Sub

    Private Sub BtnOk_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnOK.Click
        Dim parentDlg As OutputDialog = Me.Parent
        If parentDlg.WriteToXls Then
            inputTargetFileName()
        Else
            Me.NavigationService.Navigate(New LoadDataFromCsvPage2Result)
        End If
    End Sub

    Private Sub inputTargetFileName()
        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_OutputTargetXlsx) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_OutputTargetXlsx)
        Dim filepicker As New Microsoft.Win32.SaveFileDialog With {.FileName = My.Settings.lastfile_OutputTargetXlsx, .Filter = "Excel-Dateien|*.xlsx",
                                                        .InitialDirectory = defaultDir, .DefaultExt = "xlsx", .Title = "Antworten Zieldatei wählen"}
        If filepicker.ShowDialog Then
            My.Settings.lastfile_OutputTargetXlsx = filepicker.FileName
            My.Settings.Save()

            Me.NavigationService.Navigate(New LoadDataFromCsvPage2Result)
        End If
    End Sub
End Class
