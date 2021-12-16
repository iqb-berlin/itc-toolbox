Public Class OutputReadReportPage
    Private Sub Me_Loaded() Handles Me.Loaded
        Dim SearchDir As New IO.DirectoryInfo(My.Settings.lastdir_OutputSource)
        Dim LogFileCount As Integer = 0
        Dim ResponsesFileCount As Integer = 0
        Dim parentDlg As OutputDialog = Me.Parent
        For Each fi As IO.FileInfo In SearchDir.GetFiles("*.csv", IO.SearchOption.AllDirectories)
            Try
                Dim readFile As System.IO.TextReader = New IO.StreamReader(fi.FullName)
                Dim line As String = readFile.ReadLine()
                If line = OutputDialog.LogFileFirstLine Then
                    LogFileCount += 1
                ElseIf line = OutputDialog.ResponsesFileFirstLine OrElse line = OutputDialog.ResponsesFileFirstLineLegacy Then
                    ResponsesFileCount += 1
                Else
                    Me.MBUC.AddMessage("w: Datei nicht erkannt: " + fi.Name)
                End If
            Catch ex As Exception
                Me.MBUC.AddMessage("e: Fehler beim Lesen der Datei " + fi.Name + "; noch geöffnet?")
            End Try
        Next
        For Each fi As IO.FileInfo In SearchDir.GetFiles("*.yaml", IO.SearchOption.AllDirectories)
            Try
                Dim fileString As String = IO.File.ReadAllText(fi.FullName)
                Dim deserializer As New YamlDotNet.Serialization.Deserializer
                Dim yamlData As OutputConfig = deserializer.Deserialize(fileString, GetType(OutputConfig))
                If yamlData.bookletSizes IsNot Nothing Then
                    If parentDlg.outputConfig.bookletSizes Is Nothing Then
                        parentDlg.outputConfig.bookletSizes = yamlData.bookletSizes
                    Else
                        For Each booklet As KeyValuePair(Of String, Long) In yamlData.bookletSizes
                            If Not parentDlg.outputConfig.bookletSizes.ContainsKey(booklet.Key) Then parentDlg.outputConfig.bookletSizes.Add(booklet.Key, booklet.Value)
                        Next
                    End If
                End If
                If yamlData.omitUnits IsNot Nothing Then
                    If parentDlg.outputConfig.omitUnits Is Nothing Then
                        parentDlg.outputConfig.omitUnits = yamlData.omitUnits
                    Else
                        For Each unitId As String In yamlData.omitUnits
                            If Not parentDlg.outputConfig.omitUnits.Contains(unitId) Then parentDlg.outputConfig.omitUnits.Add(unitId)
                        Next
                    End If
                End If
                If yamlData.variables IsNot Nothing Then
                    If parentDlg.outputConfig.variables Is Nothing Then
                        parentDlg.outputConfig.variables = yamlData.variables
                    Else
                        For Each varDef As KeyValuePair(Of String, Dictionary(Of String, List(Of String))) In yamlData.variables
                            If Not parentDlg.outputConfig.variables.ContainsKey(varDef.Key) Then parentDlg.outputConfig.variables.Add(varDef.Key, varDef.Value)
                        Next
                    End If
                End If
            Catch ex As Exception
                Me.MBUC.AddMessage("w: Fehler beim Lesen der Datei " + fi.Name + "; Syntaxfehler?")
            End Try
        Next
        If parentDlg.outputConfig.omitUnits IsNot Nothing AndAlso parentDlg.outputConfig.omitUnits.Count > 0 Then Me.MBUC.AddMessage("yaml: " + parentDlg.outputConfig.omitUnits.Count.ToString + " Units definiert zum Ignorieren.")
        If parentDlg.outputConfig.variables IsNot Nothing AndAlso parentDlg.outputConfig.variables.Count > 0 Then Me.MBUC.AddMessage("yaml: " + parentDlg.outputConfig.variables.Count.ToString + " Units mit Variablen-Umbenennungen definiert.")
        If parentDlg.outputConfig.bookletSizes IsNot Nothing AndAlso parentDlg.outputConfig.bookletSizes.Count > 0 Then Me.MBUC.AddMessage("yaml: " + parentDlg.outputConfig.bookletSizes.Count.ToString + " Testheft-Größen definiert.")
        Me.MBUC.AddMessage(LogFileCount.ToString + " Log-Dateien und " + ResponsesFileCount.ToString + " Antwortdateien erkannt.")
    End Sub

    Private Sub BtnOk_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnOK.Click
        inputTargetFileName()
    End Sub

    Private Sub inputTargetFileName()
        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_OutputTargetXlsx) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_OutputTargetXlsx)
        Dim filepicker As New Microsoft.Win32.SaveFileDialog With {.FileName = My.Settings.lastfile_OutputTargetXlsx, .Filter = "Excel-Dateien|*.xlsx",
                                                        .InitialDirectory = defaultDir, .DefaultExt = "xlsx", .Title = "Antworten Zieldatei wählen"}
        If filepicker.ShowDialog Then
            My.Settings.lastfile_OutputTargetXlsx = filepicker.FileName
            My.Settings.Save()

            Me.NavigationService.Navigate(New OutputResultPage)
        End If
    End Sub
End Class
