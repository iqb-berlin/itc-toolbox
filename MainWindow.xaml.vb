Imports iqb.lib.components
Class MainWindow
    Private Sub MainApplication_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf MyUnhandledExceptionEventHandler

        Me.Title = My.Application.Info.AssemblyName

        DialogFactory.MainWindow = Me

        Dim ContinueStart As Boolean = True
        Dim ErrMsg As String = "Es gibt ein Problem bei dem Versuch, die alten lokalen Programmeinstellungen zu laden. Bitte deinstallieren Sie die Anwendung über die Systemsteuerung und installieren Sie sie dann erneut!"
        Dim UserConfigFilename As String = ""
        Try
            'neue Programmversion -> alte Settings holen
            If Not My.Settings.updated Then
                My.Settings.Upgrade()
                My.Settings.updated = True
                My.Settings.Save()
            End If
        Catch ex As System.Configuration.ConfigurationException
            ContinueStart = False
            If ex.InnerException Is Nothing Then
                Debug.Print("Configuration.ConfigurationException ohne InnerException")
            Else
                ErrMsg += " Alternativ können Sie die unten genannte Datei löschen (Achtung: Apps ist ein verstecktes Verzeichnis)." + vbNewLine + vbNewLine + ex.InnerException.Message
                Debug.Print(ex.InnerException.Message)
                Dim pos As Integer = ex.InnerException.Message.IndexOf("(")
                If pos > 0 Then
                    UserConfigFilename = ex.InnerException.Message.Substring(pos + 1)
                    pos = UserConfigFilename.IndexOf("\user.config ")
                    If pos > 0 Then
                        UserConfigFilename = UserConfigFilename.Substring(0, pos) + "\user.config"
                        Debug.Print(">>" + UserConfigFilename + "<<")
                    Else
                        UserConfigFilename = ""
                    End If
                End If
            End If
        End Try

        If Not ContinueStart Then
            If Not String.IsNullOrEmpty(UserConfigFilename) AndAlso
                UserConfigFilename.IndexOfAny(IO.Path.GetInvalidFileNameChars()) < 0 AndAlso
                IO.File.Exists(UserConfigFilename) Then
                Try
                    IO.File.Delete(UserConfigFilename)
                    ErrMsg = "Die lokalen Programmeinstellungen mussten gelöscht werden. Bitte starten Sie die Anwendung erneut!"
                Catch ex As Exception
                    ErrMsg += vbNewLine + vbNewLine + "Löschen gescheitert: " + ex.Message
                End Try
            End If
            DialogFactory.MsgError(Me, Me.Title, ErrMsg)
            Me.Close()
        End If

        CommandBindings.Add(New CommandBinding(ApplicationCommands.Help, AddressOf HandleHelpExecuted))
    End Sub

    Private Sub HandleHelpExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_Yaml) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_Yaml)
        Dim filepicker As New Microsoft.Win32.SaveFileDialog With {.FileName = My.Settings.lastfile_Yaml, .Filter = "Yaml-Dateien|*.yaml",
                                                           .DefaultExt = "xlsx", .Title = "Konfiguration - Wähle Ziel-Datei"}
        If filepicker.ShowDialog Then
            My.Settings.lastfile_Yaml = filepicker.FileName
            My.Settings.Save()

            Dim fileString As String = IO.File.ReadAllText(My.Settings.lastfile_Yaml)
            Dim deserializer As New YamlDotNet.Serialization.Deserializer
            Dim yamlData As OutputConfig = deserializer.Deserialize(fileString, GetType(OutputConfig))
            Debug.Print("yo")
            'Dim yamlData As New OutputConfig With {
            '    .bookletSizes = New Dictionary(Of String, Integer) From {
            '        {"THTLK1", 123445},
            '        {"THTLK2", 1234466}
            '    },
            '    .omitUnits = New List(Of String) From {
            '        "ER153ex", "ER342ex"
            '    },
            '    .replaceVariables = New Dictionary(Of String, Dictionary(Of String, List(Of String)))
            '}
            'yamlData.replaceVariables.Add("ER888", New Dictionary(Of String, List(Of String)) From {
            '        {"ER888a", New List(Of String) From {"canvasElement23", "canvasElement24", "canvasElement25"}},
            '        {"ER888b", New List(Of String) From {"canvasElement28"}}
            '    })
            'yamlData.replaceVariables.Add("ER889", New Dictionary(Of String, List(Of String)) From {
            '        {"ER889a", New List(Of String) From {"canvasElement31"}},
            '        {"ER889b", New List(Of String) From {"canvasElement42"}}
            '    })
            'Dim serializer As New YamlDotNet.Serialization.Serializer
            'Dim yamlSerialzed As String = serializer.Serialize(yamlData)
            'IO.File.WriteAllText(My.Settings.lastfile_Yaml, yamlSerialzed)
        End If

    End Sub

    '############################################
    Private Sub MyUnhandledExceptionEventHandler(sender As Object, e As UnhandledExceptionEventArgs)
        Dim MsgText As String = "??"
        If TypeOf (e.ExceptionObject) Is System.Exception Then
            Dim myException As System.Exception = e.ExceptionObject
            MsgText = myException.Message
            If myException.InnerException IsNot Nothing Then MsgText += "; " + myException.InnerException.Message
            If Not String.IsNullOrEmpty(myException.StackTrace) Then
                If myException.StackTrace.Length > 500 Then
                    MsgText += vbNewLine + myException.StackTrace.Substring(0, 500) + "..."
                Else
                    MsgText += vbNewLine + myException.StackTrace
                End If
            End If
        ElseIf TypeOf (e.ExceptionObject) Is Runtime.CompilerServices.RuntimeWrappedException Then
            Dim myException As Runtime.CompilerServices.RuntimeWrappedException = e.ExceptionObject
            If myException.InnerException IsNot Nothing Then MsgText += "; " + myException.InnerException.Message
            If Not String.IsNullOrEmpty(myException.StackTrace) Then
                If myException.StackTrace.Length > 500 Then
                    MsgText += vbNewLine + myException.StackTrace.Substring(0, 500) + "..."
                Else
                    MsgText += vbNewLine + myException.StackTrace
                End If
            End If
        End If

        DialogFactory.MsgError(Me, "Absturz " + My.Application.Info.AssemblyName, "Die Anwendung hat einen unerwarteten Abbruch erlitten. Folgende Informationen könnten bei der Fehlersuche helfen:" +
                               vbNewLine + vbNewLine + MsgText)

        Me.Close()
    End Sub
    Private Sub BtnLoginXlsx_Click(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub BtnBookletXlsx_Click(sender As Object, e As RoutedEventArgs)
        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_BookletXlsx) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_BookletXlsx)
        Dim filepicker As New Microsoft.Win32.OpenFileDialog With {.FileName = My.Settings.lastfile_BookletXlsx, .Filter = "Excel-Dateien|*.xlsx",
            .InitialDirectory = defaultDir, .DefaultExt = "Xlsx", .Title = "BookletXlsx - Wähle Datei"}
        If filepicker.ShowDialog Then
            My.Settings.lastfile_BookletXlsx = filepicker.FileName
            My.Settings.Save()

            Dim ActionDlg As New BookletXlsxDialog() With {.Owner = Me, .Title = "BookletXlsx"}
            ActionDlg.ShowDialog()
        End If
    End Sub

    Private Sub BtnSysCheck_Click(sender As Object, e As RoutedEventArgs)
        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_SysCheckCsv) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_SysCheckCsv)
        Dim filepickerSource As New Microsoft.Win32.OpenFileDialog With {.FileName = My.Settings.lastfile_SysCheckCsv, .Filter = "CSV-Dateien|*.csv",
            .InitialDirectory = defaultDir, .DefaultExt = "csv", .Title = "SysCheck - Wähle Quell-Datei"}
        If filepickerSource.ShowDialog Then
            My.Settings.lastfile_SysCheckCsv = filepickerSource.FileName
            My.Settings.Save()
            Dim csvData = New transformCsv2Xlsx(My.Settings.lastfile_SysCheckCsv)
            defaultDir = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            If Not String.IsNullOrEmpty(My.Settings.lastfile_SysCheckXlsx) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_SysCheckXlsx)
            Dim filepicker As New Microsoft.Win32.SaveFileDialog With {.FileName = My.Settings.lastfile_SysCheckXlsx, .Filter = "Excel-Dateien|*.xlsx",
                                                           .DefaultExt = "xlsx", .Title = "SysCheck - Wähle Ziel-Datei"}
            If filepicker.ShowDialog Then
                My.Settings.lastfile_SysCheckXlsx = filepicker.FileName
                My.Settings.Save()

                csvData.ToXlsx(filepicker.FileName)
            End If
        End If
    End Sub

    Private Sub BtnResponses_Click(sender As Object, e As RoutedEventArgs)
        Dim folderpicker As New System.Windows.Forms.FolderBrowserDialog With {.Description = "Wählen des Quellverzeichnisses für die Csv-Dateien",
                                                        .ShowNewFolderButton = False, .SelectedPath = My.Settings.lastdir_OutputSource}
        If folderpicker.ShowDialog() AndAlso Not String.IsNullOrEmpty(folderpicker.SelectedPath) Then
            My.Settings.lastdir_OutputSource = folderpicker.SelectedPath
            My.Settings.Save()

            Dim myDlg As New OutputDialog With {.Owner = Me}
            myDlg.ShowDialog()
        End If
    End Sub
    Private Sub HyperlinkClick(sender As Object, e As RoutedEventArgs)
        Dim linkcontrol As System.Windows.Documents.Hyperlink = sender
        Dim NavUri As Uri = linkcontrol.NavigateUri
        Process.Start(New ProcessStartInfo(NavUri.AbsoluteUri))
        e.Handled = True
    End Sub

    Private Sub BtnBookletXmlNew_Click(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub BtnBookletXmlOpen_Click(sender As Object, e As RoutedEventArgs)

    End Sub

    Private Sub BtnLoginPoolXlsx_Click(sender As Object, e As RoutedEventArgs)
        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_OutputTargetXlsx) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_OutputTargetXlsx)
        Dim filepicker As New Microsoft.Win32.SaveFileDialog With {.FileName = My.Settings.lastfile_OutputTargetXlsx, .Filter = "Excel-Dateien|*.xlsx",
                                                        .InitialDirectory = defaultDir, .DefaultExt = "xlsx", .Title = "Logins Zieldatei wählen"}
        If filepicker.ShowDialog Then
            My.Settings.lastfile_OutputTargetXlsx = filepicker.FileName
            My.Settings.Save()

            Dim myDlg As New LoginsXlsxDialog With {.Owner = Me, .Title = "Logins/Codes erzeugen"}
            myDlg.ShowDialog()
        End If
    End Sub
End Class
