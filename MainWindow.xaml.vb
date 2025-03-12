Imports iqb.lib.components
Class MainWindow
    Private SqliteDB As SQLiteConnector = Nothing

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

        If ContinueStart Then
            CommandBindings.Add(New CommandBinding(AppCommands.AppExit, AddressOf HandleAppExitExecuted))
            CommandBindings.Add(New CommandBinding(AppCommands.ImportFromTestcenter, AddressOf HandleImportFromTestcenterExecuted))
            CommandBindings.Add(New CommandBinding(AppCommands.ImportFromCsv, AddressOf HandleImportFromCsvExecuted))
            CommandBindings.Add(New CommandBinding(AppCommands.DBNew, AddressOf HandleDBNewExecuted))
            CommandBindings.Add(New CommandBinding(AppCommands.DBOpen, AddressOf HandleDBOpenExecuted))
            CommandBindings.Add(New CommandBinding(AppCommands.DBCopyTo, AddressOf HandleDBCopyToExecuted, AddressOf HandleDBCopyToCanExecute))
        Else
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
    End Sub

    Private Sub HandleHelpExecuted(sender As Object, e As ExecutedRoutedEventArgs)
        Dim myDlg As New iqb.lib.components.AppAboutDialog With {.Owner = Me}
        myDlg.ShowDialog()
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

            Dim myDlg As New CodesXlsxDialog With {.Owner = Me, .Title = "Logins/Codes erzeugen"}
            myDlg.ShowDialog()
        End If
    End Sub

    Private Sub BtnLoginXlsxTemplate_Click(sender As Object, e As RoutedEventArgs)
        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_LoginXlsx) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_LoginXlsx)
        Dim filepicker As New Microsoft.Win32.OpenFileDialog With {.FileName = IO.Path.GetFileName(My.Settings.lastfile_LoginXlsx), .Filter = "Excel-Dateien|*.xlsx",
            .InitialDirectory = defaultDir, .DefaultExt = "Xlsx", .Title = "Logins Xlsx/Xml erzeugen - Wähle Datei"}
        If filepicker.ShowDialog Then
            My.Settings.lastfile_LoginXlsx = filepicker.FileName
            My.Settings.Save()

            Dim ActionDlg As New LoginsTemplateXlsxDialog() With {.Owner = Me, .Title = "Logins Xlsx/Xml erzeugen"}
            ActionDlg.ShowDialog()
        End If
    End Sub

    Private Sub BtnLoginXlsxToDocx_Click(sender As Object, e As RoutedEventArgs)
        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_LoginXlsx) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_LoginXlsx)
        Dim filepicker As New Microsoft.Win32.OpenFileDialog With {.FileName = IO.Path.GetFileName(My.Settings.lastfile_LoginXlsx), .Filter = "Excel-Dateien|*.xlsx",
            .InitialDirectory = defaultDir, .DefaultExt = "Xlsx", .Title = "Logins Docx erzeugen - Wähle Datei"}
        If filepicker.ShowDialog Then
            My.Settings.lastfile_LoginXlsx = filepicker.FileName
            My.Settings.Save()

            Dim ActionDlg As New LoginsXlsxToDocxDialog() With {.Owner = Me, .Title = "Logins Docx erzeugen"}
            ActionDlg.ShowDialog()
        End If
    End Sub

    Private Sub BtnGetTestcenterReviewsData_Click(sender As Object, e As RoutedEventArgs)
        Dim ActionDlg As New LoadDataFromTestcenterDialog(TestcenterReadMode.Reviews) With {.Owner = Me, .Title = "Reviews aus Testcenter laden und speichern"}
        ActionDlg.ShowDialog()
    End Sub

    Private Sub BtnTestcenterToJson_Click(sender As Object, e As RoutedEventArgs)
        Dim ActionDlg As New LoadDataFromTestcenterDialog(TestcenterReadMode.Responses, True) With {
            .Owner = Me, .Title = "Antworten und Logs aus Testcenter laden und speichern"}
        ActionDlg.ShowDialog()
        updateGroupCount()
    End Sub

    Private Sub BtnMergeDataLoadTC_Click(sender As Object, e As RoutedEventArgs)
        Dim ActionDlg As New LoadDataFromTestcenterDialog(TestcenterReadMode.Responses, False) With {.Owner = Me, .Title = "Antworten und Logs aus Testcenter laden"}
        If ActionDlg.ShowDialog() Then updateGroupCount()
    End Sub

    Private Sub BtnMergeDataLoadCsv_Click(sender As Object, e As RoutedEventArgs)
        Dim folderpicker As New System.Windows.Forms.FolderBrowserDialog With {.Description = "Wählen des Quellverzeichnisses für die Csv-Dateien",
                                                        .ShowNewFolderButton = False, .SelectedPath = My.Settings.lastdir_OutputSource}
        If folderpicker.ShowDialog() AndAlso Not String.IsNullOrEmpty(folderpicker.SelectedPath) Then
            My.Settings.lastdir_OutputSource = folderpicker.SelectedPath
            My.Settings.Save()

            Dim ActionDlg As New OutputDialog(False) With {.Owner = Me}
            If ActionDlg.ShowDialog() Then
                updateGroupCount()
            End If
        End If
    End Sub

    Private Sub BtnDataLoadJson_Click(sender As Object, e As RoutedEventArgs)
        Dim ActionDlg As New readJsonFilesDialog() With {.Owner = Me, .Title = "Einlesen TC-JSON"}
        ActionDlg.ShowDialog()
        updateGroupCount()
    End Sub
    Private Sub BtnMergeDataClear_Click(sender As Object, e As RoutedEventArgs)
        globalOutputStore.clear()
        updateGroupCount()
    End Sub

    Private Sub BtnMergeDataSaveJson_Click(sender As Object, e As RoutedEventArgs)
        If globalOutputStore.personDataFull.Count > 0 Then
            Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            If Not String.IsNullOrEmpty(My.Settings.lastfile_OutputTargetJson) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_OutputTargetJson)
            Dim filepicker As New Microsoft.Win32.SaveFileDialog With {.FileName = My.Settings.lastfile_OutputTargetJson, .Filter = "JSON-Dateien|*.json",
                                                            .InitialDirectory = defaultDir, .DefaultExt = "json", .Title = "JSON Zieldatei wählen"}
            If filepicker.ShowDialog Then
                My.Settings.lastfile_OutputTargetJson = filepicker.FileName
                My.Settings.Save()

                JsonReadWrite.Write(filepicker.FileName)
                DialogFactory.Msg(Me, "DataMerge", "fertig")
            End If
        Else
            DialogFactory.MsgError(Me, "DataMerge", "JSON-Output kann nur aus dem Volldaten-Store erzeugt werden (derzeit keine Daten).")
        End If
    End Sub

    Private Sub BtnMergeDataSaveJsonByGroup_Click(sender As Object, e As RoutedEventArgs)
        If globalOutputStore.personDataFull.Count > 0 Then
            Dim folderpicker As New Forms.FolderBrowserDialog With {.Description = "Zielverzeichnis für die JSON-Dateien",
                                                            .ShowNewFolderButton = True, .SelectedPath = My.Settings.lastfolder_OutputTarget}
            If folderpicker.ShowDialog() AndAlso Not String.IsNullOrEmpty(folderpicker.SelectedPath) Then
                My.Settings.lastfolder_OutputTarget = folderpicker.SelectedPath
                My.Settings.Save()

                JsonReadWrite.WriteByGroup(folderpicker.SelectedPath)
                JsonReadWrite.WriteBigData(folderpicker.SelectedPath)
                DialogFactory.Msg(Me, "DataMerge", "fertig")
            End If
        Else
            DialogFactory.MsgError(Me, "DataMerge", "JSON-Output kann nur aus dem Volldaten-Store erzeugt werden (derzeit keine Daten).")
        End If
    End Sub

    Private Sub updateGroupCount()
        TBStoreCountFull.Text = globalOutputStore.personDataFull.Count.ToString

        TBStoreCountBlobs.Text = globalOutputStore.bigData.Count.ToString
        TBStoreCountResponses.Text = globalOutputStore.personResponses.Count.ToString
        'TBStoreCountLogs.Text = globalOutputStore.personLogs.Count.ToString
        TBStoreCountBooklets.Text = globalOutputStore.bookletSizes.Count.ToString
    End Sub

    Private Sub BtnMergeDataSaveXlsx_Click(sender As Object, e As RoutedEventArgs)
        If globalOutputStore.personDataFull.Count + globalOutputStore.personResponses.Count > 0 Then
            Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            If Not String.IsNullOrEmpty(My.Settings.lastfile_OutputTargetXlsx) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_OutputTargetXlsx)
            Dim filepicker As New Microsoft.Win32.SaveFileDialog With {.FileName = My.Settings.lastfile_OutputTargetXlsx, .Filter = "Excel-Dateien|*.xlsx",
                                                            .InitialDirectory = defaultDir, .DefaultExt = "xlsx", .Title = "Xlsx Zieldatei wählen"}
            If filepicker.ShowDialog Then
                My.Settings.lastfile_OutputTargetXlsx = filepicker.FileName
                My.Settings.Save()


                Dim ActionDlg As New ToXlsxDialog() With {.Owner = Me, .Title = "Schreiben Xslx-Output"}
                ActionDlg.ShowDialog()
            End If
        Else
            DialogFactory.MsgError(Me, "DataMerge", "JSON-Output kann nur aus dem Volldaten-Store oder dem Antwort-Store erzeugt werden (derzeit keine Daten).")
        End If
    End Sub

    Private Sub BtnWriteSqlite_Click(sender As Object, e As RoutedEventArgs)
        If globalOutputStore.personDataFull.Count + globalOutputStore.personResponses.Count > 0 Then
            Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            If Not String.IsNullOrEmpty(My.Settings.lastfile_OutputTargetSqlite) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_OutputTargetSqlite)
            Dim filepicker As New Microsoft.Win32.SaveFileDialog With {
                .FileName = My.Settings.lastfile_OutputTargetSqlite, .Filter = "SQLite-Dateien|*.sqlite",
                .InitialDirectory = defaultDir, .DefaultExt = "sqlite", .Title = "SQLite Zieldatei wählen"}
            If filepicker.ShowDialog Then
                My.Settings.lastfile_OutputTargetSqlite = filepicker.FileName
                My.Settings.Save()

                Dim ActionDlg As New ToSqliteDialog() With {.Owner = Me, .Title = "Schreiben SQLite-Output"}
                ActionDlg.ShowDialog()
            End If
        Else
            DialogFactory.MsgError(Me, "SQLite", "SQLite-Output kann nur aus dem Volldaten-Store oder dem Antwort-Store erzeugt werden (derzeit keine Daten).")
        End If
    End Sub

    Private Sub HandleAppExitExecuted(ByVal sender As Object, ByVal e As ExecutedRoutedEventArgs)
        Me.Close()
    End Sub

    ' ################################################################################################
    Private Sub UpdateSqliteDBInfo()
        If Me.SqliteDB Is Nothing Then
            TBDBInfo.Text = "Keine Daten"
        Else
            TBDBInfo.Text = Me.SqliteDB.dbCreator + ": " + Me.SqliteDB.dbCreatedDateTime
        End If
    End Sub

    Private Sub HandleImportFromTestcenterExecuted(ByVal sender As Object, ByVal e As ExecutedRoutedEventArgs)
        DialogFactory.Msg(Me, "yoyo", "HandleImportFromTestcenterExecuted")
    End Sub

    Private Sub HandleImportFromCsvExecuted(ByVal sender As Object, ByVal e As ExecutedRoutedEventArgs)
        DialogFactory.Msg(Me, "yoyo", "HandleImportFromCsvExecuted")
    End Sub

    Private Sub HandleDBNewExecuted(ByVal sender As Object, ByVal e As ExecutedRoutedEventArgs)
        DBNewOrOpen(True)
    End Sub

    Private Sub HandleDBOpenExecuted(ByVal sender As Object, ByVal e As ExecutedRoutedEventArgs)
        DBNewOrOpen(False)
    End Sub

    Private Sub DBNewOrOpen(create As Boolean)
        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_SqliteDB) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_SqliteDB)
        Dim filepicker As New Microsoft.Win32.OpenFileDialog With {.FileName = IO.Path.GetFileName(My.Settings.lastfile_SqliteDB), .Filter = "Sqlite-Dateien|*.sqlite",
            .InitialDirectory = defaultDir, .DefaultExt = "sqlite", .Title = "Datenbank-Datei wählen"}
        If filepicker.ShowDialog Then
            My.Settings.lastfile_SqliteDB = filepicker.FileName
            My.Settings.Save()

            Me.SqliteDB = New SQLiteConnector(My.Settings.lastfile_SqliteDB)
            UpdateSqliteDBInfo()
        End If
    End Sub
    Private Sub HandleDBCopyToExecuted(ByVal sender As Object, ByVal e As ExecutedRoutedEventArgs)
        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_SqliteDB) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_SqliteDB)
        Dim filepicker As New Microsoft.Win32.SaveFileDialog With {.FileName = My.Settings.lastfile_SqliteDB, .Filter = "Sqlite-Dateien|*.sqlite",
            .CheckFileExists = True, .InitialDirectory = defaultDir, .DefaultExt = "sqlite", .Title = "Datenbank-Datei wählen"}
        If filepicker.ShowDialog Then
            My.Settings.lastfile_SqliteDB = filepicker.FileName
            My.Settings.Save()

            DialogFactory.Msg(Me, "yoyo", "HandleDBCopyToExecuted")
            UpdateSqliteDBInfo()
        End If
    End Sub

    Private Function HandleDBCopyToCanExecute(ByVal sender As Object, ByVal e As System.Windows.Input.CanExecuteRoutedEventArgs) As Boolean
        Dim myreturn As Boolean = Me.SqliteDB IsNot Nothing
        e.CanExecute = myreturn
        Return myreturn
    End Function
End Class
