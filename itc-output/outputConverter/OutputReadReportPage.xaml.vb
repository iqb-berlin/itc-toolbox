Public Class OutputReadReportPage
    Private Sub Me_Loaded() Handles Me.Loaded
        Dim SearchDir As New IO.DirectoryInfo(My.Settings.lastdir_OutputSource)
        Dim LogFileCount As Integer = 0
        Dim ResponsesFileCount As Integer = 0
        For Each fi As IO.FileInfo In SearchDir.GetFiles("*.csv", IO.SearchOption.AllDirectories)
            Try
                Dim readFile As System.IO.TextReader = New IO.StreamReader(fi.FullName)
                Dim line As String = readFile.ReadLine()
                If line = OutputDialog.LogFileFirstLine Then
                    LogFileCount += 1
                ElseIf line = OutputDialog.ResponsesFileFirstLine Then
                    ResponsesFileCount += 1
                Else
                    Me.MBUC.AddMessage("w: Datei nicht erkannt: " + fi.Name)
                End If
            Catch ex As Exception
                Me.MBUC.AddMessage("e: Fehler beim Lesen der Datei " + fi.Name + "; noch geöffnet?")
            End Try
        Next
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

    Private Sub BtnOKWith_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnOKWith.Click
        Dim defaultDir As String = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If Not String.IsNullOrEmpty(My.Settings.lastfile_BookletSizeTxt) Then defaultDir = IO.Path.GetDirectoryName(My.Settings.lastfile_BookletSizeTxt)
        Dim filepicker As New Microsoft.Win32.OpenFileDialog With {.FileName = My.Settings.lastfile_BookletSizeTxt, .Filter = "Txt-Dateien|*.txt",
                                                                          .InitialDirectory = defaultDir, .DefaultExt = "txt", .Title = "BookletTxt - Wähle Datei"}
        If filepicker.ShowDialog AndAlso Not String.IsNullOrEmpty(filepicker.FileName) Then
            My.Settings.lastfile_BookletSizeTxt = filepicker.FileName
            My.Settings.Save()

            Me.MBUC.AddMessage("Lese Bookletdatei")
            Dim bookletline As String
            Dim readFile As System.IO.TextReader = New IO.StreamReader(My.Settings.lastfile_BookletSizeTxt)
            Try
                bookletline = readFile.ReadLine()
            Catch ex As Exception
                bookletline = ""
                Me.MBUC.AddMessage("e:Fehler beim Lesen der Bookletdatei: " + ex.Message)
            End Try
            If Not String.IsNullOrEmpty(bookletline) Then
                Dim myParent As OutputDialog = Me.Parent
                myParent.bookletSize.Clear()
                Do While bookletline IsNot Nothing
                    Dim lineSplits As String() = bookletline.Split({" "}, StringSplitOptions.RemoveEmptyEntries)
                    If lineSplits.Count = 2 Then
                        Dim tryInt As Integer = 0
                        If Long.TryParse(lineSplits(1), tryInt) AndAlso Not myParent.bookletSize.ContainsKey(lineSplits(0).ToUpper()) Then
                            myParent.bookletSize.Add(lineSplits(0).ToUpper(), tryInt)
                        End If
                    End If
                    bookletline = readFile.ReadLine()
                Loop
                Me.MBUC.AddMessage(myParent.bookletSize.Count.ToString + " Einträge für Booklet-Größe gelesen")

                inputTargetFileName()
            Else
                Me.MBUC.AddMessage("e:Bookletdatei ist leer")
            End If
        End If
    End Sub
End Class
