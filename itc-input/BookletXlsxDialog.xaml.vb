Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports iqb.lib.openxml

Public Class BookletXlsxDialog
    Private ErrorMessages As Dictionary(Of String, List(Of String))
    Private Warnings As Dictionary(Of String, List(Of String))

#Region "Vorspann"
    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Me_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded
        BtnClose.Visibility = Windows.Visibility.Collapsed
        BtnEditor.Visibility = Windows.Visibility.Collapsed

        ErrorMessages = New Dictionary(Of String, List(Of String))
        Warnings = New Dictionary(Of String, List(Of String))

        Process1_bw = New ComponentModel.BackgroundWorker With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
        Process1_bw.RunWorkerAsync()
    End Sub

    Private WithEvents Process1_bw As ComponentModel.BackgroundWorker = Nothing

    Private Sub BtnCancel_Click() Handles BtnCancel.Click
        If Process1_bw IsNot Nothing AndAlso Process1_bw.IsBusy Then
            Process1_bw.CancelAsync()
            BtnCancel.IsEnabled = False
        Else
            DialogResult = False
        End If
    End Sub

    Private Sub BtnClose_Click() Handles BtnClose.Click
        DialogResult = False
    End Sub

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ComponentModel.ProgressChangedEventArgs) Handles Process1_bw.ProgressChanged
        Me.APBUC.UpdateProgressState(e.ProgressPercentage)
        If Not String.IsNullOrEmpty(e.UserState) Then MBUC.AddMessage(e.UserState)
    End Sub

    Private Sub Process1_bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As ComponentModel.RunWorkerCompletedEventArgs) Handles Process1_bw.RunWorkerCompleted
        APBUC.Value = 0.0#

        If ErrorMessages.Count > 0 Then
            MBUC.AddMessage("h: Fehler")
            For Each eMsg As KeyValuePair(Of String, List(Of String)) In ErrorMessages
                MBUC.AddMessage("e: " + eMsg.Key)
                For Each m As String In From s As String In eMsg.Value Order By s
                    MBUC.AddMessage("e: " + vbTab + m)
                Next
            Next
        End If

        If Warnings.Count > 0 Then
            MBUC.AddMessage("h: Warnungen")
            For Each wMsg As KeyValuePair(Of String, List(Of String)) In Warnings
                MBUC.AddMessage("w: " + wMsg.Key)
                For Each m As String In From s As String In wMsg.Value Order By s
                    MBUC.AddMessage("w: " + vbTab + m)
                Next
            Next
        End If


        MBUC.AddMessage("beendet")
        BtnCancel.Visibility = Windows.Visibility.Collapsed

        If e.Cancelled Then MBUC.AddMessage("durch Nutzer abgebrochen.")

        BtnClose.Visibility = Windows.Visibility.Visible
        BtnEditor.Visibility = Windows.Visibility.Visible
    End Sub
#End Region

#Region "Process1_bw_DoWork"

    '######################################################################################
    '######################################################################################
    Private Sub Process1_bw_DoWork(ByVal sender As Object, ByVal e As ComponentModel.DoWorkEventArgs) Handles Process1_bw.DoWork
        Dim myworker As ComponentModel.BackgroundWorker = sender

        myworker.ReportProgress(20.0#, "Öffne Datei")
        Dim sourceFileName = My.Settings.lastfile_BookletXlsx

        Dim sourceFile As Byte()
        Try
            sourceFile = IO.File.ReadAllBytes(sourceFileName)
        Catch ex As Exception
            myworker.ReportProgress(20.0#, "e:Konnte Datei " + sourceFileName + " nicht lesen (noch geöffnet?)")
            sourceFile = Nothing
        End Try

        If sourceFile IsNot Nothing Then
            Using MemStream As New IO.MemoryStream()
                MemStream.Write(sourceFile, 0, sourceFile.Length)
                Using sourceXLS As SpreadsheetDocument = SpreadsheetDocument.Open(MemStream, False)
                    myworker.ReportProgress(20.0#, "Lese Datei " + sourceFileName)

                    Dim unitsIDRef As String = xlsxFactory.GetDefinedNameValue(sourceXLS, "unitsCol.ID")
                    Dim unitsSheetName As String = xlsxFactory.GetWorksheetNameFromRefStr(unitsIDRef)
                    Dim unitsColID As String = xlsxFactory.GetColumnFromRefStr(unitsIDRef)
                    Dim unitsFirstRow As Integer = xlsxFactory.GetRowFromRefStr(unitsIDRef) + 1
                    Dim unitsColTitle As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "unitsCol.Title"))

                    Dim testletsIDRef As String = xlsxFactory.GetDefinedNameValue(sourceXLS, "testletsCol.ID")
                    Dim testletsSheetName As String = xlsxFactory.GetWorksheetNameFromRefStr(testletsIDRef)
                    Dim testletsColID As String = xlsxFactory.GetColumnFromRefStr(testletsIDRef)
                    Dim testletsFirstRow As Integer = xlsxFactory.GetRowFromRefStr(testletsIDRef) + 1
                    Dim testletsColTitle As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "testletsCol.Title"))
                    Dim testletsColUnits As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "testletsCol.Units"))

                    '-----------------------------------------------------------------------------
                    Dim unitsSheet As WorksheetPart = xlsxFactory.GetWorksheetPart(sourceXLS, unitsSheetName)
                    Dim testletsSheet As WorksheetPart = xlsxFactory.GetWorksheetPart(sourceXLS, testletsSheetName)

                    Dim AllUnits As New Dictionary(Of String, unitdata)
                    Dim AllTestlets As New Dictionary(Of String, testletdata)
                    Dim XMLBookletFiles As New Dictionary(Of String, XDocument)

                    '-----------------------------------------------------------------------------
                    '-----------------------------------------------------------------------------
                    myworker.ReportProgress(20.0#, "h:Lese Units")

                    Dim Zeile As Integer = unitsFirstRow
                    Dim unitId As String = ""
                    Dim fatal_error As Boolean = False
                    Do
                        If myworker.CancellationPending() Then Exit Do
                        unitId = xlsxFactory.GetCellValue(sourceXLS, unitsSheetName, unitsColID + Zeile.ToString)
                        If Not String.IsNullOrEmpty(unitId) Then
                            Dim unitTitle As String = xlsxFactory.GetCellValue(sourceXLS, unitsSheetName, unitsColTitle + Zeile.ToString)
                            If String.IsNullOrEmpty(unitTitle) Then
                                myworker.ReportProgress(20.0#, "e:Unit-Titel fehlt (unit '" + unitId + "', Zeile " + Zeile.ToString + ")")
                                fatal_error = True
                            ElseIf AllUnits.ContainsKey(unitid) Then
                                myworker.ReportProgress(20.0#, "e:Unit-ID doppelt (unit '" + unitId + "', Zeile " + Zeile.ToString + ")")
                                fatal_error = True
                            Else
                                AllUnits.Add(unitId, New unitdata With {.id = unitId, .title = unitTitle})
                            End If
                        End If
                        Zeile += 1
                    Loop Until String.IsNullOrEmpty(unitId)
                    myworker.ReportProgress(20.0#, AllUnits.Count.ToString + " Units gefunden")

                    '-----------------------------------------------------------------------------
                    '-----------------------------------------------------------------------------
                    myworker.ReportProgress(20.0#, "h:Lese Testlets")

                    Zeile = testletsFirstRow
                    Dim testletId As String = ""
                    Do
                        If myworker.CancellationPending() Then Exit Do
                        testletId = xlsxFactory.GetCellValue(sourceXLS, testletsSheetName, testletsColID + Zeile.ToString)
                        If Not String.IsNullOrEmpty(testletId) Then
                            Dim testletTitle As String = xlsxFactory.GetCellValue(sourceXLS, testletsSheetName, testletsColTitle + Zeile.ToString)
                            If AllTestlets.ContainsKey(testletId) Then
                                myworker.ReportProgress(20.0#, "e:Testlet-ID doppelt (testlet '" + testletId + "', Zeile " + Zeile.ToString + ")")
                                fatal_error = True
                            Else
                                Dim testletUnits As String = xlsxFactory.GetCellValue(sourceXLS, testletsSheetName, testletsColUnits + Zeile.ToString)
                                If String.IsNullOrEmpty(testletUnits) Then
                                    myworker.ReportProgress(20.0#, "e:Testlet ohne Units (testlet '" + testletId + "', Zeile " + Zeile.ToString + ")")
                                    fatal_error = True
                                Else
                                    Dim myTestlet As New testletdata With {.id = testletId, .title = testletTitle}
                                    For Each u As String In testletUnits.Split(" ")
                                        If AllUnits.ContainsKey(u) Then
                                            myTestlet.units.Add(AllUnits.Item(u))
                                        Else
                                            myworker.ReportProgress(20.0#, "e:Unit für Testlet nicht gefunden: '" + u + "' (testlet '" + testletId + "', Zeile " + Zeile.ToString + ")")
                                            fatal_error = True
                                        End If
                                    Next
                                    If myTestlet.units.Count > 0 Then
                                        AllTestlets.Add(testletId, myTestlet)
                                    End If
                                End If
                            End If
                        End If
                        Zeile += 1
                    Loop Until String.IsNullOrEmpty(testletId)
                    myworker.ReportProgress(20.0#, AllTestlets.Count.ToString + " Testlets gefunden")

                    If Not fatal_error Then
                        Dim CodePromptPrefix As String = xlsxFactory.GetCellValueFromRefStr(sourceXLS, xlsxFactory.GetDefinedNameValue(sourceXLS, "startlock.prefix"))
                        Dim CodePromptSuffix As String = xlsxFactory.GetCellValueFromRefStr(sourceXLS, xlsxFactory.GetDefinedNameValue(sourceXLS, "startlock.suffix"))
                        'Dim customtextsKey As String = xlsxFactory.GetDefinedNameValue(sourceXLS, "customText.Key")
                        'Dim customtextsSheetName As String = xlsxFactory.GetWorksheetNameFromRefStr(customtextsKey)
                        'Dim customtextsColKey As String = xlsxFactory.GetColumnFromRefStr(customtextsKey)
                        'Dim customtextsFirstRow As Integer = xlsxFactory.GetRowFromRefStr(customtextsKey) + 1
                        'Dim customtextsColValue As String = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "customText.Value"))

                        Dim customtexts As New Dictionary(Of String, String)
                        'Zeile = customtextsFirstRow

                        'Dim key As String = ""
                        'Do
                        '    If myworker.CancellationPending() Then Exit Do
                        '    key = xlsxFactory.GetCellValue(sourceXLS, customtextsSheetName, customtextsColKey + Zeile.ToString)
                        '    If Not String.IsNullOrEmpty(key) Then
                        '        Dim v As String = xlsxFactory.GetCellValue(sourceXLS, customtextsSheetName, customtextsColValue + Zeile.ToString)
                        '        If Not String.IsNullOrEmpty(v) AndAlso Not customtexts.ContainsKey(key) Then
                        '            customtexts.Add(key, v)
                        '        End If
                        '    End If
                        '    Zeile += 1
                        'Loop Until String.IsNullOrEmpty(key)

                        '-----------------------------------------------------------------------------
                        '-----------------------------------------------------------------------------

                        Dim targetPath As String = IO.Path.GetDirectoryName(sourceFileName) + IO.Path.DirectorySeparatorChar
                        Dim bookletCount As Integer = 0

                        '-----------------------------------------------------------------------------
                        '-----------------------------------------------------------------------------
                        myworker.ReportProgress(20.0#, "h:Schreibe Booklets XML")
                        Dim bookletsIDRef As String = xlsxFactory.GetDefinedNameValue(sourceXLS, "bookletsCol.ID")
                        Dim bookletsSheetName As String = xlsxFactory.GetWorksheetNameFromRefStr(bookletsIDRef)
                        Dim bookletsColID As String = ""
                        Dim bookletsFirstRow As Integer = 0
                        Dim bookletsColTitle As String = ""
                        Dim bookletsColDescription As String = ""
                        Dim bookletsColFirstElement As String = ""
                        Dim bookletsSheet As WorksheetPart = Nothing
                        Dim bookletHeaderRow As Integer = 0
                        Dim bookletId As String = ""

                        If String.IsNullOrEmpty(bookletsSheetName) Then
                            myworker.ReportProgress(20.0#, "e:bookletsCol.ID nicht gefunden - ignoriere")
                        Else
                            bookletsColID = xlsxFactory.GetColumnFromRefStr(bookletsIDRef)
                            bookletsFirstRow = xlsxFactory.GetRowFromRefStr(bookletsIDRef) + 1
                            bookletsColTitle = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "bookletsCol.Title"))
                            bookletsColDescription = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "bookletsCol.Description"))
                            bookletsColFirstElement = xlsxFactory.GetColumnFromRefStr(xlsxFactory.GetDefinedNameValue(sourceXLS, "bookletsCol.FirstElement"))
                            bookletsSheet = xlsxFactory.GetWorksheetPart(sourceXLS, bookletsSheetName)

                            Zeile = bookletsFirstRow
                            bookletHeaderRow = Zeile - 1

                            Do
                                If myworker.CancellationPending() Then Exit Do
                                bookletId = xlsxFactory.GetCellValue(sourceXLS, bookletsSheetName, bookletsColID + Zeile.ToString)
                                If Not String.IsNullOrEmpty(bookletId) Then
                                    If bookletId.IndexOf(" ") >= 0 Then
                                    myworker.ReportProgress(20.0#, "e:Booklet-Id enthält Leerzeichen (Booklet '" + bookletId + "', Zeile " + Zeile.ToString + ")")
                                Else
                                    Dim bookletTitle As String = xlsxFactory.GetCellValue(sourceXLS, bookletsSheetName, bookletsColTitle + Zeile.ToString)
                                        If String.IsNullOrEmpty(bookletTitle) Then
                                            myworker.ReportProgress(20.0#, "e:Booklet-Title fehlt (Booklet '" + bookletId + "', Zeile " + Zeile.ToString + ")")
                                        Else
                                            Dim bookletDescription As String = xlsxFactory.GetCellValue(sourceXLS, bookletsSheetName, bookletsColTitle + Zeile.ToString)

                                            Dim BookletXmlFile As XDocument = <?xml version="1.0" encoding="utf-8"?>
                                                                              <Booklet xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                                                                                  xsi:noNamespaceSchemaLocation="https://raw.githubusercontent.com/iqb-berlin/testcenter-backend/9.1.1/definitions/vo_Booklet.xsd">

                                                                              </Booklet>
                                            fatal_error = False
                                            BookletXmlFile.Root.Add(
                                                    <Metadata>
                                                        <Id><%= bookletId %></Id>
                                                        <Label><%= bookletTitle %></Label>
                                                        <Description><%= bookletDescription %></Description>
                                                    </Metadata>
                                            )
                                            BookletXmlFile.Root.Add(
                                                    <BookletConfig>
                                                        <Config key="force_presentation_complete">ON</Config>
                                                        <Config key="unit_navibuttons">FULL</Config>
                                                    </BookletConfig>
                                            )
                                            'If customtexts.Count > 0 Then
                                            '    Dim xct As XElement = <CustomTexts></CustomTexts>
                                            '    For Each ct As KeyValuePair(Of String, String) In customtexts
                                            '        xct.Add(<Text key=<%= ct.Key %>><%= ct.Value %></Text>)
                                            '    Next
                                            '    BookletXmlFile.Root.Add(xct)
                                            'End If

                                            Dim CurrentCol As String = bookletsColFirstElement
                                            Dim HeaderContent As String = ""
                                            Dim CodeCount As Integer = 0
                                            Dim XUnits As XElement = <Units/>
                                            '-----------------------------------------------------------------------------
                                            Do
                                                HeaderContent = xlsxFactory.GetCellValue(sourceXLS, bookletsSheetName, CurrentCol + bookletHeaderRow.ToString)
                                                If Not String.IsNullOrEmpty(HeaderContent) Then
                                                    If HeaderContent = "Unit" Then
                                                        Dim CellContent As String = xlsxFactory.GetCellValue(sourceXLS, bookletsSheetName, CurrentCol + Zeile.ToString)
                                                        If Not String.IsNullOrEmpty(CellContent) Then
                                                            If Not AllUnits.ContainsKey(CellContent) Then
                                                                myworker.ReportProgress(20.0#, "e:Booklet-Unit unbekannt (Booklet '" + bookletId + "', Zeile " + Zeile.ToString + ", Spalte " + CurrentCol + ")")
                                                                fatal_error = True
                                                            Else
                                                                XUnits.Add(AllUnits.Item(CellContent).toXml("", ""))
                                                            End If
                                                        End If
                                                    ElseIf HeaderContent = "Block-ID" Then
                                                        Dim CellContent As String = xlsxFactory.GetCellValue(sourceXLS, bookletsSheetName, CurrentCol + Zeile.ToString)
                                                        If Not String.IsNullOrEmpty(CellContent) Then
                                                            If Not AllTestlets.ContainsKey(CellContent) Then
                                                                myworker.ReportProgress(20.0#, "e:Booklet-Testlet unbekannt (Booklet '" + bookletId + "', Zeile " + Zeile.ToString + ", Spalte " + CurrentCol + ")")
                                                                fatal_error = True
                                                            Else
                                                                Dim Code As String = ""
                                                                Dim CodePrompt As String = ""
                                                                Dim MaxTime As String = ""
                                                                Dim TempCol As String = xlsxFactory.GetNextColumn(CurrentCol)
                                                                Dim TempHeaderContent As String = xlsxFactory.GetCellValue(sourceXLS, bookletsSheetName, TempCol + bookletHeaderRow.ToString)
                                                                If TempHeaderContent = "Code" Then
                                                                    Code = xlsxFactory.GetCellValue(sourceXLS, bookletsSheetName, TempCol + Zeile.ToString)
                                                                    CodeCount += 1
                                                                    CodePrompt = CodePromptPrefix + CodeCount.ToString + CodePromptSuffix
                                                                    CurrentCol = TempCol
                                                                ElseIf TempHeaderContent = "max Time" Then
                                                                    MaxTime = xlsxFactory.GetCellValue(sourceXLS, bookletsSheetName, TempCol + Zeile.ToString)
                                                                    CurrentCol = TempCol
                                                                End If
                                                                TempCol = xlsxFactory.GetNextColumn(TempCol)
                                                                TempHeaderContent = xlsxFactory.GetCellValue(sourceXLS, bookletsSheetName, TempCol + bookletHeaderRow.ToString)
                                                                If TempHeaderContent = "Code" Then
                                                                    Code = xlsxFactory.GetCellValue(sourceXLS, bookletsSheetName, TempCol + Zeile.ToString)
                                                                    CodeCount += 1
                                                                    CodePrompt = CodePromptPrefix + CodeCount.ToString + CodePromptSuffix
                                                                    CurrentCol = TempCol
                                                                ElseIf TempHeaderContent = "max Time" Then
                                                                    MaxTime = xlsxFactory.GetCellValue(sourceXLS, bookletsSheetName, TempCol + Zeile.ToString)
                                                                    CurrentCol = TempCol
                                                                End If
                                                                XUnits.Add(AllTestlets.Item(CellContent).toXml(Code, CodePrompt, MaxTime))
                                                            End If
                                                        End If
                                                    ElseIf HeaderContent = "max Time" OrElse HeaderContent = "Code" Then
                                                        'ignore
                                                    Else
                                                        myworker.ReportProgress(20.0#, "w:Booklet-Spaltenkopf unbekannt '" + HeaderContent + "'")
                                                    End If
                                                    CurrentCol = xlsxFactory.GetNextColumn(CurrentCol)
                                                End If
                                            Loop Until String.IsNullOrEmpty(HeaderContent)
                                            '-----------------------------------------------------------------------------

                                            If Not fatal_error Then
                                                BookletXmlFile.Root.Add(XUnits)
                                                Try
                                                    BookletXmlFile.Save(targetPath + "booklet" + bookletId + ".xml")
                                                    myworker.ReportProgress(20.0#, "booklet" + bookletId + ".xml")
                                                    bookletCount += 1
                                                Catch ex As Exception
                                                    Dim msg As String = ex.Message
                                                    If ex.InnerException IsNot Nothing Then msg += vbNewLine + ex.InnerException.Message
                                                    myworker.ReportProgress(20.0#, "e: Konnte booklet '" + bookletId + "' nicht schreiben (" + msg + ")")
                                                End Try
                                            End If
                                        End If
                                    End If
                                End If
                                Zeile += 1
                            Loop Until String.IsNullOrEmpty(bookletId)
                        End If

                        myworker.ReportProgress(20.0#, bookletCount.ToString + " Booklets gespeichert")
                    End If
                End Using
            End Using
        End If
    End Sub


    Private Sub AddWarningMessage(MessageStr As String, MessageParameter As String)
        If Not Warnings.ContainsKey(MessageStr) Then Warnings.Add(MessageStr, New List(Of String))
        Warnings.Item(MessageStr).Add(MessageParameter)
    End Sub
    Private Sub AddErrorMessage(MessageStr As String, MessageParameter As String)
        If Not ErrorMessages.ContainsKey(MessageStr) Then ErrorMessages.Add(MessageStr, New List(Of String))
        ErrorMessages.Item(MessageStr).Add(MessageParameter)
    End Sub

    Private Sub BtnEditor_Click() Handles BtnEditor.Click
        Try
            Dim txtFN As String = IO.Path.GetTempPath + IO.Path.DirectorySeparatorChar + "TestCenter" + Guid.NewGuid.ToString + ".txt"
            IO.File.WriteAllBytes(txtFN, System.Text.Encoding.Unicode.GetBytes(MBUC.Text))

            Dim proc As New Process
            With proc.StartInfo
                .FileName = txtFN
                .WindowStyle = ProcessWindowStyle.Normal
            End With
            proc.Start()

            Me.DialogResult = True
        Catch ex As Exception
            Dim msg As String = ex.Message
            If ex.InnerException IsNot Nothing Then msg += vbNewLine + ex.InnerException.Message
            iqb.lib.components.DialogFactory.MsgError(iqb.lib.components.DialogFactory.GetParentWindow(Me), "Übertragen Meldungen in Texteditor", msg)
        End Try
    End Sub
#End Region


End Class

'#############################################################################
Public Class unitdata
    Public id As String = ""
    Public title As String = ""
    Public Function toXml(titlePrefix As String, labelShort As String) As XElement
        If String.IsNullOrEmpty(labelShort) Then
            Return <Unit id=<%= id %> label=<%= titlePrefix + title %>/>
        Else
            Return <Unit id=<%= id %> label=<%= titlePrefix + title %> labelshort=<%= labelShort %>/>
        End If
    End Function
End Class

'#############################################################################
Public Class testletdata
    Public id As String
    Public title As String = ""
    Public units As New List(Of unitdata)

    Public Function toXml(codeToEnter As String, codeToEnterPrompt As String, timeMax As String) As XElement
        Dim myreturn As XElement = <Testlet id=<%= id %> label=<%= title %>/>
        If Not String.IsNullOrEmpty(codeToEnter) OrElse Not String.IsNullOrEmpty(timeMax) Then
            Dim XRestrictions As XElement = <Restrictions/>
            If Not String.IsNullOrEmpty(codeToEnter) Then XRestrictions.Add(<CodeToEnter code=<%= codeToEnter %>><%= codeToEnterPrompt %></CodeToEnter>)
            If Not String.IsNullOrEmpty(timeMax) Then XRestrictions.Add(<TimeMax minutes=<%= timeMax %>/>)
            myreturn.Add(XRestrictions)
        End If

        Dim unitcounter As Integer = 1
        Dim doUnitNumbering As Boolean = units.Count > 1
        For Each u As unitdata In units
            If doUnitNumbering Then
                If u.id.EndsWith("End") Then
                    myreturn.Add(u.toXml("", ""))
                Else
                    myreturn.Add(u.toXml(unitcounter.ToString + ". ", unitcounter.ToString))
                End If
                unitcounter += 1
            Else
                myreturn.Add(u.toXml("", ""))
            End If
        Next

        Return myreturn
    End Function

End Class