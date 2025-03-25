Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports iqb.lib.openxml
Imports System.ComponentModel

Public Enum SubformMode
    None
    Columns
    Rows
End Enum

Public Class WriteXlsxConfig
    Public targetXlsxFilename As String
    Public writeResponsesValues As Boolean
    Public writeResponsesStatus As Boolean
    Public writeResponsesCodes As Boolean
    Public writeResponsesScores As Boolean
    Public writeSessions As Boolean
    Public subformMode As SubformMode
    Public sourceDatabase As SQLiteConnector = Nothing
End Class

Class WriteOutputToXlsx
    Public Shared Sub Write(
                           workbookTemplate As Byte(),
                           worker As BackgroundWorker,
                           e As DoWorkEventArgs,
                           config As WriteXlsxConfig
                           )
        worker.ReportProgress(0.0#, "Ermittle Variablen")
        Dim AllVariables As New List(Of String)
        If config.sourceDatabase Is Nothing Then
            AllVariables = WriteOutputToXlsx.getVariableListFromStore(config.subformMode = SubformMode.Columns)
        Else
            AllVariables = config.sourceDatabase.getVariableList(config.subformMode = SubformMode.Columns)
        End If
        worker.ReportProgress(0.0#, AllVariables.Count.ToString + " Spalten.")

        If AllVariables.Count > 0 Then
            Dim peopleList As Dictionary(Of String, String)
            If config.sourceDatabase Is Nothing Then
                peopleList = (From kvp As KeyValuePair(Of String, Person)
                                  In globalOutputStore.personDataFull).ToDictionary(Function(a) a.Key, Function(a) a.Key)
            Else
                peopleList = config.sourceDatabase.getPeopleList()
            End If

            Using MemStream As New IO.MemoryStream()
                MemStream.Write(workbookTemplate, 0, workbookTemplate.Length)
                Using ZielXLS As SpreadsheetDocument = SpreadsheetDocument.Open(MemStream, True)
                    Dim myStyles As ExcelStyleDefs = xlsxFactory.AddIQBStandardStyles(ZielXLS.WorkbookPart)
                    '########################################################
                    'Responses
                    '########################################################
                    Dim TableValues As WorksheetPart = Nothing
                    If config.writeResponsesValues Then TableValues = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "Antworten")
                    Dim TableStatus As WorksheetPart = Nothing
                    If config.writeResponsesStatus Then TableStatus = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "Status")
                    Dim TableCodes As WorksheetPart = Nothing
                    If config.writeResponsesCodes Then TableCodes = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "Codes")
                    Dim TableScores As WorksheetPart = Nothing
                    If config.writeResponsesScores Then TableScores = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "Scores")

                    worker.ReportProgress(0.0#, "Schreibe Daten")

                    Dim myRow As Integer = 1
                    If TableValues IsNot Nothing Then
                        xlsxFactory.SetCellValueString("A", myRow, TableValues, "ID", CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("A", TableValues, 20)
                        xlsxFactory.SetCellValueString("B", myRow, TableValues, "Group", CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("B", TableValues, 10)
                        xlsxFactory.SetCellValueString("C", myRow, TableValues, "Login+Code", CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("C", TableValues, 10)
                    End If
                    If TableStatus IsNot Nothing Then
                        xlsxFactory.SetCellValueString("A", myRow, TableStatus, "ID", CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("A", TableStatus, 20)
                        xlsxFactory.SetCellValueString("B", myRow, TableStatus, "Group", CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("B", TableStatus, 10)
                        xlsxFactory.SetCellValueString("C", myRow, TableStatus, "Login+Code", CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("C", TableStatus, 10)
                    End If
                    If TableScores IsNot Nothing Then
                        xlsxFactory.SetCellValueString("A", myRow, TableScores, "ID", CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("A", TableScores, 20)
                        xlsxFactory.SetCellValueString("B", myRow, TableScores, "Group", CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("B", TableScores, 10)
                        xlsxFactory.SetCellValueString("C", myRow, TableScores, "Login+Code", CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("C", TableScores, 10)
                    End If
                    If TableCodes IsNot Nothing Then
                        xlsxFactory.SetCellValueString("A", myRow, TableCodes, "ID", CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("A", TableCodes, 20)
                        xlsxFactory.SetCellValueString("B", myRow, TableCodes, "Group", CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("B", TableCodes, 10)
                        xlsxFactory.SetCellValueString("C", myRow, TableCodes, "Login+Code", CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth("C", TableCodes, 10)
                    End If

                    Dim myColumn As String = "D"
                    Dim Columns As New Dictionary(Of String, String)

                    Dim progressMax As Integer = AllVariables.Count
                    Dim progressCount As Integer = 1
                    Dim stepMax As Integer = 5
                    Dim stepCount As Integer = 1
                    Dim progressValue As Double = 0.0#

                    For Each s As String In From v As String In AllVariables Order By v Select v
                        progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                        worker.ReportProgress(progressValue, "")
                        progressCount += 1
                        If TableValues IsNot Nothing Then
                            xlsxFactory.SetCellValueString(myColumn, myRow, TableValues, s, CellFormatting.RowHeader2, myStyles)
                            xlsxFactory.SetColumnWidth(myColumn, TableValues, 10)
                        End If
                        If TableStatus IsNot Nothing Then
                            xlsxFactory.SetCellValueString(myColumn, myRow, TableStatus, s, CellFormatting.RowHeader2, myStyles)
                            xlsxFactory.SetColumnWidth(myColumn, TableStatus, 10)
                        End If
                        If TableScores IsNot Nothing Then
                            xlsxFactory.SetCellValueString(myColumn, myRow, TableScores, s, CellFormatting.RowHeader2, myStyles)
                            xlsxFactory.SetColumnWidth(myColumn, TableScores, 10)
                        End If
                        If TableCodes IsNot Nothing Then
                            xlsxFactory.SetCellValueString(myColumn, myRow, TableCodes, s, CellFormatting.RowHeader2, myStyles)
                            xlsxFactory.SetColumnWidth(myColumn, TableCodes, 10)
                        End If
                        Columns.Add(s, myColumn)
                        myColumn = xlsxFactory.GetNextColumn(myColumn)
                    Next

                    Dim BookletUnits As New Dictionary(Of String, List(Of String)) 'für unten TechTable

                    progressMax = peopleList.Count
                    progressCount = 1
                    stepCount += 1
                    For Each personKey As KeyValuePair(Of String, String) In
                        From p As KeyValuePair(Of String, String) In peopleList Order By p.Key
                        If worker.CancellationPending Then
                            e.Cancel = True
                            Exit For
                        End If
                        progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                        worker.ReportProgress(progressValue, personKey.Key)
                        progressCount += 1
                        Dim personData As Person
                        If config.sourceDatabase Is Nothing Then
                            personData = globalOutputStore.personDataFull.Item(personKey.Key)
                        Else
                            personData = config.sourceDatabase.getPersonResponses(personKey.Value)
                        End If

                        Dim subforms As List(Of String)
                        If config.subformMode = SubformMode.Rows Then
                            subforms = (From b As Booklet In personData.booklets
                                        From u As Unit In b.units
                                        From sf As SubForm In u.subforms Select sf.id).Distinct.ToList()
                        Else
                            subforms = New List(Of String)
                            subforms.Add("")
                        End If

                        For Each subformKey As String In From sf As String In subforms Order By sf
                            Dim myRowDataValues As New List(Of RowData)
                            Dim myRowDataStatus As New List(Of RowData)
                            Dim myRowDataScores As New List(Of RowData)
                            Dim myRowDataCodes As New List(Of RowData)
                            Dim hasData As Boolean = False
                            myRowDataValues.Add(New RowData With {.Column = "A", .Value = personKey.Key, .CellType = CellTypes.str})
                            myRowDataValues.Add(New RowData With {.Column = "B", .Value = personData.group, .CellType = CellTypes.str})
                            myRowDataValues.Add(New RowData With {.Column = "C",
                                                   .Value = personData.login + personData.code + IIf(String.IsNullOrEmpty(subformKey), "", "_" + subformKey),
                                                   .CellType = CellTypes.str})
                            For Each rd As RowData In myRowDataValues
                                myRowDataStatus.Add(New RowData With {.Column = rd.Column, .Value = rd.Value, .CellType = rd.CellType})
                                myRowDataScores.Add(New RowData With {.Column = rd.Column, .Value = rd.Value, .CellType = rd.CellType})
                                myRowDataCodes.Add(New RowData With {.Column = rd.Column, .Value = rd.Value, .CellType = rd.CellType})
                            Next

                            If config.subformMode = SubformMode.Columns Then
                                For Each unit As Unit In
                                    From b As Booklet In personData.booklets
                                    From u As Unit In b.units Select u
                                    For Each subform As SubForm In unit.subforms
                                        For Each r As ResponseData In subform.responses
                                            hasData = True
                                            Dim columnKey As String = Columns.Item(WriteOutputToXlsx.getColumnKey(unit, r.id, subform.id))
                                            myRowDataValues.Add(
                                                New RowData With {.Column = columnKey, .Value = r.value, .CellType = CellTypes.str})
                                            myRowDataStatus.Add(
                                                New RowData With {.Column = columnKey, .Value = r.status, .CellType = CellTypes.str})
                                            myRowDataScores.Add(
                                                New RowData With {.Column = columnKey, .Value = r.score, .CellType = CellTypes.str})
                                            myRowDataCodes.Add(
                                                New RowData With {.Column = columnKey, .Value = r.code, .CellType = CellTypes.str})
                                        Next
                                    Next
                                Next
                            Else
                                For Each unit As Unit In
                                    From b As Booklet In personData.booklets
                                    From u As Unit In b.units Select u
                                    Dim subform As SubForm = (From sf As SubForm In unit.subforms Where sf.id = subformKey).FirstOrDefault
                                    If subform IsNot Nothing Then
                                        For Each r As ResponseData In subform.responses
                                            hasData = True
                                            Dim columnKey As String = Columns.Item(IIf(String.IsNullOrEmpty(unit.alias), unit.id, unit.alias) + r.id)
                                            myRowDataValues.Add(
                                                New RowData With {.Column = columnKey, .Value = r.value, .CellType = CellTypes.str})
                                            myRowDataStatus.Add(
                                                New RowData With {.Column = columnKey, .Value = r.status, .CellType = CellTypes.str})
                                            myRowDataScores.Add(
                                                New RowData With {.Column = columnKey, .Value = r.score, .CellType = CellTypes.str})
                                            myRowDataCodes.Add(
                                                New RowData With {.Column = columnKey, .Value = r.code, .CellType = CellTypes.str})
                                        Next
                                    End If
                                Next
                            End If
                            If hasData Then
                                myRow += 1
                                If TableValues IsNot Nothing Then xlsxFactory.AppendRow(myRow, myRowDataValues, TableValues)
                                If TableStatus IsNot Nothing Then xlsxFactory.AppendRow(myRow, myRowDataStatus, TableStatus)
                                If TableScores IsNot Nothing Then xlsxFactory.AppendRow(myRow, myRowDataScores, TableScores)
                                If TableCodes IsNot Nothing Then xlsxFactory.AppendRow(myRow, myRowDataCodes, TableCodes)
                            End If
                        Next
                    Next

                    '########################################################
                    If config.writeSessions Then
                        Dim TableSessions As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "Sessions")
                        xlsxFactory.SetCellValueString("A", 1, TableSessions, "coming soon", CellFormatting.RowHeader2, myStyles)

                    End If

                    ''########################################################
                    ''TimeOnPage
                    ''########################################################
                    'progressMax = globalOutputStore.personDataFull.Count
                    'progressCount = 1
                    'stepCount += 1
                    'Dim TableTimeOnUnit As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "TimeOnUnit")
                    'myRow = 1
                    'xlsxFactory.SetCellValueString("A", myRow, TableTimeOnUnit, "ID", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("A", TableTimeOnUnit, 20)
                    'xlsxFactory.SetCellValueString("B", myRow, TableTimeOnUnit, "Group", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("B", TableTimeOnUnit, 10)
                    'xlsxFactory.SetCellValueString("C", myRow, TableTimeOnUnit, "Login+Code", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("C", TableTimeOnUnit, 10)
                    'xlsxFactory.SetCellValueString("D", myRow, TableTimeOnUnit, "Booklet", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("D", TableTimeOnUnit, 10)
                    'xlsxFactory.SetCellValueString("E", myRow, TableTimeOnUnit, "Unit", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("E", TableTimeOnUnit, 10)
                    'xlsxFactory.SetCellValueString("F", myRow, TableTimeOnUnit, "Start At", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("F", TableTimeOnUnit, 10)
                    'xlsxFactory.SetCellValueString("G", myRow, TableTimeOnUnit, "Player Load Time", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("G", TableTimeOnUnit, 10)
                    'xlsxFactory.SetCellValueString("H", myRow, TableTimeOnUnit, "Stay Time", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("H", TableTimeOnUnit, 10)
                    'xlsxFactory.SetCellValueString("I", myRow, TableTimeOnUnit, "Was Paused", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("I", TableTimeOnUnit, 10)
                    'xlsxFactory.SetCellValueString("J", myRow, TableTimeOnUnit, "Lost Focus", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("J", TableTimeOnUnit, 10)
                    'xlsxFactory.SetCellValueString("K", myRow, TableTimeOnUnit, "Responses Some Time", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("K", TableTimeOnUnit, 10)
                    'xlsxFactory.SetCellValueString("L", myRow, TableTimeOnUnit, "Responses Complete Time", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("L", TableTimeOnUnit, 10)

                    'For Each testPerson As KeyValuePair(Of String, Person) In From p As KeyValuePair(Of String, Person) In globalOutputStore.personDataFull Order By p.Key
                    '    If worker.CancellationPending Then
                    '        e.Cancel = True
                    '        Exit For
                    '    End If
                    '    progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                    '    worker.ReportProgress(progressValue, "")
                    '    progressCount += 1

                    '    For Each booklet As Booklet In From b As Booklet In testPerson.Value.booklets Order By b.id
                    '        For Each unit As Unit In From u As Unit In booklet.units Order By u.alias
                    '            Dim topData As TimeOnPageData = unit.getTimeOnPageData()
                    '            Dim myRowData As New List(Of RowData)
                    '            myRowData.Add(New RowData With {.Column = "A", .Value = testPerson.Value.group + testPerson.Value.login +
                    '                          testPerson.Value.code + booklet.id, .CellType = CellTypes.str})
                    '            myRowData.Add(New RowData With {.Column = "B", .Value = testPerson.Value.group, .CellType = CellTypes.str})
                    '            myRowData.Add(New RowData With {.Column = "C", .Value = testPerson.Value.login + testPerson.Value.code, .CellType = CellTypes.str})
                    '            myRowData.Add(New RowData With {.Column = "D", .Value = booklet.id, .CellType = CellTypes.str})
                    '            myRowData.Add(New RowData With {.Column = "E", .Value = unit.alias, .CellType = CellTypes.str})
                    '            myRowData.Add(New RowData With {.Column = "F", .Value = topData.navigationStart, .CellType = CellTypes.int})
                    '            myRowData.Add(New RowData With {.Column = "G", .Value = topData.playerLoadTime, .CellType = CellTypes.int})
                    '            myRowData.Add(New RowData With {.Column = "H", .Value = topData.stayTime, .CellType = CellTypes.int})
                    '            myRowData.Add(New RowData With {.Column = "I", .Value = topData.wasPaused.ToString, .CellType = CellTypes.str})
                    '            myRowData.Add(New RowData With {.Column = "J", .Value = topData.lostFocus.ToString, .CellType = CellTypes.str})
                    '            myRowData.Add(New RowData With {.Column = "K", .Value = topData.responseProgressTimeSome, .CellType = CellTypes.int})
                    '            myRowData.Add(New RowData With {.Column = "L", .Value = topData.responseProgressTimeComplete, .CellType = CellTypes.int})
                    '            myRow += 1
                    '            xlsxFactory.AppendRow(myRow, myRowData, TableTimeOnUnit)
                    '        Next
                    '    Next
                    'Next

                    ''########################################################
                    ''TechData
                    ''########################################################
                    'Dim TableTechData As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "TechData")
                    'Dim currentUser As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent
                    'Dim currentUserName As String = currentUser.Name.Substring(currentUser.Name.IndexOf("\") + 1)

                    'xlsxFactory.SetCellValueString("A", 1, TableTechData, "Antworten und Log-Daten IQB-Testcenter", CellFormatting.Null, myStyles)
                    'xlsxFactory.SetCellValueString("A", 2, TableTechData, "konvertiert mit " + My.Application.Info.AssemblyName + " V" +
                    '                               My.Application.Info.Version.Major.ToString + "." + My.Application.Info.Version.Minor.ToString + "." +
                    '                               My.Application.Info.Version.Build.ToString + " am " + DateTime.Now.ToShortDateString + " " + DateTime.Now.ToShortTimeString +
                    '                               " (" + currentUserName + ")", CellFormatting.Null, myStyles)

                    'myRow = 4

                    'xlsxFactory.SetCellValueString("A", myRow, TableTechData, "ID", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("A", TableTechData, 30)
                    'xlsxFactory.SetCellValueString("B", myRow, TableTechData, "Start at", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("B", TableTechData, 20)
                    'xlsxFactory.SetCellValueString("C", myRow, TableTechData, "loadcomplete after", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("C", TableTechData, 20)
                    'xlsxFactory.SetCellValueString("D", myRow, TableTechData, "loadspeed", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("D", TableTechData, 20)
                    'xlsxFactory.SetCellValueString("E", myRow, TableTechData, "firstUnitRunning after", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("E", TableTechData, 20)
                    'xlsxFactory.SetCellValueString("F", myRow, TableTechData, "os", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("F", TableTechData, 20)
                    'xlsxFactory.SetCellValueString("G", myRow, TableTechData, "browser", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("G", TableTechData, 20)
                    'xlsxFactory.SetCellValueString("H", myRow, TableTechData, "screen", CellFormatting.RowHeader2, myStyles)
                    'xlsxFactory.SetColumnWidth("H", TableTechData, 20)

                    'progressMax = globalOutputStore.personDataFull.Count
                    'progressCount = 1
                    'stepCount += 1
                    'For Each testPerson As KeyValuePair(Of String, Person) In From p As KeyValuePair(Of String, Person) In globalOutputStore.personDataFull Order By p.Key
                    '    If worker.CancellationPending Then
                    '        e.Cancel = True
                    '        Exit For
                    '    End If
                    '    progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                    '    worker.ReportProgress(progressValue, "")
                    '    progressCount += 1

                    '    For Each booklet As Booklet In From b As Booklet In testPerson.Value.booklets Order By b.id
                    '        myRow += 1
                    '        Dim myRowData As New List(Of RowData)
                    '        Dim techData As BookletTechData = booklet.getTechData(globalOutputStore.bookletSizes)
                    '        myRowData.Add(New RowData With {.Column = "A", .Value = booklet.id, .CellType = CellTypes.str})
                    '        myRowData.Add(New RowData With {.Column = "B", .Value = techData.loadStart, .CellType = CellTypes.int})
                    '        myRowData.Add(New RowData With {.Column = "C", .Value = techData.loadTimeCompleteTS, .CellType = CellTypes.int})
                    '        myRowData.Add(New RowData With {.Column = "D", .Value = techData.loadspeed, .CellType = CellTypes.dec})
                    '        myRowData.Add(New RowData With {.Column = "E", .Value = techData.firstUnitRunningAfterMS, .CellType = CellTypes.int})
                    '        myRowData.Add(New RowData With {.Column = "F", .Value = techData.os, .CellType = CellTypes.str})
                    '        myRowData.Add(New RowData With {.Column = "G", .Value = techData.browser, .CellType = CellTypes.str})
                    '        myRowData.Add(New RowData With {.Column = "H", .Value = techData.screen, .CellType = CellTypes.str})

                    '        xlsxFactory.AppendRow(myRow, myRowData, TableTechData)
                    '    Next
                    'Next

                End Using
                worker.ReportProgress(100.0#, "Speichern Datei")
                Try
                    Using fs As New IO.FileStream(config.targetXlsxFilename, IO.FileMode.Create)
                        MemStream.WriteTo(fs)
                    End Using
                Catch ex As Exception
                    worker.ReportProgress(100.0#, "e: Konnte Datei nicht schreiben: " + ex.Message)
                End Try
            End Using
        End If
    End Sub

    Public Shared Sub WriteLite(
                           workbookTemplate As Byte(),
                           worker As BackgroundWorker,
                           e As DoWorkEventArgs,
                           targetXlsxFilename As String
                           )
        Dim AllVariables As New List(Of String)
        worker.ReportProgress(0.0#, "Ermittle Variablen")

        If globalOutputStore.personDataFull.Count > 0 Then
            For Each p As KeyValuePair(Of String, Person) In globalOutputStore.personDataFull
                For Each b As Booklet In p.Value.booklets
                    For Each u As Unit In b.units
                        For Each rSub As SubForm In u.subforms
                            Dim varPrefix As String = u.alias
                            For Each r As ResponseData In rSub.responses
                                If r.status = "VALUE_CHANGED" AndAlso Not AllVariables.Contains(varPrefix + r.id) Then AllVariables.Add(varPrefix + r.id)
                            Next
                        Next
                    Next
                Next
            Next
        Else
            If globalOutputStore.personResponses.Count > 0 Then
                For Each p As PersonResponses In globalOutputStore.personResponses
                    For Each sf As SubForm In p.subforms
                        For Each r In sf.responses
                            If r.status = "VALUE_CHANGED" AndAlso Not AllVariables.Contains(r.id) Then AllVariables.Add(r.id)
                        Next
                    Next
                Next
            End If
        End If
        worker.ReportProgress(0.0#, AllVariables.Count.ToString + " Variablen gefunden.")

        If AllVariables.Count > 0 Then
            Using MemStream As New IO.MemoryStream()
                MemStream.Write(workbookTemplate, 0, workbookTemplate.Length)
                Using ZielXLS As SpreadsheetDocument = SpreadsheetDocument.Open(MemStream, True)
                    Dim myStyles As ExcelStyleDefs = xlsxFactory.AddIQBStandardStyles(ZielXLS.WorkbookPart)
                    '########################################################
                    'Responses
                    '########################################################
                    Dim TableResponses As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "Responses")
                    worker.ReportProgress(0.0#, "Schreibe Daten")

                    Dim myRow As Integer = 1
                    xlsxFactory.SetCellValueString("A", myRow, TableResponses, "ID", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("A", TableResponses, 20)
                    xlsxFactory.SetCellValueString("B", myRow, TableResponses, "Group", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("B", TableResponses, 10)
                    xlsxFactory.SetCellValueString("C", myRow, TableResponses, "Login+Code", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("C", TableResponses, 10)
                    xlsxFactory.SetCellValueString("D", myRow, TableResponses, "Booklet", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("D", TableResponses, 10)

                    Dim myColumn As String = "E"
                    Dim Columns As New Dictionary(Of String, String)

                    Dim progressMax As Integer = AllVariables.Count
                    Dim progressCount As Integer = 1
                    Dim stepMax As Integer = 5
                    Dim stepCount As Integer = 1
                    Dim progressValue As Double = 0.0#

                    For Each s As String In From v As String In AllVariables Order By v Select v
                        progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                        worker.ReportProgress(progressValue, "")
                        progressCount += 1
                        xlsxFactory.SetCellValueString(myColumn, myRow, TableResponses, s, CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth(myColumn, TableResponses, 10)
                        Columns.Add(s, myColumn)
                        myColumn = xlsxFactory.GetNextColumn(myColumn)
                    Next


                    progressMax = IIf(globalOutputStore.personDataFull.Count > 0, globalOutputStore.personDataFull.Count, globalOutputStore.personResponses.Count)
                    progressCount = 1
                    stepCount += 1

                    If globalOutputStore.personDataFull.Count > 0 Then
                        For Each person As Person In
                            From kvp As KeyValuePair(Of String, Person) In globalOutputStore.personDataFull
                            Order By kvp.Key Select kvp.Value
                            If worker.CancellationPending Then
                                e.Cancel = True
                                Exit For
                            End If
                            progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                            worker.ReportProgress(progressValue, "")
                            progressCount += 1
                            For Each booklet As Booklet In From b As Booklet In person.booklets Order By b.id
                                Dim allSubForms As List(Of String) = (From u As Unit In booklet.units From sf As SubForm In u.subforms Select sf.id).Distinct.ToList
                                For Each subFormId As String In From sf As String In allSubForms Order By sf
                                    Dim myRowDataResponses As New List(Of RowData)
                                    Dim personKey As String = person.group + person.login + person.code + subFormId
                                    myRowDataResponses.Add(New RowData With {.Column = "A", .Value = personKey + booklet.id, .CellType = CellTypes.str})
                                    myRowDataResponses.Add(New RowData With {.Column = "B", .Value = person.group, .CellType = CellTypes.str})
                                    myRowDataResponses.Add(New RowData With {.Column = "C", .Value = person.login + person.code + IIf(String.IsNullOrEmpty(subFormId), "", "_" + subFormId), .CellType = CellTypes.str})
                                    myRowDataResponses.Add(New RowData With {.Column = "D", .Value = booklet.id, .CellType = CellTypes.str})
                                    For Each unit As Unit In booklet.units
                                        Dim varPrefix As String = unit.alias
                                        Dim mySubForm As SubForm = (From sf As SubForm In unit.subforms Where sf.id = subFormId).FirstOrDefault
                                        If mySubForm IsNot Nothing Then
                                            For Each response As ResponseData In mySubForm.responses
                                                If response.status = "VALUE_CHANGED" AndAlso AllVariables.Contains(varPrefix + response.id) Then
                                                    myRowDataResponses.Add(New RowData With {.Column = Columns.Item(varPrefix + response.id), .Value = response.value, .CellType = CellTypes.str})
                                                End If
                                            Next
                                        End If
                                    Next
                                    myRow += 1
                                    xlsxFactory.AppendRow(myRow, myRowDataResponses, TableResponses, myStyles)
                                Next
                            Next
                        Next
                    Else
                        For Each person As PersonResponses In
                            From p As PersonResponses In globalOutputStore.personResponses
                            Let key = p.group + p.login + p.code + p.booklet
                            Order By key
                            Select p
                            If worker.CancellationPending Then
                                e.Cancel = True
                                Exit For
                            End If
                            progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                            worker.ReportProgress(progressValue, "")
                            progressCount += 1
                            For Each sf As SubForm In person.subforms
                                Dim myRowDataResponses As New List(Of RowData)
                                Dim personKey As String = person.group + person.login + person.code + sf.id
                                myRowDataResponses.Add(New RowData With {.Column = "A", .Value = personKey + person.booklet, .CellType = CellTypes.str})
                                myRowDataResponses.Add(New RowData With {.Column = "B", .Value = person.group, .CellType = CellTypes.str})
                                myRowDataResponses.Add(New RowData With {.Column = "C", .Value = person.login + person.code + sf.id, .CellType = CellTypes.str})
                                myRowDataResponses.Add(New RowData With {.Column = "D", .Value = person.booklet, .CellType = CellTypes.str})
                                For Each r As ResponseData In sf.responses
                                    If AllVariables.Contains(r.id) Then
                                        myRowDataResponses.Add(New RowData With {.Column = Columns.Item(r.id), .Value = r.value, .CellType = CellTypes.str})
                                    End If
                                Next
                                myRow += 1
                                xlsxFactory.AppendRow(myRow, myRowDataResponses, TableResponses, myStyles)
                            Next
                        Next
                    End If

                End Using
                worker.ReportProgress(100.0#, "Speichern Datei")
                Try
                    Using fs As New IO.FileStream(targetXlsxFilename, IO.FileMode.Create)
                        MemStream.WriteTo(fs)
                    End Using
                Catch ex As Exception
                    worker.ReportProgress(100.0#, "e: Konnte Datei nicht schreiben: " + ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Shared Function getVariableListFromStore(addSubformSuffix As Boolean) As List(Of String)
        Dim returnList As New List(Of String)
        For Each p As KeyValuePair(Of String, Person) In globalOutputStore.personDataFull
            For Each b As Booklet In p.Value.booklets
                For Each u As Unit In b.units
                    For Each rSub As SubForm In u.subforms
                        For Each r As ResponseData In rSub.responses
                            Dim varId As String = WriteOutputToXlsx.getColumnKey(u, r.id, IIf(addSubformSuffix, rSub.id, ""))
                            If Not returnList.Contains(varId) Then returnList.Add(varId)
                        Next
                    Next
                Next
            Next
        Next
        Return returnList
    End Function

    Public Shared Function getColumnKey(unit As Unit, variableId As String, subformKey As String) As String
        Return IIf(String.IsNullOrEmpty(unit.alias), unit.id, unit.alias) + variableId + IIf(String.IsNullOrEmpty(subformKey), "", "##" + subformKey)
    End Function
End Class
