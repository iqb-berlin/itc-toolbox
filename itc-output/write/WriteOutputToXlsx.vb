Imports Newtonsoft.Json
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports DocumentFormat.OpenXml.Packaging
Imports iqb.lib.openxml
Imports System.ComponentModel
Class WriteOutputToXlsx
    Public Shared Sub Write(
                           workbookTemplate As Byte(),
                           worker As BackgroundWorker,
                           e As DoWorkEventArgs,
                           targetXlsxFilename As String
                           )
        Dim AllVariables As New List(Of String)
        worker.ReportProgress(0.0#, "Ermittle Variablen")

        For Each p As KeyValuePair(Of String, Person) In globalOutputStore.personData
            For Each b As Booklet In p.Value.booklets
                For Each u As Unit In b.units
                    For Each rSub As SubForm In u.subforms
                        Dim varPrefix As String = u.alias
                        If Not String.IsNullOrEmpty(rSub.id) Then varPrefix += "##" + rSub.id
                        For Each r As ResponseData In rSub.responses
                            If r.status = "VALUE_CHANGED" AndAlso Not AllVariables.Contains(varPrefix + "##" + r.id) Then AllVariables.Add(varPrefix + "##" + r.id)
                        Next
                    Next
                Next
            Next
        Next
        For Each p As PersonResponses In globalOutputStore.personResponses
            For Each r As ResponseData In p.responses
                If r.status = "VALUE_CHANGED" AndAlso Not AllVariables.Contains(r.id) Then AllVariables.Add(r.id)
            Next
        Next
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
                    Dim TableStatus As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "Status")
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

                    xlsxFactory.SetCellValueString("A", myRow, TableStatus, "ID", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("A", TableStatus, 20)
                    xlsxFactory.SetCellValueString("B", myRow, TableStatus, "Group", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("B", TableStatus, 10)
                    xlsxFactory.SetCellValueString("C", myRow, TableStatus, "Login+Code", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("C", TableStatus, 10)
                    xlsxFactory.SetCellValueString("D", myRow, TableStatus, "Booklet", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("D", TableStatus, 10)
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
                        xlsxFactory.SetCellValueString(myColumn, myRow, TableStatus, s, CellFormatting.RowHeader2, myStyles)
                        xlsxFactory.SetColumnWidth(myColumn, TableStatus, 10)
                        Columns.Add(s, myColumn)
                        myColumn = xlsxFactory.GetNextColumn(myColumn)
                    Next

                    Dim BookletUnits As New Dictionary(Of String, List(Of String)) 'für unten TechTable

                    progressMax = globalOutputStore.personData.Count
                    progressCount = 1
                    stepCount += 1
                    For Each person As Person In
                        From kvp As KeyValuePair(Of String, Person) In globalOutputStore.personData Order By kvp.Key Select kvp.Value
                        If worker.CancellationPending Then
                            e.Cancel = True
                            Exit For
                        End If
                        progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                        worker.ReportProgress(progressValue, "")
                        progressCount += 1
                        For Each booklet As Booklet In From b As Booklet In person.booklets Order By b.id
                            Dim myRowDataResponses As New List(Of RowData)
                            Dim myRowDataStatus As New List(Of RowData)
                            For Each unit As Unit In booklet.units
                                For Each rSub As SubForm In unit.subforms
                                    Dim varPrefix As String = unit.alias
                                    If Not String.IsNullOrEmpty(rSub.id) Then varPrefix += "##" + rSub.id
                                    For Each response As ResponseData In rSub.responses
                                        If AllVariables.Contains(varPrefix + "##" + response.id) Then
                                            'If Not BookletUnits.ContainsKey(unitData.bookletname) Then BookletUnits.Add(unitData.bookletname, New List(Of String))
                                            'If Not BookletUnits.Item(unitData.bookletname).Contains(unitData.unitname) Then BookletUnits.Item(unitData.bookletname).Add(unitData.unitname)
                                            If myRowDataResponses.Count = 0 Then
                                                Dim personKey As String = person.group + person.login + person.code
                                                myRowDataResponses.Add(New RowData With {.Column = "A", .Value = personKey + booklet.id, .CellType = CellTypes.str})
                                                myRowDataResponses.Add(New RowData With {.Column = "B", .Value = person.group, .CellType = CellTypes.str})
                                                myRowDataResponses.Add(New RowData With {.Column = "C", .Value = person.login + person.code + IIf(String.IsNullOrEmpty(rSub.id), "", "_" + rSub.id), .CellType = CellTypes.str})
                                                myRowDataResponses.Add(New RowData With {.Column = "D", .Value = booklet.id, .CellType = CellTypes.str})

                                                'myRowDataStatus.Add(New RowData With {.Column = "A", .Value = personKey + unitdata.bookletname, .CellType = CellTypes.str})
                                                'myRowDataStatus.Add(New RowData With {.Column = "B", .Value = unitdata.groupname, .CellType = CellTypes.str})
                                                'myRowDataStatus.Add(New RowData With {.Column = "C", .Value = unitdata.loginname + unitdata.code + IIf(String.IsNullOrEmpty(subPerson), "", "_" + subPerson), .CellType = CellTypes.str})
                                                'myRowDataStatus.Add(New RowData With {.Column = "D", .Value = unitdata.bookletname, .CellType = CellTypes.str})
                                            End If
                                            myRowDataResponses.Add(New RowData With {.Column = Columns.Item(varPrefix + "##" + response.id), .Value = response.value, .CellType = CellTypes.str})
                                            'myRowDataStatus.Add(New RowData With {.Column = Columns.Item(unitData.unitname + "##" + rd.id), .Value = rd.status, .CellType = CellTypes.str})
                                        End If
                                    Next
                                Next
                            Next
                            If myRowDataResponses.Count > 0 Then
                                myRow += 1
                                xlsxFactory.AppendRow(myRow, myRowDataResponses, TableResponses)
                                'xlsxFactory.AppendRow(myRow, myRowDataStatus, TableStatus)
                            End If
                        Next
                    Next

                    progressMax = globalOutputStore.personResponses.Count
                    progressCount = 1
                    stepCount += 1
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
                        Dim myRowDataResponses As New List(Of RowData)
                        Dim myRowDataStatus As New List(Of RowData)
                        For Each response As ResponseData In person.responses
                            If AllVariables.Contains(response.id) Then
                                If myRowDataResponses.Count = 0 Then
                                    Dim personKey As String = person.group + person.login + person.code
                                    myRowDataResponses.Add(New RowData With {.Column = "A", .Value = personKey + person.booklet, .CellType = CellTypes.str})
                                    myRowDataResponses.Add(New RowData With {.Column = "B", .Value = person.group, .CellType = CellTypes.str})
                                    myRowDataResponses.Add(New RowData With {.Column = "C", .Value = person.login + person.code, .CellType = CellTypes.str})
                                    myRowDataResponses.Add(New RowData With {.Column = "D", .Value = person.booklet, .CellType = CellTypes.str})
                                End If
                                myRowDataResponses.Add(New RowData With {.Column = Columns.Item(response.id), .Value = response.value, .CellType = CellTypes.str})
                            End If
                        Next
                        If myRowDataResponses.Count > 0 Then
                            myRow += 1
                            xlsxFactory.AppendRow(myRow, myRowDataResponses, TableResponses)
                            'xlsxFactory.AppendRow(myRow, myRowDataStatus, TableStatus)
                        End If
                    Next

                    '########################################################
                    'TimeOnPage
                    '########################################################
                    progressMax = globalOutputStore.personData.Count
                    progressCount = 1
                    stepCount += 1
                    Dim TableTimeOnUnit As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "TimeOnUnit")
                    myRow = 1
                    xlsxFactory.SetCellValueString("A", myRow, TableTimeOnUnit, "ID", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("A", TableTimeOnUnit, 20)
                    xlsxFactory.SetCellValueString("B", myRow, TableTimeOnUnit, "Group", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("B", TableTimeOnUnit, 10)
                    xlsxFactory.SetCellValueString("C", myRow, TableTimeOnUnit, "Login+Code", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("C", TableTimeOnUnit, 10)
                    xlsxFactory.SetCellValueString("D", myRow, TableTimeOnUnit, "Booklet", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("D", TableTimeOnUnit, 10)
                    xlsxFactory.SetCellValueString("E", myRow, TableTimeOnUnit, "Unit", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("E", TableTimeOnUnit, 10)
                    xlsxFactory.SetCellValueString("F", myRow, TableTimeOnUnit, "Start At", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("F", TableTimeOnUnit, 10)
                    xlsxFactory.SetCellValueString("G", myRow, TableTimeOnUnit, "Player Load Time", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("G", TableTimeOnUnit, 10)
                    xlsxFactory.SetCellValueString("H", myRow, TableTimeOnUnit, "Stay Time", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("H", TableTimeOnUnit, 10)
                    xlsxFactory.SetCellValueString("I", myRow, TableTimeOnUnit, "Was Paused", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("I", TableTimeOnUnit, 10)
                    xlsxFactory.SetCellValueString("J", myRow, TableTimeOnUnit, "Lost Focus", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("J", TableTimeOnUnit, 10)
                    xlsxFactory.SetCellValueString("K", myRow, TableTimeOnUnit, "Responses Some Time", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("K", TableTimeOnUnit, 10)
                    xlsxFactory.SetCellValueString("L", myRow, TableTimeOnUnit, "Responses Complete Time", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("L", TableTimeOnUnit, 10)

                    For Each testPerson As KeyValuePair(Of String, Person) In From p As KeyValuePair(Of String, Person) In globalOutputStore.personData Order By p.Key
                        If worker.CancellationPending Then
                            e.Cancel = True
                            Exit For
                        End If
                        progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                        worker.ReportProgress(progressValue, "")
                        progressCount += 1

                        For Each booklet As Booklet In From b As Booklet In testPerson.Value.booklets Order By b.id
                            For Each unit As Unit In From u As Unit In booklet.units Order By u.alias
                                Dim topData As TimeOnPageData = unit.getTimeOnPageData()
                                Dim myRowData As New List(Of RowData)
                                myRowData.Add(New RowData With {.Column = "A", .Value = testPerson.Value.group + testPerson.Value.login +
                                              testPerson.Value.code + booklet.id, .CellType = CellTypes.str})
                                myRowData.Add(New RowData With {.Column = "B", .Value = testPerson.Value.group, .CellType = CellTypes.str})
                                myRowData.Add(New RowData With {.Column = "C", .Value = testPerson.Value.login + testPerson.Value.code, .CellType = CellTypes.str})
                                myRowData.Add(New RowData With {.Column = "D", .Value = booklet.id, .CellType = CellTypes.str})
                                myRowData.Add(New RowData With {.Column = "E", .Value = unit.alias, .CellType = CellTypes.str})
                                myRowData.Add(New RowData With {.Column = "F", .Value = topData.navigationStart, .CellType = CellTypes.int})
                                myRowData.Add(New RowData With {.Column = "G", .Value = topData.playerLoadTime, .CellType = CellTypes.int})
                                myRowData.Add(New RowData With {.Column = "H", .Value = topData.stayTime, .CellType = CellTypes.int})
                                myRowData.Add(New RowData With {.Column = "I", .Value = topData.wasPaused.ToString, .CellType = CellTypes.str})
                                myRowData.Add(New RowData With {.Column = "J", .Value = topData.lostFocus.ToString, .CellType = CellTypes.str})
                                myRowData.Add(New RowData With {.Column = "K", .Value = topData.responseProgressTimeSome, .CellType = CellTypes.int})
                                myRowData.Add(New RowData With {.Column = "L", .Value = topData.responseProgressTimeComplete, .CellType = CellTypes.int})
                                myRow += 1
                                xlsxFactory.AppendRow(myRow, myRowData, TableTimeOnUnit)
                            Next
                        Next
                    Next

                    '########################################################
                    'TechData
                    '########################################################
                    Dim TableTechData As WorksheetPart = xlsxFactory.InsertWorksheet(ZielXLS.WorkbookPart, "TechData")
                    Dim currentUser As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent
                    Dim currentUserName As String = currentUser.Name.Substring(currentUser.Name.IndexOf("\") + 1)

                    xlsxFactory.SetCellValueString("A", 1, TableTechData, "Antworten und Log-Daten IQB-Testcenter", CellFormatting.Null, myStyles)
                    xlsxFactory.SetCellValueString("A", 2, TableTechData, "konvertiert mit " + My.Application.Info.AssemblyName + " V" +
                                                   My.Application.Info.Version.Major.ToString + "." + My.Application.Info.Version.Minor.ToString + "." +
                                                   My.Application.Info.Version.Build.ToString + " am " + DateTime.Now.ToShortDateString + " " + DateTime.Now.ToShortTimeString +
                                                   " (" + currentUserName + ")", CellFormatting.Null, myStyles)

                    myRow = 4

                    xlsxFactory.SetCellValueString("A", myRow, TableTechData, "ID", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("A", TableTechData, 30)
                    xlsxFactory.SetCellValueString("B", myRow, TableTechData, "Start at", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("B", TableTechData, 20)
                    xlsxFactory.SetCellValueString("C", myRow, TableTechData, "loadcomplete after", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("C", TableTechData, 20)
                    xlsxFactory.SetCellValueString("D", myRow, TableTechData, "loadspeed", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("D", TableTechData, 20)
                    xlsxFactory.SetCellValueString("E", myRow, TableTechData, "firstUnitRunning after", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("E", TableTechData, 20)
                    xlsxFactory.SetCellValueString("F", myRow, TableTechData, "os", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("F", TableTechData, 20)
                    xlsxFactory.SetCellValueString("G", myRow, TableTechData, "browser", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("G", TableTechData, 20)
                    xlsxFactory.SetCellValueString("H", myRow, TableTechData, "screen", CellFormatting.RowHeader2, myStyles)
                    xlsxFactory.SetColumnWidth("H", TableTechData, 20)

                    progressMax = globalOutputStore.personData.Count
                    progressCount = 1
                    stepCount += 1
                    For Each testPerson As KeyValuePair(Of String, Person) In From p As KeyValuePair(Of String, Person) In globalOutputStore.personData Order By p.Key
                        If worker.CancellationPending Then
                            e.Cancel = True
                            Exit For
                        End If
                        progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                        worker.ReportProgress(progressValue, "")
                        progressCount += 1

                        For Each booklet As Booklet In From b As Booklet In testPerson.Value.booklets Order By b.id
                            myRow += 1
                            Dim myRowData As New List(Of RowData)
                            Dim techData As BookletTechData = booklet.getTechData(globalOutputStore.bookletSizes)
                            myRowData.Add(New RowData With {.Column = "A", .Value = booklet.id, .CellType = CellTypes.str})
                            myRowData.Add(New RowData With {.Column = "B", .Value = techData.loadStart, .CellType = CellTypes.int})
                            myRowData.Add(New RowData With {.Column = "C", .Value = techData.loadTimeCompleteTS, .CellType = CellTypes.int})
                            myRowData.Add(New RowData With {.Column = "D", .Value = techData.loadspeed, .CellType = CellTypes.dec})
                            myRowData.Add(New RowData With {.Column = "E", .Value = techData.firstUnitRunningAfterMS, .CellType = CellTypes.int})
                            myRowData.Add(New RowData With {.Column = "F", .Value = techData.os, .CellType = CellTypes.str})
                            myRowData.Add(New RowData With {.Column = "G", .Value = techData.browser, .CellType = CellTypes.str})
                            myRowData.Add(New RowData With {.Column = "H", .Value = techData.screen, .CellType = CellTypes.str})

                            xlsxFactory.AppendRow(myRow, myRowData, TableTechData)
                        Next
                    Next

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
End Class
