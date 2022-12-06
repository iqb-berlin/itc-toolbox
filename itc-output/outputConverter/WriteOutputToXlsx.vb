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
                           AllVariables As List(Of String),
                           AllPeople As Dictionary(Of String, Dictionary(Of String, List(Of UnitLineData))),
                           myTestPersonList As TestPersonList,
                           bookletSizes As Dictionary(Of String, Long),
                           targetXlsxFilename As String
                           )
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

                Dim BookletUnits As New Dictionary(Of String, List(Of String)) 'für unten TechTable

                progressMax = AllPeople.Count
                progressCount = 1
                stepCount += 1
                For Each persondata As Dictionary(Of String, List(Of UnitLineData)) In
                    From kvp As KeyValuePair(Of String, Dictionary(Of String, List(Of UnitLineData))) In AllPeople Order By kvp.Key Select kvp.Value
                    If worker.CancellationPending Then
                        e.Cancel = True
                        Exit For
                    End If
                    progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                    worker.ReportProgress(progressValue, "")
                    progressCount += 1
                    For Each bookletData As KeyValuePair(Of String, List(Of UnitLineData)) In
                        From kvp2 As KeyValuePair(Of String, List(Of UnitLineData)) In persondata Order By kvp2.Key

                        Dim bookletPeople As List(Of String) = (From ud As UnitLineData In bookletData.Value
                                                                From s As KeyValuePair(Of String, List(Of ResponseData)) In ud.responses
                                                                Select s.Key).Distinct.ToList()
                        bookletPeople.Sort()
                        For Each subPerson As String In bookletPeople
                            Dim myRowData As New List(Of RowData)
                            For Each unitData As UnitLineData In bookletData.Value
                                If Not BookletUnits.ContainsKey(unitData.bookletname) Then BookletUnits.Add(unitData.bookletname, New List(Of String))
                                If Not BookletUnits.Item(unitData.bookletname).Contains(unitData.unitname) Then BookletUnits.Item(unitData.bookletname).Add(unitData.unitname)
                                If unitData.responses.ContainsKey(subPerson) Then
                                    Dim respData As List(Of ResponseData) = unitData.responses.Item(subPerson)
                                    If respData.Count > 0 Then
                                        If myRowData.Count = 0 Then
                                            myRowData.Add(New RowData With {.Column = "A", .Value = unitData.personKey + unitData.bookletname, .CellType = CellTypes.str})
                                            myRowData.Add(New RowData With {.Column = "B", .Value = unitData.groupname, .CellType = CellTypes.str})
                                            myRowData.Add(New RowData With {.Column = "C", .Value = unitData.loginname + unitData.code + IIf(String.IsNullOrEmpty(subPerson), "", "_" + subPerson), .CellType = CellTypes.str})
                                            myRowData.Add(New RowData With {.Column = "D", .Value = unitData.bookletname, .CellType = CellTypes.str})
                                        End If
                                        For Each rd As ResponseData In respData
                                            myRowData.Add(New RowData With {.Column = Columns.Item(unitData.unitname + "##" + rd.variableId), .Value = rd.value, .CellType = CellTypes.str})
                                        Next
                                    End If
                                End If
                            Next
                            If myRowData.Count > 0 Then
                                myRow += 1
                                xlsxFactory.AppendRow(myRow, myRowData, TableResponses)
                            End If
                        Next
                    Next
                Next


                '########################################################
                'TimeOnPage
                '########################################################
                progressMax = myTestPersonList.Count
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

                For Each tc As KeyValuePair(Of String, TestPerson) In myTestPersonList
                    If worker.CancellationPending Then
                        e.Cancel = True
                        Exit For
                    End If
                    progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                    worker.ReportProgress(progressValue, "")
                    progressCount += 1
                    For Each tou As TimeOnUnit In tc.Value.GetTimeOnUnitList

                        Dim myRowData As New List(Of RowData)
                        myRowData.Add(New RowData With {.Column = "A", .Value = tc.Value.group + tc.Value.login + tc.Value.code + tc.Value.booklet, .CellType = CellTypes.str})
                        myRowData.Add(New RowData With {.Column = "B", .Value = tc.Value.group, .CellType = CellTypes.str})
                        myRowData.Add(New RowData With {.Column = "C", .Value = tc.Value.login + tc.Value.code, .CellType = CellTypes.str})
                        myRowData.Add(New RowData With {.Column = "D", .Value = tc.Value.booklet, .CellType = CellTypes.str})
                        myRowData.Add(New RowData With {.Column = "E", .Value = tou.unit, .CellType = CellTypes.str})
                        myRowData.Add(New RowData With {.Column = "F", .Value = tou.navigationStart, .CellType = CellTypes.int})
                        myRowData.Add(New RowData With {.Column = "G", .Value = tou.playerLoadTime, .CellType = CellTypes.int})
                        myRowData.Add(New RowData With {.Column = "H", .Value = tou.stayTime, .CellType = CellTypes.int})
                        myRowData.Add(New RowData With {.Column = "I", .Value = tou.wasPaused.ToString, .CellType = CellTypes.str})
                        myRowData.Add(New RowData With {.Column = "J", .Value = tou.lostFocus.ToString, .CellType = CellTypes.str})
                        myRowData.Add(New RowData With {.Column = "K", .Value = tou.responseProgressTimeSome, .CellType = CellTypes.int})
                        myRowData.Add(New RowData With {.Column = "L", .Value = tou.responseProgressTimeComplete, .CellType = CellTypes.int})
                        myRow += 1
                        xlsxFactory.AppendRow(myRow, myRowData, TableTimeOnUnit)
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

                progressMax = myTestPersonList.Count
                progressCount = 1
                stepCount += 1
                For Each tc As KeyValuePair(Of String, TestPerson) In myTestPersonList
                    If worker.CancellationPending Then
                        e.Cancel = True
                        Exit For
                    End If
                    progressValue = progressCount * (100 / stepMax) / progressMax + (100 / stepMax) * (stepCount - 1)
                    worker.ReportProgress(progressValue, "")
                    progressCount += 1

                    myRow += 1
                    Dim myRowData As New List(Of RowData)
                    myRowData.Add(New RowData With {.Column = "A", .Value = tc.Key, .CellType = CellTypes.str})
                    myRowData.Add(New RowData With {.Column = "B", .Value = tc.Value.loadStart, .CellType = CellTypes.int})
                    myRowData.Add(New RowData With {.Column = "C", .Value = tc.Value.loadtime, .CellType = CellTypes.int})
                    myRowData.Add(New RowData With {.Column = "D", .Value = tc.Value.loadspeed(bookletSizes).ToString(), .CellType = CellTypes.dec})
                    myRowData.Add(New RowData With {.Column = "E", .Value = tc.Value.getFirstPlayerRunning, .CellType = CellTypes.int})
                    myRowData.Add(New RowData With {.Column = "F", .Value = tc.Value.os, .CellType = CellTypes.str})
                    myRowData.Add(New RowData With {.Column = "G", .Value = tc.Value.browser, .CellType = CellTypes.str})
                    myRowData.Add(New RowData With {.Column = "H", .Value = tc.Value.screen, .CellType = CellTypes.str})

                    xlsxFactory.AppendRow(myRow, myRowData, TableTechData)
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
    End Sub
End Class
