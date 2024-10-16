Imports Newtonsoft.Json
Public Class JsonReadWrite
    Public Shared Sub Write(targetJsonFilename As String)
        Using file As New IO.StreamWriter(targetJsonFilename)
            Dim js As New JsonSerializer()
            js.Formatting = Formatting.Indented
            js.Serialize(file,
                         From p As KeyValuePair(Of String, Person) In globalOutputStore.personDataFull
                         Select p.Value)
        End Using
    End Sub

    Public Shared Sub WriteByGroup(targetFoldername As String)
        Dim groups As List(Of String) = (From p As KeyValuePair(Of String, Person) In globalOutputStore.personDataFull Select p.Value.group).Distinct.ToList
        For Each g As String In groups
            Using file As New IO.StreamWriter(targetFoldername + IO.Path.DirectorySeparatorChar + g + ".json")
                Dim js As New JsonSerializer()
                js.Formatting = Formatting.Indented
                js.Serialize(file,
                         From p As KeyValuePair(Of String, Person) In globalOutputStore.personDataFull Where p.Value.group = g
                         Select p.Value)
            End Using
        Next
    End Sub

    Public Shared Sub WriteBigData(targetFoldername As String)
        For Each big As KeyValuePair(Of String, String) In globalOutputStore.bigData
            Using file As New IO.StreamWriter(targetFoldername + IO.Path.DirectorySeparatorChar + big.Key)
                file.Write(big.Value)
            End Using
        Next
    End Sub

    Public Shared Sub WriteBooklets(targetJsonFilename As String)
        Using file As New IO.StreamWriter(targetJsonFilename)
            Dim js As New JsonSerializer()
            js.Formatting = Formatting.Indented
            js.Serialize(file, globalOutputStore.bookletSizes)
        End Using
    End Sub

    Public Shared Sub ReadFull(sourceJsonFilenames As String(), myworker As ComponentModel.BackgroundWorker)
        Dim progressMax As Integer = sourceJsonFilenames.Length
        Dim progressCount As Integer = 1
        Dim progressValue As Double
        For Each fn In sourceJsonFilenames
            progressValue = progressCount * (100 / progressMax)
            progressCount += 1
            myworker.ReportProgress(progressValue, IO.Path.GetFileName(fn))
            Try
                Using file As New IO.StreamReader(fn)
                    Dim js As New JsonSerializer()
                    Dim groupData As List(Of Person) = js.Deserialize(file, GetType(List(Of Person)))
                    For Each p As Person In groupData
                        globalOutputStore.personDataFull.Add(p.group + p.login + p.code, p)
                    Next
                End Using
            Catch ex As Exception
                myworker.ReportProgress(progressValue, "Fehler " + IO.Path.GetFileName(fn) + ": " + ex.Message)
            End Try
        Next
    End Sub

    Public Shared Sub ReadResponsesOnly(sourceJsonFilenames As String(), myworker As ComponentModel.BackgroundWorker,
                                        ignoreDisplayed As Boolean, ignoreNotReached As Boolean)
        Dim progressMax As Integer = sourceJsonFilenames.Length
        Dim progressCount As Integer = 1
        Dim progressValue As Double = 0.0#
        For Each fn In sourceJsonFilenames
            progressValue = progressCount * (100 / progressMax)
            progressCount += 1
            myworker.ReportProgress(progressValue, IO.Path.GetFileName(fn))
            Try
                Using file As New IO.StreamReader(fn)
                    Dim js As New JsonSerializer()
                    Dim groupData As List(Of Person) = js.Deserialize(file, GetType(List(Of Person)))
                    For Each p As Person In groupData
                        For Each b As Booklet In p.booklets
                            Dim newPR As New PersonResponses With {.group = p.group, .login = p.login, .code = p.code, .booklet = b.id, .subforms = New List(Of SubForm)}
                            Dim hasResponses As Boolean = False
                            For Each u As Unit In b.units
                                For Each sf As SubForm In u.subforms
                                    Dim newSubform As New SubForm With {.id = sf.id, .responses = New List(Of ResponseData)}
                                    For Each r As ResponseData In sf.responses
                                        If (Not ignoreDisplayed OrElse r.status <> ResponseSymbols.STATUS_DISPLAYED) AndAlso
                                                (Not ignoreNotReached OrElse r.status <> ResponseSymbols.STATUS_NOT_REACHED) Then
                                            r.id = u.alias + r.id
                                            newSubform.responses.Add(r)
                                        End If
                                    Next
                                    If newSubform.responses.Count > 0 Then
                                        hasResponses = True
                                        Dim sfToFill As SubForm = (From sfp As SubForm In newPR.subforms Where sfp.id = newSubform.id).FirstOrDefault
                                        If sfToFill Is Nothing Then
                                            newPR.subforms.Add(newSubform)
                                        Else
                                            sfToFill.responses.AddRange(newSubform.responses)
                                        End If
                                    End If
                                Next
                            Next
                            If hasResponses Then globalOutputStore.personResponses.Add(newPR)
                        Next
                    Next
                End Using
            Catch ex As Exception
                myworker.ReportProgress(progressValue, "Fehler " + IO.Path.GetFileName(fn) + ": " + ex.Message)
            End Try
        Next
    End Sub

    Public Shared Sub ReadLogsOnly(sourceJsonFilenames As String(), myworker As ComponentModel.BackgroundWorker)
        Dim progressMax As Integer = sourceJsonFilenames.Length
        Dim progressCount As Integer = 1
        Dim progressValue As Double = 0.0#
        For Each fn In sourceJsonFilenames
            progressValue = progressCount * (100 / progressMax)
            progressCount += 1
            myworker.ReportProgress(progressValue, IO.Path.GetFileName(fn))
            Try
                Using file As New IO.StreamReader(fn)
                    Dim js As New JsonSerializer()
                    Dim groupData As List(Of Person) = js.Deserialize(file, GetType(List(Of Person)))
                    For Each p As Person In groupData
                        For Each b As Booklet In p.booklets
                            Dim newPR As New PersonResponses With {.group = p.group, .login = p.login, .code = p.code, .booklet = b.id, .subforms = New List(Of SubForm)}
                            Dim hasResponses As Boolean = False
                            For Each u As Unit In b.units
                                For Each sf As SubForm In u.subforms
                                    'Dim newSubform As New SubForm With {.id = sf.id, .responses = New List(Of ResponseData)}
                                    'For Each r As ResponseData In sf.responses
                                    '    If (Not ignoreDisplayed OrElse r.status <> ResponseSymbols.STATUS_DISPLAYED) AndAlso
                                    '            (Not ignoreNotReached OrElse r.status <> ResponseSymbols.STATUS_NOT_REACHED) Then
                                    '        r.id = u.alias + r.id
                                    '        newSubform.responses.Add(r)
                                    '    End If
                                    'Next
                                    'If newSubform.responses.Count > 0 Then
                                    '    hasResponses = True
                                    '    Dim sfToFill As SubForm = (From sfp As SubForm In newPR.subforms Where sfp.id = newSubform.id).FirstOrDefault
                                    '    If sfToFill Is Nothing Then
                                    '        newPR.subforms.Add(newSubform)
                                    '    Else
                                    '        sfToFill.responses.AddRange(newSubform.responses)
                                    '    End If
                                    'End If
                                Next
                            Next
                            If hasResponses Then globalOutputStore.personResponses.Add(newPR)
                        Next
                    Next
                End Using
            Catch ex As Exception
                myworker.ReportProgress(progressValue, "Fehler " + IO.Path.GetFileName(fn) + ": " + ex.Message)
            End Try
        Next
    End Sub
End Class
