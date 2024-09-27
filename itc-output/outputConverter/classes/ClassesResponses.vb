Imports Newtonsoft.Json

Public Class ResponseSymbols
    Public Const STATUS_UNSET = "UNSET"
    Public Const STATUS_ERROR = "ERROR"
    Public Const STATUS_VALUE_CHANGED = "VALUE_CHANGED"
End Class

Public Class ResponseData
    Public ReadOnly id As String
    Public ReadOnly status As String
    Public ReadOnly value As String
    Public Sub New(id As String, v As String, st As String)
        Me.id = id
        value = v
        status = st
    End Sub
End Class

Public Class SingleFormResponseData
    Public subformId As String
    Public responses As List(Of ResponseData)
End Class

Class ResponseChunk
    Public id As String
    Public content As String
    Public type As String
    Public ts As String
End Class
Public Class ResponseChunkData
    Public id As String
    Public type As String
    Public ts As String
    Public variables As List(Of String)
End Class

Class ResponseChunkDAO
    Public id As String
    Public content As String
    Public ts As Long
    Public responseType As String
End Class

Public Class LastStateEntry
    Public key As String
    Public value As String
End Class

Public Class UnitLineData
    Public groupname As String
    Public loginname As String
    Public code As String
    Public bookletname As String
    Public unitname As String
    Public laststate As List(Of LastStateEntry)
    Public responses As List(Of SingleFormResponseData)
    Public responseChunks As List(Of ResponseChunkData)

    Public Shared Function fromCsvLine(line As String,
                                       renameVariables As Dictionary(Of String, Dictionary(Of String, List(Of String))),
                                       csvSeparator As String, replaceBigdata As Boolean) As UnitLineData
        Dim returnUnitData As New UnitLineData
        Dim responseChunks As New List(Of ResponseChunk)
        returnUnitData.laststate = New List(Of LastStateEntry)

        Dim separator As String = """" + csvSeparator + """"
        Dim lineSplits As String() = Text.RegularExpressions.Regex.Split(line, separator)
        If lineSplits.Length > 5 Then
            returnUnitData.groupname = lineSplits(0).Substring(1)
            returnUnitData.loginname = lineSplits(1)
            returnUnitData.code = lineSplits(2)
            returnUnitData.bookletname = lineSplits(3)
            returnUnitData.unitname = lineSplits(4)
            Dim startPos As Integer = lineSplits(0).Length + lineSplits(1).Length + lineSplits(2).Length + lineSplits(3).Length + lineSplits(4).Length
            Dim residual As String = line.Substring(startPos + 5 * separator.Length)
            Dim stateStartPos As Integer = residual.LastIndexOf(separator)
            Dim dataPartsString As String = ""
            If stateStartPos > 0 Then
                Dim lastStateString As String = residual.Substring(stateStartPos + separator.Length)
                lastStateString = lastStateString.Substring(0, lastStateString.Length - 1).Replace("""""", """")
                If Not String.IsNullOrEmpty(lastStateString) Then
                    Try
                        Dim stateDict As Dictionary(Of String, String) = JsonConvert.DeserializeObject(lastStateString, GetType(Dictionary(Of String, String)))
                        For Each state As KeyValuePair(Of String, String) In stateDict
                            returnUnitData.laststate.Add(New LastStateEntry With {.key = state.Key, .value = state.Value})
                        Next
                    Catch ex As Exception
                        returnUnitData.laststate.Add(New LastStateEntry With {.key = "state", .value = lastStateString})
                    End Try
                End If
                dataPartsString = residual.Substring(0, stateStartPos)
            Else
                dataPartsString = residual
            End If

            dataPartsString = dataPartsString.Replace("""""", """")
            Try
                For Each rCh As ResponseChunkDAO In JsonConvert.DeserializeObject(dataPartsString, GetType(List(Of ResponseChunkDAO)))
                    responseChunks.Add(New ResponseChunk With {
                        .id = rCh.id,
                        .content = rCh.content,
                        .type = rCh.responseType,
                        .ts = rCh.ts.ToString
                    })
                Next
            Catch ex As Exception
                responseChunks.Add(New ResponseChunk With {
                        .id = "all-error",
                        .content = dataPartsString,
                        .type = "?",
                        .ts = "?"
                    })
            End Try
        End If

        returnUnitData.responses = New List(Of SingleFormResponseData)
        returnUnitData.responseChunks = New List(Of ResponseChunkData)
        If responseChunks.Count > 0 Then
            Dim varRenameDef As Dictionary(Of String, List(Of String)) = Nothing
            If renameVariables IsNot Nothing AndAlso renameVariables.ContainsKey(returnUnitData.unitname) Then varRenameDef = renameVariables.Item(returnUnitData.unitname)
            For Each responseChunk As ResponseChunk In responseChunks
                Dim dataToAdd As List(Of SingleFormResponseData) = Nothing
                Select Case responseChunk.type
                    Case "IQBVisualUnitPlayerV2.1.0"
                        dataToAdd = setResponsesDan(responseChunk.content, varRenameDef)
                    Case "unknown"
                        dataToAdd = setResponsesAbi(responseChunk.content)
                    Case "iqb-simple-player@1.0.0"
                        dataToAdd = setResponsesSimplePlayerLegacy(responseChunk.content, varRenameDef)
                    Case "iqb-aspect-player@0.1.1", "iqb-standard@1.0.0", "iqb-standard@1.0", "iqb-standard@1.1"
                        dataToAdd = setResponsesIqbStandard(responseChunk.content, replaceBigdata)
                    Case Else
                        dataToAdd = setResponsesKeyValue(responseChunk.content, varRenameDef)
                End Select
                If dataToAdd IsNot Nothing AndAlso dataToAdd.Count > 0 Then
                    Dim newChunk = New ResponseChunkData() With {.id = responseChunk.id, .ts = responseChunk.ts,
                        .type = responseChunk.type, .variables = New List(Of String)}
                    returnUnitData.responses.AddRange(dataToAdd)
                    For Each kvp As SingleFormResponseData In dataToAdd
                        newChunk.variables.AddRange(From v In kvp.responses Select v.id)
                    Next
                    returnUnitData.responseChunks.Add(newChunk)
                End If
            Next
        End If
        Return returnUnitData
    End Function

    Public Shared Function fromTestcenterAPI(responseData As ResponseDTO, replaceBigdata As Boolean) As UnitLineData
        Dim returnUnitData As New UnitLineData With {
            .groupname = responseData.groupname, .bookletname = responseData.bookletname, .code = responseData.code,
            .loginname = responseData.loginname, .unitname = responseData.unitname, .laststate = New List(Of LastStateEntry),
            .responses = New List(Of SingleFormResponseData), .responseChunks = New List(Of ResponseChunkData)
            }
        Dim tmpLastState As String = responseData.laststate
        If Not String.IsNullOrEmpty(tmpLastState) Then
            tmpLastState = tmpLastState.Replace("""""", """")
            Try
                Dim stateDict As Dictionary(Of String, String) = JsonConvert.DeserializeObject(tmpLastState, GetType(Dictionary(Of String, String)))
                For Each state As KeyValuePair(Of String, String) In stateDict
                    returnUnitData.laststate.Add(New LastStateEntry With {.key = state.Key, .value = state.Value})
                Next
            Catch ex As Exception
                returnUnitData.laststate.Add(New LastStateEntry With {.key = "state", .value = tmpLastState})
            End Try
        End If

        If responseData.responses.Count > 0 Then
            Dim varRenameDef As Dictionary(Of String, List(Of String)) = Nothing
            For Each responseChunk As ResponseDataDTO In responseData.responses
                Dim dataToAdd As List(Of SingleFormResponseData) = Nothing
                Select Case responseChunk.responseType
                    Case "IQBVisualUnitPlayerV2.1.0"
                        dataToAdd = setResponsesDan(responseChunk.content, varRenameDef)
                    Case "unknown"
                        dataToAdd = setResponsesAbi(responseChunk.content)
                    Case "iqb-simple-player@1.0.0"
                        dataToAdd = setResponsesSimplePlayerLegacy(responseChunk.content, varRenameDef)
                    Case "iqb-aspect-player@0.1.1", "iqb-standard@1.0.0", "iqb-standard@1.0", "iqb-standard@1.1"
                        dataToAdd = setResponsesIqbStandard(responseChunk.content, replaceBigdata)
                    Case Else
                        dataToAdd = setResponsesKeyValue(responseChunk.content, varRenameDef)
                End Select
                If dataToAdd IsNot Nothing Then
                    Dim newChunk = New ResponseChunkData() With {.id = responseChunk.id, .ts = responseChunk.ts,
                        .type = responseChunk.responseType, .variables = New List(Of String)}
                    returnUnitData.responses.AddRange(dataToAdd)
                    For Each kvp As SingleFormResponseData In dataToAdd
                        newChunk.variables.AddRange(From v In kvp.responses Select v.id)
                    Next
                    returnUnitData.responseChunks.Add(newChunk)
                End If
            Next
        End If
        Return returnUnitData
    End Function

    Private Shared Function setResponsesDan(responseString As String, varRenameDef As Dictionary(Of String, List(Of String))) As List(Of SingleFormResponseData)
        Dim myreturn As New SingleFormResponseData With {.subformId = "", .responses = New List(Of ResponseData)}
        Dim localdata As New Dictionary(Of String, Linq.JToken)
        Try
            localdata = JsonConvert.DeserializeObject(responseString, GetType(Dictionary(Of String, Linq.JToken)))
        Catch ex As Exception
            myreturn.responses.Add(New ResponseData(ResponseSymbols.STATUS_ERROR, "Converter Dan failed: " + ex.Message, ResponseSymbols.STATUS_ERROR))
        End Try
        If localdata.Count > 0 Then
            Dim foundRadioButtonGroups As New Dictionary(Of String, Integer)
            Dim valuesChanged As New Dictionary(Of String, Boolean)
            If localdata.ContainsKey("responsesGiven") Then valuesChanged = localdata.Item("responsesGiven").ToObject(Of Dictionary(Of String, Boolean))
            For Each s As KeyValuePair(Of String, Linq.JToken) In localdata
                If s.Key <> "responsesGiven" AndAlso s.Key <> "pagesViewed" Then
                    Dim varName As String = s.Key
                    Dim varValue As String = s.Value.ToString
                    Dim ignoreVar As Boolean = False
                    Dim valueChanged As Boolean = False
                    If varRenameDef IsNot Nothing Then
                        For Each varNameDef As KeyValuePair(Of String, List(Of String)) In varRenameDef
                            If varNameDef.Value.Contains(varName) Then
                                If varNameDef.Key = "__omit__" Then
                                    ignoreVar = True
                                Else
                                    If varNameDef.Value.Count > 1 Then
                                        If Not foundRadioButtonGroups.ContainsKey(varNameDef.Key) Then foundRadioButtonGroups.Add(varNameDef.Key, 0)
                                        If varValue = "true" Then
                                            foundRadioButtonGroups.Item(varNameDef.Key) = varNameDef.Value.IndexOf(varName) + 1
                                        End If
                                        ignoreVar = True
                                    Else
                                        If valuesChanged IsNot Nothing AndAlso valuesChanged.ContainsKey(varName) Then valueChanged = valuesChanged.Item(varName)
                                        varName = varNameDef.Key
                                    End If
                                End If
                                Exit For
                            End If
                        Next
                    End If
                    If Not ignoreVar Then
                        myreturn.responses.Add(New ResponseData(varName, varValue, IIf(valueChanged, ResponseSymbols.STATUS_VALUE_CHANGED, ResponseSymbols.STATUS_UNSET)))
                    End If
                End If
            Next
            For Each radioVariable As KeyValuePair(Of String, Integer) In foundRadioButtonGroups
                myreturn.responses.Add(New ResponseData(radioVariable.Key, radioVariable.Value.ToString,
                                              IIf(radioVariable.Value > 0, ResponseSymbols.STATUS_VALUE_CHANGED, ResponseSymbols.STATUS_UNSET)))
            Next
        End If

        Return New List(Of SingleFormResponseData) From {[myreturn]}
    End Function

    Private Shared Function setResponsesAbi(responseString As String) As List(Of SingleFormResponseData)
        Dim myreturn As New Dictionary(Of String, List(Of ResponseData))
        Dim localdata As New Dictionary(Of String, String)
        Try
            localdata = JsonConvert.DeserializeObject(responseString, GetType(Dictionary(Of String, String)))
        Catch ex As Exception
            myreturn.Add("", New List(Of ResponseData) From {New ResponseData(ResponseSymbols.STATUS_ERROR, "Converter Abi failed: " + ex.Message, ResponseSymbols.STATUS_ERROR)})
        End Try
        If localdata.Count > 0 Then
            Dim testeeData As New Dictionary(Of Integer, Dictionary(Of String, String))
            For Each singleResponse As KeyValuePair(Of String, String) In localdata
                'find out person
                Dim pIndex As Integer = 0
                Dim pPos As Integer = singleResponse.Key.LastIndexOf("_")
                If pPos > 1 AndAlso Integer.TryParse(singleResponse.Key.Substring(pPos + 1), pIndex) Then
                    If pIndex > 0 Then
                        If Not testeeData.ContainsKey(pIndex) Then testeeData.Add(pIndex, New Dictionary(Of String, String))
                        Dim varname As String = singleResponse.Key.Substring(0, pPos)
                        If Not testeeData.Item(pIndex).ContainsKey(varname) Then testeeData.Item(pIndex).Add(varname, singleResponse.Value)
                    End If
                End If
                If pIndex <= 0 Then
                    If Not myreturn.ContainsKey("") Then myreturn.Add("", New List(Of ResponseData))
                    myreturn.Item("").Add(New ResponseData(singleResponse.Key, singleResponse.Value, ResponseSymbols.STATUS_UNSET))
                End If
            Next
            If testeeData.Count > 0 Then
                For Each td As KeyValuePair(Of Integer, Dictionary(Of String, String)) In testeeData
                    For Each v As KeyValuePair(Of String, String) In td.Value
                        If Not myreturn.ContainsKey(td.Key.ToString) Then myreturn.Add(td.Key.ToString, New List(Of ResponseData))
                        myreturn.Item(td.Key.ToString).Add(New ResponseData(v.Key, v.Value, ResponseSymbols.STATUS_UNSET))
                    Next
                Next
            End If
        End If
        Dim returnList As New List(Of SingleFormResponseData)
        For Each r As KeyValuePair(Of String, List(Of ResponseData)) In myreturn
            returnList.Add(New SingleFormResponseData With {.subformId = r.Key, .responses = r.Value})
        Next
        Return returnList
    End Function

    Private Shared Function setResponsesKeyValue(responseString As String, varRenameDef As Dictionary(Of String, List(Of String))) As List(Of SingleFormResponseData)
        Dim myreturn As New SingleFormResponseData With {.subformId = "", .responses = New List(Of ResponseData)}
        Dim localdata As New Dictionary(Of String, Linq.JToken)
        Try
            localdata = JsonConvert.DeserializeObject(responseString, GetType(Dictionary(Of String, Linq.JToken)))
        Catch ex As Exception
            myreturn.responses.Add(New ResponseData(ResponseSymbols.STATUS_ERROR, "Converter KeyValue failed: " + ex.Message, ResponseSymbols.STATUS_ERROR))
        End Try
        If localdata.Count > 0 Then
            Dim foundRadioButtonGroups As New Dictionary(Of String, Integer)
            For Each s As KeyValuePair(Of String, Linq.JToken) In localdata
                Dim varName As String = s.Key
                Dim varValue As String = s.Value.ToString
                Dim ignoreVar As Boolean = False
                If varRenameDef IsNot Nothing Then
                    For Each varNameDef As KeyValuePair(Of String, List(Of String)) In varRenameDef
                        If varNameDef.Value.Contains(varName) Then
                            If varNameDef.Key = "__omit__" Then
                                ignoreVar = True
                            Else
                                If varNameDef.Value.Count > 1 Then
                                    If Not foundRadioButtonGroups.ContainsKey(varNameDef.Key) Then foundRadioButtonGroups.Add(varNameDef.Key, 0)
                                    If varValue.Length > 0 AndAlso varValue.Trim.Substring(0, 1).ToUpper() = "T" Then
                                        foundRadioButtonGroups.Item(varNameDef.Key) = varNameDef.Value.IndexOf(varName) + 1
                                    End If
                                    ignoreVar = True
                                Else
                                    varName = varNameDef.Key
                                End If
                            End If
                            Exit For
                        End If
                    Next
                End If
                If Not ignoreVar Then myreturn.responses.Add(New ResponseData(varName, s.Value, ResponseSymbols.STATUS_UNSET))
            Next
            For Each radioVariable As KeyValuePair(Of String, Integer) In foundRadioButtonGroups
                myreturn.responses.Add(New ResponseData(radioVariable.Key, radioVariable.Value.ToString,
                                              IIf(radioVariable.Value > 0, ResponseSymbols.STATUS_VALUE_CHANGED, ResponseSymbols.STATUS_UNSET)))
            Next
        End If

        Return New List(Of SingleFormResponseData) From {[myreturn]}
    End Function

    Private Shared Function setResponsesSimplePlayerLegacy(responseString As String, varRenameDef As Dictionary(Of String, List(Of String))) As List(Of SingleFormResponseData)
        Dim myreturn As New SingleFormResponseData With {.subformId = "", .responses = New List(Of ResponseData)}
        Dim localdata As New Dictionary(Of String, Linq.JObject)
        Try
            localdata = JsonConvert.DeserializeObject(responseString, GetType(Dictionary(Of String, Linq.JObject)))
        Catch ex As Exception
            myreturn.responses.Add(New ResponseData(ResponseSymbols.STATUS_ERROR, "Converter SimplePlayerLegacy failed: " + ex.Message, ResponseSymbols.STATUS_ERROR))
        End Try
        If localdata.ContainsKey("answers") Then
            Return setResponsesKeyValue(JsonConvert.SerializeObject(localdata.Item("answers")), varRenameDef)
        Else
            Return New List(Of SingleFormResponseData) From {[myreturn]}
        End If
    End Function

    Private Shared Function setResponsesIqbStandard(responseString As String, replaceBigdata As Boolean) As List(Of SingleFormResponseData)
        Dim myreturn As New Dictionary(Of String, List(Of ResponseData))
        Dim localdata As New List(Of Dictionary(Of String, Linq.JToken))
        Dim conversionErrorMessage As String = ""
        Try
            localdata = JsonConvert.DeserializeObject(responseString, GetType(List(Of Dictionary(Of String, Linq.JToken))))
        Catch ex As Exception
            conversionErrorMessage = "Converter iqb-standard failed: " + ex.Message
        End Try
        If Not String.IsNullOrEmpty(conversionErrorMessage) Then
            Try
                localdata.Add(JsonConvert.DeserializeObject(responseString, GetType(Dictionary(Of String, Linq.JToken))))
            Catch ex As Exception
                myreturn.Add("", New List(Of ResponseData))
                myreturn.Item("").Add(New ResponseData(ResponseSymbols.STATUS_ERROR, conversionErrorMessage, ResponseSymbols.STATUS_ERROR))
            End Try
        End If
        If localdata.Count > 0 Then
            For Each entry As Dictionary(Of String, Linq.JToken) In localdata
                If entry.ContainsKey("id") AndAlso entry.ContainsKey("value") Then
                    Dim myJToken As Linq.JToken = entry.Item("value")
                    Dim newValue As String = "#null#"
                    If myJToken.Type <> Linq.JTokenType.Null Then
                        newValue = entry.Item("value").ToString
                        Const bigDataMarker = "data:application/octet-stream;base64"
                        If newValue.IndexOf(bigDataMarker) = 0 AndAlso replaceBigdata Then
                            newValue = bigDataMarker + " - hash: " + newValue.GetHashCode().ToString
                        Else
                            newValue = newValue.Replace(vbNewLine, "")
                        End If
                    End If
                    Dim subform As String = ""
                    If entry.ContainsKey("subform") Then subform = entry.Item("subform")
                    If Not myreturn.ContainsKey(subform) Then myreturn.Add(subform, New List(Of ResponseData))
                    If entry.ContainsKey("status") Then
                        myreturn.Item(subform).Add(New ResponseData(entry.Item("id").ToString, newValue, entry.Item("status").ToString))
                    Else
                        myreturn.Item(subform).Add(New ResponseData(entry.Item("id").ToString, newValue, ResponseSymbols.STATUS_UNSET))
                    End If
                End If
            Next
        End If
        Dim returnList As New List(Of SingleFormResponseData)
        For Each r As KeyValuePair(Of String, List(Of ResponseData)) In myreturn
            returnList.Add(New SingleFormResponseData With {.subformId = r.Key, .responses = r.Value})
        Next
        Return returnList
    End Function
End Class