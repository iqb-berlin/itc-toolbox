﻿Imports Newtonsoft.Json
Class ResponseData
    Public Const STATUS_UNSET = "UNSET"
    Public Const STATUS_ERROR = "ERROR"
    Public Const STATUS_VALUE_CHANGED = "VALUE_CHANGED"
    Public variableId As String
    Public status As String
    Public value As String
    Public code As Integer
    Public score As Integer
    Public Sub New(id As String, v As String, st As String)
        variableId = id
        value = v
        status = st
        code = 0
        score = 0
    End Sub
End Class

Class ResponseChunk
    Public id As String
    Public content As String
    Public responseType As String
    Public responseTimestamp As String
End Class

Class ResponseChunkDAO
    Public id As String
    Public content As String
    Public ts As Long
    Public responseType As String
End Class

Class UnitLineData
    Public groupname As String
    Public loginname As String
    Public code As String
    Public bookletname As String
    Public unitname As String
    Public laststate As Dictionary(Of String, String)
    Public responses As Dictionary(Of String, List(Of ResponseData))
    Public ReadOnly Property personKey As String
        Get
            Return groupname + loginname + code
        End Get
    End Property
    Public ReadOnly Property hasResponses As Boolean
        Get
            Return responses IsNot Nothing AndAlso responses.Count > 0 AndAlso responses.First.Value.Count > 0
        End Get
    End Property

    Public Shared Function fromCsvLine(line As String, legacyMode As Boolean,
                                       renameVariables As Dictionary(Of String, Dictionary(Of String, List(Of String))),
                                       csvSeparator As String) As UnitLineData
        Dim returnUnitData As New UnitLineData
        Dim responseChunks As New List(Of ResponseChunk)
        returnUnitData.laststate = New Dictionary(Of String, String)
        Dim position As Integer = 0
        Dim separatorActive As Boolean = True
        Dim tmpStr As String = ""
        Dim tmpResponses As String = ""
        Dim tmpResponseType As String = ""
        Dim tmpResponseTimestamp As String = ""
        Dim tmpLastState As String = ""
        For Each c As Char In line
            If c = csvSeparator Then
                If separatorActive Then
                    If Not String.IsNullOrEmpty(tmpStr) Then
                        If tmpStr.Substring(0, 1) = """" Then
                            If tmpStr.Substring(tmpStr.Length - 1, 1) = """" Then
                                tmpStr = tmpStr.Substring(1, tmpStr.Length - 2)
                            End If
                        End If
                        Select Case position
                            Case 0
                                returnUnitData.groupname = tmpStr
                            Case 1
                                returnUnitData.loginname = tmpStr
                            Case 2
                                returnUnitData.code = tmpStr
                            Case 3
                                returnUnitData.bookletname = tmpStr
                            Case 4
                                returnUnitData.unitname = tmpStr
                            Case 5
                                If Not legacyMode Then tmpResponses = tmpStr
                            Case 6
                                If legacyMode Then
                                    tmpResponses = tmpStr
                                Else
                                    tmpLastState = tmpStr
                                End If
                            Case 7
                                If legacyMode Then tmpResponseType = tmpStr
                            Case 8
                                If legacyMode Then tmpResponseTimestamp = tmpStr
                            Case 10
                                If legacyMode Then tmpLastState = tmpStr
                        End Select
                        tmpStr = ""
                    End If
                    position += 1
                Else
                    tmpStr += c
                End If
            Else
                tmpStr += c
                If c = """" Then separatorActive = Not separatorActive
            End If
        Next
        If Not String.IsNullOrEmpty(tmpResponses) Then
            tmpResponses = tmpResponses.Replace("""""", """")
            If legacyMode Then
                responseChunks.Add(New ResponseChunk With {
                    .id = "all",
                    .content = tmpResponses,
                    .responseType = tmpResponseType,
                    .responseTimestamp = tmpResponseTimestamp
                })
            Else
                Try
                    For Each rCh As ResponseChunkDAO In JsonConvert.DeserializeObject(tmpResponses, GetType(List(Of ResponseChunkDAO)))
                        responseChunks.Add(New ResponseChunk With {
                            .id = rCh.id,
                            .content = rCh.content,
                            .responseType = rCh.responseType,
                            .responseTimestamp = rCh.ts.ToString
                        })
                    Next
                Catch ex As Exception
                    responseChunks.Add(New ResponseChunk With {
                        .id = "all-error",
                        .content = tmpResponses,
                        .responseType = tmpResponseType,
                        .responseTimestamp = tmpResponseTimestamp
                    })
                End Try
            End If
        End If
        If Not String.IsNullOrEmpty(tmpLastState) Then
            tmpLastState = tmpLastState.Replace("""""", """")
            Try
                returnUnitData.laststate = JsonConvert.DeserializeObject(tmpLastState, GetType(Dictionary(Of String, String)))
            Catch ex As Exception
                returnUnitData.laststate.Add("state", tmpLastState)
            End Try
        End If

        returnUnitData.responses = New Dictionary(Of String, List(Of ResponseData))
        If responseChunks.Count > 0 Then
            Dim varRenameDef As Dictionary(Of String, List(Of String)) = Nothing
            If renameVariables IsNot Nothing AndAlso renameVariables.ContainsKey(returnUnitData.unitname) Then varRenameDef = renameVariables.Item(returnUnitData.unitname)
            For Each responseChunk As ResponseChunk In responseChunks
                Dim dataToAdd As Dictionary(Of String, List(Of ResponseData)) = Nothing
                Select Case responseChunk.responseType
                    Case "IQBVisualUnitPlayerV2.1.0"
                        dataToAdd = setResponsesDan(responseChunk.content, varRenameDef)
                    Case "unknown"
                        dataToAdd = setResponsesAbi(responseChunk.content)
                    Case "iqb-simple-player@1.0.0"
                        dataToAdd = setResponsesSimplePlayerLegacy(responseChunk.content, varRenameDef)
                    Case "iqb-aspect-player@0.1.1", "iqb-standard@1.0.0", "iqb-standard@1.0", "iqb-standard@1.1"
                        dataToAdd = setResponsesIqbStandard(responseChunk.content)
                    Case Else
                        dataToAdd = setResponsesKeyValue(responseChunk.content, varRenameDef)
                End Select
                If dataToAdd IsNot Nothing Then
                    For Each kvp As KeyValuePair(Of String, List(Of ResponseData)) In dataToAdd
                        If Not returnUnitData.responses.ContainsKey(kvp.Key) Then returnUnitData.responses.Add(kvp.Key, New List(Of ResponseData))
                        Dim respList As List(Of ResponseData) = returnUnitData.responses.Item(kvp.Key)
                        respList.AddRange(kvp.Value)
                    Next
                End If
            Next
        End If
        Return returnUnitData
    End Function

    Public Shared Function fromTestcenterAPI(responseData As ResponseDTO) As UnitLineData
        Dim returnUnitData As New UnitLineData With {
            .groupname = responseData.groupname, .bookletname = responseData.bookletname, .code = responseData.code,
            .loginname = responseData.loginname, .unitname = responseData.unitname, .laststate = New Dictionary(Of String, String),
            .responses = New Dictionary(Of String, List(Of ResponseData))
            }
        Dim tmpLastState As String = responseData.laststate
        If Not String.IsNullOrEmpty(tmpLastState) Then
            tmpLastState = tmpLastState.Replace("""""", """")
            Try
                returnUnitData.laststate = JsonConvert.DeserializeObject(tmpLastState, GetType(Dictionary(Of String, String)))
            Catch ex As Exception
                returnUnitData.laststate.Add("state", tmpLastState)
            End Try
        End If

        If responseData.responses.Count > 0 Then
            Dim varRenameDef As Dictionary(Of String, List(Of String)) = Nothing
            For Each responseChunk As ResponseDataDTO In responseData.responses
                Dim dataToAdd As Dictionary(Of String, List(Of ResponseData)) = Nothing
                Select Case responseChunk.responseType
                    Case "IQBVisualUnitPlayerV2.1.0"
                        dataToAdd = setResponsesDan(responseChunk.content, varRenameDef)
                    Case "unknown"
                        dataToAdd = setResponsesAbi(responseChunk.content)
                    Case "iqb-simple-player@1.0.0"
                        dataToAdd = setResponsesSimplePlayerLegacy(responseChunk.content, varRenameDef)
                    Case "iqb-aspect-player@0.1.1", "iqb-standard@1.0.0", "iqb-standard@1.0", "iqb-standard@1.1"
                        dataToAdd = setResponsesIqbStandard(responseChunk.content)
                    Case Else
                        dataToAdd = setResponsesKeyValue(responseChunk.content, varRenameDef)
                End Select
                If dataToAdd IsNot Nothing Then
                    For Each kvp As KeyValuePair(Of String, List(Of ResponseData)) In dataToAdd
                        If Not returnUnitData.responses.ContainsKey(kvp.Key) Then returnUnitData.responses.Add(kvp.Key, New List(Of ResponseData))
                        Dim respList As List(Of ResponseData) = returnUnitData.responses.Item(kvp.Key)
                        respList.AddRange(kvp.Value)
                    Next
                End If
            Next
        End If
        Return returnUnitData
    End Function

    Private Shared Function setResponsesDan(responseString As String, varRenameDef As Dictionary(Of String, List(Of String))) As Dictionary(Of String, List(Of ResponseData))
        Dim myreturn As New List(Of ResponseData)
        Dim localdata As New Dictionary(Of String, Linq.JToken)
        Try
            localdata = JsonConvert.DeserializeObject(responseString, GetType(Dictionary(Of String, Linq.JToken)))
        Catch ex As Exception
            myreturn.Add(New ResponseData(ResponseData.STATUS_ERROR, "Converter Dan failed: " + ex.Message, ResponseData.STATUS_ERROR))
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
                        myreturn.Add(New ResponseData(varName, varValue, IIf(valueChanged, ResponseData.STATUS_VALUE_CHANGED, ResponseData.STATUS_UNSET)))
                    End If
                End If
            Next
            For Each radioVariable As KeyValuePair(Of String, Integer) In foundRadioButtonGroups
                myreturn.Add(New ResponseData(radioVariable.Key, radioVariable.Value.ToString,
                                              IIf(radioVariable.Value > 0, ResponseData.STATUS_VALUE_CHANGED, ResponseData.STATUS_UNSET)))
            Next
        End If

        Return New Dictionary(Of String, List(Of ResponseData)) From {{"", myreturn}}
    End Function

    Private Shared Function setResponsesAbi(responseString As String) As Dictionary(Of String, List(Of ResponseData))
        Dim myreturn As New Dictionary(Of String, List(Of ResponseData))
        Dim localdata As New Dictionary(Of String, String)
        Try
            localdata = JsonConvert.DeserializeObject(responseString, GetType(Dictionary(Of String, String)))
        Catch ex As Exception
            myreturn.Add("", New List(Of ResponseData) From {New ResponseData(ResponseData.STATUS_ERROR, "Converter Abi failed: " + ex.Message, ResponseData.STATUS_ERROR)})
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
                    myreturn.Item("").Add(New ResponseData(singleResponse.Key, singleResponse.Value, ResponseData.STATUS_UNSET))
                End If
            Next
            If testeeData.Count > 0 Then
                For Each td As KeyValuePair(Of Integer, Dictionary(Of String, String)) In testeeData
                    For Each v As KeyValuePair(Of String, String) In td.Value
                        If Not myreturn.ContainsKey(td.Key.ToString) Then myreturn.Add(td.Key.ToString, New List(Of ResponseData))
                        myreturn.Item(td.Key.ToString).Add(New ResponseData(v.Key, v.Value, ResponseData.STATUS_UNSET))
                    Next
                Next
            End If
        End If
        Return myreturn
    End Function

    Private Shared Function setResponsesKeyValue(responseString As String, varRenameDef As Dictionary(Of String, List(Of String))) As Dictionary(Of String, List(Of ResponseData))
        Dim myreturn As New List(Of ResponseData)
        Dim localdata As New Dictionary(Of String, Linq.JToken)
        Try
            localdata = JsonConvert.DeserializeObject(responseString, GetType(Dictionary(Of String, Linq.JToken)))
        Catch ex As Exception
            myreturn.Add(New ResponseData(ResponseData.STATUS_ERROR, "Converter KeyValue failed: " + ex.Message, ResponseData.STATUS_ERROR))
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
                If Not ignoreVar Then myreturn.Add(New ResponseData(varName, s.Value, ResponseData.STATUS_UNSET))
            Next
            For Each radioVariable As KeyValuePair(Of String, Integer) In foundRadioButtonGroups
                myreturn.Add(New ResponseData(radioVariable.Key, radioVariable.Value.ToString,
                                              IIf(radioVariable.Value > 0, ResponseData.STATUS_VALUE_CHANGED, ResponseData.STATUS_UNSET)))
            Next
        End If

        Return New Dictionary(Of String, List(Of ResponseData)) From {{"", myreturn}}
    End Function

    Private Shared Function setResponsesSimplePlayerLegacy(responseString As String, varRenameDef As Dictionary(Of String, List(Of String))) As Dictionary(Of String, List(Of ResponseData))
        Dim myreturn As New List(Of ResponseData)
        Dim localdata As New Dictionary(Of String, Linq.JObject)
        Try
            localdata = JsonConvert.DeserializeObject(responseString, GetType(Dictionary(Of String, Linq.JObject)))
        Catch ex As Exception
            myreturn.Add(New ResponseData(ResponseData.STATUS_ERROR, "Converter SimplePlayerLegacy failed: " + ex.Message, ResponseData.STATUS_ERROR))
        End Try
        If localdata.ContainsKey("answers") Then
            Return setResponsesKeyValue(JsonConvert.SerializeObject(localdata.Item("answers")), varRenameDef)
        Else
            Return New Dictionary(Of String, List(Of ResponseData)) From {{"", myreturn}}
        End If
    End Function

    Private Shared Function setResponsesIqbStandard(responseString As String) As Dictionary(Of String, List(Of ResponseData))
        Dim myreturn As New Dictionary(Of String, List(Of ResponseData))
        Dim localdata As New List(Of Dictionary(Of String, Linq.JToken))
        Try
            localdata = JsonConvert.DeserializeObject(responseString, GetType(List(Of Dictionary(Of String, Linq.JToken))))
        Catch ex As Exception
            myreturn.Add("", New List(Of ResponseData))
            myreturn.Item("").Add(New ResponseData(ResponseData.STATUS_ERROR, "Converter iqb-standard failed: " + ex.Message, ResponseData.STATUS_ERROR))
        End Try
        If localdata.Count > 0 Then
            For Each entry As Dictionary(Of String, Linq.JToken) In localdata
                If entry.ContainsKey("id") AndAlso entry.ContainsKey("value") Then
                    Dim myJToken As Linq.JToken = entry.Item("value")
                    Dim newValue As String = "#null#"
                    If myJToken.Type <> Linq.JTokenType.Null Then newValue = entry.Item("value").ToString.Replace(vbNewLine, "")
                    Dim subform As String = ""
                    If entry.ContainsKey("subform") Then subform = entry.Item("subform")
                    If Not myreturn.ContainsKey(subform) Then myreturn.Add(subform, New List(Of ResponseData))
                    If entry.ContainsKey("status") Then
                        myreturn.Item(subform).Add(New ResponseData(entry.Item("id").ToString, newValue, entry.Item("status").ToString))
                    Else
                        myreturn.Item(subform).Add(New ResponseData(entry.Item("id").ToString, newValue, ResponseData.STATUS_UNSET))
                    End If
                End If
            Next
        End If
        Return myreturn
    End Function
End Class