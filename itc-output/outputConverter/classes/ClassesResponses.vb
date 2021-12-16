Imports Newtonsoft.Json
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

Class UnitLineData
    Public groupname As String
    Public loginname As String
    Public code As String
    Public bookletname As String
    Public unitname As String
    Private responseChunks As Dictionary(Of String, String)
    Public responseType As String
    Public responseTimestamp As String
    Public laststate As Dictionary(Of String, String)
    Public responses As Dictionary(Of String, List(Of ResponseData))
    Public ReadOnly Property personKey As String
        Get
            Return groupname + loginname + code
        End Get
    End Property

    Public Sub New(line As String, legacyMode As Boolean, renameVariables As Dictionary(Of String, Dictionary(Of String, List(Of String))))
        responseChunks = New Dictionary(Of String, String)
        laststate = New Dictionary(Of String, String)
        Dim position As Integer = 0
        Dim semicolonActive As Boolean = True
        Dim tmpStr As String = ""
        Dim tmpResponses As String = ""
        Dim tmpLastState As String = ""
        For Each c As Char In line
            If c = ";" Then
                If semicolonActive Then
                    If Not String.IsNullOrEmpty(tmpStr) Then
                        If tmpStr.Substring(0, 1) = """" Then
                            If tmpStr.Substring(tmpStr.Length - 1, 1) = """" Then
                                tmpStr = tmpStr.Substring(1, tmpStr.Length - 2)
                            End If
                        End If
                        Select Case position
                            Case 0
                                groupname = tmpStr
                            Case 1
                                loginname = tmpStr
                            Case 2
                                code = tmpStr
                            Case 3
                                bookletname = tmpStr
                            Case 4
                                unitname = tmpStr
                            Case 5
                                If Not legacyMode Then tmpResponses = tmpStr
                            Case 6
                                If legacyMode Then
                                    tmpResponses = tmpStr
                                Else
                                    responseType = tmpStr
                                End If
                            Case 7
                                If legacyMode Then
                                    responseType = tmpStr
                                Else
                                    responseTimestamp = tmpStr
                                End If
                            Case 8
                                If legacyMode Then
                                    responseTimestamp = tmpStr
                                Else
                                    tmpLastState = tmpStr
                                End If
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
                If c = """" Then semicolonActive = Not semicolonActive
            End If
        Next
        If Not String.IsNullOrEmpty(tmpResponses) Then
            tmpResponses = tmpResponses.Replace("""""", """")
            If legacyMode Then
                responseChunks.Add("all", tmpResponses)
            Else
                Try
                    tmpResponses = tmpResponses.Replace("\\", "\\\\")
                    tmpResponses = tmpResponses.Replace("\b", "\\b")
                    tmpResponses = tmpResponses.Replace("\f", "\\f")
                    tmpResponses = tmpResponses.Replace("\r", "\\r")
                    tmpResponses = tmpResponses.Replace("\n", "\\n")
                    tmpResponses = tmpResponses.Replace("\t", "\\t")
                    responseChunks = JsonConvert.DeserializeObject(tmpResponses, GetType(Dictionary(Of String, String)))
                Catch ex As Exception
                    responseChunks.Add("all", tmpResponses)
                End Try
            End If
        End If
        If Not String.IsNullOrEmpty(tmpLastState) Then
            tmpLastState = tmpLastState.Replace("""""", """")
            Try
                laststate = JsonConvert.DeserializeObject(tmpLastState, GetType(Dictionary(Of String, String)))
            Catch ex As Exception
                laststate.Add("state", tmpLastState)
            End Try
        End If

        responses = New Dictionary(Of String, List(Of ResponseData))
        If responseChunks.Count > 0 Then
            Dim varRenameDef As Dictionary(Of String, List(Of String)) = Nothing
            If renameVariables IsNot Nothing AndAlso renameVariables.ContainsKey(unitname) Then varRenameDef = renameVariables.Item(unitname)
            For Each responseChunk As KeyValuePair(Of String, String) In responseChunks
                Dim dataToAdd As Dictionary(Of String, List(Of ResponseData)) = Nothing
                Select Case responseType
                    Case "IQBVisualUnitPlayerV2.1.0"
                        dataToAdd = setResponsesDan(responseChunk.Value, varRenameDef)
                    Case "unknown"
                        dataToAdd = setResponsesAbi(responseChunk.Value)
                    Case "iqb-simple-player@1.0.0"
                        dataToAdd = setResponsesSimplePlayerLegacy(responseChunk.Value, varRenameDef)
                    Case "iqb-aspect-player@0.1.1", "iqb-standard@1.0.0"
                        dataToAdd = setResponsesIqbStandard(responseChunk.Value)
                    Case Else
                        dataToAdd = setResponsesKeyValue(responseChunk.Value, varRenameDef)
                End Select
                If dataToAdd IsNot Nothing Then
                    For Each kvp As KeyValuePair(Of String, List(Of ResponseData)) In dataToAdd
                        If Not responses.ContainsKey(kvp.Key) Then responses.Add(kvp.Key, New List(Of ResponseData))
                        Dim respList As List(Of ResponseData) = responses.Item(kvp.Key)
                        respList.AddRange(kvp.Value)
                    Next
                End If
            Next
        End If
    End Sub

    Private Function setResponsesDan(responseString As String, varRenameDef As Dictionary(Of String, List(Of String))) As Dictionary(Of String, List(Of ResponseData))
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

    Private Function setResponsesAbi(responseString As String) As Dictionary(Of String, List(Of ResponseData))
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

    Private Function setResponsesKeyValue(responseString As String, varRenameDef As Dictionary(Of String, List(Of String))) As Dictionary(Of String, List(Of ResponseData))
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

    Private Function setResponsesSimplePlayerLegacy(responseString As String, varRenameDef As Dictionary(Of String, List(Of String))) As Dictionary(Of String, List(Of ResponseData))
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

    Private Function setResponsesIqbStandard(responseString As String) As Dictionary(Of String, List(Of ResponseData))
        Dim myreturn As New List(Of ResponseData)
        Dim localdata As New List(Of Dictionary(Of String, Linq.JToken))
        Try
            localdata = JsonConvert.DeserializeObject(responseString, GetType(List(Of Dictionary(Of String, Linq.JToken))))
        Catch ex As Exception
            myreturn.Add(New ResponseData(ResponseData.STATUS_ERROR, "Converter iqb-standard failed: " + ex.Message, ResponseData.STATUS_ERROR))
        End Try
        If localdata.Count > 0 Then
            For Each entry As Dictionary(Of String, Linq.JToken) In localdata
                If entry.ContainsKey("id") AndAlso entry.ContainsKey("value") Then
                    If entry.ContainsKey("status") Then
                        myreturn.Add(New ResponseData(entry.Item("id").ToString, entry.Item("value").ToString, entry.Item("status").ToString))
                    Else
                        myreturn.Add(New ResponseData(entry.Item("id").ToString, entry.Item("value").ToString, ResponseData.STATUS_UNSET))
                    End If
                End If
            Next
        End If
        Return New Dictionary(Of String, List(Of ResponseData)) From {{"", myreturn}}
    End Function
End Class