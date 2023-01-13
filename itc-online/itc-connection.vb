Imports Newtonsoft.Json
Public Class ITCConnection
    Public selectedWorkspace As Integer = 0
    Private _url As String
    Public ReadOnly Property url() As String
        Get
            Return _url
        End Get
    End Property
    Private _lastErrorMsgText As String
    Private tokenStr As String
    Public accessTo As New Dictionary(Of Integer, String)
    Public ReadOnly Property lastErrorMsgText() As String
        Get
            Return _lastErrorMsgText
        End Get
    End Property
    Private _response_string As String
    Public ReadOnly Property response_string() As String
        Get
            Return _response_string
        End Get
    End Property
    Public Sub New(url As String, credentials As Net.NetworkCredential, worker As ComponentModel.BackgroundWorker)
        Me._url = url
        Dim resp As Net.WebResponse = Nothing
        If worker IsNot Nothing Then worker.ReportProgress(10.0#)
        Try
            Dim uri As New Uri(url + "/session/admin")
            Dim requ As Net.WebRequest = Net.WebRequest.Create(uri)
            requ.Method = "PUT"
            requ.ContentType = "application/json"
            Dim enc As New Text.UTF8Encoding
            Dim dataBin As Byte() = enc.GetBytes("{""name"":""" + credentials.UserName + """,""password"": """ + credentials.Password + """}")
            requ.ContentLength = dataBin.Length
            Dim s As IO.Stream = requ.GetRequestStream()
            s.Write(dataBin, 0, dataBin.Length)
            s.Close()

            resp = requ.GetResponse
        Catch ex As Exception
            resp = Nothing
            _lastErrorMsgText = ex.Message
            If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
        End Try
        If resp IsNot Nothing Then
            If worker IsNot Nothing Then worker.ReportProgress(20.0#)
            Using WebReader As New System.IO.StreamReader(resp.GetResponseStream(), Text.Encoding.UTF8)
                Try
                    _response_string = WebReader.ReadToEnd()
                    Dim localdata As Dictionary(Of String, Linq.JToken) = JsonConvert.DeserializeObject(_response_string, GetType(Dictionary(Of String, Linq.JToken)))
                    Me.tokenStr = localdata.Item("token").ToObject(Of String)
                    Dim tmpAccessTo As Dictionary(Of String, List(Of String)) = localdata.Item("access").ToObject(Of Dictionary(Of String, List(Of String)))
                    Me.accessTo = tmpAccessTo.Item("workspaceAdmin").ToDictionary(Function(a) Integer.Parse(a), Function(a) "Workspace " + a)
                    Me._lastErrorMsgText = ""
                Catch ex As Exception
                    _lastErrorMsgText = ex.Message
                    If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
                End Try
            End Using
            Dim maxProgressValue As Integer = accessTo.Count
            Dim progressValue As Integer = 0
            Dim wsIdList As New List(Of Integer)(accessTo.Keys)
            For Each workspaceId As Integer In wsIdList
                progressValue += 1
                If worker IsNot Nothing Then worker.ReportProgress(progressValue * 80 / maxProgressValue + 20.0#)
                Me.accessTo.Item(workspaceId) = GetWorkspaceName(workspaceId)
            Next
        End If
    End Sub

    Private Function GetWorkspaceName(wsId As Integer) As String
        Dim myReturn As String = wsId.ToString
        Dim resp As Net.WebResponse
        Try
            Dim uri As New Uri(Me._url + "/workspace/" + wsId.ToString)
            Dim requ As Net.WebRequest = Net.WebRequest.Create(uri)
            requ.Method = "GET"
            requ.ContentType = "application/json"
            requ.Headers.Item("AuthToken") = Me.tokenStr
            resp = requ.GetResponse
        Catch ex As Exception
            resp = Nothing
            _lastErrorMsgText = ex.Message
            If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
        End Try
        If resp IsNot Nothing Then
            Using WebReader As New System.IO.StreamReader(resp.GetResponseStream(), Text.Encoding.UTF8)
                Try
                    _response_string = WebReader.ReadToEnd()
                    Dim localdata As Dictionary(Of String, String) = JsonConvert.DeserializeObject(_response_string, GetType(Dictionary(Of String, String)))
                    myReturn = localdata.Item("name")
                    Me._lastErrorMsgText = ""
                Catch ex As Exception
                    _lastErrorMsgText = ex.Message
                    If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
                End Try
            End Using
        End If
        Return myReturn
    End Function

    Public Function getDataGroups() As List(Of GroupDataDTO)
        Dim myReturn As New List(Of GroupDataDTO)
        Dim resp As Net.WebResponse
        _lastErrorMsgText = ""
        Try
            Dim uri As New Uri(Me._url + "/workspace/" + selectedWorkspace.ToString + "/results")
            Dim requ As Net.WebRequest = Net.WebRequest.Create(uri)
            requ.Method = "GET"
            requ.ContentType = "application/json"
            requ.Headers.Item("AuthToken") = Me.tokenStr
            resp = requ.GetResponse
        Catch ex As Exception
            resp = Nothing
            _lastErrorMsgText = ex.Message
            If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
        End Try
        If resp IsNot Nothing Then
            Using WebReader As New System.IO.StreamReader(resp.GetResponseStream(), Text.Encoding.UTF8)
                Try
                    _response_string = WebReader.ReadToEnd()
                    myReturn = JsonConvert.DeserializeObject(_response_string, GetType(List(Of GroupDataDTO)))
                    Me._lastErrorMsgText = ""
                Catch ex As Exception
                    _lastErrorMsgText = ex.Message
                    If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
                End Try
            End Using
        End If
        Return myReturn
    End Function

    Public Function getBooklets() As List(Of BookletDTO)
        Dim myReturn As New List(Of BookletDTO)
        Dim resp As Net.WebResponse
        _lastErrorMsgText = ""
        Try
            Dim uri As New Uri(Me._url + "/workspace/" + selectedWorkspace.ToString + "/files")
            Dim requ As Net.WebRequest = Net.WebRequest.Create(uri)
            requ.Method = "GET"
            requ.ContentType = "application/json"
            requ.Headers.Item("AuthToken") = Me.tokenStr
            resp = requ.GetResponse
        Catch ex As Exception
            resp = Nothing
            _lastErrorMsgText = ex.Message
            If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
        End Try
        If resp IsNot Nothing Then
            Using WebReader As New System.IO.StreamReader(resp.GetResponseStream(), Text.Encoding.UTF8)
                Try
                    _response_string = WebReader.ReadToEnd()
                    Dim allFiles As Dictionary(Of String, Linq.JToken) = JsonConvert.DeserializeObject(_response_string, GetType(Dictionary(Of String, Linq.JToken)))
                    If allFiles.ContainsKey("Booklet") Then myReturn = allFiles.Item("Booklet").ToObject(Of List(Of BookletDTO))
                    Me._lastErrorMsgText = ""
                Catch ex As Exception
                    _lastErrorMsgText = ex.Message
                    If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
                End Try
            End Using
        End If
        Return myReturn
    End Function

    Public Function getLogs(dataGroupId As String) As List(Of LogEntryDTO)
        Dim myReturn As New List(Of LogEntryDTO)
        Dim resp As Net.WebResponse
        _lastErrorMsgText = ""
        Try
            Dim uri As New Uri(Me._url + "/workspace/" + selectedWorkspace.ToString + "/report/log?dataIds=" + dataGroupId)
            Dim requ As Net.WebRequest = Net.WebRequest.Create(uri)
            requ.Method = "GET"
            requ.ContentType = "application/json"
            requ.Headers.Item("AuthToken") = Me.tokenStr
            resp = requ.GetResponse
        Catch ex As Exception
            resp = Nothing
            _lastErrorMsgText = ex.Message
            If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
        End Try
        If resp IsNot Nothing Then
            Using WebReader As New System.IO.StreamReader(resp.GetResponseStream(), Text.Encoding.UTF8)
                Try
                    _response_string = WebReader.ReadToEnd()
                    myReturn = JsonConvert.DeserializeObject(_response_string, GetType(List(Of LogEntryDTO)))
                    Me._lastErrorMsgText = ""
                Catch ex As Exception
                    _lastErrorMsgText = ex.Message
                    If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
                End Try
            End Using
        End If
        Return myReturn
    End Function

    Public Function getResponses(dataGroupId As String) As List(Of ResponseDTO)
        Dim myReturn As New List(Of ResponseDTO)
        Dim resp As Net.WebResponse
        _lastErrorMsgText = ""
        Try
            Dim uri As New Uri(Me._url + "/workspace/" + selectedWorkspace.ToString + "/report/response?dataIds=" + dataGroupId)
            Dim requ As Net.WebRequest = Net.WebRequest.Create(uri)
            requ.Method = "GET"
            requ.ContentType = "application/json"
            requ.Headers.Item("AuthToken") = Me.tokenStr
            resp = requ.GetResponse
        Catch ex As Exception
            resp = Nothing
            _lastErrorMsgText = ex.Message
            If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
        End Try
        If resp IsNot Nothing Then
            Using WebReader As New System.IO.StreamReader(resp.GetResponseStream(), Text.Encoding.UTF8)
                Try
                    _response_string = WebReader.ReadToEnd()
                    myReturn = JsonConvert.DeserializeObject(_response_string, GetType(List(Of ResponseDTO)))
                    Me._lastErrorMsgText = ""
                Catch ex As Exception
                    _lastErrorMsgText = ex.Message
                    If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
                End Try
            End Using
        End If
        Return myReturn
    End Function

    Public Function getReviews(dataGroupId As String) As List(Of ReviewDTO)
        Dim myReturn As New List(Of ReviewDTO)
        Dim resp As Net.WebResponse
        _lastErrorMsgText = ""
        Try
            Dim uri As New Uri(Me._url + "/workspace/" + selectedWorkspace.ToString + "/report/review?dataIds=" + dataGroupId)
            Dim requ As Net.WebRequest = Net.WebRequest.Create(uri)
            requ.Method = "GET"
            requ.ContentType = "application/json"
            requ.Headers.Item("AuthToken") = Me.tokenStr
            resp = requ.GetResponse
        Catch ex As Exception
            resp = Nothing
            _lastErrorMsgText = ex.Message
            If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
        End Try
        If resp IsNot Nothing Then
            Using WebReader As New System.IO.StreamReader(resp.GetResponseStream(), Text.Encoding.UTF8)
                Try
                    _response_string = WebReader.ReadToEnd()
                    If Not String.IsNullOrWhiteSpace(_response_string) Then
                        Dim tmpDictList As List(Of Dictionary(Of String, String)) = JsonConvert.DeserializeObject(_response_string, GetType(List(Of Dictionary(Of String, String))))
                        For Each rev As Dictionary(Of String, String) In tmpDictList
                            Dim newReview As New ReviewDTO With {
                            .bookletname = rev.Item("bookletname"),
                            .code = rev.Item("code"),
                            .entry = rev.Item("entry"),
                            .groupname = rev.Item("groupname"),
                            .loginname = rev.Item("loginname"),
                            .priority = rev.Item("priority"),
                            .reviewTime = rev.Item("reviewtime"),
                            .unitname = rev.Item("unitname")
                            }
                            newReview.categoryContent = rev.Item("category: content") IsNot Nothing
                            newReview.categoryDesign = rev.Item("category: design") IsNot Nothing
                            newReview.categoryTech = rev.Item("category: tech") IsNot Nothing
                            myReturn.Add(newReview)
                        Next
                    End If
                    Me._lastErrorMsgText = ""
                Catch ex As Exception
                    _lastErrorMsgText = ex.Message
                    If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
                End Try
            End Using
        End If
        Return myReturn
    End Function
End Class
