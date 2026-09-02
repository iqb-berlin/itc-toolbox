Imports System.ComponentModel
Imports System.Net
Imports Newtonsoft.Json

Public Class ITCConnection
    Private _BufferSize As Integer = 4096
    Public Property BufferSize() As Integer
        Get
            Return _BufferSize
        End Get
        Set(ByVal value As Integer)
            _BufferSize = value
        End Set
    End Property
    Public Shared ReadOnly validBookletDependencies() As String = {"containsUnit", "usesPlayer", "isDefinedBy"}
    Public selectedWorkspace As Integer = 0
    Public selectedWorkspaceName As String
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
        ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(
            AddressOf AcceptAllCertifications)
        Dim resp As Net.WebResponse = Nothing
        If worker IsNot Nothing Then worker.ReportProgress(10.0#)
        Dim loginBody As String = JsonConvert.SerializeObject(New Dictionary(Of String, String) From {
            {"name", credentials.UserName}, {"password", credentials.Password}})
        Try
            resp = SendJsonRequest(url + "/session/admin", "PUT", loginBody)
        Catch ex As Net.WebException
            resp = Nothing
            Dim errorBody As String = ReadErrorBody(ex)
            Dim statusCode As Net.HttpStatusCode = 0
            If TypeOf ex.Response Is Net.HttpWebResponse Then statusCode = CType(ex.Response, Net.HttpWebResponse).StatusCode
            If statusCode = Net.HttpStatusCode.BadRequest AndAlso
                    (errorBody.Contains("Challenge") OrElse errorBody.Contains("Brute Force")) Then
                ' Testcenter ab Version 18.2 mit aktivem Brute-Force-Schutz: Anmeldung erfordert Altcha-Challenge
                Try
                    resp = CreateSessionWithChallenge(credentials, worker)
                Catch ex2 As Net.WebException
                    resp = Nothing
                    _lastErrorMsgText = ex2.Message
                    Dim errorBody2 As String = ReadErrorBody(ex2)
                    If Not String.IsNullOrEmpty(errorBody2) Then _lastErrorMsgText += " / " + errorBody2
                Catch ex2 As Exception
                    resp = Nothing
                    _lastErrorMsgText = ex2.Message
                    If ex2.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex2.InnerException.Message
                End Try
            Else
                _lastErrorMsgText = ex.Message
                If Not String.IsNullOrEmpty(errorBody) Then _lastErrorMsgText += " / " + errorBody
            End If
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

    Public Shared Function AcceptAllCertifications(sender As Object,
                                                   certification As System.Security.Cryptography.X509Certificates.X509Certificate,
                                                   chain As System.Security.Cryptography.X509Certificates.X509Chain,
                                                   sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
        Return True
    End Function

    Private Function SendJsonRequest(requestUrl As String, method As String, jsonBody As String) As Net.WebResponse
        Dim requ As Net.WebRequest = Net.WebRequest.Create(New Uri(requestUrl))
        requ.Method = method
        requ.ContentType = "application/json"
        Dim enc As New Text.UTF8Encoding
        Dim dataBin As Byte() = enc.GetBytes(jsonBody)
        requ.ContentLength = dataBin.Length
        Using s As IO.Stream = requ.GetRequestStream()
            s.Write(dataBin, 0, dataBin.Length)
        End Using
        Return requ.GetResponse
    End Function

    Private Shared Function ReadErrorBody(ex As Net.WebException) As String
        If ex.Response Is Nothing Then Return ""
        Try
            Using rdr As New IO.StreamReader(ex.Response.GetResponseStream())
                Return rdr.ReadToEnd()
            End Using
        Catch
            Return ""
        End Try
    End Function

    Private Function CreateSessionWithChallenge(credentials As Net.NetworkCredential, worker As ComponentModel.BackgroundWorker) As Net.WebResponse
        If worker IsNot Nothing Then worker.ReportProgress(12.0#)
        Dim challengeRequestBody As String = JsonConvert.SerializeObject(New Dictionary(Of String, String) From {
            {"loginType", "admin"}, {"name", credentials.UserName}, {"password", credentials.Password}})
        Dim challengeJson As String
        Using challengeResp As Net.WebResponse = SendJsonRequest(Me._url + "/session/challenge", "POST", challengeRequestBody)
            Using rdr As New IO.StreamReader(challengeResp.GetResponseStream(), Text.Encoding.UTF8)
                challengeJson = rdr.ReadToEnd()
            End Using
        End Using
        Dim challenge As Dictionary(Of String, Linq.JToken) = JsonConvert.DeserializeObject(challengeJson, GetType(Dictionary(Of String, Linq.JToken)))
        Dim salt As String = challenge.Item("salt").ToObject(Of String)
        Dim targetHash As String = challenge.Item("challenge").ToObject(Of String)
        Dim maxNumber As Long = challenge.Item("maxNumber").ToObject(Of Long)
        If worker IsNot Nothing Then worker.ReportProgress(15.0#)
        Dim number As Long = SolveAltchaChallenge(salt, targetHash, maxNumber)
        If number < 0 Then Throw New Exception("Brute-Force-Schutz: Die Challenge des Servers konnte nicht gelöst werden.")
        If worker IsNot Nothing Then worker.ReportProgress(18.0#)
        Dim sessionBody As String = JsonConvert.SerializeObject(New Dictionary(Of String, Object) From {
            {"algorithm", challenge.Item("algorithm").ToObject(Of String)},
            {"challenge", targetHash},
            {"salt", salt},
            {"signature", challenge.Item("signature").ToObject(Of String)},
            {"number", number}})
        Return SendJsonRequest(Me._url + "/session", "PUT", sessionBody)
    End Function

    Private Shared Function SolveAltchaChallenge(salt As String, targetHash As String, maxNumber As Long) As Long
        Dim targetBytes As Byte() = New Byte(targetHash.Length \ 2 - 1) {}
        For i As Integer = 0 To targetBytes.Length - 1
            targetBytes(i) = Convert.ToByte(targetHash.Substring(i * 2, 2), 16)
        Next
        Dim enc As New Text.UTF8Encoding
        Using sha As System.Security.Cryptography.SHA256 = System.Security.Cryptography.SHA256.Create()
            For n As Long = 0 To maxNumber
                Dim hashBytes As Byte() = sha.ComputeHash(enc.GetBytes(salt + n.ToString))
                Dim match As Boolean = hashBytes.Length = targetBytes.Length
                For i As Integer = 0 To hashBytes.Length - 1
                    If Not match Then Exit For
                    If hashBytes(i) <> targetBytes(i) Then match = False
                Next
                If match Then Return n
            Next
        End Using
        Return -1
    End Function

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

    Public Function getFullFileList() As List(Of WorkspaceFileDTO)
        Dim myReturn As New List(Of WorkspaceFileDTO)
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
                    For Each f As KeyValuePair(Of String, Linq.JToken) In allFiles
                        myReturn.AddRange(f.Value.ToObject(Of List(Of WorkspaceFileDTO)))
                    Next
                    Me._lastErrorMsgText = ""
                Catch ex As Exception
                    _lastErrorMsgText = ex.Message
                    If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
                End Try
            End Using
        End If
        Return myReturn
    End Function

    Public Function getFileDependencies(fileList As List(Of String)) As List(Of WorkspaceFileDTO)
        Dim myReturn As New List(Of WorkspaceFileDTO)
        Dim resp As Net.WebResponse
        _lastErrorMsgText = ""
        Try
            Dim uri As New Uri(Me._url + "/workspace/" + selectedWorkspace.ToString + "/files-dependencies")
            Dim requ As Net.WebRequest = Net.WebRequest.Create(uri)
            requ.Method = "POST"
            requ.ContentType = "application/json"
            requ.Headers.Item("AuthToken") = Me.tokenStr

            Dim bodyContent As String = "{""body"":" + JsonConvert.SerializeObject(fileList) + "}"

            requ.Method = "POST"
            requ.ContentType = "application/json"
            Dim enc As New Text.UTF8Encoding
            Dim dataBin As Byte() = enc.GetBytes(bodyContent)
            requ.ContentLength = dataBin.Length
            Dim s As IO.Stream = requ.GetRequestStream()
            s.Write(dataBin, 0, dataBin.Length)
            s.Close()

            resp = requ.GetResponse

        Catch ex As Net.WebException
            Dim rep As Net.HttpWebResponse = ex.Response
            Using rdr As New IO.StreamReader(rep.GetResponseStream())
                Dim errorMsg As String = rdr.ReadToEnd()
                Debug.Print(errorMsg)
                _lastErrorMsgText = ex.Message + " / " + errorMsg
            End Using
            resp = Nothing
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
                    For Each f As KeyValuePair(Of String, Linq.JToken) In allFiles
                        myReturn.AddRange(f.Value.ToObject(Of List(Of WorkspaceFileDTO)))
                    Next
                    Me._lastErrorMsgText = ""
                Catch ex As Exception
                    _lastErrorMsgText = ex.Message
                    If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
                End Try
            End Using
        End If
        Return myReturn
    End Function

    Public Function getFile(fileInfo As WorkspaceFileDTO, ByRef TargetFileStream As IO.FileStream, myBW As BackgroundWorker) As Boolean
        Dim myReturn As Boolean = False
        Dim resp As Net.WebResponse
        _lastErrorMsgText = ""
        Try
            Dim uri As New Uri(Me._url + "/workspace/" + selectedWorkspace.ToString + "/file/" + fileInfo.type + "/" + fileInfo.name)
            Dim requ As Net.WebRequest = Net.WebRequest.Create(uri)
            requ.Method = "GET"
            requ.Headers.Item("AuthToken") = Me.tokenStr
            resp = requ.GetResponse
        Catch ex As Exception
            resp = Nothing
            _lastErrorMsgText = ex.Message
            If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
        End Try
        If resp IsNot Nothing Then
            Using WebReader As New System.IO.BinaryReader(resp.GetResponseStream())
                Dim Buffer As Byte() = Nothing
                Try
                    Do
                        If Not myBW.CancellationPending Then
                            Buffer = WebReader.ReadBytes(_BufferSize)
                            If Buffer.Length > 0 Then TargetFileStream.Write(Buffer, 0, Buffer.Length)
                        End If
                    Loop Until myBW.CancellationPending OrElse Buffer.Length < BufferSize
                    myReturn = True
                Catch ex As Exception
                    _lastErrorMsgText = ex.Message
                    If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
                End Try
            End Using
        End If
        Return myReturn
    End Function

    Public Function getCsvFile(type As CsvType, dataGroupId As String, ByRef targetStream As IO.MemoryStream, myBW As BackgroundWorker) As Boolean
        Dim myReturn As Boolean = False
        Dim resp As Net.WebResponse
        _lastErrorMsgText = ""
        Dim url As String = Me._url + "/workspace/" + selectedWorkspace.ToString + "/report/" + IIf(type = CsvType.RESPONSES, "response", "log")
        Try
            Dim uri As New Uri(url + "?dataIds=" + dataGroupId + "&useNewVersion=false")
            Dim requ As HttpWebRequest = HttpWebRequest.Create(uri)
            requ.Method = "GET"
            requ.Accept = "text/csv"
            requ.Headers.Item("AuthToken") = Me.tokenStr
            resp = requ.GetResponse
        Catch ex As Exception
            resp = Nothing
            _lastErrorMsgText = ex.Message
            If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
        End Try
        If resp IsNot Nothing Then
            Using WebReader As New System.IO.BinaryReader(resp.GetResponseStream())
                Dim Buffer As Byte() = Nothing
                Try
                    Do
                        If Not myBW.CancellationPending Then
                            Buffer = WebReader.ReadBytes(_BufferSize)
                            If Buffer.Length > 0 Then targetStream.Write(Buffer, 0, Buffer.Length)
                        End If
                    Loop Until myBW.CancellationPending OrElse Buffer.Length < BufferSize
                    myReturn = True
                Catch ex As Exception
                    _lastErrorMsgText = ex.Message
                    If ex.InnerException IsNot Nothing Then _lastErrorMsgText += vbNewLine + ex.InnerException.Message
                End Try
            End Using
        End If
        Return myReturn
    End Function

    Public Function getBookletSizes() As Dictionary(Of String, Long)
        Dim bookletSizes As New Dictionary(Of String, Long)
        _lastErrorMsgText = ""
        Dim allFilesEmptyDependency As List(Of WorkspaceFileDTO) = getFullFileList()
        Dim bookletNames As List(Of String) = (From f As WorkspaceFileDTO In allFilesEmptyDependency
                                               Where f.type = "Booklet" Select f.name).ToList
        Dim bookletsWithDependency As New List(Of WorkspaceFileDTO)
        Const dependencyBatchSize As Integer = 600
        For batchStart As Integer = 0 To bookletNames.Count - 1 Step dependencyBatchSize
            Dim batch As List(Of String) = bookletNames.Skip(batchStart).Take(dependencyBatchSize).ToList
            bookletsWithDependency.AddRange(getFileDependencies(batch))
            If Not String.IsNullOrEmpty(_lastErrorMsgText) Then Exit For
        Next
        If String.IsNullOrEmpty(_lastErrorMsgText) Then
            Dim flatFileSizes As Dictionary(Of String, Long) = (From f As WorkspaceFileDTO In allFilesEmptyDependency).ToDictionary(Function(a) a.name, Function(a) a.size)

            For Each booklet As WorkspaceFileDTO In bookletsWithDependency
                Dim bookletSize As Long = booklet.size
                For Each d As FileDependencyDTO In booklet.dependencies
                    If validBookletDependencies.Contains(d.relationship_type) AndAlso flatFileSizes.ContainsKey(d.object_name) Then bookletSize += flatFileSizes.Item(d.object_name)
                Next
                bookletSizes.Add(booklet.id, bookletSize)
            Next
        End If

        Return bookletSizes
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
                    If Not String.IsNullOrEmpty(_response_string) Then myReturn = JsonConvert.DeserializeObject(_response_string, GetType(List(Of LogEntryDTO)))
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
                    If Not String.IsNullOrEmpty(_response_string) Then myReturn = JsonConvert.DeserializeObject(_response_string, GetType(List(Of ResponseDTO)))
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
