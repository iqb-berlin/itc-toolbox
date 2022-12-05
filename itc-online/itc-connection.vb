Imports Newtonsoft.Json
Public Class ITCConnection
    Private _url As String
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
    Public Sub New(url As String, credents As Net.NetworkCredential)
        Me._url = url
        Dim resp As Net.WebResponse = Nothing
        Try
            Dim uri As New Uri(url + "/session/admin")
            Dim requ As Net.WebRequest = Net.WebRequest.Create(uri)
            requ.Method = "PUT"
            requ.ContentType = "application/json"
            Dim enc As New Text.UTF8Encoding
            Dim dataBin As Byte() = enc.GetBytes("{""name"":""" + credents.UserName + """,""password"": """ + credents.Password + """}")
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
        End If
    End Sub

End Class
