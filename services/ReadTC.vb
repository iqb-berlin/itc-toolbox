Imports Newtonsoft.Json
Imports System.Net.Http

Public Class ReadTC
    Const url = "https://www.iqb-testcenter.de"
    Shared credentials As New Net.NetworkCredential("mmee", "tutti22")

    Public Shared Async Function f1() As Task(Of String)
        Dim resp As Net.WebResponse = Nothing
        Dim _response_string As String
        Dim tokenStr As String
        Dim accessTo As New Dictionary(Of Integer, String)

        Dim client As New HttpClient()
        Dim content As New StringContent("{""name"":""" + credentials.UserName + """,""password"": """ + credentials.Password + """}", System.Text.Encoding.UTF8)
        Using response As HttpResponseMessage = Await client.PutAsync(url + "/session/admin", content)
            response.EnsureSuccessStatusCode()
            Dim responseBody As String = Await response.Content.ReadAsStringAsync()
            Console.WriteLine(responseBody)
        End Using

        Return "yoyo"
    End Function
End Class
