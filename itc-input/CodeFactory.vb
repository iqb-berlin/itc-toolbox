Public Class CodeFactory
    Private Const codeCharacters = "abcdefghprqstuvxyz"
    Private Const codeNumbers = "2345679"
    Public Shared Function GetNewCode(codeLen As Integer) As String
        Dim newCode As String = ""
        Dim isNumber As Boolean = False
        Randomize()
        Do
            newCode = newCode & IIf(isNumber, Mid(codeNumbers, Int(codeNumbers.Length * Rnd() + 1), 1), Mid(codeCharacters, Int(codeCharacters.Length * Rnd() + 1), 1))
            isNumber = Not isNumber
        Loop Until newCode.Length = codeLen
        Return newCode
    End Function

    Public Shared Function GetNewCodeList(codeLen As Integer, codeCount As Integer) As List(Of String)
        Dim codeList As New List(Of String)
        Randomize()
        For i As Integer = 1 To codeCount
            Dim newCode As String
            Do
                newCode = ""
                Dim isNumber As Boolean = False
                Do
                    newCode = newCode & IIf(isNumber, Mid(codeNumbers, Int(codeNumbers.Length * Rnd() + 1), 1), Mid(codeCharacters, Int(codeCharacters.Length * Rnd() + 1), 1))
                    isNumber = Not isNumber
                Loop Until newCode.Length = codeLen
            Loop While codeList.Contains(newCode)
            codeList.Add(newCode)
        Next
        Return codeList
    End Function
End Class

Public Class groupdata
    Public id As String = ""
    Public name1 As String = ""
    Public name2 As String = ""
    Public numberLogins As Integer = 0
    Public numberLoginsPlus As Integer = 0
    Public numberReviews As Integer = 0
    Public logins As New List(Of logindata)
    Public Function toXml(Optional bookletName As String = "Booklet1") As XElement
        Dim myreturn As XElement = <Group id=<%= id %> label=<%= name1 + " - " + name2 %>></Group>
        For Each login As logindata In logins
            myreturn.Add(login.toXml(bookletName))
        Next
        Return myreturn
    End Function
End Class

'#############################################################################
Public Class logindata
    Public login As String
    Public password As String = ""
    Public mode As String = "run-hot-return"

    Public Function toXml(Optional bookletName As String = "Booklet1") As XElement
        Dim myreturn As XElement = <Login mode=<%= mode %> name=<%= login %>>
                                   </Login>
        If mode <> "monitor-group" Then myreturn.Add(<Booklet><%= bookletName %></Booklet>)
        If Not String.IsNullOrEmpty(password) Then myreturn.SetAttributeValue("pw", password)
        Return myreturn
    End Function

End Class