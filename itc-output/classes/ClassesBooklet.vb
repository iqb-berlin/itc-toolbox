Imports System.Windows.Forms
Imports DocumentFormat.OpenXml.InkML
Imports Microsoft.VisualBasic.Logging
Imports Newtonsoft.Json

Public Class PersonList
    Inherits SortedDictionary(Of String, Person)
    Public Sub SetFirstUnitEnter(g As String, l As String, c As String, b As String, value As Long)
        If Not Me.ContainsKey(g + l + c) Then Me.Add(g + l + c, New Person(g, l, c))
        Dim myPerson As Person = Me.Item(g + l + c)
        Dim myBooklet As Booklet = (From bt As Booklet In myPerson.booklets Where bt.id = b).FirstOrDefault
        If myBooklet Is Nothing Then
            myBooklet = New Booklet(b)
            myPerson.booklets.Add(myBooklet)
        End If
        'myBooklet.firstUnitEnterTS = value
    End Sub
    Public Sub AddLogEntry(g As String, l As String, c As String, b As String, timestamp As Long, unit As String, event_key As String, event_parameter As String)
        If Not Me.ContainsKey(g + l + c) Then Me.Add(g + l + c, New Person(g, l, c))
        Me.Item(g + l + c).addLogEntry(b, timestamp, unit, event_key, event_parameter)
    End Sub
    Public Sub AddUnitData(unitdata As UnitLineData)
        Dim personKey As String = unitdata.groupname + unitdata.loginname + unitdata.code
        If Not Me.ContainsKey(personKey) Then Me.Add(personKey, New Person(unitdata.groupname, unitdata.loginname, unitdata.code))
        Dim myPerson As Person = Me.Item(personKey)
        Dim myBooklet As Booklet = (From bt As Booklet In myPerson.booklets Where bt.id = unitdata.bookletname).FirstOrDefault
        If myBooklet Is Nothing Then
            myBooklet = New Booklet(unitdata.bookletname)
            myPerson.booklets.Add(myBooklet)
        End If
        'could lead to double unit entries - to be solved later
        Dim myUnit As Unit = (From u As Unit In myBooklet.units Where u.alias = unitdata.unitname).FirstOrDefault
        If myUnit Is Nothing OrElse myUnit.subforms.Count > 0 OrElse myUnit.laststate.Count > 0 OrElse myUnit.chunks.Count > 0 Then
            myBooklet.units.Add(New Unit(unitdata.unitname) With {
                            .chunks = unitdata.chunks, .laststate = unitdata.laststate, .subforms = unitdata.subforms})
        Else
            myUnit.chunks = unitdata.chunks
            myUnit.laststate = unitdata.laststate
            myUnit.subforms = unitdata.subforms
        End If
    End Sub

    Public Function ToUnitLineData() As Dictionary(Of String, Dictionary(Of String, List(Of UnitLineData)))
        Dim returnDict As New Dictionary(Of String, Dictionary(Of String, List(Of UnitLineData)))

        Return returnDict
    End Function
End Class

Public Class Person
    Public group As String
    Public login As String
    Public code As String
    Public booklets As List(Of Booklet)

    Public Sub New(g As String, l As String, c As String)
        group = g
        login = l
        code = c
        booklets = New List(Of Booklet)
    End Sub

    Public Sub addLogEntry(bookletName As String, timestamp As Long, unit As String, event_key As String, event_parameter As String)
        Dim myBooklet As Booklet = (From b As Booklet In booklets Where b.id = bookletName).FirstOrDefault
        If myBooklet Is Nothing Then
            myBooklet = New Booklet(bookletName)
            booklets.Add(myBooklet)
        End If
        If String.IsNullOrEmpty(unit) Then
            myBooklet.logs.Add(New LogEntry(timestamp, event_key, event_parameter))
            If event_key = "LOADCOMPLETE" Then
                Dim sysdata As Dictionary(Of String, String) = Nothing
                event_parameter = event_parameter.Replace("\""", """")
                Try
                    sysdata = JsonConvert.DeserializeObject(event_parameter, GetType(Dictionary(Of String, String)))
                Catch ex As Exception
                    sysdata = Nothing
                    Debug.Print("sysdata json convert failed: " + ex.Message)
                End Try
                myBooklet.addSession(timestamp, sysdata)
            End If
        Else
            Dim myUnit As Unit = (From u As Unit In myBooklet.units Where u.alias = unit).FirstOrDefault
            If myUnit Is Nothing Then
                myUnit = New Unit(unit)
                myBooklet.units.Add(myUnit)
            End If
            myUnit.logs.Add(New LogEntry(timestamp, event_key, event_parameter))
        End If
    End Sub
End Class

Public Structure LogEntry
    Public ts As Long
    Public key As String
    Public parameter As String
    Public Sub New(ts As Long, key As String, parameter As String)
        Me.ts = ts
        Me.key = key
        Me.parameter = parameter
    End Sub
End Structure

Public Class TimeOnPageData
    Public navigationStart As Long = 0
    Public playerLoadTime As Long = 0
    Public stayTime As Long = 0
    Public responseProgressTimeSome As Long = 0
    Public responseProgressTimeComplete As Long = 0
    Public wasPaused As Boolean = False
    Public lostFocus As Boolean = False
End Class

Public Class Unit
    Public id As String
    Public [alias] As String
    Public laststate As List(Of LastStateEntry)
    Public subforms As List(Of SubForm)
    Public chunks As List(Of ResponseChunkData)
    Public logs As List(Of LogEntry)

    Public Sub New(unitId As String, Optional unitAlias As String = Nothing)
        id = unitId
        [alias] = IIf(String.IsNullOrEmpty(unitAlias), unitId, unitAlias)
        laststate = New List(Of LastStateEntry)
        subforms = New List(Of SubForm)
        chunks = New List(Of ResponseChunkData)
        logs = New List(Of LogEntry)
    End Sub
    Public Function getTimeOnPageData() As TimeOnPageData
        Return New TimeOnPageData
    End Function
End Class

Public Class Session
    Public browser As String
    Public os As String
    Public screen As String
    Public ts As Long
    Public loadCompleteMS As Long
End Class

Public Class BookletTechData
    Public loadspeed As Long
    Public firstUnitRunningAfterMS As Long
    Public browser As String
    Public os As String
    Public screen As String
    Public loadStart As Long
    Public loadTimeCompleteTS As Long
End Class

Public Class Booklet
    Public id As String
    'Public firstUnitEnterTS As Long
    Public logs As List(Of LogEntry)
    Public units As List(Of Unit)
    Public sessions As List(Of Session)

    Public Sub New(bookletId As String)
        id = bookletId
        logs = New List(Of LogEntry)
        units = New List(Of Unit)
        sessions = New List(Of Session)
    End Sub

    Public Sub addSession(timestamp As Long, sysdata As Dictionary(Of String, String))
        Dim newSession As New Session With {.ts = timestamp, .loadCompleteMS = 0}
        If sysdata IsNot Nothing Then
            If sysdata.ContainsKey("browserVersion") AndAlso sysdata.ContainsKey("browserName") Then newSession.browser = sysdata.Item("browserName") + " " + sysdata.Item("browserVersion")
            If sysdata.ContainsKey("osName") Then newSession.os = sysdata.Item("osName")
            If sysdata.ContainsKey("screenSizeWidth") AndAlso sysdata.ContainsKey("screenSizeHeight") Then newSession.screen = sysdata.Item("screenSizeWidth") + " x " + sysdata.Item("screenSizeHeight")
            If sysdata.ContainsKey("loadTime") Then newSession.loadCompleteMS = Long.Parse(sysdata.Item("loadTime"))
        End If
        sessions.Add(newSession)
    End Sub

    Public Function getTechData(bookletSizes As Dictionary(Of String, Long)) As BookletTechData
        Return New BookletTechData
    End Function

End Class
