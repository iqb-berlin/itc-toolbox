Imports Newtonsoft.Json


Public Class LogEvent
    Public timestamp As Long
    Public unit As String = ""
    Public key As String = ""
    Public parameter As String = ""
End Class

Class TimeOnPage
    Public page As String = ""
    Public millisec As Long = 0
    Public count As Integer = 0
End Class

Public Class TimeOnUnit
    Public unit As String = ""
    Public navigationStart As Long = 0
    Public playerLoadTime As Long = 0
    Public stayTime As Long = 0
    Public wasPaused As Boolean = False
    Public responseProgressTimeSome As Long = 0
    Public responseProgressTimeComplete As Long = 0
    Public lostFocus As Boolean = False
End Class

Public Class TestPerson
    Public group As String
    Public login As String
    Public code As String
    Public booklet As String
    Public log As List(Of LogEvent)
    Public loadStart As Long = 0

    Public ReadOnly Property loadtime() As Long
        Get
            Return _loadtime
        End Get
    End Property
    Private _loadtime As Long
    Private _firstUnitEnter As Long
    Public Property firstUnitEnter() As Long
        Get
            Return _firstUnitEnter
        End Get
        Set(ByVal value As Long)
            If value < _firstUnitEnter Then _firstUnitEnter = value
        End Set
    End Property

    Private _browser As String
    Public ReadOnly Property browser() As String
        Get
            Return Me._browser
        End Get
    End Property
    Private _os As String
    Public ReadOnly Property os() As String
        Get
            Return Me._os
        End Get
    End Property
    Private _screen As String
    Public ReadOnly Property screen() As String
        Get
            Return Me._screen
        End Get
    End Property

    Public Sub New(g As String, l As String, c As String, b As String)
        group = g
        login = l
        code = c
        booklet = b
        _loadtime = 0
        _firstUnitEnter = Long.MaxValue
        log = New List(Of LogEvent)
        _browser = "?"
        _os = "?"
        _screen = "?"
    End Sub
    Public Function loadspeed(bookletsizelist As Dictionary(Of String, Long)) As Double
        Dim myreturn As Double = 0.0
        If _loadtime > 0 AndAlso bookletsizelist IsNot Nothing AndAlso bookletsizelist.ContainsKey(booklet) Then
            myreturn = bookletsizelist.Item(booklet) / Me._loadtime
        End If
        Return myreturn
    End Function
    Public Sub AddLogEvent(timestamp As Long, unit As String, event_key As String, event_parameter As String)
        log.Add(New LogEvent With {.timestamp = timestamp, .unit = unit, .key = event_key, .parameter = event_parameter})
    End Sub
    Public Function GetTimeOnUnitList() As List(Of TimeOnUnit)
        Dim myReturn As New List(Of TimeOnUnit)
        Dim currentTimeOnUnit As TimeOnUnit = Nothing
        Dim pageStart As Long = 0
        For Each logEntry As LogEvent In From le As LogEvent In Me.log Order By le.timestamp
            Select Case logEntry.key
                Case "CURRENT_UNIT_ID"
                    If currentTimeOnUnit IsNot Nothing Then
                        currentTimeOnUnit.stayTime = logEntry.timestamp - currentTimeOnUnit.navigationStart - currentTimeOnUnit.playerLoadTime
                        myReturn.Add(currentTimeOnUnit)
                    End If
                    currentTimeOnUnit = New TimeOnUnit With {.unit = logEntry.parameter, .navigationStart = logEntry.timestamp}
                Case "PLAYER"
                    If currentTimeOnUnit IsNot Nothing AndAlso logEntry.parameter = "RUNNING" Then
                        currentTimeOnUnit.playerLoadTime = logEntry.timestamp - currentTimeOnUnit.navigationStart
                    End If
                Case "FOCUS"
                    If currentTimeOnUnit IsNot Nothing AndAlso logEntry.parameter = "HAS_NOT" Then
                        currentTimeOnUnit.lostFocus = True
                    End If
                Case "CONTROLLER"
                    If currentTimeOnUnit IsNot Nothing Then
                        If logEntry.parameter = "PAUSED" Then
                            currentTimeOnUnit.wasPaused = True
                        ElseIf logEntry.parameter = "TERMINATED" Then
                            currentTimeOnUnit.stayTime = logEntry.timestamp - currentTimeOnUnit.navigationStart - currentTimeOnUnit.playerLoadTime
                            myReturn.Add(currentTimeOnUnit)
                            currentTimeOnUnit = Nothing
                        End If
                    End If
                Case "RESPONSE_PROGRESS"
                    If currentTimeOnUnit IsNot Nothing Then
                        If logEntry.parameter = "some" Then
                            currentTimeOnUnit.responseProgressTimeSome = logEntry.timestamp - currentTimeOnUnit.navigationStart - currentTimeOnUnit.playerLoadTime
                        ElseIf logEntry.parameter = "complete" Then
                            currentTimeOnUnit.responseProgressTimeComplete = logEntry.timestamp - currentTimeOnUnit.navigationStart - currentTimeOnUnit.playerLoadTime
                        End If
                    End If
            End Select
        Next
        Return myReturn
    End Function
    Public Function GetResponsesCompleteAllUnitCount(unitsOnly As List(Of String)) As Integer
        Dim unitList As New List(Of String)
        'For Each logList As KeyValuePair(Of Long, List(Of LogEvent)) In Me.log
        '    For Each logEntry As LogEvent In logList.Value
        '        If logEntry.key = "RESPONSESCOMPLETE" AndAlso logEntry.parameter = "all" AndAlso unitsOnly.Contains(logEntry.unit) AndAlso Not unitList.Contains(logEntry.unit) Then
        '            unitList.Add(logEntry.unit)
        '        End If
        '    Next
        'Next

        Return unitList.Count
    End Function
    Public Sub SetSysdata(timestamp As Long, sysdata As Dictionary(Of String, String))
        _browser = "?"
        _os = "?"
        _screen = "?"
        If sysdata IsNot Nothing Then
            If sysdata.ContainsKey("browserVersion") AndAlso sysdata.ContainsKey("browserName") Then _browser = sysdata.Item("browserName") + " " + sysdata.Item("browserVersion")
            If sysdata.ContainsKey("osName") Then _os = sysdata.Item("osName")
            If sysdata.ContainsKey("screenSizeWidth") AndAlso sysdata.ContainsKey("screenSizeHeight") Then _screen = sysdata.Item("screenSizeWidth") + " x " + sysdata.Item("screenSizeHeight")
            If sysdata.ContainsKey("loadTime") Then
                _loadtime = Long.Parse(sysdata.Item("loadTime"))
                loadStart = timestamp - _loadtime
            End If
        End If
    End Sub

    Function getFirstPlayerRunning() As Long
        Dim tsQuery As List(Of Long) = (From ev As LogEvent In Me.log Where ev.key = "PLAYER" AndAlso ev.parameter = "RUNNING" Select ev.timestamp).ToList()
        If tsQuery IsNot Nothing AndAlso tsQuery.Count > 0 Then
            Return tsQuery.Min() - loadStart
        Else
            Return 0
        End If
    End Function

    Function getLastActivity() As Long
        If Me.log.Count > 0 Then
            Return (From ev As LogEvent In Me.log Select ev.timestamp).Max() - loadStart
        Else
            Return 0
        End If
    End Function

End Class

Public Class TestPersonList
    Inherits SortedDictionary(Of String, TestPerson)
    Public Sub SetFirstUnitEnter(g As String, l As String, c As String, b As String, value As Long)
        If Not Me.ContainsKey(g + l + c + b) Then Me.Add(g + l + c + b, New TestPerson(g, l, c, b))
        Me.Item(g + l + c + b).firstUnitEnter = value
    End Sub
    Public Sub SetSysdata(timestamp As Long, g As String, l As String, c As String, b As String, sysdata As Dictionary(Of String, String))
        If Not Me.ContainsKey(g + l + c + b) Then Me.Add(g + l + c + b, New TestPerson(g, l, c, b))
        Me.Item(g + l + c + b).SetSysdata(timestamp, sysdata)
    End Sub
    Public Sub AddLogEvent(g As String, l As String, c As String, b As String, timestamp As Long, unit As String, event_key As String, event_parameter As String)
        If Not Me.ContainsKey(g + l + c + b) Then Me.Add(g + l + c + b, New TestPerson(g, l, c, b))
        Me.Item(g + l + c + b).AddLogEvent(timestamp, unit, event_key, event_parameter)
    End Sub
End Class