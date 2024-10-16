Imports Newtonsoft.Json

Public Class LogSymbols
    Public Const LogFileFirstLineLegacy = "groupname;loginname;code;bookletname;unitname;timestamp;logentry"
    Public Const LogFileFirstLine2024 = "groupname;loginname;code;bookletname;unitname;originalUnitId;timestamp;logentry"
End Class
Public Enum CsvLogFileType
    Legacy
    v2024
End Enum
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

Public Class UnitLineDataLog
    Public groupname As String
    Public loginname As String
    Public code As String
    Public bookletname As String
    Public unitname As String
    Public timestamp As Long
    Public originalUnitId As String
    Public eventKey As String
    Public eventParameter As String

    Public Shared Function fromCsvLine(line As String, fileType As CsvLogFileType) As UnitLineDataLog
        Dim lineSplits As String() = line.Split({""";"}, StringSplitOptions.RemoveEmptyEntries)
        Dim expectedParamCount As Integer = 7
        If fileType = CsvLogFileType.v2024 Then expectedParamCount = 8

        If lineSplits.Count <> expectedParamCount Then lineSplits = line.Split({";"}, StringSplitOptions.None)
        If lineSplits.Count = expectedParamCount Then
            Dim returnLogEntry As New UnitLineDataLog With {
                .groupname = IIf(lineSplits(0).Substring(0, 1) = """", lineSplits(0).Substring(1), lineSplits(0)),
                .loginname = IIf(lineSplits(1).Substring(0, 1) = """", lineSplits(1).Substring(1), lineSplits(1)),
                .code = "",
                .bookletname = IIf(lineSplits(3).Substring(0, 1) = """", lineSplits(3).Substring(1).ToUpper, lineSplits(3).ToUpper),
                .unitname = "",
                .originalUnitId = "",
                .eventKey = "",
                .eventParameter = "",
                .timestamp = 0
            }
            If Not String.IsNullOrEmpty(lineSplits(2)) Then returnLogEntry.code = IIf(lineSplits(2).Substring(0, 1) = """", lineSplits(2).Substring(1), lineSplits(2))
            If Not String.IsNullOrEmpty(lineSplits(4)) Then
                returnLogEntry.unitname = IIf(lineSplits(4).Substring(0, 1) = """", lineSplits(4).Substring(1), lineSplits(4))
                If fileType = CsvLogFileType.v2024 Then returnLogEntry.originalUnitId = IIf(lineSplits(5).Substring(0, 1) = """", lineSplits(5).Substring(1), lineSplits(4))
                If String.IsNullOrEmpty(returnLogEntry.originalUnitId) Then returnLogEntry.originalUnitId = returnLogEntry.unitname
            End If
            Dim offset As Integer = IIf(fileType = CsvLogFileType.Legacy, 0, 1)
            Dim timestampStr As String = IIf(lineSplits(5 + offset).Substring(0, 1) = """", lineSplits(5 + offset).Substring(1), lineSplits(5 + offset))
            If timestampStr.IndexOf("E+") > 0 Then
                returnLogEntry.timestamp = Long.Parse(timestampStr, System.Globalization.NumberStyles.Float)
            Else
                returnLogEntry.timestamp = Long.Parse(timestampStr)
            End If
            Dim entry As String = lineSplits(6 + offset)
            If Not String.IsNullOrEmpty(entry) AndAlso entry.Substring(0, 1) = """" Then
                entry = entry.Substring(1, entry.Length - 2).Replace("""""", """")
            End If

            returnLogEntry.eventKey = entry
            If returnLogEntry.eventKey.IndexOf(" : ") > 0 Then
                returnLogEntry.eventParameter = returnLogEntry.eventKey.Substring(returnLogEntry.eventKey.IndexOf(" : ") + 3)
                If returnLogEntry.eventParameter.IndexOf("""") = 0 AndAlso
                    returnLogEntry.eventParameter.LastIndexOf("""") = returnLogEntry.eventParameter.Length - 1 Then
                    returnLogEntry.eventParameter = returnLogEntry.eventParameter.Substring(1, returnLogEntry.eventParameter.Length - 2)
                    returnLogEntry.eventParameter = returnLogEntry.eventParameter.Replace("""""", """")
                    returnLogEntry.eventParameter = returnLogEntry.eventParameter.Replace("\\", "\")
                End If
                returnLogEntry.eventKey = returnLogEntry.eventKey.Substring(0, returnLogEntry.eventKey.IndexOf(" : "))
            ElseIf returnLogEntry.eventKey.IndexOf(" = ") > 0 Then
                returnLogEntry.eventParameter = returnLogEntry.eventKey.Substring(returnLogEntry.eventKey.IndexOf(" = ") + 3)
                returnLogEntry.eventKey = returnLogEntry.eventKey.Substring(0, returnLogEntry.eventKey.IndexOf(" = "))
            End If
            Return returnLogEntry
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function fromTestcenterAPI(log As LogEntryDTO) As UnitLineDataLog
        Dim returnLogEntry As New UnitLineDataLog With {
                .groupname = log.groupname,
                .loginname = log.loginname,
                .code = log.code,
                .bookletname = log.bookletname,
                .unitname = log.unitname,
                .originalUnitId = log.originalUnitId,
                .eventKey = log.logentry,
                .eventParameter = "",
                .timestamp = log.timestamp
            }
        If String.IsNullOrEmpty(returnLogEntry.originalUnitId) Then returnLogEntry.originalUnitId = returnLogEntry.unitname
        If returnLogEntry.eventKey.IndexOf(" : ") > 0 Then
            returnLogEntry.eventParameter = returnLogEntry.eventKey.Substring(returnLogEntry.eventKey.IndexOf(" : ") + 3)
            If returnLogEntry.eventParameter.IndexOf("""") = 0 AndAlso
                returnLogEntry.eventParameter.LastIndexOf("""") = returnLogEntry.eventParameter.Length - 1 Then
                returnLogEntry.eventParameter = returnLogEntry.eventParameter.Substring(1, returnLogEntry.eventParameter.Length - 2)
                returnLogEntry.eventParameter = returnLogEntry.eventParameter.Replace("""""", """")
                returnLogEntry.eventParameter = returnLogEntry.eventParameter.Replace("\\", "\")
            End If
            returnLogEntry.eventKey = returnLogEntry.eventKey.Substring(0, returnLogEntry.eventKey.IndexOf(" : "))
        ElseIf returnLogEntry.eventKey.IndexOf(" = ") > 0 Then
            returnLogEntry.eventParameter = returnLogEntry.eventKey.Substring(returnLogEntry.eventKey.IndexOf(" = ") + 3)
            returnLogEntry.eventKey = returnLogEntry.eventKey.Substring(0, returnLogEntry.eventKey.IndexOf(" = "))
        End If
        Return returnLogEntry
    End Function
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
