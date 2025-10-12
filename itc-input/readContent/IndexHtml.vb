Imports System.Net.Mime

Public Class IndexHtml
    Private Shared Sub addFileDataToTable(f As WorkspaceFileDTO, tableRow As XElement)
        tableRow.Add(<td><%= f.name %></td>)
        tableRow.Add(<td><%= GetFileSizeString(f.size) %></td>)
        Dim fileDateTime As New DateTime(1970, 1, 1, 0, 0, 0, 0)
        fileDateTime = fileDateTime.AddSeconds(f.modificationTime)
        tableRow.Add(<td><%= fileDateTime.ToShortDateString + " " + fileDateTime.ToShortTimeString %></td>)
    End Sub
    Public Shared Sub WriteTo(TargetFilename As String, url As String, workspaceName As String, files As List(Of WorkspaceFileDTO))
        Dim TargetDocX As XDocument = <?xml version="1.0" encoding="utf-8"?>
                                      <html>
                                          <head>
                                              <meta http-equiv="content-type" content="text/html; charset=UTF-8"/>
                                              <title>itc-ToolBox</title>
                                              <style>
                                                body {
                                                    font-family: Helvetica;
                                                    font-size: 1em;
                                                    margin: 0;
                                                    padding-left: 30px;
                                                    padding-right: 10px;
                                                    padding-top: 10px;
                                                }
                                                h1 {
                                                    color: blue;
                                                }
                                                table td {
                                                    padding: 2px 4px;
                                                }
                                                table td.int {
                                                    text-align: right;
                                                }
                                                table td.error {
                                                    color: red;
                                                }
                                                table tr td {
                                                    vertical-align: top;
                                                    border: 1px dotted;
                                                }
                                                table th {
                                                    background-color: lightblue;
                                                    border: 1px dotted;
                                                }
                                                table th.w20 {
                                                    width: 20%;
                                                }
                                                table th.w30 {
                                                    width: 30%;
                                                }
                                                table th.w40 {
                                                    width: 40%;
                                                }
                                                table {
                                                    border: 1px dotted;
                                                    width: 100%;
                                                }
                                                h3 {
                                                    background-color: AEEDEA;
                                                }
                                             </style>
                                          </head>
                                          <body/>
                                      </html>


        TargetDocX.Root.Add(<p>IQB-Testcenter Download Content-Dateien</p>)
        TargetDocX.Root.Add(<p style="font-size:24pt;"><%= workspaceName %></p>)
        Dim headerData As XElement = <ul></ul>
        headerData.Add(<li>URL Testcenter: <%= url.Replace("/api", "") %></li>)
        headerData.Add(<li>Download-Datum: <%= Date.Now.ToString("d.M.yyyy H:mm") %></li>)
        Dim maxModiValue As Long = (From f In files Select f.modificationTime).Max
        Dim modiDateTime As New DateTime(1970, 1, 1, 0, 0, 0, 0)
        modiDateTime = modiDateTime.AddSeconds(maxModiValue)
        headerData.Add(<li>Letzte Änderung: <%= modiDateTime.ToString("d.M.yyyy H:mm") %></li>)
        TargetDocX.Root.Add(headerData)

        Dim FileDataTable As XElement
        Dim habschonRessources As New List(Of String)
        Dim hasRows As Boolean

        'Booklets
        FileDataTable = <table><tr><th>ID</th><th>Titel</th><th>Datei</th><th>Größe</th><th>Datum</th></tr></table>
        hasRows = False
        For Each f As WorkspaceFileDTO In From fileOfType In files Where fileOfType.type = "Booklet" Order By fileOfType.id
            Dim TableRow As XElement = <tr/>
            hasRows = True
            TableRow.Add(<td><%= f.id %></td>)
            Dim title As String = ""
            If f.info IsNot Nothing Then
                If Not String.IsNullOrEmpty(f.info.label) Then title = f.info.label
            End If
            TableRow.Add(<td><%= title %></td>)
            addFileDataToTable(f, TableRow)
            FileDataTable.Add(TableRow)
        Next
        If hasRows Then
            TargetDocX.Root.Add(<h1>Booklets</h1>)
            TargetDocX.Root.Add(FileDataTable)
        End If

        'Units
        FileDataTable = <table><tr><th class="w30">Unit</th><th>Player</th><th class="w30">Datei</th><th>Größe</th><th>Datum</th></tr></table>
        hasRows = False
        For Each f As WorkspaceFileDTO In From fileOfType In files Where fileOfType.type = "Unit" Order By fileOfType.id
            Dim TableRow As XElement = <tr/>
            hasRows = True
            Dim unitInfo As String = f.id
            If f.info IsNot Nothing Then
                If f.info.label IsNot Nothing AndAlso Not String.IsNullOrEmpty(f.info.label.Trim) Then unitInfo += " - " + f.info.label.Trim
                If f.info.description IsNot Nothing AndAlso Not String.IsNullOrEmpty(f.info.description.Trim) Then unitInfo += " (" + f.info.description.Trim + ")"
            End If
            TableRow.Add(<td><%= unitInfo %></td>)
            Dim otherUnitFiles As New List(Of String)
            If f.dependencies IsNot Nothing Then
                TableRow.Add(<td><%= String.Join(", ", From d As FileDependencyDTO In f.dependencies Where d.relationship_type = "usesPlayer" Select d.object_name) %></td>)
                otherUnitFiles = (From d As FileDependencyDTO In f.dependencies
                                  Where d.relationship_type = "isDefinedBy" OrElse d.relationship_type = "usesScheme"
                                  Select d.object_name).ToList
            End If
            addFileDataToTable(f, TableRow)
            FileDataTable.Add(TableRow)
            If otherUnitFiles.Count > 0 Then
                For Each depObject As String In otherUnitFiles
                    Dim ressourceToAdd As WorkspaceFileDTO = (From r In files Where r.name = depObject).FirstOrDefault
                    If ressourceToAdd IsNot Nothing Then
                        Dim RessourceTableRow As XElement = <tr/>
                        RessourceTableRow.Add(<td></td>)
                        RessourceTableRow.Add(<td></td>)
                        addFileDataToTable(ressourceToAdd, RessourceTableRow)
                        FileDataTable.Add(RessourceTableRow)
                        habschonRessources.Add(depObject)
                    End If
                Next
            End If
        Next
        If hasRows Then
            TargetDocX.Root.Add(<h1>Units</h1>)
            TargetDocX.Root.Add(FileDataTable)
        End If

        'Testtakers
        FileDataTable = <table><tr><th>ID</th><th>Anzahl Logins</th><th>Datei</th><th>Größe</th><th>Datum</th></tr></table>
        hasRows = False
        For Each f As WorkspaceFileDTO In From fileOfType In files Where fileOfType.type = "Testtakers" Order By fileOfType.id
            Dim TableRow As XElement = <tr/>
            hasRows = True
            TableRow.Add(<td><%= f.id %></td>)
            Dim loginCount As String = "0"
            If f.info IsNot Nothing Then loginCount = f.info.testtakers.ToString("#,##0")
            TableRow.Add(<td><%= loginCount %></td>)
            addFileDataToTable(f, TableRow)
            FileDataTable.Add(TableRow)
        Next
        If hasRows Then
            TargetDocX.Root.Add(<h1>Testtakers</h1>)
            TargetDocX.Root.Add(FileDataTable)
        End If

        'Others
        Dim contentTypes As New List(Of String) From {"SysCheck", "Resource"}
        For Each contentType As String In contentTypes
            FileDataTable = <table><tr><th class="w30">Datei</th><th>Größe</th><th>Datum</th></tr></table>
            hasRows = False
            For Each f As WorkspaceFileDTO In From fileOfType In files Where fileOfType.type = contentType Order By fileOfType.id
                If contentType <> "Resource" OrElse Not habschonRessources.Contains(f.name) Then
                    Dim TableRow As XElement = <tr/>
                    hasRows = True
                    addFileDataToTable(f, TableRow)
                    FileDataTable.Add(TableRow)
                End If
            Next
            If hasRows Then
                TargetDocX.Root.Add(<h1><%= contentType %></h1>)
                TargetDocX.Root.Add(FileDataTable)
            End If
        Next

        Dim XDocIndexHtmlContent As String = TargetDocX.ToString()
        IO.File.WriteAllText(TargetFilename, XDocIndexHtmlContent)
    End Sub

    Private Shared deDK As Globalization.CultureInfo = Globalization.CultureInfo.CreateSpecificCulture("de-DE")
    Private Shared Function GetFileSizeString(FileSizeInt As Integer) As String
        If FileSizeInt = 0 Then
            Return "?"
        ElseIf FileSizeInt <1000 Then
            Return "<1 kB"
        ElseIf FileSizeInt <1000000 Then
            Return Math.Round(FileSizeInt / 1000).ToString + " kB"
        ElseIf FileSizeInt < 20000000 Then
            Return (FileSizeInt / 1000000).ToString("##0.#", deDK) + " MB"
        Else
            Return (FileSizeInt / 1000000).ToString("#,##0", deDK) + " MB"
        End If
    End Function
End Class
