Public Class IndexHtml
    Public Shared Sub WriteTo(TargetFilename As String, files As List(Of WorkspaceFileDTO))
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
                                                    background-color: lightgreen;
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


        TargetDocX.Root.Add(<p style="font-size:24pt;">IQB-Testcenter Download Content-Dateien</p>)
        TargetDocX.Root.Add(<p style="margin-top:0px;"><em><%= Date.Now.ToString("d.M.yyyy H:mm") %></em></p>)

        Dim contentTypes As New List(Of String) From {"Booklet", "SysCheck", "Testtakers", "Unit", "Resource"}
        For Each contentType As String In contentTypes
            TargetDocX.Root.Add(<h1><%= contentType %></h1>)
            Dim FileDataTable As XElement = <table><tr><th class="w30">Datei</th><th>Größe</th><th>Datum</th><th class="w20">Bearbeiter</th><th class="w30">Beschreibung</th></tr></table>

            For Each f As WorkspaceFileDTO In From fileOfType In files Where fileOfType.type = contentType Order By fileOfType.name
                Dim TableRow As XElement = <tr/>
                TableRow.Add(<td><%= f.name %></td>)
                TableRow.Add(<td><%= GetFileSizeString(f.size) %></td>)
                FileDataTable.Add(TableRow)
            Next

            TargetDocX.Root.Add(FileDataTable)
        Next

        Dim XDocIndexHtmlContent As String = TargetDocX.ToString()
        IO.File.WriteAllText(TargetFilename, XDocIndexHtmlContent)
    End Sub

    Private Shared deDK As Globalization.CultureInfo = Globalization.CultureInfo.CreateSpecificCulture("de-DE")
    Private Shared Function GetFileSizeString(FileSizeInt As Integer) As String
        If FileSizeInt = 0 Then
            Return "?"
        ElseIf FileSizeInt < 1000 Then
            Return "< 1 kB"
        ElseIf FileSizeInt < 1000000 Then
            Return Math.Round(FileSizeInt / 1000).ToString + " kB"
        ElseIf FileSizeInt < 20000000 Then
            Return (FileSizeInt / 1000000).ToString("##0.#", deDK) + " MB"
        Else
            Return (FileSizeInt / 1000000).ToString("#,##0", deDK) + " MB"
        End If
    End Function
End Class
