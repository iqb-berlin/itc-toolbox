Imports System.Data.Common
Imports System.Data.SQLite
Imports System.Globalization
Imports iqb.lib.windows

Public Class SQLiteConnector
    Const currentDbVersion = 1
    Private fileName As String
    Public ReadOnly dbVersion As Integer
    Public ReadOnly dbCreator As String
    Public ReadOnly dbCreatedDateTime As String
    Public Sub New(dbFileName As String)
        fileName = dbFileName
        Dim addFullSchema As Boolean = Not IO.File.Exists(fileName)
        Using sqliteConnection As SQLiteConnection = GetOpenConnection(False)
            If addFullSchema Then
                Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                    cmd.CommandText = "
SELECT 1;
PRAGMA foreign_keys=OFF;
BEGIN TRANSACTION;
CREATE TABLE [db_info] (
  [key] text NOT NULL
, [value] text NOT NULL
);
CREATE TABLE [person] (
  [id] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL
, [group] text NOT NULL
, [login]  NOT NULL
, [code]  NULL
);
CREATE TABLE [bookletInfo] (
  [id] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL
, [name] text NOT NULL
, [size] bigint DEFAULT (0) NOT NULL
);
CREATE TABLE [booklet] (
  [id] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL
, [infoId] bigint NOT NULL
, [personId] bigint NOT NULL
, [lastTs] bigint DEFAULT (0) NOT NULL
, [firstTs] bigint DEFAULT (0) NOT NULL
, CONSTRAINT [FK_booklet_0_0] FOREIGN KEY ([infoId]) REFERENCES [bookletInfo] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
, CONSTRAINT [FK_booklet_1_0] FOREIGN KEY ([personId]) REFERENCES [person] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [unit] (
  [id] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL
, [bookletId] bigint NOT NULL
, [name] text NOT NULL
, [alias] text NULL
, CONSTRAINT [FK_unit_0_0] FOREIGN KEY ([bookletId]) REFERENCES [booklet] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [subform] (
  [id] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL
, [unitId] bigint NOT NULL
, [key] text NULL
, CONSTRAINT [FK_subform_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [response] (
  [subformId] bigint NOT NULL
, [variableId] text NOT NULL
, [value] text NULL
, [status] text NOT NULL
, [code] bigint DEFAULT (0) NOT NULL
, [score] bigint DEFAULT (0) NOT NULL
, CONSTRAINT [FK_response_0_0] FOREIGN KEY ([subformId]) REFERENCES [subform] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [chunk] (
  [unitId] bigint NOT NULL
, [key] text NOT NULL
, [type] text NULL
, [variables] text NULL
, [ts] bigint NULL
, CONSTRAINT [FK_chunk_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [unitLastState] (
  [unitId] bigint NOT NULL
, [key] text NOT NULL
, [value] text NULL
, CONSTRAINT [FK_unitLastState_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [unitLog] (
  [unitId] bigint NOT NULL
, [key] text NOT NULL
, [parameter] text NULL
, [ts] bigint NULL
, CONSTRAINT [FK_unitLog_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [bookletLog] (
  [bookletId] bigint NOT NULL
, [key] text NOT NULL
, [parameter] text NULL
, [ts] bigint NULL
, CONSTRAINT [FK_bookletLog_0_0] FOREIGN KEY ([bookletId]) REFERENCES [booklet] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [session] (
  [bookletId] bigint NOT NULL
, [browser] text NULL
, [os] text NULL
, [screen] text NULL
, [ts] bigint NULL
, [loadCompleteMS] bigint NULL
, CONSTRAINT [FK_session_0_0] FOREIGN KEY ([bookletId]) REFERENCES [booklet] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TRIGGER [fki_bookletLog_bookletId_booklet_id] BEFORE Insert ON [bookletLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table bookletLog violates foreign key constraint FK_bookletLog_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fki_booklet_infoId_bookletInfo_id] BEFORE Insert ON [booklet] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table booklet violates foreign key constraint FK_booklet_0_0') WHERE (SELECT id FROM bookletInfo WHERE  id = NEW.infoId) IS NULL; END;
CREATE TRIGGER [fki_booklet_personId_person_id] BEFORE Insert ON [booklet] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table booklet violates foreign key constraint FK_booklet_1_0') WHERE (SELECT id FROM person WHERE  id = NEW.personId) IS NULL; END;
CREATE TRIGGER [fki_chunk_unitId_unit_id] BEFORE Insert ON [chunk] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table chunk violates foreign key constraint FK_chunk_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_response_subformId_subform_id] BEFORE Insert ON [response] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table response violates foreign key constraint FK_response_0_0') WHERE (SELECT id FROM subform WHERE  id = NEW.subformId) IS NULL; END;
CREATE TRIGGER [fki_session_bookletId_booklet_id] BEFORE Insert ON [session] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table session violates foreign key constraint FK_session_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fki_subform_unitId_unit_id] BEFORE Insert ON [subform] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table subform violates foreign key constraint FK_subform_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_unitLastState_unitId_unit_id] BEFORE Insert ON [unitLastState] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table unitLastState violates foreign key constraint FK_unitLastState_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_unitLog_unitId_unit_id] BEFORE Insert ON [unitLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table unitLog violates foreign key constraint FK_unitLog_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_unit_bookletId_booklet_id] BEFORE Insert ON [unit] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table unit violates foreign key constraint FK_unit_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fku_bookletLog_bookletId_booklet_id] BEFORE Update ON [bookletLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table bookletLog violates foreign key constraint FK_bookletLog_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fku_booklet_infoId_bookletInfo_id] BEFORE Update ON [booklet] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table booklet violates foreign key constraint FK_booklet_0_0') WHERE (SELECT id FROM bookletInfo WHERE  id = NEW.infoId) IS NULL; END;
CREATE TRIGGER [fku_booklet_personId_person_id] BEFORE Update ON [booklet] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table booklet violates foreign key constraint FK_booklet_1_0') WHERE (SELECT id FROM person WHERE  id = NEW.personId) IS NULL; END;
CREATE TRIGGER [fku_chunk_unitId_unit_id] BEFORE Update ON [chunk] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table chunk violates foreign key constraint FK_chunk_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_response_subformId_subform_id] BEFORE Update ON [response] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table response violates foreign key constraint FK_response_0_0') WHERE (SELECT id FROM subform WHERE  id = NEW.subformId) IS NULL; END;
CREATE TRIGGER [fku_session_bookletId_booklet_id] BEFORE Update ON [session] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table session violates foreign key constraint FK_session_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fku_subform_unitId_unit_id] BEFORE Update ON [subform] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table subform violates foreign key constraint FK_subform_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_unitLastState_unitId_unit_id] BEFORE Update ON [unitLastState] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table unitLastState violates foreign key constraint FK_unitLastState_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_unitLog_unitId_unit_id] BEFORE Update ON [unitLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table unitLog violates foreign key constraint FK_unitLog_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_unit_bookletId_booklet_id] BEFORE Update ON [unit] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table unit violates foreign key constraint FK_unit_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
COMMIT;"
                    cmd.ExecuteNonQuery()
                End Using
                dbVersion = 1
                Dim now As DateTime = DateTime.Now
                dbCreatedDateTime = now.ToShortDateString + " " + now.ToShortTimeString
                dbCreator = ADFactory.GetMyNameLong
                Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                    cmd.CommandText = "
                        INSERT INTO [db_info] ([key],[value]) VALUES ('name', 'IQB-Testcenter-Output');
                        INSERT INTO [db_info] ([key],[value]) VALUES ('dbVersion', '" + dbVersion.ToString + "');
                        INSERT INTO [db_info] ([key],[value]) VALUES ('dbCreator', '" + dbCreator + "');
                        INSERT INTO [db_info] ([key],[value]) VALUES ('dbCreatedDateTime', '" + dbCreatedDateTime + "');
                    "
                    cmd.ExecuteNonQuery()
                End Using
            Else
                Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                    cmd.CommandText = "SELECT * FROM [db_info];"
                    Dim dbReader As SQLiteDataReader = cmd.ExecuteReader()
                    While dbReader.Read()
                        Dim key As String = dbReader.GetString(0)
                        Dim value As String = dbReader.GetString(1)
                        Select Case key
                            Case "dbVersion" : dbVersion = Long.Parse(value)
                            Case "dbCreator" : dbCreator = value
                            Case "dbCreatedDateTime" : dbCreatedDateTime = value
                        End Select
                    End While
                End Using
            End If
        End Using
    End Sub

    Public Function GetOpenConnection(ReadOnlyMode As Boolean) As SQLiteConnection
        Dim fact As DbProviderFactory = DbProviderFactories.GetFactory("System.Data.SQLite")
        Dim sqliteConnection As SQLiteConnection = fact.CreateConnection()
        sqliteConnection.ConnectionString = "Data Source=" + fileName + IIf(ReadOnlyMode, ";Read Only=True;", "")
        sqliteConnection.Open()
        Return sqliteConnection
    End Function

    Public Function GetCoreData(closeConnection As Boolean) As String
        Dim returnText As String = ""
        Using sqliteConnection As SQLiteConnection = GetOpenConnection(True)
            Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                cmd.CommandText = "SELECT COUNT(*) FROM [person];"
                Dim personCount As Long = cmd.ExecuteScalar()
                cmd.CommandText = "SELECT COUNT(*) FROM [response];"
                Dim responseCount As Long = cmd.ExecuteScalar()
                Dim deCulture = CultureInfo.CreateSpecificCulture("de-DE")
                returnText = "Anzahl Personen: " + personCount.ToString("N0", deCulture) +
                    ", Anzahl Antwortdaten: " + responseCount.ToString("N0", deCulture)
            End Using
        End Using
        If closeConnection Then Me.CloseConnection()
        Return returnText
    End Function

    Public Sub CloseConnection()
        GC.Collect()
        GC.WaitForPendingFinalizers()
    End Sub

    Public Sub addPerson(p As Person)
        Dim addedBookletCount As Integer = 0
        Dim updatedBookletCount As Integer = 0
        Dim ignoredBookletCount As Integer = 0
        Using sqliteConnection As SQLiteConnection = GetOpenConnection(False)
            Dim personDbId As Long = -1
            Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                cmd.CommandText = "select [id] from [person] where [group]='" + p.group +
                    "' and [login]='" + p.login + "' and [code]='" + p.code + "' LIMIT 1;"
                Dim dbReader As SQLiteDataReader = cmd.ExecuteReader()
                While dbReader.Read()
                    personDbId = dbReader.GetInt64(0)
                End While
            End Using
            Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                If personDbId < 0 Then
                    cmd.CommandText = "
BEGIN TRANSACTION;
INSERT INTO [person] ([group],[login],[code]) VALUES ('" + p.group + "', '" + p.login + "', '" + p.code + "');
SELECT last_insert_rowid();
COMMIT;"
                    personDbId = cmd.ExecuteScalar()
                End If
                For Each b As Booklet In p.booklets
                    If b.lastTS = 0 OrElse b.firstTS = 0 Then b.setTimestamps()

                    cmd.CommandText = "
select booklet.id, booklet.lastTs, booklet.firstTs, booklet.infoId, bookletInfo.name from [booklet] 
join [bookletInfo] on booklet.infoId = bookletInfo.id
where booklet.personId = @personId AND bookletInfo.name = @bookletName LIMIT 1"
                    cmd.Parameters.AddWithValue("@personId", personDbId)
                    cmd.Parameters.AddWithValue("@bookletName", b.id)
                    Dim dbReader As SQLiteDataReader = cmd.ExecuteReader()
                    Dim bookletDbId As Long = -1
                    Dim lastTs As Long = -1
                    Dim firstTs As Long = -1
                    Dim infoId As Long = -1
                    While dbReader.Read()
                        bookletDbId = dbReader.GetInt64(0)
                        lastTs = dbReader.GetInt64(1)
                        firstTs = dbReader.GetInt64(2)
                        infoId = dbReader.GetInt64(3)
                    End While
                    dbReader.Close()

                    If bookletDbId >= 0 AndAlso b.lastTS = lastTs AndAlso b.firstTS = b.firstTS Then
                        ignoredBookletCount += 1
                    Else
                        If bookletDbId >= 0 Then
                            cmd.CommandText = "
DELETE FROM [booklet] WHERE id=" + bookletDbId.ToString + ";"
                            cmd.ExecuteNonQuery()
                            updatedBookletCount += 1
                        Else
                            cmd.CommandText = "select id from [bookletInfo] where name = @bookletName LIMIT 1"
                            cmd.Parameters.AddWithValue("@bookletName", b.id)
                            dbReader = cmd.ExecuteReader()
                            While dbReader.Read()
                                infoId = dbReader.GetInt64(0)
                            End While
                            dbReader.Close()
                            If infoId < 0 Then
                                cmd.CommandText = "
BEGIN TRANSACTION;
INSERT INTO [bookletInfo] ([name]) VALUES ('" + b.id + "');
SELECT last_insert_rowid();
COMMIT;"
                                infoId = cmd.ExecuteScalar()
                            End If

                            addedBookletCount += 1
                        End If
                        cmd.CommandText = "
BEGIN TRANSACTION;
INSERT INTO [booklet] ([personId],[infoId],[firstTs],[lastTs]) VALUES (" + personDbId.ToString + ", " +
infoId.ToString + "," + b.firstTS.ToString + "," + b.lastTS.ToString + ");
SELECT last_insert_rowid();
COMMIT;"
                        bookletDbId = cmd.ExecuteScalar()
                        cmd.CommandText = ""
                        For Each bLog As LogEntry In b.logs
                            cmd.CommandText += "
INSERT INTO [bookletLog] ([bookletId],[key],[parameter],[ts]) VALUES (" +
    bookletDbId.ToString + ", '" + bLog.key + "','" + bLog.parameter + "'," + bLog.ts.ToString + ");"
                        Next
                        For Each s As Session In b.sessions
                            cmd.CommandText += "
INSERT INTO [session] ([bookletId],[browser],[os],[ts],[screen],[loadCompleteMS]) VALUES (" +
    bookletDbId.ToString + ", '" + s.browser + "','" + s.os + "'," + s.ts.ToString + ",'" + s.screen + "'," + s.loadCompleteMS.ToString + ");"
                        Next
                        If Not String.IsNullOrEmpty(cmd.CommandText) Then
                            cmd.CommandText = "BEGIN TRANSACTION;" + cmd.CommandText + "COMMIT;"
                            cmd.ExecuteNonQuery()
                        End If

                        For Each u As Unit In b.units
                            cmd.CommandText = "
BEGIN TRANSACTION;
INSERT INTO [unit] ([bookletId],[name],[alias]) VALUES (" + bookletDbId.ToString + ", '" + u.id + "', '" + u.alias + "');
SELECT last_insert_rowid();
COMMIT;"
                            Dim lastInsert_UnitId As Long = cmd.ExecuteScalar()

                            cmd.CommandText = ""
                            For Each ls As LastStateEntry In u.laststate
                                cmd.CommandText += "
INSERT INTO [unitLastState] ([unitId],[key],[value]) VALUES (" +
        lastInsert_UnitId.ToString + ", '" + ls.key + "','" + ls.value + "');"
                            Next
                            For Each l As LogEntry In u.logs
                                cmd.CommandText += "
INSERT INTO [unitLog] ([unitId],[key],[parameter],[ts]) VALUES (" +
        lastInsert_UnitId.ToString + ", '" + l.key + "','" + l.parameter + "'," + l.ts.ToString + ");"
                            Next
                            For Each r As ResponseChunkData In u.chunks
                                cmd.CommandText += "
INSERT INTO [chunk] ([unitId],[key],[type],[variables],[ts]) VALUES (" +
        lastInsert_UnitId.ToString + ", '" + r.id + "','" + r.type + "','" + String.Join(" ", r.variables) + "'," + r.ts.ToString + ");"
                            Next
                            If Not String.IsNullOrEmpty(cmd.CommandText) Then
                                cmd.CommandText = "BEGIN TRANSACTION;" + cmd.CommandText + "COMMIT;"
                                cmd.ExecuteNonQuery()
                            End If

                            For Each sf As SubForm In u.subforms
                                cmd.CommandText = "
BEGIN TRANSACTION;
INSERT INTO [subform] ([unitId],[key]) VALUES (" + lastInsert_UnitId.ToString + ", '" + sf.id + "');
SELECT last_insert_rowid();
COMMIT;"
                                Dim lastInsert_SubformId As Long = cmd.ExecuteScalar()
                                cmd.CommandText = ""
                                For Each resp As ResponseData In sf.responses
                                    cmd.CommandText += "
INSERT INTO [response] ([subformId],[variableId],[value],[status],[code],[score]) VALUES (" +
        lastInsert_SubformId.ToString + ", '" + resp.id + "','" + resp.value.Replace("'", "''") + "','" + resp.status + "'," +
                                        resp.code.ToString + "," + resp.score.ToString + ");"
                                Next
                                If Not String.IsNullOrEmpty(cmd.CommandText) Then
                                    cmd.CommandText = "BEGIN TRANSACTION;" + cmd.CommandText + "COMMIT;"
                                    cmd.ExecuteNonQuery()
                                End If
                            Next
                        Next
                    End If
                Next
            End Using
        End Using
    End Sub

    Public Function addBooklets(bookletSizes As Dictionary(Of String, Long)) As Integer
        Dim addedBookletCount As Integer = 0
        Using sqliteConnection As SQLiteConnection = GetOpenConnection(False)
            Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                Dim dbReader As SQLiteDataReader
                For Each b As KeyValuePair(Of String, Long) In bookletSizes
                    Dim infoId As Long = -1
                    cmd.CommandText = "select id from [bookletInfo] where name = @bookletName LIMIT 1"
                    cmd.Parameters.AddWithValue("@bookletName", b.Key)
                    dbReader = cmd.ExecuteReader()
                    While dbReader.Read()
                        infoId = dbReader.GetInt64(0)
                    End While
                    dbReader.Close()
                    cmd.Parameters.AddWithValue("@bookletSize", b.Value)
                    If infoId < 0 Then
                        cmd.CommandText = "INSERT INTO [bookletInfo] ([name],[size]) VALUES (@bookletName,@bookletSize);"
                        cmd.ExecuteScalar()
                        addedBookletCount += 1
                    Else
                        cmd.CommandText = "UPDATE [bookletInfo] SET [size]= @bookletSize where id = @bookletId"
                        cmd.Parameters.AddWithValue("@bookletId", infoId)
                        cmd.ExecuteScalar()
                    End If
                Next
            End Using
            Me.CloseConnection()
            Return addedBookletCount
        End Using
    End Function

    Function hasSubforms() As Boolean
        Dim firstSubformResponseDbId As Long = -1
        Using sqliteConnection As SQLiteConnection = GetOpenConnection(True)
            Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                cmd.CommandText = "SELECT [id],[key] FROM [subform] where key not in ('') limit 1;"
                Dim dbReader As SQLiteDataReader = cmd.ExecuteReader()
                While dbReader.Read()
                    firstSubformResponseDbId = dbReader.GetInt64(0)
                End While
                dbReader.Close()
            End Using
        End Using
        Return firstSubformResponseDbId >= 0
    End Function

    Public Function getVariableList(addSubformSuffix As Boolean) As List(Of String)
        Dim returnList As New List(Of String)
        Using sqliteConnection As SQLiteConnection = GetOpenConnection(True)
            Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                cmd.CommandText = "
select response.variableId,subform.key,response.subformId,subform.unitId,unit.id,unit.name,unit.alias from [response]
join [subform] on subform.id = response.subformId
join [unit] on unit.id = subform.unitId;"
                Dim dbReader As SQLiteDataReader = cmd.ExecuteReader()
                While dbReader.Read()
                    Dim unitName As String = dbReader.GetString(5)
                    Dim unitAlias As String = dbReader.GetString(6)
                    Dim variableId As String = IIf(String.IsNullOrEmpty(unitAlias), unitName, unitAlias) + dbReader.GetString(0)
                    Dim subformKey As String = dbReader.GetString(1)
                    If Not String.IsNullOrEmpty(subformKey) AndAlso addSubformSuffix Then variableId += "##" + subformKey
                    If Not returnList.Contains(variableId) Then returnList.Add(variableId)
                End While
                dbReader.Close()
            End Using
        End Using
        Return returnList
    End Function

    Public Function getPeopleList() As Dictionary(Of String, String)
        Dim returnList As New Dictionary(Of String, String)
        Using sqliteConnection As SQLiteConnection = GetOpenConnection(True)
            Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                cmd.CommandText = "select [id],[group],[login],[code] from [person];"
                Dim dbReader As SQLiteDataReader = cmd.ExecuteReader()
                While dbReader.Read()
                    Dim personDbId As Long = dbReader.GetInt64(0)
                    Dim personKey As String = dbReader.GetString(1) + dbReader.GetString(2) + dbReader.GetString(3)
                    If Not returnList.ContainsKey(personKey) Then returnList.Add(personKey, personDbId.ToString)
                End While
            End Using
        End Using
        Return returnList
    End Function

    Public Function getPersonResponses(dbIdString As String) As Person
        Dim dbPerson As Person = Nothing
        Using sqliteConnection As SQLiteConnection = GetOpenConnection(True)
            Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                cmd.CommandText = "select [group],[login],[code] from [person] where [id]=" + dbIdString + " LIMIT 1;"
                Dim dbReader As SQLiteDataReader = cmd.ExecuteReader()
                While dbReader.Read()
                    dbPerson = New Person(dbReader.GetString(0), dbReader.GetString(1), dbReader.GetString(2))
                End While
                dbReader.close()
                cmd.CommandText = "
select booklet.id,booklet.personId,unit.id,unit.bookletId,unit.name,unit.alias from [unit]
join [booklet] on booklet.id = unit.bookletId where booklet.personId = " + dbIdString + ";"
                dbReader = cmd.ExecuteReader()
                Dim dummyBooklet As New Booklet("XYZ")
                dbPerson.booklets.Add(dummyBooklet)
                Dim unitList As New Dictionary(Of String, Unit)
                While dbReader.Read()
                    Dim unitDbId As Long = dbReader.GetInt64(2)
                    unitList.Add(unitDbId.ToString, New Unit(dbReader.GetString(4), dbReader.GetString(5)))
                End While
                dbReader.Close()
                For Each dbUnit As KeyValuePair(Of String, Unit) In unitList
                    cmd.CommandText = "
select subform.id,subform.unitId,subform.key,response.variableId,response.value,response.status,response.code,response.score from [response]
join [subform] on subform.id = response.subformId where subform.unitId = " + dbUnit.Key + ";"
                    dbReader = cmd.ExecuteReader()
                    Dim subformList As New Dictionary(Of String, List(Of ResponseData))
                    While dbReader.Read()
                        Dim subformKey As String = dbReader.GetString(2)
                        If Not subformList.ContainsKey(subformKey) Then subformList.Add(subformKey, New List(Of ResponseData))
                        subformList.Item(subformKey).Add(New ResponseData(
                                                         dbReader.GetString(3), dbReader.GetString(4), dbReader.GetString(5)) With {
                                                         .code = dbReader.GetInt64(6), .score = dbReader.GetInt64(7)})
                    End While
                    dbReader.Close()
                    For Each sf As KeyValuePair(Of String, List(Of ResponseData)) In subformList
                        dbUnit.Value.subforms.Add(New SubForm() With {.id = sf.Key, .responses = sf.Value})
                    Next
                    dummyBooklet.units.Add(dbUnit.Value)
                Next
            End Using
        End Using
        Return dbPerson
    End Function
End Class
