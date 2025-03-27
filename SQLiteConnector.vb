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
    Public ReadOnly dbLastChanger As String
    Public ReadOnly dbLastChangedDateTime As String
    Public Sub New(dbFileName As String, create As Boolean)
        fileName = dbFileName
        Dim fileExists As Boolean = IO.File.Exists(fileName)
        If create AndAlso fileExists Then
            IO.File.Delete(fileName)
            fileExists = False
        End If
        Using sqliteConnection As SQLiteConnection = GetOpenConnection(False)
            If Not fileExists Then
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
, CONSTRAINT [FK_booklet_0_0] FOREIGN KEY ([personId]) REFERENCES [person] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
, CONSTRAINT [FK_booklet_1_0] FOREIGN KEY ([infoId]) REFERENCES [bookletInfo] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
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
CREATE TABLE [bookletLog] (
  [bookletId] bigint NOT NULL
, [key] text NOT NULL
, [parameter] text NULL
, [ts] bigint NULL
, CONSTRAINT [FK_bookletLog_0_0] FOREIGN KEY ([bookletId]) REFERENCES [booklet] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [unit] (
  [id] INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL
, [bookletId] bigint NOT NULL
, [name] text NOT NULL
, [alias] text NULL
, CONSTRAINT [FK_unit_0_0] FOREIGN KEY ([bookletId]) REFERENCES [booklet] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [unitLog] (
  [unitId] bigint NOT NULL
, [key] text NOT NULL
, [parameter] text NULL
, [ts] bigint NULL
, CONSTRAINT [FK_unitLog_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [unitLastState] (
  [unitId] bigint NOT NULL
, [key] text NOT NULL
, [value] text NULL
, CONSTRAINT [FK_unitLastState_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [chunk] (
  [unitId] bigint NOT NULL
, [key] text NOT NULL
, [type] text NULL
, [variables] text NULL
, [ts] bigint NULL
, CONSTRAINT [FK_chunk_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TABLE [response] (
  [unitId] bigint NOT NULL
, [variableId] text NOT NULL
, [status] text NOT NULL
, [value] text NULL
, [subform] text NULL
, [code] bigint DEFAULT (0) NOT NULL
, [score] bigint DEFAULT (0) NOT NULL
, CONSTRAINT [FK_response_0_0] FOREIGN KEY ([unitId]) REFERENCES [unit] ([id]) ON DELETE CASCADE ON UPDATE NO ACTION
);
CREATE TRIGGER [fki_booklet_personId_person_id] BEFORE Insert ON [booklet] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table booklet violates foreign key constraint FK_booklet_0_0') WHERE (SELECT id FROM person WHERE  id = NEW.personId) IS NULL; END;
CREATE TRIGGER [fku_booklet_personId_person_id] BEFORE Update ON [booklet] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table booklet violates foreign key constraint FK_booklet_0_0') WHERE (SELECT id FROM person WHERE  id = NEW.personId) IS NULL; END;
CREATE TRIGGER [fki_booklet_infoId_bookletInfo_id] BEFORE Insert ON [booklet] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table booklet violates foreign key constraint FK_booklet_1_0') WHERE (SELECT id FROM bookletInfo WHERE  id = NEW.infoId) IS NULL; END;
CREATE TRIGGER [fku_booklet_infoId_bookletInfo_id] BEFORE Update ON [booklet] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table booklet violates foreign key constraint FK_booklet_1_0') WHERE (SELECT id FROM bookletInfo WHERE  id = NEW.infoId) IS NULL; END;
CREATE TRIGGER [fki_session_bookletId_booklet_id] BEFORE Insert ON [session] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table session violates foreign key constraint FK_session_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fku_session_bookletId_booklet_id] BEFORE Update ON [session] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table session violates foreign key constraint FK_session_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fki_bookletLog_bookletId_booklet_id] BEFORE Insert ON [bookletLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table bookletLog violates foreign key constraint FK_bookletLog_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fku_bookletLog_bookletId_booklet_id] BEFORE Update ON [bookletLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table bookletLog violates foreign key constraint FK_bookletLog_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fki_unit_bookletId_booklet_id] BEFORE Insert ON [unit] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table unit violates foreign key constraint FK_unit_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fku_unit_bookletId_booklet_id] BEFORE Update ON [unit] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table unit violates foreign key constraint FK_unit_0_0') WHERE (SELECT id FROM booklet WHERE  id = NEW.bookletId) IS NULL; END;
CREATE TRIGGER [fki_unitLog_unitId_unit_id] BEFORE Insert ON [unitLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table unitLog violates foreign key constraint FK_unitLog_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_unitLog_unitId_unit_id] BEFORE Update ON [unitLog] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table unitLog violates foreign key constraint FK_unitLog_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_unitLastState_unitId_unit_id] BEFORE Insert ON [unitLastState] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table unitLastState violates foreign key constraint FK_unitLastState_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_unitLastState_unitId_unit_id] BEFORE Update ON [unitLastState] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table unitLastState violates foreign key constraint FK_unitLastState_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_chunk_unitId_unit_id] BEFORE Insert ON [chunk] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table chunk violates foreign key constraint FK_chunk_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_chunk_unitId_unit_id] BEFORE Update ON [chunk] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table chunk violates foreign key constraint FK_chunk_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fki_response_unitId_unit_id] BEFORE Insert ON [response] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Insert on table response violates foreign key constraint FK_response_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
CREATE TRIGGER [fku_response_unitId_unit_id] BEFORE Update ON [response] FOR EACH ROW BEGIN SELECT RAISE(ROLLBACK, 'Update on table response violates foreign key constraint FK_response_0_0') WHERE (SELECT id FROM unit WHERE  id = NEW.unitId) IS NULL; END;
COMMIT;"
                    cmd.ExecuteNonQuery()
                End Using
                dbVersion = 1
                Dim now As DateTime = DateTime.Now
                dbCreatedDateTime = now.ToShortDateString + " " + now.ToShortTimeString
                dbCreator = ADFactory.GetMyNameLong
                dbLastChangedDateTime = dbCreatedDateTime
                dbLastChanger = dbCreator
                Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                    cmd.CommandText = "
BEGIN TRANSACTION;
INSERT INTO [db_info] ([key],[value]) VALUES ('name', 'IQB-Testcenter-Output');
INSERT INTO [db_info] ([key],[value]) VALUES ('version', '" + dbVersion.ToString + "');
INSERT INTO [db_info] ([key],[value]) VALUES ('created_By', '" + dbCreator + "');
INSERT INTO [db_info] ([key],[value]) VALUES ('created_DateTime', '" + dbCreatedDateTime + "');
INSERT INTO [db_info] ([key],[value]) VALUES ('lastchanged_By', '" + dbLastChanger + "');
INSERT INTO [db_info] ([key],[value]) VALUES ('lastchanged_DateTime', '" + dbLastChangedDateTime + "');
INSERT INTO [db_info] ([key],[value]) VALUES ('number_of_people', '0');
INSERT INTO [db_info] ([key],[value]) VALUES ('number_of_responses', '0');
COMMIT;"
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
                            Case "version" : dbVersion = Long.Parse(value)
                            Case "created_By" : dbCreator = value
                            Case "created_DateTime" : dbCreatedDateTime = value
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

    Public Function WriteDbInfoData(closeConnection As Boolean) As String
        Dim returnText As String = ""
        Using sqliteConnection As SQLiteConnection = GetOpenConnection(False)
            Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                cmd.CommandText = "
PRAGMA journal_mode=OFF;
PRAGMA synchronous=OFF;"
                cmd.ExecuteNonQuery()

                cmd.CommandText = "SELECT COUNT(*) FROM [person];"
                Dim personCount As Long = cmd.ExecuteScalar()
                cmd.CommandText = "SELECT COUNT(*) FROM [response];"
                Dim responseCount As Long = cmd.ExecuteScalar()

                cmd.CommandText = "BEGIN TRANSACTION;"
                Dim deCulture = CultureInfo.CreateSpecificCulture("de-DE")
                cmd.CommandText += "UPDATE [db_info] SET [value]= '" + personCount.ToString("N0", deCulture) + "' where key = 'number_of_people';"
                cmd.CommandText += "UPDATE [db_info] SET [value]= '" + responseCount.ToString("N0", deCulture) + "' where key = 'number_of_responses';"
                Dim now As DateTime = DateTime.Now
                cmd.CommandText += "UPDATE [db_info] SET [value]= '" + now.ToShortDateString + " " + now.ToShortTimeString + "' where key = 'lastchanged_DateTime';"
                cmd.CommandText += "UPDATE [db_info] SET [value]= '" + ADFactory.GetMyNameLong + "' where key = 'lastchanged_By';"
                cmd.CommandText += "COMMIT;"
                cmd.ExecuteScalar()
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
                cmd.CommandText = "
PRAGMA journal_mode=OFF;
PRAGMA synchronous=OFF;"
                cmd.ExecuteNonQuery()
                cmd.CommandText = "
select [id] from [person] where [group]='" + p.group +
                    "' and [login]='" + p.login + "' and [code]='" + p.code + "' LIMIT 1;"
                Dim dbReader As SQLiteDataReader = cmd.ExecuteReader()
                While dbReader.Read()
                    personDbId = dbReader.GetInt64(0)
                End While
                dbReader.Close()
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
                    dbReader = cmd.ExecuteReader()
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
                                cmd.CommandText = ""
                                For Each resp As ResponseData In sf.responses
                                    cmd.CommandText += "
INSERT INTO [response] ([unitId],[subform],[variableId],[value],[status],[code],[score]) VALUES (" + lastInsert_UnitId.ToString + ",'" +
                                    IIf(String.IsNullOrEmpty(sf.id), "", sf.id) + "', '" + resp.id + "','" + resp.value.Replace("'", "''") + "','" + resp.status + "'," +
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
                    cmd.CommandText = "
PRAGMA journal_mode=OFF;
PRAGMA synchronous=OFF;"
                    cmd.ExecuteNonQuery()

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
                cmd.CommandText = "SELECT 1 FROM [response] where subform not in ('') limit 1;"
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
select response.variableId,response.subform,response.unitId,unit.id,unit.name,unit.alias from [response]
join [unit] on unit.id = response.unitId;"
                Dim dbReader As SQLiteDataReader = cmd.ExecuteReader()
                While dbReader.Read()
                    Dim unitName As String = dbReader.GetString(4)
                    Dim unitAlias As String = dbReader.GetString(5)
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
select response.unitId,response.subform,response.variableId,response.value,response.status,response.code,response.score
from [response] where response.unitId = " + dbUnit.Key + ";"
                    dbReader = cmd.ExecuteReader()
                    Dim subformList As New Dictionary(Of String, List(Of ResponseData))
                    While dbReader.Read()
                        Dim subformKey As String = dbReader.GetString(1)
                        If Not subformList.ContainsKey(subformKey) Then subformList.Add(subformKey, New List(Of ResponseData))
                        subformList.Item(subformKey).Add(New ResponseData(
                                                         dbReader.GetString(2), dbReader.GetString(3), dbReader.GetString(4)) With {
                                                         .code = dbReader.GetInt64(5), .score = dbReader.GetInt64(6)})
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

    Public Function getSessionReports(personDbId As Long) As List(Of SessionReport)
        Dim mySessionReports As New List(Of SessionReport)
        Using sqliteConnection As SQLiteConnection = GetOpenConnection(True)
            Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                Dim bNames As New Dictionary(Of Long, String)
                Dim bSizes As New Dictionary(Of Long, Long)
                cmd.CommandText = "
select booklet.id,bookletInfo.name,bookletInfo.size from [booklet]
join [bookletInfo] on bookletInfo.id = booklet.infoId where booklet.personId=" + personDbId.ToString + ";"
                Dim dbReader As SQLiteDataReader = cmd.ExecuteReader()
                While dbReader.Read()
                    Dim bookletDbId As Long = dbReader.GetInt64(0)
                    bNames.Add(bookletDbId, dbReader.GetString(1))
                    bSizes.Add(bookletDbId, dbReader.GetInt64(2))
                End While
                dbReader.Close()
                For Each b As KeyValuePair(Of Long, String) In bNames
                    Dim unitsLastChanged As Dictionary(Of String, Long) = getUnitsLastChanged(b.Key)
                    If unitsLastChanged.Count > 0 Then
                        Dim bookletSize As Long = bSizes.Item(b.Key)
                        Dim localSessionReports As New List(Of SessionReport)
                        cmd.CommandText = "
select [browser],[os],[screen],[ts],[loadCompleteMS] from [session] where [bookletId]=" + b.Key.ToString + ";"
                        dbReader = cmd.ExecuteReader()
                        While dbReader.Read()
                            Dim rep As New SessionReport
                            rep.browser = dbReader.GetString(0)
                            rep.os = dbReader.GetString(1)
                            rep.screen = dbReader.GetString(2)
                            rep.ts = dbReader.GetInt64(3)
                            rep.loadCompleteMS = dbReader.GetInt64(4)
                            rep.sessionStartTs = rep.ts - rep.loadCompleteMS
                            rep.booklet = b.Value
                            If rep.loadCompleteMS > 0 Then
                                rep.contentLoadSpeed = bookletSize / rep.loadCompleteMS
                            Else
                                rep.contentLoadSpeed = 0
                            End If
                            localSessionReports.Add(rep)
                        End While
                        dbReader.Close()
                        If localSessionReports.Count > 0 Then
                            If localSessionReports.Count = 1 Then
                                Dim rep As SessionReport = localSessionReports.First
                                rep.sessionNumber = 1
                                rep.firstUnitTS = unitsLastChanged.Values.Min
                                rep.lastUnitTS = unitsLastChanged.Values.Max
                                rep.unitsWithResponse = (From u As KeyValuePair(Of String, Long) In unitsLastChanged
                                                         Order By u.Value Select u.Key).ToList
                                mySessionReports.Add(rep)
                            Else
                                Dim sessionCount As Integer = localSessionReports.Count
                                Dim lastSessionStart As Long = Long.MaxValue
                                Dim localSessionReportsReverseOrder As New List(Of SessionReport)
                                For Each rep As SessionReport In (From r As SessionReport In localSessionReports Order By r.sessionStartTs).Reverse
                                    rep.sessionNumber = sessionCount
                                    sessionCount -= 1
                                    Dim myUnits As Dictionary(Of String, Long) = (From u As KeyValuePair(Of String, Long) In unitsLastChanged
                                                                                  Where u.Value < lastSessionStart AndAlso
                                                                                      u.Value >= rep.sessionStartTs).ToDictionary(Function(a) a.Key, Function(a) a.Value)
                                    If myUnits.Count > 0 Then
                                        rep.firstUnitTS = myUnits.Values.Min
                                        rep.lastUnitTS = myUnits.Values.Max
                                        rep.unitsWithResponse = (From u As KeyValuePair(Of String, Long) In myUnits
                                                                 Order By u.Value Select u.Key).ToList
                                    Else
                                        rep.firstUnitTS = 0
                                        rep.lastUnitTS = 0
                                        rep.unitsWithResponse = New List(Of String)
                                    End If
                                    lastSessionStart = rep.sessionStartTs
                                    localSessionReportsReverseOrder.Add(rep)
                                Next
                                localSessionReports.AddRange((From rep As SessionReport In localSessionReportsReverseOrder).Reverse)
                            End If
                        End If
                    End If
                Next
            End Using
        End Using

        Return mySessionReports
    End Function

    Public Function getUnitsLastChanged(bookletId As Long) As Dictionary(Of String, Long)
        Dim unitsLastChanged As New Dictionary(Of String, Long)
        Using sqliteConnection As SQLiteConnection = GetOpenConnection(True)
            Using cmd As SQLiteCommand = sqliteConnection.CreateCommand()
                cmd.CommandText = "
select unit.id,unit.bookletId,unit.name,unit.alias from [unit] where unit.bookletId=" + bookletId.ToString + ";"
                Dim dbReader As SQLiteDataReader = cmd.ExecuteReader()
                Dim allUnits As New Dictionary(Of Long, String)
                While dbReader.Read()
                    Dim unitDbId As Long = dbReader.GetInt64(0)
                    Dim unitName As String = dbReader.GetString(2)
                    Dim unitAlias As String = dbReader.GetString(3)
                    allUnits.Add(unitDbId, IIf(String.IsNullOrEmpty(unitAlias), unitName, unitAlias))
                End While
                dbReader.Close()
                Dim unitVariablesWithValues As New Dictionary(Of Long, String)
                For Each u As KeyValuePair(Of Long, String) In allUnits
                    Dim varId As String = Nothing
                    cmd.CommandText = "
select [variableId] from [response]
where not [status] in ('DISPLAYED','UNSET','NOT_REACHED') and unitId=" + u.Key.ToString + " LIMIT 1;"
                    dbReader = cmd.ExecuteReader()
                    While dbReader.Read()
                        varId = dbReader.GetString(0)
                    End While
                    dbReader.Close()
                    If Not String.IsNullOrEmpty(varId) Then unitVariablesWithValues.Add(u.Key, u.Value)
                Next
                For Each u As KeyValuePair(Of Long, String) In unitVariablesWithValues
                    cmd.CommandText = "select [unitId],[ts] from [chunk] where unitId=" + u.Key.ToString + ";"
                    dbReader = cmd.ExecuteReader()
                    Dim lastChangedTs As Long = 0
                    While dbReader.Read()
                        Dim ts As Long = dbReader.GetInt64(1)
                        If lastChangedTs < ts Then lastChangedTs = ts
                    End While
                    dbReader.Close()
                    If lastChangedTs > 0 Then
                        If unitsLastChanged.ContainsKey(u.Value) Then
                            Debug.Print("two entries for unit for same booklet: booklet-id " + bookletId.ToString +
                                        ", unit " + u.Value + "; take newest")
                            Dim oldTS As Long = unitsLastChanged.Item(u.Value)
                            If oldTS < lastChangedTs Then unitsLastChanged.Item(u.Value) = lastChangedTs
                        Else
                            unitsLastChanged.Add(u.Value, lastChangedTs)
                        End If
                    End If
                Next
            End Using
        End Using
        Return unitsLastChanged
    End Function
End Class
