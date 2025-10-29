' Arquivo: DatabaseManager.vb (VERSÃO COM TABELA ANNOUNCEMENTS)
Imports System.Data.SQLite
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Globalization ' Necessário para DateTime parsing

Public Module DatabaseManager
    Private ReadOnly DbFolder As String = "Z:\Engenharia\iLogic\CitrotecAddinData"
    Private ReadOnly DbFile As String = "usuarios.db"
    Private ReadOnly DbPath As String = Path.Combine(DbFolder, DbFile)
    Private ReadOnly ConnectionString As String = $"Data Source={DbPath};Version=3;FailIfMissing=False;"

    Public Sub InitializeDatabase()
        If Not Directory.Exists(DbFolder) Then
            Directory.CreateDirectory(DbFolder)
        End If

        Dim dbInitialized As Boolean = File.Exists(DbPath) AndAlso New FileInfo(DbPath).Length > 0

        If Not dbInitialized Then
            If File.Exists(DbPath) Then File.Delete(DbPath)

            Using conn As New SQLiteConnection(ConnectionString)
                conn.Open()
                ' Cria Tabela Usuarios
                Dim sqlUsers As String = "CREATE TABLE Usuarios (" &
                                    "Id INTEGER PRIMARY KEY AUTOINCREMENT, " &
                                    "NomeUsuario TEXT UNIQUE NOT NULL, " &
                                    "NomeCompleto TEXT, " &
                                    "HashSenha TEXT NOT NULL, " &
                                    "Avatar TEXT, " &
                                    "Nivel INTEGER DEFAULT 1, " &
                                    "Aprovado INTEGER DEFAULT 0, " &
                                    "Sexo TEXT, " &
                                    "UserTag TEXT, " &
                                    "ResetSenhaObrigatorio INTEGER DEFAULT 0);"
                Using command As New SQLiteCommand(sqlUsers, conn)
                    command.ExecuteNonQuery()
                End Using

                ' Cria usuário admin padrão
                Dim adminHash = HashPassword("admin")
                Dim sqlAdmin As String = "INSERT INTO Usuarios (NomeUsuario, NomeCompleto, HashSenha, Avatar, Nivel, Aprovado, Sexo, UserTag) VALUES ('admin', 'Administrador', @hash, '', 3, 1, 'M', 'A.D.M.')"
                Using command As New SQLiteCommand(sqlAdmin, conn)
                    command.Parameters.AddWithValue("@hash", adminHash)
                    command.ExecuteNonQuery()
                End Using

                ' ***** NOVO: Cria Tabela Announcements *****
                Dim sqlAnnouncements As String = "CREATE TABLE Announcements (" &
                                              "Id INTEGER PRIMARY KEY AUTOINCREMENT, " &
                                              "AdminUserName TEXT NOT NULL, " &
                                              "Content TEXT NOT NULL, " &
                                              "Timestamp TEXT NOT NULL);" ' Usando TEXT para data/hora (ISO 8601)
                Using command As New SQLiteCommand(sqlAnnouncements, conn)
                    command.ExecuteNonQuery()
                End Using
                ' ***** FIM NOVO *****

            End Using
        Else
            UpdateDatabaseSchema()
        End If
    End Sub

    Private Sub UpdateDatabaseSchema()
        Using conn As New SQLiteConnection(ConnectionString)
            conn.Open()

            ' --- Verifica colunas da tabela Usuarios ---
            Dim sqlCheckUsers As String = "PRAGMA table_info(Usuarios);"
            Dim userColumns As New List(Of String)
            Using cmdCheck As New SQLiteCommand(sqlCheckUsers, conn)
                Using reader As SQLiteDataReader = cmdCheck.ExecuteReader()
                    While reader.Read()
                        userColumns.Add(reader("name").ToString().ToLower())
                    End While
                End Using
            End Using
            ' Adiciona colunas faltantes em Usuarios (se necessário)
            If Not userColumns.Contains("nomecompleto") Then AddColumn(conn, "Usuarios", "NomeCompleto TEXT")
            If Not userColumns.Contains("sexo") Then AddColumn(conn, "Usuarios", "Sexo TEXT")
            If Not userColumns.Contains("usertag") Then AddColumn(conn, "Usuarios", "UserTag TEXT")
            If Not userColumns.Contains("resetsenhaobrigatorio") Then AddColumn(conn, "Usuarios", "ResetSenhaObrigatorio INTEGER DEFAULT 0")


            ' --- ***** NOVO: Verifica e cria tabela Announcements se faltar ***** ---
            Dim sqlCheckAnnouncements As String = "SELECT name FROM sqlite_master WHERE type='table' AND name='Announcements';"
            Dim announcementTableExists As Boolean = False
            Using cmdCheck As New SQLiteCommand(sqlCheckAnnouncements, conn)
                Using reader As SQLiteDataReader = cmdCheck.ExecuteReader()
                    announcementTableExists = reader.HasRows
                End Using
            End Using

            If Not announcementTableExists Then
                Dim sqlCreateAnnouncements As String = "CREATE TABLE Announcements (" &
                                                      "Id INTEGER PRIMARY KEY AUTOINCREMENT, " &
                                                      "AdminUserName TEXT NOT NULL, " &
                                                      "Content TEXT NOT NULL, " &
                                                      "Timestamp TEXT NOT NULL);"
                Using command As New SQLiteCommand(sqlCreateAnnouncements, conn)
                    command.ExecuteNonQuery()
                End Using
            End If
            ' --- ***** FIM NOVO ***** ---

        End Using
    End Sub

    ' Função auxiliar para adicionar colunas
    Private Sub AddColumn(conn As SQLiteConnection, tableName As String, columnDefinition As String)
        Dim sqlAlter As String = $"ALTER TABLE {tableName} ADD COLUMN {columnDefinition};"
        Using cmdAlter As New SQLiteCommand(sqlAlter, conn)
            cmdAlter.ExecuteNonQuery()
        End Using
    End Sub

    ' --- Funções existentes (HashPassword, GetUserByUsername, AuthenticateUser, AddUser, GetAllUsers, UpdateUser, DeleteUser, etc.) ---
    ' Mantenha todas as suas funções existentes aqui...
    Public Function HashPassword(password As String) As String
        Using sha256 As SHA256 = SHA256.Create()
            Dim bytes As Byte() = sha256.ComputeHash(Encoding.UTF8.GetBytes(password))
            Dim builder As New StringBuilder()
            For i As Integer = 0 To bytes.Length - 1
                builder.Append(bytes(i).ToString("x2"))
            Next
            Return builder.ToString()
        End Using
    End Function

    Public Function GetUserByUsername(username As String) As User
        Using conn As New SQLiteConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "SELECT * FROM Usuarios WHERE LOWER(NomeUsuario) = LOWER(@user)"
            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@user", username)
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        Return New User With {
                            .Id = Convert.ToInt32(reader("Id")),
                            .NomeUsuario = reader("NomeUsuario").ToString(),
                            .NomeCompleto = If(reader.IsDBNull(reader.GetOrdinal("NomeCompleto")), "", reader("NomeCompleto").ToString()),
                            .Avatar = If(reader.IsDBNull(reader.GetOrdinal("Avatar")), "", reader("Avatar").ToString()),
                            .Nivel = Convert.ToInt32(reader("Nivel")),
                            .Aprovado = Convert.ToBoolean(reader("Aprovado")),
                            .Sexo = If(reader.IsDBNull(reader.GetOrdinal("Sexo")), "", reader("Sexo").ToString()),
                            .UserTag = If(reader.IsDBNull(reader.GetOrdinal("UserTag")), "", reader("UserTag").ToString()),
                            .ResetSenhaObrigatorio = Convert.ToBoolean(reader("ResetSenhaObrigatorio"))
                        }
                    Else
                        Return Nothing
                    End If
                End Using
            End Using
        End Using
    End Function

    Public Function AuthenticateUser(username As String, password As String) As User
        Dim hash = HashPassword(password)
        Using conn As New SQLiteConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "SELECT * FROM Usuarios WHERE LOWER(NomeUsuario) = LOWER(@user) AND HashSenha = @hash AND Aprovado = 1"
            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@user", username)
                command.Parameters.AddWithValue("@hash", hash)
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        Return New User With {
                            .Id = Convert.ToInt32(reader("Id")),
                            .NomeUsuario = reader("NomeUsuario").ToString(),
                            .NomeCompleto = If(reader.IsDBNull(reader.GetOrdinal("NomeCompleto")), "", reader("NomeCompleto").ToString()),
                            .Avatar = If(reader.IsDBNull(reader.GetOrdinal("Avatar")), "", reader("Avatar").ToString()),
                            .Nivel = Convert.ToInt32(reader("Nivel")),
                            .Aprovado = Convert.ToBoolean(reader("Aprovado")),
                            .Sexo = If(reader.IsDBNull(reader.GetOrdinal("Sexo")), "", reader("Sexo").ToString()),
                            .UserTag = If(reader.IsDBNull(reader.GetOrdinal("UserTag")), "", reader("UserTag").ToString()),
                            .ResetSenhaObrigatorio = Convert.ToBoolean(reader("ResetSenhaObrigatorio"))
                        }
                    Else
                        Return Nothing
                    End If
                End Using
            End Using
        End Using
    End Function

    Public Function AddUser(username As String, nomeCompleto As String, password As String, sexo As String, userTag As String, Optional avatar As String = "") As Boolean
        Dim hash = HashPassword(password)
        Using conn As New SQLiteConnection(ConnectionString)
            Try
                conn.Open()
                Dim sql As String = "INSERT INTO Usuarios (NomeUsuario, NomeCompleto, HashSenha, Avatar, Nivel, Aprovado, Sexo, UserTag) VALUES (@user, @nome, @hash, @avatar, 1, 1, @sexo, @userTag)" ' Nível padrão 1, Aprovado padrão 1 (pode mudar se precisar de aprovação)
                Using command As New SQLiteCommand(sql, conn)
                    command.Parameters.AddWithValue("@user", username)
                    command.Parameters.AddWithValue("@nome", nomeCompleto)
                    command.Parameters.AddWithValue("@hash", hash)
                    command.Parameters.AddWithValue("@avatar", avatar)
                    command.Parameters.AddWithValue("@sexo", sexo)
                    command.Parameters.AddWithValue("@userTag", userTag)
                    command.ExecuteNonQuery()
                End Using
                Return True
            Catch ex As SQLiteException
                Return False ' Provavelmente usuário duplicado
            End Try
        End Using
    End Function

    Public Function GetAllUsers() As List(Of User)
        Dim users As New List(Of User)()
        Using conn As New SQLiteConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "SELECT * FROM Usuarios"
            Using command As New SQLiteCommand(sql, conn)
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    While reader.Read()
                        users.Add(New User With {
                            .Id = Convert.ToInt32(reader("Id")),
                            .NomeUsuario = reader("NomeUsuario").ToString(),
                            .NomeCompleto = If(reader.IsDBNull(reader.GetOrdinal("NomeCompleto")), "", reader("NomeCompleto").ToString()),
                            .Avatar = If(reader.IsDBNull(reader.GetOrdinal("Avatar")), "", reader("Avatar").ToString()),
                            .Nivel = Convert.ToInt32(reader("Nivel")),
                            .Aprovado = Convert.ToBoolean(reader("Aprovado")),
                            .Sexo = If(reader.IsDBNull(reader.GetOrdinal("Sexo")), "", reader("Sexo").ToString()),
                            .UserTag = If(reader.IsDBNull(reader.GetOrdinal("UserTag")), "", reader("UserTag").ToString()),
                            .ResetSenhaObrigatorio = Convert.ToBoolean(reader("ResetSenhaObrigatorio"))
                        })
                    End While
                End Using
            End Using
        End Using
        Return users
    End Function

    Public Sub UpdateUser(user As User)
        Using conn As New SQLiteConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "UPDATE Usuarios SET NomeUsuario = @user, NomeCompleto = @nome, Avatar = @avatar, Nivel = @nivel, Aprovado = @aprovado, Sexo = @sexo, UserTag = @userTag, ResetSenhaObrigatorio = @reset WHERE Id = @id"
            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@user", user.NomeUsuario)
                command.Parameters.AddWithValue("@nome", user.NomeCompleto)
                command.Parameters.AddWithValue("@avatar", user.Avatar)
                command.Parameters.AddWithValue("@nivel", user.Nivel)
                command.Parameters.AddWithValue("@aprovado", user.Aprovado)
                command.Parameters.AddWithValue("@sexo", user.Sexo)
                command.Parameters.AddWithValue("@userTag", user.UserTag)
                command.Parameters.AddWithValue("@reset", user.ResetSenhaObrigatorio)
                command.Parameters.AddWithValue("@id", user.Id)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub DeleteUser(userId As Integer)
        Using conn As New SQLiteConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "DELETE FROM Usuarios WHERE Id = @id"
            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@id", userId)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub FlagUserForPasswordReset(userId As Integer)
        Using conn As New SQLiteConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "UPDATE Usuarios SET ResetSenhaObrigatorio = 1 WHERE Id = @id"
            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@id", userId)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function UpdateUserPassword(userId As Integer, newPassword As String) As Boolean
        Dim newHash = HashPassword(newPassword)
        Using conn As New SQLiteConnection(ConnectionString)
            Try
                conn.Open()
                Dim sql As String = "UPDATE Usuarios SET HashSenha = @hash, ResetSenhaObrigatorio = 0 WHERE Id = @id"
                Using command As New SQLiteCommand(sql, conn)
                    command.Parameters.AddWithValue("@hash", newHash)
                    command.Parameters.AddWithValue("@id", userId)
                    command.ExecuteNonQuery()
                End Using
                Return True
            Catch ex As Exception
                Return False
            End Try
        End Using
    End Function

    ' --- ***** NOVO: Funções para Announcements ***** ---

    ' Adiciona um novo anúncio ao banco
    Public Sub AddAnnouncement(adminUserName As String, content As String)
        Using conn As New SQLiteConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "INSERT INTO Announcements (AdminUserName, Content, Timestamp) VALUES (@admin, @content, @timestamp)"
            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@admin", adminUserName)
                command.Parameters.AddWithValue("@content", content)
                ' Salva a data/hora no formato ISO 8601 (YYYY-MM-DD HH:MM:SS) recomendado para SQLite
                command.Parameters.AddWithValue("@timestamp", DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss"))
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub DeleteAnnouncement(announcementId As Integer)
        Using conn As New SQLiteConnection(ConnectionString)
            conn.Open()
            Dim sql As String = "DELETE FROM Announcements WHERE Id = @id"
            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@id", announcementId)
                command.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function GetAllAnnouncements() As List(Of Announcement)
        Dim announcements As New List(Of Announcement)()
        Using conn As New SQLiteConnection(ConnectionString)
            conn.Open()
            ' Seleciona todos os anúncios, ordenados do mais recente para o mais antigo
            Dim sql As String = "SELECT Id, AdminUserName, Content, Timestamp FROM Announcements ORDER BY Timestamp DESC"
            Using command As New SQLiteCommand(sql, conn)
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim timestampStr = reader("Timestamp").ToString()
                        Dim timestampDt As DateTime
                        ' Tenta converter a string do banco para DateTime
                        If DateTime.TryParseExact(timestampStr, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal, timestampDt) Then
                            announcements.Add(New Announcement With {
                                .Id = Convert.ToInt32(reader("Id")),
                                .AdminUserName = reader("AdminUserName").ToString(),
                                .Content = reader("Content").ToString(),
                                .Timestamp = timestampDt.ToLocalTime() ' Converte para hora local
                            })
                        End If
                    End While
                End Using
            End Using
        End Using
        Return announcements
    End Function

    ' Busca anúncios dos últimos 7 dias
    Public Function GetRecentAnnouncements() As List(Of Announcement)
        Dim announcements As New List(Of Announcement)()
        Using conn As New SQLiteConnection(ConnectionString)
            conn.Open()
            ' Seleciona anúncios onde o Timestamp é maior ou igual à data/hora de 7 dias atrás
            Dim sql As String = "SELECT Id, AdminUserName, Content, Timestamp FROM Announcements WHERE Timestamp >= date('now', '-7 days') ORDER BY Timestamp DESC"
            Using command As New SQLiteCommand(sql, conn)
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    While reader.Read()
                        Dim timestampStr = reader("Timestamp").ToString()
                        Dim timestampDt As DateTime
                        ' Tenta converter a string do banco para DateTime (assume formato ISO 8601 UTC)
                        If DateTime.TryParseExact(timestampStr, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal Or DateTimeStyles.AdjustToUniversal, timestampDt) Then
                            announcements.Add(New Announcement With {
                                .Id = Convert.ToInt32(reader("Id")),
                                .AdminUserName = reader("AdminUserName").ToString(),
                                .Content = reader("Content").ToString(),
                                .Timestamp = timestampDt.ToLocalTime() ' Converte de UTC para hora local
                            })
                        End If
                    End While
                End Using
            End Using
        End Using
        Return announcements
    End Function

    ' --- ***** FIM NOVO ***** ---

End Module