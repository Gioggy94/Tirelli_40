Imports System.Data.SqlClient

Public Class MC_DatabaseService

    Private Function GetConnection() As SqlConnection
        Return New SqlConnection(Homepage.sap_tirelli)
    End Function

    ' ──────────────────────────────────────────────
    ' MACCHINE
    ' ──────────────────────────────────────────────

    Public Function GetMacchine(soloAttive As Boolean) As List(Of MC_Macchina)
        Dim lista As New List(Of MC_Macchina)
        Dim sql = "SELECT ID,Matricola,NomeMacchina,Modello,TipoMacchina," &
                  "ClienteFinale,AnnoCostruzione,LinguaCodice,Attiva,Note," &
                  "DataCreazione,DataModifica " &
                  "FROM [Tirelli_40].dbo.Macchine " &
                  If(soloAttive, "WHERE Attiva=1 ", "") &
                  "ORDER BY Matricola"
        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            cn.Open()
            Using rd = cmd.ExecuteReader()
                While rd.Read()
                    lista.Add(ReadMacchina(rd))
                End While
            End Using
        End Using
        Return lista
    End Function

    Public Function GetMacchinaByMatricola(matricola As String) As MC_Macchina
        Dim sql = "SELECT ID,Matricola,NomeMacchina,Modello,TipoMacchina," &
                  "ClienteFinale,AnnoCostruzione,LinguaCodice,Attiva,Note," &
                  "DataCreazione,DataModifica " &
                  "FROM [Tirelli_40].dbo.Macchine WHERE Matricola=@M"
        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            cmd.Parameters.AddWithValue("@M", matricola)
            cn.Open()
            Using rd = cmd.ExecuteReader()
                If rd.Read() Then Return ReadMacchina(rd)
            End Using
        End Using
        Return Nothing
    End Function

    Public Function SalvaMacchina(m As MC_Macchina) As Integer
        If m.ID = 0 Then Return InsertMacchina(m)
        UpdateMacchina(m)
        Return m.ID
    End Function

    Private Function InsertMacchina(m As MC_Macchina) As Integer
        Dim sql = "INSERT INTO [Tirelli_40].dbo.Macchine " &
                  "(Matricola,NomeMacchina,Modello,TipoMacchina,ClienteFinale," &
                  "AnnoCostruzione,LinguaCodice,Attiva,Note) " &
                  "VALUES (@Mat,@Nome,@Mod,@Tipo,@Cli,@Anno,@Lng,@Att,@Note);" &
                  "SELECT SCOPE_IDENTITY();"
        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            AddMacchinaParams(cmd, m)
            cn.Open()
            Return CInt(cmd.ExecuteScalar())
        End Using
    End Function

    Private Sub UpdateMacchina(m As MC_Macchina)
        Dim sql = "UPDATE [Tirelli_40].dbo.Macchine SET " &
                  "Matricola=@Mat,NomeMacchina=@Nome,Modello=@Mod," &
                  "TipoMacchina=@Tipo,ClienteFinale=@Cli," &
                  "AnnoCostruzione=@Anno,LinguaCodice=@Lng," &
                  "Attiva=@Att,Note=@Note,DataModifica=GETDATE() " &
                  "WHERE ID=@ID"
        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            AddMacchinaParams(cmd, m)
            cmd.Parameters.AddWithValue("@ID", m.ID)
            cn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub EliminaMacchina(id As Integer)
        Using cn = GetConnection(),
              cmd As New SqlCommand("DELETE FROM [Tirelli_40].dbo.Macchine WHERE ID=@ID", cn)
            cmd.Parameters.AddWithValue("@ID", id)
            cn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub AddMacchinaParams(cmd As SqlCommand, m As MC_Macchina)
        cmd.Parameters.AddWithValue("@Mat",  m.Matricola)
        cmd.Parameters.AddWithValue("@Nome", m.NomeMacchina)
        cmd.Parameters.AddWithValue("@Mod",  m.Modello)
        cmd.Parameters.AddWithValue("@Tipo", If(m.TipoMacchina, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@Cli",  If(m.ClienteFinale, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@Anno", If(m.AnnoCostruzione.HasValue, CObj(m.AnnoCostruzione.Value), DBNull.Value))
        cmd.Parameters.AddWithValue("@Lng",  m.LinguaCodice)
        cmd.Parameters.AddWithValue("@Att",  m.Attiva)
        cmd.Parameters.AddWithValue("@Note", If(m.Note, CObj(DBNull.Value)))
    End Sub

    Private Function ReadMacchina(rd As SqlDataReader) As MC_Macchina
        Return New MC_Macchina With {
            .ID              = rd.GetInt32(0),
            .Matricola       = rd.GetString(1),
            .NomeMacchina    = rd.GetString(2),
            .Modello         = rd.GetString(3),
            .TipoMacchina    = If(rd.IsDBNull(4), "", rd.GetString(4)),
            .ClienteFinale   = If(rd.IsDBNull(5), "", rd.GetString(5)),
            .AnnoCostruzione = If(rd.IsDBNull(6), Nothing, CType(rd.GetInt32(6), Integer?)),
            .LinguaCodice    = rd.GetString(7),
            .Attiva          = rd.GetBoolean(8),
            .Note            = If(rd.IsDBNull(9), "", rd.GetString(9)),
            .DataCreazione   = rd.GetDateTime(10),
            .DataModifica    = If(rd.IsDBNull(11), Nothing, CType(rd.GetDateTime(11), DateTime?))
        }
    End Function

    ' ──────────────────────────────────────────────
    ' FOTOCELLULE
    ' ──────────────────────────────────────────────

    Public Function GetFotocellule(macchinaID As Integer) As List(Of MC_Fotocellula)
        Dim lista As New List(Of MC_Fotocellula)
        Dim sql = "SELECT ID,MacchinaID,Codice,Marca,Modello,TipoRilevazione," &
                  "Posizione,TensioneLavoro,UscitaLogica,DistanzaRilev," &
                  "NoteInstallaz,DataCreazione " &
                  "FROM [Tirelli_40].dbo.Fotocellule WHERE MacchinaID=@MID ORDER BY Codice"
        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            cmd.Parameters.AddWithValue("@MID", macchinaID)
            cn.Open()
            Using rd = cmd.ExecuteReader()
                While rd.Read()
                    lista.Add(ReadFotocellula(rd))
                End While
            End Using
        End Using
        Return lista
    End Function

    Public Function SalvaFotocellula(f As MC_Fotocellula) As Integer
        If f.ID = 0 Then
            Dim sql = "INSERT INTO [Tirelli_40].dbo.Fotocellule " &
                      "(MacchinaID,Codice,Marca,Modello,TipoRilevazione," &
                      "Posizione,TensioneLavoro,UscitaLogica,DistanzaRilev,NoteInstallaz) " &
                      "VALUES (@MID,@Cod,@Mar,@Mod,@Tipo,@Pos,@Tens,@Usc,@Dist,@Note);" &
                      "SELECT SCOPE_IDENTITY();"
            Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
                AddFotocParam(cmd, f)
                cn.Open()
                Return CInt(cmd.ExecuteScalar())
            End Using
        Else
            Dim sql = "UPDATE [Tirelli_40].dbo.Fotocellule SET " &
                      "Codice=@Cod,Marca=@Mar,Modello=@Mod,TipoRilevazione=@Tipo," &
                      "Posizione=@Pos,TensioneLavoro=@Tens,UscitaLogica=@Usc," &
                      "DistanzaRilev=@Dist,NoteInstallaz=@Note " &
                      "WHERE ID=@ID"
            Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
                AddFotocParam(cmd, f)
                cmd.Parameters.AddWithValue("@ID", f.ID)
                cn.Open()
                cmd.ExecuteNonQuery()
                Return f.ID
            End Using
        End If
    End Function

    Public Sub EliminaFotocellula(id As Integer)
        Using cn = GetConnection(),
              cmd As New SqlCommand("DELETE FROM [Tirelli_40].dbo.Fotocellule WHERE ID=@ID", cn)
            cmd.Parameters.AddWithValue("@ID", id)
            cn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub AddFotocParam(cmd As SqlCommand, f As MC_Fotocellula)
        cmd.Parameters.AddWithValue("@MID",  f.MacchinaID)
        cmd.Parameters.AddWithValue("@Cod",  f.Codice)
        cmd.Parameters.AddWithValue("@Mar",  If(f.Marca, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@Mod",  If(f.Modello, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@Tipo", If(f.TipoRilevazione, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@Pos",  If(f.Posizione, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@Tens", If(f.TensioneLavoro, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@Usc",  If(f.UscitaLogica, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@Dist", If(f.DistanzaRilev, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@Note", If(f.NoteInstallaz, CObj(DBNull.Value)))
    End Sub

    Private Function ReadFotocellula(rd As SqlDataReader) As MC_Fotocellula
        Return New MC_Fotocellula With {
            .ID             = rd.GetInt32(0),
            .MacchinaID     = rd.GetInt32(1),
            .Codice         = rd.GetString(2),
            .Marca          = If(rd.IsDBNull(3), "", rd.GetString(3)),
            .Modello        = If(rd.IsDBNull(4), "", rd.GetString(4)),
            .TipoRilevazione= If(rd.IsDBNull(5), "", rd.GetString(5)),
            .Posizione      = If(rd.IsDBNull(6), "", rd.GetString(6)),
            .TensioneLavoro = If(rd.IsDBNull(7), "", rd.GetString(7)),
            .UscitaLogica   = If(rd.IsDBNull(8), "", rd.GetString(8)),
            .DistanzaRilev  = If(rd.IsDBNull(9), "", rd.GetString(9)),
            .NoteInstallaz  = If(rd.IsDBNull(10), "", rd.GetString(10)),
            .DataCreazione  = rd.GetDateTime(11)
        }
    End Function

    ' ──────────────────────────────────────────────
    ' CODICI ERRORE
    ' ──────────────────────────────────────────────

    Public Function GetCodiciErrore(macchinaID As Integer) As List(Of MC_CodiceErrore)
        Dim lista As New List(Of MC_CodiceErrore)
        Dim sql = "SELECT ID,MacchinaID,CodiceErrore,Titolo,Descrizione," &
                  "Causa,Rimedio,Gravita,NomeScreenshot,PathScreenshot,DataCreazione " &
                  "FROM [Tirelli_40].dbo.CodiciErrore WHERE MacchinaID=@MID ORDER BY CodiceErrore"
        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            cmd.Parameters.AddWithValue("@MID", macchinaID)
            cn.Open()
            Using rd = cmd.ExecuteReader()
                While rd.Read()
                    lista.Add(ReadCodiceErrore(rd))
                End While
            End Using
        End Using
        Return lista
    End Function

    Public Function SalvaCodiceErrore(e As MC_CodiceErrore) As Integer
        If e.ID = 0 Then
            Dim sql = "INSERT INTO [Tirelli_40].dbo.CodiciErrore " &
                      "(MacchinaID,CodiceErrore,Titolo,Descrizione,Causa,Rimedio," &
                      "Gravita,NomeScreenshot,PathScreenshot) " &
                      "VALUES (@MID,@Cod,@Tit,@Desc,@Cau,@Rim,@Grav,@NomeS,@PathS);" &
                      "SELECT SCOPE_IDENTITY();"
            Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
                AddErroreParams(cmd, e)
                cn.Open()
                Return CInt(cmd.ExecuteScalar())
            End Using
        Else
            Dim sql = "UPDATE [Tirelli_40].dbo.CodiciErrore SET " &
                      "CodiceErrore=@Cod,Titolo=@Tit,Descrizione=@Desc," &
                      "Causa=@Cau,Rimedio=@Rim,Gravita=@Grav," &
                      "NomeScreenshot=@NomeS,PathScreenshot=@PathS " &
                      "WHERE ID=@ID"
            Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
                AddErroreParams(cmd, e)
                cmd.Parameters.AddWithValue("@ID", e.ID)
                cn.Open()
                cmd.ExecuteNonQuery()
                Return e.ID
            End Using
        End If
    End Function

    Public Sub EliminaCodiceErrore(id As Integer)
        Using cn = GetConnection(),
              cmd As New SqlCommand("DELETE FROM [Tirelli_40].dbo.CodiciErrore WHERE ID=@ID", cn)
            cmd.Parameters.AddWithValue("@ID", id)
            cn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Sub AddErroreParams(cmd As SqlCommand, e As MC_CodiceErrore)
        cmd.Parameters.AddWithValue("@MID",  e.MacchinaID)
        cmd.Parameters.AddWithValue("@Cod",  e.Codice)
        cmd.Parameters.AddWithValue("@Tit",  e.Titolo)
        cmd.Parameters.AddWithValue("@Desc", If(e.Descrizione, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@Cau",  If(e.Causa, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@Rim",  If(e.Rimedio, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@Grav", If(e.Gravita, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@NomeS",If(e.NomeScreenshot, CObj(DBNull.Value)))
        cmd.Parameters.AddWithValue("@PathS",If(e.PathScreenshot, CObj(DBNull.Value)))
    End Sub

    Private Function ReadCodiceErrore(rd As SqlDataReader) As MC_CodiceErrore
        Return New MC_CodiceErrore With {
            .ID             = rd.GetInt32(0),
            .MacchinaID     = rd.GetInt32(1),
            .Codice         = rd.GetString(2),
            .Titolo         = rd.GetString(3),
            .Descrizione    = If(rd.IsDBNull(4), "", rd.GetString(4)),
            .Causa          = If(rd.IsDBNull(5), "", rd.GetString(5)),
            .Rimedio        = If(rd.IsDBNull(6), "", rd.GetString(6)),
            .Gravita        = If(rd.IsDBNull(7), "Avviso", rd.GetString(7)),
            .NomeScreenshot = If(rd.IsDBNull(8), "", rd.GetString(8)),
            .PathScreenshot = If(rd.IsDBNull(9), "", rd.GetString(9)),
            .DataCreazione  = rd.GetDateTime(10)
        }
    End Function

    ' ──────────────────────────────────────────────
    ' LINGUE
    ' ──────────────────────────────────────────────

    Public Function GetLingue() As List(Of MC_Lingua)
        Dim lista As New List(Of MC_Lingua)
        Using cn = GetConnection(),
              cmd As New SqlCommand("SELECT Codice,Nome,Attiva FROM [Tirelli_40].dbo.Lingue WHERE Codice IN ('IT','EN','FR') ORDER BY CASE Codice WHEN 'IT' THEN 0 WHEN 'EN' THEN 1 WHEN 'FR' THEN 2 END", cn)
            cn.Open()
            Using rd = cmd.ExecuteReader()
                While rd.Read()
                    lista.Add(New MC_Lingua With {
                        .Codice = rd.GetString(0),
                        .Nome = rd.GetString(1),
                        .Attiva = rd.GetBoolean(2)
                    })
                End While
            End Using
        End Using
        Return lista
    End Function

    Public Function TestConnessione() As Boolean
        Try
            Using cn = GetConnection()
                cn.Open()
                Return True
            End Using
        Catch
            Return False
        End Try
    End Function

End Class
