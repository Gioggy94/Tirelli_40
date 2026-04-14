Imports System.Data.SqlClient

Public Class MC_DatabaseService

    Public Sub New()
        Try
            CreaTabelleSeNonEsistono()
        Catch
        End Try
    End Sub

    Private Function GetConnection() As SqlConnection
        Return New SqlConnection(Homepage.sap_tirelli)
    End Function

    ' ──────────────────────────────────────────────
    ' SETUP TABELLE LOOKUP
    ' ──────────────────────────────────────────────

    Private Sub CreaTabelleSeNonEsistono()
        Dim sql = "
IF NOT EXISTS (SELECT * FROM [Tirelli_40].sys.tables WHERE name='MC_Modelli')
CREATE TABLE [Tirelli_40].dbo.MC_Modelli (
    ID   int IDENTITY(1,1) PRIMARY KEY,
    Nome nvarchar(100) NOT NULL,
    DataCreazione datetime DEFAULT GETDATE()
);
IF NOT EXISTS (SELECT * FROM [Tirelli_40].sys.tables WHERE name='MC_TipiMacchina')
CREATE TABLE [Tirelli_40].dbo.MC_TipiMacchina (
    ID   int IDENTITY(1,1) PRIMARY KEY,
    Nome nvarchar(100) NOT NULL,
    DataCreazione datetime DEFAULT GETDATE()
);
IF NOT EXISTS (SELECT * FROM [Tirelli_40].sys.columns WHERE object_id=OBJECT_ID('[Tirelli_40].dbo.Macchine') AND name='PesoKg')
    ALTER TABLE [Tirelli_40].dbo.Macchine ADD PesoKg float NULL;
IF NOT EXISTS (SELECT * FROM [Tirelli_40].sys.columns WHERE object_id=OBJECT_ID('[Tirelli_40].dbo.Macchine') AND name='ConsumoAria')
    ALTER TABLE [Tirelli_40].dbo.Macchine ADD ConsumoAria float NULL;
IF NOT EXISTS (SELECT * FROM [Tirelli_40].sys.columns WHERE object_id=OBJECT_ID('[Tirelli_40].dbo.Macchine') AND name='Corrente')
    ALTER TABLE [Tirelli_40].dbo.Macchine ADD Corrente nvarchar(100) NULL;
IF NOT EXISTS (SELECT * FROM [Tirelli_40].sys.columns WHERE object_id=OBJECT_ID('[Tirelli_40].dbo.Macchine') AND name='Tensione')
    ALTER TABLE [Tirelli_40].dbo.Macchine ADD Tensione nvarchar(100) NULL;
IF NOT EXISTS (SELECT * FROM [Tirelli_40].sys.tables WHERE name='MC_TipiFotocellule')
CREATE TABLE [Tirelli_40].dbo.MC_TipiFotocellule (
    ID   int IDENTITY(1,1) PRIMARY KEY,
    Nome nvarchar(100) NOT NULL,
    DataCreazione datetime DEFAULT GETDATE()
);
IF NOT EXISTS (SELECT * FROM [Tirelli_40].sys.tables WHERE name='MC_CatalogoFotocellule')
CREATE TABLE [Tirelli_40].dbo.MC_CatalogoFotocellule (
    ID            int IDENTITY(1,1) PRIMARY KEY,
    Codice        nvarchar(100) NOT NULL,
    TipoID        int NOT NULL,
    PathImmagine  nvarchar(500) NULL,
    DataCreazione datetime DEFAULT GETDATE()
);
IF NOT EXISTS (SELECT * FROM [Tirelli_40].sys.columns WHERE object_id=OBJECT_ID('[Tirelli_40].dbo.Fotocellule') AND name='CatalogoID')
    ALTER TABLE [Tirelli_40].dbo.Fotocellule ADD CatalogoID int NULL;"
        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            cn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ' ──────────────────────────────────────────────
    ' MACCHINE
    ' ──────────────────────────────────────────────

    Public Function GetMacchine(soloAttive As Boolean) As List(Of MC_Macchina)
        Dim lista As New List(Of MC_Macchina)
        Dim sql = "SELECT ID,Matricola,NomeMacchina,Modello,TipoMacchina," &
                  "ClienteFinale,AnnoCostruzione,LinguaCodice,Attiva,Note," &
                  "DataCreazione,DataModifica,PesoKg,ConsumoAria,Corrente,Tensione " &
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
                  "DataCreazione,DataModifica,PesoKg,ConsumoAria,Corrente,Tensione " &
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

    ' Ricerca macchine su AS400 + arricchimento locale
    Public Function GetMacchineAS400(filtroMatricola As String, filtroCliente As String) As List(Of MC_Macchina)
        Dim lista As New List(Of MC_Macchina)
        Dim fm = If(filtroMatricola?.Trim(), "")
        Dim fc = If(filtroCliente?.Trim(), "")

        Dim inner = "SELECT trim(t0.matricola) as matricola, " &
                    "trim(t0.itemname) as itemname, " &
                    "trim(t0.dscli_fatt) as dscli_fatt " &
                    "FROM TIR90VIS.JGALCOM t0 " &
                    "WHERE t0.matricola <> '''' " &
                    "AND substring(t0.matricola,1,1) = ''M'' "
        If Not String.IsNullOrEmpty(fm) Then
            inner &= $"AND UPPER(t0.matricola) LIKE ''%{EscAs400(fm)}%'' "
        End If
        If Not String.IsNullOrEmpty(fc) Then
            Dim fcE = EscAs400(fc)
            inner &= $"AND (UPPER(t0.dscli_fatt) LIKE ''%{fcE}%'' " &
                     $"OR UPPER(t0.codice_finale) LIKE ''%{fcE}%'') "
        End If
        inner &= "ORDER BY t0.matricola DESC FETCH FIRST 200 ROWS ONLY"

        Dim sql = "SELECT oq.matricola, oq.itemname, oq.dscli_fatt, " &
                  "ISNULL(m.ID, 0) as ID, " &
                  "ISNULL(m.Modello, '') as Modello, " &
                  "ISNULL(m.TipoMacchina, '') as TipoMacchina, " &
                  "ISNULL(m.LinguaCodice, 'IT') as LinguaCodice, " &
                  "ISNULL(m.Note, '') as Note, " &
                  "m.PesoKg, m.ConsumoAria, " &
                  "ISNULL(m.Corrente, '') as Corrente, " &
                  "ISNULL(m.Tensione, '') as Tensione " &
                  $"FROM OPENQUERY(AS400, '{inner}') oq " &
                  "LEFT JOIN [Tirelli_40].dbo.Macchine m ON m.Matricola = oq.matricola COLLATE SQL_Latin1_General_CP1_CI_AS"

        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            cn.Open()
            Using rd = cmd.ExecuteReader()
                While rd.Read()
                    lista.Add(New MC_Macchina With {
                        .Matricola     = rd.GetString(0).Trim(),
                        .NomeMacchina  = If(rd.IsDBNull(1), "", rd.GetString(1).Trim()),
                        .ClienteFinale = If(rd.IsDBNull(2), "", rd.GetString(2).Trim()),
                        .ID            = rd.GetInt32(3),
                        .Modello       = rd.GetString(4),
                        .TipoMacchina  = rd.GetString(5),
                        .LinguaCodice  = rd.GetString(6),
                        .Note          = rd.GetString(7),
                        .PesoKg        = If(rd.IsDBNull(8), Nothing, CType(rd.GetDouble(8), Double?)),
                        .ConsumoAria   = If(rd.IsDBNull(9), Nothing, CType(rd.GetDouble(9), Double?)),
                        .Corrente      = rd.GetString(10),
                        .Tensione      = rd.GetString(11)
                    })
                End While
            End Using
        End Using
        Return lista
    End Function

    Private Shared Function EscAs400(s As String) As String
        Return s.ToUpper().Replace("'", "''''")
    End Function

    ' UPSERT dati locali per una macchina (crea record in Macchine se non esiste)
    Public Function SalvaExtraMacchina(m As MC_Macchina) As Integer
        Dim sql = "IF EXISTS (SELECT 1 FROM [Tirelli_40].dbo.Macchine WHERE Matricola=@Mat) " &
                  "  UPDATE [Tirelli_40].dbo.Macchine SET " &
                  "    NomeMacchina=@Nome, Modello=@Mod, TipoMacchina=@Tipo, " &
                  "    ClienteFinale=@Cli, LinguaCodice=@Lng, Note=@Note, " &
                  "    PesoKg=@Peso, ConsumoAria=@Aria, Corrente=@Cor, Tensione=@Ten, " &
                  "    DataModifica=GETDATE() " &
                  "  WHERE Matricola=@Mat " &
                  "ELSE " &
                  "  INSERT INTO [Tirelli_40].dbo.Macchine " &
                  "    (Matricola,NomeMacchina,Modello,TipoMacchina,ClienteFinale,LinguaCodice,Attiva,Note,PesoKg,ConsumoAria,Corrente,Tensione) " &
                  "  VALUES (@Mat,@Nome,@Mod,@Tipo,@Cli,@Lng,1,@Note,@Peso,@Aria,@Cor,@Ten); " &
                  "SELECT ID FROM [Tirelli_40].dbo.Macchine WHERE Matricola=@Mat"
        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            cmd.Parameters.AddWithValue("@Mat",  m.Matricola)
            cmd.Parameters.AddWithValue("@Nome", If(m.NomeMacchina, ""))
            cmd.Parameters.AddWithValue("@Mod",  If(m.Modello, ""))
            cmd.Parameters.AddWithValue("@Tipo", If(m.TipoMacchina, ""))
            cmd.Parameters.AddWithValue("@Cli",  If(m.ClienteFinale, ""))
            cmd.Parameters.AddWithValue("@Lng",  If(String.IsNullOrEmpty(m.LinguaCodice), "IT", m.LinguaCodice))
            cmd.Parameters.AddWithValue("@Note", If(m.Note, ""))
            cmd.Parameters.AddWithValue("@Peso", If(m.PesoKg.HasValue, CObj(m.PesoKg.Value), DBNull.Value))
            cmd.Parameters.AddWithValue("@Aria", If(m.ConsumoAria.HasValue, CObj(m.ConsumoAria.Value), DBNull.Value))
            cmd.Parameters.AddWithValue("@Cor",  If(m.Corrente, ""))
            cmd.Parameters.AddWithValue("@Ten",  If(m.Tensione, ""))
            cn.Open()
            Return CInt(cmd.ExecuteScalar())
        End Using
    End Function

    ' ──────────────────────────────────────────────
    ' MODELLI
    ' ──────────────────────────────────────────────

    Public Function GetModelli() As List(Of MC_Modello)
        Dim lista As New List(Of MC_Modello)
        Using cn = GetConnection(),
              cmd As New SqlCommand("SELECT ID, Nome FROM [Tirelli_40].dbo.MC_Modelli ORDER BY Nome", cn)
            cn.Open()
            Using rd = cmd.ExecuteReader()
                While rd.Read()
                    lista.Add(New MC_Modello With {.ID = rd.GetInt32(0), .Nome = rd.GetString(1)})
                End While
            End Using
        End Using
        Return lista
    End Function

    Public Sub SalvaModello(nome As String)
        Using cn = GetConnection(),
              cmd As New SqlCommand("INSERT INTO [Tirelli_40].dbo.MC_Modelli (Nome) VALUES (@N)", cn)
            cmd.Parameters.AddWithValue("@N", nome)
            cn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub EliminaModello(id As Integer)
        Using cn = GetConnection(),
              cmd As New SqlCommand("DELETE FROM [Tirelli_40].dbo.MC_Modelli WHERE ID=@ID", cn)
            cmd.Parameters.AddWithValue("@ID", id)
            cn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ' ──────────────────────────────────────────────
    ' TIPI MACCHINA
    ' ──────────────────────────────────────────────

    Public Function GetTipiMacchina() As List(Of MC_TipoMacchina)
        Dim lista As New List(Of MC_TipoMacchina)
        Using cn = GetConnection(),
              cmd As New SqlCommand("SELECT ID, Nome FROM [Tirelli_40].dbo.MC_TipiMacchina ORDER BY Nome", cn)
            cn.Open()
            Using rd = cmd.ExecuteReader()
                While rd.Read()
                    lista.Add(New MC_TipoMacchina With {.ID = rd.GetInt32(0), .Nome = rd.GetString(1)})
                End While
            End Using
        End Using
        Return lista
    End Function

    Public Sub SalvaTipoMacchina(nome As String)
        Using cn = GetConnection(),
              cmd As New SqlCommand("INSERT INTO [Tirelli_40].dbo.MC_TipiMacchina (Nome) VALUES (@N)", cn)
            cmd.Parameters.AddWithValue("@N", nome)
            cn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub EliminaTipoMacchina(id As Integer)
        Using cn = GetConnection(),
              cmd As New SqlCommand("DELETE FROM [Tirelli_40].dbo.MC_TipiMacchina WHERE ID=@ID", cn)
            cmd.Parameters.AddWithValue("@ID", id)
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
            .DataModifica    = If(rd.IsDBNull(11), Nothing, CType(rd.GetDateTime(11), DateTime?)),
            .PesoKg          = If(rd.FieldCount > 12 AndAlso Not rd.IsDBNull(12), CType(rd.GetDouble(12), Double?), Nothing),
            .ConsumoAria     = If(rd.FieldCount > 13 AndAlso Not rd.IsDBNull(13), CType(rd.GetDouble(13), Double?), Nothing),
            .Corrente        = If(rd.FieldCount > 14 AndAlso Not rd.IsDBNull(14), rd.GetString(14), ""),
            .Tensione        = If(rd.FieldCount > 15 AndAlso Not rd.IsDBNull(15), rd.GetString(15), "")
        }
    End Function

    ' ──────────────────────────────────────────────
    ' FOTOCELLULE
    ' ──────────────────────────────────────────────

    Public Function GetFotocellule(macchinaID As Integer) As List(Of MC_Fotocellula)
        Dim lista As New List(Of MC_Fotocellula)
        Dim sql = "SELECT f.ID, f.MacchinaID, ISNULL(f.CatalogoID,0), " &
                  "ISNULL(c.Codice,'') AS Codice, " &
                  "ISNULL(t.Nome,'') AS TipoNome, " &
                  "ISNULL(c.PathImmagine,'') AS PathImmagine, " &
                  "f.DataCreazione " &
                  "FROM [Tirelli_40].dbo.Fotocellule f " &
                  "LEFT JOIN [Tirelli_40].dbo.MC_CatalogoFotocellule c ON c.ID = f.CatalogoID " &
                  "LEFT JOIN [Tirelli_40].dbo.MC_TipiFotocellule t ON t.ID = c.TipoID " &
                  "WHERE f.MacchinaID=@MID ORDER BY c.Codice"
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
        Dim sql = "INSERT INTO [Tirelli_40].dbo.Fotocellule (MacchinaID,CatalogoID) " &
                  "VALUES (@MID,@CatID); SELECT SCOPE_IDENTITY();"
        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            cmd.Parameters.AddWithValue("@MID",   f.MacchinaID)
            cmd.Parameters.AddWithValue("@CatID", f.CatalogoID)
            cn.Open()
            Return CInt(cmd.ExecuteScalar())
        End Using
    End Function

    Public Sub EliminaFotocellula(id As Integer)
        Using cn = GetConnection(),
              cmd As New SqlCommand("DELETE FROM [Tirelli_40].dbo.Fotocellule WHERE ID=@ID", cn)
            cmd.Parameters.AddWithValue("@ID", id)
            cn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Function ReadFotocellula(rd As SqlDataReader) As MC_Fotocellula
        Return New MC_Fotocellula With {
            .ID            = rd.GetInt32(0),
            .MacchinaID    = rd.GetInt32(1),
            .CatalogoID    = rd.GetInt32(2),
            .Codice        = rd.GetString(3),
            .TipoNome      = rd.GetString(4),
            .PathImmagine  = rd.GetString(5),
            .DataCreazione = rd.GetDateTime(6)
        }
    End Function

    ' ──────────────────────────────────────────────
    ' CATALOGO FOTOCELLULE
    ' ──────────────────────────────────────────────

    Public Function GetTipiFotocellula() As List(Of MC_TipoFotocellula)
        Dim lista As New List(Of MC_TipoFotocellula)
        Using cn = GetConnection(),
              cmd As New SqlCommand("SELECT ID,Nome FROM [Tirelli_40].dbo.MC_TipiFotocellule ORDER BY Nome", cn)
            cn.Open()
            Using rd = cmd.ExecuteReader()
                While rd.Read()
                    lista.Add(New MC_TipoFotocellula With {.ID = rd.GetInt32(0), .Nome = rd.GetString(1)})
                End While
            End Using
        End Using
        Return lista
    End Function

    Public Sub SalvaTipoFotocellula(nome As String)
        Using cn = GetConnection(),
              cmd As New SqlCommand("INSERT INTO [Tirelli_40].dbo.MC_TipiFotocellule (Nome) VALUES (@N)", cn)
            cmd.Parameters.AddWithValue("@N", nome)
            cn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub EliminaTipoFotocellula(id As Integer)
        Using cn = GetConnection(),
              cmd As New SqlCommand("DELETE FROM [Tirelli_40].dbo.MC_TipiFotocellule WHERE ID=@ID", cn)
            cmd.Parameters.AddWithValue("@ID", id)
            cn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Function GetCatalogoFotocellule(Optional filtro As String = "") As List(Of MC_CatalogoFotocellula)
        Dim lista As New List(Of MC_CatalogoFotocellula)
        Dim where = If(String.IsNullOrWhiteSpace(filtro), "",
                       $" AND (UPPER(c.Codice) LIKE '%{filtro.ToUpper().Replace("'", "''")}%' " &
                       $"OR UPPER(t.Nome) LIKE '%{filtro.ToUpper().Replace("'", "''")}%')")
        Dim sql = "SELECT c.ID, c.Codice, c.TipoID, ISNULL(t.Nome,'') as TipoNome, " &
                  "ISNULL(c.PathImmagine,'') as PathImmagine, c.DataCreazione " &
                  "FROM [Tirelli_40].dbo.MC_CatalogoFotocellule c " &
                  "LEFT JOIN [Tirelli_40].dbo.MC_TipiFotocellule t ON t.ID = c.TipoID " &
                  $"WHERE 1=1{where} ORDER BY c.Codice"
        Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
            cn.Open()
            Using rd = cmd.ExecuteReader()
                While rd.Read()
                    lista.Add(New MC_CatalogoFotocellula With {
                        .ID           = rd.GetInt32(0),
                        .Codice       = rd.GetString(1),
                        .TipoID       = rd.GetInt32(2),
                        .TipoNome     = rd.GetString(3),
                        .PathImmagine = rd.GetString(4),
                        .DataCreazione= rd.GetDateTime(5)
                    })
                End While
            End Using
        End Using
        Return lista
    End Function

    Public Function GetCatalogoFotocellula(id As Integer) As MC_CatalogoFotocellula
        Return GetCatalogoFotocellule().FirstOrDefault(Function(x) x.ID = id)
    End Function

    Public Function SalvaCatalogoFotocellula(c As MC_CatalogoFotocellula) As Integer
        If c.ID = 0 Then
            Dim sql = "INSERT INTO [Tirelli_40].dbo.MC_CatalogoFotocellule (Codice,TipoID,PathImmagine) " &
                      "VALUES (@Cod,@TipoID,@Path); SELECT SCOPE_IDENTITY();"
            Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
                cmd.Parameters.AddWithValue("@Cod",   c.Codice)
                cmd.Parameters.AddWithValue("@TipoID",c.TipoID)
                cmd.Parameters.AddWithValue("@Path",  If(String.IsNullOrEmpty(c.PathImmagine), CObj(DBNull.Value), c.PathImmagine))
                cn.Open()
                Return CInt(cmd.ExecuteScalar())
            End Using
        Else
            Dim sql = "UPDATE [Tirelli_40].dbo.MC_CatalogoFotocellule SET Codice=@Cod,TipoID=@TipoID,PathImmagine=@Path WHERE ID=@ID"
            Using cn = GetConnection(), cmd As New SqlCommand(sql, cn)
                cmd.Parameters.AddWithValue("@Cod",   c.Codice)
                cmd.Parameters.AddWithValue("@TipoID",c.TipoID)
                cmd.Parameters.AddWithValue("@Path",  If(String.IsNullOrEmpty(c.PathImmagine), CObj(DBNull.Value), c.PathImmagine))
                cmd.Parameters.AddWithValue("@ID",    c.ID)
                cn.Open()
                cmd.ExecuteNonQuery()
                Return c.ID
            End Using
        End If
    End Function

    Public Sub EliminaCatalogoFotocellula(id As Integer)
        Using cn = GetConnection(),
              cmd As New SqlCommand("DELETE FROM [Tirelli_40].dbo.MC_CatalogoFotocellule WHERE ID=@ID", cn)
            cmd.Parameters.AddWithValue("@ID", id)
            cn.Open()
            cmd.ExecuteNonQuery()
        End Using
    End Sub

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
