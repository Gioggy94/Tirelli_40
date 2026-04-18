Imports System.Data.SqlClient
Imports System.Drawing.Printing

Public Class Form_Ferretto
    Public id_da_stampare As Integer = 0
    Private id_ultimo_timer As Integer = 0  ' usato solo dal timer, indipendente dal click manuale
    Public Stampante_Selezionata As Boolean
    Public stampa_descrizione As String
    Public stampa_progetto As String
    Public stampa_codice_articolo_padre As String

    ' Helper: converte DBNull/Nothing in stringa vuota e fa Trim
    Private Shared Function S(val As Object) As String
        If IsDBNull(val) OrElse val Is Nothing Then Return ""
        Return val.ToString().Trim()
    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        movimenti(DataGridView1)
    End Sub

    Sub movimenti(par_datagridview As DataGridView)
        par_datagridview.Rows.Clear()

        Dim filterId As String = TextBox4.Text.Trim()
        Dim filtroItem As String = TextBox2.Text.Trim()          ' Codice SAP → item
        Dim filtroMag As String = TextBox3.Text.Trim()           ' Mag → Magazzino
        Dim filtroCommessa As String = TextBox5.Text.Trim()      ' Commessa → COM
        Dim filtroSotto As String = TextBox6.Text.Trim()         ' Sottocommessa → ID_Value

        ' CTE necessaria: COM/Magazzino/ID_Value sono alias CASE e non filtrabili
        ' direttamente nel WHERE della query originale.
        Dim sql As String = "
WITH base AS (
    SELECT TOP (50)
        [id], [recordStatus], [recordWritingDate], [recordImportationDate],
        [plantId], [response], [listType], [listNumber],
        CASE
            WHEN CHARINDEX('OD:', listNumber) > 0 AND CHARINDEX('-COM:', listNumber) > 0
                 AND CHARINDEX('-COM:', listNumber) > CHARINDEX('OD:', listNumber)
            THEN SUBSTRING(listNumber, CHARINDEX('OD:', listNumber) + 3,
                 CHARINDEX('-COM:', listNumber) - CHARINDEX('OD:', listNumber) - 3)
            ELSE NULL
        END AS OD_Number,
        CASE
            WHEN listNumber LIKE 'OC-%'
            THEN SUBSTRING(listNumber, 4, LEN(listNumber))
            ELSE NULL
        END AS OC_Number,
        CASE
            WHEN CHARINDEX('COM:', listNumber) > 0 AND CHARINDEX('-ART:', listNumber) > 0
                 AND CHARINDEX('-ART:', listNumber) > CHARINDEX('COM:', listNumber)
            THEN SUBSTRING(listNumber, CHARINDEX('COM:', listNumber) + 4,
                 CHARINDEX('-ART:', listNumber) - CHARINDEX('COM:', listNumber) - 4)
            WHEN CHARINDEX('COM:', listNumber) > 0 AND CHARINDEX('-ID:', listNumber) > 0
                 AND CHARINDEX('-ID:', listNumber) > CHARINDEX('COM:', listNumber)
            THEN SUBSTRING(listNumber, CHARINDEX('COM:', listNumber) + 4,
                 CHARINDEX('-ID:', listNumber) - CHARINDEX('COM:', listNumber) - 4)
            WHEN CHARINDEX('COM:', listNumber) > 0 AND CHARINDEX('-MG:', listNumber) > 0
                 AND CHARINDEX('-MG:', listNumber) > CHARINDEX('COM:', listNumber)
            THEN SUBSTRING(listNumber, CHARINDEX('COM:', listNumber) + 4,
                 CHARINDEX('-MG:', listNumber) - CHARINDEX('COM:', listNumber) - 4)
            ELSE NULL
        END AS COM,
        CASE
            WHEN CHARINDEX('-ID:', listNumber) > 0 AND CHARINDEX('-MG:', listNumber) > 0
                 AND CHARINDEX('-MG:', listNumber) > CHARINDEX('-ID:', listNumber)
            THEN REPLACE(SUBSTRING(listNumber, CHARINDEX('-ID:', listNumber) + 3,
                 CHARINDEX('-MG:', listNumber) - CHARINDEX('-ID:', listNumber) - 3), ':', '')
            ELSE NULL
        END AS ID_Value,
        CASE
            WHEN CHARINDEX('-MG:', listNumber) > 0
            THEN SUBSTRING(listNumber, CHARINDEX('-MG:', listNumber) + 4, LEN(listNumber))
            ELSE NULL
        END AS Magazzino,
        [lineNumber], [item], [batch], [serialNumber],
        [requestedQty], [processedQty], [errorCause], [wmsGenerated]
    FROM [FGWmsErp].[dbo].[LISTS_RESULT]
    WHERE response <> 11 AND listtype = 0 AND [processedQty] > 0
    ORDER BY id DESC
)
SELECT * FROM base
WHERE 1=1
  AND (@filterId         = '' OR CAST(id AS VARCHAR) LIKE '%' + @filterId + '%')
  AND (@filtroItem       = '' OR LTRIM(RTRIM(COALESCE(item,'')))      LIKE '%' + @filtroItem + '%')
  AND (@filtroMag        = '' OR LTRIM(RTRIM(COALESCE(Magazzino,''))) LIKE '%' + @filtroMag + '%')
  AND (@filtroCommessa   = '' OR LTRIM(RTRIM(COALESCE(COM,'')))       LIKE '%' + @filtroCommessa + '%')
  AND (@filtroSotto      = '' OR LTRIM(RTRIM(COALESCE(ID_Value,'')))  LIKE '%' + @filtroSotto + '%')"

        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Cnn.Open()
            Using CMD As New SqlCommand(sql, Cnn)
                CMD.Parameters.AddWithValue("@filterId", filterId)
                CMD.Parameters.AddWithValue("@filtroItem", filtroItem)
                CMD.Parameters.AddWithValue("@filtroMag", filtroMag)
                CMD.Parameters.AddWithValue("@filtroCommessa", filtroCommessa)
                CMD.Parameters.AddWithValue("@filtroSotto", filtroSotto)

                Using reader As SqlDataReader = CMD.ExecuteReader()
                    Do While reader.Read()
                        par_datagridview.Rows.Add(
                            reader("id"),
                            reader("recordStatus"),
                            reader("recordWritingDate"),
                            reader("recordImportationDate"),
                            reader("plantId"),
                            reader("response"),
                            reader("listType"),
                            S(reader("listNumber")),
                            S(reader("OD_Number")),
                            S(reader("OC_Number")),
                            S(reader("COM")),
                            S(reader("ID_Value")),
                            S(reader("Magazzino")),
                            reader("lineNumber"),
                            S(reader("item")),
                            "",
                            S(reader("batch")),
                            S(reader("serialNumber")),
                            reader("requestedQty"),
                            reader("processedQty"),
                            S(reader("errorCause")),
                            reader("wmsGenerated")
                        )
                    Loop
                End Using
            End Using
        End Using

        ' --- Lookup descrizioni AS400: una sola query con IN sui codici distinti ---
        Dim codici As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        For Each row As DataGridViewRow In par_datagridview.Rows
            Dim cod = S(row.Cells("item").Value)
            If cod <> "" Then codici.Add(cod)
        Next

        If codici.Count > 0 Then
            ' Ogni codice viene racchiuso in ''...'' (doppi apici per SQL Server → singoli per AS400)
            Dim inList = String.Join(",", codici.Select(Function(x) "''" & x.Replace("'", "''''") & "''"))
            Dim sqlDesc = "SELECT trim(CODE) AS code, trim(DES_CODE) AS des_code " &
                          "FROM OPENQUERY(AS400, 'SELECT trim(CODE) AS CODE, trim(DES_CODE) AS DES_CODE " &
                          "FROM S786FAD1.TIR90VIS.JGALART WHERE trim(CODE) IN (" & inList & ")')"

            Dim descMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Using Cnn2 As New SqlConnection(Homepage.sap_tirelli)
                Cnn2.Open()
                Using CMD2 As New SqlCommand(sqlDesc, Cnn2)
                    Using reader2 As SqlDataReader = CMD2.ExecuteReader()
                        Do While reader2.Read()
                            descMap(S(reader2("code"))) = S(reader2("des_code"))
                        Loop
                    End Using
                End Using
            End Using

            For Each row As DataGridViewRow In par_datagridview.Rows
                Dim cod = S(row.Cells("item").Value)
                If cod <> "" Then
                    Dim desc As String = ""
                    descMap.TryGetValue(cod, desc)
                    row.Cells("descrizione_articolo").Value = desc
                End If
            Next
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.AutoGenerateColumns = False
        genera_colonne(DataGridView1)
        movimenti(DataGridView1)
    End Sub

    Sub genera_colonne(par_datagridview As DataGridView)
        par_datagridview.Columns.Clear()
        par_datagridview.AutoGenerateColumns = False
        par_datagridview.Columns.Add("id", "ID")
        par_datagridview.Columns.Add("recordStatus", "Stato Record")
        par_datagridview.Columns.Add("recordWritingDate", "Data Scrittura")
        par_datagridview.Columns.Add("recordImportationDate", "Data Importazione")
        par_datagridview.Columns.Add("plantId", "Plant")
        par_datagridview.Columns.Add("response", "Response")
        par_datagridview.Columns.Add("listType", "Tipo Lista")
        par_datagridview.Columns.Add("listNumber", "Numero Lista")
        par_datagridview.Columns.Add("OD_Number", "N° OD")
        par_datagridview.Columns.Add("OC_Number", "N° OC")
        par_datagridview.Columns.Add("COM", "COM")
        par_datagridview.Columns.Add("ID_Value", "ID Valore")
        par_datagridview.Columns.Add("Magazzino", "Magazzino")
        par_datagridview.Columns.Add("lineNumber", "N° Riga")
        par_datagridview.Columns.Add("item", "Articolo")
        par_datagridview.Columns.Add("descrizione_articolo", "Descrizione")
        par_datagridview.Columns.Add("batch", "Batch")
        par_datagridview.Columns.Add("serialNumber", "Matricola")
        par_datagridview.Columns.Add("requestedQty", "Qtà Richiesta")
        par_datagridview.Columns.Add("processedQty", "Qtà Processata")
        par_datagridview.Columns.Add("errorCause", "Causa Errore")
        par_datagridview.Columns.Add("wmsGenerated", "WMS Generato")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellBorderStyleChanged(sender As Object, e As EventArgs) Handles DataGridView1.CellBorderStyleChanged

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            id_da_stampare = DataGridView1.Rows(e.RowIndex).Cells(columnName:="id").Value
            RichTextBox1.Text = id_da_stampare
        End If
    End Sub

    ' =============================================
    ' VARIABILI GLOBALI DI STAMPA
    ' =============================================
    Dim stampa_OD As String = ""
    Dim stampa_OC As String = ""
    Dim stampa_COM As String = ""
    Dim stampa_ID As String = ""
    Dim stampa_articolo As String = ""
    Dim stampa_qta_richiesta As String = ""
    Dim stampa_qta_processata As String = ""
    Dim stampa_numero_scontrino As String = ""
    Dim stampa_baia As String = ""
    Dim stampa_stato_odp As String = ""
    Dim stampa_magazzino_destinazione As String = ""
    Dim codice_gruppo_art_padre As String = ""
    Dim stampa_tipo_OC As String = ""
    Dim stampa_numero_OC As String = ""
    Dim Stampa_Ubicazione_Macchina As String = ""

    ' =============================================
    ' CLICK PULSANTE STAMPA
    ' =============================================
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If id_da_stampare = 0 Then
            MessageBox.Show("Seleziona una riga da stampare.")
            Exit Sub
        End If
        Stampa_stampa(id_da_stampare)
    End Sub

    ' =============================================
    ' CARICA DATI E AVVIA STAMPA
    ' =============================================
    Sub Stampa_stampa(par_id_stampa As Integer)

        Dim sql As String = "
SELECT [id], [listNumber], [item], [requestedQty], [processedQty],
    CASE
        WHEN CHARINDEX('OD:', listNumber) > 0 AND CHARINDEX('-COM:', listNumber) > 0
             AND CHARINDEX('-COM:', listNumber) > CHARINDEX('OD:', listNumber)
        THEN SUBSTRING(listNumber, CHARINDEX('OD:', listNumber) + 3,
             CHARINDEX('-COM:', listNumber) - CHARINDEX('OD:', listNumber) - 3)
        ELSE NULL
    END AS OD_Number,
    CASE
        WHEN listNumber LIKE 'OC-%'
        THEN SUBSTRING(listNumber, 4, LEN(listNumber))
        ELSE NULL
    END AS OC_Number,
    CASE
        WHEN CHARINDEX('COM:', listNumber) > 0 AND CHARINDEX('-ART:', listNumber) > 0
             AND CHARINDEX('-ART:', listNumber) > CHARINDEX('COM:', listNumber)
        THEN SUBSTRING(listNumber, CHARINDEX('COM:', listNumber) + 4,
             CHARINDEX('-ART:', listNumber) - CHARINDEX('COM:', listNumber) - 4)
        WHEN CHARINDEX('COM:', listNumber) > 0 AND CHARINDEX('-ID:', listNumber) > 0
             AND CHARINDEX('-ID:', listNumber) > CHARINDEX('COM:', listNumber)
        THEN SUBSTRING(listNumber, CHARINDEX('COM:', listNumber) + 4,
             CHARINDEX('-ID:', listNumber) - CHARINDEX('COM:', listNumber) - 4)
        WHEN CHARINDEX('COM:', listNumber) > 0 AND CHARINDEX('-MG:', listNumber) > 0
             AND CHARINDEX('-MG:', listNumber) > CHARINDEX('COM:', listNumber)
        THEN SUBSTRING(listNumber, CHARINDEX('COM:', listNumber) + 4,
             CHARINDEX('-MG:', listNumber) - CHARINDEX('COM:', listNumber) - 4)
        ELSE NULL
    END AS COM,
    CASE
        WHEN CHARINDEX('-ID:', listNumber) > 0 AND CHARINDEX('-MG:', listNumber) > 0
             AND CHARINDEX('-MG:', listNumber) > CHARINDEX('-ID:', listNumber)
        THEN REPLACE(SUBSTRING(listNumber, CHARINDEX('-ID:', listNumber) + 3,
             CHARINDEX('-MG:', listNumber) - CHARINDEX('-ID:', listNumber) - 3), ':', '')
        ELSE NULL
    END AS ID_Value
FROM [FGWmsErp].[dbo].[LISTS_RESULT]
WHERE id = @id"

        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Cnn.Open()
            Using CMD As New SqlCommand(sql, Cnn)
                CMD.Parameters.AddWithValue("@id", par_id_stampa)
                Using reader As SqlDataReader = CMD.ExecuteReader()
                    If reader.Read() Then
                        stampa_numero_scontrino = S(reader("id"))
                        stampa_OD = S(reader("OD_Number"))
                        stampa_OC = S(reader("OC_Number"))
                        stampa_COM = S(reader("COM"))
                        stampa_ID = S(reader("ID_Value"))
                        stampa_articolo = S(reader("item"))
                        stampa_qta_richiesta = S(reader("requestedQty"))
                        stampa_qta_processata = S(reader("processedQty"))
                        stampa_descrizione = Magazzino.OttieniDettagliAnagrafica(stampa_articolo).Descrizione
                        stampa_progetto = Scheda_tecnica.Ottieni_cliente_papa_macchina(stampa_COM).progetto


                        If stampa_qta_processata = "" OrElse stampa_qta_processata = "0" OrElse stampa_qta_processata = "0.000" Then
                            MessageBox.Show("Nessuna quantità processata, stampa annullata.")
                            Exit Sub
                        End If
                    Else
                        MessageBox.Show("ID non trovato: " & par_id_stampa)
                        Exit Sub
                    End If
                End Using
            End Using
        End Using

        ' --- tipo/numero OC (derivati da stampa_OC già caricato) ---
        stampa_numero_OC = stampa_OC
        stampa_tipo_OC = If(stampa_OC.Length > 0, stampa_OC.Substring(0, 1), "")

        ' --- BAIA, STATO e GRUPPO ARTICOLO dall'OD ---
        If stampa_OD <> "" Then
            Dim info_odp = ODP_Form.ottieni_informazioni_odp("Numero", 0, stampa_OD)
            stampa_baia = info_odp.nome_baia
            stampa_stato_odp = info_odp.stato
            stampa_magazzino_destinazione = info_odp.warehouse
            Stampa_Ubicazione_Macchina = stampa_baia
            If info_odp.itemcode <> "" Then
                codice_gruppo_art_padre = Magazzino.OttieniDettagliAnagrafica(info_odp.itemcode).CodiceGruppo
            Else
                codice_gruppo_art_padre = ""
            End If
        Else
            stampa_baia = ""
            stampa_stato_odp = ""
            stampa_magazzino_destinazione = ""
            Stampa_Ubicazione_Macchina = ""
            codice_gruppo_art_padre = ""
        End If

        ' --- STAMPA ---
        If Stampante_Selezionata = False Then
            Dim sel As New PrintDialog
            sel.Document = Scontrino
            If sel.ShowDialog() = DialogResult.OK Then
                Stampante_Selezionata = True
                Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Scontrino", 185, 250)
                Scontrino.Print()
            End If
        Else
            Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Scontrino", 185, 250)
            Scontrino.Print()
        End If
    End Sub

    Private Sub Scontrino_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles Scontrino.PrintPage

        Dim Penna As New Pen(Color.Black)
        Dim fTitolo As New Font("Calibri", 6, FontStyle.Italic)
        Dim fSmall As New Font("Calibri", 7, FontStyle.Regular)
        Dim fGrande As New Font("Calibri", 14, FontStyle.Bold)
        Dim fArticolo As New Font("Calibri", 14, FontStyle.Italic)
        Dim fQta As New Font("Calibri", 11, FontStyle.Bold)
        Dim fMatricola As New Font("Calibri", 12, FontStyle.Bold)
        Dim fDesc As New Font("Calibri", 8, FontStyle.Italic)

        Dim g As System.Drawing.Graphics = e.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

        ' --- BORDO ESTERNO ---
        g.DrawRectangle(Penna, 1, 1, 183, 245)

        ' --- INTESTAZIONE: numero scontrino ---
        g.DrawString("N° " & stampa_numero_scontrino, fTitolo, Brushes.Gray, 3, 2)

        ' RIGA 1: [N° OD + Stato (sx)] [N° OC (dx)]
        g.DrawRectangle(Penna, 3, 10, 85, 45)
        g.DrawString("N° OD", fTitolo, Brushes.Black, 5, 12)
        g.DrawString(stampa_OD, fGrande, Brushes.Black, 5, 20)
        g.DrawString("Stato: " & stampa_stato_odp, fSmall, Brushes.Gray, 5, 38)

        g.DrawRectangle(Penna, 95, 10, 87, 45)
        g.DrawString("N° OC", fTitolo, Brushes.Black, 97, 12)
        g.DrawString(stampa_OC, fGrande, Brushes.Black, 97, 20)

        ' RIGA 2: [COM (sx 70px)] [ID (40px)] [Mag (dx 52px)]
        g.DrawRectangle(Penna, 3, 60, 70, 35)
        g.DrawString("COM", fTitolo, Brushes.Black, 5, 62)
        g.DrawString(stampa_COM, fGrande, Brushes.Black, 5, 72)

        g.DrawRectangle(Penna, 78, 60, 40, 35)
        g.DrawString("ID", fTitolo, Brushes.Black, 80, 62)
        g.DrawString(stampa_ID, fGrande, Brushes.Black, 80, 72)

        g.DrawRectangle(Penna, 123, 60, 59, 35)
        g.DrawString("Mag", fTitolo, Brushes.Black, 125, 62)
        g.DrawString(stampa_magazzino_destinazione, fSmall, Brushes.Black, 125, 72)

        ' RIGA 3: [Articolo (sx 120px)] [Baia/QE/CDS (dx 62px) — altezza 50 per 2 righe CDS]
        Dim articolo_troncato As String = stampa_articolo.Substring(0, Math.Min(8, stampa_articolo.Length))
        g.DrawRectangle(Penna, 3, 100, 117, 35)
        g.DrawString("Articolo", fTitolo, Brushes.Black, 5, 102)
        g.DrawString(articolo_troncato, fArticolo, Brushes.Black, 5, 110)

        g.DrawRectangle(Penna, 125, 100, 57, 50)
        g.DrawString("Baia", fTitolo, Brushes.Black, 127, 102)

        ' Testo baia: Q.E. / ubicazione macchina / CDS OC — coerente con Form_stampe
        If codice_gruppo_art_padre = "63" Then
            g.DrawString("Q.E.", fMatricola, Brushes.Black, 127, 111)
        ElseIf stampa_tipo_OC = "" OrElse stampa_tipo_OC = " " OrElse stampa_tipo_OC = "B" Then
            g.DrawString(Stampa_Ubicazione_Macchina, fMatricola, Brushes.Black, 127, 111)
        Else
            g.DrawString("CDS " & stampa_tipo_OC, fMatricola, Brushes.Black, 127, 108)
            g.DrawString("OC " & stampa_numero_OC, fDesc, Brushes.Black, 127, 125)
        End If

        ' RIGA 4: Qtà Processata (larghezza piena) — spostata per cella Baia più alta
        g.DrawRectangle(Penna, 3, 155, 179, 30)
        g.DrawString("Qtà Processata", fTitolo, Brushes.Black, 5, 157)
        g.DrawString(stampa_qta_processata, fQta, Brushes.Black, 5, 166)

        ' RIGA 5: Descrizione articolo
        Dim desc_safe As String = If(stampa_descrizione, "")
        Dim desc_troncata As String = desc_safe.Substring(0, Math.Min(35, desc_safe.Length))
        g.DrawRectangle(Penna, 3, 189, 179, 18)
        g.DrawString("Descrizione", fTitolo, Brushes.Black, 5, 190)
        g.DrawString(desc_troncata, fSmall, Brushes.Black, 5, 198)

        ' RIGA 6: Progetto
        Dim prog_safe As String = If(stampa_progetto, "")
        Dim progetto_troncato As String = prog_safe.Substring(0, Math.Min(35, prog_safe.Length))
        g.DrawRectangle(Penna, 3, 211, 179, 18)
        g.DrawString("Progetto", fTitolo, Brushes.Black, 5, 212)
        g.DrawString(progetto_troncato, fSmall, Brushes.Black, 5, 220)

        ' FOOTER: data/ora + utente (una sola query)
        Dim DataOggi As String = DateTime.Now.ToString("dd/MM/yy HH:mm")
        g.DrawString(DataOggi, fSmall, Brushes.Black, 120, 232)

        Dim d = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO)
        Dim NomeTroncato As String = (d.COGNOME & " " & d.NOME).Substring(0, Math.Min(12, (d.COGNOME & " " & d.NOME).Length))
        g.DrawString(NomeTroncato, fSmall, Brushes.Black, 5, 232)

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim id_auto As Integer = max_id_ferretto()
        If id_auto <> id_ultimo_timer Then
            id_ultimo_timer = id_auto
            Stampa_stampa(id_auto)
        End If
    End Sub

    Public Function max_id_ferretto() As Integer
        Dim id As Integer = 0
        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Cnn.Open()
            Using CMD As New SqlCommand(
                "SELECT TOP 1 [id] FROM [FGWmsErp].[dbo].[LISTS_RESULT]
                 WHERE response <> 11 AND listtype = 0 AND [processedQty] > 0
                 ORDER BY id DESC", Cnn)
                Using reader As SqlDataReader = CMD.ExecuteReader()
                    If reader.Read() Then id = CInt(reader("id"))
                End Using
            End Using
        End Using
        Return id
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        id_ultimo_timer = max_id_ferretto()   ' allinea subito per non stampare al primo tick
        Timer1.Interval = CInt(TextBox7.Text)
        Timer1.Start()
        Button3.Visible = True
        Panel1.BackColor = Color.Lime
        Button2.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Timer1.Stop()
        Button2.Visible = True
        Button3.Visible = False
        Panel1.BackColor = Color.IndianRed
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim ms As Integer
        If Integer.TryParse(TextBox7.Text, ms) Then
            Timer1.Interval = ms
        Else
            MessageBox.Show("Inserire un valore numerico in millisecondi.")
        End If
    End Sub
End Class
