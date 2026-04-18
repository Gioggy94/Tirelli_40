Imports System.Data
Imports System.Data.SqlClient

Public Class Solleciti_OA

    Private _datiOA As DataTable
    Private _codFornSelezionato As String = ""
    Private _descFornSelezionato As String = ""
    Private _aggiornandoAnteprima As Boolean = False
    Private _htmlAnteprima As String = ""
    Private _fornSollecito As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

    Private Sub Solleciti_OA_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CreaTabellaFornitatoriSollecito()
        CreaTabellaLog()
        CaricaFornitatoriSollecito()
        ImpostaListaFornitori()
        ImpostaGriglia()
        ImpostaGrigliaStatistiche()
        ImpostaGrigliaLog()
        AggiornaLog()
    End Sub

    Private Sub Solleciti_OA_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        scStatistiche.SplitterDistance = scStatistiche.Width \ 2
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' Setup UI
    ' ─────────────────────────────────────────────────────────

    Private Sub ImpostaListaFornitori()
        lvFornitori.Columns.Clear()
        lvFornitori.View = View.Details
        lvFornitori.FullRowSelect = True
        lvFornitori.GridLines = True
        lvFornitori.MultiSelect = True
        lvFornitori.Columns.Add("Sol.", 38, HorizontalAlignment.Center)
        lvFornitori.Columns.Add("Fornitore", 185)
        lvFornitori.Columns.Add("Righe", 50, HorizontalAlignment.Right)
        lvFornitori.Columns.Add("Scad.", 50, HorizontalAlignment.Right)
        lvFornitori.Columns.Add("% Sc.", 52, HorizontalAlignment.Right)
    End Sub

    Private Sub ImpostaGriglia()
        dgvOrdini.Columns.Clear()
        dgvOrdini.AutoGenerateColumns = False
        dgvOrdini.AllowUserToAddRows = False
        dgvOrdini.AllowUserToDeleteRows = False
        dgvOrdini.ReadOnly = False
        dgvOrdini.MultiSelect = True
        dgvOrdini.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvOrdini.RowHeadersVisible = False
        dgvOrdini.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255)
        dgvOrdini.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue
        dgvOrdini.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvOrdini.EnableHeadersVisualStyles = False
        dgvOrdini.ColumnHeadersHeight = 26

        Dim colSel As New DataGridViewCheckBoxColumn
        colSel.Name = "colSel"
        colSel.HeaderText = ""
        colSel.Width = 28
        colSel.ReadOnly = False
        dgvOrdini.Columns.Add(colSel)

        Dim aggiungiCol = Sub(nome As String, header As String, larghezza As Integer)
                              Dim c As New DataGridViewTextBoxColumn
                              c.Name = nome
                              c.HeaderText = header
                              c.Width = larghezza
                              c.ReadOnly = True
                              dgvOrdini.Columns.Add(c)
                          End Sub

        aggiungiCol("colNumdoc", "N. Ordine", 95)
        aggiungiCol("colAcquisitore", "Acquisitore", 100)
        aggiungiCol("colCommessa", "Commessa", 90)
        aggiungiCol("colSottocommessa", "Sottocom.", 75)
        aggiungiCol("colMatricola", "Matricola", 85)
        aggiungiCol("colCodart", "Codice", 105)
        aggiungiCol("colDesCode", "Descrizione", 215)
        aggiungiCol("colDataImmissione", "Data ord.", 85)
        aggiungiCol("colDataRichiesta", "Data consegna", 95)
        aggiungiCol("colQtaOrd", "Q.ord.", 70)
        aggiungiCol("colQtaEnt", "Q.ric.", 70)
        aggiungiCol("colQtaRes", "Q.res.", 70)
        aggiungiCol("colDisegno", "Disegno", 105)
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' Utility: parsing data AS400 (formato YYYYMMDD intero)
    ' ─────────────────────────────────────────────────────────

    Private Function ParseDataAS400(valore As Object) As Date?
        If IsDBNull(valore) OrElse valore Is Nothing Then Return Nothing
        Dim s As String = valore.ToString().Trim()
        If s.Length = 8 Then
            Try
                Dim y = Integer.Parse(s.Substring(0, 4))
                Dim m = Integer.Parse(s.Substring(4, 2))
                Dim d = Integer.Parse(s.Substring(6, 2))
                If y > 1900 AndAlso m >= 1 AndAlso m <= 12 AndAlso d >= 1 AndAlso d <= 31 Then
                    Return New Date(y, m, d)
                End If
            Catch
            End Try
        End If
        Return Nothing
    End Function

    ' ─────────────────────────────────────────────────────────
    ' Carica dati da AS400
    ' ─────────────────────────────────────────────────────────

    Sub CaricaDati()
        Cursor = Cursors.WaitCursor
        lblStato.Text = "Caricamento dati AS400..."
        lblStato.Refresh()
        Application.DoEvents()
        Try
            Dim CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Dim CMD As New SqlCommand
            CMD.Connection = CNN
            CMD.CommandTimeout = 120
            CMD.CommandText =
                "SELECT * FROM OPENQUERY(AS400,
'SELECT DOC, NUMDOC, CODART, DES_CODE, DATA_IMMISSIONE, DATA_RICHIESTA,
 QTA_ORD, QTA_ENT, EVASO, ID_COMM, MATRICOLA, COD_FORN, DESC_FOR,
 MAG_ORD, STAMPATO, DISEGNO, COMMESSA, SOTTOCOMM, DIP_INS, ACQUISITORE
 FROM TIR90VIS.JGALORD_03
 WHERE DOC = ''OA'' and (substring(codart,1,1)=''0'' or substring(codart,1,1)=''C'' or substring(codart,1,1)=''D'') 
   AND EVASO <> ''S'''  -- Qui servono 3 apici totali
)"
            Dim DA As New SqlDataAdapter(CMD)
            _datiOA = New DataTable
            DA.Fill(_datiOA)
            CNN.Close()

            lblStato.Text = _datiOA.Rows.Count & " righe caricate"
            AggiornaTabellaFornitori()
        Catch ex As Exception
            lblStato.Text = "Errore: " & ex.Message
            MessageBox.Show("Errore caricamento dati:" & vbCrLf & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Cursor = Cursors.Default
        End Try
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' Aggiorna lista fornitori a sinistra
    ' ─────────────────────────────────────────────────────────

    Sub AggiornaTabellaFornitori()
        If _datiOA Is Nothing Then Return
        Dim oggi = DateTime.Today
        Dim filtroComm = txtFiltroCommessa.Text.Trim().ToUpper()
        Dim soloScaduti = chkSoloScaduti.Checked
        Dim soloSollecito = chkSoloSollecito.Checked
        Dim filtroForn = If(cmbFiltroFornitore.SelectedIndex > 0, CStr(cmbFiltroFornitore.SelectedItem), "")
        Dim filtroAcq = If(cmbFiltroAcquisitore.SelectedIndex > 0, CStr(cmbFiltroAcquisitore.SelectedItem), "")

        Dim gruppi = _datiOA.AsEnumerable().
            Where(Function(r)
                      If filtroComm <> "" AndAlso Not r("commessa").ToString().Trim().ToUpper().Contains(filtroComm) Then Return False
                      If filtroForn <> "" AndAlso r("desc_for").ToString().Trim() <> filtroForn Then Return False
                      If filtroAcq <> "" AndAlso r("acquisitore").ToString().Trim() <> filtroAcq Then Return False
                      Return True
                  End Function).
            GroupBy(Function(r) r("cod_forn").ToString().Trim()).
            Select(Function(g)
                       Dim scadute = g.Where(Function(r)
                                                 Dim dr = ParseDataAS400(r("data_richiesta"))
                                                 Return dr.HasValue AndAlso dr.Value < oggi
                                             End Function).Count()
                       Return New With {
                           .CodForn = g.Key,
                           .DescFor = g.First()("desc_for").ToString().Trim(),
                           .TotRighe = g.Count(),
                           .RigheScadute = scadute,
                           .PctScadute = If(g.Count() > 0, Math.Round(scadute * 100.0 / g.Count(), 0), 0.0)
                       }
                   End Function).
            Where(Function(f) (Not soloScaduti OrElse f.RigheScadute > 0) AndAlso
                              (Not soloSollecito OrElse _fornSollecito.Contains(f.CodForn))).
            OrderBy(Function(f) f.DescFor).
            ToList()

        ' Rigenera combo fornitori (solo all'avvio o su richiesta esplicita)
        If cmbFiltroFornitore.Items.Count <= 1 Then
            cmbFiltroFornitore.Items.Clear()
            cmbFiltroFornitore.Items.Add("(tutti)")
            For Each desc In _datiOA.AsEnumerable().
                    Select(Function(r) r("desc_for").ToString().Trim()).
                    Distinct().OrderBy(Function(s) s)
                cmbFiltroFornitore.Items.Add(desc)
            Next
            cmbFiltroFornitore.SelectedIndex = 0
        End If

        ' Rigenera combo acquisitore (solo all'avvio o su richiesta esplicita)
        If cmbFiltroAcquisitore.Items.Count <= 1 Then
            cmbFiltroAcquisitore.Items.Clear()
            cmbFiltroAcquisitore.Items.Add("(tutti)")
            For Each acq In _datiOA.AsEnumerable().
                    Select(Function(r) r("acquisitore").ToString().Trim()).
                    Distinct().OrderBy(Function(s) s)
                cmbFiltroAcquisitore.Items.Add(acq)
            Next
            cmbFiltroAcquisitore.SelectedIndex = 0
        End If

        lvFornitori.Items.Clear()
        For Each f In gruppi
            Dim inSollecito = _fornSollecito.Contains(f.CodForn)
            Dim item As New ListViewItem(If(inSollecito, "★", ""))
            item.SubItems.Add(f.DescFor & " [" & f.CodForn & "]")
            item.SubItems.Add(f.TotRighe.ToString())
            item.SubItems.Add(f.RigheScadute.ToString())
            item.SubItems.Add(If(f.TotRighe > 0, f.PctScadute.ToString("F0") & "%", "-"))
            item.Tag = f.CodForn
            If inSollecito Then
                item.BackColor = Color.FromArgb(230, 255, 230)
                item.Font = New Font(lvFornitori.Font, FontStyle.Bold)
            End If
            If f.RigheScadute > 0 Then item.ForeColor = Color.DarkRed
            lvFornitori.Items.Add(item)
        Next
        lblConteggioFornitori.Text = gruppi.Count & " fornitori  (" & _fornSollecito.Count & " da sollecitare)"
        AggiornaStatistiche()
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' Aggiorna griglia ordini per il fornitore selezionato
    ' ─────────────────────────────────────────────────────────

    Sub AggiornaTabellaOrdini(codForn As String)
        _aggiornandoAnteprima = True
        dgvOrdini.Rows.Clear()
        _codFornSelezionato = codForn

        If _datiOA Is Nothing OrElse codForn = "" Then
            _aggiornandoAnteprima = False
            Return
        End If

        Dim oggi = DateTime.Today
        Dim filtroComm = txtFiltroCommessa.Text.Trim().ToUpper()

        Dim filtroAcq = If(cmbFiltroAcquisitore.SelectedIndex > 0, CStr(cmbFiltroAcquisitore.SelectedItem), "")

        Dim righe = _datiOA.AsEnumerable().
            Where(Function(r)
                      If r("cod_forn").ToString().Trim() <> codForn Then Return False
                      If filtroComm <> "" AndAlso Not r("commessa").ToString().Trim().ToUpper().Contains(filtroComm) Then Return False
                      If filtroAcq <> "" AndAlso r("acquisitore").ToString().Trim() <> filtroAcq Then Return False
                      Return True
                  End Function).
            OrderBy(Function(r) r("data_richiesta").ToString()).
            ToList()

        For Each r In righe
            Dim qtaOrd = 0.0
            Dim qtaEnt = 0.0
            Try : qtaOrd = CDbl(r("qta_ord")) : Catch : End Try
            Try : qtaEnt = CDbl(r("qta_ent")) : Catch : End Try

            Dim dataRich = ParseDataAS400(r("data_richiesta"))
            Dim dataImmiss = ParseDataAS400(r("data_immissione"))
            Dim scaduta = dataRich.HasValue AndAlso dataRich.Value < oggi

            ' Applica filtro "solo scaduti" anche alla griglia ordini
            If chkSoloScaduti.Checked AndAlso Not scaduta Then Continue For

            Dim idx = dgvOrdini.Rows.Add(
                True,
                r("numdoc").ToString().Trim(),
                r("acquisitore").ToString().Trim(),
                r("commessa").ToString().Trim(),
                r("sottocomm").ToString().Trim(),
                r("matricola").ToString().Trim(),
                r("codart").ToString().Trim(),
                r("des_code").ToString().Trim(),
                If(dataImmiss.HasValue, dataImmiss.Value.ToString("dd/MM/yyyy"), ""),
                If(dataRich.HasValue, dataRich.Value.ToString("dd/MM/yyyy"), ""),
                qtaOrd.ToString("N0"),
                qtaEnt.ToString("N0"),
                (qtaOrd - qtaEnt).ToString("N0"),
                r("disegno").ToString().Trim()
            )

            If scaduta Then
                With dgvOrdini.Rows(idx).DefaultCellStyle
                    .BackColor = Color.FromArgb(255, 205, 205)
                    .ForeColor = Color.DarkRed
                End With
            End If
        Next

        Dim totScadute = righe.Where(Function(r)
                                         Dim dr = ParseDataAS400(r("data_richiesta"))
                                         Return dr.HasValue AndAlso dr.Value < oggi
                                     End Function).Count()
        lblConteggio.Text = righe.Count & " righe  (" & totScadute & " scadute)"

        _aggiornandoAnteprima = False
        AggiornaAnteprima()
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' Anteprima email
    ' ─────────────────────────────────────────────────────────

    Sub AggiornaAnteprima()
        If _aggiornandoAnteprima Then Return

        Dim descFor = _descFornSelezionato
        If descFor = "" AndAlso _codFornSelezionato <> "" Then
            For Each item As ListViewItem In lvFornitori.Items
                If item.Tag.ToString() = _codFornSelezionato Then
                    descFor = item.Text
                    Exit For
                End If
            Next
        End If

        Dim righeSelezionate As New List(Of DataGridViewRow)
        For Each row As DataGridViewRow In dgvOrdini.Rows
            If row.Cells("colSel").Value IsNot Nothing AndAlso CBool(row.Cells("colSel").Value) Then
                righeSelezionate.Add(row)
            End If
        Next

        If righeSelezionate.Count = 0 Then
            rtbAnteprima.Text = "(nessuna riga selezionata)"
            _htmlAnteprima = ""
            Return
        End If

        If txtOggetto.Text = "" Then
            txtOggetto.Text = "Sollecito ordini di acquisto - Tirelli S.r.l."
        End If

        ' ── Genera HTML ──────────────────────────────────────────────────
        Dim hsb As New System.Text.StringBuilder
        hsb.AppendLine("<html><body style='font-family:Calibri,Arial,sans-serif;font-size:11pt;'>")
        hsb.AppendLine($"<p>Spett.le {Net.WebUtility.HtmlEncode(descFor)},</p>")
        hsb.AppendLine("<p>con la presente Vi sollecitiamo la consegna dei seguenti ordini di acquisto ancora in attesa di evasione:</p>")
        hsb.AppendLine("<table border='1' cellpadding='5' cellspacing='0' style='border-collapse:collapse;font-size:10pt;'>")
        hsb.AppendLine("<tr style='background-color:#1F5FAF;color:white;font-weight:bold;'>")
        For Each h In {"N. Ordine", "Commessa", "Sottocom.", "Matricola", "Codice", "Disegno", "Descrizione", "Data ordine", "Data consegna", "Q.ord.", "Q.ric.", "Q.res."}
            hsb.AppendLine($"<th style='padding:5px 8px;'>{h}</th>")
        Next
        hsb.AppendLine("</tr>")

        Dim oggi = DateTime.Today
        Dim txtSb As New System.Text.StringBuilder  ' testo per anteprima rtb
        txtSb.AppendLine($"A: {descFor}")
        txtSb.AppendLine("Oggetto: " & txtOggetto.Text)
        txtSb.AppendLine()
        txtSb.AppendLine(String.Format("{0,-10} {1,-10} {2,-8} {3,-10} {4,-12} {5,-26} {6,-12} {7,-12} {8,6} {9,6} {10,6}",
                                        "N.Ordine", "Commessa", "Sottocom", "Matricola", "Codice", "Descrizione",
                                        "Dt.Ordine", "Dt.Consegna", "Q.ord", "Q.ric", "Q.res"))
        txtSb.AppendLine(New String("-"c, 115))

        For Each row In righeSelezionate
            Dim numdoc = If(TryCast(row.Cells("colNumdoc").Value, String), "")
            Dim comm = If(TryCast(row.Cells("colCommessa").Value, String), "")
            Dim sotto = If(TryCast(row.Cells("colSottocommessa").Value, String), "")
            Dim matr = If(TryCast(row.Cells("colMatricola").Value, String), "")
            Dim codart = If(TryCast(row.Cells("colCodart").Value, String), "")
            Dim disegno = If(TryCast(row.Cells("colDisegno").Value, String), "")
            Dim desc = If(TryCast(row.Cells("colDesCode").Value, String), "")
            Dim dtOrd = If(TryCast(row.Cells("colDataImmissione").Value, String), "")
            Dim dtCons = If(TryCast(row.Cells("colDataRichiesta").Value, String), "")
            Dim qtaOrd = If(TryCast(row.Cells("colQtaOrd").Value, String), "")
            Dim qtaEnt = If(TryCast(row.Cells("colQtaEnt").Value, String), "")
            Dim qtaRes = If(TryCast(row.Cells("colQtaRes").Value, String), "")

            ' Scaduta = data consegna in passato
            Dim dtConsDate As Date
            Dim scaduta = Date.TryParse(dtCons, dtConsDate) AndAlso dtConsDate < oggi
            Dim rowBg = If(scaduta, "background-color:#FFD0D0;", "")

            hsb.AppendLine($"<tr style='{rowBg}'>")
            For Each v In {numdoc, comm, sotto, matr, codart, disegno, desc, dtOrd, dtCons, qtaOrd, qtaEnt, qtaRes}
                hsb.AppendLine($"<td style='padding:4px 8px;'>{Net.WebUtility.HtmlEncode(v)}</td>")
            Next
            hsb.AppendLine("</tr>")

            Dim descShort = If(desc.Length > 24, desc.Substring(0, 24), desc)
            txtSb.AppendLine(String.Format("{0,-10} {1,-10} {2,-8} {3,-10} {4,-12} {5,-26} {6,-12} {7,-12} {8,6} {9,6} {10,6}",
                                            numdoc, comm, sotto, matr, codart, descShort, dtOrd, dtCons, qtaOrd, qtaEnt, qtaRes))
        Next

        hsb.AppendLine("</table>")
        hsb.AppendLine("<p>Vi chiediamo cortesemente di confermarci la data di consegna prevista per ciascun articolo.</p>")
        hsb.AppendLine("<p>In attesa di un vostro riscontro, porgiamo distinti saluti.<br/><br/>Tirelli S.r.l. - Ufficio Acquisti</p>")
        hsb.AppendLine("</body></html>")

        _htmlAnteprima = hsb.ToString()
        rtbAnteprima.Text = txtSb.ToString()
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' Setup griglia statistiche
    ' ─────────────────────────────────────────────────────────

    Sub ImpostaGrigliaStatistiche()
        dgvStatAcquisitore.Columns.Clear()
        dgvStatAcquisitore.AutoGenerateColumns = False
        dgvStatAcquisitore.AllowUserToAddRows = False
        dgvStatAcquisitore.ReadOnly = True
        dgvStatAcquisitore.RowHeadersVisible = False
        dgvStatAcquisitore.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvStatAcquisitore.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue
        dgvStatAcquisitore.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvStatAcquisitore.EnableHeadersVisualStyles = False
        dgvStatAcquisitore.BackgroundColor = Color.White
        dgvStatAcquisitore.BorderStyle = BorderStyle.None

        Dim addCol = Sub(nome As String, header As String, larghezza As Integer, align As DataGridViewContentAlignment)
                         Dim c As New DataGridViewTextBoxColumn
                         c.Name = nome
                         c.HeaderText = header
                         c.Width = larghezza
                         c.ReadOnly = True
                         c.DefaultCellStyle.Alignment = align
                         dgvStatAcquisitore.Columns.Add(c)
                     End Sub

        addCol("colAcq", "Acquisitore", 180, DataGridViewContentAlignment.MiddleLeft)
        addCol("colAcqTot", "Tot. Righe", 80, DataGridViewContentAlignment.MiddleRight)
        addCol("colAcqScad", "Scadute", 75, DataGridViewContentAlignment.MiddleRight)
        addCol("colAcqPct", "% Scad.", 70, DataGridViewContentAlignment.MiddleRight)
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' Setup e caricamento griglia log solleciti
    ' ─────────────────────────────────────────────────────────

    Sub ImpostaGrigliaLog()
        dgvLog.Columns.Clear()
        dgvLog.AutoGenerateColumns = False
        dgvLog.AllowUserToAddRows = False
        dgvLog.ReadOnly = True
        dgvLog.RowHeadersVisible = False
        dgvLog.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvLog.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue
        dgvLog.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvLog.EnableHeadersVisualStyles = False
        dgvLog.BackgroundColor = Color.White
        dgvLog.BorderStyle = BorderStyle.None
        dgvLog.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255)

        Dim addCol = Sub(nome As String, header As String, larghezza As Integer, align As DataGridViewContentAlignment)
                         Dim c As New DataGridViewTextBoxColumn
                         c.Name = nome
                         c.HeaderText = header
                         c.Width = larghezza
                         c.ReadOnly = True
                         c.DefaultCellStyle.Alignment = align
                         dgvLog.Columns.Add(c)
                     End Sub

        addCol("logDataOra", "Data / Ora", 130, DataGridViewContentAlignment.MiddleCenter)
        addCol("logUtente", "Utente", 110, DataGridViewContentAlignment.MiddleLeft)
        addCol("logCodForn", "Cod. Forn.", 100, DataGridViewContentAlignment.MiddleLeft)
        addCol("logDescFor", "Fornitore", 200, DataGridViewContentAlignment.MiddleLeft)
        addCol("logEmail", "Email", 210, DataGridViewContentAlignment.MiddleLeft)
        addCol("logNRighe", "N. Righe", 70, DataGridViewContentAlignment.MiddleRight)
    End Sub

    Sub AggiornaLog()
        Try
            dgvLog.Rows.Clear()
            Using cn As New SqlConnection(Homepage.sap_tirelli)
                cn.Open()
                Dim sql =
                    "SELECT TOP 200 DataOra, Utente, CodForn, DescFor, Email, NRighe
                     FROM [Tirelli_40].dbo.SollecitiLog
                     ORDER BY DataOra DESC"
                Using cmd As New SqlCommand(sql, cn)
                    Using rd = cmd.ExecuteReader()
                        While rd.Read()
                            dgvLog.Rows.Add(
                                CDate(rd("DataOra")).ToString("dd/MM/yyyy HH:mm"),
                                rd("Utente").ToString(),
                                rd("CodForn").ToString(),
                                rd("DescFor").ToString(),
                                rd("Email").ToString(),
                                rd("NRighe").ToString())
                        End While
                    End Using
                End Using
            End Using
        Catch
        End Try
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' Statistiche scaduti globali e per acquisitore
    ' ─────────────────────────────────────────────────────────

    Sub AggiornaStatistiche()
        If _datiOA Is Nothing Then
            lblStatGenerale.Text = ""
            dgvStatAcquisitore.Rows.Clear()
            Return
        End If

        Dim oggi = DateTime.Today
        Dim filtroComm = txtFiltroCommessa.Text.Trim().ToUpper()
        Dim filtroForn = If(cmbFiltroFornitore.SelectedIndex > 0, CStr(cmbFiltroFornitore.SelectedItem), "")
        Dim filtroAcq = If(cmbFiltroAcquisitore.SelectedIndex > 0, CStr(cmbFiltroAcquisitore.SelectedItem), "")

        Dim righe = _datiOA.AsEnumerable().
            Where(Function(r)
                      If filtroComm <> "" AndAlso Not r("commessa").ToString().Trim().ToUpper().Contains(filtroComm) Then Return False
                      If filtroForn <> "" AndAlso r("desc_for").ToString().Trim() <> filtroForn Then Return False
                      If filtroAcq <> "" AndAlso r("acquisitore").ToString().Trim() <> filtroAcq Then Return False
                      Return True
                  End Function).ToList()

        Dim totRighe = righe.Count
        Dim totScadute = righe.Where(Function(r)
                                         Dim dr = ParseDataAS400(r("data_richiesta"))
                                         Return dr.HasValue AndAlso dr.Value < oggi
                                     End Function).Count()
        Dim pctTot = If(totRighe > 0, Math.Round(totScadute * 100.0 / totRighe, 1), 0.0)

        lblStatGenerale.Text = $"Totale righe OA: {totRighe}     Scadute: {totScadute}     ({pctTot}%)"

        Dim gruppiAcq = righe.
            GroupBy(Function(r) r("acquisitore").ToString().Trim()).
            Select(Function(g)
                       Dim scadute = g.Where(Function(r)
                                                 Dim dr = ParseDataAS400(r("data_richiesta"))
                                                 Return dr.HasValue AndAlso dr.Value < oggi
                                             End Function).Count()
                       Return New With {
                           .Acquisitore = If(g.Key = "", "(non assegnato)", g.Key),
                           .TotRighe = g.Count(),
                           .Scadute = scadute,
                           .Pct = If(g.Count() > 0, Math.Round(scadute * 100.0 / g.Count(), 1), 0.0)
                       }
                   End Function).
            OrderByDescending(Function(a) a.Scadute).
            ToList()

        dgvStatAcquisitore.Rows.Clear()
        For Each a In gruppiAcq
            Dim idx = dgvStatAcquisitore.Rows.Add(a.Acquisitore, a.TotRighe, a.Scadute, a.Pct.ToString("F1") & "%")
            If a.Scadute > 0 Then
                dgvStatAcquisitore.Rows(idx).DefaultCellStyle.BackColor = Color.FromArgb(255, 215, 215)
            End If
        Next
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' Lookup email fornitore da AS400 (VISTA_CONTATTI_FORNITORI)
    ' ─────────────────────────────────────────────────────────

    Private Function LookupEmailFornitore(codForn As String) As String
        If String.IsNullOrWhiteSpace(codForn) Then Return ""
        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnn
                    cmd.CommandTimeout = 30
                    cmd.CommandText = String.Format(
                        "SELECT TOP 1 INDIRIZZO_EMAIL " &
                        "FROM OPENQUERY(AS400, " &
                        "'SELECT INDIRIZZO_EMAIL FROM TIR90VIS.VISTA_CONTATTI_FORNITORI " &
                        " WHERE MODALITA_SPEDIZIONE=''P'' AND trim(CONTO_CONTABILE)=''{0}''')",
                        codForn.Trim().Replace("'", "''"))
                    Dim result = cmd.ExecuteScalar()
                    If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                        Return result.ToString().Trim()
                    End If
                End Using
            End Using
        Catch
        End Try
        Return ""
    End Function

    ' ─────────────────────────────────────────────────────────
    ' Prepara mail Outlook per il fornitore selezionato
    ' ─────────────────────────────────────────────────────────

    Sub PreparaMail()
        If String.IsNullOrWhiteSpace(_htmlAnteprima) Then
            MessageBox.Show("Nessuna riga selezionata. Selezionare almeno una riga.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Try
            Dim objOutlook As Object = CreateObject("Outlook.Application")
            Dim objMail As Object = objOutlook.CreateItem(0)
            Dim accountAcquisti As Object = Nothing
            For Each acc As Object In objOutlook.Session.Accounts
                If acc.SmtpAddress.ToLower() = "acquisti@tirelli.net" Then
                    accountAcquisti = acc
                    Exit For
                End If
            Next
            Dim nRigheSelezionate As Integer = 0
            For Each row As DataGridViewRow In dgvOrdini.Rows
                If row.Cells("colSel").Value IsNot Nothing AndAlso CBool(row.Cells("colSel").Value) Then
                    nRigheSelezionate += 1
                End If
            Next
            With objMail
                If accountAcquisti IsNot Nothing Then .SendUsingAccount = accountAcquisti
                .To = txtEmail.Text.Trim()
                .BCC = "giovanni.tirelli@tirelli.net; stefano.bruno@tirelli.net"
                .Subject = txtOggetto.Text
                .HTMLBody = _htmlAnteprima
                .Display()
            End With
            LogSollecito(_codFornSelezionato, _descFornSelezionato, txtEmail.Text.Trim(), nRigheSelezionate)
            AggiornaLog()
            objMail = Nothing
            objOutlook = Nothing
        Catch ex As Exception
            MessageBox.Show("Errore apertura Outlook: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' Prepara mail per tutti i fornitori selezionati in lista
    ' ─────────────────────────────────────────────────────────

    Sub PreparaTutteMail()
        If _datiOA Is Nothing OrElse lvFornitori.Items.Count = 0 Then
            MessageBox.Show("Caricare prima i dati.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim itemsDaElaborare As New List(Of ListViewItem)
        If lvFornitori.SelectedItems.Count > 0 Then
            For Each item As ListViewItem In lvFornitori.SelectedItems
                itemsDaElaborare.Add(item)
            Next
        Else
            For Each item As ListViewItem In lvFornitori.Items
                itemsDaElaborare.Add(item)
            Next
        End If

        Dim risposta = MessageBox.Show(
            "Verranno create " & itemsDaElaborare.Count & " bozze email (una per fornitore)." & vbCrLf &
            "Continuare?",
            "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If risposta <> DialogResult.Yes Then Return

        Dim savedCod = _codFornSelezionato
        Dim savedDesc = _descFornSelezionato
        txtOggetto.Text = "Sollecito ordini di acquisto - Tirelli S.r.l."

        For Each item As ListViewItem In itemsDaElaborare
            Dim codForn = item.Tag.ToString()
            _descFornSelezionato = item.Text
            AggiornaTabellaOrdini(codForn)
            For Each row As DataGridViewRow In dgvOrdini.Rows
                row.Cells("colSel").Value = True
            Next
            txtEmail.Text = LookupEmailFornitore(codForn)
            AggiornaAnteprima()
            PreparaMail()
        Next

        ' Ripristina selezione precedente
        _descFornSelezionato = savedDesc
        If savedCod <> "" Then AggiornaTabellaOrdini(savedCod)
        MessageBox.Show("Create " & itemsDaElaborare.Count & " bozze email.", "Completato", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' Event handlers
    ' ─────────────────────────────────────────────────────────

    Private Sub btnCarica_Click(sender As Object, e As EventArgs) Handles btnCarica.Click
        CaricaDati()
    End Sub

    Private Sub btnAggiornaLog_Click(sender As Object, e As EventArgs) Handles btnAggiornaLog.Click
        AggiornaLog()
    End Sub

    Private Sub lvFornitori_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvFornitori.SelectedIndexChanged
        If lvFornitori.SelectedItems.Count = 0 Then Return
        Dim item = lvFornitori.SelectedItems(0)
        Dim codForn = item.Tag.ToString()
        _descFornSelezionato = item.Text
        AggiornaTabellaOrdini(codForn)
        ' Precompila destinatario email da AS400
        Dim email = LookupEmailFornitore(codForn)
        txtEmail.Text = email
        If email = "" Then
            txtEmail.BackColor = System.Drawing.Color.FromArgb(255, 230, 230)  ' rosato se non trovata
        Else
            txtEmail.BackColor = System.Drawing.Color.FromArgb(230, 255, 230)  ' verde se trovata
        End If
    End Sub

    Private Sub btnSelTutti_Click(sender As Object, e As EventArgs) Handles btnSelTutti.Click
        _aggiornandoAnteprima = True
        For Each row As DataGridViewRow In dgvOrdini.Rows
            row.Cells("colSel").Value = True
        Next
        _aggiornandoAnteprima = False
        AggiornaAnteprima()
    End Sub

    Private Sub btnDeselTutti_Click(sender As Object, e As EventArgs) Handles btnDeselTutti.Click
        _aggiornandoAnteprima = True
        For Each row As DataGridViewRow In dgvOrdini.Rows
            row.Cells("colSel").Value = False
        Next
        _aggiornandoAnteprima = False
        AggiornaAnteprima()
    End Sub

    Private Sub btnAggiornaAnteprima_Click(sender As Object, e As EventArgs) Handles btnAggiornaAnteprima.Click
        AggiornaAnteprima()
    End Sub

    Private Sub btnPreparaMail_Click(sender As Object, e As EventArgs) Handles btnPreparaMail.Click
        If lvFornitori.SelectedItems.Count <= 1 Then
            ' Comportamento normale: un solo fornitore
            PreparaMail()
        Else
            ' Più fornitori selezionati → una bozza per ciascuno
            Dim savedCod = _codFornSelezionato
            Dim savedDesc = _descFornSelezionato
            Dim savedEmail = txtEmail.Text

            For Each item As ListViewItem In lvFornitori.SelectedItems
                Dim codForn = item.Tag.ToString()
                _descFornSelezionato = item.SubItems(1).Text
                AggiornaTabellaOrdini(codForn)
                For Each row As DataGridViewRow In dgvOrdini.Rows
                    row.Cells("colSel").Value = True
                Next
                txtEmail.Text = LookupEmailFornitore(codForn)
                AggiornaAnteprima()
                PreparaMail()
            Next

            ' Ripristina selezione precedente
            _descFornSelezionato = savedDesc
            txtEmail.Text = savedEmail
            If savedCod <> "" Then AggiornaTabellaOrdini(savedCod)
        End If
    End Sub

    Private Sub btnTutteMail_Click(sender As Object, e As EventArgs) Handles btnTutteMail.Click
        PreparaTutteMail()
    End Sub

    Private Sub dgvOrdini_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvOrdini.CurrentCellDirtyStateChanged
        If dgvOrdini.IsCurrentCellDirty Then
            dgvOrdini.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub dgvOrdini_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvOrdini.CellValueChanged
        If e.ColumnIndex = 0 AndAlso Not _aggiornandoAnteprima Then
            AggiornaAnteprima()
        End If
    End Sub

    Private Sub chkSoloScaduti_CheckedChanged(sender As Object, e As EventArgs) Handles chkSoloScaduti.CheckedChanged
        If _datiOA IsNot Nothing Then AggiornaTabellaFornitori()
    End Sub

    Private Sub chkSoloSollecito_CheckedChanged(sender As Object, e As EventArgs) Handles chkSoloSollecito.CheckedChanged
        If _datiOA IsNot Nothing Then AggiornaTabellaFornitori()
    End Sub

    Private Sub txtFiltroCommessa_TextChanged(sender As Object, e As EventArgs) Handles txtFiltroCommessa.TextChanged
        If _datiOA IsNot Nothing Then AggiornaTabellaFornitori()
    End Sub

    Private Sub cmbFiltroFornitore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFiltroFornitore.SelectedIndexChanged
        If _datiOA IsNot Nothing Then AggiornaTabellaFornitori()
    End Sub

    Private Sub cmbFiltroAcquisitore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFiltroAcquisitore.SelectedIndexChanged
        If _datiOA IsNot Nothing Then AggiornaTabellaFornitori()
    End Sub

    Private Sub btnToggleSollecito_Click(sender As Object, e As EventArgs) Handles btnToggleSollecito.Click
        If lvFornitori.SelectedItems.Count = 0 Then
            MessageBox.Show("Seleziona un fornitore dalla lista.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If
        Dim item = lvFornitori.SelectedItems(0)
        Dim codForn = item.Tag.ToString()
        Dim descFor = item.SubItems(1).Text  ' SubItems(0)=Sol., SubItems(1)=nome fornitore
        Try
            ToggleFornitoreSollecito(codForn, descFor)
            CaricaFornitatoriSollecito()
            AggiornaTabellaFornitori()
        Catch ex As Exception
            MessageBox.Show("Errore: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub lvFornitori_DoubleClick(sender As Object, e As EventArgs) Handles lvFornitori.DoubleClick
        btnToggleSollecito_Click(sender, e)
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' DB — Fornitori da sollecitare
    ' ─────────────────────────────────────────────────────────

    Private Sub CreaTabellaLog()
        Try
            Using cn As New SqlConnection(Homepage.sap_tirelli)
                cn.Open()
                Dim sql =
                    "IF NOT EXISTS (SELECT * FROM [Tirelli_40].sys.tables WHERE name='SollecitiLog')
                     CREATE TABLE [Tirelli_40].dbo.SollecitiLog (
                         Id          int IDENTITY(1,1) PRIMARY KEY,
                         DataOra     datetime NOT NULL DEFAULT GETDATE(),
                         Utente      nvarchar(100) NOT NULL,
                         CodForn     nvarchar(30),
                         DescFor     nvarchar(250),
                         Email       nvarchar(250),
                         NRighe      int
                     )"
                Call New SqlCommand(sql, cn).ExecuteNonQuery()
            End Using
        Catch
        End Try
    End Sub

    Private Sub LogSollecito(codForn As String, descFor As String, email As String, nRighe As Integer)
        Try
            Using cn As New SqlConnection(Homepage.sap_tirelli)
                cn.Open()
                Dim sql =
                    "INSERT INTO [Tirelli_40].dbo.SollecitiLog (Utente, CodForn, DescFor, Email, NRighe)
                     VALUES (@Utente, @CodForn, @DescFor, @Email, @NRighe)"
                Using cmd As New SqlCommand(sql, cn)
                    cmd.Parameters.AddWithValue("@Utente", Environment.UserName)
                    cmd.Parameters.AddWithValue("@CodForn", If(codForn, ""))
                    cmd.Parameters.AddWithValue("@DescFor", If(descFor, ""))
                    cmd.Parameters.AddWithValue("@Email", If(email, ""))
                    cmd.Parameters.AddWithValue("@NRighe", nRighe)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
        Catch
        End Try
    End Sub

    Private Sub CreaTabellaFornitatoriSollecito()
        Try
            Using cn As New SqlConnection(Homepage.sap_tirelli)
                cn.Open()
                Dim sql = "IF NOT EXISTS (SELECT * FROM [Tirelli_40].sys.tables WHERE name='FornitatoriSollecito') " &
                          "CREATE TABLE [Tirelli_40].dbo.FornitatoriSollecito (" &
                          "  CodForn nvarchar(30) PRIMARY KEY, " &
                          "  DescFor nvarchar(250), " &
                          "  DataAggiunta datetime DEFAULT GETDATE())"
                Call New SqlCommand(sql, cn).ExecuteNonQuery()
            End Using
        Catch
        End Try
    End Sub

    Private Sub CaricaFornitatoriSollecito()
        _fornSollecito.Clear()
        Try
            Using cn As New SqlConnection(Homepage.sap_tirelli)
                cn.Open()
                Using cmd As New SqlCommand("SELECT CodForn FROM [Tirelli_40].dbo.FornitatoriSollecito", cn)
                    Using rd = cmd.ExecuteReader()
                        While rd.Read()
                            _fornSollecito.Add(rd.GetString(0).Trim())
                        End While
                    End Using
                End Using
            End Using
        Catch
        End Try
    End Sub

    Private Sub ToggleFornitoreSollecito(codForn As String, descFor As String)
        Using cn As New SqlConnection(Homepage.sap_tirelli)
            cn.Open()
            If _fornSollecito.Contains(codForn) Then
                Using cmd As New SqlCommand("DELETE FROM [Tirelli_40].dbo.FornitatoriSollecito WHERE CodForn=@C", cn)
                    cmd.Parameters.AddWithValue("@C", codForn)
                    cmd.ExecuteNonQuery()
                End Using
            Else
                Using cmd As New SqlCommand("INSERT INTO [Tirelli_40].dbo.FornitatoriSollecito (CodForn, DescFor) VALUES (@C, @D)", cn)
                    cmd.Parameters.AddWithValue("@C", codForn)
                    cmd.Parameters.AddWithValue("@D", descFor)
                    cmd.ExecuteNonQuery()
                End Using
            End If
        End Using
    End Sub

End Class
