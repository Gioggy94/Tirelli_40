Imports System.Data
Imports System.Data.SqlClient

Public Class Solleciti_OA

    Private _datiOA As DataTable
    Private _codFornSelezionato As String = ""
    Private _descFornSelezionato As String = ""
    Private _aggiornandoAnteprima As Boolean = False

    Private Sub Solleciti_OA_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ImpostaListaFornitori()
        ImpostaGriglia()
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
        lvFornitori.Columns.Add("Fornitore", 210)
        lvFornitori.Columns.Add("Righe", 55, HorizontalAlignment.Right)
        lvFornitori.Columns.Add("Scad.", 55, HorizontalAlignment.Right)
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
        aggiungiCol("colCodart", "Codice", 105)
        aggiungiCol("colDesCode", "Descrizione", 215)
        aggiungiCol("colDataRichiesta", "Data rich.", 95)
        aggiungiCol("colDataImmissione", "Data ord.", 85)
        aggiungiCol("colQtaOrd", "Q.ord.", 70)
        aggiungiCol("colQtaEnt", "Q.ric.", 70)
        aggiungiCol("colQtaRes", "Q.res.", 70)
        aggiungiCol("colIdComm", "Commessa", 90)
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
                "SELECT * FROM OPENQUERY(AS400, " &
                "'SELECT numdoc, codart, des_code, data_immissione, data_richiesta, " &
                " qta_ord, qta_ent, evaso, id_comm, cod_forn, desc_for, disegno " &
                " FROM TIR90VIS.JGALord t0 " &
                " WHERE DOC = ''OA'' " &
                " AND evaso <> ''S''')"
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
        Dim filtroForn = If(cmbFiltroFornitore.SelectedIndex > 0, CStr(cmbFiltroFornitore.SelectedItem), "")

        Dim gruppi = _datiOA.AsEnumerable().
            Where(Function(r)
                      If filtroComm <> "" AndAlso Not r("id_comm").ToString().Trim().ToUpper().Contains(filtroComm) Then Return False
                      If filtroForn <> "" AndAlso r("desc_for").ToString().Trim() <> filtroForn Then Return False
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
                           .RigheScadute = scadute
                       }
                   End Function).
            Where(Function(f) Not soloScaduti OrElse f.RigheScadute > 0).
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

        lvFornitori.Items.Clear()
        For Each f In gruppi
            Dim item As New ListViewItem(f.DescFor & " [" & f.CodForn & "]")
            item.SubItems.Add(f.TotRighe.ToString())
            item.SubItems.Add(f.RigheScadute.ToString())
            item.Tag = f.CodForn
            If f.RigheScadute > 0 Then item.ForeColor = Color.DarkRed
            lvFornitori.Items.Add(item)
        Next
        lblConteggioFornitori.Text = gruppi.Count & " fornitori"
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

        Dim righe = _datiOA.AsEnumerable().
            Where(Function(r)
                      If r("cod_forn").ToString().Trim() <> codForn Then Return False
                      If filtroComm <> "" AndAlso Not r("id_comm").ToString().Trim().ToUpper().Contains(filtroComm) Then Return False
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

            Dim idx = dgvOrdini.Rows.Add(
                True,
                r("numdoc").ToString().Trim(),
                r("codart").ToString().Trim(),
                r("des_code").ToString().Trim(),
                If(dataRich.HasValue, dataRich.Value.ToString("dd/MM/yyyy"), ""),
                If(dataImmiss.HasValue, dataImmiss.Value.ToString("dd/MM/yyyy"), ""),
                qtaOrd.ToString("N0"),
                qtaEnt.ToString("N0"),
                (qtaOrd - qtaEnt).ToString("N0"),
                r("id_comm").ToString().Trim(),
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
            Return
        End If

        Dim sb As New System.Text.StringBuilder
        sb.AppendLine("Spett.le " & descFor & ",")
        sb.AppendLine()
        sb.AppendLine("con la presente Vi sollecitiamo la consegna dei seguenti ordini di acquisto ancora in attesa di evasione:")
        sb.AppendLine()

        Dim hdr = String.Format("{0,-12} {1,-14} {2,-32} {3,8} {4,8} {5,12}  {6}",
                                 "N. Ordine", "Codice", "Descrizione", "Q.ord.", "Q.ric.", "Data rich.", "Commessa")
        sb.AppendLine(hdr)
        sb.AppendLine(New String("-"c, Math.Max(hdr.Length, 90)))

        For Each row In righeSelezionate
            Dim numdoc = If(TryCast(row.Cells("colNumdoc").Value, String), "")
            Dim codart = If(TryCast(row.Cells("colCodart").Value, String), "")
            Dim desc = If(TryCast(row.Cells("colDesCode").Value, String), "")
            If desc.Length > 30 Then desc = desc.Substring(0, 30)
            Dim qtaOrd = If(TryCast(row.Cells("colQtaOrd").Value, String), "")
            Dim qtaEnt = If(TryCast(row.Cells("colQtaEnt").Value, String), "")
            Dim dataRich = If(TryCast(row.Cells("colDataRichiesta").Value, String), "")
            Dim idComm = If(TryCast(row.Cells("colIdComm").Value, String), "")
            sb.AppendLine(String.Format("{0,-12} {1,-14} {2,-32} {3,8} {4,8} {5,12}  {6}",
                                         numdoc, codart, desc, qtaOrd, qtaEnt, dataRich, idComm))
        Next

        sb.AppendLine()
        sb.AppendLine("Vi chiediamo cortesemente di confermarci la data di consegna prevista per ciascun articolo.")
        sb.AppendLine()
        sb.AppendLine("In attesa di un vostro riscontro, porgiamo distinti saluti.")
        sb.AppendLine()
        sb.AppendLine("Tirelli S.r.l. - Ufficio Acquisti")

        rtbAnteprima.Text = sb.ToString()

        If txtOggetto.Text = "" Then
            txtOggetto.Text = "Sollecito ordini di acquisto - Tirelli S.r.l."
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────
    ' Prepara mail Outlook per il fornitore selezionato
    ' ─────────────────────────────────────────────────────────

    Sub PreparaMail()
        If rtbAnteprima.Text.Trim() = "" OrElse rtbAnteprima.Text.Trim() = "(nessuna riga selezionata)" Then
            MessageBox.Show("Nessuna riga selezionata. Selezionare almeno una riga.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        Try
            Dim objOutlook As Object = CreateObject("Outlook.Application")
            Dim objMail As Object = objOutlook.CreateItem(0)
            With objMail
                .To = txtEmail.Text.Trim()
                .Subject = txtOggetto.Text
                .Body = rtbAnteprima.Text
                .Display()
            End With
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
            _descFornSelezionato = item.Text
            AggiornaTabellaOrdini(item.Tag.ToString())
            For Each row As DataGridViewRow In dgvOrdini.Rows
                row.Cells("colSel").Value = True
            Next
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

    Private Sub lvFornitori_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lvFornitori.SelectedIndexChanged
        If lvFornitori.SelectedItems.Count = 0 Then Return
        Dim item = lvFornitori.SelectedItems(0)
        _descFornSelezionato = item.Text
        AggiornaTabellaOrdini(item.Tag.ToString())
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
        PreparaMail()
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

    Private Sub txtFiltroCommessa_TextChanged(sender As Object, e As EventArgs) Handles txtFiltroCommessa.TextChanged
        If _datiOA IsNot Nothing Then AggiornaTabellaFornitori()
    End Sub

    Private Sub cmbFiltroFornitore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFiltroFornitore.SelectedIndexChanged
        If _datiOA IsNot Nothing Then AggiornaTabellaFornitori()
    End Sub

End Class
