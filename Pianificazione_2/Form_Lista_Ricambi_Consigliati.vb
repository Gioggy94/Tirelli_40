Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports PdfSharp.Pdf
Imports PD = PdfSharp.Drawing

Public Class Form_Lista_Ricambi_Consigliati

    Public commessa As String = ""
    Public n_rev As Integer = 0

    ' Dizionario codart -> qtapia totale impiegata nella commessa (da JGALIMP)
    Private _impieghi As New Dictionary(Of String, Decimal)(StringComparer.OrdinalIgnoreCase)

    ' Nome lista attualmente visualizzata
    Private _nomeLista As String = "Lista 1"

    ' Colonne griglia
    Private Const COL_IMG As String = "Immagine"
    Private Const COL_CODICE As String = "Codice"
    Private Const COL_DESC As String = "Descrizione"
    Private Const COL_DESC_SUP As String = "Desc_Sup"
    Private Const COL_QTA As String = "Quantita"
    Private Const COL_COSTO As String = "Costo"
    Private Const COL_COSTO_TOT As String = "CostoTot"
    Private Const COL_PZ_IMP As String = "PzImpiegati"

    Private _isLoading As Boolean = False

    ' ─────────────────────────────────────────────────────────────────
    '  LOAD
    ' ─────────────────────────────────────────────────────────────────
    Private Sub Form_Lista_Ricambi_Consigliati_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ApplicaStile()
        ImpostaGriglia()

        lblCommessa.Text = "Commessa: " & commessa & "   Rev.: " & n_rev

        cmbMoltiplicatore.Items.AddRange(New Object() {"1,0", "1,05", "1,1", "1,15", "1,2", "1,25", "1,3", "1,5", "2,0"})
        cmbMoltiplicatore.SelectedIndex = 0

        CaricaImpieghi()
        CaricaListeDisponibili()
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  STILE
    ' ─────────────────────────────────────────────────────────────────
    Private Sub ApplicaStile()
        Dim navy As System.Drawing.Color = System.Drawing.Color.FromArgb(22, 45, 84)
        Dim navyHover As System.Drawing.Color = System.Drawing.Color.FromArgb(30, 63, 122)

        btnSalva.BackColor = navy
        btnSalva.ForeColor = System.Drawing.Color.White
        btnSalva.FlatStyle = FlatStyle.Flat
        btnSalva.FlatAppearance.BorderSize = 0
        btnSalva.FlatAppearance.MouseOverBackColor = navyHover

        btnChiudi.FlatStyle = FlatStyle.Flat
        btnAggiungi.FlatStyle = FlatStyle.Flat
        btnElimina.FlatStyle = FlatStyle.Flat

        btnNuovaLista.BackColor = System.Drawing.Color.FromArgb(60, 90, 150)
        btnNuovaLista.ForeColor = System.Drawing.Color.White
        btnNuovaLista.FlatStyle = FlatStyle.Flat
        btnNuovaLista.FlatAppearance.BorderSize = 0

        btnExportExcel.BackColor = System.Drawing.Color.FromArgb(32, 120, 60)
        btnExportExcel.ForeColor = System.Drawing.Color.White
        btnExportExcel.FlatStyle = FlatStyle.Flat
        btnExportExcel.FlatAppearance.BorderSize = 0

        btnExportPdf.BackColor = System.Drawing.Color.FromArgb(170, 50, 50)
        btnExportPdf.ForeColor = System.Drawing.Color.White
        btnExportPdf.FlatStyle = FlatStyle.Flat
        btnExportPdf.FlatAppearance.BorderSize = 0

        lblCommessa.ForeColor = System.Drawing.Color.White
        lblCommessa.Font = New System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
        lblTotale.ForeColor = navy
        lblTotale.Font = New System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  SETUP GRIGLIA
    ' ─────────────────────────────────────────────────────────────────
    Private Sub ImpostaGriglia()
        dgvRicambi.AutoGenerateColumns = False
        dgvRicambi.AllowUserToAddRows = False
        dgvRicambi.RowHeadersVisible = False
        dgvRicambi.BackgroundColor = System.Drawing.Color.White
        dgvRicambi.BorderStyle = BorderStyle.None
        dgvRicambi.GridColor = System.Drawing.Color.FromArgb(210, 220, 235)
        dgvRicambi.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        Dim navy As System.Drawing.Color = System.Drawing.Color.FromArgb(22, 45, 84)
        dgvRicambi.ColumnHeadersDefaultCellStyle.BackColor = navy
        dgvRicambi.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White
        dgvRicambi.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Segoe UI", 8.5F, System.Drawing.FontStyle.Bold)
        dgvRicambi.EnableHeadersVisualStyles = False
        dgvRicambi.RowsDefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(210, 225, 245)
        dgvRicambi.RowsDefaultCellStyle.SelectionForeColor = navy
        dgvRicambi.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(245, 248, 255)

        dgvRicambi.Columns.Clear()
        dgvRicambi.RowTemplate.Height = 60

        ' Immagine articolo
        Dim colImg As New DataGridViewImageColumn()
        colImg.Name = COL_IMG
        colImg.HeaderText = ""
        colImg.Width = 65
        colImg.ReadOnly = True
        colImg.ImageLayout = DataGridViewImageCellLayout.Zoom
        colImg.DefaultCellStyle.NullValue = Nothing
        dgvRicambi.Columns.Add(colImg)

        Dim colCodice As New DataGridViewTextBoxColumn()
        colCodice.Name = COL_CODICE
        colCodice.HeaderText = "Codice"
        colCodice.Width = 110
        dgvRicambi.Columns.Add(colCodice)

        Dim colDesc As New DataGridViewTextBoxColumn()
        colDesc.Name = COL_DESC
        colDesc.HeaderText = "Descrizione"
        colDesc.Width = 220
        colDesc.ReadOnly = True
        colDesc.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 245, 255)
        dgvRicambi.Columns.Add(colDesc)

        Dim colDescSup As New DataGridViewTextBoxColumn()
        colDescSup.Name = COL_DESC_SUP
        colDescSup.HeaderText = "Desc. Sup."
        colDescSup.Width = 180
        colDescSup.ReadOnly = True
        colDescSup.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(240, 245, 255)
        dgvRicambi.Columns.Add(colDescSup)

        Dim colQta As New DataGridViewTextBoxColumn()
        colQta.Name = COL_QTA
        colQta.HeaderText = "Qtà"
        colQta.Width = 65
        colQta.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        dgvRicambi.Columns.Add(colQta)

        Dim colCosto As New DataGridViewTextBoxColumn()
        colCosto.Name = COL_COSTO
        colCosto.HeaderText = "Costo unit."
        colCosto.Width = 90
        colCosto.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        colCosto.DefaultCellStyle.Format = "N2"
        dgvRicambi.Columns.Add(colCosto)

        Dim colCostoTot As New DataGridViewTextBoxColumn()
        colCostoTot.Name = COL_COSTO_TOT
        colCostoTot.HeaderText = "Costo tot. (x molt.)"
        colCostoTot.Width = 130
        colCostoTot.ReadOnly = True
        colCostoTot.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        colCostoTot.DefaultCellStyle.Format = "N2"
        colCostoTot.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(235, 255, 235)
        colCostoTot.DefaultCellStyle.Font = New System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)
        dgvRicambi.Columns.Add(colCostoTot)

        Dim colPzImp As New DataGridViewTextBoxColumn()
        colPzImp.Name = COL_PZ_IMP
        colPzImp.HeaderText = "Pz impiegati in macchina"
        colPzImp.Width = 140
        colPzImp.ReadOnly = True
        colPzImp.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        colPzImp.DefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(255, 248, 230)
        dgvRicambi.Columns.Add(colPzImp)
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  CARICA IMPIEGHI DA JGALIMP (per commessa)
    ' ─────────────────────────────────────────────────────────────────
    Private Sub CaricaImpieghi()
        _impieghi.Clear()
        If String.IsNullOrWhiteSpace(commessa) Then Return

        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnn
                    cmd.CommandTimeout = 60
                    cmd.CommandText = "
SELECT trim(t10.codart) AS codart, SUM(t10.qtapia) AS qtapia
FROM OPENQUERY([AS400], '
    SELECT codart, qtapia, matricola
    FROM S786FAD1.TIR90VIS.JGALIMP
    WHERE evaso_odp <> ''S''
      AND trim(matricola) = ''" & commessa.Replace("'", "''") & "''
') AS t10
GROUP BY trim(t10.codart)"

                    Using rd As SqlDataReader = cmd.ExecuteReader()
                        While rd.Read()
                            Dim cod As String = rd("codart").ToString().Trim()
                            Dim qta As Decimal = 0
                            Decimal.TryParse(rd("qtapia").ToString(), qta)
                            If Not String.IsNullOrEmpty(cod) Then
                                If _impieghi.ContainsKey(cod) Then
                                    _impieghi(cod) += qta
                                Else
                                    _impieghi(cod) = qta
                                End If
                            End If
                        End While
                    End Using
                End Using
            End Using
        Catch
            ' Impieghi non critici
        End Try
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  CARICA LISTE DISPONIBILI PER (COMMESSA, REV)
    ' ─────────────────────────────────────────────────────────────────
    Private Sub CaricaListeDisponibili()
        _isLoading = True
        Dim nomePrecedente As String = _nomeLista
        cmbNomeLista.Items.Clear()

        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnn
                    cmd.CommandText = "
SELECT DISTINCT NomeLista
FROM [Tirelli_40].[dbo].[Lista_Ricambi_Consigliati_Righe]
WHERE Commessa=@c AND Rev=@r
ORDER BY NomeLista"
                    cmd.Parameters.AddWithValue("@c", commessa)
                    cmd.Parameters.AddWithValue("@r", n_rev)
                    Using rd As SqlDataReader = cmd.ExecuteReader()
                        While rd.Read()
                            cmbNomeLista.Items.Add(rd(0).ToString())
                        End While
                    End Using
                End Using
            End Using
        Catch
        End Try

        If cmbNomeLista.Items.Count = 0 Then
            cmbNomeLista.Items.Add("Lista 1")
        End If

        Dim idx As Integer = cmbNomeLista.Items.IndexOf(nomePrecedente)
        _isLoading = False
        cmbNomeLista.SelectedIndex = If(idx >= 0, idx, 0)
        ' SelectedIndexChanged scatena CaricaRighe
    End Sub

    Private Sub cmbNomeLista_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbNomeLista.SelectedIndexChanged
        If _isLoading Then Return
        If cmbNomeLista.SelectedItem IsNot Nothing Then
            _nomeLista = cmbNomeLista.SelectedItem.ToString()
            CaricaRighe()
        End If
    End Sub

    Private Sub btnNuovaLista_Click(sender As Object, e As EventArgs) Handles btnNuovaLista.Click
        Dim defaultNome As String = "Lista " & (cmbNomeLista.Items.Count + 1)
        Dim nome As String = InputBox("Inserisci il nome della nuova lista:", "Nuova lista ricambi", defaultNome)
        If String.IsNullOrWhiteSpace(nome) Then Return
        nome = nome.Trim()

        If cmbNomeLista.Items.Contains(nome) Then
            cmbNomeLista.SelectedItem = nome
            Return
        End If

        _isLoading = True
        cmbNomeLista.Items.Add(nome)
        _isLoading = False
        _nomeLista = nome
        cmbNomeLista.SelectedItem = nome
        dgvRicambi.Rows.Clear()
        AggiornaTotale()
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  CARICA RIGHE SALVATE DAL DB SQL
    ' ─────────────────────────────────────────────────────────────────
    Private Sub CaricaRighe()
        _isLoading = True
        dgvRicambi.Rows.Clear()

        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnn
                    cmd.CommandText = "
SELECT Moltiplicatore, Codice, Descrizione, DescrizioneSup, Quantita, Costo, CostoTot
FROM [Tirelli_40].[dbo].[Lista_Ricambi_Consigliati_Righe]
WHERE Commessa=@comm AND Rev=@rev AND NomeLista=@nome
ORDER BY ID"
                    cmd.Parameters.AddWithValue("@comm", commessa)
                    cmd.Parameters.AddWithValue("@rev", n_rev)
                    cmd.Parameters.AddWithValue("@nome", _nomeLista)

                    Using rd As SqlDataReader = cmd.ExecuteReader()
                        Dim primaRiga As Boolean = True
                        While rd.Read()
                            If primaRiga Then
                                Dim molStr As String = rd("Moltiplicatore").ToString().Replace(".", ",")
                                If cmbMoltiplicatore.Items.Contains(molStr) Then
                                    cmbMoltiplicatore.Text = molStr
                                Else
                                    cmbMoltiplicatore.Text = molStr
                                End If
                                primaRiga = False
                            End If

                            Dim cod As String = rd("Codice").ToString().Trim()
                            Dim pzImp As String = ""
                            If _impieghi.ContainsKey(cod) Then
                                pzImp = _impieghi(cod).ToString("0.####")
                            End If

                            dgvRicambi.Rows.Add(
                                CaricaImmagine(cod),
                                cod,
                                rd("Descrizione").ToString(),
                                rd("DescrizioneSup").ToString(),
                                rd("Quantita").ToString(),
                                rd("Costo").ToString(),
                                rd("CostoTot").ToString(),
                                pzImp)
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Errore caricamento lista: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try

        _isLoading = False
        AggiornaTotale()
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  LOOKUP ARTICOLO SU AS400 (JGALART)
    ' ─────────────────────────────────────────────────────────────────
    Private Function LookupArticolo(codice As String) As (Descrizione As String, DescSup As String, Costo As Decimal, Trovato As Boolean)
        Dim result = (Descrizione:="", DescSup:="", Costo:=0D, Trovato:=False)
        If String.IsNullOrWhiteSpace(codice) Then Return result

        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnn
                    cmd.CommandTimeout = 30
                    cmd.CommandText = String.Format("
SELECT trim(CODE) AS CODE, DES_CODE AS Descrizione, LNG_CODE AS DescrizioneSup,
       COSTO_STD AS Costo
FROM OPENQUERY([AS400], 'SELECT CODE, DES_CODE, LNG_CODE, COSTO_STD
                          FROM S786FAD1.TIR90VIS.JGALART
                          WHERE trim(CODE) = ''{0}''') T10",
                        codice.Trim().Replace("'", "''"))

                    Using rd As SqlDataReader = cmd.ExecuteReader()
                        If rd.Read() Then
                            result.Descrizione = rd("Descrizione").ToString()
                            result.DescSup = rd("DescrizioneSup").ToString()
                            Decimal.TryParse(rd("Costo").ToString().Replace(",", "."),
                                             System.Globalization.NumberStyles.Any,
                                             System.Globalization.CultureInfo.InvariantCulture,
                                             result.Costo)
                            result.Trovato = True
                        End If
                    End Using
                End Using
            End Using
        Catch
        End Try

        Return result
    End Function

    ' ─────────────────────────────────────────────────────────────────
    '  CARICA IMMAGINE ARTICOLO
    ' ─────────────────────────────────────────────────────────────────
    Private Function CaricaImmagine(codice As String) As Image
        If String.IsNullOrWhiteSpace(codice) Then Return Nothing
        Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & codice & ".PNG"
        If Not File.Exists(percorso) Then Return Nothing
        Try
            Using fs As New FileStream(percorso, FileMode.Open, FileAccess.Read)
                Using tmp As Image = Image.FromStream(fs)
                    Return New Bitmap(tmp)
                End Using
            End Using
        Catch
            Return Nothing
        End Try
    End Function

    ' ─────────────────────────────────────────────────────────────────
    '  CELLENDEDITEDIT — quando si lascia la cella codice
    ' ─────────────────────────────────────────────────────────────────
    Private Sub dgvRicambi_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvRicambi.CellEndEdit
        If _isLoading Then Return
        Dim row As DataGridViewRow = dgvRicambi.Rows(e.RowIndex)

        If e.ColumnIndex = dgvRicambi.Columns(COL_CODICE).Index Then
            Dim codice As String = If(row.Cells(COL_CODICE).Value IsNot Nothing, row.Cells(COL_CODICE).Value.ToString().Trim().ToUpper(), "")
            If Not String.IsNullOrEmpty(codice) Then
                row.Cells(COL_CODICE).Value = codice
                row.Cells(COL_IMG).Value = CaricaImmagine(codice)
                Dim info = LookupArticolo(codice)
                If info.Trovato Then
                    row.Cells(COL_DESC).Value = info.Descrizione
                    row.Cells(COL_DESC_SUP).Value = info.DescSup
                    If row.Cells(COL_COSTO).Value Is Nothing OrElse row.Cells(COL_COSTO).Value.ToString() = "" Then
                        row.Cells(COL_COSTO).Value = info.Costo
                    End If
                    If _impieghi.ContainsKey(codice) Then
                        row.Cells(COL_PZ_IMP).Value = _impieghi(codice).ToString("0.####")
                    Else
                        row.Cells(COL_PZ_IMP).Value = "—"
                    End If
                Else
                    row.Cells(COL_PZ_IMP).Value = "Codice non trovato"
                End If
            End If
            RicalcolaCostoTot(row)
        End If

        If e.ColumnIndex = dgvRicambi.Columns(COL_QTA).Index OrElse
           e.ColumnIndex = dgvRicambi.Columns(COL_COSTO).Index Then
            RicalcolaCostoTot(row)
        End If

        AggiornaTotale()
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  RICALCOLA COSTO TOT PER RIGA
    ' ─────────────────────────────────────────────────────────────────
    Private Sub RicalcolaCostoTot(row As DataGridViewRow)
        Dim qta As Decimal = 0
        Dim costo As Decimal = 0
        Dim molt As Decimal = 1

        Decimal.TryParse(If(row.Cells(COL_QTA).Value IsNot Nothing, row.Cells(COL_QTA).Value.ToString().Replace(",", "."), ""),
                         System.Globalization.NumberStyles.Any,
                         System.Globalization.CultureInfo.InvariantCulture, qta)
        Decimal.TryParse(If(row.Cells(COL_COSTO).Value IsNot Nothing, row.Cells(COL_COSTO).Value.ToString().Replace(",", "."), ""),
                         System.Globalization.NumberStyles.Any,
                         System.Globalization.CultureInfo.InvariantCulture, costo)
        Decimal.TryParse(cmbMoltiplicatore.Text.Replace(",", "."),
                         System.Globalization.NumberStyles.Any,
                         System.Globalization.CultureInfo.InvariantCulture, molt)

        row.Cells(COL_COSTO_TOT).Value = System.Math.Round(qta * costo * molt, 2)
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  TOTALE LISTINO
    ' ─────────────────────────────────────────────────────────────────
    Private Sub AggiornaTotale()
        Dim tot As Decimal = 0
        For Each row As DataGridViewRow In dgvRicambi.Rows
            Dim v As Decimal = 0
            If row.Cells(COL_COSTO_TOT).Value IsNot Nothing Then
                Decimal.TryParse(row.Cells(COL_COSTO_TOT).Value.ToString().Replace(",", "."),
                                 System.Globalization.NumberStyles.Any,
                                 System.Globalization.CultureInfo.InvariantCulture, v)
                tot += v
            End If
        Next
        lblTotale.Text = "TOTALE: € " & tot.ToString("N2")
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  MOLTIPLICATORE CAMBIATO → ricalcola tutto
    ' ─────────────────────────────────────────────────────────────────
    Private Sub cmbMoltiplicatore_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMoltiplicatore.SelectedIndexChanged
        If _isLoading Then Return
        RicalcolaTutteLeRighe()
    End Sub

    Private Sub cmbMoltiplicatore_TextChanged(sender As Object, e As EventArgs) Handles cmbMoltiplicatore.TextChanged
        If _isLoading Then Return
        RicalcolaTutteLeRighe()
    End Sub

    Private Sub RicalcolaTutteLeRighe()
        For Each row As DataGridViewRow In dgvRicambi.Rows
            RicalcolaCostoTot(row)
        Next
        AggiornaTotale()
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  AGGIUNGI RIGA
    ' ─────────────────────────────────────────────────────────────────
    Private Sub btnAggiungi_Click(sender As Object, e As EventArgs) Handles btnAggiungi.Click
        dgvRicambi.Rows.Add(Nothing, "", "", "", "1", "", "", "")
        Dim lastIdx As Integer = dgvRicambi.Rows.Count - 1
        dgvRicambi.CurrentCell = dgvRicambi.Rows(lastIdx).Cells(COL_CODICE)
        dgvRicambi.BeginEdit(True)
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  ELIMINA RIGA SELEZIONATA
    ' ─────────────────────────────────────────────────────────────────
    Private Sub btnElimina_Click(sender As Object, e As EventArgs) Handles btnElimina.Click
        If dgvRicambi.CurrentRow IsNot Nothing Then
            dgvRicambi.Rows.RemoveAt(dgvRicambi.CurrentRow.Index)
            AggiornaTotale()
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  SALVA
    ' ─────────────────────────────────────────────────────────────────
    Private Sub btnSalva_Click(sender As Object, e As EventArgs) Handles btnSalva.Click
        dgvRicambi.EndEdit()

        Dim molt As Decimal = 1
        Decimal.TryParse(cmbMoltiplicatore.Text.Replace(",", "."),
                         System.Globalization.NumberStyles.Any,
                         System.Globalization.CultureInfo.InvariantCulture, molt)

        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Using tran As SqlTransaction = cnn.BeginTransaction()
                    Try
                        Using cmdDel As New SqlCommand()
                            cmdDel.Connection = cnn
                            cmdDel.Transaction = tran
                            cmdDel.CommandText = "DELETE FROM [Tirelli_40].[dbo].[Lista_Ricambi_Consigliati_Righe] WHERE Commessa=@c AND Rev=@r AND NomeLista=@nome"
                            cmdDel.Parameters.AddWithValue("@c", commessa)
                            cmdDel.Parameters.AddWithValue("@r", n_rev)
                            cmdDel.Parameters.AddWithValue("@nome", _nomeLista)
                            cmdDel.ExecuteNonQuery()
                        End Using

                        For Each row As DataGridViewRow In dgvRicambi.Rows
                            Dim cod As String = If(row.Cells(COL_CODICE).Value IsNot Nothing, row.Cells(COL_CODICE).Value.ToString().Trim(), "")
                            If String.IsNullOrEmpty(cod) Then Continue For

                            Dim qta As Decimal = 0
                            Dim costo As Decimal = 0
                            Dim costoTot As Decimal = 0
                            Decimal.TryParse(If(row.Cells(COL_QTA).Value IsNot Nothing, row.Cells(COL_QTA).Value.ToString().Replace(",", "."), ""),
                                             System.Globalization.NumberStyles.Any,
                                             System.Globalization.CultureInfo.InvariantCulture, qta)
                            Decimal.TryParse(If(row.Cells(COL_COSTO).Value IsNot Nothing, row.Cells(COL_COSTO).Value.ToString().Replace(",", "."), ""),
                                             System.Globalization.NumberStyles.Any,
                                             System.Globalization.CultureInfo.InvariantCulture, costo)
                            Decimal.TryParse(If(row.Cells(COL_COSTO_TOT).Value IsNot Nothing, row.Cells(COL_COSTO_TOT).Value.ToString().Replace(",", "."), ""),
                                             System.Globalization.NumberStyles.Any,
                                             System.Globalization.CultureInfo.InvariantCulture, costoTot)

                            Using cmdIns As New SqlCommand()
                                cmdIns.Connection = cnn
                                cmdIns.Transaction = tran
                                cmdIns.CommandText = "
INSERT INTO [Tirelli_40].[dbo].[Lista_Ricambi_Consigliati_Righe]
(Commessa, Rev, NomeLista, Moltiplicatore, Codice, Descrizione, DescrizioneSup, Quantita, Costo, CostoTot)
VALUES (@comm, @rev, @nome, @molt, @cod, @des, @dessup, @qta, @costo, @costoTot)"
                                cmdIns.Parameters.AddWithValue("@comm", commessa)
                                cmdIns.Parameters.AddWithValue("@rev", n_rev)
                                cmdIns.Parameters.AddWithValue("@nome", _nomeLista)
                                cmdIns.Parameters.AddWithValue("@molt", molt)
                                cmdIns.Parameters.AddWithValue("@cod", cod)
                                cmdIns.Parameters.AddWithValue("@des", If(row.Cells(COL_DESC).Value IsNot Nothing, row.Cells(COL_DESC).Value.ToString(), ""))
                                cmdIns.Parameters.AddWithValue("@dessup", If(row.Cells(COL_DESC_SUP).Value IsNot Nothing, row.Cells(COL_DESC_SUP).Value.ToString(), ""))
                                cmdIns.Parameters.AddWithValue("@qta", qta)
                                cmdIns.Parameters.AddWithValue("@costo", costo)
                                cmdIns.Parameters.AddWithValue("@costoTot", costoTot)
                                cmdIns.ExecuteNonQuery()
                            End Using
                        Next

                        tran.Commit()
                        MsgBox("Lista ricambi salvata correttamente.", MsgBoxStyle.Information)
                    Catch ex As Exception
                        tran.Rollback()
                        MsgBox("Errore salvataggio: " & ex.Message, MsgBoxStyle.Critical)
                    End Try
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Errore connessione: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  EXPORT EXCEL
    ' ─────────────────────────────────────────────────────────────────
    Private Sub btnExportExcel_Click(sender As Object, e As EventArgs) Handles btnExportExcel.Click
        dgvRicambi.EndEdit()
        If dgvRicambi.Rows.Count = 0 Then
            MsgBox("Nessuna riga da esportare.", MsgBoxStyle.Information)
            Return
        End If

        Dim sfd As New SaveFileDialog()
        sfd.Filter = "Excel (*.xlsx)|*.xlsx"
        sfd.FileName = "Ricambi_" & commessa & "_" & _nomeLista.Replace(" ", "_") & ".xlsx"
        If sfd.ShowDialog() <> DialogResult.OK Then Return

        Try
            Using doc As SpreadsheetDocument = SpreadsheetDocument.Create(sfd.FileName, SpreadsheetDocumentType.Workbook)
                Dim wbPart As WorkbookPart = doc.AddWorkbookPart()
                wbPart.Workbook = New Workbook()

                Dim wsPart As WorksheetPart = wbPart.AddNewPart(Of WorksheetPart)()
                Dim sheetData As New SheetData()
                wsPart.Worksheet = New Worksheet(sheetData)

                Dim sheets As Sheets = doc.WorkbookPart.Workbook.AppendChild(New Sheets())
                sheets.Append(New Sheet() With {
                    .Id = doc.WorkbookPart.GetIdOfPart(wsPart),
                    .SheetId = 1,
                    .Name = "Lista Ricambi"
                })

                ' Intestazione documento
                sheetData.AppendChild(ExcelRiga({"Commessa: " & commessa & "  Rev.: " & n_rev & "  Lista: " & _nomeLista}))
                sheetData.AppendChild(ExcelRiga({"Data: " & Now.ToString("dd/MM/yyyy")}))
                sheetData.AppendChild(New Row())

                ' Intestazione colonne
                sheetData.AppendChild(ExcelRiga({"Codice", "Descrizione", "Desc. Supplementare", "Qtà", "Costo unit. (€)", "Costo tot. (€)", "Pz in macchina"}))

                ' Righe dati
                Dim totale As Decimal = 0
                For Each row As DataGridViewRow In dgvRicambi.Rows
                    Dim cod As String = ValCella(row, COL_CODICE)
                    If String.IsNullOrEmpty(cod) Then Continue For
                    Dim qta As String = FormatDecimalCella(ValCella(row, COL_QTA), "0.##")
                    Dim costo As String = FormatDecimalCella(ValCella(row, COL_COSTO), "N2")
                    Dim costoTot As String = FormatDecimalCella(ValCella(row, COL_COSTO_TOT), "N2")
                    Dim ctDec As Decimal = 0
                    Decimal.TryParse(ValCella(row, COL_COSTO_TOT).Replace(",", "."), System.Globalization.NumberStyles.Any,
                                     System.Globalization.CultureInfo.InvariantCulture, ctDec)
                    totale += ctDec
                    sheetData.AppendChild(ExcelRiga({
                        cod,
                        ValCella(row, COL_DESC),
                        ValCella(row, COL_DESC_SUP),
                        qta,
                        costo,
                        costoTot,
                        ValCella(row, COL_PZ_IMP)
                    }))
                Next

                ' Riga totale
                sheetData.AppendChild(New Row())
                sheetData.AppendChild(ExcelRiga({"", "", "", "", "TOTALE", totale.ToString("N2"), ""}))

                ' Riga moltiplicatore
                sheetData.AppendChild(ExcelRiga({"", "", "", "", "Moltiplicatore: " & cmbMoltiplicatore.Text, "", ""}))

                wbPart.Workbook.Save()
            End Using

            If MsgBox("File salvato. Aprirlo ora?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                Process.Start(New ProcessStartInfo(sfd.FileName) With {.UseShellExecute = True})
            End If
        Catch ex As Exception
            MsgBox("Errore esportazione Excel: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Function ExcelRiga(valori() As String) As Row
        Dim r As New Row()
        For Each v As String In valori
            Dim cell As New Cell()
            cell.DataType = CellValues.InlineString
            cell.InlineString = New InlineString(New Text(SanitizzaXml(v)))
            r.AppendChild(cell)
        Next
        Return r
    End Function

    ''' <summary>Tronca il testo per farlo stare nella larghezza indicata, aggiungendo "…".</summary>
    Private Shared Function TroncaTesto(gfx As PD.XGraphics, testo As String, font As PD.XFont, larghezzaMax As Double) As String
        If String.IsNullOrEmpty(testo) Then Return testo
        If gfx.MeasureString(testo, font).Width <= larghezzaMax Then Return testo
        Dim s As String = testo
        Do While s.Length > 0 AndAlso gfx.MeasureString(s & "…", font).Width > larghezzaMax
            s = s.Substring(0, s.Length - 1)
        Loop
        Return If(s.Length > 0, s & "…", "")
    End Function

    ''' <summary>Rimuove caratteri non validi in XML (es. 0x1A da AS400).</summary>
    Private Shared Function SanitizzaXml(s As String) As String
        If String.IsNullOrEmpty(s) Then Return s
        Dim sb As New System.Text.StringBuilder(s.Length)
        For Each c As Char In s
            Dim n As Integer = AscW(c)
            If n = 9 OrElse n = 10 OrElse n = 13 OrElse (n >= 32 AndAlso n <= 55295) OrElse (n >= 57344 AndAlso n <= 65533) Then
                sb.Append(c)
            End If
        Next
        Return sb.ToString()
    End Function

    Private Function ValCella(row As DataGridViewRow, colName As String) As String
        Dim v = row.Cells(colName).Value
        Return If(v IsNot Nothing, v.ToString(), "")
    End Function

    ''' <summary>Parsa un valore decimale dal grid e lo formatta con il formato indicato.</summary>
    Private Shared Function FormatDecimalCella(raw As String, formato As String) As String
        If String.IsNullOrWhiteSpace(raw) Then Return ""
        Dim d As Decimal = 0
        Decimal.TryParse(raw.Replace(",", "."), System.Globalization.NumberStyles.Any,
                         System.Globalization.CultureInfo.InvariantCulture, d)
        Return d.ToString(formato)
    End Function

    ' ─────────────────────────────────────────────────────────────────
    '  EXPORT PDF OFFERTA
    ' ─────────────────────────────────────────────────────────────────
    Private Sub btnExportPdf_Click(sender As Object, e As EventArgs) Handles btnExportPdf.Click
        dgvRicambi.EndEdit()
        If dgvRicambi.Rows.Count = 0 Then
            MsgBox("Nessuna riga da esportare.", MsgBoxStyle.Information)
            Return
        End If

        Dim sfd As New SaveFileDialog()
        sfd.Filter = "PDF (*.pdf)|*.pdf"
        sfd.FileName = "Offerta_Ricambi_" & commessa & "_" & _nomeLista.Replace(" ", "_") & ".pdf"
        If sfd.ShowDialog() <> DialogResult.OK Then Return

        Try
            Dim doc As New PdfDocument()
            doc.Info.Title = "Lista Ricambi Consigliati - " & commessa
            doc.Info.Author = "Tirelli S.r.l."

            Dim navy As PD.XColor = PD.XColor.FromArgb(255, 22, 45, 84)
            Dim green As PD.XColor = PD.XColor.FromArgb(255, 220, 245, 220)
            Dim white As PD.XColor = PD.XColor.FromArgb(255, 255, 255, 255)
            Dim altRow As PD.XColor = PD.XColor.FromArgb(255, 245, 248, 255)
            Dim lineColor As PD.XColor = PD.XColor.FromArgb(255, 210, 220, 235)

            Dim fontTitle As New PD.XFont("Arial", 16, PD.XFontStyleEx.Bold)
            Dim fontSub As New PD.XFont("Arial", 10, PD.XFontStyleEx.Bold)
            Dim fontNorm As New PD.XFont("Arial", 8.5)
            Dim fontBold As New PD.XFont("Arial", 8.5, PD.XFontStyleEx.Bold)

            Dim pageW As Double = PD.XUnit.FromMillimeter(297).Point
            Dim pageH As Double = PD.XUnit.FromMillimeter(210).Point
            Dim marginL As Double = 36
            Dim contentW As Double = pageW - marginL - 36

            Dim colW() As Double = {90, 220, 190, 45, 80, 85, 60}

            Dim rowH As Double = 18
            Dim headerH As Double = 22

            Dim page As PdfPage = doc.AddPage()
            page.Width = pageW
            page.Height = pageH
            Dim gfx As PD.XGraphics = PD.XGraphics.FromPdfPage(page)

            Dim y As Double = 30

            ' ── Banner intestazione ──────────────────────────────
            gfx.DrawRectangle(New PD.XSolidBrush(navy), marginL, y, contentW, 44)
            gfx.DrawString("LISTA RICAMBI CONSIGLIATI", fontTitle, PD.XBrushes.White,
                           New PD.XRect(marginL + 10, y + 10, contentW - 160, 30), PD.XStringFormats.CenterLeft)
            ' Logo Tirelli (se presente)
            Dim percorsoLogo As String = System.IO.Path.Combine(Application.StartupPath, "Tirelli.png")
            If System.IO.File.Exists(percorsoLogo) Then
                Dim logo As PD.XImage = PD.XImage.FromFile(percorsoLogo)
                Dim logoH As Double = 36
                Dim logoW As Double = logo.PixelWidth * logoH / logo.PixelHeight
                gfx.DrawImage(logo, marginL + contentW - logoW - 8, y + 4, logoW, logoH)
                logo.Dispose()
            Else
                gfx.DrawString("Tirelli S.r.l.", New PD.XFont("Arial", 10, PD.XFontStyleEx.BoldItalic), PD.XBrushes.White,
                               New PD.XRect(marginL, y + 10, contentW - 10, 30), PD.XStringFormats.CenterRight)
            End If
            y += 52

            ' ── Info commessa ────────────────────────────────────
            gfx.DrawString("Commessa: " & commessa & "   Rev.: " & n_rev,
                           fontSub, New PD.XSolidBrush(navy), marginL, y)
            gfx.DrawString("Lista: " & _nomeLista,
                           fontSub, New PD.XSolidBrush(navy), marginL + 280, y)
            gfx.DrawString("Data: " & Now.ToString("dd/MM/yyyy"),
                           fontNorm, PD.XBrushes.Black, contentW + marginL - 120, y)
            y += 20

            ' ── Moltiplicatore ───────────────────────────────────
            gfx.DrawString("Moltiplicatore applicato: " & cmbMoltiplicatore.Text,
                           fontNorm, PD.XBrushes.Gray, marginL, y)
            y += 18

            ' ── Intestazione tabella ─────────────────────────────
            Dim headers() As String = {"Codice", "Descrizione", "Desc. Supplementare", "Qtà", "Costo unit.", "Costo tot.", "Pz in mac."}
            Dim x As Double = marginL
            gfx.DrawRectangle(New PD.XSolidBrush(navy), x, y, contentW, headerH)
            For i As Integer = 0 To headers.Length - 1
                Dim align As PD.XStringFormat = If(i >= 3, PD.XStringFormats.CenterRight, PD.XStringFormats.CenterLeft)
                Dim padding As Double = If(i >= 3, -4, 4)
                gfx.DrawString(headers(i), fontBold, PD.XBrushes.White,
                               New PD.XRect(x + padding, y, colW(i), headerH), align)
                x += colW(i)
            Next
            y += headerH

            ' ── Righe dati ───────────────────────────────────────
            Dim rowIdx As Integer = 0
            Dim totale As Decimal = 0

            For Each row As DataGridViewRow In dgvRicambi.Rows
                Dim cod As String = ValCella(row, COL_CODICE)
                If String.IsNullOrEmpty(cod) Then Continue For

                If y + rowH > pageH - 60 Then
                    gfx.Dispose()
                    page = doc.AddPage()
                    page.Width = pageW
                    page.Height = pageH
                    gfx = PD.XGraphics.FromPdfPage(page)
                    y = 30
                    x = marginL
                    gfx.DrawRectangle(New PD.XSolidBrush(navy), x, y, contentW, headerH)
                    For i As Integer = 0 To headers.Length - 1
                        Dim align As PD.XStringFormat = If(i >= 3, PD.XStringFormats.CenterRight, PD.XStringFormats.CenterLeft)
                        Dim padding As Double = If(i >= 3, -4, 4)
                        gfx.DrawString(headers(i), fontBold, PD.XBrushes.White,
                                       New PD.XRect(x + padding, y, colW(i), headerH), align)
                        x += colW(i)
                    Next
                    y += headerH
                    rowIdx = 0
                End If

                Dim bgColor As PD.XColor = If(rowIdx Mod 2 = 0, white, altRow)
                gfx.DrawRectangle(New PD.XSolidBrush(bgColor), marginL, y, contentW, rowH)

                Dim cells() As String = {
                    cod,
                    ValCella(row, COL_DESC),
                    ValCella(row, COL_DESC_SUP),
                    FormatDecimalCella(ValCella(row, COL_QTA), "0.##"),
                    FormatDecimalCella(ValCella(row, COL_COSTO), "N2"),
                    FormatDecimalCella(ValCella(row, COL_COSTO_TOT), "N2"),
                    ValCella(row, COL_PZ_IMP)
                }

                Dim ctDec As Decimal = 0
                Decimal.TryParse(cells(5).Replace(",", "."), System.Globalization.NumberStyles.Any,
                                 System.Globalization.CultureInfo.InvariantCulture, ctDec)
                totale += ctDec

                If Not String.IsNullOrEmpty(cells(5)) AndAlso ctDec > 0 Then
                    Dim xCostoTot As Double = marginL + colW(0) + colW(1) + colW(2) + colW(3) + colW(4)
                    gfx.DrawRectangle(New PD.XSolidBrush(green), xCostoTot, y, colW(5), rowH)
                End If

                x = marginL
                For i As Integer = 0 To cells.Length - 1
                    Dim align As PD.XStringFormat = If(i >= 3, PD.XStringFormats.CenterRight, PD.XStringFormats.CenterLeft)
                    Dim padding As Double = If(i >= 3, -4, 4)
                    Dim fnt As PD.XFont = If(i = 5, fontBold, fontNorm)
                    Dim testo As String = TroncaTesto(gfx, cells(i), fnt, colW(i) - 8)
                    gfx.DrawString(testo, fnt, PD.XBrushes.Black,
                                   New PD.XRect(x + padding, y + 1, colW(i) - 4, rowH - 2), align)
                    x += colW(i)
                Next

                gfx.DrawLine(New PD.XPen(lineColor, 0.3), marginL, y + rowH, marginL + contentW, y + rowH)

                y += rowH
                rowIdx += 1
            Next

            ' ── Riga totale ──────────────────────────────────────
            y += 6
            Dim xTotLabel As Double = marginL + colW(0) + colW(1) + colW(2) + colW(3) + colW(4)
            gfx.DrawRectangle(New PD.XSolidBrush(navy), xTotLabel - 70, y, 70 + colW(5), 22)
            gfx.DrawString("TOTALE:", fontBold, PD.XBrushes.White,
                           New PD.XRect(xTotLabel - 68, y + 2, 66, 18), PD.XStringFormats.CenterRight)
            gfx.DrawString("€ " & totale.ToString("N2"), fontBold, PD.XBrushes.White,
                           New PD.XRect(xTotLabel - 4, y + 2, colW(5) - 2, 18), PD.XStringFormats.CenterRight)

            gfx.Dispose()
            doc.Save(sfd.FileName)

            If MsgBox("PDF salvato. Aprirlo ora?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                Process.Start(New ProcessStartInfo(sfd.FileName) With {.UseShellExecute = True})
            End If
        Catch ex As Exception
            MsgBox("Errore generazione PDF: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  CHIUDI
    ' ─────────────────────────────────────────────────────────────────
    Private Sub btnChiudi_Click(sender As Object, e As EventArgs) Handles btnChiudi.Click
        Me.Close()
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  TASTO INVIO NELLA GRIGLIA: sposta alla cella successiva
    ' ─────────────────────────────────────────────────────────────────
    Private Sub dgvRicambi_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvRicambi.KeyDown
        If e.KeyCode = Keys.Return Then
            e.SuppressKeyPress = True
            SendKeys.Send("{TAB}")
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  PROPRIETA' PUBBLICA: numero righe salvate (usata da Scheda Tecnica)
    ' ─────────────────────────────────────────────────────────────────
    Public ReadOnly Property NumeroRighe As Integer
        Get
            Return dgvRicambi.Rows.Count
        End Get
    End Property

End Class
