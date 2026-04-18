Imports System.Data.SqlClient
Imports System.Data

Public Class Form_Importa_Lista_Ricambi

    Public commessaDestinazione As String = ""
    Public revDestinazione As Integer = 0

    ' ─────────────────────────────────────────────────────────────────
    '  LOAD
    ' ─────────────────────────────────────────────────────────────────
    Private Sub Form_Importa_Lista_Ricambi_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Importa lista ricambi da altra commessa"
        Me.Size = New System.Drawing.Size(820, 520)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Font = New System.Drawing.Font("Segoe UI", 8.5!)
        ApplicaStile()
        CercaCommesse("")
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  STILE
    ' ─────────────────────────────────────────────────────────────────
    Private Sub ApplicaStile()
        Dim navy As System.Drawing.Color = System.Drawing.Color.FromArgb(22, 45, 84)

        btnImporta.BackColor = navy
        btnImporta.ForeColor = System.Drawing.Color.White
        btnImporta.FlatStyle = FlatStyle.Flat
        btnImporta.FlatAppearance.BorderSize = 0

        btnAnnulla.FlatStyle = FlatStyle.Flat

        For Each dgv As DataGridView In {dgvCommesse, dgvListe}
            dgv.ColumnHeadersDefaultCellStyle.BackColor = navy
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = System.Drawing.Color.White
            dgv.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Bold)
            dgv.EnableHeadersVisualStyles = False
            dgv.BackgroundColor = System.Drawing.Color.White
            dgv.BorderStyle = BorderStyle.None
            dgv.GridColor = System.Drawing.Color.FromArgb(210, 220, 235)
            dgv.RowsDefaultCellStyle.SelectionBackColor = System.Drawing.Color.FromArgb(210, 225, 245)
            dgv.RowsDefaultCellStyle.SelectionForeColor = navy
            dgv.AlternatingRowsDefaultCellStyle.BackColor = System.Drawing.Color.FromArgb(245, 248, 255)
        Next
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  CERCA COMMESSE (che hanno liste salvate, join JGALCOM)
    ' ─────────────────────────────────────────────────────────────────
    Private Sub CercaCommesse(filtro As String)
        dgvListe.DataSource = Nothing
        dgvListe.Columns.Clear()
        lblListe.Text = "Liste disponibili"
        btnImporta.Enabled = False

        Dim filtroUpper As String = filtro.Trim().ToUpper()

        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnn
                    cmd.CommandTimeout = 30
                    cmd.CommandText = "
SELECT r.Commessa,
       ISNULL(RTRIM(j.itemname), '') AS Macchina,
       ISNULL(RTRIM(j.dscli_fatt), '') AS Cliente,
       COUNT(DISTINCT r.NomeLista) AS NListe
FROM [Tirelli_40].[dbo].[Lista_Ricambi_Consigliati_Righe] r
LEFT JOIN [AS400].[S786FAD1].[TIR90VIS].[JGALCOM] j
    ON RTRIM(LTRIM(j.matricola)) = r.Commessa
WHERE r.Commessa <> @commDest
  AND (@filtro = ''
       OR UPPER(r.Commessa) LIKE '%' + @filtro + '%'
       OR UPPER(ISNULL(j.itemname,'')) LIKE '%' + @filtro + '%'
       OR UPPER(ISNULL(j.dscli_fatt,'')) LIKE '%' + @filtro + '%')
GROUP BY r.Commessa, j.itemname, j.dscli_fatt
ORDER BY r.Commessa DESC"
                    cmd.Parameters.AddWithValue("@commDest", commessaDestinazione)
                    cmd.Parameters.AddWithValue("@filtro", filtroUpper)

                    Dim dt As New DataTable()
                    Using rd As SqlDataReader = cmd.ExecuteReader()
                        dt.Load(rd)
                    End Using

                    dgvCommesse.AutoGenerateColumns = False
                    dgvCommesse.DataSource = Nothing
                    dgvCommesse.Columns.Clear()
                    dgvCommesse.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "colComm", .HeaderText = "Commessa", .DataPropertyName = "Commessa", .Width = 80})
                    dgvCommesse.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "colMacc", .HeaderText = "Macchina", .DataPropertyName = "Macchina", .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill})
                    dgvCommesse.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "colCli", .HeaderText = "Cliente", .DataPropertyName = "Cliente", .Width = 120})
                    dgvCommesse.Columns.Add(New DataGridViewTextBoxColumn() With {
                        .Name = "colN", .HeaderText = "Liste", .DataPropertyName = "NListe", .Width = 44,
                        .DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleRight}})
                    dgvCommesse.DataSource = dt
                    lblCommesse.Text = "Commesse trovate: " & dt.Rows.Count
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Errore ricerca commesse: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  SELEZIONE COMMESSA → carica le sue liste
    ' ─────────────────────────────────────────────────────────────────
    Private Sub dgvCommesse_SelectionChanged(sender As Object, e As EventArgs) Handles dgvCommesse.SelectionChanged
        btnImporta.Enabled = False
        dgvListe.DataSource = Nothing
        dgvListe.Columns.Clear()
        If dgvCommesse.CurrentRow Is Nothing Then Return

        Dim commSorg As String = dgvCommesse.CurrentRow.Cells("colComm").Value?.ToString()
        If String.IsNullOrEmpty(commSorg) Then Return

        lblListe.Text = "Liste di " & commSorg

        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnn
                    cmd.CommandText = "
SELECT NomeLista, COUNT(*) AS NArticoli, SUM(CostoTot) AS Totale
FROM [Tirelli_40].[dbo].[Lista_Ricambi_Consigliati_Righe]
WHERE Commessa=@c
GROUP BY NomeLista
ORDER BY NomeLista"
                    cmd.Parameters.AddWithValue("@c", commSorg)

                    Dim dt As New DataTable()
                    Using rd As SqlDataReader = cmd.ExecuteReader()
                        dt.Load(rd)
                    End Using

                    dgvListe.AutoGenerateColumns = False
                    dgvListe.Columns.Add(New DataGridViewTextBoxColumn() With {.Name = "colNome", .HeaderText = "Nome lista", .DataPropertyName = "NomeLista", .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill})
                    dgvListe.Columns.Add(New DataGridViewTextBoxColumn() With {
                        .Name = "colArt", .HeaderText = "Art.", .DataPropertyName = "NArticoli", .Width = 40,
                        .DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleRight}})
                    dgvListe.Columns.Add(New DataGridViewTextBoxColumn() With {
                        .Name = "colTot", .HeaderText = "€ tot", .DataPropertyName = "Totale", .Width = 80,
                        .DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N2"}})
                    dgvListe.DataSource = dt
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Errore caricamento liste: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub dgvListe_SelectionChanged(sender As Object, e As EventArgs) Handles dgvListe.SelectionChanged
        btnImporta.Enabled = (dgvListe.CurrentRow IsNot Nothing)
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  CERCA
    ' ─────────────────────────────────────────────────────────────────
    Private Sub btnCerca_Click(sender As Object, e As EventArgs) Handles btnCerca.Click
        CercaCommesse(txtFiltro.Text)
    End Sub

    Private Sub txtFiltro_KeyDown(sender As Object, e As KeyEventArgs) Handles txtFiltro.KeyDown
        If e.KeyCode = Keys.Return Then CercaCommesse(txtFiltro.Text)
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  IMPORTA
    ' ─────────────────────────────────────────────────────────────────
    Private Sub btnImporta_Click(sender As Object, e As EventArgs) Handles btnImporta.Click
        If dgvCommesse.CurrentRow Is Nothing OrElse dgvListe.CurrentRow Is Nothing Then Return

        Dim commSorg As String = dgvCommesse.CurrentRow.Cells("colComm").Value?.ToString()
        Dim nomeListaSorg As String = dgvListe.CurrentRow.Cells("colNome").Value?.ToString()
        If String.IsNullOrEmpty(commSorg) OrElse String.IsNullOrEmpty(nomeListaSorg) Then Return

        ' Nome lista nella destinazione — proponi lo stesso, permetti di cambiarlo
        Dim nomeListaDest As String = InputBox(
            "Nome della lista nella commessa " & commessaDestinazione & ":",
            "Importa lista", nomeListaSorg)
        If String.IsNullOrWhiteSpace(nomeListaDest) Then Return
        nomeListaDest = nomeListaDest.Trim()

        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                ' Controlla se esiste già
                Dim esiste As Integer = 0
                Using cmdChk As New SqlCommand(
                    "SELECT COUNT(*) FROM [Tirelli_40].[dbo].[Lista_Ricambi_Consigliati_Righe] WHERE Commessa=@c AND Rev=@r AND NomeLista=@n",
                    cnn)
                    cmdChk.Parameters.AddWithValue("@c", commessaDestinazione)
                    cmdChk.Parameters.AddWithValue("@r", revDestinazione)
                    cmdChk.Parameters.AddWithValue("@n", nomeListaDest)
                    esiste = CInt(cmdChk.ExecuteScalar())
                End Using

                If esiste > 0 Then
                    If MsgBox("La lista """ & nomeListaDest & """ esiste già nella commessa " & commessaDestinazione & ". Sovrascrivere?",
                              MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation) <> MsgBoxResult.Yes Then Return
                    Using cmdDel As New SqlCommand(
                        "DELETE FROM [Tirelli_40].[dbo].[Lista_Ricambi_Consigliati_Righe] WHERE Commessa=@c AND Rev=@r AND NomeLista=@n",
                        cnn)
                        cmdDel.Parameters.AddWithValue("@c", commessaDestinazione)
                        cmdDel.Parameters.AddWithValue("@r", revDestinazione)
                        cmdDel.Parameters.AddWithValue("@n", nomeListaDest)
                        cmdDel.ExecuteNonQuery()
                    End Using
                End If

                ' Copia righe
                Using cmdIns As New SqlCommand()
                    cmdIns.Connection = cnn
                    cmdIns.CommandText = "
INSERT INTO [Tirelli_40].[dbo].[Lista_Ricambi_Consigliati_Righe]
    (Commessa, Rev, NomeLista, Moltiplicatore, Codice, Descrizione, DescrizioneSup, Quantita, Costo, CostoTot)
SELECT @commDest, @revDest, @nomeDest, Moltiplicatore, Codice, Descrizione, DescrizioneSup, Quantita, Costo, CostoTot
FROM [Tirelli_40].[dbo].[Lista_Ricambi_Consigliati_Righe]
WHERE Commessa=@commSorg AND NomeLista=@nomeSorg"
                    cmdIns.Parameters.AddWithValue("@commDest", commessaDestinazione)
                    cmdIns.Parameters.AddWithValue("@revDest", revDestinazione)
                    cmdIns.Parameters.AddWithValue("@nomeDest", nomeListaDest)
                    cmdIns.Parameters.AddWithValue("@commSorg", commSorg)
                    cmdIns.Parameters.AddWithValue("@nomeSorg", nomeListaSorg)
                    Dim n As Integer = cmdIns.ExecuteNonQuery()
                    MsgBox("Importate " & n & " righe nella lista """ & nomeListaDest & """.", MsgBoxStyle.Information)
                End Using
            End Using
            Me.Close()
        Catch ex As Exception
            MsgBox("Errore importazione: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnAnnulla_Click(sender As Object, e As EventArgs) Handles btnAnnulla.Click
        Me.Close()
    End Sub

End Class
