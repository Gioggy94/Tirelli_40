Imports System.Data.SqlClient

Public Class Entrate_merci_storico

    Private Sub Entrate_merci_storico_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cmbStato.Items.Add("")
        cmbStato.Items.Add("OK")
        cmbStato.Items.Add("Codice non in questo ordine")
        cmbStato.Items.Add("Ordine non trovato")
        cmbStato.SelectedIndex = 0
        dtpDal.Value = Date.Today.AddMonths(-3)
        dtpAl.Value = Date.Today
        Carica()
    End Sub

    Private Sub Carica()
        Dim where As New List(Of String)
        Dim params As New List(Of SqlParameter)

        where.Add("DDT_Data >= @dal")
        params.Add(New SqlParameter("@dal", dtpDal.Value.Date))
        where.Add("DDT_Data <= @al")
        params.Add(New SqlParameter("@al", dtpAl.Value.Date))

        If txtFornitore.Text.Trim() <> "" Then
            where.Add("Fornitore LIKE @forn")
            params.Add(New SqlParameter("@forn", "%" & txtFornitore.Text.Trim() & "%"))
        End If
        If txtCodice.Text.Trim() <> "" Then
            where.Add("(Codice_Articolo LIKE @cod OR Disegno LIKE @cod)")
            params.Add(New SqlParameter("@cod", "%" & txtCodice.Text.Trim() & "%"))
        End If
        If txtOrdine.Text.Trim() <> "" Then
            where.Add("Ordine_Acquisto LIKE @ord")
            params.Add(New SqlParameter("@ord", "%" & txtOrdine.Text.Trim() & "%"))
        End If
        If cmbStato.SelectedItem IsNot Nothing AndAlso cmbStato.SelectedItem.ToString() <> "" Then
            where.Add("Stato LIKE @stato")
            params.Add(New SqlParameter("@stato", "%" & cmbStato.SelectedItem.ToString() & "%"))
        End If
        If txtDipendente.Text.Trim() <> "" Then
            where.Add("(T1.FIRSTNAME LIKE @dip OR T1.LASTNAME LIKE @dip)")
            params.Add(New SqlParameter("@dip", "%" & txtDipendente.Text.Trim() & "%"))
        End If

        Dim sql As String = "SELECT t0.ID, t0.DDT_Numero, t0.DDT_Data, t0.Fornitore, t0.Ordine_Acquisto, " &
                            "t0.Codice_Articolo, t0.Disegno, t0.UM, t0.Quantita, t0.Stato, t0.PDF_File, t0.Data_Inserimento, t0.Utente, " &
                            "CONCAT(T1.FIRSTNAME, ' ', T1.LASTNAME) AS Dipendente " &
                            "FROM [TIRELLI_40].[dbo].[Entrate_merci] t0 " &
                            "LEFT JOIN [TIRELLI_40].[dbo].[OHEM] t1 ON T0.Utente_Galileo = T1.Galileo"
        If where.Count > 0 Then sql &= " WHERE " & String.Join(" AND ", where)
        sql &= " ORDER BY t0.Data_Inserimento DESC, t0.ID DESC"

        Try
            Using conn As New SqlConnection(Homepage.sap_tirelli)
                conn.Open()
                Dim da As New SqlDataAdapter(sql, conn)
                For Each p As SqlParameter In params
                    da.SelectCommand.Parameters.Add(p)
                Next
                Dim dt As New DataTable()
                da.Fill(dt)
                dgvStorico.DataSource = dt
                lblCount.Text = $"{dt.Rows.Count} righe"
            End Using
        Catch ex As Exception
            MsgBox("Errore caricamento: " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' Colora le righe per stato
        For Each row As DataGridViewRow In dgvStorico.Rows
            Dim stato As String = If(row.Cells("Stato").Value IsNot Nothing, row.Cells("Stato").Value.ToString(), "")
            If stato = "OK" Then
                row.DefaultCellStyle.BackColor = Color.FromArgb(190, 235, 190)
            ElseIf stato.StartsWith("Q DDT") Then
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 170)
            ElseIf stato = "Codice non in questo ordine" Then
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 210, 140)
            ElseIf stato = "Ordine non trovato" Then
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 180, 180)
            End If
        Next
    End Sub

    Private Sub btnCerca_Click(sender As Object, e As EventArgs) Handles btnCerca.Click
        Carica()
    End Sub

    Private Sub btnElimina_Click(sender As Object, e As EventArgs) Handles btnElimina.Click
        Dim ids As New List(Of Integer)
        For Each row As DataGridViewRow In dgvStorico.SelectedRows
            If row.Cells("ID").Value IsNot Nothing Then
                ids.Add(CInt(row.Cells("ID").Value))
            End If
        Next
        If ids.Count = 0 Then
            MsgBox("Seleziona almeno una riga da eliminare.", MsgBoxStyle.Information)
            Return
        End If
        If MsgBox($"Eliminare {ids.Count} righe selezionate?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) <> MsgBoxResult.Yes Then Return

        Try
            Using conn As New SqlConnection(Homepage.sap_tirelli)
                conn.Open()
                Dim cmd As New SqlCommand($"DELETE FROM [TIRELLI_40].[dbo].[Entrate_merci] WHERE ID IN ({String.Join(",", ids)})", conn)
                cmd.ExecuteNonQuery()
            End Using
            Carica()
        Catch ex As Exception
            MsgBox("Errore eliminazione: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnChiudi_Click(sender As Object, e As EventArgs) Handles btnChiudi.Click
        Me.Close()
    End Sub

    Private Sub txtFiltro_KeyDown(sender As Object, e As KeyEventArgs) Handles txtFornitore.KeyDown, txtCodice.KeyDown, txtOrdine.KeyDown, txtDipendente.KeyDown
        If e.KeyCode = Keys.Enter Then Carica()
    End Sub

End Class
