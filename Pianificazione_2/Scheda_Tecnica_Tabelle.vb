Imports System.Data.SqlClient

Public Class Scheda_Tecnica_Tabelle

    Private dgvTipologie As DataGridView
    Private dgvModelli As DataGridView
    Private navy As Color = Color.FromArgb(22, 45, 84)

    Private Sub Scheda_Tecnica_Tabelle_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Font = New Font("Segoe UI", 9)
        CreaLayout()
        RiempiTipologie()
        RiempiModelli()
    End Sub

    Private Sub CreaLayout()
        ' Bottom close button
        Dim pnlBottom As New Panel
        pnlBottom.Height = 40
        pnlBottom.Dock = DockStyle.Bottom

        Dim btnClose As New Button
        btnClose.Text = "Chiudi"
        btnClose.Width = 100
        btnClose.Height = 28
        btnClose.Left = (Me.ClientSize.Width - 100) \ 2
        btnClose.Top = 6
        btnClose.FlatStyle = FlatStyle.Flat
        btnClose.BackColor = navy
        btnClose.ForeColor = Color.White
        btnClose.FlatAppearance.BorderSize = 0
        AddHandler btnClose.Click, Sub(s, ev) Me.Close()
        pnlBottom.Controls.Add(btnClose)

        ' Split panel
        Dim split As New SplitContainer
        split.Dock = DockStyle.Fill
        split.Orientation = Orientation.Vertical
        split.SplitterDistance = 350

        ' === Left: Tipologie ===
        Dim gbTip As New GroupBox
        gbTip.Text = "Tipologia macchina"
        gbTip.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        gbTip.ForeColor = navy
        gbTip.Dock = DockStyle.Fill

        dgvTipologie = CreaGriglia()
        dgvTipologie.Columns(1).HeaderText = "Tipologia"

        Dim pnlTipBtn As New Panel
        pnlTipBtn.Height = 32
        pnlTipBtn.Dock = DockStyle.Bottom

        Dim btnAddTip As New Button
        btnAddTip.Text = "Aggiungi"
        btnAddTip.Width = 90
        btnAddTip.Location = New Point(4, 4)
        btnAddTip.FlatStyle = FlatStyle.Flat
        AddHandler btnAddTip.Click, AddressOf BtnAddTipologia_Click

        Dim btnDelTip As New Button
        btnDelTip.Text = "Elimina"
        btnDelTip.Width = 90
        btnDelTip.Location = New Point(98, 4)
        btnDelTip.FlatStyle = FlatStyle.Flat
        AddHandler btnDelTip.Click, AddressOf BtnDelTipologia_Click

        pnlTipBtn.Controls.Add(btnAddTip)
        pnlTipBtn.Controls.Add(btnDelTip)
        gbTip.Controls.Add(dgvTipologie)
        gbTip.Controls.Add(pnlTipBtn)
        split.Panel1.Controls.Add(gbTip)

        ' === Right: Modelli ===
        Dim gbMod As New GroupBox
        gbMod.Text = "Modello macchina"
        gbMod.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        gbMod.ForeColor = navy
        gbMod.Dock = DockStyle.Fill

        dgvModelli = CreaGriglia()
        dgvModelli.Columns(1).HeaderText = "Modello"

        Dim pnlModBtn As New Panel
        pnlModBtn.Height = 32
        pnlModBtn.Dock = DockStyle.Bottom

        Dim btnAddMod As New Button
        btnAddMod.Text = "Aggiungi"
        btnAddMod.Width = 90
        btnAddMod.Location = New Point(4, 4)
        btnAddMod.FlatStyle = FlatStyle.Flat
        AddHandler btnAddMod.Click, AddressOf BtnAddModello_Click

        Dim btnDelMod As New Button
        btnDelMod.Text = "Elimina"
        btnDelMod.Width = 90
        btnDelMod.Location = New Point(98, 4)
        btnDelMod.FlatStyle = FlatStyle.Flat
        AddHandler btnDelMod.Click, AddressOf BtnDelModello_Click

        pnlModBtn.Controls.Add(btnAddMod)
        pnlModBtn.Controls.Add(btnDelMod)
        gbMod.Controls.Add(dgvModelli)
        gbMod.Controls.Add(pnlModBtn)
        split.Panel2.Controls.Add(gbMod)

        Me.Controls.Add(split)
        Me.Controls.Add(pnlBottom)
    End Sub

    Private Function CreaGriglia() As DataGridView
        Dim dgv As New DataGridView
        dgv.Dock = DockStyle.Fill
        dgv.ColumnCount = 2
        dgv.Columns(0).Name = "id"
        dgv.Columns(0).HeaderText = "ID"
        dgv.Columns(0).Width = 45
        dgv.Columns(0).ReadOnly = True
        dgv.Columns(1).Name = "valore"
        dgv.Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        dgv.AllowUserToAddRows = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.BorderStyle = BorderStyle.None
        dgv.BackgroundColor = Color.White
        dgv.GridColor = Color.FromArgb(210, 220, 235)
        dgv.ColumnHeadersDefaultCellStyle.BackColor = navy
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        dgv.EnableHeadersVisualStyles = False
        dgv.RowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 225, 245)
        dgv.RowsDefaultCellStyle.SelectionForeColor = navy
        Return dgv
    End Function

    Sub RiempiTipologie()
        dgvTipologie.Rows.Clear()
        Try
            Using Cnn As New SqlConnection(Homepage.sap_tirelli)
                Cnn.Open()
                Using CMD As New SqlCommand("SELECT id, valore FROM [TIRELLI_40].[dbo].[ST_lookup_tipologia_macchina] ORDER BY valore", Cnn)
                    Using r = CMD.ExecuteReader()
                        While r.Read()
                            dgvTipologie.Rows.Add(r("id"), r("valore"))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Errore lettura tipologie: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Sub RiempiModelli()
        dgvModelli.Rows.Clear()
        Try
            Using Cnn As New SqlConnection(Homepage.sap_tirelli)
                Cnn.Open()
                Using CMD As New SqlCommand("SELECT id, valore FROM [TIRELLI_40].[dbo].[ST_lookup_modello_macchina] ORDER BY valore", Cnn)
                    Using r = CMD.ExecuteReader()
                        While r.Read()
                            dgvModelli.Rows.Add(r("id"), r("valore"))
                        End While
                    End Using
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Errore lettura modelli: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub BtnAddTipologia_Click(sender As Object, e As EventArgs)
        Dim valore = InputBox("Inserisci tipologia macchina:", "Nuova tipologia")
        If valore.Trim() = "" Then Return
        Try
            Using Cnn As New SqlConnection(Homepage.sap_tirelli)
                Cnn.Open()
                Using CMD As New SqlCommand("INSERT INTO [TIRELLI_40].[dbo].[ST_lookup_tipologia_macchina] (valore) VALUES (@v)", Cnn)
                    CMD.Parameters.AddWithValue("@v", valore.Trim())
                    CMD.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Errore inserimento: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        RiempiTipologie()
    End Sub

    Private Sub BtnDelTipologia_Click(sender As Object, e As EventArgs)
        If dgvTipologie.SelectedRows.Count = 0 Then Return
        Dim id = dgvTipologie.SelectedRows(0).Cells("id").Value
        Dim valore = dgvTipologie.SelectedRows(0).Cells("valore").Value
        If MessageBox.Show($"Eliminare la tipologia ""{valore}""?", "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                Using Cnn As New SqlConnection(Homepage.sap_tirelli)
                    Cnn.Open()
                    Using CMD As New SqlCommand("DELETE FROM [TIRELLI_40].[dbo].[ST_lookup_tipologia_macchina] WHERE id=@id", Cnn)
                        CMD.Parameters.AddWithValue("@id", id)
                        CMD.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Errore eliminazione: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
            RiempiTipologie()
        End If
    End Sub

    Private Sub BtnAddModello_Click(sender As Object, e As EventArgs)
        Dim valore = InputBox("Inserisci modello macchina:", "Nuovo modello")
        If valore.Trim() = "" Then Return
        Try
            Using Cnn As New SqlConnection(Homepage.sap_tirelli)
                Cnn.Open()
                Using CMD As New SqlCommand("INSERT INTO [TIRELLI_40].[dbo].[ST_lookup_modello_macchina] (valore) VALUES (@v)", Cnn)
                    CMD.Parameters.AddWithValue("@v", valore.Trim())
                    CMD.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Errore inserimento: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
        RiempiModelli()
    End Sub

    Private Sub BtnDelModello_Click(sender As Object, e As EventArgs)
        If dgvModelli.SelectedRows.Count = 0 Then Return
        Dim id = dgvModelli.SelectedRows(0).Cells("id").Value
        Dim valore = dgvModelli.SelectedRows(0).Cells("valore").Value
        If MessageBox.Show($"Eliminare il modello ""{valore}""?", "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Try
                Using Cnn As New SqlConnection(Homepage.sap_tirelli)
                    Cnn.Open()
                    Using CMD As New SqlCommand("DELETE FROM [TIRELLI_40].[dbo].[ST_lookup_modello_macchina] WHERE id=@id", Cnn)
                        CMD.Parameters.AddWithValue("@id", id)
                        CMD.ExecuteNonQuery()
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Errore eliminazione: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
            RiempiModelli()
        End If
    End Sub

End Class
