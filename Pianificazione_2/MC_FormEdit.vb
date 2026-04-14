Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

' Form per associare Modello/Tipo/Lingua a una macchina proveniente da AS400
Public Class MC_FrmEditMacchina
    Inherits Form

    Private _m As MC_Macchina
    Private _db As MC_DatabaseService

    Public Sub New(m As MC_Macchina, db As MC_DatabaseService)
        _m  = m
        _db = db
        Me.Text            = "Associa dati manuale"
        Me.Size            = New Size(480, 380)
        Me.StartPosition   = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.Font            = New Font("Segoe UI", 9)

        Dim y = 16

        ' ── Campi readonly da AS400 ──────────────────────
        For Each pair In {("Matricola:", m.Matricola), ("Nome macchina:", m.NomeMacchina), ("Cliente:", m.ClienteFinale)}
            Me.Controls.Add(New Label() With {.Text = pair.Item1, .Location = New Point(16, y + 3), .Size = New Size(120, 20)})
            Me.Controls.Add(New Label() With {
                .Text = pair.Item2, .Location = New Point(140, y + 3), .Size = New Size(300, 20),
                .Font = New Font("Segoe UI", 9, FontStyle.Bold), .ForeColor = Color.FromArgb(40, 40, 40)
            })
            y += 28
        Next
        y += 8

        ' ── Modello (dropdown) ───────────────────────────
        Me.Controls.Add(New Label() With {.Text = "Modello:", .Location = New Point(16, y + 3), .Size = New Size(120, 20)})
        Dim cmbMod As New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList, .Location = New Point(140, y), .Size = New Size(280, 24)
        }
        Dim modelli = _db.GetModelli()
        cmbMod.DataSource    = modelli
        cmbMod.DisplayMember = "Nome"
        cmbMod.ValueMember   = "ID"
        Dim selMod = modelli.FirstOrDefault(Function(x) x.Nome = _m.Modello)
        If selMod IsNot Nothing Then cmbMod.SelectedItem = selMod
        Me.Controls.Add(cmbMod)
        y += 32

        ' ── Tipo macchina (dropdown) ─────────────────────
        Me.Controls.Add(New Label() With {.Text = "Tipo macchina:", .Location = New Point(16, y + 3), .Size = New Size(120, 20)})
        Dim cmbTipo As New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList, .Location = New Point(140, y), .Size = New Size(280, 24)
        }
        Dim tipi = _db.GetTipiMacchina()
        cmbTipo.DataSource    = tipi
        cmbTipo.DisplayMember = "Nome"
        cmbTipo.ValueMember   = "ID"
        Dim selTipo = tipi.FirstOrDefault(Function(x) x.Nome = _m.TipoMacchina)
        If selTipo IsNot Nothing Then cmbTipo.SelectedItem = selTipo
        Me.Controls.Add(cmbTipo)
        y += 32

        ' ── Lingua ──────────────────────────────────────
        Me.Controls.Add(New Label() With {.Text = "Lingua:", .Location = New Point(16, y + 3), .Size = New Size(120, 20)})
        Dim cmbLng As New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList, .Location = New Point(140, y), .Size = New Size(120, 24)
        }
        For Each l In {"IT", "EN", "FR", "ES", "DE"} : cmbLng.Items.Add(l) : Next
        cmbLng.SelectedItem = If(String.IsNullOrEmpty(_m.LinguaCodice), "IT", _m.LinguaCodice)
        Me.Controls.Add(cmbLng)
        y += 32

        ' ── Note ────────────────────────────────────────
        Me.Controls.Add(New Label() With {.Text = "Note:", .Location = New Point(16, y + 3), .Size = New Size(120, 20)})
        Dim txtNote As New TextBox() With {
            .Text = _m.Note, .Location = New Point(140, y), .Size = New Size(295, 54), .Multiline = True
        }
        Me.Controls.Add(txtNote)
        y += 66

        ' ── Pulsanti ────────────────────────────────────
        Dim btnOk  As New Button() With {.Text = "Salva",   .Location = New Point(140, y), .Size = New Size(100, 32)}
        Dim btnAnn As New Button() With {.Text = "Annulla", .Location = New Point(250, y), .Size = New Size(100, 32), .DialogResult = DialogResult.Cancel}
        MC_PanelBuild.StyleButton(btnOk, True)
        MC_PanelBuild.StyleButton(btnAnn, False)
        Me.Controls.AddRange({btnOk, btnAnn})
        Me.CancelButton = btnAnn

        AddHandler btnOk.Click, Sub(s, e)
            _m.Modello      = If(cmbMod.SelectedItem IsNot Nothing, DirectCast(cmbMod.SelectedItem, MC_Modello).Nome, "")
            _m.TipoMacchina = If(cmbTipo.SelectedItem IsNot Nothing, DirectCast(cmbTipo.SelectedItem, MC_TipoMacchina).Nome, "")
            _m.LinguaCodice = cmbLng.SelectedItem?.ToString()
            _m.Note         = txtNote.Text.Trim()
            Try
                _m.ID = _db.SalvaExtraMacchina(_m)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Catch ex As Exception
                MessageBox.Show("Errore salvataggio: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Sub
End Class

' Form per gestire le voci dei dropdown (Modelli / Tipi macchina)
Public Class MC_FrmGestisciLookup
    Inherits Form

    Private _tipo As String   ' "Modelli" oppure "TipiMacchina"
    Private _db As MC_DatabaseService
    Private _lst As ListBox

    Public Sub New(tipo As String, db As MC_DatabaseService)
        _tipo = tipo
        _db   = db
        Me.Text            = If(tipo = "Modelli", "Gestisci Modelli", "Gestisci Tipi macchina")
        Me.Size            = New Size(360, 380)
        Me.StartPosition   = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.Font            = New Font("Segoe UI", 9)

        _lst = New ListBox() With {
            .Location = New Point(16, 16), .Size = New Size(310, 260), .Font = New Font("Segoe UI", 10)
        }
        Me.Controls.Add(_lst)

        Dim txtNuovo As New TextBox() With {.Location = New Point(16, 284), .Size = New Size(200, 24)}
        Dim btnAgg   As New Button() With {.Text = "+ Aggiungi", .Location = New Point(224, 283), .Size = New Size(100, 26)}
        Dim btnEl    As New Button() With {.Text = "Elimina selezionato", .Location = New Point(16, 316), .Size = New Size(170, 28)}
        Dim btnChiudi As New Button() With {.Text = "Chiudi", .Location = New Point(200, 316), .Size = New Size(80, 28), .DialogResult = DialogResult.OK}
        MC_PanelBuild.StyleButton(btnAgg, True)
        MC_PanelBuild.StyleButton(btnEl, False)
        MC_PanelBuild.StyleButton(btnChiudi, False)
        Me.Controls.AddRange({txtNuovo, btnAgg, btnEl, btnChiudi})
        Me.AcceptButton = btnChiudi

        AddHandler btnAgg.Click, Sub(s, e)
            Dim nome = txtNuovo.Text.Trim()
            If String.IsNullOrEmpty(nome) Then Return
            Try
                If _tipo = "Modelli" Then _db.SalvaModello(nome) Else _db.SalvaTipoMacchina(nome)
                txtNuovo.Clear()
                Ricarica()
            Catch ex As Exception
                MessageBox.Show("Errore: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        AddHandler btnEl.Click, Sub(s, e)
            If _lst.SelectedItem Is Nothing Then Return
            If MessageBox.Show($"Eliminare '{_lst.SelectedItem}'?", "Conferma",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then Return
            Try
                Dim id = CInt(_lst.SelectedValue)
                If _tipo = "Modelli" Then _db.EliminaModello(id) Else _db.EliminaTipoMacchina(id)
                Ricarica()
            Catch ex As Exception
                MessageBox.Show("Errore: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Ricarica()
    End Sub

    Private Sub Ricarica()
        If _tipo = "Modelli" Then
            _lst.DataSource    = _db.GetModelli()
            _lst.DisplayMember = "Nome"
            _lst.ValueMember   = "ID"
        Else
            _lst.DataSource    = _db.GetTipiMacchina()
            _lst.DisplayMember = "Nome"
            _lst.ValueMember   = "ID"
        End If
    End Sub
End Class

' ══════════════════════════════════════════════
Public Class MC_FrmEditCodiceErrore
    Inherits Form

    Private _err As MC_CodiceErrore
    Private _db As MC_DatabaseService
    Private _macchinaID As Integer

    Public Sub New(err As MC_CodiceErrore, macchinaID As Integer, db As MC_DatabaseService)
        _err = If(err IsNot Nothing, err, New MC_CodiceErrore() With {.MacchinaID = macchinaID})
        _macchinaID = macchinaID
        _db = db
        Me.Text = If(err IsNot Nothing AndAlso err.ID > 0, "Modifica codice errore", "Nuovo codice errore")
        Me.Size = New Size(580, 560)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Font = New Font("Segoe UI", 9)

        Dim txtCodice As New TextBox() With {.Text = _err.Codice, .Location = New Point(150, 20), .Size = New Size(390, 24)}
        Dim txtTitolo As New TextBox() With {.Text = _err.Titolo, .Location = New Point(150, 54), .Size = New Size(390, 24)}
        Me.Controls.Add(New Label() With {.Text = "Codice* (es. ALM_001):", .Location = New Point(16, 23), .Size = New Size(130, 20)})
        Me.Controls.Add(New Label() With {.Text = "Titolo*:", .Location = New Point(16, 57), .Size = New Size(130, 20)})
        Me.Controls.AddRange({txtCodice, txtTitolo})

        Me.Controls.Add(New Label() With {.Text = "Gravità:", .Location = New Point(16, 89), .Size = New Size(130, 20)})
        Dim cmbGrav As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Location = New Point(150, 86), .Size = New Size(130, 24)}
        For Each g In {"Avviso", "Allarme", "Blocco"} : cmbGrav.Items.Add(g) : Next
        cmbGrav.SelectedItem = If(String.IsNullOrEmpty(_err.Gravita), "Avviso", _err.Gravita)
        Me.Controls.Add(cmbGrav)

        Me.Controls.Add(New Label() With {.Text = "Descrizione:", .Location = New Point(16, 123), .AutoSize = True})
        Dim txtDesc As New TextBox() With {.Text = _err.Descrizione, .Location = New Point(16, 143), .Size = New Size(524, 70), .Multiline = True}
        Me.Controls.Add(txtDesc)

        Me.Controls.Add(New Label() With {.Text = "Causa:", .Location = New Point(16, 222), .AutoSize = True})
        Dim txtCausa As New TextBox() With {.Text = _err.Causa, .Location = New Point(16, 242), .Size = New Size(524, 60), .Multiline = True}
        Me.Controls.Add(txtCausa)

        Me.Controls.Add(New Label() With {.Text = "Rimedio:", .Location = New Point(16, 312), .AutoSize = True})
        Dim txtRim As New TextBox() With {.Text = _err.Rimedio, .Location = New Point(16, 332), .Size = New Size(524, 60), .Multiline = True}
        Me.Controls.Add(txtRim)

        Me.Controls.Add(New Label() With {.Text = "Screenshot:", .Location = New Point(16, 402), .AutoSize = True})
        Dim txtScreen As New TextBox() With {.Text = _err.NomeScreenshot, .Location = New Point(16, 422), .Size = New Size(400, 24)}
        Dim btnScreen As New Button() With {.Text = "Sfoglia...", .Location = New Point(424, 422), .Size = New Size(80, 24)}
        MC_PanelBuild.StyleButton(btnScreen, False)
        Me.Controls.AddRange({txtScreen, btnScreen})
        AddHandler btnScreen.Click, Sub(s, e)
            Using ofd As New OpenFileDialog() With {.Filter = "Immagini|*.png;*.jpg;*.jpeg;*.bmp"}
                If ofd.ShowDialog() = DialogResult.OK Then
                    txtScreen.Text = IO.Path.GetFileName(ofd.FileName)
                    _err.PathScreenshot = ofd.FileName
                End If
            End Using
        End Sub

        Dim btnOk  As New Button() With {.Text = "Salva",   .Location = New Point(150, 466), .Size = New Size(100, 32)}
        Dim btnAnn As New Button() With {.Text = "Annulla", .Location = New Point(260, 466), .Size = New Size(100, 32), .DialogResult = DialogResult.Cancel}
        MC_PanelBuild.StyleButton(btnOk, True)
        MC_PanelBuild.StyleButton(btnAnn, False)
        Me.Controls.AddRange({btnOk, btnAnn})
        Me.CancelButton = btnAnn

        AddHandler btnOk.Click, Sub(s, e)
            _err.MacchinaID     = _macchinaID
            _err.Codice         = txtCodice.Text.Trim()
            _err.Titolo         = txtTitolo.Text.Trim()
            _err.Gravita        = cmbGrav.SelectedItem?.ToString()
            _err.Descrizione    = txtDesc.Text.Trim()
            _err.Causa          = txtCausa.Text.Trim()
            _err.Rimedio        = txtRim.Text.Trim()
            _err.NomeScreenshot = txtScreen.Text.Trim()
            If String.IsNullOrEmpty(_err.Codice) OrElse String.IsNullOrEmpty(_err.Titolo) Then
                MessageBox.Show("Codice e Titolo sono obbligatori.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Try
                _db.SalvaCodiceErrore(_err)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Catch ex As Exception
                MessageBox.Show("Errore: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Sub
End Class

' ══════════════════════════════════════════════
Public Class MC_FrmTestoGenerato
    Inherits Form

    Public Sub New(titolo As String, testo As String)
        Me.Text = titolo
        Me.Size = New Size(760, 580)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.Font = New Font("Segoe UI", 9)

        Dim rtb As New RichTextBox() With {
            .Dock = DockStyle.Fill, .Text = testo,
            .Font = New Font("Segoe UI", 9.5F), .ReadOnly = False, .BackColor = Color.White
        }
        Dim pnlBtn As New Panel() With {.Dock = DockStyle.Bottom, .Height = 46, .BackColor = Color.FromArgb(245, 245, 242)}
        Dim btnCopia  As New Button() With {.Text = "Copia testo", .Location = New Point(12, 8), .Size = New Size(120, 30)}
        Dim btnChiudi As New Button() With {.Text = "Chiudi",      .Location = New Point(142, 8), .Size = New Size(80, 30), .DialogResult = DialogResult.OK}
        MC_PanelBuild.StyleButton(btnCopia, True)
        MC_PanelBuild.StyleButton(btnChiudi, False)
        AddHandler btnCopia.Click, Sub(s, e) Clipboard.SetText(rtb.Text)
        pnlBtn.Controls.AddRange({btnCopia, btnChiudi})
        Me.Controls.Add(rtb)
        Me.Controls.Add(pnlBtn)
    End Sub
End Class

' ══════════════════════════════════════════════
' Form per gestire i tipi di fotocellula (catalogo)
Public Class MC_FrmGestisciLookupFotoc
    Inherits Form

    Private _db As MC_DatabaseService
    Private _lst As ListBox

    Public Sub New(tipo As String, db As MC_DatabaseService)
        _db = db
        Me.Text            = "Gestisci tipi fotocellula"
        Me.Size            = New Size(360, 380)
        Me.StartPosition   = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.Font            = New Font("Segoe UI", 9)

        _lst = New ListBox() With {.Location = New Point(16, 16), .Size = New Size(310, 260), .Font = New Font("Segoe UI", 10)}
        Me.Controls.Add(_lst)

        Dim txtNuovo  As New TextBox() With {.Location = New Point(16, 284), .Size = New Size(200, 24)}
        Dim btnAgg    As New Button() With {.Text = "+ Aggiungi", .Location = New Point(224, 283), .Size = New Size(100, 26)}
        Dim btnEl     As New Button() With {.Text = "Elimina selezionato", .Location = New Point(16, 316), .Size = New Size(170, 28)}
        Dim btnChiudi As New Button() With {.Text = "Chiudi", .Location = New Point(200, 316), .Size = New Size(80, 28), .DialogResult = DialogResult.OK}
        MC_PanelBuild.StyleButton(btnAgg, True) : MC_PanelBuild.StyleButton(btnEl, False) : MC_PanelBuild.StyleButton(btnChiudi, False)
        Me.Controls.AddRange({txtNuovo, btnAgg, btnEl, btnChiudi})
        Me.AcceptButton = btnChiudi

        AddHandler btnAgg.Click, Sub(s, e)
            Dim nome = txtNuovo.Text.Trim()
            If String.IsNullOrEmpty(nome) Then Return
            Try : _db.SalvaTipoFotocellula(nome) : txtNuovo.Clear() : Ricarica()
            Catch ex As Exception : MessageBox.Show("Errore: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
        AddHandler btnEl.Click, Sub(s, e)
            If _lst.SelectedItem Is Nothing Then Return
            If MessageBox.Show($"Eliminare '{_lst.SelectedItem}'?", "Conferma",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Warning) <> DialogResult.Yes Then Return
            Try : _db.EliminaTipoFotocellula(CInt(_lst.SelectedValue)) : Ricarica()
            Catch ex As Exception : MessageBox.Show("Errore: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
        Ricarica()
    End Sub

    Private Sub Ricarica()
        _lst.DataSource    = _db.GetTipiFotocellula()
        _lst.DisplayMember = "Nome"
        _lst.ValueMember   = "ID"
    End Sub
End Class

' ══════════════════════════════════════════════
' Form per creare/modificare una voce del catalogo fotocellule
Public Class MC_FrmEditCatalogoFotocellula
    Inherits Form

    Private _cat As MC_CatalogoFotocellula
    Private _db  As MC_DatabaseService

    Public Sub New(cat As MC_CatalogoFotocellula, db As MC_DatabaseService)
        _cat = If(cat IsNot Nothing, cat, New MC_CatalogoFotocellula())
        _db  = db
        Me.Text            = If(cat IsNot Nothing AndAlso cat.ID > 0, "Modifica fotocellula catalogo", "Nuova fotocellula catalogo")
        Me.Size            = New Size(560, 440)
        Me.StartPosition   = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.Font            = New Font("Segoe UI", 9)

        Dim y = 20

        ' Codice
        Me.Controls.Add(New Label() With {.Text = "Codice*:", .Location = New Point(16, y + 3), .Size = New Size(100, 20)})
        Dim txtCodice As New TextBox() With {.Text = _cat.Codice, .Location = New Point(120, y), .Size = New Size(300, 24)}
        Me.Controls.Add(txtCodice)
        y += 36

        ' Tipo (dropdown)
        Me.Controls.Add(New Label() With {.Text = "Tipo*:", .Location = New Point(16, y + 3), .Size = New Size(100, 20)})
        Dim cmbTipo As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Location = New Point(120, y), .Size = New Size(240, 24)}
        Dim btnGestTipi As New Button() With {.Text = "Gestisci...", .Location = New Point(368, y), .Size = New Size(90, 24)}
        MC_PanelBuild.StyleButton(btnGestTipi, False)
        Me.Controls.AddRange({cmbTipo, btnGestTipi})
        y += 36

        Dim RicaricaTipi As Action = Sub()
            Dim tipi = _db.GetTipiFotocellula()
            cmbTipo.DataSource    = tipi
            cmbTipo.DisplayMember = "Nome"
            cmbTipo.ValueMember   = "ID"
            Dim sel = tipi.FirstOrDefault(Function(t) t.ID = _cat.TipoID)
            If sel IsNot Nothing Then cmbTipo.SelectedItem = sel
        End Sub
        Try : RicaricaTipi() : Catch : End Try

        AddHandler btnGestTipi.Click, Sub(s, e)
            Using f As New MC_FrmGestisciLookupFotoc("TipiFotocellule", _db) : f.ShowDialog(Me) : End Using
            RicaricaTipi()
        End Sub

        ' Immagine
        Me.Controls.Add(New Label() With {.Text = "Immagine:", .Location = New Point(16, y + 3), .Size = New Size(100, 20)})
        Dim txtPath As New TextBox() With {.Text = _cat.PathImmagine, .Location = New Point(120, y), .Size = New Size(240, 24), .ReadOnly = True}
        Dim btnSfoglia As New Button() With {.Text = "Sfoglia...", .Location = New Point(368, y), .Size = New Size(90, 24)}
        MC_PanelBuild.StyleButton(btnSfoglia, False)
        Me.Controls.AddRange({txtPath, btnSfoglia})
        y += 36

        ' Anteprima immagine
        Dim pic As New PictureBox() With {
            .Location = New Point(120, y), .Size = New Size(330, 200),
            .SizeMode = PictureBoxSizeMode.Zoom, .BackColor = Color.FromArgb(245, 245, 242),
            .BorderStyle = BorderStyle.FixedSingle
        }
        Me.Controls.Add(pic)
        If Not String.IsNullOrEmpty(_cat.PathImmagine) AndAlso IO.File.Exists(_cat.PathImmagine) Then
            Try : pic.Image = Image.FromFile(_cat.PathImmagine) : Catch : End Try
        End If
        y += 210

        AddHandler btnSfoglia.Click, Sub(s, e)
            Using ofd As New OpenFileDialog() With {.Filter = "Immagini|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif|Tutti|*.*"}
                If ofd.ShowDialog() = DialogResult.OK Then
                    txtPath.Text = ofd.FileName
                    Try : pic.Image = Image.FromFile(ofd.FileName) : Catch : End Try
                End If
            End Using
        End Sub

        ' Pulsanti salva/annulla
        Dim btnOk  As New Button() With {.Text = "Salva",   .Location = New Point(120, y), .Size = New Size(100, 32)}
        Dim btnAnn As New Button() With {.Text = "Annulla", .Location = New Point(228, y), .Size = New Size(90, 32), .DialogResult = DialogResult.Cancel}
        MC_PanelBuild.StyleButton(btnOk, True) : MC_PanelBuild.StyleButton(btnAnn, False)
        Me.Controls.AddRange({btnOk, btnAnn})
        Me.CancelButton = btnAnn

        AddHandler btnOk.Click, Sub(s, e)
            If String.IsNullOrWhiteSpace(txtCodice.Text) Then
                MessageBox.Show("Il codice è obbligatorio.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Return
            End If
            If cmbTipo.SelectedItem Is Nothing Then
                MessageBox.Show("Seleziona un tipo.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Return
            End If
            _cat.Codice       = txtCodice.Text.Trim()
            _cat.TipoID       = DirectCast(cmbTipo.SelectedItem, MC_TipoFotocellula).ID
            _cat.PathImmagine = txtPath.Text.Trim()
            Try
                _cat.ID = _db.SalvaCatalogoFotocellula(_cat)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Catch ex As Exception
                MessageBox.Show("Errore: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
    End Sub
End Class

' ══════════════════════════════════════════════
' Picker per selezionare una fotocellula dal catalogo
Public Class MC_FrmSelezionaFotocellula
    Inherits Form

    Public ReadOnly Property Selezionata As MC_CatalogoFotocellula

    Public Sub New(db As MC_DatabaseService)
        Me.Text            = "Seleziona fotocellula dal catalogo"
        Me.Size            = New Size(760, 520)
        Me.StartPosition   = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox     = False
        Me.Font            = New Font("Segoe UI", 9)

        ' Filtro
        Dim pnlTop As New Panel() With {.Dock = DockStyle.Top, .Height = 40}
        Dim lblF As New Label() With {.Text = "Cerca:", .AutoSize = True, .Location = New Point(8, 11)}
        Dim txtF As New TextBox() With {.Location = New Point(52, 8), .Size = New Size(260, 24)}
        pnlTop.Controls.AddRange({lblF, txtF})
        Me.Controls.Add(pnlTop)

        ' Layout split: griglia sx + preview dx
        Dim pnlSplit As New Panel() With {.Dock = DockStyle.Fill}

        Dim pnlPrev As New Panel() With {.Dock = DockStyle.Right, .Width = 240, .BackColor = Color.FromArgb(245, 245, 242), .Padding = New Padding(8)}
        Dim pic As New PictureBox() With {.Dock = DockStyle.Fill, .SizeMode = PictureBoxSizeMode.Zoom, .BackColor = Color.FromArgb(245, 245, 242)}
        Dim lblPrev As New Label() With {.Dock = DockStyle.Bottom, .Height = 36, .Font = New Font("Segoe UI", 8.5F), .TextAlign = ContentAlignment.MiddleCenter, .ForeColor = Color.FromArgb(60, 60, 60)}
        pnlPrev.Controls.AddRange({pic, lblPrev})

        Dim dgv As New DataGridView() With {
            .Dock = DockStyle.Fill, .AllowUserToAddRows = False, .ReadOnly = True,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .RowHeadersVisible = False, .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle, .Font = New Font("Segoe UI", 9.5F)
        }
        Dim BLUE_LIGHT As Color = Color.FromArgb(230, 241, 251)
        Dim BLUE_DARK  As Color = Color.FromArgb(24, 95, 165)
        dgv.ColumnHeadersDefaultCellStyle.BackColor = BLUE_LIGHT
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = BLUE_DARK
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI Semibold", 9)
        dgv.EnableHeadersVisualStyles = False

        pnlSplit.Controls.Add(dgv)
        pnlSplit.Controls.Add(pnlPrev)

        ' Pulsanti
        Dim pnlBot As New Panel() With {.Dock = DockStyle.Bottom, .Height = 46}
        Dim btnOk  As New Button() With {.Text = "Seleziona", .Location = New Point(8, 8),   .Size = New Size(120, 30)}
        Dim btnAnn As New Button() With {.Text = "Annulla",   .Location = New Point(136, 8), .Size = New Size(90, 30), .DialogResult = DialogResult.Cancel}
        MC_PanelBuild.StyleButton(btnOk, True) : MC_PanelBuild.StyleButton(btnAnn, False)
        pnlBot.Controls.AddRange({btnOk, btnAnn})

        Me.Controls.Add(pnlSplit)
        Me.Controls.Add(pnlBot)
        Me.CancelButton = btnAnn

        Dim Carica As Action = Sub()
            dgv.Rows.Clear() : dgv.Columns.Clear()
            For Each h In {"Codice", "Tipo"} : dgv.Columns.Add(h, h) : Next
            Try
                For Each c In db.GetCatalogoFotocellule(txtF.Text.Trim())
                    dgv.Rows.Add(c.Codice, c.TipoNome)
                    dgv.Rows(dgv.Rows.Count - 1).Tag = c
                Next
            Catch : End Try
        End Sub

        AddHandler dgv.SelectionChanged, Sub(s, e)
            If dgv.SelectedRows.Count = 0 Then pic.Image = Nothing : lblPrev.Text = "" : Return
            Dim cat = TryCast(dgv.SelectedRows(0).Tag, MC_CatalogoFotocellula) : If cat Is Nothing Then Return
            lblPrev.Text = $"{cat.Codice}{vbLf}{cat.TipoNome}"
            If Not String.IsNullOrEmpty(cat.PathImmagine) AndAlso IO.File.Exists(cat.PathImmagine) Then
                Try : pic.Image = Image.FromFile(cat.PathImmagine) : Catch : pic.Image = Nothing : End Try
            Else
                pic.Image = Nothing
            End If
        End Sub

        AddHandler txtF.TextChanged, Sub(s, e) Carica()

        AddHandler btnOk.Click, Sub(s, e)
            If dgv.SelectedRows.Count = 0 Then Return
            _Selezionata = TryCast(dgv.SelectedRows(0).Tag, MC_CatalogoFotocellula)
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub

        AddHandler dgv.CellDoubleClick, Sub(s, e)
            If dgv.SelectedRows.Count = 0 Then Return
            _Selezionata = TryCast(dgv.SelectedRows(0).Tag, MC_CatalogoFotocellula)
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End Sub

        Carica()
    End Sub
End Class

' ══════════════════════════════════════════════
Public Class MC_FrmImpostazioni
    Inherits Form

    Public Sub New()
        Me.Text = "Impostazioni ManualCraft"
        Me.Size = New Size(560, 200)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Font = New Font("Segoe UI", 9)

        Me.Controls.Add(New Label() With {
            .Text = "Connessione DB: usa Homepage.sap_tirelli (Tirelli_40)",
            .Location = New Point(16, 16), .AutoSize = True, .ForeColor = Color.Gray
        })

        Me.Controls.Add(New Label() With {.Text = "Chiave API Anthropic:", .Location = New Point(16, 44), .AutoSize = True})
        Dim keyPath = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "anthropic_key.txt")
        Dim currentKey = If(IO.File.Exists(keyPath), IO.File.ReadAllText(keyPath).Trim(), "")
        Dim txtApi As New TextBox() With {.Text = currentKey, .Location = New Point(16, 64), .Size = New Size(510, 24), .PasswordChar = "*"c}
        Dim chkShow As New CheckBox() With {.Text = "Mostra chiave", .Location = New Point(16, 94), .AutoSize = True}
        AddHandler chkShow.CheckedChanged, Sub(s, e) txtApi.PasswordChar = If(chkShow.Checked, Nothing, "*"c)
        Me.Controls.AddRange({txtApi, chkShow})

        Dim btnSalva As New Button() With {.Text = "Salva", .Location = New Point(16, 124), .Size = New Size(100, 32)}
        MC_PanelBuild.StyleButton(btnSalva, True)
        AddHandler btnSalva.Click, Sub(s, e)
            Try
                IO.File.WriteAllText(keyPath, txtApi.Text.Trim())
                MessageBox.Show("Chiave API salvata.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Close()
            Catch ex As Exception
                MessageBox.Show("Errore: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Dim btnAnn As New Button() With {.Text = "Annulla", .Location = New Point(126, 124), .Size = New Size(80, 32), .DialogResult = DialogResult.Cancel}
        MC_PanelBuild.StyleButton(btnAnn, False)
        Me.Controls.AddRange({btnSalva, btnAnn})
        Me.CancelButton = btnAnn
    End Sub
End Class
