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
Public Class MC_FrmEditFotocellula
    Inherits Form

    Private _fc As MC_Fotocellula
    Private _db As MC_DatabaseService
    Private _macchinaID As Integer

    Public Sub New(fc As MC_Fotocellula, macchinaID As Integer, db As MC_DatabaseService)
        _fc = If(fc IsNot Nothing, fc, New MC_Fotocellula() With {.MacchinaID = macchinaID})
        _macchinaID = macchinaID
        _db = db
        Me.Text = If(fc IsNot Nothing, "Modifica fotocellula", "Nuova fotocellula")
        Me.Size = New Size(520, 520)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Font = New Font("Segoe UI", 9)

        Dim fields As New Dictionary(Of String, TextBox)
        Dim defs As (String, String)() = {
            ("Codice* (es. FC-01):", _fc.Codice),
            ("Marca:", _fc.Marca),
            ("Modello:", _fc.Modello),
            ("Posizione:", _fc.Posizione),
            ("Tensione lavoro:", _fc.TensioneLavoro),
            ("Uscita logica:", _fc.UscitaLogica),
            ("Distanza rilevazione:", _fc.DistanzaRilev)
        }
        Dim y = 20
        For Each d In defs
            Me.Controls.Add(New Label() With {.Text = d.Item1, .Location = New Point(16, y + 3), .Size = New Size(160, 20)})
            Dim txt As New TextBox() With {.Text = d.Item2, .Location = New Point(180, y), .Size = New Size(290, 24)}
            Me.Controls.Add(txt)
            fields(d.Item1) = txt
            y += 32
        Next

        Me.Controls.Add(New Label() With {.Text = "Tipo rilevazione:", .Location = New Point(16, y + 3), .Size = New Size(160, 20)})
        Dim cmbTipo As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Location = New Point(180, y), .Size = New Size(180, 24)}
        For Each t In {"Barriera", "Riflessione", "Prossimità", "Laser", "A fibra ottica"} : cmbTipo.Items.Add(t) : Next
        cmbTipo.SelectedItem = If(String.IsNullOrEmpty(_fc.TipoRilevazione), "Barriera", _fc.TipoRilevazione)
        Me.Controls.Add(cmbTipo)
        y += 32

        Me.Controls.Add(New Label() With {.Text = "Note installazione:", .Location = New Point(16, y + 3), .Size = New Size(160, 20)})
        Dim txtNote As New TextBox() With {.Text = _fc.NoteInstallaz, .Location = New Point(180, y), .Size = New Size(290, 60), .Multiline = True}
        Me.Controls.Add(txtNote)
        y += 72

        Dim btnOk  As New Button() With {.Text = "Salva",   .Location = New Point(180, y), .Size = New Size(100, 32)}
        Dim btnAnn As New Button() With {.Text = "Annulla", .Location = New Point(290, y), .Size = New Size(100, 32), .DialogResult = DialogResult.Cancel}
        MC_PanelBuild.StyleButton(btnOk, True)
        MC_PanelBuild.StyleButton(btnAnn, False)
        Me.Controls.AddRange({btnOk, btnAnn})
        Me.CancelButton = btnAnn

        AddHandler btnOk.Click, Sub(s, e)
            _fc.MacchinaID      = _macchinaID
            _fc.Codice          = fields("Codice* (es. FC-01):").Text.Trim()
            _fc.Marca           = fields("Marca:").Text.Trim()
            _fc.Modello         = fields("Modello:").Text.Trim()
            _fc.Posizione       = fields("Posizione:").Text.Trim()
            _fc.TensioneLavoro  = fields("Tensione lavoro:").Text.Trim()
            _fc.UscitaLogica    = fields("Uscita logica:").Text.Trim()
            _fc.DistanzaRilev   = fields("Distanza rilevazione:").Text.Trim()
            _fc.TipoRilevazione = cmbTipo.SelectedItem?.ToString()
            _fc.NoteInstallaz   = txtNote.Text.Trim()
            If String.IsNullOrEmpty(_fc.Codice) Then
                MessageBox.Show("Il codice è obbligatorio.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Try
                _db.SalvaFotocellula(_fc)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Catch ex As Exception
                MessageBox.Show("Errore: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
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
