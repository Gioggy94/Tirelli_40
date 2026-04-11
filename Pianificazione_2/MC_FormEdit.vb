Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Public Class MC_FrmEditMacchina
    Inherits Form

    Private _m As MC_Macchina
    Private _db As MC_DatabaseService
    Private _fields As New Dictionary(Of String, Control)

    Public Sub New(m As MC_Macchina, db As MC_DatabaseService)
        _m = If(m IsNot Nothing, m, New MC_Macchina())
        _db = db
        Me.Text = If(m IsNot Nothing, "Modifica macchina", "Nuova macchina")
        Me.Size = New Size(500, 460)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Font = New Font("Segoe UI", 9)

        Dim y = 20
        Dim defs As (String, String)() = {
            ("Matricola*:", _m.Matricola),
            ("Nome macchina*:", _m.NomeMacchina),
            ("Modello*:", _m.Modello),
            ("Tipo macchina:", _m.TipoMacchina),
            ("Cliente finale:", _m.ClienteFinale),
            ("Anno costruzione:", If(_m.AnnoCostruzione.HasValue, _m.AnnoCostruzione.Value.ToString(), ""))
        }
        For Each d In defs
            Me.Controls.Add(New Label() With {.Text = d.Item1, .Location = New Point(20, y + 3), .Size = New Size(140, 20)})
            Dim txt As New TextBox() With {.Text = d.Item2, .Location = New Point(165, y), .Size = New Size(295, 24)}
            Me.Controls.Add(txt)
            _fields(d.Item1) = txt
            y += 34
        Next

        Me.Controls.Add(New Label() With {.Text = "Lingua:", .Location = New Point(20, y + 3), .Size = New Size(140, 20)})
        Dim cmbLng As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Location = New Point(165, y), .Size = New Size(120, 24)}
        For Each l In {"IT", "EN", "FR", "ES", "DE"} : cmbLng.Items.Add(l) : Next
        cmbLng.SelectedItem = If(String.IsNullOrEmpty(_m.LinguaCodice), "IT", _m.LinguaCodice)
        Me.Controls.Add(cmbLng)
        _fields("Lingua:") = cmbLng
        y += 34

        Dim chkAtt As New CheckBox() With {.Text = "Macchina attiva", .Checked = _m.Attiva, .Location = New Point(165, y), .AutoSize = True}
        Me.Controls.Add(chkAtt)
        y += 40

        Dim btnOk  As New Button() With {.Text = "Salva",   .Location = New Point(165, y), .Size = New Size(100, 32)}
        Dim btnAnn As New Button() With {.Text = "Annulla", .Location = New Point(275, y), .Size = New Size(100, 32), .DialogResult = DialogResult.Cancel}
        MC_PanelBuild.StyleButton(btnOk, True)
        MC_PanelBuild.StyleButton(btnAnn, False)
        Me.Controls.AddRange({btnOk, btnAnn})
        Me.CancelButton = btnAnn

        AddHandler btnOk.Click, Sub(s, e)
            _m.Matricola     = _fields("Matricola*:").Text.Trim()
            _m.NomeMacchina  = _fields("Nome macchina*:").Text.Trim()
            _m.Modello       = _fields("Modello*:").Text.Trim()
            _m.TipoMacchina  = _fields("Tipo macchina:").Text.Trim()
            _m.ClienteFinale = _fields("Cliente finale:").Text.Trim()
            Dim anno As Integer
            If Integer.TryParse(_fields("Anno costruzione:").Text.Trim(), anno) Then _m.AnnoCostruzione = anno
            _m.LinguaCodice = cmbLng.SelectedItem?.ToString()
            _m.Attiva       = chkAtt.Checked
            If String.IsNullOrEmpty(_m.Matricola) OrElse String.IsNullOrEmpty(_m.NomeMacchina) Then
                MessageBox.Show("Matricola e Nome macchina sono obbligatori.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Try
                _db.SalvaMacchina(_m)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Catch ex As Exception
                MessageBox.Show("Errore salvataggio: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub
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
