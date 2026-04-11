Imports System.Drawing
Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Module MC_PanelBuild

    Private ReadOnly BLUE_DARK As Color = Color.FromArgb(24, 95, 165)
    Private ReadOnly BLUE_LIGHT As Color = Color.FromArgb(230, 241, 251)
    Private ReadOnly BORDER_C As Color = Color.FromArgb(220, 220, 220)
    Private ReadOnly FONT_TITLE As New Font("Segoe UI Semibold", 11)
    Private ReadOnly FONT_LABEL As New Font("Segoe UI", 8.5F)
    Private ReadOnly FONT_BODY As New Font("Segoe UI", 9.5F)

    ' ════════════════════════════════════════════
    ' PANEL HOME
    ' ════════════════════════════════════════════

    Public Function BuildPanelHome(owner As MC_FrmMain, db As MC_DatabaseService) As Panel
        Dim pnl As New Panel()
        Dim lblTitle As New Label() With {
            .Text = "Home macchina", .Font = New Font("Segoe UI Semibold", 16),
            .ForeColor = Color.FromArgb(40, 40, 40), .AutoSize = True, .Location = New Point(0, 0)
        }
        Dim card As Panel = BuildCard(0, 40, 600, 370, "Selezione macchina")

        ' ── Ricerca per matricola ──────────────────────────────────────
        Dim lblMat As New Label() With {.Text = "Codice matricola:", .Location = New Point(16, 50), .AutoSize = True, .Font = FONT_LABEL}
        Dim txtMat As New TextBox() With {.Location = New Point(160, 47), .Size = New Size(180, 24), .Font = FONT_BODY}
        Dim btnCerca As New Button() With {.Text = "Cerca", .Location = New Point(350, 46), .Size = New Size(80, 26), .Font = FONT_BODY}
        StyleButton(btnCerca, True)

        ' ── Campi risultato ───────────────────────────────────────────
        Dim y = 90
        Dim txtNome   As New TextBox() With {.Location = New Point(160, y),      .Size = New Size(380, 24), .Font = FONT_BODY, .ReadOnly = True, .BackColor = Color.FromArgb(245, 245, 242)}
        Dim txtCliente As New TextBox() With {.Location = New Point(160, y + 64), .Size = New Size(380, 24), .Font = FONT_BODY, .ReadOnly = True, .BackColor = Color.FromArgb(245, 245, 242)}
        Dim txtLingua As New TextBox() With {.Location = New Point(160, y + 96),  .Size = New Size(100, 24), .Font = FONT_BODY, .ReadOnly = True, .BackColor = Color.FromArgb(245, 245, 242)}

        ' Modello (ComboBox)
        Dim cmbModello As New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Location = New Point(160, y + 32), .Size = New Size(280, 24), .Font = FONT_BODY
        }
        ' Tipo macchina (ComboBox)
        Dim cmbTipo As New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Location = New Point(160, y + 64), .Size = New Size(280, 24), .Font = FONT_BODY
        }

        ' Etichette
        For Each pair In {("Nome macchina:", y), ("Modello:", y + 32), ("Tipo macchina:", y + 64), ("Cliente finale:", y + 96), ("Lingua:", y + 128)}
            card.Controls.Add(New Label() With {.Text = pair.Item1, .Location = New Point(16, pair.Item2 + 3), .Size = New Size(140, 20), .Font = FONT_LABEL})
        Next
        txtCliente = New TextBox() With {.Location = New Point(160, y + 96),  .Size = New Size(380, 24), .Font = FONT_BODY, .ReadOnly = True, .BackColor = Color.FromArgb(245, 245, 242)}
        txtLingua  = New TextBox() With {.Location = New Point(160, y + 128), .Size = New Size(100, 24), .Font = FONT_BODY, .ReadOnly = True, .BackColor = Color.FromArgb(245, 245, 242)}

        Dim btnSeleziona As New Button() With {.Text = "Imposta come attiva", .Location = New Point(16, y + 168), .Size = New Size(170, 32), .Font = FONT_BODY}
        StyleButton(btnSeleziona, True)

        ' Carica voci dropdown
        Dim CaricaDropdown As Action = Sub()
            Dim modelli = db.GetModelli()
            Dim tipi    = db.GetTipiMacchina()
            cmbModello.DataSource    = modelli   : cmbModello.DisplayMember = "Nome" : cmbModello.ValueMember = "ID"
            cmbTipo.DataSource       = tipi      : cmbTipo.DisplayMember    = "Nome" : cmbTipo.ValueMember    = "ID"
        End Sub
        Try : CaricaDropdown() : Catch : End Try

        AddHandler btnCerca.Click, Sub(s, e)
            Try
                CaricaDropdown()
                Dim m = db.GetMacchinaByMatricola(txtMat.Text.Trim())
                If m IsNot Nothing Then
                    txtNome.Text    = m.NomeMacchina
                    txtCliente.Text = m.ClienteFinale
                    txtLingua.Text  = m.LinguaCodice
                    Dim selMod = DirectCast(cmbModello.DataSource, List(Of MC_Modello)).FirstOrDefault(Function(x) x.Nome = m.Modello)
                    If selMod IsNot Nothing Then cmbModello.SelectedItem = selMod
                    Dim selTipo = DirectCast(cmbTipo.DataSource, List(Of MC_TipoMacchina)).FirstOrDefault(Function(x) x.Nome = m.TipoMacchina)
                    If selTipo IsNot Nothing Then cmbTipo.SelectedItem = selTipo
                    card.Tag = m
                Else
                    MessageBox.Show("Matricola non trovata.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch ex As Exception
                MessageBox.Show("Errore DB: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        AddHandler btnSeleziona.Click, Sub(s, e)
            If card.Tag Is Nothing Then Return
            Dim m = DirectCast(card.Tag, MC_Macchina)
            m.Modello      = If(cmbModello.SelectedItem IsNot Nothing, DirectCast(cmbModello.SelectedItem, MC_Modello).Nome, "")
            m.TipoMacchina = If(cmbTipo.SelectedItem IsNot Nothing, DirectCast(cmbTipo.SelectedItem, MC_TipoMacchina).Nome, "")
            m.LinguaCodice = txtLingua.Text.Trim()
            Try
                m.ID = db.SalvaExtraMacchina(m)
                owner.SetMacchinaCorrente(m)
                MessageBox.Show("Macchina impostata come attiva.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Errore salvataggio: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        card.Controls.AddRange({lblMat, txtMat, btnCerca, txtNome, cmbModello, cmbTipo, txtCliente, txtLingua, btnSeleziona})
        pnl.Controls.AddRange({lblTitle, card})

        AddHandler pnl.SizeChanged, Sub(s, e)
            Dim cx = Math.Max(24, (pnl.ClientSize.Width - card.Width) \ 2)
            card.Left = cx
            lblTitle.Left = cx
        End Sub

        Return pnl
    End Function

    ' ════════════════════════════════════════════
    ' PANEL ANAGRAFICA MACCHINE
    ' ════════════════════════════════════════════

    Public Function BuildPanelMacchine(owner As MC_FrmMain, db As MC_DatabaseService) As Panel
        Dim pnl As New Panel()

        ' ── Titolo ───────────────────────────────────────────────────────
        Dim pnlTitle As New Panel() With {.Dock = DockStyle.Top, .Height = 50}
        pnlTitle.Controls.Add(New Label() With {
            .Text = "Anagrafica macchine", .Font = New Font("Segoe UI Semibold", 16),
            .ForeColor = Color.FromArgb(40, 40, 40), .AutoSize = True, .Location = New Point(0, 8)
        })

        ' ── Barra di ricerca ─────────────────────────────────────────────
        Dim pnlSearch As New Panel() With {.Dock = DockStyle.Top, .Height = 42}
        Dim lblMat As New Label() With {.Text = "Matricola:", .AutoSize = True, .Location = New Point(0, 11), .Font = FONT_LABEL}
        Dim txtMat As New TextBox() With {.Location = New Point(68, 8), .Size = New Size(150, 24), .Font = FONT_BODY}
        Dim lblCli As New Label() With {.Text = "Cliente:", .AutoSize = True, .Location = New Point(228, 11), .Font = FONT_LABEL}
        Dim txtCli As New TextBox() With {.Location = New Point(278, 8), .Size = New Size(200, 24), .Font = FONT_BODY}
        Dim btnCerca As New Button() With {.Text = "Cerca", .Location = New Point(488, 7), .Size = New Size(80, 26), .Font = FONT_BODY}
        StyleButton(btnCerca, True)
        pnlSearch.Controls.AddRange({lblMat, txtMat, lblCli, txtCli, btnCerca})

        ' ── Pulsanti azione ──────────────────────────────────────────────
        Dim pnlBtns As New Panel() With {.Dock = DockStyle.Bottom, .Height = 48}
        Dim btnAssoc   As New Button() With {.Text = "Associa dati manuale", .Location = New Point(0, 8), .Size = New Size(170, 32), .Font = FONT_BODY}
        Dim btnImposta As New Button() With {.Text = "Imposta come attiva",  .Location = New Point(180, 8), .Size = New Size(160, 32), .Font = FONT_BODY}
        Dim btnGestMod As New Button() With {.Text = "Modelli...",           .Location = New Point(360, 8), .Size = New Size(100, 32), .Font = FONT_BODY}
        Dim btnGestTip As New Button() With {.Text = "Tipi macchina...",     .Location = New Point(470, 8), .Size = New Size(130, 32), .Font = FONT_BODY}
        StyleButton(btnAssoc, True) : StyleButton(btnImposta, True)
        StyleButton(btnGestMod, False) : StyleButton(btnGestTip, False)
        pnlBtns.Controls.AddRange({btnAssoc, btnImposta, btnGestMod, btnGestTip})

        ' ── Griglia ──────────────────────────────────────────────────────
        Dim dgv As New DataGridView() With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False, .ReadOnly = True,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .RowHeadersVisible = False, .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle, .Font = FONT_BODY, .Name = "dgvMacchine"
        }
        dgv.ColumnHeadersDefaultCellStyle.BackColor = BLUE_LIGHT
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = BLUE_DARK
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI Semibold", 9)
        dgv.EnableHeadersVisualStyles = False

        ' ── Azione di ricerca (riusata da più handler) ───────────────────
        Dim Cerca As Action = Sub()
            dgv.Rows.Clear() : dgv.Columns.Clear()
            For Each col In {"Matricola", "Nome macchina", "Cliente", "Modello", "Tipo macchina", "Lingua"}
                dgv.Columns.Add(col, col)
            Next
            Try
                For Each m In db.GetMacchineAS400(txtMat.Text.Trim(), txtCli.Text.Trim())
                    dgv.Rows.Add(m.Matricola, m.NomeMacchina, m.ClienteFinale, m.Modello, m.TipoMacchina, m.LinguaCodice)
                    dgv.Rows(dgv.Rows.Count - 1).Tag = m
                Next
            Catch ex As Exception
                MessageBox.Show("Errore ricerca AS400: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        AddHandler btnCerca.Click, Sub(s, e) Cerca()
        AddHandler txtMat.KeyDown, Sub(s, e) If e.KeyCode = Keys.Return Then Cerca()
        AddHandler txtCli.KeyDown, Sub(s, e) If e.KeyCode = Keys.Return Then Cerca()

        AddHandler btnAssoc.Click, Sub(s, e)
            Dim m = GetSelectedMacchina(dgv) : If m Is Nothing Then Return
            Using f As New MC_FrmEditMacchina(m, db)
                If f.ShowDialog(owner) = DialogResult.OK Then Cerca()
            End Using
        End Sub

        AddHandler btnImposta.Click, Sub(s, e)
            Dim m = GetSelectedMacchina(dgv) : If m Is Nothing Then Return
            If m.ID = 0 Then m.ID = db.SalvaExtraMacchina(m)
            owner.SetMacchinaCorrente(m)
            MessageBox.Show($"Macchina '{m.NomeMacchina}' ({m.Matricola}) impostata come attiva.",
                            "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        AddHandler btnGestMod.Click, Sub(s, e)
            Using f As New MC_FrmGestisciLookup("Modelli", db)
                f.ShowDialog(owner)
            End Using
        End Sub
        AddHandler btnGestTip.Click, Sub(s, e)
            Using f As New MC_FrmGestisciLookup("TipiMacchina", db)
                f.ShowDialog(owner)
            End Using
        End Sub

        ' Ordine aggiunta: Fill prima, poi Bottom, poi Top (ordine z per docking corretto)
        pnl.Controls.Add(dgv)
        pnl.Controls.Add(pnlBtns)
        pnl.Controls.Add(pnlSearch)
        pnl.Controls.Add(pnlTitle)
        Return pnl
    End Function

    ' La ricerca è interna al pannello — questa sub rimane per compatibilità con la navigazione
    Public Sub RicaricaMacchine(pnl As Panel, db As MC_DatabaseService)
    End Sub

    Private Function GetSelectedMacchina(dgv As DataGridView) As MC_Macchina
        If dgv.SelectedRows.Count = 0 Then
            MessageBox.Show("Seleziona una riga.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return Nothing
        End If
        Return TryCast(dgv.SelectedRows(0).Tag, MC_Macchina)
    End Function

    ' ════════════════════════════════════════════
    ' PANEL FOTOCELLULE
    ' ════════════════════════════════════════════

    Public Function BuildPanelFotocellule(owner As MC_FrmMain, db As MC_DatabaseService, ai As MC_AnthropicService) As Panel
        Dim pnl As New Panel()
        Dim lblTitle As New Label() With {.Text = "Fotocellule — cap. 5.1", .Font = New Font("Segoe UI Semibold", 16), .ForeColor = Color.FromArgb(40, 40, 40), .AutoSize = True, .Location = New Point(0, 0)}
        Dim lblSub As New Label() With {.Text = "Anagrafica fotocellule. Popola automaticamente il capitolo 5.1 del manuale.", .Font = FONT_LABEL, .ForeColor = Color.Gray, .AutoSize = True, .Location = New Point(0, 30)}

        Dim dgv As New DataGridView() With {
            .Location = New Point(0, 58), .Size = New Size(900, 380),
            .AllowUserToAddRows = False, .ReadOnly = True,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .RowHeadersVisible = False, .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle, .Font = FONT_BODY, .Name = "dgvFotoc"
        }
        dgv.ColumnHeadersDefaultCellStyle.BackColor = BLUE_LIGHT
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = BLUE_DARK
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI Semibold", 9)
        dgv.EnableHeadersVisualStyles = False

        Dim btnNuova As New Button() With {.Text = "+ Aggiungi", .Location = New Point(0, 450), .Size = New Size(120, 32), .Font = FONT_BODY}
        Dim btnModifica As New Button() With {.Text = "Modifica", .Location = New Point(130, 450), .Size = New Size(100, 32), .Font = FONT_BODY}
        Dim btnElimina As New Button() With {.Text = "Elimina", .Location = New Point(240, 450), .Size = New Size(100, 32), .Font = FONT_BODY}
        Dim btnGenCap As New Button() With {.Text = "Genera cap. 5.1 ✦", .Location = New Point(600, 450), .Size = New Size(160, 32), .Font = FONT_BODY}
        StyleButton(btnNuova, True) : StyleButton(btnModifica, False) : StyleButton(btnElimina, False) : StyleButton(btnGenCap, True)

        AddHandler btnNuova.Click, Sub(s, e)
                                       Dim m = owner.GetMacchinaCorrente() : If m Is Nothing Then Return
                                       Using f As New MC_FrmEditFotocellula(Nothing, m.ID, db)
                                           If f.ShowDialog(owner) = DialogResult.OK Then RicaricaFotocellule(pnl, db, m)
                                       End Using
                                   End Sub
        AddHandler btnModifica.Click, Sub(s, e)
                                          If dgv.SelectedRows.Count = 0 Then Return
                                          Dim fc = TryCast(dgv.SelectedRows(0).Tag, MC_Fotocellula) : If fc Is Nothing Then Return
                                          Using f As New MC_FrmEditFotocellula(fc, fc.MacchinaID, db)
                                              If f.ShowDialog(owner) = DialogResult.OK Then RicaricaFotocellule(pnl, db, owner.GetMacchinaCorrente())
                                          End Using
                                      End Sub
        AddHandler btnElimina.Click, Sub(s, e)
                                         If dgv.SelectedRows.Count = 0 Then Return
                                         Dim fc = TryCast(dgv.SelectedRows(0).Tag, MC_Fotocellula) : If fc Is Nothing Then Return
                                         If MessageBox.Show($"Eliminare '{fc.Codice}'?", "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
                                             db.EliminaFotocellula(fc.ID) : RicaricaFotocellule(pnl, db, owner.GetMacchinaCorrente())
                                         End If
                                     End Sub
        AddHandler btnGenCap.Click, Async Sub(s, e)
                                        Dim m = owner.GetMacchinaCorrente() : If m Is Nothing Then Return
                                        Dim lista = db.GetFotocellule(m.ID)
                                        If lista.Count = 0 Then MessageBox.Show("Nessuna fotocellula registrata.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information) : Return
                                        btnGenCap.Enabled = False : btnGenCap.Text = "Generazione..."
                                        Try
                                            Dim testo = Await ai.GeneraCapitoloFotocellule(lista, m, owner.GetLinguaSelezionata())
                                            Using f As New MC_FrmTestoGenerato("Cap. 5.1 – Fotocellule", testo)
                                                f.ShowDialog(owner)
                                            End Using
                                        Catch ex As Exception
                                            MessageBox.Show("Errore AI: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        Finally
                                            btnGenCap.Enabled = True : btnGenCap.Text = "Genera cap. 5.1 ✦"
                                        End Try
                                    End Sub

        pnl.Controls.AddRange({lblTitle, lblSub, dgv, btnNuova, btnModifica, btnElimina, btnGenCap})
        Return pnl
    End Function

    Public Sub RicaricaFotocellule(pnl As Panel, db As MC_DatabaseService, m As MC_Macchina)
        Dim dgv = TryCast(pnl.Controls.Find("dgvFotoc", True).FirstOrDefault(), DataGridView)
        If dgv Is Nothing Then Return
        dgv.Rows.Clear() : dgv.Columns.Clear()
        For Each col In {"Codice", "Marca", "Modello", "Tipo", "Posizione", "Tensione", "Uscita"}
            dgv.Columns.Add(col, col)
        Next
        If m Is Nothing Then Return
        For Each f In db.GetFotocellule(m.ID)
            dgv.Rows.Add(f.Codice, f.Marca, f.Modello, f.TipoRilevazione, f.Posizione, f.TensioneLavoro, f.UscitaLogica)
            dgv.Rows(dgv.Rows.Count - 1).Tag = f
        Next
    End Sub

    ' ════════════════════════════════════════════
    ' PANEL SOFTWARE PLC
    ' ════════════════════════════════════════════

    Public Function BuildPanelSoftware(owner As MC_FrmMain, db As MC_DatabaseService, ai As MC_AnthropicService) As Panel
        Dim pnl As New Panel()
        Dim lblTitle As New Label() With {.Text = "Analisi software PLC / HMI", .Font = New Font("Segoe UI Semibold", 16), .ForeColor = Color.FromArgb(40, 40, 40), .AutoSize = True, .Location = New Point(0, 0)}

        ' Card 1 - File sorgente
        Dim card1 = BuildCard(0, 50, 580, 210, "Carica file sorgente PLC")
        Dim lblInfo1 As New Label() With {.Text = "Carica il file sorgente (.txt, .zip, .gx3, ...). L'AI estrarrà automaticamente tutti i codici errore.", .Location = New Point(16, 46), .Size = New Size(540, 40), .Font = FONT_LABEL, .ForeColor = Color.Gray}
        Dim lblFile1 As New Label() With {.Text = "Nessun file selezionato", .Location = New Point(16, 100), .Size = New Size(440, 20), .Font = FONT_LABEL, .ForeColor = Color.Gray}
        Dim btnSfog1 As New Button() With {.Text = "Sfoglia...", .Location = New Point(16, 130), .Size = New Size(100, 30), .Font = FONT_BODY}
        Dim btnAnal1 As New Button() With {.Text = "Analizza con AI ✦", .Location = New Point(126, 130), .Size = New Size(160, 30), .Font = FONT_BODY}
        StyleButton(btnSfog1, False) : StyleButton(btnAnal1, True)

        Dim filePlcPath As String = ""
        AddHandler btnSfog1.Click, Sub(s, e)
                                       Using ofd As New OpenFileDialog() With {.Filter = "File PLC|*.txt;*.zip;*.gx3;*.mer;*.prj;*.7z|Tutti|*.*"}
                                           If ofd.ShowDialog() = DialogResult.OK Then filePlcPath = ofd.FileName : lblFile1.Text = Path.GetFileName(filePlcPath)
                                       End Using
                                   End Sub
        AddHandler btnAnal1.Click, Async Sub(s, e)
                                       Dim m = owner.GetMacchinaCorrente()
                                       If m Is Nothing OrElse String.IsNullOrEmpty(filePlcPath) Then MessageBox.Show("Seleziona prima un file.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information) : Return
                                       btnAnal1.Enabled = False : btnAnal1.Text = "Analisi in corso..."
                                       Try
                                           Dim contenuto = File.ReadAllText(filePlcPath)
                                           Dim errori = Await ai.AnalizzaSoftwarePLC(contenuto, m.NomeMacchina, owner.GetLinguaSelezionata())
                                           If MessageBox.Show($"Trovati {errori.Count} codici errore. Importarli?", "Risultato", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                                               For Each ce In errori : ce.MacchinaID = m.ID : db.SalvaCodiceErrore(ce) : Next
                                               MessageBox.Show($"{errori.Count} codici errore importati.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                           End If
                                       Catch ex As Exception
                                           MessageBox.Show("Errore: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                       Finally
                                           btnAnal1.Enabled = True : btnAnal1.Text = "Analizza con AI ✦"
                                       End Try
                                   End Sub
        card1.Controls.AddRange({lblInfo1, lblFile1, btnSfog1, btnAnal1})

        ' Card 2 - Screenshot HMI
        Dim card2 = BuildCard(0, 280, 580, 250, "Analisi screenshot schermata errore HMI")
        Dim lblInfo2 As New Label() With {.Text = "Carica uno screenshot della schermata errore. L'AI riconosce il codice e lo descrive.", .Location = New Point(16, 46), .Size = New Size(540, 40), .Font = FONT_LABEL, .ForeColor = Color.Gray}
        Dim picPreview As New PictureBox() With {.Location = New Point(16, 95), .Size = New Size(200, 110), .SizeMode = PictureBoxSizeMode.Zoom, .BorderStyle = BorderStyle.FixedSingle, .BackColor = Color.FromArgb(245, 245, 245)}
        Dim lblFile2 As New Label() With {.Text = "Nessuna immagine", .Location = New Point(226, 95), .Size = New Size(330, 20), .Font = FONT_LABEL, .ForeColor = Color.Gray}
        Dim btnSfog2 As New Button() With {.Text = "Sfoglia immagine...", .Location = New Point(226, 120), .Size = New Size(150, 30), .Font = FONT_BODY}
        Dim btnAnal2 As New Button() With {.Text = "Analizza con AI ✦", .Location = New Point(226, 160), .Size = New Size(160, 30), .Font = FONT_BODY}
        StyleButton(btnSfog2, False) : StyleButton(btnAnal2, True)

        Dim fileImgPath As String = ""
        AddHandler btnSfog2.Click, Sub(s, e)
                                       Using ofd As New OpenFileDialog() With {.Filter = "Immagini|*.png;*.jpg;*.jpeg;*.bmp"}
                                           If ofd.ShowDialog() = DialogResult.OK Then
                                               fileImgPath = ofd.FileName : lblFile2.Text = Path.GetFileName(fileImgPath)
                                               Try : picPreview.Image = Image.FromFile(fileImgPath) : Catch : End Try
                                           End If
                                       End Using
                                   End Sub
        AddHandler btnAnal2.Click, Async Sub(s, e)
                                       Dim m = owner.GetMacchinaCorrente()
                                       If m Is Nothing OrElse String.IsNullOrEmpty(fileImgPath) Then MessageBox.Show("Seleziona prima un'immagine.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information) : Return
                                       btnAnal2.Enabled = False : btnAnal2.Text = "Analisi in corso..."
                                       Try
                                           Dim ce = Await ai.AnalizzaScreenshotErrore(fileImgPath, m.NomeMacchina, owner.GetLinguaSelezionata())
                                           ce.MacchinaID = m.ID : ce.NomeScreenshot = Path.GetFileName(fileImgPath) : ce.PathScreenshot = fileImgPath
                                           Using f As New MC_FrmEditCodiceErrore(ce, m.ID, db)
                                               f.Text = "Verifica e salva errore rilevato"
                                               f.ShowDialog(owner)
                                           End Using
                                       Catch ex As Exception
                                           MessageBox.Show("Errore AI: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                       Finally
                                           btnAnal2.Enabled = True : btnAnal2.Text = "Analizza con AI ✦"
                                       End Try
                                   End Sub
        card2.Controls.AddRange({lblInfo2, picPreview, lblFile2, btnSfog2, btnAnal2})

        pnl.Controls.AddRange({lblTitle, card1, card2})
        Return pnl
    End Function

    ' ════════════════════════════════════════════
    ' PANEL CODICI ERRORE
    ' ════════════════════════════════════════════

    Public Function BuildPanelErrori(owner As MC_FrmMain, db As MC_DatabaseService, ai As MC_AnthropicService) As Panel
        Dim pnl As New Panel()
        Dim lblTitle As New Label() With {.Text = "Codici errore", .Font = New Font("Segoe UI Semibold", 16), .ForeColor = Color.FromArgb(40, 40, 40), .AutoSize = True, .Location = New Point(0, 0)}

        Dim dgv As New DataGridView() With {
            .Location = New Point(0, 50), .Size = New Size(900, 380),
            .AllowUserToAddRows = False, .ReadOnly = True,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .RowHeadersVisible = False, .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle, .Font = FONT_BODY, .Name = "dgvErrori"
        }
        dgv.ColumnHeadersDefaultCellStyle.BackColor = BLUE_LIGHT
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = BLUE_DARK
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI Semibold", 9)
        dgv.EnableHeadersVisualStyles = False

        Dim btnNuovo As New Button() With {.Text = "+ Aggiungi", .Location = New Point(0, 448), .Size = New Size(120, 32), .Font = FONT_BODY}
        Dim btnModifica As New Button() With {.Text = "Modifica", .Location = New Point(130, 448), .Size = New Size(100, 32), .Font = FONT_BODY}
        Dim btnElimina As New Button() With {.Text = "Elimina", .Location = New Point(240, 448), .Size = New Size(100, 32), .Font = FONT_BODY}
        Dim btnGenCap As New Button() With {.Text = "Genera capitolo ✦", .Location = New Point(590, 448), .Size = New Size(160, 32), .Font = FONT_BODY}
        StyleButton(btnNuovo, True) : StyleButton(btnModifica, False) : StyleButton(btnElimina, False) : StyleButton(btnGenCap, True)

        AddHandler btnNuovo.Click, Sub(s, e)
                                       Dim m = owner.GetMacchinaCorrente() : If m Is Nothing Then Return
                                       Using f As New MC_FrmEditCodiceErrore(Nothing, m.ID, db)
                                           If f.ShowDialog(owner) = DialogResult.OK Then RicaricaErrori(pnl, db, m)
                                       End Using
                                   End Sub
        AddHandler btnModifica.Click, Sub(s, e)
                                          If dgv.SelectedRows.Count = 0 Then Return
                                          Dim ce = TryCast(dgv.SelectedRows(0).Tag, MC_CodiceErrore) : If ce Is Nothing Then Return
                                          Using f As New MC_FrmEditCodiceErrore(ce, ce.MacchinaID, db)
                                              If f.ShowDialog(owner) = DialogResult.OK Then RicaricaErrori(pnl, db, owner.GetMacchinaCorrente())
                                          End Using
                                      End Sub
        AddHandler btnElimina.Click, Sub(s, e)
                                         If dgv.SelectedRows.Count = 0 Then Return
                                         Dim ce = TryCast(dgv.SelectedRows(0).Tag, MC_CodiceErrore) : If ce Is Nothing Then Return
                                         If MessageBox.Show($"Eliminare '{ce.Codice}'?", "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
                                             db.EliminaCodiceErrore(ce.ID) : RicaricaErrori(pnl, db, owner.GetMacchinaCorrente())
                                         End If
                                     End Sub
        AddHandler btnGenCap.Click, Async Sub(s, e)
                                        Dim m = owner.GetMacchinaCorrente() : If m Is Nothing Then Return
                                        Dim lista = db.GetCodiciErrore(m.ID)
                                        If lista.Count = 0 Then MessageBox.Show("Nessun codice errore.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information) : Return
                                        btnGenCap.Enabled = False : btnGenCap.Text = "Generazione..."
                                        Try
                                            Dim testo = Await ai.GeneraCapitoloErrori(lista, m, owner.GetLinguaSelezionata())
                                            Using f As New MC_FrmTestoGenerato("Codici errore", testo)
                                                f.ShowDialog(owner)
                                            End Using
                                        Catch ex As Exception
                                            MessageBox.Show("Errore AI: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        Finally
                                            btnGenCap.Enabled = True : btnGenCap.Text = "Genera capitolo ✦"
                                        End Try
                                    End Sub

        pnl.Controls.AddRange({lblTitle, dgv, btnNuovo, btnModifica, btnElimina, btnGenCap})
        Return pnl
    End Function

    Public Sub RicaricaErrori(pnl As Panel, db As MC_DatabaseService, m As MC_Macchina)
        Dim dgv = TryCast(pnl.Controls.Find("dgvErrori", True).FirstOrDefault(), DataGridView)
        If dgv Is Nothing Then Return
        dgv.Rows.Clear() : dgv.Columns.Clear()
        For Each col In {"Codice", "Gravità", "Titolo", "Screenshot"} : dgv.Columns.Add(col, col) : Next
        If m Is Nothing Then Return
        For Each ce In db.GetCodiciErrore(m.ID)
            dgv.Rows.Add(ce.Codice, ce.Gravita, ce.Titolo, If(String.IsNullOrEmpty(ce.NomeScreenshot), "—", ce.NomeScreenshot))
            Dim row = dgv.Rows(dgv.Rows.Count - 1)
            row.Tag = ce
            If ce.Gravita = "Blocco" Then row.DefaultCellStyle.ForeColor = Color.DarkRed
            If ce.Gravita = "Allarme" Then row.DefaultCellStyle.ForeColor = Color.DarkOrange
        Next
    End Sub

    ' ════════════════════════════════════════════
    ' PANEL GENERA MANUALE
    ' ════════════════════════════════════════════

    Public Function BuildPanelGenera(owner As MC_FrmMain, db As MC_DatabaseService, ai As MC_AnthropicService, word As MC_WordService) As Panel
        Dim pnl As New Panel()
        Dim lblTitle As New Label() With {.Text = "Genera manuale", .Font = New Font("Segoe UI Semibold", 16), .ForeColor = Color.FromArgb(40, 40, 40), .AutoSize = True, .Location = New Point(0, 0)}
        Dim card = BuildCard(0, 50, 620, 380, "Configurazione generazione")

        Dim lblMac As New Label() With {.Text = "Macchina:", .Location = New Point(16, 50), .AutoSize = True, .Font = FONT_LABEL}
        Dim lblMacVal As New Label() With {.Text = "—", .Location = New Point(160, 50), .Size = New Size(420, 20), .Font = FONT_BODY, .ForeColor = BLUE_DARK, .Name = "lblMacValGenera"}
        Dim lblRev As New Label() With {.Text = "Revisione:", .Location = New Point(16, 84), .AutoSize = True, .Font = FONT_LABEL}
        Dim txtRev As New TextBox() With {.Text = "Rev. 1", .Location = New Point(160, 81), .Size = New Size(120, 24), .Font = FONT_BODY}
        Dim lblLng As New Label() With {.Text = "Lingua:", .Location = New Point(16, 118), .AutoSize = True, .Font = FONT_LABEL}
        Dim cmbLng As New ComboBox() With {.DropDownStyle = ComboBoxStyle.DropDownList, .Location = New Point(160, 115), .Size = New Size(150, 24), .Font = FONT_BODY}
        Try
            Dim lingue = db.GetLingue()
            cmbLng.DataSource = lingue : cmbLng.DisplayMember = "Nome" : cmbLng.SelectedIndex = 0
        Catch : cmbLng.DataSource = Nothing : cmbLng.Items.Add("Italiano") : cmbLng.SelectedIndex = 0 : End Try

        Dim chkAiFotoc As New CheckBox() With {.Text = "Genera testo cap. 5.1 con AI", .Location = New Point(16, 152), .AutoSize = True, .Font = FONT_BODY, .Checked = True}
        Dim chkAiErr As New CheckBox() With {.Text = "Genera testo codici errore con AI", .Location = New Point(16, 178), .AutoSize = True, .Font = FONT_BODY, .Checked = True}
        Dim lblOut As New Label() With {.Text = "Cartella output:", .Location = New Point(16, 214), .AutoSize = True, .Font = FONT_LABEL}
        Dim txtOut As New TextBox() With {.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop), .Location = New Point(160, 211), .Size = New Size(360, 24), .Font = FONT_BODY}
        Dim btnOut As New Button() With {.Text = "...", .Location = New Point(530, 211), .Size = New Size(36, 24), .Font = FONT_BODY}
        StyleButton(btnOut, False)
        AddHandler btnOut.Click, Sub(s, e)
                                     Using fbd As New FolderBrowserDialog()
                                         If fbd.ShowDialog() = DialogResult.OK Then txtOut.Text = fbd.SelectedPath
                                     End Using
                                 End Sub

        Dim pbar As New ProgressBar() With {.Location = New Point(16, 252), .Size = New Size(560, 16), .Style = ProgressBarStyle.Marquee, .Visible = False}
        Dim lblStatus As New Label() With {.Text = "", .Location = New Point(16, 274), .Size = New Size(560, 20), .Font = FONT_LABEL, .ForeColor = Color.Gray}
        Dim btnGenera As New Button() With {.Text = "Genera manuale .docx ✦", .Location = New Point(16, 300), .Size = New Size(200, 36), .Font = FONT_BODY}
        StyleButton(btnGenera, True)

        AddHandler btnGenera.Click, Async Sub(s, e)
                                        Dim m = owner.GetMacchinaCorrente()
                                        If m Is Nothing Then MessageBox.Show("Nessuna macchina selezionata.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Return
                                        Dim lingua = If(TryCast(cmbLng.SelectedItem, MC_Lingua)?.Codice, "IT")
                                        Dim revisione = txtRev.Text.Trim()
                                        Dim nomeFile = Path.Combine(txtOut.Text.Trim(), $"{m.Matricola}_{m.Modello}_Manuale_{revisione.Replace(" ", "_")}_{lingua}.docx")
                                        btnGenera.Enabled = False : pbar.Visible = True : lblStatus.Text = "Caricamento dati..."
                                        Try
                                            Dim fotocellule = db.GetFotocellule(m.ID)
                                            Dim errori = db.GetCodiciErrore(m.ID)
                                            Dim testoFotoc = ""
                                            Dim testoErr = ""
                                            If chkAiFotoc.Checked AndAlso fotocellule.Count > 0 Then
                                                lblStatus.Text = "Generazione cap. 5.1 con AI..."
                                                testoFotoc = Await ai.GeneraCapitoloFotocellule(fotocellule, m, lingua)
                                            End If
                                            If chkAiErr.Checked AndAlso errori.Count > 0 Then
                                                lblStatus.Text = "Generazione codici errore con AI..."
                                                testoErr = Await ai.GeneraCapitoloErrori(errori, m, lingua)
                                            End If
                                            lblStatus.Text = "Creazione documento Word..."
                                            word.GeneraManuale(nomeFile, m, fotocellule, errori, testoFotoc, testoErr, revisione, lingua)
                                            lblStatus.Text = "Completato!"
                                            If MessageBox.Show($"Manuale generato:{vbCrLf}{nomeFile}{vbCrLf}{vbCrLf}Aprire il file?",
                        "Completato", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then
                                                Process.Start(New ProcessStartInfo(nomeFile) With {.UseShellExecute = True})
                                            End If
                                        Catch ex As Exception
                                            MessageBox.Show("Errore generazione: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        Finally
                                            btnGenera.Enabled = True : pbar.Visible = False
                                        End Try
                                    End Sub

        card.Controls.AddRange({lblMac, lblMacVal, lblRev, txtRev, lblLng, cmbLng,
                                 chkAiFotoc, chkAiErr, lblOut, txtOut, btnOut,
                                 pbar, lblStatus, btnGenera})
        pnl.Controls.AddRange({lblTitle, card})
        Return pnl
    End Function

    Public Sub ImpostaMacchinaGenera(pnl As Panel, m As MC_Macchina)
        Dim lbl = TryCast(pnl.Controls.Find("lblMacValGenera", True).FirstOrDefault(), Label)
        If lbl IsNot Nothing AndAlso m IsNot Nothing Then lbl.Text = $"{m.NomeMacchina}  (Mat. {m.Matricola})"
    End Sub

    ' ════════════════════════════════════════════
    ' HELPER UI
    ' ════════════════════════════════════════════

    Public Function BuildCard(x As Integer, y As Integer, w As Integer, h As Integer, title As String) As Panel
        Dim pnl As New Panel() With {.Location = New Point(x, y), .Size = New Size(w, h), .BackColor = Color.White, .BorderStyle = BorderStyle.FixedSingle}
        Dim header As New Panel() With {.Dock = DockStyle.Top, .Height = 38, .BackColor = Color.White}
        header.Controls.Add(New Label() With {.Text = title, .Font = FONT_TITLE, .ForeColor = Color.FromArgb(40, 40, 40), .Location = New Point(16, 10), .AutoSize = True})
        header.Controls.Add(New Panel() With {.Dock = DockStyle.Bottom, .Height = 1, .BackColor = BORDER_C})
        pnl.Controls.Add(header)
        Return pnl
    End Function

    Public Sub StyleButton(btn As Button, primary As Boolean)
        btn.FlatStyle = FlatStyle.Flat
        btn.FlatAppearance.BorderSize = 1
        btn.Cursor = Cursors.Hand
        If primary Then
            btn.BackColor = Color.FromArgb(24, 95, 165)
            btn.ForeColor = Color.White
            btn.FlatAppearance.BorderColor = Color.FromArgb(24, 95, 165)
        Else
            btn.BackColor = Color.White
            btn.ForeColor = Color.FromArgb(60, 60, 60)
            btn.FlatAppearance.BorderColor = Color.FromArgb(180, 180, 180)
        End If
    End Sub

End Module
