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

        ' ── TabControl ───────────────────────────────────────────────────
        Dim tabs As New TabControl() With {.Dock = DockStyle.Fill, .Font = New Font("Segoe UI", 9.5F)}

        ' ═══════════════════════════════════════
        ' TAB 1 — Ricerca macchina
        ' ═══════════════════════════════════════
        Dim tabRicerca As New TabPage("Ricerca macchina")

        Dim pnlSearch As New Panel() With {.Dock = DockStyle.Top, .Height = 42}
        Dim lblMat As New Label() With {.Text = "Matricola:", .AutoSize = True, .Location = New Point(0, 11), .Font = FONT_LABEL}
        Dim txtMat As New TextBox() With {.Location = New Point(68, 8), .Size = New Size(150, 24), .Font = FONT_BODY}
        Dim lblCli As New Label() With {.Text = "Cliente:", .AutoSize = True, .Location = New Point(228, 11), .Font = FONT_LABEL}
        Dim txtCli As New TextBox() With {.Location = New Point(278, 8), .Size = New Size(200, 24), .Font = FONT_BODY}
        Dim btnCerca As New Button() With {.Text = "Cerca", .Location = New Point(488, 7), .Size = New Size(80, 26), .Font = FONT_BODY}
        StyleButton(btnCerca, True)
        pnlSearch.Controls.AddRange({lblMat, txtMat, lblCli, txtCli, btnCerca})

        Dim pnlBtns As New Panel() With {.Dock = DockStyle.Bottom, .Height = 48}
        Dim btnImposta As New Button() With {.Text = "Imposta come attiva", .Location = New Point(0, 8), .Size = New Size(160, 32), .Font = FONT_BODY}
        Dim btnGestMod As New Button() With {.Text = "Modelli...",          .Location = New Point(170, 8), .Size = New Size(100, 32), .Font = FONT_BODY}
        Dim btnGestTip As New Button() With {.Text = "Tipi macchina...",    .Location = New Point(280, 8), .Size = New Size(130, 32), .Font = FONT_BODY}
        StyleButton(btnImposta, True)
        StyleButton(btnGestMod, False) : StyleButton(btnGestTip, False)
        pnlBtns.Controls.AddRange({btnImposta, btnGestMod, btnGestTip})

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

        ' Stato macchina selezionata nella tab (condiviso tra le tab)
        Dim macchinaSelezionata As MC_Macchina = Nothing

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

        AddHandler btnImposta.Click, Sub(s, e)
            Dim m = GetSelectedMacchina(dgv) : If m Is Nothing Then Return
            If m.ID = 0 Then m.ID = db.SalvaExtraMacchina(m)
            macchinaSelezionata = m
            owner.SetMacchinaCorrente(m)
            MessageBox.Show($"Macchina '{m.NomeMacchina}' ({m.Matricola}) impostata come attiva.",
                            "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Sub

        AddHandler dgv.SelectionChanged, Sub(s, e)
            If dgv.SelectedRows.Count > 0 Then
                macchinaSelezionata = TryCast(dgv.SelectedRows(0).Tag, MC_Macchina)
            End If
        End Sub

        AddHandler btnGestMod.Click, Sub(s, e)
            Using f As New MC_FrmGestisciLookup("Modelli", db) : f.ShowDialog(owner) : End Using
        End Sub
        AddHandler btnGestTip.Click, Sub(s, e)
            Using f As New MC_FrmGestisciLookup("TipiMacchina", db) : f.ShowDialog(owner) : End Using
        End Sub

        ' Ordine controlli nella tab (Fill prima, poi Bottom, poi Top)
        tabRicerca.Controls.Add(dgv)
        tabRicerca.Controls.Add(pnlBtns)
        tabRicerca.Controls.Add(pnlSearch)

        ' ═══════════════════════════════════════
        ' TAB 2 — Dati anagrafici
        ' ═══════════════════════════════════════
        Dim tabAnagrafica As New TabPage("Dati anagrafici")

        Dim lblInfoAn As New Label() With {
            .Text = "Seleziona una macchina nella tab 'Ricerca' per modificarne i dati.",
            .Font = FONT_LABEL, .ForeColor = Color.Gray, .AutoSize = True, .Location = New Point(16, 16)
        }

        Dim yA = 52
        Dim lblNomeA As New Label() With {.Text = "Nome macchina:", .Location = New Point(16, yA + 3), .Size = New Size(130, 20), .Font = FONT_LABEL}
        Dim txtNomeA As New TextBox() With {.Location = New Point(152, yA), .Size = New Size(340, 24), .Font = FONT_BODY}
        yA += 34

        Dim lblModA As New Label() With {.Text = "Modello:", .Location = New Point(16, yA + 3), .Size = New Size(130, 20), .Font = FONT_LABEL}
        Dim cmbModA As New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Location = New Point(152, yA), .Size = New Size(280, 24), .Font = FONT_BODY
        }
        yA += 34

        Dim lblTipoA As New Label() With {.Text = "Tipologia:", .Location = New Point(16, yA + 3), .Size = New Size(130, 20), .Font = FONT_LABEL}
        Dim cmbTipoA As New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Location = New Point(152, yA), .Size = New Size(280, 24), .Font = FONT_BODY
        }
        yA += 34

        Dim lblLngA As New Label() With {.Text = "Lingua:", .Location = New Point(16, yA + 3), .Size = New Size(130, 20), .Font = FONT_LABEL}
        Dim cmbLngA As New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Location = New Point(152, yA), .Size = New Size(120, 24), .Font = FONT_BODY
        }
        For Each l In {"IT", "EN", "FR", "ES", "DE"} : cmbLngA.Items.Add(l) : Next
        cmbLngA.SelectedIndex = 0
        yA += 50

        Dim btnSalvaAn As New Button() With {.Text = "Salva dati anagrafici", .Location = New Point(152, yA), .Size = New Size(180, 32), .Font = FONT_BODY}
        StyleButton(btnSalvaAn, True)

        ' Carica dropdown modelli/tipi
        Dim CaricaDropdownAn As Action = Sub()
            Try
                Dim modelli = db.GetModelli()
                Dim tipi    = db.GetTipiMacchina()
                cmbModA.DataSource    = modelli : cmbModA.DisplayMember = "Nome" : cmbModA.ValueMember = "ID"
                cmbTipoA.DataSource   = tipi    : cmbTipoA.DisplayMember = "Nome" : cmbTipoA.ValueMember = "ID"
            Catch : End Try
        End Sub
        CaricaDropdownAn()

        ' Popola campi quando si entra nella tab
        AddHandler tabAnagrafica.Enter, Sub(s, e)
            CaricaDropdownAn()
            Dim m = macchinaSelezionata
            If m Is Nothing Then Return
            txtNomeA.Text = m.NomeMacchina
            Dim selMod = TryCast(cmbModA.DataSource, List(Of MC_Modello))?.FirstOrDefault(Function(x) x.Nome = m.Modello)
            If selMod IsNot Nothing Then cmbModA.SelectedItem = selMod
            Dim selTipo = TryCast(cmbTipoA.DataSource, List(Of MC_TipoMacchina))?.FirstOrDefault(Function(x) x.Nome = m.TipoMacchina)
            If selTipo IsNot Nothing Then cmbTipoA.SelectedItem = selTipo
            If cmbLngA.Items.Contains(m.LinguaCodice) Then cmbLngA.SelectedItem = m.LinguaCodice
        End Sub

        AddHandler btnSalvaAn.Click, Sub(s, e)
            Dim m = macchinaSelezionata
            If m Is Nothing Then
                MessageBox.Show("Seleziona prima una macchina nella tab 'Ricerca'.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            m.NomeMacchina = txtNomeA.Text.Trim()
            m.Modello      = If(cmbModA.SelectedItem IsNot Nothing, DirectCast(cmbModA.SelectedItem, MC_Modello).Nome, "")
            m.TipoMacchina = If(cmbTipoA.SelectedItem IsNot Nothing, DirectCast(cmbTipoA.SelectedItem, MC_TipoMacchina).Nome, "")
            m.LinguaCodice = cmbLngA.SelectedItem?.ToString()
            Try
                m.ID = db.SalvaExtraMacchina(m)
                If owner.GetMacchinaCorrente()?.Matricola = m.Matricola Then owner.SetMacchinaCorrente(m)
                MessageBox.Show("Dati anagrafici salvati.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Errore salvataggio: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        tabAnagrafica.Controls.AddRange({lblInfoAn, lblNomeA, txtNomeA, lblModA, cmbModA,
                                         lblTipoA, cmbTipoA, lblLngA, cmbLngA, btnSalvaAn})

        ' ═══════════════════════════════════════
        ' TAB 3 — Dati tecnici
        ' ═══════════════════════════════════════
        Dim tabTecnici As New TabPage("Dati tecnici")

        Dim lblInfoTec As New Label() With {
            .Text = "Seleziona una macchina nella tab 'Ricerca' per modificarne i dati tecnici.",
            .Font = FONT_LABEL, .ForeColor = Color.Gray, .AutoSize = True, .Location = New Point(16, 16)
        }

        Dim yT = 52
        Dim lblPeso As New Label() With {.Text = "Peso macchina (kg):", .Location = New Point(16, yT + 3), .Size = New Size(160, 20), .Font = FONT_LABEL}
        Dim txtPeso As New TextBox() With {.Location = New Point(184, yT), .Size = New Size(140, 24), .Font = FONT_BODY}
        yT += 34

        Dim lblAria As New Label() With {.Text = "Consumo aria (Nl/min):", .Location = New Point(16, yT + 3), .Size = New Size(160, 20), .Font = FONT_LABEL}
        Dim txtAria As New TextBox() With {.Location = New Point(184, yT), .Size = New Size(140, 24), .Font = FONT_BODY}
        yT += 34

        Dim lblCor As New Label() With {.Text = "Corrente:", .Location = New Point(16, yT + 3), .Size = New Size(160, 20), .Font = FONT_LABEL}
        Dim txtCor As New TextBox() With {.Location = New Point(184, yT), .Size = New Size(240, 24), .Font = FONT_BODY}
        yT += 34

        Dim lblTen As New Label() With {.Text = "Tensione:", .Location = New Point(16, yT + 3), .Size = New Size(160, 20), .Font = FONT_LABEL}
        Dim txtTen As New TextBox() With {.Location = New Point(184, yT), .Size = New Size(240, 24), .Font = FONT_BODY}
        yT += 50

        Dim btnSalvaTec As New Button() With {.Text = "Salva dati tecnici", .Location = New Point(184, yT), .Size = New Size(160, 32), .Font = FONT_BODY}
        StyleButton(btnSalvaTec, True)

        ' Popola campi quando si entra nella tab
        AddHandler tabTecnici.Enter, Sub(s, e)
            Dim m = macchinaSelezionata
            If m Is Nothing Then Return
            txtPeso.Text = If(m.PesoKg.HasValue, m.PesoKg.Value.ToString("G"), "")
            txtAria.Text = If(m.ConsumoAria.HasValue, m.ConsumoAria.Value.ToString("G"), "")
            txtCor.Text  = m.Corrente
            txtTen.Text  = m.Tensione
        End Sub

        AddHandler btnSalvaTec.Click, Sub(s, e)
            Dim m = macchinaSelezionata
            If m Is Nothing Then
                MessageBox.Show("Seleziona prima una macchina nella tab 'Ricerca'.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            Dim peso As Double, aria As Double
            m.PesoKg      = If(Double.TryParse(txtPeso.Text.Trim().Replace(",", "."), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, peso), CType(peso, Double?), Nothing)
            m.ConsumoAria = If(Double.TryParse(txtAria.Text.Trim().Replace(",", "."), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, aria), CType(aria, Double?), Nothing)
            m.Corrente    = txtCor.Text.Trim()
            m.Tensione    = txtTen.Text.Trim()
            Try
                m.ID = db.SalvaExtraMacchina(m)
                MessageBox.Show("Dati tecnici salvati.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Errore salvataggio: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        tabTecnici.Controls.AddRange({lblInfoTec, lblPeso, txtPeso, lblAria, txtAria,
                                      lblCor, txtCor, lblTen, txtTen, btnSalvaTec})

        ' ═══════════════════════════════════════
        ' TAB 4 — Tutte le macchine (filtro)
        ' ═══════════════════════════════════════
        Dim tabTutte As New TabPage("Tutte le macchine")

        Dim pnlFiltro As New Panel() With {.Dock = DockStyle.Top, .Height = 42}
        Dim lblFiltro As New Label() With {.Text = "Filtra:", .AutoSize = True, .Location = New Point(0, 11), .Font = FONT_LABEL}
        Dim txtFiltro As New TextBox() With {.Location = New Point(44, 8), .Size = New Size(300, 24), .Font = FONT_BODY}
        Dim btnRicarica As New Button() With {.Text = "Aggiorna", .Location = New Point(360, 7), .Size = New Size(90, 26), .Font = FONT_BODY}
        StyleButton(btnRicarica, False)
        pnlFiltro.Controls.AddRange({lblFiltro, txtFiltro, btnRicarica})

        Dim dgvTutte As New DataGridView() With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False, .ReadOnly = True,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .RowHeadersVisible = False, .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle, .Font = FONT_BODY, .Name = "dgvTutte"
        }
        dgvTutte.ColumnHeadersDefaultCellStyle.BackColor = BLUE_LIGHT
        dgvTutte.ColumnHeadersDefaultCellStyle.ForeColor = BLUE_DARK
        dgvTutte.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI Semibold", 9)
        dgvTutte.EnableHeadersVisualStyles = False

        ' DataTable sorgente per il filtro
        Dim dtTutte As New System.Data.DataTable()
        Dim colHeaders = {"Matricola", "Nome macchina", "Cliente", "Modello", "Tipologia", "Lingua",
                          "Peso (kg)", "Consumo aria (Nl/min)", "Corrente", "Tensione"}
        For Each h In colHeaders : dtTutte.Columns.Add(h) : Next

        Dim CaricaTutte As Action = Sub()
            dtTutte.Rows.Clear()
            Try
                For Each m In db.GetMacchine(False)
                    dtTutte.Rows.Add(
                        m.Matricola, m.NomeMacchina, m.ClienteFinale, m.Modello, m.TipoMacchina, m.LinguaCodice,
                        If(m.PesoKg.HasValue, CObj(m.PesoKg.Value), DBNull.Value),
                        If(m.ConsumoAria.HasValue, CObj(m.ConsumoAria.Value), DBNull.Value),
                        m.Corrente, m.Tensione)
                Next
            Catch ex As Exception
                MessageBox.Show("Errore caricamento: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Dim dv As New System.Data.DataView(dtTutte)
            If Not String.IsNullOrWhiteSpace(txtFiltro.Text) Then
                Dim f = txtFiltro.Text.Trim().Replace("'", "''")
                Dim filtri = colHeaders.Select(Function(c) $"CONVERT([{c}], System.String) LIKE '%{f}%'")
                dv.RowFilter = String.Join(" OR ", filtri)
            End If
            dgvTutte.DataSource = dv
        End Sub

        AddHandler tabTutte.Enter, Sub(s, e) CaricaTutte()
        AddHandler btnRicarica.Click, Sub(s, e) CaricaTutte()
        AddHandler txtFiltro.TextChanged, Sub(s, e) CaricaTutte()

        ' Ordine: Fill prima, poi Top
        tabTutte.Controls.Add(dgvTutte)
        tabTutte.Controls.Add(pnlFiltro)

        ' ── Assembla TabControl ───────────────────────────────────────────
        tabs.TabPages.AddRange({tabRicerca, tabAnagrafica, tabTecnici, tabTutte})

        pnl.Controls.Add(tabs)
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

        Dim btnAggiungi As New Button() With {.Text = "+ Aggiungi dal catalogo", .Location = New Point(0, 450), .Size = New Size(180, 32), .Font = FONT_BODY}
        Dim btnRimuovi  As New Button() With {.Text = "Rimuovi",               .Location = New Point(190, 450), .Size = New Size(100, 32), .Font = FONT_BODY}
        Dim btnGenCap   As New Button() With {.Text = "Genera cap. 5.1 ✦",    .Location = New Point(620, 450), .Size = New Size(160, 32), .Font = FONT_BODY}
        StyleButton(btnAggiungi, True) : StyleButton(btnRimuovi, False) : StyleButton(btnGenCap, True)

        AddHandler btnAggiungi.Click, Sub(s, e)
            Dim m = owner.GetMacchinaCorrente() : If m Is Nothing Then Return
            Using picker As New MC_FrmSelezionaFotocellula(db)
                If picker.ShowDialog(owner) = DialogResult.OK AndAlso picker.Selezionata IsNot Nothing Then
                    db.SalvaFotocellula(New MC_Fotocellula() With {.MacchinaID = m.ID, .CatalogoID = picker.Selezionata.ID})
                    RicaricaFotocellule(pnl, db, m)
                End If
            End Using
        End Sub

        AddHandler btnRimuovi.Click, Sub(s, e)
            If dgv.SelectedRows.Count = 0 Then Return
            Dim fc = TryCast(dgv.SelectedRows(0).Tag, MC_Fotocellula) : If fc Is Nothing Then Return
            If MessageBox.Show($"Rimuovere '{fc.Codice}'?", "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
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

        ' Anteprima immagine sulla selezione
        AddHandler dgv.SelectionChanged, Sub(s, e)
            If dgv.SelectedRows.Count = 0 Then Return
            Dim fc = TryCast(dgv.SelectedRows(0).Tag, MC_Fotocellula)
        End Sub

        pnl.Controls.AddRange({lblTitle, lblSub, dgv, btnAggiungi, btnRimuovi, btnGenCap})
        Return pnl
    End Function

    Public Sub RicaricaFotocellule(pnl As Panel, db As MC_DatabaseService, m As MC_Macchina)
        Dim dgv = TryCast(pnl.Controls.Find("dgvFotoc", True).FirstOrDefault(), DataGridView)
        If dgv Is Nothing Then Return
        dgv.Rows.Clear() : dgv.Columns.Clear()
        For Each col In {"Codice", "Tipo", "Immagine"}
            dgv.Columns.Add(col, col)
        Next
        If m Is Nothing Then Return
        For Each f In db.GetFotocellule(m.ID)
            dgv.Rows.Add(f.Codice, f.TipoNome, If(String.IsNullOrEmpty(f.PathImmagine), "", "✔"))
            dgv.Rows(dgv.Rows.Count - 1).Tag = f
        Next
    End Sub

    ' ════════════════════════════════════════════
    ' PANEL CATALOGO FOTOCELLULE
    ' ════════════════════════════════════════════

    Public Function BuildPanelCatalogo(owner As MC_FrmMain, db As MC_DatabaseService) As Panel
        Dim pnl As New Panel()

        ' ── Titolo ───────────────────────────────────────────────────────
        Dim pnlTitle As New Panel() With {.Dock = DockStyle.Top, .Height = 50}
        pnlTitle.Controls.Add(New Label() With {
            .Text = "Catalogo fotocellule", .Font = New Font("Segoe UI Semibold", 16),
            .ForeColor = Color.FromArgb(40, 40, 40), .AutoSize = True, .Location = New Point(0, 8)
        })

        ' ── Barra ricerca + bottoni ───────────────────────────────────────
        Dim pnlTop As New Panel() With {.Dock = DockStyle.Top, .Height = 44}
        Dim lblFiltro As New Label() With {.Text = "Cerca:", .AutoSize = True, .Location = New Point(0, 12), .Font = FONT_LABEL}
        Dim txtFiltro As New TextBox() With {.Location = New Point(46, 9), .Size = New Size(200, 24), .Font = FONT_BODY}
        Dim btnNuovo  As New Button() With {.Text = "+ Nuovo",     .Location = New Point(260, 8), .Size = New Size(100, 28), .Font = FONT_BODY}
        Dim btnMod    As New Button() With {.Text = "Modifica",    .Location = New Point(368, 8), .Size = New Size(90, 28), .Font = FONT_BODY}
        Dim btnDel    As New Button() With {.Text = "Elimina",     .Location = New Point(466, 8), .Size = New Size(90, 28), .Font = FONT_BODY}
        Dim btnTipi   As New Button() With {.Text = "Tipi...",     .Location = New Point(566, 8), .Size = New Size(80, 28), .Font = FONT_BODY}
        StyleButton(btnNuovo, True) : StyleButton(btnMod, False) : StyleButton(btnDel, False) : StyleButton(btnTipi, False)
        pnlTop.Controls.AddRange({lblFiltro, txtFiltro, btnNuovo, btnMod, btnDel, btnTipi})

        ' ── Layout split: griglia sx + anteprima dx ───────────────────────
        Dim pnlSplit As New Panel() With {.Dock = DockStyle.Fill}

        ' Pannello anteprima immagine (destra)
        Dim pnlPreview As New Panel() With {
            .Dock = DockStyle.Right, .Width = 280,
            .BackColor = Color.White,
            .Padding = New Padding(8)
        }
        pnlPreview.Controls.Add(New Panel() With {.Dock = DockStyle.Left, .Width = 1, .BackColor = BORDER_C})
        Dim lblPrevTitolo As New Label() With {
            .Text = "Anteprima immagine", .Font = FONT_LABEL, .ForeColor = Color.Gray,
            .Dock = DockStyle.Top, .Height = 24, .TextAlign = ContentAlignment.MiddleCenter
        }
        Dim picPreview As New PictureBox() With {
            .Dock = DockStyle.Fill, .SizeMode = PictureBoxSizeMode.Zoom,
            .BackColor = Color.FromArgb(245, 245, 242)
        }
        Dim lblPrevInfo As New Label() With {
            .Dock = DockStyle.Bottom, .Height = 40, .Font = FONT_LABEL,
            .ForeColor = Color.FromArgb(60, 60, 60), .TextAlign = ContentAlignment.MiddleCenter
        }
        pnlPreview.Controls.AddRange({picPreview, lblPrevTitolo, lblPrevInfo})

        ' Griglia catalogo (sinistra)
        Dim dgv As New DataGridView() With {
            .Dock = DockStyle.Fill,
            .AllowUserToAddRows = False, .ReadOnly = True,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            .RowHeadersVisible = False, .BackgroundColor = Color.White,
            .BorderStyle = BorderStyle.FixedSingle, .Font = FONT_BODY, .Name = "dgvCatalogo"
        }
        dgv.ColumnHeadersDefaultCellStyle.BackColor = BLUE_LIGHT
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = BLUE_DARK
        dgv.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI Semibold", 9)
        dgv.EnableHeadersVisualStyles = False

        ' ── Azione di caricamento ─────────────────────────────────────────
        Dim Carica As Action = Sub()
            dgv.Rows.Clear() : dgv.Columns.Clear()
            For Each h In {"Codice", "Tipo", "Immagine"}
                dgv.Columns.Add(h, h)
            Next
            Try
                For Each c In db.GetCatalogoFotocellule(txtFiltro.Text.Trim())
                    dgv.Rows.Add(c.Codice, c.TipoNome, If(String.IsNullOrEmpty(c.PathImmagine), "", "✔"))
                    dgv.Rows(dgv.Rows.Count - 1).Tag = c
                Next
            Catch ex As Exception
                MessageBox.Show("Errore caricamento: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ' Aggiorna anteprima quando cambia la selezione
        AddHandler dgv.SelectionChanged, Sub(s, e)
            If dgv.SelectedRows.Count = 0 Then
                picPreview.Image = Nothing : lblPrevInfo.Text = "" : Return
            End If
            Dim cat = TryCast(dgv.SelectedRows(0).Tag, MC_CatalogoFotocellula)
            If cat Is Nothing Then Return
            lblPrevInfo.Text = $"{cat.Codice}{vbLf}{cat.TipoNome}"
            If Not String.IsNullOrEmpty(cat.PathImmagine) AndAlso File.Exists(cat.PathImmagine) Then
                Try
                    picPreview.Image = Image.FromFile(cat.PathImmagine)
                Catch
                    picPreview.Image = Nothing
                End Try
            Else
                picPreview.Image = Nothing
            End If
        End Sub

        AddHandler txtFiltro.TextChanged, Sub(s, e) Carica()

        AddHandler btnNuovo.Click, Sub(s, e)
            Using f As New MC_FrmEditCatalogoFotocellula(Nothing, db)
                If f.ShowDialog(owner) = DialogResult.OK Then Carica()
            End Using
        End Sub

        AddHandler btnMod.Click, Sub(s, e)
            If dgv.SelectedRows.Count = 0 Then Return
            Dim cat = TryCast(dgv.SelectedRows(0).Tag, MC_CatalogoFotocellula) : If cat Is Nothing Then Return
            Using f As New MC_FrmEditCatalogoFotocellula(cat, db)
                If f.ShowDialog(owner) = DialogResult.OK Then Carica()
            End Using
        End Sub

        AddHandler btnDel.Click, Sub(s, e)
            If dgv.SelectedRows.Count = 0 Then Return
            Dim cat = TryCast(dgv.SelectedRows(0).Tag, MC_CatalogoFotocellula) : If cat Is Nothing Then Return
            If MessageBox.Show($"Eliminare '{cat.Codice}' dal catalogo?", "Conferma",
                               MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = DialogResult.Yes Then
                db.EliminaCatalogoFotocellula(cat.ID) : Carica()
            End If
        End Sub

        AddHandler btnTipi.Click, Sub(s, e)
            Using f As New MC_FrmGestisciLookupFotoc("TipiFotocellule", db)
                f.ShowDialog(owner)
            End Using
        End Sub

        ' Carica al primo accesso al pannello
        AddHandler pnl.VisibleChanged, Sub(s, e) If pnl.Visible Then Carica()

        pnlSplit.Controls.Add(dgv)
        pnlSplit.Controls.Add(pnlPreview)

        pnl.Controls.Add(pnlSplit)
        pnl.Controls.Add(pnlTop)
        pnl.Controls.Add(pnlTitle)
        Return pnl
    End Function

    ' ════════════════════════════════════════════
    ' PANEL SOFTWARE PLC → GENERAZIONE MANUALE
    ' ════════════════════════════════════════════

    Public Function BuildPanelSoftware(owner As MC_FrmMain, db As MC_DatabaseService, ai As MC_AnthropicService) As Panel
        Dim pnl As New Panel()

        ' ── Titolo ───────────────────────────────────────────────────────
        Dim pnlTitle As New Panel() With {.Dock = DockStyle.Top, .Height = 50}
        pnlTitle.Controls.Add(New Label() With {
            .Text = "Analisi programma macchina", .Font = New Font("Segoe UI Semibold", 16),
            .ForeColor = Color.FromArgb(40, 40, 40), .AutoSize = True, .Location = New Point(0, 8)
        })

        ' ── Barra caricamento file ────────────────────────────────────────
        Dim pnlLoad As New Panel() With {.Dock = DockStyle.Top, .Height = 44, .BackColor = Color.White}
        pnlLoad.Controls.Add(New Panel() With {.Dock = DockStyle.Bottom, .Height = 1, .BackColor = BORDER_C})
        Dim lblPath As New Label() With {
            .Text = "Nessun programma caricato", .AutoSize = False,
            .Location = New Point(8, 12), .Size = New Size(480, 20),
            .Font = FONT_LABEL, .ForeColor = Color.Gray
        }
        Dim btnSfoglia As New Button() With {.Text = "Carica programma...", .Location = New Point(500, 8), .Size = New Size(150, 28), .Font = FONT_BODY}
        Dim btnApri    As New Button() With {.Text = "Apri nel software",   .Location = New Point(658, 8), .Size = New Size(140, 28), .Font = FONT_BODY, .Enabled = False}
        Dim btnAnalizza As New Button() With {.Text = "Analizza e genera manuale ✦", .Location = New Point(806, 8), .Size = New Size(210, 28), .Font = FONT_BODY, .Enabled = False}
        StyleButton(btnSfoglia, False) : StyleButton(btnApri, False) : StyleButton(btnAnalizza, True)
        pnlLoad.Controls.AddRange({lblPath, btnSfoglia, btnApri, btnAnalizza})

        ' ── Status bar ───────────────────────────────────────────────────
        Dim pnlStatus As New Panel() With {.Dock = DockStyle.Bottom, .Height = 28, .BackColor = Color.FromArgb(245, 245, 242)}
        pnlStatus.Controls.Add(New Panel() With {.Dock = DockStyle.Top, .Height = 1, .BackColor = BORDER_C})
        Dim lblStatus As New Label() With {
            .Text = "Carica un file di progetto PLC (.zip, .gx3, .gxw, .ap15, .zap15, .mer, .prj, .txt) per iniziare.",
            .Dock = DockStyle.Fill, .Font = FONT_LABEL, .ForeColor = Color.Gray,
            .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(8, 0, 0, 0)
        }
        pnlStatus.Controls.Add(lblStatus)

        ' ── TabControl output ─────────────────────────────────────────────
        Dim tabs As New TabControl() With {.Dock = DockStyle.Fill, .Font = New Font("Segoe UI", 9.5F)}

        Dim tabOp  As New TabPage("5.1 – Descrizione operazione")
        Dim tabCmd As New TabPage("5.2 – Comandi e sensori")
        Dim tabAlr As New TabPage("9.3 – Allarmi")

        Dim MakeRtb As Func(Of RichTextBox) = Function()
            Return New RichTextBox() With {
                .Dock = DockStyle.Fill, .Font = New Font("Segoe UI", 9.5F),
                .BackColor = Color.White, .ReadOnly = False
            }
        End Function
        Dim rtbOp  = MakeRtb() : tabOp.Controls.Add(rtbOp)
        Dim rtbCmd = MakeRtb() : tabCmd.Controls.Add(rtbCmd)
        Dim rtbAlr = MakeRtb() : tabAlr.Controls.Add(rtbAlr)

        tabs.TabPages.AddRange({tabOp, tabCmd, tabAlr})

        ' ── Stato ────────────────────────────────────────────────────────
        Dim filePath As String = ""

        ' ── Bottone Sfoglia ───────────────────────────────────────────────
        AddHandler btnSfoglia.Click, Sub(s, e)
            Using ofd As New OpenFileDialog() With {
                .Filter = "Progetto PLC|*.zip;*.gx3;*.gxw;*.ap15;*.zap15;*.mer;*.prj;*.7z;*.txt|Tutti|*.*",
                .Title = "Seleziona file programma PLC"
            }
                If ofd.ShowDialog() = DialogResult.OK Then
                    filePath = ofd.FileName
                    lblPath.Text = filePath
                    lblPath.ForeColor = Color.FromArgb(40, 40, 40)
                    btnApri.Enabled    = True
                    btnAnalizza.Enabled = True
                    lblStatus.Text = $"File caricato: {Path.GetFileName(filePath)} — pronto per l'analisi."
                End If
            End Using
        End Sub

        ' ── Apri nel software nativo ──────────────────────────────────────
        AddHandler btnApri.Click, Sub(s, e)
            Try
                Diagnostics.Process.Start(New Diagnostics.ProcessStartInfo(filePath) With {.UseShellExecute = True})
            Catch ex As Exception
                MessageBox.Show("Impossibile aprire il file: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End Sub

        ' ── Analizza ─────────────────────────────────────────────────────
        AddHandler btnAnalizza.Click, Async Sub(s, e)
            Dim m = owner.GetMacchinaCorrente()
            If m Is Nothing OrElse String.IsNullOrEmpty(filePath) Then
                MessageBox.Show("Seleziona prima una macchina attiva e un file programma.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning) : Return
            End If
            btnAnalizza.Enabled = False : btnAnalizza.Text = "Analisi in corso..." : lblStatus.Text = "Lettura file..."
            Try
                Dim contenuto = EstraiContenutoProgramma(filePath)
                lblStatus.Text = $"Programma letto ({contenuto.Length:N0} caratteri). Invio ad AI per analisi..."

                Dim risultati = Await ai.AnalizzaProgrammaCompleto(contenuto, m, owner.GetLinguaSelezionata())

                If risultati.ContainsKey("operazione") Then rtbOp.Text  = risultati("operazione")
                If risultati.ContainsKey("comandi")    Then rtbCmd.Text = risultati("comandi")
                If risultati.ContainsKey("allarmi")    Then rtbAlr.Text = risultati("allarmi")

                ' Importa allarmi nel DB se presenti come JSON
                If risultati.ContainsKey("allarmi_json") Then
                    Dim errori = ai.ParseAllarmiJson(risultati("allarmi_json"))
                    If errori.Count > 0 AndAlso
                       MessageBox.Show($"Trovati {errori.Count} allarmi. Importarli nei Codici errore?",
                                       "Importa allarmi", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                        For Each ce In errori : ce.MacchinaID = m.ID : db.SalvaCodiceErrore(ce) : Next
                        MessageBox.Show($"{errori.Count} allarmi importati.", "OK", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                End If

                tabs.SelectedIndex = 0
                lblStatus.Text = "Analisi completata. Rivedi e correggi il testo generato se necessario."
            Catch ex As Exception
                MessageBox.Show("Errore analisi: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                lblStatus.Text = "Errore durante l'analisi."
            Finally
                btnAnalizza.Enabled = True : btnAnalizza.Text = "Analizza e genera manuale ✦"
            End Try
        End Sub

        pnl.Controls.Add(tabs)
        pnl.Controls.Add(pnlStatus)
        pnl.Controls.Add(pnlLoad)
        pnl.Controls.Add(pnlTitle)
        Return pnl
    End Function

    ' Estrae il testo leggibile da un file PLC (zip o testo diretto)
    Private Function EstraiContenutoProgramma(filePath As String) As String
        Dim ext = Path.GetExtension(filePath).ToLowerInvariant()

        ' File zip: estrai tutti i testi leggibili
        If ext = ".zip" OrElse ext = ".7z" OrElse ext = ".zap15" Then
            Dim sb As New System.Text.StringBuilder()
            Try
                Using za = System.IO.Compression.ZipFile.OpenRead(filePath)
                    Dim textExts = {".txt", ".xml", ".csv", ".st", ".fbd", ".il", ".ld", ".sfc", ".cfc", ".exp", ".xte", ".xdb", ".ini", ".cfg", ".json"}
                    Dim entries = za.Entries.Where(Function(e) textExts.Any(Function(x) e.Name.ToLower().EndsWith(x))).OrderBy(Function(e) e.Length).ToList()
                    For Each entry In entries
                        If sb.Length > 150000 Then Exit For
                        Try
                            Using sr As New StreamReader(entry.Open())
                                sb.AppendLine($"=== FILE: {entry.FullName} ===")
                                sb.AppendLine(sr.ReadToEnd())
                                sb.AppendLine()
                            End Using
                        Catch : End Try
                    Next
                End Using
            Catch : End Try
            Return If(sb.Length > 0, sb.ToString(), File.ReadAllText(filePath))
        End If

        ' File testo diretto
        Try
            Return File.ReadAllText(filePath)
        Catch
            Return ""
        End Try
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
