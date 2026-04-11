Imports System.Drawing
Imports System.Windows.Forms

Public Class MC_FrmMain
    Inherits Form

    ' ── Servizi ──────────────────────────────────
    Private ReadOnly _db   As New MC_DatabaseService()
    Private ReadOnly _ai   As New MC_AnthropicService()
    Private ReadOnly _word As New MC_WordService()

    ' ── Stato corrente ────────────────────────────
    Private _macchinaCorrente As MC_Macchina = Nothing

    ' ── Controlli layout ─────────────────────────
    Private WithEvents pnlSidebar As Panel
    Private WithEvents pnlContent As Panel
    Private lblMacchinaAttiva As Label
    Private cmbLingua As ComboBox

    ' ── Pulsanti sidebar ─────────────────────────
    Private WithEvents btnHome         As Button
    Private WithEvents btnMacchine     As Button
    Private WithEvents btnFotocellule  As Button
    Private WithEvents btnSoftware     As Button
    Private WithEvents btnErrori       As Button
    Private WithEvents btnGenera       As Button
    Private WithEvents btnImpostazioni As Button

    ' ── Pannelli sezioni ─────────────────────────
    Private pnlHome        As Panel
    Private pnlMacchine    As Panel
    Private pnlFotocellule As Panel
    Private pnlSoftware    As Panel
    Private pnlErrori      As Panel
    Private pnlGenera      As Panel

    Private _btnCorrente As Button = Nothing

    Public Sub New()
        SetupUI()
        CaricaLingue()
        MostraPannello(pnlHome, btnHome)
    End Sub

    Private Sub SetupUI()
        Me.Text = "ManualCraft CCMS – Tirelli"
        Me.Size = New Size(1200, 800)
        Me.MinimumSize = New Size(950, 650)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = Color.FromArgb(245, 245, 242)
        Me.Font = New Font("Segoe UI", 9)

        BuildPanels()
        BuildSidebar()
    End Sub

    ' ══════════════════════════════════════════════
    ' SIDEBAR
    ' ══════════════════════════════════════════════

    Private Sub BuildSidebar()
        pnlSidebar = New Panel() With {
            .Dock = DockStyle.Left, .Width = 220, .BackColor = Color.White
        }

        ' Logo
        Dim pnlLogo As New Panel() With {.Dock = DockStyle.Top, .Height = 70, .BackColor = Color.White}
        pnlLogo.Controls.Add(New Label() With {
            .Text = "ManualCraft", .Font = New Font("Segoe UI Semibold", 13),
            .ForeColor = Color.FromArgb(24, 95, 165), .Location = New Point(16, 14), .AutoSize = True
        })
        pnlLogo.Controls.Add(New Label() With {
            .Text = "CCMS Tirelli", .Font = New Font("Segoe UI", 8),
            .ForeColor = Color.Gray, .Location = New Point(18, 38), .AutoSize = True
        })
        pnlLogo.Controls.Add(New Panel() With {
            .Dock = DockStyle.Bottom, .Height = 1, .BackColor = Color.FromArgb(230, 230, 230)
        })

        ' Card macchina attiva
        Dim pnlMac As New Panel() With {
            .Dock = DockStyle.Top, .Height = 80,
            .BackColor = Color.FromArgb(230, 241, 251), .Padding = New Padding(12)
        }
        Dim lblTitoloMac As New Label() With {
            .Text = "MACCHINA ATTIVA", .Font = New Font("Segoe UI", 7, FontStyle.Bold),
            .ForeColor = Color.FromArgb(24, 95, 165), .Location = New Point(12, 10), .AutoSize = True
        }
        lblMacchinaAttiva = New Label() With {
            .Text = "Nessuna macchina selezionata",
            .Font = New Font("Segoe UI", 9), .ForeColor = Color.FromArgb(12, 68, 124),
            .Location = New Point(12, 30), .Size = New Size(196, 36), .AutoEllipsis = True
        }
        pnlMac.Controls.AddRange({lblTitoloMac, lblMacchinaAttiva})

        ' Selezione lingua
        Dim pnlLng As New Panel() With {.Dock = DockStyle.Top, .Height = 50, .BackColor = Color.White, .Padding = New Padding(8, 6, 8, 6)}
        Dim lblLingua As New Label() With {
            .Text = "Lingua:", .AutoSize = True, .Location = New Point(12, 16),
            .ForeColor = Color.FromArgb(100, 100, 100), .Font = New Font("Segoe UI", 8)
        }
        cmbLingua = New ComboBox() With {
            .DropDownStyle = ComboBoxStyle.DropDownList,
            .Size = New Size(130, 24), .Location = New Point(60, 13), .Font = New Font("Segoe UI", 9)
        }
        pnlLng.Controls.AddRange({lblLingua, cmbLingua})

        ' Bottoni navigazione
        Dim pnlNav As New Panel() With {.Dock = DockStyle.Fill, .BackColor = Color.White}

        Dim navDefs As (String, Button)() = {
            ("  Home macchina",       New Button()),
            ("  Anagrafica macchine", New Button()),
            ("  Fotocellule (5.1)",   New Button()),
            ("  Software PLC",        New Button()),
            ("  Codici errore",       New Button()),
            ("  Genera manuale",      New Button()),
            ("  Impostazioni",        New Button())
        }

        btnHome         = navDefs(0).Item2
        btnMacchine     = navDefs(1).Item2
        btnFotocellule  = navDefs(2).Item2
        btnSoftware     = navDefs(3).Item2
        btnErrori       = navDefs(4).Item2
        btnGenera       = navDefs(5).Item2
        btnImpostazioni = navDefs(6).Item2

        Dim y = 0
        For Each item In navDefs
            Dim btn = item.Item2
            btn.Text      = item.Item1
            btn.Size      = New Size(220, 42)
            btn.Location  = New Point(0, y)
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 0
            btn.TextAlign = ContentAlignment.MiddleLeft
            btn.Font      = New Font("Segoe UI", 9.5F)
            btn.ForeColor = Color.FromArgb(90, 90, 90)
            btn.BackColor = Color.White
            btn.Cursor    = Cursors.Hand
            pnlNav.Controls.Add(btn)
            y += 42
        Next

        Dim sep As New Panel() With {
            .Size = New Size(220, 1), .Location = New Point(0, 252),
            .BackColor = Color.FromArgb(230, 230, 230)
        }
        pnlNav.Controls.Add(sep)

        pnlSidebar.Controls.Add(pnlNav)
        pnlSidebar.Controls.Add(pnlLng)
        pnlSidebar.Controls.Add(pnlMac)
        pnlSidebar.Controls.Add(pnlLogo)
        pnlSidebar.Controls.Add(New Panel() With {
            .Dock = DockStyle.Right, .Width = 1,
            .BackColor = Color.FromArgb(220, 220, 220)
        })

        Me.Controls.Add(pnlSidebar)
    End Sub

    Private Sub SetNavActive(btn As Button)
        If _btnCorrente IsNot Nothing Then
            _btnCorrente.BackColor = Color.White
            _btnCorrente.ForeColor = Color.FromArgb(90, 90, 90)
            _btnCorrente.Font = New Font("Segoe UI", 9.5F)
        End If
        btn.BackColor = Color.FromArgb(230, 241, 251)
        btn.ForeColor = Color.FromArgb(24, 95, 165)
        btn.Font      = New Font("Segoe UI", 9.5F, FontStyle.Bold)
        _btnCorrente  = btn
    End Sub

    ' ══════════════════════════════════════════════
    ' PANNELLI
    ' ══════════════════════════════════════════════

    Private Sub BuildPanels()
        pnlContent = New Panel() With {
            .Dock = DockStyle.Fill,
            .Padding = New Padding(24),
            .BackColor = Color.FromArgb(245, 245, 242)
        }
        Me.Controls.Add(pnlContent)

        pnlHome        = MC_PanelBuild.BuildPanelHome(Me, _db)
        pnlMacchine    = MC_PanelBuild.BuildPanelMacchine(Me, _db)
        pnlFotocellule = MC_PanelBuild.BuildPanelFotocellule(Me, _db, _ai)
        pnlSoftware    = MC_PanelBuild.BuildPanelSoftware(Me, _db, _ai)
        pnlErrori      = MC_PanelBuild.BuildPanelErrori(Me, _db, _ai)
        pnlGenera      = MC_PanelBuild.BuildPanelGenera(Me, _db, _ai, _word)

        For Each pnl In {pnlHome, pnlMacchine, pnlFotocellule, pnlSoftware, pnlErrori, pnlGenera}
            pnl.Dock    = DockStyle.Fill
            pnl.Visible = False
            pnlContent.Controls.Add(pnl)
        Next
    End Sub

    Private Sub MostraPannello(pannello As Panel, btn As Button)
        For Each pnl In {pnlHome, pnlMacchine, pnlFotocellule, pnlSoftware, pnlErrori, pnlGenera}
            pnl.Visible = False
        Next
        pannello.Visible = True
        SetNavActive(btn)
    End Sub

    ' ══════════════════════════════════════════════
    ' NAVIGAZIONE
    ' ══════════════════════════════════════════════

    Private Sub btnHome_Click(s As Object, e As EventArgs) Handles btnHome.Click
        MostraPannello(pnlHome, btnHome)
    End Sub

    Private Sub btnMacchine_Click(s As Object, e As EventArgs) Handles btnMacchine.Click
        MostraPannello(pnlMacchine, btnMacchine)
    End Sub

    Private Sub btnFotocellule_Click(s As Object, e As EventArgs) Handles btnFotocellule.Click
        If Not CheckMacchina() Then Return
        MC_PanelBuild.RicaricaFotocellule(pnlFotocellule, _db, _macchinaCorrente)
        MostraPannello(pnlFotocellule, btnFotocellule)
    End Sub

    Private Sub btnSoftware_Click(s As Object, e As EventArgs) Handles btnSoftware.Click
        If Not CheckMacchina() Then Return
        MostraPannello(pnlSoftware, btnSoftware)
    End Sub

    Private Sub btnErrori_Click(s As Object, e As EventArgs) Handles btnErrori.Click
        If Not CheckMacchina() Then Return
        MC_PanelBuild.RicaricaErrori(pnlErrori, _db, _macchinaCorrente)
        MostraPannello(pnlErrori, btnErrori)
    End Sub

    Private Sub btnGenera_Click(s As Object, e As EventArgs) Handles btnGenera.Click
        If Not CheckMacchina() Then Return
        MC_PanelBuild.ImpostaMacchinaGenera(pnlGenera, _macchinaCorrente)
        MostraPannello(pnlGenera, btnGenera)
    End Sub

    Private Sub btnImpostazioni_Click(s As Object, e As EventArgs) Handles btnImpostazioni.Click
        Using f As New MC_FrmImpostazioni()
            f.ShowDialog(Me)
        End Using
    End Sub

    ' ──────────────────────────────────────────────
    ' UTILITY pubbliche (usate da MC_PanelBuild)
    ' ──────────────────────────────────────────────

    Public Sub SetMacchinaCorrente(m As MC_Macchina)
        _macchinaCorrente = m
        If m IsNot Nothing Then
            lblMacchinaAttiva.Text = $"{m.NomeMacchina}{vbLf}Mat. {m.Matricola} · {m.ClienteFinale}"
        Else
            lblMacchinaAttiva.Text = "Nessuna macchina selezionata"
        End If
    End Sub

    Public Function GetMacchinaCorrente() As MC_Macchina
        Return _macchinaCorrente
    End Function

    Public Function GetLinguaSelezionata() As String
        If cmbLingua.SelectedItem IsNot Nothing Then
            Return DirectCast(cmbLingua.SelectedItem, MC_Lingua).Codice
        End If
        Return "IT"
    End Function

    Private Sub CaricaLingue()
        Try
            Dim lingue = _db.GetLingue()
            cmbLingua.DataSource    = lingue
            cmbLingua.DisplayMember = "Nome"
            cmbLingua.ValueMember   = "Codice"
            cmbLingua.SelectedIndex = 0
        Catch
            cmbLingua.DataSource = Nothing
            cmbLingua.Items.Add(New MC_Lingua() With {.Codice = "IT", .Nome = "Italiano"})
            cmbLingua.SelectedIndex = 0
        End Try
    End Sub

    Private Function CheckMacchina() As Boolean
        If _macchinaCorrente Is Nothing Then
            MessageBox.Show(
                "Seleziona prima una macchina dalla sezione 'Anagrafica macchine'.",
                "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return False
        End If
        Return True
    End Function

End Class
