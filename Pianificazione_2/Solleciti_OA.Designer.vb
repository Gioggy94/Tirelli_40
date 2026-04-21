<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Solleciti_OA
    Inherits System.Windows.Forms.Form

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.txbStatus = New System.Windows.Forms.TextBox()
        Me.tlpMain = New System.Windows.Forms.TableLayoutPanel()
        Me.gbFiltri = New System.Windows.Forms.GroupBox()
        Me.lblFiltroAcq = New System.Windows.Forms.Label()
        Me.cmbFiltroAcquisitore = New System.Windows.Forms.ComboBox()
        Me.gbStatistiche = New System.Windows.Forms.GroupBox()
        Me.pnlStatTop = New System.Windows.Forms.Panel()
        Me.lblStatGenerale = New System.Windows.Forms.Label()
        Me.scStatistiche = New System.Windows.Forms.SplitContainer()
        Me.dgvStatAcquisitore = New System.Windows.Forms.DataGridView()
        Me.pnlLogTop = New System.Windows.Forms.Panel()
        Me.btnAggiornaLog = New System.Windows.Forms.Button()
        Me.btnInviaReport = New System.Windows.Forms.Button()
        Me.dgvLog = New System.Windows.Forms.DataGridView()
        Me.lblFiltroForn = New System.Windows.Forms.Label()
        Me.cmbFiltroFornitore = New System.Windows.Forms.ComboBox()
        Me.lblFiltroComm = New System.Windows.Forms.Label()
        Me.txtFiltroCommessa = New System.Windows.Forms.TextBox()
        Me.chkSoloScaduti = New System.Windows.Forms.CheckBox()
        Me.chkSoloSollecito = New System.Windows.Forms.CheckBox()
        Me.btnCarica = New System.Windows.Forms.Button()
        Me.lblStato = New System.Windows.Forms.Label()
        Me.scMain = New System.Windows.Forms.SplitContainer()
        Me.gbFornitori = New System.Windows.Forms.GroupBox()
        Me.lvFornitori = New System.Windows.Forms.ListView()
        Me.pnlFornBottom = New System.Windows.Forms.Panel()
        Me.lblConteggioFornitori = New System.Windows.Forms.Label()
        Me.btnToggleSollecito = New System.Windows.Forms.Button()
        Me.gbOrdini = New System.Windows.Forms.GroupBox()
        Me.pnlBtnsOrdini = New System.Windows.Forms.Panel()
        Me.btnSelTutti = New System.Windows.Forms.Button()
        Me.btnDeselTutti = New System.Windows.Forms.Button()
        Me.lblConteggio = New System.Windows.Forms.Label()
        Me.dgvOrdini = New System.Windows.Forms.DataGridView()
        Me.gbAnteprima = New System.Windows.Forms.GroupBox()
        Me.pnlMailHeader = New System.Windows.Forms.Panel()
        Me.lblA = New System.Windows.Forms.Label()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.lblOggetto = New System.Windows.Forms.Label()
        Me.txtOggetto = New System.Windows.Forms.TextBox()
        Me.pnlBtnsAnteprima = New System.Windows.Forms.Panel()
        Me.btnAggiornaAnteprima = New System.Windows.Forms.Button()
        Me.btnPreparaMail = New System.Windows.Forms.Button()
        Me.btnTutteMail = New System.Windows.Forms.Button()
        Me.rtbAnteprima = New System.Windows.Forms.RichTextBox()

        Me.tlpMain.SuspendLayout()
        Me.gbFiltri.SuspendLayout()
        Me.gbStatistiche.SuspendLayout()
        Me.pnlStatTop.SuspendLayout()
        CType(Me.scStatistiche, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.scStatistiche.Panel1.SuspendLayout()
        Me.scStatistiche.Panel2.SuspendLayout()
        Me.scStatistiche.SuspendLayout()
        CType(Me.dgvStatAcquisitore, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlLogTop.SuspendLayout()
        CType(Me.dgvLog, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.scMain, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.scMain.Panel1.SuspendLayout()
        Me.scMain.Panel2.SuspendLayout()
        Me.scMain.SuspendLayout()
        Me.gbFornitori.SuspendLayout()
        Me.gbOrdini.SuspendLayout()
        Me.pnlBtnsOrdini.SuspendLayout()
        CType(Me.dgvOrdini, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gbAnteprima.SuspendLayout()
        Me.pnlMailHeader.SuspendLayout()
        Me.pnlBtnsAnteprima.SuspendLayout()
        Me.SuspendLayout()

        ' ── tlpMain ──────────────────────────────────────────
        Me.tlpMain.ColumnCount = 1
        Me.tlpMain.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpMain.Controls.Add(Me.gbFiltri, 0, 0)
        Me.tlpMain.Controls.Add(Me.scMain, 0, 1)
        Me.tlpMain.Controls.Add(Me.gbStatistiche, 0, 2)
        Me.tlpMain.Controls.Add(Me.gbAnteprima, 0, 3)
        Me.tlpMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tlpMain.Location = New System.Drawing.Point(0, 0)
        Me.tlpMain.Margin = New System.Windows.Forms.Padding(0)
        Me.tlpMain.Name = "tlpMain"
        Me.tlpMain.RowCount = 4
        Me.tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 90.0!))
        Me.tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 160.0!))
        Me.tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 255.0!))
        Me.tlpMain.Size = New System.Drawing.Size(1400, 860)
        Me.tlpMain.TabIndex = 0

        ' ── gbFiltri ─────────────────────────────────────────
        Me.gbFiltri.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbFiltri.Margin = New System.Windows.Forms.Padding(4, 3, 4, 2)
        Me.gbFiltri.Name = "gbFiltri"
        Me.gbFiltri.Text = "Filtri"
        Me.gbFiltri.TabIndex = 0
        Me.gbFiltri.Controls.Add(Me.lblFiltroForn)
        Me.gbFiltri.Controls.Add(Me.cmbFiltroFornitore)
        Me.gbFiltri.Controls.Add(Me.lblFiltroComm)
        Me.gbFiltri.Controls.Add(Me.txtFiltroCommessa)
        Me.gbFiltri.Controls.Add(Me.chkSoloScaduti)
        Me.gbFiltri.Controls.Add(Me.chkSoloSollecito)
        Me.gbFiltri.Controls.Add(Me.btnCarica)
        Me.gbFiltri.Controls.Add(Me.lblStato)
        Me.gbFiltri.Controls.Add(Me.lblFiltroAcq)
        Me.gbFiltri.Controls.Add(Me.cmbFiltroAcquisitore)

        ' lblFiltroForn
        Me.lblFiltroForn.AutoSize = True
        Me.lblFiltroForn.Location = New System.Drawing.Point(8, 24)
        Me.lblFiltroForn.Name = "lblFiltroForn"
        Me.lblFiltroForn.Text = "Fornitore:"

        ' cmbFiltroFornitore
        Me.cmbFiltroFornitore.Location = New System.Drawing.Point(72, 21)
        Me.cmbFiltroFornitore.Name = "cmbFiltroFornitore"
        Me.cmbFiltroFornitore.Size = New System.Drawing.Size(210, 22)
        Me.cmbFiltroFornitore.TabIndex = 0
        Me.cmbFiltroFornitore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFiltroFornitore.DropDownWidth = 380

        ' lblFiltroComm
        Me.lblFiltroComm.AutoSize = True
        Me.lblFiltroComm.Location = New System.Drawing.Point(292, 24)
        Me.lblFiltroComm.Name = "lblFiltroComm"
        Me.lblFiltroComm.Text = "Commessa:"

        ' txtFiltroCommessa
        Me.txtFiltroCommessa.Location = New System.Drawing.Point(366, 21)
        Me.txtFiltroCommessa.Name = "txtFiltroCommessa"
        Me.txtFiltroCommessa.Size = New System.Drawing.Size(120, 22)
        Me.txtFiltroCommessa.TabIndex = 1

        ' chkSoloScaduti
        Me.chkSoloScaduti.AutoSize = True
        Me.chkSoloScaduti.Location = New System.Drawing.Point(500, 23)
        Me.chkSoloScaduti.Name = "chkSoloScaduti"
        Me.chkSoloScaduti.Text = "Solo scaduti"
        Me.chkSoloScaduti.TabIndex = 2

        ' chkSoloSollecito
        Me.chkSoloSollecito.AutoSize = True
        Me.chkSoloSollecito.Location = New System.Drawing.Point(610, 23)
        Me.chkSoloSollecito.Name = "chkSoloSollecito"
        Me.chkSoloSollecito.Text = "Solo da sollecitare"
        Me.chkSoloSollecito.TabIndex = 3
        Me.chkSoloSollecito.ForeColor = System.Drawing.Color.DarkGreen

        ' btnCarica
        Me.btnCarica.Location = New System.Drawing.Point(765, 18)
        Me.btnCarica.Name = "btnCarica"
        Me.btnCarica.Size = New System.Drawing.Size(120, 28)
        Me.btnCarica.TabIndex = 4
        Me.btnCarica.Text = "Carica dati"
        Me.btnCarica.BackColor = System.Drawing.Color.SteelBlue
        Me.btnCarica.ForeColor = System.Drawing.Color.White
        Me.btnCarica.UseVisualStyleBackColor = False

        ' lblStato
        Me.lblStato.AutoSize = False
        Me.lblStato.Location = New System.Drawing.Point(898, 24)
        Me.lblStato.Name = "lblStato"
        Me.lblStato.Size = New System.Drawing.Size(450, 18)
        Me.lblStato.ForeColor = System.Drawing.Color.Navy
        Me.lblStato.Text = ""

        ' lblFiltroAcq
        Me.lblFiltroAcq.AutoSize = True
        Me.lblFiltroAcq.Location = New System.Drawing.Point(8, 50)
        Me.lblFiltroAcq.Name = "lblFiltroAcq"
        Me.lblFiltroAcq.Text = "Acquisitore:"

        ' cmbFiltroAcquisitore
        Me.cmbFiltroAcquisitore.Location = New System.Drawing.Point(90, 47)
        Me.cmbFiltroAcquisitore.Name = "cmbFiltroAcquisitore"
        Me.cmbFiltroAcquisitore.Size = New System.Drawing.Size(210, 22)
        Me.cmbFiltroAcquisitore.TabIndex = 5
        Me.cmbFiltroAcquisitore.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFiltroAcquisitore.DropDownWidth = 280

        ' ── scMain ───────────────────────────────────────────
        Me.scMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.scMain.Margin = New System.Windows.Forms.Padding(4, 2, 4, 2)
        Me.scMain.Name = "scMain"
        Me.scMain.Orientation = System.Windows.Forms.Orientation.Vertical
        Me.scMain.SplitterDistance = 330
        Me.scMain.SplitterWidth = 5
        Me.scMain.TabIndex = 1

        ' ── scMain.Panel1 → gbFornitori ──────────────────────
        Me.scMain.Panel1.Controls.Add(Me.gbFornitori)

        Me.gbFornitori.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbFornitori.Name = "gbFornitori"
        Me.gbFornitori.Text = "Fornitori con OA aperti"
        Me.gbFornitori.TabIndex = 0
        Me.gbFornitori.Controls.Add(Me.lvFornitori)
        Me.gbFornitori.Controls.Add(Me.pnlFornBottom)

        Me.lvFornitori.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lvFornitori.Name = "lvFornitori"
        Me.lvFornitori.TabIndex = 0

        ' pnlFornBottom — contiene conteggio + pulsante toggle sollecito
        Me.pnlFornBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFornBottom.Height = 28
        Me.pnlFornBottom.Name = "pnlFornBottom"
        Me.pnlFornBottom.Controls.Add(Me.lblConteggioFornitori)
        Me.pnlFornBottom.Controls.Add(Me.btnToggleSollecito)

        Me.lblConteggioFornitori.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblConteggioFornitori.Name = "lblConteggioFornitori"
        Me.lblConteggioFornitori.Text = ""
        Me.lblConteggioFornitori.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblConteggioFornitori.Padding = New System.Windows.Forms.Padding(4, 0, 0, 0)
        Me.lblConteggioFornitori.ForeColor = System.Drawing.Color.DimGray

        Me.btnToggleSollecito.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnToggleSollecito.Width = 175
        Me.btnToggleSollecito.Name = "btnToggleSollecito"
        Me.btnToggleSollecito.Text = "★ Aggiungi/Rimuovi sollecito"
        Me.btnToggleSollecito.TabIndex = 1
        Me.btnToggleSollecito.BackColor = System.Drawing.Color.DarkGreen
        Me.btnToggleSollecito.ForeColor = System.Drawing.Color.White
        Me.btnToggleSollecito.UseVisualStyleBackColor = False
        Me.btnToggleSollecito.Font = New System.Drawing.Font("Segoe UI", 8.5F)

        ' ── scMain.Panel2 → gbOrdini ─────────────────────────
        Me.scMain.Panel2.Controls.Add(Me.gbOrdini)

        Me.gbOrdini.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbOrdini.Name = "gbOrdini"
        Me.gbOrdini.Text = "Ordini del fornitore selezionato"
        Me.gbOrdini.TabIndex = 0
        ' Add dgvOrdini first (Fill), then pnlBtnsOrdini (Bottom - higher z, docks first)
        Me.gbOrdini.Controls.Add(Me.dgvOrdini)
        Me.gbOrdini.Controls.Add(Me.pnlBtnsOrdini)

        ' pnlBtnsOrdini
        Me.pnlBtnsOrdini.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBtnsOrdini.Height = 34
        Me.pnlBtnsOrdini.Name = "pnlBtnsOrdini"
        Me.pnlBtnsOrdini.Controls.Add(Me.btnSelTutti)
        Me.pnlBtnsOrdini.Controls.Add(Me.btnDeselTutti)
        Me.pnlBtnsOrdini.Controls.Add(Me.lblConteggio)

        Me.btnSelTutti.Location = New System.Drawing.Point(3, 4)
        Me.btnSelTutti.Name = "btnSelTutti"
        Me.btnSelTutti.Size = New System.Drawing.Size(105, 25)
        Me.btnSelTutti.TabIndex = 0
        Me.btnSelTutti.Text = "Seleziona tutti"

        Me.btnDeselTutti.Location = New System.Drawing.Point(113, 4)
        Me.btnDeselTutti.Name = "btnDeselTutti"
        Me.btnDeselTutti.Size = New System.Drawing.Size(118, 25)
        Me.btnDeselTutti.TabIndex = 1
        Me.btnDeselTutti.Text = "Deseleziona tutti"

        Me.lblConteggio.AutoSize = False
        Me.lblConteggio.Location = New System.Drawing.Point(240, 8)
        Me.lblConteggio.Name = "lblConteggio"
        Me.lblConteggio.Size = New System.Drawing.Size(350, 18)
        Me.lblConteggio.ForeColor = System.Drawing.Color.DimGray
        Me.lblConteggio.Text = ""

        ' dgvOrdini
        Me.dgvOrdini.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvOrdini.Name = "dgvOrdini"
        Me.dgvOrdini.TabIndex = 0
        Me.dgvOrdini.BackgroundColor = System.Drawing.Color.White
        Me.dgvOrdini.BorderStyle = System.Windows.Forms.BorderStyle.None

        ' ── gbStatistiche ────────────────────────────────────
        Me.gbStatistiche.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbStatistiche.Margin = New System.Windows.Forms.Padding(4, 2, 4, 2)
        Me.gbStatistiche.Name = "gbStatistiche"
        Me.gbStatistiche.Text = "Statistiche scaduti"
        Me.gbStatistiche.TabIndex = 3
        Me.gbStatistiche.Controls.Add(Me.scStatistiche)
        Me.gbStatistiche.Controls.Add(Me.pnlStatTop)

        Me.pnlStatTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlStatTop.Height = 26
        Me.pnlStatTop.Name = "pnlStatTop"
        Me.pnlStatTop.BackColor = System.Drawing.Color.FromArgb(235, 242, 250)
        Me.pnlStatTop.Controls.Add(Me.lblStatGenerale)

        Me.lblStatGenerale.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblStatGenerale.Name = "lblStatGenerale"
        Me.lblStatGenerale.Text = ""
        Me.lblStatGenerale.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblStatGenerale.Padding = New System.Windows.Forms.Padding(6, 0, 0, 0)
        Me.lblStatGenerale.Font = New System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Bold)
        Me.lblStatGenerale.ForeColor = System.Drawing.Color.DarkRed

        ' scStatistiche
        Me.scStatistiche.Dock = System.Windows.Forms.DockStyle.Fill
        Me.scStatistiche.Name = "scStatistiche"
        Me.scStatistiche.Orientation = System.Windows.Forms.Orientation.Vertical
        Me.scStatistiche.SplitterDistance = 400
        Me.scStatistiche.SplitterWidth = 4
        Me.scStatistiche.TabIndex = 1

        Me.scStatistiche.Panel1.Controls.Add(Me.dgvStatAcquisitore)
        Me.scStatistiche.Panel2.Controls.Add(Me.dgvLog)
        Me.scStatistiche.Panel2.Controls.Add(Me.pnlLogTop)

        Me.dgvStatAcquisitore.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvStatAcquisitore.Name = "dgvStatAcquisitore"
        Me.dgvStatAcquisitore.TabIndex = 0
        Me.dgvStatAcquisitore.BackgroundColor = System.Drawing.Color.White
        Me.dgvStatAcquisitore.BorderStyle = System.Windows.Forms.BorderStyle.None

        ' pnlLogTop
        Me.pnlLogTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLogTop.Height = 28
        Me.pnlLogTop.Name = "pnlLogTop"
        Me.pnlLogTop.Controls.Add(Me.btnInviaReport)
        Me.pnlLogTop.Controls.Add(Me.btnAggiornaLog)

        Me.btnAggiornaLog.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnAggiornaLog.Width = 120
        Me.btnAggiornaLog.Name = "btnAggiornaLog"
        Me.btnAggiornaLog.Text = "↻ Aggiorna log"
        Me.btnAggiornaLog.TabIndex = 0
        Me.btnAggiornaLog.UseVisualStyleBackColor = True

        Me.btnInviaReport.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnInviaReport.Width = 150
        Me.btnInviaReport.Name = "btnInviaReport"
        Me.btnInviaReport.Text = "📧 Invia report settimanale"
        Me.btnInviaReport.TabIndex = 1
        Me.btnInviaReport.UseVisualStyleBackColor = True

        ' dgvLog
        Me.dgvLog.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvLog.Name = "dgvLog"
        Me.dgvLog.TabIndex = 1
        Me.dgvLog.BackgroundColor = System.Drawing.Color.White
        Me.dgvLog.BorderStyle = System.Windows.Forms.BorderStyle.None

        ' ── gbAnteprima ──────────────────────────────────────
        Me.gbAnteprima.Dock = System.Windows.Forms.DockStyle.Fill
        Me.gbAnteprima.Margin = New System.Windows.Forms.Padding(4, 2, 4, 3)
        Me.gbAnteprima.Name = "gbAnteprima"
        Me.gbAnteprima.Text = "Anteprima mail"
        Me.gbAnteprima.TabIndex = 2
        ' Add rtbAnteprima first (Fill, lowest z), then pnlBtnsAnteprima (Right), then pnlMailHeader (Top, highest z)
        Me.gbAnteprima.Controls.Add(Me.rtbAnteprima)
        Me.gbAnteprima.Controls.Add(Me.pnlBtnsAnteprima)
        Me.gbAnteprima.Controls.Add(Me.pnlMailHeader)

        ' pnlMailHeader
        Me.pnlMailHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMailHeader.Height = 30
        Me.pnlMailHeader.Name = "pnlMailHeader"
        Me.pnlMailHeader.Controls.Add(Me.lblA)
        Me.pnlMailHeader.Controls.Add(Me.txtEmail)
        Me.pnlMailHeader.Controls.Add(Me.lblOggetto)
        Me.pnlMailHeader.Controls.Add(Me.txtOggetto)

        Me.lblA.AutoSize = True
        Me.lblA.Location = New System.Drawing.Point(3, 7)
        Me.lblA.Name = "lblA"
        Me.lblA.Text = "A:"

        Me.txtEmail.Location = New System.Drawing.Point(22, 4)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(360, 22)
        Me.txtEmail.TabIndex = 0

        Me.lblOggetto.AutoSize = True
        Me.lblOggetto.Location = New System.Drawing.Point(392, 7)
        Me.lblOggetto.Name = "lblOggetto"
        Me.lblOggetto.Text = "Oggetto:"

        Me.txtOggetto.Location = New System.Drawing.Point(450, 4)
        Me.txtOggetto.Name = "txtOggetto"
        Me.txtOggetto.Size = New System.Drawing.Size(700, 22)
        Me.txtOggetto.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right Or System.Windows.Forms.AnchorStyles.Top
        Me.txtOggetto.TabIndex = 1

        ' pnlBtnsAnteprima
        Me.pnlBtnsAnteprima.Dock = System.Windows.Forms.DockStyle.Right
        Me.pnlBtnsAnteprima.Width = 178
        Me.pnlBtnsAnteprima.Name = "pnlBtnsAnteprima"
        Me.pnlBtnsAnteprima.Padding = New System.Windows.Forms.Padding(4)
        Me.pnlBtnsAnteprima.Controls.Add(Me.btnAggiornaAnteprima)
        Me.pnlBtnsAnteprima.Controls.Add(Me.btnPreparaMail)
        Me.pnlBtnsAnteprima.Controls.Add(Me.btnTutteMail)

        Me.btnAggiornaAnteprima.Location = New System.Drawing.Point(5, 5)
        Me.btnAggiornaAnteprima.Name = "btnAggiornaAnteprima"
        Me.btnAggiornaAnteprima.Size = New System.Drawing.Size(167, 28)
        Me.btnAggiornaAnteprima.TabIndex = 0
        Me.btnAggiornaAnteprima.Text = "Aggiorna anteprima"

        Me.btnPreparaMail.Location = New System.Drawing.Point(5, 40)
        Me.btnPreparaMail.Name = "btnPreparaMail"
        Me.btnPreparaMail.Size = New System.Drawing.Size(167, 35)
        Me.btnPreparaMail.TabIndex = 1
        Me.btnPreparaMail.Text = "Prepara mail in Outlook"
        Me.btnPreparaMail.BackColor = System.Drawing.Color.SteelBlue
        Me.btnPreparaMail.ForeColor = System.Drawing.Color.White
        Me.btnPreparaMail.UseVisualStyleBackColor = False

        Me.btnTutteMail.Location = New System.Drawing.Point(5, 83)
        Me.btnTutteMail.Name = "btnTutteMail"
        Me.btnTutteMail.Size = New System.Drawing.Size(167, 28)
        Me.btnTutteMail.TabIndex = 2
        Me.btnTutteMail.Text = "Prepara mail selezione..."

        ' rtbAnteprima
        Me.rtbAnteprima.Dock = System.Windows.Forms.DockStyle.Fill
        Me.rtbAnteprima.Name = "rtbAnteprima"
        Me.rtbAnteprima.Font = New System.Drawing.Font("Courier New", 9.0!)
        Me.rtbAnteprima.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.rtbAnteprima.TabIndex = 0
        Me.rtbAnteprima.ReadOnly = True
        Me.rtbAnteprima.BackColor = System.Drawing.Color.White

        ' txbStatus
        Me.txbStatus.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.txbStatus.ReadOnly = True
        Me.txbStatus.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txbStatus.BackColor = System.Drawing.Color.FromArgb(240, 240, 240)
        Me.txbStatus.Font = New System.Drawing.Font("Segoe UI", 8.5!)
        Me.txbStatus.Height = 20
        Me.txbStatus.TabStop = False
        Me.txbStatus.Name = "txbStatus"

        ' ── Form ─────────────────────────────────────────────
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1400, 860)
        Me.Controls.Add(Me.tlpMain)
        Me.Controls.Add(Me.txbStatus)
        Me.MinimumSize = New System.Drawing.Size(1050, 650)
        Me.Name = "Solleciti_OA"
        Me.Text = "Solleciti Ordini di Acquisto"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized

        Me.tlpMain.ResumeLayout(False)
        Me.gbFiltri.ResumeLayout(False)
        Me.gbFiltri.PerformLayout()
        Me.gbStatistiche.ResumeLayout(False)
        Me.pnlStatTop.ResumeLayout(False)
        Me.scStatistiche.Panel1.ResumeLayout(False)
        Me.scStatistiche.Panel2.ResumeLayout(False)
        CType(Me.scStatistiche, System.ComponentModel.ISupportInitialize).EndInit()
        Me.scStatistiche.ResumeLayout(False)
        CType(Me.dgvStatAcquisitore, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlLogTop.ResumeLayout(False)
        CType(Me.dgvLog, System.ComponentModel.ISupportInitialize).EndInit()
        Me.scMain.Panel1.ResumeLayout(False)
        Me.scMain.Panel2.ResumeLayout(False)
        CType(Me.scMain, System.ComponentModel.ISupportInitialize).EndInit()
        Me.scMain.ResumeLayout(False)
        Me.gbFornitori.ResumeLayout(False)
        Me.gbOrdini.ResumeLayout(False)
        Me.pnlBtnsOrdini.ResumeLayout(False)
        CType(Me.dgvOrdini, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gbAnteprima.ResumeLayout(False)
        Me.pnlMailHeader.ResumeLayout(False)
        Me.pnlMailHeader.PerformLayout()
        Me.pnlBtnsAnteprima.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    Friend WithEvents txbStatus As System.Windows.Forms.TextBox
    Friend WithEvents tlpMain As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents gbFiltri As System.Windows.Forms.GroupBox
    Friend WithEvents lblFiltroForn As System.Windows.Forms.Label
    Friend WithEvents cmbFiltroFornitore As System.Windows.Forms.ComboBox
    Friend WithEvents lblFiltroComm As System.Windows.Forms.Label
    Friend WithEvents txtFiltroCommessa As System.Windows.Forms.TextBox
    Friend WithEvents chkSoloScaduti As System.Windows.Forms.CheckBox
    Friend WithEvents chkSoloSollecito As System.Windows.Forms.CheckBox
    Friend WithEvents btnCarica As System.Windows.Forms.Button
    Friend WithEvents lblStato As System.Windows.Forms.Label
    Friend WithEvents scMain As System.Windows.Forms.SplitContainer
    Friend WithEvents gbFornitori As System.Windows.Forms.GroupBox
    Friend WithEvents lvFornitori As System.Windows.Forms.ListView
    Friend WithEvents pnlFornBottom As System.Windows.Forms.Panel
    Friend WithEvents lblConteggioFornitori As System.Windows.Forms.Label
    Friend WithEvents btnToggleSollecito As System.Windows.Forms.Button
    Friend WithEvents gbOrdini As System.Windows.Forms.GroupBox
    Friend WithEvents pnlBtnsOrdini As System.Windows.Forms.Panel
    Friend WithEvents btnSelTutti As System.Windows.Forms.Button
    Friend WithEvents btnDeselTutti As System.Windows.Forms.Button
    Friend WithEvents lblConteggio As System.Windows.Forms.Label
    Friend WithEvents dgvOrdini As System.Windows.Forms.DataGridView
    Friend WithEvents gbAnteprima As System.Windows.Forms.GroupBox
    Friend WithEvents pnlMailHeader As System.Windows.Forms.Panel
    Friend WithEvents lblA As System.Windows.Forms.Label
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents lblOggetto As System.Windows.Forms.Label
    Friend WithEvents txtOggetto As System.Windows.Forms.TextBox
    Friend WithEvents pnlBtnsAnteprima As System.Windows.Forms.Panel
    Friend WithEvents btnAggiornaAnteprima As System.Windows.Forms.Button
    Friend WithEvents btnPreparaMail As System.Windows.Forms.Button
    Friend WithEvents btnTutteMail As System.Windows.Forms.Button
    Friend WithEvents rtbAnteprima As System.Windows.Forms.RichTextBox
    Friend WithEvents gbStatistiche As System.Windows.Forms.GroupBox
    Friend WithEvents pnlStatTop As System.Windows.Forms.Panel
    Friend WithEvents lblStatGenerale As System.Windows.Forms.Label
    Friend WithEvents dgvStatAcquisitore As System.Windows.Forms.DataGridView
    Friend WithEvents lblFiltroAcq As System.Windows.Forms.Label
    Friend WithEvents cmbFiltroAcquisitore As System.Windows.Forms.ComboBox
    Friend WithEvents scStatistiche As System.Windows.Forms.SplitContainer
    Friend WithEvents pnlLogTop As System.Windows.Forms.Panel
    Friend WithEvents btnAggiornaLog As System.Windows.Forms.Button
    Friend WithEvents btnInviaReport As System.Windows.Forms.Button
    Friend WithEvents dgvLog As System.Windows.Forms.DataGridView

End Class
