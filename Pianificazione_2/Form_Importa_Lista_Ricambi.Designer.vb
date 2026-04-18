Partial Class Form_Importa_Lista_Ricambi
    Inherits System.Windows.Forms.Form

    Private Sub InitializeComponent()
        Me.pnlTop = New System.Windows.Forms.Panel()
        Me.lblFiltro = New System.Windows.Forms.Label()
        Me.txtFiltro = New System.Windows.Forms.TextBox()
        Me.btnCerca = New System.Windows.Forms.Button()
        Me.lblCommesse = New System.Windows.Forms.Label()
        Me.pnlCentro = New System.Windows.Forms.TableLayoutPanel()
        Me.grpCommesse = New System.Windows.Forms.GroupBox()
        Me.dgvCommesse = New System.Windows.Forms.DataGridView()
        Me.grpListe = New System.Windows.Forms.GroupBox()
        Me.lblListe = New System.Windows.Forms.Label()
        Me.dgvListe = New System.Windows.Forms.DataGridView()
        Me.pnlBottom = New System.Windows.Forms.Panel()
        Me.btnImporta = New System.Windows.Forms.Button()
        Me.btnAnnulla = New System.Windows.Forms.Button()
        Me.pnlTop.SuspendLayout()
        Me.pnlCentro.SuspendLayout()
        Me.grpCommesse.SuspendLayout()
        CType(Me.dgvCommesse, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpListe.SuspendLayout()
        CType(Me.dgvListe, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.lblFiltro)
        Me.pnlTop.Controls.Add(Me.txtFiltro)
        Me.pnlTop.Controls.Add(Me.btnCerca)
        Me.pnlTop.Controls.Add(Me.lblCommesse)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Height = 44
        Me.pnlTop.Padding = New System.Windows.Forms.Padding(8, 8, 8, 0)
        '
        'lblFiltro
        '
        Me.lblFiltro.AutoSize = True
        Me.lblFiltro.Location = New System.Drawing.Point(8, 14)
        Me.lblFiltro.Text = "Cerca commessa / macchina:"
        '
        'txtFiltro
        '
        Me.txtFiltro.Location = New System.Drawing.Point(200, 11)
        Me.txtFiltro.Size = New System.Drawing.Size(280, 22)
        '
        'btnCerca
        '
        Me.btnCerca.Location = New System.Drawing.Point(488, 10)
        Me.btnCerca.Size = New System.Drawing.Size(80, 24)
        Me.btnCerca.Text = "Cerca"
        '
        'lblCommesse
        '
        Me.lblCommesse.AutoSize = True
        Me.lblCommesse.Location = New System.Drawing.Point(580, 14)
        Me.lblCommesse.ForeColor = System.Drawing.Color.Gray
        Me.lblCommesse.Text = ""
        '
        'pnlCentro
        '
        Me.pnlCentro.ColumnCount = 2
        Me.pnlCentro.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 60.0!))
        Me.pnlCentro.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40.0!))
        Me.pnlCentro.Controls.Add(Me.grpCommesse, 0, 0)
        Me.pnlCentro.Controls.Add(Me.grpListe, 1, 0)
        Me.pnlCentro.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlCentro.Padding = New System.Windows.Forms.Padding(6)
        Me.pnlCentro.RowCount = 1
        Me.pnlCentro.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        '
        'grpCommesse
        '
        Me.grpCommesse.Controls.Add(Me.dgvCommesse)
        Me.grpCommesse.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpCommesse.Text = "Commesse con lista ricambi"
        Me.grpCommesse.Padding = New System.Windows.Forms.Padding(4)
        '
        'dgvCommesse
        '
        Me.dgvCommesse.AllowUserToAddRows = False
        Me.dgvCommesse.AllowUserToDeleteRows = False
        Me.dgvCommesse.AllowUserToResizeRows = False
        Me.dgvCommesse.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvCommesse.MultiSelect = False
        Me.dgvCommesse.ReadOnly = True
        Me.dgvCommesse.RowHeadersVisible = False
        Me.dgvCommesse.RowTemplate.Height = 20
        Me.dgvCommesse.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        '
        'grpListe
        '
        Me.grpListe.Controls.Add(Me.lblListe)
        Me.grpListe.Controls.Add(Me.dgvListe)
        Me.grpListe.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpListe.Text = "Liste disponibili"
        Me.grpListe.Padding = New System.Windows.Forms.Padding(4)
        '
        'lblListe
        '
        Me.lblListe.AutoSize = True
        Me.lblListe.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblListe.Font = New System.Drawing.Font("Segoe UI", 8.0!, System.Drawing.FontStyle.Italic)
        Me.lblListe.ForeColor = System.Drawing.Color.Gray
        Me.lblListe.Text = "← Seleziona una commessa"
        Me.lblListe.Padding = New System.Windows.Forms.Padding(0, 2, 0, 2)
        '
        'dgvListe
        '
        Me.dgvListe.AllowUserToAddRows = False
        Me.dgvListe.AllowUserToDeleteRows = False
        Me.dgvListe.AllowUserToResizeRows = False
        Me.dgvListe.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvListe.MultiSelect = False
        Me.dgvListe.ReadOnly = True
        Me.dgvListe.RowHeadersVisible = False
        Me.dgvListe.RowTemplate.Height = 20
        Me.dgvListe.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        '
        'pnlBottom
        '
        Me.pnlBottom.Controls.Add(Me.btnImporta)
        Me.pnlBottom.Controls.Add(Me.btnAnnulla)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Height = 44
        Me.pnlBottom.Padding = New System.Windows.Forms.Padding(8)
        '
        'btnImporta
        '
        Me.btnImporta.Enabled = False
        Me.btnImporta.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnImporta.Location = New System.Drawing.Point(8, 8)
        Me.btnImporta.Size = New System.Drawing.Size(200, 28)
        Me.btnImporta.Text = "Importa lista selezionata →"
        '
        'btnAnnulla
        '
        Me.btnAnnulla.Location = New System.Drawing.Point(220, 8)
        Me.btnAnnulla.Size = New System.Drawing.Size(80, 28)
        Me.btnAnnulla.Text = "Annulla"
        '
        'Form_Importa_Lista_Ricambi
        '
        Me.Controls.Add(Me.pnlCentro)
        Me.Controls.Add(Me.pnlTop)
        Me.Controls.Add(Me.pnlBottom)
        Me.MinimumSize = New System.Drawing.Size(700, 420)
        Me.Name = "Form_Importa_Lista_Ricambi"
        Me.pnlTop.ResumeLayout(False)
        Me.pnlTop.PerformLayout()
        Me.pnlCentro.ResumeLayout(False)
        Me.grpCommesse.ResumeLayout(False)
        CType(Me.dgvCommesse, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpListe.ResumeLayout(False)
        Me.grpListe.PerformLayout()
        CType(Me.dgvListe, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlBottom.ResumeLayout(False)
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents pnlTop As Panel
    Friend WithEvents lblFiltro As Label
    Friend WithEvents txtFiltro As TextBox
    Friend WithEvents btnCerca As Button
    Friend WithEvents lblCommesse As Label
    Friend WithEvents pnlCentro As TableLayoutPanel
    Friend WithEvents grpCommesse As GroupBox
    Friend WithEvents dgvCommesse As DataGridView
    Friend WithEvents grpListe As GroupBox
    Friend WithEvents lblListe As Label
    Friend WithEvents dgvListe As DataGridView
    Friend WithEvents pnlBottom As Panel
    Friend WithEvents btnImporta As Button
    Friend WithEvents btnAnnulla As Button

End Class
