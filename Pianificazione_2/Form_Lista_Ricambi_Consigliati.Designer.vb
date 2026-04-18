<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form_Lista_Ricambi_Consigliati
    Inherits System.Windows.Forms.Form

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.panelHeader = New System.Windows.Forms.Panel()
        Me.lblCommessa = New System.Windows.Forms.Label()
        Me.lblMoltiplicatore = New System.Windows.Forms.Label()
        Me.cmbMoltiplicatore = New System.Windows.Forms.ComboBox()
        Me.lblNomeLista = New System.Windows.Forms.Label()
        Me.cmbNomeLista = New System.Windows.Forms.ComboBox()
        Me.btnNuovaLista = New System.Windows.Forms.Button()
        Me.dgvRicambi = New System.Windows.Forms.DataGridView()
        Me.panelBottom = New System.Windows.Forms.Panel()
        Me.lblTotale = New System.Windows.Forms.Label()
        Me.panelButtons = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnAggiungi = New System.Windows.Forms.Button()
        Me.btnElimina = New System.Windows.Forms.Button()
        Me.btnExportExcel = New System.Windows.Forms.Button()
        Me.btnExportPdf = New System.Windows.Forms.Button()
        Me.btnSalva = New System.Windows.Forms.Button()
        Me.btnChiudi = New System.Windows.Forms.Button()
        Me.panelHeader.SuspendLayout()
        CType(Me.dgvRicambi, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panelBottom.SuspendLayout()
        Me.panelButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'panelHeader
        '
        Me.panelHeader.BackColor = System.Drawing.Color.FromArgb(22, 45, 84)
        Me.panelHeader.Controls.Add(Me.btnNuovaLista)
        Me.panelHeader.Controls.Add(Me.cmbNomeLista)
        Me.panelHeader.Controls.Add(Me.lblNomeLista)
        Me.panelHeader.Controls.Add(Me.cmbMoltiplicatore)
        Me.panelHeader.Controls.Add(Me.lblMoltiplicatore)
        Me.panelHeader.Controls.Add(Me.lblCommessa)
        Me.panelHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.panelHeader.Location = New System.Drawing.Point(0, 0)
        Me.panelHeader.Name = "panelHeader"
        Me.panelHeader.Size = New System.Drawing.Size(1050, 80)
        Me.panelHeader.TabIndex = 0
        '
        'lblCommessa
        '
        Me.lblCommessa.AutoSize = True
        Me.lblCommessa.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Bold)
        Me.lblCommessa.ForeColor = System.Drawing.Color.White
        Me.lblCommessa.Location = New System.Drawing.Point(12, 10)
        Me.lblCommessa.Name = "lblCommessa"
        Me.lblCommessa.Size = New System.Drawing.Size(200, 20)
        Me.lblCommessa.TabIndex = 0
        Me.lblCommessa.Text = "Commessa: —"
        '
        'lblMoltiplicatore
        '
        Me.lblMoltiplicatore.AutoSize = True
        Me.lblMoltiplicatore.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblMoltiplicatore.ForeColor = System.Drawing.Color.White
        Me.lblMoltiplicatore.Location = New System.Drawing.Point(450, 13)
        Me.lblMoltiplicatore.Name = "lblMoltiplicatore"
        Me.lblMoltiplicatore.Size = New System.Drawing.Size(78, 15)
        Me.lblMoltiplicatore.TabIndex = 1
        Me.lblMoltiplicatore.Text = "Moltiplicatore:"
        '
        'cmbMoltiplicatore
        '
        Me.cmbMoltiplicatore.BackColor = System.Drawing.Color.White
        Me.cmbMoltiplicatore.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.cmbMoltiplicatore.Location = New System.Drawing.Point(540, 9)
        Me.cmbMoltiplicatore.Name = "cmbMoltiplicatore"
        Me.cmbMoltiplicatore.Size = New System.Drawing.Size(80, 25)
        Me.cmbMoltiplicatore.TabIndex = 2
        '
        'lblNomeLista
        '
        Me.lblNomeLista.AutoSize = True
        Me.lblNomeLista.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.lblNomeLista.ForeColor = System.Drawing.Color.FromArgb(200, 220, 255)
        Me.lblNomeLista.Location = New System.Drawing.Point(12, 50)
        Me.lblNomeLista.Name = "lblNomeLista"
        Me.lblNomeLista.Size = New System.Drawing.Size(56, 15)
        Me.lblNomeLista.TabIndex = 3
        Me.lblNomeLista.Text = "Lista:"
        '
        'cmbNomeLista
        '
        Me.cmbNomeLista.BackColor = System.Drawing.Color.White
        Me.cmbNomeLista.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.cmbNomeLista.Location = New System.Drawing.Point(60, 47)
        Me.cmbNomeLista.Name = "cmbNomeLista"
        Me.cmbNomeLista.Size = New System.Drawing.Size(220, 23)
        Me.cmbNomeLista.TabIndex = 4
        '
        'btnNuovaLista
        '
        Me.btnNuovaLista.BackColor = System.Drawing.Color.FromArgb(60, 90, 150)
        Me.btnNuovaLista.FlatAppearance.BorderSize = 0
        Me.btnNuovaLista.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnNuovaLista.Font = New System.Drawing.Font("Segoe UI", 8.5!)
        Me.btnNuovaLista.ForeColor = System.Drawing.Color.White
        Me.btnNuovaLista.Location = New System.Drawing.Point(290, 46)
        Me.btnNuovaLista.Name = "btnNuovaLista"
        Me.btnNuovaLista.Size = New System.Drawing.Size(110, 25)
        Me.btnNuovaLista.TabIndex = 5
        Me.btnNuovaLista.Text = "+ Nuova lista"
        Me.btnNuovaLista.UseVisualStyleBackColor = False
        '
        'dgvRicambi
        '
        Me.dgvRicambi.AllowUserToAddRows = False
        Me.dgvRicambi.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvRicambi.Location = New System.Drawing.Point(0, 80)
        Me.dgvRicambi.Name = "dgvRicambi"
        Me.dgvRicambi.RowHeadersVisible = False
        Me.dgvRicambi.Size = New System.Drawing.Size(1050, 490)
        Me.dgvRicambi.TabIndex = 1
        '
        'panelBottom
        '
        Me.panelBottom.BackColor = System.Drawing.Color.FromArgb(235, 240, 250)
        Me.panelBottom.Controls.Add(Me.lblTotale)
        Me.panelBottom.Controls.Add(Me.panelButtons)
        Me.panelBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.panelBottom.Location = New System.Drawing.Point(0, 570)
        Me.panelBottom.Name = "panelBottom"
        Me.panelBottom.Size = New System.Drawing.Size(1050, 50)
        Me.panelBottom.TabIndex = 2
        '
        'lblTotale
        '
        Me.lblTotale.AutoSize = True
        Me.lblTotale.Font = New System.Drawing.Font("Segoe UI", 11.0!, System.Drawing.FontStyle.Bold)
        Me.lblTotale.ForeColor = System.Drawing.Color.FromArgb(22, 45, 84)
        Me.lblTotale.Location = New System.Drawing.Point(12, 14)
        Me.lblTotale.Name = "lblTotale"
        Me.lblTotale.Size = New System.Drawing.Size(200, 20)
        Me.lblTotale.TabIndex = 0
        Me.lblTotale.Text = "TOTALE: € 0,00"
        '
        'panelButtons
        '
        Me.panelButtons.Anchor = CType(System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right, System.Windows.Forms.AnchorStyles)
        Me.panelButtons.Controls.Add(Me.btnAggiungi)
        Me.panelButtons.Controls.Add(Me.btnElimina)
        Me.panelButtons.Controls.Add(Me.btnExportExcel)
        Me.panelButtons.Controls.Add(Me.btnExportPdf)
        Me.panelButtons.Controls.Add(Me.btnSalva)
        Me.panelButtons.Controls.Add(Me.btnChiudi)
        Me.panelButtons.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft
        Me.panelButtons.Location = New System.Drawing.Point(390, 8)
        Me.panelButtons.Name = "panelButtons"
        Me.panelButtons.Size = New System.Drawing.Size(650, 35)
        Me.panelButtons.TabIndex = 1
        '
        'btnAggiungi
        '
        Me.btnAggiungi.BackColor = System.Drawing.Color.FromArgb(60, 130, 60)
        Me.btnAggiungi.FlatAppearance.BorderSize = 0
        Me.btnAggiungi.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnAggiungi.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.btnAggiungi.ForeColor = System.Drawing.Color.White
        Me.btnAggiungi.Location = New System.Drawing.Point(3, 3)
        Me.btnAggiungi.Name = "btnAggiungi"
        Me.btnAggiungi.Size = New System.Drawing.Size(100, 28)
        Me.btnAggiungi.TabIndex = 0
        Me.btnAggiungi.Text = "+ Aggiungi riga"
        Me.btnAggiungi.UseVisualStyleBackColor = False
        '
        'btnElimina
        '
        Me.btnElimina.BackColor = System.Drawing.Color.FromArgb(180, 60, 60)
        Me.btnElimina.FlatAppearance.BorderSize = 0
        Me.btnElimina.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnElimina.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.btnElimina.ForeColor = System.Drawing.Color.White
        Me.btnElimina.Location = New System.Drawing.Point(109, 3)
        Me.btnElimina.Name = "btnElimina"
        Me.btnElimina.Size = New System.Drawing.Size(90, 28)
        Me.btnElimina.TabIndex = 1
        Me.btnElimina.Text = "Elimina riga"
        Me.btnElimina.UseVisualStyleBackColor = False
        '
        'btnExportExcel
        '
        Me.btnExportExcel.BackColor = System.Drawing.Color.FromArgb(32, 120, 60)
        Me.btnExportExcel.FlatAppearance.BorderSize = 0
        Me.btnExportExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnExportExcel.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.btnExportExcel.ForeColor = System.Drawing.Color.White
        Me.btnExportExcel.Location = New System.Drawing.Point(205, 3)
        Me.btnExportExcel.Name = "btnExportExcel"
        Me.btnExportExcel.Size = New System.Drawing.Size(95, 28)
        Me.btnExportExcel.TabIndex = 2
        Me.btnExportExcel.Text = "Export Excel"
        Me.btnExportExcel.UseVisualStyleBackColor = False
        '
        'btnExportPdf
        '
        Me.btnExportPdf.BackColor = System.Drawing.Color.FromArgb(170, 50, 50)
        Me.btnExportPdf.FlatAppearance.BorderSize = 0
        Me.btnExportPdf.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnExportPdf.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.btnExportPdf.ForeColor = System.Drawing.Color.White
        Me.btnExportPdf.Location = New System.Drawing.Point(306, 3)
        Me.btnExportPdf.Name = "btnExportPdf"
        Me.btnExportPdf.Size = New System.Drawing.Size(95, 28)
        Me.btnExportPdf.TabIndex = 3
        Me.btnExportPdf.Text = "Offerta PDF"
        Me.btnExportPdf.UseVisualStyleBackColor = False
        '
        'btnSalva
        '
        Me.btnSalva.BackColor = System.Drawing.Color.FromArgb(22, 45, 84)
        Me.btnSalva.FlatAppearance.BorderSize = 0
        Me.btnSalva.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSalva.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnSalva.ForeColor = System.Drawing.Color.White
        Me.btnSalva.Location = New System.Drawing.Point(407, 3)
        Me.btnSalva.Name = "btnSalva"
        Me.btnSalva.Size = New System.Drawing.Size(80, 28)
        Me.btnSalva.TabIndex = 4
        Me.btnSalva.Text = "Salva"
        Me.btnSalva.UseVisualStyleBackColor = False
        '
        'btnChiudi
        '
        Me.btnChiudi.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnChiudi.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.btnChiudi.Location = New System.Drawing.Point(493, 3)
        Me.btnChiudi.Name = "btnChiudi"
        Me.btnChiudi.Size = New System.Drawing.Size(70, 28)
        Me.btnChiudi.TabIndex = 5
        Me.btnChiudi.Text = "Chiudi"
        Me.btnChiudi.UseVisualStyleBackColor = True
        '
        'Form_Lista_Ricambi_Consigliati
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1050, 620)
        Me.Controls.Add(Me.dgvRicambi)
        Me.Controls.Add(Me.panelBottom)
        Me.Controls.Add(Me.panelHeader)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.MinimumSize = New System.Drawing.Size(950, 550)
        Me.Name = "Form_Lista_Ricambi_Consigliati"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Lista Ricambi Consigliati"
        Me.panelHeader.ResumeLayout(False)
        Me.panelHeader.PerformLayout()
        CType(Me.dgvRicambi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panelBottom.ResumeLayout(False)
        Me.panelBottom.PerformLayout()
        Me.panelButtons.ResumeLayout(False)
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents panelHeader As System.Windows.Forms.Panel
    Friend WithEvents lblCommessa As System.Windows.Forms.Label
    Friend WithEvents lblMoltiplicatore As System.Windows.Forms.Label
    Friend WithEvents cmbMoltiplicatore As System.Windows.Forms.ComboBox
    Friend WithEvents lblNomeLista As System.Windows.Forms.Label
    Friend WithEvents cmbNomeLista As System.Windows.Forms.ComboBox
    Friend WithEvents btnNuovaLista As System.Windows.Forms.Button
    Friend WithEvents dgvRicambi As System.Windows.Forms.DataGridView
    Friend WithEvents panelBottom As System.Windows.Forms.Panel
    Friend WithEvents lblTotale As System.Windows.Forms.Label
    Friend WithEvents panelButtons As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents btnAggiungi As System.Windows.Forms.Button
    Friend WithEvents btnElimina As System.Windows.Forms.Button
    Friend WithEvents btnExportExcel As System.Windows.Forms.Button
    Friend WithEvents btnExportPdf As System.Windows.Forms.Button
    Friend WithEvents btnSalva As System.Windows.Forms.Button
    Friend WithEvents btnChiudi As System.Windows.Forms.Button

End Class
