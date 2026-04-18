<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Entrate_merci_storico
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()>
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
        Me.pnlTop = New System.Windows.Forms.Panel()
        Me.lblDal = New System.Windows.Forms.Label()
        Me.dtpDal = New System.Windows.Forms.DateTimePicker()
        Me.lblAl = New System.Windows.Forms.Label()
        Me.dtpAl = New System.Windows.Forms.DateTimePicker()
        Me.lblFornitore = New System.Windows.Forms.Label()
        Me.txtFornitore = New System.Windows.Forms.TextBox()
        Me.lblCodice = New System.Windows.Forms.Label()
        Me.txtCodice = New System.Windows.Forms.TextBox()
        Me.lblOrdine = New System.Windows.Forms.Label()
        Me.txtOrdine = New System.Windows.Forms.TextBox()
        Me.lblStato = New System.Windows.Forms.Label()
        Me.cmbStato = New System.Windows.Forms.ComboBox()
        Me.lblDipendente = New System.Windows.Forms.Label()
        Me.txtDipendente = New System.Windows.Forms.TextBox()
        Me.lblBollaID = New System.Windows.Forms.Label()
        Me.txtBollaID = New System.Windows.Forms.TextBox()
        Me.lblDDT = New System.Windows.Forms.Label()
        Me.txtDDT = New System.Windows.Forms.TextBox()
        Me.btnCerca = New System.Windows.Forms.Button()
        Me.pnlBottom = New System.Windows.Forms.Panel()
        Me.lblCount = New System.Windows.Forms.Label()
        Me.btnElimina = New System.Windows.Forms.Button()
        Me.btnChiudi = New System.Windows.Forms.Button()
        Me.pnlEstratto = New System.Windows.Forms.Panel()
        Me.btnPrepara = New System.Windows.Forms.Button()
        Me.lblOrdiniTxt = New System.Windows.Forms.Label()
        Me.rtbOrdini = New System.Windows.Forms.RichTextBox()
        Me.btnCopiaOrdini = New System.Windows.Forms.Button()
        Me.lblCodiciTxt = New System.Windows.Forms.Label()
        Me.rtbCodici = New System.Windows.Forms.RichTextBox()
        Me.btnCopiaCodici = New System.Windows.Forms.Button()
        Me.dgvStorico = New System.Windows.Forms.DataGridView()
        Me.pnlTop.SuspendLayout()
        Me.pnlBottom.SuspendLayout()
        Me.pnlEstratto.SuspendLayout()
        CType(Me.dgvStorico, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.BackColor = System.Drawing.Color.FromArgb(45, 85, 140)
        Me.pnlTop.Controls.Add(Me.btnCerca)
        Me.pnlTop.Controls.Add(Me.cmbStato)
        Me.pnlTop.Controls.Add(Me.lblStato)
        Me.pnlTop.Controls.Add(Me.txtOrdine)
        Me.pnlTop.Controls.Add(Me.lblOrdine)
        Me.pnlTop.Controls.Add(Me.txtCodice)
        Me.pnlTop.Controls.Add(Me.lblCodice)
        Me.pnlTop.Controls.Add(Me.txtFornitore)
        Me.pnlTop.Controls.Add(Me.lblFornitore)
        Me.pnlTop.Controls.Add(Me.dtpAl)
        Me.pnlTop.Controls.Add(Me.lblAl)
        Me.pnlTop.Controls.Add(Me.dtpDal)
        Me.pnlTop.Controls.Add(Me.lblDal)
        Me.pnlTop.Controls.Add(Me.txtDipendente)
        Me.pnlTop.Controls.Add(Me.lblDipendente)
        Me.pnlTop.Controls.Add(Me.txtBollaID)
        Me.pnlTop.Controls.Add(Me.lblBollaID)
        Me.pnlTop.Controls.Add(Me.txtDDT)
        Me.pnlTop.Controls.Add(Me.lblDDT)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(1300, 90)
        Me.pnlTop.TabIndex = 0
        '
        'lblDal
        '
        Me.lblDal.AutoSize = True
        Me.lblDal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblDal.ForeColor = System.Drawing.Color.White
        Me.lblDal.Location = New System.Drawing.Point(10, 20)
        Me.lblDal.Name = "lblDal"
        Me.lblDal.Text = "Dal:"
        '
        'dtpDal
        '
        Me.dtpDal.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpDal.Location = New System.Drawing.Point(40, 16)
        Me.dtpDal.Name = "dtpDal"
        Me.dtpDal.Size = New System.Drawing.Size(100, 20)
        Me.dtpDal.TabIndex = 0
        '
        'lblAl
        '
        Me.lblAl.AutoSize = True
        Me.lblAl.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblAl.ForeColor = System.Drawing.Color.White
        Me.lblAl.Location = New System.Drawing.Point(148, 20)
        Me.lblAl.Name = "lblAl"
        Me.lblAl.Text = "Al:"
        '
        'dtpAl
        '
        Me.dtpAl.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpAl.Location = New System.Drawing.Point(168, 16)
        Me.dtpAl.Name = "dtpAl"
        Me.dtpAl.Size = New System.Drawing.Size(100, 20)
        Me.dtpAl.TabIndex = 1
        '
        'lblFornitore
        '
        Me.lblFornitore.AutoSize = True
        Me.lblFornitore.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblFornitore.ForeColor = System.Drawing.Color.White
        Me.lblFornitore.Location = New System.Drawing.Point(280, 20)
        Me.lblFornitore.Name = "lblFornitore"
        Me.lblFornitore.Text = "Fornitore:"
        '
        'txtFornitore
        '
        Me.txtFornitore.Location = New System.Drawing.Point(352, 16)
        Me.txtFornitore.Name = "txtFornitore"
        Me.txtFornitore.Size = New System.Drawing.Size(160, 20)
        Me.txtFornitore.TabIndex = 2
        '
        'lblCodice
        '
        Me.lblCodice.AutoSize = True
        Me.lblCodice.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblCodice.ForeColor = System.Drawing.Color.White
        Me.lblCodice.Location = New System.Drawing.Point(522, 20)
        Me.lblCodice.Name = "lblCodice"
        Me.lblCodice.Text = "Codice:"
        '
        'txtCodice
        '
        Me.txtCodice.Location = New System.Drawing.Point(578, 16)
        Me.txtCodice.Name = "txtCodice"
        Me.txtCodice.Size = New System.Drawing.Size(130, 20)
        Me.txtCodice.TabIndex = 3
        '
        'lblOrdine
        '
        Me.lblOrdine.AutoSize = True
        Me.lblOrdine.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblOrdine.ForeColor = System.Drawing.Color.White
        Me.lblOrdine.Location = New System.Drawing.Point(718, 20)
        Me.lblOrdine.Name = "lblOrdine"
        Me.lblOrdine.Text = "Ordine:"
        '
        'txtOrdine
        '
        Me.txtOrdine.Location = New System.Drawing.Point(770, 16)
        Me.txtOrdine.Name = "txtOrdine"
        Me.txtOrdine.Size = New System.Drawing.Size(100, 20)
        Me.txtOrdine.TabIndex = 4
        '
        'lblStato
        '
        Me.lblStato.AutoSize = True
        Me.lblStato.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblStato.ForeColor = System.Drawing.Color.White
        Me.lblStato.Location = New System.Drawing.Point(880, 20)
        Me.lblStato.Name = "lblStato"
        Me.lblStato.Text = "Stato:"
        '
        'cmbStato
        '
        Me.cmbStato.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbStato.Location = New System.Drawing.Point(924, 16)
        Me.cmbStato.Name = "cmbStato"
        Me.cmbStato.Size = New System.Drawing.Size(200, 21)
        Me.cmbStato.TabIndex = 5
        '
        'lblDipendente
        '
        Me.lblDipendente.AutoSize = True
        Me.lblDipendente.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblDipendente.ForeColor = System.Drawing.Color.White
        Me.lblDipendente.Location = New System.Drawing.Point(10, 58)
        Me.lblDipendente.Name = "lblDipendente"
        Me.lblDipendente.Text = "Dipendente:"
        '
        'txtDipendente
        '
        Me.txtDipendente.Location = New System.Drawing.Point(100, 54)
        Me.txtDipendente.Name = "txtDipendente"
        Me.txtDipendente.Size = New System.Drawing.Size(200, 20)
        Me.txtDipendente.TabIndex = 7
        '
        'lblBollaID
        '
        Me.lblBollaID.AutoSize = True
        Me.lblBollaID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblBollaID.ForeColor = System.Drawing.Color.White
        Me.lblBollaID.Location = New System.Drawing.Point(315, 58)
        Me.lblBollaID.Name = "lblBollaID"
        Me.lblBollaID.Text = "Bolla ID:"
        '
        'txtBollaID
        '
        Me.txtBollaID.Location = New System.Drawing.Point(385, 54)
        Me.txtBollaID.Name = "txtBollaID"
        Me.txtBollaID.Size = New System.Drawing.Size(80, 20)
        Me.txtBollaID.TabIndex = 8
        '
        'lblDDT
        '
        Me.lblDDT.AutoSize = True
        Me.lblDDT.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblDDT.ForeColor = System.Drawing.Color.White
        Me.lblDDT.Location = New System.Drawing.Point(480, 58)
        Me.lblDDT.Name = "lblDDT"
        Me.lblDDT.Text = "DDT N°:"
        '
        'txtDDT
        '
        Me.txtDDT.Location = New System.Drawing.Point(550, 54)
        Me.txtDDT.Name = "txtDDT"
        Me.txtDDT.Size = New System.Drawing.Size(100, 20)
        Me.txtDDT.TabIndex = 9
        '
        'btnCerca
        '
        Me.btnCerca.BackColor = System.Drawing.Color.FromArgb(0, 150, 80)
        Me.btnCerca.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnCerca.ForeColor = System.Drawing.Color.White
        Me.btnCerca.Location = New System.Drawing.Point(1138, 30)
        Me.btnCerca.Name = "btnCerca"
        Me.btnCerca.Size = New System.Drawing.Size(100, 28)
        Me.btnCerca.TabIndex = 6
        Me.btnCerca.Text = "Cerca"
        Me.btnCerca.UseVisualStyleBackColor = False
        '
        'pnlBottom
        '
        Me.pnlBottom.BackColor = System.Drawing.Color.FromArgb(240, 240, 240)
        Me.pnlBottom.Controls.Add(Me.lblCount)
        Me.pnlBottom.Controls.Add(Me.btnElimina)
        Me.pnlBottom.Controls.Add(Me.btnChiudi)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 648)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(1300, 52)
        Me.pnlBottom.TabIndex = 1
        '
        'lblCount
        '
        Me.lblCount.AutoSize = True
        Me.lblCount.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblCount.Location = New System.Drawing.Point(12, 18)
        Me.lblCount.Name = "lblCount"
        Me.lblCount.Text = ""
        '
        'btnElimina
        '
        Me.btnElimina.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnElimina.BackColor = System.Drawing.Color.FromArgb(180, 40, 40)
        Me.btnElimina.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnElimina.ForeColor = System.Drawing.Color.White
        Me.btnElimina.Location = New System.Drawing.Point(1068, 12)
        Me.btnElimina.Name = "btnElimina"
        Me.btnElimina.Size = New System.Drawing.Size(140, 28)
        Me.btnElimina.TabIndex = 1
        Me.btnElimina.Text = "Elimina selezionati"
        Me.btnElimina.UseVisualStyleBackColor = False
        '
        'btnChiudi
        '
        Me.btnChiudi.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnChiudi.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.btnChiudi.Location = New System.Drawing.Point(1218, 10)
        Me.btnChiudi.Name = "btnChiudi"
        Me.btnChiudi.Size = New System.Drawing.Size(70, 32)
        Me.btnChiudi.TabIndex = 2
        Me.btnChiudi.Text = "Chiudi"
        Me.btnChiudi.UseVisualStyleBackColor = True
        '
        'pnlEstratto
        '
        Me.pnlEstratto.BackColor = System.Drawing.Color.FromArgb(225, 225, 225)
        Me.pnlEstratto.Controls.Add(Me.btnPrepara)
        Me.pnlEstratto.Controls.Add(Me.lblOrdiniTxt)
        Me.pnlEstratto.Controls.Add(Me.rtbOrdini)
        Me.pnlEstratto.Controls.Add(Me.btnCopiaOrdini)
        Me.pnlEstratto.Controls.Add(Me.lblCodiciTxt)
        Me.pnlEstratto.Controls.Add(Me.rtbCodici)
        Me.pnlEstratto.Controls.Add(Me.btnCopiaCodici)
        Me.pnlEstratto.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlEstratto.Name = "pnlEstratto"
        Me.pnlEstratto.Size = New System.Drawing.Size(1300, 85)
        Me.pnlEstratto.TabIndex = 3
        '
        'btnPrepara
        '
        Me.btnPrepara.BackColor = System.Drawing.Color.FromArgb(0, 110, 180)
        Me.btnPrepara.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnPrepara.ForeColor = System.Drawing.Color.White
        Me.btnPrepara.Location = New System.Drawing.Point(10, 27)
        Me.btnPrepara.Name = "btnPrepara"
        Me.btnPrepara.Size = New System.Drawing.Size(130, 30)
        Me.btnPrepara.TabIndex = 0
        Me.btnPrepara.Text = "Prepara riepilogo"
        Me.btnPrepara.UseVisualStyleBackColor = False
        '
        'lblOrdiniTxt
        '
        Me.lblOrdiniTxt.AutoSize = True
        Me.lblOrdiniTxt.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.lblOrdiniTxt.Location = New System.Drawing.Point(150, 8)
        Me.lblOrdiniTxt.Name = "lblOrdiniTxt"
        Me.lblOrdiniTxt.Text = "Ordini acquisto:"
        '
        'rtbOrdini
        '
        Me.rtbOrdini.Location = New System.Drawing.Point(150, 24)
        Me.rtbOrdini.Name = "rtbOrdini"
        Me.rtbOrdini.Size = New System.Drawing.Size(478, 54)
        Me.rtbOrdini.TabIndex = 1
        Me.rtbOrdini.Text = ""
        '
        'btnCopiaOrdini
        '
        Me.btnCopiaOrdini.Location = New System.Drawing.Point(632, 30)
        Me.btnCopiaOrdini.Name = "btnCopiaOrdini"
        Me.btnCopiaOrdini.Size = New System.Drawing.Size(26, 42)
        Me.btnCopiaOrdini.TabIndex = 10
        Me.btnCopiaOrdini.Text = "C"
        Me.btnCopiaOrdini.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.btnCopiaOrdini.UseVisualStyleBackColor = True
        '
        'lblCodiciTxt
        '
        Me.lblCodiciTxt.AutoSize = True
        Me.lblCodiciTxt.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.lblCodiciTxt.Location = New System.Drawing.Point(670, 8)
        Me.lblCodiciTxt.Name = "lblCodiciTxt"
        Me.lblCodiciTxt.Text = "Codici articolo:"
        '
        'rtbCodici
        '
        Me.rtbCodici.Location = New System.Drawing.Point(670, 24)
        Me.rtbCodici.Name = "rtbCodici"
        Me.rtbCodici.Size = New System.Drawing.Size(590, 54)
        Me.rtbCodici.TabIndex = 2
        Me.rtbCodici.Text = ""
        '
        'btnCopiaCodici
        '
        Me.btnCopiaCodici.Location = New System.Drawing.Point(1264, 30)
        Me.btnCopiaCodici.Name = "btnCopiaCodici"
        Me.btnCopiaCodici.Size = New System.Drawing.Size(26, 42)
        Me.btnCopiaCodici.TabIndex = 11
        Me.btnCopiaCodici.Text = "C"
        Me.btnCopiaCodici.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.btnCopiaCodici.UseVisualStyleBackColor = True
        '
        'dgvStorico
        '
        Me.dgvStorico.AllowUserToAddRows = False
        Me.dgvStorico.AllowUserToDeleteRows = False
        Me.dgvStorico.BackgroundColor = System.Drawing.Color.White
        Me.dgvStorico.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgvStorico.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvStorico.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvStorico.Location = New System.Drawing.Point(0, 90)
        Me.dgvStorico.Name = "dgvStorico"
        Me.dgvStorico.ReadOnly = True
        Me.dgvStorico.RowHeadersWidth = 30
        Me.dgvStorico.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvStorico.MultiSelect = True
        Me.dgvStorico.Size = New System.Drawing.Size(1300, 558)
        Me.dgvStorico.TabIndex = 2
        '
        'Entrate_merci_storico
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1300, 700)
        Me.Controls.Add(Me.dgvStorico)
        Me.Controls.Add(Me.pnlEstratto)
        Me.Controls.Add(Me.pnlBottom)
        Me.Controls.Add(Me.pnlTop)
        Me.MinimumSize = New System.Drawing.Size(1000, 600)
        Me.Name = "Entrate_merci_storico"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Storico Entrate Merce"
        Me.pnlTop.ResumeLayout(False)
        Me.pnlTop.PerformLayout()
        Me.pnlBottom.ResumeLayout(False)
        Me.pnlBottom.PerformLayout()
        Me.pnlEstratto.ResumeLayout(False)
        Me.pnlEstratto.PerformLayout()
        CType(Me.dgvStorico, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents pnlTop As Panel
    Friend WithEvents lblDal As Label
    Friend WithEvents dtpDal As DateTimePicker
    Friend WithEvents lblAl As Label
    Friend WithEvents dtpAl As DateTimePicker
    Friend WithEvents lblFornitore As Label
    Friend WithEvents txtFornitore As TextBox
    Friend WithEvents lblCodice As Label
    Friend WithEvents txtCodice As TextBox
    Friend WithEvents lblOrdine As Label
    Friend WithEvents txtOrdine As TextBox
    Friend WithEvents lblStato As Label
    Friend WithEvents cmbStato As ComboBox
    Friend WithEvents lblDipendente As Label
    Friend WithEvents txtDipendente As TextBox
    Friend WithEvents lblBollaID As Label
    Friend WithEvents txtBollaID As TextBox
    Friend WithEvents lblDDT As Label
    Friend WithEvents txtDDT As TextBox
    Friend WithEvents btnCerca As Button
    Friend WithEvents pnlBottom As Panel
    Friend WithEvents lblCount As Label
    Friend WithEvents btnElimina As Button
    Friend WithEvents btnChiudi As Button
    Friend WithEvents pnlEstratto As Panel
    Friend WithEvents btnPrepara As Button
    Friend WithEvents lblOrdiniTxt As Label
    Friend WithEvents rtbOrdini As RichTextBox
    Friend WithEvents btnCopiaOrdini As Button
    Friend WithEvents lblCodiciTxt As Label
    Friend WithEvents rtbCodici As RichTextBox
    Friend WithEvents btnCopiaCodici As Button
    Friend WithEvents dgvStorico As DataGridView
End Class
