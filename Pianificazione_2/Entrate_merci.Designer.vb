<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Entrate_merci
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()> _
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

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.pnlTop = New System.Windows.Forms.Panel()
        Me.lblFile = New System.Windows.Forms.Label()
        Me.txtFilePath = New System.Windows.Forms.TextBox()
        Me.btnSfoglia = New System.Windows.Forms.Button()
        Me.lblApiKey = New System.Windows.Forms.Label()
        Me.txtApiKey = New System.Windows.Forms.TextBox()
        Me.btnAnalizza = New System.Windows.Forms.Button()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.pnlBottom = New System.Windows.Forms.Panel()
        Me.lblMsg = New System.Windows.Forms.Label()
        Me.btnChiudi = New System.Windows.Forms.Button()
        Me.btnSalva = New System.Windows.Forms.Button()
        Me.btnStorico = New System.Windows.Forms.Button()
        Me.btnDeselTutto = New System.Windows.Forms.Button()
        Me.btnSelTutto = New System.Windows.Forms.Button()
        Me.dgvRighe = New System.Windows.Forms.DataGridView()
        Me.colSel = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.colDDT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colData = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colFornitore = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colOrdine = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colCodice = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colDescrizione = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colUM = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colQuantita = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.pnlTop.SuspendLayout()
        Me.pnlBottom.SuspendLayout()
        CType(Me.dgvRighe, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.BackColor = System.Drawing.Color.FromArgb(CType(CType(45, Byte), Integer), CType(CType(85, Byte), Integer), CType(CType(140, Byte), Integer))
        Me.pnlTop.Controls.Add(Me.lblStatus)
        Me.pnlTop.Controls.Add(Me.btnAnalizza)
        Me.pnlTop.Controls.Add(Me.txtApiKey)
        Me.pnlTop.Controls.Add(Me.lblApiKey)
        Me.pnlTop.Controls.Add(Me.btnSfoglia)
        Me.pnlTop.Controls.Add(Me.txtFilePath)
        Me.pnlTop.Controls.Add(Me.lblFile)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(1184, 115)
        Me.pnlTop.TabIndex = 0
        '
        'lblFile
        '
        Me.lblFile.AutoSize = True
        Me.lblFile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblFile.ForeColor = System.Drawing.Color.White
        Me.lblFile.Location = New System.Drawing.Point(12, 15)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(90, 15)
        Me.lblFile.Text = "File DDT (PDF):"
        '
        'txtFilePath
        '
        Me.txtFilePath.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFilePath.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!)
        Me.txtFilePath.Location = New System.Drawing.Point(108, 12)
        Me.txtFilePath.Name = "txtFilePath"
        Me.txtFilePath.ReadOnly = True
        Me.txtFilePath.Size = New System.Drawing.Size(850, 21)
        Me.txtFilePath.TabIndex = 1
        '
        'btnSfoglia
        '
        Me.btnSfoglia.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSfoglia.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnSfoglia.Location = New System.Drawing.Point(966, 10)
        Me.btnSfoglia.Name = "btnSfoglia"
        Me.btnSfoglia.Size = New System.Drawing.Size(100, 26)
        Me.btnSfoglia.TabIndex = 2
        Me.btnSfoglia.Text = "Sfoglia..."
        Me.btnSfoglia.UseVisualStyleBackColor = True
        '
        'lblApiKey
        '
        Me.lblApiKey.AutoSize = True
        Me.lblApiKey.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblApiKey.ForeColor = System.Drawing.Color.White
        Me.lblApiKey.Location = New System.Drawing.Point(12, 50)
        Me.lblApiKey.Name = "lblApiKey"
        Me.lblApiKey.Size = New System.Drawing.Size(90, 15)
        Me.lblApiKey.Text = "Chiave API Claude:"
        '
        'txtApiKey
        '
        Me.txtApiKey.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtApiKey.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!)
        Me.txtApiKey.Location = New System.Drawing.Point(108, 47)
        Me.txtApiKey.Name = "txtApiKey"
        Me.txtApiKey.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.txtApiKey.Size = New System.Drawing.Size(740, 21)
        Me.txtApiKey.TabIndex = 3
        '
        'btnAnalizza
        '
        Me.btnAnalizza.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAnalizza.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(150, Byte), Integer), CType(CType(80, Byte), Integer))
        Me.btnAnalizza.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.0!, System.Drawing.FontStyle.Bold)
        Me.btnAnalizza.ForeColor = System.Drawing.Color.White
        Me.btnAnalizza.Location = New System.Drawing.Point(860, 40)
        Me.btnAnalizza.Name = "btnAnalizza"
        Me.btnAnalizza.Size = New System.Drawing.Size(206, 60)
        Me.btnAnalizza.TabIndex = 4
        Me.btnAnalizza.Text = "Analizza PDF"
        Me.btnAnalizza.UseVisualStyleBackColor = False
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Italic)
        Me.lblStatus.ForeColor = System.Drawing.Color.LightYellow
        Me.lblStatus.Location = New System.Drawing.Point(12, 85)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(50, 15)
        Me.lblStatus.Text = "Pronto."
        '
        'pnlBottom
        '
        Me.pnlBottom.BackColor = System.Drawing.Color.FromArgb(CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer), CType(CType(240, Byte), Integer))
        Me.pnlBottom.Controls.Add(Me.lblMsg)
        Me.pnlBottom.Controls.Add(Me.btnChiudi)
        Me.pnlBottom.Controls.Add(Me.btnSalva)
        Me.pnlBottom.Controls.Add(Me.btnStorico)
        Me.pnlBottom.Controls.Add(Me.btnDeselTutto)
        Me.pnlBottom.Controls.Add(Me.btnSelTutto)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 680)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(1184, 52)
        Me.pnlBottom.TabIndex = 1
        '
        'btnSelTutto
        '
        Me.btnSelTutto.Location = New System.Drawing.Point(10, 12)
        Me.btnSelTutto.Name = "btnSelTutto"
        Me.btnSelTutto.Size = New System.Drawing.Size(130, 28)
        Me.btnSelTutto.TabIndex = 0
        Me.btnSelTutto.Text = "Seleziona tutto"
        Me.btnSelTutto.UseVisualStyleBackColor = True
        '
        'btnDeselTutto
        '
        Me.btnDeselTutto.Location = New System.Drawing.Point(148, 12)
        Me.btnDeselTutto.Name = "btnDeselTutto"
        Me.btnDeselTutto.Size = New System.Drawing.Size(130, 28)
        Me.btnDeselTutto.TabIndex = 1
        Me.btnDeselTutto.Text = "Deseleziona tutto"
        Me.btnDeselTutto.UseVisualStyleBackColor = True
        '
        'btnSalva
        '
        Me.btnSalva.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSalva.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(120, Byte), Integer), CType(CType(60, Byte), Integer))
        Me.btnSalva.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.btnSalva.ForeColor = System.Drawing.Color.White
        Me.btnSalva.Location = New System.Drawing.Point(940, 10)
        Me.btnSalva.Name = "btnSalva"
        Me.btnSalva.Size = New System.Drawing.Size(140, 32)
        Me.btnSalva.TabIndex = 2
        Me.btnSalva.Text = "Salva in SQL"
        Me.btnSalva.UseVisualStyleBackColor = False
        '
        'btnStorico
        '
        Me.btnStorico.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStorico.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.btnStorico.Location = New System.Drawing.Point(786, 10)
        Me.btnStorico.Name = "btnStorico"
        Me.btnStorico.Size = New System.Drawing.Size(140, 32)
        Me.btnStorico.TabIndex = 5
        Me.btnStorico.Text = "Storico"
        Me.btnStorico.UseVisualStyleBackColor = True
        '
        'btnChiudi
        '
        Me.btnChiudi.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnChiudi.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Bold)
        Me.btnChiudi.Location = New System.Drawing.Point(1090, 10)
        Me.btnChiudi.Name = "btnChiudi"
        Me.btnChiudi.Size = New System.Drawing.Size(82, 32)
        Me.btnChiudi.TabIndex = 3
        Me.btnChiudi.Text = "Chiudi"
        Me.btnChiudi.UseVisualStyleBackColor = True
        '
        'lblMsg
        '
        Me.lblMsg.AutoSize = True
        Me.lblMsg.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblMsg.Location = New System.Drawing.Point(290, 18)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(0, 15)
        '
        'dgvRighe
        '
        Me.dgvRighe.AllowUserToAddRows = False
        Me.dgvRighe.AllowUserToDeleteRows = False
        Me.dgvRighe.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvRighe.BackgroundColor = System.Drawing.Color.White
        Me.dgvRighe.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgvRighe.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvRighe.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colSel, Me.colDDT, Me.colData, Me.colFornitore, Me.colOrdine, Me.colCodice, Me.colDescrizione, Me.colUM, Me.colQuantita})
        Me.dgvRighe.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvRighe.Location = New System.Drawing.Point(0, 115)
        Me.dgvRighe.Name = "dgvRighe"
        Me.dgvRighe.RowHeadersWidth = 30
        Me.dgvRighe.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvRighe.Size = New System.Drawing.Size(1184, 565)
        Me.dgvRighe.TabIndex = 2
        '
        'colSel
        '
        Me.colSel.HeaderText = ""
        Me.colSel.Name = "colSel"
        Me.colSel.Width = 30
        '
        'colDDT
        '
        Me.colDDT.HeaderText = "DDT N°"
        Me.colDDT.Name = "colDDT"
        Me.colDDT.Width = 80
        '
        'colData
        '
        Me.colData.HeaderText = "Data"
        Me.colData.Name = "colData"
        Me.colData.Width = 90
        '
        'colFornitore
        '
        Me.colFornitore.HeaderText = "Fornitore"
        Me.colFornitore.Name = "colFornitore"
        Me.colFornitore.Width = 200
        '
        'colOrdine
        '
        Me.colOrdine.HeaderText = "N° Ordine"
        Me.colOrdine.Name = "colOrdine"
        Me.colOrdine.Width = 100
        '
        'colCodice
        '
        Me.colCodice.HeaderText = "Codice"
        Me.colCodice.Name = "colCodice"
        Me.colCodice.Width = 120
        '
        'colDescrizione
        '
        Me.colDescrizione.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.colDescrizione.HeaderText = "Descrizione"
        Me.colDescrizione.Name = "colDescrizione"
        '
        'colUM
        '
        Me.colUM.HeaderText = "UM"
        Me.colUM.Name = "colUM"
        Me.colUM.Width = 50
        '
        'colQuantita
        '
        Me.colQuantita.HeaderText = "Quantita"
        Me.colQuantita.Name = "colQuantita"
        Me.colQuantita.Width = 80
        '
        'Entrate_merci
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1184, 732)
        Me.Controls.Add(Me.dgvRighe)
        Me.Controls.Add(Me.pnlBottom)
        Me.Controls.Add(Me.pnlTop)
        Me.MinimumSize = New System.Drawing.Size(900, 600)
        Me.Name = "Entrate_merci"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Riconoscimento Bolle di Entrata Merce"
        Me.pnlTop.ResumeLayout(False)
        Me.pnlTop.PerformLayout()
        Me.pnlBottom.ResumeLayout(False)
        Me.pnlBottom.PerformLayout()
        CType(Me.dgvRighe, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents pnlTop As Panel
    Friend WithEvents lblFile As Label
    Friend WithEvents txtFilePath As TextBox
    Friend WithEvents btnSfoglia As Button
    Friend WithEvents lblApiKey As Label
    Friend WithEvents txtApiKey As TextBox
    Friend WithEvents btnAnalizza As Button
    Friend WithEvents lblStatus As Label
    Friend WithEvents pnlBottom As Panel
    Friend WithEvents btnSelTutto As Button
    Friend WithEvents btnDeselTutto As Button
    Friend WithEvents btnSalva As Button
    Friend WithEvents btnStorico As Button
    Friend WithEvents btnChiudi As Button
    Friend WithEvents lblMsg As Label
    Friend WithEvents dgvRighe As DataGridView
    Friend WithEvents colSel As DataGridViewCheckBoxColumn
    Friend WithEvents colDDT As DataGridViewTextBoxColumn
    Friend WithEvents colData As DataGridViewTextBoxColumn
    Friend WithEvents colFornitore As DataGridViewTextBoxColumn
    Friend WithEvents colOrdine As DataGridViewTextBoxColumn
    Friend WithEvents colCodice As DataGridViewTextBoxColumn
    Friend WithEvents colDescrizione As DataGridViewTextBoxColumn
    Friend WithEvents colUM As DataGridViewTextBoxColumn
    Friend WithEvents colQuantita As DataGridViewTextBoxColumn
End Class
