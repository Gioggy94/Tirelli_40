<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Commesse_magazzino_ODP
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
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

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla mediante l'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle13 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle14 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle15 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Button_OC = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button_CDS = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.DataGridView_commesse_odp = New System.Windows.Forms.DataGridView()
        Me.Commessa_Tab = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Descrizione = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Cliente_data = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Cliente_F = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Consegna_data = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TOT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OK_TOT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OK_Comple_TOT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TOT_PREM = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OK_TOT_PREM = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OK_COMPLETABILE_TOT_PREM = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Tot_MONT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OK_TOT_MONT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.OK_completabile_MONT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel1.SuspendLayout()
        Me.Panel7.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.DataGridView_commesse_odp, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel1.Controls.Add(Me.Panel7)
        Me.Panel1.Controls.Add(Me.Panel6)
        Me.Panel1.Controls.Add(Me.Panel5)
        Me.Panel1.Controls.Add(Me.Panel4)
        Me.Panel1.Controls.Add(Me.Panel3)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1141, 78)
        Me.Panel1.TabIndex = 163
        '
        'Panel7
        '
        Me.Panel7.Controls.Add(Me.Button1)
        Me.Panel7.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel7.Location = New System.Drawing.Point(497, 0)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(142, 78)
        Me.Panel7.TabIndex = 167
        '
        'Button1
        '
        Me.Button1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button1.Location = New System.Drawing.Point(0, 0)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(142, 78)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Visualizza per materiale"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.Button3)
        Me.Panel6.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel6.Location = New System.Drawing.Point(1059, 0)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(82, 78)
        Me.Panel6.TabIndex = 165
        '
        'Button3
        '
        Me.Button3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(0, 0)
        Me.Button3.Margin = New System.Windows.Forms.Padding(2)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(82, 78)
        Me.Button3.TabIndex = 158
        Me.Button3.Text = "X"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.GroupBox6)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel5.Location = New System.Drawing.Point(290, 0)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(207, 78)
        Me.Panel5.TabIndex = 164
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.TextBox1)
        Me.GroupBox6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox6.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox6.ForeColor = System.Drawing.Color.Black
        Me.GroupBox6.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(207, 78)
        Me.GroupBox6.TabIndex = 161
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "Ricerca per Commessa"
        '
        'TextBox1
        '
        Me.TextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(3, 20)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(201, 47)
        Me.TextBox1.TabIndex = 159
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.Button_OC)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel4.Location = New System.Drawing.Point(145, 0)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(145, 78)
        Me.Panel4.TabIndex = 163
        '
        'Button_OC
        '
        Me.Button_OC.BackColor = System.Drawing.Color.MediumSpringGreen
        Me.Button_OC.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button_OC.Location = New System.Drawing.Point(0, 0)
        Me.Button_OC.Name = "Button_OC"
        Me.Button_OC.Size = New System.Drawing.Size(145, 78)
        Me.Button_OC.TabIndex = 94
        Me.Button_OC.Text = "ORDINI CLIENTE"
        Me.Button_OC.UseVisualStyleBackColor = False
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.Button_CDS)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(145, 78)
        Me.Panel3.TabIndex = 162
        '
        'Button_CDS
        '
        Me.Button_CDS.BackColor = System.Drawing.Color.Plum
        Me.Button_CDS.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button_CDS.Location = New System.Drawing.Point(0, 0)
        Me.Button_CDS.Name = "Button_CDS"
        Me.Button_CDS.Size = New System.Drawing.Size(145, 78)
        Me.Button_CDS.TabIndex = 93
        Me.Button_CDS.Text = "ODP PER CDS"
        Me.Button_CDS.UseVisualStyleBackColor = False
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.DataGridView_commesse_odp)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 78)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1141, 689)
        Me.Panel2.TabIndex = 164
        '
        'DataGridView_commesse_odp
        '
        Me.DataGridView_commesse_odp.AllowDrop = True
        Me.DataGridView_commesse_odp.AllowUserToAddRows = False
        Me.DataGridView_commesse_odp.AllowUserToDeleteRows = False
        Me.DataGridView_commesse_odp.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView_commesse_odp.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        Me.DataGridView_commesse_odp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlDark
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView_commesse_odp.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView_commesse_odp.ColumnHeadersHeight = 50
        Me.DataGridView_commesse_odp.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Commessa_Tab, Me.Descrizione, Me.Cliente_data, Me.Cliente_F, Me.Consegna_data, Me.TOT, Me.OK_TOT, Me.OK_Comple_TOT, Me.TOT_PREM, Me.OK_TOT_PREM, Me.OK_COMPLETABILE_TOT_PREM, Me.Tot_MONT, Me.OK_TOT_MONT, Me.OK_completabile_MONT})
        Me.DataGridView_commesse_odp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView_commesse_odp.GridColor = System.Drawing.SystemColors.ActiveBorder
        Me.DataGridView_commesse_odp.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.DataGridView_commesse_odp.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView_commesse_odp.Name = "DataGridView_commesse_odp"
        Me.DataGridView_commesse_odp.RowHeadersVisible = False
        Me.DataGridView_commesse_odp.RowHeadersWidth = 51
        Me.DataGridView_commesse_odp.RowTemplate.DefaultCellStyle.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DataGridView_commesse_odp.RowTemplate.Height = 40
        Me.DataGridView_commesse_odp.Size = New System.Drawing.Size(1141, 689)
        Me.DataGridView_commesse_odp.TabIndex = 91
        '
        'Commessa_Tab
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Bold)
        Me.Commessa_Tab.DefaultCellStyle = DataGridViewCellStyle2
        Me.Commessa_Tab.FillWeight = 120.0!
        Me.Commessa_Tab.HeaderText = "Commessa"
        Me.Commessa_Tab.MinimumWidth = 6
        Me.Commessa_Tab.Name = "Commessa_Tab"
        Me.Commessa_Tab.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Commessa_Tab.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'Descrizione
        '
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Descrizione.DefaultCellStyle = DataGridViewCellStyle3
        Me.Descrizione.FillWeight = 200.0!
        Me.Descrizione.HeaderText = "Descrizione"
        Me.Descrizione.MinimumWidth = 20
        Me.Descrizione.Name = "Descrizione"
        '
        'Cliente_data
        '
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cliente_data.DefaultCellStyle = DataGridViewCellStyle4
        Me.Cliente_data.FillWeight = 120.3767!
        Me.Cliente_data.HeaderText = "Cliente"
        Me.Cliente_data.MinimumWidth = 6
        Me.Cliente_data.Name = "Cliente_data"
        '
        'Cliente_F
        '
        DataGridViewCellStyle5.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cliente_F.DefaultCellStyle = DataGridViewCellStyle5
        Me.Cliente_F.FillWeight = 115.1934!
        Me.Cliente_F.HeaderText = "Cliente_F"
        Me.Cliente_F.MinimumWidth = 20
        Me.Cliente_F.Name = "Cliente_F"
        '
        'Consegna_data
        '
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!)
        Me.Consegna_data.DefaultCellStyle = DataGridViewCellStyle6
        Me.Consegna_data.FillWeight = 79.57105!
        Me.Consegna_data.HeaderText = "Consegna"
        Me.Consegna_data.MinimumWidth = 6
        Me.Consegna_data.Name = "Consegna_data"
        '
        'TOT
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle7.BackColor = System.Drawing.Color.Yellow
        Me.TOT.DefaultCellStyle = DataGridViewCellStyle7
        Me.TOT.HeaderText = "TOT"
        Me.TOT.Name = "TOT"
        '
        'OK_TOT
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle8.BackColor = System.Drawing.Color.Yellow
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.Format = "N0"
        DataGridViewCellStyle8.NullValue = Nothing
        Me.OK_TOT.DefaultCellStyle = DataGridViewCellStyle8
        Me.OK_TOT.HeaderText = "OK/TOT"
        Me.OK_TOT.Name = "OK_TOT"
        '
        'OK_Comple_TOT
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle9.BackColor = System.Drawing.Color.Yellow
        DataGridViewCellStyle9.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle9.Format = "N0"
        DataGridViewCellStyle9.NullValue = Nothing
        Me.OK_Comple_TOT.DefaultCellStyle = DataGridViewCellStyle9
        Me.OK_Comple_TOT.HeaderText = "OK+Completabile/TOT"
        Me.OK_Comple_TOT.Name = "OK_Comple_TOT"
        '
        'TOT_PREM
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle10.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        DataGridViewCellStyle10.Format = "N0"
        DataGridViewCellStyle10.NullValue = Nothing
        Me.TOT_PREM.DefaultCellStyle = DataGridViewCellStyle10
        Me.TOT_PREM.HeaderText = "TOT PREM"
        Me.TOT_PREM.Name = "TOT_PREM"
        '
        'OK_TOT_PREM
        '
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle11.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        DataGridViewCellStyle11.Format = "N0"
        DataGridViewCellStyle11.NullValue = Nothing
        Me.OK_TOT_PREM.DefaultCellStyle = DataGridViewCellStyle11
        Me.OK_TOT_PREM.HeaderText = "OK/TOT PREM"
        Me.OK_TOT_PREM.Name = "OK_TOT_PREM"
        '
        'OK_COMPLETABILE_TOT_PREM
        '
        DataGridViewCellStyle12.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle12.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        DataGridViewCellStyle12.Format = "N0"
        DataGridViewCellStyle12.NullValue = Nothing
        Me.OK_COMPLETABILE_TOT_PREM.DefaultCellStyle = DataGridViewCellStyle12
        Me.OK_COMPLETABILE_TOT_PREM.HeaderText = "OK+Completabile/TOT PREM"
        Me.OK_COMPLETABILE_TOT_PREM.Name = "OK_COMPLETABILE_TOT_PREM"
        '
        'Tot_MONT
        '
        DataGridViewCellStyle13.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle13.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle13.Format = "N0"
        Me.Tot_MONT.DefaultCellStyle = DataGridViewCellStyle13
        Me.Tot_MONT.FillWeight = 70.0!
        Me.Tot_MONT.HeaderText = "TOT MONT"
        Me.Tot_MONT.Name = "Tot_MONT"
        '
        'OK_TOT_MONT
        '
        DataGridViewCellStyle14.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle14.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle14.Format = "N0"
        Me.OK_TOT_MONT.DefaultCellStyle = DataGridViewCellStyle14
        Me.OK_TOT_MONT.FillWeight = 70.0!
        Me.OK_TOT_MONT.HeaderText = "OK/TOT MONT"
        Me.OK_TOT_MONT.Name = "OK_TOT_MONT"
        '
        'OK_completabile_MONT
        '
        DataGridViewCellStyle15.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle15.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        DataGridViewCellStyle15.Format = "N0"
        Me.OK_completabile_MONT.DefaultCellStyle = DataGridViewCellStyle15
        Me.OK_completabile_MONT.FillWeight = 70.0!
        Me.OK_completabile_MONT.HeaderText = "OK+Completabile/TOT MONT"
        Me.OK_completabile_MONT.Name = "OK_completabile_MONT"
        '
        'Commesse_magazzino_ODP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1141, 767)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "Commesse_magazzino_ODP"
        Me.Text = "Commesse_magazzino_ODP"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel7.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.DataGridView_commesse_odp, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel6 As Panel
    Friend WithEvents Button3 As Button
    Friend WithEvents Panel5 As Panel
    Friend WithEvents GroupBox6 As GroupBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Button_OC As Button
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Button_CDS As Button
    Friend WithEvents Panel2 As Panel
    Friend WithEvents DataGridView_commesse_odp As DataGridView
    Friend WithEvents Commessa_Tab As DataGridViewButtonColumn
    Friend WithEvents Descrizione As DataGridViewTextBoxColumn
    Friend WithEvents Cliente_data As DataGridViewTextBoxColumn
    Friend WithEvents Cliente_F As DataGridViewTextBoxColumn
    Friend WithEvents Consegna_data As DataGridViewTextBoxColumn
    Friend WithEvents TOT As DataGridViewTextBoxColumn
    Friend WithEvents OK_TOT As DataGridViewTextBoxColumn
    Friend WithEvents OK_Comple_TOT As DataGridViewTextBoxColumn
    Friend WithEvents TOT_PREM As DataGridViewTextBoxColumn
    Friend WithEvents OK_TOT_PREM As DataGridViewTextBoxColumn
    Friend WithEvents OK_COMPLETABILE_TOT_PREM As DataGridViewTextBoxColumn
    Friend WithEvents Tot_MONT As DataGridViewTextBoxColumn
    Friend WithEvents OK_TOT_MONT As DataGridViewTextBoxColumn
    Friend WithEvents OK_completabile_MONT As DataGridViewTextBoxColumn
    Friend WithEvents Panel7 As Panel
    Friend WithEvents Button1 As Button
End Class
