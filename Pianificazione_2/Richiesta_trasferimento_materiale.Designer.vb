<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Richiesta_trasferimento_materiale
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
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.DataGridView = New System.Windows.Forms.DataGridView()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Linenum = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Visorder = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Flag = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Codice = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Descrizione = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Disegno = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Q = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Tras = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Da_tras = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Q_richiesta = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Da_magazzino = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.A_Magazzino = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Giacenza = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Q_RT = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Docentry_odp_ = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Docentry_oc_ = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Docnum = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Commessa = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.Panel4.SuspendLayout()
        CType(Me.DataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(53, 766)
        Me.Panel1.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel2.Controls.Add(Me.TableLayoutPanel1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(53, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1013, 99)
        Me.Panel2.TabIndex = 1
        '
        'Button4
        '
        Me.Button4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(911, 2)
        Me.Button4.Margin = New System.Windows.Forms.Padding(2)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(100, 95)
        Me.Button4.TabIndex = 180
        Me.Button4.Text = "X"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel3.Controls.Add(Me.Button1)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(53, 699)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1013, 67)
        Me.Panel3.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Gold
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Location = New System.Drawing.Point(908, 22)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(86, 24)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Aggiungere"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.DataGridView)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel4.Location = New System.Drawing.Point(53, 99)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(1013, 600)
        Me.Panel4.TabIndex = 3
        '
        'DataGridView
        '
        Me.DataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Linenum, Me.Visorder, Me.Flag, Me.Codice, Me.Descrizione, Me.Disegno, Me.Q, Me.Tras, Me.Da_tras, Me.Q_richiesta, Me.Da_magazzino, Me.A_Magazzino, Me.Giacenza, Me.Q_RT, Me.Docentry_odp_, Me.Docentry_oc_, Me.Docnum, Me.Commessa})
        Me.DataGridView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView.Name = "DataGridView"
        Me.DataGridView.RowHeadersVisible = False
        Me.DataGridView.RowHeadersWidth = 10
        Me.DataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView.Size = New System.Drawing.Size(1013, 600)
        Me.DataGridView.TabIndex = 174
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 10
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.Button4, 9, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1013, 99)
        Me.TableLayoutPanel1.TabIndex = 181
        '
        'Linenum
        '
        Me.Linenum.FillWeight = 60.0!
        Me.Linenum.HeaderText = "Linenum"
        Me.Linenum.MinimumWidth = 15
        Me.Linenum.Name = "Linenum"
        Me.Linenum.Visible = False
        '
        'Visorder
        '
        Me.Visorder.FillWeight = 80.0!
        Me.Visorder.HeaderText = "Visorder"
        Me.Visorder.MinimumWidth = 15
        Me.Visorder.Name = "Visorder"
        Me.Visorder.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Visorder.Visible = False
        '
        'Flag
        '
        Me.Flag.FillWeight = 50.0!
        Me.Flag.HeaderText = "Selezionato"
        Me.Flag.Name = "Flag"
        '
        'Codice
        '
        Me.Codice.HeaderText = "Codice"
        Me.Codice.MinimumWidth = 6
        Me.Codice.Name = "Codice"
        '
        'Descrizione
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Descrizione.DefaultCellStyle = DataGridViewCellStyle1
        Me.Descrizione.FillWeight = 200.0!
        Me.Descrizione.HeaderText = "Descrizione"
        Me.Descrizione.MinimumWidth = 15
        Me.Descrizione.Name = "Descrizione"
        '
        'Disegno
        '
        Me.Disegno.HeaderText = "Disegno"
        Me.Disegno.Name = "Disegno"
        '
        'Q
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.Format = "N2"
        Me.Q.DefaultCellStyle = DataGridViewCellStyle2
        Me.Q.FillWeight = 80.0!
        Me.Q.HeaderText = "Q"
        Me.Q.MinimumWidth = 15
        Me.Q.Name = "Q"
        '
        'Tras
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.Format = "N2"
        Me.Tras.DefaultCellStyle = DataGridViewCellStyle3
        Me.Tras.FillWeight = 80.0!
        Me.Tras.HeaderText = "Tras"
        Me.Tras.MinimumWidth = 15
        Me.Tras.Name = "Tras"
        '
        'Da_tras
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.Format = "N2"
        Me.Da_tras.DefaultCellStyle = DataGridViewCellStyle4
        Me.Da_tras.FillWeight = 80.0!
        Me.Da_tras.HeaderText = "Da Tras"
        Me.Da_tras.MinimumWidth = 15
        Me.Da_tras.Name = "Da_tras"
        '
        'Q_richiesta
        '
        DataGridViewCellStyle5.Format = "N2"
        DataGridViewCellStyle5.NullValue = Nothing
        Me.Q_richiesta.DefaultCellStyle = DataGridViewCellStyle5
        Me.Q_richiesta.HeaderText = "Q richiesta"
        Me.Q_richiesta.Name = "Q_richiesta"
        '
        'Da_magazzino
        '
        Me.Da_magazzino.HeaderText = "Da magazzino"
        Me.Da_magazzino.Name = "Da_magazzino"
        '
        'A_Magazzino
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.A_Magazzino.DefaultCellStyle = DataGridViewCellStyle6
        Me.A_Magazzino.FillWeight = 80.0!
        Me.A_Magazzino.HeaderText = "A magazzino"
        Me.A_Magazzino.MinimumWidth = 15
        Me.A_Magazzino.Name = "A_Magazzino"
        '
        'Giacenza
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Giacenza.DefaultCellStyle = DataGridViewCellStyle7
        Me.Giacenza.FillWeight = 80.0!
        Me.Giacenza.HeaderText = "Giacenza"
        Me.Giacenza.MinimumWidth = 15
        Me.Giacenza.Name = "Giacenza"
        '
        'Q_RT
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Q_RT.DefaultCellStyle = DataGridViewCellStyle8
        Me.Q_RT.HeaderText = "Q RT"
        Me.Q_RT.Name = "Q_RT"
        '
        'Docentry_odp_
        '
        Me.Docentry_odp_.HeaderText = "Docentry_odp"
        Me.Docentry_odp_.Name = "Docentry_odp_"
        Me.Docentry_odp_.Visible = False
        '
        'Docentry_oc_
        '
        Me.Docentry_oc_.HeaderText = "Docentry_oc"
        Me.Docentry_oc_.Name = "Docentry_oc_"
        Me.Docentry_oc_.Visible = False
        '
        'Docnum
        '
        Me.Docnum.HeaderText = "Docnum"
        Me.Docnum.Name = "Docnum"
        '
        'Commessa
        '
        Me.Commessa.HeaderText = "Commessa"
        Me.Commessa.Name = "Commessa"
        '
        'Richiesta_trasferimento_materiale
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1066, 766)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "Richiesta_trasferimento_materiale"
        Me.Text = "Richiesta_trasferimento_materiale"
        Me.Panel2.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        CType(Me.DataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Panel4 As Panel
    Friend WithEvents DataGridView As DataGridView
    Friend WithEvents Button4 As Button
    Friend WithEvents Button1 As Button
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents Linenum As DataGridViewTextBoxColumn
    Friend WithEvents Visorder As DataGridViewTextBoxColumn
    Friend WithEvents Flag As DataGridViewCheckBoxColumn
    Friend WithEvents Codice As DataGridViewTextBoxColumn
    Friend WithEvents Descrizione As DataGridViewTextBoxColumn
    Friend WithEvents Disegno As DataGridViewTextBoxColumn
    Friend WithEvents Q As DataGridViewTextBoxColumn
    Friend WithEvents Tras As DataGridViewTextBoxColumn
    Friend WithEvents Da_tras As DataGridViewTextBoxColumn
    Friend WithEvents Q_richiesta As DataGridViewTextBoxColumn
    Friend WithEvents Da_magazzino As DataGridViewTextBoxColumn
    Friend WithEvents A_Magazzino As DataGridViewTextBoxColumn
    Friend WithEvents Giacenza As DataGridViewTextBoxColumn
    Friend WithEvents Q_RT As DataGridViewTextBoxColumn
    Friend WithEvents Docentry_odp_ As DataGridViewTextBoxColumn
    Friend WithEvents Docentry_oc_ As DataGridViewTextBoxColumn
    Friend WithEvents Docnum As DataGridViewTextBoxColumn
    Friend WithEvents Commessa As DataGridViewTextBoxColumn
End Class
