<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ODP_Tree
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
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nodo1")
        Dim TreeNode2 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nodo0", New System.Windows.Forms.TreeNode() {TreeNode1})
        Dim TreeNode3 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nodo2")
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ODP_Tree))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Lbl_Matricola = New System.Windows.Forms.Label()
        Me.TV_Progetto = New System.Windows.Forms.TreeView()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button11 = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.TXT_ODP = New System.Windows.Forms.Label()
        Me.Cmd_Exit = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.Numero_ODP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Disegno_odp = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Commessa_odp = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Stato_odp = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Panel21 = New System.Windows.Forms.Panel()
        Me.Panel36 = New System.Windows.Forms.Panel()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Panel35 = New System.Windows.Forms.Panel()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel5 = New System.Windows.Forms.TableLayoutPanel()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel6 = New System.Windows.Forms.TableLayoutPanel()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.GroupBox12 = New System.Windows.Forms.GroupBox()
        Me.GroupBox14 = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel9 = New System.Windows.Forms.TableLayoutPanel()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.Cmd_Indietro = New System.Windows.Forms.Button()
        Me.Cmd_Avanti = New System.Windows.Forms.Button()
        Me.Txt_DocNum = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Seleziona = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Livello = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.N_ODP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Commessa = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Codice = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Disegno = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Stato = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Column3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Pos = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Lotto = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Fant = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel21.SuspendLayout()
        Me.Panel36.SuspendLayout()
        Me.Panel35.SuspendLayout()
        Me.TableLayoutPanel5.SuspendLayout()
        Me.TableLayoutPanel6.SuspendLayout()
        Me.GroupBox12.SuspendLayout()
        Me.GroupBox14.SuspendLayout()
        Me.TableLayoutPanel9.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Lbl_Matricola)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Left
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(249, 100)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Matricola"
        '
        'Lbl_Matricola
        '
        Me.Lbl_Matricola.AutoSize = True
        Me.Lbl_Matricola.Font = New System.Drawing.Font("Microsoft Sans Serif", 36.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_Matricola.ForeColor = System.Drawing.Color.White
        Me.Lbl_Matricola.Location = New System.Drawing.Point(22, 23)
        Me.Lbl_Matricola.Name = "Lbl_Matricola"
        Me.Lbl_Matricola.Size = New System.Drawing.Size(199, 55)
        Me.Lbl_Matricola.TabIndex = 0
        Me.Lbl_Matricola.Text = "M01234"
        '
        'TV_Progetto
        '
        Me.TV_Progetto.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TV_Progetto.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TV_Progetto.Location = New System.Drawing.Point(3, 3)
        Me.TV_Progetto.Name = "TV_Progetto"
        TreeNode1.Name = "Nodo1"
        TreeNode1.Text = "Nodo1"
        TreeNode2.Name = "Nodo0"
        TreeNode2.Text = "Nodo0"
        TreeNode3.Name = "Nodo2"
        TreeNode3.Text = "Nodo2"
        Me.TV_Progetto.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode2, TreeNode3})
        Me.TV_Progetto.ShowPlusMinus = False
        Me.TV_Progetto.Size = New System.Drawing.Size(499, 739)
        Me.TV_Progetto.TabIndex = 1
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Gray
        Me.Panel1.Controls.Add(Me.Button11)
        Me.Panel1.Controls.Add(Me.GroupBox2)
        Me.Panel1.Controls.Add(Me.Cmd_Exit)
        Me.Panel1.Controls.Add(Me.GroupBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1264, 100)
        Me.Panel1.TabIndex = 2
        '
        'Button11
        '
        Me.Button11.BackColor = System.Drawing.Color.Lime
        Me.Button11.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button11.Dock = System.Windows.Forms.DockStyle.Left
        Me.Button11.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button11.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button11.ForeColor = System.Drawing.Color.Maroon
        Me.Button11.Image = CType(resources.GetObject("Button11.Image"), System.Drawing.Image)
        Me.Button11.Location = New System.Drawing.Point(498, 0)
        Me.Button11.Name = "Button11"
        Me.Button11.Size = New System.Drawing.Size(284, 100)
        Me.Button11.TabIndex = 183
        Me.Button11.Text = "ORDINE DI PRODUZIONE"
        Me.Button11.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.Button11.UseVisualStyleBackColor = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.CheckBox3)
        Me.GroupBox2.Controls.Add(Me.TXT_ODP)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Left
        Me.GroupBox2.ForeColor = System.Drawing.Color.White
        Me.GroupBox2.Location = New System.Drawing.Point(249, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(249, 100)
        Me.GroupBox2.TabIndex = 182
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Ordine di Produzione"
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Checked = True
        Me.CheckBox3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox3.Location = New System.Drawing.Point(16, 71)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(79, 17)
        Me.CheckBox3.TabIndex = 1
        Me.CheckBox3.Text = "Solo gruppi"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'TXT_ODP
        '
        Me.TXT_ODP.AutoSize = True
        Me.TXT_ODP.Font = New System.Drawing.Font("Microsoft Sans Serif", 36.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TXT_ODP.ForeColor = System.Drawing.Color.White
        Me.TXT_ODP.Location = New System.Drawing.Point(34, 23)
        Me.TXT_ODP.Name = "TXT_ODP"
        Me.TXT_ODP.Size = New System.Drawing.Size(0, 55)
        Me.TXT_ODP.TabIndex = 0
        '
        'Cmd_Exit
        '
        Me.Cmd_Exit.Dock = System.Windows.Forms.DockStyle.Right
        Me.Cmd_Exit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Cmd_Exit.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cmd_Exit.Location = New System.Drawing.Point(1125, 0)
        Me.Cmd_Exit.Margin = New System.Windows.Forms.Padding(2)
        Me.Cmd_Exit.Name = "Cmd_Exit"
        Me.Cmd_Exit.Size = New System.Drawing.Size(139, 100)
        Me.Cmd_Exit.TabIndex = 181
        Me.Cmd_Exit.Text = "X"
        Me.Cmd_Exit.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.TableLayoutPanel1)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(0, 100)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1264, 745)
        Me.Panel2.TabIndex = 3
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.TV_Progetto, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel3, 2, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 1, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1264, 745)
        Me.TableLayoutPanel1.TabIndex = 2
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.TableLayoutPanel3)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(1139, 2)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(2)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(123, 741)
        Me.Panel3.TabIndex = 186
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 1
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.DataGridView2, 0, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.Panel21, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel5, 0, 4)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel6, 0, 2)
        Me.TableLayoutPanel3.Controls.Add(Me.GroupBox12, 0, 3)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel3.Margin = New System.Windows.Forms.Padding(2)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 5
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 19.04762!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 38.09524!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 14.28571!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(123, 741)
        Me.TableLayoutPanel3.TabIndex = 0
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView2.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView2.BorderStyle = System.Windows.Forms.BorderStyle.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView2.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Numero_ODP, Me.Disegno_odp, Me.Commessa_odp, Me.Stato_odp})
        Me.DataGridView2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView2.Location = New System.Drawing.Point(3, 144)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.RowHeadersVisible = False
        Me.DataGridView2.RowHeadersWidth = 123
        Me.DataGridView2.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView2.Size = New System.Drawing.Size(117, 276)
        Me.DataGridView2.TabIndex = 185
        '
        'Numero_ODP
        '
        Me.Numero_ODP.FillWeight = 80.0!
        Me.Numero_ODP.HeaderText = "N°"
        Me.Numero_ODP.MinimumWidth = 15
        Me.Numero_ODP.Name = "Numero_ODP"
        Me.Numero_ODP.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'Disegno_odp
        '
        Me.Disegno_odp.HeaderText = "Disegno"
        Me.Disegno_odp.Name = "Disegno_odp"
        Me.Disegno_odp.Visible = False
        '
        'Commessa_odp
        '
        Me.Commessa_odp.HeaderText = "COMM"
        Me.Commessa_odp.MinimumWidth = 15
        Me.Commessa_odp.Name = "Commessa_odp"
        Me.Commessa_odp.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'Stato_odp
        '
        Me.Stato_odp.FillWeight = 50.0!
        Me.Stato_odp.HeaderText = "Stato"
        Me.Stato_odp.Name = "Stato_odp"
        '
        'Panel21
        '
        Me.Panel21.Controls.Add(Me.Panel36)
        Me.Panel21.Controls.Add(Me.Panel35)
        Me.Panel21.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel21.Location = New System.Drawing.Point(3, 3)
        Me.Panel21.Name = "Panel21"
        Me.Panel21.Size = New System.Drawing.Size(117, 135)
        Me.Panel21.TabIndex = 1
        '
        'Panel36
        '
        Me.Panel36.Controls.Add(Me.CheckBox2)
        Me.Panel36.Controls.Add(Me.CheckBox1)
        Me.Panel36.Location = New System.Drawing.Point(0, 47)
        Me.Panel36.Name = "Panel36"
        Me.Panel36.Size = New System.Drawing.Size(130, 277)
        Me.Panel36.TabIndex = 2
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.Location = New System.Drawing.Point(3, 27)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(111, 17)
        Me.CheckBox2.TabIndex = 1
        Me.CheckBox2.Text = "Etichetta cassetta"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(3, 9)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(87, 17)
        Me.CheckBox1.TabIndex = 0
        Me.CheckBox1.Text = "Distinta ODP"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Panel35
        '
        Me.Panel35.Controls.Add(Me.Button2)
        Me.Panel35.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel35.Location = New System.Drawing.Point(0, 0)
        Me.Panel35.Name = "Panel35"
        Me.Panel35.Size = New System.Drawing.Size(117, 47)
        Me.Panel35.TabIndex = 1
        '
        'Button2
        '
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.Transparent
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Location = New System.Drawing.Point(0, 0)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(117, 47)
        Me.Button2.TabIndex = 55
        Me.Button2.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage
        Me.Button2.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel5
        '
        Me.TableLayoutPanel5.ColumnCount = 2
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel5.Controls.Add(Me.Button7, 1, 0)
        Me.TableLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel5.Location = New System.Drawing.Point(3, 636)
        Me.TableLayoutPanel5.Name = "TableLayoutPanel5"
        Me.TableLayoutPanel5.RowCount = 2
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel5.Size = New System.Drawing.Size(117, 102)
        Me.TableLayoutPanel5.TabIndex = 3
        '
        'Button7
        '
        Me.Button7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button7.Location = New System.Drawing.Point(61, 3)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(53, 45)
        Me.Button7.TabIndex = 1
        Me.Button7.Text = "Crea/aggiorna lotto di prelievo"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel6
        '
        Me.TableLayoutPanel6.ColumnCount = 2
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel6.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel6.Controls.Add(Me.Button5, 1, 0)
        Me.TableLayoutPanel6.Controls.Add(Me.Button3, 0, 0)
        Me.TableLayoutPanel6.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel6.Location = New System.Drawing.Point(3, 426)
        Me.TableLayoutPanel6.Name = "TableLayoutPanel6"
        Me.TableLayoutPanel6.RowCount = 1
        Me.TableLayoutPanel6.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel6.Size = New System.Drawing.Size(117, 99)
        Me.TableLayoutPanel6.TabIndex = 186
        '
        'Button5
        '
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button5.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.Location = New System.Drawing.Point(60, 2)
        Me.Button5.Margin = New System.Windows.Forms.Padding(2)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(55, 95)
        Me.Button5.TabIndex = 155
        Me.Button5.Text = "X"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(3, 3)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(52, 93)
        Me.Button3.TabIndex = 0
        Me.Button3.Text = ">"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'GroupBox12
        '
        Me.GroupBox12.Controls.Add(Me.GroupBox14)
        Me.GroupBox12.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox12.Location = New System.Drawing.Point(3, 531)
        Me.GroupBox12.Name = "GroupBox12"
        Me.GroupBox12.Size = New System.Drawing.Size(117, 99)
        Me.GroupBox12.TabIndex = 187
        Me.GroupBox12.TabStop = False
        Me.GroupBox12.Text = "Lotto di prelievo"
        '
        'GroupBox14
        '
        Me.GroupBox14.Controls.Add(Me.TableLayoutPanel9)
        Me.GroupBox14.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox14.Location = New System.Drawing.Point(3, 16)
        Me.GroupBox14.Name = "GroupBox14"
        Me.GroupBox14.Size = New System.Drawing.Size(111, 80)
        Me.GroupBox14.TabIndex = 5
        Me.GroupBox14.TabStop = False
        Me.GroupBox14.Text = "Numero lotto prelievo"
        '
        'TableLayoutPanel9
        '
        Me.TableLayoutPanel9.ColumnCount = 4
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.0!))
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 40.0!))
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.0!))
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.0!))
        Me.TableLayoutPanel9.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 15.0!))
        Me.TableLayoutPanel9.Controls.Add(Me.Button8, 3, 0)
        Me.TableLayoutPanel9.Controls.Add(Me.Cmd_Indietro, 0, 0)
        Me.TableLayoutPanel9.Controls.Add(Me.Cmd_Avanti, 2, 0)
        Me.TableLayoutPanel9.Controls.Add(Me.Txt_DocNum, 1, 0)
        Me.TableLayoutPanel9.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel9.Location = New System.Drawing.Point(3, 16)
        Me.TableLayoutPanel9.Name = "TableLayoutPanel9"
        Me.TableLayoutPanel9.RowCount = 1
        Me.TableLayoutPanel9.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel9.Size = New System.Drawing.Size(105, 61)
        Me.TableLayoutPanel9.TabIndex = 0
        '
        'Button8
        '
        Me.Button8.Font = New System.Drawing.Font("Webdings", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Button8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Button8.Location = New System.Drawing.Point(88, 3)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(14, 55)
        Me.Button8.TabIndex = 175
        Me.Button8.Text = "L"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'Cmd_Indietro
        '
        Me.Cmd_Indietro.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Cmd_Indietro.Font = New System.Drawing.Font("Wingdings", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Cmd_Indietro.Location = New System.Drawing.Point(3, 3)
        Me.Cmd_Indietro.Name = "Cmd_Indietro"
        Me.Cmd_Indietro.Size = New System.Drawing.Size(12, 55)
        Me.Cmd_Indietro.TabIndex = 1
        Me.Cmd_Indietro.Text = "ç"
        Me.Cmd_Indietro.UseVisualStyleBackColor = True
        '
        'Cmd_Avanti
        '
        Me.Cmd_Avanti.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Cmd_Avanti.Font = New System.Drawing.Font("Wingdings", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Cmd_Avanti.Location = New System.Drawing.Point(70, 3)
        Me.Cmd_Avanti.Name = "Cmd_Avanti"
        Me.Cmd_Avanti.Size = New System.Drawing.Size(12, 55)
        Me.Cmd_Avanti.TabIndex = 2
        Me.Cmd_Avanti.Text = "è"
        Me.Cmd_Avanti.UseVisualStyleBackColor = True
        '
        'Txt_DocNum
        '
        Me.Txt_DocNum.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Txt_DocNum.Font = New System.Drawing.Font("Microsoft Sans Serif", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Txt_DocNum.Location = New System.Drawing.Point(21, 3)
        Me.Txt_DocNum.Name = "Txt_DocNum"
        Me.Txt_DocNum.Size = New System.Drawing.Size(43, 47)
        Me.Txt_DocNum.TabIndex = 0
        Me.Txt_DocNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 1
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.DataGridView1, 0, 1)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(508, 3)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 3
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 60.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 20.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(626, 739)
        Me.TableLayoutPanel2.TabIndex = 2
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Seleziona, Me.Livello, Me.N_ODP, Me.Commessa, Me.Codice, Me.Column2, Me.Disegno, Me.Stato, Me.Column3, Me.Pos, Me.Lotto, Me.Fant})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 150)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidth = 123
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(620, 437)
        Me.DataGridView1.TabIndex = 185
        '
        'Seleziona
        '
        Me.Seleziona.FillWeight = 50.0!
        Me.Seleziona.HeaderText = "+/-"
        Me.Seleziona.Name = "Seleziona"
        Me.Seleziona.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Seleziona.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'Livello
        '
        Me.Livello.HeaderText = "Livello"
        Me.Livello.Name = "Livello"
        '
        'N_ODP
        '
        Me.N_ODP.HeaderText = "N° ODP"
        Me.N_ODP.Name = "N_ODP"
        '
        'Commessa
        '
        Me.Commessa.HeaderText = "Commessa"
        Me.Commessa.Name = "Commessa"
        '
        'Codice
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Codice.DefaultCellStyle = DataGridViewCellStyle3
        Me.Codice.HeaderText = "Codice"
        Me.Codice.Name = "Codice"
        '
        'Column2
        '
        Me.Column2.HeaderText = "Descrizione"
        Me.Column2.Name = "Column2"
        '
        'Disegno
        '
        Me.Disegno.HeaderText = "Disegno"
        Me.Disegno.Name = "Disegno"
        '
        'Stato
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Stato.DefaultCellStyle = DataGridViewCellStyle4
        Me.Stato.FillWeight = 50.0!
        Me.Stato.HeaderText = "Stato"
        Me.Stato.Name = "Stato"
        '
        'Column3
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.Format = "N0"
        DataGridViewCellStyle5.NullValue = Nothing
        Me.Column3.DefaultCellStyle = DataGridViewCellStyle5
        Me.Column3.HeaderText = "Qta"
        Me.Column3.Name = "Column3"
        '
        'Pos
        '
        Me.Pos.HeaderText = "Pos"
        Me.Pos.Name = "Pos"
        '
        'Lotto
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Lotto.DefaultCellStyle = DataGridViewCellStyle6
        Me.Lotto.HeaderText = "Lotto Prel"
        Me.Lotto.Name = "Lotto"
        '
        'Fant
        '
        Me.Fant.HeaderText = "Fant"
        Me.Fant.Name = "Fant"
        '
        'ODP_Tree
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(1264, 845)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "ODP_Tree"
        Me.Text = "Albero Montaggio"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.Panel3.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel21.ResumeLayout(False)
        Me.Panel36.ResumeLayout(False)
        Me.Panel36.PerformLayout()
        Me.Panel35.ResumeLayout(False)
        Me.TableLayoutPanel5.ResumeLayout(False)
        Me.TableLayoutPanel6.ResumeLayout(False)
        Me.GroupBox12.ResumeLayout(False)
        Me.GroupBox14.ResumeLayout(False)
        Me.TableLayoutPanel9.ResumeLayout(False)
        Me.TableLayoutPanel9.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Lbl_Matricola As Label
    Friend WithEvents TV_Progetto As TreeView
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Cmd_Exit As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents TXT_ODP As Label
    Friend WithEvents Button11 As Button
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents TableLayoutPanel3 As TableLayoutPanel
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents Numero_ODP As DataGridViewTextBoxColumn
    Friend WithEvents Disegno_odp As DataGridViewTextBoxColumn
    Friend WithEvents Commessa_odp As DataGridViewTextBoxColumn
    Friend WithEvents Stato_odp As DataGridViewTextBoxColumn
    Friend WithEvents Panel21 As Panel
    Friend WithEvents Panel36 As Panel
    Friend WithEvents CheckBox2 As CheckBox
    Friend WithEvents CheckBox1 As CheckBox
    Friend WithEvents Panel35 As Panel
    Friend WithEvents Button2 As Button
    Friend WithEvents TableLayoutPanel5 As TableLayoutPanel
    Friend WithEvents Button7 As Button
    Friend WithEvents TableLayoutPanel6 As TableLayoutPanel
    Friend WithEvents Button5 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents GroupBox12 As GroupBox
    Friend WithEvents GroupBox14 As GroupBox
    Friend WithEvents TableLayoutPanel9 As TableLayoutPanel
    Friend WithEvents Button8 As Button
    Friend WithEvents Cmd_Indietro As Button
    Friend WithEvents Cmd_Avanti As Button
    Friend WithEvents Txt_DocNum As TextBox
    Friend WithEvents CheckBox3 As CheckBox
    Friend WithEvents Seleziona As DataGridViewCheckBoxColumn
    Friend WithEvents Livello As DataGridViewTextBoxColumn
    Friend WithEvents N_ODP As DataGridViewTextBoxColumn
    Friend WithEvents Commessa As DataGridViewTextBoxColumn
    Friend WithEvents Codice As DataGridViewTextBoxColumn
    Friend WithEvents Column2 As DataGridViewTextBoxColumn
    Friend WithEvents Disegno As DataGridViewTextBoxColumn
    Friend WithEvents Stato As DataGridViewTextBoxColumn
    Friend WithEvents Column3 As DataGridViewTextBoxColumn
    Friend WithEvents Pos As DataGridViewTextBoxColumn
    Friend WithEvents Lotto As DataGridViewTextBoxColumn
    Friend WithEvents Fant As DataGridViewTextBoxColumn
End Class
