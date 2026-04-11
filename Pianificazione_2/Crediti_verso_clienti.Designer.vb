<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Crediti_verso_clienti
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Seleziona = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Docentry = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Fattura = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Data_fattura = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Data_scadenza = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Overdue = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BP_CODE = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.BP_name = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Finale_bp_code = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Final_BP_Name = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Year = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Department = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Saldo_BP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Salesman = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Country = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Total = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Paid = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Adjustment = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Credit = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Payment_group = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 90.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel3, 0, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(988, 549)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 1
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.DataGridView1, 0, 2)
        Me.TableLayoutPanel2.Controls.Add(Me.Button4, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel4, 0, 1)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(101, 3)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 3
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15.0!))
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 75.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(884, 543)
        Me.TableLayoutPanel2.TabIndex = 0
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Seleziona, Me.Docentry, Me.Fattura, Me.Data_fattura, Me.Data_scadenza, Me.Overdue, Me.BP_CODE, Me.BP_name, Me.Finale_bp_code, Me.Final_BP_Name, Me.Year, Me.Department, Me.Saldo_BP, Me.Salesman, Me.Country, Me.Total, Me.Paid, Me.Adjustment, Me.Credit, Me.Payment_group})
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 138)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.RowHeadersVisible = False
        Me.DataGridView1.RowHeadersWidth = 123
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.DataGridView1.Size = New System.Drawing.Size(878, 402)
        Me.DataGridView1.TabIndex = 186
        '
        'Button4
        '
        Me.Button4.Dock = System.Windows.Forms.DockStyle.Right
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.Location = New System.Drawing.Point(707, 2)
        Me.Button4.Margin = New System.Windows.Forms.Padding(2)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(175, 50)
        Me.Button4.TabIndex = 180
        Me.Button4.Text = "X"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.ColumnCount = 6
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 10.0!))
        Me.TableLayoutPanel4.Controls.Add(Me.GroupBox4, 3, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.GroupBox3, 2, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.GroupBox2, 1, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.GroupBox1, 0, 0)
        Me.TableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(3, 57)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 2
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(878, 75)
        Me.TableLayoutPanel4.TabIndex = 187
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.TextBox4)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox4.Location = New System.Drawing.Point(441, 3)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(140, 31)
        Me.GroupBox4.TabIndex = 3
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Deparment"
        '
        'TextBox4
        '
        Me.TextBox4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBox4.Location = New System.Drawing.Point(3, 16)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(134, 20)
        Me.TextBox4.TabIndex = 1
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.TextBox3)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.Location = New System.Drawing.Point(295, 3)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(140, 31)
        Me.GroupBox3.TabIndex = 2
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Nome cliente"
        '
        'TextBox3
        '
        Me.TextBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBox3.Location = New System.Drawing.Point(3, 16)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(134, 20)
        Me.TextBox3.TabIndex = 1
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.TextBox2)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(149, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(140, 31)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Codice cliente"
        '
        'TextBox2
        '
        Me.TextBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBox2.Location = New System.Drawing.Point(3, 16)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(134, 20)
        Me.TextBox2.TabIndex = 1
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TextBox1)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(140, 31)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "N° Fattura"
        '
        'TextBox1
        '
        Me.TextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBox1.Location = New System.Drawing.Point(3, 16)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(134, 20)
        Me.TextBox1.TabIndex = 0
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 1
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.Button1, 0, 2)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 3)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 4
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.63327!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.63327!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 16.63327!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.1002!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(92, 543)
        Me.TableLayoutPanel3.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button1.Location = New System.Drawing.Point(3, 183)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(86, 84)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Crea PDF"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Seleziona
        '
        Me.Seleziona.FillWeight = 50.0!
        Me.Seleziona.HeaderText = "Seleziona"
        Me.Seleziona.Name = "Seleziona"
        '
        'Docentry
        '
        Me.Docentry.FillWeight = 70.0!
        Me.Docentry.HeaderText = "Docentry"
        Me.Docentry.MinimumWidth = 15
        Me.Docentry.Name = "Docentry"
        Me.Docentry.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Docentry.Visible = False
        '
        'Fattura
        '
        Me.Fattura.FillWeight = 70.0!
        Me.Fattura.HeaderText = "Invoice"
        Me.Fattura.MinimumWidth = 6
        Me.Fattura.Name = "Fattura"
        Me.Fattura.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Fattura.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'Data_fattura
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.Format = "d"
        DataGridViewCellStyle2.NullValue = Nothing
        Me.Data_fattura.DefaultCellStyle = DataGridViewCellStyle2
        Me.Data_fattura.FillWeight = 70.0!
        Me.Data_fattura.HeaderText = "Invoice date"
        Me.Data_fattura.MinimumWidth = 6
        Me.Data_fattura.Name = "Data_fattura"
        Me.Data_fattura.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'Data_scadenza
        '
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle3.Format = "d"
        DataGridViewCellStyle3.NullValue = Nothing
        Me.Data_scadenza.DefaultCellStyle = DataGridViewCellStyle3
        Me.Data_scadenza.FillWeight = 70.0!
        Me.Data_scadenza.HeaderText = "Data scadenza"
        Me.Data_scadenza.MinimumWidth = 6
        Me.Data_scadenza.Name = "Data_scadenza"
        '
        'Overdue
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.Format = "N0"
        DataGridViewCellStyle4.NullValue = Nothing
        Me.Overdue.DefaultCellStyle = DataGridViewCellStyle4
        Me.Overdue.FillWeight = 90.0!
        Me.Overdue.HeaderText = "Overdue"
        Me.Overdue.MinimumWidth = 6
        Me.Overdue.Name = "Overdue"
        '
        'BP_CODE
        '
        Me.BP_CODE.FillWeight = 80.0!
        Me.BP_CODE.HeaderText = "BP Code"
        Me.BP_CODE.MinimumWidth = 6
        Me.BP_CODE.Name = "BP_CODE"
        '
        'BP_name
        '
        Me.BP_name.FillWeight = 200.0!
        Me.BP_name.HeaderText = "BP Name"
        Me.BP_name.MinimumWidth = 6
        Me.BP_name.Name = "BP_name"
        Me.BP_name.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'Finale_bp_code
        '
        Me.Finale_bp_code.FillWeight = 50.0!
        Me.Finale_bp_code.HeaderText = "Final BP Code"
        Me.Finale_bp_code.MinimumWidth = 6
        Me.Finale_bp_code.Name = "Finale_bp_code"
        Me.Finale_bp_code.Visible = False
        '
        'Final_BP_Name
        '
        Me.Final_BP_Name.FillWeight = 200.0!
        Me.Final_BP_Name.HeaderText = "Final_BP_name"
        Me.Final_BP_Name.MinimumWidth = 6
        Me.Final_BP_Name.Name = "Final_BP_Name"
        '
        'Year
        '
        Me.Year.FillWeight = 60.0!
        Me.Year.HeaderText = "Year"
        Me.Year.MinimumWidth = 6
        Me.Year.Name = "Year"
        '
        'Department
        '
        Me.Department.HeaderText = "Department"
        Me.Department.MinimumWidth = 6
        Me.Department.Name = "Department"
        Me.Department.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'Saldo_BP
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle5.Format = "C0"
        DataGridViewCellStyle5.NullValue = Nothing
        Me.Saldo_BP.DefaultCellStyle = DataGridViewCellStyle5
        Me.Saldo_BP.HeaderText = "Saldo BP"
        Me.Saldo_BP.MinimumWidth = 6
        Me.Saldo_BP.Name = "Saldo_BP"
        '
        'Salesman
        '
        Me.Salesman.HeaderText = "Salesman"
        Me.Salesman.MinimumWidth = 6
        Me.Salesman.Name = "Salesman"
        '
        'Country
        '
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Country.DefaultCellStyle = DataGridViewCellStyle6
        Me.Country.HeaderText = "Country"
        Me.Country.MinimumWidth = 6
        Me.Country.Name = "Country"
        '
        'Total
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle7.Format = "C0"
        DataGridViewCellStyle7.NullValue = Nothing
        Me.Total.DefaultCellStyle = DataGridViewCellStyle7
        Me.Total.HeaderText = "Total"
        Me.Total.MinimumWidth = 6
        Me.Total.Name = "Total"
        '
        'Paid
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle8.Format = "C0"
        DataGridViewCellStyle8.NullValue = Nothing
        Me.Paid.DefaultCellStyle = DataGridViewCellStyle8
        Me.Paid.FillWeight = 50.0!
        Me.Paid.HeaderText = "Paid"
        Me.Paid.MinimumWidth = 6
        Me.Paid.Name = "Paid"
        '
        'Adjustment
        '
        DataGridViewCellStyle9.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle9.Format = "C0"
        DataGridViewCellStyle9.NullValue = Nothing
        Me.Adjustment.DefaultCellStyle = DataGridViewCellStyle9
        Me.Adjustment.FillWeight = 50.0!
        Me.Adjustment.HeaderText = "Adjustment"
        Me.Adjustment.MinimumWidth = 6
        Me.Adjustment.Name = "Adjustment"
        '
        'Credit
        '
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight
        DataGridViewCellStyle10.Format = "C0"
        DataGridViewCellStyle10.NullValue = Nothing
        Me.Credit.DefaultCellStyle = DataGridViewCellStyle10
        Me.Credit.HeaderText = "Credit"
        Me.Credit.MinimumWidth = 6
        Me.Credit.Name = "Credit"
        '
        'Payment_group
        '
        Me.Payment_group.FillWeight = 80.0!
        Me.Payment_group.HeaderText = "Payment group"
        Me.Payment_group.MinimumWidth = 6
        Me.Payment_group.Name = "Payment_group"
        Me.Payment_group.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'Crediti_verso_clienti
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(988, 549)
        Me.ControlBox = False
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Name = "Crediti_verso_clienti"
        Me.Text = "Crediti_verso_clienti"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents Button4 As Button
    Friend WithEvents TableLayoutPanel3 As TableLayoutPanel
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents TableLayoutPanel4 As TableLayoutPanel
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Seleziona As DataGridViewCheckBoxColumn
    Friend WithEvents Docentry As DataGridViewTextBoxColumn
    Friend WithEvents Fattura As DataGridViewButtonColumn
    Friend WithEvents Data_fattura As DataGridViewTextBoxColumn
    Friend WithEvents Data_scadenza As DataGridViewTextBoxColumn
    Friend WithEvents Overdue As DataGridViewTextBoxColumn
    Friend WithEvents BP_CODE As DataGridViewTextBoxColumn
    Friend WithEvents BP_name As DataGridViewTextBoxColumn
    Friend WithEvents Finale_bp_code As DataGridViewTextBoxColumn
    Friend WithEvents Final_BP_Name As DataGridViewTextBoxColumn
    Friend WithEvents Year As DataGridViewTextBoxColumn
    Friend WithEvents Department As DataGridViewTextBoxColumn
    Friend WithEvents Saldo_BP As DataGridViewTextBoxColumn
    Friend WithEvents Salesman As DataGridViewTextBoxColumn
    Friend WithEvents Country As DataGridViewTextBoxColumn
    Friend WithEvents Total As DataGridViewTextBoxColumn
    Friend WithEvents Paid As DataGridViewTextBoxColumn
    Friend WithEvents Adjustment As DataGridViewTextBoxColumn
    Friend WithEvents Credit As DataGridViewTextBoxColumn
    Friend WithEvents Payment_group As DataGridViewTextBoxColumn
End Class
