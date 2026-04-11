<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Seleziona_BP_Trattamento
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
        Me.DataGrid_BP = New System.Windows.Forms.DataGridView()
        Me.Nome_BP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Cmd_Seleziona = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Cmd_Exit = New System.Windows.Forms.Button()
        CType(Me.DataGrid_BP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGrid_BP
        '
        Me.DataGrid_BP.AllowUserToAddRows = False
        Me.DataGrid_BP.AllowUserToDeleteRows = False
        Me.DataGrid_BP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGrid_BP.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Nome_BP, Me.Cmd_Seleziona})
        Me.DataGrid_BP.Location = New System.Drawing.Point(12, 77)
        Me.DataGrid_BP.Name = "DataGrid_BP"
        Me.DataGrid_BP.ReadOnly = True
        Me.DataGrid_BP.RowHeadersVisible = False
        Me.DataGrid_BP.Size = New System.Drawing.Size(776, 361)
        Me.DataGrid_BP.TabIndex = 0
        '
        'Nome_BP
        '
        Me.Nome_BP.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Nome_BP.FillWeight = 80.0!
        Me.Nome_BP.HeaderText = "Nome"
        Me.Nome_BP.Name = "Nome_BP"
        Me.Nome_BP.ReadOnly = True
        '
        'Cmd_Seleziona
        '
        Me.Cmd_Seleziona.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.Cmd_Seleziona.FillWeight = 20.0!
        Me.Cmd_Seleziona.HeaderText = "Seleziona"
        Me.Cmd_Seleziona.Name = "Cmd_Seleziona"
        Me.Cmd_Seleziona.ReadOnly = True
        '
        'Cmd_Exit
        '
        Me.Cmd_Exit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Cmd_Exit.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cmd_Exit.Location = New System.Drawing.Point(661, 11)
        Me.Cmd_Exit.Margin = New System.Windows.Forms.Padding(2)
        Me.Cmd_Exit.Name = "Cmd_Exit"
        Me.Cmd_Exit.Size = New System.Drawing.Size(128, 61)
        Me.Cmd_Exit.TabIndex = 191
        Me.Cmd_Exit.Text = "X"
        Me.Cmd_Exit.UseVisualStyleBackColor = True
        '
        'Form_Seleziona_BP_Trattamento
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.ControlBox = False
        Me.Controls.Add(Me.Cmd_Exit)
        Me.Controls.Add(Me.DataGrid_BP)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "Form_Seleziona_BP_Trattamento"
        Me.Text = "Form_Seleziona_BP_Trattamento"
        CType(Me.DataGrid_BP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGrid_BP As DataGridView
    Friend WithEvents Nome_BP As DataGridViewTextBoxColumn
    Friend WithEvents Cmd_Seleziona As DataGridViewButtonColumn
    Friend WithEvents Cmd_Exit As Button
End Class
