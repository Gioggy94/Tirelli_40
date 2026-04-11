<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Visualizza_Articolo
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
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Grp_Articolo = New System.Windows.Forms.GroupBox()
        Me.Lbl_Descrizione = New System.Windows.Forms.Label()
        Me.Lbl_Codice = New System.Windows.Forms.Label()
        Me.Cmd_Annulla = New System.Windows.Forms.Button()
        Me.DataGridView_magazzino = New System.Windows.Forms.DataGridView()
        Me.Magazzino = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.A_MAGA = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CONF_ = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ORD_ = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Grp_Articolo.SuspendLayout()
        CType(Me.DataGridView_magazzino, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Grp_Articolo
        '
        Me.Grp_Articolo.Controls.Add(Me.Lbl_Descrizione)
        Me.Grp_Articolo.Controls.Add(Me.Lbl_Codice)
        Me.Grp_Articolo.Location = New System.Drawing.Point(12, 12)
        Me.Grp_Articolo.Name = "Grp_Articolo"
        Me.Grp_Articolo.Size = New System.Drawing.Size(776, 48)
        Me.Grp_Articolo.TabIndex = 189
        Me.Grp_Articolo.TabStop = False
        Me.Grp_Articolo.Text = "Articolo"
        '
        'Lbl_Descrizione
        '
        Me.Lbl_Descrizione.AutoSize = True
        Me.Lbl_Descrizione.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_Descrizione.Location = New System.Drawing.Point(120, 16)
        Me.Lbl_Descrizione.Name = "Lbl_Descrizione"
        Me.Lbl_Descrizione.Size = New System.Drawing.Size(70, 24)
        Me.Lbl_Descrizione.TabIndex = 1
        Me.Lbl_Descrizione.Text = "Codice"
        '
        'Lbl_Codice
        '
        Me.Lbl_Codice.AutoSize = True
        Me.Lbl_Codice.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_Codice.Location = New System.Drawing.Point(6, 16)
        Me.Lbl_Codice.Name = "Lbl_Codice"
        Me.Lbl_Codice.Size = New System.Drawing.Size(70, 24)
        Me.Lbl_Codice.TabIndex = 0
        Me.Lbl_Codice.Text = "Codice"
        '
        'Cmd_Annulla
        '
        Me.Cmd_Annulla.Location = New System.Drawing.Point(697, 401)
        Me.Cmd_Annulla.Name = "Cmd_Annulla"
        Me.Cmd_Annulla.Size = New System.Drawing.Size(91, 37)
        Me.Cmd_Annulla.TabIndex = 188
        Me.Cmd_Annulla.Text = "Annulla"
        Me.Cmd_Annulla.UseVisualStyleBackColor = True
        '
        'DataGridView_magazzino
        '
        Me.DataGridView_magazzino.AllowUserToAddRows = False
        Me.DataGridView_magazzino.AllowUserToDeleteRows = False
        Me.DataGridView_magazzino.AllowUserToResizeColumns = False
        Me.DataGridView_magazzino.AllowUserToResizeRows = False
        Me.DataGridView_magazzino.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView_magazzino.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView_magazzino.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView_magazzino.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Magazzino, Me.A_MAGA, Me.CONF_, Me.ORD_})
        Me.DataGridView_magazzino.Enabled = False
        Me.DataGridView_magazzino.Location = New System.Drawing.Point(12, 66)
        Me.DataGridView_magazzino.Name = "DataGridView_magazzino"
        Me.DataGridView_magazzino.ReadOnly = True
        Me.DataGridView_magazzino.RowHeadersVisible = False
        Me.DataGridView_magazzino.RowHeadersWidth = 123
        Me.DataGridView_magazzino.Size = New System.Drawing.Size(776, 329)
        Me.DataGridView_magazzino.TabIndex = 187
        '
        'Magazzino
        '
        Me.Magazzino.HeaderText = "Magazzino"
        Me.Magazzino.MinimumWidth = 15
        Me.Magazzino.Name = "Magazzino"
        Me.Magazzino.ReadOnly = True
        '
        'A_MAGA
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.Format = "N2"
        DataGridViewCellStyle4.NullValue = Nothing
        Me.A_MAGA.DefaultCellStyle = DataGridViewCellStyle4
        Me.A_MAGA.HeaderText = "A Magazzino"
        Me.A_MAGA.MinimumWidth = 15
        Me.A_MAGA.Name = "A_MAGA"
        Me.A_MAGA.ReadOnly = True
        '
        'CONF_
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.Format = "N2"
        DataGridViewCellStyle5.NullValue = Nothing
        Me.CONF_.DefaultCellStyle = DataGridViewCellStyle5
        Me.CONF_.HeaderText = "Confermato"
        Me.CONF_.MinimumWidth = 15
        Me.CONF_.Name = "CONF_"
        Me.CONF_.ReadOnly = True
        '
        'ORD_
        '
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle6.Format = "N2"
        Me.ORD_.DefaultCellStyle = DataGridViewCellStyle6
        Me.ORD_.HeaderText = "Ordinato"
        Me.ORD_.MinimumWidth = 15
        Me.ORD_.Name = "ORD_"
        Me.ORD_.ReadOnly = True
        '
        'Form_Visualizza_Articolo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Grp_Articolo)
        Me.Controls.Add(Me.Cmd_Annulla)
        Me.Controls.Add(Me.DataGridView_magazzino)
        Me.Name = "Form_Visualizza_Articolo"
        Me.Text = "Form_Visualizza_Articolo"
        Me.Grp_Articolo.ResumeLayout(False)
        Me.Grp_Articolo.PerformLayout()
        CType(Me.DataGridView_magazzino, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Grp_Articolo As GroupBox
    Friend WithEvents Lbl_Descrizione As Label
    Friend WithEvents Lbl_Codice As Label
    Friend WithEvents Cmd_Annulla As Button
    Friend WithEvents DataGridView_magazzino As DataGridView
    Friend WithEvents Magazzino As DataGridViewTextBoxColumn
    Friend WithEvents A_MAGA As DataGridViewTextBoxColumn
    Friend WithEvents CONF_ As DataGridViewTextBoxColumn
    Friend WithEvents ORD_ As DataGridViewTextBoxColumn
End Class
