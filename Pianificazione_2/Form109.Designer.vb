<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form109
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
        Me.TextBox_Commessa = New System.Windows.Forms.TextBox()
        Me.TextBox_descrizione_commessa = New System.Windows.Forms.TextBox()
        Me.TextBox_OC = New System.Windows.Forms.TextBox()
        Me.TextBox_consegna = New System.Windows.Forms.TextBox()
        Me.Label_commessa = New System.Windows.Forms.Label()
        Me.Label_Descrizione_commessa = New System.Windows.Forms.Label()
        Me.Label_OC = New System.Windows.Forms.Label()
        Me.Label_conesgna = New System.Windows.Forms.Label()
        Me.Button_inserisci = New System.Windows.Forms.Button()
        Me.Label_cliente = New System.Windows.Forms.Label()
        Me.TextBox_cliente = New System.Windows.Forms.TextBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox_Commessa
        '
        Me.TextBox_Commessa.Location = New System.Drawing.Point(32, 39)
        Me.TextBox_Commessa.Name = "TextBox_Commessa"
        Me.TextBox_Commessa.Size = New System.Drawing.Size(101, 20)
        Me.TextBox_Commessa.TabIndex = 0
        '
        'TextBox_descrizione_commessa
        '
        Me.TextBox_descrizione_commessa.Location = New System.Drawing.Point(161, 39)
        Me.TextBox_descrizione_commessa.Name = "TextBox_descrizione_commessa"
        Me.TextBox_descrizione_commessa.Size = New System.Drawing.Size(337, 20)
        Me.TextBox_descrizione_commessa.TabIndex = 1
        '
        'TextBox_OC
        '
        Me.TextBox_OC.Location = New System.Drawing.Point(504, 39)
        Me.TextBox_OC.Name = "TextBox_OC"
        Me.TextBox_OC.Size = New System.Drawing.Size(101, 20)
        Me.TextBox_OC.TabIndex = 3
        '
        'TextBox_consegna
        '
        Me.TextBox_consegna.Location = New System.Drawing.Point(750, 39)
        Me.TextBox_consegna.Name = "TextBox_consegna"
        Me.TextBox_consegna.Size = New System.Drawing.Size(101, 20)
        Me.TextBox_consegna.TabIndex = 4
        '
        'Label_commessa
        '
        Me.Label_commessa.AutoSize = True
        Me.Label_commessa.Location = New System.Drawing.Point(52, 23)
        Me.Label_commessa.Name = "Label_commessa"
        Me.Label_commessa.Size = New System.Drawing.Size(58, 13)
        Me.Label_commessa.TabIndex = 5
        Me.Label_commessa.Text = "Commessa"
        '
        'Label_Descrizione_commessa
        '
        Me.Label_Descrizione_commessa.AutoSize = True
        Me.Label_Descrizione_commessa.Location = New System.Drawing.Point(158, 23)
        Me.Label_Descrizione_commessa.Name = "Label_Descrizione_commessa"
        Me.Label_Descrizione_commessa.Size = New System.Drawing.Size(115, 13)
        Me.Label_Descrizione_commessa.TabIndex = 6
        Me.Label_Descrizione_commessa.Text = "Descrizione commessa"
        '
        'Label_OC
        '
        Me.Label_OC.AutoSize = True
        Me.Label_OC.Location = New System.Drawing.Point(539, 20)
        Me.Label_OC.Name = "Label_OC"
        Me.Label_OC.Size = New System.Drawing.Size(22, 13)
        Me.Label_OC.TabIndex = 8
        Me.Label_OC.Text = "OC"
        '
        'Label_conesgna
        '
        Me.Label_conesgna.AutoSize = True
        Me.Label_conesgna.Location = New System.Drawing.Point(767, 20)
        Me.Label_conesgna.Name = "Label_conesgna"
        Me.Label_conesgna.Size = New System.Drawing.Size(55, 13)
        Me.Label_conesgna.TabIndex = 9
        Me.Label_conesgna.Text = "Consegna"
        '
        'Button_inserisci
        '
        Me.Button_inserisci.Location = New System.Drawing.Point(750, 100)
        Me.Button_inserisci.Name = "Button_inserisci"
        Me.Button_inserisci.Size = New System.Drawing.Size(156, 29)
        Me.Button_inserisci.TabIndex = 10
        Me.Button_inserisci.Text = "Inserisci"
        Me.Button_inserisci.UseVisualStyleBackColor = True
        '
        'Label_cliente
        '
        Me.Label_cliente.AutoSize = True
        Me.Label_cliente.Location = New System.Drawing.Point(650, 20)
        Me.Label_cliente.Name = "Label_cliente"
        Me.Label_cliente.Size = New System.Drawing.Size(39, 13)
        Me.Label_cliente.TabIndex = 12
        Me.Label_cliente.Text = "Cliente"
        '
        'TextBox_cliente
        '
        Me.TextBox_cliente.Location = New System.Drawing.Point(622, 39)
        Me.TextBox_cliente.Name = "TextBox_cliente"
        Me.TextBox_cliente.Size = New System.Drawing.Size(101, 20)
        Me.TextBox_cliente.TabIndex = 11
        '
        'ComboBox1
        '
        Me.ComboBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"MACCHINE", "CDS", "TUTTO"})
        Me.ComboBox1.Location = New System.Drawing.Point(3, 16)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(149, 21)
        Me.ComboBox1.TabIndex = 13
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ComboBox1)
        Me.GroupBox1.Location = New System.Drawing.Point(857, 20)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(155, 44)
        Me.GroupBox1.TabIndex = 14
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Gestione"
        '
        'Form109
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1012, 141)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Label_cliente)
        Me.Controls.Add(Me.TextBox_cliente)
        Me.Controls.Add(Me.Button_inserisci)
        Me.Controls.Add(Me.Label_conesgna)
        Me.Controls.Add(Me.Label_OC)
        Me.Controls.Add(Me.Label_Descrizione_commessa)
        Me.Controls.Add(Me.Label_commessa)
        Me.Controls.Add(Me.TextBox_consegna)
        Me.Controls.Add(Me.TextBox_OC)
        Me.Controls.Add(Me.TextBox_descrizione_commessa)
        Me.Controls.Add(Me.TextBox_Commessa)
        Me.Name = "Form109"
        Me.Text = "Form3"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TextBox_Commessa As TextBox
    Friend WithEvents TextBox_descrizione_commessa As TextBox
    Friend WithEvents TextBox_OC As TextBox
    Friend WithEvents TextBox_consegna As TextBox
    Friend WithEvents Label_commessa As Label
    Friend WithEvents Label_Descrizione_commessa As Label
    Friend WithEvents Label_OC As Label
    Friend WithEvents Label_conesgna As Label
    Friend WithEvents Button_inserisci As Button
    Friend WithEvents Label_cliente As Label
    Friend WithEvents TextBox_cliente As TextBox
    Friend WithEvents ComboBox1 As ComboBox
    Friend WithEvents GroupBox1 As GroupBox
End Class
