<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ConfigPassword
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ConfigPassword))
        Me.TXTPAssword = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BTN_OK = New System.Windows.Forms.Button()
        Me.BTNCancel = New System.Windows.Forms.Button()
        Me.LBLError = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'TXTPAssword
        '
        Me.TXTPAssword.Location = New System.Drawing.Point(8, 31)
        Me.TXTPAssword.Margin = New System.Windows.Forms.Padding(2)
        Me.TXTPAssword.Name = "TXTPAssword"
        Me.TXTPAssword.Size = New System.Drawing.Size(297, 20)
        Me.TXTPAssword.TabIndex = 0
        Me.TXTPAssword.Text = "T1r3l11@4zero!?"
        Me.TXTPAssword.UseSystemPasswordChar = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 6)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(284, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Inserisci la password per aprire le configurazioni applicative"
        '
        'BTN_OK
        '
        Me.BTN_OK.Location = New System.Drawing.Point(71, 93)
        Me.BTN_OK.Margin = New System.Windows.Forms.Padding(2)
        Me.BTN_OK.Name = "BTN_OK"
        Me.BTN_OK.Size = New System.Drawing.Size(55, 25)
        Me.BTN_OK.TabIndex = 2
        Me.BTN_OK.Text = "OK"
        Me.BTN_OK.UseVisualStyleBackColor = True
        '
        'BTNCancel
        '
        Me.BTNCancel.Location = New System.Drawing.Point(169, 93)
        Me.BTNCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.BTNCancel.Name = "BTNCancel"
        Me.BTNCancel.Size = New System.Drawing.Size(55, 25)
        Me.BTNCancel.TabIndex = 3
        Me.BTNCancel.Text = "Annulla"
        Me.BTNCancel.UseVisualStyleBackColor = True
        '
        'LBLError
        '
        Me.LBLError.AutoSize = True
        Me.LBLError.ForeColor = System.Drawing.Color.Red
        Me.LBLError.Location = New System.Drawing.Point(8, 66)
        Me.LBLError.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.LBLError.Name = "LBLError"
        Me.LBLError.Size = New System.Drawing.Size(0, 13)
        Me.LBLError.TabIndex = 4
        '
        'ConfigPassword
        '
        Me.AcceptButton = Me.BTN_OK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(312, 125)
        Me.Controls.Add(Me.LBLError)
        Me.Controls.Add(Me.BTNCancel)
        Me.Controls.Add(Me.BTN_OK)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TXTPAssword)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ConfigPassword"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Password"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TXTPAssword As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents BTN_OK As Button
    Friend WithEvents BTNCancel As Button
    Friend WithEvents LBLError As Label
End Class
