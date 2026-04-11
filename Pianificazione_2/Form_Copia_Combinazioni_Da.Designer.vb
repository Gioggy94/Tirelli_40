<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Copia_Combinazioni_Da
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
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Txt_Destinazione = New System.Windows.Forms.TextBox()
        Me.Cmd_Copia = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Txt_Matricola_Da = New System.Windows.Forms.TextBox()
        Me.Cmd_Esci = New System.Windows.Forms.Button()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Txt_Destinazione)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 69)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(136, 47)
        Me.GroupBox2.TabIndex = 7
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Destinazione"
        '
        'Txt_Destinazione
        '
        Me.Txt_Destinazione.Enabled = False
        Me.Txt_Destinazione.Location = New System.Drawing.Point(7, 20)
        Me.Txt_Destinazione.Name = "Txt_Destinazione"
        Me.Txt_Destinazione.Size = New System.Drawing.Size(123, 20)
        Me.Txt_Destinazione.TabIndex = 0
        '
        'Cmd_Copia
        '
        Me.Cmd_Copia.Location = New System.Drawing.Point(154, 68)
        Me.Cmd_Copia.Name = "Cmd_Copia"
        Me.Cmd_Copia.Size = New System.Drawing.Size(88, 48)
        Me.Cmd_Copia.TabIndex = 6
        Me.Cmd_Copia.Text = "Copia"
        Me.Cmd_Copia.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Txt_Matricola_Da)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(327, 50)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Copia Combinazioni da ..."
        '
        'Txt_Matricola_Da
        '
        Me.Txt_Matricola_Da.Location = New System.Drawing.Point(6, 19)
        Me.Txt_Matricola_Da.Name = "Txt_Matricola_Da"
        Me.Txt_Matricola_Da.Size = New System.Drawing.Size(315, 20)
        Me.Txt_Matricola_Da.TabIndex = 0
        '
        'Cmd_Esci
        '
        Me.Cmd_Esci.Location = New System.Drawing.Point(248, 68)
        Me.Cmd_Esci.Name = "Cmd_Esci"
        Me.Cmd_Esci.Size = New System.Drawing.Size(88, 48)
        Me.Cmd_Esci.TabIndex = 4
        Me.Cmd_Esci.Text = "Esci"
        Me.Cmd_Esci.UseVisualStyleBackColor = True
        '
        'Form_Copia_Combinazioni_Da
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(347, 123)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Cmd_Copia)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Cmd_Esci)
        Me.Name = "Form_Copia_Combinazioni_Da"
        Me.Text = "Form_Copia_Combinazioni_Da"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Txt_Destinazione As TextBox
    Friend WithEvents Cmd_Copia As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Txt_Matricola_Da As TextBox
    Friend WithEvents Cmd_Esci As Button
End Class
