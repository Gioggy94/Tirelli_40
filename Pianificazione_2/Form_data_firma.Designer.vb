<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Data_Firma
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
        Me.Cmd_Annulla = New System.Windows.Forms.Button()
        Me.Cmd_Conferma = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Combo_Dipendenti = New System.Windows.Forms.ComboBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.MonthCalendar1 = New System.Windows.Forms.MonthCalendar()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Cmd_Annulla
        '
        Me.Cmd_Annulla.Location = New System.Drawing.Point(162, 272)
        Me.Cmd_Annulla.Name = "Cmd_Annulla"
        Me.Cmd_Annulla.Size = New System.Drawing.Size(105, 38)
        Me.Cmd_Annulla.TabIndex = 19
        Me.Cmd_Annulla.Text = "Annulla"
        Me.Cmd_Annulla.UseVisualStyleBackColor = True
        '
        'Cmd_Conferma
        '
        Me.Cmd_Conferma.Location = New System.Drawing.Point(12, 272)
        Me.Cmd_Conferma.Name = "Cmd_Conferma"
        Me.Cmd_Conferma.Size = New System.Drawing.Size(105, 38)
        Me.Cmd_Conferma.TabIndex = 18
        Me.Cmd_Conferma.Text = "Conferma"
        Me.Cmd_Conferma.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Combo_Dipendenti)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 209)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(255, 56)
        Me.GroupBox2.TabIndex = 17
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Firma"
        '
        'Combo_Dipendenti
        '
        Me.Combo_Dipendenti.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Dipendenti.FormattingEnabled = True
        Me.Combo_Dipendenti.Location = New System.Drawing.Point(12, 19)
        Me.Combo_Dipendenti.Name = "Combo_Dipendenti"
        Me.Combo_Dipendenti.Size = New System.Drawing.Size(227, 21)
        Me.Combo_Dipendenti.TabIndex = 10
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.MonthCalendar1)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(255, 190)
        Me.GroupBox1.TabIndex = 16
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Data"
        '
        'MonthCalendar1
        '
        Me.MonthCalendar1.Location = New System.Drawing.Point(12, 16)
        Me.MonthCalendar1.Name = "MonthCalendar1"
        Me.MonthCalendar1.TabIndex = 0
        '
        'Form_Data_Firma
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(276, 321)
        Me.Controls.Add(Me.Cmd_Annulla)
        Me.Controls.Add(Me.Cmd_Conferma)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form_Data_Firma"
        Me.Text = "Form_Data_Firma"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Cmd_Annulla As Button
    Friend WithEvents Cmd_Conferma As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Combo_Dipendenti As ComboBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents MonthCalendar1 As MonthCalendar
End Class
