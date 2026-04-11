<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Presence
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
        Me.components = New System.ComponentModel.Container()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Timer_Ferretto = New System.Windows.Forms.Timer(Me.components)
        Me.Timer_revisioni = New System.Windows.Forms.Timer(Me.components)
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Timer_stampante = New System.Windows.Forms.Timer(Me.components)
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Timer_giornaliero = New System.Windows.Forms.Timer(Me.components)
        Me.Timer_notturno = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(84, 142)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(112, 39)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Forza"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Timer_Ferretto
        '
        Me.Timer_Ferretto.Interval = 1000
        '
        'Timer_revisioni
        '
        Me.Timer_revisioni.Interval = 10000
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(230, 147)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(84, 34)
        Me.Button2.TabIndex = 1
        Me.Button2.Text = "Esci dal Job"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Timer_stampante
        '
        Me.Timer_stampante.Interval = 60000
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(150, 74)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(100, 34)
        Me.Button3.TabIndex = 2
        Me.Button3.Text = "ORdini inutili"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Timer_giornaliero
        '
        Me.Timer_giornaliero.Interval = 60000
        '
        'Timer_notturno
        '
        Me.Timer_notturno.Interval = 60000
        '
        'Presence
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(374, 207)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Presence"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Presence"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As Button
    Friend WithEvents Timer_Ferretto As Timer
    Friend WithEvents Timer_revisioni As Timer
    Friend WithEvents Button2 As Button
    Friend WithEvents Timer_stampante As Timer
    Friend WithEvents Button3 As Button
    Friend WithEvents Timer_giornaliero As Timer
    Friend WithEvents Timer_notturno As Timer
End Class
