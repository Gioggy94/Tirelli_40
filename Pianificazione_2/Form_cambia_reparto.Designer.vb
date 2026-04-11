<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Cambia_Reparto
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
        Me.Btn_Cancella = New System.Windows.Forms.Button()
        Me.Cmd_Seleziona = New System.Windows.Forms.Button()
        Me.Grp_Reparti = New System.Windows.Forms.GroupBox()
        Me.Combo_Reparti = New System.Windows.Forms.ComboBox()
        Me.Grp_Reparti.SuspendLayout()
        Me.SuspendLayout()
        '
        'Btn_Cancella
        '
        Me.Btn_Cancella.Location = New System.Drawing.Point(451, 33)
        Me.Btn_Cancella.Name = "Btn_Cancella"
        Me.Btn_Cancella.Size = New System.Drawing.Size(75, 23)
        Me.Btn_Cancella.TabIndex = 5
        Me.Btn_Cancella.Text = "&Cancella"
        Me.Btn_Cancella.UseVisualStyleBackColor = True
        '
        'Cmd_Seleziona
        '
        Me.Cmd_Seleziona.Location = New System.Drawing.Point(451, 4)
        Me.Cmd_Seleziona.Name = "Cmd_Seleziona"
        Me.Cmd_Seleziona.Size = New System.Drawing.Size(75, 23)
        Me.Cmd_Seleziona.TabIndex = 4
        Me.Cmd_Seleziona.Text = "&Seleziona"
        Me.Cmd_Seleziona.UseVisualStyleBackColor = True
        '
        'Grp_Reparti
        '
        Me.Grp_Reparti.Controls.Add(Me.Combo_Reparti)
        Me.Grp_Reparti.Location = New System.Drawing.Point(8, 6)
        Me.Grp_Reparti.Name = "Grp_Reparti"
        Me.Grp_Reparti.Size = New System.Drawing.Size(424, 51)
        Me.Grp_Reparti.TabIndex = 3
        Me.Grp_Reparti.TabStop = False
        Me.Grp_Reparti.Text = "Seleziona Reparto"
        '
        'Combo_Reparti
        '
        Me.Combo_Reparti.FormattingEnabled = True
        Me.Combo_Reparti.Location = New System.Drawing.Point(7, 20)
        Me.Combo_Reparti.Name = "Combo_Reparti"
        Me.Combo_Reparti.Size = New System.Drawing.Size(411, 21)
        Me.Combo_Reparti.TabIndex = 0
        '
        'Form_Cambia_Reparto
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(535, 62)
        Me.ControlBox = False
        Me.Controls.Add(Me.Btn_Cancella)
        Me.Controls.Add(Me.Cmd_Seleziona)
        Me.Controls.Add(Me.Grp_Reparti)
        Me.Name = "Form_Cambia_Reparto"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form_cambia_reparto"
        Me.TopMost = True
        Me.Grp_Reparti.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Btn_Cancella As Button
    Friend WithEvents Cmd_Seleziona As Button
    Friend WithEvents Grp_Reparti As GroupBox
    Friend WithEvents Combo_Reparti As ComboBox
End Class
