<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form107
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
        Me.LabelDisegnoSAP = New System.Windows.Forms.Label()
        Me.LabelDis = New System.Windows.Forms.Label()
        Me.LabelDescrizioneSAP = New System.Windows.Forms.Label()
        Me.LabelCodiceSAP = New System.Windows.Forms.Label()
        Me.LabelDesc = New System.Windows.Forms.Label()
        Me.LabelCodice = New System.Windows.Forms.Label()
        Me.Button_conferma = New System.Windows.Forms.Button()
        Me.TextBox_quantità = New System.Windows.Forms.TextBox()
        Me.Label_quantità_modifica = New System.Windows.Forms.Label()
        Me.Button_disegno = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'LabelDisegnoSAP
        '
        Me.LabelDisegnoSAP.AutoSize = True
        Me.LabelDisegnoSAP.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDisegnoSAP.Location = New System.Drawing.Point(160, 345)
        Me.LabelDisegnoSAP.Name = "LabelDisegnoSAP"
        Me.LabelDisegnoSAP.Size = New System.Drawing.Size(110, 20)
        Me.LabelDisegnoSAP.TabIndex = 74
        Me.LabelDisegnoSAP.Text = "DisegnoSAP"
        Me.LabelDisegnoSAP.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelDis
        '
        Me.LabelDis.AutoSize = True
        Me.LabelDis.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDis.Location = New System.Drawing.Point(58, 348)
        Me.LabelDis.Name = "LabelDis"
        Me.LabelDis.Size = New System.Drawing.Size(59, 16)
        Me.LabelDis.TabIndex = 73
        Me.LabelDis.Text = "Disegno"
        '
        'LabelDescrizioneSAP
        '
        Me.LabelDescrizioneSAP.AutoSize = True
        Me.LabelDescrizioneSAP.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDescrizioneSAP.Location = New System.Drawing.Point(158, 322)
        Me.LabelDescrizioneSAP.Name = "LabelDescrizioneSAP"
        Me.LabelDescrizioneSAP.Size = New System.Drawing.Size(97, 13)
        Me.LabelDescrizioneSAP.TabIndex = 72
        Me.LabelDescrizioneSAP.Text = "DescrizioneSAP"
        Me.LabelDescrizioneSAP.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'LabelCodiceSAP
        '
        Me.LabelCodiceSAP.AutoSize = True
        Me.LabelCodiceSAP.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCodiceSAP.Location = New System.Drawing.Point(158, 290)
        Me.LabelCodiceSAP.Name = "LabelCodiceSAP"
        Me.LabelCodiceSAP.Size = New System.Drawing.Size(99, 20)
        Me.LabelCodiceSAP.TabIndex = 71
        Me.LabelCodiceSAP.Text = "CodiceSAP"
        Me.LabelCodiceSAP.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LabelDesc
        '
        Me.LabelDesc.AutoSize = True
        Me.LabelDesc.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelDesc.Location = New System.Drawing.Point(56, 325)
        Me.LabelDesc.Name = "LabelDesc"
        Me.LabelDesc.Size = New System.Drawing.Size(79, 16)
        Me.LabelDesc.TabIndex = 70
        Me.LabelDesc.Text = "Descrizione"
        '
        'LabelCodice
        '
        Me.LabelCodice.AutoSize = True
        Me.LabelCodice.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LabelCodice.Location = New System.Drawing.Point(56, 290)
        Me.LabelCodice.Name = "LabelCodice"
        Me.LabelCodice.Size = New System.Drawing.Size(51, 16)
        Me.LabelCodice.TabIndex = 69
        Me.LabelCodice.Text = "Codice"
        '
        'Button_conferma
        '
        Me.Button_conferma.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_conferma.Location = New System.Drawing.Point(525, 213)
        Me.Button_conferma.Name = "Button_conferma"
        Me.Button_conferma.Size = New System.Drawing.Size(125, 66)
        Me.Button_conferma.TabIndex = 68
        Me.Button_conferma.Text = "CONFERMA"
        Me.Button_conferma.UseVisualStyleBackColor = True
        '
        'TextBox_quantità
        '
        Me.TextBox_quantità.Font = New System.Drawing.Font("Microsoft Sans Serif", 36.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox_quantità.Location = New System.Drawing.Point(293, 103)
        Me.TextBox_quantità.Name = "TextBox_quantità"
        Me.TextBox_quantità.Size = New System.Drawing.Size(226, 62)
        Me.TextBox_quantità.TabIndex = 67
        Me.TextBox_quantità.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label_quantità_modifica
        '
        Me.Label_quantità_modifica.AutoSize = True
        Me.Label_quantità_modifica.Font = New System.Drawing.Font("Microsoft Sans Serif", 24.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_quantità_modifica.Location = New System.Drawing.Point(138, 30)
        Me.Label_quantità_modifica.Name = "Label_quantità_modifica"
        Me.Label_quantità_modifica.Size = New System.Drawing.Size(567, 37)
        Me.Label_quantità_modifica.TabIndex = 66
        Me.Label_quantità_modifica.Text = "Indicare la quantità dei pezzi modificati"
        '
        'Button_disegno
        '
        Me.Button_disegno.Location = New System.Drawing.Point(276, 345)
        Me.Button_disegno.Name = "Button_disegno"
        Me.Button_disegno.Size = New System.Drawing.Size(75, 23)
        Me.Button_disegno.TabIndex = 96
        Me.Button_disegno.Text = "Disegno"
        Me.Button_disegno.UseVisualStyleBackColor = True
        '
        'Form107
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.Button_disegno)
        Me.Controls.Add(Me.LabelDisegnoSAP)
        Me.Controls.Add(Me.LabelDis)
        Me.Controls.Add(Me.LabelDescrizioneSAP)
        Me.Controls.Add(Me.LabelCodiceSAP)
        Me.Controls.Add(Me.LabelDesc)
        Me.Controls.Add(Me.LabelCodice)
        Me.Controls.Add(Me.Button_conferma)
        Me.Controls.Add(Me.TextBox_quantità)
        Me.Controls.Add(Me.Label_quantità_modifica)
        Me.Name = "Form107"
        Me.Text = "Form17"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LabelDisegnoSAP As Label
    Friend WithEvents LabelDis As Label
    Friend WithEvents LabelDescrizioneSAP As Label
    Friend WithEvents LabelCodiceSAP As Label
    Friend WithEvents LabelDesc As Label
    Friend WithEvents LabelCodice As Label
    Friend WithEvents Button_conferma As Button
    Friend WithEvents TextBox_quantità As TextBox
    Friend WithEvents Label_quantità_modifica As Label
    Friend WithEvents Button_disegno As Button
End Class
