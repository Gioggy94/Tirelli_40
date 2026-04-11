<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Carica_Video
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
        Me.Cmd_Guarda_Video = New System.Windows.Forms.Button()
        Me.Cmd_Salva_Video = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Cmd_Sfoglia = New System.Windows.Forms.Button()
        Me.TXT_Input = New System.Windows.Forms.TextBox()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TXT_Info = New System.Windows.Forms.TextBox()
        Me.Cmd_Esci = New System.Windows.Forms.Button()
        Me.Crp_Commessa = New System.Windows.Forms.GroupBox()
        Me.Lbl_Nome_File = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.Crp_Commessa.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Cmd_Guarda_Video
        '
        Me.Cmd_Guarda_Video.Location = New System.Drawing.Point(446, 175)
        Me.Cmd_Guarda_Video.Name = "Cmd_Guarda_Video"
        Me.Cmd_Guarda_Video.Size = New System.Drawing.Size(110, 52)
        Me.Cmd_Guarda_Video.TabIndex = 17
        Me.Cmd_Guarda_Video.Text = "&Guarda Video"
        Me.Cmd_Guarda_Video.UseVisualStyleBackColor = True
        '
        'Cmd_Salva_Video
        '
        Me.Cmd_Salva_Video.Location = New System.Drawing.Point(562, 175)
        Me.Cmd_Salva_Video.Name = "Cmd_Salva_Video"
        Me.Cmd_Salva_Video.Size = New System.Drawing.Size(110, 52)
        Me.Cmd_Salva_Video.TabIndex = 16
        Me.Cmd_Salva_Video.Text = "&Salva Video"
        Me.Cmd_Salva_Video.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Cmd_Sfoglia)
        Me.GroupBox2.Controls.Add(Me.TXT_Input)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 121)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(776, 48)
        Me.GroupBox2.TabIndex = 14
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Seleziona File"
        '
        'Cmd_Sfoglia
        '
        Me.Cmd_Sfoglia.Location = New System.Drawing.Point(686, 14)
        Me.Cmd_Sfoglia.Name = "Cmd_Sfoglia"
        Me.Cmd_Sfoglia.Size = New System.Drawing.Size(84, 27)
        Me.Cmd_Sfoglia.TabIndex = 9
        Me.Cmd_Sfoglia.Text = "&Sfoglia"
        Me.Cmd_Sfoglia.UseVisualStyleBackColor = True
        '
        'TXT_Input
        '
        Me.TXT_Input.Enabled = False
        Me.TXT_Input.Location = New System.Drawing.Point(19, 19)
        Me.TXT_Input.Name = "TXT_Input"
        Me.TXT_Input.Size = New System.Drawing.Size(661, 20)
        Me.TXT_Input.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TXT_Info)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 65)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(776, 48)
        Me.GroupBox1.TabIndex = 13
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Informazioni Aggiuntive (es. Fronte, Retro, Distribuzione, ecc.)"
        '
        'TXT_Info
        '
        Me.TXT_Info.Location = New System.Drawing.Point(19, 19)
        Me.TXT_Info.Name = "TXT_Info"
        Me.TXT_Info.Size = New System.Drawing.Size(743, 20)
        Me.TXT_Info.TabIndex = 0
        '
        'Cmd_Esci
        '
        Me.Cmd_Esci.Location = New System.Drawing.Point(678, 175)
        Me.Cmd_Esci.Name = "Cmd_Esci"
        Me.Cmd_Esci.Size = New System.Drawing.Size(110, 52)
        Me.Cmd_Esci.TabIndex = 15
        Me.Cmd_Esci.Text = "&Esci"
        Me.Cmd_Esci.UseVisualStyleBackColor = True
        '
        'Crp_Commessa
        '
        Me.Crp_Commessa.Controls.Add(Me.Lbl_Nome_File)
        Me.Crp_Commessa.Location = New System.Drawing.Point(12, 12)
        Me.Crp_Commessa.Name = "Crp_Commessa"
        Me.Crp_Commessa.Size = New System.Drawing.Size(776, 47)
        Me.Crp_Commessa.TabIndex = 12
        Me.Crp_Commessa.TabStop = False
        Me.Crp_Commessa.Text = "Nome del File"
        '
        'Lbl_Nome_File
        '
        Me.Lbl_Nome_File.AutoSize = True
        Me.Lbl_Nome_File.Location = New System.Drawing.Point(16, 23)
        Me.Lbl_Nome_File.Name = "Lbl_Nome_File"
        Me.Lbl_Nome_File.Size = New System.Drawing.Size(57, 13)
        Me.Lbl_Nome_File.TabIndex = 5
        Me.Lbl_Nome_File.Text = "Nome_File"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.TextBox1)
        Me.GroupBox3.Location = New System.Drawing.Point(309, 175)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(131, 54)
        Me.GroupBox3.TabIndex = 18
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Bit rate (qualità video)"
        Me.GroupBox3.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBox1.Location = New System.Drawing.Point(3, 16)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(125, 20)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = "2000"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 48.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(166, 243)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(464, 73)
        Me.Label1.TabIndex = 19
        Me.Label1.Text = "Caricamento..."
        Me.Label1.Visible = False
        '
        'Form_Carica_Video
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(810, 338)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.Cmd_Guarda_Video)
        Me.Controls.Add(Me.Cmd_Salva_Video)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Cmd_Esci)
        Me.Controls.Add(Me.Crp_Commessa)
        Me.Name = "Form_Carica_Video"
        Me.Text = "Form_Carica_Video"
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.Crp_Commessa.ResumeLayout(False)
        Me.Crp_Commessa.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Cmd_Guarda_Video As Button
    Friend WithEvents Cmd_Salva_Video As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Cmd_Sfoglia As Button
    Friend WithEvents TXT_Input As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents TXT_Info As TextBox
    Friend WithEvents Cmd_Esci As Button
    Friend WithEvents Crp_Commessa As GroupBox
    Friend WithEvents Lbl_Nome_File As Label
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents Label1 As Label
End Class
