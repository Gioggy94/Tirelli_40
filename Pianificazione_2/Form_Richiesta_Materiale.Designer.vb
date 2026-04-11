<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Richiesta_Materiale
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TXT_ODP = New System.Windows.Forms.TextBox()
        Me.DataGrid_Materiale = New System.Windows.Forms.DataGridView()
        Me.Grp_Lunghezza = New System.Windows.Forms.GroupBox()
        Me.Txt_Lunghezza = New System.Windows.Forms.TextBox()
        Me.Grp_Qta = New System.Windows.Forms.GroupBox()
        Me.Txt_Qta = New System.Windows.Forms.TextBox()
        Me.Grp_Materiale = New System.Windows.Forms.GroupBox()
        Me.List_Materiale = New System.Windows.Forms.ListBox()
        Me.Cmd_Invia = New System.Windows.Forms.Button()
        Me.Cmd_Annulla = New System.Windows.Forms.Button()
        Me.Cmd_Home = New System.Windows.Forms.Button()
        Me.Grp_Descrizione = New System.Windows.Forms.GroupBox()
        Me.Txt_Descrizione = New System.Windows.Forms.TextBox()
        Me.Txt_Codice = New System.Windows.Forms.TextBox()
        Me.Cmd_Aggiungi = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox14 = New System.Windows.Forms.GroupBox()
        Me.Combo_Utente = New System.Windows.Forms.ComboBox()
        Me.GroupBox13 = New System.Windows.Forms.GroupBox()
        Me.Combo_Mittente = New System.Windows.Forms.ComboBox()
        Me.Grp_Commessa = New System.Windows.Forms.GroupBox()
        Me.Txt_Commessa = New System.Windows.Forms.TextBox()
        Me.GroupBox1.SuspendLayout()
        CType(Me.DataGrid_Materiale, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Grp_Lunghezza.SuspendLayout()
        Me.Grp_Qta.SuspendLayout()
        Me.Grp_Materiale.SuspendLayout()
        Me.Grp_Descrizione.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox14.SuspendLayout()
        Me.GroupBox13.SuspendLayout()
        Me.Grp_Commessa.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TXT_ODP)
        Me.GroupBox1.Location = New System.Drawing.Point(13, 13)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(344, 51)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Ordine di Produzione"
        '
        'TXT_ODP
        '
        Me.TXT_ODP.Enabled = False
        Me.TXT_ODP.Location = New System.Drawing.Point(7, 20)
        Me.TXT_ODP.Name = "TXT_ODP"
        Me.TXT_ODP.Size = New System.Drawing.Size(321, 20)
        Me.TXT_ODP.TabIndex = 0
        '
        'DataGrid_Materiale
        '
        Me.DataGrid_Materiale.AllowUserToAddRows = False
        Me.DataGrid_Materiale.AllowUserToDeleteRows = False
        Me.DataGrid_Materiale.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGrid_Materiale.Location = New System.Drawing.Point(13, 124)
        Me.DataGrid_Materiale.Name = "DataGrid_Materiale"
        Me.DataGrid_Materiale.ReadOnly = True
        Me.DataGrid_Materiale.Size = New System.Drawing.Size(1422, 476)
        Me.DataGrid_Materiale.TabIndex = 1
        '
        'Grp_Lunghezza
        '
        Me.Grp_Lunghezza.Controls.Add(Me.Txt_Lunghezza)
        Me.Grp_Lunghezza.Location = New System.Drawing.Point(12, 684)
        Me.Grp_Lunghezza.Name = "Grp_Lunghezza"
        Me.Grp_Lunghezza.Size = New System.Drawing.Size(218, 60)
        Me.Grp_Lunghezza.TabIndex = 2
        Me.Grp_Lunghezza.TabStop = False
        Me.Grp_Lunghezza.Text = "Lunghezza"
        '
        'Txt_Lunghezza
        '
        Me.Txt_Lunghezza.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Txt_Lunghezza.Location = New System.Drawing.Point(7, 19)
        Me.Txt_Lunghezza.Name = "Txt_Lunghezza"
        Me.Txt_Lunghezza.Size = New System.Drawing.Size(197, 31)
        Me.Txt_Lunghezza.TabIndex = 0
        '
        'Grp_Qta
        '
        Me.Grp_Qta.Controls.Add(Me.Txt_Qta)
        Me.Grp_Qta.Location = New System.Drawing.Point(13, 750)
        Me.Grp_Qta.Name = "Grp_Qta"
        Me.Grp_Qta.Size = New System.Drawing.Size(217, 66)
        Me.Grp_Qta.TabIndex = 3
        Me.Grp_Qta.TabStop = False
        Me.Grp_Qta.Text = "Quantità"
        '
        'Txt_Qta
        '
        Me.Txt_Qta.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Txt_Qta.Location = New System.Drawing.Point(7, 20)
        Me.Txt_Qta.Name = "Txt_Qta"
        Me.Txt_Qta.Size = New System.Drawing.Size(197, 31)
        Me.Txt_Qta.TabIndex = 0
        '
        'Grp_Materiale
        '
        Me.Grp_Materiale.Controls.Add(Me.List_Materiale)
        Me.Grp_Materiale.Location = New System.Drawing.Point(406, 606)
        Me.Grp_Materiale.Name = "Grp_Materiale"
        Me.Grp_Materiale.Size = New System.Drawing.Size(860, 210)
        Me.Grp_Materiale.TabIndex = 4
        Me.Grp_Materiale.TabStop = False
        Me.Grp_Materiale.Text = "Materiale"
        '
        'List_Materiale
        '
        Me.List_Materiale.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.List_Materiale.FormattingEnabled = True
        Me.List_Materiale.ItemHeight = 25
        Me.List_Materiale.Location = New System.Drawing.Point(7, 20)
        Me.List_Materiale.Name = "List_Materiale"
        Me.List_Materiale.Size = New System.Drawing.Size(847, 179)
        Me.List_Materiale.TabIndex = 0
        '
        'Cmd_Invia
        '
        Me.Cmd_Invia.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cmd_Invia.Location = New System.Drawing.Point(1272, 750)
        Me.Cmd_Invia.Name = "Cmd_Invia"
        Me.Cmd_Invia.Size = New System.Drawing.Size(163, 52)
        Me.Cmd_Invia.TabIndex = 5
        Me.Cmd_Invia.Text = "Invia"
        Me.Cmd_Invia.UseVisualStyleBackColor = True
        '
        'Cmd_Annulla
        '
        Me.Cmd_Annulla.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cmd_Annulla.Location = New System.Drawing.Point(1272, 692)
        Me.Cmd_Annulla.Name = "Cmd_Annulla"
        Me.Cmd_Annulla.Size = New System.Drawing.Size(163, 52)
        Me.Cmd_Annulla.TabIndex = 6
        Me.Cmd_Annulla.Text = "Annulla"
        Me.Cmd_Annulla.UseVisualStyleBackColor = True
        '
        'Cmd_Home
        '
        Me.Cmd_Home.Font = New System.Drawing.Font("Webdings", 36.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(2, Byte))
        Me.Cmd_Home.Location = New System.Drawing.Point(1286, 12)
        Me.Cmd_Home.Name = "Cmd_Home"
        Me.Cmd_Home.Size = New System.Drawing.Size(149, 94)
        Me.Cmd_Home.TabIndex = 7
        Me.Cmd_Home.Text = "H"
        Me.Cmd_Home.UseVisualStyleBackColor = True
        '
        'Grp_Descrizione
        '
        Me.Grp_Descrizione.Controls.Add(Me.Txt_Descrizione)
        Me.Grp_Descrizione.Controls.Add(Me.Txt_Codice)
        Me.Grp_Descrizione.Location = New System.Drawing.Point(13, 606)
        Me.Grp_Descrizione.Name = "Grp_Descrizione"
        Me.Grp_Descrizione.Size = New System.Drawing.Size(387, 72)
        Me.Grp_Descrizione.TabIndex = 3
        Me.Grp_Descrizione.TabStop = False
        Me.Grp_Descrizione.Text = "Selezione"
        '
        'Txt_Descrizione
        '
        Me.Txt_Descrizione.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Txt_Descrizione.Location = New System.Drawing.Point(7, 44)
        Me.Txt_Descrizione.Name = "Txt_Descrizione"
        Me.Txt_Descrizione.Size = New System.Drawing.Size(374, 22)
        Me.Txt_Descrizione.TabIndex = 1
        '
        'Txt_Codice
        '
        Me.Txt_Codice.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Txt_Codice.Location = New System.Drawing.Point(7, 19)
        Me.Txt_Codice.Name = "Txt_Codice"
        Me.Txt_Codice.Size = New System.Drawing.Size(374, 22)
        Me.Txt_Codice.TabIndex = 0
        '
        'Cmd_Aggiungi
        '
        Me.Cmd_Aggiungi.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cmd_Aggiungi.Location = New System.Drawing.Point(236, 684)
        Me.Cmd_Aggiungi.Name = "Cmd_Aggiungi"
        Me.Cmd_Aggiungi.Size = New System.Drawing.Size(163, 132)
        Me.Cmd_Aggiungi.TabIndex = 8
        Me.Cmd_Aggiungi.Text = "Aggiungi"
        Me.Cmd_Aggiungi.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.GroupBox14)
        Me.GroupBox2.Controls.Add(Me.GroupBox13)
        Me.GroupBox2.Location = New System.Drawing.Point(363, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(917, 75)
        Me.GroupBox2.TabIndex = 48
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Mittente"
        '
        'GroupBox14
        '
        Me.GroupBox14.Controls.Add(Me.Combo_Utente)
        Me.GroupBox14.Location = New System.Drawing.Point(413, 13)
        Me.GroupBox14.Name = "GroupBox14"
        Me.GroupBox14.Size = New System.Drawing.Size(498, 56)
        Me.GroupBox14.TabIndex = 1
        Me.GroupBox14.TabStop = False
        Me.GroupBox14.Text = "Utente"
        '
        'Combo_Utente
        '
        Me.Combo_Utente.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_Utente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Utente.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_Utente.FormattingEnabled = True
        Me.Combo_Utente.Location = New System.Drawing.Point(6, 16)
        Me.Combo_Utente.Name = "Combo_Utente"
        Me.Combo_Utente.Size = New System.Drawing.Size(486, 28)
        Me.Combo_Utente.TabIndex = 2
        '
        'GroupBox13
        '
        Me.GroupBox13.Controls.Add(Me.Combo_Mittente)
        Me.GroupBox13.Location = New System.Drawing.Point(6, 13)
        Me.GroupBox13.Name = "GroupBox13"
        Me.GroupBox13.Size = New System.Drawing.Size(401, 56)
        Me.GroupBox13.TabIndex = 0
        Me.GroupBox13.TabStop = False
        Me.GroupBox13.Text = "Reparto"
        '
        'Combo_Mittente
        '
        Me.Combo_Mittente.BackColor = System.Drawing.SystemColors.Window
        Me.Combo_Mittente.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Mittente.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Combo_Mittente.FormattingEnabled = True
        Me.Combo_Mittente.Location = New System.Drawing.Point(6, 17)
        Me.Combo_Mittente.Name = "Combo_Mittente"
        Me.Combo_Mittente.Size = New System.Drawing.Size(389, 28)
        Me.Combo_Mittente.TabIndex = 1
        '
        'Grp_Commessa
        '
        Me.Grp_Commessa.Controls.Add(Me.Txt_Commessa)
        Me.Grp_Commessa.Location = New System.Drawing.Point(12, 67)
        Me.Grp_Commessa.Name = "Grp_Commessa"
        Me.Grp_Commessa.Size = New System.Drawing.Size(344, 51)
        Me.Grp_Commessa.TabIndex = 1
        Me.Grp_Commessa.TabStop = False
        Me.Grp_Commessa.Text = "Commessa"
        '
        'Txt_Commessa
        '
        Me.Txt_Commessa.Enabled = False
        Me.Txt_Commessa.Location = New System.Drawing.Point(7, 20)
        Me.Txt_Commessa.Name = "Txt_Commessa"
        Me.Txt_Commessa.Size = New System.Drawing.Size(321, 20)
        Me.Txt_Commessa.TabIndex = 0
        '
        'Form_Richiesta_Materiale
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1447, 828)
        Me.Controls.Add(Me.Grp_Commessa)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Cmd_Aggiungi)
        Me.Controls.Add(Me.Grp_Descrizione)
        Me.Controls.Add(Me.Cmd_Home)
        Me.Controls.Add(Me.Cmd_Annulla)
        Me.Controls.Add(Me.Cmd_Invia)
        Me.Controls.Add(Me.Grp_Materiale)
        Me.Controls.Add(Me.Grp_Qta)
        Me.Controls.Add(Me.Grp_Lunghezza)
        Me.Controls.Add(Me.DataGrid_Materiale)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "Form_Richiesta_Materiale"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Richiesta Materiale"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.DataGrid_Materiale, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Grp_Lunghezza.ResumeLayout(False)
        Me.Grp_Lunghezza.PerformLayout()
        Me.Grp_Qta.ResumeLayout(False)
        Me.Grp_Qta.PerformLayout()
        Me.Grp_Materiale.ResumeLayout(False)
        Me.Grp_Descrizione.ResumeLayout(False)
        Me.Grp_Descrizione.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox14.ResumeLayout(False)
        Me.GroupBox13.ResumeLayout(False)
        Me.Grp_Commessa.ResumeLayout(False)
        Me.Grp_Commessa.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents TXT_ODP As TextBox
    Friend WithEvents DataGrid_Materiale As DataGridView
    Friend WithEvents Grp_Lunghezza As GroupBox
    Friend WithEvents Txt_Lunghezza As TextBox
    Friend WithEvents Grp_Qta As GroupBox
    Friend WithEvents Txt_Qta As TextBox
    Friend WithEvents Grp_Materiale As GroupBox
    Friend WithEvents List_Materiale As ListBox
    Friend WithEvents Cmd_Invia As Button
    Friend WithEvents Cmd_Annulla As Button
    Friend WithEvents Cmd_Home As Button
    Friend WithEvents Grp_Descrizione As GroupBox
    Friend WithEvents Txt_Descrizione As TextBox
    Friend WithEvents Txt_Codice As TextBox
    Friend WithEvents Cmd_Aggiungi As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox14 As GroupBox
    Friend WithEvents Combo_Utente As ComboBox
    Friend WithEvents GroupBox13 As GroupBox
    Friend WithEvents Combo_Mittente As ComboBox
    Friend WithEvents Grp_Commessa As GroupBox
    Friend WithEvents Txt_Commessa As TextBox
End Class
