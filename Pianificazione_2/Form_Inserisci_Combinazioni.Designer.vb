<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_Inserisci_Combinazioni
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
        Me.Cmd_Copia_Da = New System.Windows.Forms.Button()
        Me.Cmd_Elimina = New System.Windows.Forms.Button()
        Me.Cmd_Nuova = New System.Windows.Forms.Button()
        Me.Cmd_Azione = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Cmd_Canc = New System.Windows.Forms.Button()
        Me.Lbl_ID = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Txt_Vel_Effettiva = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Txt_Vel_Richiesta = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Lbl_Firma = New System.Windows.Forms.Label()
        Me.Combo_Dipendenti = New System.Windows.Forms.ComboBox()
        Me.Txt_Note = New System.Windows.Forms.TextBox()
        Me.Lbl_Note = New System.Windows.Forms.Label()
        Me.Check_Video = New System.Windows.Forms.CheckBox()
        Me.Check_Collaudato = New System.Windows.Forms.CheckBox()
        Me.Txt_Ricetta = New System.Windows.Forms.TextBox()
        Me.Lbl_Ricetta = New System.Windows.Forms.Label()
        Me.DataGrid_Combinazione = New System.Windows.Forms.DataGridView()
        Me.Nome = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Automatico = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Rimuovi = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Lst_Combinazioni = New System.Windows.Forms.ListBox()
        Me.Lbl_Commessa = New System.Windows.Forms.Label()
        Me.Grp_Campioni = New System.Windows.Forms.GroupBox()
        Me.Lst_Campioni = New System.Windows.Forms.ListBox()
        Me.Txt_Lista_Campioni = New System.Windows.Forms.TextBox()
        Me.Cmd_Esci = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Lst_BP = New System.Windows.Forms.ListBox()
        Me.Cmd_Cerca_BP = New System.Windows.Forms.Button()
        Me.TXT_Bp = New System.Windows.Forms.TextBox()
        Me.GroupBox3.SuspendLayout()
        CType(Me.DataGrid_Combinazione, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.Grp_Campioni.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Cmd_Copia_Da
        '
        Me.Cmd_Copia_Da.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Cmd_Copia_Da.Location = New System.Drawing.Point(630, 12)
        Me.Cmd_Copia_Da.Name = "Cmd_Copia_Da"
        Me.Cmd_Copia_Da.Size = New System.Drawing.Size(73, 40)
        Me.Cmd_Copia_Da.TabIndex = 20
        Me.Cmd_Copia_Da.Text = "Copia da ..."
        Me.Cmd_Copia_Da.UseVisualStyleBackColor = True
        '
        'Cmd_Elimina
        '
        Me.Cmd_Elimina.Location = New System.Drawing.Point(709, 12)
        Me.Cmd_Elimina.Name = "Cmd_Elimina"
        Me.Cmd_Elimina.Size = New System.Drawing.Size(73, 40)
        Me.Cmd_Elimina.TabIndex = 19
        Me.Cmd_Elimina.Text = "Elimina"
        Me.Cmd_Elimina.UseVisualStyleBackColor = True
        '
        'Cmd_Nuova
        '
        Me.Cmd_Nuova.Location = New System.Drawing.Point(788, 12)
        Me.Cmd_Nuova.Name = "Cmd_Nuova"
        Me.Cmd_Nuova.Size = New System.Drawing.Size(73, 40)
        Me.Cmd_Nuova.TabIndex = 18
        Me.Cmd_Nuova.Text = "Nuova"
        Me.Cmd_Nuova.UseVisualStyleBackColor = True
        '
        'Cmd_Azione
        '
        Me.Cmd_Azione.Location = New System.Drawing.Point(866, 12)
        Me.Cmd_Azione.Name = "Cmd_Azione"
        Me.Cmd_Azione.Size = New System.Drawing.Size(73, 40)
        Me.Cmd_Azione.TabIndex = 17
        Me.Cmd_Azione.Text = "Inserisci"
        Me.Cmd_Azione.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Cmd_Canc)
        Me.GroupBox3.Controls.Add(Me.Lbl_ID)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.Txt_Vel_Effettiva)
        Me.GroupBox3.Controls.Add(Me.Label2)
        Me.GroupBox3.Controls.Add(Me.Txt_Vel_Richiesta)
        Me.GroupBox3.Controls.Add(Me.Label1)
        Me.GroupBox3.Controls.Add(Me.Lbl_Firma)
        Me.GroupBox3.Controls.Add(Me.Combo_Dipendenti)
        Me.GroupBox3.Controls.Add(Me.Txt_Note)
        Me.GroupBox3.Controls.Add(Me.Lbl_Note)
        Me.GroupBox3.Controls.Add(Me.Check_Video)
        Me.GroupBox3.Controls.Add(Me.Check_Collaudato)
        Me.GroupBox3.Controls.Add(Me.Txt_Ricetta)
        Me.GroupBox3.Controls.Add(Me.Lbl_Ricetta)
        Me.GroupBox3.Controls.Add(Me.DataGrid_Combinazione)
        Me.GroupBox3.Location = New System.Drawing.Point(515, 58)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(503, 327)
        Me.GroupBox3.TabIndex = 16
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Combinazione Attuale"
        '
        'Cmd_Canc
        '
        Me.Cmd_Canc.Location = New System.Drawing.Point(438, 297)
        Me.Cmd_Canc.Name = "Cmd_Canc"
        Me.Cmd_Canc.Size = New System.Drawing.Size(52, 21)
        Me.Cmd_Canc.TabIndex = 16
        Me.Cmd_Canc.Text = "Canc."
        Me.Cmd_Canc.UseVisualStyleBackColor = True
        '
        'Lbl_ID
        '
        Me.Lbl_ID.AutoSize = True
        Me.Lbl_ID.Location = New System.Drawing.Point(417, 252)
        Me.Lbl_ID.Name = "Lbl_ID"
        Me.Lbl_ID.Size = New System.Drawing.Size(37, 13)
        Me.Lbl_ID.TabIndex = 15
        Me.Lbl_ID.Text = "12345"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(352, 252)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(27, 13)
        Me.Label3.TabIndex = 14
        Me.Label3.Text = "ID : "
        '
        'Txt_Vel_Effettiva
        '
        Me.Txt_Vel_Effettiva.Location = New System.Drawing.Point(390, 271)
        Me.Txt_Vel_Effettiva.Name = "Txt_Vel_Effettiva"
        Me.Txt_Vel_Effettiva.Size = New System.Drawing.Size(100, 20)
        Me.Txt_Vel_Effettiva.TabIndex = 13
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(292, 278)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(87, 13)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Velocità Effettiva"
        '
        'Txt_Vel_Richiesta
        '
        Me.Txt_Vel_Richiesta.Location = New System.Drawing.Point(106, 271)
        Me.Txt_Vel_Richiesta.Name = "Txt_Vel_Richiesta"
        Me.Txt_Vel_Richiesta.Size = New System.Drawing.Size(100, 20)
        Me.Txt_Vel_Richiesta.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 278)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(92, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Velocità Richiesta"
        '
        'Lbl_Firma
        '
        Me.Lbl_Firma.AutoSize = True
        Me.Lbl_Firma.Location = New System.Drawing.Point(8, 305)
        Me.Lbl_Firma.Name = "Lbl_Firma"
        Me.Lbl_Firma.Size = New System.Drawing.Size(32, 13)
        Me.Lbl_Firma.TabIndex = 9
        Me.Lbl_Firma.Text = "Firma"
        '
        'Combo_Dipendenti
        '
        Me.Combo_Dipendenti.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Combo_Dipendenti.FormattingEnabled = True
        Me.Combo_Dipendenti.Location = New System.Drawing.Point(44, 297)
        Me.Combo_Dipendenti.Name = "Combo_Dipendenti"
        Me.Combo_Dipendenti.Size = New System.Drawing.Size(388, 21)
        Me.Combo_Dipendenti.TabIndex = 8
        '
        'Txt_Note
        '
        Me.Txt_Note.Location = New System.Drawing.Point(44, 181)
        Me.Txt_Note.Multiline = True
        Me.Txt_Note.Name = "Txt_Note"
        Me.Txt_Note.Size = New System.Drawing.Size(302, 84)
        Me.Txt_Note.TabIndex = 7
        '
        'Lbl_Note
        '
        Me.Lbl_Note.AutoSize = True
        Me.Lbl_Note.Location = New System.Drawing.Point(8, 181)
        Me.Lbl_Note.Name = "Lbl_Note"
        Me.Lbl_Note.Size = New System.Drawing.Size(30, 13)
        Me.Lbl_Note.TabIndex = 6
        Me.Lbl_Note.Text = "Note"
        '
        'Check_Video
        '
        Me.Check_Video.AutoSize = True
        Me.Check_Video.Location = New System.Drawing.Point(352, 204)
        Me.Check_Video.Name = "Check_Video"
        Me.Check_Video.Size = New System.Drawing.Size(102, 17)
        Me.Check_Video.TabIndex = 5
        Me.Check_Video.Text = "Video Effettuato"
        Me.Check_Video.UseVisualStyleBackColor = True
        '
        'Check_Collaudato
        '
        Me.Check_Collaudato.AutoSize = True
        Me.Check_Collaudato.Location = New System.Drawing.Point(352, 181)
        Me.Check_Collaudato.Name = "Check_Collaudato"
        Me.Check_Collaudato.Size = New System.Drawing.Size(145, 17)
        Me.Check_Collaudato.TabIndex = 4
        Me.Check_Collaudato.Text = "Combinazione Collaudata"
        Me.Check_Collaudato.UseVisualStyleBackColor = True
        '
        'Txt_Ricetta
        '
        Me.Txt_Ricetta.Location = New System.Drawing.Point(396, 227)
        Me.Txt_Ricetta.Name = "Txt_Ricetta"
        Me.Txt_Ricetta.Size = New System.Drawing.Size(100, 20)
        Me.Txt_Ricetta.TabIndex = 2
        '
        'Lbl_Ricetta
        '
        Me.Lbl_Ricetta.AutoSize = True
        Me.Lbl_Ricetta.Location = New System.Drawing.Point(349, 230)
        Me.Lbl_Ricetta.Name = "Lbl_Ricetta"
        Me.Lbl_Ricetta.Size = New System.Drawing.Size(41, 13)
        Me.Lbl_Ricetta.TabIndex = 1
        Me.Lbl_Ricetta.Text = "Ricetta"
        '
        'DataGrid_Combinazione
        '
        Me.DataGrid_Combinazione.AllowUserToAddRows = False
        Me.DataGrid_Combinazione.AllowUserToDeleteRows = False
        Me.DataGrid_Combinazione.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGrid_Combinazione.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Nome, Me.Automatico, Me.Rimuovi})
        Me.DataGrid_Combinazione.Location = New System.Drawing.Point(7, 19)
        Me.DataGrid_Combinazione.Name = "DataGrid_Combinazione"
        Me.DataGrid_Combinazione.ReadOnly = True
        Me.DataGrid_Combinazione.RowHeadersVisible = False
        Me.DataGrid_Combinazione.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        Me.DataGrid_Combinazione.RowTemplate.Height = 44
        Me.DataGrid_Combinazione.Size = New System.Drawing.Size(490, 156)
        Me.DataGrid_Combinazione.TabIndex = 0
        '
        'Nome
        '
        Me.Nome.HeaderText = "Nome"
        Me.Nome.MinimumWidth = 6
        Me.Nome.Name = "Nome"
        Me.Nome.ReadOnly = True
        Me.Nome.Width = 125
        '
        'Automatico
        '
        Me.Automatico.HeaderText = "Automatico"
        Me.Automatico.MinimumWidth = 6
        Me.Automatico.Name = "Automatico"
        Me.Automatico.ReadOnly = True
        Me.Automatico.Width = 125
        '
        'Rimuovi
        '
        Me.Rimuovi.HeaderText = "Rimuovi"
        Me.Rimuovi.MinimumWidth = 6
        Me.Rimuovi.Name = "Rimuovi"
        Me.Rimuovi.ReadOnly = True
        Me.Rimuovi.Text = "Rimuovi"
        Me.Rimuovi.Width = 125
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Lst_Combinazioni)
        Me.GroupBox2.Location = New System.Drawing.Point(515, 391)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(503, 108)
        Me.GroupBox2.TabIndex = 15
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Elenco Conbinazioni"
        '
        'Lst_Combinazioni
        '
        Me.Lst_Combinazioni.FormattingEnabled = True
        Me.Lst_Combinazioni.Location = New System.Drawing.Point(11, 16)
        Me.Lst_Combinazioni.Name = "Lst_Combinazioni"
        Me.Lst_Combinazioni.Size = New System.Drawing.Size(486, 82)
        Me.Lst_Combinazioni.TabIndex = 0
        '
        'Lbl_Commessa
        '
        Me.Lbl_Commessa.AutoSize = True
        Me.Lbl_Commessa.Font = New System.Drawing.Font("Calibri", 21.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Lbl_Commessa.Location = New System.Drawing.Point(514, 9)
        Me.Lbl_Commessa.Name = "Lbl_Commessa"
        Me.Lbl_Commessa.Size = New System.Drawing.Size(100, 36)
        Me.Lbl_Commessa.TabIndex = 14
        Me.Lbl_Commessa.Text = "M0000"
        '
        'Grp_Campioni
        '
        Me.Grp_Campioni.Controls.Add(Me.Lst_Campioni)
        Me.Grp_Campioni.Controls.Add(Me.Txt_Lista_Campioni)
        Me.Grp_Campioni.Location = New System.Drawing.Point(263, 12)
        Me.Grp_Campioni.Name = "Grp_Campioni"
        Me.Grp_Campioni.Size = New System.Drawing.Size(245, 487)
        Me.Grp_Campioni.TabIndex = 13
        Me.Grp_Campioni.TabStop = False
        Me.Grp_Campioni.Text = "Campioni"
        '
        'Lst_Campioni
        '
        Me.Lst_Campioni.FormattingEnabled = True
        Me.Lst_Campioni.Location = New System.Drawing.Point(7, 72)
        Me.Lst_Campioni.Name = "Lst_Campioni"
        Me.Lst_Campioni.Size = New System.Drawing.Size(227, 407)
        Me.Lst_Campioni.TabIndex = 2
        '
        'Txt_Lista_Campioni
        '
        Me.Txt_Lista_Campioni.Location = New System.Drawing.Point(7, 20)
        Me.Txt_Lista_Campioni.Name = "Txt_Lista_Campioni"
        Me.Txt_Lista_Campioni.Size = New System.Drawing.Size(227, 20)
        Me.Txt_Lista_Campioni.TabIndex = 0
        '
        'Cmd_Esci
        '
        Me.Cmd_Esci.Location = New System.Drawing.Point(945, 12)
        Me.Cmd_Esci.Name = "Cmd_Esci"
        Me.Cmd_Esci.Size = New System.Drawing.Size(73, 40)
        Me.Cmd_Esci.TabIndex = 12
        Me.Cmd_Esci.Text = "Esci"
        Me.Cmd_Esci.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Lst_BP)
        Me.GroupBox1.Controls.Add(Me.Cmd_Cerca_BP)
        Me.GroupBox1.Controls.Add(Me.TXT_Bp)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(245, 487)
        Me.GroupBox1.TabIndex = 11
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Business Partner"
        '
        'Lst_BP
        '
        Me.Lst_BP.FormattingEnabled = True
        Me.Lst_BP.Location = New System.Drawing.Point(7, 72)
        Me.Lst_BP.Name = "Lst_BP"
        Me.Lst_BP.Size = New System.Drawing.Size(227, 407)
        Me.Lst_BP.TabIndex = 2
        '
        'Cmd_Cerca_BP
        '
        Me.Cmd_Cerca_BP.Location = New System.Drawing.Point(159, 43)
        Me.Cmd_Cerca_BP.Name = "Cmd_Cerca_BP"
        Me.Cmd_Cerca_BP.Size = New System.Drawing.Size(75, 23)
        Me.Cmd_Cerca_BP.TabIndex = 1
        Me.Cmd_Cerca_BP.Text = "&Cerca"
        Me.Cmd_Cerca_BP.UseVisualStyleBackColor = True
        '
        'TXT_Bp
        '
        Me.TXT_Bp.Location = New System.Drawing.Point(7, 20)
        Me.TXT_Bp.Name = "TXT_Bp"
        Me.TXT_Bp.Size = New System.Drawing.Size(227, 20)
        Me.TXT_Bp.TabIndex = 0
        '
        'Form_Inserisci_Combinazioni
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1024, 506)
        Me.ControlBox = False
        Me.Controls.Add(Me.Cmd_Copia_Da)
        Me.Controls.Add(Me.Cmd_Elimina)
        Me.Controls.Add(Me.Cmd_Nuova)
        Me.Controls.Add(Me.Cmd_Azione)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.Lbl_Commessa)
        Me.Controls.Add(Me.Grp_Campioni)
        Me.Controls.Add(Me.Cmd_Esci)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "Form_Inserisci_Combinazioni"
        Me.Text = "Form_Inserisci_Combinazioni"
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        CType(Me.DataGrid_Combinazione, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.Grp_Campioni.ResumeLayout(False)
        Me.Grp_Campioni.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Cmd_Copia_Da As Button
    Friend WithEvents Cmd_Elimina As Button
    Friend WithEvents Cmd_Nuova As Button
    Friend WithEvents Cmd_Azione As Button
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents Cmd_Canc As Button
    Friend WithEvents Lbl_ID As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Txt_Vel_Effettiva As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Txt_Vel_Richiesta As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Lbl_Firma As Label
    Friend WithEvents Combo_Dipendenti As ComboBox
    Friend WithEvents Txt_Note As TextBox
    Friend WithEvents Lbl_Note As Label
    Friend WithEvents Check_Video As CheckBox
    Friend WithEvents Check_Collaudato As CheckBox
    Friend WithEvents Txt_Ricetta As TextBox
    Friend WithEvents Lbl_Ricetta As Label
    Friend WithEvents DataGrid_Combinazione As DataGridView
    Friend WithEvents Nome As DataGridViewTextBoxColumn
    Friend WithEvents Automatico As DataGridViewCheckBoxColumn
    Friend WithEvents Rimuovi As DataGridViewButtonColumn
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Lst_Combinazioni As ListBox
    Friend WithEvents Lbl_Commessa As Label
    Friend WithEvents Grp_Campioni As GroupBox
    Friend WithEvents Lst_Campioni As ListBox
    Friend WithEvents Txt_Lista_Campioni As TextBox
    Friend WithEvents Cmd_Esci As Button
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Lst_BP As ListBox
    Friend WithEvents Cmd_Cerca_BP As Button
    Friend WithEvents TXT_Bp As TextBox
End Class
