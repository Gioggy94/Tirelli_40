<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Lavorazioni_MES
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Lavorazioni_MES))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.ComboBox_dipendente = New System.Windows.Forms.ComboBox()
        Me.Label_gruppo_articolo_F = New System.Windows.Forms.Label()
        Me.Label_gruppo_Articoli = New System.Windows.Forms.Label()
        Me.Label_disegno_F = New System.Windows.Forms.Label()
        Me.Label_Disegno = New System.Windows.Forms.Label()
        Me.Label_codice_ODP = New System.Windows.Forms.Label()
        Me.Label_Codice_ODP_F = New System.Windows.Forms.Label()
        Me.Label_fase_F = New System.Windows.Forms.Label()
        Me.Label_Fase = New System.Windows.Forms.Label()
        Me.Label_commessa_F = New System.Windows.Forms.Label()
        Me.Label_commessa = New System.Windows.Forms.Label()
        Me.Label_descrizione = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label_numero_odp = New System.Windows.Forms.Label()
        Me.Label_numero_ODP_F = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ComboBox_risorse = New System.Windows.Forms.ComboBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.Button_stop = New System.Windows.Forms.Button()
        Me.Button_start = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.Panel9 = New System.Windows.Forms.Panel()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.DataGridView_lavorazioni = New System.Windows.Forms.DataGridView()
        Me.N_ODP = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Documento = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ODP = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Codice = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Descrizione = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Disegno = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Quantità = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Commessa = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Nome_macchina = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Cliente = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Dipendente = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Risorsa = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Data = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Start = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.Panel6.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.Panel8.SuspendLayout()
        Me.Panel9.SuspendLayout()
        Me.Panel7.SuspendLayout()
        CType(Me.DataGridView_lavorazioni, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ComboBox_dipendente
        '
        Me.ComboBox_dipendente.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ComboBox_dipendente.Font = New System.Drawing.Font("Microsoft Sans Serif", 21.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox_dipendente.FormattingEnabled = True
        Me.ComboBox_dipendente.Location = New System.Drawing.Point(3, 16)
        Me.ComboBox_dipendente.Name = "ComboBox_dipendente"
        Me.ComboBox_dipendente.Size = New System.Drawing.Size(1081, 41)
        Me.ComboBox_dipendente.TabIndex = 80
        '
        'Label_gruppo_articolo_F
        '
        Me.Label_gruppo_articolo_F.AutoEllipsis = True
        Me.Label_gruppo_articolo_F.AutoSize = True
        Me.Label_gruppo_articolo_F.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_gruppo_articolo_F.ForeColor = System.Drawing.Color.Black
        Me.Label_gruppo_articolo_F.Location = New System.Drawing.Point(82, 79)
        Me.Label_gruppo_articolo_F.Name = "Label_gruppo_articolo_F"
        Me.Label_gruppo_articolo_F.Size = New System.Drawing.Size(83, 15)
        Me.Label_gruppo_articolo_F.TabIndex = 136
        Me.Label_gruppo_articolo_F.Text = "Descrizione"
        '
        'Label_gruppo_Articoli
        '
        Me.Label_gruppo_Articoli.AutoSize = True
        Me.Label_gruppo_Articoli.ForeColor = System.Drawing.Color.Black
        Me.Label_gruppo_Articoli.Location = New System.Drawing.Point(6, 79)
        Me.Label_gruppo_Articoli.Name = "Label_gruppo_Articoli"
        Me.Label_gruppo_Articoli.Size = New System.Drawing.Size(79, 13)
        Me.Label_gruppo_Articoli.TabIndex = 135
        Me.Label_gruppo_Articoli.Text = "Gruppo articolo"
        '
        'Label_disegno_F
        '
        Me.Label_disegno_F.AutoEllipsis = True
        Me.Label_disegno_F.AutoSize = True
        Me.Label_disegno_F.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_disegno_F.ForeColor = System.Drawing.Color.Black
        Me.Label_disegno_F.Location = New System.Drawing.Point(82, 99)
        Me.Label_disegno_F.Name = "Label_disegno_F"
        Me.Label_disegno_F.Size = New System.Drawing.Size(60, 15)
        Me.Label_disegno_F.TabIndex = 134
        Me.Label_disegno_F.Text = "Disegno"
        '
        'Label_Disegno
        '
        Me.Label_Disegno.AutoSize = True
        Me.Label_Disegno.ForeColor = System.Drawing.Color.Black
        Me.Label_Disegno.Location = New System.Drawing.Point(6, 99)
        Me.Label_Disegno.Name = "Label_Disegno"
        Me.Label_Disegno.Size = New System.Drawing.Size(46, 13)
        Me.Label_Disegno.TabIndex = 133
        Me.Label_Disegno.Text = "Disegno"
        '
        'Label_codice_ODP
        '
        Me.Label_codice_ODP.AutoSize = True
        Me.Label_codice_ODP.ForeColor = System.Drawing.Color.Black
        Me.Label_codice_ODP.Location = New System.Drawing.Point(6, 39)
        Me.Label_codice_ODP.Name = "Label_codice_ODP"
        Me.Label_codice_ODP.Size = New System.Drawing.Size(66, 13)
        Me.Label_codice_ODP.TabIndex = 132
        Me.Label_codice_ODP.Text = "Codice ODP"
        '
        'Label_Codice_ODP_F
        '
        Me.Label_Codice_ODP_F.AutoEllipsis = True
        Me.Label_Codice_ODP_F.AutoSize = True
        Me.Label_Codice_ODP_F.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_Codice_ODP_F.ForeColor = System.Drawing.Color.Black
        Me.Label_Codice_ODP_F.Location = New System.Drawing.Point(82, 39)
        Me.Label_Codice_ODP_F.Name = "Label_Codice_ODP_F"
        Me.Label_Codice_ODP_F.Size = New System.Drawing.Size(84, 15)
        Me.Label_Codice_ODP_F.TabIndex = 131
        Me.Label_Codice_ODP_F.Text = "Codice ODP"
        '
        'Label_fase_F
        '
        Me.Label_fase_F.AutoEllipsis = True
        Me.Label_fase_F.AutoSize = True
        Me.Label_fase_F.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_fase_F.ForeColor = System.Drawing.Color.Black
        Me.Label_fase_F.Location = New System.Drawing.Point(82, 139)
        Me.Label_fase_F.Name = "Label_fase_F"
        Me.Label_fase_F.Size = New System.Drawing.Size(38, 15)
        Me.Label_fase_F.TabIndex = 130
        Me.Label_fase_F.Text = "Fase"
        '
        'Label_Fase
        '
        Me.Label_Fase.AutoSize = True
        Me.Label_Fase.ForeColor = System.Drawing.Color.Black
        Me.Label_Fase.Location = New System.Drawing.Point(6, 139)
        Me.Label_Fase.Name = "Label_Fase"
        Me.Label_Fase.Size = New System.Drawing.Size(30, 13)
        Me.Label_Fase.TabIndex = 129
        Me.Label_Fase.Text = "Fase"
        '
        'Label_commessa_F
        '
        Me.Label_commessa_F.AutoEllipsis = True
        Me.Label_commessa_F.AutoSize = True
        Me.Label_commessa_F.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_commessa_F.ForeColor = System.Drawing.Color.Black
        Me.Label_commessa_F.Location = New System.Drawing.Point(82, 119)
        Me.Label_commessa_F.Name = "Label_commessa_F"
        Me.Label_commessa_F.Size = New System.Drawing.Size(97, 15)
        Me.Label_commessa_F.TabIndex = 128
        Me.Label_commessa_F.Text = "Ordine cliente"
        '
        'Label_commessa
        '
        Me.Label_commessa.AutoSize = True
        Me.Label_commessa.ForeColor = System.Drawing.Color.Black
        Me.Label_commessa.Location = New System.Drawing.Point(6, 119)
        Me.Label_commessa.Name = "Label_commessa"
        Me.Label_commessa.Size = New System.Drawing.Size(58, 13)
        Me.Label_commessa.TabIndex = 127
        Me.Label_commessa.Text = "Commessa"
        '
        'Label_descrizione
        '
        Me.Label_descrizione.AutoEllipsis = True
        Me.Label_descrizione.AutoSize = True
        Me.Label_descrizione.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_descrizione.ForeColor = System.Drawing.Color.Black
        Me.Label_descrizione.Location = New System.Drawing.Point(82, 59)
        Me.Label_descrizione.Name = "Label_descrizione"
        Me.Label_descrizione.Size = New System.Drawing.Size(83, 15)
        Me.Label_descrizione.TabIndex = 126
        Me.Label_descrizione.Text = "Descrizione"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(6, 59)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(62, 13)
        Me.Label3.TabIndex = 125
        Me.Label3.Text = "Descrizione"
        '
        'Label_numero_odp
        '
        Me.Label_numero_odp.AutoSize = True
        Me.Label_numero_odp.ForeColor = System.Drawing.Color.Black
        Me.Label_numero_odp.Location = New System.Drawing.Point(6, 19)
        Me.Label_numero_odp.Name = "Label_numero_odp"
        Me.Label_numero_odp.Size = New System.Drawing.Size(70, 13)
        Me.Label_numero_odp.TabIndex = 124
        Me.Label_numero_odp.Text = "Numero ODP"
        '
        'Label_numero_ODP_F
        '
        Me.Label_numero_ODP_F.AutoEllipsis = True
        Me.Label_numero_ODP_F.AutoSize = True
        Me.Label_numero_ODP_F.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label_numero_ODP_F.ForeColor = System.Drawing.Color.Black
        Me.Label_numero_ODP_F.Location = New System.Drawing.Point(82, 19)
        Me.Label_numero_ODP_F.Name = "Label_numero_ODP_F"
        Me.Label_numero_ODP_F.Size = New System.Drawing.Size(17, 15)
        Me.Label_numero_ODP_F.TabIndex = 123
        Me.Label_numero_ODP_F.Text = "N"
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.SteelBlue
        Me.GroupBox1.Controls.Add(Me.ComboBox_dipendente)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1087, 107)
        Me.GroupBox1.TabIndex = 141
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Dipendente"
        '
        'ComboBox_risorse
        '
        Me.ComboBox_risorse.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ComboBox_risorse.Font = New System.Drawing.Font("Microsoft Sans Serif", 20.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComboBox_risorse.FormattingEnabled = True
        Me.ComboBox_risorse.Location = New System.Drawing.Point(3, 16)
        Me.ComboBox_risorse.Name = "ComboBox_risorse"
        Me.ComboBox_risorse.Size = New System.Drawing.Size(1081, 39)
        Me.ComboBox_risorse.TabIndex = 138
        '
        'GroupBox2
        '
        Me.GroupBox2.BackColor = System.Drawing.Color.SteelBlue
        Me.GroupBox2.Controls.Add(Me.ComboBox_risorse)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.ForeColor = System.Drawing.Color.White
        Me.GroupBox2.Location = New System.Drawing.Point(3, 116)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(1087, 107)
        Me.GroupBox2.TabIndex = 142
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Risorsa"
        '
        'Button1
        '
        Me.Button1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(0, 0)
        Me.Button1.Margin = New System.Windows.Forms.Padding(2)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(64, 54)
        Me.Button1.TabIndex = 144
        Me.Button1.Text = "X"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel1.Controls.Add(Me.Button5)
        Me.Panel1.Controls.Add(Me.Button4)
        Me.Panel1.Controls.Add(Me.Panel6)
        Me.Panel1.Controls.Add(Me.Panel5)
        Me.Panel1.Controls.Add(Me.Panel2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(251, 857)
        Me.Panel1.TabIndex = 145
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(56, 585)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(142, 52)
        Me.Button5.TabIndex = 144
        Me.Button5.Text = "Statistiche"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(49, 510)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(150, 65)
        Me.Button4.TabIndex = 143
        Me.Button4.Text = "Visualizza Excel storico lavorazioni"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.GroupBox5)
        Me.Panel6.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel6.Location = New System.Drawing.Point(0, 335)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(251, 109)
        Me.Panel6.TabIndex = 142
        Me.Panel6.Visible = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Button2)
        Me.GroupBox5.Controls.Add(Me.Button3)
        Me.GroupBox5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox5.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox5.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox5.Size = New System.Drawing.Size(251, 109)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Lavorazioni OC"
        '
        'Button2
        '
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.ForeColor = System.Drawing.Color.White
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Location = New System.Drawing.Point(123, 15)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(126, 92)
        Me.Button2.TabIndex = 142
        Me.Button2.Text = "STOP"
        Me.Button2.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button3.Dock = System.Windows.Forms.DockStyle.Left
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.ForeColor = System.Drawing.Color.White
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.Location = New System.Drawing.Point(2, 15)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(121, 92)
        Me.Button3.TabIndex = 140
        Me.Button3.Text = "START"
        Me.Button3.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.GroupBox4)
        Me.Panel5.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel5.Location = New System.Drawing.Point(0, 226)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(251, 109)
        Me.Panel5.TabIndex = 141
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Button_stop)
        Me.GroupBox4.Controls.Add(Me.Button_start)
        Me.GroupBox4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox4.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox4.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBox4.Size = New System.Drawing.Size(251, 109)
        Me.GroupBox4.TabIndex = 0
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Lavorazioni ODP"
        '
        'Button_stop
        '
        Me.Button_stop.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button_stop.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Button_stop.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_stop.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_stop.ForeColor = System.Drawing.Color.White
        Me.Button_stop.Image = CType(resources.GetObject("Button_stop.Image"), System.Drawing.Image)
        Me.Button_stop.Location = New System.Drawing.Point(123, 15)
        Me.Button_stop.Name = "Button_stop"
        Me.Button_stop.Size = New System.Drawing.Size(126, 92)
        Me.Button_stop.TabIndex = 142
        Me.Button_stop.Text = "STOP"
        Me.Button_stop.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage
        Me.Button_stop.UseVisualStyleBackColor = True
        '
        'Button_start
        '
        Me.Button_start.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button_start.Dock = System.Windows.Forms.DockStyle.Left
        Me.Button_start.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button_start.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button_start.ForeColor = System.Drawing.Color.White
        Me.Button_start.Image = CType(resources.GetObject("Button_start.Image"), System.Drawing.Image)
        Me.Button_start.Location = New System.Drawing.Point(2, 15)
        Me.Button_start.Name = "Button_start"
        Me.Button_start.Size = New System.Drawing.Size(121, 92)
        Me.Button_start.TabIndex = 140
        Me.Button_start.Text = "START"
        Me.Button_start.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage
        Me.Button_start.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.Controls.Add(Me.GroupBox3)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(251, 226)
        Me.Panel2.TabIndex = 0
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label_gruppo_articolo_F)
        Me.GroupBox3.Controls.Add(Me.Label_descrizione)
        Me.GroupBox3.Controls.Add(Me.Label_gruppo_Articoli)
        Me.GroupBox3.Controls.Add(Me.Label_numero_ODP_F)
        Me.GroupBox3.Controls.Add(Me.Label_disegno_F)
        Me.GroupBox3.Controls.Add(Me.Label_numero_odp)
        Me.GroupBox3.Controls.Add(Me.Label_Disegno)
        Me.GroupBox3.Controls.Add(Me.Label3)
        Me.GroupBox3.Controls.Add(Me.Label_codice_ODP)
        Me.GroupBox3.Controls.Add(Me.Label_commessa)
        Me.GroupBox3.Controls.Add(Me.Label_Codice_ODP_F)
        Me.GroupBox3.Controls.Add(Me.Label_commessa_F)
        Me.GroupBox3.Controls.Add(Me.Label_fase_F)
        Me.GroupBox3.Controls.Add(Me.Label_Fase)
        Me.GroupBox3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox3.ForeColor = System.Drawing.Color.White
        Me.GroupBox3.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(251, 226)
        Me.GroupBox3.TabIndex = 137
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Dettagli ordine"
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.White
        Me.Panel3.Controls.Add(Me.TableLayoutPanel1)
        Me.Panel3.Controls.Add(Me.Panel8)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel3.Location = New System.Drawing.Point(251, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1157, 226)
        Me.Panel3.TabIndex = 146
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox1, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox2, 0, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(1093, 226)
        Me.TableLayoutPanel1.TabIndex = 146
        '
        'Panel8
        '
        Me.Panel8.Controls.Add(Me.Panel9)
        Me.Panel8.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel8.Location = New System.Drawing.Point(1093, 0)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(64, 226)
        Me.Panel8.TabIndex = 145
        '
        'Panel9
        '
        Me.Panel9.Controls.Add(Me.Button1)
        Me.Panel9.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel9.Location = New System.Drawing.Point(0, 0)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(64, 54)
        Me.Panel9.TabIndex = 0
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.SteelBlue
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel4.Location = New System.Drawing.Point(1250, 226)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(158, 631)
        Me.Panel4.TabIndex = 147
        '
        'Panel7
        '
        Me.Panel7.Controls.Add(Me.DataGridView_lavorazioni)
        Me.Panel7.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel7.Location = New System.Drawing.Point(251, 226)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(999, 631)
        Me.Panel7.TabIndex = 148
        '
        'DataGridView_lavorazioni
        '
        Me.DataGridView_lavorazioni.AllowDrop = True
        Me.DataGridView_lavorazioni.AllowUserToAddRows = False
        Me.DataGridView_lavorazioni.AllowUserToDeleteRows = False
        Me.DataGridView_lavorazioni.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView_lavorazioni.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView_lavorazioni.BorderStyle = System.Windows.Forms.BorderStyle.None
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlDark
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView_lavorazioni.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView_lavorazioni.ColumnHeadersHeight = 30
        Me.DataGridView_lavorazioni.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.N_ODP, Me.Documento, Me.ODP, Me.Codice, Me.Descrizione, Me.Disegno, Me.Quantità, Me.Commessa, Me.Nome_macchina, Me.Cliente, Me.Dipendente, Me.Risorsa, Me.Data, Me.Start})
        Me.DataGridView_lavorazioni.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView_lavorazioni.GridColor = System.Drawing.SystemColors.ActiveBorder
        Me.DataGridView_lavorazioni.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.DataGridView_lavorazioni.Location = New System.Drawing.Point(0, 0)
        Me.DataGridView_lavorazioni.Name = "DataGridView_lavorazioni"
        Me.DataGridView_lavorazioni.RowHeadersVisible = False
        Me.DataGridView_lavorazioni.RowHeadersWidth = 51
        Me.DataGridView_lavorazioni.RowTemplate.Height = 35
        Me.DataGridView_lavorazioni.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView_lavorazioni.Size = New System.Drawing.Size(999, 631)
        Me.DataGridView_lavorazioni.TabIndex = 140
        '
        'N_ODP
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.FloralWhite
        Me.N_ODP.DefaultCellStyle = DataGridViewCellStyle2
        Me.N_ODP.HeaderText = "ID"
        Me.N_ODP.MinimumWidth = 6
        Me.N_ODP.Name = "N_ODP"
        Me.N_ODP.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.N_ODP.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.N_ODP.Visible = False
        '
        'Documento
        '
        Me.Documento.HeaderText = "Documento"
        Me.Documento.MinimumWidth = 6
        Me.Documento.Name = "Documento"
        Me.Documento.Visible = False
        '
        'ODP
        '
        Me.ODP.FillWeight = 50.0!
        Me.ODP.HeaderText = "ODP"
        Me.ODP.MinimumWidth = 6
        Me.ODP.Name = "ODP"
        '
        'Codice
        '
        Me.Codice.FillWeight = 50.0!
        Me.Codice.HeaderText = "Codice"
        Me.Codice.MinimumWidth = 6
        Me.Codice.Name = "Codice"
        '
        'Descrizione
        '
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Descrizione.DefaultCellStyle = DataGridViewCellStyle3
        Me.Descrizione.FillWeight = 200.0!
        Me.Descrizione.HeaderText = "Descrizione"
        Me.Descrizione.MinimumWidth = 20
        Me.Descrizione.Name = "Descrizione"
        '
        'Disegno
        '
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Disegno.DefaultCellStyle = DataGridViewCellStyle4
        Me.Disegno.FillWeight = 80.0!
        Me.Disegno.HeaderText = "Disegno"
        Me.Disegno.MinimumWidth = 6
        Me.Disegno.Name = "Disegno"
        Me.Disegno.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Disegno.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        '
        'Quantità
        '
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Quantità.DefaultCellStyle = DataGridViewCellStyle5
        Me.Quantità.FillWeight = 60.0!
        Me.Quantità.HeaderText = "Quantità"
        Me.Quantità.MinimumWidth = 6
        Me.Quantità.Name = "Quantità"
        '
        'Commessa
        '
        Me.Commessa.FillWeight = 60.0!
        Me.Commessa.HeaderText = "Commessa"
        Me.Commessa.MinimumWidth = 6
        Me.Commessa.Name = "Commessa"
        '
        'Nome_macchina
        '
        Me.Nome_macchina.HeaderText = "Nome macchina"
        Me.Nome_macchina.Name = "Nome_macchina"
        '
        'Cliente
        '
        Me.Cliente.HeaderText = "Cliente"
        Me.Cliente.Name = "Cliente"
        '
        'Dipendente
        '
        DataGridViewCellStyle6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dipendente.DefaultCellStyle = DataGridViewCellStyle6
        Me.Dipendente.FillWeight = 150.0!
        Me.Dipendente.HeaderText = "Dipendente"
        Me.Dipendente.MinimumWidth = 6
        Me.Dipendente.Name = "Dipendente"
        Me.Dipendente.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'Risorsa
        '
        Me.Risorsa.FillWeight = 120.0!
        Me.Risorsa.HeaderText = "Risorsa"
        Me.Risorsa.MinimumWidth = 6
        Me.Risorsa.Name = "Risorsa"
        '
        'Data
        '
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle7.Format = "M"
        DataGridViewCellStyle7.NullValue = Nothing
        Me.Data.DefaultCellStyle = DataGridViewCellStyle7
        Me.Data.FillWeight = 70.0!
        Me.Data.HeaderText = "Data"
        Me.Data.MinimumWidth = 6
        Me.Data.Name = "Data"
        '
        'Start
        '
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Start.DefaultCellStyle = DataGridViewCellStyle8
        Me.Start.FillWeight = 60.0!
        Me.Start.HeaderText = "Start"
        Me.Start.MinimumWidth = 6
        Me.Start.Name = "Start"
        '
        'Lavorazioni_MES
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1408, 857)
        Me.ControlBox = False
        Me.Controls.Add(Me.Panel7)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "Lavorazioni_MES"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Lavorazioni"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.Panel6.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.Panel8.ResumeLayout(False)
        Me.Panel9.ResumeLayout(False)
        Me.Panel7.ResumeLayout(False)
        CType(Me.DataGridView_lavorazioni, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ComboBox_dipendente As ComboBox
    Friend WithEvents Label_gruppo_articolo_F As Label
    Friend WithEvents Label_gruppo_Articoli As Label
    Friend WithEvents Label_disegno_F As Label
    Friend WithEvents Label_Disegno As Label
    Friend WithEvents Label_codice_ODP As Label
    Friend WithEvents Label_Codice_ODP_F As Label
    Friend WithEvents Label_fase_F As Label
    Friend WithEvents Label_Fase As Label
    Friend WithEvents Label_commessa_F As Label
    Friend WithEvents Label_commessa As Label
    Friend WithEvents Label_descrizione As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label_numero_odp As Label
    Friend WithEvents Label_numero_ODP_F As Label
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents ComboBox_risorse As ComboBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents Panel5 As Panel
    Friend WithEvents Button_start As Button
    Friend WithEvents Button_stop As Button
    Friend WithEvents Panel4 As Panel
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents Panel6 As Panel
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Panel8 As Panel
    Friend WithEvents Panel9 As Panel
    Friend WithEvents Panel7 As Panel
    Friend WithEvents Button4 As Button
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents DataGridView_lavorazioni As DataGridView
    Friend WithEvents N_ODP As DataGridViewButtonColumn
    Friend WithEvents Documento As DataGridViewTextBoxColumn
    Friend WithEvents ODP As DataGridViewTextBoxColumn
    Friend WithEvents Codice As DataGridViewTextBoxColumn
    Friend WithEvents Descrizione As DataGridViewTextBoxColumn
    Friend WithEvents Disegno As DataGridViewButtonColumn
    Friend WithEvents Quantità As DataGridViewTextBoxColumn
    Friend WithEvents Commessa As DataGridViewTextBoxColumn
    Friend WithEvents Nome_macchina As DataGridViewTextBoxColumn
    Friend WithEvents Cliente As DataGridViewTextBoxColumn
    Friend WithEvents Dipendente As DataGridViewTextBoxColumn
    Friend WithEvents Risorsa As DataGridViewTextBoxColumn
    Friend WithEvents Data As DataGridViewTextBoxColumn
    Friend WithEvents Start As DataGridViewTextBoxColumn
    Friend WithEvents Button5 As Button
End Class
