<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Dashboard_pianificazione
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Dashboard_pianificazione))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Label_commessa = New System.Windows.Forms.Label()
        Me.Button_aggiungi_riga = New System.Windows.Forms.Button()
        Me.Label_descrizione = New System.Windows.Forms.Label()
        Me.Label_consegna = New System.Windows.Forms.Label()
        Me.Label_cliente = New System.Windows.Forms.Label()
        Me.Label_cliente_finale = New System.Windows.Forms.Label()
        Me.MonthCalendar_data_fine = New System.Windows.Forms.MonthCalendar()
        Me.TextBox_data_fine = New System.Windows.Forms.TextBox()
        Me.TextBox_data_inizio = New System.Windows.Forms.TextBox()
        Me.ComboBox_dipendente = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ComboBox_risorse = New System.Windows.Forms.ComboBox()
        Me.MonthCalendar_data_inizio = New System.Windows.Forms.MonthCalendar()
        Me.Button_aggiorna_excel = New System.Windows.Forms.Button()
        Me.TextBox_giorni_lav = New System.Windows.Forms.TextBox()
        Me.Label_attivita_F = New System.Windows.Forms.Label()
        Me.TextBox_attivita = New System.Windows.Forms.TextBox()
        Me.CheckBox_dipendenti = New System.Windows.Forms.CheckBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.ComboBox_unità = New System.Windows.Forms.ComboBox()
        Me.DataGridView_Risorse = New System.Windows.Forms.DataGridView()
        Me.Id = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Risorsa = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Risorsa_dec = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.dip = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Data_I = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Data_F = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Unità = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Attivita_Tab = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Inizio = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Fine = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Inizio_fine = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Modifica = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Delete = New System.Windows.Forms.DataGridViewButtonColumn()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel6 = New System.Windows.Forms.TableLayoutPanel()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.FlowLayoutPanel2 = New System.Windows.Forms.FlowLayoutPanel()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.GroupBox11 = New System.Windows.Forms.GroupBox()
        Me.GroupBox12 = New System.Windows.Forms.GroupBox()
        Me.GroupBox13 = New System.Windows.Forms.GroupBox()
        Me.GroupBox14 = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.GroupBox15 = New System.Windows.Forms.GroupBox()
        Me.GroupBox16 = New System.Windows.Forms.GroupBox()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.GroupBox10 = New System.Windows.Forms.GroupBox()
        Me.GroupBox17 = New System.Windows.Forms.GroupBox()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.TableLayoutPanel5 = New System.Windows.Forms.TableLayoutPanel()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.ComboBox_stato = New System.Windows.Forms.ComboBox()
        Me.GroupBox9 = New System.Windows.Forms.GroupBox()
        Me.Panel1.SuspendLayout()
        CType(Me.DataGridView_Risorse, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TableLayoutPanel6.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.FlowLayoutPanel2.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox11.SuspendLayout()
        Me.GroupBox12.SuspendLayout()
        Me.GroupBox13.SuspendLayout()
        Me.GroupBox14.SuspendLayout()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.GroupBox15.SuspendLayout()
        Me.GroupBox16.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox10.SuspendLayout()
        Me.GroupBox17.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.TableLayoutPanel5.SuspendLayout()
        Me.GroupBox9.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label_commessa
        '
        resources.ApplyResources(Me.Label_commessa, "Label_commessa")
        Me.Label_commessa.Name = "Label_commessa"
        '
        'Button_aggiungi_riga
        '
        Me.Button_aggiungi_riga.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        resources.ApplyResources(Me.Button_aggiungi_riga, "Button_aggiungi_riga")
        Me.Button_aggiungi_riga.Name = "Button_aggiungi_riga"
        Me.Button_aggiungi_riga.UseVisualStyleBackColor = False
        '
        'Label_descrizione
        '
        resources.ApplyResources(Me.Label_descrizione, "Label_descrizione")
        Me.Label_descrizione.Name = "Label_descrizione"
        '
        'Label_consegna
        '
        resources.ApplyResources(Me.Label_consegna, "Label_consegna")
        Me.Label_consegna.Name = "Label_consegna"
        '
        'Label_cliente
        '
        resources.ApplyResources(Me.Label_cliente, "Label_cliente")
        Me.Label_cliente.Name = "Label_cliente"
        '
        'Label_cliente_finale
        '
        resources.ApplyResources(Me.Label_cliente_finale, "Label_cliente_finale")
        Me.Label_cliente_finale.Name = "Label_cliente_finale"
        '
        'MonthCalendar_data_fine
        '
        resources.ApplyResources(Me.MonthCalendar_data_fine, "MonthCalendar_data_fine")
        Me.MonthCalendar_data_fine.MaxSelectionCount = 999
        Me.MonthCalendar_data_fine.Name = "MonthCalendar_data_fine"
        '
        'TextBox_data_fine
        '
        resources.ApplyResources(Me.TextBox_data_fine, "TextBox_data_fine")
        Me.TextBox_data_fine.Name = "TextBox_data_fine"
        '
        'TextBox_data_inizio
        '
        resources.ApplyResources(Me.TextBox_data_inizio, "TextBox_data_inizio")
        Me.TextBox_data_inizio.Name = "TextBox_data_inizio"
        '
        'ComboBox_dipendente
        '
        Me.ComboBox_dipendente.FormattingEnabled = True
        resources.ApplyResources(Me.ComboBox_dipendente, "ComboBox_dipendente")
        Me.ComboBox_dipendente.Name = "ComboBox_dipendente"
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'ComboBox_risorse
        '
        resources.ApplyResources(Me.ComboBox_risorse, "ComboBox_risorse")
        Me.ComboBox_risorse.FormattingEnabled = True
        Me.ComboBox_risorse.Name = "ComboBox_risorse"
        '
        'MonthCalendar_data_inizio
        '
        resources.ApplyResources(Me.MonthCalendar_data_inizio, "MonthCalendar_data_inizio")
        Me.MonthCalendar_data_inizio.MaxSelectionCount = 999
        Me.MonthCalendar_data_inizio.Name = "MonthCalendar_data_inizio"
        '
        'Button_aggiorna_excel
        '
        Me.Button_aggiorna_excel.BackColor = System.Drawing.Color.Green
        Me.Button_aggiorna_excel.Cursor = System.Windows.Forms.Cursors.Arrow
        resources.ApplyResources(Me.Button_aggiorna_excel, "Button_aggiorna_excel")
        Me.Button_aggiorna_excel.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Button_aggiorna_excel.Name = "Button_aggiorna_excel"
        Me.Button_aggiorna_excel.UseVisualStyleBackColor = False
        '
        'TextBox_giorni_lav
        '
        resources.ApplyResources(Me.TextBox_giorni_lav, "TextBox_giorni_lav")
        Me.TextBox_giorni_lav.Name = "TextBox_giorni_lav"
        '
        'Label_attivita_F
        '
        resources.ApplyResources(Me.Label_attivita_F, "Label_attivita_F")
        Me.Label_attivita_F.Name = "Label_attivita_F"
        '
        'TextBox_attivita
        '
        resources.ApplyResources(Me.TextBox_attivita, "TextBox_attivita")
        Me.TextBox_attivita.Name = "TextBox_attivita"
        '
        'CheckBox_dipendenti
        '
        resources.ApplyResources(Me.CheckBox_dipendenti, "CheckBox_dipendenti")
        Me.CheckBox_dipendenti.Name = "CheckBox_dipendenti"
        Me.CheckBox_dipendenti.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.CheckBox_dipendenti)
        Me.Panel1.Controls.Add(Me.TextBox_attivita)
        Me.Panel1.Controls.Add(Me.Label_attivita_F)
        Me.Panel1.Controls.Add(Me.ComboBox_dipendente)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Button_aggiungi_riga)
        resources.ApplyResources(Me.Panel1, "Panel1")
        Me.Panel1.Name = "Panel1"
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        resources.ApplyResources(Me.Button2, "Button2")
        Me.Button2.Name = "Button2"
        Me.Button2.UseVisualStyleBackColor = False
        '
        'ComboBox_unità
        '
        resources.ApplyResources(Me.ComboBox_unità, "ComboBox_unità")
        Me.ComboBox_unità.FormattingEnabled = True
        Me.ComboBox_unità.Items.AddRange(New Object() {resources.GetString("ComboBox_unità.Items"), resources.GetString("ComboBox_unità.Items1"), resources.GetString("ComboBox_unità.Items2"), resources.GetString("ComboBox_unità.Items3"), resources.GetString("ComboBox_unità.Items4"), resources.GetString("ComboBox_unità.Items5"), resources.GetString("ComboBox_unità.Items6"), resources.GetString("ComboBox_unità.Items7"), resources.GetString("ComboBox_unità.Items8"), resources.GetString("ComboBox_unità.Items9")})
        Me.ComboBox_unità.Name = "ComboBox_unità"
        '
        'DataGridView_Risorse
        '
        Me.DataGridView_Risorse.AllowDrop = True
        Me.DataGridView_Risorse.AllowUserToAddRows = False
        Me.DataGridView_Risorse.AllowUserToDeleteRows = False
        Me.DataGridView_Risorse.AllowUserToOrderColumns = True
        Me.DataGridView_Risorse.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridView_Risorse.BackgroundColor = System.Drawing.SystemColors.ActiveCaption
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.ControlDark
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView_Risorse.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        resources.ApplyResources(Me.DataGridView_Risorse, "DataGridView_Risorse")
        Me.DataGridView_Risorse.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Id, Me.Risorsa, Me.Risorsa_dec, Me.dip, Me.Data_I, Me.Data_F, Me.Unità, Me.Attivita_Tab, Me.Inizio, Me.Fine, Me.Inizio_fine, Me.Modifica, Me.Delete})
        Me.DataGridView_Risorse.Cursor = System.Windows.Forms.Cursors.Default
        Me.DataGridView_Risorse.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridView_Risorse.GridColor = System.Drawing.SystemColors.ActiveBorder
        Me.DataGridView_Risorse.Name = "DataGridView_Risorse"
        Me.DataGridView_Risorse.RowHeadersVisible = False
        Me.DataGridView_Risorse.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        '
        'Id
        '
        Me.Id.FillWeight = 118.7561!
        resources.ApplyResources(Me.Id, "Id")
        Me.Id.Name = "Id"
        '
        'Risorsa
        '
        Me.Risorsa.FillWeight = 119.4907!
        resources.ApplyResources(Me.Risorsa, "Risorsa")
        Me.Risorsa.Name = "Risorsa"
        Me.Risorsa.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'Risorsa_dec
        '
        Me.Risorsa_dec.FillWeight = 120.3767!
        resources.ApplyResources(Me.Risorsa_dec, "Risorsa_dec")
        Me.Risorsa_dec.Name = "Risorsa_dec"
        '
        'dip
        '
        resources.ApplyResources(Me.dip, "dip")
        Me.dip.Name = "dip"
        '
        'Data_I
        '
        Me.Data_I.FillWeight = 115.1934!
        resources.ApplyResources(Me.Data_I, "Data_I")
        Me.Data_I.Name = "Data_I"
        '
        'Data_F
        '
        Me.Data_F.FillWeight = 79.57105!
        resources.ApplyResources(Me.Data_F, "Data_F")
        Me.Data_F.Name = "Data_F"
        '
        'Unità
        '
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.Unità.DefaultCellStyle = DataGridViewCellStyle2
        Me.Unità.FillWeight = 75.54593!
        resources.ApplyResources(Me.Unità, "Unità")
        Me.Unità.Items.AddRange(New Object() {"0", "1", "2", "3", "4", "5", "6"})
        Me.Unità.Name = "Unità"
        Me.Unità.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'Attivita_Tab
        '
        resources.ApplyResources(Me.Attivita_Tab, "Attivita_Tab")
        Me.Attivita_Tab.Name = "Attivita_Tab"
        '
        'Inizio
        '
        resources.ApplyResources(Me.Inizio, "Inizio")
        Me.Inizio.Name = "Inizio"
        '
        'Fine
        '
        resources.ApplyResources(Me.Fine, "Fine")
        Me.Fine.Name = "Fine"
        '
        'Inizio_fine
        '
        resources.ApplyResources(Me.Inizio_fine, "Inizio_fine")
        Me.Inizio_fine.Name = "Inizio_fine"
        Me.Inizio_fine.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Inizio_fine.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable
        '
        'Modifica
        '
        resources.ApplyResources(Me.Modifica, "Modifica")
        Me.Modifica.Name = "Modifica"
        Me.Modifica.Text = "✎"
        Me.Modifica.UseColumnTextForButtonValue = True
        '
        'Delete
        '
        resources.ApplyResources(Me.Delete, "Delete")
        Me.Delete.Name = "Delete"
        Me.Delete.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        Me.Delete.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic
        Me.Delete.Text = "Ø"
        Me.Delete.UseColumnTextForButtonValue = True
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Blue
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Arrow
        resources.ApplyResources(Me.Button1, "Button1")
        Me.Button1.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Button1.Name = "Button1"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'TableLayoutPanel1
        '
        resources.ApplyResources(Me.TableLayoutPanel1, "TableLayoutPanel1")
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel5, 0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        '
        'TableLayoutPanel2
        '
        resources.ApplyResources(Me.TableLayoutPanel2, "TableLayoutPanel2")
        Me.TableLayoutPanel2.Controls.Add(Me.DataGridView_Risorse, 1, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.TableLayoutPanel6, 0, 0)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        '
        'TableLayoutPanel6
        '
        resources.ApplyResources(Me.TableLayoutPanel6, "TableLayoutPanel6")
        Me.TableLayoutPanel6.Controls.Add(Me.GroupBox7, 0, 1)
        Me.TableLayoutPanel6.Controls.Add(Me.FlowLayoutPanel1, 0, 0)
        Me.TableLayoutPanel6.Name = "TableLayoutPanel6"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.TableLayoutPanel3)
        resources.ApplyResources(Me.GroupBox7, "GroupBox7")
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.TabStop = False
        '
        'TableLayoutPanel3
        '
        resources.ApplyResources(Me.TableLayoutPanel3, "TableLayoutPanel3")
        Me.TableLayoutPanel3.Controls.Add(Me.FlowLayoutPanel2, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.Panel1, 0, 2)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel4, 0, 1)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        '
        'FlowLayoutPanel2
        '
        Me.FlowLayoutPanel2.Controls.Add(Me.GroupBox8)
        Me.FlowLayoutPanel2.Controls.Add(Me.GroupBox11)
        Me.FlowLayoutPanel2.Controls.Add(Me.GroupBox12)
        Me.FlowLayoutPanel2.Controls.Add(Me.GroupBox13)
        Me.FlowLayoutPanel2.Controls.Add(Me.GroupBox14)
        resources.ApplyResources(Me.FlowLayoutPanel2, "FlowLayoutPanel2")
        Me.FlowLayoutPanel2.Name = "FlowLayoutPanel2"
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.ComboBox_risorse)
        resources.ApplyResources(Me.GroupBox8, "GroupBox8")
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.TabStop = False
        '
        'GroupBox11
        '
        Me.GroupBox11.Controls.Add(Me.TextBox_data_inizio)
        resources.ApplyResources(Me.GroupBox11, "GroupBox11")
        Me.GroupBox11.Name = "GroupBox11"
        Me.GroupBox11.TabStop = False
        '
        'GroupBox12
        '
        Me.GroupBox12.Controls.Add(Me.ComboBox_unità)
        resources.ApplyResources(Me.GroupBox12, "GroupBox12")
        Me.GroupBox12.Name = "GroupBox12"
        Me.GroupBox12.TabStop = False
        '
        'GroupBox13
        '
        Me.GroupBox13.Controls.Add(Me.TextBox_giorni_lav)
        resources.ApplyResources(Me.GroupBox13, "GroupBox13")
        Me.GroupBox13.Name = "GroupBox13"
        Me.GroupBox13.TabStop = False
        '
        'GroupBox14
        '
        Me.GroupBox14.Controls.Add(Me.TextBox_data_fine)
        resources.ApplyResources(Me.GroupBox14, "GroupBox14")
        Me.GroupBox14.Name = "GroupBox14"
        Me.GroupBox14.TabStop = False
        '
        'TableLayoutPanel4
        '
        resources.ApplyResources(Me.TableLayoutPanel4, "TableLayoutPanel4")
        Me.TableLayoutPanel4.Controls.Add(Me.GroupBox15, 0, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.GroupBox16, 1, 0)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        '
        'GroupBox15
        '
        Me.GroupBox15.Controls.Add(Me.MonthCalendar_data_inizio)
        resources.ApplyResources(Me.GroupBox15, "GroupBox15")
        Me.GroupBox15.Name = "GroupBox15"
        Me.GroupBox15.TabStop = False
        '
        'GroupBox16
        '
        Me.GroupBox16.Controls.Add(Me.MonthCalendar_data_fine)
        resources.ApplyResources(Me.GroupBox16, "GroupBox16")
        Me.GroupBox16.Name = "GroupBox16"
        Me.GroupBox16.TabStop = False
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.Controls.Add(Me.GroupBox1)
        Me.FlowLayoutPanel1.Controls.Add(Me.GroupBox9)
        Me.FlowLayoutPanel1.Controls.Add(Me.GroupBox2)
        Me.FlowLayoutPanel1.Controls.Add(Me.GroupBox4)
        Me.FlowLayoutPanel1.Controls.Add(Me.GroupBox5)
        Me.FlowLayoutPanel1.Controls.Add(Me.GroupBox10)
        Me.FlowLayoutPanel1.Controls.Add(Me.GroupBox17)
        Me.FlowLayoutPanel1.Controls.Add(Me.GroupBox6)
        resources.ApplyResources(Me.FlowLayoutPanel1, "FlowLayoutPanel1")
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label_commessa)
        resources.ApplyResources(Me.GroupBox1, "GroupBox1")
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.TabStop = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.Label_descrizione)
        resources.ApplyResources(Me.GroupBox2, "GroupBox2")
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.TabStop = False
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.Label_cliente)
        resources.ApplyResources(Me.GroupBox4, "GroupBox4")
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.TabStop = False
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.Label_cliente_finale)
        resources.ApplyResources(Me.GroupBox5, "GroupBox5")
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.TabStop = False
        '
        'GroupBox10
        '
        Me.GroupBox10.Controls.Add(Me.Button1)
        resources.ApplyResources(Me.GroupBox10, "GroupBox10")
        Me.GroupBox10.Name = "GroupBox10"
        Me.GroupBox10.TabStop = False
        '
        'GroupBox17
        '
        Me.GroupBox17.Controls.Add(Me.Button3)
        resources.ApplyResources(Me.GroupBox17, "GroupBox17")
        Me.GroupBox17.Name = "GroupBox17"
        Me.GroupBox17.TabStop = False
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.SystemColors.Info
        resources.ApplyResources(Me.Button3, "Button3")
        Me.Button3.Name = "Button3"
        Me.Button3.UseVisualStyleBackColor = False
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.Label_consegna)
        resources.ApplyResources(Me.GroupBox6, "GroupBox6")
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.TabStop = False
        '
        'TableLayoutPanel5
        '
        resources.ApplyResources(Me.TableLayoutPanel5, "TableLayoutPanel5")
        Me.TableLayoutPanel5.Controls.Add(Me.Button4, 9, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.Button_aggiorna_excel, 0, 0)
        Me.TableLayoutPanel5.Name = "TableLayoutPanel5"
        '
        'Button4
        '
        resources.ApplyResources(Me.Button4, "Button4")
        Me.Button4.Name = "Button4"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'ComboBox_stato
        '
        resources.ApplyResources(Me.ComboBox_stato, "ComboBox_stato")
        Me.ComboBox_stato.FormattingEnabled = True
        Me.ComboBox_stato.Items.AddRange(New Object() {resources.GetString("ComboBox_stato.Items"), resources.GetString("ComboBox_stato.Items1"), resources.GetString("ComboBox_stato.Items2")})
        Me.ComboBox_stato.Name = "ComboBox_stato"
        '
        'GroupBox9
        '
        Me.GroupBox9.Controls.Add(Me.ComboBox_stato)
        resources.ApplyResources(Me.GroupBox9, "GroupBox9")
        Me.GroupBox9.Name = "GroupBox9"
        Me.GroupBox9.TabStop = False
        '
        'Dashboard_pianificazione
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ControlBox = False
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.MinimizeBox = False
        Me.Name = "Dashboard_pianificazione"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.DataGridView_Risorse, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel6.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.FlowLayoutPanel2.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox11.ResumeLayout(False)
        Me.GroupBox11.PerformLayout()
        Me.GroupBox12.ResumeLayout(False)
        Me.GroupBox13.ResumeLayout(False)
        Me.GroupBox13.PerformLayout()
        Me.GroupBox14.ResumeLayout(False)
        Me.GroupBox14.PerformLayout()
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.GroupBox15.ResumeLayout(False)
        Me.GroupBox16.ResumeLayout(False)
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox10.ResumeLayout(False)
        Me.GroupBox17.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.TableLayoutPanel5.ResumeLayout(False)
        Me.GroupBox9.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Label_commessa As Label
    Friend WithEvents Button_aggiungi_riga As Button
    Friend WithEvents Label_descrizione As Label
    Friend WithEvents Label_consegna As Label
    Friend WithEvents Label_cliente As Label
    Friend WithEvents Label_cliente_finale As Label
    Friend WithEvents MonthCalendar_data_fine As MonthCalendar
    Friend WithEvents TextBox_data_fine As TextBox
    Friend WithEvents TextBox_data_inizio As TextBox
    Friend WithEvents ComboBox_dipendente As ComboBox
    Friend WithEvents Label2 As Label
    Friend WithEvents ComboBox_risorse As ComboBox
    Friend WithEvents MonthCalendar_data_inizio As MonthCalendar
    Friend WithEvents Button_aggiorna_excel As Button
    Friend WithEvents TextBox_giorni_lav As TextBox
    Friend WithEvents Label_attivita_F As Label
    Friend WithEvents TextBox_attivita As TextBox
    Friend WithEvents CheckBox_dipendenti As CheckBox
    Friend WithEvents Panel1 As Panel
    Friend WithEvents DataGridView_Risorse As DataGridView
    Friend WithEvents ComboBox_unità As ComboBox
    Friend WithEvents Button1 As Button
    Friend WithEvents TableLayoutPanel1 As TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As TableLayoutPanel
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents GroupBox5 As GroupBox
    Friend WithEvents GroupBox6 As GroupBox
    Friend WithEvents GroupBox7 As GroupBox
    Friend WithEvents TableLayoutPanel3 As TableLayoutPanel
    Friend WithEvents GroupBox10 As GroupBox
    Friend WithEvents FlowLayoutPanel2 As FlowLayoutPanel
    Friend WithEvents GroupBox8 As GroupBox
    Friend WithEvents GroupBox11 As GroupBox
    Friend WithEvents GroupBox12 As GroupBox
    Friend WithEvents GroupBox13 As GroupBox
    Friend WithEvents GroupBox14 As GroupBox
    Friend WithEvents TableLayoutPanel4 As TableLayoutPanel
    Friend WithEvents GroupBox15 As GroupBox
    Friend WithEvents GroupBox16 As GroupBox
    Friend WithEvents TableLayoutPanel5 As TableLayoutPanel
    Friend WithEvents Button4 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Id As DataGridViewTextBoxColumn
    Friend WithEvents Risorsa As DataGridViewTextBoxColumn
    Friend WithEvents Risorsa_dec As DataGridViewTextBoxColumn
    Friend WithEvents dip As DataGridViewTextBoxColumn
    Friend WithEvents Data_I As DataGridViewTextBoxColumn
    Friend WithEvents Data_F As DataGridViewTextBoxColumn
    Friend WithEvents Unità As DataGridViewComboBoxColumn
    Friend WithEvents Attivita_Tab As DataGridViewTextBoxColumn
    Friend WithEvents Inizio As DataGridViewTextBoxColumn
    Friend WithEvents Fine As DataGridViewTextBoxColumn
    Friend WithEvents Inizio_fine As DataGridViewTextBoxColumn
    Friend WithEvents Modifica As DataGridViewButtonColumn
    Friend WithEvents Delete As DataGridViewButtonColumn
    Friend WithEvents GroupBox17 As GroupBox
    Friend WithEvents Button3 As Button
    Friend WithEvents TableLayoutPanel6 As TableLayoutPanel
    Friend WithEvents GroupBox9 As GroupBox
    Friend WithEvents ComboBox_stato As ComboBox
End Class
