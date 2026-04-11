Imports System.Data.SqlClient
Imports Tirelli.ODP_Form
Imports Npgsql

Public Class Modulo_baia

#Region "Proprietà e stato"

    Public numero_baia As Integer

    Private isDragging As Boolean = False
    Private startPoint As Point

    Public stato_montaggio As Boolean = False
    Public stato_elettrico As Boolean = False
    Public stato_software As Boolean = False
    Public stato_collaudo As Boolean = False

    Public contatore_mec As Integer = 0
    Public contatore_el As Integer = 0
    Public contatore_soft As Integer = 0
    Public contatore_col As Integer = 0
    Public contatore_altro As Integer = 0

    Public Property Titolo As String
        Get
            Return Label1.Text
        End Get
        Set(value As String)
            Label1.Text = value
        End Set
    End Property

    ' Costanti WBS
    Private Const WBS_MONTAGGIO As String = "4.3"
    Private Const WBS_ELETTRICO As String = "4.4"
    Private Const WBS_SOFTWARE As String = "4.11"
    Private Const WBS_COLLAUDO As String = "5"
    Private Const WBS_FAT As String = "999"
    Private Const WBS_SPEDIZIONE As String = "7"

    ' Shortcut al form principale
    Private ReadOnly Property FLayout As Form_layout_CAP_1
        Get
            Return Form_layout_CAP_1
        End Get
    End Property

#End Region

#Region "Load e inizializzazione"

    Private Sub modulo_baia_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        AggiungiHandlerAiControlli(Me)
        AbilitaDragSuFigli(Me)
        BloccaDropInTuttiICtrlDelModulo(Me)
        FlowLayoutPanel1.AllowDrop = False
        GroupBox101.AllowDrop = False
    End Sub

    Public Sub inizializza_modulo(par_commessa As String, pnl As Panel)
        Label4.Text = FLayout.giorni_lavorativi_tra(
            FLayout.ingresso_commessa(par_commessa, "Officina"), Date.Today).ToString() & " GG"
        aggiorna_collaudo(par_commessa)
        trova_tempi_montaggio(par_commessa)
        inserisci_dipendenti(par_commessa)
        alleggerisci_modulo()
        aggiusta_dimensioni()
    End Sub

#End Region

#Region "Drag & Drop"

    Private Sub BloccaDropInTuttiICtrlDelModulo(ctrl As Control)
        ctrl.AllowDrop = True
        AddHandler ctrl.DragEnter, AddressOf BloccaDrop
        AddHandler ctrl.DragOver, AddressOf BloccaDrop
        AddHandler ctrl.DragDrop, AddressOf BloccaDrop
        For Each child As Control In ctrl.Controls
            BloccaDropInTuttiICtrlDelModulo(child)
        Next
    End Sub

    Private Sub BloccaDrop(sender As Object, e As DragEventArgs)
        e.Effect = DragDropEffects.None
    End Sub

    Private Sub AbilitaDragSuFigli(ctrl As Control)
        AddHandler ctrl.MouseDown, AddressOf Controllo_MouseDown
        AddHandler ctrl.MouseMove, AddressOf Controllo_MouseMove
        AddHandler ctrl.MouseUp, AddressOf Controllo_MouseUp
        For Each child As Control In ctrl.Controls
            AbilitaDragSuFigli(child)
        Next
    End Sub

    Private Sub Controllo_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            isDragging = True
            startPoint = e.Location
        End If
    End Sub

    Private Sub Controllo_MouseMove(sender As Object, e As MouseEventArgs)
        If Not isDragging Then Return
        Dim soglia = SystemInformation.DragSize
        If Math.Abs(e.X - startPoint.X) > soglia.Width \ 2 OrElse
           Math.Abs(e.Y - startPoint.Y) > soglia.Height \ 2 Then
            isDragging = False
            FLayout.tipo_spostamento = "SPOST"
            FLayout.parametro_trascinato = "COMMESSA"
            DoDragDrop(Label1.Text, DragDropEffects.Copy)
        End If
    End Sub

    Private Sub Controllo_MouseUp(sender As Object, e As MouseEventArgs)
        isDragging = False
    End Sub

    Private Sub Label1_MouseDown(sender As Object, e As MouseEventArgs) Handles Label1.MouseDown
        If e.Button = MouseButtons.Left Then
            FLayout.tipo_spostamento = "SPOST"
            Label1.DoDragDrop(Label1.Text, DragDropEffects.Copy)
        End If
    End Sub

    Private Sub modulo_baia_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        If e.Button = MouseButtons.Left Then
            FLayout.tipo_spostamento = "SPOST"
            DoDragDrop(Label1.Text, DragDropEffects.Copy)
        End If
    End Sub

    Private Sub modulo_baia_DragEnter(sender As Object, e As DragEventArgs)
        e.Effect = If(FLayout.parametro_trascinato = "RISORSA", DragDropEffects.Copy, DragDropEffects.None)
    End Sub

#End Region

#Region "UI – Dimensioni e visibilità"

    Public Sub aggiusta_dimensioni()
        Me.Margin = New Padding(0)
        Me.AutoSize = True
        Me.AutoSizeMode = AutoSizeMode.GrowAndShrink
        Me.BorderStyle = BorderStyle.FixedSingle

        Dim stato = FLayout.check_baia_layout_A_numero_baia(Label1.Text, numero_baia).Stato
        Select Case stato
            Case "O"
                Me.BackColor = Color.LightBlue
            Case "P"
                Me.BackColor = Color.LemonChiffon
        End Select

        Me.PerformLayout()
    End Sub

    Sub alleggerisci_modulo()
        Dim timerAttivi = Timer_montaggio.Enabled OrElse Timer_EL.Enabled OrElse
                          Timer_COLL.Enabled OrElse Timer_SOFT.Enabled
        Dim panelVisibili = Panel5.Visible OrElse Panel6.Visible OrElse
                            Panel7.Visible OrElse Panel8.Visible

        TableLayoutPanel1.Visible = timerAttivi OrElse panelVisibili
        TableLayoutPanel3.Visible = panelVisibili
    End Sub

#End Region

#Region "Collaudo e montaggio"

    Sub aggiorna_collaudo(par_commessa As String)
        Dim info = Scheda_tecnica.Ottieni_numero_combinazioni(par_commessa)
        If info.Numero_combinazioni > 0 Then
            ProgressBar1.Value = info.Numero_collaudati * 100 \ info.Numero_combinazioni
            GroupBox101.Text = "Coll - " & info.Numero_collaudati & " / " & info.Numero_combinazioni
            GroupBox101.Visible = True
        Else
            ProgressBar1.Value = 0
            GroupBox101.Visible = False
        End If
    End Sub

    Public Sub trova_tempi_montaggio(par_commessa As String)
        Label13.Text = ""
        Dim dataScadenza As Date = Date.MinValue

        Using conn As New NpgsqlConnection(Homepage.JPM_TIRELLI)
            conn.Open()
            Using CMD As New NpgsqlCommand(
                "SELECT W.WBSLVLCOD AS ATT_WBS,
                        substring(W.WBSLVLCOD,1,3) AS Prime_3_WBS,
                        T.TSKRSDTSSSTR AS ATT_DTPIAINI,
                        T.TSKRSDTSSEND AS ATT_DTPIAFIN
                 FROM PRJTSK T
                 LEFT JOIN PRJTSKDET TD ON TD.TSKUID = T.UID
                 LEFT JOIN ANGWBSLVL W ON T.WBSLVLUID = W.UID
                 LEFT JOIN PRJ P ON T.PRJUID = P.UID
                 WHERE T.LOGDEL = 0 AND P.PRJcod = @Commessa AND T.TSKKND = '1'
                 AND (substring(W.WBSLVLCOD,1,3) IN (@wbs_m, @wbs_e, @wbs_s)
                      OR W.WBSLVLCOD IN (@wbs_c, @wbs_f, @wbs_sp))", conn)

                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                CMD.Parameters.AddWithValue("@wbs_m", WBS_MONTAGGIO)
                CMD.Parameters.AddWithValue("@wbs_e", WBS_ELETTRICO)
                CMD.Parameters.AddWithValue("@wbs_s", WBS_SOFTWARE)
                CMD.Parameters.AddWithValue("@wbs_c", WBS_COLLAUDO)
                CMD.Parameters.AddWithValue("@wbs_f", WBS_FAT)
                CMD.Parameters.AddWithValue("@wbs_sp", WBS_SPEDIZIONE)

                Using reader = CMD.ExecuteReader()
                    Dim haRisultati As Boolean = False

                    Do While reader.Read()
                        haRisultati = True
                        Dim wbs As String = reader("ATT_WBS").ToString()
                        Dim wbs3 As String = reader("Prime_3_WBS").ToString()
                        Dim dtIni As Date = CDate(reader("ATT_DTPIAINI")).Date
                        Dim dtFin As Date = CDate(reader("ATT_DTPIAFIN")).Date

                        Select Case True
                            Case wbs3 = WBS_MONTAGGIO
                                GestisciFaseTimer(dtIni, dtFin, Timer_montaggio,
                                                  AddressOf montaggio_completato, AddressOf DISATTIVA_MONTAGGIO)
                            Case wbs3 = WBS_ELETTRICO
                                GestisciFaseTimer(dtIni, dtFin, Timer_EL,
                                                  AddressOf EL_completato, AddressOf DISATTIVA_EL)
                            Case wbs3 = WBS_SOFTWARE
                                GestisciFaseTimer(dtIni, dtFin, Timer_SOFT,
                                                  AddressOf SOFT_completato, AddressOf DISATTIVA_SOFTW)
                            Case wbs = WBS_COLLAUDO
                                GestisciFaseTimer(dtIni, dtFin, Timer_COLL,
                                                  AddressOf coll_completato, AddressOf DISATTIVA_COLL)
                            Case wbs = WBS_FAT OrElse wbs = WBS_SPEDIZIONE
                                Label13.Text = If(wbs = WBS_FAT, "FAT ", "CON ") & dtFin.ToString("dd/MM")
                                dataScadenza = dtFin
                        End Select
                    Loop

                    If Not haRisultati Then
                        Disattivatutti()
                    End If
                End Using
            End Using
        End Using

        AggiornaColuoreScadenza(dataScadenza)
    End Sub

    ''' <summary>Avvia/completa/disattiva una fase in base alle date pianificate.</summary>
    Private Sub GestisciFaseTimer(dtIni As Date, dtFin As Date, timer As Timer,
                                   completato As Action, disattiva As Action)
        If dtIni <= Today AndAlso Today <= dtFin Then
            timer.Start()
        ElseIf dtFin < Today Then
            completato()
        Else
            timer.Stop()
            disattiva()
        End If
    End Sub

    Private Sub Disattivatutti()
        Timer_montaggio.Stop() : DISATTIVA_MONTAGGIO()
        Timer_EL.Stop() : DISATTIVA_EL()
        Timer_COLL.Stop() : DISATTIVA_COLL()
        Timer_SOFT.Stop() : DISATTIVA_SOFTW()
    End Sub

    Private Sub AggiornaColuoreScadenza(dataScadenza As Date)
        If dataScadenza = Date.MinValue Then Return

        Dim oggi = Date.Today
        Dim lunedi = oggi.AddDays(1 - If(oggi.DayOfWeek = DayOfWeek.Sunday, 7, CInt(oggi.DayOfWeek)))

        Label13.ForeColor =
            If(dataScadenza < lunedi.AddDays(14), Color.Red,
            If(dataScadenza < lunedi.AddDays(28), Color.Orange,
            If(dataScadenza < lunedi.AddDays(42), Color.Yellow,
            Color.Green)))
    End Sub

#End Region

#Region "Fasi – Attiva / Disattiva / Completato"

    Sub ATTIVA_MONTAGGIO()
        Panel1.Visible = True
        stato_montaggio = True
    End Sub

    Sub DISATTIVA_MONTAGGIO()
        Panel1.Visible = False
        stato_montaggio = False
    End Sub

    Sub montaggio_completato()
        Panel1.Visible = True
        Panel1.BackColor = Color.Lime
        Label6.ForeColor = Color.Black
    End Sub

    Sub ATTIVA_EL()
        Panel2.Visible = True
        stato_elettrico = True
    End Sub

    Sub DISATTIVA_EL()
        Panel2.Visible = False
        stato_elettrico = False
    End Sub

    Sub EL_completato()
        Panel2.Visible = True
        Panel2.BackColor = Color.Lime
        Label7.ForeColor = Color.Black
    End Sub

    Sub ATTIVA_SOFTW()
        Panel3.Visible = True
        stato_software = True
    End Sub

    Sub DISATTIVA_SOFTW()
        Panel3.Visible = False
        stato_software = False
    End Sub

    Sub SOFT_completato()
        Panel3.Visible = True
        Panel3.BackColor = Color.Lime
        Label8.ForeColor = Color.Black
    End Sub

    Sub ATTIVA_COLL()
        Panel4.Visible = True
        stato_collaudo = True
    End Sub

    Sub DISATTIVA_COLL()
        Panel4.Visible = False
        stato_collaudo = False
    End Sub

    Sub coll_completato()
        Panel4.Visible = True
        Panel4.BackColor = Color.Lime
        Label9.ForeColor = Color.Black
    End Sub

#End Region

#Region "Timer Tick"

    Private Sub Timer_montaggio_Tick(sender As Object, e As EventArgs) Handles Timer_montaggio.Tick
        If stato_montaggio Then DISATTIVA_MONTAGGIO() Else ATTIVA_MONTAGGIO()
    End Sub

    Private Sub Timer_EL_Tick(sender As Object, e As EventArgs) Handles Timer_EL.Tick
        If stato_elettrico Then DISATTIVA_EL() Else ATTIVA_EL()
    End Sub

    Private Sub Timer_COLL_Tick(sender As Object, e As EventArgs) Handles Timer_COLL.Tick
        If stato_collaudo Then DISATTIVA_COLL() Else ATTIVA_COLL()
    End Sub

    Private Sub Timer_SOFT_Tick(sender As Object, e As EventArgs) Handles Timer_SOFT.Tick
        If stato_software Then DISATTIVA_SOFTW() Else ATTIVA_SOFTW()
    End Sub

#End Region

#Region "Risorse / Dipendenti"

    Sub inserisci_dipendenti(commessa As String)
        contatore_mec = 0 : contatore_el = 0
        contatore_soft = 0 : contatore_col = 0 : contatore_altro = 0

        Dim risorse As New List(Of String)

        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "SELECT Risorsa FROM [Tirelli_40].[dbo].[Layout_CAP1_risorse]
                 WHERE commessa = @commessa", CNN)
                CMD.Parameters.AddWithValue("@commessa", commessa)
                Using reader = CMD.ExecuteReader()
                    Do While reader.Read()
                        risorse.Add(reader("Risorsa").ToString())
                    Loop
                End Using
            End Using
        End Using

        ' Recupera tutti i gruppi in un'unica query NPgsql
        ' invece di N connessioni separate
        If risorse.Count = 0 Then Return

        Dim uids = String.Join(",", risorse.Select(Function(r) "'" & r & "'"))

        Using conn As New NpgsqlConnection(Homepage.JPM_TIRELLI)
            conn.Open()
            Using CMD As New NpgsqlCommand(
                "SELECT res.uid, grp.grpcod AS codice_gruppo
                 FROM angres res
                 LEFT JOIN angresgrp rg ON res.uid = rg.resuid AND rg.prjgrppri = -1
                 LEFT JOIN anggrp grp ON grp.uid = rg.grpuid
                 WHERE res.logdel = 0 AND res.uid IN (" & uids & ")", conn)

                Using reader = CMD.ExecuteReader()
                    Do While reader.Read()
                        Select Case reader("codice_gruppo").ToString()
                            Case "MONT MECC TIR", "MONT MECC KTF", "MONT MECC BRB"
                                contatore_mec += 1
                            Case "ELETTRICO"
                                contatore_el += 1
                            Case "COLL TIR", "TRASFERTISTA"
                                contatore_col += 1
                        End Select
                    Loop
                End Using
            End Using
        End Using

        If contatore_mec > 0 Then Label5.Text = contatore_mec : Panel5.Visible = True
        If contatore_el > 0 Then Label10.Text = contatore_el : Panel6.Visible = True
        If contatore_soft > 0 Then Label11.Text = contatore_soft : Panel7.Visible = True
        If contatore_col > 0 Then Label12.Text = contatore_col : Panel8.Visible = True
    End Sub

#End Region

#Region "Click e eventi UI"

    Public Event BaiaCliccata(valore As String)

    Private Sub AggiungiHandlerAiControlli(parent As Control)
        For Each ctrl As Control In parent.Controls
            AddHandler ctrl.Click, AddressOf TuttoCliccato
            If ctrl.HasChildren Then AggiungiHandlerAiControlli(ctrl)
        Next
    End Sub

    Private Sub TuttoCliccato(sender As Object, e As EventArgs)
        RaiseEvent BaiaCliccata(Label1.Text)
        FLayout.TextBox9.Text = Label1.Text
    End Sub

#End Region

#Region "Menu contestuale"

    Private Sub SpedizioneToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SpedizioneToolStripMenuItem.Click
        FLayout.cancella_record_baia(Label1.Text, numero_baia, FLayout.zona)
        FLayout.inserisci_record_baia_log(Label1.Text, numero_baia, "OUT")
        FLayout.inserisci_record_baia_spedizione(Label1.Text, numero_baia, "OUT")
        FLayout.check_presenza_commessa_baia_layout(numero_baia, FLayout.zona)
        MsgBox(Label1.Text & " spedita con successo")
    End Sub

    Private Sub CancellaMacchinaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CancellaMacchinaToolStripMenuItem.Click
        If MessageBox.Show("Vuoi definitivamente cancellare ogni record di " & Label1.Text,
                       "Cancella", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Return

        FLayout.cancella_record_baia(Label1.Text, numero_baia, FLayout.zona)
        FLayout.cancella_record_baia_log_per_baia(Label1.Text, numero_baia)
        FLayout.check_presenza_commessa_baia_layout(numero_baia, FLayout.zona)
        MsgBox(Label1.Text & " cancellata dalla baia")
    End Sub

    Private Sub CollaudoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CollaudoToolStripMenuItem.Click
        Commesse_MES.SCHEDA_COMMESSA(Label1.Text)
        Form_Scheda_Collaudi.Lbl_Commessa.Text = Label1.Text
        Form_Scheda_Collaudi.inizializzazione_form(Label1.Text)
        Form_Scheda_Collaudi.Show()
    End Sub

    Private Sub PianificazioneToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PianificazioneToolStripMenuItem.Click
        FLayout.date_jpm_commessa(Label1.Text, FLayout.DataGridView2)
        FLayout.TextBox6.Text = Label1.Text
        FLayout.TabControl1.SelectedTab = FLayout.TabPage2
    End Sub

    Private Sub SpostaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SpostaToolStripMenuItem.Click
        If Label1.Text >= "M04000" Then
            Scheda_tecnica.Close()
            Scheda_tecnica.Show()
            Scheda_tecnica.BringToFront()
            Scheda_tecnica.inizializza_scheda_tecnica(Label1.Text)
            Try
                Scheda_tecnica.codice_bp_campione = Label1.Text
            Catch
            End Try
        End If
    End Sub

    Private Sub PianificatoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PianificatoToolStripMenuItem.Click
        FLayout.aggiorna_stato_baia(Label1.Text, "P", FLayout.zona, numero_baia)
        FLayout.check_presenza_commessa_baia_layout(numero_baia, FLayout.zona)
        MsgBox(Label1.Text & " modificata a Pianificato")
    End Sub

    Private Sub RilasciatoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RilasciatoToolStripMenuItem.Click
        If FLayout.check_baia_layout_aperte(Label1.Text).Stato = "O" Then
            MsgBox("Impossibile rilasciare la macchina. Rimuovere prima da " &
                   FLayout.check_baia_layout(Label1.Text).Nome_baia)
            Return
        End If
        FLayout.aggiorna_stato_baia(Label1.Text, "O", FLayout.zona, numero_baia)
        FLayout.check_presenza_commessa_baia_layout(numero_baia, FLayout.zona)
        MsgBox(Label1.Text & " modificata a rilasciato")
    End Sub

#End Region

    Public Class dettagli_collaudo_odp
        Public docnum As String
    End Class

End Class