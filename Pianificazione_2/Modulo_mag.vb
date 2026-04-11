Imports System.Data.SqlClient
Imports Npgsql

Public Class Modulo_mag

#Region "Proprietà e stato"

    Public numero_baia As Integer
    Private isDragging As Boolean = False
    Private startPoint As Point

    Private Const WBS_FAT As String = "999"
    Private Const WBS_SPEDIZIONE As String = "7"

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
    End Sub

    Public Sub inizializza_modulo(par_commessa As String, pnl As Panel)
        Label4.Text = FLayout.giorni_lavorativi_tra(
            FLayout.ingresso_commessa(par_commessa, "Magazzino"), Date.Today).ToString() & " GG"

        If FLayout.stato_kpi Then
            SITUAZIONE_MAGAZZINO(par_commessa)
            trova_tempi_montaggio(par_commessa)
            Label5.Visible = True
            GroupBox1.Visible = True
            GroupBox2.Visible = True
        Else
            Label5.Visible = False
            GroupBox1.Visible = False
            GroupBox2.Visible = False
        End If

        conta_n_odp(par_commessa)
        aggiusta_dimensioni()
    End Sub

    Public Sub aggiusta_dimensioni()
        Me.Margin = New Padding(0)
        Me.AutoSize = True
        Me.AutoSizeMode = AutoSizeMode.GrowAndShrink
        Me.BackColor = Color.LightBlue
        Me.BorderStyle = BorderStyle.FixedSingle
        Me.PerformLayout()
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

    Private Sub CancellareMacchinaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CancellareMacchinaToolStripMenuItem.Click
        If MessageBox.Show("Vuoi definitivamente cancellare ogni record di " & Label1.Text,
                           "Cancella", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then Return

        FLayout.cancella_record_baia(Label1.Text, numero_baia, FLayout.zona)
        FLayout.cancella_record_baia_log_per_baia(Label1.Text, numero_baia)
        FLayout.check_presenza_commessa_baia_layout(numero_baia, FLayout.zona)
        MsgBox(Label1.Text & " cancellata dalla baia")
    End Sub

    Private Sub MandaInOfficinaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MandaInOfficinaToolStripMenuItem.Click
        FLayout.cancella_record_baia(Label1.Text, numero_baia, FLayout.zona)
        FLayout.inserisci_record_baia_log(Label1.Text, numero_baia, "OUT")
        FLayout.check_presenza_commessa_baia_layout(numero_baia, FLayout.zona)
        MsgBox(Label1.Text & " liberata per officina")
    End Sub

    Private Sub SpedizioneToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SpedizioneToolStripMenuItem.Click
        FLayout.cancella_record_baia(Label1.Text, numero_baia, FLayout.zona)
        FLayout.inserisci_record_baia_log(Label1.Text, numero_baia, "OUT")
        FLayout.inserisci_record_baia_spedizione(Label1.Text, numero_baia, "OUT")
        FLayout.check_presenza_commessa_baia_layout(numero_baia, FLayout.zona)
        MsgBox(Label1.Text & " spedita con successo")
    End Sub

#End Region

#Region "Dati – ODP e tempi"

    Sub conta_n_odp(par_commessa As String)
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "SELECT COUNT(t0.docnum) AS N_ODP
                 FROM [TIRELLISRLDB].DBO.owor t0
                 WHERE t0.u_prg_azs_commessa = @Commessa
                 AND (t0.status = 'P' OR t0.status = 'R')", CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                Dim result = CMD.ExecuteScalar()
                Label6.Text = If(result IsNot Nothing, "N° ODP: " & result.ToString(), "")
            End Using
        End Using
    End Sub

    Public Sub trova_tempi_montaggio(par_commessa As String)
        Label5.Text = ""
        Dim dataScadenza As Date = Date.MinValue

        Using conn As New NpgsqlConnection(Homepage.JPM_TIRELLI)
            conn.Open()
            Using CMD As New NpgsqlCommand(
                "SELECT W.WBSLVLCOD AS ATT_WBS, T.TSKRSDTSSEND AS ATT_DTPIAFIN
                 FROM PRJTSK T
                 LEFT JOIN ANGWBSLVL W ON T.WBSLVLUID = W.UID
                 LEFT JOIN PRJ P ON T.PRJUID = P.UID
                 WHERE T.LOGDEL = 0 AND P.PRJcod = @Commessa AND T.TSKKND = '1'
                 AND W.WBSLVLCOD IN (@wbs_f, @wbs_sp)", conn)

                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                CMD.Parameters.AddWithValue("@wbs_f", WBS_FAT)
                CMD.Parameters.AddWithValue("@wbs_sp", WBS_SPEDIZIONE)

                Using reader = CMD.ExecuteReader()
                    Do While reader.Read()
                        Dim wbs As String = reader("ATT_WBS").ToString()
                        Dim dtFin As Date = CDate(reader("ATT_DTPIAFIN")).Date
                        Label5.Text = If(wbs = WBS_FAT, "FAT ", "CON ") & dtFin.ToString("dd/MM")
                        dataScadenza = dtFin
                    Loop
                End Using
            End Using
        End Using

        AggiornaColuoreScadenza(dataScadenza)
    End Sub

    Private Sub AggiornaColuoreScadenza(dataScadenza As Date)
        If dataScadenza = Date.MinValue Then Return
        Dim oggi = Date.Today
        Dim lunedi = oggi.AddDays(1 - If(oggi.DayOfWeek = DayOfWeek.Sunday, 7, CInt(oggi.DayOfWeek)))

        Label5.ForeColor =
            If(dataScadenza < lunedi.AddDays(14), Color.Red,
            If(dataScadenza < lunedi.AddDays(28), Color.Orange,
            If(dataScadenza < lunedi.AddDays(42), Color.Yellow,
            Color.Green)))
    End Sub

#End Region

#Region "Situazione magazzino"

    Sub SITUAZIONE_MAGAZZINO(par_commessa As String)
        GroupBox1.Visible = True
        GroupBox2.Visible = True
        PanelBar.Controls.Clear()
        Panelbarmont.Controls.Clear()

        Dim trasferito, trasferibile, mancante As Decimal
        Dim trasferito_mont, trasferibile_mont, mancante_mont As Decimal

        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(QuerySituazioneMagazzino(), CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                Using reader = CMD.ExecuteReader()
                    If reader.Read() Then
                        CalcolaPercentuali(reader,
                                           trasferito, trasferibile, mancante,
                                           trasferito_mont, trasferibile_mont, mancante_mont)
                    End If
                End Using
            End Using
        End Using

        CreaBarraProgresso(PanelBar, trasferito, trasferibile, mancante)
        CreaBarraProgresso(Panelbarmont, trasferito_mont, trasferibile_mont, mancante_mont)
    End Sub

    ''' <summary>Calcola tutte le percentuali dalla riga del reader.</summary>
    Private Sub CalcolaPercentuali(reader As SqlDataReader,
                                   ByRef trasferito As Decimal, ByRef trasferibile As Decimal, ByRef mancante As Decimal,
                                   ByRef trasferito_mont As Decimal, ByRef trasferibile_mont As Decimal, ByRef mancante_mont As Decimal)
        Dim totale As Decimal = CDec(reader("Totale"))
        Dim totale_mont As Decimal = CDec(reader("N prem")) + CDec(reader("N mont"))

        ' Barra codici
        Dim trf As Decimal = CDec(reader("Trasferiti"))
        Dim trfb As Decimal = CDec(reader("Trasferibile"))
        trasferito = If(trf = 0, 0, Math.Round(trf / totale * 100, 1))
        trasferibile = If(trf + trfb = 0, 0,
            Math.Round((trf + trfb) / totale * 100 - trf / totale * 100, 1))
        mancante = Math.Max(0, Math.Round(100 - trasferito - trasferibile, 1))

        ' Barra montaggio
        If totale_mont > 0 Then
            Dim trfP As Decimal = CDec(reader("Trasferiti PREM"))
            Dim trfM As Decimal = CDec(reader("Trasferiti MONT"))
            Dim trfbP As Decimal = CDec(reader("Trasferibile PREM"))
            Dim trfbM As Decimal = CDec(reader("Trasferibile MONT"))
            Dim trfTot = trfP + trfM
            Dim trfbTot = trfbP + trfbM

            trasferito_mont = If(trfTot = 0, 0, Math.Round(trfTot / totale_mont * 100, 1))
            trasferibile_mont = If(trfTot + trfbTot = 0, 0,
                Math.Round((trfTot + trfbTot) / totale_mont * 100 - trfTot / totale_mont * 100, 1))
            mancante_mont = Math.Max(0, Math.Round(100 - trasferito_mont - trasferibile_mont, 1))
        End If
    End Sub

    ''' <summary>Crea i 3 pannelli colorati (verde/giallo/rosso) nella barra indicata.</summary>
    Private Sub CreaBarraProgresso(barraPanel As Panel,
                                   trasferito As Decimal, trasferibile As Decimal, mancante As Decimal)
        Dim totW = barraPanel.Width
        Dim h = barraPanel.Height

        Dim wTrf = CInt(totW * trasferito / 100)
        Dim wTrfb = CInt(totW * trasferibile / 100)
        Dim wMan = totW - wTrf - wTrfb

        barraPanel.Controls.Add(CreaPannello(Color.Lime, h, wTrf, 0, trasferito, Color.Black))
        barraPanel.Controls.Add(CreaPannello(Color.Yellow, h, wTrfb, wTrf, trasferibile, Color.Black))
        barraPanel.Controls.Add(CreaPannello(Color.OrangeRed, h, wMan, wTrf + wTrfb, mancante, Color.White))
    End Sub

    ''' <summary>Crea un singolo pannello colorato con eventuale label percentuale.</summary>
    Private Function CreaPannello(colore As Color, altezza As Integer, larghezza As Integer,
                                   left As Integer, percentuale As Decimal, foreColor As Color) As Panel
        Dim pnl As New Panel With {
            .BackColor = colore,
            .Height = altezza,
            .Width = larghezza,
            .Left = left,
            .Top = 0
        }
        If percentuale >= 5 Then
            pnl.Controls.Add(New Label With {
                .Text = $"{Math.Round(percentuale)}%",
                .ForeColor = foreColor,
                .BackColor = Color.Transparent,
                .TextAlign = ContentAlignment.MiddleCenter,
                .Dock = DockStyle.Fill
            })
        End If
        Return pnl
    End Function

    ''' <summary>Query SQL per la situazione magazzino (estratta per leggibilità).</summary>
    Private Function QuerySituazioneMagazzino() As String
        Return "SELECT t20.commessa,
                SUM(T20.Totale) AS Totale,
                SUM(t20.Trasferiti) AS Trasferiti,
                SUM(t20.[da trasferire]) AS [da trasferire],
                SUM(t20.trasferibile) AS Trasferibile,
                SUM(T20.[N PREM]) AS [N PREM],
                SUM(T20.[N MONT]) AS [N MONT],
                SUM(T20.[Trasferiti PREM]) AS [Trasferiti PREM],
                SUM(T20.[Trasferiti MONT]) AS [Trasferiti MONT],
                SUM(t20.[Trasferibile PREM]) AS [Trasferibile PREM],
                SUM(t20.[Trasferibile MONT]) AS [Trasferibile MONT]
         FROM (
             SELECT T10.N_ODP, t10.commessa,
                    SUM(t10.N) AS Totale,
                    SUM(t10.Trasferiti) AS Trasferiti,
                    SUM(t10.[da trasferire]) AS [da trasferire],
                    SUM(t10.mag01)+SUM(t10.magfer)+SUM(t10.SCA)+SUM(t10.mag03)+SUM(t10.mut) AS Trasferibile,
                    T10.[N PREM], T10.[N MONT], T10.[Trasferiti PREM], T10.[Trasferiti MONT],
                    CASE WHEN t14.U_Fase='p01501'
                         THEN SUM(t10.mag01)+SUM(t10.magfer)+SUM(t10.SCA)+SUM(t10.mag03)+SUM(t10.mut)
                         ELSE 0 END AS [Trasferibile PREM],
                    CASE WHEN t14.U_Fase='p02001'
                         THEN SUM(t10.mag01)+SUM(t10.magfer)+SUM(t10.SCA)+SUM(t10.mag03)+SUM(t10.mut)
                         ELSE 0 END AS [Trasferibile MONT]
             FROM (
                 SELECT T0.U_PRG_AZS_Commessa AS commessa, T0.DocNum AS N_ODP,
                        SUM(CASE WHEN SUBSTRING(t3.itemcode,1,1) IN ('D','C') THEN 1 ELSE 0 END) AS N,
                        SUM(CASE WHEN t3.U_prg_wip_qtadatrasf=0 AND SUBSTRING(t3.itemcode,1,1) IN ('D','C') THEN 1 ELSE 0 END) AS trasferiti,
                        SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 AND t4.dfltwh='01' AND t3.U_prg_wip_qtadatrasf<=t5.onhand AND SUBSTRING(t3.itemcode,1,1) IN ('D','C') THEN 1 ELSE 0 END) AS mag01,
                        SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 AND t3.U_prg_wip_qtadatrasf>t5.onhand AND t3.U_prg_wip_qtadatrasf<=t6.onhand AND SUBSTRING(t3.itemcode,1,1) IN ('D','C') THEN 1 ELSE 0 END) AS magfer,
                        SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 AND t4.dfltwh='SCA' AND t3.U_prg_wip_qtadatrasf<=t7.onhand AND SUBSTRING(t3.itemcode,1,1) IN ('D','C') THEN 1 ELSE 0 END) AS SCA,
                        SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 AND t4.dfltwh='03' AND t3.U_prg_wip_qtadatrasf<=t8.onhand AND SUBSTRING(t3.itemcode,1,1) IN ('D','C') THEN 1 ELSE 0 END) AS mag03,
                        SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 AND t4.dfltwh='MUT' AND t3.U_prg_wip_qtadatrasf<=t9.onhand AND SUBSTRING(t3.itemcode,1,1) IN ('D','C') THEN 1 ELSE 0 END) AS mut,
                        SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 AND SUBSTRING(t3.itemcode,1,1) IN ('D','C') THEN 1 ELSE 0 END) AS [da trasferire],
                        CASE WHEN T0.U_Fase='P01501' THEN SUM(CASE WHEN SUBSTRING(t3.itemcode,1,1) IN ('D','C') THEN 1 ELSE 0 END) ELSE 0 END AS [N PREM],
                        CASE WHEN T0.U_Fase='P02001' THEN SUM(CASE WHEN SUBSTRING(t3.itemcode,1,1) IN ('D','C') THEN 1 ELSE 0 END) ELSE 0 END AS [N MONT],
                        CASE WHEN T0.U_Fase='P01501' THEN SUM(CASE WHEN t3.U_prg_wip_qtadatrasf=0 AND SUBSTRING(t3.itemcode,1,1) IN ('D','C') THEN 1 ELSE 0 END) ELSE 0 END AS [Trasferiti PREM],
                        CASE WHEN T0.U_Fase='P02001' THEN SUM(CASE WHEN t3.U_prg_wip_qtadatrasf=0 AND SUBSTRING(t3.itemcode,1,1) IN ('D','C') THEN 1 ELSE 0 END) ELSE 0 END AS [Trasferiti MONT]
                 FROM OWOR T0
                 INNER JOIN OITM T1 ON t0.itemcode = t1.itemcode
                 LEFT JOIN [dbo].[@FASE] T2 ON T0.U_Fase = T2.Code
                 LEFT JOIN wor1 t3 ON t3.docentry = t0.docentry
                 INNER JOIN OITM T4 ON T4.itemcode = t3.itemcode
                 LEFT JOIN OITW T5 ON T5.ITEMCODE=T3.ITEMCODE AND T5.WHSCODE='01'
                 LEFT JOIN OITW T6 ON T6.ITEMCODE=T3.ITEMCODE AND T6.WHSCODE='ferretto'
                 LEFT JOIN OITW T7 ON T7.ITEMCODE=T3.ITEMCODE AND T7.WHSCODE='SCA'
                 LEFT JOIN OITW T8 ON T8.ITEMCODE=T3.ITEMCODE AND T8.WHSCODE='03'
                 LEFT JOIN OITW T9 ON T9.ITEMCODE=T3.ITEMCODE AND T9.WHSCODE='MUT'
                 LEFT JOIN OWOR T10 ON T10.ITEMCODE=T3.ITEMCODE AND T10.STATUS IN ('P','R') AND T10.U_PRODUZIONE='ASSEMBL'
                 WHERE T0.status IN ('P','R') AND T0.U_PRODUZIONE='ASSEMBL'
                 AND t3.itemtype=4 AND T0.U_PRG_AZS_Commessa=@Commessa
                 AND T10.DOCNUM IS NULL AND T4.DfltWH <> '03'
                 GROUP BY T0.DocNum, T0.ItemCode, T1.U_Disegno, T2.Name, t1.itemname,
                          T0.PlannedQty, T0.U_Fase, T0.U_stato, t0.status, T0.U_PRG_AZS_Commessa
             ) AS t10
             LEFT JOIN oitm t1 ON t10.commessa = t1.itemcode
             LEFT JOIN rdr1 t2 ON t1.itemcode = t2.itemcode AND t2.OpenQty > 0
             LEFT JOIN ordr t3 ON t3.docentry = t2.docentry AND t3.docstatus = 'O'
             LEFT JOIN OWOR T14 ON t14.docnum = T10.N_ODP
             GROUP BY N_ODP, t10.commessa, t1.itemname, t3.cardname,
                      T1.U_Final_customer_name, T3.U_Clientefinale, T3.DocDueDate,
                      T10.[N PREM], T10.[N MONT], T10.[Trasferiti PREM], T10.[Trasferiti MONT], t14.U_Fase
         ) AS T20
         GROUP BY t20.commessa
         ORDER BY t20.commessa"
    End Function

#End Region

End Class