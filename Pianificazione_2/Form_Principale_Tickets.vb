Imports System.IO
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop


Public Class Pianificazione_Tickets

    Public Administrator As Integer

    Public Elenco_Reparti(1000) As String
    Public Reparti_Caricati As Integer
    ' Public Mittente_Mail As String
    Public Password_Mail As String
    Public Ultima_Riga As String
    'Public Business As String
    Public CODICE_REPARTO As Integer
    Public filtro_reparto As String
    Public filtro_reparto_task As String
    Public filtro_commessa As String
    Public filtro_id As String
    Public filtro_id_padre As String
    Public filtro_cliente As String
    Public filtro_mittente_padre As String
    Public filtro_business As String
    Public filtro_utente_padre As String
    Public filtro_utente As String
    Public filtro_riunione As String
    Public filtro_articolo As String
    Public filtro_assegnato As String
    Public riga As Integer
    Public status_1 As String = "t0.aperto=1"
    Public status_2 As String = "t10.aperto=1"
    Public Declare Ansi Function ExtractIconEx Lib "Shell32.dll" _
    (ByVal lpszFile As String,
    ByVal nIconIndex As Integer, ByVal phIconLarge As IntPtr(),
    ByVal phIconSmall As IntPtr(), ByVal nIcons As Integer) _
    As Integer

    Public form_visualizzato As String = "Ticket"

    Public variabile_iniziazione As Integer = 0
    Public filtro_contenuto As String


    Private Sub Pianificazione_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        ApplicaStile()
    End Sub

    Private Sub ApplicaStile()
        Dim navy As Color = Color.FromArgb(22, 45, 84)
        Dim navyDark As Color = Color.FromArgb(10, 26, 55)
        Dim navyHover As Color = Color.FromArgb(30, 63, 122)
        Dim bgApp As Color = Color.FromArgb(238, 242, 247)
        Dim textColor As Color = Color.FromArgb(40, 60, 90)
        Dim fontBold As New Font("Segoe UI", 8.5, FontStyle.Bold)
        Dim fontUI As New Font("Segoe UI", 8.5, FontStyle.Regular)
        Dim fontSmall As New Font("Segoe UI", 8, FontStyle.Regular)

        Me.BackColor = bgApp
        TableLayoutPanel2.BackColor = navy
        TableLayoutPanel4.BackColor = navyDark

        ' GroupBox nei filtri
        For Each gb As GroupBox In New GroupBox() {Grp_Reparto, GroupBox2, GroupBox3, GroupBox4,
            GroupBox5, GroupBox6, GroupBox7, GroupBox8, GroupBox9, GroupBox10,
            GroupBox11, GroupBox12, GroupBox13, GroupBox14, GroupBox15, GroupBox1}
            gb.BackColor = navy
            gb.ForeColor = Color.White
            gb.Font = fontBold
        Next

        ' RadioButton stato (GroupBox5) — toggle button
        For Each rb As RadioButton In New RadioButton() {RadioButton1, RadioButton2, RadioButton3}
            rb.Appearance = Appearance.Button
            rb.FlatStyle = FlatStyle.Flat
            rb.BackColor = Color.FromArgb(14, 32, 68)
            rb.ForeColor = Color.FromArgb(180, 205, 235)
            rb.Font = New Font("Segoe UI", 7.5, FontStyle.Bold)
            rb.TextAlign = ContentAlignment.MiddleCenter
            rb.FlatAppearance.CheckedBackColor = navyHover
            rb.FlatAppearance.BorderColor = Color.FromArgb(40, 70, 130)
            rb.FlatAppearance.MouseOverBackColor = Color.FromArgb(30, 58, 115)
        Next

        ' RadioButton visualizzazione (GroupBox11: Reparto / Tutti) — toggle button
        GroupBox11.BackColor = Color.FromArgb(14, 32, 68)
        GroupBox11.ForeColor = Color.FromArgb(180, 205, 235)
        GroupBox11.Font = fontBold
        For Each rb As RadioButton In New RadioButton() {RadioButton4, RadioButton5}
            rb.Appearance = Appearance.Button
            rb.FlatStyle = FlatStyle.Flat
            rb.BackColor = Color.FromArgb(10, 24, 52)
            rb.ForeColor = Color.White
            rb.Font = New Font("Segoe UI", 8, FontStyle.Bold)
            rb.TextAlign = ContentAlignment.MiddleCenter
            rb.FlatAppearance.CheckedBackColor = navyHover
            rb.FlatAppearance.BorderColor = Color.FromArgb(40, 70, 130)
            rb.FlatAppearance.MouseOverBackColor = Color.FromArgb(30, 58, 115)
        Next

        ' TextBox
        For Each tb As TextBox In New TextBox() {TextBox1, TextBox2, TextBox3, TextBox4, TextBox5,
            TextBox6, TextBox7, TextBox8, TextBox9, TextBox10, TextBox11}
            tb.BackColor = Color.White
            tb.ForeColor = textColor
            tb.Font = New Font("Segoe UI", 9)
            tb.BorderStyle = BorderStyle.FixedSingle
        Next

        RichTextBox1.BackColor = Color.White
        RichTextBox1.ForeColor = textColor
        RichTextBox1.Font = New Font("Segoe UI", 9)
        RichTextBox1.BorderStyle = BorderStyle.FixedSingle

        Lbl_Nome_Reparto.ForeColor = Color.White
        Lbl_Nome_Reparto.Font = fontBold
        Lbl_Nome_Reparto.BackColor = Color.Transparent
        Label1.ForeColor = Color.White
        Label1.Font = fontBold
        Label1.BackColor = Color.Transparent

        ' Pulsanti barra superiore
        For Each btn As Button In New Button() {Button1, Button3}
            btn.BackColor = navyDark
            btn.ForeColor = Color.White
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderColor = navyHover
            btn.FlatAppearance.MouseOverBackColor = navyHover
            btn.Font = fontBold
        Next

        Cmd_Cambia.BackColor = navyHover
        Cmd_Cambia.ForeColor = Color.White
        Cmd_Cambia.FlatStyle = FlatStyle.Flat
        Cmd_Cambia.FlatAppearance.BorderColor = navy
        Cmd_Cambia.FlatAppearance.MouseOverBackColor = Color.FromArgb(50, 85, 150)
        Cmd_Cambia.Font = fontUI

        ' Pulsanti barra inferiore
        For Each btn As Button In New Button() {Button2, Button4, Button5, Button15, Cmd_Nuovo}
            btn.BackColor = navy
            btn.ForeColor = Color.White
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderColor = navyHover
            btn.FlatAppearance.MouseOverBackColor = navyHover
            btn.Font = fontUI
        Next

        TabControl1.BackColor = bgApp
        TabControl1.Font = fontBold

        ' DataGridView1
        With DataGridView1
            .BackgroundColor = bgApp
            .BorderStyle = BorderStyle.None
            .GridColor = Color.FromArgb(200, 210, 225)
            .EnableHeadersVisualStyles = False
            .ColumnHeadersDefaultCellStyle.BackColor = navy
            .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            .ColumnHeadersDefaultCellStyle.Font = fontBold
            .ColumnHeadersDefaultCellStyle.Padding = New Padding(4, 0, 0, 0)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 248, 252)
            .DefaultCellStyle.Font = fontSmall
            .DefaultCellStyle.SelectionBackColor = navyHover
            .DefaultCellStyle.SelectionForeColor = Color.White
        End With

        ' DataGridView2
        With DataGridView2
            .BackgroundColor = bgApp
            .BorderStyle = BorderStyle.None
            .GridColor = Color.FromArgb(200, 210, 225)
            .EnableHeadersVisualStyles = False
            .ColumnHeadersDefaultCellStyle.BackColor = navy
            .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            .ColumnHeadersDefaultCellStyle.Font = fontBold
            .ColumnHeadersDefaultCellStyle.Padding = New Padding(4, 0, 0, 0)
            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 248, 252)
            .DefaultCellStyle.Font = fontSmall
            .DefaultCellStyle.SelectionBackColor = navyHover
            .DefaultCellStyle.SelectionForeColor = Color.White
        End With

        ' DataGridView3 (Statistiche)
        With DataGridView3
            .BackgroundColor = bgApp
            .BorderStyle = BorderStyle.None
            .GridColor = Color.FromArgb(200, 210, 225)
            .EnableHeadersVisualStyles = False
            .ColumnHeadersDefaultCellStyle.BackColor = navy
            .ColumnHeadersDefaultCellStyle.ForeColor = Color.White
            .ColumnHeadersDefaultCellStyle.Font = fontBold
            .AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(245, 248, 252)
            .DefaultCellStyle.Font = New Font("Segoe UI", 10)
            .DefaultCellStyle.SelectionBackColor = navyHover
            .DefaultCellStyle.SelectionForeColor = Color.White
        End With

        ' Dashboard panel
        Panel_Dashboard.BackColor = navyDark
        Panel_Dashboard.Padding = New Padding(0)

        ' Separa dashboard dal TabControl con un bordo sottile a destra
        Panel_Dashboard.Margin = New Padding(0, 0, 2, 0)
    End Sub

    ' ----------------------------------------------------------------
    ' Dashboard reparto
    ' ----------------------------------------------------------------
    Private _dashBuilt As Boolean = False
    Private _lblDashReparto As Label
    Private _lblDashTicketNum As Label
    Private _lblDashGiorniNum As Label
    Private _barTicket As Panel   ' barra accent colorata ticket
    Private _barGiorni As Panel   ' barra accent colorata giorni

    Sub AggiornaDashboard()
        If Not _dashBuilt Then
            CostribuisciDashboard()
            _dashBuilt = True
        End If

        Dim ticketAperti As Integer = 0
        Dim giorniMedi As Integer = 0

        Dim cnn As New SqlConnection(Homepage.sap_tirelli)
        Try
            cnn.Open()
            Dim cmd As New SqlCommand(
                "SELECT COUNT(*) AS Ticket_Aperti,
                        COALESCE(AVG(DATEDIFF(day, t0.Data_Creazione, GETDATE())), 0) AS Giorni_Medi
                 FROM [TIRELLI_40].[DBO].coll_tickets t0
                 WHERE t0.aperto = 1 AND t0.destinatario = @rep", cnn)
            cmd.Parameters.AddWithValue("@rep", CODICE_REPARTO)
            Dim r As SqlDataReader = cmd.ExecuteReader()
            If r.Read() Then
                ticketAperti = CInt(r("Ticket_Aperti"))
                giorniMedi = CInt(r("Giorni_Medi"))
            End If
            r.Close()
        Catch
        Finally
            cnn.Close()
        End Try

        _lblDashReparto.Text = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).nome_reparto
        _lblDashTicketNum.Text = ticketAperti.ToString()
        _lblDashGiorniNum.Text = giorniMedi.ToString()

        ' Colore adattivo sui giorni medi
        Dim colGiorni As Color
        If giorniMedi <= 2 Then
            colGiorni = Color.FromArgb(60, 210, 120)        ' verde
        ElseIf giorniMedi <= 4 Then
            colGiorni = Color.FromArgb(255, 215, 50)        ' giallo
        ElseIf giorniMedi <= 5 Then
            colGiorni = Color.FromArgb(255, 145, 30)        ' arancione
        Else
            colGiorni = Color.FromArgb(240, 65, 65)         ' rosso
        End If
        _lblDashGiorniNum.ForeColor = colGiorni
        _barGiorni.BackColor = colGiorni
    End Sub

    Private Sub CostribuisciDashboard()
        ' ── Palette ──────────────────────────────────────────────
        Dim bg As Color = Color.FromArgb(11, 22, 46)
        Dim cardBg As Color = Color.FromArgb(18, 38, 74)
        Dim cardBgAlt As Color = Color.FromArgb(14, 28, 58)
        Dim accentBlue As Color = Color.FromArgb(70, 158, 255)
        Dim accentAmber As Color = Color.FromArgb(255, 168, 40)
        Dim muted As Color = Color.FromArgb(110, 145, 190)
        Dim divider As Color = Color.FromArgb(22, 44, 84)

        Panel_Dashboard.BackColor = bg
        Panel_Dashboard.Controls.Clear()

        ' ── HEADER ───────────────────────────────────────────────
        Dim pHeader As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 64,
            .BackColor = Color.FromArgb(7, 14, 34),
            .Padding = New Padding(16, 10, 10, 6)
        }
        Dim lblCaption As New Label With {
            .Text = "IL MIO REPARTO",
            .Font = New Font("Segoe UI", 7, FontStyle.Bold),
            .ForeColor = accentBlue,
            .AutoSize = False,
            .Dock = DockStyle.Top,
            .Height = 18,
            .TextAlign = ContentAlignment.BottomLeft
        }
        _lblDashReparto = New Label With {
            .Text = "—",
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = False,
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleLeft
        }
        pHeader.Controls.Add(_lblDashReparto)
        pHeader.Controls.Add(lblCaption)
        Panel_Dashboard.Controls.Add(pHeader)

        Panel_Dashboard.Controls.Add(New Panel With {.Dock = DockStyle.Top, .Height = 1, .BackColor = divider})

        ' ── CARD TICKET APERTI ───────────────────────────────────
        Dim pCard1 As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 148,
            .BackColor = cardBg
        }
        _barTicket = New Panel With {
            .Dock = DockStyle.Left,
            .Width = 5,
            .BackColor = accentBlue
        }
        Dim pC1Inner As New Panel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.Transparent,
            .Padding = New Padding(14, 12, 14, 6)
        }
        Dim lbl1Cap As New Label With {
            .Text = "TICKET APERTI",
            .Font = New Font("Segoe UI", 7.5, FontStyle.Bold),
            .ForeColor = muted,
            .AutoSize = False,
            .Dock = DockStyle.Top,
            .Height = 22,
            .TextAlign = ContentAlignment.MiddleLeft
        }
        _lblDashTicketNum = New Label With {
            .Text = "—",
            .Font = New Font("Segoe UI", 52, FontStyle.Bold),
            .ForeColor = Color.White,
            .AutoSize = False,
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleCenter
        }
        pC1Inner.Controls.Add(_lblDashTicketNum)
        pC1Inner.Controls.Add(lbl1Cap)
        pCard1.Controls.Add(pC1Inner)
        pCard1.Controls.Add(_barTicket)
        Panel_Dashboard.Controls.Add(pCard1)

        Panel_Dashboard.Controls.Add(New Panel With {.Dock = DockStyle.Top, .Height = 1, .BackColor = divider})

        ' ── CARD GIORNI MEDI ─────────────────────────────────────
        Dim pCard2 As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 148,
            .BackColor = cardBg
        }
        _barGiorni = New Panel With {
            .Dock = DockStyle.Left,
            .Width = 5,
            .BackColor = accentAmber
        }
        Dim pC2Inner As New Panel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.Transparent,
            .Padding = New Padding(14, 12, 14, 6)
        }
        Dim lbl2Cap As New Label With {
            .Text = "GIORNI MEDI APERTURA",
            .Font = New Font("Segoe UI", 7.5, FontStyle.Bold),
            .ForeColor = muted,
            .AutoSize = False,
            .Dock = DockStyle.Top,
            .Height = 22,
            .TextAlign = ContentAlignment.MiddleLeft
        }
        _lblDashGiorniNum = New Label With {
            .Text = "—",
            .Font = New Font("Segoe UI", 52, FontStyle.Bold),
            .ForeColor = accentAmber,
            .AutoSize = False,
            .Dock = DockStyle.Fill,
            .TextAlign = ContentAlignment.MiddleCenter
        }
        pC2Inner.Controls.Add(_lblDashGiorniNum)
        pC2Inner.Controls.Add(lbl2Cap)
        pCard2.Controls.Add(pC2Inner)
        pCard2.Controls.Add(_barGiorni)
        Panel_Dashboard.Controls.Add(pCard2)

        Panel_Dashboard.Controls.Add(New Panel With {.Dock = DockStyle.Top, .Height = 1, .BackColor = divider})

        ' ── LEGENDA ──────────────────────────────────────────────
        Dim pLeg As New Panel With {
            .Dock = DockStyle.Top,
            .Height = 68,
            .BackColor = cardBgAlt,
            .Padding = New Padding(16, 8, 14, 6)
        }
        Dim lblLegTitle As New Label With {
            .Text = "SCALA TEMPI",
            .Font = New Font("Segoe UI", 7, FontStyle.Bold),
            .ForeColor = muted,
            .AutoSize = False,
            .Dock = DockStyle.Top,
            .Height = 18,
            .TextAlign = ContentAlignment.BottomLeft
        }
        Dim pLegDots As New FlowLayoutPanel With {
            .Dock = DockStyle.Fill,
            .BackColor = Color.Transparent,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = True,
            .Padding = New Padding(0, 2, 0, 0)
        }
        For Each entry In New (String, Color)() {
            ("≤2 gg",  Color.FromArgb(60, 210, 120)),
            ("≤4 gg",  Color.FromArgb(255, 215, 50)),
            ("≤5 gg",  Color.FromArgb(255, 145, 30)),
            (">5 gg",   Color.FromArgb(240, 65, 65))}
            Dim dot As New Label With {
                .Text = "● " & entry.Item1,
                .Font = New Font("Segoe UI", 7.5, FontStyle.Regular),
                .ForeColor = entry.Item2,
                .AutoSize = True,
                .Margin = New Padding(0, 0, 8, 0)
            }
            pLegDots.Controls.Add(dot)
        Next
        pLeg.Controls.Add(pLegDots)
        pLeg.Controls.Add(lblLegTitle)
        Panel_Dashboard.Controls.Add(pLeg)

        ' ── PULSANTE AGGIORNA ────────────────────────────────────
        Dim btnRicarica As New Button With {
            .Text = "↻   Aggiorna",
            .Dock = DockStyle.Bottom,
            .Height = 44,
            .BackColor = Color.FromArgb(30, 63, 122),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat,
            .Font = New Font("Segoe UI", 9.5, FontStyle.Bold),
            .Cursor = Cursors.Hand
        }
        btnRicarica.FlatAppearance.BorderSize = 0
        btnRicarica.FlatAppearance.MouseOverBackColor = Color.FromArgb(48, 92, 160)
        AddHandler btnRicarica.Click, Sub(s, ev) AggiornaDashboard()
        Panel_Dashboard.Controls.Add(btnRicarica)
    End Sub

    Sub inizializzazione_form()
        variabile_iniziazione = 1
        Lbl_Nome_Reparto.Text = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).nome_reparto
        Carica_Reparti()
        riempi_tickets(DataGridView1)
        AggiornaDashboard()
        variabile_iniziazione = 1
    End Sub






    Private Sub Cmd_Cambia_Click(sender As Object, e As EventArgs) Handles Cmd_Cambia.Click
        Form_Cambia_Reparto.Show()

    End Sub


    Private Sub Cmd_Nuovo_Click(sender As Object, e As EventArgs) Handles Cmd_Nuovo.Click
        Ultima_Riga = ""
        Form_nuovo_ticket.Show()
        Form_nuovo_ticket.Inserimento_dipendenti()

        Form_nuovo_ticket.Reparto = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto
        Form_nuovo_ticket.Administrator = 1
        Form_nuovo_ticket.Startup()
        Form_nuovo_ticket.ComboBox2.Text = Homepage.business

    End Sub

    Private Function Reparto(Num_Reparto As Integer) As String
        Return Elenco_Reparti(Num_Reparto)
    End Function

    Private Sub Carica_Reparti()




        Dim Cnn_Reparto As New SqlConnection
        Cnn_Reparto.ConnectionString = Homepage.sap_tirelli
        Cnn_Reparto.Open()
        Dim Cmd_Reparto As New SqlCommand
        Dim Reader_Reparto As SqlDataReader

        Cmd_Reparto.Connection = Cnn_Reparto
        Cmd_Reparto.CommandText = "SELECT Id_Reparto,Descrizione 
FROM [TIRELLI_40].[DBO].COLL_Reparti
WHERE active ='Y' "
        Reader_Reparto = Cmd_Reparto.ExecuteReader()


        Do While Reader_Reparto.Read()
            Elenco_Reparti(Reader_Reparto("Id_Reparto")) = Reader_Reparto("Descrizione")
        Loop
        Cnn_Reparto.Close()
    End Sub


    Private Sub Cmd_Esci_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        riempi_tickets(DataGridView1)
        AggiornaDashboard()
    End Sub


    Sub riempi_tickets(par_datagridview As DataGridView)

        Dim contatore As Integer = 0

        If RadioButton4.Checked = True Then
            filtro_reparto = "and t0.destinatario= '" & CODICE_REPARTO & "'"

        Else
            filtro_reparto = ""
        End If
        par_datagridview.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        CMD_SAP_2.CommandText = "

SELECT *
from
(
select
t0.[Id_Ticket]
,coalesce(t5.[Descrizione_Motivo],'') as 'Descrizione_Motivo'
      ,t0.[Commessa], 
case WHEN t12.itemname is not null then t12.itemname when t6.itemname is null then '' else t6.itemname end as 'Itemname',
	  case when substring(t0.COMMESSA,1,3)='CDS' THEN case when t12.U_Final_customer_name is null then t11.custmrName else t12.U_Final_customer_name end when t7.cardname is null then t6.u_final_customer_name else t7.cardname end as 'Cliente'
      ,t0.[Data_Creazione]
	  , DATEDIFF(day,t0.[Data_Creazione], getdate()) as 'giorni'
      ,t0.[Data_Chiusura]
      ,t0.[Data_Prevista_Chiusura]
      ,t0.[Aperto]
      ,t0.[Descrizione] 
	  ,t4.Descrizione as 'Mittente_padre'
      ,t1.[descrizione] as 'Mittente'
      ,t2.[descrizione] as 'Destinatario'
      ,t0.[Immagine]
      ,t0.[Id_Padre]
      ,t0.[Business]
, t0.oggetto
      ,t0.[Utente], concat(t9.firstname,' ', t9.lastname) as 'Nome_utente'
, concat(t10.firstname,' ', t10.lastname) as 'Utente_padre'
, case when t0.assegnato is null then '' else concat(t8.firstname,' ', t8.lastname) end as 'Assegnato'
      ,t0.[Data_chiusura_totale], case when t0.aperto =1 then 'Y' else 'N' end as 'stato'
,coalesce(t0.tpr,'') as 'TPR'
,coalesce(t0.riunione,'') as 'Riunione'

from  [TIRELLI_40].[DBO].coll_tickets t0 
  left join [TIRELLI_40].[DBO].COLL_Reparti t1 on t1.Id_Reparto=t0.Mittente
  left join [TIRELLI_40].[DBO].COLL_Reparti t2 on t2.Id_Reparto=t0.destinatario
  left join [TIRELLI_40].[DBO].COLL_Tickets t3 on t3.Id_Ticket= t0.id_padre
  left join [TIRELLI_40].[DBO].COLL_Reparti t4 on t4.Id_Reparto=t3.Mittente
  left join [TIRELLI_40].[DBO].COLL_motivazione t5 on t5.Id_Motivo = t0.Motivazione
LEFT JOIN [TIRELLISRLDB].[DBO].oitm t6 on t6.itemcode=t0.[Commessa]
left join [TIRELLISRLDB].[DBO].ocrd t7 on t7.cardcode=t6.u_final_customer_code
left join [TIRELLI_40].[DBO].ohem t8 on t8.empid=t0.assegnato
left join [TIRELLI_40].[DBO].ohem t9 on t9.empid=t0.utente
left join [TIRELLI_40].[DBO].ohem t10 on t10.empid=t3.utente
left join [TIRELLISRLDB].[DBO].oscl t11 on cast(t11.callid as varchar) = CAST(substring(t0.COMMESSA,4,999) AS VARCHAR) and substring(t0.COMMESSA,1,3)='CDS'
left join [TIRELLISRLDB].[DBO].oitm t12 on t12.itemcode=t11.itemcode
" & filtro_articolo & "
left join
(select t0.[Id_Padre], max(t0.[Id_Ticket]) as 'Ticket_max' from [TIRELLI_40].[DBO].coll_tickets t0 group by t0.[Id_Padre] ) a on t0.[Id_Ticket]=a.[Ticket_max]

 where " & status_1 & " " & filtro_reparto & " " & filtro_commessa & " " & filtro_id & " " & filtro_id_padre & "
)
as t10

 where " & status_2 & " " & filtro_cliente & " " & filtro_mittente_padre & " " & filtro_business & " " & filtro_utente_padre & " " & filtro_utente & filtro_riunione & filtro_contenuto & filtro_assegnato & " 

  order by t10.giorni DESC"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader_2.Read()
            Try


                par_datagridview.Rows.Add(cmd_SAP_reader_2("Id_Ticket"), cmd_SAP_reader_2("Id_padre"), cmd_SAP_reader_2("Descrizione_Motivo"), cmd_SAP_reader_2("Commessa"), cmd_SAP_reader_2("Itemname"), cmd_SAP_reader_2("Cliente"), cmd_SAP_reader_2("Business"), cmd_SAP_reader_2("Mittente_padre"), cmd_SAP_reader_2("Mittente"), cmd_SAP_reader_2("Destinatario"), cmd_SAP_reader_2("Assegnato"), cmd_SAP_reader_2("Data_creazione"), cmd_SAP_reader_2("Giorni"), cmd_SAP_reader_2("Stato"), cmd_SAP_reader_2("Oggetto"), cmd_SAP_reader_2("Riunione"), cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("TPR"))
            Catch ex As Exception

            End Try
            contatore += 1
        Loop

        Label1.Text = contatore

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        par_datagridview.ClearSelection()

    End Sub



    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)
        If variabile_iniziazione = 0 Then

        Else
            If form_visualizzato = "Ticket" Then
                riempi_tickets(DataGridView1)

            ElseIf form_visualizzato = "task" Then
                riempi_tasks()
            End If
        End If


    End Sub



    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If e.RowIndex < 0 Then Return
        Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

        ' Riga "Richiesta di Miglioria" → sfondo grigio
        If row.Cells("Causale").Value?.ToString() = "Richiesta di Miglioria" Then
            row.DefaultCellStyle.BackColor = Color.LightGray
            row.DefaultCellStyle.ForeColor = Color.DimGray
        End If

        ' Colonna stato aperto/chiuso
        If e.ColumnIndex = DataGridView1.Columns("Open").Index Then
            If row.Cells("Open").Value?.ToString() = "Y" Then
                e.CellStyle.BackColor = Color.FromArgb(220, 60, 60)
                e.CellStyle.ForeColor = Color.White
            Else
                e.CellStyle.BackColor = Color.FromArgb(50, 170, 80)
                e.CellStyle.ForeColor = Color.White
            End If
        End If

        ' Colonna giorni: verde → giallo → arancione → rosso in base all'età
        If e.ColumnIndex = DataGridView1.Columns("Giorni").Index Then
            Dim giorni As Integer
            If Integer.TryParse(row.Cells("Giorni").Value?.ToString(), giorni) Then
                If giorni <= 2 Then
                    e.CellStyle.BackColor = Color.FromArgb(144, 238, 144)
                ElseIf giorni <= 4 Then
                    e.CellStyle.BackColor = Color.FromArgb(255, 230, 80)
                ElseIf giorni <= 5 Then
                    e.CellStyle.BackColor = Color.FromArgb(255, 160, 20)
                    e.CellStyle.ForeColor = Color.White
                Else
                    e.CellStyle.BackColor = Color.FromArgb(200, 40, 40)
                    e.CellStyle.ForeColor = Color.White
                End If
            End If
        End If
    End Sub

    ''' <summary>Imposta un filtro LIKE e ricarica la griglia.</summary>
    Private Sub ApplicaFiltroLike(ByRef filtro As String, ByVal valore As String, ByVal campo As String)
        filtro = If(String.IsNullOrEmpty(valore), "", "and " & campo & " Like '%%" & valore & "%%'")
        riempi_tickets(DataGridView1)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ApplicaFiltroLike(filtro_commessa, TextBox1.Text, "t0.commessa")
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        If e.RowIndex >= 0 Then
            riga = e.RowIndex

            If e.ColumnIndex = DataGridView1.Columns.IndexOf(ID) Then

                Dim new_form_visualizza_ticket = New Form_Visualizza_Ticket

                new_form_visualizza_ticket.Show()



                new_form_visualizza_ticket.Show()
                new_form_visualizza_ticket.Inserimento_dipendenti()

                new_form_visualizza_ticket.Reparto = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto
                new_form_visualizza_ticket.Administrator = 1
                new_form_visualizza_ticket.Txt_Id.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ID").Value
                new_form_visualizza_ticket.Startup()

            End If
        End If

    End Sub

    Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
        Form_Visualizza_Ticket.Show()
        Form_Visualizza_Ticket.Inserimento_dipendenti()

        Form_Visualizza_Ticket.Reparto = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto
        Form_Visualizza_Ticket.Administrator = 1
        Form_Visualizza_Ticket.Txt_Id.Text = DataGridView1.Rows(riga).Cells(columnName:="ID").Value
        Form_Visualizza_Ticket.Startup()

    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "[]" Then

            Me.WindowState = FormWindowState.Maximized
            Button1.Text = "Riduci"
        ElseIf Button1.Text = "Riduci" Then
            Me.WindowState = FormWindowState.Normal
            Button1.Text = "[]"
        End If

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        ApplicaFiltroLike(filtro_cliente, TextBox2.Text, "t10.cliente")
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        status_1 = "t0.aperto=1"
        status_2 = "t10.aperto=1"
        If variabile_iniziazione = 0 Then

        Else
            riempi_tickets(DataGridView1)
        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        status_1 = "t0.aperto=0"
        status_2 = "t10.aperto=0"
        If variabile_iniziazione = 0 Then
        Else
            riempi_tickets(DataGridView1)
        End If

    End Sub



    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        ApplicaFiltroLike(filtro_mittente_padre, TextBox3.Text, "t10.mittente_padre")
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        ApplicaFiltroLike(filtro_id, TextBox4.Text, "t0.[Id_Ticket]")
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        ApplicaFiltroLike(filtro_business, TextBox5.Text, "t10.business")
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        ApplicaFiltroLike(filtro_id_padre, TextBox6.Text, "t0.[Id_padre]")
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        ApplicaFiltroLike(filtro_utente_padre, TextBox7.Text, "t10.Utente_padre")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Process.Start("\\tirfs01\tirelli\00-Responsible\KPI\Analisi tickets.xlsx")
    End Sub

    Sub riempi_tasks()

        DataGridView2.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        CMD_SAP_2.CommandText = "select t0.id,t0.oc, t1.cardname
, CASE WHEN T5.CARDNAME IS NULL THEN '' ELSE t5.cardname end as 'Cliente_finale'
, t0.task, t2.Nome_task, t0.reparto,t3.Descrizione,t0.riferimento
, t4.riferimento as 'Nome_riferimento', t0.giorni,t0.stato,t0.linenum
, t0.data_inizio, t0.data_fine, t0.id_link, t0.Data_chiusura_task, t0.Ora_chiusura_task 

from [Tirelli_40].[dbo].[Pianificazione_CDS] t0 inner join ordr t1 on t0.oc=t1.docnum
left join [Tirelli_40].[dbo].[Pianificazione_CDS_TASK] t2 on t2.id =t0.task
left join [TIRELLI_40].[DBO].COLL_Reparti t3 on t0.reparto=t3.Id_Reparto

  left join [Tirelli_40].[dbo].[Pianificazione_CDS_Riferimenti] t4 on t0.Riferimento=t4.id
left join ocrd t5 on t5.cardcode=t1.u_CODICEBP

inner join

(select t0.oc, min(t0.linenum) as 'linenum'
from [Tirelli_40].[dbo].[Pianificazione_CDS] t0 where t0.stato='P'
group by t0.oc) A on A.linenum=t0.Linenum and t0.oc=a.oc where 0=0 " & filtro_reparto_task & "

"



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader_2.Read()

            DataGridView2.Rows.Add(cmd_SAP_reader_2("Id"), cmd_SAP_reader_2("OC"), cmd_SAP_reader_2("cardname"), cmd_SAP_reader_2("Cliente_finale"), cmd_SAP_reader_2("Task"), cmd_SAP_reader_2("Nome_task"), cmd_SAP_reader_2("Reparto"), cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("riferimento"), cmd_SAP_reader_2("Nome_riferimento"), cmd_SAP_reader_2("Giorni"), cmd_SAP_reader_2("Stato"), cmd_SAP_reader_2("Linenum"), cmd_SAP_reader_2("Data_inizio"), cmd_SAP_reader_2("Data_fine"), cmd_SAP_reader_2("Id_link"), cmd_SAP_reader_2("Data_chiusura_task"), cmd_SAP_reader_2("Ora_chiusura_task"))

        Loop



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        DataGridView2.ClearSelection()

    End Sub

    Private Sub tabpage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Enter

        riempi_tickets(DataGridView1)
        form_visualizzato = "Ticket"

    End Sub

    Private Sub tabpage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter
        riempi_tasks()
        form_visualizzato = "task"
    End Sub

    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        riempi_statistiche()
    End Sub

    Sub riempi_statistiche()
        DataGridView3.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Try
            Cnn1.Open()
            Dim cmd As New SqlCommand
            cmd.Connection = Cnn1
            cmd.CommandText = "
SELECT coalesce(t2.Descrizione, '(Senza reparto)') AS Reparto,
       COUNT(*) AS Ticket_Aperti,
       AVG(DATEDIFF(day, t0.Data_Creazione, GETDATE())) AS Giorni_Medi
FROM [TIRELLI_40].[DBO].coll_tickets t0
LEFT JOIN [TIRELLI_40].[DBO].COLL_Reparti t2 ON t2.Id_Reparto = t0.destinatario
WHERE t0.aperto = 1
GROUP BY t0.destinatario, t2.Descrizione
ORDER BY COUNT(*) DESC"
            Dim reader As SqlDataReader = cmd.ExecuteReader()
            Do While reader.Read()
                DataGridView3.Rows.Add(reader("Reparto"), reader("Ticket_Aperti"), reader("Giorni_Medi"))
            Loop
            reader.Close()
        Catch ex As Exception
            MessageBox.Show("Errore caricamento statistiche: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Finally
            Cnn1.Close()
        End Try
    End Sub



    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.RowIndex >= 0 Then
            riga = e.RowIndex

            If e.ColumnIndex = DataGridView2.Columns.IndexOf(ID_) Then
                Task_visualizza.id_task = DataGridView2.Rows(e.RowIndex).Cells(columnName:="ID_").Value

                Task_visualizza.Show()


            End If
        End If
    End Sub



    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        ApplicaFiltroLike(filtro_utente, TextBox8.Text, "t10.nome_Utente")
    End Sub



    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If variabile_iniziazione = 0 Then

        Else
            If form_visualizzato = "Ticket" Then
                riempi_tickets(DataGridView1)

            ElseIf form_visualizzato = "task" Then
                riempi_tasks()
            End If
        End If
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form_impostazioni_ticket.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        status_1 = "(t0.aperto=1 or t0.aperto=0)"
        status_2 = "(t10.aperto=1 or t10.aperto=0)"
        If variabile_iniziazione = 0 Then

        Else
            riempi_tickets(DataGridView1)
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim par_datagridview As DataGridView = DataGridView1
        ' Creare un'applicazione Excel
        Dim excelApp As New Excel.Application
        excelApp.Visible = True ' Mostrare Excel all'utente

        ' Creare un nuovo foglio di lavoro
        Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add
        Dim excelWorksheet As Excel.Worksheet = CType(excelWorkbook.Worksheets(1), Excel.Worksheet)

        ' Aggiungere intestazioni alla prima riga del foglio di lavoro (facoltativo)
        For col As Integer = 1 To par_datagridview.Columns.Count
            excelWorksheet.Cells(1, col) = par_datagridview.Columns(col - 1).HeaderText
        Next

        ' Aggiungere dati alla DataGridView al foglio di lavoro
        For row As Integer = 0 To par_datagridview.Rows.Count - 1
            For col As Integer = 0 To par_datagridview.Columns.Count - 1
                excelWorksheet.Cells(row + 2, col + 1) = par_datagridview.Rows(row).Cells(col).Value
            Next
        Next

        ' Salvare il file Excel
        Dim saveFileDialog As New SaveFileDialog
        saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            excelWorkbook.SaveAs(saveFileDialog.FileName)
            MessageBox.Show("Esportazione completata con successo!", "Successo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        ' Chiudere Excel
        excelApp.Quit()
        ReleaseComObject(excelApp)
    End Sub

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        filtro_riunione = If(String.IsNullOrEmpty(TextBox9.Text), "", "and coalesce(t10.riunione,'') Like '%%" & TextBox9.Text & "%%'")
        riempi_tickets(DataGridView1)
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        filtro_articolo = If(String.IsNullOrEmpty(TextBox10.Text), "", " inner join [Tirelli_40].[dbo].[COLL_Riferimenti] t13 on t13.codice_sap Like '%%" & TextBox10.Text & "%%' and t13.Rif_Ticket=t0.[Id_Ticket] ")
        riempi_tickets(DataGridView1)
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        ApplicaFiltroLike(filtro_contenuto, RichTextBox1.Text, "t10.descrizione")
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        ApplicaFiltroLike(filtro_assegnato, TextBox11.Text, "t10.Assegnato")
    End Sub
End Class
