Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Diagnostics
Imports System.IO
Imports System.Net.Http
Imports System.Net.Mail
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Security.Policy
Imports System.Threading
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView
Imports System.Windows.Media.Media3D
Imports AxFOXITREADERLib
Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word

Imports Excel = Microsoft.Office.Interop.Excel



Public Class Homepage
    Public Form_precedente As Integer
    Public inter As Integer = 0
    Public orario_chiusura As String
    Public orario_chiusura_FINE As String
    Public dashboard As String
    Public ora_spegnimento As Integer
    Public MINUTO_spegnimento As Integer
    Public stringa_connessione As String

    Public Percorso_Immagini_TICKETS As String
    Public Percorso_Schede As String
    Public Mittente_Mail As String
    Public business As String
    Public Percorso_procedure As String
    Public Percorso_immagini As String
    Public percorso_server As String
    Public percorso_disegni_generico As String
    Public percorso_cartelle_macchine As String
    Public percorso_progetti As String
    Public percorso_OFF As String
    Public percorso_DXF As String
    Public percorso_ODA As String
    Public percorso_disegni As String
    Public percorso_DWF As String
    Public percorso_offerte_vendita As String
    Public percorso_acquisti As String
    Public PERCORSO_DOCUMENTO_ODP_ETICHETTA As String
    Public PERCORSO_DOCUMENTO_ODP As String
    Public PERCORSO_CARTELLE_OPPORTUNITà As String
    Public PERCORSO_QUALITA As String
    Public PERCORSO_DOCUMENTO_OC As String
    Public PERCORSO_statistiche_lavorazioni As String
    Public percorso_PDM_BRB As String
    Public Percorso_Immagini_TICKETS_HELPDESK As String
    Public Percorso_FILE_TICKETS_HELPDESK As String
    Public percorso_costificatore_seng As String
    Public sap_tirelli As String
    Public sap_4life As String
    Public sap_prova As String
    Public JPM_TIRELLI As String
    Public tempo_stampe_scontrini As Integer = 2000



    Public Elenco_dipendenti_UT(1000) As String
    Public Elenco_nome_sap_tirelli_UT(1000) As String

    Public indice_UT As Integer
    Public totem As String
    Public ID_SALVATO As String
    'Public UTENTE_NOME_SALVATO As String


    'Public codice_licenza_sap_tirelli As String
    'Public Branch As String
    Public Centro_di_costo As String



    Public da_aggiornare_ini As String
    Public PROGRAMMA As String
    Public password_mail As String


    Public Stampante_Selezionata As Boolean = False

    Public azienda As String = "Tirelli"
    Public logo_azienda As String = "T:\00-Tirelli 4.0\Immagini generiche\TIRELLI_blue_nopayoff_360x250.png"

    Public colore_sfondo As New Color
    Public commessa As String
    ' Variabile di stato per controllare il ciclo
    Private stopCycle As Boolean = False
    'GIOVANNI
    Public ERP_provenienza As String = "GALILEO"
    'Public ERP_provenienza As String = "SAP"





    Private Sub Button_pianificazione_Click(sender As Object, e As EventArgs) Handles Button_pianificazione.Click

        Pianificazione.Show()

    End Sub

    Private Sub MES_Click(sender As Object, e As EventArgs) Handles MES.Click

        Commesse_magazzino.magazzino = 0
        Commesse_MES.Commesse_odp_aperte(Commesse_MES.DataGridView_commesse, Commesse_MES.TextBox_commessa.Text, Commesse_MES.TextBox1.Text, Commesse_MES.TextBox2.Text, Commesse_MES.CheckBox1.Checked)
        Commesse_MES.Show()
        Commesse_MES.Refresh()

    End Sub

    Private Sub Button_MU_Click(sender As Object, e As EventArgs) Handles Button_MU.Click




        Dashboard_MU_New.inizializzazione_dashboard_mu()


        Dashboard_MU_New.Show()


    End Sub



    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim kpi As Boolean = False
        If kpi = True Then


            Form_KPI.avvia_KPI("kpi_sales.xlsx", 15000)
            Form_KPI.Timer1.Start()
            Me.Hide()
        Else


            Try
                leggi_ini_orario()
            Catch ex As Exception

            End Try



            Try
                leggi_ini_server()
            Catch ex As Exception

            End Try

            Try
                leggi_ini_computer()

            Catch ex As Exception

            End Try

            'Timer1.Start()


            operazioni_dopo_lettura_ini()


        End If
        'Funzioni_utili.Show()

        ApplicaStile()

    End Sub

    ' ═══════════════════════════════════════════════════════════════════
    '  STILE MODERNO — brand Tirelli (navy #162D54)
    ' ═══════════════════════════════════════════════════════════════════
    Private Sub ApplicaStile()

        ' ── Palette ───────────────────────────────────────────────────
        Dim navy As Color = Color.FromArgb(22, 45, 84)          ' #162D54  brand Tirelli
        Dim navyDark As Color = Color.FromArgb(10, 26, 55)      ' #0A1A37  zona logo
        Dim navyHover As Color = Color.FromArgb(30, 63, 122)    ' #1E3F7A  hover sidebar
        Dim bgApp As Color = Color.FromArgb(238, 242, 247)      ' #EEF2F7  sfondo app
        Dim bgCard As Color = Color.White                        ' card moduli
        Dim cardBorder As Color = Color.FromArgb(210, 220, 235) ' bordo card
        Dim cardHover As Color = Color.FromArgb(232, 239, 252)  ' hover card
        Dim cardPress As Color = Color.FromArgb(210, 225, 245)  ' click card
        Dim sideText As Color = Color.FromArgb(195, 212, 238)   ' testo sec. sidebar
        Dim fontUI As String = "Segoe UI"

        ' ── Form ──────────────────────────────────────────────────────
        Me.BackColor = bgApp

        ' ── Sidebar ───────────────────────────────────────────────────
        Panel1.BackColor = navy
        Panel2.BackColor = navyDark
        Panel23.BackColor = navy
        Label5.ForeColor = sideText
        Label5.Font = New System.Drawing.Font(fontUI, 7.5F, System.Drawing.FontStyle.Regular)
        Label5.BackColor = Color.Transparent

        Dim sidebarBtns() As System.Windows.Forms.Button = {Button13, Button_dashboard, Button7, Button24}
        For Each b As System.Windows.Forms.Button In sidebarBtns
            b.BackColor = Color.Transparent
            b.ForeColor = Color.White
            b.FlatStyle = FlatStyle.Flat
            b.FlatAppearance.BorderSize = 0
            b.FlatAppearance.MouseOverBackColor = navyHover
            b.FlatAppearance.MouseDownBackColor = navyDark
            b.Font = New System.Drawing.Font(fontUI, 10.5F, System.Drawing.FontStyle.Regular)
            b.TextAlign = ContentAlignment.MiddleCenter
            b.Padding = New Padding(0, 4, 0, 4)
        Next

        ' ── Top bar ───────────────────────────────────────────────────
        TableLayoutPanel5.BackColor = Color.White
        TableLayoutPanel4.BackColor = bgApp

        Dim topFont As New System.Drawing.Font(fontUI, 8.5F, System.Drawing.FontStyle.Bold)
        Dim labelFont As New System.Drawing.Font(fontUI, 8.5F, System.Drawing.FontStyle.Regular)
        For Each gb As GroupBox In {GroupBox5, GroupBox2, GroupBox3, GroupBox1, GroupBox6}
            gb.Font = topFont
            gb.ForeColor = navy
            gb.BackColor = Color.White
        Next
        FlowLayoutPanel2.BackColor = Color.White
        For Each lbl As Label In {Label2, Label3, Label6, Label4, Label7, Label8, Label1}
            lbl.Font = labelFont
            lbl.ForeColor = Color.FromArgb(40, 60, 90)
        Next

        Button15.FlatStyle = FlatStyle.Flat
        Button15.BackColor = navy
        Button15.ForeColor = Color.White
        Button15.Font = New System.Drawing.Font(fontUI, 9.5F, System.Drawing.FontStyle.Bold)
        Button15.FlatAppearance.BorderSize = 0
        Button15.FlatAppearance.MouseOverBackColor = navyHover
        Button15.FlatAppearance.MouseDownBackColor = navyDark
        Button18.FlatStyle = FlatStyle.Flat
        Button18.BackColor = navy
        Button18.ForeColor = Color.White
        Button18.Font = New System.Drawing.Font(fontUI, 8.0F, System.Drawing.FontStyle.Regular)
        Button18.FlatAppearance.BorderSize = 0

        ' ── Grid moduli (4×3) ─────────────────────────────────────────
        TableLayoutPanel1.BackColor = bgApp
        TableLayoutPanel1.Padding = New Padding(6)
        TableLayoutPanel1.CellBorderStyle = TableLayoutPanelCellBorderStyle.None

        Dim gridBtns() As System.Windows.Forms.Button = {
            Button2, Button_UT, Button_UA,
            Button_pianificazione, Button4, Button_MU, Button11,
            MES, Button1, Button6, Button16
        }
        For Each b As System.Windows.Forms.Button In gridBtns
            b.BackColor = bgCard
            b.ForeColor = navy
            b.FlatStyle = FlatStyle.Flat
            b.FlatAppearance.BorderColor = cardBorder
            b.FlatAppearance.BorderSize = 1
            b.FlatAppearance.MouseOverBackColor = cardHover
            b.FlatAppearance.MouseDownBackColor = cardPress
            b.Font = New System.Drawing.Font(fontUI, 13.0F, System.Drawing.FontStyle.Bold)
            b.TextAlign = ContentAlignment.BottomCenter
            b.TextImageRelation = TextImageRelation.ImageAboveText
            b.Padding = New Padding(0, 14, 0, 12)
            b.Margin = New Padding(5)
        Next

        ' ── Pannello destro ───────────────────────────────────────────
        Panel16.BackColor = navy
        Panel17.BackColor = navyDark
        TableLayoutPanel3.BackColor = navy
        TableLayoutPanel7.BackColor = navy
        TableLayoutPanel8.BackColor = navy

        Dim rightBtns() As System.Windows.Forms.Button = {Button5, Button10, Button14, Button12}
        For Each b As System.Windows.Forms.Button In rightBtns
            b.BackColor = Color.Transparent
            b.ForeColor = Color.White
            b.FlatStyle = FlatStyle.Flat
            b.FlatAppearance.BorderSize = 0
            b.FlatAppearance.MouseOverBackColor = navyHover
            b.FlatAppearance.MouseDownBackColor = navyDark
            b.Font = New System.Drawing.Font(fontUI, 10.0F, System.Drawing.FontStyle.Bold)
            b.TextAlign = ContentAlignment.BottomCenter
            b.TextImageRelation = TextImageRelation.ImageAboveText
        Next
        ' Button8 (icona grande, nessun testo)
        Button8.BackColor = Color.Transparent
        Button8.FlatStyle = FlatStyle.Flat
        Button8.FlatAppearance.BorderSize = 0
        Button8.FlatAppearance.MouseOverBackColor = navyHover

        ' Bottoni mini (lavorazioni aperte, JPM, consuntivo, guida JPM)
        For Each b As System.Windows.Forms.Button In {Button_lavorazioni_aperte, Button17, Button_manodopera, Button22}
            b.BackColor = Color.Transparent
            b.ForeColor = sideText
            b.FlatStyle = FlatStyle.Flat
            b.FlatAppearance.BorderSize = 0
            b.FlatAppearance.MouseOverBackColor = navyHover
            b.Font = New System.Drawing.Font(fontUI, 7.5F, System.Drawing.FontStyle.Regular)
        Next

        ' Bottone chiudi/minimizza
        Button3.BackColor = Color.FromArgb(180, 40, 40)
        Button3.ForeColor = Color.White
        Button3.FlatStyle = FlatStyle.Flat
        Button3.FlatAppearance.BorderSize = 0
        Button3.FlatAppearance.MouseOverBackColor = Color.FromArgb(210, 55, 55)
        Button3.Font = New System.Drawing.Font(fontUI, 22.0F, System.Drawing.FontStyle.Bold)

        ' PictureBox logo su sfondo navy
        PictureBox2.BackColor = navyDark
        PictureBox1.BackColor = navyDark

    End Sub

    Sub kpi()

    End Sub

    Public Function Trova_regola_dist(PAR_commessa As String)

        Dim Regola_distribuzione As String = ""



        Dim Cnn_Matricola As New SqlConnection
        Dim Cmd_Matricola As New SqlCommand
        Dim Cmd_Matricola_Reader As SqlDataReader

        Cnn_Matricola.ConnectionString = sap_tirelli
        Cnn_Matricola.Open()
        Cmd_Matricola.Connection = Cnn_Matricola
        Cmd_Matricola.CommandText = "
select coalesce(t11.ocrcode,'') as 'ocrcode'
from
(
select max(t1.docentry) as 'docentry'
from TIRELLISRLDB.DBO.ordr t0 
inner join TIRELLISRLDB.DBO.rdr1 t1 on t0.docentry=t1.docentry
where t1.itemcode='" & PAR_commessa & "'
)
as t10 left join TIRELLISRLDB.DBO.rdr1 t11 on t11.docentry=t10.docentry and t11.itemcode='" & PAR_commessa & "'"

        Cmd_Matricola_Reader = Cmd_Matricola.ExecuteReader
        If Cmd_Matricola_Reader.Read() Then



            Regola_distribuzione = Cmd_Matricola_Reader("ocrcode")



        End If
        Cmd_Matricola_Reader.Close()
        Cnn_Matricola.Close()
        Return Regola_distribuzione
    End Function





    Sub operazioni_dopo_lettura_ini()

        da_aggiornare_ini = "NO"

        If totem = Nothing Then

            Discriminante_Totem_PC.Show()
            Me.Enabled = False
        End If

        If totem = "Y" Then
            'Form_Cambia_Reparto.Show()
            'Me.Enabled = False
        ElseIf totem = "Y" And trova_Dettagli_dipendente(ID_SALVATO).codice_reparto <> Nothing Then
            Label4.Text = "TOTEM"
            Label1.Text = "TOTEM"
        End If

        If totem = "N" And ID_SALVATO = Nothing Then
            Form_gestione_utente.Show()
            Me.Enabled = False
        ElseIf totem = "N" And ID_SALVATO <> Nothing Then

            Label1.Text = "PC"
        End If



        da_aggiornare_ini = "SI"
        Try
            If ID_SALVATO = 221 Then
                Me.WindowState = FormWindowState.Minimized

                'Timer1.Stop()
                Timer2.Stop()
                Presence.Show()
                'ellys
            ElseIf ID_SALVATO = 181 Then
                Presence.Timer_revisioni.Start()
                'Gianluca
            ElseIf ID_SALVATO = 4 Then
                Presence.Timer_revisioni.Start()

                '    Gio
                'ElseIf ID_SALVATO = 75 Then
                '    Presence.Timer_revisioni.Start()



            End If
        Catch ex As Exception

        End Try



        Try
            If trova_Dettagli_dipendente(ID_SALVATO).codice_reparto = "7" Or trova_Dettagli_dipendente(ID_SALVATO).codice_reparto = "10" Or trova_Dettagli_dipendente(ID_SALVATO).codice_reparto = "16" Or trova_Dettagli_dipendente(ID_SALVATO).codice_reparto = "23" Then
                Button9.Visible = True
            Else


            End If
        Catch ex As Exception

        End Try



        assegna_dati_dipendente_a_label()

    End Sub


    Sub assegna_dati_dipendente_a_label()
        Label2.Text = trova_DETTAGLI_dipendente(ID_SALVATO).nome
        Label3.Text = trova_Dettagli_dipendente(ID_SALVATO).cognome
        Label6.Text = ID_SALVATO
        Label9.Text = trova_Dettagli_dipendente(ID_SALVATO).utente_galileo
        Label4.Text = trova_Dettagli_dipendente(ID_SALVATO).NOME_REPARTO
        Label7.Text = trova_Dettagli_dipendente(ID_SALVATO).codice_reparto
        Label8.Text = trova_Dettagli_dipendente(ID_SALVATO).id_reparto_ticket
        Form_cambia_dipendente.Combo_dipendenti.Text = trova_Dettagli_dipendente(ID_SALVATO).cognome & " " & trova_Dettagli_dipendente(ID_SALVATO).nome

    End Sub
    Public Function trova_Dettagli_dipendente(par_id As String)
        Dim dettagli As New Dettaglidipendente()

        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = sap_tirelli
        Cnn.Open()

        Dim CMD_sap_tirelli As New SqlCommand
        Dim cmd_sap_tirelli_reader As SqlDataReader
        CMD_sap_tirelli.CommandTimeout = 0
        CMD_sap_tirelli.Connection = Cnn

        If ERP_provenienza = "SAP" Then
            CMD_sap_tirelli.CommandText = "SELECT T0.[lastName] as 'Cognome' , T0.[firstName] AS 'Nome', T1.[name] 
, case when T2.USER_CODE is null then '' else t2.user_code end AS 'Nome_sap_tirelli'

, t0.dept, t3.descrizione,t3.id_reparto as 'Reparto'
, case when T0.[userid] is null then '50' else t0.userid end as 'Codice_licenza_erp_tirelli'

, coalesce(t0.u_codice_pdm,'') as 'Codice_PDM'
,COALESCE(T0.U_REPARTO_TICKETs,0) AS 'Codice_reparto_ticket'
,coalesce(t0.[Galileo],'') as 'Utente_galileo'
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
        LEFT JOIN [TIRELLIsrldb].[dbo].OUSR T2 ON T0.[userId] = T2.[USERID] 
left join [TIRELLI_40].[DBO].coll_reparti t3 on t3.ID_REPARTO=t0.u_reparto_tickets
where t0.empid='" & par_id & "' "

        ElseIf ERP_provenienza = "GALILEO" Then
            CMD_sap_tirelli.CommandText = "SELECT T0.[lastName] as 'Cognome' 
, T0.[firstName] AS 'Nome', T1.[name] , case when T2.USER_CODE is null then '' else t2.user_code end AS 'Nome_sap_tirelli'

, t0.dept
, coalesce(t3.descrizione,'') as 'Descrizione'
,coalesce(t3.id_reparto,0) as 'Reparto'
, case when T0.[userid] is null then '50' else t0.userid end as 'Codice_licenza_erp_tirelli'

, coalesce(t0.u_codice_pdm,'') as 'Codice_PDM'
,COALESCE(T0.U_REPARTO_TICKETs,0) AS 'Codice_reparto_ticket'
,coalesce(t0.[Galileo],'') as 'Utente_galileo'
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
        LEFT JOIN [TIRELLIsrldb].[dbo].OUSR T2 ON T0.[userId] = T2.[USERID] 
left join [TIRELLI_40].[DBO].coll_reparti t3 on t3.ID_REPARTO=t0.u_reparto_tickets
where t0.empid='" & par_id & "' OR t0.empid='" & Replace(par_id, "T", "") & "'  "

        ElseIf ERP_provenienza = "GALILEOPOST" Then

            CMD_sap_tirelli.CommandText = "select 
t0.cogn_dip as 'Cognome'
, t0.nome_dip as 'Nome'
,t0.qual_des as 'name'
,'' as 'Nome_ERP'
,0 as 'Position'
,t0.qual_dip as 'Dept'
,t0.qual_des as 'Descrizione'
,t0.rep_dip as 'Reparto'
,t0.prof_gal as 'Codice_licenza_Erp_tirelli'
,0 as 'Branch0'
,'' as 'Costcenter'
,'' as 'Codice_pdm'
,COALESCE(T0.U_REPARTO_TICKET,0) AS 'Codice_reparto_ticket'

FROM [Tirelli_40_restored].[dbo].[JGAL_Dipendenti] t0
where t0.cod_dip='" & "T" & par_id.ToString("000") & "' OR t0.cod_dip='" & par_id & "' "

        End If
        cmd_sap_tirelli_reader = CMD_sap_tirelli.ExecuteReader


        If cmd_sap_tirelli_reader.Read() Then

            dettagli.COGNOME = cmd_sap_tirelli_reader("Cognome")
            dettagli.NOME = cmd_sap_tirelli_reader("nome")
            dettagli.ID_REPARTO_TICKET = cmd_sap_tirelli_reader("Codice_reparto_ticket")
            dettagli.NOME_REPARTO_TICKET = cmd_sap_tirelli_reader("descrizione")
            dettagli.utente_galileo = cmd_sap_tirelli_reader("utente_galileo")
            If Not DBNull.Value.Equals(cmd_sap_tirelli_reader("reparto")) Then
                dettagli.CODICE_REPARTO = cmd_sap_tirelli_reader("reparto")
            Else
                dettagli.CODICE_REPARTO = 9999
            End If
            If Not DBNull.Value.Equals(cmd_sap_tirelli_reader("descrizione")) Then
                dettagli.NOME_REPARTO = cmd_sap_tirelli_reader("descrizione")
                If totem = "N" Then
                    Label4.Text = dettagli.NOME_REPARTO
                End If
            End If
            dettagli.utente_sap_salvato = cmd_sap_tirelli_reader("Codice_licenza_erp_tirelli")
        Else
            dettagli.COGNOME = ""
            dettagli.NOME = ""
        End If

        cmd_sap_tirelli_reader.Close()
        Cnn.Close()
        Return dettagli
    End Function

    Public Class Dettaglidipendente
        Public Descrizione As String
        Public COGNOME As String
        Public NOME As String
        Public CODICE_REPARTO As String
        Public NOME_REPARTO As String
        Public ID_REPARTO_TICKET As String
        Public NOME_REPARTO_TICKET As String
        Public utente_sap_salvato As String
        Public utente_galileo As String

    End Class






    Sub partenza_screensaver()
        If dashboard = "Y" Then
            Timer2.Stop()
            Timer2.Start()
            Count.Start()
        Else
            Timer2.Stop()
            Count.Stop()
        End If
    End Sub



    Sub Aggiorna_stato_programma()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = sap_tirelli


        Cnn.Open()

        Dim Cmd_sap_tirelli As New SqlCommand

        'Inserisco i valori nell'odp
        Cmd_sap_tirelli.Connection = Cnn
        Cmd_sap_tirelli.CommandText = "update Config_4_0 set value='N' where date='Programma'"
        Cmd_sap_tirelli.ExecuteNonQuery()

        Cnn.Close()


    End Sub 'Aggiorno la data dell'odp



    Private Sub Button_manodopera_Click(sender As Object, e As EventArgs) Handles Button_manodopera.Click

        Consuntivo1.dipendente_manodopera = Nothing
        Consuntivo1.ComboBox_dipendente.Text = Nothing
        Consuntivo1.data_selezione = Nothing
        Consuntivo1.risorsa_manodopera = Nothing
        Consuntivo1.ComboBox_risorse.Text = Nothing


        Lavorazioni_MES.inserimento_dipendenti_MES(Consuntivo1.ComboBox_dipendente, Consuntivo1.Elenco_dipendenti)
        Consuntivo1.Inserimento_risorse()
        If totem = "N" Then
            Consuntivo1.ComboBox_dipendente.Text = trova_Dettagli_dipendente(ID_SALVATO).cognome & " " & trova_Dettagli_dipendente(ID_SALVATO).nome
        End If

        Consuntivo1.Show()
        If Lavorazioni_MES.ComboBox_dipendente.SelectedIndex > 0 Then
            Consuntivo1.Lavorazioni_aperte(Lavorazioni_MES.Elenco_dipendenti_MES(Lavorazioni_MES.ComboBox_dipendente.SelectedIndex), Lavorazioni_MES.DataGridView_lavorazioni)

        End If



    End Sub

    Private Sub Button_UA_Click(sender As Object, e As EventArgs) Handles Button_UA.Click

        Acquisti.Show()

        Acquisti.Inserimento_dipendenti()
        '  Acquisti.Inserimento_fasi(Acquisti.ComboBox4)
        '  If Acquisti.ComboBox1.Text = "RDO" Then
        '  Acquisti.lista_documenti(Acquisti.DataGridView1)
        ' End If

    End Sub



    Sub leggi_ini_server()

        Dim File_INI_Stream As StreamReader
        Dim CryptedIniFilePath As String = ".\MES.INI"
        Dim decypheredStream As Stream = Nothing

        'CryptFile("T1r3l11@4zero!?", CryptedIniFilePath)

        If File.Exists(CryptedIniFilePath) Then
            decypheredStream = DecryptFile("T1r3l11@4zero!?", CryptedIniFilePath)
        End If

        Dim Str_Lettura As String
        If Not IsNothing(decypheredStream) Then

            File_INI_Stream = New StreamReader(decypheredStream)
            Str_Lettura = File_INI_Stream.ReadLine

            Pianificazione.ConfigDictionary = New Dictionary(Of String, String)

            Do While Not Str_Lettura Is Nothing

                If Str_Lettura.StartsWith("[SAP_TIRELLI]=") Then

                    sap_tirelli = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)
                    ' sap_tirelli = "Data Source=srvtirsap01.corp.arol-group.com;Initial Catalog=TIRELLISRLDB;Persist Security Info=True;User ID=sa;Password=123B1Admin"

                    Pianificazione.ConfigDictionary.Add("SAP_TIRELLI", sap_tirelli)
                End If

                If Str_Lettura.StartsWith("[JPM_TIRELLI]=") Then

                    JPM_tirelli = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)


                    Pianificazione.ConfigDictionary.Add("JPM_TIRELLI", JPM_TIRELLI)
                End If

                If Str_Lettura.StartsWith("[SAP_4LIFE]=") Then
                    sap_4life = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("SAP_4LIFE", sap_4life)
                End If

                If Str_Lettura.StartsWith("[SAP_PROVA]=") Then
                    sap_prova = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("SAP_PROVA", sap_prova)
                End If

                If Str_Lettura.StartsWith("[PWD]=") Then

                    password_mail = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PWD", password_mail)

                End If



                If Str_Lettura.StartsWith("[DB]=") Then
                    stringa_connessione = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("DB", stringa_connessione)

                End If


                If Str_Lettura.StartsWith("[IMG]=") Then
                    Percorso_immagini = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("IMG", Percorso_immagini)

                End If

                If Str_Lettura.StartsWith("[IMG_TICKETS]=") Then
                    Percorso_Immagini_TICKETS = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("IMG_TICKETS", Percorso_Immagini_TICKETS)

                End If

                If Str_Lettura.StartsWith("[IMG_TICKETS_HELPDESK]=") Then
                    Percorso_Immagini_TICKETS_HELPDESK = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("IMG_TICKETS_HELPDESK", Percorso_Immagini_TICKETS_HELPDESK)

                End If

                If Str_Lettura.StartsWith("[FILE_TICKETS_HELPDESK]=") Then
                    Percorso_FILE_TICKETS_HELPDESK = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("FILE_TICKETS_HELPDESK", Percorso_FILE_TICKETS_HELPDESK)

                End If

                If Str_Lettura.StartsWith("[SCH]=") Then
                    Percorso_Schede = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("SCH", Percorso_Schede)

                End If

                If Str_Lettura.StartsWith("[EML]=") Then
                    Mittente_Mail = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("EML", Mittente_Mail)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_PROCEDURE]=") Then
                    Percorso_procedure = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_PROCEDURE", Percorso_procedure)


                End If

                If Str_Lettura.StartsWith("[PROGRAMMA]=") Then
                    PROGRAMMA = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PROGRAMMA", PROGRAMMA)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_SERVER]=") Then
                    percorso_server = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_SERVER", percorso_server)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_DISEGNI_GENERICO]=") Then
                    percorso_disegni_generico = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_DISEGNI_GENERICO", percorso_disegni_generico)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_CARTELLE_MACCHINE]=") Then
                    percorso_cartelle_macchine = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_CARTELLE_MACCHINE", percorso_cartelle_macchine)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_PROGETTI]=") Then
                    percorso_progetti = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_PROGETTI", percorso_progetti)

                End If

                'If Str_Lettura.StartsWith("[PERCORSO_OFF]=") Then
                '    percorso_OFF = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                '    Pianificazione.ConfigDictionary.Add("PERCORSO_OFF", percorso_OFF)

                'End If


                If Str_Lettura.StartsWith("[PERCORSO_DXF]=") Then
                    percorso_DXF = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_DXF", percorso_DXF)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_ODA]=") Then
                    percorso_ODA = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_ODA", percorso_ODA)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_DISEGNI]=") Then
                    percorso_disegni = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_DISEGNI", percorso_disegni)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_DWF]=") Then
                    percorso_DWF = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_DWF", percorso_DWF)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_OFFERTE_VENDITA]=") Then
                    percorso_offerte_vendita = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_OFFERTE_VENDITA", percorso_offerte_vendita)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_ACQUISTI]=") Then
                    percorso_acquisti = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_ACQUISTI", percorso_acquisti)

                End If
                If Str_Lettura.StartsWith("[PERCORSO_DOCUMENTO_ODP]=") Then
                    PERCORSO_DOCUMENTO_ODP = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_DOCUMENTO_ODP", PERCORSO_DOCUMENTO_ODP)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_DOCUMENTO_ODP_ETICHETTA]=") Then
                    PERCORSO_DOCUMENTO_ODP_ETICHETTA = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_DOCUMENTO_ODP_ETICHETTA", PERCORSO_DOCUMENTO_ODP_ETICHETTA)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_CARTELLE_OPPORTUNITA]=") Then
                    PERCORSO_CARTELLE_OPPORTUNITà = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_CARTELLE_OPPORTUNITA", PERCORSO_CARTELLE_OPPORTUNITà)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_QUALITA]=") Then
                    PERCORSO_QUALITA = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_QUALITA", PERCORSO_QUALITA)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_DOCUMENTO_OC]=") Then
                    PERCORSO_DOCUMENTO_OC = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_DOCUMENTO_OC", PERCORSO_DOCUMENTO_OC)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_STATISTICHE_LAVORAZIONI]=") Then
                    PERCORSO_statistiche_lavorazioni = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_STATISTICHE_LAVORAZIONI", PERCORSO_statistiche_lavorazioni)

                End If

                If Str_Lettura.StartsWith("[PERCORSO_PDM_BRB]=") Then

                    percorso_PDM_BRB = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("PERCORSO_PDM_BRB", percorso_PDM_BRB)

                End If

                If Str_Lettura.StartsWith("[COSTIFICATORE_SENG]=") Then

                    percorso_costificatore_seng = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                    Pianificazione.ConfigDictionary.Add("COSTIFICATORE_SENG", percorso_costificatore_seng)

                End If



                Str_Lettura = File_INI_Stream.ReadLine
            Loop

            File_INI_Stream.Close()
        End If


    End Sub

    ''test encrypt

    Public Function DecryptFile(ByVal password As String,
                         ByVal filePath As String) As Stream

        Dim retVal As Stream = Nothing
        Dim ms As MemoryStream = New MemoryStream()

        Using in_stream As New FileStream(filePath,
        FileMode.Open, FileAccess.Read)
            CryptStream(password, in_stream, ms,
                    False)
        End Using

        retVal = New MemoryStream(ms.ToArray())
        retVal.Seek(0, SeekOrigin.Begin)

        Return retVal
    End Function

    Public Sub CryptFile(ByVal password As String,
                         ByVal filePath As String)

        Dim outFilePath As String =
            String.Format(".\{0}_crypted{1}",
                Path.GetFileNameWithoutExtension(filePath),
                Path.GetExtension(filePath))


        Using in_stream As New FileStream(filePath,
        FileMode.Open, FileAccess.Read)
            Using out_stream As New FileStream(outFilePath,
            FileMode.Create, FileAccess.Write)
                CryptStream(password, in_stream, out_stream,
                    True)
            End Using
        End Using

        File.Delete(filePath)
        File.Move(outFilePath, filePath)

    End Sub

    Public Sub CryptStream(ByVal password As String,
        ByVal in_stream As Stream, ByVal out_stream As Stream, ByVal encrypt As Boolean)
        ' Make an AES service provider.
        Dim aes_provider As New AesCryptoServiceProvider()

        ' Find a valid key size for this provider.
        Dim key_size_bits As Integer = 0
        For i As Integer = 1024 To 1 Step -1
            If (aes_provider.ValidKeySize(i)) Then
                key_size_bits = i
                Exit For
            End If
        Next i
        Debug.Assert(key_size_bits > 0)
        Console.WriteLine("Key size: " & key_size_bits)

        ' Get the block size for this provider.
        Dim block_size_bits As Integer = aes_provider.BlockSize

        ' Generate the key and initialization vector.
        Dim key() As Byte = Nothing
        Dim iv() As Byte = Nothing
        Dim salt() As Byte = {&H0, &H0, &H1, &H2, &H3, &H4,
        &H5, &H6, &HF1, &HF0, &HEE, &H21, &H22, &H45}
        MakeKeyAndIV(password, salt, key_size_bits,
        block_size_bits, key, iv)

        ' Make the encryptor or decryptor.
        Dim crypto_transform As ICryptoTransform
        If (encrypt) Then
            crypto_transform =
            aes_provider.CreateEncryptor(key, iv)
        Else
            crypto_transform =
            aes_provider.CreateDecryptor(key, iv)
        End If

        ' Attach a crypto stream to the output stream.
        ' Closing crypto_stream sometimes throws an
        ' exception if the decryption didn't work
        ' (e.g. if we use the wrong password).
        Try
            Using crypto_stream As New CryptoStream(out_stream,
            crypto_transform, CryptoStreamMode.Write)
                ' Encrypt or decrypt the file.
                Const block_size As Integer = 1024
                Dim buffer(block_size) As Byte
                Dim bytes_read As Integer
                Do
                    ' Read some bytes.
                    bytes_read = in_stream.Read(buffer, 0,
                    block_size)
                    If (bytes_read = 0) Then Exit Do

                    ' Write the bytes into the CryptoStream.
                    crypto_stream.Write(buffer, 0, bytes_read)
                Loop
            End Using
        Catch
        End Try

        crypto_transform.Dispose()
    End Sub

    Private Sub MakeKeyAndIV(ByVal password As String, ByVal _
    salt() As Byte, ByVal key_size_bits As Integer, ByVal _
    block_size_bits As Integer, ByRef key() As Byte, ByRef _
    iv() As Byte)
        Dim derive_bytes As New Rfc2898DeriveBytes(password,
        salt, 1000)

        key = derive_bytes.GetBytes(key_size_bits / 8)
        iv = derive_bytes.GetBytes(block_size_bits / 8)
    End Sub

    ''end test encrypt

    Sub leggi_ini_orario()
        Dim File_INI_Stream As StreamReader
        Dim Str_Lettura As String
        If File.Exists(".\orario.INI") Then

            File_INI_Stream = My.Computer.FileSystem.OpenTextFileReader(".\orario.INI")
            Str_Lettura = File_INI_Stream.ReadLine
            Do While Not Str_Lettura Is Nothing
                If Str_Lettura.StartsWith("[ORARIO_CHIUSURA]=") Then
                    orario_chiusura = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                End If
                If Str_Lettura.StartsWith("[ORA_SPEGNIMENTO]=") Then
                    ora_spegnimento = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                End If

                If Str_Lettura.StartsWith("[MINUTO_SPEGNIMENTO]=") Then
                    MINUTO_spegnimento = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                End If

                If Str_Lettura.StartsWith("[ORARIO_CHIUSURA_FINE]=") Then
                    orario_chiusura_FINE = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                End If

                Str_Lettura = File_INI_Stream.ReadLine

            Loop
            File_INI_Stream.Close()


        End If


    End Sub



    Sub leggi_ini_computer()
        Dim File_INI_Stream As StreamReader
        Dim Str_Lettura As String
        If File.Exists("C:\MES\MES.INI") Then

            File_INI_Stream = My.Computer.FileSystem.OpenTextFileReader("C:\MES\MES.INI")
            Str_Lettura = File_INI_Stream.ReadLine
            Do While Not Str_Lettura Is Nothing

                If Str_Lettura.StartsWith("[COMMESSA]=") Then

                    Pianificazione.commessa = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                End If

                If Str_Lettura.StartsWith("[DASHBOARD]=") Then

                    dashboard = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                End If

                If Str_Lettura.StartsWith("[TOTEM]=") Then

                    totem = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                End If

                If Str_Lettura.StartsWith("[ID_SALVATO]=") Then

                    ID_SALVATO = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                End If

                ' If Str_Lettura.StartsWith("[UTENTE_SALVATO]=") Then

                'UTENTE_NOME_SALVATO = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                '  End If




                'If Str_Lettura.StartsWith("[CODICE_REPARTO]=") Then

                '    codice_reparto = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)

                'End If

                If Str_Lettura.StartsWith("[BUSINESS]=") Then
                    business = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)
                End If

                If Str_Lettura.StartsWith("[TEMPO_TIMER_STAMPE_SCONTRINI]=") Then
                    tempo_stampe_scontrini = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)
                End If


                'If Str_Lettura.StartsWith("[UTENTE_SAP_SALVATO]=") Then
                '    UTENTE_sap_SALVATO = Str_Lettura.Remove(0, Str_Lettura.IndexOf("=") + 1)
                'End If

                Str_Lettura = File_INI_Stream.ReadLine

            Loop
            File_INI_Stream.Close()
        Else
            If Directory.Exists("C:\MES\") Then
                Aggiorna_INI_COMPUTER()

                Return
            Else
                Directory.CreateDirectory("C:\MES\")
                Aggiorna_INI_COMPUTER()

                Return
            End If
        End If

        If business = Nothing Then
            business = "CONTINUING"
            Aggiorna_INI_COMPUTER()
        End If



    End Sub

    Public Sub Aggiorna_INI_COMPUTER()
        Dim File_INI_Stream As StreamWriter
        File_INI_Stream = My.Computer.FileSystem.OpenTextFileWriter("C:\MES\MES.INI", False)
        File_INI_Stream.WriteLine("[COMMESSA]=" & Pianificazione.commessa)
        File_INI_Stream.WriteLine("[DASHBOARD]=" & dashboard)
        File_INI_Stream.WriteLine("[TOTEM]=" & totem)
        File_INI_Stream.WriteLine("[ID_SALVATO]=" & ID_SALVATO)
        'File_INI_Stream.WriteLine("[UTENTE_SALVATO]=" & UTENTE_NOME_SALVATO)
        'File_INI_Stream.WriteLine("[CODICE_REPARTO]=" & codice_reparto)
        File_INI_Stream.WriteLine("[BUSINESS]=" & business)
        ' File_INI_Stream.WriteLine("[UTENTE_SAP_SALVATO]=" & UTENTE_sap_SALVATO)
        File_INI_Stream.WriteLine("[TEMPO_TIMER_STAMPE_SCONTRINI]=" & tempo_stampe_scontrini)


        File_INI_Stream.Close()
        File_INI_Stream = Nothing
    End Sub



    Private Sub Button_UT_Click(sender As Object, e As EventArgs) Handles Button_UT.Click

        Disambiguazione_UT.Show()







    End Sub

    Private Sub Button_lavorazioni_aperte_Click(sender As Object, e As EventArgs) Handles Button_lavorazioni_aperte.Click
        Form_precedente = 0
        Lavorazioni_MES.ComboBox_dipendente.Text = ""
        Lavorazioni_MES.ComboBox_risorse.Text = ""



        Lavorazioni_MES.inserimento_dipendenti_MES(Lavorazioni_MES.ComboBox_dipendente, Lavorazioni_MES.Elenco_dipendenti_MES)
        Lavorazioni_MES.Inserimento_risorse_MES(Lavorazioni_MES.ComboBox_risorse)
        Lavorazioni_MES.GroupBox3.Hide()
        Lavorazioni_MES.Button_start.Hide()
        Lavorazioni_MES.Button_stop.Show()
        Lavorazioni_MES.GroupBox1.Hide()
        Lavorazioni_MES.GroupBox2.Hide()

        Lavorazioni_MES.Lavorazioni_aperte(Lavorazioni_MES.DataGridView_lavorazioni, 0, 1)

        Lavorazioni_MES.Show()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Form_gestione_campioni.Show()


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.WindowState = FormWindowState.Minimized


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Sales_disambiguazione.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        ' Ripristina la finestra se è minimizzata
        If Magazzino.WindowState = FormWindowState.Minimized Then
            Magazzino.WindowState = FormWindowState.Normal
        End If

        ' Porta la finestra in primo piano
        Magazzino.BringToFront()
        Magazzino.Activate()
        Magazzino.Show()




        Commesse_magazzino.magazzino = 1


        'new_form_magazzino.Inserimento_dipendenti(new_form_magazzino.Combodipendenti)


        'new_form_magazzino.Show()



    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs)
        leggi_ini_server()
        MsgBox(Mid(Now, 12, 8))


    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        leggi_ini_computer()


    End Sub


    Sub mostra_dashboard()

        Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa)

        Mostra.lavorazioni_su_commessa()
        Mostra.lavorazioni_per_reparto()
        Mostra.Lavorazioni_aperte_mostra()
        Mostra.SITUAZIONE_MAGAZZINO()
        Mostra.Gantt()
        Mostra.CHART_AVANZAMENTO()
        Mostra.tickets()
        Mostra.Collaudo()
        Mostra.leggi_chi_è_responsabile()
        Mostra.DataGridView_materiale_mancante.Hide()
        Mostra.GroupBox13.Visible = False
        Mostra.GroupBox17.Visible = False
        Mostra.Label_descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Descrizione_commessa
        Mostra.Label_ordine_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).ordine_cliente_commessa
        Mostra.Label_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_commessa
        Mostra.Label_cliente_finale.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_finale_commessa
        Mostra.Button_commessa.Text = commessa
        Mostra.Label_consegna.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Consegna_commessa
        Mostra.Label3.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Giorni_alla_consegna
        Mostra.Label_destinazione.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).destinazione

        'Mostra.Label_consegna.Text = Consegna_commessa


    End Sub

    Private Sub Count_Tick(sender As Object, e As EventArgs) Handles Count.Tick
        Mousepos.Text = MousePosition.X & MousePosition.Y

    End Sub


    Private Sub Mousepos_TextChanged(sender As Object, e As EventArgs) Handles Mousepos.TextChanged
        Timer2.Stop()
        Timer2.Start()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button_dashboard.Click
        leggi_ini_computer()
        Mostra.Owner = Me
        Mostra.Hide()
        mostra_dashboard()
        Mostra.Show()

    End Sub



    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Form_gestione_utente.Show()






    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        CQ.Show()

    End Sub



    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Pianificazione_Tickets.Show()
        Pianificazione_Tickets.CODICE_REPARTO = trova_Dettagli_dipendente(ID_SALVATO).codice_reparto
        Pianificazione_Tickets.inizializzazione_form()

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Scheda_commessa_Pianificazione.inizializzazione = True
        Scheda_commessa_Pianificazione.Show()
        Scheda_commessa_Pianificazione.BringToFront()
        Scheda_commessa_Pianificazione.WindowState = FormWindowState.Maximized

        Scheda_commessa_Pianificazione.carica_commesse(Scheda_commessa_Pianificazione.DataGridView, Scheda_commessa_Pianificazione.TextBox1.Text, Scheda_commessa_Pianificazione.TextBox2.Text, Scheda_commessa_Pianificazione.filtro_cliente_f, Scheda_commessa_Pianificazione.filtro_n_progetto, Scheda_commessa_Pianificazione.filtro_nome_progetto_commessa, Scheda_commessa_Pianificazione.TextBox16.Text.ToUpper, Scheda_commessa_Pianificazione.filtro_desc_sup, "", "")
        Scheda_commessa_Pianificazione.inizializzazione = False
    End Sub






    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Form_Cambia_Reparto.Show()

    End Sub



    Private Sub Button11_Click(sender As Object, e As EventArgs)
        Distinta_base_form.Show()
    End Sub

    Private Sub Cmd_Consumabili_Click(sender As Object, e As EventArgs) Handles Cmd_Consumabili.Click
        Form_Richiesta_Materiale.Show()

        Form_Richiesta_Materiale.Txt_Commessa.Text = ""
        Form_Richiesta_Materiale.TXT_ODP.Text = ""
        Form_Richiesta_Materiale.Home_Lista()

    End Sub



    Private Sub Button11_Click_1(sender As Object, e As EventArgs) Handles Button11.Click
        Funzioni_utili.Show()

    End Sub




    Sub test_database_internet()


        Dim Cnn3 As New SqlConnection
        sap_tirelli = "Data Source=sql11.freesqldatabase.com;Initial Catalog=sql11500881;Persist Security Info=True;User ID=sql11500881;Password=XXD3DNYzFD"
        Cnn3.ConnectionString = sap_tirelli
        Cnn3.Open()

        Dim CMD_sap_tirelli_3 As New SqlCommand

        CMD_sap_tirelli_3.Connection = Cnn3


        CMD_sap_tirelli_3.CommandText = "insert into opor (DocEntry) values (123)"

        CMD_sap_tirelli_3.ExecuteNonQuery()
        Cnn3.Close()


    End Sub





    Private Sub Button14_Click(sender As Object, e As EventArgs)
        Form_cambia_dipendente.Show()
    End Sub








    Private Sub Button16_Click(sender As Object, e As EventArgs)
        Form_Cambia_Reparto.Show()
    End Sub





    Sub esegui_stored_procedure(par_object_type As String, par_transaction_type As String, par_num_of_cols_in_key As Integer, par_list_of_key_cols_tab_del As String, par_list_of_cols As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = sap_tirelli
        Cnn.Open()

        Dim CMD_sap_tirelli As New SqlCommand
        Dim cmd_sap_tirelli_reader As SqlDataReader
        CMD_sap_tirelli.Connection = Cnn




        CMD_sap_tirelli.Connection = Cnn
        CMD_sap_tirelli.CommandText = "EXECUTE SBO_SP_TransactionNotification @object_type='" & par_object_type & "', @transaction_type='" & par_transaction_type & "',@num_of_cols_in_key=" & par_num_of_cols_in_key & " ,@list_of_key_cols_tab_del='" & par_list_of_key_cols_tab_del & "',@list_of_cols_val_tab_del ='" & par_list_of_cols & "'"
        cmd_sap_tirelli_reader = CMD_sap_tirelli.ExecuteReader

        If cmd_sap_tirelli_reader.Read() Then

            Dim errore As Integer = cmd_sap_tirelli_reader("@error")
            Dim error_message As String = cmd_sap_tirelli_reader("@error_message")

            MsgBox(errore)
            MsgBox(error_message)
        End If


        cmd_sap_tirelli_reader.Close()
        cnn.Close()


    End Sub

    Private Sub Button17_Click_1(sender As Object, e As EventArgs)
        Form_Cambia_Reparto.Show()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            sap_tirelli = sap_4life
            azienda = "4LIFE"
            colore_sfondo = Color.YellowGreen
            logo_azienda = "\\tirfs01\00-Tirelli 4.0\Immagini generiche\Logo 4 Life.jpg"




        End If
        PictureBox1.Image = Image.FromFile(logo_azienda)
    End Sub


    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        Disambiguazione_payments.Show()
    End Sub



    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            sap_tirelli = sap_prova
            azienda = "PROVASRL"
            colore_sfondo = Color.Aqua
            logo_azienda = "\\tirfs01\00-Tirelli 4.0\Immagini generiche\Prova.jpg"

        End If
        PictureBox1.Image = Image.FromFile(logo_azienda)
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            'sap_tirelli = Pianificazione.SAP_TIRELLI
            azienda = "Tirelli"
            If ERP_provenienza = "SAP" Then


                colore_sfondo = Color.PowderBlue
            Else
                colore_sfondo = Color.Aquamarine
            End If

            Try
                logo_azienda = "\\tirfs01\00-Tirelli 4.0\Immagini generiche\TIRELLI_trasp_nopayoff_360x250.png"
            Catch ex As Exception

            End Try


        End If
        Try
            PictureBox1.Image = Image.FromFile(logo_azienda)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button16_Click_1(sender As Object, e As EventArgs) Handles Button16.Click
        Help_desk_disambiguazione.Show()
    End Sub


    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked = True Then

            Centro_di_costo = "KTF01"

        End If

    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then

            Centro_di_costo = "TIR01"

        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs)
        Dim stopwatch As New Stopwatch()
        stopwatch.Start()

        Dim filePath As String = "\\tirfs01\12-Engineering\Mechanical\Disegni Meccanici\pdf-dxf\pdf\000077.pdf"

        ' Verifica se il file esiste prima di tentare di aprirlo
        If System.IO.File.Exists(filePath) Then
            Try
                Process.Start(filePath)
            Catch ex As Exception
                MessageBox.Show("Errore nell'apertura del file: " & ex.Message)
            End Try
        Else
            MessageBox.Show("Il file specificato non esiste.")
        End If

        stopwatch.Stop()
        MessageBox.Show("Tempo impiegato per aprire il PDF: " & stopwatch.ElapsedMilliseconds & " millisecondi")
    End Sub

    Private Sub Button14_Click_1(sender As Object, e As EventArgs)
        Dim stopwatch As New Stopwatch()
        stopwatch.Start()

        Dim filePath As String = "W:\Tecnico\000077.pdf"

        ' Verifica se il file esiste prima di tentare di aprirlo
        If System.IO.File.Exists(filePath) Then
            Try
                Process.Start(filePath)
            Catch ex As Exception
                MessageBox.Show("Errore nell'apertura del file: " & ex.Message)
            End Try
        Else
            MessageBox.Show("Il file specificato non esiste.")
        End If

        stopwatch.Stop()
        MessageBox.Show("Tempo impiegato per aprire il PDF: " & stopwatch.ElapsedMilliseconds & " millisecondi")

    End Sub

    Private Sub Button13_Click_1(sender As Object, e As EventArgs) Handles Button13.Click
        Dim res As DialogResult = ConfigPassword.ShowDialog()
        If res = DialogResult.OK Then

            Configurazioni.CaricaConfigurazioni()
            Dim result As DialogResult = Configurazioni.ShowDialog()

            If result Then
                leggi_ini_server()
            End If

        End If
    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        If RadioButton6.Checked = True Then

            Centro_di_costo = "BRB01"

        End If
    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        If RadioButton7.Checked = True Then

            Centro_di_costo = "OH01"

        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Presence.Show()
    End Sub

    Private Sub Button14_Click_2(sender As Object, e As EventArgs)
        Presence.invia_report_ordini_in_garanzia(percorso_server & "00-Tirelli 4.0\Report\Ordini in garanzia-recall.xlsx")
        Presence.Show()
        MsgBox("Report inviato")

    End Sub


    ' Importiamo la funzione SetCursorPos dalla libreria user32.dll
    <DllImport("user32.dll")>
    Private Shared Function SetCursorPos(x As Integer, y As Integer) As Boolean
    End Function



    ' Funzione per rilasciare gli oggetti COM
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            If Not obj Is Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub






    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Discriminante_Totem_PC.Show()
    End Sub








    ' Importare la funzione GetAsyncKeyState per rilevare la pressione del tasto Esc
    <DllImport("user32.dll")>
        Private Shared Function GetAsyncKeyState(ByVal vKey As Integer) As Short
        End Function

    Private Sub Button19_Click(sender As Object, e As EventArgs)
        Form_KPI.Show()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs)
        Presence.Show()
    End Sub

    Private Sub Button14_Click_3(sender As Object, e As EventArgs)
        Form_configuratore_vendita.Show()
    End Sub

    Private Sub Button14_Click_4(sender As Object, e As EventArgs) Handles Button14.Click
        Form_layout_CAP_1.Show()
    End Sub

    Private Sub Button17_Click_2(sender As Object, e As EventArgs) Handles Button17.Click
        Process.Start("https://jpm.tirelli.net/jpm-share/?path=dclOlnCfg")
    End Sub

    Private Sub Button19_Click_1(sender As Object, e As EventArgs) Handles Button19.Click
        test_insert()
        MsgBox("Fatto")
    End Sub

    Sub test_select()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT TOP (1000) [CODE]
      ,[COD_FERR]
      ,[DES_CODE]
      ,[UBI_CODE]
      ,[LNG_CODE]
      ,[DISEGNO]
      ,[STAT_CODE]
      ,[COD_TRAT]
      ,[DESC_TRAT]
      ,[GRUP_ART]
      ,[DESC_GRP]
      ,[COSTO_STD]
      ,[QTA_SAFE]
      ,[QTA_MIOR]
      ,[SOGG_COL]
      ,[PROD_FOR]
      ,[CODAR_FOR]
      ,[UBI_SEC]
      ,[CODE_BRB]
      ,[UMIS]
      ,[COD_FOR]
      ,[DESC_FOR]
      ,[TIPO_PARTE]
      ,[GEST_COMM]
      ,[MOTIV_MAG]
      ,[DATA_STOCK]
      ,[CHECK_DB]
  FROM [Tirelli_40].[dbo].[JGAL_Articoli]

where code='C00001'
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() Then

            MsgBox(cmd_SAP_reader("DES_CODE"))
        End If

        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub test_insert()


        Dim Cnn3 As New SqlConnection

        Cnn3.ConnectionString = sap_tirelli
        Cnn3.Open()

        Dim CMD_sap_tirelli_3 As New SqlCommand

        CMD_sap_tirelli_3.Connection = Cnn3


        CMD_sap_tirelli_3.CommandText = "INSERT INTO PTI90DAT.YPCMOV0F (PROFYP, DT01YP, CDDTYP, DTOPYP,      
NROPYP, RIGAYP, DTMOYP, ORPRYP, CDFAYP,                 CAMOYP,     
TEMOYP) VALUES('TIR40', 20260114, '01', 20260114, 400, 2, 20260113, 
25001238, '010' ,            '01', 0,25)"

        CMD_sap_tirelli_3.ExecuteNonQuery()
        Cnn3.Close()


    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Presence.azzera_disegni_mancanti()
        Presence.check_presenza_disegno()
        MsgBox("FINE")
        Beep()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        'Form_IA.Show()
        'Form_gemini.Show()
        Form_stampe.Show()
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Process.Start(percorso_server & "00-Tirelli 4.0\File\Vari\Guida_JPM.pdf")
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        Dashboard_MU_New.inserisci_lavorazione_a_Galileo_macchina("030", 123456, "0.25")
        MsgBox("Manodopera inserita con successo")
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        Process.Start("\\tirfs01\00-Tirelli 4.0\Guida Galileo.docx")
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        Form_Movimenti_magazzino.Show()
    End Sub
End Class