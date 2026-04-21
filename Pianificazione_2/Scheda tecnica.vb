Imports System.Data.SqlClient
Imports System.Linq
Imports System.IO
Imports System.Reflection.Emit
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Threading.Tasks
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TreeView






Public Class Scheda_tecnica
    Public bp_code As String
    Public final_bp_code As String
    Public bp_code_galileo As String
    Public final_bp_code_galileo As String

    Public codice_bp_campione As String
    Public cartella_macchina As String
    Public codice_commessa As String
    Public Elenco_dipendenti(1000) As String
    Public Elenco_reparti(1000) As String
    Public id_utente As Integer
    Public numero_ultima_revisione As Integer = 0
    Private N_rev_visualizza As Integer
    Public numero_combinazioni As Integer = 0

    Private cbTipologiaMacchina As ComboBox
    Private cbModelloMacchina As ComboBox

    Private Controllo_generico(50) As Byte
    Private Configurazione_macchina(50) As Byte ' Private binary array of 50 bytes
    Private soffiatura_accessorio(50) As Byte
    Private riempimento_accessorio(50) As Byte

    Private tappatura_dettagli(50) As Byte
    Private tappatura_optional(50) As Byte
    Private tappatura_accessori(50) As Byte
    Private Etichettatura_accessori(50) As Byte
    Private Etichettatura_piattelli(50) As Byte
    Private num_collaudati As Integer

    Private Sub Scheda_tecnica_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TabControl1.DrawMode = TabDrawMode.OwnerDrawFixed
        '    Me.BackColor = Homepage.colore_sfondo
        AggiungiCampiTipologiaModello()
        ApplicaStile()
        Inserimento_dipendenti()
        CaricaTipologieModelli()
        For Each page As TabPage In TabControl1.TabPages
            For Each ctrl As Control In page.Controls
                AggiungiGestoreEvento(ctrl)
            Next
        Next
    End Sub

    Private Sub AggiungiCampiTipologiaModello()
        Dim navy As Color = Color.FromArgb(22, 45, 84)

        Dim pnlMacchina As New Panel
        pnlMacchina.Height = 64
        pnlMacchina.Dock = DockStyle.Top
        pnlMacchina.BackColor = Color.FromArgb(230, 240, 255)
        pnlMacchina.Padding = New Padding(2)

        Dim gbTip As New GroupBox
        gbTip.Text = "Tipologia macchina"
        gbTip.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        gbTip.ForeColor = navy
        gbTip.Left = 0
        gbTip.Top = 0
        gbTip.Width = 200
        gbTip.Height = 60
        cbTipologiaMacchina = New ComboBox
        cbTipologiaMacchina.Dock = DockStyle.Fill
        cbTipologiaMacchina.Font = New Font("Segoe UI", 9)
        cbTipologiaMacchina.Name = "cbTipologiaMacchina"
        gbTip.Controls.Add(cbTipologiaMacchina)

        Dim gbMod As New GroupBox
        gbMod.Text = "Modello macchina"
        gbMod.Font = New Font("Segoe UI", 8.5F, FontStyle.Bold)
        gbMod.ForeColor = navy
        gbMod.Left = 200
        gbMod.Top = 0
        gbMod.Width = 193
        gbMod.Height = 60
        cbModelloMacchina = New ComboBox
        cbModelloMacchina.Dock = DockStyle.Fill
        cbModelloMacchina.Font = New Font("Segoe UI", 9)
        cbModelloMacchina.Name = "cbModelloMacchina"
        gbMod.Controls.Add(cbModelloMacchina)

        Dim btnGestisci As New Button
        btnGestisci.Text = "..."
        btnGestisci.Font = New Font("Segoe UI", 7.5F)
        btnGestisci.Left = 393
        btnGestisci.Top = 0
        btnGestisci.Width = 24
        btnGestisci.Height = 60
        btnGestisci.FlatStyle = FlatStyle.Flat
        btnGestisci.BackColor = navy
        btnGestisci.ForeColor = Color.White
        btnGestisci.FlatAppearance.BorderSize = 0
        AddHandler btnGestisci.Click, AddressOf BtnGestisciTabelleST_Click

        pnlMacchina.Controls.Add(gbTip)
        pnlMacchina.Controls.Add(gbMod)
        pnlMacchina.Controls.Add(btnGestisci)

        Panel2.Controls.Add(pnlMacchina)
    End Sub

    Private Sub CaricaTipologieModelli()
        If cbTipologiaMacchina Is Nothing OrElse cbModelloMacchina Is Nothing Then Return
        Dim savedTip = cbTipologiaMacchina.Text
        Dim savedMod = cbModelloMacchina.Text
        cbTipologiaMacchina.Items.Clear()
        cbModelloMacchina.Items.Clear()
        Try
            Using Cnn As New SqlConnection(Homepage.sap_tirelli)
                Cnn.Open()
                Using CMD As New SqlCommand("SELECT valore FROM [TIRELLI_40].[dbo].[ST_lookup_tipologia_macchina] ORDER BY valore", Cnn)
                    Using r = CMD.ExecuteReader()
                        While r.Read()
                            cbTipologiaMacchina.Items.Add(r("valore").ToString())
                        End While
                    End Using
                End Using
                Using CMD As New SqlCommand("SELECT valore FROM [TIRELLI_40].[dbo].[ST_lookup_modello_macchina] ORDER BY valore", Cnn)
                    Using r = CMD.ExecuteReader()
                        While r.Read()
                            cbModelloMacchina.Items.Add(r("valore").ToString())
                        End While
                    End Using
                End Using
            End Using
        Catch
        End Try
        cbTipologiaMacchina.Text = savedTip
        cbModelloMacchina.Text = savedMod
    End Sub

    Private Sub BtnGestisciTabelleST_Click(sender As Object, e As EventArgs)
        Using f As New Scheda_Tecnica_Tabelle
            f.ShowDialog()
        End Using
        CaricaTipologieModelli()
    End Sub

    Private Sub ApplicaStile()
        Dim navy As Color = Color.FromArgb(22, 45, 84)
        Dim navyHover As Color = Color.FromArgb(30, 63, 122)
        Dim navyDark As Color = Color.FromArgb(10, 26, 55)
        Dim fontUI As String = "Segoe UI"

        Me.Font = New Font(fontUI, 9)

        ' Pulsanti principali: stile flat coerente con Homepage
        For Each btn As Button In Me.Controls.OfType(Of Button)()
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 1
            btn.Font = New Font(fontUI, 9, FontStyle.Regular)
        Next

        ' Pulsanti di navigazione/azione: sfondo navy
        Dim btnNavy() As Button = {Button3}  ' chiudi
        For Each btn As Button In btnNavy
            If btn IsNot Nothing Then
                btn.BackColor = navy
                btn.ForeColor = Color.White
                btn.FlatAppearance.BorderSize = 0
                btn.FlatAppearance.MouseOverBackColor = navyHover
                btn.FlatAppearance.MouseDownBackColor = navyDark
            End If
        Next

        ' Label intestazione commessa
        For Each lbl As System.Windows.Forms.Label In New System.Windows.Forms.Label() {Label1, Label2}
            lbl.Font = New Font(fontUI, 10, FontStyle.Bold)
            lbl.ForeColor = navy
        Next

        ' Label secondarie
        For Each lbl As System.Windows.Forms.Label In New System.Windows.Forms.Label() {Label3, Label4, Label5, Label6}
            lbl.Font = New Font(fontUI, 9, FontStyle.Regular)
            lbl.ForeColor = Color.FromArgb(40, 60, 90)
        Next

        ' DataGridView revisioni
        DataGridView_revisione.BorderStyle = BorderStyle.None
        DataGridView_revisione.BackgroundColor = Color.White
        DataGridView_revisione.GridColor = Color.FromArgb(210, 220, 235)
        DataGridView_revisione.ColumnHeadersDefaultCellStyle.BackColor = navy
        DataGridView_revisione.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        DataGridView_revisione.ColumnHeadersDefaultCellStyle.Font = New Font(fontUI, 8.5F, FontStyle.Bold)
        DataGridView_revisione.EnableHeadersVisualStyles = False
        DataGridView_revisione.RowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 225, 245)
        DataGridView_revisione.RowsDefaultCellStyle.SelectionForeColor = navy
    End Sub

    Private Sub AggiungiGestoreEvento(ctrl As Control)
        If TypeOf ctrl Is TextBox OrElse TypeOf ctrl Is RichTextBox OrElse TypeOf ctrl Is ComboBox Then
            AddHandler ctrl.TextChanged, AddressOf EvidenziaInTempoReale
        End If

        For Each child As Control In ctrl.Controls
            AggiungiGestoreEvento(child)
        Next
    End Sub
    'draaaaaaaa

    Private Sub EvidenziaInTempoReale(sender As Object, e As EventArgs)
        Dim ctrl As Control = CType(sender, Control)

        ' Evidenzia solo se compilato
        If ctrl.Text.Trim() = "" Or ctrl.Text.Trim() = "0" Then
            ctrl.BackColor = Color.White
            ctrl.Font = New Font(ctrl.Font, FontStyle.Regular)
        Else
            ctrl.BackColor = Color.LightYellow
            ctrl.Font = New Font(ctrl.Font, FontStyle.Regular)
        End If

        ' Ridisegna la tab contenente il controllo
        Dim tp As TabPage = FindParentTabPage(ctrl)
        If tp IsNot Nothing Then
            TabControl1.Invalidate() ' Forza il redraw della tab
        End If
    End Sub

    Private Function FindParentTabPage(ctrl As Control) As TabPage
        Dim p As Control = ctrl
        While p IsNot Nothing
            If TypeOf p Is TabPage Then Return CType(p, TabPage)
            p = p.Parent
        End While
        Return Nothing
    End Function

    Private Sub TabControl1_DrawItem(sender As Object, e As DrawItemEventArgs) Handles TabControl1.DrawItem
        Dim tp As TabPage = TabControl1.TabPages(e.Index)
        Dim text As String = tp.Text
        Dim font As Font = TabControl1.Font
        Dim fore As Brush = Brushes.Black

        ' Tab da evidenziare con sfondo colorato
        Dim tabDaEvidenziare As String() = {"Documentazione", "Caratteristiche_tecniche", "Caratteristiche_EL", "Certificazioni", "Formati"}
        Dim tabCondizionali As String() = {"Soffiatura", "Riempimento", "Tappatura", "Termosaldatrice", "Etichettatrice"}

        ' Sfondo diverso se la tab è da evidenziare
        If tabDaEvidenziare.Contains(tp.Name) Or (tabCondizionali.Contains(tp.Text) AndAlso TabPageHaEvidenziati(tp)) Then
            e.Graphics.FillRectangle(Brushes.Bisque, e.Bounds) ' colore di sfondo a scelta
        Else
            e.Graphics.FillRectangle(SystemBrushes.Control, e.Bounds)
        End If

        ' Misura il testo
        Dim size As SizeF = e.Graphics.MeasureString(text, font)

        ' Posizione centrata verticalmente, piccolo padding a sinistra
        Dim x As Integer = e.Bounds.Left + 2
        Dim y As Integer = e.Bounds.Top + (e.Bounds.Height - size.Height) / 2

        ' Disegna il testo
        e.Graphics.DrawString(text, font, fore, x, y)
    End Sub

    Private Function TabPageHaEvidenziati(tp As TabPage) As Boolean
        For Each ctrl As Control In tp.Controls
            If ctrl.BackColor = Color.LightYellow Then Return True
            If ctrl.HasChildren Then
                If TabPageHaEvidenziatiInContenitore(ctrl) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function TabPageHaEvidenziatiInContenitore(parent As Control) As Boolean
        For Each c As Control In parent.Controls
            If c.BackColor = Color.LightYellow Then Return True
            If c.HasChildren Then
                If TabPageHaEvidenziatiInContenitore(c) Then Return True
            End If
        Next
        Return False
    End Function


    Sub riempi_scheda_tecnica(par_codice_commessa As String, par_n_rev As Integer)

        Using Cnn1 As New SqlConnection(Homepage.sap_tirelli)
            Cnn1.Open()

            Using CMD_SAP_2 As New SqlCommand()
                Dim cmd_SAP_reader_2 As SqlDataReader


                CMD_SAP_2.Connection = Cnn1
                CMD_SAP_2.CommandText = "SELECT t0.[ID]
      ,t0.[Commessa]
      ,t0.[Rev]
,t0.note
      ,t0.[Velocita]
      ,t0.[Tipo_nastro]
      ,t0.[Tipo_nastro_commento]
      ,t0.[Altezza_lavoro]
      ,t0.[Altezza_lavoro_commento]
      ,t0.[Scarto]
      ,t0.[Nastro_Scarto]
      ,t0.[Pulsantiera]
      ,t0.[Pulsantiera_commento]
      ,t0.[Allacciamenti_elettrici]
      ,t0.[Allacciamenti_elettrici_Commento]
      ,t0.[Allacciamenti_pneumatici]
      ,t0.[Allacciamenti_pneumatici_commento]
      ,t0.[Marca_motiri]
      ,t0.[Industria_4_0]
      ,t0.[Teleassistenza]
, coalesce(t0.[Marca_plc],'') as 'Marca_plc'
      ,t0.[Controllo_presenza_tappo]
      ,t0.[Controllo_mal_tappato]
      ,t0.[Controllo_presenza_pescante]
      ,t0.[Controllo_presenza_oggetto]
      ,t0.[telecamera_e_pannello]
      ,t0.[Etichetta_trasparente]
      ,t0.[Illuminatore]
      ,t0.[Ambiente_standard]
      ,t0.[ambiente_standard_prod_inf]
      ,t0.[Ambiente_installazione_Atex_1]
      ,t0.[Ambiente_installazione_Atex_2]
      ,t0.[Hazloc]
      ,t0.[Documenti_aggiuntivi]
      ,t0.[Documenti_aggiuntivi_commento]
      ,t0.[Certificazioni_materiali]
      ,t0.[Lingua_manuale]
      ,t0.[Lingua_manuale_commento]
      ,t0.[Layout]
      ,t0.[Lista_ricambi_consigliata]
      ,t0.[Lista_ricambi_dettagli]
      ,t0.[Lista_ricambi_importo]
      ,t0.[accessorio_tipo_macchina]
      ,t0.[accessorio_sottotipo_macchina]
      ,t0.[accessorio_diametro]
      ,t0.[soffiatura_tipo_trattamento]
      ,t0.[soffiatura_ugello]
      ,t0.[soffiatura_accessorio_lavaggio]
      ,t0.[soffiatura_accessorio_copertura_totale]
      ,t0.[soffiatura_accessorio_aspirazione_polveri]
      ,t0.[soffiatura_accessorio_ionizzazione]
      ,t0.[soffiatura_accessorio_commento]
,coalesce(t0.Soffiatura_numero_ugelli,0) as 'Soffiatura_numero_ugelli'
      ,t0.[Riempimento_sottotipo_macchina]
      ,t0.[Riempimento_N_rubinetti]
      ,t0.[Riempimento_tipologia]
      ,t0.[Riempimento_tipologia_commento]
      ,t0.[Riempimento_carrello_riempimento]
      ,t0.[Riempimento_carrello_commento]
      ,t0.[Riempimento_carrello_lavaggio]
      ,t0.[Riempimento_carrello_lavaggio_commento]
      ,t0.[Riempimento_movimentazione_rubinetti]
      ,t0.[Riempimento_accessorio_azoto]
      ,t0.[Riempimento_accessorio_pompa_carico_prodotto]
      ,t0.[Riempimento_accessorio_pompa_carico_lavaggio]
      ,t0.[Riempimento_accessorio_serbatoio_esterno]
      ,t0.[Riempimento_accessorio_soffiatura]
      ,t0.[Riempimento_accessorio_circuito_riscaldamento]
      ,t0.[Riempimento_accessorio_raccogli_goccia]
      ,t0.[Riempimento_accessorio_commento]
,coalesce(t0.[Riempimento_salita_discesa],'') as 'Riempimento_salita_discesa'
      ,t0.[Tappatura_Tipo_Macchina]
      ,t0.[Tappatura_Sottotipo_Macchina]
,coalesce(t0.Tappatura_fornitura_torretta,'') as 'Tappatura_fornitura_torretta'
      ,t0.[Tappatura_numero_teste]
      ,t0.[Tappatura_tappo_trattato_trigger]
      ,t0.[Tappatura_tappo_trattato_pompetta]
      ,t0.[Tappatura_chiusura_trattata_vite]
      ,t0.[Tappatura_chiusura_trattata_pressione]
      ,t0.[Tappatura_tappo_trattato]
      ,t0.[Tappatura_sottotappo_trattato]
      ,t0.[Tappatura_pompetta_trattata]
      ,t0.[Tappatura_stiratura_pompetta_elettronica]
      ,t0.[Tappatura_stiratura_pompetta_pneumatica]
      ,t0.[Tappatura_optional_pompetta_Caricamento_automatica]
      ,t0.[Tappatura_optional_pompetta_centratori]
      ,t0.[Tappatura_optional_pompetta_testa_bordatrice]
      ,t0.[Tappatura_optional_pompetta_testa_trigger]
      ,t0.[Tappatura_optional_pompetta_blocco_ingresso]
      ,t0.[Tappatura_optional_pompetta_antirotazione]
      ,t0.[Tappatura_optional_pompetta_pressetta]
      ,t0.[Tappatura_optional_pompetta_sistema_termosaldatura]
      ,t0.[Tappatura_optional_pompetta_stella]
      ,t0.[Tappatura_optional_pompetta_a_inseguimento]
      ,t0.[Termosaldatura_sottotipo_macchina]
      ,t0.[Termosaldatura_N_teste]
      ,t0.[Termosaldatura_tipologia_sigillatura]
      ,t0.[Termosaldatura_azoto]
      ,t0.[Termosladatura_Passo_termosaldatrice]
      ,t0.[Etichettatura_Sottotipo_macchina]
      ,t0.[Etichettatura_Nome]
      
      ,t0.[Etichettatura_Tipo_testata]
      ,t0.[Etichettatura_tipo_testata_commento]
      ,t0.[Etichettatura_accessorio_stiratura_3_rulli]
      ,t0.[Etichettatura_accessorio_centratore_3_rulli]
      ,t0.[Etichettatura_accessorio_blocco_ingresso]
      ,t0.[Etichettatura_accessorio_centratore_orbitale]
      ,t0.[Etichettatura_accessorio_Divisore_rotante]
      ,t0.[Etichettatura_accessorio_Nastrino_superiore]
      ,t0.[Etichettatura_accessorio_Stiratura_tondi_con_contrasto]
      ,t0.[Etichettatura_accessorio_timbratore_inkjet_nastrino_inkjet]
      ,t0.[Etichettatura_accessorio_timbratore_laser]
      ,t0.[Etichettatura_accessorio_Lettore_fine_bobina]
      ,t0.[Etichettatura_accessorio_gruppo_stampa_barban]
      ,t0.[Etichettatura_accessorio_commento]
,coALESCE(t0.Etichettatura_regolazione_giostra,'') AS 'Etichettatura_regolazione_giostra'
   ,t0.[Etichettatura_diametro_primitivo]
      ,t0.[Etichettatura_Numero_piattelli]
      ,t0.[Etichettatura_senso_rotazione]


,t0.[Etichettatura_coclea]
,	t0.[Etichettatura_piattello_gomma_vulcanizzata]
,	t0.[Etichettatura_piattello_gomma_sagoma]
	,t0.[Etichettatura_piattello_espulsore_molla]
	,t0.[Etichettatura_piattello_perno_espulsore] 
	,t0.[Etichettatura_piattello_orientamento_meccanico] 
	,t0.[Etichettatura_piattello_orientamento_elettronico]
,coalesce(A.N,0) as 'Numero_gruppi'


 ,coalesce(t0.[light_washable],'') as 'light_washable'
           ,coalesce(t0.[alleggerita],'') as 'alleggerita'
           ,t0.[family_feeling]
           ,t0.[protezioni_porte]
           ,t0.[Marca_pannello_operatore]
           ,t0.[Dimensioni_pannello_operatore]
,COALESCE(t0.[serbatoio],'') AS 'SERBATOIO'
           ,t0.[Accessori_serbatoio]
           ,t0.[Lavaggio]
           ,t0.[Ingresso_prodotto]
           ,t0.[Pompa_Prodotto]
           ,t0.[CIP]
           ,t0.[Circuito_riempimento]
           ,t0.[Livello_lavaggio]
           ,t0.[Guarnizioni]
,t0.[Tappatura_accessori_presenza_tappo]
           ,t0.[Tappatura_tappo_maltappato]
           ,t0.[Tappatura_presenza_prescante]
           ,t0.[Tappatura_presenza_oggetto]
           ,t0.[Tappatura_termosigillatura]
           ,t0.[Tappatura_fornitura_canale_tappi]
           ,t0.[Tappatura_CF_Canale_Tappi]
 ,coalesce(t0.[Tipo_catena],'') as 'Tipo_catena'
      ,coalesce(t0.[By_pass],'') as 'By_pass'
      ,coalesce(t0.[Condizionamento_quadro],'') as 'Condizionamento_quadro'
      ,coalesce(t0.[Generico_UTE],'') as 'Generico_UTE'
      ,coalesce(t0.[Colori_Attrezzature],'') as 'Colori_Attrezzature'
      ,coalesce(t0.[Etichettatura_tetto_antipolvere],'') as 'Etichettatura_tetto_antipolvere'
      ,coalesce(t0.[Etichettatura_Lamiera_forata],'') as 'Etichettatura_Lamiera_forata'
      ,coalesce(t0.[Etichettatura_Camma_piattello],'') as 'Etichettatura_Camma_piattello'


      ,coalesce(t0.[Etichettatura_altezza_albero_Centrale],'') as 'Etichettatura_altezza_albero_Centrale'
      ,coalesce(t0.[Etichettatura_tipo_testina],'') as 'Etichettatura_tipo_testina'
      ,coalesce(t0.[Etichettatura_tipo_molla],'') as 'Etichettatura_tipo_molla'
      ,coalesce(t0.[Etichettatura_pressurizzazione_bottiglia],'') as 'Etichettatura_pressurizzazione_bottiglia'
      ,coalesce(t0.[Etichettatura_lavaggio],'') as 'Etichettatura_lavaggio'
      ,coalesce(t0.[Etichettatura_stella_entrata],'') as 'Etichettatura_stella_entrata'
      ,coalesce(t0.[Etichettatura_stella_entrata_diametro],'') as 'Etichettatura_stella_entrata_diametro'
      ,coalesce(t0.[Etichettatura_stella_uscita],'') as 'Etichettatura_stella_uscita'
      ,coalesce(t0.[Etichettatura_stella_uscita_diametro],'') as 'Etichettatura_stella_uscita_diametro'
,coalesce(t0.[FAT_concordata],'') as 'FAT_concordata'
,coalesce(t1.stato,'') as 'Stato_scheda'
,coalesce(t0.[Tipologia_macchina],'') as 'Tipologia_macchina'
,coalesce(t0.[Modello_macchina],'') as 'Modello_macchina'





  FROM [TIRELLI_40].[dbo].[Scheda_Tecnica_valori] t0
left join 
( select sum (case when t0.id is null then 0 else 1 end ) as 'N', t0.commessa as 'Comm'
from [TIRELLI_40].[DBO].[BRB_Gruppi_etichettaggio] t0
where t0.commessa='" & par_codice_commessa & "' 
group by t0.commessa) A on A.comm=Commessa

left join [Tirelli_40].[dbo].[Scheda_tecnica_revisioni] t1 on t1.commessa='" & par_codice_commessa & "' and t1.numero='" & par_n_rev & "'
where t0.commessa='" & par_codice_commessa & "' and t0.rev ='" & par_n_rev & "'
"

                cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

                If cmd_SAP_reader_2.Read() Then
                    ComboBox70.Text = cmd_SAP_reader_2("Stato_scheda")
                    RichTextBox44.Text = cmd_SAP_reader_2("Note")
                    TextBox2.Text = cmd_SAP_reader_2("Velocita")
                    ComboBox4.Text = cmd_SAP_reader_2("Tipo_nastro")
                    TextBox13.Text = cmd_SAP_reader_2("Tipo_nastro_commento")
                    TextBox1.Text = cmd_SAP_reader_2("Altezza_lavoro")
                    ComboBox3.Text = cmd_SAP_reader_2("Altezza_lavoro_commento")
                    ComboBox1.Text = cmd_SAP_reader_2("Scarto")
                    ComboBox2.Text = cmd_SAP_reader_2("Nastro_Scarto")
                    ComboBox10.Text = cmd_SAP_reader_2("Pulsantiera")
                    TextBox3.Text = cmd_SAP_reader_2("Pulsantiera_commento")
                    ComboBox7.Text = cmd_SAP_reader_2("Allacciamenti_elettrici")
                    RichTextBox1.Text = cmd_SAP_reader_2("Allacciamenti_elettrici_Commento")
                    ComboBox8.Text = cmd_SAP_reader_2("Allacciamenti_pneumatici")
                    RichTextBox2.Text = cmd_SAP_reader_2("Allacciamenti_pneumatici_commento")
                    TextBox12.Text = cmd_SAP_reader_2("Marca_motiri")
                    ComboBox9.Text = cmd_SAP_reader_2("Industria_4_0")
                    ComboBox6.Text = cmd_SAP_reader_2("Teleassistenza")
                    ComboBox54.Text = cmd_SAP_reader_2("Marca_PLC")
                    CheckBox24.Checked = Convert.ToBoolean(cmd_SAP_reader_2("Controllo_presenza_tappo"))
                    CheckBox25.Checked = Convert.ToBoolean(cmd_SAP_reader_2("Controllo_mal_tappato"))
                    CheckBox26.Checked = Convert.ToBoolean(cmd_SAP_reader_2("Controllo_presenza_pescante"))
                    CheckBox27.Checked = Convert.ToBoolean(cmd_SAP_reader_2("Controllo_presenza_oggetto"))
                    CheckBox28.Checked = Convert.ToBoolean(cmd_SAP_reader_2("telecamera_e_pannello"))
                    CheckBox29.Checked = Convert.ToBoolean(cmd_SAP_reader_2("Etichetta_trasparente"))
                    CheckBox30.Checked = Convert.ToBoolean(cmd_SAP_reader_2("Illuminatore"))
                    CheckBox52.Checked = Convert.ToBoolean(cmd_SAP_reader_2("Ambiente_standard"))
                    CheckBox53.Checked = Convert.ToBoolean(cmd_SAP_reader_2("ambiente_standard_prod_inf"))
                    CheckBox54.Checked = Convert.ToBoolean(cmd_SAP_reader_2("Ambiente_installazione_Atex_1"))
                    CheckBox55.Checked = Convert.ToBoolean(cmd_SAP_reader_2("Ambiente_installazione_Atex_2"))
                    CheckBox56.Checked = Convert.ToBoolean(cmd_SAP_reader_2("Hazloc"))
                    CheckBox5.Checked = Convert.ToBoolean(cmd_SAP_reader_2("light_washable"))
                    CheckBox6.Checked = Convert.ToBoolean(cmd_SAP_reader_2("alleggerita"))
                    ComboBox19.Text = cmd_SAP_reader_2("Documenti_aggiuntivi")
                    TextBox10.Text = cmd_SAP_reader_2("Documenti_aggiuntivi_commento")
                    ComboBox5.Text = cmd_SAP_reader_2("Certificazioni_materiali")
                    ComboBox33.Text = cmd_SAP_reader_2("Lingua_manuale")
                    TextBox11.Text = cmd_SAP_reader_2("Lingua_manuale_commento")
                    TextBox8.Text = cmd_SAP_reader_2("Layout")

                    CheckBox50.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Lista_ricambi_consigliata”))

                    ComboBox36.Text = cmd_SAP_reader_2(“Lista_ricambi_dettagli”)
                    TextBox14.Text = cmd_SAP_reader_2(“Lista_ricambi_importo”)
                    ComboBox18.Text = cmd_SAP_reader_2(“accessorio_tipo_macchina”)
                    ComboBox20.Text = cmd_SAP_reader_2(“accessorio_sottotipo_macchina”)
                    ComboBox21.Text = cmd_SAP_reader_2(“accessorio_diametro”)
                    ComboBox38.Text = cmd_SAP_reader_2(“soffiatura_tipo_trattamento”)
                    ComboBox39.Text = cmd_SAP_reader_2(“soffiatura_ugello”)
                    ComboBox68.Text = cmd_SAP_reader_2(“Soffiatura_numero_ugelli”)


                    CheckBox51.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“soffiatura_accessorio_lavaggio”))
                    CheckBox57.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“soffiatura_accessorio_copertura_totale”))
                    CheckBox58.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“soffiatura_accessorio_aspirazione_polveri”))
                    CheckBox59.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“soffiatura_accessorio_ionizzazione”))
                    RichTextBox3.Text = cmd_SAP_reader_2(“soffiatura_accessorio_commento”)
                    ComboBox11.Text = cmd_SAP_reader_2(“Riempimento_sottotipo_macchina”)
                    ComboBox12.Text = cmd_SAP_reader_2(“Riempimento_N_rubinetti”)
                    ComboBox14.Text = cmd_SAP_reader_2(“Riempimento_tipologia”)
                    TextBox4.Text = cmd_SAP_reader_2(“Riempimento_tipologia_commento”)
                    ComboBox16.Text = cmd_SAP_reader_2(“Riempimento_carrello_riempimento”)
                    TextBox5.Text = cmd_SAP_reader_2(“Riempimento_carrello_commento”)
                    ComboBox17.Text = cmd_SAP_reader_2(“Riempimento_carrello_lavaggio”)
                    TextBox15.Text = cmd_SAP_reader_2(“Riempimento_carrello_lavaggio_commento”)
                    ComboBox22.Text = cmd_SAP_reader_2(“Riempimento_movimentazione_rubinetti”)
                    ComboBox67.Text = cmd_SAP_reader_2("Riempimento_salita_discesa")
                    CheckBox43.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Riempimento_accessorio_azoto”))
                    CheckBox44.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Riempimento_accessorio_pompa_carico_prodotto”))
                    CheckBox45.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Riempimento_accessorio_pompa_carico_lavaggio”))
                    CheckBox46.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Riempimento_accessorio_serbatoio_esterno”))
                    CheckBox47.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Riempimento_accessorio_soffiatura”))
                    CheckBox48.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Riempimento_accessorio_circuito_riscaldamento”))
                    CheckBox49.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Riempimento_accessorio_raccogli_goccia”))
                    TextBox6.Text = cmd_SAP_reader_2(“Riempimento_accessorio_commento”)
                    ComboBox23.Text = cmd_SAP_reader_2(“Tappatura_Tipo_Macchina”)
                    ComboBox24.Text = cmd_SAP_reader_2(“Tappatura_Sottotipo_Macchina”)
                    ComboBox69.Text = cmd_SAP_reader_2(“Tappatura_fornitura_torretta”)
                    ComboBox25.Text = cmd_SAP_reader_2(“Tappatura_numero_teste”)

                    CheckBox12.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_stiratura_pompetta_elettronica”))
                    CheckBox13.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_stiratura_pompetta_pneumatica”))
                    CheckBox14.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_optional_pompetta_Caricamento_automatica”))
                    CheckBox15.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_optional_pompetta_centratori”))
                    CheckBox16.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_optional_pompetta_testa_bordatrice”))
                    CheckBox17.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_optional_pompetta_testa_trigger”))
                    CheckBox18.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_optional_pompetta_blocco_ingresso”))
                    CheckBox19.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_optional_pompetta_antirotazione”))
                    CheckBox20.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_optional_pompetta_pressetta”))
                    CheckBox21.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_optional_pompetta_sistema_termosaldatura”))
                    CheckBox22.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_optional_pompetta_stella”))
                    CheckBox23.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_optional_pompetta_a_inseguimento”))
                    ComboBox30.Text = cmd_SAP_reader_2(“Termosaldatura_sottotipo_macchina”)
                    ComboBox31.Text = cmd_SAP_reader_2(“Termosaldatura_N_teste”)
                    ComboBox32.Text = cmd_SAP_reader_2(“Termosaldatura_tipologia_sigillatura”)
                    ComboBox37.Text = cmd_SAP_reader_2(“Termosaldatura_azoto”)
                    TextBox9.Text = cmd_SAP_reader_2(“Termosladatura_Passo_termosaldatrice”)
                    ComboBox26.Text = cmd_SAP_reader_2(“Etichettatura_Sottotipo_macchina”)
                    ComboBox27.Text = cmd_SAP_reader_2(“Etichettatura_Nome”)

                    ComboBox29.Text = cmd_SAP_reader_2(“Etichettatura_Tipo_testata”)
                    TextBox7.Text = cmd_SAP_reader_2(“Etichettatura_tipo_testata_commento”)
                    CheckBox39.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_accessorio_stiratura_3_rulli”))
                    CheckBox40.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_accessorio_centratore_3_rulli”))
                    CheckBox38.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_accessorio_blocco_ingresso”))
                    CheckBox37.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_accessorio_centratore_orbitale”))
                    CheckBox36.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_accessorio_Divisore_rotante”))
                    CheckBox35.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_accessorio_Nastrino_superiore”))
                    CheckBox34.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_accessorio_Stiratura_tondi_con_contrasto”))
                    CheckBox33.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_accessorio_timbratore_inkjet_nastrino_inkjet”))
                    CheckBox32.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_accessorio_timbratore_laser”))
                    CheckBox31.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_accessorio_Lettore_fine_bobina”))
                    CheckBox41.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_accessorio_gruppo_stampa_barban”))
                    RichTextBox4.Text = cmd_SAP_reader_2(“Etichettatura_accessorio_commento”)



                    TextBox18.Text = cmd_SAP_reader_2(“Etichettatura_diametro_primitivo”)
                    ComboBox50.Text = cmd_SAP_reader_2(“Etichettatura_Numero_piattelli”)
                    Label10.Text = cmd_SAP_reader_2(“Numero_gruppi”)
                    ComboBox51.Text = cmd_SAP_reader_2(“Etichettatura_senso_rotazione”)
                    ComboBox53.Text = cmd_SAP_reader_2(“Etichettatura_regolazione_giostra”)

                    ComboBox52.Text = cmd_SAP_reader_2(“Etichettatura_coclea”)
                    CheckBox1.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_piattello_gomma_vulcanizzata”))
                    CheckBox2.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_piattello_gomma_sagoma”))
                    CheckBox3.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_piattello_espulsore_molla”))
                    CheckBox4.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_piattello_perno_espulsore”))
                    CheckBox7.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_piattello_orientamento_meccanico”))
                    CheckBox60.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Etichettatura_piattello_orientamento_elettronico”))


                    CheckBox5.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“light_washable”))
                    CheckBox6.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“alleggerita”))
                    ComboBox13.Text = cmd_SAP_reader_2(“family_feeling”)
                    ComboBox15.Text = cmd_SAP_reader_2(“protezioni_porte”)
                    ComboBox34.Text = cmd_SAP_reader_2(“Marca_pannello_operatore”)
                    ComboBox15.Text = cmd_SAP_reader_2(“protezioni_porte”)
                    TextBox17.Text = cmd_SAP_reader_2(“Dimensioni_pannello_operatore”)


                    ComboBox35.Text = cmd_SAP_reader_2("serbatoio”)
                    ComboBox40.Text = cmd_SAP_reader_2(“Accessori_serbatoio”)
                    ComboBox41.Text = cmd_SAP_reader_2(“Lavaggio”)
                    ComboBox42.Text = cmd_SAP_reader_2(“Ingresso_prodotto”)
                    ComboBox43.Text = cmd_SAP_reader_2(“Pompa_Prodotto”)
                    ComboBox44.Text = cmd_SAP_reader_2(“CIP”)
                    ComboBox46.Text = cmd_SAP_reader_2(“Circuito_riempimento”)
                    ComboBox47.Text = cmd_SAP_reader_2(“Livello_lavaggio”)
                    ComboBox48.Text = cmd_SAP_reader_2(“Guarnizioni”)


                    CheckBox11.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_accessori_presenza_tappo”))
                    CheckBox42.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_tappo_maltappato”))
                    CheckBox10.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_presenza_prescante”))
                    CheckBox9.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_presenza_oggetto”))
                    CheckBox8.Checked = Convert.ToBoolean(cmd_SAP_reader_2(“Tappatura_termosigillatura”))

                    ComboBox49.Text = cmd_SAP_reader_2(“Tappatura_fornitura_canale_tappi”)
                    ComboBox45.Text = cmd_SAP_reader_2(“Tappatura_CF_Canale_Tappi”)

                    TextBox16.Text = cmd_SAP_reader_2(“Tipo_catena”)
                    ComboBox57.Text = cmd_SAP_reader_2(“By_pass”)
                    ComboBox58.Text = cmd_SAP_reader_2(“Condizionamento_quadro”)

                    If cmd_SAP_reader_2(“Generico_UTE”) <> "" Then
                        RichTextBox5.Text = cmd_SAP_reader_2(“Generico_UTE”)
                    End If

                    ComboBox66.Text = cmd_SAP_reader_2(“Colori_Attrezzature”)
                    ComboBox28.Text = cmd_SAP_reader_2(“Etichettatura_tetto_antipolvere”)
                    ComboBox55.Text = cmd_SAP_reader_2(“Etichettatura_Lamiera_forata”)
                    ComboBox56.Text = cmd_SAP_reader_2(“Etichettatura_Camma_piattello”)



                    ComboBox59.Text = cmd_SAP_reader_2(“Etichettatura_altezza_albero_Centrale”)
                    ComboBox60.Text = cmd_SAP_reader_2(“Etichettatura_tipo_testina”)
                    ComboBox61.Text = cmd_SAP_reader_2(“Etichettatura_tipo_molla”)
                    ComboBox62.Text = cmd_SAP_reader_2(“Etichettatura_pressurizzazione_bottiglia”)
                    ComboBox63.Text = cmd_SAP_reader_2(“Etichettatura_lavaggio”)
                    ComboBox64.Text = cmd_SAP_reader_2(“Etichettatura_stella_entrata”)
                    ComboBox65.Text = cmd_SAP_reader_2(“Etichettatura_stella_uscita”)
                    TextBox19.Text = cmd_SAP_reader_2(“Etichettatura_stella_entrata_diametro”)
                    TextBox20.Text = cmd_SAP_reader_2(“Etichettatura_stella_uscita_diametro”)

                    RichTextBox6.Text = cmd_SAP_reader_2(“FAT_concordata”)

                    If cbTipologiaMacchina IsNot Nothing Then cbTipologiaMacchina.Text = cmd_SAP_reader_2(“Tipologia_macchina”)
                    If cbModelloMacchina IsNot Nothing Then cbModelloMacchina.Text = cmd_SAP_reader_2(“Modello_macchina”)




                End If
                cmd_SAP_reader_2.Close()
            End Using ' CMD_SAP_2
        End Using ' Cnn1

    End Sub



    Public Async Function inizializza_scheda_tecnica(par_codice_commessa As String) As Task
        codice_commessa = par_codice_commessa

        compila_anagrafica(par_codice_commessa)

        ' Chiamata asincrona al modulo cardini
        Await crea_modulo_cardini(FlowLayoutPanel13, Label1.Text, Label2.Text, Homepage.JPM_TIRELLI)

        elenca_revisioni(par_codice_commessa)
        trova_ultima_revisione(par_codice_commessa)
        Label7.Text = numero_ultima_revisione

        ' Se mostra_file_async è Async, va Await
        '  If mostra_file_async() IsNot Nothing Then
        Await mostra_file_async(LinkLabel2.Text, TreeView1)
        ' End If

        riempi_scheda_tecnica(par_codice_commessa, numero_ultima_revisione)
        Progetto.trova_ultima_revisione_progetto(Button26.Text, Button12.Text)
        Progetto.riempi_scheda_tecnica_progetto(Button26.Text, Progetto.numero_ultima_revisione, "Scheda_tecnica")
        AggiornaStatoRicambi(par_codice_commessa, numero_ultima_revisione)
        Me.Refresh()
        ' Funzione Async termina qui implicitamentedatagr

        Dim nomeProgetto As String = trova_percorso_documenti(Button26.Text, "PROGETTO", Button12.Text)
        Dim nomeSottoCartella As String = "Azioni contenitive"
        Dim percorsoRelativoPerTreeView As String = nomeProgetto & "\" & nomeSottoCartella

        mostra_file_async_progetto(percorsoRelativoPerTreeView, TreeView3)

    End Function

    Sub compila_anagrafica(par_codice_Commessa As String)
        Using Cnn1 As New SqlConnection(Homepage.sap_tirelli)
            Cnn1.Open()

            Using CMD_SAP_2 As New SqlCommand()
                Dim cmd_SAP_reader_2 As SqlDataReader


                CMD_SAP_2.Connection = Cnn1
                If Homepage.ERP_provenienza = "SAP" Then
                    CMD_SAP_2.CommandText = "Select t10.itemcode, coalesce(t15.absentry, 0) as 'N_progetto', coalesce(t15.name,'') as 'Nome_progetto',coalesce(CONCAT(T16.LASTNAME,' ',T16.FIRSTNAME),'') AS 'PM', t13.itemname, case when t12.cardcode is null then '' else t12.cardcode end as 'Cardcode', case when t12.cardname is null then '' else t12.cardname end as 'Cardname', t12.docduedate,case when t12.u_destinazione is null then '' else t12.u_destinazione end as 'u_destinazione', case when t12.u_codicebp is null then '' else t12.u_codicebp end as 'codice_Cliente_finale', case when t14.cardname is null then '' else t14.cardname end  as 'Cliente_F' 

            
			
from
(
SELECT t99.itemcode, max(t0.docentry) as 'Docentry'
from oitm t99 left join rdr1 t0 on t99.itemcode=t0.itemcode
where substring(t99.itemcode,1,1)='M' and t99.itemcode ='" & par_codice_Commessa & "'
group by t99.itemcode
)
as t10 left join rdr1 t11 on t11.itemcode=t10.itemcode and t11.docentry=t10.docentry
left join ordr t12 on t12.docentry=t11.docentry
left join [TIRELLISRLDB].[DBO].oitm t13 on t13.itemcode=t10.itemcode
left join [TIRELLISRLDB].[DBO].ocrd t14 on t14.cardcode=t12.u_codicebp
left join [TIRELLISRLDB].[DBO].opmg t15 on t15.absentry=t13.u_progetto
LEFT Join [TIRELLI_40].[dbo].OHEM T16 ON T15.OWNER=T16.EMPID

order by t10.docentry DESC, t10.itemcode"
                Else
                    CMD_SAP_2.CommandText =
                "SELECT top 100 trim(t10.matricola) as 'Itemcode'
,    trim(numero_progetto) as 'N_progetto'
,    trim(numero_progetto) as 'Nome_progetto'
,t10.DESC_pm as 'PM'
		,t10.itemname
	,t10.codice_cliente as 'Codice_cliente_finale'
,t10.codice_finale as 'Cliente_f'
,t10.cli_fatt as 'Cardcode'
,t10.dscli_fatt as 'Cardname'
		--,CONVERT(date, CONVERT(char(8), DATA_CONSEGNA), 112) AS DocDueDate
		,'01/01/1990' as 'Docduedate'
		, T10.DSNAZ_FINALE as 'u_destinazione'
, t10.desc_supp

 ,t10.codice_finale as 'Cliente_f'


		, t10.numero_progetto as 'absentry',
		T10.NAME_progetto AS 'DESC_PROGETTO',
		'' as 'Nome_stato',
        '' as 'Livello_rischio_totale', '' as 'Name',
        t10.pm as 'CODICE_PM'
        ,t10.brand AS 'CODICE_BRAND',
		T10.DESC_BRAND AS 'BRAND',
		'' as 'Baia'
		, '' as 'Zona'
		,T10.NOME_STATO AS 'STATO_COMMESSA'
                ,coalesce(t10.Codice_sap_cliente_finale,'') as 'Codice_sap_cliente_finale'
                 ,coalesce(t10.Codice_sap_cliente_fatturazione,'') as 'Codice_sap_cliente_fatturazione'

FROM OPENQUERY(AS400, '
     SELECT  T0.*,
        T1.CONTO AS CONTO_FATTURAZIONE,
		trim(t1.codesap) as Codice_sap_cliente_finale
		,trim(t2.codesap) as Codice_sap_cliente_fatturazione
        ,T2.CONTO AS CONTO_CLIENTE
    FROM TIR90VIS.JGALCOM t0
	left join S786FAD1.TIR90VIS.JGALACF T1 on t1.conto=t0.cli_fatt
	left join S786FAD1.TIR90VIS.JGALACF T2 on t2.conto=t0.codice_cliente

    WHERE 
--SUBSTRING(T0.matricola,1,1) = ''M'' AND
UPPER(t0.matricola)=''" & par_codice_Commessa & "''

  
ORDER BY T0.matricola DESC

limit 100  
') T10"
                End If


                cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

                If cmd_SAP_reader_2.Read() Then

                    Label1.Text = cmd_SAP_reader_2("itemcode")
                    Label2.Text = cmd_SAP_reader_2("itemname")
                    Label16.Text = cmd_SAP_reader_2("u_destinazione")
                    If cmd_SAP_reader_2("docduedate") IsNot DBNull.Value Then
                        Dim data As DateTime = Convert.ToDateTime(cmd_SAP_reader_2("docduedate"))
                        Label17.Text = data.ToString("dd/MM/yyyy") ' Converte la data nel formato "gg/mm/YYYY" e la assegna a Label17.Text
                    Else
                        ' Se il valore è DBNull, puoi assegnare un valore predefinito o vuoto a Label17.Text
                        Label17.Text = "Valore non disponibile"
                        ' oppure Label17.Text = String.Empty
                    End If
                    'Label17.Text = cmd_SAP_reader_2("docduedate")
                    Dim clienteInfo = Ottieni_cliente_papa_macchina(par_codice_Commessa)
                    If cmd_SAP_reader_2("cardcode").ToString() = "" Then
                        bp_code = clienteInfo.cardcode_sap
                        Label3.Text = clienteInfo.cardname
                        final_bp_code = clienteInfo.final_cardcode_sap
                        Label4.Text = clienteInfo.final_cardname
                    Else
                        bp_code = cmd_SAP_reader_2("cardcode").ToString()
                        Label3.Text = cmd_SAP_reader_2("cardname").ToString()
                        final_bp_code = cmd_SAP_reader_2("codice_Cliente_finale").ToString()
                        Label4.Text = cmd_SAP_reader_2("Cliente_F").ToString()
                    End If
                    bp_code_galileo = clienteInfo.cardcode_galileo
                    final_bp_code = clienteInfo.final_cardcode_galileo
                    final_bp_code_galileo = clienteInfo.final_cardcode_galileo

                    If Homepage.ERP_provenienza = "SAP" Then

                    Else
                        bp_code = cmd_SAP_reader_2("Codice_sap_cliente_fatturazione")
                        final_bp_code = cmd_SAP_reader_2("Codice_sap_cliente_finale")
                    End If

                    'cartella_macchina = cmd_SAP_reader_2("u_cartella_macchina")
                    cartella_macchina = trova_percorso_documenti(cmd_SAP_reader_2("itemcode"), "COMMESSA", "")

                    Dim valore As String = cmd_SAP_reader_2("N_progetto")
                    Dim numero As Integer = Integer.Parse(System.Text.RegularExpressions.Regex.Match(valore, "\d+").Value)


                    Button26.Text = numero
                    Button12.Text = valore
                    Label5.Text = cmd_SAP_reader_2("Nome_progetto")
                    Label6.Text = cmd_SAP_reader_2("PM")

                End If
                cmd_SAP_reader_2.Close()
            End Using ' CMD_SAP_2
        End Using ' Cnn1

        If cartella_macchina = "" Or cartella_macchina = "-" Then
            Try
                trova_cartella_macchina()
            Catch ex As Exception

            End Try

        End If

        LinkLabel2.Text = cartella_macchina

    End Sub



    Public Function trova_percorso_documenti(par_codice As String, par_tipo As String, par_codice_progetto As String) As String
        Dim percorso As String = ""
        Using Cnn1 As New SqlConnection(Homepage.sap_tirelli)
            Cnn1.Open()
            Using CMD As New SqlCommand(
                "SELECT TOP 1 [Percorso] FROM [Tirelli_40].[dbo].[Percorsi_Documentale]
                 WHERE (CODICE=@codice OR CODICE=@codice_prog) AND tipo=@tipo", Cnn1)
                CMD.Parameters.AddWithValue("@codice", par_codice)
                CMD.Parameters.AddWithValue("@codice_prog", par_codice_progetto)
                CMD.Parameters.AddWithValue("@tipo", par_tipo)
                Using reader As SqlDataReader = CMD.ExecuteReader()
                    If reader.Read() Then percorso = reader("percorso").ToString()
                End Using
            End Using
        End Using
        Return percorso
    End Function

    Sub elenca_revisioni(par_codice_Commessa As String)
        DataGridView_revisione.Rows.Clear()
        Using Cnn1 As New SqlConnection(Homepage.sap_tirelli)
            Cnn1.Open()
            Using CMD As New SqlCommand(
                "SELECT t0.numero, t0.utente, CONCAT(T1.LASTNAME,' ',T1.FIRSTNAME) AS 'Nome_utente',
                        t0.Data, t0.ora, COALESCE(t0.stato,'') AS Stato, COALESCE(t0.note,'') AS Note
                 FROM [Tirelli_40].[dbo].[Scheda_tecnica_revisioni] t0
                 LEFT JOIN [TIRELLI_40].[dbo].ohem t1 ON t1.empid = t0.utente
                 WHERE t0.commessa = @commessa
                 ORDER BY t0.ID DESC", Cnn1)
                CMD.Parameters.AddWithValue("@commessa", par_codice_Commessa)
                Using reader As SqlDataReader = CMD.ExecuteReader()
                    Do While reader.Read()
                        DataGridView_revisione.Rows.Add(
                            reader("numero"), reader("utente"), reader("nome_utente"),
                            reader("data"), reader("ora"), reader("Stato"), reader("Note"))
                    Loop
                End Using
            End Using
        End Using
    End Sub

    Sub Inserimento_dipendenti()
        ComboBox_utente.Items.Clear()
        Dim dip = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO)
        Dim codRep As String = dip.codice_reparto
        Dim codRepSenzaT As String = codRep.Replace("T", "")
        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Cnn.Open()
            Using CMD As New SqlCommand(
                "SELECT T0.[empID] AS 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome'
                 FROM [TIRELLI_40].[dbo].OHEM T0
                 JOIN [TIRELLI_40].[dbo].oudp t1 ON T0.[dept] = t1.code
                 JOIN [TIRELLI_40].[DBO].COLL_Reparti t2 ON (t2.sap_id_reparto = t1.code OR t2.sap_id_reparto_2 = t1.code)
                 WHERE t0.active = 'Y'
                   AND (CAST(t2.id_reparto AS VARCHAR) = @rep OR CAST(t2.id_reparto AS VARCHAR) = @repNoT)
                 ORDER BY T0.[lastName] + ' ' + T0.[firstName]", Cnn)
                CMD.Parameters.AddWithValue("@rep", codRep)
                CMD.Parameters.AddWithValue("@repNoT", codRepSenzaT)
                Using reader As SqlDataReader = CMD.ExecuteReader()
                    Dim indice As Integer = 0
                    Do While reader.Read()
                        Elenco_dipendenti(indice) = reader("Codice dipendenti").ToString()
                        ComboBox_utente.Items.Add(reader("Nome").ToString())
                        If reader("Codice dipendenti").ToString() = Homepage.ID_SALVATO.ToString() Then
                            ComboBox_utente.Text = reader("Nome").ToString()
                        End If
                        indice += 1
                    Loop
                End Using
            End Using
        End Using
    End Sub

    Sub trova_cartella_macchina()
        Dim cartella_padre As String
        Dim cliente As String
        Dim cartella_esistente As String = ""

        If Label4.Text = "" Then
            cliente = Label3.Text
        Else
            cliente = Label4.Text
        End If

        Dim rootDirectory As String = Homepage.percorso_cartelle_macchine

        Dim directories As String() = System.IO.Directory.GetDirectories(rootDirectory, Strings.Left(codice_commessa, 4) & "*")

        For Each directory As String In directories
            cartella_padre = directory
        Next
        Dim sottocartella As String
        sottocartella = Replace(cartella_padre, rootDirectory, "")

        rootDirectory = cartella_padre

        directories = System.IO.Directory.GetDirectories(rootDirectory, codice_commessa & "*")

        For Each directory As String In directories
            cartella_esistente = directory
        Next


        If cartella_esistente = "" Then




            Directory.CreateDirectory(cartella_padre & "\" & codice_commessa & " " & Label2.Text & " - " & cliente)
            LinkLabel2.Text = sottocartella & "\" & codice_commessa & " " & Label2.Text & " - " & cliente

            Aggiorna_percorso_macchina(LinkLabel2.Text, Label1.Text, "COMMESSA")



        Else



            ' LinkLabel2.Text = cartella_esistente
            cartella_macchina = Replace(cartella_esistente, Homepage.percorso_cartelle_macchine, "")
            LinkLabel2.Text = cartella_macchina
            Aggiorna_percorso_macchina(LinkLabel2.Text, Label1.Text, "COMMESSA")


        End If
    End Sub





    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If MessageBox.Show($"ATTENZIONE! Se non è stato premuto il tasto aggiorna nessuna modifica è stata salvata. Vuoi uscire lo stesso?", "ESCI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            Me.Close()

        End If


    End Sub


    Private Sub ComboBox_reparto_SelectedIndexChanged(sender As Object, e As EventArgs)
        'Homepage.codice_reparto = Elenco_reparti(ComboBox_reparto.SelectedIndex)
        'Homepage.nome_reparto = ComboBox_reparto.Text


        Inserimento_dipendenti()
    End Sub

    Private Sub ComboBox_utente_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_utente.SelectedIndexChanged
        Homepage.ID_SALVATO = Elenco_dipendenti(ComboBox_utente.SelectedIndex)
        'Homepage.UTENTE_NOME_SALVATO = ComboBox_utente.Text
        id_utente = Elenco_dipendenti(ComboBox_utente.SelectedIndex)

        Homepage.Aggiorna_INI_COMPUTER()


    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        If Button26.Text <> 0 Then




            Progetto.Show()
            Progetto.BringToFront()
            Progetto.absentry = Button26.Text
            Progetto.codice_progetto = Button12.Text
            Progetto.inizializza_progetto()

        Else
            MsgBox("Nessun progetto è assegnato a questa commessa")

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)

        trova_cartella_macchina()
    End Sub





    Sub inserisci_numero_nuova_revisione(par_codice_commessa As String, par_stato As String, Par_note As String)
        If Par_note = Nothing Then
            Par_note = ""
        End If
        trova_ultima_revisione(codice_commessa)
        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Cnn.Open()
            Using CMD As New SqlCommand(
                "INSERT INTO [Tirelli_40].[dbo].[Scheda_tecnica_revisioni]
                 ([Commessa],[Numero],[utente],[Data],[ora],[stato],[Note])
                 VALUES (@commessa, @num+1, @utente, GETDATE(), CONVERT(VARCHAR,GETDATE(),108), @stato, @note)", Cnn)
                CMD.Parameters.AddWithValue("@commessa", par_codice_commessa)
                CMD.Parameters.AddWithValue("@num", numero_ultima_revisione)
                CMD.Parameters.AddWithValue("@utente", Homepage.ID_SALVATO)
                CMD.Parameters.AddWithValue("@stato", par_stato)
                CMD.Parameters.AddWithValue("@note", Par_note)
                CMD.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Sub inserisci_valori_scheda_tecnica(par_codice_commessa As String)


        trova_ultima_revisione(codice_commessa)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "INSERT INTO [Tirelli_40].[dbo].[Scheda_Tecnica_valori]
           ([Commessa]
           ,[Rev]
,note
           ,[Velocita]
           ,[Tipo_nastro]
,tipo_nastro_commento
           ,[Altezza_lavoro]
,Altezza_lavoro_commento
           ,[Scarto]
           ,[Nastro_Scarto]
           ,[Pulsantiera]
           ,[Pulsantiera_commento]
           ,[Allacciamenti_elettrici]
           ,[Allacciamenti_elettrici_Commento]
           ,[Allacciamenti_pneumatici]
           ,[Allacciamenti_pneumatici_commento]
           ,[Marca_motiri]
           ,[Industria_4_0]
           ,[Teleassistenza]
,[Marca_plc]
           ,[Controllo_presenza_tappo]
,[Controllo_mal_tappato]
           ,[Controllo_presenza_pescante]
           ,[Controllo_presenza_oggetto]
           ,[telecamera_e_pannello]
           ,[Etichetta_trasparente]
           ,[Illuminatore]
,[Ambiente_standard]
           ,[ambiente_standard_prod_inf]
           ,[Ambiente_installazione_Atex_1]
           ,[Ambiente_installazione_Atex_2]
           ,[Hazloc]
,[light_washable]
,[alleggerita]
           ,[Documenti_aggiuntivi]
           ,[Documenti_aggiuntivi_commento]
           ,[Certificazioni_materiali]
           ,[Lingua_manuale]
           ,[Lingua_manuale_commento]
           ,[Layout]
,[Lista_ricambi_consigliata]
,[Lista_ricambi_dettagli]
,[Lista_ricambi_importo]
           ,[accessorio_tipo_macchina]
           ,[accessorio_sottotipo_macchina]
           ,[accessorio_diametro]
           ,[soffiatura_tipo_trattamento]
           ,[soffiatura_ugello]
           ,[soffiatura_accessorio_lavaggio]
           ,[soffiatura_accessorio_copertura_totale]
           ,[soffiatura_accessorio_aspirazione_polveri]
           ,[soffiatura_accessorio_ionizzazione]
           ,[soffiatura_accessorio_commento]

           ,[Riempimento_sottotipo_macchina]
           ,[Riempimento_N_rubinetti]
           ,[Riempimento_tipologia]
,[Riempimento_tipologia_commento]
           ,[Riempimento_carrello_riempimento]
           ,[Riempimento_carrello_commento]
           ,[Riempimento_carrello_lavaggio]
           ,[Riempimento_carrello_lavaggio_commento]
           ,[Riempimento_movimentazione_rubinetti]
 ,[Riempimento_salita_discesa]

           ,[Riempimento_accessorio_azoto]
           ,[Riempimento_accessorio_pompa_carico_prodotto]
           ,[Riempimento_accessorio_pompa_carico_lavaggio]
           ,[Riempimento_accessorio_serbatoio_esterno]
           ,[Riempimento_accessorio_soffiatura]
           ,[Riempimento_accessorio_circuito_riscaldamento]
           ,[Riempimento_accessorio_raccogli_goccia]
           ,[Riempimento_accessorio_commento]
,[Tappatura_Tipo_Macchina]
           ,[Tappatura_Sottotipo_Macchina]
           ,[Tappatura_numero_teste]
           ,[Tappatura_tappo_trattato_trigger]
           ,[Tappatura_tappo_trattato_pompetta]
           ,[Tappatura_chiusura_trattata_vite]
           ,[Tappatura_chiusura_trattata_pressione]
           ,[Tappatura_tappo_trattato]
           ,[Tappatura_sottotappo_trattato]
           ,[Tappatura_pompetta_trattata]
           ,[Tappatura_stiratura_pompetta_elettronica]
           ,[Tappatura_stiratura_pompetta_pneumatica]
           ,[Tappatura_optional_pompetta_Caricamento_automatica]
           ,[Tappatura_optional_pompetta_centratori]
           ,[Tappatura_optional_pompetta_testa_bordatrice]
           ,[Tappatura_optional_pompetta_testa_trigger]
           ,[Tappatura_optional_pompetta_blocco_ingresso]
           ,[Tappatura_optional_pompetta_antirotazione]
           ,[Tappatura_optional_pompetta_pressetta]
           ,[Tappatura_optional_pompetta_sistema_termosaldatura]
           ,[Tappatura_optional_pompetta_stella]
           ,[Tappatura_optional_pompetta_a_inseguimento]
           ,[Termosaldatura_sottotipo_macchina]
           ,[Termosaldatura_N_teste]
           ,[Termosaldatura_tipologia_sigillatura]
           ,[Termosaldatura_azoto]
           ,[Termosladatura_Passo_termosaldatrice]
           ,[Etichettatura_Sottotipo_macchina]
           ,[Etichettatura_Nome]
           
           ,[Etichettatura_Tipo_testata]
           ,[Etichettatura_tipo_testata_commento]
           ,[Etichettatura_accessorio_stiratura_3_rulli]
           ,[Etichettatura_accessorio_centratore_3_rulli]
           ,[Etichettatura_accessorio_blocco_ingresso]
           ,[Etichettatura_accessorio_centratore_orbitale]
           ,[Etichettatura_accessorio_Divisore_rotante]
           ,[Etichettatura_accessorio_Nastrino_superiore]
           ,[Etichettatura_accessorio_Stiratura_tondi_con_contrasto]
           ,[Etichettatura_accessorio_timbratore_inkjet_nastrino_inkjet]
           ,[Etichettatura_accessorio_timbratore_laser]
           ,[Etichettatura_accessorio_Lettore_fine_bobina]
           ,[Etichettatura_accessorio_gruppo_stampa_barban]
           ,[Etichettatura_accessorio_commento]
,[Etichettatura_diametro_primitivo]
           ,[Etichettatura_Numero_piattelli]
           ,[Etichettatura_senso_rotazione]
,Etichettatura_regolazione_giostra


	,[Etichettatura_coclea] 
	,[Etichettatura_piattello_gomma_vulcanizzata] 
	,[Etichettatura_piattello_gomma_sagoma]
	,[Etichettatura_piattello_espulsore_molla] 
	,[Etichettatura_piattello_perno_espulsore]
	,[Etichettatura_piattello_orientamento_meccanico]
	,[Etichettatura_piattello_orientamento_elettronico]



,[Family_feeling]
,[protezioni_porte]
,[Marca_pannello_operatore]
,[Dimensioni_pannello_operatore]
,[serbatoio]
           ,[Accessori_serbatoio]
           ,[Lavaggio]
           ,[Ingresso_prodotto]
           ,[Pompa_Prodotto]
           ,[CIP]
           ,[Circuito_riempimento]
           ,[Livello_lavaggio]
           ,[Guarnizioni]
,[Tappatura_accessori_presenza_tappo]
           ,[Tappatura_tappo_maltappato]
           ,[Tappatura_presenza_prescante]
           ,[Tappatura_presenza_oggetto]
           ,[Tappatura_termosigillatura]
           ,[Tappatura_fornitura_canale_tappi]
           ,[Tappatura_CF_Canale_Tappi]
,Tipo_catena
,By_pass
,Condizionamento_quadro
,Generico_UTE
,Colori_Attrezzature
,Etichettatura_tetto_antipolvere
,Etichettatura_Lamiera_forata
,Etichettatura_Camma_piattello
 ,[Etichettatura_altezza_albero_Centrale]
      ,[Etichettatura_tipo_testina]
      ,[Etichettatura_tipo_molla]
      ,[Etichettatura_pressurizzazione_bottiglia]
      ,[Etichettatura_lavaggio]
      ,[Etichettatura_stella_entrata]
      ,[Etichettatura_stella_entrata_diametro]
      ,[Etichettatura_stella_uscita]
      ,[Etichettatura_stella_uscita_diametro]
,[Fat_concordata]
,Tappatura_fornitura_torretta
,Soffiatura_numero_ugelli
,[Tipologia_macchina]
,[Modello_macchina]





)
           
     VALUES
           ('" & par_codice_commessa & "'
           ," & numero_ultima_revisione & "+1
           ,'" & Replace(RichTextBox44.Text, "'", "") & "'
           ,'" & TextBox2.Text & "'
           ,'" & ComboBox4.Text & "'
,'" & TextBox13.Text & "'
           ,'" & ComboBox3.Text & "'
,'" & TextBox1.Text & "'
           ,'" & ComboBox1.Text & "'
           ,'" & ComboBox2.Text & "'
           ,'" & ComboBox10.Text & "'
           ,'" & TextBox3.Text & "'
           ,'" & ComboBox7.Text & "'
           ,'" & RichTextBox1.Text & "'
           ,'" & ComboBox8.Text & "'
           ,'" & RichTextBox2.Text & "'
           ,'" & TextBox12.Text & "'
           ,'" & ComboBox9.Text & "'
           ,'" & ComboBox6.Text & "'
           ,'" & ComboBox54.Text & "'
           ," & Controllo_generico(1) & "
," & Controllo_generico(2) & "
          , " & Controllo_generico(3) & "
         , " & Controllo_generico(4) & "
          , " & Controllo_generico(5) & "
           ," & Controllo_generico(6) & "
         ,  " & Controllo_generico(7) & "

," & Configurazione_macchina(1) & "
           ," & Configurazione_macchina(2) & "
           ," & Configurazione_macchina(3) & "
           ," & Configurazione_macchina(4) & "
           ," & Configurazione_macchina(5) & "
           ," & Configurazione_macchina(6) & "
           ," & Configurazione_macchina(7) & "
           ,'" & ComboBox19.Text & "'
           ,'" & TextBox10.Text & "'
           ,'" & ComboBox5.Text & "'
           ,'" & ComboBox33.Text & "'
           ,'" & TextBox11.Text & "'
           ,'" & TextBox8.Text & "'
           ," & Configurazione_macchina(6) & "
,'" & ComboBox36.Text & "'
,'" & TextBox14.Text & "'
           ,'" & ComboBox18.Text & "'
           ,'" & ComboBox20.Text & "'
           ,'" & ComboBox21.Text & "'
           ,'" & ComboBox38.Text & "'
           ,'" & ComboBox39.Text & "'
           ," & soffiatura_accessorio(1) & "
           ," & soffiatura_accessorio(2) & "
           ," & soffiatura_accessorio(3) & "
           ," & soffiatura_accessorio(4) & "
           ,'" & RichTextBox3.Text & "'

           ,'" & ComboBox11.Text & "'
           ,'" & ComboBox12.Text & "'
           ,'" & ComboBox14.Text & "'
,'" & TextBox4.Text & "'
           ,'" & ComboBox16.Text & "'
           ,'" & TextBox5.Text & "'
           ,'" & ComboBox17.Text & "'
           ,'" & TextBox15.Text & "'
           ,'" & ComboBox22.Text & "'
,'" & ComboBox67.Text & "'
          
           ," & riempimento_accessorio(1) & "
           ," & riempimento_accessorio(2) & "
           ," & riempimento_accessorio(3) & "
           ," & riempimento_accessorio(4) & "
           ," & riempimento_accessorio(5) & "
           ," & riempimento_accessorio(6) & "
           ," & riempimento_accessorio(7) & "
           ,'" & TextBox6.Text & "'






 ,'" & ComboBox23.Text & "'
           ,'" & ComboBox24.Text & "'
           ,'" & ComboBox25.Text & "'
           ," & tappatura_dettagli(1) & "
           ," & tappatura_dettagli(2) & "
           ," & tappatura_dettagli(3) & "
           ," & tappatura_dettagli(4) & "
," & tappatura_dettagli(5) & "
," & tappatura_dettagli(6) & "
," & tappatura_dettagli(7) & "
," & tappatura_dettagli(8) & "
," & tappatura_dettagli(9) & "
           ," & tappatura_optional(1) & "
," & tappatura_optional(2) & "
," & tappatura_optional(3) & "
," & tappatura_optional(4) & "
," & tappatura_optional(5) & "
," & tappatura_optional(6) & "
," & tappatura_optional(7) & "
," & tappatura_optional(8) & "
," & tappatura_optional(9) & "
," & tappatura_optional(10) & "
        
           ,'" & ComboBox30.Text & "'
           ,'" & ComboBox31.Text & "'
           ,'" & ComboBox32.Text & "'
           ,'" & ComboBox37.Text & "'
           ,'" & TextBox9.Text & "'
           ,'" & ComboBox26.Text & "'
           ,'" & ComboBox27.Text & "'
           ,'" & ComboBox29.Text & "'
           ,'" & TextBox7.Text & "'
           ," & Etichettatura_accessori(1) & "
           ," & Etichettatura_accessori(2) & "
           ," & Etichettatura_accessori(3) & "
           ," & Etichettatura_accessori(4) & "
           ," & Etichettatura_accessori(5) & "
           ," & Etichettatura_accessori(6) & "
           ," & Etichettatura_accessori(7) & "
           ," & Etichettatura_accessori(8) & "
           ," & Etichettatura_accessori(9) & "
           ," & Etichettatura_accessori(10) & "
           ," & Etichettatura_accessori(11) & "
           ,'" & RichTextBox4.Text & "'

           ,'" & TextBox18.Text & "'
           ,'" & ComboBox50.Text & "'
           ,'" & ComboBox51.Text & "'
           ,'" & ComboBox53.Text & "'

           ,'" & ComboBox52.Text & "'
           ," & Etichettatura_piattelli(1) & "
           ," & Etichettatura_piattelli(2) & "
           ," & Etichettatura_piattelli(3) & "
           ," & Etichettatura_piattelli(4) & "
           ," & Etichettatura_piattelli(5) & "
           ," & Etichettatura_piattelli(6) & "


,'" & ComboBox13.Text & "'
,'" & ComboBox15.Text & "'
,'" & ComboBox34.Text & "'
,'" & Replace(TextBox17.Text, "'", " ") & "'
,'" & ComboBox35.Text & "'
,'" & ComboBox40.Text & "'
,'" & ComboBox41.Text & "'
,'" & ComboBox42.Text & "'
,'" & ComboBox43.Text & "'
,'" & ComboBox44.Text & "'
,'" & ComboBox46.Text & "'
,'" & ComboBox47.Text & "'
,'" & ComboBox48.Text & "'

," & tappatura_accessori(1) & "
," & tappatura_accessori(2) & "
," & tappatura_accessori(3) & "
," & tappatura_accessori(4) & "
," & tappatura_accessori(5) & "
,'" & ComboBox49.Text & "'
,'" & ComboBox45.Text & "'
,'" & Replace(TextBox16.Text, "'", " ") & "'
,'" & Replace(ComboBox57.Text, "'", " ") & "'
,'" & Replace(ComboBox58.Text, "'", " ") & "'
,'" & Replace(RichTextBox5.Text, "'", " ") & "'
,'" & Replace(ComboBox66.Text, "'", " ") & "'
,'" & Replace(ComboBox28.Text, "'", " ") & "'
,'" & Replace(ComboBox55.Text, "'", " ") & "'
,'" & Replace(ComboBox56.Text, "'", " ") & "'




,'" & Replace(ComboBox59.Text, "'", " ") & "'
,'" & Replace(ComboBox60.Text, "'", " ") & "'
,'" & Replace(ComboBox61.Text, "'", " ") & "'
,'" & Replace(ComboBox62.Text, "'", " ") & "'
,'" & Replace(ComboBox63.Text, "'", " ") & "'
,'" & Replace(ComboBox64.Text, "'", " ") & "'
,'" & Replace(TextBox19.Text, "'", " ") & "'
,'" & Replace(ComboBox65.Text, "'", " ") & "'
,'" & Replace(TextBox20.Text, "'", " ") & "'
,'" & Replace(RichTextBox6.Text, "'", " ") & "'
,'" & Replace(ComboBox69.Text, "'", " ") & "'
,'" & ComboBox68.Text & "'
,'" & Replace(If(cbTipologiaMacchina IsNot Nothing, cbTipologiaMacchina.Text, ""), "'", " ") & "'
,'" & Replace(If(cbModelloMacchina IsNot Nothing, cbModelloMacchina.Text, ""), "'", " ") & "'


)"




        CMD_SAP_3.ExecuteNonQuery()

        Cnn3.Close()


    End Sub

    Sub trova_ultima_revisione(par_codice_commessa)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "
               Select  coalesce(t11.numero,0) as 'Ultima_rev',t11.utente, coalesce(CONCAT(T12.LASTNAME,' ',T12.FIRSTNAME),'-') as 'Nome_utente', t11.data,t11.ora
,coalesce(t11.stato,'') as 'Stato_scheda'
from
(
SELECT MAX(t0.id) as 'Ultimo_id'
     
  FROM [Tirelli_40].[dbo].[Scheda_tecnica_revisioni] t0
where t0.commessa ='" & par_codice_commessa & "'

)
as t10 left join [Tirelli_40].[dbo].[Scheda_tecnica_revisioni] t11 on t10.ultimo_id=t11.id
left join [TIRELLI_40].[dbo].ohem t12 on t12.empid=t11.utente

"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            numero_ultima_revisione = cmd_SAP_reader("ultima_Rev")
            ComboBox70.Text = cmd_SAP_reader("Stato_scheda")

            If Not cmd_SAP_reader("Data") Is System.DBNull.Value Then
                Label8.Text = cmd_SAP_reader("Data") & " | " & cmd_SAP_reader("ORA")
            Else
                Label8.Text = "-"
            End If


            Label9.Text = cmd_SAP_reader("Nome_utente")

        Else
            numero_ultima_revisione = 0
            Label8.Text = "-"
            Label9.Text = "-"
        End If

        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim STATO_SCHEDA As String = ""
        STATO_SCHEDA = ComboBox70.Text
        inserisci_valori_scheda_tecnica(codice_commessa)
        inserisci_numero_nuova_revisione(codice_commessa, STATO_SCHEDA, Replace(RichTextBox8.Text, "'", ""))
        elenca_revisioni(codice_commessa)
        trova_ultima_revisione(codice_commessa)
        Label7.Text = numero_ultima_revisione
        ' Label7.Text = numero_ultima_revisione + 1
        RichTextBox8.Text = ""
        MsgBox("Revisione N° " & numero_ultima_revisione & " inserita con successo")
    End Sub

    Private Sub CheckBox24_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox24.CheckedChanged
        If CheckBox24.Checked = False Then
            Controllo_generico(1) = 0
        Else
            Controllo_generico(1) = 1
        End If
    End Sub

    Private Sub CheckBox25_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox25.CheckedChanged
        If CheckBox25.Checked = False Then
            Controllo_generico(2) = 0
        Else
            Controllo_generico(2) = 1
        End If
    End Sub

    Private Sub CheckBox26_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox26.CheckedChanged
        If CheckBox26.Checked = False Then
            Controllo_generico(3) = 0
        Else
            Controllo_generico(3) = 1
        End If
    End Sub

    Private Sub CheckBox27_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox27.CheckedChanged
        If CheckBox27.Checked = False Then
            Controllo_generico(4) = 0
        Else
            Controllo_generico(4) = 1
        End If
    End Sub

    Private Sub CheckBox28_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox28.CheckedChanged
        If CheckBox28.Checked = False Then
            Controllo_generico(5) = 0
        Else
            Controllo_generico(5) = 1
        End If
    End Sub

    Private Sub CheckBox29_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox29.CheckedChanged
        If CheckBox29.Checked = False Then
            Controllo_generico(6) = 0
        Else
            Controllo_generico(6) = 1
        End If
    End Sub

    Private Sub CheckBox30_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox30.CheckedChanged
        If CheckBox30.Checked = False Then
            Controllo_generico(7) = 0
        Else
            Controllo_generico(7) = 1
        End If
    End Sub

    Private Sub ComboBox20_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox20.SelectedIndexChanged
        If ComboBox20.Text = "Piatto di Alimentazione" Then
            GroupBox43.Visible = True
            GroupBox43.Text = "Diametro piatto"
            ComboBox21.Items.Clear()
            ComboBox21.Items.Add("-")
            ComboBox21.Items.Add("Diam. 1000")
            ComboBox21.Items.Add("Diam. 1200")
            ComboBox21.Items.Add("Diam. 1500")
        ElseIf ComboBox20.Text = "Piatto di Raccolta" Then
            GroupBox43.Visible = True
            GroupBox43.Text = "Diametro piatto"
            ComboBox21.Items.Clear()
            ComboBox21.Items.Add("-")
            ComboBox21.Items.Add("Diam. 1000")
            ComboBox21.Items.Add("Diam. 1200")
            ComboBox21.Items.Add("Diam. 1500")
        ElseIf ComboBox20.Text = "Nastri di raffreddamento" Then
            GroupBox43.Visible = True
            GroupBox43.Text = "Tipo di nastri"
            ComboBox21.Items.Clear()
            ComboBox21.Items.Add("-")
            ComboBox21.Items.Add("Sistema di nastri")
            ComboBox21.Items.Add("Camera coibentata")
            ComboBox21.Items.Add("Raffreddamento forzato con Chiller ")


        End If
    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedIndexChanged

        GroupBox38.Visible = True
        GroupBox39.Visible = True
        GroupBox44.Visible = True
        GroupBox45.Visible = True
        GroupBox46.Visible = True
        ComboBox12.Items.Clear()
        ComboBox12.Items.Add("-")
        ComboBox12.Items.Add("1")
        ComboBox12.Items.Add("2")
        ComboBox12.Items.Add("3")
        ComboBox12.Items.Add("4")
        ComboBox12.Items.Add("6")
        ComboBox12.Items.Add("8")
        ComboBox12.Items.Add("10")
        ComboBox12.Items.Add("12")
        ComboBox12.Items.Add("14")
        ComboBox12.Items.Add("16")
        ComboBox12.Items.Add("20")
        ComboBox12.Items.Add("30")
        ComboBox12.Items.Add("36")
        ComboBox12.Items.Add("40")



        ComboBox14.Items.Clear()
        ComboBox14.Items.Add("Pistoni a comando pneumatico")
        ComboBox14.Items.Add("Pistoni a comando brushless")
        ComboBox14.Items.Add("Pistoni a comando stepper")
        ComboBox14.Items.Add("A flussimetri magnetici")
        ComboBox14.Items.Add("A flussimetri massici")
        ComboBox14.Items.Add("A pompe peristaltiche")
        ComboBox14.Items.Add("A livello (leggero vuoto)")
        ComboBox14.Items.Add("A peso")






        ComboBox16.Items.Clear()
        ComboBox16.Items.Add("-")
        ComboBox16.Items.Add("Si")
        ComboBox16.Items.Add("No")
        ComboBox16.Items.Add("altro")


        CheckBox43.Text = "Azoto"
        CheckBox44.Text = "Pompa di carico prodotto"
        CheckBox45.Text = "Pompa di carico lavaggio"
        CheckBox46.Text = "Serbatoio esterno"
        CheckBox47.Text = "Soffiatura"
        CheckBox48.Text = "Circuito di riscaldamento"
        CheckBox49.Text = "Raccogli goccia"



        ComboBox22.Items.Clear()
        ComboBox22.Items.Add("-")
        ComboBox22.Items.Add("Inseguimento")
        ComboBox22.Items.Add("Fissa")

    End Sub

    Private Sub ComboBox18_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox18.SelectedIndexChanged
        GroupBox42.Visible = True
        GroupBox43.Visible = True

        ComboBox20.Items.Clear()

        If ComboBox18.Text = "Accessorio" Then
            ComboBox20.Items.Add("Controllo peso")
            ComboBox20.Items.Add("Elevatore")
            ComboBox20.Items.Add("Estrattore flaconi")
            ComboBox20.Items.Add("Nastri di raffreddamento")
            ComboBox20.Items.Add("Piatto di Alimentazione")
            ComboBox20.Items.Add("Piatto di Raccolta")
            ComboBox20.Items.Add("Sistema di alimentazione automatica con elevatore")
            ComboBox20.Items.Add("Telaio INKJET")
        End If
    End Sub

    Private Sub CheckBox52_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox52.CheckedChanged
        If CheckBox52.Checked = False Then
            Configurazione_macchina(1) = 0
        Else
            Configurazione_macchina(1) = 1
        End If
    End Sub

    Private Sub CheckBox53_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox53.CheckedChanged
        If CheckBox53.Checked = False Then
            Configurazione_macchina(2) = 0
        Else
            Configurazione_macchina(2) = 1
        End If
    End Sub

    Private Sub CheckBox54_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox54.CheckedChanged
        If CheckBox54.Checked = False Then
            Configurazione_macchina(3) = 0
        Else
            Configurazione_macchina(3) = 1
        End If
    End Sub

    Private Sub CheckBox55_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox55.CheckedChanged
        If CheckBox55.Checked = False Then
            Configurazione_macchina(4) = 0
        Else
            Configurazione_macchina(4) = 1
        End If
    End Sub

    Private Sub CheckBox56_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox56.CheckedChanged
        If CheckBox56.Checked = False Then
            Configurazione_macchina(5) = 0
        Else
            Configurazione_macchina(5) = 1
        End If
    End Sub

    Private Sub CheckBox50_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox50.CheckedChanged
        If CheckBox50.Checked = False Then
            Configurazione_macchina(6) = 0
        Else
            Configurazione_macchina(6) = 1
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  LISTA RICAMBI CONSIGLIATI — indicatore e apertura form
    ' ─────────────────────────────────────────────────────────────────

    Sub AggiornaStatoRicambi(par_commessa As String, par_rev As Integer)
        If dgvRicambiScheda Is Nothing Then Return
        ' Resetta grid
        dgvRicambiScheda.AutoGenerateColumns = False
        dgvRicambiScheda.DataSource = Nothing
        dgvRicambiScheda.Columns.Clear()

        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnn
                    cmd.CommandText = "
SELECT NomeLista, COUNT(*) AS NArticoli, SUM(CostoTot) AS Totale
FROM [Tirelli_40].[dbo].[Lista_Ricambi_Consigliati_Righe]
WHERE Commessa=@c AND Rev=@r
GROUP BY NomeLista
ORDER BY NomeLista"
                    cmd.Parameters.AddWithValue("@c", par_commessa)
                    cmd.Parameters.AddWithValue("@r", par_rev)

                    Dim dt As New DataTable()
                    Using rd As SqlDataReader = cmd.ExecuteReader()
                        dt.Load(rd)
                    End Using

                    Dim nListe As Integer = dt.Rows.Count

                    If nListe = 0 Then
                        lblStatoRicambi.Text = "Nessuna lista salvata"
                        lblStatoRicambi.ForeColor = Color.Gray
                        btnApriRicambi.BackColor = SystemColors.Control
                        btnApriRicambi.ForeColor = SystemColors.ControlText
                    Else
                        lblStatoRicambi.Text = "✔ " & nListe & " list" & If(nListe = 1, "a", "e") & " ricambi"
                        lblStatoRicambi.ForeColor = Color.FromArgb(0, 128, 0)
                        lblStatoRicambi.Font = New Font("Segoe UI", 8.0!, FontStyle.Bold Or FontStyle.Italic)
                        btnApriRicambi.BackColor = Color.FromArgb(22, 45, 84)
                        btnApriRicambi.ForeColor = Color.White

                        ' Stile grid
                        Dim navy As Color = Color.FromArgb(22, 45, 84)
                        dgvRicambiScheda.ColumnHeadersDefaultCellStyle.BackColor = navy
                        dgvRicambiScheda.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
                        dgvRicambiScheda.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 7.0!, FontStyle.Bold)
                        dgvRicambiScheda.EnableHeadersVisualStyles = False
                        dgvRicambiScheda.Font = New Font("Segoe UI", 7.5!)
                        dgvRicambiScheda.RowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(210, 225, 245)
                        dgvRicambiScheda.RowsDefaultCellStyle.SelectionForeColor = navy
                        dgvRicambiScheda.GridColor = Color.FromArgb(210, 220, 235)

                        ' Solo nome lista, n. articoli e totale
                        dgvRicambiScheda.Columns.Add(New DataGridViewTextBoxColumn() With {
                            .Name = "colLista", .HeaderText = "Nome lista", .DataPropertyName = "NomeLista", .FillWeight = 55})
                        dgvRicambiScheda.Columns.Add(New DataGridViewTextBoxColumn() With {
                            .Name = "colN", .HeaderText = "Art.", .DataPropertyName = "NArticoli", .FillWeight = 20,
                            .DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleRight}})
                        dgvRicambiScheda.Columns.Add(New DataGridViewTextBoxColumn() With {
                            .Name = "colTot", .HeaderText = "€ tot", .DataPropertyName = "Totale", .FillWeight = 35,
                            .DefaultCellStyle = New DataGridViewCellStyle() With {.Alignment = DataGridViewContentAlignment.MiddleRight, .Format = "N2"}})

                        dgvRicambiScheda.DataSource = dt
                    End If
                End Using
            End Using
        Catch ex As Exception
            lblStatoRicambi.Text = "Errore: " & ex.Message
            lblStatoRicambi.ForeColor = Color.OrangeRed
        End Try
    End Sub

    Private Sub btnApriRicambi_Click(sender As Object, e As EventArgs) Handles btnApriRicambi.Click
        Dim frm As New Form_Lista_Ricambi_Consigliati()
        frm.commessa = codice_commessa
        frm.n_rev = numero_ultima_revisione
        frm.ShowDialog()
        AggiornaStatoRicambi(codice_commessa, numero_ultima_revisione)
    End Sub

    Private Sub btnEliminaLista_Click(sender As Object, e As EventArgs) Handles btnEliminaLista.Click
        If dgvRicambiScheda.CurrentRow Is Nothing Then
            MsgBox("Seleziona una lista da eliminare.", MsgBoxStyle.Information)
            Return
        End If
        Dim nomeLista As String = dgvRicambiScheda.CurrentRow.Cells("colLista").Value?.ToString()
        If String.IsNullOrEmpty(nomeLista) Then Return
        If MsgBox("Eliminare la lista """ & nomeLista & """ e tutte le sue righe?",
                  MsgBoxStyle.YesNo Or MsgBoxStyle.Exclamation, "Conferma eliminazione") <> MsgBoxResult.Yes Then Return
        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnn
                    cmd.CommandText = "DELETE FROM [Tirelli_40].[dbo].[Lista_Ricambi_Consigliati_Righe] WHERE Commessa=@c AND Rev=@r AND NomeLista=@n"
                    cmd.Parameters.AddWithValue("@c", codice_commessa)
                    cmd.Parameters.AddWithValue("@r", numero_ultima_revisione)
                    cmd.Parameters.AddWithValue("@n", nomeLista)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
            AggiornaStatoRicambi(codice_commessa, numero_ultima_revisione)
        Catch ex As Exception
            MsgBox("Errore eliminazione: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnImportaLista_Click(sender As Object, e As EventArgs) Handles btnImportaLista.Click
        Dim frm As New Form_Importa_Lista_Ricambi()
        frm.commessaDestinazione = codice_commessa
        frm.revDestinazione = numero_ultima_revisione
        frm.ShowDialog()
        AggiornaStatoRicambi(codice_commessa, numero_ultima_revisione)
    End Sub

    Private Sub CheckBox51_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox51.CheckedChanged
        If CheckBox51.Checked = False Then
            soffiatura_accessorio(1) = 0
        Else
            soffiatura_accessorio(1) = 1
        End If
    End Sub

    Private Sub CheckBox57_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox57.CheckedChanged
        If CheckBox57.Checked = False Then
            soffiatura_accessorio(2) = 0
        Else
            soffiatura_accessorio(2) = 1
        End If
    End Sub

    Private Sub CheckBox58_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox58.CheckedChanged
        If CheckBox58.Checked = False Then
            soffiatura_accessorio(3) = 0
        Else
            soffiatura_accessorio(3) = 1
        End If
    End Sub

    Private Sub CheckBox59_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox59.CheckedChanged
        If CheckBox59.Checked = False Then
            soffiatura_accessorio(4) = 0
        Else
            soffiatura_accessorio(4) = 1
        End If
    End Sub

    Private Sub CheckBox43_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox43.CheckedChanged
        If CheckBox43.Checked = False Then
            riempimento_accessorio(1) = 0
        Else
            riempimento_accessorio(1) = 1
        End If
    End Sub

    Private Sub CheckBox44_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox44.CheckedChanged
        If CheckBox44.Checked = False Then
            riempimento_accessorio(2) = 0
        Else
            riempimento_accessorio(2) = 1
        End If
    End Sub

    Private Sub CheckBox45_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox45.CheckedChanged
        If CheckBox45.Checked = False Then
            riempimento_accessorio(3) = 0
        Else
            riempimento_accessorio(3) = 1
        End If
    End Sub

    Private Sub CheckBox46_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox46.CheckedChanged
        If CheckBox46.Checked = False Then
            riempimento_accessorio(4) = 0
        Else
            riempimento_accessorio(4) = 1
        End If
    End Sub

    Private Sub CheckBox47_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox47.CheckedChanged
        If CheckBox47.Checked = False Then
            riempimento_accessorio(5) = 0
        Else
            riempimento_accessorio(5) = 1
        End If
    End Sub

    Private Sub CheckBox48_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox48.CheckedChanged
        If CheckBox48.Checked = False Then
            riempimento_accessorio(6) = 0
        Else
            riempimento_accessorio(6) = 1
        End If
    End Sub

    Private Sub CheckBox49_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox49.CheckedChanged
        If CheckBox49.Checked = False Then
            riempimento_accessorio(7) = 0
        Else
            riempimento_accessorio(7) = 1
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ComboBox23_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox23.SelectedIndexChanged
        If ComboBox23.SelectedIndex < 0 Then
            GroupBox48.Visible = False
            GroupBox49.Visible = False
        Else
            GroupBox48.Visible = True
            GroupBox49.Visible = True
        End If

        If ComboBox23.Text = "Tappatore lineare" Then
            GroupBox48.Visible = True
            GroupBox48.Text = "Sottotipo di macchina"
            ComboBox24.Items.Clear()
            ComboBox24.Items.Add("-")
            ComboBox24.Items.Add("Mininebel")
            ComboBox24.Items.Add("Nebel")
            ComboBox24.Items.Add("RO 1 E")
            ComboBox24.Items.Add("RO 1 E S")
            ComboBox24.Items.Add("RO 1 E C")

            GroupBox49.Visible = True
            ComboBox25.Items.Clear()
            ComboBox25.Items.Add("-")
            ComboBox25.Items.Add("1")
            ComboBox25.Items.Add("2")
            ComboBox25.Items.Add("3")

        ElseIf ComboBox23.Text = "Tappatore rotativo (Ro)" Then

            GroupBox48.Text = "Tipo di azionamento"
            ComboBox24.Items.Clear()
            ComboBox24.Items.Add("-")
            ComboBox24.Items.Add("Meccanico")
            ComboBox24.Items.Add("Elettronico")
            ComboBox24.Items.Add("Camma virtuale")

            GroupBox49.Visible = True
            ComboBox25.Items.Clear()
            ComboBox25.Items.Add("-")
            ComboBox25.Items.Add("3")
            ComboBox25.Items.Add("4")
            ComboBox25.Items.Add("6")
            ComboBox25.Items.Add("8")
            ComboBox25.Items.Add("10")
            ComboBox25.Items.Add("12")
            ComboBox25.Items.Add("16")
            ComboBox25.Items.Add("20")



            GroupBox53.Visible = True
            GroupBox53.Text = "Tipologia stiratura"
            CheckBox12.Text = "Stella di stiratura pompetta elettronica"
            CheckBox13.Text = "Stella di stiratura pompetta pneumatica"

        ElseIf ComboBox23.Text = "Tappatore a inseguimento" Then

            GroupBox48.Visible = True
            GroupBox48.Text = "Tipologia di oscar"
            ComboBox24.Items.Clear()
            ComboBox24.Items.Add("-")
            ComboBox24.Items.Add("Oscar 13 azionamenti")
            ComboBox24.Items.Add("Oscar 19 azionamenti")

            GroupBox47.Visible = True
            ComboBox25.Items.Clear()
            ComboBox25.Items.Add("-")
            ComboBox25.Items.Add("1")
            ComboBox25.Items.Add("2")
            ComboBox25.Items.Add("3")





        End If
    End Sub



    Private Sub ComboBox26_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox26.SelectedIndexChanged
        ComboBox27.Items.Clear()

        If ComboBox26.Text = "Lineare" Then

            ComboBox27.Items.Add("-")
            ComboBox27.Items.Add("Miniecho")
            ComboBox27.Items.Add("Delta")
            ComboBox27.Items.Add("Bravo")
            ComboBox27.Items.Add("ADV")

            GroupBox69.Visible = True


        ElseIf ComboBox26.Text = "Rotativa" Then
            ComboBox27.Items.Add("-")
            ComboBox27.Items.Add("Tango")
            ComboBox27.Items.Add("BRB")

            GroupBox69.Visible = False
        End If
    End Sub


    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        If CheckBox12.Checked = False Then
            tappatura_dettagli(8) = 0
        Else
            tappatura_dettagli(8) = 1
        End If
    End Sub



    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        If CheckBox14.Checked = False Then
            tappatura_optional(1) = 0
        Else
            tappatura_optional(1) = 1
        End If
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        If CheckBox15.Checked = False Then
            tappatura_optional(2) = 0
        Else
            tappatura_optional(2) = 1
        End If
    End Sub

    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        If CheckBox16.Checked = False Then
            tappatura_optional(3) = 0
        Else
            tappatura_optional(3) = 1
        End If
    End Sub

    Private Sub CheckBox17_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox17.CheckedChanged
        If CheckBox17.Checked = False Then
            tappatura_optional(4) = 0
        Else
            tappatura_optional(4) = 1
        End If
    End Sub

    Private Sub CheckBox18_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox18.CheckedChanged
        If CheckBox18.Checked = False Then
            tappatura_optional(5) = 0
        Else
            tappatura_optional(5) = 1
        End If
    End Sub

    Private Sub CheckBox19_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox19.CheckedChanged
        If CheckBox19.Checked = False Then
            tappatura_optional(6) = 0
        Else
            tappatura_optional(6) = 1
        End If
    End Sub

    Private Sub CheckBox20_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox20.CheckedChanged
        If CheckBox20.Checked = False Then
            tappatura_optional(7) = 0
        Else
            tappatura_optional(7) = 1
        End If
    End Sub

    Private Sub CheckBox21_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox21.CheckedChanged
        If CheckBox21.Checked = False Then
            tappatura_optional(8) = 0
        Else
            tappatura_optional(8) = 1
        End If
    End Sub

    Private Sub CheckBox22_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox22.CheckedChanged
        If CheckBox22.Checked = False Then
            tappatura_optional(9) = 0
        Else
            tappatura_optional(9) = 1
        End If
    End Sub

    Private Sub CheckBox23_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox23.CheckedChanged
        If CheckBox23.Checked = False Then
            tappatura_optional(10) = 0
        Else
            tappatura_optional(10) = 1
        End If
    End Sub

    Private Sub CheckBox39_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox39.CheckedChanged
        If CheckBox39.Checked = False Then
            Etichettatura_accessori(1) = 0
        Else
            Etichettatura_accessori(1) = 1
        End If
    End Sub

    Private Sub CheckBox40_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox40.CheckedChanged
        If CheckBox40.Checked = False Then
            Etichettatura_accessori(2) = 0
        Else
            Etichettatura_accessori(2) = 1
        End If
    End Sub

    Private Sub CheckBox38_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox38.CheckedChanged
        If CheckBox38.Checked = False Then
            Etichettatura_accessori(3) = 0
        Else
            Etichettatura_accessori(3) = 1
        End If
    End Sub

    Private Sub CheckBox37_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox37.CheckedChanged
        If CheckBox37.Checked = False Then
            Etichettatura_accessori(4) = 0
        Else
            Etichettatura_accessori(4) = 1
        End If
    End Sub

    Private Sub CheckBox36_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox36.CheckedChanged
        If CheckBox36.Checked = False Then
            Etichettatura_accessori(5) = 0
        Else
            Etichettatura_accessori(5) = 1
        End If
    End Sub

    Private Sub CheckBox35_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox35.CheckedChanged
        If CheckBox35.Checked = False Then
            Etichettatura_accessori(6) = 0
        Else
            Etichettatura_accessori(6) = 1
        End If
    End Sub

    Private Sub CheckBox34_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox34.CheckedChanged
        If CheckBox34.Checked = False Then
            Etichettatura_accessori(7) = 0
        Else
            Etichettatura_accessori(7) = 1
        End If

    End Sub

    Private Sub CheckBox33_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox33.CheckedChanged
        If CheckBox33.Checked = False Then
            Etichettatura_accessori(8) = 0
        Else
            Etichettatura_accessori(8) = 1
        End If
    End Sub

    Private Sub CheckBox32_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox32.CheckedChanged
        If CheckBox32.Checked = False Then
            Etichettatura_accessori(9) = 0
        Else
            Etichettatura_accessori(9) = 1
        End If
    End Sub

    Private Sub CheckBox31_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox31.CheckedChanged
        If CheckBox31.Checked = False Then
            Etichettatura_accessori(10) = 0
        Else
            Etichettatura_accessori(10) = 1
        End If
    End Sub

    Private Sub CheckBox41_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox41.CheckedChanged
        If CheckBox34.Checked = False Then
            Etichettatura_accessori(11) = 0
        Else
            Etichettatura_accessori(11) = 1
        End If
    End Sub

    Private Async Sub formati_Click(sender As Object, e As EventArgs) Handles Formati.Enter


        ' Avvia le operazioni in background senza bloccare l'interfaccia utente

        Dim percorso_sap As String = Homepage.sap_tirelli
        Dim percorso_immagini As String = Homepage.Percorso_immagini


        riempi_datagridview_campioni(DataGridView3, codice_bp_campione, bp_code_galileo, final_bp_code_galileo, percorso_immagini, percorso_sap)

        riempi_datagridview_combinazioni(DataGridView1, codice_commessa, Homepage.sap_tirelli)


    End Sub


    Sub riempi_datagridview_campioni(par_datagridview As DataGridView, par_codice_bp_1 As String, par_codice_bp_2 As String, par_codice_bp_3 As String, par_percorso_immagini As String, par_percorso_Sap As String)
        Try



            ' Mostra la riga di caricamento
            par_datagridview.Invoke(Sub()
                                        par_datagridview.Rows.Clear()
                                        par_datagridview.Rows.Add(Nothing, "Caricamento in corso...", Nothing, Nothing, Nothing, Nothing)
                                    End Sub)
        Catch ex As Exception

        End Try

        Task.Run(Sub()
                     ' Crea una lista temporanea per i dati
                     Dim righe As New List(Of Object())

                     ' Apertura connessione SQL
                     Using Cnn1 As New SqlConnection(par_percorso_Sap)
                         Cnn1.Open()

                         Using CMD_SAP_2 As New SqlCommand("
                     SELECT t0.id_campione,  
                            t1.INIZIALE_SIGLA + T0.NOME AS Nome,
                            CASE WHEN COALESCE(t0.immagine, '') = '' THEN 'Bianco.JPG' ELSE t0.immagine END AS immagine,

                            t1.descrizione AS Tipo,
                            COALESCE(t0.Dato_6, '') AS Dato_6,
                            t0.descrizione
                     FROM [TIRELLI_40].[DBO].coll_campioni t0 
                     LEFT JOIN [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t1 
                         ON t0.TIPO_campione = T1.ID_TIPO_CAMPIONE
                     WHERE t0.codice_bp_galileo IN (CAST('" & par_codice_bp_1 & "' AS INTEGER), 
                                            CAST('" & par_codice_bp_2 & "' AS INTEGER),  
                                            CAST('" & par_codice_bp_3 & "' AS INTEGER))
                     ORDER BY t1.INIZIALE_SIGLA, CAST(SUBSTRING(T0.NOME,1,99) AS INTEGER)", Cnn1)

                             Using cmd_SAP_reader_2 As SqlDataReader = CMD_SAP_2.ExecuteReader()
                                 While cmd_SAP_reader_2.Read()
                                     Dim idCampione As Object = cmd_SAP_reader_2("id_campione")
                                     Dim nome As Object = cmd_SAP_reader_2("Nome")
                                     Dim tipo As Object = cmd_SAP_reader_2("Tipo")
                                     Dim dato6 As Object = cmd_SAP_reader_2("Dato_6")
                                     Dim descrizione As Object = cmd_SAP_reader_2("descrizione")
                                     Dim percorsoImmagine As String = par_percorso_immagini & cmd_SAP_reader_2("immagine")

                                     ' Caricamento immagine con fallback
                                     Dim MyImage As Bitmap = Nothing
                                     Try
                                         MyImage = Image.FromFile(percorsoImmagine)
                                     Catch ex As Exception
                                         MyImage = Image.FromFile(par_percorso_immagini & "Bianco.JPG")
                                     End Try

                                     ' Ridimensionamento immagine mantenendo le proporzioni con altezza massima
                                     Dim altezzaMassima As Integer = 80 ' Puoi modificare questo valore
                                     Dim ratio As Double = MyImage.Width / MyImage.Height
                                     Dim nuovaAltezza As Integer = Math.Min(altezzaMassima, MyImage.Height)
                                     Dim nuovaLarghezza As Integer = CInt(nuovaAltezza * ratio)
                                     Dim resizedImage As New Bitmap(MyImage, New Size(nuovaLarghezza, nuovaAltezza))

                                     ' Aggiunta alla lista temporanea
                                     righe.Add(New Object() {idCampione, nome, tipo, resizedImage, dato6, descrizione})
                                 End While
                             End Using
                         End Using
                     End Using


                     ' Aggiorna il DataGridView sulla UI Thread
                     Try
                         par_datagridview.Invoke(Sub()
                                                     par_datagridview.Rows.Clear()
                                                     For Each riga As Object() In righe
                                                         par_datagridview.Rows.Add(riga)
                                                     Next
                                                     par_datagridview.ClearSelection()
                                                 End Sub)
                     Catch ex As Exception

                     End Try

                 End Sub)
    End Sub

    Sub trova_info_campione(par_datagridview As DataGridView, par_codice_campione As String)



        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.id_campione,  t1.INIZIALE_SIGLA + T0.NOME as 'Nome'
, case when (t0.immagine is null or t0.immagine ='') then '\\192.168.0.150\k\Tecnico\DISEGNI ELETTRICI\Basi per Sviluppo software\TWSM\Img_Totem_Collaudi\N_A.JPG' else t0.immagine end as 'immagine'
, t1.descrizione as 'Tipo'
, coalesce(t0.Dato_6,'') as 'Dato_6', t0.descrizione
from [TIRELLI_40].[DBO].coll_campioni t0 
left  join  [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t1 on t0.TIPO_campione= T1.ID_TIPO_CAMPIONE
where t0.id_campione='" & par_codice_campione & "'


"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then


            Dim MyImage As Bitmap

            Try
                MyImage = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine"))
            Catch ex As Exception
                MyImage = Image.FromFile(Homepage.Percorso_immagini & "Bianco.JPG")

            End Try

            par_datagridview.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("Tipo"), MyImage, cmd_SAP_reader_2("dato_6"), cmd_SAP_reader_2("descrizione")) 'Image.FromFile(cmd_SAP_reader_2("immagine"))


        End If


        cmd_SAP_reader_2.Close()
        Cnn1.Close()




    End Sub

    '    Sub riempi_datagridview_combinazioni(par_datagridview As DataGridView, par_codice_commessa As String)


    '        par_datagridview.Rows.Clear()
    '        par_datagridview.Columns(columnName:="vel_richiesta").Visible = False
    '        par_datagridview.Columns(columnName:="immagine_1").Visible = False
    '        par_datagridview.Columns(columnName:="immagine_2").Visible = False
    '        par_datagridview.Columns(columnName:="immagine_3").Visible = False
    '        par_datagridview.Columns(columnName:="immagine_3").Visible = False
    '        par_datagridview.Columns(columnName:="immagine_4").Visible = False
    '        par_datagridview.Columns(columnName:="immagine_5").Visible = False
    '        par_datagridview.Columns(columnName:="immagine_6").Visible = False
    '        par_datagridview.Columns(columnName:="immagine_6").Visible = False
    '        par_datagridview.Columns(columnName:="immagine_7").Visible = False
    '        par_datagridview.Columns(columnName:="immagine_8").Visible = False
    '        par_datagridview.Columns(columnName:="immagine_9").Visible = False
    '        par_datagridview.Columns(columnName:="immagine_10").Visible = False

    '        par_datagridview.Columns(columnName:="nome_1").Visible = False
    '        par_datagridview.Columns(columnName:="nome_2").Visible = False
    '        par_datagridview.Columns(columnName:="nome_3").Visible = False
    '        par_datagridview.Columns(columnName:="nome_4").Visible = False
    '        par_datagridview.Columns(columnName:="nome_5").Visible = False
    '        par_datagridview.Columns(columnName:="nome_6").Visible = False
    '        par_datagridview.Columns(columnName:="nome_7").Visible = False
    '        par_datagridview.Columns(columnName:="nome_8").Visible = False
    '        par_datagridview.Columns(columnName:="nome_9").Visible = False
    '        par_datagridview.Columns(columnName:="nome_10").Visible = False

    '        Dim Cnn1 As New SqlConnection
    '        Cnn1.ConnectionString = Homepage.sap_tirelli
    '        Cnn1.Open()


    '        Dim CMD_SAP_2 As New SqlCommand
    '        Dim cmd_SAP_reader_2 As SqlDataReader


    '        CMD_SAP_2.Connection = Cnn1
    '        CMD_SAP_2.CommandText = "SELECT t0.numero_combinazione,coalesce(t0.[Tipo],'') as 'Tipo', t0.id_combinazione,t0.vel_richiesta, t0.campione_1
    ', t11.INIZIALE_SIGLA + T1.NOME   as 'Nome_1'
    ',case when coalesce(t1.immagine,'')='' then 'Bianco.JPG' else t1.immagine end as 'Immagine_1'
    ', t0.campione_2
    ', t12.INIZIALE_SIGLA + T2.NOME  as 'Nome_2'
    ',case when coalesce(t2.immagine,'')='' then 'Bianco.JPG' else t2.immagine end as 'Immagine_2'
    ', t0.campione_3
    ',t13.INIZIALE_SIGLA + T3.NOME  as 'Nome_3'
    ',case when coalesce(t3.immagine,'')='' then 'Bianco.JPG' else t3.immagine end as 'Immagine_3' 
    ', t0.campione_4,t14.INIZIALE_SIGLA + T4.NOME  as 'Nome_4'
    ',case when coalesce(t4.immagine,'')='' then 'Bianco.JPG' else t4.immagine end as 'Immagine_4', t0.campione_5,t15.INIZIALE_SIGLA + T5.NOME  as 'Nome_5'
    ',case when coalesce(t5.immagine,'')='' then 'Bianco.JPG' else t5.immagine end as 'Immagine_5'
    ', t0.campione_6,t16.INIZIALE_SIGLA + T6.NOME  as 'Nome_6' 
    ',case when coalesce(t6.immagine,'')='' then 'Bianco.JPG' else t6.immagine end as 'Immagine_6'
    ', t0.campione_7, t17.INIZIALE_SIGLA + T7.NOME  as 'Nome_7'
    ',case when coalesce(t7.immagine,'')='' then 'Bianco.JPG' else t7.immagine end as 'Immagine_7'
    ', t0.campione_8,t18.INIZIALE_SIGLA + T8.NOME  as 'Nome_8'
    ',case when coalesce(t8.immagine,'')='' then 'Bianco.JPG' else t8.immagine end as 'Immagine_8', t0.campione_9
    ',t19.INIZIALE_SIGLA + T9.NOME  as 'Nome_9',
    'case when coalesce(t9.immagine,'')='' then 'Bianco.JPG' else t9.immagine end as 'Immagine_9', t0.campione_10
    ',t20.INIZIALE_SIGLA + T10.NOME  as 'Nome_10'
    ',case when coalesce(t10.immagine,'')='' then 'Bianco.JPG' else t10.immagine end as 'Immagine_10'
    ',coalesce([Collaudato],0) as 'Collaudato'

    'FROM [TIRELLI_40].[DBO].COLL_Combinazioni t0
    'left join [TIRELLI_40].[DBO].coll_campioni t1 on t0.campione_1=t1.id_campione
    'left join [TIRELLI_40].[DBO].coll_campioni t2 on t0.campione_2=t2.id_campione
    'left join [TIRELLI_40].[DBO].coll_campioni t3 on t0.campione_3=t3.id_campione
    'left join [TIRELLI_40].[DBO].coll_campioni t4 on t0.campione_4=t4.id_campione
    'left join [TIRELLI_40].[DBO].coll_campioni t5 on t0.campione_5=t5.id_campione
    'left join [TIRELLI_40].[DBO].coll_campioni t6 on t0.campione_6=t6.id_campione
    'left join [TIRELLI_40].[DBO].coll_campioni t7 on t0.campione_7=t7.id_campione
    'left join [TIRELLI_40].[DBO].coll_campioni t8 on t0.campione_8=t8.id_campione
    'left join [TIRELLI_40].[DBO].coll_campioni t9 on t0.campione_9=t9.id_campione
    'left join [TIRELLI_40].[DBO].coll_campioni t10 on t0.campione_10=t10.id_campione

    'left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t11 on t1.TIPO_campione= T11.ID_TIPO_CAMPIONE
    'left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t12 on t2.TIPO_campione= T12.ID_TIPO_CAMPIONE
    'left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t13 on t3.TIPO_campione= T13.ID_TIPO_CAMPIONE
    'left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t14 on t4.TIPO_campione= T14.ID_TIPO_CAMPIONE
    'left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t15 on t5.TIPO_campione= T15.ID_TIPO_CAMPIONE
    'left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t16 on t6.TIPO_campione= T16.ID_TIPO_CAMPIONE
    'left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t17 on t7.TIPO_campione= T17.ID_TIPO_CAMPIONE
    'left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t18 on t8.TIPO_campione= T18.ID_TIPO_CAMPIONE
    'left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t19 on t9.TIPO_campione= T19.ID_TIPO_CAMPIONE
    'left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t20 on t10.TIPO_campione= T20.ID_TIPO_CAMPIONE

    'where t0.commessa='" & par_codice_commessa & "'
    'order by
    't0.numero_combinazione,
    't11.INIZIALE_SIGLA ,  cast(substring(T1.NOME,1,99) as integer)"


    '        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
    '        Dim contatore As Integer = 0


    '        Do While cmd_SAP_reader_2.Read()
    '            par_datagridview.Columns("vel_richiesta").Visible = True
    '            Dim rowIndex As Integer = par_datagridview.Rows.Add(cmd_SAP_reader_2("numero_combinazione"), cmd_SAP_reader_2("id_combinazione"))

    '            ' Popola campioni
    '            For i As Integer = 1 To 10
    '                Dim campioneField As String = $"campione_{i}"
    '                If Not cmd_SAP_reader_2(campioneField) Is System.DBNull.Value Then
    '                    par_datagridview.Rows(rowIndex).Cells(campioneField).Value = cmd_SAP_reader_2(campioneField)
    '                End If
    '            Next

    '            ' Popola immagini
    '            For i As Integer = 1 To 10
    '                Dim immagineField As String = $"Immagine_{i}"
    '                If Not cmd_SAP_reader_2(immagineField) Is System.DBNull.Value AndAlso cmd_SAP_reader_2(immagineField) <> "Bianco.JPG" Then
    '                    Try
    '                        par_datagridview.Rows(rowIndex).Cells(immagineField).Value = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2(immagineField))
    '                        par_datagridview.Columns(immagineField).Visible = True
    '                    Catch ex As Exception
    '                        ' Log o gestione dell'errore
    '                    End Try
    '                End If
    '            Next

    '            ' Popola nomi
    '            For i As Integer = 1 To 10
    '                Dim nomeField As String = $"nome_{i}"
    '                If Not cmd_SAP_reader_2(nomeField) Is System.DBNull.Value Then
    '                    par_datagridview.Rows(rowIndex).Cells(nomeField).Value = cmd_SAP_reader_2(nomeField)
    '                    par_datagridview.Columns(nomeField).Visible = True
    '                End If
    '            Next

    '            ' Altri campi
    '            par_datagridview.Rows(rowIndex).Cells("vel_richiesta").Value = cmd_SAP_reader_2("vel_richiesta")
    '            par_datagridview.Rows(rowIndex).Cells("Collaudo").Value = cmd_SAP_reader_2("Collaudato")
    '            Try
    '                par_datagridview.Rows(rowIndex).Cells("Tipo_combinazione").Value = cmd_SAP_reader_2("Tipo")
    '            Catch ex As Exception

    '            End Try

    '            par_datagridview.Rows(rowIndex).Cells("ID_combinazione").Value = cmd_SAP_reader_2("id_combinazione")

    '        Loop



    '        cmd_SAP_reader_2.Close()
    '        Cnn1.Close()

    '        par_datagridview.ClearSelection()

    '    End Sub



    Public Async Sub riempi_datagridview_combinazioni(par_datagridview As DataGridView, par_codice_commessa As String, PAR_CONNESSIONE_SAP As String)

        ' Pulizia iniziale UI (sul thread principale)
        par_datagridview.Rows.Clear()

        '' Nascondo tutte le colonne immagini e nomi in partenza
        'For Each col As DataGridViewColumn In par_datagridview.Columns
        '    If col.Name Like "immagine_*" OrElse col.Name Like "nome_*" Then
        '        col.Visible = False
        '    End If
        'Next

        ' 👉 Riga di caricamento
        Dim idxLoading As Integer = par_datagridview.Rows.Add()
        par_datagridview.Rows(idxLoading).Cells(0).Value = "🔄 Caricamento..."
        par_datagridview.Rows(idxLoading).DefaultCellStyle.ForeColor = Color.Gray
        par_datagridview.Rows(idxLoading).DefaultCellStyle.Font = New Font(par_datagridview.Font, FontStyle.Italic)



        For Each col As DataGridViewColumn In par_datagridview.Columns
            If col.Name Like "vel_richiesta" OrElse col.Name Like "immagine_*" OrElse col.Name Like "nome_*" Then
                col.Visible = False
            End If
        Next

        ' Esegui in background il caricamento dei dati da SQL
        Dim dati As New List(Of Dictionary(Of String, Object))

        Await Task.Run(Sub()
                           Using Cnn1 As New SqlConnection(PAR_CONNESSIONE_SAP)
                               Cnn1.Open()
                               Using CMD_SAP_2 As New SqlCommand()
                                   CMD_SAP_2.Connection = Cnn1
                                   CMD_SAP_2.CommandText = "SELECT t0.numero_combinazione,coalesce(t0.[Tipo],'') as 'Tipo', t0.id_combinazione,t0.vel_richiesta, t0.campione_1
    , t11.INIZIALE_SIGLA + T1.NOME   as 'Nome_1'
    ,case when coalesce(t1.immagine,'')='' then 'Bianco.JPG' else t1.immagine end as 'Immagine_1'
    , t0.campione_2
    , t12.INIZIALE_SIGLA + T2.NOME  as 'Nome_2'
    ,case when coalesce(t2.immagine,'')='' then 'Bianco.JPG' else t2.immagine end as 'Immagine_2'
    , t0.campione_3
    ,t13.INIZIALE_SIGLA + T3.NOME  as 'Nome_3'
    ,case when coalesce(t3.immagine,'')='' then 'Bianco.JPG' else t3.immagine end as 'Immagine_3' 
    , t0.campione_4,t14.INIZIALE_SIGLA + T4.NOME  as 'Nome_4'
    ,case when coalesce(t4.immagine,'')='' then 'Bianco.JPG' else t4.immagine end as 'Immagine_4', t0.campione_5,t15.INIZIALE_SIGLA + T5.NOME  as 'Nome_5'
    ,case when coalesce(t5.immagine,'')='' then 'Bianco.JPG' else t5.immagine end as 'Immagine_5'
    , t0.campione_6,t16.INIZIALE_SIGLA + T6.NOME  as 'Nome_6' 
    ,case when coalesce(t6.immagine,'')='' then 'Bianco.JPG' else t6.immagine end as 'Immagine_6'
    , t0.campione_7, t17.INIZIALE_SIGLA + T7.NOME  as 'Nome_7'
    ,case when coalesce(t7.immagine,'')='' then 'Bianco.JPG' else t7.immagine end as 'Immagine_7'
    , t0.campione_8,t18.INIZIALE_SIGLA + T8.NOME  as 'Nome_8'
    ,case when coalesce(t8.immagine,'')='' then 'Bianco.JPG' else t8.immagine end as 'Immagine_8', t0.campione_9
    ,t19.INIZIALE_SIGLA + T9.NOME  as 'Nome_9',
    case when coalesce(t9.immagine,'')='' then 'Bianco.JPG' else t9.immagine end as 'Immagine_9', t0.campione_10
    ,t20.INIZIALE_SIGLA + T10.NOME  as 'Nome_10'
    ,case when coalesce(t10.immagine,'')='' then 'Bianco.JPG' else t10.immagine end as 'Immagine_10'
    ,coalesce([Collaudato],0) as 'Collaudato'

    FROM [TIRELLI_40].[DBO].COLL_Combinazioni t0
    left join [TIRELLI_40].[DBO].coll_campioni t1 on t0.campione_1=t1.id_campione
    left join [TIRELLI_40].[DBO].coll_campioni t2 on t0.campione_2=t2.id_campione
    left join [TIRELLI_40].[DBO].coll_campioni t3 on t0.campione_3=t3.id_campione
    left join [TIRELLI_40].[DBO].coll_campioni t4 on t0.campione_4=t4.id_campione
    left join [TIRELLI_40].[DBO].coll_campioni t5 on t0.campione_5=t5.id_campione
    left join [TIRELLI_40].[DBO].coll_campioni t6 on t0.campione_6=t6.id_campione
    left join [TIRELLI_40].[DBO].coll_campioni t7 on t0.campione_7=t7.id_campione
    left join [TIRELLI_40].[DBO].coll_campioni t8 on t0.campione_8=t8.id_campione
    left join [TIRELLI_40].[DBO].coll_campioni t9 on t0.campione_9=t9.id_campione
    left join [TIRELLI_40].[DBO].coll_campioni t10 on t0.campione_10=t10.id_campione

    left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t11 on t1.TIPO_campione= T11.ID_TIPO_CAMPIONE
    left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t12 on t2.TIPO_campione= T12.ID_TIPO_CAMPIONE
    left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t13 on t3.TIPO_campione= T13.ID_TIPO_CAMPIONE
    left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t14 on t4.TIPO_campione= T14.ID_TIPO_CAMPIONE
    left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t15 on t5.TIPO_campione= T15.ID_TIPO_CAMPIONE
    left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t16 on t6.TIPO_campione= T16.ID_TIPO_CAMPIONE
    left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t17 on t7.TIPO_campione= T17.ID_TIPO_CAMPIONE
    left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t18 on t8.TIPO_campione= T18.ID_TIPO_CAMPIONE
    left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t19 on t9.TIPO_campione= T19.ID_TIPO_CAMPIONE
    left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t20 on t10.TIPO_campione= T20.ID_TIPO_CAMPIONE

    where t0.commessa='" & par_codice_commessa & "'
    order by
    t0.numero_combinazione,
    t11.INIZIALE_SIGLA ,  cast(substring(T1.NOME,1,99) as integer)"
                                   Using reader = CMD_SAP_2.ExecuteReader()
                                       While reader.Read()
                                           Dim row As New Dictionary(Of String, Object)
                                           For i As Integer = 0 To reader.FieldCount - 1
                                               row(reader.GetName(i)) = reader(i)
                                           Next
                                           dati.Add(row)
                                       End While
                                   End Using
                               End Using
                           End Using
                       End Sub)

        ' Ora aggiorno la DataGridView sul thread UI
        par_datagridview.Rows.Clear() ' 👉 rimuovo la riga "Caricamento..." prima di riempire


        '' 🔑 reset colonne immagine/nome ad invisibili prima di popolare
        'For Each col As DataGridViewColumn In par_datagridview.Columns
        '    If col.Name Like "immagine_*" OrElse col.Name Like "nome_*" Then
        '        col.Visible = False
        '    End If
        'Next

        ' Ora aggiorno la DataGridView sul thread UI
        For Each r In dati
            Dim rowIndex As Integer = par_datagridview.Rows.Add(r("numero_combinazione"), r("id_combinazione"))

            ' Campioni
            For i As Integer = 1 To 10
                Dim campioneField = $"campione_{i}"
                If r.ContainsKey(campioneField) AndAlso Not IsDBNull(r(campioneField)) Then
                    par_datagridview.Rows(rowIndex).Cells(campioneField).Value = r(campioneField)
                End If
            Next

            ' Immagini
            For i As Integer = 1 To 10
                Dim immagineField = $"Immagine_{i}"
                If r.ContainsKey(immagineField) AndAlso Not IsDBNull(r(immagineField)) AndAlso r(immagineField).ToString() <> "Bianco.JPG" Then
                    Try
                        Dim imgPath = IO.Path.Combine(Homepage.Percorso_immagini, r(immagineField).ToString())
                        If IO.File.Exists(imgPath) Then
                            par_datagridview.Rows(rowIndex).Cells(immagineField).Value = Image.FromFile(imgPath)
                            par_datagridview.Columns(immagineField).Visible = True ' 👉 visibile SOLO se trovata immagine valida
                        End If
                    Catch ex As Exception
                        ' Ignora immagine non caricata
                    End Try
                End If
            Next

            ' Nomi
            For i As Integer = 1 To 10
                Dim nomeField = $"Nome_{i}"
                Dim nomeColonna = $"nome_{i}"
                ' Cerca la chiave nel dizionario senza distinzione tra maiuscole e minuscole
                Dim chiaveEffettiva = r.Keys.FirstOrDefault(Function(k) String.Equals(k, nomeField, StringComparison.OrdinalIgnoreCase))
                If chiaveEffettiva IsNot Nothing AndAlso Not IsDBNull(r(chiaveEffettiva)) Then
                    par_datagridview.Rows(rowIndex).Cells(nomeColonna).Value = r(chiaveEffettiva)
                    par_datagridview.Columns(nomeColonna).Visible = True
                End If
            Next

            ' Altri campi
            If r.ContainsKey("vel_richiesta") Then
                par_datagridview.Rows(rowIndex).Cells("vel_richiesta").Value = r("vel_richiesta")
                par_datagridview.Columns("vel_richiesta").Visible = True
            End If
            If r.ContainsKey("Collaudato") Then
                par_datagridview.Rows(rowIndex).Cells("Collaudo").Value = r("Collaudato")
            End If
            If r.ContainsKey("Tipo") Then
                par_datagridview.Rows(rowIndex).Cells("Tipo_combinazione").Value = r("Tipo")
            End If
            If r.ContainsKey("id_combinazione") Then
                par_datagridview.Rows(rowIndex).Cells("ID_combinazione").Value = r("id_combinazione")
            End If
        Next

        ' 🔍 Nascondi colonne immagini vuote
        For Each col As DataGridViewColumn In par_datagridview.Columns
            If col.Name Like "Immagine_*" Then
                Dim haImmagine As Boolean = False

                For Each row As DataGridViewRow In par_datagridview.Rows
                    Dim cellValue = row.Cells(col.Index).Value
                    If cellValue IsNot Nothing AndAlso TypeOf cellValue Is Image Then
                        haImmagine = True
                        Exit For
                    End If
                Next

                col.Visible = haImmagine
            End If
        Next
        par_datagridview.ClearSelection()
    End Sub


    Public Sub ScaricaImmaginiDaDataGridView(PAR_DATAGRIDVIEW As DataGridView)
        ' Cartella di destinazione per le immagini scaricate
        Dim destinazione As String = "C:\Percorso\Campioni\" & Label1.Text & "\"

        ' Assicurati che la cartella di destinazione esista
        If Not Directory.Exists(destinazione) Then
            Directory.CreateDirectory(destinazione)
        End If

        ' Scorri tutte le righe e colonne del DataGridView
        For i As Integer = 0 To PAR_DATAGRIDVIEW.Rows.Count - 1
            For j As Integer = 0 To PAR_DATAGRIDVIEW.Columns.Count - 1
                ' Verifica che la cella contenga un'immagine
                If TypeOf PAR_DATAGRIDVIEW.Rows(i).Cells(j).Value Is Image Then
                    ' Ottieni l'immagine dalla cella
                    Dim img As Image = DirectCast(PAR_DATAGRIDVIEW.Rows(i).Cells(j).Value, Image)

                    ' Determina il nome del file usando il valore della cella alla sinistra
                    Dim nomeFile As String
                    If j > 0 AndAlso Not IsDBNull(PAR_DATAGRIDVIEW.Rows(i).Cells(j - 1).Value) Then
                        nomeFile = Path.Combine(destinazione, $"{PAR_DATAGRIDVIEW.Rows(i).Cells(j - 1).Value}.png")
                    Else
                        nomeFile = Path.Combine(destinazione, $"Immagine_{i}_{j}.png")
                    End If

                    ' Salva l'immagine con il nome determinato
                    img.Save(nomeFile, System.Drawing.Imaging.ImageFormat.Png)
                End If
            Next
        Next
        Process.Start(destinazione)
        MessageBox.Show("Tutte le immagini sono state scaricate con successo!")
    End Sub



    Public Function Ottieni_numero_combinazioni(par_commessa As String) As DettagliCombinazioni

        Dim dettagli As New DettagliCombinazioni()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT 
    t0.commessa, sum(case when t0.numero_combinazione is null then 0 else 1 end ) as 'N_combinazioni'
    , sum(coalesce(case when t0.collaudato>0 then 1 else 0 end,0)) as 'N_Collaudati'



    FROM [TIRELLI_40].[DBO].COLL_Combinazioni t0


    where t0.commessa='" & par_commessa & "'
    group by t0.commessa
    "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then
            dettagli.Numero_combinazioni = cmd_SAP_reader("N_combinazioni")
            dettagli.Numero_collaudati = cmd_SAP_reader("N_Collaudati")



        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
        Return dettagli
    End Function

    Public Class DettagliCombinazioni
        Public Numero_combinazioni As Integer
        Public Numero_collaudati As Integer

    End Class

    Private Sub Cmd_Inserisci_Click(sender As Object, e As EventArgs) Handles Cmd_Inserisci.Click
        Form_nuovo_campione.Show()
        Form_nuovo_campione.inizializza_form()

        If final_bp_code = "" Then
            'Form_Inserisci_Campioni.Codice_BP = bp_code
            Form_nuovo_campione.Codice_BP_selezionato = bp_code
        Else
            Form_nuovo_campione.Codice_BP_selezionato = final_bp_code

        End If
        Form_nuovo_campione.Codice_BP = bp_code
        Form_nuovo_campione.Codice_BP_finale = final_bp_code
    End Sub



    Private Sub Cmd_Inserimento_Combinazioni_Click(sender As Object, e As EventArgs) Handles Cmd_Inserimento_Combinazioni.Click




        Form_Nuova_combinazione.codice_bp = bp_code_galileo

        Form_Nuova_combinazione.codice_bp_finale = final_bp_code_galileo



        Form_Nuova_combinazione.Label4.Text = Form_Nuova_combinazione.TROVA_MAX_COMBINAZIONE(codice_commessa)
        Form_Nuova_combinazione.codice_commessa = codice_commessa
        Form_Nuova_combinazione.Show()




    End Sub




    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then




            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_1) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_2) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_3) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_4) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_5) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_6) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_7) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_8) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_9) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_10) Then


                Form_campione_visualizza.id_campione = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex - 2).Value
                Form_campione_visualizza.Show()
                Form_campione_visualizza.BringToFront()
                Form_campione_visualizza.inizializza_form()






            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(Numero) Then


                If final_bp_code = "" Then
                    Form_Nuova_combinazione.codice_bp = bp_code
                Else
                    Form_Nuova_combinazione.codice_bp = final_bp_code
                End If

                Form_Nuova_combinazione.codice_commessa = codice_commessa
                Form_Nuova_combinazione.Show()
                Form_Nuova_combinazione.DataGridView2.Rows.Clear()
                Form_Nuova_combinazione.ID_combinazione_salvata = DataGridView1.Rows(e.RowIndex).Cells(columnName:="id_combinazione").Value
                Form_Nuova_combinazione.info_combinazioni(Form_Nuova_combinazione.DataGridView2, DataGridView1.Rows(e.RowIndex).Cells(columnName:="id_combinazione").Value)

            End If

        End If
    End Sub



    Private Sub TextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress, TextBox18.KeyPress, TextBox2.KeyPress
        ' Consenti solo numeri interi, il punto decimale e non consentire il carattere singolo apostrofo nelle TextBox2, TextBox7
        Dim textBox As TextBox = DirectCast(sender, TextBox)

        If (Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "." AndAlso e.KeyChar <> ControlChars.Back) AndAlso textBox IsNot TextBox17 Then
            e.Handled = True
        End If

        ' Impedisci la comparsa della lettera "G" nella TextBox4
        If e.KeyChar = "'" AndAlso textBox Is TextBox17 Then
            e.Handled = True
        End If

        ' Consenti solo un punto decimale
        If e.KeyChar = "." AndAlso textBox.Text.Contains(".") Then
            e.Handled = True
        End If
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub DataGridView_revisione_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_revisione.CellContentClick

    End Sub

    Private Sub DataGridView_revisione_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_revisione.CellClick
        N_rev_visualizza = DataGridView_revisione.Rows(e.RowIndex).Cells(columnName:="N_Rev").Value
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            riempi_scheda_tecnica(codice_commessa, N_rev_visualizza)
            Label7.Text = N_rev_visualizza
            MsgBox("Stai ora visualizzando la revisione " & N_rev_visualizza)
        Catch ex As Exception
            MsgBox("Selezionare un numero di revisione")
        End Try

    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs)

    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        trova_cartella_macchina()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Console.WriteLine(Homepage.percorso_cartelle_macchine)



            LinkLabel2.Text = Replace(FolderBrowserDialog1.SelectedPath.ToUpper(), Homepage.percorso_cartelle_macchine.ToUpper(), "")


            Aggiorna_percorso_macchina(LinkLabel2.Text, Label1.Text, "COMMESSA")

        End If
    End Sub



    Sub Aggiorna_percorso_macchina(par_percorso As String, par_macchina As String, PAR_TIPO As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        par_percorso = Replace(par_percorso.ToUpper(), Homepage.percorso_cartelle_macchine.ToUpper(), "")

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        '        Cmd_SAP.CommandText = "UPDATE [TIRELLISRLDB].[DBO].OITM 
        'SET OITM.U_CARTELLA_MACCHINA='" & par_percorso & "' 
        'where oitm.itemcode='" & par_macchina & "'"

        Cmd_SAP.CommandText = "DELETE [Tirelli_40].[dbo].[Percorsi_Documentale]
  WHERE CODICE='" & par_macchina & "' and tipo='" & PAR_TIPO & "'

INSERT INTO [Tirelli_40].[dbo].[Percorsi_Documentale]
([Tipo]
      ,[Codice]
      ,[Percorso])

VALUES
('" & PAR_TIPO & "'
,'" & par_macchina & "'
,'" & par_percorso & "' )
"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub



    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            Process.Start(Homepage.percorso_cartelle_macchine & LinkLabel2.Text)
        Catch ex As Exception
            MsgBox("Il percorso non esiste")
        End Try
    End Sub

    Private Sub TextBoxes_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress, TextBox2.KeyPress, TextBox4.KeyPress, TextBox5.KeyPress, TextBox6.KeyPress, TextBox7.KeyPress, TextBox8.KeyPress, TextBox9.KeyPress, TextBox10.KeyPress, TextBox11.KeyPress, TextBox13.KeyPress, TextBox14.KeyPress, TextBox15.KeyPress ' Aggiungi qui tutti i tuoi TextBox
        Dim textBox As TextBox = CType(sender, TextBox) ' Ottieni il TextBox attuale

        If e.KeyChar = "'" Or e.KeyChar = "," Then
            e.Handled = True ' Impedisce l'inserimento del carattere '
        End If
    End Sub

    Private Sub RichTextBoxes_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles RichTextBox3.KeyPress, RichTextBox4.KeyPress
        Dim richTextBox As RichTextBox = CType(sender, RichTextBox) ' Ottieni la RichTextBox attuale

        If e.KeyChar = "'" Or e.KeyChar = "," Then
            e.Handled = True ' Impedisce l'inserimento del carattere '
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = False Then
            Configurazione_macchina(6) = 0
        Else
            Configurazione_macchina(6) = 1
        End If
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = False Then
            Configurazione_macchina(7) = 0
        Else
            Configurazione_macchina(7) = 1
        End If
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        If CheckBox11.Checked = False Then
            tappatura_accessori(1) = 0
        Else
            tappatura_accessori(1) = 1
        End If
    End Sub

    Private Sub CheckBox42_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox42.CheckedChanged
        If CheckBox42.Checked = False Then
            tappatura_accessori(2) = 0
        Else
            tappatura_accessori(2) = 1
        End If
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        If CheckBox10.Checked = False Then
            tappatura_accessori(3) = 0
        Else
            tappatura_accessori(3) = 1
        End If
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        If CheckBox9.Checked = False Then
            tappatura_accessori(4) = 0
        Else
            tappatura_accessori(4) = 1
        End If
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        If CheckBox8.Checked = False Then
            tappatura_accessori(5) = 0
        Else
            tappatura_accessori(5) = 1
        End If
    End Sub
    Sub trova_gruppi_etichettaggio(par_datagridview As DataGridView, par_commessa As String)

        par_datagridview.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        Dim contatore As Integer = 1

        CMD_SAP.CommandText = "SELECT t0.[ID]
      ,t0.[N]
      ,t0.[Commessa]
      ,t0.[Tecnologia]
,t0.note
,coalesce(tipo_stazione,'') as 'Tipo_stazione'
  FROM [TIRELLI_40].[DBO].[BRB_Gruppi_etichettaggio] t0

 where t0.commessa='" & par_commessa & "'
order by t0.n"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            Dim immagineFile As String = ""
            If cmd_SAP_reader("Tipo_stazione") <> "" Then
                If cmd_SAP_reader("Tipo_stazione") = "Mensola interna" Then
                    immagineFile = Homepage.Percorso_Immagini_TICKETS & Form_Codici_vendita.trova_immagine_codice("Y00147").ToString
                ElseIf cmd_SAP_reader("Tipo_stazione") = "Stazione esterna modulo fisso" Then
                    immagineFile = Homepage.Percorso_Immagini_TICKETS & Form_Codici_vendita.trova_immagine_codice("Y00148").ToString
                ElseIf cmd_SAP_reader("Tipo_stazione") = "Carrello con ruote" Then
                    immagineFile = Homepage.Percorso_Immagini_TICKETS & Form_Codici_vendita.trova_immagine_codice("Y00149").ToString
                Else
                    immagineFile = Homepage.Percorso_Immagini_TICKETS & "Bianco.JPG"
                End If

            Else
                immagineFile = Homepage.Percorso_Immagini_TICKETS & "Bianco.JPG"
            End If

            ' Controlla se l'immagine esiste prima di aggiungerla

            ' Carica l'immagine
            Dim image As Image

            Try
                image = Image.FromFile(immagineFile)
            Catch ex As Exception

            End Try


            ' Imposta l'altezza massima desiderata
            Dim maxHeight As Integer = 60
            Dim scaleFactor As Double = maxHeight / image.Height
            Dim newWidth As Integer = CInt(image.Width * scaleFactor)
            Dim newSize As New Size(newWidth, maxHeight)

            ' Crea l'immagine ridimensionata
            Dim smallImage As New Bitmap(image, newSize)


            par_datagridview.Rows.Add(cmd_SAP_reader("ID"), contatore, cmd_SAP_reader("Tecnologia"), cmd_SAP_reader("Note"), smallImage)
            contatore += 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()


        par_datagridview.ClearSelection()


    End Sub 'Inserisco le risorse nella combo box


    Public Function trova_SENSO_ORIENTAMENTO_commessa(par_commessa As String)
        Dim senso As String = ""
        trova_ultima_revisione(par_commessa)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        Dim contatore As Integer = 1

        CMD_SAP.CommandText = "Select coalesce(Etichettatura_senso_rotazione,'') as 'Senso_rotazione'



  FROM [TIRELLI_40].[dbo].[Scheda_Tecnica_valori]

where commessa='" & par_commessa & "' and rev ='" & numero_ultima_revisione & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() Then
            senso = cmd_SAP_reader("Senso_rotazione")

        Else
            senso = ""
        End If
        cmd_SAP_reader.Close()
        Cnn.Close()


        Return senso


    End Function 'Inserisco le risorse nella combo box

    Private Sub etichettatrice_Click(sender As Object, e As EventArgs) Handles Etichettatrice.Enter

        trova_gruppi_etichettaggio(DataGridView2, Label1.Text)
        ' PictureBox2.Image = Homepage.Percorso_Immagini_TICKETS & Form_Codici_vendita.trova_immagine_codice("Y00066")

        assegna_foto(PictureBox2, Form_Codici_vendita.trova_immagine_codice("Y00066").ToString)
        assegna_foto(PictureBox3, Form_Codici_vendita.trova_immagine_codice("Y00033").ToString)
        assegna_foto(PictureBox4, "Piattello.png")

    End Sub

    Sub assegna_foto(par_picture_box As PictureBox, par_immagine As String)
        Dim MyImage As Bitmap


        Try
            MyImage = New Bitmap(Homepage.Percorso_Immagini_TICKETS & par_immagine)
        Catch ex As Exception
        End Try
        par_picture_box.Image = CType(MyImage, Image)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form_Gruppi_Etichettaggio.commessa = Label1.Text
        Form_Gruppi_Etichettaggio.stato_gruppo = "Nuovo"
        Form_Gruppi_Etichettaggio.inizializza_form()
        Form_Gruppi_Etichettaggio.Show()

    End Sub





    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click

        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.ColumnIndex = DataGridView2.Columns.IndexOf(N_) Then
            Form_Gruppi_Etichettaggio.commessa = Label1.Text
            Form_Gruppi_Etichettaggio.id = DataGridView2.Rows(e.RowIndex).Cells(columnName:="ID").Value
            Form_Gruppi_Etichettaggio.N = DataGridView2.Rows(e.RowIndex).Cells(columnName:="N_").Value
            Form_Gruppi_Etichettaggio.stato_gruppo = "Visualizza"
            Form_Gruppi_Etichettaggio.inizializza_form()
            Form_Gruppi_Etichettaggio.Show()
        End If
    End Sub

    Private Sub GroupBox40_Enter(sender As Object, e As EventArgs) Handles GroupBox40.Enter

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub TableLayoutPanel25_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel25.Paint

    End Sub



    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Dim par_datagrdiview As DataGridView = DataGridView3
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = par_datagrdiview.Columns.IndexOf(Immagine_) Then



                Form_campione_visualizza.id_campione = par_datagrdiview.Rows(e.RowIndex).Cells(columnName:="Campione_").Value
                Form_campione_visualizza.Show()
                Form_campione_visualizza.BringToFront()
                Form_campione_visualizza.inizializza_form()

            End If


        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Rows(e.RowIndex).Cells("Collaudo").Value = 1 Then
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
        End If
        If DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Value = "M" Then
            DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Style.BackColor = Color.Orange
        ElseIf DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Value = "CDS" Then
            DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Style.BackColor = Color.YellowGreen
        End If
    End Sub



    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = False Then
            Etichettatura_piattelli(1) = 0
        Else
            Etichettatura_piattelli(1) = 1
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = False Then
            Etichettatura_piattelli(2) = 0
        Else
            Etichettatura_piattelli(2) = 1
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = False Then
            Etichettatura_piattelli(3) = 0
        Else
            Etichettatura_piattelli(3) = 1
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = False Then
            Etichettatura_piattelli(4) = 0
        Else
            Etichettatura_piattelli(4) = 1
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = False Then
            Etichettatura_piattelli(5) = 0
        Else
            Etichettatura_piattelli(5) = 1
        End If
    End Sub

    Private Sub CheckBox60_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox60.CheckedChanged
        If CheckBox60.Checked = False Then
            Etichettatura_piattelli(6) = 0
        Else
            Etichettatura_piattelli(6) = 1
        End If
    End Sub

    Private Sub RichTextBox5_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox5.TextChanged

    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        If CheckBox13.Checked = False Then
            tappatura_dettagli(9) = 0
        Else
            tappatura_dettagli(9) = 1
        End If
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        ' Visualizza una InputBox per chiedere la matricola all'utente
        Dim matricola As String = InputBox("Indicare la commessa da cui copiare la scheda tecnica:", "Duplicazione scheda tecnica")
        Dim REVISIONE_DA_COPIARE As Integer
        ' Controlla se l'utente ha inserito qualcosa o ha annullato l'input
        If matricola <> "" Then
            ' Esegui le azioni necessarie con la variabile 'matricola'

            trova_ultima_revisione(matricola)
            REVISIONE_DA_COPIARE = numero_ultima_revisione
            trova_ultima_revisione(Label1.Text)
            DUPLICA_SCHEDA_TECNICA(matricola, Label1.Text, REVISIONE_DA_COPIARE, numero_ultima_revisione + 1)
            inserisci_numero_nuova_revisione(Label1.Text, ComboBox70.Text, Replace(RichTextBox8.Text, "'", ""))
            elenca_revisioni(Label1.Text)
            trova_ultima_revisione(Label1.Text)
            Label7.Text = numero_ultima_revisione
            ' Label7.Text = numero_ultima_revisione + 1
            MsgBox("Revisione N° " & numero_ultima_revisione & " inserita con successo, copiata da " & matricola)
            inizializza_scheda_tecnica(Label1.Text)
        Else
            ' Esegui azioni alternative o mostra un messaggio di avviso
            MessageBox.Show("Nessuna matricola inserita.")
        End If
    End Sub

    Sub DUPLICA_SCHEDA_TECNICA(par_commessa_fonte As String, par_commessa_destinazione As String, par_revisione_FONTE As Integer, par_revisione_DESTINAZIONE As Integer)
        Dim CNN5 As New SqlConnection
        CNN5.ConnectionString = Homepage.sap_tirelli

        Try
            CNN5.Open()

            Dim CMD_SAP_5 As New SqlCommand
            CMD_SAP_5.Connection = CNN5

            CMD_SAP_5.CommandText = "
DECLARE @cols AS NVARCHAR(MAX)
DECLARE @sql AS NVARCHAR(MAX)
DECLARE @COMMESSA AS NVARCHAR(MAX) = @p_commessa
DECLARE @COMMESSA_NUOVA AS NVARCHAR(MAX) = @p_commessa_nuova
DECLARE @REV AS INTEGER = @p_rev
DECLARE @REV_NUOVA AS INTEGER = @p_rev_nuova

-- Solo nome tabella per interrogare INFORMATION_SCHEMA
DECLARE @table_name AS NVARCHAR(MAX) = 'Scheda_Tecnica_valori'
DECLARE @table_full AS NVARCHAR(MAX) = '[tirelli_40].[dbo].[' + @table_name + ']'

-- Costruzione delle colonne da copiare (tranne ID, commessa, rev)
SET @cols = (
    SELECT STUFF((SELECT ', ' + QUOTENAME(COLUMN_NAME)
                  FROM [tirelli_40].INFORMATION_SCHEMA.COLUMNS
                  WHERE TABLE_NAME = @table_name
                    AND COLUMN_NAME NOT IN ('ID', 'commessa', 'rev')
                  FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
)

-- Costruzione SQL dinamico
SET @sql = '
INSERT INTO ' + @table_full + ' (' + @cols + ', commessa, rev)
SELECT ' + @cols + ',
       ''' + @COMMESSA_NUOVA + ''',
       ' + CAST(@REV_NUOVA AS NVARCHAR) + '
FROM ' + @table_full + '
WHERE commessa = ''' + @COMMESSA + ''' AND rev = ' + CAST(@REV AS NVARCHAR)

-- Esecuzione
EXEC sp_executesql @sql
"

            ' Parametri passati correttamente
            CMD_SAP_5.Parameters.AddWithValue("@p_commessa", par_commessa_fonte)
            CMD_SAP_5.Parameters.AddWithValue("@p_commessa_nuova", par_commessa_destinazione)
            CMD_SAP_5.Parameters.AddWithValue("@p_rev", par_revisione_FONTE)
            CMD_SAP_5.Parameters.AddWithValue("@p_rev_nuova", par_revisione_DESTINAZIONE)

            CMD_SAP_5.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Errore: " & ex.Message)
        Finally
            CNN5.Close()
        End Try
    End Sub


    Sub elimina_record_scheda_tecnica(par_commessa_destinazione As String)
        Dim CNN5 As New SqlConnection
        CNN5.ConnectionString = Homepage.sap_tirelli
        CNN5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = CNN5
        CMD_SAP_5.CommandText = "DELETE [Tirelli_40].[dbo].[Scheda_Tecnica_valori]
WHERE COMMESSA='" & par_commessa_destinazione & "'"
        CMD_SAP_5.ExecuteNonQuery()

        CMD_SAP_5.CommandText = "DELETE [Tirelli_40].[dbo].[Scheda_Tecnica_valori]
WHERE COMMESSA='" & par_commessa_destinazione & "'"
        CMD_SAP_5.ExecuteNonQuery()



        CNN5.Close()
    End Sub

    Private Sub ComboBox51_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox51.SelectedIndexChanged
        If ComboBox51.SelectedIndex = 1 Then
            PictureBox1.Image = Image.FromFile(Homepage.percorso_server & "00-Tirelli 4.0\Immagini\Img Scheda tecnica\Macchina oraria.png")
            ' Imposta la modalità di ridimensionamento
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom

            ' Adatta la larghezza mantenendo l'aspetto proporzionale
            PictureBox1.Width = PictureBox1.Image.Width
            PictureBox1.Height = PictureBox1.Image.Height
        ElseIf ComboBox51.SelectedIndex = 2 Then

            PictureBox1.Image = Image.FromFile(Homepage.percorso_server & "00-Tirelli 4.0\Immagini\Img Scheda tecnica\Macchina antioraria.png")
            ' Imposta la modalità di ridimensionamento
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom

            ' Adatta la larghezza mantenendo l'aspetto proporzionale
            PictureBox1.Width = PictureBox1.Image.Width
            PictureBox1.Height = PictureBox1.Image.Height
        Else GroupBox113.Visible = False
        End If

        Dim senso_rotazione As String = ComboBox51.Text



        If senso_rotazione = "SX-DX(CW)" Then
            PictureBox5.Image = Image.FromFile(Homepage.percorso_server & "00-Tirelli 4.0\Immagini\Img Scheda tecnica\Oraria.png")
            ' Imposta la modalità di ridimensionamento

        ElseIf senso_rotazione = "DX-SX(CCW)" Then
            PictureBox5.Image = Image.FromFile(Homepage.percorso_server & "00-Tirelli 4.0\Immagini\Img Scheda tecnica\Antioraria.png")


        End If
        PictureBox5.SizeMode = PictureBoxSizeMode.Zoom




    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        If ComboBox51.SelectedIndex = 1 Then
            Process.Start(Homepage.percorso_server & "00-Tirelli 4.0\Immagini\Img Scheda tecnica\Macchina oraria.pdf")
        ElseIf ComboBox51.SelectedIndex = 2 Then

            Process.Start(Homepage.percorso_server & "00-Tirelli 4.0\Immagini\Img Scheda tecnica\Macchina antioraria.pdf")
        Else
            MsgBox("Nessun file presente")
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ScaricaImmaginiDaDataGridView(DataGridView1)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Process.Start("\\tirfs01\tirelli\00-Tirelli 4.0\Schede tecniche.xlsx")
    End Sub
    Public Async Function mostra_file_async(par_percorso As String, par_treeview As TreeView) As Task
        Dim rootDirectoryPath As String = Homepage.percorso_cartelle_macchine & par_percorso

        ' Esegui l'operazione pesante in background
        Await Task.Run(Sub()
                           ' Pulisce la TreeView e aggiunge il nodo "Caricamento..."
                           par_treeview.Invoke(Sub()
                                                   par_treeview.Nodes.Clear()
                                                   Dim loadingNode As New TreeNode("🔄 Caricamento...")
                                                   par_treeview.Nodes.Add(loadingNode)
                                               End Sub)

                           ' Recupera la directory root
                           Dim rootDirectory As New DirectoryInfo(rootDirectoryPath)
                           Dim rootNode As New TreeNode(rootDirectory.Name) With {.Tag = rootDirectory}

                           ' Popola TreeView in background
                           AddDirectories(rootNode, par_treeview)
                           Addfiles(rootNode, par_treeview)

                           ' Aggiorna la TreeView con i dati finali
                           Try
                               par_treeview.Invoke(Sub()
                                                       par_treeview.Nodes.Clear()
                                                       par_treeview.Nodes.Add(rootNode)
                                                       par_treeview.ExpandAll()
                                                       par_treeview.AllowDrop = True
                                                   End Sub)
                           Catch ex As Exception

                           End Try

                       End Sub)
    End Function

    Public Async Function mostra_file_async_progetto(par_percorso As String, par_treeview As TreeView) As Task
        Dim rootDirectoryPath As String = Homepage.percorso_progetti & par_percorso

        ' Esegui l'operazione pesante in background
        Await Task.Run(Sub()
                           ' Pulisce la TreeView e aggiunge il nodo "Caricamento..."
                           par_treeview.Invoke(Sub()
                                                   par_treeview.Nodes.Clear()
                                                   Dim loadingNode As New TreeNode("🔄 Caricamento...")
                                                   par_treeview.Nodes.Add(loadingNode)
                                               End Sub)

                           ' Recupera la directory root
                           Dim rootDirectory As New DirectoryInfo(rootDirectoryPath)
                           Dim rootNode As New TreeNode(rootDirectory.Name) With {.Tag = rootDirectory}

                           ' Popola TreeView in background
                           AddDirectories(rootNode, par_treeview)
                           Addfiles(rootNode, par_treeview)

                           ' Aggiorna la TreeView con i dati finali
                           Try
                               par_treeview.Invoke(Sub()
                                                       par_treeview.Nodes.Clear()
                                                       par_treeview.Nodes.Add(rootNode)
                                                       par_treeview.ExpandAll()
                                                       par_treeview.AllowDrop = True
                                                   End Sub)
                           Catch ex As Exception

                           End Try

                       End Sub)

    End Function



    Public Sub AddDirectories(parentNode As TreeNode, par_treeview As TreeView)
        Dim parentDirectory As DirectoryInfo = TryCast(parentNode.Tag, DirectoryInfo)
        If parentDirectory Is Nothing OrElse Not parentDirectory.Exists Then Exit Sub

        Debug.WriteLine("📂 Scansiono cartella: " & parentDirectory.FullName)
        Try


            ' Aggiungi icona cartella in modo sicuro
            par_treeview.Invoke(Sub() AggiungiIconaCartella())

            ' Aggiunge tutte le cartelle come nodi figli
            For Each directory As DirectoryInfo In parentDirectory.GetDirectories()
                Dim directoryNode As New TreeNode(directory.Name) With {
        .Tag = directory,
        .ImageKey = "folder",           ' Imposta l'icona della cartella
        .SelectedImageKey = "folder"     ' Imposta l'icona selezionata
    }

                ' Aggiunge il nodo in modo thread-safe
                par_treeview.Invoke(Sub() parentNode.Nodes.Add(directoryNode))

                ' Aggiunge anche i file nella cartella
                Addfiles(directoryNode, par_treeview)

                ' Chiamata ricorsiva per le sottocartelle
                AddDirectories(directoryNode, par_treeview)
            Next
        Catch ex As Exception

        End Try
    End Sub


    ' Funzione per aggiungere l'icona della cartella in modo sicuro
    Private Sub AggiungiIconaCartella()
        If Not ImageList1.Images.ContainsKey("folder") Then
            ImageList1.Images.Add("folder", SystemIcons.WinLogo)
        End If
    End Sub

    Public Sub Addfiles(parentNode As TreeNode, par_treeview As TreeView)
        Dim parentDirectory As DirectoryInfo = TryCast(parentNode.Tag, DirectoryInfo)
        If parentDirectory Is Nothing OrElse Not parentDirectory.Exists Then Exit Sub

        Try
            For Each file As FileInfo In parentDirectory.GetFiles()
                Dim fileNode As New TreeNode(file.Name & " (" & file.LastWriteTime.ToString() & ")") With {.Tag = file}

                ' Sostituzione del percorso di rete
                Dim filepath As String = file.FullName.Replace("\\tirfs01\Tirelli", "T:")

                ' Ottiene l'icona del file
                Dim fileIcon As Icon = SystemIcons.WinLogo
                Try
                    fileIcon = Icon.ExtractAssociatedIcon(filepath)
                Catch ex As Exception
                    Debug.WriteLine("⚠️ ERRORE: Impossibile estrarre icona per " & filepath)
                End Try

                ' Aggiunge l'icona in modo sicuro
                par_treeview.Invoke(Sub() AggiungiIconaFile(file.Extension, fileIcon))

                fileNode.ImageKey = file.Extension

                ' Aggiunge il nodo file in modo thread-safe
                par_treeview.Invoke(Sub() parentNode.Nodes.Add(fileNode))
            Next
        Catch ex As Exception
            Debug.WriteLine("❌ ERRORE in Addfiles: " & ex.Message)
        End Try
    End Sub


    ' Funzione per aggiungere l'icona di un file in modo sicuro
    Private Sub AggiungiIconaFile(extension As String, icon As Icon)
        If Not ImageList1.Images.ContainsKey(extension) Then
            ImageList1.Images.Add(extension, icon)
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        mostra_file_async(LinkLabel2.Text, TreeView1)
    End Sub
    Private Sub Apri_file_Click(sender As Object, e As EventArgs) Handles Apri_file.Click
        ' Verifica se il nodo selezionato è un file
        If TypeOf TreeView1.SelectedNode.Tag Is FileInfo Then
            ' Se il nodo selezionato è un file, apri il file
            Dim file As FileInfo = DirectCast(TreeView1.SelectedNode.Tag, FileInfo)
            Process.Start(file.FullName)
        ElseIf TypeOf TreeView1.SelectedNode.Tag Is DirectoryInfo Then
            ' Se il nodo selezionato è una directory, apri la cartella
            Dim directory As DirectoryInfo = DirectCast(TreeView1.SelectedNode.Tag, DirectoryInfo)
            Process.Start("explorer.exe", directory.FullName)
        End If
    End Sub
    Private Sub RinominaFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RinominaFileToolStripMenuItem.Click

        ' Controlla se un nodo è stato selezionato
        If TreeView1.SelectedNode IsNot Nothing Then
            Dim file As FileInfo = TryCast(TreeView1.SelectedNode.Tag, FileInfo)

            ' Apri una finestra di dialogo per consentire all'utente di inserire il nuovo nome del file
            Dim newFileName As String = InputBox("Inserisci il nuovo nome del file", "Rinomina file", file.Name)

            If Not String.IsNullOrEmpty(newFileName) Then
                ' Rinomina il file
                Dim newFilePath As String = Path.Combine(file.DirectoryName, newFileName)
                FileSystem.Rename(file.FullName, newFilePath)
                mostra_file_async(LinkLabel2.Text, TreeView1)


            End If
        End If

    End Sub


    Private Sub Elimina_file_Click_1(sender As Object, e As EventArgs) Handles Elimina_file.Click
        ' Verifica se il nodo selezionato è un file
        If TypeOf TreeView1.SelectedNode.Tag Is FileInfo Then
            ' Chiedi all'utente conferma prima di eliminare il file
            Dim file As FileInfo = DirectCast(TreeView1.SelectedNode.Tag, FileInfo)
            If MessageBox.Show($"Sei sicuro di voler eliminare il file '{file.Name}'?", "Elimina file", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                file.Delete()
                TreeView1.SelectedNode.Remove()
            End If
        ElseIf TypeOf TreeView1.SelectedNode.Tag Is DirectoryInfo Then
            ' Chiedi all'utente conferma prima di eliminare la directory
            Dim directory As DirectoryInfo = DirectCast(TreeView1.SelectedNode.Tag, DirectoryInfo)
            If MessageBox.Show($"Sei sicuro di voler eliminare la directory '{directory.Name}' e tutti i file contenuti al suo interno?", "Elimina directory", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                directory.Delete(True)
                TreeView1.SelectedNode.Remove()
            End If
        End If
    End Sub
    Private Sub TreeView1_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick
        ' Verifica se il nodo selezionato è una directory
        If TypeOf e.Node.Tag Is FileInfo Then
            ' Apri il file con l'applicazione predefinita
            Dim file As FileInfo = DirectCast(e.Node.Tag, FileInfo)
            Process.Start(file.FullName)
        ElseIf TypeOf e.Node.Tag Is DirectoryInfo Then
            ' Se il nodo selezionato è una directory, apri la cartella
            Dim directory As DirectoryInfo = DirectCast(e.Node.Tag, DirectoryInfo)
            Process.Start("explorer.exe", directory.FullName)
        End If


    End Sub


    Public Function Ottieni_cliente_papa_macchina(par_Codice_macchina As String) As Dettaglicliente

        Dim dettagli As New Dettaglicliente()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn


        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP.CommandText = "          Select  top 100 t30.itemcode, t30.itemname, t30.desc_supp
, t30.cliente
,t30.U_CodiceBP
, t30.cliente_finale,
t30.codice_cliente,
t30.numero_progetto,t30.absentry,t30.name,t30.pm,t30.u_country_of_delivery
,t30.brand,t30.baia
,t30.zona, coalesce(t32.ordine,999) as 'Ordine'
, coalesce(t32.nome,'') as 'Nome_stato'
, t33.livello_rischio_totale
from
(
select t20.itemcode,t20.U_CodiceBP, t20.itemname, t20.desc_supp, t20.cliente, t20.cliente_finale,
t20.codice_cliente,
t20.numero_progetto,t20.absentry,t20.name,t20.pm,t20.u_country_of_delivery
,t20.brand,t20.baia,t20.zona
, max(coalesce(t21.rev,0)) as 'rev_max'
from
(
Select coalesce(t12.U_CodiceBP,'') as 'U_CodiceBP',t10.itemcode, t13.itemname,
coalesce(t13.frgnname,'') as 'Desc_supp',
  COALESCE(t14.cardname, t20.CARDNAME) As 'Cliente'
, COALESCE(t15.cardname,coalesce(t20.CARDNAME,t13.u_final_customer_name)) AS 'Cliente_finale'
,  case when t15.cardCODE is null AND T14.CARDCODE IS NULL then T13.U_FINAL_CUSTOMER_CODE WHEN T15.CARDCODE IS NULL THEN T14.CARDCODE ELSE T15.CARDCODE end AS 'Codice_cliente'
, t16.docnum as 'Numero_progetto' ,t16.[AbsEntry], t16.name
,concat(t17.lastname,' ' , t17.firstname) as 'PM'
, coalesce(t12.u_destinazione,t13.u_country_of_delivery) as 'u_country_of_delivery', coalesce(t13.u_brand,'') as 'Brand'
,coalesce(t19.nome_baia,'') as 'baia'
,coalesce(t19.[Zona],'') as 'Zona'
from
(
Select t7.itemcode, max(t0.docentry) As 'Docentry'
From oitm t7 left Join rdr1 t0 on t7.itemcode=t0.itemcode
Left Join ordr t1 on t1.docentry=t0.docentry And T1.CANCELED='N'
  where t7.itemcode= '" & par_Codice_macchina & "' 
group by t7.itemcode
)
as t10 left join rdr1 t11 on t11.itemcode = t10.itemcode And t11.docentry =t10.docentry
Left Join ordr t12 on t12.docentry=t11.docentry
Left Join oitm t13 on t13.itemcode=t10.itemcode
Left Join ocrd t14 on t14.cardcode=t12.cardcode
Left Join ocrd t15 on t15.cardcode=t12.U_CodiceBP
Left Join opmg t16 on t16.[AbsEntry]=t13.u_progetto
Left Join [TIRELLI_40].[dbo].ohem t17 on t17.empid=t16.owner
Left Join [Tirelli_40].[dbo].[Layout_CAP1] t18 on t18.commessa=t10.itemcode And T18.STATO='O'
                        Left Join [Tirelli_40].[dbo].[Layout_CAP1_nomi] T19 ON T19.NUMERO_baia =t18.baia
Left Join OCRD T20 ON T20.CARDCODE=T13.U_Final_customer_Code

)
as t20
left join [Tirelli_40].[dbo].[Scheda_Tecnica_valori_progetto] t21 on t21.n_progetto=t20.numero_progetto

where T20.ITEMcode = '" & par_Codice_macchina & "' 

group by t20.itemcode, t20.itemname, t20.desc_supp, t20.cliente, t20.cliente_finale,
t20.codice_cliente,t20.U_CodiceBP,
t20.numero_progetto,t20.absentry,t20.name,t20.pm,t20.u_country_of_delivery
,t20.brand,t20.baia,t20.zona
)
as t30
left join [Tirelli_40].[dbo].[Scheda_tecnica_revisioni_progetto] t31 on t31.n_progetto=t30.numero_progetto and t31.Numero=t30.rev_max
LEFT JOIN [Tirelli_40].[dbo].[Scheda_Tecnica_stato_progetto] T32 ON T32.ID=T31.STATO
left join [Tirelli_40].[dbo].[Scheda_Tecnica_valori_progetto] t33 on t33.n_progetto=t30.numero_progetto and t33.rev=t30.rev_max
order by t30.itemcode DESc"

        Else
            CMD_SAP.CommandText =
"SELECT top 100 
trim(t10.matricola) as 'Itemcode', t10.itemname, t10.desc_supp
, T10.DSCLI_FATT as 'Cliente'
, T10.CLI_FATT as 'Codice_cliente',
        t10.codice_finale as 'Cliente_finale',
		t10.codice_cliente as 'Codice_cliente_finale'
		,t10.codice_cliente_SAP
,t10.codice_cliente_finale_SAP
		, t10.itemcode as 'absentry',
        trim(t10.itemcode) as 'Numero_progetto',
		T10.NAME_progetto AS 'DESC_PROGETTO',
		'' as 'Nome_stato',
        '' as 'Livello_rischio_totale', '' as 'Name',
        t10.pm as 'CODICE_PM'
		,t10.DESC_pm as 'PM'
		, T10.DSNAZ_FINALE as u_country_of_delivery,
        t10.brand AS 'CODICE_BRAND',
		trim(T10.DESC_BRAND) AS 'BRAND',
		coalesce(t12.Nome_Baia,'') as 'Baia'
		,coalesce(t12.Zona,'') as 'Zona'
		,DATA_CONSEGNA
		,T10.NOME_STATO AS 'STATO_COMMESSA'
FROM OPENQUERY(AS400, '
    SELECT 
	t0.matricola
	, t0.itemname
	, t0.desc_supp
, T0.DSCLI_FATT
, T0.CLI_FATT 
,t0.codice_finale
,t0.codice_cliente
,coalesce(t1.codesap,'''') as codice_cliente_SAP
,coalesce(t2.codesap,'''') as codice_cliente_finale_SAP
, t0.itemcode
,T0.NAME_progetto
,t0.pm
,t0.DESC_pm
,T0.DSNAZ_FINALE
,t0.brand
,T0.DESC_BRAND
,t0.DATA_CONSEGNA
,T0.NOME_STATO

    FROM TIR90VIS.JGALCOM t0
	left join TIR90VIS.JGALACF t1 on t0.CLI_FATT=t1.conto
	left join TIR90VIS.JGALACF t2 on t0.codice_cliente=t2.conto
    WHERE 
t0.matricola<>'''' and
       UPPER(t0.matricola) = ''" & par_Codice_macchina & "''
       
ORDER BY t0.matricola DESC

limit 100  
') T10 LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1] t11
    ON t11.[Stato]<>'P' and TRIM(t10.matricola) COLLATE SQL_Latin1_General_CP1_CI_AS = t11.commessa COLLATE SQL_Latin1_General_CP1_CI_AS
	left join [Tirelli_40].[dbo].[Layout_CAP1_nomi] t12 on t12.NUMERO_BAIA=t11.Baia"
        End If



        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() Then
            If Homepage.ERP_provenienza = "SAP" Then
                dettagli.cardcode_sap = cmd_SAP_reader("codice_cliente").ToString().Trim()
                dettagli.cardname = cmd_SAP_reader("cliente").ToString().Trim()
                dettagli.final_cardcode_sap = cmd_SAP_reader("U_CodiceBP").ToString().Trim()
                dettagli.final_cardname = cmd_SAP_reader("cliente_finale").ToString().Trim()
            Else
                dettagli.cardcode_sap = cmd_SAP_reader("codice_cliente_SAP").ToString().Trim()
                dettagli.cardname = cmd_SAP_reader("Cliente").ToString().Trim()
                dettagli.final_cardcode_sap = cmd_SAP_reader("codice_cliente_finale_SAP").ToString().Trim()
                dettagli.final_cardname = cmd_SAP_reader("Cliente_finale").ToString().Trim()
                dettagli.cardcode_galileo = cmd_SAP_reader("Codice_cliente").ToString().Trim()
                dettagli.final_cardcode_galileo = cmd_SAP_reader("Codice_cliente_finale").ToString().Trim()
                dettagli.progetto = cmd_SAP_reader("numero_progetto").ToString().Trim()
            End If
        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
        Return dettagli
    End Function

    Public Class Dettaglicliente
        Public cardcode_sap As String = ""
        Public cardname As String = ""
        Public final_cardcode_sap As String = ""
        Public final_cardname As String = ""
        Public cardcode_galileo As String = ""
        Public final_cardcode_galileo As String = ""
        Public progetto As String = ""
    End Class

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Scheda_commessa_Pianificazione.ExportVisibleColumnsToExcel(DataGridView1)
    End Sub

    Private Sub DataGridView_revisione_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_revisione.CellFormatting

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Scheda_commessa_Pianificazione.ExportVisibleColumnsToExcel(DataGridView3)
    End Sub

    Public Async Function crea_modulo_cardini(
    par_flowalayoutpanel As FlowLayoutPanel,
    par_commessa As String,
    par_nome_commessa As String,
    par_conn_string As String) As Task

        ' Crea il modulo ma non lo aggiungi subito
        Dim modulo As New Modulo_cardini()
        modulo.Visible = False
        modulo.inizializza_modulo_vuoto(par_commessa)

        ' Carica i dati pesanti in background
        Dim dati = Await Task.Run(Function()
                                      Return modulo.CaricaDati(par_commessa, par_conn_string)
                                  End Function)

        ' Ora aggiorni tutto sul thread UI
        par_flowalayoutpanel.SuspendLayout()
        par_flowalayoutpanel.Controls.Add(modulo)
        par_flowalayoutpanel.Controls.SetChildIndex(modulo, 0)
        par_flowalayoutpanel.ResumeLayout()

        modulo.VisualizzaDati(dati, par_nome_commessa)
        modulo.Visible = True
    End Function

    Private Sub RichTextBox44_KeyPress(sender As Object, e As KeyPressEventArgs) Handles RichTextBox44.KeyPress
        If e.KeyChar = "'"c Then
            e.Handled = True ' blocca il carattere
        End If
    End Sub

    Private Sub TreeView3_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView3.NodeMouseDoubleClick
        Try
            ' 1. Verifica se il nodo cliccato ha un Tag valido
            If e.Node.Tag Is Nothing Then Exit Sub

            ' 2. Gestione se il nodo è un FILE
            If TypeOf e.Node.Tag Is IO.FileInfo Then
                Dim file As IO.FileInfo = DirectCast(e.Node.Tag, IO.FileInfo)

                ' Verifica che il file esista ancora prima di provare ad aprirlo
                If file.Exists Then
                    Process.Start(New ProcessStartInfo(file.FullName) With {.UseShellExecute = True})
                Else
                    MsgBox("Il file non è più disponibile nel percorso: " & file.FullName, MsgBoxStyle.Critical)
                End If

                ' 3. Gestione se il nodo è una CARTELLA (Directory)
            ElseIf TypeOf e.Node.Tag Is IO.DirectoryInfo Then
                Dim directory As IO.DirectoryInfo = DirectCast(e.Node.Tag, IO.DirectoryInfo)

                If directory.Exists Then
                    Process.Start("explorer.exe", directory.FullName)
                Else
                    MsgBox("La cartella non è raggiungibile.", MsgBoxStyle.Exclamation)
                End If
            End If

        Catch ex As Exception
            MsgBox("Impossibile aprire l'elemento: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If Button26.Text <> 0 Then


            Progetto.Show()
            Progetto.BringToFront()
            Progetto.absentry = Button26.Text
            Progetto.codice_progetto = Button12.Text
            Progetto.inizializza_progetto()

        Else
            MsgBox("Nessun progetto è assegnato a questa commessa")

        End If
    End Sub
End Class