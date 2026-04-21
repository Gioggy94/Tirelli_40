Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Messaging
Imports System
Imports System.Windows.Forms
Imports System.Windows.Media.Media3D






Public Class Magazzino
    Public riga As Integer
    Public Codice_SAP As String
    ' Public Descrizione As String
    '  Public Descrizione_SUP As String
    ' Public osservazioni As String
    Public codice_disegno As String
    '  Public disegno As String
    Public magazzino As String = 0
    ' Public ubicazione As String
    ' Public gruppo As String
    Public magazzino_tot As Decimal
    Public Confermato_tot As Decimal
    Public ordinato_tot As Decimal
    Public disponibile As Decimal
    Public Ass_TOT As Decimal
    '  Public Gestito_a_ferretto As String
    ' Public Max_docentry_traferimenti As Integer
    ' Public docentryodp As Integer
    ' Public Max_DOCNUM_traferimenti As Integer
    Public Linenum_ODP As Integer
    'Public periodo_contabile As String
    'Public serie_RT As String
    'Public serie_trasferimento As String
    Public quantita_ODP As String
    Public quantita_trasferibile As String
    Public check As Integer
    Public Numeratore_OIVL As String
    'Public Prezzo_listino_acquisto As String
    Public MESSAGEID As String
    'Public commessa_ODP As String
    Public magazzino_destinazione As String
    Public magazzino_partenza As String
    Public Elenco_dipendenti(1000) As String
    Public Elenco_produttori(1000) As String
    Public Elenco_nome_SAP(1000) As String
    Public Elenco_gestione(1000) As String
    'Public Codicedip As Integer
    'Public absentry As String
    ' Public stringa_trasferimento As String
    Public Da_trasferire_riga As Decimal
    Public Trasferito_riga As String
    Public Documento As String
    'Public BP_CODE_commessa As String
    'Public BP_name_commessa As String
    Public stato_ODP As String
    ' Public docentryoc As Integer
    Public docnum_OC As String
    'Public ref1 As String
    Public giacenza As Decimal
    Public Password_mag = "mag21"
    'Public test As String
    Public consumo_y As Integer
    Public consumo_y_1 As Integer
    Public consumo_y_2 As Integer
    Public nuovo_valore As Integer
    Public nuovo_valore_string As String
    'Public Soggetto_collaudo As String
    'Public produttore As String
    ' Public catalogo As String
    Public filtro_disegno As String
    Public id_mag As Integer

    Public percorso_Allegato As String
    Public filename_allegato As String
    Public estensione_allegato As String
    Public codice_fornitore As String

    Private filtro_doc As String
    Private filtro_N_doc As String
    Private filtro_osservazioni As String
    Private filtro_mag As String
    Private filtro_Comm As String
    Private provenienza_codice As String = "TIR01"
    Public qualita_tot As Decimal

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub


    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim tb As TextBox = DirectCast(sender, TextBox)
        provenienza_codice = "TIR01"

        ' --- Logica per il ridimensionamento del Font ---
        ' Partiamo da una dimensione base (es. 10) se il testo è corto
        If tb.Text.Length <= 6 Then
            tb.Font = New Font(tb.Font.FontFamily, 18, tb.Font.Style)
        Else
            ' Se siamo dal settimo carattere in su, controlliamo l'ingombro
            Dim g As Graphics = tb.CreateGraphics()
            Dim textSize As SizeF = g.MeasureString(tb.Text, tb.Font)

            ' Se il testo è più largo della TextBox (meno un piccolo margine di sicurezza)
            ' riduciamo il font finché non ci sta o finché non raggiunge un limite minimo (es. 6pt)
            Dim currentSize As Single = tb.Font.Size
            While textSize.Width > (tb.Width - 5) AndAlso currentSize > 6
                currentSize -= 0.5
                Dim tempFont As New Font(tb.Font.FontFamily, currentSize, tb.Font.Style)
                textSize = g.MeasureString(tb.Text, tempFont)
                tb.Font = tempFont
            End While
            g.Dispose()
        End If

        ' --- Tua logica originale ---
        If Len(tb.Text) >= 6 Then
            start_magazzino(TabControl1, tb.Text.ToUpper(), "")
        End If
    End Sub

    Sub start_magazzino(par_tab_control As TabControl, par_codice_sap As String, par_codice_brb As String)
        ' 1. Gestione sicura della connessione
        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Try
                Cnn.Open()
                Documento = ""

                Using CMD_SAP As New SqlCommand()
                    CMD_SAP.Connection = Cnn

                    ' 2. Costruzione query con Parametri per evitare SQL Injection
                    If Homepage.ERP_provenienza = "SAP" Then
                        If provenienza_codice = "TIR01" Then
                            CMD_SAP.CommandText = "SELECT t0.itemcode, t0.itemname, COALESCE(t1.code, '') as code " &
                                             "FROM oitm t0 LEFT JOIN oitt t1 ON t0.itemcode = t1.code " &
                                             "WHERE t0.itemcode = @codice"
                            CMD_SAP.Parameters.AddWithValue("@codice", par_codice_sap)
                        Else
                            CMD_SAP.CommandText = "SELECT t0.itemcode, t0.itemname, COALESCE(t1.code, '') as code " &
                                             "FROM oitm t0 LEFT JOIN oitt t1 ON t0.itemcode = t1.code " &
                                             "WHERE COALESCE(t0.u_codice_brb,'') = @brb AND COALESCE(t0.u_codice_brb,'') <> ''"
                            CMD_SAP.Parameters.AddWithValue("@brb", par_codice_brb)
                        End If
                    Else
                        ' Esempio per AS400 (nota: OPENQUERY non accetta parametri facilmente, 
                        ' ma per coerenza usiamo una variabile pulita)
                        CMD_SAP.CommandText = String.Format("SELECT trim(CODE) AS itemCODE, DES_CODE AS itemname, CHECK_DB AS 'Code' " &
                                         "FROM OPENQUERY(AS400, 'SELECT * FROM S786FAD1.TIR90VIS.JGALART WHERE code = ''{0}''') T10",
                                         par_codice_sap.Replace("'", "''"))
                    End If

                    Using cmd_SAP_reader As SqlDataReader = CMD_SAP.ExecuteReader()
                        If cmd_SAP_reader.Read() Then
                            ' Assegnazione variabili locali
                            Codice_SAP = cmd_SAP_reader("itemcode").ToString()
                            par_codice_sap = Codice_SAP
                            Button1.Visible = True

                            ' --- OTTIMIZZAZIONE CHIAVE: Chiamata singola alla funzione ---
                            Dim dett = OttieniDettagliAnagrafica(par_codice_sap)
                            ' -------------------------------------------------------------

                            ' Logica UI Distinta Base
                            If dett.Distinta_base = "N" Then
                                Button1.Text = "Crea distinta base"
                                Button1.BackColor = Color.Yellow
                            Else
                                Button1.Text = "Visualizza distinta base"
                                Button1.BackColor = Color.Lime
                            End If

                            ' Gestione Tab
                            Select Case TabControl1.SelectedTab.Name ' Usare il nome è più robusto dell'oggetto
                                Case TabPage1.Name : trasferito(Codice_SAP, DataGridView_trasferito)
                                Case TabPage2.Name : ordinato(Codice_SAP, DataGridView_ordinato)
                                Case TabPage4.Name : rt_aperte(DataGridView1, Codice_SAP)
                            End Select

                            ROF()

                            ' Aggiornamento controlli con i dati caricati in memoria (dett)
                            DateTimePicker4.Value = DateTime.Today.AddDays(-30)
                            Lista_registrazioni(Codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, filtro_doc, filtro_N_doc, filtro_osservazioni, filtro_mag, filtro_Comm)
                            consumi()
                            allegati(DataGridView2, par_codice_sap, TabControl2)

                            ' Popolamento campi (molto più veloce perché non interroga il DB ogni volta)
                            TextBox1.Text = dett.Disegno
                            ComboBox3.Text = dett.gestione_magazzino
                            RichTextBox1.Text = dett.motivazione_stock
                            DateTimePicker1.Value = dett.data_valutazione
                            TextBox_descrizione.Text = dett.Descrizione
                            TextBox3.Text = dett.Descrizione_SUP
                            Label22.Text = dett.Osservazioni
                            ComboBox1.Text = dett.Produttore
                            TextBox8.Text = dett.Catalogo
                            Label20.Text = dett.n_mag
                            Label21.Text = dett.n_cass
                            Label6.Text = "€ " & dett.Prezzo_listino_acquisto.ToString("N2")
                            Label19.Text = dett.minordrqty
                            Label7.Text = dett.trattamento
                            Label4.Text = dett.Ubicazione
                            Label_gestito_a_ferretto.Text = dett.Gestito_a_ferretto
                            Label5.Text = dett.codice_brb
                            Label2.Text = dett.unita_misura
                            Label3.Text = dett.nome_fornitore
                            ComboBox2.Text = dett.Gruppo
                            codice_fornitore = dett.codice_fornitore

                            visualizza_picture(dett.Disegno, PictureBox2)

                            If provenienza_codice <> "TIR01" Then TextBox2.Text = par_codice_sap

                            RadioButton1.Checked = (dett.attivo = "Y")
                            RadioButton2.Checked = (dett.attivo <> "Y")

                            TableLayoutPanel9.Visible = True
                            TabControl1.Visible = True
                            giacenze_magazzino(DataGridView_magazzino, Codice_SAP)

                            DataGridView4.ClearSelection()
                            DataGridView_trasferito.ClearSelection()
                        Else
                            ' Gestione record non trovato
                            Button1.Visible = False
                            TableLayoutPanel9.Visible = False
                            TabControl1.Visible = False
                        End If
                    End Using
                End Using
            Catch ex As Exception
                MessageBox.Show("Errore durante il caricamento: " & ex.Message)
            End Try
        End Using ' La connessione viene chiusa automaticamente qui
    End Sub

    Public Function OttieniDettagliAnagrafica(par_Codice_SAP As String) As DettagliAnagrafica

        Dim dettagli As New DettagliAnagrafica()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.CommandTimeout = 0

        CMD_SAP.Connection = Cnn
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP.CommandText = "select coalesce(T0.[QryGroup10],'N') as 'Gestito a Ferretto',
            COALESCE( t0.itemname ,'') as 'itemname',
        COALESCE(t0.u_ubicazione ,'' ) as 'ubicazione',
        case when t0.frgnname is null then '' else t0.frgnname end as 'frgnname' ,
        case when t0.u_disegno is null then '' else t0.U_disegno end as 'U_disegno',
        case when t0.validcomm is null then '' else t0.validcomm end as 'validcomm',
case when t0.u_prg_tir_trattamento is null then '' else t0.u_prg_tir_trattamento end as 'Trattamento',
        T1.[ItmsGrpNam] as 'Gruppo',
case when T2.PRICE is null then 0 else t2.price end as 'price', case when t0.minlevel is null then 0 else t0.minlevel end as 'punto_rio' 
, case when t0.minordrqty is null then 0 else t0.minordrqty end as 'minordrqty',
CASE WHEN T0.[U_PRG_QLT_HasTC] IS NULL THEN '' ELSE T0.[U_PRG_QLT_HasTC] END AS 'Soggetto_collaudo'   
,case when t3.firmname is null then '' else t3.firmname end as 'Produttore'
, case when t0.suppcatnum is null then '' else t0.suppcatnum end as 'Catalogo'
,coalesce(case when 
coalesce(t0.u_ubicazione_labelling,'') <>'' then concat('CAP3: ',t0.u_ubicazione_labelling) else '' end,'') as 'Ubicazione_labelling'

,coalesce(t0.u_codice_brb,'') as 'Codice_BRB'
,case 
when t0.frozenfor='Y' and t0.frozenfrom<=getdate() and t0.frozento>=getdate() then 'N' 
when t0.validfor='Y' and t0.validfrom<=getdate() and t0.validto>=getdate() then 'Y' 
when t0.validfor='Y' and t0.validfrom>getdate() OR t0.validto<=getdate() then 'N' 
when t0.validfor='Y' and t0.validfrom is null then 'Y'
when t0.FROZENFOR='N' and t0.validfrom is null then 'Y'

else 'N' end AS 'ATTIVO'
,coalesce(t0.invntryuom,'') as 'UM'
,coalesce(t0.cardcode,'') as 'Codice_fornitore'
,coalesce(t4.cardname,'') as 'nome_fornitore'
,T0.[PrcrmntMtd]
,coalesce(t0.u_final_customer_name,'') as 'Cliente'
,coalesce(t0.u_gestione_magazzino,'') as 'Gestione_magazzino'
,coalesce(t0.u_ubimag,'') as 'Motivazione_stock',
coalesce(t0.u_data_valutazione_stock,'') as 'U_data_valutazione'
,CASE WHEN T6.CODE IS NULL THEN 'N' ELSE case when coalesce(t7.min,999999999) = 999999999 then 'N' else 'Y'end END AS 'DB'
, coalesce(t7.min,999999999) as 'Check_db'
, coalesce(t8.magazzino,0) as 'n_mag'
, coalesce(t8.cassetto,0) as 'n_cass'
,coalesce(t0.u_progetto,0) as 'N_progetto'
,T1.[ItmsGrpCod] as 'CodiceGruppo'

from oitm t0 INNER JOIN OITB T1 ON T0.[ItmsGrpCod] = T1.[ItmsGrpCod]
LEFT JOIN ITM1 T2 ON T2.ITEMCODE=T0.ITEMCODE
left JOIN OMRC T3 ON T0.[FirmCode] = T3.[FirmCode]
left join ocrd t4 on t4.cardcode=t0.cardcode
left join ufd1 t5 on t5.fieldid=103 and tableid='OITM' AND T5.fldvalue=t0.u_gestione_magazzino
left join oitt t6 on t6.code=t0.itemcode
left join 
(
select min(t0.visorder) as 'Min' from itt1 t0 where t0.father='" & par_Codice_SAP & "'
) 
t7 on t7.min >=0
left join [Tirelli_40].[dbo].[Cassetto_codici] t8 on t8.[Codice]=t0.itemcode

        where t0.itemcode= '" & par_Codice_SAP & "' AND T2.PRICELIST=2"
        Else


            CMD_SAP.CommandText =
"SELECT   
CDDT AS CDDT,
trim(CODE) AS CODE,
COD_FERR AS [Gestito a Ferretto],
DES_CODE AS itemname,
UBI_CODE AS ubicazione,
LNG_CODE AS frgnname,
trim(DISEGNO) AS u_disegno,
STAT_CODE AS validcomm,
coalesce(COD_TRAT,'') AS codice_Trattamento,
coalesce(DESC_TRAT,'') AS  Trattamento,
GRUP_ART AS CodiceGruppo,
DESC_GRP AS Gruppo,
COSTO_STD AS Price,
punto_rio AS punto_rio,
QTA_SAFE AS qta_sicurezza,
lotto_min AS minordrqty,
SOGG_COL AS Soggetto_collaudo,
PROD_FOR AS Produttore,
CODAR_FOR AS Catalogo,
UBI_SEC AS ubicazione_labelling,
CODE_BRB AS Codice_BRB,
UMIS AS UM,
COD_FOR AS Codice_fornitore,
DESC_FOR AS nome_fornitore,
TIPO_PARTE AS PrcrmntMtd,
GEST_COMM AS Gestione_magazzino,
MOTIV_MAG AS motivazione_stock,
CONVERT(DATE, CAST(DATA_STOCK AS CHAR(8)), 112)  AS 'u_data_valutazione',
CHECK_DB AS DB
,coalesce(T11.MAGAZZINO,0) AS n_mag
,coalesce(T11.cassETTO,0) AS N_cass
,stat_code as 'Attivo'
,'' as 'Cliente'
,999 AS N_progetto

FROM OPENQUERY(AS400, '
    SELECT *
    FROM S786FAD1.TIR90VIS.JGALART
    WHERE code = ''" & par_Codice_SAP & "''
') T10 
left join [Tirelli_40].[dbo].[Cassetto_codici] T11 ON T10.code COLLATE SQL_Latin1_General_CP850_CI_AS
 = T11.CODICE COLLATE SQL_Latin1_General_CP850_CI_AS"

        End If


        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then
            dettagli.Descrizione = cmd_SAP_reader("itemname")
            dettagli.Descrizione_SUP = cmd_SAP_reader("frgnname")
            dettagli.Osservazioni = cmd_SAP_reader("validcomm")
            dettagli.Disegno = cmd_SAP_reader("u_disegno")
            dettagli.codice_brb = cmd_SAP_reader("Codice_BRB")
            dettagli.attivo = cmd_SAP_reader("ATTIVO")
            dettagli.unita_misura = cmd_SAP_reader("UM")
            dettagli.codice_fornitore = cmd_SAP_reader("Codice_fornitore")
            dettagli.nome_fornitore = cmd_SAP_reader("nome_fornitore")
            dettagli.Cliente = cmd_SAP_reader("cliente")
            dettagli.Distinta_base = cmd_SAP_reader("DB")
            dettagli.n_mag = cmd_SAP_reader("n_mag")
            dettagli.n_cass = cmd_SAP_reader("N_cass")

            If cmd_SAP_reader("ubicazione") = "" And cmd_SAP_reader("ubicazione_labelling") <> "" Then
                dettagli.Ubicazione = cmd_SAP_reader("ubicazione_labelling")

            Else

                dettagli.Ubicazione = cmd_SAP_reader("ubicazione")

            End If

            dettagli.Gestito_a_ferretto = cmd_SAP_reader("Gestito a Ferretto")
            dettagli.Gruppo = cmd_SAP_reader("Gruppo")
            dettagli.CodiceGruppo = cmd_SAP_reader("CodiceGruppo")
            dettagli.Prezzo_listino_acquisto = cmd_SAP_reader("Price")

            dettagli.gestione_magazzino = cmd_SAP_reader("Gestione_magazzino")
            dettagli.motivazione_stock = cmd_SAP_reader("motivazione_stock")
            dettagli.data_valutazione = cmd_SAP_reader("u_data_valutazione")


            dettagli.Minimo = Math.Round(cmd_SAP_reader("punto_rio"))
            Label18.Text = dettagli.Minimo



            dettagli.minordrqty = Math.Round(cmd_SAP_reader("minordrqty"))
            Label19.Text = dettagli.minordrqty
            dettagli.trattamento = cmd_SAP_reader("Trattamento")

            dettagli.Soggetto_collaudo = cmd_SAP_reader("Soggetto_collaudo")
            dettagli.Produttore = cmd_SAP_reader("Produttore")
            dettagli.Catalogo = cmd_SAP_reader("Catalogo")
            dettagli.Approvvigionamento = cmd_SAP_reader("PrcrmntMtd")
            dettagli.u_progetto = cmd_SAP_reader("N_progetto")


            dettagli.Test = "SI"
        Else
            dettagli.Test = "NO"

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
        Return dettagli
    End Function


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        codice_disegno = TextBox1.Text
        visualizza_disegno(codice_disegno)
    End Sub

    Sub visualizza_disegno(parametro_codice_disegno As String)
        Dim num_foglio As Integer = 1

        If File.Exists(Homepage.percorso_disegni_generico & "PDF\" & parametro_codice_disegno & ".PDF") Then
            Process.Start(Homepage.percorso_disegni_generico & "PDF\" & parametro_codice_disegno & ".PDF")
        ElseIf File.Exists(Homepage.percorso_disegni_generico & "PDF\" & parametro_codice_disegno & "_foglio_" & num_foglio & ".PDF") Then
            Process.Start(Homepage.percorso_disegni_generico & "PDF\" & parametro_codice_disegno & "_foglio_" & num_foglio & ".PDF")
        Else
            MsgBox("PDF non trovato")
        End If




    End Sub

    Sub visualizza_picture(parametro_codice_disegno As String, par_picturebox As PictureBox)

        Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & parametro_codice_disegno & ".PNG"

        ' Pulisce sempre prima (evita immagini bloccate)
        If par_picturebox.Image IsNot Nothing Then
            par_picturebox.Image.Dispose()
            par_picturebox.Image = Nothing
        End If

        If File.Exists(percorso) Then
            Using fs As New FileStream(percorso, FileMode.Open, FileAccess.Read)
                par_picturebox.Image = Image.FromStream(fs)
            End Using

            par_picturebox.SizeMode = PictureBoxSizeMode.Zoom
        Else
            ' Se non trova l'immagine resta vuota
            par_picturebox.Image = Nothing
        End If

    End Sub

    Sub apri_picture(parametro_codice_disegno As String)

        Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & parametro_codice_disegno & ".PNG"

        If File.Exists(percorso) Then
            Process.Start(percorso)
        End If

    End Sub



    Sub trova_ID()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "SELECT 'o', max(case when t0.id is null then 0 else t0.id end )+1 as 'ID' from [Tirelli_40].[dbo].[INVENTARIO] t0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then

            If Not cmd_SAP_reader("ID") Is System.DBNull.Value Then
                id_mag = cmd_SAP_reader("ID")
            Else
                id_mag = 1
            End If
        Else
            id_mag = 1
        End If

        Cnn.Close()
        cmd_SAP_reader.Close()


    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then
            riga = e.RowIndex
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button_inventario.Click

        'Inventario.magazzini()
        Inventario.Show()

        'Lavorazioni_MES.inserimento_dipendenti_MES(Inventario.ComboBox_DIPENDENTE, Lavorazioni_MES.Elenco_dipendenti_MES)
    End Sub



    Sub giacenze_magazzino(par_datagridview As DataGridView, par_codice_sap As String)

        Dim Cnn1 As New SqlConnection
        par_datagridview.Rows.Clear()
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_2.CommandText = "SELECT 
0 as QTA_QUA,
0 as qta_ass,
T0.[WhsCode]
, CASE WHEN T0.[OnHand] is null then 0 else T0.[OnHand] END AS 'onhand' 
, case when T0.[IsCommited] is null then 0 else T0.[IsCommited] end as 'iscommited' 
, case when T0.[OnOrder] is null then 0 else T0.[OnOrder] end as 'onorder'  
FROM OITW T0 WHERE (T0.[OnHand]<>0 or t0.iscommited<>0 or t0.onorder<>0) 
and t0.itemcode='" & par_codice_sap & "'"


        Else

            CMD_SAP_2.CommandText = "SELECT *
FROM OPENQUERY(AS400, '
     SELECT 
        cod_mag AS whscode,
        qta_mag -qta_ass AS onhand,
		qta_ass,
		QTA_QUA,
        qta_imp + qta_ven AS iscommited,
        qta_acq +qta_odp AS onorder
    FROM S786FAD1.TIR90VIS.JGALMAG
    WHERE codart = ''" & par_codice_sap & "''
') T10
WHERE NOT (
 onhand = 0
    AND iscommited = 0
    AND onorder = 0
	AND QTA_QUA = 0
and qta_ass=0
)"

        End If
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()


            par_datagridview.Rows.Add(cmd_SAP_reader_2("whscode"),
                                      cmd_SAP_reader_2("onhand"),
                                      cmd_SAP_reader_2("QTA_ass"),
                                      cmd_SAP_reader_2("QTA_QUA"),
                                      cmd_SAP_reader_2("iscommited"),
                                      cmd_SAP_reader_2("onorder"))
        Loop


        Cnn1.Close()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP.CommandText = " select  
0 as 'Qualita_tot'
,0 as 'Ass_TOT'
,sum(case when T0.[OnHand] is null then 0 else T0.[OnHand] end ) as 'Magazzino_TOT'
, sum(case when T0.[iscoMMited] is null then 0 else T0.[iscoMMited] end) as 'Confermato_TOT'
, sum(case when T0.[onorder] is null then 0 else T0.[onorder] end) as 'ordinato_TOT'
,  sum(case when T0.[OnHand] is null then 0 else T0.[OnHand] end-case when T0.[iscoMMited] is null then 0 else T0.[iscoMMited] end+case when T0.[onorder] is null then 0 else T0.[onorder] end) as 'Disponibile'
FROM OITW T0 WHERE (T0.[OnHand]>0 or t0.iscommited>0 or t0.onorder>0) and t0.itemcode='" & par_codice_sap & "'"

        Else
            CMD_SAP.CommandText = "SELECT *
FROM OPENQUERY(AS400, '
    SELECT 
        
        sum(qta_mag-qta_ass) AS Magazzino_TOT,
sum(qta_ass) as Ass_TOT,
sum(QTA_QUA) AS qualita_TOT,
        sum(qta_imp + qta_ven) AS Confermato_TOT,
        sum(qta_acq +qta_odp) AS ordinato_TOT,
		sum(qta_disp) as Disponibile
		
    FROM S786FAD1.TIR90VIS.JGALMAG
    WHERE codart = ''" & par_codice_sap & "''
	
') T10"


        End If
        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            If Not cmd_SAP_reader("Magazzino_TOT") Is System.DBNull.Value Then
                magazzino_tot = cmd_SAP_reader("Magazzino_TOT")
            Else
                magazzino_tot = 0
            End If

            If Not cmd_SAP_reader("Ass_TOT") Is System.DBNull.Value Then
                Ass_TOT = cmd_SAP_reader("Ass_TOT")
            Else
                Ass_TOT = 0
            End If

            If Not cmd_SAP_reader("qualita_TOT") Is System.DBNull.Value Then
                qualita_tot = cmd_SAP_reader("qualita_TOT")
            Else
                qualita_tot = 0
            End If

            If Not cmd_SAP_reader("Confermato_TOT") Is System.DBNull.Value Then
                Confermato_tot = cmd_SAP_reader("Confermato_TOT")
            Else
                Confermato_tot = 0
            End If

            If Not cmd_SAP_reader("ordinato_TOT") Is System.DBNull.Value Then
                ordinato_tot = cmd_SAP_reader("ordinato_TOT")
            Else
                ordinato_tot = 0
            End If

            If Not cmd_SAP_reader("disponibile") Is System.DBNull.Value Then
                disponibile = cmd_SAP_reader("disponibile")
            Else
                disponibile = 0
            End If

        Else
            magazzino_tot = 0
            qualita_tot = 0
            Confermato_tot = 0
            ordinato_tot = 0
            disponibile = 0
            Ass_TOT = 0

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
        par_datagridview.Rows.Add("TOTALE", magazzino_tot, Ass_TOT, qualita_tot, Confermato_tot, ordinato_tot, disponibile)
        par_datagridview.ClearSelection()
    End Sub

    Sub trasferito(par_codice_sap As String, par_datagridview As DataGridView)
        Dim WIP As Decimal = 0
        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_2.CommandText = "Select '' AS nome_baia ,'' as 'Saldo_imp','' as 'codmag_im', t0.[Documento] , T0.[ODP],t0.progressivo, t0.[OC],t0.[Codice], T0.ITEMNAME,t0.[Q.tà pianificata], T0.[Trasferito], t0.[Da trasferire], t0.[U_prg_azs_commessa], t0.[U_utilizz], t0.[status] , t0.[u_PRODUZIONE],T0.LINENUM, t0.resname, t0.startdate, T0.DIV
from (
SELECT 'ODP' as 'Documento', T1.[DocNum] as 'ODP',coalesce(t1.u_progressivo_commessa,'') as 'Progressivo', '' as 'OC',T2.[ItemCode] as 'Codice', T1.PRODNAME AS 'ITEMNAME',t0.[PlannedQty]-coalesce(t0.issuedqty,0) as 'Q.tà pianificata', coalesce(T0.[U_PRG_WIP_QtaSpedita] , 0)-coalesce(t0.issuedqty,0) AS 'Trasferito', T0.[U_PRG_WIP_QtaDaTrasf] as 'Da trasferire', t1.U_prg_azs_commessa, case when t1.[U_utilizz] is null or t1.u_utilizz='' then coalesce(t4.u_final_customer_name,'') else t1.u_utilizz end as 'u_utilizz', t1.status , T1.u_PRODUZIONE,T0.LINENUM, t3.resname, t1.startdate
,case when coalesce(t5.location,'')='13' then 'BRB01' ELSE 'TIR01' END AS 'DIV'
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] 
inner join oitm t2 on t2.itemcode= t0.itemcode
left join orsc t3 on t3.visrescode=t1.u_fase
left join oitm t4 on t4.itemcode=t1.U_PRG_AZS_Commessa
left join owhs t5 on t5.whscode=t0.wareHouse

WHERE T0.[ItemCode] = '" & par_codice_sap & "' AND  (T1.Status <> N'L' )  AND  (T1.Status <> N'C' )

union all


SELECT 'OC','','',T1.DOCNUM AS 'OC',T0.[ItemCode], '',  T0.[OpenQty], T0.[U_Trasferito] as 'Totale trasferito', T0.[U_Datrasferire] as 'Da trasferire' , COALESCE(T1.U_MATRCDS,CONCAT('_',T1.DOCNUM)) AS 'CDS', T1.CARDNAME,'','' ,T0.LINENUM,'', t1.docduedate

,COALESCE(T0.OcrCode,'') AS 'DIV'
FROM RDR1 T0  INNER JOIN ORDR T1 ON T0.[DocEntry] = T1.[DocEntry]



WHERE T1.DocStatus = N'o' and T0.[OpenCreQty] >0 and T0.[ItemCode] ='" & par_codice_sap & "'

) as t0
group by
t0.[Documento] , T0.[ODP],t0.progressivo, t0.[OC],t0.[Codice], T0.ITEMNAME,t0.[Q.tà pianificata], T0.[Trasferito], t0.[Da trasferire], t0.[U_prg_azs_commessa], t0.[U_utilizz], t0.[status] , t0.[u_PRODUZIONE],T0.LINENUM,t0.resname , t0.startdate,T0.DIV
order by t0.startdate"
        Else
            CMD_SAP_2.CommandText = "


SELECT
    t10.documento,
    t10.odp,
    t10.progressivo,
    COALESCE(t10.oc, '')          AS oc,
    t10.codart                    AS codice,
    t10.itemname,
    t10.qtapia                    AS [Q.tà pianificata],
    t10.qtatra                    AS Trasferito,
    t10.qtadatra                  AS [Da trasferire],
    trim(t10.matricola)                 AS U_prg_azs_commessa,
	cod_commessa,
	cod_sottocommessa,
coalesce(t12.[Nome_Baia],'') as 'Nome_baia' ,
    'MANCA'                       AS U_utilizz,
    t10.status,
    'MANCA'                       AS U_produzione,
    999                           AS Linenum,
    'MANCA'                       AS resname,
    CONVERT(DATE, CAST(t10.Dtasca AS CHAR(8)), 112) AS startdate,
    mag_ver                           AS DIV,
codmag_im
,saldo_imp

FROM OPENQUERY([AS400], '
    SELECT
        documento,
        odp,
        progressivo,
        oc,
        codart,
        itemname,
        qtapia,
        qtatra,
        qtadatra,
        matricola,
        status,
        Dtasca,
codmag_im,
mag_ver as mag_ver
,saldo_imp
,cod_commessa as cod_commessa
,cod_sottocommessa as cod_sottocommessa
    FROM S786FAD1.TIR90VIS.JGALIMP
    WHERE evaso_odp <> ''S'' and codart=''" & par_codice_sap & "''
') AS t10
LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1] t11
       ON t11.commessa COLLATE SQL_Latin1_General_CP850_CI_AS
        = TRIM(t10.matricola) COLLATE SQL_Latin1_General_CP850_CI_AS
      AND t11.stato = 'O'
left join [Tirelli_40].[dbo].[Layout_CAP1_nomi] t12 on t12.numero_baia= t11.baia

"

        End If


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader_2("documento"),
                                      cmd_SAP_reader_2("ODP"),
                                      cmd_SAP_reader_2("progressivo"),
                                      cmd_SAP_reader_2("ITEMNAME"),
                                      cmd_SAP_reader_2("OC"),
                                      cmd_SAP_reader_2("Q.tà pianificata"),
                                      cmd_SAP_reader_2("Trasferito"),
                                      cmd_SAP_reader_2("Da trasferire"),
                                      cmd_SAP_reader_2("U_prg_azs_commessa"),
                                      cmd_SAP_reader_2("Nome_baia"),
                                      Business_partner_della_commessa(cmd_SAP_reader_2("U_prg_azs_commessa")).nome_bp,
cmd_SAP_reader_2("codmag_im"),
            cmd_SAP_reader_2("DIV"),
                                      cmd_SAP_reader_2("status"),
                                      cmd_SAP_reader_2("LINENUM"),
                                      cmd_SAP_reader_2("u_produzione"),
                                      cmd_SAP_reader_2("Resname"),
cmd_SAP_reader_2("saldo_imp"))

            WIP += cmd_SAP_reader_2("Trasferito")
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        par_datagridview.ClearSelection()
        Button25.Text = WIP.ToString("0.################")
    End Sub

    Sub ROF()
        If Homepage.ERP_provenienza = "SAP" Then
            Button13.Text = 0
            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()


            Dim CMD_SAP_2 As New SqlCommand
            Dim cmd_SAP_reader_2 As SqlDataReader


            CMD_SAP_2.Connection = Cnn1

            CMD_SAP_2.CommandText = "SELECT  cast(sum(case when T0.[OpenQty] is null then 0 else T0.[OpenQty] end ) as decimal) as 'Openqty' FROM PQT1 T0  INNER JOIN OPQT T1 ON T0.[DocEntry] = T1.[DocEntry] 
    WHERE T0.linestatus='O' and t0.itemcode='" & Codice_SAP & "'
"



            cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

            If cmd_SAP_reader_2.Read() Then

                If Not cmd_SAP_reader_2("Openqty") Is System.DBNull.Value Then
                    Button13.Text = cmd_SAP_reader_2("Openqty")
                Else
                    Button13.Text = 0
                End If


            End If

            cmd_SAP_reader_2.Close()
            Cnn1.Close()
        Else

        End If
    End Sub

    Sub ordinato(par_codice_sap As String, par_datagridview As DataGridView)

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_2.CommandText = "SELECT '' as 'Mag_ord', 'OA' as 'documento',T0.[DocNum] as 'N_documento', T0.[CardName] as 'Fornitore', T1.[OpenQty] as 'Q',T0.[DocDate] as 'Data_R'
, T1.[ShipDate] as 'Data_C'
, t1.u_prg_azs_commessa as 'Commessa'
,t3.U_Final_customer_name
,t2.lastname +' ' +t2.firstname as 'Acquisitore'

 FROM OPOR T0  INNER JOIN POR1 T1 ON T0.[DocEntry] = T1.[DocEntry]
INNER JOIN [TIRELLI_40].[dbo].OHEM T2 ON T0.[OwnerCode] = T2.[empID]
left join oitm t3 on t3.itemcode=t1.u_prg_azs_commessa
 WHERE T1.[OpenQty] >0 and t1.itemcode='" & par_codice_sap & "'

UNION ALL

SELECT '' as 'Mag_ord','ODP', T0.[DocNum], T0.[U_PRODUZIONE], T0.[PlannedQty], T0.[PostDate], T0.[DueDate], T0.[U_PRG_AZS_Commessa] ,t0.U_UTILIZZ, '' FROM OWOR T0 
WHERE (T0.[Status] ='P' or  T0.[Status] ='R') and  T0.[Type] ='S' AND t0.itemcode='" & par_codice_sap & "'"

        Else
            CMD_SAP_2.CommandText = "SELECT *
FROM OPENQUERY(AS400, '
    SELECT t0.doc as documento 
	, t0.numdoc as n_documento
, t0.qta_ord as Q

,DATE(
            SUBSTR(CHAR(t0.data_immissione),1,4) || ''-'' ||
            SUBSTR(CHAR(t0.data_immissione),5,2) || ''-'' ||
            SUBSTR(CHAR(t0.data_immissione),7,2)
        ) AS Data_R

,DATE(
            SUBSTR(CHAR(t0.data_richiesta),1,4) || ''-'' ||
            SUBSTR(CHAR(t0.data_richiesta),5,2) || ''-'' ||
            SUBSTR(CHAR(t0.data_richiesta),7,2)
        ) AS data_C

,trim(t0.matricola) as Commessa
,''Manca'' as U_final_customer_name
,t0.desc_for as fornitore
,''Manca'' as acquisitore
,mag_ord
    FROM TIR90VIS.JGALord t0
    WHERE 
	codart = ''" & par_codice_sap & "'' and evaso <>''S''

     ORDER BY Numdoc DESC 
 
') T10

"

        End If

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader_2("documento"),
                                      cmd_SAP_reader_2("N_documento"),
                                      Math.Round(cmd_SAP_reader_2("Q"), 2),
                                      cmd_SAP_reader_2("Data_R"),
                                      cmd_SAP_reader_2("Data_C"),
                                      cmd_SAP_reader_2("mag_ord"),
                                      cmd_SAP_reader_2("Commessa"),
                                      Business_partner_della_commessa(cmd_SAP_reader_2("Commessa")).nome_bp,
                                      cmd_SAP_reader_2("Fornitore"),
                                      cmd_SAP_reader_2("Acquisitore"))
        Loop





        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        par_datagridview.ClearSelection()
    End Sub

    Sub rt_aperte(par_datagridview As DataGridView, par_codice_sap As String)
        If Homepage.ERP_provenienza = "SAP" Then
            par_datagridview.Rows.Clear()
            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()


            Dim CMD_SAP_2 As New SqlCommand
            Dim cmd_SAP_reader_2 As SqlDataReader


            CMD_SAP_2.Connection = Cnn1



            CMD_SAP_2.CommandText = "select t0.itemcode,t1.docnum,coalesce(t1.comments,'') as 'Osservazioni'
, t1.DocDate, t0.FromWhsCod,t0.WhsCode, t0.Quantity
from WTQ1 t0 inner join owtq t1 on t0.docentry=t1.docentry
where t0.linestatus='O' and t0.itemcode='" & par_codice_sap & "'"

            cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

            Do While cmd_SAP_reader_2.Read()
                par_datagridview.Rows.Add(cmd_SAP_reader_2("docnum"), cmd_SAP_reader_2("docdate"), cmd_SAP_reader_2("FromWhsCod"), cmd_SAP_reader_2("whscode"), cmd_SAP_reader_2("quantity"), cmd_SAP_reader_2("Osservazioni"))
            Loop

            cmd_SAP_reader_2.Close()
            Cnn1.Close()
            par_datagridview.ClearSelection()
        Else

        End If
    End Sub

    Sub rof_aperte(par_datagridview As DataGridView, par_codice_sap As String)

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT  t1.docnum, t0.OpenQty, t1.DocDate,t1.DocDueDate,t0.U_PRG_AZS_Commessa, t1.cardcode,t1.cardname, t2.lastname
FROM PQT1 T0  INNER JOIN OPQT T1 ON T0.[DocEntry] = T1.[DocEntry] 
left JOIN [TIRELLI_40].[dbo].OHEM T2 ON T1.[OwnerCode] = T2.[empID]
    WHERE T0.linestatus='O' and t0.itemcode='" & par_codice_sap & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader_2("docnum"), cmd_SAP_reader_2("OpenQty"), cmd_SAP_reader_2("DocDate"), cmd_SAP_reader_2("DocDueDate"), cmd_SAP_reader_2("U_PRG_AZS_Commessa"), cmd_SAP_reader_2("cardcode"), cmd_SAP_reader_2("cardname"), cmd_SAP_reader_2("Lastname"))
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        par_datagridview.ClearSelection()
    End Sub



    Private Sub Button_confermato_TOT_Click(sender As Object, e As EventArgs)
        trasferito(Codice_SAP, DataGridView_trasferito)
        DataGridView_trasferito.Show()
        DataGridView_ordinato.Hide()
    End Sub


    Private Sub DataGridView_trasferito_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_trasferito.CellClick

        If e.RowIndex >= 0 Then
            riga = e.RowIndex
            Documento = DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="Doc").Value

            If Documento = "ODP" Then
                FORM6.ODP = DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="ODP").Value

            ElseIf Documento = "OC" Then
                docnum_OC = DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="OC").Value
            End If

            Linenum_ODP = DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="Linenum").Value
            Dim commessa_odp As String
            Try
                commessa_odp = DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="Comm").Value
            Catch ex As Exception
                commessa_odp = ""
            End Try

            stato_ODP = DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="stato").Value
            Business_partner_della_commessa(commessa_odp)
            quantita_ODP = Math.Round(DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="Q").Value, 3)
            'check_disponibilità_magazzino()
            'TextBox3.Text = quantita_trasferibile

            If e.ColumnIndex = DataGridView_trasferito.Columns.IndexOf(ODP) And DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="Doc").Value = "ODP" Then





                ODP_Form.docnum_odp = DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="ODP").Value
                ODP_Form.Show()
                ODP_Form.inizializza_form(DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="ODP").Value)




            ElseIf e.ColumnIndex = DataGridView_trasferito.Columns.IndexOf(OC) And DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="Doc").Value = "OC" Then

                Form_nuova_offerta.Show()
                Form_nuova_offerta.TextBox10.Text = DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="OC").Value
                Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                Form_nuova_offerta.inizializzazione_form(DataGridView_trasferito.Rows(e.RowIndex).Cells(columnName:="OC").Value, "ORDR", "rdr1", "ORDINE")


            End If

        End If
    End Sub

    Public Function Business_partner_della_commessa(par_codice_commessa As String)

        Dim codice As New scopri_codice_bp_nome_bp_OITM
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then


            CMD_SAP_2.CommandText = "SELECT coalesce(T0.[u_final_customer_code] ,'') as 'u_final_customer_code'
, coalesce( t1.cardname, '' ) as 'cardname' 
FROM oitm T0 
left join ocrd t1 on t0.u_final_customer_code =t1.cardcode 
WHERE T0.[itemcode] ='" & par_codice_commessa & "'"
        Else
            CMD_SAP_2.CommandText = "SELECT top 
1 trim(t10.matricola) as 'Itemcode', t10.itemname, t10.desc_supp
, T10.DSCLI_FATT as 'Cliente'

      ,  t10.codice_finale as 'Cardname'
		,t10.codice_cliente as 'u_final_customer_code'
       , trim(t10.itemcode) as 'Numero_progetto',
		T10.NAME_progetto AS 'DESC_PROGETTO',
		'' as 'Nome_stato',
        '' as 'Livello_rischio_totale', '' as 'Name',
        t10.pm as 'CODICE_PM'
		,t10.DESC_pm as 'PM'
		, T10.DSNAZ_FINALE as u_country_of_delivery,
        t10.brand AS 'CODICE_BRAND',
		T10.DESC_BRAND AS 'BRAND',
		'' as 'Baia'
		, '' as 'Zona'
		,DATA_CONSEGNA
		,T10.NOME_STATO AS 'STATO_COMMESSA'
FROM OPENQUERY(AS400, '
    SELECT *
    FROM TIR90VIS.JGALCOM t0
    WHERE 
t0.matricola=''" & par_codice_commessa & "'' and t0.matricola<>''''
      
AND LEFT(UPPER(t0.itemcode),2) <> ''TZ''
       
  
ORDER BY t0.matricola DESC

limit 100  
') T10"
        End If
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            codice.codice_BP = cmd_SAP_reader_2("u_final_customer_code")
            codice.nome_BP = cmd_SAP_reader_2("cardname")
        Else
            codice.codice_BP = ""
            codice.nome_BP = ""
        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        Return codice
    End Function

    Sub check_disponibilità_magazzino()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_2.CommandText = "SELECT T0.[ONHAND] 
FROM OITW T0 WHERE T0.[WhsCode] ='" & magazzino_partenza & "' AND  T0.[ItemCode] ='" & Codice_SAP & "'"
        Else
            CMD_SAP_2.CommandText = "SELECT 0 as 'onhand'"
        End If


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            If cmd_SAP_reader_2("ONHAND") >= quantita_ODP Then
                quantita_trasferibile = quantita_ODP
            ElseIf cmd_SAP_reader_2("ONHAND") >= 0 Then
                quantita_trasferibile = Math.Round(cmd_SAP_reader_2("ONHAND"), 3)
            Else
                quantita_trasferibile = 0
            End If
        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    Sub check_quantità_minore_da_trasf(par_documento As String, par_numero_odp As Integer, par_numero_oc As Integer)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        If par_documento = "ODP" Then

            CMD_SAP_2.Connection = Cnn1
            CMD_SAP_2.CommandText = "select t0.u_prg_wip_qtadatrasf as 'Da_trasferire',t0.u_prg_wip_qtaspedita as 'Trasferito' FROM WOR1 T0 WHERE T0.LINENUM=" & Linenum_ODP & " AND T0.DOCENTRY =" & DOCENTRY_documento(par_numero_odp, par_numero_oc, par_documento).Docentryodp & ""

        ElseIf par_documento = "OC" Then

            CMD_SAP_2.Connection = Cnn1
            CMD_SAP_2.CommandText = "SELECT t0.u_trasferito as 'Trasferito', t0.u_datrasferire as 'Da_trasferire' FROM RDR1 T0 WHERE T0.[LineNum] =" & Linenum_ODP & " and t0.docentry=" & DOCENTRY_documento(par_numero_odp, par_numero_oc, par_documento).Docentryoc & ""
        End If
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            Da_trasferire_riga = cmd_SAP_reader_2("Da_trasferire")


            Trasferito_riga = cmd_SAP_reader_2("Trasferito")



        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    Sub check_giacenza()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select t0.onhand from oitw t0 where t0.itemcode= '" & Codice_SAP & "' and t0.whscode='" & magazzino_partenza & "'"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            giacenza = cmd_SAP_reader_2("onhand")


        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub


    Function DOCENTRY_Trasferimenti()
        Dim Max_docentry_traferimenti As Integer
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT MAX(T0.DOCENTRY) as 'Docentry' FROM OWTR T0"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            Max_docentry_traferimenti = cmd_SAP_reader_2("docentry")
        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        Return Max_docentry_traferimenti
    End Function

    Function DOCNUM_Trasferimenti()
        Dim Max_DOCNUM_traferimenti As Integer
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT MAX(T0.docnum) as 'DOCNUM' FROM OWTR T0"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            Max_DOCNUM_traferimenti = cmd_SAP_reader_2("DOCNUM")
        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        Return Max_DOCNUM_traferimenti
    End Function


    Public Function DOCENTRY_documento(par_ODP As String, par_docnum_OC As String, par_documento As String) As scopri_docentry_documento



        Dim docentry As New scopri_docentry_documento()


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        CMD_SAP_2.Connection = Cnn1
        If par_documento = "ODP" Then

            CMD_SAP_2.CommandText = "SELECT T0.DOCENTRY as 'Docentry' FROM OWOR T0 where t0.docnum=" & par_ODP & ""
        Else
            CMD_SAP_2.CommandText = "SELECT T0.DOCENTRY as 'Docentry' FROM ORDR T0 where t0.docnum=" & par_docnum_OC & ""
        End If

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then

            If par_documento = "ODP" Then
                docentry.Docentryodp = cmd_SAP_reader_2("docentry")
                docentry.Docentryoc = 0
            ElseIf par_documento = "OC" Then
                docentry.Docentryoc = cmd_SAP_reader_2("docentry")
                docentry.Docentryodp = 0

            End If

        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        Return docentry
    End Function

    Function Trova_PERIODO_contabile()

        Dim periodo_contabile As String
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.AbsEntry, T0.Code, T0.Name, T0.F_RefDate, T0.T_RefDate, T0.F_DueDate, T0.T_DueDate, T0.F_TaxDate, T0.T_TaxDate, T0.Free2, T0.Free3, T0.DataSource, T0.SubNum, T0.Addition, T0.AddNum, T0.Category, T0.Indicator, T0.UpdateDate, T0.WasStatChd, T0.PeriodStat 
FROM OFPR T0 WHERE T0.F_REFDATE<=GETDATE() AND GETDATE()<=T0.T_REFDATE"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            periodo_contabile = cmd_SAP_reader_2("INDICATOR")
            ''absentry = cmd_SAP_reader_2("ABSENTRY")
        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        Return periodo_contabile

    End Function

    Function trova_absentry()
        Dim absentry As String
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.AbsEntry, T0.Code, T0.Name, T0.F_RefDate, T0.T_RefDate, T0.F_DueDate, T0.T_DueDate, T0.F_TaxDate, T0.T_TaxDate, T0.Free2, T0.Free3, T0.DataSource, T0.SubNum, T0.Addition, T0.AddNum, T0.Category, T0.Indicator, T0.UpdateDate, T0.WasStatChd, T0.PeriodStat 
FROM OFPR T0 WHERE T0.F_REFDATE<=GETDATE() AND GETDATE()<=T0.T_REFDATE"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then

            absentry = cmd_SAP_reader_2("ABSENTRY")
        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        Return absentry

    End Function


    Function Trova_serie_RT()

        Dim serie_rt As String
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "Select top 1 t0.series from owtq t0 order by t0.docentry DESC"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            serie_rt = cmd_SAP_reader_2("series")
        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        Return serie_rt
    End Function

    Public Function Trova_serie_Trasferimento()
        Dim serie_trasferimento As String
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "Select top 1 t0.series from owtr t0 order by t0.docentry DESC"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            serie_trasferimento = cmd_SAP_reader_2("series")
        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        Return serie_trasferimento

    End Function

    Sub Trova_message_id()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select max(messageid) as 'AUTOKEY'
from oilm "

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then

            MESSAGEID = cmd_SAP_reader_2("autokey")
        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    Sub Aggiusta_numeratore_messageid()


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE ONNM SET AUTOKEY=" & MESSAGEID & "+3
FROM ONNM
WHERE OBJECTCODE='10000048'"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub


    'aggiorna_da_trasferire(Documento,quantita_trasferibile,Linenum_ODP)
    Sub aggiorna_da_trasferire(par_Documento As String, par_quantita_trasferibile As String, par_Linenum_ODP As String, par_numero_odp As Integer, par_numero_oc As Integer, par_stringa_trasferimento As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn

        If par_stringa_trasferimento = "Reso" Then
            If par_Documento = "ODP" Then

                If ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).u_produzione = "INT" Then
                    Cmd_SAP.CommandText = "update t0 set t0.u_qta_richiesta_wip= t0.u_qta_richiesta_wip- " & par_quantita_trasferibile & " FROM WOR1 T0 WHERE T0.LINENUM=" & par_Linenum_ODP & " AND T0.DOCENTRY =" & DOCENTRY_documento(par_numero_odp, par_numero_oc, par_Documento).Docentryodp & ""
                Else
                    Cmd_SAP.CommandText = "update t0 set t0.u_prg_wip_qtadatrasf= t0.u_prg_wip_qtadatrasf+ " & par_quantita_trasferibile & ", t0.u_prg_wip_qtaspedita = t0.u_prg_wip_qtaspedita - " & par_quantita_trasferibile & " FROM WOR1 T0 WHERE T0.LINENUM=" & par_Linenum_ODP & " AND T0.DOCENTRY =" & DOCENTRY_documento(par_numero_odp, par_numero_oc, par_Documento).Docentryodp & ""

                End If


            ElseIf par_Documento = "OC" Then
                Cmd_SAP.CommandText = "update t0 set T0.[U_Datrasferire]= T0.[U_Datrasferire]+ " & par_quantita_trasferibile & ", t0.u_trasferito = t0.u_trasferito - " & par_quantita_trasferibile & " FROM RDR1 T0 WHERE T0.LINENUM=" & par_Linenum_ODP & " AND T0.DOCENTRY =" & DOCENTRY_documento(par_numero_odp, par_numero_oc, par_Documento).Docentryoc & ""
            End If
            Cmd_SAP.ExecuteNonQuery()
        Else

            If par_Documento = "ODP" Then

                If ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).u_produzione = "INT" Then
                    Cmd_SAP.CommandText = "update t0 set t0.u_qta_richiesta_wip= t0.u_qta_richiesta_wip+ " & par_quantita_trasferibile & " FROM WOR1 T0 WHERE T0.LINENUM=" & par_Linenum_ODP & " AND T0.DOCENTRY =" & DOCENTRY_documento(par_numero_odp, par_numero_oc, par_Documento).Docentryodp & ""
                Else
                    Cmd_SAP.CommandText = "update t0 set t0.u_prg_wip_qtadatrasf= t0.u_prg_wip_qtadatrasf- " & par_quantita_trasferibile & ", t0.u_prg_wip_qtaspedita = t0.u_prg_wip_qtaspedita + " & par_quantita_trasferibile & " 
FROM WOR1 T0 WHERE T0.LINENUM=" & par_Linenum_ODP & " AND T0.DOCENTRY =" & DOCENTRY_documento(par_numero_odp, par_numero_oc, par_Documento).Docentryodp & ""
                End If

            ElseIf par_Documento = "OC" Then
                Cmd_SAP.CommandText = "update t0 set T0.[U_Datrasferire]= T0.[U_Datrasferire]- " & par_quantita_trasferibile & ", t0.u_trasferito = t0.u_trasferito + " & par_quantita_trasferibile & " FROM RDR1 T0 WHERE T0.LINENUM=" & par_Linenum_ODP & " AND T0.DOCENTRY =" & DOCENTRY_documento(par_numero_odp, par_numero_oc, par_Documento).Docentryoc & ""
            End If

            Cmd_SAP.ExecuteNonQuery()
        End If
        Cnn.Close()

    End Sub

    Sub aggiorna_qta_richiesta_per_wip(par_Documento As String, par_numero_documento As String, par_quantita_trasferibile As String, par_Linenum_ODP As String)





        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn



        If par_Documento = "ODP" Then

            Cmd_SAP.CommandText = "update t0 set t0.U_Qta_richiesta_wip= case when t0.U_Qta_richiesta_wip is null then  0 else t0.U_Qta_richiesta_wip  end+ " & par_quantita_trasferibile & "
FROM WOR1 T0 inner join owor t1 on t0.docentry=t1.docentry WHERE T0.LINENUM=" & par_Linenum_ODP & " AND T1.docnum ='" & par_numero_documento & "'"

        End If

        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub


    'aggiorna_OITW(quantita_trasferibile, magazzino_partenza, magazzino_destinazione, Codice_SAP)
    Sub aggiorna_OITW(par_quantita_trasferibile As String, par_magazzino_partenza As String, par_magazzino_destinazione As String, par_Codice_SAP As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "update t0 set t0.ONHAND= t0.ONHAND - " & par_quantita_trasferibile & " FROM OITW T0 WHERE T0.WHSCODE='" & par_magazzino_partenza & "' AND T0.itemcode ='" & par_Codice_SAP & "'"
        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = "update t0 set t0.ONHAND= t0.ONHAND + " & par_quantita_trasferibile & " FROM OITW T0 WHERE T0.WHSCODE='" & par_magazzino_destinazione & "' AND T0.itemcode ='" & par_Codice_SAP & "'"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub

    Sub aggiorna_NNM1_trasferimento()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "update t0 set t0.nextnumber=" & DOCNUM_Trasferimenti() & "+1 
FROM NNM1 T0 
WHERE T0.[Series] ='" & Trova_serie_Trasferimento() & "' and t0.indicator='" & Trova_PERIODO_contabile() & "' AND T0.OBJECTCODE='67'"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()


    End Sub

    Sub AGGIUSTA_docentry()
        Dim Cnn5 As New SqlConnection
        Cnn5.ConnectionString = Homepage.sap_tirelli
        Cnn5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = Cnn5
        CMD_SAP_5.CommandText = "UPDATE ONNM SET AUTOKEY ='" & DOCENTRY_Trasferimenti() & "'+1 WHERE OBJECTCODE='67'"
        CMD_SAP_5.ExecuteNonQuery()


        Cnn5.Close()

    End Sub

    'metto_wip_nel_magazzino_riga(Documento,magazzino_destinazione,Linenum_ODP)
    Sub metto_wip_nel_magazzino_riga(par_Documento As String, par_magazzino_destinazione As String, par_Linenum_ODP As String, par_numero_odp As Integer, par_numero_oc As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        If par_Documento = "ODP" Then

            If ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).u_produzione = "INT" Then

            Else
                Cmd_SAP.CommandText = "update t0 set T0.[wareHouse]= '" & par_magazzino_destinazione & "' 
FROM WOR1 T0 WHERE T0.LINENUM=" & par_Linenum_ODP & " AND T0.DOCENTRY =" & DOCENTRY_documento(par_numero_odp, par_numero_oc, par_Documento).Docentryodp & " "
            End If

        ElseIf par_Documento = "OC" Then
            Cmd_SAP.CommandText = "update t0 set T0.[whscode]= '" & par_magazzino_destinazione & "' FROM rdr1 T0 WHERE T0.LINENUM=" & par_Linenum_ODP & " AND T0.DOCENTRY =" & DOCENTRY_documento(par_numero_odp, par_numero_oc, par_Documento).Docentryoc & "  "
        End If
        If par_Documento = "ODP" Then

            If ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).u_produzione = "INT" Then
            Else
                Cmd_SAP.ExecuteNonQuery()
            End If

        Else
            Cmd_SAP.ExecuteNonQuery()
        End If


        Cnn.Close()

    End Sub

    Sub NEW_OILM()



        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "-- Dichiarazioni delle variabili
DECLARE @max_autokey INT
DECLARE @max_messageid INT

-- Calcola i valori massimi
SELECT @max_autokey = MAX(TransSeq) FROM OIVL
SELECT @max_messageid = MAX(MessageID) FROM OILM

-- Elimina la tabella temporanea
DROP TABLE #TempTable
-- Crea una tabella temporanea
CREATE TABLE #TempTable (
    TransSeq INT,
    MessageID INT,
    DocEntry INT,
	quantity DECIMAL,
	inqty DECIMAL,
	outqty DECIMAL,
	whscode VARCHAR(255),
	docdate DATE,
	itemcode VARCHAR(255),
	cardcode VARCHAR(255),
	stringa_trasferimento VARCHAR(255),
	documento VARCHAR(255),
	N_documento INT,
	commessa VARCHAR(255),
	price DECIMAL,
	dscription VARCHAR(255),
	docnum INT,
	usersign INT,
	orario_Stringa varchar(255)
)

-- Inserisci i dati nella tabella temporanea
INSERT INTO #TempTable (TransSeq, MessageID, DocEntry, quantity, inqty, outqty, whscode, docdate, itemcode, cardcode, stringa_trasferimento, documento, N_documento, commessa, price, dscription, docnum, usersign,orario_stringa)
SELECT 
    @max_autokey + 1 AS TransSeq,
    @max_messageid + 1 AS MessageID,
    t10.docentry,
	t10.quantity,
	t10.inqty,
	t10.outqty,
	t10.whscode,
	t10.docdate,
	t10.itemcode,
	t10.cardcode,
	t10.stringa_trasferimento,
	t10.documento, 
	t10.N_documento,
	t10.commessa,
	t10.price,
	t10.dscription,
	t10.docnum,
	t10.usersign,
	CONCAT(DATEPART(HOUR, t10.docdate),DATEPART(MINUTE, t10.docdate))



FROM 
    (
     
select '+' as 'Movimento', t3.docnum,t3.cardcode, t1.docentry,t3.docdate,t0.itemcode, t0.onhand, t1.whscode,t1.quantity as 'inqty',0 as 'outqty',t0.onhand+t1.quantity as 'Result',  t2.messageid
,case when t1.fromwhscod ='WIP' or t1.fromwhscod ='BWIP' then 'Reso' else 'Trasferimento' end as 'Stringa_trasferimento'
, case when T1.[U_PRG_AZS_OpDocEntry] > 0 then 'ODP' 
when T1.[U_PRG_AZS_OcDocEntry] > 0 then 'OC'
else''
end as 'Documento'

, case when  T1.[U_PRG_AZS_OpDocNum] > 0 then  T1.[U_PRG_AZS_OpDocNum]
when  T1.[U_PRG_AZS_OcDocNum] > 0 then  T1.[U_PRG_AZS_OcDocNum]
else''
end as 'N_Documento'
,t4.price
,coalesce(t5.u_prg_azs_commessa,'') as 'Commessa'
,t1.dscription
,t3.usersign
,t1.quantity


from oitw t0
inner join wtr1 t1 on t0.itemcode=t1.itemcode and (t1.whscode=t0.whscode)
left join oilm t2 on t1.docentry=t2.docentry 
AND t1.whscode=t2.loccode
left join owtr t3 on t3.docentry=t1.docentry
inner join itm1 t4 on t4.itemcode =t1.itemcode and t4.pricelist=2
left join owor t5 on t5.docentry=T1.[U_PRG_AZS_OpDocEntry] and t5.docentry>0


where t0.onhand<0 and 
t2.messageid is null

union all 
select '-', t3.docnum,t3.cardcode, t1.docentry,t3.docdate,t0.itemcode, t4.onhand, t1.fromwhscod,0 as 'inqty',t1.quantity as 'outqty',t4.onhand-t1.quantity as 'Result',  t2.messageid
,case when t1.fromwhscod ='WIP' or t1.fromwhscod ='BWIP' then 'Reso' else 'Trasferimento' end as 'Stringa_trasferimento'
, case when T1.[U_PRG_AZS_OpDocEntry] > 0 then 'ODP' 
when T1.[U_PRG_AZS_OcDocEntry] > 0 then 'OC'
else''
end as 'Documento'
, case when  T1.[U_PRG_AZS_OpDocNum] > 0 then  T1.[U_PRG_AZS_OpDocNum]
when  T1.[U_PRG_AZS_OcDocNum] > 0 then  T1.[U_PRG_AZS_OcDocNum]
else''
end as 'N_Documento'
,t5.price
,coalesce(t6.u_prg_azs_commessa,'') as 'Commessa'
,t1.dscription
,t3.usersign
,t1.quantity
from oitw t0
inner join wtr1 t1 on t0.itemcode=t1.itemcode and (t1.whscode=t0.whscode)
left join oilm t2 on t1.docentry=t2.docentry AND t1.fromwhscod=t2.loccode
left join owtr t3 on t3.docentry=t1.docentry
inner join oitw t4 on t4.itemcode=t1.itemcode and (t4.whscode=t1.fromwhscod)
inner join itm1 t5 on t5.itemcode =t1.itemcode and t5.pricelist=2
left join owor t6 on t6.docentry=T1.[U_PRG_AZS_OpDocEntry] and t6.docentry>0
where t0.onhand<0 and 
t2.messageid is null
)

     AS t10

---- Esegui gli inserimenti utilizzando i valori dalla tabella temporanea
INSERT INTO OIVL (OIVL.TransType, OIVL.CreatedBy, OIVL.BASE_REF, OIVL.DocLineNum, OIVL.DocDate, OIVL.CreateTime, OIVL.ItemCode, OIVL.InQty, OIVL.OutQty, OIVL.Price, OIVL.Currency, OIVL.Rate, OIVL.TrnsfrAct, OIVL.PriceDifAc, OIVL.VarianceAc, OIVL.ReturnAct, OIVL.ExcRateAct, OIVL.ClearAct, OIVL.CostAct, OIVL.WipAct, OIVL.OpenStock, OIVL.CreateDate, OIVL.PriceDiff,OIVL.TransSeq, OIVL.InvntAct, OIVL.SubLineNum, OIVL.AppObjLine, OIVL.Expenses, OIVL.OpenExp ,OIVL.Allocation, OIVL.OpenAlloc,OIVL.ExpAlloc, OIVL.OExpAlloc, OIVL.OpenPDiff ,OIVL.ExchDiff, OIVL.OpenEDiff, OIVL.NegInvAdjs, OIVL.OpenNegInv, OIVL.NegStckAct, OIVL.BTransVal, OIVL.VarVal, OIVL.BExpVal, OIVL.CogsVal, OIVL.BNegAVal, OIVL.IOffIncAcc, OIVL.IOffIncVal, OIVL.DOffDecAcc, OIVL.DOffDecVal, OIVL.DecAcc ,OIVL.DecVal, OIVL.WipVal, OIVL.WipVarAcc, OIVL.WipVarVal, OIVL.IncAct, OIVL.IncVal, OIVL.ExpCAcc, OIVL.CostMethod, OIVL.MessageID , OIVL.LocType, OIVL.LocCode, OIVL.PostStatus, OIVL.SumStock, OIVL.OpenCogs, OIVL.OpenQty, OIVL.TreeID, OIVL.ParentID, OIVL.PAOffAcc, OIVL.PAOffVal, OIVL.OpenPAOff, OIVL.PAAcc, OIVL.PAVal, OIVL.OpenPA, OIVL.LinkArc, OIVL.VersionNum, OIVL.BSubLineNo, OIVL.WipDebCred,oivl.usersign)

select top 1'67',t10.docentry, t10.docnum,'0',t10.docdate,t10.orario_stringa,t10.itemcode,t10.inqty,t10.outqty,t10.price,'EUR','0','','','','','','','','','0',t10.docdATE,0,@max_Autokey+1,'','-1','-1',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'',' ',@max_messageid+1,'64',t10.whscode,'N','0',0,'0',@max_autokey+1,'-1','',0,'0','',0,'0','N','10.00.140.04','-1','',t10.usersign
FROM 
    #TempTable t10

INSERT INTO OILM (OILM.MessageID, OILM.DocEntry, OILM.TransType, OILM.DocLineNum, OILM.Quantity, OILM.EffectQty, OILM.LocType, OILM.LocCode, OILM.TotalLC, OILM.TotalFC, OILM.TotalSC, OILM.BaseAbsEnt, OILM.BaseType, OILM.BaseCurr, OILM.Currency, OILM.AccumType, OILM.ActionType, OILM.ExpensesLC, OILM.ExpensesFC, OILM.ExpensesSC, OILM.DocDueDate, OILM.ItemCode, OILM.BPCardCode, OILM.DocDate, OILM.DocRate, OILM.Comment, OILM.JrnlMemo, OILM.Ref1, OILM.Ref2, OILM.BaseLine, OILM.SnBType, OILM.CreateTime, OILM.DataSource, OILM.CreateDate, OILM.OcrCode, OILM.OcrCode2, OILM.OcrCode3, OILM.OcrCode4, OILM.OcrCode5, OILM.DocPrice, OILM.CardName, OILM.Dscription, OILM.TreeType, OILM.ApplObj, OILM.AppObjAbs, OILM.AppObjType, OILM.AppObjLine, OILM.BASE_REF, OILM.TransSeqRf, OILM.LayerIDRef, OILM.VersionNum, OILM.PriceRate, OILM.PriceCurr, OILM.DocTotal, OILM.Price, OILM.CIShbQty, OILM.SubLineNum, OILM.PrjCode, OILM.SlpCode, OILM.TaxDate, OILM.UseDocPric, OILM.VendorNum, OILM.SerialNum, OILM.BlockNum, OILM.ImportLog, OILM.Location, OILM.DocPrcRate, OILM.DocPrcCurr, OILM.CgsOcrCod, OILM.CgsOcrCod2, OILM.CgsOcrCod3, OILM.CgsOcrCod4, OILM.CgsOcrCod5, OILM.BSubLineNo, OILM.AppSubLine, OILM.SysRate, OILM.ExFromRpt, OILM.Ref3, OILM.EnSetCost, OILM.RetCost, OILM.DocAction, OILM.UseShpdGd, OILM.AddTotalLC, OILM.AddExpLC, OILM.IsNegLnQty, OILM.StgSeqNum, OILM.StgEntry, OILM.StgDesc, oilm.usersign ) 

select top 1 @max_messageid+1,t10.docentry,'67','0',t10.quantity,t10.quantity,'64',t10.whscode,'0','0','0',t10.docentry,'67','','EUR','1','1','0','0','0',t10.docdate,t10.itemcode,t10.cardcode,t10.docdate,'0',concat(t10.stringa_trasferimento,' ',t10.whscode,' ', t10.documento , t10.N_documento,' ',t10.commessa ),concat(t10.stringa_trasferimento,' ',t10.whscode,'-'),t10.docnum,t10.commessa,'0','4',t10.orario_stringa,'I',t10.docdate,'','','','','',t10.price,t10.cardcode,t10.dscription,'N','-1','-1','','-1',t10.docnum,'-1','-1','10.00.140.04','0','EUR',t10.price,t10.price,'0','-1','','-1',t10.docdate,'N','','','','','','0','EUR','','','','','','-1','-1','1','N','','N','0','1','N','0','0','N','0','0','', t10.usersign

FROM 
    #TempTable t10

-- Elimina la tabella temporanea
DROP TABLE #TempTable

-- Aggiorna il valore AUTOKEY in ONNM
UPDATE ONNM SET AUTOKEY = @max_messageid + 2
WHERE OBJECTCODE = '10000048'"

        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub

    Sub Trova_NUMERATORE_OIVL()

        Dim Cnn1 As New SqlConnection


        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select max(OIVL.TransSeq) as 'Autokey'
from oivl"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            Numeratore_OIVL = cmd_SAP_reader_2("AUTOKEY")
        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub


    'OIVL(par_Codice_SAP,par_quantita_trasferibile,par_magazzino_partenza,par_magazzino_destinazione)
    Sub OIVL_IVL1_OIVK(par_Codice_SAP As String, par_quantita_trasferibile As String, par_magazzino_partenza As String, par_magazzino_destinazione As String, par_utente_sap As String, par_prezzo_listino_acquisto As String)

        par_prezzo_listino_acquisto = Replace(par_prezzo_listino_acquisto, ",", ".")
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO OIVL (OIVL.TransType, OIVL.CreatedBy, OIVL.BASE_REF, OIVL.DocLineNum, OIVL.DocDate, OIVL.CreateTime, OIVL.ItemCode, OIVL.InQty, OIVL.OutQty, OIVL.Price, OIVL.Currency, OIVL.Rate, OIVL.TrnsfrAct, OIVL.PriceDifAc, OIVL.VarianceAc, OIVL.ReturnAct, OIVL.ExcRateAct, OIVL.ClearAct, OIVL.CostAct, OIVL.WipAct, OIVL.OpenStock, OIVL.CreateDate, OIVL.PriceDiff,OIVL.TransSeq, OIVL.InvntAct, OIVL.SubLineNum, OIVL.AppObjLine, OIVL.Expenses, OIVL.OpenExp ,OIVL.Allocation, OIVL.OpenAlloc,OIVL.ExpAlloc, OIVL.OExpAlloc, OIVL.OpenPDiff ,OIVL.ExchDiff, OIVL.OpenEDiff, OIVL.NegInvAdjs, OIVL.OpenNegInv, OIVL.NegStckAct, OIVL.BTransVal, OIVL.VarVal, OIVL.BExpVal, OIVL.CogsVal, OIVL.BNegAVal, OIVL.IOffIncAcc, OIVL.IOffIncVal, OIVL.DOffDecAcc, OIVL.DOffDecVal, OIVL.DecAcc ,OIVL.DecVal, OIVL.WipVal, OIVL.WipVarAcc, OIVL.WipVarVal, OIVL.IncAct, OIVL.IncVal, OIVL.ExpCAcc, OIVL.CostMethod, OIVL.MessageID , OIVL.LocType, OIVL.LocCode, OIVL.PostStatus, OIVL.SumStock, OIVL.OpenCogs, OIVL.OpenQty, OIVL.TreeID, OIVL.ParentID, OIVL.PAOffAcc, OIVL.PAOffVal, OIVL.OpenPAOff, OIVL.PAAcc, OIVL.PAVal, OIVL.OpenPA, OIVL.LinkArc, OIVL.VersionNum, OIVL.BSubLineNo, OIVL.WipDebCred,oivl.usersign)

VALUES
('67'," & DOCENTRY_Trasferimenti() & "," & DOCNUM_Trasferimenti() & ",'0',CONVERT(date, GETDATE()),CONCAT(DATEPART(HOUR, GETDATE()),DATEPART(MINUTE, GETDATE())),'" & par_Codice_SAP & "','0','" & par_quantita_trasferibile & "','" & par_prezzo_listino_acquisto & "','EUR','0','','','','','','','','','0',CONVERT(date, GETDATE()),0,'" & Numeratore_OIVL & "'+1,'','-1','-1',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'',' ','" & MESSAGEID & "'+1,'64','" & par_magazzino_partenza & "','N','0',0,'0','" & Numeratore_OIVL & "'+1,'-1','',0,'0','',0,'0','N','10.00.140.04','-1','','" & par_utente_sap & "')"
        Cmd_SAP.ExecuteNonQuery()


        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO OIVL (OIVL.TransType, OIVL.CreatedBy, OIVL.BASE_REF, OIVL.DocLineNum, OIVL.DocDate, OIVL.CreateTime, OIVL.ItemCode, OIVL.InQty, OIVL.OutQty, OIVL.Price, OIVL.Currency, OIVL.Rate, OIVL.TrnsfrAct, OIVL.PriceDifAc, OIVL.VarianceAc, OIVL.ReturnAct, OIVL.ExcRateAct, OIVL.ClearAct, OIVL.CostAct, OIVL.WipAct, OIVL.OpenStock, OIVL.CreateDate, OIVL.PriceDiff,OIVL.TransSeq, OIVL.InvntAct, OIVL.SubLineNum, OIVL.AppObjLine, OIVL.Expenses, OIVL.OpenExp ,OIVL.Allocation, OIVL.OpenAlloc,OIVL.ExpAlloc, OIVL.OExpAlloc, OIVL.OpenPDiff ,OIVL.ExchDiff, OIVL.OpenEDiff, OIVL.NegInvAdjs, OIVL.OpenNegInv, OIVL.NegStckAct, OIVL.BTransVal, OIVL.VarVal, OIVL.BExpVal, OIVL.CogsVal, OIVL.BNegAVal, OIVL.IOffIncAcc, OIVL.IOffIncVal, OIVL.DOffDecAcc, OIVL.DOffDecVal, OIVL.DecAcc ,OIVL.DecVal, OIVL.WipVal, OIVL.WipVarAcc, OIVL.WipVarVal, OIVL.IncAct, OIVL.IncVal, OIVL.ExpCAcc, OIVL.CostMethod, OIVL.MessageID , OIVL.LocType, OIVL.LocCode, OIVL.PostStatus, OIVL.SumStock, OIVL.OpenCogs, OIVL.OpenQty, OIVL.TreeID, OIVL.ParentID, OIVL.PAOffAcc, OIVL.PAOffVal, OIVL.OpenPAOff, OIVL.PAAcc, OIVL.PAVal, OIVL.OpenPA, OIVL.LinkArc, OIVL.VersionNum, OIVL.BSubLineNo, OIVL.WipDebCred,oivl.usersign)

VALUES
('67'," & DOCENTRY_Trasferimenti() & "," & DOCNUM_Trasferimenti() & ",'0',CONVERT(date, GETDATE()),CONCAT(DATEPART(HOUR, GETDATE()),DATEPART(MINUTE, GETDATE())),'" & par_Codice_SAP & "','" & par_quantita_trasferibile & "','0','" & par_prezzo_listino_acquisto & "','EUR','0','','','','','','','','','0',CONVERT(date, GETDATE()),0," & Numeratore_OIVL & "+2,'','-1','-1',0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,'',' ','" & MESSAGEID & "'+1,'64','" & par_magazzino_destinazione & "','N','0',0,'0','2947807','-1','',0,'0','',0,'0','N','10.00.140.04','-1','','" & par_utente_sap & "')"
        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO IVL1 (IVL1.TransSeq, IVL1.LayerID, IVL1.CalcPrice, IVL1.Balance, IVL1.TransValue, IVL1.LayerInQty, IVL1.LayerOutQ, IVL1.RevalTotal) 
VALUES ('" & Numeratore_OIVL & "'+1,0,0,0,0,0,'" & par_quantita_trasferibile & "',0)"
        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO IVL1 (IVL1.TransSeq, IVL1.LayerID, IVL1.CalcPrice, IVL1.Balance, IVL1.TransValue, IVL1.LayerInQty, IVL1.LayerOutQ, IVL1.RevalTotal) 
VALUES ('" & Numeratore_OIVL & "'+2,0,0,0,0,'" & par_quantita_trasferibile & "',0,0)"
        Cmd_SAP.ExecuteNonQuery()


        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO OIVK ([TransSeq],[LayerID],[RootID],[TransNum],[Instance],[INMTransSe]) VALUES ('" & Numeratore_OIVL & "'+1,0,-1,'" & Numeratore_OIVL & "'+1,0,'" & Numeratore_OIVL & "'+1)"
        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO OIVK ([TransSeq],[LayerID],[RootID],[TransNum],[Instance],[INMTransSe]) VALUES ('" & Numeratore_OIVL & "'+2,0,-1,'" & Numeratore_OIVL & "'+2,0,'" & Numeratore_OIVL & "'+2)"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub


    Sub IVL1(par_quantita_trasferibile As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO IVL1 (IVL1.TransSeq, IVL1.LayerID, IVL1.CalcPrice, IVL1.Balance, IVL1.TransValue, IVL1.LayerInQty, IVL1.LayerOutQ, IVL1.RevalTotal) 
VALUES ('" & Numeratore_OIVL & "'+1,0,0,0,0,0,'" & par_quantita_trasferibile & "',0)"
        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO IVL1 (IVL1.TransSeq, IVL1.LayerID, IVL1.CalcPrice, IVL1.Balance, IVL1.TransValue, IVL1.LayerInQty, IVL1.LayerOutQ, IVL1.RevalTotal) 
VALUES ('" & Numeratore_OIVL & "'+2,0,0,0,0,'" & par_quantita_trasferibile & "',0,0)"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub

    Sub OIVK()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO OIVK ([TransSeq],[LayerID],[RootID],[TransNum],[Instance],[INMTransSe]) VALUES ('" & Numeratore_OIVL & "'+1,0,-1,'" & Numeratore_OIVL & "'+1,0,'" & Numeratore_OIVL & "'+1)"
        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO OIVK ([TransSeq],[LayerID],[RootID],[TransNum],[Instance],[INMTransSe]) VALUES ('" & Numeratore_OIVL & "'+2,0,-1,'" & Numeratore_OIVL & "'+2,0,'" & Numeratore_OIVL & "'+2)"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub

    'AWOR(FORM6.ODP,Linenum_ODP)
    Sub AWOR(par_numero_ODP As String, par_Linenum_ODP As String, par_utente_sap As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO AWOR (AWOR.DOCENTRY,AWOR.UPDATEDATE,AWOR.LOGINSTANC, AWOR.DocNum, AWOR.Series, AWOR.ItemCode, AWOR.Status, AWOR.Type, AWOR.PlannedQty, AWOR.CmpltQty, AWOR.RjctQty, AWOR.PostDate, AWOR.DueDate, AWOR.OriginAbs, AWOR.OriginNum, AWOR.OriginType, AWOR.UserSign, AWOR.Comments, AWOR.CloseDate, AWOR.RlsDate, AWOR.CardCode, AWOR.Warehouse, AWOR.Uom, AWOR.LineDirty, AWOR.JrnlMemo, AWOR.TransId, AWOR.CreateDate, AWOR.Printed, AWOR.OcrCode, AWOR.PIndicator, AWOR.OcrCode2, AWOR.OcrCode3, AWOR.OcrCode4, AWOR.OcrCode5, AWOR.SeqCode, AWOR.Serial, AWOR.SeriesStr, AWOR.SubStr, AWOR.Project, AWOR.SupplCode, AWOR.UomEntry, AWOR.PickRmrk, AWOR.SysCloseDt, AWOR.SysCloseTm, AWOR.CloseVerNm, AWOR.StartDate, AWOR.ObjType, AWOR.ProdName, AWOR.Priority, AWOR.RouDatCalc, AWOR.UpdAlloc, AWOR.CreateTS, AWOR.UpdateTS, AWOR.VersionNum, AWOR.AtcEntry, AWOR.AsChild, AWOR.LinkToObj, AWOR.ProcItms, AWOR.U_UTILIZZ, AWOR.U_PRODUZIONE, AWOR.U_Totcosto, AWOR.U_PRG_AZS_Terzista, AWOR.U_PRG_AZS_RdrLineNum, AWOR.U_PRG_AZS_FromDate, AWOR.U_PRG_AZS_FromHour, AWOR.U_PRG_CLV_Fattibil, AWOR.U_PRG_TOTCOSTO, AWOR.U_STAMPATO, AWOR.U_MATRIC, AWOR.U_PRG_AZS_Commessa, AWOR.U_Primadatadiconsegna, AWOR.U_Consumomediomensile, AWOR.U_Permag, AWOR.U_LPONE, AWOR.U_Inventario, AWOR.U_ODPPadre, AWOR.U_Distintabase, AWOR.U_Aggiornaprezzo, AWOR.U_Collaudatore, AWOR.U_Elettrico, AWOR.U_Assemblatore, AWOR.U_Lavorazione, AWOR.U_Lavorazione_in_corso, AWOR.U_Lavoratore, AWOR.U_Data_ora_inizio, AWOR.U_Data_ora_fine, AWOR.U_Disegno, AWOR.U_Fase, AWOR.U_Stato, AWOR.U_PRG_AZS_CreatedBy, AWOR.U_PRG_WMS_Exp, AWOR.U_PRG_WMS_ExpDate, AWOR.U_Data_cons_MES, AWOR.U_Priorita_MES)

SELECT t0.DOCENTRY,GETDATE(), MAX(T1.LOGINSTANC)+1, T0.DocNum, T0.Series, T0.ItemCode, T0.Status, T0.Type, T0.PlannedQty, T0.CmpltQty, T0.RjctQty, T0.PostDate, T0.DueDate, T0.OriginAbs, T0.OriginNum, T0.OriginType, '" & par_utente_sap & "', T0.Comments, T0.CloseDate, T0.RlsDate, T0.CardCode, T0.Warehouse, T0.Uom, '" & par_Linenum_ODP & "', T0.JrnlMemo, T0.TransId, T0.CreateDate, T0.Printed, T0.OcrCode, T0.PIndicator, T0.OcrCode2, T0.OcrCode3, T0.OcrCode4, T0.OcrCode5, T0.SeqCode, T0.Serial, T0.SeriesStr, T0.SubStr, T0.Project, T0.SupplCode, T0.UomEntry, T0.PickRmrk, T0.SysCloseDt, T0.SysCloseTm, T0.CloseVerNm, T0.StartDate, T0.ObjType, T0.ProdName, T0.Priority, T0.RouDatCalc, T0.UpdAlloc, T0.CreateTS, CONCAT(DATEPART(HOUR,GETDATE()),DATEPART(MINUTE,GETDATE()),DATEPART(SECOND,GETDATE())), T0.VersionNum, T0.AtcEntry, T0.AsChild, T0.LinkToObj, T0.ProcItms, T0.U_UTILIZZ, T0.U_PRODUZIONE, T0.U_Totcosto, T0.U_PRG_AZS_Terzista, T0.U_PRG_AZS_RdrLineNum, T0.U_PRG_AZS_FromDate, T0.U_PRG_AZS_FromHour, T0.U_PRG_CLV_Fattibil, T0.U_PRG_TOTCOSTO, T0.U_STAMPATO, T0.U_MATRIC, T0.U_PRG_AZS_Commessa, T0.U_Primadatadiconsegna, T0.U_Consumomediomensile, T0.U_Permag, T0.U_LPONE, T0.U_Inventario, T0.U_ODPPadre, T0.U_Distintabase, T0.U_Aggiornaprezzo, T0.U_Collaudatore, T0.U_Elettrico, T0.U_Assemblatore, T0.U_Lavorazione, T0.U_Lavorazione_in_corso, T0.U_Lavoratore, T0.U_Data_ora_inizio, T0.U_Data_ora_fine, T0.U_Disegno, T0.U_Fase, T0.U_Stato, T0.U_PRG_AZS_CreatedBy, T0.U_PRG_WMS_Exp, T0.U_PRG_WMS_ExpDate, T0.U_Data_cons_MES, T0.U_Priorita_MES FROM OWOR T0 INNER JOIN AWOR T1 ON T0.DOCENTRY=T1.DOCENTRY 
WHERE T0.DOCNUM='" & par_numero_ODP & "' GROUP BY t0.DOCENTRY, T0.DocNum, T0.Series, T0.ItemCode, T0.Status, T0.Type, T0.PlannedQty, T0.CmpltQty, T0.RjctQty, T0.PostDate, T0.DueDate, T0.OriginAbs, T0.OriginNum, T0.OriginType, T0.UserSign, T0.Comments, T0.CloseDate, T0.RlsDate, T0.CardCode, T0.Warehouse, T0.Uom, T0.LineDirty, T0.JrnlMemo, T0.TransId, T0.CreateDate, T0.Printed, T0.OcrCode, T0.PIndicator, T0.OcrCode2, T0.OcrCode3, T0.OcrCode4, T0.OcrCode5, T0.SeqCode, T0.Serial, T0.SeriesStr, T0.SubStr, T0.Project, T0.SupplCode, T0.UomEntry, T0.PickRmrk, T0.SysCloseDt, T0.SysCloseTm, T0.CloseVerNm, T0.StartDate, T0.ObjType, T0.ProdName, T0.Priority, T0.RouDatCalc, T0.UpdAlloc, T0.CreateTS, T0.UpdateTS, T0.VersionNum, T0.AtcEntry, T0.AsChild, T0.LinkToObj, T0.ProcItms, T0.U_UTILIZZ, T0.U_PRODUZIONE, T0.U_Totcosto, T0.U_PRG_AZS_Terzista, T0.U_PRG_AZS_RdrLineNum, T0.U_PRG_AZS_FromDate, T0.U_PRG_AZS_FromHour, T0.U_PRG_CLV_Fattibil, T0.U_PRG_TOTCOSTO, T0.U_STAMPATO, T0.U_MATRIC, T0.U_PRG_AZS_Commessa, T0.U_Primadatadiconsegna, T0.U_Consumomediomensile, T0.U_Permag, T0.U_LPONE, T0.U_Inventario, T0.U_ODPPadre, T0.U_Distintabase, T0.U_Aggiornaprezzo, T0.U_Collaudatore, T0.U_Elettrico, T0.U_Assemblatore, T0.U_Lavorazione, T0.U_Lavorazione_in_corso, T0.U_Lavoratore, T0.U_Data_ora_inizio, T0.U_Data_ora_fine, T0.U_Disegno, T0.U_Fase, T0.U_Stato, T0.U_PRG_AZS_CreatedBy, T0.U_PRG_WMS_Exp, T0.U_PRG_WMS_ExpDate, T0.U_Data_cons_MES, T0.U_Priorita_MES"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub


    Sub AWOR_NEW(par_utente_sap As String, par_numero_odp As Integer, par_documento As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE T11 SET T11.UPDATEDATE=GETDATE(), USERSIGN='" & par_utente_sap & "'
FROM
(
SELECT  MAX(T0.LOGINSTANC) AS 'MAX',MAX(T0.LOGINSTANC)-1 AS 'MAX-1', T0.DOCENTRY
FROM AWOR T0
WHERE T0.DOCENTRY='" & DOCENTRY_documento(par_numero_odp, 0, par_documento).Docentryodp & "'
GROUP BY T0.DOCENTRY
)
AS T10 INNER JOIN AWOR T11 ON T11.DOCENTRY=T10.DOCENTRY AND (T11.LOGINSTANC=T10.MAX OR T11.LOGINSTANC=T10.MAX-1)
"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub

    'OILM(Documento,Codice_SAP,quantita_trasferibile,magazzino_partenza,magazzino_destinazione,stringa_trasferimento,commessa_ODP)
    Sub OILM(par_Documento As String, par_Codice_SAP As String, par_quantita_trasferibile As String, par_magazzino_partenza As String, par_magazzino_destinazione As String, par_stringa_trasferimento As String, par_commessa_ODP As String, par_utente_sap As String, par_prezzo_listino_acquisto As String, par_descrizione As String, par_ref1 As String)

        Dim Cnn As New SqlConnection
        par_descrizione = Replace(par_descrizione, "'", " ")
        par_prezzo_listino_acquisto = Replace(par_prezzo_listino_acquisto, ",", ".")
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO OILM (OILM.MessageID, OILM.DocEntry, OILM.TransType, OILM.DocLineNum, OILM.Quantity, OILM.EffectQty, OILM.LocType, OILM.LocCode, OILM.TotalLC, OILM.TotalFC, OILM.TotalSC, OILM.BaseAbsEnt, OILM.BaseType, OILM.BaseCurr, OILM.Currency, OILM.AccumType, OILM.ActionType, OILM.ExpensesLC, OILM.ExpensesFC, OILM.ExpensesSC, OILM.DocDueDate, OILM.ItemCode, OILM.BPCardCode, OILM.DocDate, OILM.DocRate, OILM.Comment, OILM.JrnlMemo, OILM.Ref1, OILM.Ref2, OILM.BaseLine, OILM.SnBType, OILM.CreateTime, OILM.DataSource, OILM.CreateDate, OILM.OcrCode, OILM.OcrCode2, OILM.OcrCode3, OILM.OcrCode4, OILM.OcrCode5, OILM.DocPrice, OILM.CardName, OILM.Dscription, OILM.TreeType, OILM.ApplObj, OILM.AppObjAbs, OILM.AppObjType, OILM.AppObjLine, OILM.BASE_REF, OILM.TransSeqRf, OILM.LayerIDRef, OILM.VersionNum, OILM.PriceRate, OILM.PriceCurr, OILM.DocTotal, OILM.Price, OILM.CIShbQty, OILM.SubLineNum, OILM.PrjCode, OILM.SlpCode, OILM.TaxDate, OILM.UseDocPric, OILM.VendorNum, OILM.SerialNum, OILM.BlockNum, OILM.ImportLog, OILM.Location, OILM.DocPrcRate, OILM.DocPrcCurr, OILM.CgsOcrCod, OILM.CgsOcrCod2, OILM.CgsOcrCod3, OILM.CgsOcrCod4, OILM.CgsOcrCod5, OILM.BSubLineNo, OILM.AppSubLine, OILM.SysRate, OILM.ExFromRpt, OILM.Ref3, OILM.EnSetCost, OILM.RetCost, OILM.DocAction, OILM.UseShpdGd, OILM.AddTotalLC, OILM.AddExpLC, OILM.IsNegLnQty, OILM.StgSeqNum, OILM.StgEntry, OILM.StgDesc, oilm.usersign ) VALUES
('" & MESSAGEID & "'+1,'" & DOCENTRY_Trasferimenti() & "','67','0'," & par_quantita_trasferibile & "," & par_quantita_trasferibile & ",'64','" & par_magazzino_destinazione & "','0','0','0','" & DOCENTRY_Trasferimenti() & "','67','','EUR','1','1','0','0','0',GETDATE(),'" & par_Codice_SAP & "','" & Business_partner_della_commessa(par_commessa_ODP).codice_bp & "',convert(date,GETDATE()),'0','" & par_stringa_trasferimento & " " & par_magazzino_destinazione & " " & par_Documento & par_ref1 & " " & par_commessa_ODP & "','" & par_stringa_trasferimento & " magazzino -'," & DOCNUM_Trasferimenti() & ",'" & par_commessa_ODP & "','0','4',CONCAT(DATEPART(HOUR, GETDATE()),DATEPART(MINUTE, GETDATE())),'I',GETDATE(),'','','','','','" & par_prezzo_listino_acquisto & "','" & Business_partner_della_commessa(par_commessa_ODP).nome_bp & "','" & par_descrizione & "','N','-1','-1','','-1'," & DOCNUM_Trasferimenti() & ",'-1','-1','10.00.140.04','0','EUR','" & par_prezzo_listino_acquisto & "','" & par_prezzo_listino_acquisto & "','0','-1','','-1',convert(date,GETDATE()),'N','','','','','','0','EUR','','','','','','-1','-1','1','N','','N','0','1','N','0','0','N','0','0','', '" & par_utente_sap & "')"
        Cmd_SAP.ExecuteNonQuery()


        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO OILM (OILM.MessageID, OILM.DocEntry, OILM.TransType, OILM.DocLineNum, OILM.Quantity, OILM.EffectQty, OILM.LocType, OILM.LocCode, OILM.TotalLC, OILM.TotalFC, OILM.TotalSC, OILM.BaseAbsEnt, OILM.BaseType, OILM.BaseCurr, OILM.Currency, OILM.AccumType, OILM.ActionType, OILM.ExpensesLC, OILM.ExpensesFC, OILM.ExpensesSC, OILM.DocDueDate, OILM.ItemCode, OILM.BPCardCode, OILM.DocDate, OILM.DocRate, OILM.Comment, OILM.JrnlMemo, OILM.Ref1, OILM.Ref2, OILM.BaseLine, OILM.SnBType, OILM.CreateTime, OILM.DataSource, OILM.CreateDate, OILM.OcrCode, OILM.OcrCode2, OILM.OcrCode3, OILM.OcrCode4, OILM.OcrCode5, OILM.DocPrice, OILM.CardName, OILM.Dscription, OILM.TreeType, OILM.ApplObj, OILM.AppObjAbs, OILM.AppObjType, OILM.AppObjLine, OILM.BASE_REF, OILM.TransSeqRf, OILM.LayerIDRef, OILM.VersionNum, OILM.PriceRate, OILM.PriceCurr, OILM.DocTotal, OILM.Price, OILM.CIShbQty, OILM.SubLineNum, OILM.PrjCode, OILM.SlpCode, OILM.TaxDate, OILM.UseDocPric, OILM.VendorNum, OILM.SerialNum, OILM.BlockNum, OILM.ImportLog, OILM.Location, OILM.DocPrcRate, OILM.DocPrcCurr, OILM.CgsOcrCod, OILM.CgsOcrCod2, OILM.CgsOcrCod3, OILM.CgsOcrCod4, OILM.CgsOcrCod5, OILM.BSubLineNo, OILM.AppSubLine, OILM.SysRate, OILM.ExFromRpt, OILM.Ref3, OILM.EnSetCost, OILM.RetCost, OILM.DocAction, OILM.UseShpdGd, OILM.AddTotalLC, OILM.AddExpLC, OILM.IsNegLnQty, OILM.StgSeqNum, OILM.StgEntry, OILM.StgDesc, oilm.usersign ) VALUES
('" & MESSAGEID & "'+2,'" & DOCENTRY_Trasferimenti() & "','67','0'," & par_quantita_trasferibile & "," & par_quantita_trasferibile & ",'64','" & par_magazzino_destinazione & "','0','0','0','" & DOCENTRY_Trasferimenti() & "','67','','EUR','1','1','0','0','0',GETDATE(),'" & par_Codice_SAP & "','" & Business_partner_della_commessa(par_commessa_ODP).codice_bp & "',convert(date,GETDATE()),'0','" & par_stringa_trasferimento & " " & par_magazzino_destinazione & " " & par_Documento & par_ref1 & " " & par_commessa_ODP & "' ,'" & par_stringa_trasferimento & " -'," & DOCNUM_Trasferimenti() & ",'" & par_commessa_ODP & "','0','4',CONCAT(DATEPART(HOUR, GETDATE()),DATEPART(MINUTE, GETDATE())),'I',GETDATE(),'','','','','','" & par_prezzo_listino_acquisto & "','" & Business_partner_della_commessa(par_commessa_ODP).nome_bp & "','" & par_descrizione & "','N','-1','-1','','-1'," & DOCNUM_Trasferimenti() & ",'-1','-1','10.00.140.04','0','EUR','" & par_prezzo_listino_acquisto & "','" & par_prezzo_listino_acquisto & "','0','-1','','-1',convert(date,GETDATE()),'N','','','','','','0','EUR','','','','','','-1','-1','1','N','','N','0','1','N','0','0','N','0','0','','" & par_utente_sap & "')"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub


    Sub aggiusta_Numeratore_OIVL()
        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = " 

update ONNM set ONNM.AUTOKEY= T10.AUTOKEY+1
FROM
(
select max(T0.TransSeq) as 'Autokey'
from oivl AS T0
)
AS T10
WHERE ONNM.OBJECTCODE='10000062' or ONNM.OBJECTCODE='310000000'  

"

        '        update t0 set t0.AUTOKEY=" & Numeratore_OIVL & "+3
        'FROM ONNM t0  
        'WHERE t0.OBJECTCODE='10000062' or t0.OBJECTCODE='310000000'

        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub







    Private Sub DataGridView_trasferito_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_trasferito.CellFormatting
        Try
            Dim daTrasValue As Integer = Convert.ToInt32(DataGridView_trasferito.Rows(e.RowIndex).Cells("Da_tras").Value)
            If daTrasValue = 0 Then
                DataGridView_trasferito.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
            End If
        Catch ex As Exception
            ' Gestisci l'eccezione se necessario
        End Try

        Dim divValue As String = DataGridView_trasferito.Rows(e.RowIndex).Cells("DIV").Value.ToString()
        Select Case divValue
            Case "BRB01"
                DataGridView_trasferito.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.Yellow
            Case "TIR01"
                DataGridView_trasferito.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.LightBlue
            Case "KTF01"
                DataGridView_trasferito.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.Green
        End Select
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Me.Close()
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs)
        Commesse_magazzino_ODP.Show()
        Commesse_magazzino_ODP.Owner = Me
        Commesse_magazzino_ODP.Commesse_odp_aperte()
        Me.Hide()
    End Sub

    Sub ripristino_giacenze_corrette(par_codice As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        ' manca da assegnare il valore par_docentry_rt a baseentry perchè devo trovarlo nel values

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "update t11 set t11.onhand = t10.q
from
(
select itemcode, loccode, sum(inqty-outqty) as Q
from oivl
where itemcode='" & par_codice & "'
group by itemcode, loccode
)
as t10 left join oitw t11 on t10.itemcode=t11.itemcode and t10.loccode=t11.whscode
where t10.q<>t11.onhand"

        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()


    End Sub


    Sub Inserisci_documento_trasferimento(par_ultimo_docentry As Integer, par_ultimo_docnum As Integer, par_Documento As String, par_num_ODP As String, par_num_OC As String, par_qta_trasferimento As String, par_magazzino_partenza As String, par_magazzino_destinazione As String, par_utente_sap As String, par_docentry_rt As Integer, par_prezzo_listino_acquisto As Decimal, par_stringa_trasferimento As String, par_ref_1 As String)
        Dim prezzo As String = Replace(par_prezzo_listino_acquisto, ",", ".")


        Trova_serie_Trasferimento()
        If par_prezzo_listino_acquisto = Nothing Then
            par_prezzo_listino_acquisto = 0
        End If
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        ' manca da assegnare il valore par_docentry_rt a baseentry perchè devo trovarlo nel values

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO OWTR (T0.DOCENTRY,T0.DocNum, T0.DocType, T0.CANCELED, T0.Handwrtten, T0.Printed, T0.DocStatus, T0.InvntSttus, T0.Transfered, T0.ObjType, T0.DocDate, T0.DocDueDate, T0.CardCode, T0.CardName, T0.Address, T0.NumAtCard, T0.VatPercent, T0.VatSum, T0.VatSumFC, T0.DiscPrcnt, T0.DiscSum, T0.DiscSumFC, T0.DocCur, T0.DocRate, T0.DocTotal, T0.DocTotalFC, T0.PaidToDate, T0.PaidFC, T0.GrosProfit, T0.GrosProfFC, T0.Ref1, T0.Ref2, T0.Comments, T0.JrnlMemo, T0.TransId, T0.ReceiptNum, T0.GroupNum, T0.DocTime, T0.SlpCode, T0.TrnspCode, T0.PartSupply, T0.Confirmed, T0.GrossBase, T0.ImportEnt, T0.SummryType, T0.UpdInvnt, T0.UpdCardBal, T0.CntctCode, T0.FatherCard, T0.SysRate, T0.CurSource, T0.VatSumSy, T0.DiscSumSy, T0.DocTotalSy, T0.PaidSys, T0.FatherType, T0.GrosProfSy, T0.UpdateDate, T0.IsICT, T0.CreateDate, T0.Volume, T0.VolUnit, T0.Weight, T0.WeightUnit, T0.Series, T0.TaxDate, T0.Filler, T0.StampNum, T0.isCrin, T0.FinncPriod, T0.UserSign, T0.selfInv, T0.VatPaid, T0.VatPaidFC, T0.VatPaidSys, T0.WddStatus, T0.draftKey, T0.TotalExpns, T0.TotalExpFC, T0.TotalExpSC, T0.Address2, T0.Exported, T0.StationID, T0.Indicator, T0.NetProc, T0.AqcsTax, T0.AqcsTaxFC, T0.AqcsTaxSC, T0.CashDiscPr, T0.CashDiscnt, T0.CashDiscFC, T0.CashDiscSC, T0.ShipToCode, T0.LicTradNum, T0.PaymentRef, T0.WTSum, T0.WTSumFC, T0.WTSumSC, T0.RoundDif, T0.RoundDifFC, T0.RoundDifSy, T0.CheckDigit, T0.Form1099, T0.Box1099, T0.submitted, T0.PoPrss, T0.Rounding, T0.RevisionPo, T0.Segment, T0.ReqDate, T0.CancelDate, T0.PickStatus, T0.Pick, T0.BlockDunn, T0.PeyMethod, T0.PayBlock, T0.PayBlckRef, T0.MaxDscn, T0.Reserve, T0.Max1099, T0.CntrlBnk, T0.PickRmrk, T0.ISRCodLine, T0.ExpAppl, T0.ExpApplFC, T0.ExpApplSC, T0.Project, T0.DeferrTax, T0.LetterNum, T0.FromDate, T0.ToDate, T0.WTApplied, T0.WTAppliedF, T0.BoeReserev, T0.AgentCode, T0.WTAppliedS, T0.EquVatSum, T0.EquVatSumF, T0.EquVatSumS, T0.Installmnt, T0.VATFirst, T0.NnSbAmnt, T0.NnSbAmntSC, T0.NbSbAmntFC, T0.ExepAmnt, T0.ExepAmntSC, T0.ExepAmntFC, T0.VatDate, T0.CorrExt, T0.CorrInv, T0.NCorrInv, T0.CEECFlag, T0.BaseAmnt, T0.BaseAmntSC, T0.BaseAmntFC, T0.CtlAccount, T0.BPLId, T0.BPLName, T0.VATRegNum, T0.TxInvRptNo, T0.TxInvRptDt, T0.KVVATCode, T0.WTDetails, T0.SumAbsId, T0.SumRptDate, T0.PIndicator, T0.ManualNum, T0.UseShpdGd, T0.BaseVtAt, T0.BaseVtAtSC, T0.BaseVtAtFC, T0.NnSbVAt, T0.NnSbVAtSC, T0.NbSbVAtFC, T0.ExptVAt, T0.ExptVAtSC, T0.ExptVAtFC, T0.LYPmtAt, T0.LYPmtAtSC, T0.LYPmtAtFC, T0.ExpAnSum, T0.ExpAnSys, T0.ExpAnFrgn, T0.DocSubType, T0.DpmStatus, T0.DpmAmnt, T0.DpmAmntSC, T0.DpmAmntFC, T0.DpmDrawn, T0.DpmPrcnt, T0.PaidSum, T0.PaidSumFc, T0.PaidSumSc, T0.FolioPref, T0.FolioNum, T0.DpmAppl, T0.DpmApplFc, T0.DpmApplSc, T0.LPgFolioN, T0.Header, T0.Footer, T0.Posted, T0.OwnerCode, T0.BPChCode, T0.BPChCntc, T0.PayToCode, T0.IsPaytoBnk, T0.BnkCntry, T0.BankCode, T0.BnkAccount, T0.BnkBranch, T0.isIns, T0.TrackNo, T0.VersionNum, T0.LangCode, T0.BPNameOW, T0.BillToOW, T0.ShipToOW, T0.RetInvoice, T0.ClsDate, T0.MInvNum, T0.MInvDate, T0.SeqCode, T0.Serial, T0.SeriesStr, T0.SubStr, T0.Model, T0.TaxOnExp, T0.TaxOnExpFc, T0.TaxOnExpSc, T0.TaxOnExAp, T0.TaxOnExApF, T0.TaxOnExApS, T0.LastPmnTyp, T0.LndCstNum, T0.UseCorrVat, T0.BlkCredMmo, T0.OpenForLaC, T0.Excised, T0.ExcRefDate, T0.ExcRmvTime, T0.SrvGpPrcnt, T0.DepositNum, T0.CertNum, T0.DutyStatus, T0.AutoCrtFlw, T0.FlwRefDate, T0.FlwRefNum, T0.VatJENum, T0.DpmVat, T0.DpmVatFc, T0.DpmVatSc, T0.DpmAppVat, T0.DpmAppVatF, T0.DpmAppVatS, T0.InsurOp347, T0.IgnRelDoc, T0.BuildDesc, T0.ResidenNum, T0.Checker, T0.Payee, T0.CopyNumber, T0.SSIExmpt, T0.PQTGrpSer, T0.PQTGrpNum, T0.PQTGrpHW, T0.ReopOriDoc, T0.ReopManCls, T0.DocManClsd, T0.ClosingOpt, T0.SpecDate, T0.Ordered, T0.NTSApprov, T0.NTSWebSite, T0.NTSeTaxNo, T0.NTSApprNo, T0.PayDuMonth, T0.ExtraMonth, T0.ExtraDays, T0.CdcOffset, T0.SignMsg, T0.SignDigest, T0.CertifNum, T0.KeyVersion, T0.EDocGenTyp, T0.ESeries, T0.EDocNum, T0.EDocExpFrm, T0.OnlineQuo, T0.POSEqNum, T0.POSManufSN, T0.POSCashN, T0.EDocStatus, T0.EDocCntnt, T0.EDocProces, T0.EDocErrCod, T0.EDocErrMsg, T0.EDocCancel, T0.EDocTest, T0.EDocPrefix, T0.CUP, T0.CIG, T0.DpmAsDscnt, T0.Attachment, T0.AtcEntry, T0.SupplCode, T0.GTSRlvnt, T0.BaseDisc, T0.BaseDiscSc, T0.BaseDiscFc, T0.BaseDiscPr, T0.CreateTS, T0.UpdateTS, T0.SrvTaxRule, T0.AnnInvDecR, T0.Supplier, T0.Releaser, T0.Receiver, T0.ToWhsCode, T0.AssetDate, T0.Requester, T0.ReqName, T0.Branch, T0.Department, T0.Email, T0.Notify, T0.ReqType, T0.OriginType, T0.IsReuseNum, T0.IsReuseNFN, T0.DocDlvry, T0.PaidDpm, T0.PaidDpmF, T0.PaidDpmS, T0.EnvTypeNFe, T0.AgrNo, T0.IsAlt, T0.AltBaseTyp, T0.AltBaseEnt, T0.AuthCode, T0.StDlvDate, T0.StDlvTime, T0.EndDlvDate, T0.EndDlvTime, T0.VclPlate, T0.ElCoStatus, T0.AtDocType, T0.ElCoMsg, T0.PrintSEPA, T0.FreeChrg, T0.FreeChrgFC, T0.FreeChrgSC, T0.NfeValue, T0.FiscDocNum, T0.RelatedTyp, T0.RelatedEnt, T0.CCDEntry, T0.NfePrntFo, T0.ZrdAbs, T0.POSRcptNo, T0.FoCTax, T0.FoCTaxFC, T0.FoCTaxSC, T0.TpCusPres, T0.ExcDocDate, T0.FoCFrght, T0.FoCFrghtFC, T0.FoCFrghtSC, T0.InterimTyp, T0.PTICode, T0.Letter, T0.FolNumFrom, T0.FolNumTo, T0.FolSeries, T0.SplitTax, T0.SplitTaxFC, T0.SplitTaxSC, T0.ToBinCode, T0.PriceMode, T0.PoDropPrss, T0.PermitNo, T0.MYFtype, T0.DocTaxID, T0.DateReport, T0.RepSection, T0.ExclTaxRep, T0.PosCashReg, T0.DmpTransID, T0.ECommerBP, T0.EComerGSTN, T0.Revision, T0.RevRefNo, T0.RevRefDate, T0.RevCreRefN, T0.RevCreRefD, T0.TaxInvNo, T0.FrmBpDate, T0.GSTTranTyp, T0.BaseType,

T0.BaseEntry,

T0.ComTrade, T0.UseBilAddr, T0.IssReason, T0.ComTradeRt, T0.SplitPmnt, T0.SOIWizId, T0.SelfPosted, T0.EnBnkAcct, T0.EncryptIV, T0.DPPStatus, T0.SAPPassprt, T0.EWBGenType, T0.CtActTax, T0.CtActTaxFC, T0.CtActTaxSC, T0.EDocType, T0.QRCodeSrc, T0.AggregDoc, T0.DataVers, T0.ShipState, T0.ShipPlace, T0.CustOffice, T0.FCI, T0.U_TRASPORT, T0.U_TERMCONS, T0.U_PROF, T0.U_01, T0.U_02, T0.U_03, T0.U_04, T0.U_001, T0.U_002, T0.U_003, T0.U_004, T0.U_CodDog, T0.U_Termini, T0.U_MODCONS, T0.U_DIMENIMB, T0.U_PESOLORD, T0.U_PESONET, T0.U_MATRcds, T0.U_PROVAGEN, T0.U_PARTITA, T0.U_DataTra,


T0.U_Vettore, T0.U_NumCol, T0.U_BanApp, T0.U_Peso, T0.U_TotValOm, T0.U_FATTSER, T0.U_FATTRIF, T0.U_COLLOFF, T0.U_COLLORD, T0.U_AspEst, T0.U_CauTra, T0.U_ACura, T0.U_NUM_SPED, T0.U_CausCons, T0.U_Trasporto, T0.U_Resa, T0.U_Imballo, T0.U_DataTrasp, T0.U_OraTrasp, T0.U_Colli, T0.U_PesoN, T0.U_PesoL, T0.U_TipoFatt, T0.U_NotChiuTS, T0.U_IndDstOA, T0.U_PrzBC, T0.U_StSerLot, T0.U_LettAWB, T0.U_PRG_AZS_Incoterms, T0.U_PRG_AZS_DueDate, T0.U_PRG_AZS_DNumPro, T0.U_PRG_AZS_StLogo, T0.U_PRG_AZS_NrListBP, T0.U_PRG_AZS_NrListDoc, T0.U_PRG_AZS_PrzOFOC, T0.U_PRG_AZS_AliasOC, T0.U_PRG_AZS_GrpBP, T0.U_PRG_AZS_OrdFatt, T0.U_PRG_AZS_ShipDConf, T0.U_PRG_AZS_StQtaImpT, T0.U_PRG_AZS_IncotCity, T0.U_PRG_AZS_StatoComm, T0.U_Vettore2, T0.U_PRG_AZS_DimImb, T0.U_Aspetto_APP, T0.U_Aspetto, T0.U_Plafond, T0.U_PlafRag, T0.U_Dataricordine, T0.U_Dataricacconto, T0.U_DatariccampUT, T0.U_DataricevcampiniFAT, T0.U_Dataconscart, T0.U_DatainizioProg, T0.U_DataprevfineUT, T0.U_Datacollaudo, T0.U_DataFatcliente, T0.U_Commento, T0.U_Datalayout, T0.U_PRG_AZS_SdrDdtBp, T0.U_PRG_AZS_SdrDdtNum, T0.U_PRG_AZS_SdrDdtDate, T0.U_SelArtNS, T0.U_SelSerNS, T0.U_NumSerNS, T0.U_PRG_CVM_DocCorr, T0.U_PRG_CMP_Rate, T0.U_PRG_AZS_Vol, T0.U_PRG_AZS_DtCompIva, T0.U_Elaboratore, T0.U_Uffcompetenza, T0.U_primadataconsegna, T0.U_PRG_CMP_DataCMPEntPro, T0.U_Distributore, T0.U_DataRDO, T0.U_DataCM, T0.U_Settore, T0.U_Destinazione, T0.U_Rev, T0.U_Clientefinale, T0.U_AROLBranch, T0.U_HOT, T0.U_FatturaGAR, T0.U_ArticoloGAR, T0.U_Commessa, T0.U_ODP, T0.U_ConsegnaHOT, T0.U_PrezzoArolbranch, T0.U_Ultimamodifica, T0.U_CodiceBP, T0.U_Aggiustamentofattura, T0.U_VenditoreArol, T0.U_Inventario, T0.U_Opportunita, T0.U_PRG_AZS_Commessa, T0.U_PRG_WIP_Transfer_Type, T0.U_Causale, T0.U_NrNotaspese, T0.U_Aggiornaprezzo, T0.U_ChiudiODP, T0.U_InizioASS, T0.U_FineASS, T0.U_DataHOT, T0.U_Coeff_vendita, T0.U_Bureau_Veritas, T0.U_Categoria, T0.U_PRG_AZS_CreatedBy, T0.U_PRG_AZS_OpDocNum, T0.U_PRG_AZS_OcDocNum, T0.U_PRG_QLT_QCResult, T0.U_PRG_QLT_QCNCEmp, T0.U_PRG_WMS_Exp, T0.U_PRG_WMS_ExpDate, T0.U_B1SYS_INV_TYPE)

VALUES
(" & par_ultimo_docentry & "+1, " & par_ultimo_docnum & "+1,'I','N','N','N','O','O','N','67',GETDATE(),GETDATE(),'','','','','0','0','0','0','0','0','EUR','1'," & prezzo & "*" & par_qta_trasferimento & ",'0','0','0','0','0'," & DOCNUM_Trasferimenti() & "+1,'" & par_ref_1 & "','" & par_stringa_trasferimento & " Materiali con 4.0" & par_Documento & " Nr " & par_ref_1 & "','" & par_stringa_trasferimento & " magazzino -','','','-1',CONCAT(DATEPART(HOUR, GETDATE()),DATEPART(MINUTE, GETDATE())),'','','Y','Y','2','','N','I','N','','','1','L','0','0'," & par_qta_trasferimento & " * " & prezzo & ",'0','P','0',GETDATE(),'N',GETDATE(),'0','2','0','2','" & Trova_serie_Trasferimento() & "',GETDATE(),'" & par_magazzino_partenza & "','','N','" & trova_absentry() & "','" & par_utente_sap & "','N','0','0','0','-','','0','0','0','','N','','','N','0','0','0','0','0','0','0','','','','0','0','0','0','0','0','','','','N','N','N','N','','','','N','N','N','','N','','N','N'," & par_qta_trasferimento & " * " & prezzo & ",'','','','0','0','0','','N','','','','0','0','N','','0','0','0','0','1','N','0','0','0','0','0','0',GETDATE(),'','','','N','0','0','0','','','','','','','','','-1','','" & Trova_PERIODO_contabile() & "','','N','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','--','O','0','0','0','N','0','0','0','0','','','0','0','0','','','','Y','','','','','','','','','','N','','10.00.140.04','','N','N','N','N','','','','','','','','0','0','0','0','0','0','0','','','N','N','Y','O','','','0','','','Y','N','','','-1','0','0','0','0','0','0','N','N','','1','','','','','','','N','','','N','1','','N','N','','','','','','','','','','','','N','','','','N','','','','C','','C','','','N','N','','','','N','','','','N','0','0','0','0','','" & Mid(Now, 12, 2) & Mid(Now, 15, 2) & Mid(Now, 18, 2) & "','N','','','','','" & par_magazzino_destinazione & "','','','','','','','','12','M','N','N','0','0','0','0','-1','','N','-1','','','','','','','','','','','N','0','0','0','0','','-1','','','','','','0','0','0','','','0','0','0','','','','','','','0','0','0','','','N','','','','','','N','','','','','N','','','','','','','','-1','','E','','1','N','N','','N','','','N','','','0','0','0','F','','N','1','','','','','','','','','','','','0','0','0','0','','','','','','','','','','','','','','0','0','','','','','','','','','V','V','PA','I','','','','0','0','F','','','N','N','', '','','','N','','','Y','','','N','N','Y','','O','','','','','0','0','','','','','','','','','','','','','','','','','','N','0','0','','','DA DEFINIRE','','','','','','','','0','','','','','','0','','','0','','','0','','','','','1','','','','','','','','0','','Standard','','" & DOCENTRY_documento(par_num_ODP, par_num_OC, par_Documento).Docentryodp & "','" & DOCENTRY_documento(par_num_ODP, par_num_OC, par_Documento).Docentryoc & "','X','','N','','TD01')
"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()


    End Sub

    Sub Inserisci_documento_richiesta_trasferimento(par_docentry As Integer, par_docnum As Integer, par_Documento As String, par_num_ODP As String, par_num_OC As String, par_qta_trasferimento As String, par_magazzino_partenza As String, par_magazzino_destinazione As String, par_docentry_odp As Integer, par_docentry_oc As Integer, par_utente_sap As String, par_stringa_trasferimento As String)

        Trova_serie_RT()
        Trova_PERIODO_contabile()

        Dim ref1 As String

        If par_Documento = "ODP" Then
            ref1 = par_num_ODP
        ElseIf par_Documento = "OC" Then
            ref1 = par_num_ODP
        End If
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand



        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO OWTq (T0.DOCENTRY,T0.DocNum, T0.DocType, T0.CANCELED, T0.Handwrtten, T0.Printed, T0.DocStatus, T0.InvntSttus, T0.Transfered, T0.ObjType, T0.DocDate, T0.DocDueDate, T0.CardCode, T0.CardName, T0.Address, T0.NumAtCard, T0.VatPercent, T0.VatSum, T0.VatSumFC, T0.DiscPrcnt, T0.DiscSum, T0.DiscSumFC, T0.DocCur, T0.DocRate, T0.DocTotal, T0.DocTotalFC, T0.PaidToDate, T0.PaidFC, T0.GrosProfit, T0.GrosProfFC, T0.Ref1, T0.Ref2, T0.Comments, T0.JrnlMemo, T0.TransId, T0.ReceiptNum, T0.GroupNum, T0.DocTime, T0.SlpCode, T0.TrnspCode, T0.PartSupply, T0.Confirmed, T0.GrossBase, T0.ImportEnt, T0.SummryType, T0.UpdInvnt, T0.UpdCardBal, T0.CntctCode, T0.FatherCard, T0.SysRate, T0.CurSource, T0.VatSumSy, T0.DiscSumSy, T0.DocTotalSy, T0.PaidSys, T0.FatherType, T0.GrosProfSy, T0.UpdateDate, T0.IsICT, T0.CreateDate, T0.Volume, T0.VolUnit, T0.Weight, T0.WeightUnit, T0.Series, T0.TaxDate, T0.Filler, T0.StampNum, T0.isCrin, T0.FinncPriod, T0.UserSign, T0.selfInv, T0.VatPaid, T0.VatPaidFC, T0.VatPaidSys, T0.WddStatus, T0.draftKey, T0.TotalExpns, T0.TotalExpFC, T0.TotalExpSC, T0.Address2, T0.Exported, T0.StationID, T0.Indicator, T0.NetProc, T0.AqcsTax, T0.AqcsTaxFC, T0.AqcsTaxSC, T0.CashDiscPr, T0.CashDiscnt, T0.CashDiscFC, T0.CashDiscSC, T0.ShipToCode, T0.LicTradNum, T0.PaymentRef, T0.WTSum, T0.WTSumFC, T0.WTSumSC, T0.RoundDif, T0.RoundDifFC, T0.RoundDifSy, T0.CheckDigit, T0.Form1099, T0.Box1099, T0.submitted, T0.PoPrss, T0.Rounding, T0.RevisionPo, T0.Segment, T0.ReqDate, T0.CancelDate, T0.PickStatus, T0.Pick, T0.BlockDunn, T0.PeyMethod, T0.PayBlock, T0.PayBlckRef, T0.MaxDscn, T0.Reserve, T0.Max1099, T0.CntrlBnk, T0.PickRmrk, T0.ISRCodLine, T0.ExpAppl, T0.ExpApplFC, T0.ExpApplSC, T0.Project, T0.DeferrTax, T0.LetterNum, T0.FromDate, T0.ToDate, T0.WTApplied, T0.WTAppliedF, T0.BoeReserev, T0.AgentCode, T0.WTAppliedS, T0.EquVatSum, T0.EquVatSumF, T0.EquVatSumS, T0.Installmnt, T0.VATFirst, T0.NnSbAmnt, T0.NnSbAmntSC, T0.NbSbAmntFC, T0.ExepAmnt, T0.ExepAmntSC, T0.ExepAmntFC, T0.VatDate, T0.CorrExt, T0.CorrInv, T0.NCorrInv, T0.CEECFlag, T0.BaseAmnt, T0.BaseAmntSC, T0.BaseAmntFC, T0.CtlAccount, T0.BPLId, T0.BPLName, T0.VATRegNum, T0.TxInvRptNo, T0.TxInvRptDt, T0.KVVATCode, T0.WTDetails, T0.SumAbsId, T0.SumRptDate, T0.PIndicator, T0.ManualNum, T0.UseShpdGd, T0.BaseVtAt, T0.BaseVtAtSC, T0.BaseVtAtFC, T0.NnSbVAt, T0.NnSbVAtSC, T0.NbSbVAtFC, T0.ExptVAt, T0.ExptVAtSC, T0.ExptVAtFC, T0.LYPmtAt, T0.LYPmtAtSC, T0.LYPmtAtFC, T0.ExpAnSum, T0.ExpAnSys, T0.ExpAnFrgn, T0.DocSubType, T0.DpmStatus, T0.DpmAmnt, T0.DpmAmntSC, T0.DpmAmntFC, T0.DpmDrawn, T0.DpmPrcnt, T0.PaidSum, T0.PaidSumFc, T0.PaidSumSc, T0.FolioPref, T0.FolioNum, T0.DpmAppl, T0.DpmApplFc, T0.DpmApplSc, T0.LPgFolioN, T0.Header, T0.Footer, T0.Posted, T0.OwnerCode, T0.BPChCode, T0.BPChCntc, T0.PayToCode, T0.IsPaytoBnk, T0.BnkCntry, T0.BankCode, T0.BnkAccount, T0.BnkBranch, T0.isIns, T0.TrackNo, T0.VersionNum, T0.LangCode, T0.BPNameOW, T0.BillToOW, T0.ShipToOW, T0.RetInvoice, T0.ClsDate, T0.MInvNum, T0.MInvDate, T0.SeqCode, T0.Serial, T0.SeriesStr, T0.SubStr, T0.Model, T0.TaxOnExp, T0.TaxOnExpFc, T0.TaxOnExpSc, T0.TaxOnExAp, T0.TaxOnExApF, T0.TaxOnExApS, T0.LastPmnTyp, T0.LndCstNum, T0.UseCorrVat, T0.BlkCredMmo, T0.OpenForLaC, T0.Excised, T0.ExcRefDate, T0.ExcRmvTime, T0.SrvGpPrcnt, T0.DepositNum, T0.CertNum, T0.DutyStatus, T0.AutoCrtFlw, T0.FlwRefDate, T0.FlwRefNum, T0.VatJENum, T0.DpmVat, T0.DpmVatFc, T0.DpmVatSc, T0.DpmAppVat, T0.DpmAppVatF, T0.DpmAppVatS, T0.InsurOp347, T0.IgnRelDoc, T0.BuildDesc, T0.ResidenNum, T0.Checker, T0.Payee, T0.CopyNumber, T0.SSIExmpt, T0.PQTGrpSer, T0.PQTGrpNum, T0.PQTGrpHW, T0.ReopOriDoc, T0.ReopManCls, T0.DocManClsd, T0.ClosingOpt, T0.SpecDate, T0.Ordered, T0.NTSApprov, T0.NTSWebSite, T0.NTSeTaxNo, T0.NTSApprNo, T0.PayDuMonth, T0.ExtraMonth, T0.ExtraDays, T0.CdcOffset, T0.SignMsg, T0.SignDigest, T0.CertifNum, T0.KeyVersion, T0.EDocGenTyp, T0.ESeries, T0.EDocNum, T0.EDocExpFrm, T0.OnlineQuo, T0.POSEqNum, T0.POSManufSN, T0.POSCashN, T0.EDocStatus, T0.EDocCntnt, T0.EDocProces, T0.EDocErrCod, T0.EDocErrMsg, T0.EDocCancel, T0.EDocTest, T0.EDocPrefix, T0.CUP, T0.CIG, T0.DpmAsDscnt, T0.Attachment, T0.AtcEntry, T0.SupplCode, T0.GTSRlvnt, T0.BaseDisc, T0.BaseDiscSc, T0.BaseDiscFc, T0.BaseDiscPr, T0.CreateTS, T0.UpdateTS, T0.SrvTaxRule, T0.AnnInvDecR, T0.Supplier, T0.Releaser, T0.Receiver, T0.ToWhsCode, T0.AssetDate, T0.Requester, T0.ReqName, T0.Branch, T0.Department, T0.Email, T0.Notify, T0.ReqType, T0.OriginType, T0.IsReuseNum, T0.IsReuseNFN, T0.DocDlvry, T0.PaidDpm, T0.PaidDpmF, T0.PaidDpmS, T0.EnvTypeNFe, T0.AgrNo, T0.IsAlt, T0.AltBaseTyp, T0.AltBaseEnt, T0.AuthCode, T0.StDlvDate, T0.StDlvTime, T0.EndDlvDate, T0.EndDlvTime, T0.VclPlate, T0.ElCoStatus, T0.AtDocType, T0.ElCoMsg, T0.PrintSEPA, T0.FreeChrg, T0.FreeChrgFC, T0.FreeChrgSC, T0.NfeValue, T0.FiscDocNum, T0.RelatedTyp, T0.RelatedEnt, T0.CCDEntry, T0.NfePrntFo, T0.ZrdAbs, T0.POSRcptNo, T0.FoCTax, T0.FoCTaxFC, T0.FoCTaxSC, T0.TpCusPres, T0.ExcDocDate, T0.FoCFrght, T0.FoCFrghtFC, T0.FoCFrghtSC, T0.InterimTyp, T0.PTICode, T0.Letter, T0.FolNumFrom, T0.FolNumTo, T0.FolSeries, T0.SplitTax, T0.SplitTaxFC, T0.SplitTaxSC, T0.ToBinCode, T0.PriceMode, T0.PoDropPrss, T0.PermitNo, T0.MYFtype, T0.DocTaxID, T0.DateReport, T0.RepSection, T0.ExclTaxRep, T0.PosCashReg, T0.DmpTransID, T0.ECommerBP, T0.EComerGSTN, T0.Revision, T0.RevRefNo, T0.RevRefDate, T0.RevCreRefN, T0.RevCreRefD, T0.TaxInvNo, T0.FrmBpDate, T0.GSTTranTyp, T0.BaseType, T0.BaseEntry, T0.ComTrade, T0.UseBilAddr, T0.IssReason, T0.ComTradeRt, T0.SplitPmnt, T0.SOIWizId, T0.SelfPosted, T0.EnBnkAcct, T0.EncryptIV, T0.DPPStatus, T0.SAPPassprt, T0.EWBGenType, T0.CtActTax, T0.CtActTaxFC, T0.CtActTaxSC, T0.EDocType, T0.QRCodeSrc, T0.AggregDoc, T0.DataVers, T0.ShipState, T0.ShipPlace, T0.CustOffice, T0.FCI, T0.U_TRASPORT, T0.U_TERMCONS, T0.U_PROF, T0.U_01, T0.U_02, T0.U_03, T0.U_04, T0.U_001, T0.U_002, T0.U_003, T0.U_004, T0.U_CodDog, T0.U_Termini, T0.U_MODCONS, T0.U_DIMENIMB, T0.U_PESOLORD, T0.U_PESONET, T0.U_MATRcds, T0.U_PROVAGEN, T0.U_PARTITA, T0.U_DataTra,


T0.U_Vettore, T0.U_NumCol, T0.U_BanApp, T0.U_Peso, T0.U_TotValOm, T0.U_FATTSER, T0.U_FATTRIF, T0.U_COLLOFF, T0.U_COLLORD, T0.U_AspEst, T0.U_CauTra, T0.U_ACura, T0.U_NUM_SPED, T0.U_CausCons, T0.U_Trasporto, T0.U_Resa, T0.U_Imballo, T0.U_DataTrasp, T0.U_OraTrasp, T0.U_Colli, T0.U_PesoN, T0.U_PesoL, T0.U_TipoFatt, T0.U_NotChiuTS, T0.U_IndDstOA, T0.U_PrzBC, T0.U_StSerLot, T0.U_LettAWB, T0.U_PRG_AZS_Incoterms, T0.U_PRG_AZS_DueDate, T0.U_PRG_AZS_DNumPro, T0.U_PRG_AZS_StLogo, T0.U_PRG_AZS_NrListBP, T0.U_PRG_AZS_NrListDoc, T0.U_PRG_AZS_PrzOFOC, T0.U_PRG_AZS_AliasOC, T0.U_PRG_AZS_GrpBP, T0.U_PRG_AZS_OrdFatt, T0.U_PRG_AZS_ShipDConf, T0.U_PRG_AZS_StQtaImpT, T0.U_PRG_AZS_IncotCity, T0.U_PRG_AZS_StatoComm, T0.U_Vettore2, T0.U_PRG_AZS_DimImb, T0.U_Aspetto_APP, T0.U_Aspetto, T0.U_Plafond, T0.U_PlafRag, T0.U_Dataricordine, T0.U_Dataricacconto, T0.U_DatariccampUT, T0.U_DataricevcampiniFAT, T0.U_Dataconscart, T0.U_DatainizioProg, T0.U_DataprevfineUT, T0.U_Datacollaudo, T0.U_DataFatcliente, T0.U_Commento, T0.U_Datalayout, T0.U_PRG_AZS_SdrDdtBp, T0.U_PRG_AZS_SdrDdtNum, T0.U_PRG_AZS_SdrDdtDate, T0.U_SelArtNS, T0.U_SelSerNS, T0.U_NumSerNS, T0.U_PRG_CVM_DocCorr, T0.U_PRG_CMP_Rate, T0.U_PRG_AZS_Vol, T0.U_PRG_AZS_DtCompIva, T0.U_Elaboratore, T0.U_Uffcompetenza, T0.U_primadataconsegna, T0.U_PRG_CMP_DataCMPEntPro, T0.U_Distributore, T0.U_DataRDO, T0.U_DataCM, T0.U_Settore, T0.U_Destinazione, T0.U_Rev, T0.U_Clientefinale, T0.U_AROLBranch, T0.U_HOT, T0.U_FatturaGAR, T0.U_ArticoloGAR, T0.U_Commessa, T0.U_ODP, T0.U_ConsegnaHOT, T0.U_PrezzoArolbranch, T0.U_Ultimamodifica, T0.U_CodiceBP, T0.U_Aggiustamentofattura, T0.U_VenditoreArol, T0.U_Inventario, T0.U_Opportunita, T0.U_PRG_AZS_Commessa, T0.U_PRG_WIP_Transfer_Type, T0.U_Causale, T0.U_NrNotaspese, T0.U_Aggiornaprezzo, T0.U_ChiudiODP, T0.U_InizioASS, T0.U_FineASS, T0.U_DataHOT, T0.U_Coeff_vendita, T0.U_Bureau_Veritas, T0.U_Categoria, T0.U_PRG_AZS_CreatedBy, T0.U_PRG_AZS_OpDocNum, T0.U_PRG_AZS_OcDocNum, T0.U_PRG_QLT_QCResult, T0.U_PRG_QLT_QCNCEmp, T0.U_PRG_WMS_Exp, T0.U_PRG_WMS_ExpDate, T0.U_B1SYS_INV_TYPE)

VALUES
(" & par_docentry & ", " & par_docnum & ",'I','N','N','N','O','O','N','1250000001',GETDATE(),GETDATE(),'','','','','0','0','0','0','0','0','EUR','1',0,'0','0','0','0','0'," & par_docnum & ",'" & ref1 & "','" & par_stringa_trasferimento & "" & par_Documento & " Nr " & ref1 & "','" & par_stringa_trasferimento & " magazzino -','','','-1',CONCAT(DATEPART(HOUR, GETDATE()),DATEPART(MINUTE, GETDATE())),'','','Y','Y','2','','N','I','N','','','1','L','0','0',0,'0','P','0',GETDATE(),'N',GETDATE(),'0','2','0','2','" & Trova_serie_RT() & "',GETDATE(),'" & par_magazzino_partenza & "','','N','" & trova_absentry() & "','" & par_utente_sap & "','N','0','0','0','-','','0','0','0','','N','','','N','0','0','0','0','0','0','0','','','','0','0','0','0','0','0','','','','N','N','N','N','','','','N','N','N','','N','','N','N',0,'','','','0','0','0','','N','','','','0','0','N','','0','0','0','0','1','N','0','0','0','0','0','0',GETDATE(),'','','','N','0','0','0','','','','','','','','','-1','','" & Trova_PERIODO_contabile() & "','','N','0','0','0','0','0','0','0','0','0','0','0','0','0','0','0','--','O','0','0','0','N','0','0','0','0','','','0','0','0','','','','Y','','','','','','','','','','N','','10.00.140.04','','N','N','N','N','','','','','','','','0','0','0','0','0','0','0','','','N','N','Y','O','','','0','','','Y','N','','','-1','0','0','0','0','0','0','N','N','','1','','','','','','','N','','','N','1','','N','N','','','','','','','','','','','','N','','','','N','','','','C','','C','','','N','N','','','','N','','','','N','0','0','0','0','','" & Mid(Now, 12, 2) & Mid(Now, 15, 2) & Mid(Now, 18, 2) & "','N','','','','','" & par_magazzino_destinazione & "','','','','','','','','12','M','N','N','0','0','0','0','-1','','N','-1','','','','','','','','','','','N','0','0','0','0','','-1','','','','','','0','0','0','','','0','0','0','','','','','','','0','0','0','','','N','','','','','','N','','','','','N','','','','','','','','-1','','E','','1','N','N','','N','','','N','','','0','0','0','F','','N','1','','','','','','','','','','','','0','0','0','0','','','','','','','','','','','','','','0','0','','','','','','','','','V','V','PA','I','','','','0','0','F','','','N','N','', '','','','N','','','Y','','','N','N','Y','','O','','','','','0','0','','','','','','','','','','','','','','','','','','N','0','0','','','DA DEFINIRE','','','','','','','','0','','','','','','0','','','0','','','0','','','','','1','','','','','','','','0','','Standard','','" & par_docentry_odp & "','" & par_docentry_oc & "','X','','N','','TD01')
"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()


    End Sub




    Sub Inserisci_righe_trasferimento(PAR_DOCENTRY As Integer, par_documento As String, par_numero_ODP As Integer, par_numero_oc As Integer, par_Codice_SAP As String, par_quantita_trasferibile As String, par_magazzino_partenza As String, par_magazzino_destinazione As String, par_linenum_ODP As String, par_prezzo_listino_acquisto As String, par_descrizione As String, par_riga As Integer)
        par_prezzo_listino_acquisto = Replace(par_prezzo_listino_acquisto, ",", ".")
        par_descrizione = Replace(par_descrizione, "'", "")
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        If par_documento = "ODP" Then
            Cmd_SAP.Connection = Cnn
            Cmd_SAP.CommandText = "INSERT INTO WTR1 (WTR1.DocEntry, WTR1.LineNum, WTR1.TargetType,  WTR1.BaseType, WTR1.LineStatus, WTR1.ItemCode, WTR1.Dscription, WTR1.Quantity, WTR1.ShipDate, WTR1.OpenQty, WTR1.Price, WTR1.Currency, WTR1.Rate, WTR1.DiscPrcnt, WTR1.LineTotal, WTR1.TotalFrgn, WTR1.OpenSum, WTR1.OpenSumFC, WTR1.VendorNum, WTR1.SerialNum, WTR1.WhsCode, WTR1.SlpCode, WTR1.Commission, WTR1.TreeType, WTR1.AcctCode, WTR1.TaxStatus, WTR1.GrossBuyPr, WTR1.PriceBefDi, WTR1.DocDate, WTR1.OpenCreQty, WTR1.UseBaseUn, WTR1.SubCatNum, WTR1.BaseCard, WTR1.TotalSumSy, WTR1.OpenSumSys, WTR1.InvntSttus, WTR1.OcrCode, WTR1.Project, WTR1.CodeBars, WTR1.VatPrcnt, WTR1.VatGroup, WTR1.PriceAfVAT, WTR1.Height1, WTR1.Hght1Unit, WTR1.Height2, WTR1.Hght2Unit, WTR1.Width1, WTR1.Wdth1Unit, WTR1.Width2, WTR1.Wdth2Unit, WTR1.Length1, WTR1.Len1Unit, WTR1.length2, WTR1.Len2Unit, WTR1.Volume, WTR1.VolUnit, WTR1.Weight1, WTR1.Wght1Unit, WTR1.Weight2, WTR1.Wght2Unit, WTR1.Factor1, WTR1.Factor2, WTR1.Factor3, WTR1.Factor4, WTR1.PackQty, WTR1.UpdInvntry, WTR1.BaseDocNum, WTR1.BaseAtCard, WTR1.SWW, WTR1.VatSum, WTR1.VatSumFrgn, WTR1.VatSumSy, WTR1.FinncPriod, WTR1.ObjType, WTR1.BlockNum, WTR1.ImportLog, WTR1.DedVatSum, WTR1.DedVatSumF, WTR1.DedVatSumS, WTR1.IsAqcuistn, WTR1.DistribSum, WTR1.DstrbSumFC, WTR1.DstrbSumSC, WTR1.GrssProfit, WTR1.GrssProfSC, WTR1.GrssProfFC, WTR1.VisOrder, WTR1.INMPrice, WTR1.PoTrgNum, WTR1.PoTrgEntry, WTR1.DropShip, WTR1.PoLineNum, WTR1.Address, WTR1.TaxCode, WTR1.TaxType, WTR1.OrigItem, WTR1.BackOrdr, WTR1.FreeTxt, WTR1.PickStatus, WTR1.PickOty, WTR1.PickIdNo, WTR1.TrnsCode, WTR1.VatAppld, WTR1.VatAppldFC, WTR1.VatAppldSC, WTR1.BaseQty, WTR1.BaseOpnQty, WTR1.VatDscntPr, WTR1.WtLiable, WTR1.DeferrTax, WTR1.EquVatPer, WTR1.EquVatSum, WTR1.EquVatSumF, WTR1.EquVatSumS, WTR1.LineVat, WTR1.LineVatlF, WTR1.LineVatS, WTR1.unitMsr, WTR1.NumPerMsr, WTR1.CEECFlag, WTR1.ToStock, WTR1.ToDiff, WTR1.ExciseAmt, WTR1.TaxPerUnit, WTR1.TotInclTax, WTR1.CountryOrg, WTR1.StckDstSum, WTR1.ReleasQtty, WTR1.LineType, WTR1.TranType, WTR1.Text, WTR1.OwnerCode, WTR1.StockPrice, WTR1.ConsumeFCT, WTR1.LstByDsSum, WTR1.StckINMPr, WTR1.LstBINMPr, WTR1.StckDstFc, WTR1.StckDstSc, WTR1.LstByDsFc, WTR1.LstByDsSc, WTR1.StockSum, WTR1.StockSumFc, WTR1.StockSumSc, WTR1.StckSumApp, WTR1.StckAppFc, WTR1.StckAppSc, WTR1.ShipToCode, WTR1.ShipToDesc, WTR1.StckAppD, WTR1.StckAppDFC, WTR1.StckAppDSC, WTR1.BasePrice, WTR1.GTotal, WTR1.GTotalFC, WTR1.GTotalSC, WTR1.DistribExp, WTR1.DescOW, WTR1.DetailsOW, WTR1.GrossBase, WTR1.VatWoDpm, WTR1.VatWoDpmFc, WTR1.VatWoDpmSc, WTR1.CFOPCode, WTR1.CSTCode, WTR1.Usage, WTR1.TaxOnly, WTR1.WtCalced, WTR1.QtyToShip, WTR1.DelivrdQty, WTR1.OrderedQty, WTR1.CogsOcrCod, WTR1.CiOppLineN, WTR1.CogsAcct, WTR1.ChgAsmBoMW, WTR1.ActDelDate, WTR1.OcrCode2, WTR1.OcrCode3, WTR1.OcrCode4, WTR1.OcrCode5, WTR1.TaxDistSum, WTR1.TaxDistSFC, WTR1.TaxDistSSC, WTR1.PostTax, WTR1.Excisable, WTR1.AssblValue, WTR1.RG23APart1, WTR1.RG23APart2, WTR1.RG23CPart1, WTR1.RG23CPart2, WTR1.CogsOcrCo2, WTR1.CogsOcrCo3, WTR1.CogsOcrCo4, WTR1.CogsOcrCo5, WTR1.LnExcised, WTR1.LocCode, WTR1.StockValue, WTR1.GPTtlBasPr, WTR1.unitMsr2, WTR1.NumPerMsr2, WTR1.SpecPrice, WTR1.CSTfIPI, WTR1.CSTfPIS, WTR1.CSTfCOFINS, WTR1.ExLineNo, WTR1.isSrvCall, WTR1.PQTReqQty, WTR1.PQTReqDate, WTR1.PcDocType, WTR1.PcQuantity, WTR1.LinManClsd, WTR1.VatGrpSrc, WTR1.NoInvtryMv, WTR1.ActBaseEnt, WTR1.ActBaseLn, WTR1.ActBaseNum, WTR1.OpenRtnQty, WTR1.AgrNo, WTR1.AgrLnNum, WTR1.CredOrigin, WTR1.Surpluses, WTR1.DefBreak, WTR1.Shortages, WTR1.UomEntry, WTR1.UomEntry2, WTR1.UomCode, WTR1.UomCode2, WTR1.FromWhsCod, WTR1.NeedQty, WTR1.PartRetire, WTR1.RetireQty, WTR1.RetireAPC, WTR1.RetirAPCFC, WTR1.RetirAPCSC, WTR1.InvQty, WTR1.OpenInvQty, WTR1.EnSetCost, WTR1.RetCost, WTR1.Incoterms, WTR1.TransMod, WTR1.LineVendor, WTR1.DistribIS, WTR1.ISDistrb, WTR1.ISDistrbFC, WTR1.ISDistrbSC, WTR1.IsByPrdct, WTR1.ItemType, WTR1.PriceEdit, WTR1.PrntLnNum, WTR1.LinePoPrss, WTR1.FreeChrgBP, WTR1.TaxRelev, WTR1.LegalText, WTR1.ThirdParty, WTR1.LicTradNum, WTR1.InvQtyOnly, WTR1.UnencReasn, WTR1.ShipFromCo, WTR1.ShipFromDe, WTR1.FisrtBin, WTR1.AllocBinC, WTR1.ExpType, WTR1.ExpUUID, WTR1.ExpOpType, WTR1.DIOTNat, WTR1.MYFtype, WTR1.GPBefDisc, WTR1.ReturnRsn, WTR1.ReturnAct, WTR1.StgSeqNum, WTR1.StgEntry, WTR1.StgDesc, WTR1.ItmTaxType, WTR1.SacEntry, WTR1.NCMCode, WTR1.HsnEntry, WTR1.OriBAbsEnt, WTR1.OriBLinNum, WTR1.OriBDocTyp, WTR1.IsPrscGood, WTR1.IsCstmAct, WTR1.EncryptIV, WTR1.ExtTaxRate, WTR1.ExtTaxSum, WTR1.TaxAmtSrc, WTR1.ExtTaxSumF, WTR1.ExtTaxSumS, WTR1.StdItemId, WTR1.CommClass, WTR1.VatExEntry, WTR1.VatExLN, WTR1.NatOfTrans, WTR1.ISDtCryImp, WTR1.ISDtRgnImp, WTR1.ISOrCryExp, WTR1.ISOrRgnExp, WTR1.NVECode, WTR1.PoNum, WTR1.PoItmNum, WTR1.IndEscala, WTR1.CESTCode, WTR1.CtrSealQty, WTR1.CNJPMan, WTR1.UFFiscBene, WTR1.U_BLD_LyID, WTR1.U_BLD_NCps, WTR1.U_O01FlagU, WTR1.U_O01ProAg, WTR1.U_O01ProCA, WTR1.U_O01ProCZ, WTR1.U_O01ProDI, WTR1.U_O01PrzGr, WTR1.U_O01ScoIm, WTR1.U_BnTrian, WTR1.U_Note, WTR1.U_TrasMgEM, WTR1.U_Totval, WTR1.U_BNIncTrm, WTR1.U_BNTrnMod, WTR1.U_TestoDOC, WTR1.U_QtySup, WTR1.U_PRG_AZS_OpDocEntry, WTR1.U_PRG_AZS_OpLineNum, WTR1.U_TpForn, WTR1.U_PRG_AZS_DescrAlt, WTR1.U_PRG_AZS_PrevMPS, WTR1.U_PRG_AZS_StatoComm, WTR1.U_Colli, WTR1.U_PRG_AZS_OcDocEntry, WTR1.U_PRG_AZS_OcDocNum, WTR1.U_PRG_AZS_OcLineNum, WTR1.U_PRG_AZS_OaDocEntry, WTR1.U_PRG_AZS_OaDocNum, WTR1.U_PRG_AZS_OaLineNum, WTR1.U_Datitecncompl, WTR1.U_UTdatainiz, WTR1.U_UTfineprog, WTR1.U_inizioassel, WTR1.U_Fineassel, WTR1.U_inizassmecc, WTR1.U_fineassmecc, WTR1.U_PRG_AZS_OpDocNum, WTR1.U_PRG_AZS_Commessa, WTR1.U_PRG_AZS_NumAtCard, WTR1.U_PRG_AZS_DataRic, WTR1.U_PRG_AZS_DataCon, WTR1.U_PRG_AZS_PrzProForma, WTR1.U_PRG_CLV_PrzPia, WTR1.U_PRG_CLV_PrzLav, WTR1.U_PRG_CVM_DocAssoc, WTR1.U_B1SYS_Discount, WTR1.U_B1SYS_Discount_FC, WTR1.U_B1SYS_Discount_SC, WTR1.U_B1SYS_DiscountVat, WTR1.U_B1SYS_DiscountVtFC, WTR1.U_B1SYS_DiscountVtSC, WTR1.U_Inizcol, WTR1.U_Finecol, WTR1.U_Fineapp, WTR1.U_mod_macchina, WTR1.U_Fine_app_MU, WTR1.U_Inizio_ass_EL, WTR1.U_Fine_ass_EL, WTR1.U_Inizioapprovvigionamento, WTR1.U_DataKOM, WTR1.U_PListinoAcqu, WTR1.U_Ultimoprezzodeterminato, WTR1.U_Migliorprezzo, WTR1.U_Migliorfornitore, WTR1.U_Trasferito, WTR1.U_Datrasferire, WTR1.U_Almag01, WTR1.U_AlmagCDS, WTR1.U_Opportunita, WTR1.U_Ubicazione, WTR1.U_O01Sc1, WTR1.U_O01Sc2, WTR1.U_O01Sc3, WTR1.U_O01Sc4, WTR1.U_O01Sc5, WTR1.U_O01Sc6, WTR1.U_Ricarico, WTR1.U_Prezzoarolbranch, WTR1.U_Commissione_agente, WTR1.U_Costo, WTR1.U_Data_scheda_tecnica, WTR1.U_Data_clean_order, WTR1.U_Disegno, WTR1.U_Produttore, WTR1.U_Revisione, WTR1.U_PRG_AZS_UbiDest, WTR1.U_PRG_AZS_PrjFather, WTR1.U_PRG_AZS_QtaEvasa, WTR1.U_PRG_WIP_QtaRichMagAuto, WTR1.U_PRG_QLT_QCDlnQty, WTR1.U_PRG_QLT_QCCntQty, WTR1.U_PRG_QLT_QCNCResE, WTR1.U_PRG_QLT_QCNCResM, WTR1.U_PRG_QLT_HasTC, WTR1.U_PRG_WMS_Exp, WTR1.U_PRG_WMS_ExpDate, WTR1.U_PRG_WMS_MdMovQty, WTR1.U_Coefficiente_vendita, WTR1.U_Gestito_Ferretto, WTR1.U_Mag_ferretto)

VALUES (" & PAR_DOCENTRY & "," & par_riga & ",'-1','-1','O','" & par_Codice_SAP & "','" & par_descrizione & "'," & par_quantita_trasferibile & ",GETDATE()," & par_quantita_trasferibile & ",'" & par_prezzo_listino_acquisto & "','EUR','0','0'," & par_quantita_trasferibile & "*" & par_prezzo_listino_acquisto & ",'0'," & par_quantita_trasferibile & "*" & par_prezzo_listino_acquisto & ",'0','','','" & par_magazzino_destinazione & "','-1','0','N','','','0'," & par_prezzo_listino_acquisto & ",GETDATE()," & par_quantita_trasferibile & ",'N','',''," & par_quantita_trasferibile & "*" & par_prezzo_listino_acquisto & "," & par_quantita_trasferibile & "*" & par_prezzo_listino_acquisto & ",'O','','','','0','','0','0','','0','','0','','0','','0','','0','','0','','0','','0','','1','1','1','1','0','Y','','','','0','0','0','" & trova_absentry() & "','67','','','0','0','0','N','0','0','0','0','0','0'," & par_riga & ",'" & par_prezzo_listino_acquisto & "','','','N','','','','Y','','','','N','0','','17','0','0','0','4','4','0','N','N','0','0','0','0','0','0','0','PZ','1','S','0','0','0','0','0','','0','0','R','','','','0','','0','0','0','0','0','0','0','0','0','0','0','0','0','','','0','0','0','E','0','0','0','Y','N','N','','0','0','0','','','','N','N','0','0','0','','-1','','','','','','','','0','0','0','Y','','0','','','','','','','','','','','0','0','pz','1','N','','','','','N','0','','-1','0','N','N','N','','','','0','','','','0','0','0','-1','-1','Manuale','Manuale','" & par_magazzino_partenza & "','N','N','0','0','0','0','" & par_quantita_trasferibile & "','" & par_quantita_trasferibile & "','N','0','0','0','','N','0','0','0','N'," & par_quantita_trasferibile & ",'N','','N','N','Y','','N','','N','','','','','0','','','','','','0','-1','-1','','','','','','-1','','','','','N','N','','0','0','S','0','0','','','','','47','IT','0','','','','','','N','','0','','','-1','','N','0','0','0','0','0','0','N','','N','0','','','','0','" & DOCENTRY_documento(par_numero_ODP, par_numero_oc, par_documento).Docentryodp & "','" & par_linenum_ODP & "','NO','','','O','','','','','','','','','','','','','','','" & par_numero_ODP & "','','','','','0','0','0','','0','0','0','0','0','0','','','','','','','','','','0','0','0','0','0','0','0','0','','','','','','','','','','','0','0','','','','','','','','0','0','0','0','X','X','N','N','','0','0','','0')"
        ElseIf par_documento = "OC" Then
            Cmd_SAP.Connection = Cnn
            Cmd_SAP.CommandText = "INSERT INTO WTR1 (WTR1.DocEntry, WTR1.LineNum, WTR1.TargetType, WTR1.BaseType,  WTR1.LineStatus, WTR1.ItemCode, WTR1.Dscription, WTR1.Quantity, WTR1.ShipDate, WTR1.OpenQty, WTR1.Price, WTR1.Currency, WTR1.Rate, WTR1.DiscPrcnt, WTR1.LineTotal, WTR1.TotalFrgn, WTR1.OpenSum, WTR1.OpenSumFC, WTR1.VendorNum, WTR1.SerialNum, WTR1.WhsCode, WTR1.SlpCode, WTR1.Commission, WTR1.TreeType, WTR1.AcctCode, WTR1.TaxStatus, WTR1.GrossBuyPr, WTR1.PriceBefDi, WTR1.DocDate, WTR1.OpenCreQty, WTR1.UseBaseUn, WTR1.SubCatNum, WTR1.BaseCard, WTR1.TotalSumSy, WTR1.OpenSumSys, WTR1.InvntSttus, WTR1.OcrCode, WTR1.Project, WTR1.CodeBars, WTR1.VatPrcnt, WTR1.VatGroup, WTR1.PriceAfVAT, WTR1.Height1, WTR1.Hght1Unit, WTR1.Height2, WTR1.Hght2Unit, WTR1.Width1, WTR1.Wdth1Unit, WTR1.Width2, WTR1.Wdth2Unit, WTR1.Length1, WTR1.Len1Unit, WTR1.length2, WTR1.Len2Unit, WTR1.Volume, WTR1.VolUnit, WTR1.Weight1, WTR1.Wght1Unit, WTR1.Weight2, WTR1.Wght2Unit, WTR1.Factor1, WTR1.Factor2, WTR1.Factor3, WTR1.Factor4, WTR1.PackQty, WTR1.UpdInvntry, WTR1.BaseDocNum, WTR1.BaseAtCard, WTR1.SWW, WTR1.VatSum, WTR1.VatSumFrgn, WTR1.VatSumSy, WTR1.FinncPriod, WTR1.ObjType, WTR1.BlockNum, WTR1.ImportLog, WTR1.DedVatSum, WTR1.DedVatSumF, WTR1.DedVatSumS, WTR1.IsAqcuistn, WTR1.DistribSum, WTR1.DstrbSumFC, WTR1.DstrbSumSC, WTR1.GrssProfit, WTR1.GrssProfSC, WTR1.GrssProfFC, WTR1.VisOrder, WTR1.INMPrice, WTR1.PoTrgNum, WTR1.PoTrgEntry, WTR1.DropShip, WTR1.PoLineNum, WTR1.Address, WTR1.TaxCode, WTR1.TaxType, WTR1.OrigItem, WTR1.BackOrdr, WTR1.FreeTxt, WTR1.PickStatus, WTR1.PickOty, WTR1.PickIdNo, WTR1.TrnsCode, WTR1.VatAppld, WTR1.VatAppldFC, WTR1.VatAppldSC, WTR1.BaseQty, WTR1.BaseOpnQty, WTR1.VatDscntPr, WTR1.WtLiable, WTR1.DeferrTax, WTR1.EquVatPer, WTR1.EquVatSum, WTR1.EquVatSumF, WTR1.EquVatSumS, WTR1.LineVat, WTR1.LineVatlF, WTR1.LineVatS, WTR1.unitMsr, WTR1.NumPerMsr, WTR1.CEECFlag, WTR1.ToStock, WTR1.ToDiff, WTR1.ExciseAmt, WTR1.TaxPerUnit, WTR1.TotInclTax, WTR1.CountryOrg, WTR1.StckDstSum, WTR1.ReleasQtty, WTR1.LineType, WTR1.TranType, WTR1.Text, WTR1.OwnerCode, WTR1.StockPrice, WTR1.ConsumeFCT, WTR1.LstByDsSum, WTR1.StckINMPr, WTR1.LstBINMPr, WTR1.StckDstFc, WTR1.StckDstSc, WTR1.LstByDsFc, WTR1.LstByDsSc, WTR1.StockSum, WTR1.StockSumFc, WTR1.StockSumSc, WTR1.StckSumApp, WTR1.StckAppFc, WTR1.StckAppSc, WTR1.ShipToCode, WTR1.ShipToDesc, WTR1.StckAppD, WTR1.StckAppDFC, WTR1.StckAppDSC, WTR1.BasePrice, WTR1.GTotal, WTR1.GTotalFC, WTR1.GTotalSC, WTR1.DistribExp, WTR1.DescOW, WTR1.DetailsOW, WTR1.GrossBase, WTR1.VatWoDpm, WTR1.VatWoDpmFc, WTR1.VatWoDpmSc, WTR1.CFOPCode, WTR1.CSTCode, WTR1.Usage, WTR1.TaxOnly, WTR1.WtCalced, WTR1.QtyToShip, WTR1.DelivrdQty, WTR1.OrderedQty, WTR1.CogsOcrCod, WTR1.CiOppLineN, WTR1.CogsAcct, WTR1.ChgAsmBoMW, WTR1.ActDelDate, WTR1.OcrCode2, WTR1.OcrCode3, WTR1.OcrCode4, WTR1.OcrCode5, WTR1.TaxDistSum, WTR1.TaxDistSFC, WTR1.TaxDistSSC, WTR1.PostTax, WTR1.Excisable, WTR1.AssblValue, WTR1.RG23APart1, WTR1.RG23APart2, WTR1.RG23CPart1, WTR1.RG23CPart2, WTR1.CogsOcrCo2, WTR1.CogsOcrCo3, WTR1.CogsOcrCo4, WTR1.CogsOcrCo5, WTR1.LnExcised, WTR1.LocCode, WTR1.StockValue, WTR1.GPTtlBasPr, WTR1.unitMsr2, WTR1.NumPerMsr2, WTR1.SpecPrice, WTR1.CSTfIPI, WTR1.CSTfPIS, WTR1.CSTfCOFINS, WTR1.ExLineNo, WTR1.isSrvCall, WTR1.PQTReqQty, WTR1.PQTReqDate, WTR1.PcDocType, WTR1.PcQuantity, WTR1.LinManClsd, WTR1.VatGrpSrc, WTR1.NoInvtryMv, WTR1.ActBaseEnt, WTR1.ActBaseLn, WTR1.ActBaseNum, WTR1.OpenRtnQty, WTR1.AgrNo, WTR1.AgrLnNum, WTR1.CredOrigin, WTR1.Surpluses, WTR1.DefBreak, WTR1.Shortages, WTR1.UomEntry, WTR1.UomEntry2, WTR1.UomCode, WTR1.UomCode2, WTR1.FromWhsCod, WTR1.NeedQty, WTR1.PartRetire, WTR1.RetireQty, WTR1.RetireAPC, WTR1.RetirAPCFC, WTR1.RetirAPCSC, WTR1.InvQty, WTR1.OpenInvQty, WTR1.EnSetCost, WTR1.RetCost, WTR1.Incoterms, WTR1.TransMod, WTR1.LineVendor, WTR1.DistribIS, WTR1.ISDistrb, WTR1.ISDistrbFC, WTR1.ISDistrbSC, WTR1.IsByPrdct, WTR1.ItemType, WTR1.PriceEdit, WTR1.PrntLnNum, WTR1.LinePoPrss, WTR1.FreeChrgBP, WTR1.TaxRelev, WTR1.LegalText, WTR1.ThirdParty, WTR1.LicTradNum, WTR1.InvQtyOnly, WTR1.UnencReasn, WTR1.ShipFromCo, WTR1.ShipFromDe, WTR1.FisrtBin, WTR1.AllocBinC, WTR1.ExpType, WTR1.ExpUUID, WTR1.ExpOpType, WTR1.DIOTNat, WTR1.MYFtype, WTR1.GPBefDisc, WTR1.ReturnRsn, WTR1.ReturnAct, WTR1.StgSeqNum, WTR1.StgEntry, WTR1.StgDesc, WTR1.ItmTaxType, WTR1.SacEntry, WTR1.NCMCode, WTR1.HsnEntry, WTR1.OriBAbsEnt, WTR1.OriBLinNum, WTR1.OriBDocTyp, WTR1.IsPrscGood, WTR1.IsCstmAct, WTR1.EncryptIV, WTR1.ExtTaxRate, WTR1.ExtTaxSum, WTR1.TaxAmtSrc, WTR1.ExtTaxSumF, WTR1.ExtTaxSumS, WTR1.StdItemId, WTR1.CommClass, WTR1.VatExEntry, WTR1.VatExLN, WTR1.NatOfTrans, WTR1.ISDtCryImp, WTR1.ISDtRgnImp, WTR1.ISOrCryExp, WTR1.ISOrRgnExp, WTR1.NVECode, WTR1.PoNum, WTR1.PoItmNum, WTR1.IndEscala, WTR1.CESTCode, WTR1.CtrSealQty, WTR1.CNJPMan, WTR1.UFFiscBene, WTR1.U_BLD_LyID, WTR1.U_BLD_NCps, WTR1.U_O01FlagU, WTR1.U_O01ProAg, WTR1.U_O01ProCA, WTR1.U_O01ProCZ, WTR1.U_O01ProDI, WTR1.U_O01PrzGr, WTR1.U_O01ScoIm, WTR1.U_BnTrian, WTR1.U_Note, WTR1.U_TrasMgEM, WTR1.U_Totval, WTR1.U_BNIncTrm, WTR1.U_BNTrnMod, WTR1.U_TestoDOC, WTR1.U_QtySup, WTR1.U_PRG_AZS_OpDocEntry, WTR1.U_PRG_AZS_OpLineNum, WTR1.U_TpForn, WTR1.U_PRG_AZS_DescrAlt, WTR1.U_PRG_AZS_PrevMPS, WTR1.U_PRG_AZS_StatoComm, WTR1.U_Colli, WTR1.U_PRG_AZS_OcDocEntry, WTR1.U_PRG_AZS_OcDocNum, WTR1.U_PRG_AZS_OcLineNum, WTR1.U_PRG_AZS_OaDocEntry, WTR1.U_PRG_AZS_OaDocNum, WTR1.U_PRG_AZS_OaLineNum, WTR1.U_Datitecncompl, WTR1.U_UTdatainiz, WTR1.U_UTfineprog, WTR1.U_inizioassel, WTR1.U_Fineassel, WTR1.U_inizassmecc, WTR1.U_fineassmecc, WTR1.U_PRG_AZS_OpDocNum, WTR1.U_PRG_AZS_Commessa, WTR1.U_PRG_AZS_NumAtCard, WTR1.U_PRG_AZS_DataRic, WTR1.U_PRG_AZS_DataCon, WTR1.U_PRG_AZS_PrzProForma, WTR1.U_PRG_CLV_PrzPia, WTR1.U_PRG_CLV_PrzLav, WTR1.U_PRG_CVM_DocAssoc, WTR1.U_B1SYS_Discount, WTR1.U_B1SYS_Discount_FC, WTR1.U_B1SYS_Discount_SC, WTR1.U_B1SYS_DiscountVat, WTR1.U_B1SYS_DiscountVtFC, WTR1.U_B1SYS_DiscountVtSC, WTR1.U_Inizcol, WTR1.U_Finecol, WTR1.U_Fineapp, WTR1.U_mod_macchina, WTR1.U_Fine_app_MU, WTR1.U_Inizio_ass_EL, WTR1.U_Fine_ass_EL, WTR1.U_Inizioapprovvigionamento, WTR1.U_DataKOM, WTR1.U_PListinoAcqu, WTR1.U_Ultimoprezzodeterminato, WTR1.U_Migliorprezzo, WTR1.U_Migliorfornitore, WTR1.U_Trasferito, WTR1.U_Datrasferire, WTR1.U_Almag01, WTR1.U_AlmagCDS, WTR1.U_Opportunita, WTR1.U_Ubicazione, WTR1.U_O01Sc1, WTR1.U_O01Sc2, WTR1.U_O01Sc3, WTR1.U_O01Sc4, WTR1.U_O01Sc5, WTR1.U_O01Sc6, WTR1.U_Ricarico, WTR1.U_Prezzoarolbranch, WTR1.U_Commissione_agente, WTR1.U_Costo, WTR1.U_Data_scheda_tecnica, WTR1.U_Data_clean_order, WTR1.U_Disegno, WTR1.U_Produttore, WTR1.U_Revisione, WTR1.U_PRG_AZS_UbiDest, WTR1.U_PRG_AZS_PrjFather, WTR1.U_PRG_AZS_QtaEvasa, WTR1.U_PRG_WIP_QtaRichMagAuto, WTR1.U_PRG_QLT_QCDlnQty, WTR1.U_PRG_QLT_QCCntQty, WTR1.U_PRG_QLT_QCNCResE, WTR1.U_PRG_QLT_QCNCResM, WTR1.U_PRG_QLT_HasTC, WTR1.U_PRG_WMS_Exp, WTR1.U_PRG_WMS_ExpDate, WTR1.U_PRG_WMS_MdMovQty, WTR1.U_Coefficiente_vendita, WTR1.U_Gestito_Ferretto, WTR1.U_Mag_ferretto)

VALUES (" & PAR_DOCENTRY & "," & par_riga & ",'-1','-1','O','" & par_Codice_SAP & "','" & par_descrizione & "'," & par_quantita_trasferibile & ",GETDATE()," & par_quantita_trasferibile & ",'" & par_prezzo_listino_acquisto & "','EUR','0','0'," & par_quantita_trasferibile & "*" & par_prezzo_listino_acquisto & ",'0'," & par_quantita_trasferibile & "*" & par_prezzo_listino_acquisto & ",'0','','','" & par_magazzino_destinazione & "','-1','0','N','','','0'," & par_prezzo_listino_acquisto & ",GETDATE()," & par_quantita_trasferibile & ",'N','',''," & par_quantita_trasferibile & "*" & par_prezzo_listino_acquisto & "," & par_quantita_trasferibile & "*" & par_prezzo_listino_acquisto & ",'O','','','','0','','0','0','','0','','0','','0','','0','','0','','0','','0','','0','','1','1','1','1','0','Y','','','','0','0','0','" & trova_absentry() & "','67','','','0','0','0','N','0','0','0','0','0','0'," & par_riga & ",'" & par_prezzo_listino_acquisto & "','','','N','','','','Y','','','','N','0','','17','0','0','0','4','4','0','N','N','0','0','0','0','0','0','0','PZ','1','S','0','0','0','0','0','','0','0','R','','','','0','','0','0','0','0','0','0','0','0','0','0','0','0','0','','','0','0','0','E','0','0','0','Y','N','N','','0','0','0','','','','N','N','0','0','0','','-1','','','','','','','','0','0','0','Y','','0','','','','','','','','','','','0','0','pz','1','N','','','','','N','0','','-1','0','N','N','N','','','','0','','','','0','0','0','-1','-1','Manuale','Manuale','" & par_magazzino_partenza & "','N','N','0','0','0','0','" & par_quantita_trasferibile & "','" & par_quantita_trasferibile & "','N','0','0','0','','N','0','0','0','N'," & par_quantita_trasferibile & ",'N','','N','N','Y','','N','','N','','','','','0','','','','','','0','-1','-1','','','','','','-1','','','','','N','N','','0','0','S','0','0','','','','','47','IT','0','','','','','','N','','0','','','-1','','N','0','0','0','0','0','0','N','','N','0','','','','0','','','NO','','','O','','" & DOCENTRY_documento(par_numero_ODP, par_numero_oc, par_documento).Docentryoc & "','" & par_numero_oc & "','" & par_linenum_ODP & "','','','','','','','','','','','','','','','','0','0','0','','0','0','0','0','0','0','','','','','','','','','','0','0','0','0','0','0','0','0','','','','','','','','','','','0','0','','','','','','','','0','0','0','0','X','X','N','N','','0','0','','0')"
        Else

            Cmd_SAP.Connection = Cnn
            Cmd_SAP.CommandText = "INSERT INTO WTR1 (WTR1.DocEntry, WTR1.LineNum, WTR1.TargetType,  WTR1.BaseType, WTR1.LineStatus, WTR1.ItemCode, WTR1.Dscription, WTR1.Quantity, WTR1.ShipDate, WTR1.OpenQty, WTR1.Price, WTR1.Currency, WTR1.Rate, WTR1.DiscPrcnt, WTR1.LineTotal, WTR1.TotalFrgn, WTR1.OpenSum, WTR1.OpenSumFC, WTR1.VendorNum, WTR1.SerialNum, WTR1.WhsCode, WTR1.SlpCode, WTR1.Commission, WTR1.TreeType, WTR1.AcctCode, WTR1.TaxStatus, WTR1.GrossBuyPr, WTR1.PriceBefDi, WTR1.DocDate, WTR1.OpenCreQty, WTR1.UseBaseUn, WTR1.SubCatNum, WTR1.BaseCard, WTR1.TotalSumSy, WTR1.OpenSumSys, WTR1.InvntSttus, WTR1.OcrCode, WTR1.Project, WTR1.CodeBars, WTR1.VatPrcnt, WTR1.VatGroup, WTR1.PriceAfVAT, WTR1.Height1, WTR1.Hght1Unit, WTR1.Height2, WTR1.Hght2Unit, WTR1.Width1, WTR1.Wdth1Unit, WTR1.Width2, WTR1.Wdth2Unit, WTR1.Length1, WTR1.Len1Unit, WTR1.length2, WTR1.Len2Unit, WTR1.Volume, WTR1.VolUnit, WTR1.Weight1, WTR1.Wght1Unit, WTR1.Weight2, WTR1.Wght2Unit, WTR1.Factor1, WTR1.Factor2, WTR1.Factor3, WTR1.Factor4, WTR1.PackQty, WTR1.UpdInvntry, WTR1.BaseDocNum, WTR1.BaseAtCard, WTR1.SWW, WTR1.VatSum, WTR1.VatSumFrgn, WTR1.VatSumSy, WTR1.FinncPriod, WTR1.ObjType, WTR1.BlockNum, WTR1.ImportLog, WTR1.DedVatSum, WTR1.DedVatSumF, WTR1.DedVatSumS, WTR1.IsAqcuistn, WTR1.DistribSum, WTR1.DstrbSumFC, WTR1.DstrbSumSC, WTR1.GrssProfit, WTR1.GrssProfSC, WTR1.GrssProfFC, WTR1.VisOrder, WTR1.INMPrice, WTR1.PoTrgNum, WTR1.PoTrgEntry, WTR1.DropShip, WTR1.PoLineNum, WTR1.Address, WTR1.TaxCode, WTR1.TaxType, WTR1.OrigItem, WTR1.BackOrdr, WTR1.FreeTxt, WTR1.PickStatus, WTR1.PickOty, WTR1.PickIdNo, WTR1.TrnsCode, WTR1.VatAppld, WTR1.VatAppldFC, WTR1.VatAppldSC, WTR1.BaseQty, WTR1.BaseOpnQty, WTR1.VatDscntPr, WTR1.WtLiable, WTR1.DeferrTax, WTR1.EquVatPer, WTR1.EquVatSum, WTR1.EquVatSumF, WTR1.EquVatSumS, WTR1.LineVat, WTR1.LineVatlF, WTR1.LineVatS, WTR1.unitMsr, WTR1.NumPerMsr, WTR1.CEECFlag, WTR1.ToStock, WTR1.ToDiff, WTR1.ExciseAmt, WTR1.TaxPerUnit, WTR1.TotInclTax, WTR1.CountryOrg, WTR1.StckDstSum, WTR1.ReleasQtty, WTR1.LineType, WTR1.TranType, WTR1.Text, WTR1.OwnerCode, WTR1.StockPrice, WTR1.ConsumeFCT, WTR1.LstByDsSum, WTR1.StckINMPr, WTR1.LstBINMPr, WTR1.StckDstFc, WTR1.StckDstSc, WTR1.LstByDsFc, WTR1.LstByDsSc, WTR1.StockSum, WTR1.StockSumFc, WTR1.StockSumSc, WTR1.StckSumApp, WTR1.StckAppFc, WTR1.StckAppSc, WTR1.ShipToCode, WTR1.ShipToDesc, WTR1.StckAppD, WTR1.StckAppDFC, WTR1.StckAppDSC, WTR1.BasePrice, WTR1.GTotal, WTR1.GTotalFC, WTR1.GTotalSC, WTR1.DistribExp, WTR1.DescOW, WTR1.DetailsOW, WTR1.GrossBase, WTR1.VatWoDpm, WTR1.VatWoDpmFc, WTR1.VatWoDpmSc, WTR1.CFOPCode, WTR1.CSTCode, WTR1.Usage, WTR1.TaxOnly, WTR1.WtCalced, WTR1.QtyToShip, WTR1.DelivrdQty, WTR1.OrderedQty, WTR1.CogsOcrCod, WTR1.CiOppLineN, WTR1.CogsAcct, WTR1.ChgAsmBoMW, WTR1.ActDelDate, WTR1.OcrCode2, WTR1.OcrCode3, WTR1.OcrCode4, WTR1.OcrCode5, WTR1.TaxDistSum, WTR1.TaxDistSFC, WTR1.TaxDistSSC, WTR1.PostTax, WTR1.Excisable, WTR1.AssblValue, WTR1.RG23APart1, WTR1.RG23APart2, WTR1.RG23CPart1, WTR1.RG23CPart2, WTR1.CogsOcrCo2, WTR1.CogsOcrCo3, WTR1.CogsOcrCo4, WTR1.CogsOcrCo5, WTR1.LnExcised, WTR1.LocCode, WTR1.StockValue, WTR1.GPTtlBasPr, WTR1.unitMsr2, WTR1.NumPerMsr2, WTR1.SpecPrice, WTR1.CSTfIPI, WTR1.CSTfPIS, WTR1.CSTfCOFINS, WTR1.ExLineNo, WTR1.isSrvCall, WTR1.PQTReqQty, WTR1.PQTReqDate, WTR1.PcDocType, WTR1.PcQuantity, WTR1.LinManClsd, WTR1.VatGrpSrc, WTR1.NoInvtryMv, WTR1.ActBaseEnt, WTR1.ActBaseLn, WTR1.ActBaseNum, WTR1.OpenRtnQty, WTR1.AgrNo, WTR1.AgrLnNum, WTR1.CredOrigin, WTR1.Surpluses, WTR1.DefBreak, WTR1.Shortages, WTR1.UomEntry, WTR1.UomEntry2, WTR1.UomCode, WTR1.UomCode2, WTR1.FromWhsCod, WTR1.NeedQty, WTR1.PartRetire, WTR1.RetireQty, WTR1.RetireAPC, WTR1.RetirAPCFC, WTR1.RetirAPCSC, WTR1.InvQty, WTR1.OpenInvQty, WTR1.EnSetCost, WTR1.RetCost, WTR1.Incoterms, WTR1.TransMod, WTR1.LineVendor, WTR1.DistribIS, WTR1.ISDistrb, WTR1.ISDistrbFC, WTR1.ISDistrbSC, WTR1.IsByPrdct, WTR1.ItemType, WTR1.PriceEdit, WTR1.PrntLnNum, WTR1.LinePoPrss, WTR1.FreeChrgBP, WTR1.TaxRelev, WTR1.LegalText, WTR1.ThirdParty, WTR1.LicTradNum, WTR1.InvQtyOnly, WTR1.UnencReasn, WTR1.ShipFromCo, WTR1.ShipFromDe, WTR1.FisrtBin, WTR1.AllocBinC, WTR1.ExpType, WTR1.ExpUUID, WTR1.ExpOpType, WTR1.DIOTNat, WTR1.MYFtype, WTR1.GPBefDisc, WTR1.ReturnRsn, WTR1.ReturnAct, WTR1.StgSeqNum, WTR1.StgEntry, WTR1.StgDesc, WTR1.ItmTaxType, WTR1.SacEntry, WTR1.NCMCode, WTR1.HsnEntry, WTR1.OriBAbsEnt, WTR1.OriBLinNum, WTR1.OriBDocTyp, WTR1.IsPrscGood, WTR1.IsCstmAct, WTR1.EncryptIV, WTR1.ExtTaxRate, WTR1.ExtTaxSum, WTR1.TaxAmtSrc, WTR1.ExtTaxSumF, WTR1.ExtTaxSumS, WTR1.StdItemId, WTR1.CommClass, WTR1.VatExEntry, WTR1.VatExLN, WTR1.NatOfTrans, WTR1.ISDtCryImp, WTR1.ISDtRgnImp, WTR1.ISOrCryExp, WTR1.ISOrRgnExp, WTR1.NVECode, WTR1.PoNum, WTR1.PoItmNum, WTR1.IndEscala, WTR1.CESTCode, WTR1.CtrSealQty, WTR1.CNJPMan, WTR1.UFFiscBene, WTR1.U_BLD_LyID, WTR1.U_BLD_NCps, WTR1.U_O01FlagU, WTR1.U_O01ProAg, WTR1.U_O01ProCA, WTR1.U_O01ProCZ, WTR1.U_O01ProDI, WTR1.U_O01PrzGr, WTR1.U_O01ScoIm, WTR1.U_BnTrian, WTR1.U_Note, WTR1.U_TrasMgEM, WTR1.U_Totval, WTR1.U_BNIncTrm, WTR1.U_BNTrnMod, WTR1.U_TestoDOC, WTR1.U_QtySup, WTR1.U_PRG_AZS_OpDocEntry, WTR1.U_PRG_AZS_OpLineNum, WTR1.U_TpForn, WTR1.U_PRG_AZS_DescrAlt, WTR1.U_PRG_AZS_PrevMPS, WTR1.U_PRG_AZS_StatoComm, WTR1.U_Colli, WTR1.U_PRG_AZS_OcDocEntry, WTR1.U_PRG_AZS_OcDocNum, WTR1.U_PRG_AZS_OcLineNum, WTR1.U_PRG_AZS_OaDocEntry, WTR1.U_PRG_AZS_OaDocNum, WTR1.U_PRG_AZS_OaLineNum, WTR1.U_Datitecncompl, WTR1.U_UTdatainiz, WTR1.U_UTfineprog, WTR1.U_inizioassel, WTR1.U_Fineassel, WTR1.U_inizassmecc, WTR1.U_fineassmecc, WTR1.U_PRG_AZS_OpDocNum, WTR1.U_PRG_AZS_Commessa, WTR1.U_PRG_AZS_NumAtCard, WTR1.U_PRG_AZS_DataRic, WTR1.U_PRG_AZS_DataCon, WTR1.U_PRG_AZS_PrzProForma, WTR1.U_PRG_CLV_PrzPia, WTR1.U_PRG_CLV_PrzLav, WTR1.U_PRG_CVM_DocAssoc, WTR1.U_B1SYS_Discount, WTR1.U_B1SYS_Discount_FC, WTR1.U_B1SYS_Discount_SC, WTR1.U_B1SYS_DiscountVat, WTR1.U_B1SYS_DiscountVtFC, WTR1.U_B1SYS_DiscountVtSC, WTR1.U_Inizcol, WTR1.U_Finecol, WTR1.U_Fineapp, WTR1.U_mod_macchina, WTR1.U_Fine_app_MU, WTR1.U_Inizio_ass_EL, WTR1.U_Fine_ass_EL, WTR1.U_Inizioapprovvigionamento, WTR1.U_DataKOM, WTR1.U_PListinoAcqu, WTR1.U_Ultimoprezzodeterminato, WTR1.U_Migliorprezzo, WTR1.U_Migliorfornitore, WTR1.U_Trasferito, WTR1.U_Datrasferire, WTR1.U_Almag01, WTR1.U_AlmagCDS, WTR1.U_Opportunita, WTR1.U_Ubicazione, WTR1.U_O01Sc1, WTR1.U_O01Sc2, WTR1.U_O01Sc3, WTR1.U_O01Sc4, WTR1.U_O01Sc5, WTR1.U_O01Sc6, WTR1.U_Ricarico, WTR1.U_Prezzoarolbranch, WTR1.U_Commissione_agente, WTR1.U_Costo, WTR1.U_Data_scheda_tecnica, WTR1.U_Data_clean_order, WTR1.U_Disegno, WTR1.U_Produttore, WTR1.U_Revisione, WTR1.U_PRG_AZS_UbiDest, WTR1.U_PRG_AZS_PrjFather, WTR1.U_PRG_AZS_QtaEvasa, WTR1.U_PRG_WIP_QtaRichMagAuto, WTR1.U_PRG_QLT_QCDlnQty, WTR1.U_PRG_QLT_QCCntQty, WTR1.U_PRG_QLT_QCNCResE, WTR1.U_PRG_QLT_QCNCResM, WTR1.U_PRG_QLT_HasTC, WTR1.U_PRG_WMS_Exp, WTR1.U_PRG_WMS_ExpDate, WTR1.U_PRG_WMS_MdMovQty, WTR1.U_Coefficiente_vendita, WTR1.U_Gestito_Ferretto, WTR1.U_Mag_ferretto)

VALUES (" & PAR_DOCENTRY & "," & par_riga & ",'-1','-1','O','" & par_Codice_SAP & "','" & par_descrizione & "'," & par_quantita_trasferibile & ",GETDATE()," & par_quantita_trasferibile & ",'" & par_prezzo_listino_acquisto & "','EUR','0','0'," & par_quantita_trasferibile & "*" & par_prezzo_listino_acquisto & ",'0'," & par_quantita_trasferibile & "*" & par_prezzo_listino_acquisto & ",'0','','','" & par_magazzino_destinazione & "','-1','0','N','','','0'," & par_prezzo_listino_acquisto & ",GETDATE()," & par_quantita_trasferibile & ",'N','',''," & par_quantita_trasferibile & "*" & par_prezzo_listino_acquisto & "," & par_quantita_trasferibile & "*" & par_prezzo_listino_acquisto & ",'O','','','','0','','0','0','','0','','0','','0','','0','','0','','0','','0','','0','','1','1','1','1','0','Y','','','','0','0','0','" & trova_absentry() & "','67','','','0','0','0','N','0','0','0','0','0','0'," & par_riga & ",'" & par_prezzo_listino_acquisto & "','','','N','','','','Y','','','','N','0','','17','0','0','0','4','4','0','N','N','0','0','0','0','0','0','0','PZ','1','S','0','0','0','0','0','','0','0','R','','','','0','','0','0','0','0','0','0','0','0','0','0','0','0','0','','','0','0','0','E','0','0','0','Y','N','N','','0','0','0','','','','N','N','0','0','0','','-1','','','','','','','','0','0','0','Y','','0','','','','','','','','','','','0','0','pz','1','N','','','','','N','0','','-1','0','N','N','N','','','','0','','','','0','0','0','-1','-1','Manuale','Manuale','" & par_magazzino_partenza & "','N','N','0','0','0','0','" & par_quantita_trasferibile & "','" & par_quantita_trasferibile & "','N','0','0','0','','N','0','0','0','N'," & par_quantita_trasferibile & ",'N','','N','N','Y','','N','','N','','','','','0','','','','','','0','-1','-1','','','','','','-1','','','','','N','N','','0','0','S','0','0','','','','','47','IT','0','','','','','','N','','0','','','-1','','N','0','0','0','0','0','0','N','','N','0','','','','0','" & DOCENTRY_documento(par_numero_ODP, par_numero_oc, par_documento).Docentryodp & "','" & par_linenum_ODP & "','NO','','','O','','','','','','','','','','','','','','','" & par_numero_ODP & "','','','','','0','0','0','','0','0','0','0','0','0','','','','','','','','','','0','0','0','0','0','0','0','0','','','','','','','','','','','0','0','','','','','','','','0','0','0','0','X','X','N','N','','0','0','','0')"

        End If

        Cmd_SAP.ExecuteNonQuery()
        Cnn.Close()

    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Enter
        trasferito(Codice_SAP, DataGridView_trasferito)
    End Sub
    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter
        ordinato(Codice_SAP, DataGridView_ordinato)
    End Sub
    Private Sub TabPage4_Click(sender As Object, e As EventArgs) Handles TabPage4.Enter
        rt_aperte(DataGridView1, Codice_SAP)
    End Sub
    Private Sub TabPage9_Click(sender As Object, e As EventArgs) Handles TabPage9.Enter
        rof_aperte(DataGridView6, Codice_SAP)
    End Sub

    Sub Inserimento_dipendenti(PAR_COmBOBOX As ComboBox)

        'Dim filtro_regola_distribuzione = ""

        'If Homepage.Centro_di_costo = "BRB01" Then
        '    filtro_regola_distribuzione = " And t0.costcenter='BRB01' "
        'Else
        '    filtro_regola_distribuzione = ""
        'End If

        PAR_COmBOBOX.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[userid] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 left join oudp t1 on T0.[dept]=t1.code 
--inner join [TIRELLI_40].[DBO].COLL_Reparti t2 on (cast(t2.sap_id_reparto as varchar) =cast(t1.code as varchar) or cast(t2.sap_id_reparto_2 as varchar) =cast(t1.code as varchar))  
where t0.active='Y' AND T0.[userid] <>'' 
--and cast(t2.id_reparto as varchar)='" & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto & "' 

order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()

            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            PAR_COmBOBOX.Items.Add(cmd_SAP_reader("Nome"))

            If Elenco_dipendenti(Indice) = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato Then
                PAR_COmBOBOX.SelectedIndex = Indice
            End If

            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Sub Inserimento_produttore(par_combobox As ComboBox)

        par_combobox.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP.CommandText = "SELECT T0.FirmCode, T0.FirmName FROM OMRC T0
order by T0.FirmName
"
        Else
            CMD_SAP.CommandText = "SELECT   

*


FROM OPENQUERY(AS400, '
    SELECT    '''' AS Firmcode,
      prod_for AS Firmname
    FROM S786FAD1.TIR90VIS.JGALART
    WHERE prod_for<>  ''''
	group by prod_for
') T10"
        End If


        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()

            Elenco_produttori(Indice) = cmd_SAP_reader("FirmCode")
            par_combobox.Items.Add(cmd_SAP_reader("FirmName"))



            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Inserisco le risorse nella combo box


    Sub Inserimento_GESTIONE_MAGAZZINO(par_combobox As ComboBox)

        par_combobox.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP.CommandText = "select fldvalue, descr from ufd1
where FIELDID=103 AND TABLEID='OITM' ORDER BY INDEXID
"
        Else
            CMD_SAP.CommandText = "select '' as fldvalue, '' as 'descr'"
        End If


        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()

            Elenco_gestione(Indice) = cmd_SAP_reader("fldvalue")
            par_combobox.Items.Add(cmd_SAP_reader("descr"))



            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Inserisco le risorse nella combo box



    Sub Lista_registrazioni(par_codice_sap As String,
                        par_datagridview As DataGridView,
                        par_datetimepicker_inizio As DateTimePicker,
                        par_datetimepicker_fine As DateTimePicker,
                        par_filtro_documento As String,
                        par_filtro_n_documento As String,
                        par_filtro_osservazioni As String,
                        par_filtro_mag As String,
                        par_filtro_comm As String)

        par_datagridview.Rows.Clear()

        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Cnn.Open()

            Dim CMD_SAP As New SqlCommand()
            CMD_SAP.Connection = Cnn

            If Homepage.ERP_provenienza = "SAP" Then
                CMD_SAP.CommandText = "
select *, '' as 'Sottocommessa', '' as 'Matricola','' as 'rifmag', '' as num_ope
from
(
    SELECT  T0.[DocDate],t0.transtype,t0.transseq,
    case when t0.transtype='10000071' then 'CS' 
         when T0.[TransType] = '67' then 'Trasf MAG'  
         WHEN t0.transtype ='59' then 'EM P' 
         WHEN t0.transtype =15  then 'BC' 
         WHEN t0.transtype =18  then 'UP' 
         WHEN t0.transtype =20  then 'EM Forn' 
         when t0.transtype =60 then 'UP' 
         else t11.JrnlMemo end as 'Transname',
    T11.[Ref1],  T11.COMMENT, T0.LocCode, T0.[InQty]- T0.[OutQty] as 'Movimento', 
    CASE WHEN T0.TRANSTYPE=59 THEN T7.U_UTILIZZ 
         when T0.TRANSTYPE=60 then t10.u_utilizz
         when T0.TRANSTYPE=67 and T3.DOCENTRY= T2.U_PRG_AZS_OPDOCENTRY then t3.U_UTILIZZ
         when T0.TRANSTYPE=67 and T4.DOCENTRY= T2.U_PRG_AZS_OcDocEntry then concat('OC ',t4.cardname) 
         else T11.[CardName] end as 'Cardname',
    CASE WHEN T0.TRANSTYPE=20 THEN T1.U_PRG_AZS_COMMESSA  
         WHEN T3.U_PRG_AZS_COMMESSA <> '' THEN T3.U_PRG_AZS_COMMESSA
         when T0.TRANSTYPE=59 then t7.U_PRG_AZS_Commessa
         when T0.TRANSTYPE=60 then t10.U_PRG_AZS_Commessa
         when T0.TRANSTYPE=67 and T3.DOCENTRY= T2.U_PRG_AZS_OPDOCENTRY then t3.U_PRG_AZS_Commessa
         when T0.TRANSTYPE=67 and T4.DOCENTRY= T2.U_PRG_AZS_OCdocentry then concat('OC ',t4.cardname) 
         ELSE T4.CARDNAME  END AS 'U_PRG_AZS_COMMESSA',
    case when t0.transtype<>67 and t0.transtype<>'10000071' then t0.price end as 'Price',
    t8.lastname
    FROM OIVL T0
    LEFT JOIN PDN1 T1 ON T0.CREATEDBY=T1.DOCENTRY AND T0.DOCLINENUM=T1.LINENUM AND T0.TRANSTYPE=20
    LEFT JOIN WTR1 T2 ON T0.CREATEDBY=T2.DOCENTRY AND T0.TRANSTYPE=67
    LEFT JOIN OWOR T3 ON T3.DOCENTRY= T2.U_PRG_AZS_OPDOCENTRY
    LEFT JOIN ORDR T4 ON T4.DOCENTRY= T2.U_PRG_AZS_OcDocEntry
    LEFT JOIN [OILM] T11  ON  T0.[MessageID] = T11.[MessageID] 
    left join oign t5 on cast(t5.docnum as varchar)=cast(T11.[Ref1] as varchar) and cast(t0.transtype as varchar)='59'
    left join ign1 t6 on t6.docentry=t5.docentry and t6.itemcode=t0.itemcode and t0.doclinenum=t6.linenum
    left join owor t7 on t7.docnum=t6.baseref
    left join [TIRELLI_40].[dbo].ohem t8 on t8.userid=t0.usersign
    left join ige1 t9 on t9.itemcode=t0.itemcode and T0.CREATEDBY=T9.DOCENTRY AND T0.TRANSTYPE=60 and t0.doclinenum=t9.linenum
    left join owor t10 on t10.docnum=t9.baseref
    WHERE T0.[ItemCode] = @itemCode
      AND T0.DocDate >= @dataInizio
      AND T0.DocDate <= @dataFine
      AND T0.[InQty]-T0.[OutQty]<>0
    GROUP BY t11.taxdate, t9.baseref,t3.U_UTILIZZ,T2.U_PRG_AZS_OCdocentry,T4.DOCENTRY,T2.U_PRG_AZS_OPDOCENTRY,
             T3.DOCENTRY,t0.transseq,T0.[DocDate], T0.[TransType],T11.[Ref1], T0.[Price], T11.COMMENT, 
             T0.LOCCODE,T0.[InQty],T0.[OutQty],T11.[CardName],T1.U_PRG_AZS_COMMESSA,T3.U_PRG_AZS_COMMESSA,
             T4.CARDNAME, T0.TRANSSEQ,t7.U_PRG_AZS_Commessa, t11.JrnlMemo, t7.u_utilizz, t8.lastname, 
             t10.U_UTILIZZ, t10.U_prg_azs_commessa
) as t20
where 0=0 " & par_filtro_documento & par_filtro_n_documento & par_filtro_osservazioni & par_filtro_comm & par_filtro_mag & "
ORDER BY t20.docdate,t20.transseq;
"
            Else


                CMD_SAP.CommandText = "
SELECT *, case when cast(codcau as varchar)='10' then t10.fornitore else t10.cliente_commessa end as 'Cardname'
FROM OPENQUERY(AS400, '
    SELECT
t0.codart,
        DATE(
            SUBSTR(CHAR(t0.docdate),1,4) || ''-'' ||
            SUBSTR(CHAR(t0.docdate),5,2) || ''-'' ||
            SUBSTR(CHAR(t0.docdate),7,2)
        ) AS DocDate,
time,
		   DATE(
            SUBSTR(CHAR(t0.data_reg),1,4) || ''-'' ||
            SUBSTR(CHAR(t0.data_reg),5,2) || ''-'' ||
            SUBSTR(CHAR(t0.data_reg),7,2)
        ) AS datareg,
		num_ope,
		rig_ope,
        t0.transname,
        t0.ref1,
        trim(commento) AS Comment,
        t0.loccode,
	t0.segno,
        case when segno =''-'' then -t0.movimento else t0.movimento end as movimento,
        case when t0.codcomm='''' then trim(t2.cod_commessa) else trim(t0.codcomm) end AS U_PRG_AZS_COMMESSA,
case when T0.MATRICOLA ='''' then trim(t2.matricola) else trim(t0.matricola) end as matricola,
case when T0.SOTTOCOMMESSA='''' then trim(t2.cod_sottocommessa) else trim(t0.sottocommessa) end as sottocommessa,
       cardname as Fornitore,
	   codcau,
        t0.ds_clicom as Cliente_commessa
,
        t0.price,
        coalesce(t0.lastname,'''') as Codice_galileo
,coalesce(t1.cogn_dip,'''') as lastname
,t0.rifmag
    FROM TIR90VIS.JGALMOV t0
left join TIR90VIS.JGALDIP t1 on t1.prof_gal=t0.lastname
left join TIR90VIS.JGALodp t2 on trim(t0.commento)=t2.numodp and trim(t0.commento)<>''''
    WHERE t0.codart = ''" & par_codice_sap & "'' " & par_filtro_documento & par_filtro_n_documento & par_filtro_osservazioni & par_filtro_comm & par_filtro_mag & "
      AND DATE(
            SUBSTR(CHAR(t0.docdate),1,4) || ''-'' ||
            SUBSTR(CHAR(t0.docdate),5,2) || ''-'' ||
            SUBSTR(CHAR(t0.docdate),7,2)
          ) >= ''" & par_datetimepicker_inizio.Value.ToString("yyyy-MM-dd") & "'' 
      AND DATE(
            SUBSTR(CHAR(t0.docdate),1,4) || ''-'' ||
            SUBSTR(CHAR(t0.docdate),5,2) || ''-'' ||
            SUBSTR(CHAR(t0.docdate),7,2)
          ) <= ''" & par_datetimepicker_fine.Value.ToString("yyyy-MM-dd") & "'' 
    ORDER BY DATE(
        SUBSTR(CHAR(t0.docdate),1,4) || ''-'' ||
        SUBSTR(CHAR(t0.docdate),5,2) || ''-'' ||
        SUBSTR(CHAR(t0.docdate),7,2)
    ), num_ope,rig_ope,segno desc
') T10
"

            End If



            ' Parametri sicuri
            CMD_SAP.Parameters.AddWithValue("@itemCode", par_codice_sap)
            CMD_SAP.Parameters.AddWithValue("@dataInizio", par_datetimepicker_inizio.Value.Date)
            CMD_SAP.Parameters.AddWithValue("@dataFine", par_datetimepicker_fine.Value.Date)

            Dim cmd_SAP_reader As SqlDataReader = CMD_SAP.ExecuteReader()

            Do While cmd_SAP_reader.Read()
                par_datagridview.Rows.Add(cmd_SAP_reader("DocDate"),
                                      cmd_SAP_reader("TRANSname"),
                                      cmd_SAP_reader("REF1"),
                                      cmd_SAP_reader("Comment"),
                                      cmd_SAP_reader("LOCCODE"),
                                      cmd_SAP_reader("Movimento"),
                                      cmd_SAP_reader("U_PRG_AZS_COMMESSA"),
                                      cmd_SAP_reader("SOTTOCOMMESSA"),
                                      cmd_SAP_reader("Matricola"),
                                      cmd_SAP_reader("Cardname"),
                                      cmd_SAP_reader("Price"),
                                      cmd_SAP_reader("lastname"), cmd_SAP_reader("rifmag"), cmd_SAP_reader("num_ope"))
            Loop

            cmd_SAP_reader.Close()
        End Using

        Try
            par_datagridview.FirstDisplayedScrollingRowIndex = par_datagridview.RowCount - 1
        Catch ex As Exception
            ' Se la griglia è vuota, ignora
        End Try
    End Sub

    Sub allegati(par_datagridview As DataGridView, par_itemcode As String, par_tabcontrol As TabControl)
        If Homepage.ERP_provenienza = "SAP" Then




            Button6.BackColor = Color.OrangeRed
            par_datagridview.Rows.Clear()
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()

            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

            CMD_SAP.Connection = Cnn

            CMD_SAP.CommandText = "select top 100 t1.itemcode,t0.trgtPath, t0.filename,t0.fileext,t0.date,t0.usrid, concat(t2.lastname,' ',t2.firstname) as 'Nome_ID',  *
from atc1 t0 inner join oitm t1 on t0.absentry=t1.AtcEntry
left join [TIRELLI_40].[dbo].ohem t2 on t2.userid=t0.usrid

where t1.itemcode='" & par_itemcode & "'"

            cmd_SAP_reader = CMD_SAP.ExecuteReader

            Do While cmd_SAP_reader.Read()
                par_datagridview.Rows.Add("SAP", cmd_SAP_reader("trgtPath"), cmd_SAP_reader("filename"), cmd_SAP_reader("fileext"), cmd_SAP_reader("date"), cmd_SAP_reader("usrid"), cmd_SAP_reader("Nome_ID"))
                ' Label2.ForeColor = Color.Lime
            Loop
            cmd_SAP_reader.Close()
            Cnn.Close()



            allegati_custom(par_datagridview, par_itemcode, par_tabcontrol)


            par_datagridview.ClearSelection()

        Else

        End If
    End Sub 'Inserisco le risorse nella combo box

    Sub trova_codici_in_distinta(par_datagridview As DataGridView, par_itemcode As String)

        par_datagridview.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "select t0.father, t1.itemname,coalesce(t1.u_disegno,'') as 'u_Disegno', t0.quantity
from itt1 t0
inner join oitm t1 on t1.itemcode=t0.father

 where t0.code='" & par_itemcode & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader("father"), cmd_SAP_reader("itemname"), cmd_SAP_reader("u_disegno"), cmd_SAP_reader("quantity"))

        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()


        par_datagridview.ClearSelection()


    End Sub 'Inserisco le risorse nella combo box

    Sub trova_codice_in_odp(par_datagridview As DataGridView, par_itemcode As String)

        par_datagridview.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "SELECT T1.[DocNum], T1.[ItemCode], t3.itemname, coalesce(t3.u_disegno,'') as 'u_Disegno', T1.[PlannedQty], T1.[Status], T1.[PostDate], T1.CLOSEDATE,T0.[ItemCode], T0.[PlannedQty], T1.[U_PRG_AZS_Commessa] , coalesce(t2.itemname,'') as 'Desc_commessa', coalesce(t2.u_final_customer_name,'') as 'u_final_customer_name'

FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] 
LEFT JOIN OITM T2 ON T2.ITEMCODE=T1.[U_PRG_AZS_Commessa] 
left join oitm t3 on t3.itemcode=t1.itemcode

WHERE T0.ItemCode ='" & par_itemcode & "' AND T1.[Status]<>'C'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader("docnum"), cmd_SAP_reader("ItemCode"), cmd_SAP_reader("itemname"), cmd_SAP_reader("u_disegno"), cmd_SAP_reader("PlannedQty"), cmd_SAP_reader("status"), cmd_SAP_reader("postdate"), cmd_SAP_reader("U_PRG_AZS_Commessa"), cmd_SAP_reader("desc_commessa"), cmd_SAP_reader("u_final_customer_name"))

        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()


        par_datagridview.ClearSelection()


    End Sub 'Inserisco le risorse nella combo box

    Sub allegati_custom(par_datagridview As DataGridView, par_itemcode As String, par_tabcontrol As TabControl)


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "SELECT TOP (1000) t0.[ID]
      ,t0.[Codice]
      ,t0.[Filename]
      ,t0.[FileExt]
      ,t0.[Date]
      ,t0.[UsrID]
  FROM [TIRELLI_40].[DBO].[Allegati] t0
  inner join oitm t1 on t1.itemcode=t0.codice
  left join [TIRELLI_40].[dbo].ohem t2 on t2.empid=t0.usrid
where t1.itemcode='" & par_itemcode & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add("Custom", "", cmd_SAP_reader("filename"), cmd_SAP_reader("fileext"), cmd_SAP_reader("date"), cmd_SAP_reader("usrid"))
            'Label2.ForeColor = Color.Lime
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()




        par_datagridview.ClearSelection()


    End Sub 'Inserisco le risorse nella combo box


    Sub consumi()
        If Homepage.ERP_provenienza = "SAP" Then


            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()

            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

            CMD_SAP.Connection = Cnn

            CMD_SAP.CommandText = "declare @data_inizio_anno_1 as date
declare @data_fine_anno_1 as date
declare @giorni_considerati_per_consumo as integer
declare @giorni_considerati_per_acquisti as integer
declare @giorni_considerati_per_lead_time_fornitori as integer
declare @giorni_considerati_per_lead_time_MU as integer

set @data_inizio_anno_1 = CONVERT(DATETIME, '20221001', 112)
set @data_fine_anno_1 = CONVERT(DATETIME, '20230930', 112)
set @giorni_considerati_per_consumo = 1095
set @giorni_considerati_per_acquisti = 365
set @giorni_considerati_per_lead_time_fornitori = 1500
set @giorni_considerati_per_lead_time_mu = 730

select  substring(t0.itemcode,1,1) as '|',t0.itemcode,t0.itemname, t0.U_Disegno,t0.frozenfor, t0.u_prg_tir_materiale, t1.ItmsGrpNam, t4.firmname,t5.cardname as 'Fornitore_pref',d.docdate as 'Ultimo_acq',d.CardName as 'Ultimo_forn', t0.U_Gestione_magazzino, t0.MinLevel, t0.MinOrdrQty, a.[Consumo Y],a.[Consumo Y-1],a.[Consumo Y-2], a.[N°_scarichi], sum(case when t2.onhand is null then 0 else t2.onhand end) as 'Mag_tot',  sum(case when t2.onhand is null then 0 else t2.onhand end)- sum(case when t2.iscommited is null then 0 else t2.iscommited end) + sum(case when t2.onorder is null then 0 else t2.onorder end) as 'Disp', t3.Price AS 'Prezzo_listino_acquisto',f.Prezzo_medio_acquisto, b.Lead_time_fornitori, c.Lead_time_mu, E.TOTALE AS 'FATTURATO_FORN_Y', f.N_ordini, f.Pezzi_ordinati
from oitm t0 inner join oitb t1 on t0.ItmsGrpCod=t1.ItmsGrpCod

left join (select t1.itemcode,  sum(case when t1.TaxDate>=getdate()-@giorni_considerati_per_consumo/3 then t1.outqty else 0 end) as 'Consumo Y', sum(case when t1.TaxDate>=getdate()-@giorni_considerati_per_consumo*2/3 and t1.TaxDate<getdate()-@giorni_considerati_per_consumo*1/3 then t1.outqty else 0 end) as 'Consumo Y-1', sum( case when t1.TaxDate>=getdate()-@giorni_considerati_per_consumo and t1.TaxDate<getdate()-@giorni_considerati_per_consumo*2/3 then t1.outqty else 0 end) as 'Consumo Y-2', sum(case when t1.TaxDate>=getdate()-@giorni_considerati_per_consumo/3 and t1.OutQty>0 then 1 else 0 end) as 'N°_scarichi'
from oinm t1
where t1.ref1<>'61892' and t1.transtype <>'10000071' and t1.transtype <>'67' and t1.TaxDate>=getdate()-@giorni_considerati_per_consumo
group by t1.itemcode) A on t0.itemcode=a.itemcode

left join oitw t2 on t2.itemcode=t0.itemcode
left join itm1 t3 on t3.itemcode=t0.itemcode

left join (select t1.itemcode, sum((DATEDIFF(dd, t3.docdate,t0.docdate) + 1)
  -(DATEDIFF(wk, t3.docdate,t0.docdate) * 2)
  -(CASE WHEN DATENAME(dw, t3.docdate) = 'Sunday' THEN 1 ELSE 0 END)
  -(CASE WHEN DATENAME(dw, t0.docdate) = 'Saturday' THEN 1 ELSE 0 END))/count(t1.itemcode) as 'Lead_time_fornitori'
from opdn t0 inner join pdn1 t1 on t0.docentry=t1.docentry
left join por1 t2 on t2.docentry=t1.BaseEntry and t2.itemcode=t1.itemcode
left join opor t3 on t3.docentry=t2.docentry
where t0.docdate>=getdate()-@giorni_considerati_per_lead_time_fornitori

group by t1.itemcode) B on b.itemcode=t0.itemcode

left join (select t0.itemcode, sum((DATEDIFF(dd, t0.postdate,t0.closedate) + 1)
  -(DATEDIFF(wk, t0.postdate,t0.closedate) * 2)
  -(CASE WHEN DATENAME(dw, t0.postdate) = 'Sunday' THEN 1 ELSE 0 END)
  -(CASE WHEN DATENAME(dw, t0.closedate) = 'Saturday' THEN 1 ELSE 0 END))/count(t0.itemcode) as 'Lead_time_mu'
from owor t0
where t0.CloseDate>=getdate()-@giorni_considerati_per_lead_time_MU and substring(t0.u_produzione,1,3)='INT'
group by t0.itemcode) C on c.itemcode=t0.itemcode
left join omrc t4 on t4.FirmCode=t0.firmcode
left join ocrd t5 on t0.CardCode=t5.CardCode

left join (select t10.itemcode, t11.docdate, t11.cardname
from (
Select t0.itemcode, max(t0.docentry) as 'Docentry' from por1 t0 
group by t0.itemcode
)
as t10 inner join opor t11 on t11.docentry=t10.docentry) D on d.itemcode=t0.itemcode

LEFT JOIN (SELECT T10.ITEMCODE, SUM(T10.TOTALE) AS 'TOTALE'
FROM
(
SELECT T1.ITEMCODE, SUM(CASE WHEN T1.LINETOTAL*(100-T0.DiscPrcnt)/100 IS NULL THEN 0 ELSE T1.LINETOTAL*(100-T0.DiscPrcnt)/100 END) AS 'TOTALE'
FROM OPCH T0 INNER JOIN PCH1 T1 ON T0.DOCENTRY=T1.DOCENTRY
WHERE T0.TAXDATE>=GETDATE()-365
GROUP BY T1.ITEMCODE

UNION ALL

SELECT T1.ITEMCODE, -SUM(CASE WHEN T1.LINETOTAL*(100-T0.DiscPrcnt)/100 IS NULL THEN 0 ELSE T1.LINETOTAL*(100-T0.DiscPrcnt)/100 END ) AS 'TOTALE'
FROM ORPC T0 INNER JOIN RPC1 T1 ON T0.DOCENTRY=T1.DOCENTRY
WHERE T0.TAXDATE>=GETDATE()-365
GROUP BY T1.ITEMCODE
)
AS T10
GROUP BY T10.ITEMCODE ) E ON E.ITEMCODE=T0.ItemCode

left join (
select t10.ItemCode, COUNT(t10.itemcode) as 'N_ordini', sum(t10.quantity) as 'Pezzi_ordinati',sum(t10.Prezzo_riga) as 'Valore_ordinato', sum(t10.Prezzo_riga)/sum(t10.quantity) as 'Prezzo_medio_acquisto'
from
(
select  t1.itemcode, t1.quantity, t1.LineTotal*(100-t0.DiscPrcnt)/100 as 'Prezzo_riga', t1.LineTotal*(100-t0.DiscPrcnt)/100/t1.quantity as 'P_unitario'
from oPDN t0 LEFT JOIN PDN1 T1 ON T0.DOCENTRY=T1.DocEntry
where t0.TAXDATE>=getdate()-@giorni_considerati_per_acquisti and t0.doctype='I'

)
as t10
group by t10.ItemCode
) F on f.itemcode=t0.itemcode



where t0.itemcode='" & Codice_SAP & "' and t3.pricelist=2
group by t0.itemcode,t0.itemname, t0.U_Disegno,t0.frozenfor, t1.ItmsGrpNam,t0.U_Gestione_magazzino, t0.MinLevel, t0.MinOrdrQty, a.[Consumo Y],a.[Consumo Y-1],a.[Consumo Y-2],a.[N°_scarichi],t3.Price,b.Lead_time_fornitori, c.lead_time_mu, t4.firmname,t5.cardname, d.CardName,d.docdate,E.TOTALE,f.N_ordini,f.Pezzi_ordinati,f.Prezzo_medio_acquisto, t0.u_prg_tir_materiale"

            cmd_SAP_reader = CMD_SAP.ExecuteReader

            If cmd_SAP_reader.Read() Then


                If Not cmd_SAP_reader("Consumo Y") Is System.DBNull.Value Then
                    Label13.Text = CType(cmd_SAP_reader("Consumo Y"), Integer)
                Else
                    Label13.Text = 0
                End If

                If Not cmd_SAP_reader("Consumo Y-1") Is System.DBNull.Value Then
                    Label12.Text = CType(cmd_SAP_reader("Consumo Y-1"), Integer)
                Else
                    Label12.Text = 0
                End If

                If Not cmd_SAP_reader("Consumo Y-2") Is System.DBNull.Value Then
                    Label11.Text = CType(cmd_SAP_reader("Consumo Y-2"), Integer)
                Else
                    Label11.Text = 0
                End If

                If Not cmd_SAP_reader("N°_scarichi") Is System.DBNull.Value Then
                    Label14.Text = cmd_SAP_reader("N°_scarichi")
                Else
                    Label14.Text = ""
                End If

                If Not cmd_SAP_reader("Ultimo_forn") Is System.DBNull.Value Then
                    Label16.Text = cmd_SAP_reader("Ultimo_forn")
                Else
                    Label16.Text = ""
                End If

                If Not cmd_SAP_reader("ultimo_acq") Is System.DBNull.Value Then
                    Label15.Text = cmd_SAP_reader("ultimo_acq")
                Else
                    Label15.Text = ""
                End If




                If Not cmd_SAP_reader("Lead_time_fornitori") Is System.DBNull.Value Then
                    Label17.Text = cmd_SAP_reader("Lead_time_fornitori")
                Else
                    Label17.Text = "-"
                End If



            End If
            cmd_SAP_reader.Close()
            Cnn.Close()

        Else
            Label13.Text = 9999
            Label12.Text = 9999
            Label11.Text = 9999
            Label14.Text = 9999
            Label16.Text = 9999
            Label17.Text = "9999"
        End If

    End Sub 'Inserisco le risorse nella combo box

    Private Sub DateTimePicker1_ValueChanged_1(sender As Object, e As EventArgs)
        Lista_registrazioni(Codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, filtro_doc, filtro_N_doc, filtro_osservazioni, filtro_mag, filtro_Comm)
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs)
        Lista_registrazioni(Codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, filtro_doc, filtro_N_doc, filtro_osservazioni, filtro_mag, filtro_Comm)
    End Sub




    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button_trasferibili.Click
        Form_ritiri.Show()


    End Sub

    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        Lista_registrazioni(Codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, filtro_doc, filtro_N_doc, filtro_osservazioni, filtro_mag, filtro_Comm)
    End Sub

    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        Lista_registrazioni(Codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, filtro_doc, filtro_N_doc, filtro_osservazioni, filtro_mag, filtro_Comm)
    End Sub



    Private Sub DataGridView_ordinato_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ordinato.CellClick
        Dim par_datagridview As DataGridView
        par_datagridview = DataGridView_ordinato
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = par_datagridview.Columns.IndexOf(N_doc_) And par_datagridview.Rows(e.RowIndex).Cells(columnName:="Doc_ord").Value = "OP" Then



                ODP_Form.docnum_odp = par_datagridview.Rows(e.RowIndex).Cells(columnName:="N_doc_").Value
                ODP_Form.Show()
                ODP_Form.inizializza_form(par_datagridview.Rows(e.RowIndex).Cells(columnName:="N_doc_").Value)

            End If
        End If
    End Sub

    Sub cambiare_gestione_ferretto(par_utente_sap_aggiornatore)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE T0 SET T0.[QryGroup10]= case when t0.[QryGroup10]='Y' then 'N' else 'Y' end,  t0.usersign2='" & par_utente_sap_aggiornatore & "',t0.[UpdateDate]=getdate(),t0.updatets=concat ( DATEPART(hour, GETDATE()),DATEPART(minute, GETDATE()), DATEPART(second, GETDATE()))
FROM OITM T0 WHERE T0.[ItemCode] ='" & Codice_SAP & "' "
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Sub cambiare_gestione_minimo(par_utente_sap_aggiornatore As String)

        Dim minimo As String


        minimo = "minlevel"
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE T0 SET T0.[" & minimo & "]= " & nuovo_valore & ",  t0.usersign2='" & par_utente_sap_aggiornatore & "',t0.[UpdateDate]=getdate(), t0.updatets=concat ( DATEPART(hour, GETDATE()),DATEPART(minute, GETDATE()), DATEPART(second, GETDATE()))

FROM OITM T0 WHERE T0.[ItemCode] ='" & Codice_SAP & "' "
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Sub UPDATE_OITM(par_utente_SAP_AGGIORNAMEnto As String, par_codice_articolo As String, par_descrizione_articolo As String, par_desc_supp As String, par_codice_disegno As String, par_fornitore_preferito As String, par_catalogo_fornitore As String, par_produttore As String, par_tipo_montaggio As String, par_codice_BP As String, par_nome_bp As String, par_settore As String, par_paese As String, par_agente As String, par_brand As String, par_codice_BRB As String, par_revisione As String)

        par_descrizione_articolo = Replace(par_descrizione_articolo, "'", "")
        par_desc_supp = Replace(par_desc_supp, "'", "")
        par_codice_disegno = Replace(par_codice_disegno, "'", "")

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE T0 SET 

t0.usersign2='" & par_utente_SAP_AGGIORNAMEnto & "', t0.[UpdateDate]=getdate(), t0.updatets=concat ( DATEPART(hour, GETDATE()),DATEPART(minute, GETDATE()), DATEPART(second, GETDATE()))
,t0.itemname='" & par_descrizione_articolo & "'
,t0.frgnname ='" & par_desc_supp & "'
,t0.u_disegno='" & par_codice_disegno & "'
,t0.u_prg_tir_rev='" & par_revisione & "'


 FROM OITM T0 

WHERE T0.[ItemCode] ='" & par_codice_articolo & "' "
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Sub update_AITM(par_codice_articolo As String, par_utente_Sap As Integer) 'istanze di registro SAP
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "DECLARE @logistanza AS INT
DECLARE @aggiornatore AS INT
DECLARE @codice_Articolo AS VARCHAR(50)


SET @aggiornatore = " & par_utente_Sap & "
SET @codice_Articolo = '" & par_codice_articolo & "'

SELECT @logistanza = MAX(COALESCE(t0.loginstanc, 0))+1
FROM aitm t0
WHERE t0.itemcode = @codice_Articolo

select @logistanza=coalesce(@logistanza,1)


insert into aitm (loginstanc, [ItemCode]
      ,[ItemName]
      ,[FrgnName]
      ,[ItmsGrpCod]
      ,[CstGrpCode]
      ,[VatGourpSa]
      ,[CodeBars]
      ,[VATLiable]
      ,[PrchseItem]
      ,[SellItem]
      ,[InvntItem]
      ,[OnHand]
      ,[IsCommited]
      ,[OnOrder]
      ,[IncomeAcct]
      ,[ExmptIncom]
      ,[MaxLevel]
      ,[DfltWH]
      ,[CardCode]
      ,[SuppCatNum]
      ,[BuyUnitMsr]
      ,[NumInBuy]
      ,[ReorderQty]
      ,[MinLevel]
      ,[LstEvlPric]
      ,[LstEvlDate]
      ,[CustomPer]
      ,[Canceled]
      ,[MnufctTime]
      ,[WholSlsTax]
      ,[RetilrTax]
      ,[SpcialDisc]
      ,[DscountCod]
      ,[TrackSales]
      ,[SalUnitMsr]
      ,[NumInSale]
      ,[Consig]
      ,[QueryGroup]
      ,[Counted]
      ,[OpenBlnc]
      ,[EvalSystem]
      ,[UserSign]
      ,[FREE]
      ,[PicturName]
      ,[Transfered]
      ,[BlncTrnsfr]
      ,[UserText]
      ,[SerialNum]
      ,[CommisPcnt]
      ,[CommisSum]
      ,[CommisGrp]
      ,[TreeType]
      ,[TreeQty]
      ,[LastPurPrc]
      ,[LastPurCur]
      ,[LastPurDat]
      ,[ExitCur]
      ,[ExitPrice]
      ,[ExitWH]
      ,[AssetItem]
      ,[WasCounted]
      ,[ManSerNum]
      ,[SHeight1]
      ,[SHght1Unit]
      ,[SHeight2]
      ,[SHght2Unit]
      ,[SWidth1]
      ,[SWdth1Unit]
      ,[SWidth2]
      ,[SWdth2Unit]
      ,[SLength1]
      ,[SLen1Unit]
      ,[Slength2]
      ,[SLen2Unit]
      ,[SVolume]
      ,[SVolUnit]
      ,[SWeight1]
      ,[SWght1Unit]
      ,[SWeight2]
      ,[SWght2Unit]
      ,[BHeight1]
      ,[BHght1Unit]
      ,[BHeight2]
      ,[BHght2Unit]
      ,[BWidth1]
      ,[BWdth1Unit]
      ,[BWidth2]
      ,[BWdth2Unit]
      ,[BLength1]
      ,[BLen1Unit]
      ,[Blength2]
      ,[BLen2Unit]
      ,[BVolume]
      ,[BVolUnit]
      ,[BWeight1]
      ,[BWght1Unit]
      ,[BWeight2]
      ,[BWght2Unit]
      ,[FixCurrCms]
      ,[FirmCode]
      ,[LstSalDate]
      ,[QryGroup1]
      ,[QryGroup2]
      ,[QryGroup3]
      ,[QryGroup4]
      ,[QryGroup5]
      ,[QryGroup6]
      ,[QryGroup7]
      ,[QryGroup8]
      ,[QryGroup9]
      ,[QryGroup10]
      ,[QryGroup11]
      ,[QryGroup12]
      ,[QryGroup13]
      ,[QryGroup14]
      ,[QryGroup15]
      ,[QryGroup16]
      ,[QryGroup17]
      ,[QryGroup18]
      ,[QryGroup19]
      ,[QryGroup20]
      ,[QryGroup21]
      ,[QryGroup22]
      ,[QryGroup23]
      ,[QryGroup24]
      ,[QryGroup25]
      ,[QryGroup26]
      ,[QryGroup27]
      ,[QryGroup28]
      ,[QryGroup29]
      ,[QryGroup30]
      ,[QryGroup31]
      ,[QryGroup32]
      ,[QryGroup33]
      ,[QryGroup34]
      ,[QryGroup35]
      ,[QryGroup36]
      ,[QryGroup37]
      ,[QryGroup38]
      ,[QryGroup39]
      ,[QryGroup40]
      ,[QryGroup41]
      ,[QryGroup42]
      ,[QryGroup43]
      ,[QryGroup44]
      ,[QryGroup45]
      ,[QryGroup46]
      ,[QryGroup47]
      ,[QryGroup48]
      ,[QryGroup49]
      ,[QryGroup50]
      ,[QryGroup51]
      ,[QryGroup52]
      ,[QryGroup53]
      ,[QryGroup54]
      ,[QryGroup55]
      ,[QryGroup56]
      ,[QryGroup57]
      ,[QryGroup58]
      ,[QryGroup59]
      ,[QryGroup60]
      ,[QryGroup61]
      ,[QryGroup62]
      ,[QryGroup63]
      ,[QryGroup64]
      ,[CreateDate]
      ,[UpdateDate]
      ,[ExportCode]
      ,[SalFactor1]
      ,[SalFactor2]
      ,[SalFactor3]
      ,[SalFactor4]
      ,[PurFactor1]
      ,[PurFactor2]
      ,[PurFactor3]
      ,[PurFactor4]
      ,[SalFormula]
      ,[PurFormula]
      ,[VatGroupPu]
      ,[AvgPrice]
      ,[PurPackMsr]
      ,[PurPackUn]
      ,[SalPackMsr]
      ,[SalPackUn]
      ,[SCNCounter]
      ,[ManBtchNum]
      ,[ManOutOnly]
      ,[DataSource]
      ,[validFor]
      ,[validFrom]
      ,[validTo]
      ,[frozenFor]
      ,[frozenFrom]
      ,[frozenTo]
      ,[BlockOut]
      ,[ValidComm]
      ,[FrozenComm]
      ,[ObjType]
      ,[SWW]
      ,[Deleted]
      ,[DocEntry]
      ,[ExpensAcct]
      ,[FrgnInAcct]
      ,[ShipType]
      ,[GLMethod]
      ,[ECInAcct]
      ,[FrgnExpAcc]
      ,[ECExpAcc]
      ,[TaxType]
      ,[ByWh]
      ,[WTLiable]
      ,[ItemType]
      ,[WarrntTmpl]
      ,[BaseUnit]
      ,[CountryOrg]
      ,[StockValue]
      ,[Phantom]
      ,[IssueMthd]
      ,[FREE1]
      ,[PricingPrc]
      ,[MngMethod]
      ,[ReorderPnt]
      ,[InvntryUom]
      ,[PlaningSys]
      ,[PrcrmntMtd]
      ,[OrdrIntrvl]
      ,[OrdrMulti]
      ,[MinOrdrQty]
      ,[LeadTime]
      ,[IndirctTax]
      ,[TaxCodeAR]
      ,[TaxCodeAP]
      ,[OSvcCode]
      ,[ISvcCode]
      ,[ServiceGrp]
      ,[NCMCode]
      ,[MatType]
      ,[MatGrp]
      ,[ProductSrc]
      ,[ServiceCtg]
      ,[ItemClass]
      ,[Excisable]
      ,[ChapterID]
      ,[NotifyASN]
      ,[ProAssNum]
      ,[AssblValue]
      ,[DNFEntry]
      ,[UserSign2]
      ,[Spec]
      ,[TaxCtg]
      ,[Series]
      ,[Number]
      ,[FuelCode]
      ,[BeverTblC]
      ,[BeverGrpC]
      ,[BeverTM]
      ,[Attachment]
      ,[AtcEntry]
      ,[ToleranDay]
      ,[UgpEntry]
      ,[PUoMEntry]
      ,[SUoMEntry]
      ,[IUoMEntry]
      ,[IssuePriBy]
      ,[AssetClass]
      ,[AssetGroup]
      ,[InventryNo]
      ,[Technician]
      ,[Employee]
      ,[Location]
      ,[StatAsset]
      ,[Cession]
      ,[DeacAftUL]
      ,[AsstStatus]
      ,[CapDate]
      ,[AcqDate]
      ,[RetDate]
      ,[GLPickMeth]
      ,[NoDiscount]
      ,[MgrByQty]
      ,[AssetRmk1]
      ,[AssetRmk2]
      ,[AssetAmnt1]
      ,[AssetAmnt2]
      ,[DeprGroup]
      ,[AssetSerNo]
      ,[CntUnitMsr]
      ,[NumInCnt]
      ,[INUoMEntry]
      ,[OneBOneRec]
      ,[RuleCode]
      ,[ScsCode]
      ,[SpProdType]
      ,[IWeight1]
      ,[IWght1Unit]
      ,[IWeight2]
      ,[IWght2Unit]
      ,[CompoWH]
      ,[CreateTS]
      ,[UpdateTS]
      ,[VirtAstItm]
      ,[SouVirAsst]
      ,[InCostRoll]
      ,[PrdStdCst]
      ,[EnAstSeri]
      ,[LinkRsc]
      ,[OnHldPert]
      ,[onHldLimt]
      ,[PriceUnit]
      ,[GSTRelevnt]
      ,[SACEntry]
      ,[GstTaxCtg]
      ,[AssVal4WTR]
      ,[ExcImpQUoM]
      ,[ExcFixAmnt]
      ,[ExcRate]
      ,[SOIExc]
      ,[TNVED]
      ,[Imported]
      ,[AutoBatch]
      ,[CstmActing]
      ,[StdItemId]
      ,[CommClass]
      ,[TaxCatCode]
      ,[DataVers]
      ,[NVECode]
      ,[CESTCode]
      ,[CtrSealQty]
      ,[LegalText]
      ,[QRCodeSrc]
      ,[Traceable]
      ,[U_UBIMAG]
      ,[U_TEMPOMED]
      ,[U_MODMAC]
      ,[U_SEZIONE]
      ,[U_DESCFR]
      ,[U_DESCING]
      ,[U_UTILIZZ]
      ,[U_ARTCES]
      ,[U_BA_IsFA]
      ,[U_BA_TypID]
      ,[U_BA_NumID]
      ,[U_BA_LVAFrom]
      ,[U_BA_LVA]
      ,[U_ULTVER]
      ,[U_ULTPREZ]
      ,[U_PrzPeso]
      ,[U_PRG_AZS_DesAggArt]
      ,[U_PRG_AZS_DesAgg2Art]
      ,[U_PRG_AZS_CMM]
      ,[U_PRG_AZS_ItmsGrp2Cod]
      ,[U_PRG_AZS_GestComm]
      ,[U_PRG_AZS_CreatedBy]
      ,[U_PRG_AZS_Phantom]
      ,[U_PRG_CLV_Tipo_Lav]
      ,[U_PRG_CLV_Lav_EstAss]
      ,[U_PRG_CLV_Grezzo]
      ,[U_PRG_CLV_TipologiaLav]
      ,[U_PRG_CLV_ArtAssLav]
      ,[U_PRG_CLV_SemiLavComodo]
      ,[U_PRG_CLV_MagPrePro]
      ,[U_Anagrafica]
      ,[U_Ricambio]
      ,[U_Collaudo]
      ,[U_DocumentazioneTecnica]
      ,[U_CostoMateriale]
      ,[U_CostoPrimo]
      ,[U_Margine]
      ,[U_ListinoMinimo]
      ,[U_Famiglia]
      ,[U_ErroreWip]
      ,[U_Movimentazioni]
      ,[U_TraferitoODP]
      ,[U_Superlistino]
      ,[U_Superlistino_Vecchio]
      ,[U_Inventariato]
      ,[U_Contato_da]
      ,[U_Country_of_delivery]
      ,[U_Final_customer_Code]
      ,[U_Final_customer_name]
      ,[U_Agent]
      ,[U_Insert_date]
      ,[U_Sector]
      ,[U_Indice_di_revisione]
      ,[U_Tipo_macchina]
      ,[U_Numero_formati]
      ,[U_Gestione_magazzino]
      ,[U_Macchina_standard]
      ,[U_PRG_TIR_Rev]
      ,[U_PRG_TIR_RevProd]
      ,[U_PRG_TIR_RevProdDate]
      ,[U_Tipo_montaggio]
      ,[U_PRG_QLT_HasTC]
      ,[U_Matrice_disegno]
      ,[U_Storico_prezzi]
      ,[U_Disegno]
      ,[U_Spessore_lamiera]
      ,[U_volume]
      ,[U_PRG_TIR_Volume]
      ,[U_PRG_TIR_SpesLamiera]
      ,[U_PRG_TIR_Trattamento]
      ,[U_PRG_TIR_Materiale]
      ,[U_Ubicazione]
      ,[U_Cartella_macchina]
      ,[U_Cartella_linea]
      ,[U_Responsabile_Montaggio]
      ,[U_Responsabile_collaudo]
      ,[U_Famiglia_disegno]
      ,[U_Codice_KTF]
      ,[U_Data_valutazione_stock]
      ,[U_Progetto]
      ,[U_PRG_AZS_MACROGRART]
      ,[U_PRG_AZS_SOTFAM]
      ,[U_PRG_TIR_Explosion]
      ,[U_PRG_TIR_CSost]
      ,[U_PRG_TIR_DispPortale]
      ,[U_Made_in]
      ,[U_Brand]
      ,[U_Codice_BRB]
      ,[U_Contatto_alimentare])


	  select @logistanza,--numero loginstanc
	  [ItemCode]
      ,[Itemname]
      ,[FrgnName]
      ,[ItmsGrpCod]
      ,[CstGrpCode]
      ,[VatGourpSa]
      ,[CodeBars]
      ,[VATLiable]
      ,[PrchseItem]
      ,[SellItem]
      ,[InvntItem]
      ,[OnHand]
      ,[IsCommited]
      ,[OnOrder]
      ,[IncomeAcct]
      ,[ExmptIncom]
      ,[MaxLevel]
      ,[DfltWH]
      ,[CardCode]
      ,[SuppCatNum]
      ,[BuyUnitMsr]
      ,[NumInBuy]
      ,[ReorderQty]
      ,[MinLevel]
      ,[LstEvlPric]
      ,[LstEvlDate]
      ,[CustomPer]
      ,[Canceled]
      ,[MnufctTime]
      ,[WholSlsTax]
      ,[RetilrTax]
      ,[SpcialDisc]
      ,[DscountCod]
      ,[TrackSales]
      ,[SalUnitMsr]
      ,[NumInSale]
      ,[Consig]
      ,[QueryGroup]
      ,[Counted]
      ,[OpenBlnc]
      ,[EvalSystem]
      ,[usersign]
      ,[FREE]
      ,[PicturName]
      ,[Transfered]
      ,[BlncTrnsfr]
      ,[UserText]
      ,[SerialNum]
      ,[CommisPcnt]
      ,[CommisSum]
      ,[CommisGrp]
      ,[TreeType]
      ,[TreeQty]
      ,[LastPurPrc]
      ,[LastPurCur]
      ,[LastPurDat]
      ,[ExitCur]
      ,[ExitPrice]
      ,[ExitWH]
      ,[AssetItem]
      ,[WasCounted]
      ,[ManSerNum]
      ,[SHeight1]
      ,[SHght1Unit]
      ,[SHeight2]
      ,[SHght2Unit]
      ,[SWidth1]
      ,[SWdth1Unit]
      ,[SWidth2]
      ,[SWdth2Unit]
      ,[SLength1]
      ,[SLen1Unit]
      ,[Slength2]
      ,[SLen2Unit]
      ,[SVolume]
      ,[SVolUnit]
      ,[SWeight1]
      ,[SWght1Unit]
      ,[SWeight2]
      ,[SWght2Unit]
      ,[BHeight1]
      ,[BHght1Unit]
      ,[BHeight2]
      ,[BHght2Unit]
      ,[BWidth1]
      ,[BWdth1Unit]
      ,[BWidth2]
      ,[BWdth2Unit]
      ,[BLength1]
      ,[BLen1Unit]
      ,[Blength2]
      ,[BLen2Unit]
      ,[BVolume]
      ,[BVolUnit]
      ,[BWeight1]
      ,[BWght1Unit]
      ,[BWeight2]
      ,[BWght2Unit]
      ,[FixCurrCms]
      ,[FirmCode]
      ,[LstSalDate]
      ,[QryGroup1]
      ,[QryGroup2]
      ,[QryGroup3]
      ,[QryGroup4]
      ,[QryGroup5]
      ,[QryGroup6]
      ,[QryGroup7]
      ,[QryGroup8]
      ,[QryGroup9]
      ,[QryGroup10]
      ,[QryGroup11]
      ,[QryGroup12]
      ,[QryGroup13]
      ,[QryGroup14]
      ,[QryGroup15]
      ,[QryGroup16]
      ,[QryGroup17]
      ,[QryGroup18]
      ,[QryGroup19]
      ,[QryGroup20]
      ,[QryGroup21]
      ,[QryGroup22]
      ,[QryGroup23]
      ,[QryGroup24]
      ,[QryGroup25]
      ,[QryGroup26]
      ,[QryGroup27]
      ,[QryGroup28]
      ,[QryGroup29]
      ,[QryGroup30]
      ,[QryGroup31]
      ,[QryGroup32]
      ,[QryGroup33]
      ,[QryGroup34]
      ,[QryGroup35]
      ,[QryGroup36]
      ,[QryGroup37]
      ,[QryGroup38]
      ,[QryGroup39]
      ,[QryGroup40]
      ,[QryGroup41]
      ,[QryGroup42]
      ,[QryGroup43]
      ,[QryGroup44]
      ,[QryGroup45]
      ,[QryGroup46]
      ,[QryGroup47]
      ,[QryGroup48]
      ,[QryGroup49]
      ,[QryGroup50]
      ,[QryGroup51]
      ,[QryGroup52]
      ,[QryGroup53]
      ,[QryGroup54]
      ,[QryGroup55]
      ,[QryGroup56]
      ,[QryGroup57]
      ,[QryGroup58]
      ,[QryGroup59]
      ,[QryGroup60]
      ,[QryGroup61]
      ,[QryGroup62]
      ,[QryGroup63]
      ,[QryGroup64]
      ,[CreateDate]
      ,getdate()
      ,[ExportCode]
      ,[SalFactor1]
      ,[SalFactor2]
      ,[SalFactor3]
      ,[SalFactor4]
      ,[PurFactor1]
      ,[PurFactor2]
      ,[PurFactor3]
      ,[PurFactor4]
      ,[SalFormula]
      ,[PurFormula]
      ,[VatGroupPu]
      ,[AvgPrice]
      ,[PurPackMsr]
      ,[PurPackUn]
      ,[SalPackMsr]
      ,[SalPackUn]
      ,[SCNCounter]
      ,[ManBtchNum]
      ,[ManOutOnly]
      ,[DataSource]
      ,[validFor]
      ,[validFrom]
      ,[validTo]
      ,[frozenFor]
      ,[frozenFrom]
      ,[frozenTo]
      ,[BlockOut]
      ,[ValidComm]
      ,[FrozenComm]
      ,[ObjType]
      ,[SWW]
      ,[Deleted]
      ,[DocEntry]
      ,[ExpensAcct]
      ,[FrgnInAcct]
      ,[ShipType]
      ,[GLMethod]
      ,[ECInAcct]
      ,[FrgnExpAcc]
      ,[ECExpAcc]
      ,[TaxType]
      ,[ByWh]
      ,[WTLiable]
      ,[ItemType]
      ,[WarrntTmpl]
      ,[BaseUnit]
      ,[CountryOrg]
      ,[StockValue]
      ,[Phantom]
      ,[IssueMthd]
      ,[FREE1]
      ,[PricingPrc]
      ,[MngMethod]
      ,[ReorderPnt]
      ,[InvntryUom]
      ,[PlaningSys]
      ,[PrcrmntMtd]
      ,[OrdrIntrvl]
      ,[OrdrMulti]
      ,[MinOrdrQty]
      ,[LeadTime]
      ,[IndirctTax]
      ,[TaxCodeAR]
      ,[TaxCodeAP]
      ,[OSvcCode]
      ,[ISvcCode]
      ,[ServiceGrp]
      ,[NCMCode]
      ,[MatType]
      ,[MatGrp]
      ,[ProductSrc]
      ,[ServiceCtg]
      ,[ItemClass]
      ,[Excisable]
      ,[ChapterID]
      ,[NotifyASN]
      ,[ProAssNum]
      ,[AssblValue]
      ,[DNFEntry]
      ,@aggiornatore -- mettere utente SAP
      ,[Spec]
      ,[TaxCtg]
      ,[Series]
      ,[Number]
      ,[FuelCode]
      ,[BeverTblC]
      ,[BeverGrpC]
      ,[BeverTM]
      ,[Attachment]
      ,[AtcEntry]
      ,[ToleranDay]
      ,[UgpEntry]
      ,[PUoMEntry]
      ,[SUoMEntry]
      ,[IUoMEntry]
      ,[IssuePriBy]
      ,[AssetClass]
      ,[AssetGroup]
      ,[InventryNo]
      ,[Technician]
      ,[Employee]
      ,[Location]
      ,[StatAsset]
      ,[Cession]
      ,[DeacAftUL]
      ,[AsstStatus]
      ,[CapDate]
      ,[AcqDate]
      ,[RetDate]
      ,[GLPickMeth]
      ,[NoDiscount]
      ,[MgrByQty]
      ,[AssetRmk1]
      ,[AssetRmk2]
      ,[AssetAmnt1]
      ,[AssetAmnt2]
      ,[DeprGroup]
      ,[AssetSerNo]
      ,[CntUnitMsr]
      ,[NumInCnt]
      ,[INUoMEntry]
      ,[OneBOneRec]
      ,[RuleCode]
      ,[ScsCode]
      ,[SpProdType]
      ,[IWeight1]
      ,[IWght1Unit]
      ,[IWeight2]
      ,[IWght2Unit]
      ,[CompoWH]
      ,[CreateTS]
      , concat ( DATEPART(hour, GETDATE()),DATEPART(minute, GETDATE()), DATEPART(second, GETDATE()))
      ,[VirtAstItm]
      ,[SouVirAsst]
      ,[InCostRoll]
      ,[PrdStdCst]
      ,[EnAstSeri]
      ,[LinkRsc]
      ,[OnHldPert]
      ,[onHldLimt]
      ,[PriceUnit]
      ,[GSTRelevnt]
      ,[SACEntry]
      ,[GstTaxCtg]
      ,[AssVal4WTR]
      ,[ExcImpQUoM]
      ,[ExcFixAmnt]
      ,[ExcRate]
      ,[SOIExc]
      ,[TNVED]
      ,[Imported]
      ,[AutoBatch]
      ,[CstmActing]
      ,[StdItemId]
      ,[CommClass]
      ,[TaxCatCode]
      ,[DataVers]
      ,[NVECode]
      ,[CESTCode]
      ,[CtrSealQty]
      ,[LegalText]
      ,[QRCodeSrc]
      ,[Traceable]
      ,[U_UBIMAG]
      ,[U_TEMPOMED]
      ,[U_MODMAC]
      ,[U_SEZIONE]
      ,[U_DESCFR]
      ,[U_DESCING]
      ,[U_UTILIZZ]
      ,[U_ARTCES]
      ,[U_BA_IsFA]
      ,[U_BA_TypID]
      ,[U_BA_NumID]
      ,[U_BA_LVAFrom]
      ,[U_BA_LVA]
      ,[U_ULTVER]
      ,[U_ULTPREZ]
      ,[U_PrzPeso]
      ,[U_PRG_AZS_DesAggArt]
      ,[U_PRG_AZS_DesAgg2Art]
      ,[U_PRG_AZS_CMM]
      ,[U_PRG_AZS_ItmsGrp2Cod]
      ,[U_PRG_AZS_GestComm]
      ,[U_PRG_AZS_CreatedBy]
      ,[U_PRG_AZS_Phantom]
      ,[U_PRG_CLV_Tipo_Lav]
      ,[U_PRG_CLV_Lav_EstAss]
      ,[U_PRG_CLV_Grezzo]
      ,[U_PRG_CLV_TipologiaLav]
      ,[U_PRG_CLV_ArtAssLav]
      ,[U_PRG_CLV_SemiLavComodo]
      ,[U_PRG_CLV_MagPrePro]
      ,[U_Anagrafica]
      ,[U_Ricambio]
      ,[U_Collaudo]
      ,[U_DocumentazioneTecnica]
      ,[U_CostoMateriale]
      ,[U_CostoPrimo]
      ,[U_Margine]
      ,[U_ListinoMinimo]
      ,[U_Famiglia]
      ,[U_ErroreWip]
      ,[U_Movimentazioni]
      ,[U_TraferitoODP]
      ,[U_Superlistino]
      ,[U_Superlistino_Vecchio]
      ,[U_Inventariato]
      ,[U_Contato_da]
      ,[U_Country_of_delivery]
      ,[U_Final_customer_Code]
      ,[U_Final_customer_name]
      ,[U_Agent]
      ,[U_Insert_date]
      ,[U_Sector]
      ,[U_Indice_di_revisione]
      ,[U_Tipo_macchina]
      ,[U_Numero_formati]
      ,[U_Gestione_magazzino]
      ,[U_Macchina_standard]
      ,[U_PRG_TIR_Rev]
      ,[U_PRG_TIR_RevProd]
      ,[U_PRG_TIR_RevProdDate]
      ,[U_Tipo_montaggio]
      ,[U_PRG_QLT_HasTC]
      ,[U_Matrice_disegno]
      ,[U_Storico_prezzi]
      ,[U_Disegno]
      ,[U_Spessore_lamiera]
      ,[U_volume]
      ,[U_PRG_TIR_Volume]
      ,[U_PRG_TIR_SpesLamiera]
      ,[U_PRG_TIR_Trattamento]
      ,[U_PRG_TIR_Materiale]
      ,[U_Ubicazione]
      ,[U_Cartella_macchina]
      ,[U_Cartella_linea]
      ,[U_Responsabile_Montaggio]
      ,[U_Responsabile_collaudo]
      ,[U_Famiglia_disegno]
      ,[U_Codice_KTF]
      ,[U_Data_valutazione_stock]
      ,[U_Progetto]
      ,[U_PRG_AZS_MACROGRART]
      ,[U_PRG_AZS_SOTFAM]
      ,[U_PRG_TIR_Explosion]
      ,[U_PRG_TIR_CSost]
      ,[U_PRG_TIR_DispPortale]
      ,[U_Made_in]
      ,[U_Brand]
      ,[U_Codice_BRB]
      ,[U_Contatto_alimentare]

  FROM [TIRELLISRLDB].[dbo].[OITM]
  where itemcode=@codice_Articolo

  update oitm set loginstanc=@logistanza, usersign2=@aggiornatore ,updatets =concat ( DATEPART(hour, GETDATE()),DATEPART(minute, GETDATE()), DATEPART(second, GETDATE())), updatedate=getdate() where itemcode=@codice_Articolo "
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Sub cambiare_gestione_minimo_ordine(par_utente_sap_aggiornatore As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE T0 SET T0.[MINordrqty]= " & nuovo_valore & " ,  t0.usersign2='" & par_utente_sap_aggiornatore & "',t0.[UpdateDate]=getdate(), t0.updatets=concat ( DATEPART(hour, GETDATE()),DATEPART(minute, GETDATE()), DATEPART(second, GETDATE()))
FROM OITM T0 WHERE T0.[ItemCode] ='" & Codice_SAP & "' "
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Sub cambiare_codice_BRB(par_codice_sap As String, par_valore_stringa As String, utente_sap_aggiornatore As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE T0 SET T0.[u_codice_BRB]= '" & par_valore_stringa & "', t0.usersign2='" & utente_sap_aggiornatore & "', t0.[UpdateDate]=getdate(), t0.updatets=concat ( DATEPART(hour, GETDATE()),DATEPART(minute, GETDATE()), DATEPART(second, GETDATE()))
FROM OITM T0 WHERE T0.[ItemCode] ='" & par_codice_sap & "' "
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Sub cambiare_gestione_ubicazione(par_codice_sap As String, par_valore_stringa As String, utente_sap_aggiornatore As String)

        Dim ubicazione As String

        'If Homepage.Centro_di_costo = "BRB01" Then
        '    ubicazione = "u_ubicazione_labelling"

        'Else

        '    ubicazione = "u_ubicazione"

        'End If


        ubicazione = "u_ubicazione"


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE T0 SET T0.[" & ubicazione & "]= '" & par_valore_stringa & "', t0.usersign2='" & utente_sap_aggiornatore & "', t0.[UpdateDate]=getdate(), t0.updatets=concat ( DATEPART(hour, GETDATE()),DATEPART(minute, GETDATE()), DATEPART(second, GETDATE()))
FROM OITM T0 WHERE T0.[ItemCode] ='" & par_codice_sap & "' "
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Sub cambiare_disegno(par_codice_sap As String, par_valore_stringa As String, utente_sap_aggiornatore As String)

        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE T0 SET T0.[u_disegno]= '" & par_valore_stringa & "', t0.usersign2='" & utente_sap_aggiornatore & "', t0.[UpdateDate]=getdate(), t0.updatets=concat ( DATEPART(hour, GETDATE()),DATEPART(minute, GETDATE()), DATEPART(second, GETDATE()))
FROM OITM T0 WHERE T0.[ItemCode] ='" & par_codice_sap & "' "
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Sub aggiornare_descrizione_desc_sup_osservazioni(par_codice_sap As String, par_descrizione As String, par_descrizione_sup As String, par_osservazioni As String, par_disegno As String, utente_sap_aggiornatore As String, par_produttore As String, par_catalogo As String, par_unita_misura As String, par_codice_fornitore As String, par_gruppo_articoli As String, par_gestione_magazzino As String, par_motivazione_stock As String, par_data_stock As String)

        par_descrizione = Replace(par_descrizione, "'", " ")
        par_descrizione_sup = Replace(par_descrizione_sup, "'", " ")
        par_osservazioni = Replace(par_osservazioni, "'", " ")
        par_disegno = Replace(par_disegno, "'", " ")

        Dim data_stock As String
        If par_data_stock = "01/01/1900" Then
            data_stock = ""
        Else
            data_stock = ", t0.u_data_valutazione_stock= CONVERT(DATE,  '" & par_data_stock & "', 103)"
        End If

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE T0 SET T0.[itemname]= '" & par_descrizione & "'
,t0.frgnname= '" & par_descrizione_sup & "'
,t0.validcomm= '" & par_osservazioni & "'
,t0.u_disegno='" & par_disegno & "'
, t0.firmcode ='" & par_produttore & "'
, t0.suppcatnum='" & par_catalogo & "'
, t0.usersign2='" & utente_sap_aggiornatore & "'
, t0.cardcode ='" & par_codice_fornitore & "'
, t0.[UpdateDate]=getdate()
, t0.invntryuom= '" & par_unita_misura & "'
, t0.itmsgrpcod = '" & par_gruppo_articoli & "'
,t0.u_gestione_magazzino= '" & par_gestione_magazzino & "'
, t0.u_ubimag= '" & par_motivazione_stock & "'
" & data_stock & "
, t0.updatets=concat ( DATEPART(hour, GETDATE()),DATEPART(minute, GETDATE()), DATEPART(second, GETDATE()))
FROM OITM T0 WHERE T0.[ItemCode] ='" & par_codice_sap & "' "
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub



    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = 0 Then

            MsgBox("Non risulta un'utenza sap associata a questo utente")

        Else
            nuovo_valore = InputBox("Inserire nuovo valore minimo ordine")
            cambiare_gestione_minimo_ordine(Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
            update_AITM(Codice_SAP, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
            OttieniDettagliAnagrafica(Codice_SAP)
        End If

    End Sub


    Private Sub Combodipendenti_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles Combodipendenti.SelectedIndexChanged

        'Homepage.UTENTE_sap_SALVATO = Elenco_dipendenti(Combodipendenti.SelectedIndex)

        Homepage.Aggiorna_INI_COMPUTER()



    End Sub


    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        'MsgBox("Funzione non ancora disponibile, utilizzare il gestionale")
        'Return
        If TextBox4.Text = "" Then
            MsgBox("Inserire un numero di Ordine di produzione")
            Return
        End If
        ODP_Form.docnum_odp = TextBox4.Text
        ODP_Form.Show()
        ODP_Form.inizializza_form(TextBox4.Text)




    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        If Homepage.ERP_provenienza = "SAP" Then
            MsgBox("funzione non ancora disponibile, utilizzare il gestionale")
            Return
        End If


        If Button1.Text = "Visualizza distinta base" Then
            Distinta_base_form.Show()

            Distinta_base_form.TextBox1.Text = TextBox2.Text
        ElseIf Button1.Text = "Crea distinta base" Then
            Distinta_base_form.Show()

            Distinta_base_form.TextBox1.Text = TextBox2.Text

        End If
    End Sub




    Private Sub DataGridView_magazzino_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_magazzino.CellFormatting

        Dim par_datagridview As DataGridView
        par_datagridview = DataGridView_magazzino

        If e.RowIndex >= 0 Then

            If par_datagridview.Rows(e.RowIndex).Cells(columnName:="A_MAGA").Value = 0 Then
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="A_MAGA").Style.ForeColor = Color.White
            End If
            If par_datagridview.Rows(e.RowIndex).Cells(columnName:="Da_ass").Value = 0 Then
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="Da_ass").Style.ForeColor = Color.White
            End If
            If par_datagridview.Rows(e.RowIndex).Cells(columnName:="CQ").Value = 0 Then
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="CQ").Style.ForeColor = Color.White
            End If
            If par_datagridview.Rows(e.RowIndex).Cells(columnName:="CONF_").Value = 0 Then
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="CONF_").Style.ForeColor = Color.White
            End If
            If par_datagridview.Rows(e.RowIndex).Cells(columnName:="ORD_").Value = 0 Then
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="ORD_").Style.ForeColor = Color.White
            End If

        End If


        If par_datagridview.Rows(e.RowIndex).Cells(columnName:="MAG").Value = "TOTALE" Then
            par_datagridview.Rows(e.RowIndex).DefaultCellStyle.Font = New Font(par_datagridview.Font, FontStyle.Bold)

            If par_datagridview.Rows(e.RowIndex).Cells(columnName:="DISP").Value < 0 Then

                par_datagridview.Rows(e.RowIndex).Cells(columnName:="DISP").Style.ForeColor = Color.OrangeRed

            End If
        End If
    End Sub



    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click

        Ordine_di_produzione_lista.Show()
        Ordine_di_produzione_lista.inizializzazione_ordine_di_produzione_lista()

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Form_stato_commesse.Show()
    End Sub



    Private Sub Label7_TextChanged(sender As Object, e As EventArgs) Handles Label7.TextChanged

        If Label7.Text = "" Then
            GroupBox19.BackColor = Color.White
        Else
            GroupBox19.BackColor = Color.Gray
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        ' Giacenze.Show()
        Form_Ferretto.Show()

    End Sub

    Private Sub tabpage8_Click(sender As Object, e As EventArgs) Handles TabPage8.Enter

        trova_codici_in_distinta(DataGridView5, TextBox2.Text)
    End Sub

    Private Sub tabpage7_Click(sender As Object, e As EventArgs) Handles TabPage7.Enter

        trova_codice_in_odp(DataGridView3, TextBox2.Text)
    End Sub

    Private Sub tabpage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Enter

        consumi()

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = 0 Then

            MsgBox("Non risulta un'utenza sap associata a questo utente")

        Else
            cambiare_gestione_ferretto(Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
            update_AITM(Codice_SAP, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
            OttieniDettagliAnagrafica(Codice_SAP)
            Label_gestito_a_ferretto.Text = OttieniDettagliAnagrafica(Codice_SAP).Gestito_a_ferretto
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click

        If Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = 0 Then

            MsgBox("Non risulta un'utenza sap associata a questo utente")

        Else
            nuovo_valore = InputBox("Inserire nuovo valore minimo")
            cambiare_gestione_minimo(Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
            update_AITM(Codice_SAP, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
            OttieniDettagliAnagrafica(Codice_SAP)

        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        If Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = 0 Then

            MsgBox("Non risulta un'utenza sap associata a questo utente")

        Else
            nuovo_valore_string = InputBox("Inserire nuova ubicazione")
            cambiare_gestione_ubicazione(Codice_SAP, nuovo_valore_string, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
            update_AITM(Codice_SAP, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
            OttieniDettagliAnagrafica(Codice_SAP)
        End If

    End Sub



    Private Sub Cmd_Entrata_Merce_Click(sender As Object, e As EventArgs)

        Form_Entrate_Merci.inizializzazione_form = True
        If Combodipendenti.Text = "" Or Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = Nothing Then
            MsgBox("Selezionare un utente")
        Else


            Form_Entrate_Merci.Show()

        End If
        Form_Entrate_Merci.inizializzazione_form = False
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)
        Richiesta_trasferimento_materiale.Show()
        ' Richiesta_trasferimento_materiale.riempi_datagridview_rt()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        Form_lotto_di_prelievo.ultimo_lotto()
        Form_lotto_di_prelievo.Show()
        Form_lotto_di_prelievo.inizializzazione_lotto_di_prelievo()

    End Sub

    Private Sub Magazzino_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Inserimento_dipendenti(Combodipendenti)
        'Inserimento_produttore(ComboBox1)
        'UT.inserimento_gruppi(ComboBox2)vai

        'Inserimento_GESTIONE_MAGAZZINO(ComboBox3)
        'Me.BackColor = Homepage.colore_sfondo
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then

            filtro_disegno = ""
        Else
            filtro_disegno = " And t0.u_disegno   Like '%%" & TextBox1.Text & "%%' "
        End If

    End Sub

    Private Sub Button5_Click_2(sender As Object, e As EventArgs) Handles Button5.Click
        Commesse_magazzino.Commesse_odp_aperte()

        Commesse_magazzino.Show()
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs)
        Commesse_magazzino_ODP.Show()
        Commesse_magazzino_ODP.Owner = Me
        Commesse_magazzino_ODP.Commesse_odp_aperte()
        Me.Hide()
    End Sub

    Public Class DettagliAnagrafica
        Public Descrizione As String
        Public codice_brb As String
        Public attivo As String
        Public unita_misura As String
        Public codice_fornitore As String
        Public nome_fornitore As String
        Public Approvvigionamento As String
        Public Cliente As String
        Public gestione_magazzino As String
        Public motivazione_stock As String
        Public data_valutazione As String
        Public Minimo As String
        Public Distinta_base As String
        Public n_mag As Integer
        Public n_cass As Integer
        Public u_progetto As Integer
        Public Property Descrizione_SUP As String
        Public Property Osservazioni As String
        Public Property Disegno As String
        Public Property Ubicazione As String
        Public Property Gestito_a_ferretto As String
        Public Property Gruppo As String
        Public CodiceGruppo As String
        Public Property Prezzo_listino_acquisto As Decimal
        Public Property Label18_Text As String
        Public Property Label19_Text As String
        Public Property Label7_Text As String
        Public Property Label4_Text As String
        Public Property Soggetto_collaudo As String
        Public Property Produttore As String
        Public Property Catalogo As String
        Public Property Test As String

        Public Property minordrqty As String
        Public Property trattamento As String
    End Class

    Public Class scopri_docentry_documento
        Public Docentryodp As Integer
        Public Docentryoc As Integer

    End Class

    Public Class scopri_codice_bp_nome_bp_OITM

        Public codice_BP As String
        Public nome_BP As String

    End Class



    Private Sub Button6_Click_2(sender As Object, e As EventArgs)
        If Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = 0 Then

            MsgBox("Non risulta un'utenza sap associata a questo utente")
        Else
            Dim answer As Integer
            answer = MsgBox("Confermare di aggiornare le informazioni anagrafiche ?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
            If answer = vbYes Then
                aggiornare_descrizione_desc_sup_osservazioni(Codice_SAP, TextBox_descrizione.Text, TextBox3.Text, Label22.Text, TextBox1.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, Elenco_produttori(ComboBox1.SelectedIndex), TextBox8.Text, Label2.Text, codice_fornitore, UT.Elenco_gruppi(ComboBox2.SelectedIndex), Elenco_gestione(ComboBox3.SelectedIndex), RichTextBox1.Text, DateTimePicker1.Value)
                update_AITM(Codice_SAP, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
                OttieniDettagliAnagrafica(Codice_SAP)
                MsgBox("Anagrafica aggiornata con successo")
            End If
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        If Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = 0 Then

            MsgBox("Non risulta un'utenza sap associata a questo utente")
        Else
            Dim answer As Integer
            answer = MsgBox("Confermare di aggiornare le informazioni anagrafiche ?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
            If answer = vbYes Then
                aggiornare_descrizione_desc_sup_osservazioni(Codice_SAP, TextBox_descrizione.Text, TextBox3.Text, Label22.Text, TextBox1.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, Elenco_produttori(ComboBox1.SelectedIndex), TextBox8.Text, Label2.Text, codice_fornitore, UT.Elenco_gruppi(ComboBox2.SelectedIndex), Elenco_gestione(ComboBox3.SelectedIndex), RichTextBox1.Text, DateTimePicker1.Value)
                update_AITM(Codice_SAP, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
                OttieniDettagliAnagrafica(Codice_SAP)
                MsgBox("Anagrafica aggiornata con successo")
            End If
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        If Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = 0 Then

            MsgBox("Non risulta un'utenza sap associata a questo utente")

        Else
            Dim answer As Integer
            answer = MsgBox("Confermare di aggiornare le informazioni anagrafiche ?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
            If answer = vbYes Then
                aggiornare_descrizione_desc_sup_osservazioni(Codice_SAP, TextBox_descrizione.Text, TextBox3.Text, Label22.Text, TextBox1.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, Elenco_produttori(ComboBox1.SelectedIndex), TextBox8.Text, Label2.Text, codice_fornitore, UT.Elenco_gruppi(ComboBox2.SelectedIndex), Elenco_gestione(ComboBox3.SelectedIndex), RichTextBox1.Text, DateTimePicker1.Value)
                update_AITM(Codice_SAP, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
                OttieniDettagliAnagrafica(Codice_SAP)
                MsgBox("Anagrafica aggiornata con successo")
            End If
        End If
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click

        If Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = 0 Then

            MsgBox("Non risulta un'utenza sap associata a questo utente")

        Else
            Dim answer As Integer
            answer = MsgBox("Confermare di aggiornare le informazioni anagrafiche ?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
            If answer = vbYes Then
                Dim produttore As String

                If ComboBox1.SelectedIndex = -1 Then
                    produttore = "-1"
                Else
                    produttore = Elenco_produttori(ComboBox1.SelectedIndex)
                End If

                Dim gruppo_articoli As String

                If ComboBox2.SelectedIndex = -1 Then
                    gruppo_articoli = "183"
                Else
                    gruppo_articoli = UT.Elenco_gruppi(ComboBox2.SelectedIndex)
                End If
                aggiornare_descrizione_desc_sup_osservazioni(Codice_SAP, TextBox_descrizione.Text, TextBox3.Text, Label22.Text, TextBox1.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, produttore, TextBox8.Text, Label2.Text, codice_fornitore, gruppo_articoli, Elenco_gestione(ComboBox3.SelectedIndex), RichTextBox1.Text, DateTimePicker1.Value)

                update_AITM(Codice_SAP, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
                OttieniDettagliAnagrafica(Codice_SAP)
                MsgBox("Anagrafica aggiornata con successo")
            End If
        End If
    End Sub



    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        percorso_Allegato = DataGridView2.Rows(e.RowIndex).Cells(columnName:="path").Value
        filename_allegato = DataGridView2.Rows(e.RowIndex).Cells(columnName:="nome").Value
        estensione_allegato = DataGridView2.Rows(e.RowIndex).Cells(columnName:="estensione").Value
        Button6.BackColor = Color.Lime
    End Sub

    Private Sub Button6_Click_3(sender As Object, e As EventArgs) Handles Button6.Click
        Process.Start(percorso_Allegato & "\" & filename_allegato & "." & estensione_allegato)
    End Sub

    Private Sub DataGridView_trasferito_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_trasferito.CellContentClick

    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = DataGridView3.Columns.IndexOf(docnum) Then




                ODP_Form.docnum_odp = DataGridView3.Rows(e.RowIndex).Cells(columnName:="docnum").Value
                ODP_Form.Show()
                ODP_Form.inizializza_form(DataGridView3.Rows(e.RowIndex).Cells(columnName:="docnum").Value)

            ElseIf e.ColumnIndex = DataGridView3.Columns.IndexOf(Disegno) Then



                visualizza_disegno(DataGridView3.Rows(e.RowIndex).Cells(columnName:="Disegno").Value)



            End If
        End If

    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs)
        Business_partner.Provenienza = "Magazzino"
        Business_partner.Show()




    End Sub

    Private Sub DataGridView5_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellContentClick

    End Sub

    Private Sub DataGridView5_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellClick
        Dim par_datagridview As DataGridView = DataGridView5
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = par_datagridview.Columns.IndexOf(Disegno_padre) Then


                visualizza_disegno(par_datagridview.Rows(e.RowIndex).Cells(columnName:="Disegno_padre").Value)

            ElseIf e.ColumnIndex = par_datagridview.Columns.IndexOf(Codice_padre) Then

                Codice_SAP = par_datagridview.Rows(e.RowIndex).Cells(columnName:="Codice_padre").Value




                TextBox2.Text = Codice_SAP
                OttieniDettagliAnagrafica(Codice_SAP)


            End If
        End If

    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs)
        Form_Richieste_Trasferimento.Show()
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        If TextBox10.Text <> "" Then
            If TextBox10.Text <> "" Then
                If Homepage.ERP_provenienza = "SAP" Then
                    filtro_doc = " AND t20.transname LIKE '%" & TextBox10.Text & "%' "
                Else
                    filtro_doc = " AND UPPER(t0.transname) LIKE UPPER(''%" & TextBox10.Text & "%'') "
                End If
            Else
                filtro_doc = ""
            End If

        Else
            filtro_doc = ""
        End If
        Lista_registrazioni(Codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, filtro_doc, filtro_N_doc, filtro_osservazioni, filtro_mag, filtro_Comm)
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        If TextBox11.Text <> "" Then
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_N_doc = " AND t20.ref1 LIKE '%" & TextBox11.Text & "%' "
            Else
                filtro_N_doc = " AND UPPER(t0.ref1) LIKE UPPER(''%" & TextBox11.Text & "%'') "
            End If
        Else
            filtro_N_doc = ""
        End If
        Lista_registrazioni(Codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, filtro_doc, filtro_N_doc, filtro_osservazioni, filtro_mag, filtro_Comm)
    End Sub



    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        If TextBox12.Text <> "" Then
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_osservazioni = " AND t20.comment LIKE '%" & TextBox12.Text & "%' "
            Else
                filtro_osservazioni = " AND UPPER(t0.comment) LIKE UPPER(''%" & TextBox12.Text & "%'') "
            End If
        Else
            filtro_osservazioni = ""
        End If
        Lista_registrazioni(Codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, filtro_doc, filtro_N_doc, filtro_osservazioni, filtro_mag, filtro_Comm)
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        If TextBox13.Text <> "" Then
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_mag = " AND t20.loccode LIKE '%" & TextBox13.Text & "%' "
            Else
                filtro_mag = " AND UPPER(t0.loccode) LIKE UPPER(''%" & TextBox13.Text & "%'') "
            End If
        Else
            filtro_mag = ""
        End If
        Lista_registrazioni(Codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, filtro_doc, filtro_N_doc, filtro_osservazioni, filtro_mag, filtro_Comm)
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        If TextBox14.Text <> "" Then
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_Comm = " AND t20.U_PRG_AZS_COMMESSA LIKE '%" & TextBox14.Text & "%' "
            Else
                filtro_Comm = " AND UPPER(t0.U_PRG_AZS_COMMESSA) LIKE UPPER(''%" & TextBox14.Text & "%'') "
            End If
        Else
            filtro_Comm = ""
        End If
        Lista_registrazioni(Codice_SAP, DataGridView4, DateTimePicker4, DateTimePicker3, filtro_doc, filtro_N_doc, filtro_osservazioni, filtro_mag, filtro_Comm)
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub DataGridView6_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellContentClick

    End Sub

    Private Sub DataGridView6_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellClick
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = DataGridView6.Columns.IndexOf(DataGridViewButtonColumn2) Then

                Form_nuova_offerta.Show()

                Form_nuova_offerta.TextBox10.Text = DataGridView6.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn2").Value
                Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                Form_nuova_offerta.inizializzazione_form(DataGridView6.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn2").Value, "OPQT", "PQT1", "Richiesta_di_offerta")
            End If
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        provenienza_codice = "BRB01"
        If Len(TextBox6.Text) >= 6 Then
            start_magazzino(TabControl1, TextBox2.Text, TextBox6.Text)
        End If

    End Sub



    Sub scarica_disegni(par_codice_disegno As String)

    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        Dim percorso_Cartella As String
        percorso_Cartella = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TextBox2.Text)
        If Directory.Exists(percorso_Cartella) Then
        Else
            Directory.CreateDirectory(percorso_Cartella)
        End If

        Acquisti.trova_disegni_codice(TextBox1.Text, percorso_Cartella, Homepage.percorso_disegni_generico)
        Beep()
        Process.Start(percorso_Cartella)



    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click

        Dim par_datagridview As DataGridView = DataGridView4
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

    Private Sub BackgroundWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker.DoWork

        'Dim stopWatch As New Stopwatch()
        'Dim ts As TimeSpan = stopWatch.Elapsed

        'Dim percorso_disegni As String = CStr(e.Argument)
        'Dim pdfFile As String = percorso_disegni & "PDF\" & TextBox1.Text & ".PDF"
        '' OpenFileInBackground(percorso_disegni & "PDF\")


        'Stopwatch.Start()



        'Process.Start(percorso_disegni & "PDF\")
        ''Dim esiste = File.GetCreationTime(pdfFile)


        'stopWatch.Stop()
        'ts = stopWatch.Elapsed
        'Console.WriteLine("esiste: " & ts.TotalMilliseconds & " ms")
        'stopWatch.Restart()


        Dim percorso_disegni As String = CStr(e.Argument)
        Dim pdfFile As String = percorso_disegni & "PDF\" & TextBox1.Text & ".PDF"

        'If File.Exists(pdfFile) Then
        '    Me.Invoke(Sub()

        '                  AxFoxitCtl1.OpenFile(pdfFile)
        '                  AxFoxitCtl1.Show()
        '                  TextBox1.BackColor = Color.Lime
        '              End Sub)
        'Else
        '    Me.Invoke(Sub()
        '                  AxFoxitCtl1.Hide()
        '                  TextBox1.BackColor = Color.Red
        '              End Sub)
        'End If
    End Sub

    Public Sub OpenFileInBackground(pdfFile As String)
        Dim startInfo As New ProcessStartInfo()
        startInfo.FileName = pdfFile
        startInfo.UseShellExecute = True
        startInfo.WindowStyle = ProcessWindowStyle.Hidden
        Process.Start(startInfo)
    End Sub



    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        If Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = 0 Then

            MsgBox("Non risulta un'utenza sap associata a questo utente")

        Else
            nuovo_valore_string = InputBox("Inserire nuovo codice BRB")
            cambiare_codice_BRB(Codice_SAP, nuovo_valore_string, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvatomepage.UTENTE_sap_SALVATO)

            update_AITM(Codice_SAP, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
            OttieniDettagliAnagrafica(Codice_SAP)
        End If
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then



            If e.ColumnIndex = DataGridView4.Columns.IndexOf(N_doc) Then
                If DataGridView4.Rows(e.RowIndex).Cells(columnName:="tipo_doc").Value = "EM Forn" Then
                    Form_Entrate_Merci.inizializzazione_form = True


                    Form_Entrate_Merci.Show()
                    Form_Entrate_Merci.BringToFront()
                    Form_Entrate_Merci.Txt_DocNum.Text = DataGridView4.Rows(e.RowIndex).Cells(columnName:="N_doc").Value
                    Form_Entrate_Merci.TextBox2.Text = TextBox2.Text
                    Form_Entrate_Merci.Aggiorna_EM(Form_Entrate_Merci.Txt_DocNum.Text)
                    Form_Entrate_Merci.inizializzazione_form = False
                End If

            End If

        End If
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click

        If DataGridView4.CurrentRow IsNot Nothing Then
            Dim tipo_trasf As String = ""

            If DataGridView4.CurrentRow.Cells("Tipo_doc").Value = "Trasf MAG" Then


                Dim nDoc As Integer = Convert.ToInt32(DataGridView4.CurrentRow.Cells("N_doc").Value)
                Dim itemCode As String = TextBox2.Text.ToUpper() ' Sostituisci "ItemCode" con il nome corretto della colonna

                Trasferimento_magazzino.stampa_scontrino_da_trasf(nDoc, itemCode)
            Else
                MsgBox("Funzione valida solo per trasferimenti di magazzino")
            End If
        Else
            MessageBox.Show("Seleziona una riga valida.")
        End If
    End Sub

    Private Sub DataGridView_ordinato_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ordinato.CellContentClick

    End Sub

    Private Sub Button8_Click_2(sender As Object, e As EventArgs) Handles Button8.Click
        Form_stampe.Show()
    End Sub

    Private Sub GroupBox18_Enter(sender As Object, e As EventArgs) Handles GroupBox18.Enter


    End Sub

    Private Sub TableLayoutPanel8_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel8.Paint

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        Homepage.ERP_provenienza = "SAP"
        Homepage.sap_tirelli = "Data Source=srvtirsap01.corp.arol-group.com;Initial Catalog=TIRELLISRLDB;Persist Security Info=True;User ID=sa;Password=123B1Admin"

        If Homepage.ERP_provenienza = "SAP" Then


            Homepage.colore_sfondo = Color.PowderBlue
        Else
            Homepage.colore_sfondo = Color.Aquamarine
        End If
        Me.BackColor = Homepage.colore_sfondo

        'start_magazzino(TabControl1, TextBox2.Text.ToUpper(), "")
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        Homepage.ERP_provenienza = "GALILEO"
        Homepage.sap_tirelli = "Data Source=srvtirsap01.corp.arol-group.com;Initial Catalog=TIRELLI_40;Persist Security Info=True;User ID=sa;Password=123B1Admin"
        If Homepage.ERP_provenienza = "SAP" Then


            Homepage.colore_sfondo = Color.PowderBlue
        Else
            Homepage.colore_sfondo = Color.Aquamarine
        End If
        Me.BackColor = Homepage.colore_sfondo

        start_magazzino(TabControl1, TextBox2.Text.ToUpper(), "")
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Form_visualizza_picture.Show()
        visualizza_picture(TextBox1.Text, Form_visualizza_picture.PictureBox1)
    End Sub

    Private Sub DataGridView4_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub DataGridView4_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) _
        'Handles DataGridView4.CellPainting

        '    If e.RowIndex < 0 OrElse e.ColumnIndex < 0 Then Exit Sub

        '    Dim dgv = DataGridView4
        '    Dim colRifmag = dgv.Columns("num_ope").Index

        '    Dim currentValue = dgv.Rows(e.RowIndex).Cells(colRifmag).Value?.ToString()

        '    Dim prevValue As String = Nothing
        '    Dim nextValue As String = Nothing

        '    If e.RowIndex > 0 Then
        '        prevValue = dgv.Rows(e.RowIndex - 1).Cells(colRifmag).Value?.ToString()
        '    End If

        '    If e.RowIndex < dgv.Rows.Count - 1 Then
        '        nextValue = dgv.Rows(e.RowIndex + 1).Cells(colRifmag).Value?.ToString()
        '    End If

        '    e.Paint(e.CellBounds, DataGridViewPaintParts.All)

        '    Using pen As New Pen(Color.Black, 2)

        '        ' Bordo superiore (inizio gruppo)
        '        If currentValue <> prevValue Then
        '            e.Graphics.DrawLine(
        '            pen,
        '            e.CellBounds.Left,
        '            e.CellBounds.Top,
        '            e.CellBounds.Right,
        '            e.CellBounds.Top
        '        )
        '        End If

        '        ' Bordo inferiore (fine gruppo)
        '        If currentValue <> nextValue Then
        '            e.Graphics.DrawLine(
        '            pen,
        '            e.CellBounds.Left,
        '            e.CellBounds.Bottom - 1,
        '            e.CellBounds.Right,
        '            e.CellBounds.Bottom - 1
        '        )
        '        End If
        '    End Using

        '    e.Handled = True
    End Sub

    Private Sub DataGridView_magazzino_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_magazzino.CellContentClick

    End Sub

    Private Sub DataGridView4_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView4.CellFormatting
        Dim par_datagridview As DataGridView
        par_datagridview = DataGridView4

        If e.RowIndex >= 0 Then

            If par_datagridview.Rows(e.RowIndex).Cells(columnName:="MOV").Value < 0 Then
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="MOV").Style.ForeColor = Color.Red
            End If


        End If
    End Sub

    Private Sub Button7_Click_2(sender As Object, e As EventArgs) Handles Button7.Click
        Form_Movimenti_magazzino.Show()
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        Entrate_merci.Show()
    End Sub

    Private Sub TextBox_descrizione_TextChanged(sender As Object, e As EventArgs) Handles TextBox_descrizione.TextChanged

    End Sub
End Class

