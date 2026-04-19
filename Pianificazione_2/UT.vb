Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports System.Windows.Documents
Imports System.Diagnostics
Imports System.Net.Http
Imports System.Text
Imports Newtonsoft.Json.Linq


Public Class UT
    Public Elenco_gruppi(10000) As String

    Private filtro_fam_disegno As String
    Private _picRicercaImg As New PictureBox()
    Private _btnCercaImmagine As New Button()
    Private _lblPastaHint As New Label()
    Private visualizzazione As String = "TIRELLI"
    Private filtro_descrizione As String
    Public filtro_descrizione_supp As String
    Private filtro_disegno As String
    Private filtro_catalogo As String

    ' Cronologia conversazione AI
    Private _aiHistory As New JArray()

    Private Const AI_SYSTEM_PROMPT As String =
        "Sei un assistente per la ricerca articoli nel gestionale Tirelli (AS400/SAP)." & vbLf & vbLf &
        "SCHEMA AS400 (S786FAD1.TIR90VIS):" & vbLf &
        "JGALART: CODE=codice, DES_CODE=descrizione, DISEGNO=disegno," & vbLf &
        "  GRUP_ART=codice-gruppo, DESC_GRP=descrizione-gruppo, PROD_FOR=produttore, CODAR_FOR=codice-catalogo-fornitore," & vbLf &
        "  DESC_FOR=fornitore-preferito, UBI_CODE=ubicazione, COSTO_STD=costo," & vbLf &
        "  STAT_CODE=stato(A/I), UMIS=unità-misura, TIPO_PARTE=approvvigionamento(A/P)" & vbLf &
        "JGALMAG: CODART, MAG, QTA_MAG=giacenza, QTA_DISP=disponibile" & vbLf &
        "JGALord: CODART, DISEGNO, NUMDOC=n.ordine, QTA_ORD, DATA_RICHIESTA, EVASO(S/N), COD_FORN" & vbLf & vbLf &
        "REGOLA OBBLIGATORIA: rispondi SEMPRE con JSON puro (niente testo prima/dopo, niente ```json)." & vbLf & vbLf &
        "Se la richiesta è una RICERCA di articoli, usa:" & vbLf &
        "{""azione"":""cerca"",""messaggio"":""spiegazione breve"",""filtri"":{" & vbLf &
        "  ""codice"":"""",""descrizione"":"""",""desc_supp"":"""",""disegno"":""""," & vbLf &
        "  ""catalogo"":"""",""produttore"":"""",""fornitore"":"""",""ubicazione"":"""",""gruppo"":""""}}" & vbLf & vbLf &
        "Il campo 'gruppo' deve contenere la DESCRIZIONE del gruppo articoli (DESC_GRP), non il codice numerico." & vbLf &
        "Esempi di descrizioni gruppo: 'TAGLIO LASER', 'CUSCINETTI', 'VITI E BULLONI', ecc." & vbLf & vbLf &
        "Se la richiesta è INFORMATIVA (definizioni, spiegazioni, dati da DB non cercabili tramite filtri), usa:" & vbLf &
        "{""azione"":""info"",""messaggio"":""risposta completa""}" & vbLf & vbLf &
        "SINTASSI VALORI FILTRO (campi filtri):" & vbLf &
        "  parola     = contiene 'parola'" & vbLf &
        "  *parola    = inizia con 'parola'" & vbLf &
        "  parola*    = finisce con 'parola'" & vbLf &
        "  p1*p2      = contiene p1 E p2" & vbLf &
        "  """"        = nessun filtro su quel campo" & vbLf & vbLf &
        "ESEMPI:" & vbLf &
        "Utente: 'cerca cuscinetti SKF 6205' → {""azione"":""cerca"",""messaggio"":""Cerco cuscinetti con catalogo 6205 e produttore SKF"",""filtri"":{""codice"":"""",""descrizione"":""cuscinett*"",""desc_supp"":"""",""disegno"":"""",""catalogo"":""6205"",""produttore"":""SKF"",""fornitore"":"""",""ubicazione"":"""",""gruppo"":""""}}" & vbLf &
        "Utente: 'codici taglio laser di Galati' → {""azione"":""cerca"",""messaggio"":""Cerco articoli taglio laser con fornitore Galati"",""filtri"":{""codice"":"""",""descrizione"":"""",""desc_supp"":"""",""disegno"":"""",""catalogo"":"""",""produttore"":"""",""fornitore"":""galati"",""ubicazione"":"""",""gruppo"":""taglio laser""}}" & vbLf &
        "Utente: 'cosa significa TIPO_PARTE=A?' → {""azione"":""info"",""messaggio"":""TIPO_PARTE=A significa acquisto esterno (il componente viene acquistato, non prodotto internamente).""}"

    Sub inizializzazione_ut()
        inserimento_gruppi(ComboBox1)
        inserimento_gruppi(ComboBox9)
        inserimento_produttore()
        inserimento_FORNITORI()
        AggiungiColonnaImmagine()
    End Sub

    Private Sub AggiungiColonnaImmagine()
        If DataGridView_SAP.Columns.Contains("Immagine") Then Return
        Dim imgCol As New DataGridViewImageColumn()
        imgCol.Name = "Immagine"
        imgCol.HeaderText = ""
        imgCol.Width = 72
        imgCol.FillWeight = 25
        imgCol.Resizable = DataGridViewTriState.False
        imgCol.ImageLayout = DataGridViewImageCellLayout.Zoom
        imgCol.DefaultCellStyle.NullValue = Nothing
        DataGridView_SAP.Columns.Add(imgCol)
        imgCol.DisplayIndex = 1
        DataGridView_SAP.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        DataGridView_SAP.RowTemplate.Height = 70
    End Sub

    Sub cerca(Par_codice As String, par_descrizione As String, par_descrizone_supp As String, par_disegno As String, par_fam_disegno As String, par_gruppo_art As String, par_produttore As String, par_cat_forn As String, par_fornitore As String, par_ubicazione As String)
        Dim produttore As String
        Dim fornitore_preferito As String
        Dim ubic As String

        If ComboBox4.Text = "" Then
            produttore = ""
        Else
            produttore = " and t2.firmname like '%%" & ComboBox4.Text & "%%'"
        End If

        If ComboBox6.SelectedIndex < 0 Then
            fornitore_preferito = ""
        Else
            fornitore_preferito = " and t3.cardname= '" & ComboBox6.Text & "'"
        End If

        If TextBox4.Text = "" Then
            ubic = ""
        Else
            If visualizzazione = "TIRELLI" Then
                ubic = " And T0.[u_ubicazione] Like '%%" & TextBox4.Text & "%%' "
            Else
                ubic = " and T0.[ubicazione] like '%%" & TextBox4.Text & "%%' "
            End If
        End If

        If TextBox_disegno_SAP.Text = Nothing Then
            TextBox_disegno_SAP.Text = ""
        End If

        DataGridView_SAP.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim Cmd_SAP As New SqlCommand
        Dim Cmd_SAP_Reader As SqlDataReader

        Cmd_SAP.Connection = Cnn
        If Homepage.ERP_provenienza = "sap" Then

            Cmd_SAP.CommandText = "select top " & TextBox9.Text & " t10.codice,t10.nome,t10.[Nome supp],t10.disegno,t10.[Gruppo art],t10.FirmName, t10.SuppCatNum,t10.CardName,t10.u_ubicazione, t10.u_famiglia_disegno,
t10.onhand,t10.WIP,sum(coalesce(t11.onhand,0)+coalesce(t11.onorder,0)-coalesce(t11.iscommited,0)) as 'Disp',t10.price,t10.frozenfor
from
(
SELECT top 100 T0.[ItemCode] as 'Codice', coalesce(T0.[ItemName],'') as 'Nome', coalesce(T0.[FrgnName],'') as 'Nome supp',t0.u_disegno as 'Disegno', T1.[ItmsGrpNam] as 'Gruppo art' , t2.firmname, t0.suppcatnum,coalesce(t3.cardname,'') as 'Cardname', t0.u_ubicazione, t0.u_famiglia_disegno,
t0.onhand-t5.onhand-t6.onhand as 'onhand',t5.onhand+t6.onhand as 'WIP',t4.price,t0.frozenfor
        FROM OITM T0  INNER JOIN OITB T1 ON T0.[ItmsGrpCod] = T1.[ItmsGrpCod]
left join omrc t2 on t2.firmcode=t0.firmcode
left join ocrd t3 on t3.cardcode=t0.cardcode
inner join itm1 t4 on t4.itemcode=t0.itemcode and t4.pricelist=2
left join oitw t5 on t5.itemcode =t0.itemcode and t5.whscode='WIP'
left join oitw t6 on t6.itemcode =t0.itemcode and t6.whscode='BWIP'
        WHERE t0.itemcode Like '%%" & TextBox_codice_SAP_RICERCA.Text & "%%' " & filtro_descrizione & fornitore_preferito & filtro_catalogo & filtro_disegno & filtro_descrizione_supp & "  and t0.itmsgrpcod like '%%" & Elenco_gruppi(ComboBox1.SelectedIndex) & "%%'" & produttore & ubic & filtro_fam_disegno &
             ")
		as t10 left join oitw t11 on t10.codice=t11.itemcode

		group by t10.codice,t10.nome,t10.[Nome supp],t10.disegno,t10.[Gruppo art],t10.FirmName, t10.SuppCatNum,t10.CardName,t10.u_ubicazione, t10.u_famiglia_disegno,
t10.onhand,t10.WIP,t10.price,t10.frozenfor"
        Else
            ' Costruisci WHERE AS400 con supporto asterischi per ogni campo
            Dim where400 As String = "where 1=1"
            where400 &= CostruisciCondizioneAS400("code", Par_codice)
            where400 &= CostruisciCondizioneAS400("t0.des_code", par_descrizione)
            where400 &= CostruisciCondizioneAS400("t0.disegno", par_disegno)
            where400 &= CostruisciCondizioneAS400("t0.grup_art", par_gruppo_art)
            where400 &= CostruisciCondizioneAS400("t0.prod_for", par_produttore)
            where400 &= CostruisciCondizioneAS400("t0.codar_for", par_cat_forn)
            where400 &= CostruisciCondizioneAS400("t0.desc_for", par_fornitore)
            where400 &= CostruisciCondizioneAS400("t0.ubi_code", par_ubicazione)

            Cmd_SAP.CommandText = "select
    trim(t10.code) as 'Codice'
    ,t10.des_code as 'Nome'
    ,'MANCA' AS 'Nome supp'
    ,trim(t10.disegno) as 'Disegno'
    ,trim(t10.desc_grp) as 'Gruppo art'
    ,t10.prod_for as 'firmname'
    ,t10.codar_for as 'Suppcatnum'
    ,t10.desc_for as 'Cardname'
    ,t10.ubi_cOde as 'u_ubicazione'
    ,'manca' as 'u_famiglia_disegno'
    ,T10.onhand as 'onhand'
    ,T10.wip_tot as 'WIP'
    ,T10.disp_tot as 'DISP'
    ,T10.costo_std as 'PRICE'
    ,trim(t10.stat_code) as 'frozenfor'
from openquery(AS400,'
    select
        T0.CODE,
        T0.DES_CODE,
        T0.DISEGNO,
        T0.GRUP_ART,
        T0.DESC_GRP,
        T0.STAT_CODE,
        T0.PROD_FOR,
        T0.CODAR_FOR,
        T0.DESC_FOR,
        T0.UBI_CODE,
        T0.COSTO_STD,
        coalesce(T1.TOT_QTA, 0) as onhand,
        coalesce(T1.TOT_DISP, 0) as disp_tot,
        coalesce(T2.TOT_WIP, 0) as wip_tot
    from S786FAD1.TIR90VIS.JGALART T0
    left join (
        select
            CODART,
            sum(QTA_MAG) as TOT_QTA,
            sum(QTA_DISP) as TOT_DISP
        from TIR90VIS.JGALMAG
        group by CODART
    ) T1 ON T0.CODE = T1.CODART
    left join (
        select
            CODART,
            sum(QTATRA) as TOT_WIP
        from TIR90VIS.JGALIMP
        where EVASO_ODP <> ''S''
        group by CODART
    ) T2 ON T0.CODE = T2.CODART
" & where400 & "
'
)
as t10"

        End If

        Cmd_SAP_Reader = Cmd_SAP.ExecuteReader

        Do While Cmd_SAP_Reader.Read()
            Dim disVal = If(IsDBNull(Cmd_SAP_Reader("Disegno")), "", Cmd_SAP_Reader("Disegno").ToString().Trim())
            Dim rowIdx = DataGridView_SAP.Rows.Add(Cmd_SAP_Reader("Codice"), Cmd_SAP_Reader("Nome"), Cmd_SAP_Reader("Nome supp"), Cmd_SAP_Reader("u_famiglia_disegno"), Cmd_SAP_Reader("Disegno"), Cmd_SAP_Reader("Gruppo art"), Cmd_SAP_Reader("firmname"), Cmd_SAP_Reader("suppcatnum"), Cmd_SAP_Reader("cardname"), Cmd_SAP_Reader("u_ubicazione"), Cmd_SAP_Reader("onhand"), Cmd_SAP_Reader("wip"), Cmd_SAP_Reader("disp"), Cmd_SAP_Reader("price"), Cmd_SAP_Reader("frozenfor"))
            CaricaImmagineRiga(rowIdx, disVal)
        Loop

        Cmd_SAP_Reader.Close()
        Cnn.Close()
    End Sub

    Sub cerca_BRB()
        Dim produttore As String
        Dim catalogo_fornitore As String
        Dim ubic As String

        If ComboBox4.Text = "" Then
            produttore = ""
        Else
            produttore = " and (t10.firmname like '%%" & ComboBox4.Text & "%%' or t10.firmname like '%%" & ComboBox4.Text & "%%') "
        End If

        If TextBox2.Text = "" Then
            catalogo_fornitore = ""
        Else
            catalogo_fornitore = " And (t10.SUPPCATNUM Like '%%" & TextBox2.Text & "%%' or t10.catalogo_fornitore like '%%" & TextBox2.Text & "%%') "
        End If

        If TextBox4.Text = "" Then
            ubic = ""
        Else
            ubic = " and (T10.[u_ubicazione_labelling] like '%%" & TextBox4.Text & "%%' or T10.[ubicazione] like '%%" & TextBox4.Text & "%%' )"
        End If

        DataGridView1.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim Cmd_SAP As New SqlCommand
        Dim Cmd_SAP_Reader As SqlDataReader

        Cmd_SAP.Connection = Cnn
        If TextBox_disegno_SAP.Text = Nothing Then
            Cmd_SAP.CommandText = "		select *
from
(

select
coalesce(t1.[ID],0) as 'ID'
      ,coalesce(t1.[Codice_BRB],t0.u_codice_brb) as 'Codice_BRB'
      ,coalesce(t1.[Descrizione_BRB],t0.itemname) as 'Descrizione_BRB'
      ,coalesce(t1.[Descrizione_supp_BRB],t0.frgnname) as 'Descrizione_supp_BRB'
      ,coalesce(t1.[Catalogo_fornitore],t0.suppcatnum) as 'Catalogo_fornitore'
      ,coalesce(t1.[Fornitore],t4.cardname) as 'Fornitore'
      ,cast(coalesce(t1.[Costo],coalesce(t5.price,0)) as integer) as 'Costo'
,coalesce(t1.ubicazione,COALESCE(t0.u_ubicazione_labelling,'')) as 'Ubicazione'
,coalesce(t0.itemcode,'') as 'Itemcode'
,coalesce(t2.tirelli,'') as 'Tirelli'
, coalesce(t2.Gruppo_articoli,t0.[ItmsGrpCod]) as 'Gruppo_articoli'
,coalesce(t3.ItmsGrpNam,t6.ItmsGrpNam) as 'ItmsGrpNam'
, coalesce(t0.itemname,'') as 'Desc_tirelli'
,coalesce(t0.frgnname,'') as 'Desc_supp_tirelli'
,t0.u_codice_brb

  FROM [TIRELLISRLDB].[dbo].[OITM] t0
   left join [TIRELLI_40].[DBO].BRB_Codici t1 on t0.u_codice_BRB=t1.Codice_BRB
  left join [TIRELLI_40].[DBO].[Frontiera_PDM_BRB_SAP_Prima_Lettera] t2 on t2.brb=substring(t1.codice_BRB,1,1)
  left join [TIRELLISRLDB].[dbo].[oitb] t3 on t3.ItmsGrpCod=t2.Gruppo_articoli
  left join ocrd t4 on t4.cardcode=t0.cardcode
  left join itm1 t5 on t5.itemcode=t0.itemcode and t5.pricelist=2
left join [TIRELLISRLDB].[dbo].[oitb] t6 on t6.ItmsGrpCod=t0.[ItmsGrpCod]
left join omrc t7 on t7.firmcode=t0.firmcode

union all

sELECT  coalesce(t1.[ID],0) as 'ID'
      ,coalesce(t1.[Codice_BRB],t0.u_codice_brb) as 'Codice_BRB'
      ,coalesce(t1.[Descrizione_BRB],t0.itemname) as 'Descrizione_BRB'
      ,coalesce(t1.[Descrizione_supp_BRB],t0.frgnname) as 'Descrizione_supp_BRB'
      ,coalesce(t1.[Catalogo_fornitore],t0.suppcatnum) as 'Catalogo_fornitore'
      ,coalesce(t1.[Fornitore],t4.cardname) as 'Fornitore'
      ,cast(coalesce(t1.[Costo],coalesce(t5.price,0)) as integer) as 'Costo'
,coalesce(t1.ubicazione,t0.u_ubicazione_labelling) as 'Ubicazione'
,coalesce(t0.itemcode,'') as 'Itemcode'
,coalesce(t2.tirelli,'') as 'Tirelli'
, coalesce(t2.Gruppo_articoli,t0.[ItmsGrpCod]) as 'Gruppo_articoli'
,coalesce(t3.ItmsGrpNam,t6.ItmsGrpNam) as 'ItmsGrpNam'
, coalesce(t0.itemname,'') as 'Desc_tirelli'
,coalesce(t0.frgnname,'') as 'Desc_supp_tirelli'
,t0.u_codice_brb


 from [TIRELLI_40].[DBO].BRB_Codici t1
left join [TIRELLISRLDB].[dbo].[OITM] t0 on t0.u_codice_BRB=t1.Codice_BRB
  left join [TIRELLI_40].[DBO].[Frontiera_PDM_BRB_SAP_Prima_Lettera] t2 on t2.brb=substring(t1.codice_BRB,1,1)
  left join [TIRELLISRLDB].[dbo].[oitb] t3 on t3.ItmsGrpCod=t2.Gruppo_articoli
  left join ocrd t4 on t4.cardcode=t0.cardcode
  left join itm1 t5 on t5.itemcode=t0.itemcode and t5.pricelist=2
left join [TIRELLISRLDB].[dbo].[oitb] t6 on t6.ItmsGrpCod=t0.[ItmsGrpCod]
left join omrc t7 on t7.firmcode=t0.firmcode

        WHERE t0.itemcode  is null

)
as t10



        WHERE (t10.Codice_BRB Like '%%" & TextBox_codice_SAP_RICERCA.Text & "%%' or t10.u_codice_brb Like '%%" & TextBox_codice_SAP_RICERCA.Text & "%%')   " & filtro_descrizione_supp & filtro_descrizione & catalogo_fornitore & ubic

        End If

        Cmd_SAP_Reader = Cmd_SAP.ExecuteReader

        Do While Cmd_SAP_Reader.Read()
            DataGridView1.Rows.Add(Cmd_SAP_Reader("Codice_BRB"), Cmd_SAP_Reader("itemcode"), Cmd_SAP_Reader("Descrizione_BRB"), Cmd_SAP_Reader("descrizione_supp_Brb"), Cmd_SAP_Reader("Fornitore"), Cmd_SAP_Reader("Catalogo_fornitore"), Cmd_SAP_Reader("Ubicazione"), Cmd_SAP_Reader("Costo"), Cmd_SAP_Reader("Tirelli"), Cmd_SAP_Reader("Gruppo_articoli"), Cmd_SAP_Reader("ItmsGrpNam"))
        Loop

        Cmd_SAP_Reader.Close()
        Cnn.Close()
    End Sub

    Sub inserimento_gruppi(par_combobox As ComboBox)

        par_combobox.Items.Clear()
        par_combobox.Items.Add("")
        par_combobox.SelectedIndex = 0

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP.CommandText = "SELECT T0.[ItmsGrpCod] AS 'Gruppo', T0.[ItmsGrpNam] as 'Nome gruppo' FROM OITB T0"
        Else
            CMD_SAP.CommandText = "SELECT trim(t0.grup_art) as Gruppo, trim(t0.desc_grp) as 'Nome gruppo'
FROM OPENQUERY(AS400, '
    SELECT DISTINCT grup_art, desc_grp
    FROM S786FAD1.TIR90VIS.JGALART
    WHERE stat_code <> ''I'' AND grup_art <> '' ''
') as t0
ORDER BY t0.grup_art"
        End If

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer = 1
        Do While cmd_SAP_reader.Read()
            Elenco_gruppi(Indice) = cmd_SAP_reader("Gruppo")
            par_combobox.Items.Add(cmd_SAP_reader("Nome gruppo"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub inserimento_produttore()
        ComboBox4.Items.Clear()
        ComboBox4.Items.Add("")
        ComboBox4.SelectedIndex = 0

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn

        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP.CommandText = "Select t0.firmname as nome from tirellisrldb.dbo.omrc t0 ORDER BY T0.FIRMNAME"
        Else
            CMD_SAP.CommandText = "SELECT trim(t0.prod_for) as nome
FROM OPENQUERY(AS400, '
    SELECT DISTINCT prod_for
    FROM S786FAD1.TIR90VIS.JGALART
    WHERE stat_code <> ''I'' AND prod_for <> '' ''
') as t0
ORDER BY t0.prod_for"
        End If

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Do While cmd_SAP_reader.Read()
            Dim val = cmd_SAP_reader("nome").ToString().Trim()
            If val <> "" Then ComboBox4.Items.Add(val)
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub inserimento_FORNITORI()
        ComboBox6.Items.Clear()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD As New SqlCommand
        Dim rdr As SqlDataReader

        CMD.Connection = Cnn

        If Homepage.ERP_provenienza = "SAP" Then
            CMD.CommandText = "Select t0.CARDCODE, t0.CARDname from OCRD t0 WHERE T0.CARDTYPE='S' and t0.validfor='Y' ORDER BY T0.CARDNAME"
            rdr = CMD.ExecuteReader
            Do While rdr.Read()
                ComboBox6.Items.Add(rdr("CARDname"))
            Loop
        Else
            CMD.CommandText = "SELECT trim(t0.conto) as conto, trim(t0.ds_Conto) as ds_Conto
FROM OPENQUERY(AS400, '
    SELECT conto, ds_Conto
    FROM S786FAD1.TIR90VIS.JGALACF
    WHERE clifor=''F'' and stato<>''S''
') as t0
ORDER BY t0.ds_Conto"
            rdr = CMD.ExecuteReader
            Do While rdr.Read()
                Dim nome = rdr("ds_Conto").ToString().Trim()
                If nome <> "" Then ComboBox6.Items.Add(nome)
            Loop
        End If

        rdr.Close()
        Cnn.Close()
    End Sub

    Private Sub UT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.BackColor = Homepage.colore_sfondo
        inizializzazione_ut()
        ApplicaStileNavy()
    End Sub

    Private Sub ApplicaStileNavy()
        Dim navy As Color = Color.FromArgb(22, 45, 84)
        Dim navyDark As Color = Color.FromArgb(10, 26, 55)
        Dim navyHover As Color = Color.FromArgb(30, 63, 122)
        Dim bgApp As Color = Color.FromArgb(238, 242, 247)

        ' Barre filtri — sfondo medio, label navy leggibili su bianco
        Dim bgBar As Color = Color.FromArgb(224, 232, 245)    ' azzurro-grigio chiaro
        Panel3.BackColor = bgBar
        Panel26.BackColor = bgBar

        ' Pulsanti X e − — navy scuro su sfondo barra
        For Each b As Button In New Button() {Button3, Button20}
            b.ForeColor = navy
            b.BackColor = bgBar
            b.FlatStyle = FlatStyle.Flat
            b.FlatAppearance.BorderColor = Color.FromArgb(180, 200, 230)
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(200, 215, 240)
        Next

        ' Pulsante Legenda — nella barra filtri
        BtnLegenda.BackColor = navy
        BtnLegenda.ForeColor = Color.White
        BtnLegenda.FlatStyle = FlatStyle.Flat
        BtnLegenda.FlatAppearance.BorderColor = navyDark
        BtnLegenda.FlatAppearance.MouseOverBackColor = navyHover

        ' Pulsante Cerca
        Button_CERCA.BackColor = navy
        Button_CERCA.ForeColor = Color.White
        Button_CERCA.FlatStyle = FlatStyle.Flat
        Button_CERCA.FlatAppearance.BorderColor = navyDark
        Button_CERCA.FlatAppearance.MouseOverBackColor = navyHover

        ' GroupBox label — navy su sfondo chiaro = leggibile
        For Each gb As GroupBox In New GroupBox() {
            GroupBox2, GroupBox3, GroupBox10, GroupBox4, GroupBox22,
            GroupBox6, GroupBox14, GroupBox12, GroupBox13, GroupBox15, GroupBox27}
            gb.ForeColor = navy
            gb.BackColor = Color.White
        Next

        ' DataGridView header navy
        DataGridView_SAP.ColumnHeadersDefaultCellStyle.BackColor = navy
        DataGridView_SAP.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        DataGridView_SAP.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(238, 242, 247)
        DataGridView_SAP.BackgroundColor = Color.FromArgb(250, 252, 255)
        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = navy
        DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        DataGridView1.BackgroundColor = Color.FromArgb(250, 252, 255)

        ' Pannello AI — header più prominente con blu vivace
        Dim accentBlue As Color = Color.FromArgb(0, 120, 212)
        Dim accentDark As Color = Color.FromArgb(0, 90, 160)
        PanelAI.BackColor = bgApp
        PanelAI.BorderStyle = BorderStyle.FixedSingle
        PanelAI_Header.BackColor = accentBlue
        PanelAI_Header.Height = 50
        LabelAI_Title.ForeColor = Color.White
        LabelAI_Title.Font = New Font("Segoe UI", 11.0!, FontStyle.Bold)
        LabelAI_Title.Text = "Assistente AI — Ricerca Articoli"
        RtbChat.BackColor = Color.White
        RtbChat.Font = New Font("Segoe UI", 9.5!)
        PanelAI_Btns.BackColor = accentDark
        BtnAiInvia.BackColor = accentBlue
        BtnAiInvia.ForeColor = Color.White
        BtnAiInvia.FlatAppearance.BorderColor = accentDark
        BtnAiInvia.FlatAppearance.MouseOverBackColor = accentDark

        ' Allarga colonna AI dal 30% al 34%
        Dim tl = TryCast(PanelAI.Parent, TableLayoutPanel)
        If tl IsNot Nothing Then
            tl.ColumnStyles(0).Width = 34
            tl.ColumnStyles(1).Width = 66
        End If

        AggiungiPanelRicercaImmagine()
    End Sub

    Private Sub AggiungiPanelRicercaImmagine()
        If PanelAI.Controls.ContainsKey("PanelImgRicerca") Then Return

        Dim accentBlue As Color = Color.FromArgb(0, 120, 212)
        Dim accentDark As Color = Color.FromArgb(0, 90, 160)

        ' ── Contenitore principale ──────────────────────────────────────
        Dim pnl As New Panel()
        pnl.Name = "PanelImgRicerca"
        pnl.Dock = DockStyle.Bottom
        pnl.Height = 145
        pnl.BackColor = Color.FromArgb(230, 240, 255)
        pnl.Padding = New Padding(6)

        ' ── Separatore visivo ───────────────────────────────────────────
        Dim sep As New Panel()
        sep.Dock = DockStyle.Top
        sep.Height = 2
        sep.BackColor = accentBlue
        pnl.Controls.Add(sep)

        ' ── Titolino ────────────────────────────────────────────────────
        Dim lblTit As New Label()
        lblTit.Text = "Ricerca per immagine"
        lblTit.Dock = DockStyle.Top
        lblTit.Height = 20
        lblTit.Font = New Font("Segoe UI", 8.25!, FontStyle.Bold)
        lblTit.ForeColor = accentDark
        lblTit.TextAlign = ContentAlignment.MiddleLeft
        pnl.Controls.Add(lblTit)

        ' ── PictureBox ──────────────────────────────────────────────────
        _picRicercaImg.SizeMode = PictureBoxSizeMode.Zoom
        _picRicercaImg.BorderStyle = BorderStyle.FixedSingle
        _picRicercaImg.BackColor = Color.White
        _picRicercaImg.Width = 110
        _picRicercaImg.Height = 100
        _picRicercaImg.Location = New Point(6, 26)
        _picRicercaImg.Cursor = Cursors.Hand
        _picRicercaImg.TabStop = True
        pnl.Controls.Add(_picRicercaImg)
        AddHandler _picRicercaImg.Click, AddressOf PicRicercaImg_Click

        ' ── Hint sul PictureBox ─────────────────────────────────────────
        _lblPastaHint.Text = "Clicca qui" & vbCrLf & "e premi Ctrl+V" & vbCrLf & "per incollare"
        _lblPastaHint.AutoSize = False
        _lblPastaHint.Size = New Size(110, 100)
        _lblPastaHint.Location = New Point(6, 26)
        _lblPastaHint.TextAlign = ContentAlignment.MiddleCenter
        _lblPastaHint.ForeColor = Color.Gray
        _lblPastaHint.Font = New Font("Segoe UI", 8.0!, FontStyle.Italic)
        _lblPastaHint.BackColor = Color.Transparent
        _lblPastaHint.Cursor = Cursors.Hand
        pnl.Controls.Add(_lblPastaHint)
        AddHandler _lblPastaHint.Click, AddressOf PicRicercaImg_Click

        ' ── Pulsante cerca per immagine ─────────────────────────────────
        _btnCercaImmagine.Text = "Cerca per immagine"
        _btnCercaImmagine.Location = New Point(124, 26)
        _btnCercaImmagine.Size = New Size(145, 36)
        _btnCercaImmagine.FlatStyle = FlatStyle.Flat
        _btnCercaImmagine.BackColor = accentBlue
        _btnCercaImmagine.ForeColor = Color.White
        _btnCercaImmagine.Font = New Font("Segoe UI", 8.5!, FontStyle.Bold)
        _btnCercaImmagine.FlatAppearance.BorderColor = accentDark
        _btnCercaImmagine.Enabled = False
        pnl.Controls.Add(_btnCercaImmagine)
        AddHandler _btnCercaImmagine.Click, AddressOf BtnCercaImmagine_Click

        ' ── Label "Svuota" ──────────────────────────────────────────────
        Dim btnSvuota As New Button()
        btnSvuota.Text = "Svuota"
        btnSvuota.Location = New Point(124, 70)
        btnSvuota.Size = New Size(145, 26)
        btnSvuota.FlatStyle = FlatStyle.Flat
        btnSvuota.BackColor = Color.FromArgb(230, 240, 255)
        btnSvuota.ForeColor = accentDark
        btnSvuota.Font = New Font("Segoe UI", 8.0!)
        btnSvuota.FlatAppearance.BorderColor = Color.FromArgb(180, 200, 230)
        pnl.Controls.Add(btnSvuota)
        AddHandler btnSvuota.Click, Sub(s, ev)
                                        _picRicercaImg.Image = Nothing
                                        _lblPastaHint.Visible = True
                                        _btnCercaImmagine.Enabled = False
                                    End Sub

        PanelAI.Controls.Add(pnl)
        pnl.BringToFront()
        PanelAI_Bottom.BringToFront()
    End Sub

    Private Sub BRB_Click(sender As Object, e As EventArgs) Handles BRB.Enter
        visualizzazione = "BRB"
    End Sub

    Private Sub TIRELLI_Click(sender As Object, e As EventArgs) Handles Tirelli.Enter
        visualizzazione = "TIRELLI"
    End Sub

    Private Sub Button_CERCA_Click_1(sender As Object, e As EventArgs) Handles Button_CERCA.Click
        If visualizzazione = "TIRELLI" Then
            Dim codGruppo As String = If(ComboBox1.SelectedIndex > 0, Elenco_gruppi(ComboBox1.SelectedIndex), "")
            cerca(TextBox_codice_SAP_RICERCA.Text.ToUpper, TextBox_Descrizione_ricerca.Text.ToUpper, TextBox3.Text.ToUpper, TextBox_disegno_SAP.Text.ToUpper, TextBox6.Text.ToUpper, codGruppo, ComboBox4.Text.ToUpper, TextBox2.Text.ToUpper, ComboBox6.Text.ToUpper, TextBox4.Text.ToUpper)
        Else
            cerca_BRB()
        End If
    End Sub

    Private Sub TextBox_disegno_SAP_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox_disegno_SAP.TextChanged
        filtro_disegno = CostruisciCondizioneSAP("T0.[u_disegno]", TextBox_disegno_SAP.Text)
    End Sub

    Private Sub TextBox6_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        filtro_fam_disegno = CostruisciCondizioneSAP("t0.u_famiglia_disegno", TextBox6.Text)
    End Sub

    Private Sub DataGridView_SAP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_SAP.CellContentClick
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = DataGridView_SAP.Columns.IndexOf(Codice) Then
                Magazzino.Codice_SAP = DataGridView_SAP.Rows(e.RowIndex).Cells(columnName:="Codice").Value
                Magazzino.Show()
                Magazzino.BringToFront()
                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)

            ElseIf e.ColumnIndex = DataGridView_SAP.Columns.IndexOf(Disegno_) Then
                Magazzino.visualizza_disegno(DataGridView_SAP.Rows(e.RowIndex).Cells(columnName:="Disegno_").Value)

            ElseIf DataGridView_SAP.Columns.Contains("Immagine") AndAlso
                   e.ColumnIndex = DataGridView_SAP.Columns("Immagine").Index Then
                Dim disegnoCell = DataGridView_SAP.Rows(e.RowIndex).Cells("Disegno_").Value?.ToString()
                If Not String.IsNullOrWhiteSpace(disegnoCell) Then
                    Form_visualizza_picture.Show()
                    Magazzino.visualizza_picture(disegnoCell, Form_visualizza_picture.PictureBox1)
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView_SAP_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_SAP.CellFormatting
        If e.RowIndex < 0 Then Return
        Dim val = DataGridView_SAP.Rows(e.RowIndex).Cells("Attivo").Value?.ToString()
        If val = "Y" OrElse val = "I" Then
            DataGridView_SAP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightSlateGray
            DataGridView_SAP.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.DimGray
        Else
            DataGridView_SAP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Empty
            DataGridView_SAP.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.Empty
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Codice_Tirelli) Then
                Magazzino.Codice_SAP = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice_tirelli").Value
                Magazzino.Show()
                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)
            End If
        End If
    End Sub

    Private Sub TextBox_Descrizione_ricerca_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox_Descrizione_ricerca.TextChanged
        If visualizzazione = "TIRELLI" Then
            filtro_descrizione = CostruisciCondizioneSAP("t0.itemname", TextBox_Descrizione_ricerca.Text)
        Else
            filtro_descrizione = CostruisciCondizioneOrSAP(
                New String() {"t10.descrizione_brb", "t10.itemname"},
                TextBox_Descrizione_ricerca.Text)
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If visualizzazione = "TIRELLI" Then
            filtro_descrizione_supp = CostruisciCondizioneSAP("T0.[FrgnName]", TextBox3.Text)
        Else
            filtro_descrizione_supp = CostruisciCondizioneOrSAP(
                New String() {"t10.descrizione_supp", "t10.frgnname"},
                TextBox3.Text)
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        filtro_catalogo = CostruisciCondizioneSAP("T0.[suppcatnum]", TextBox2.Text)
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button20_Click_1(sender As Object, e As EventArgs) Handles Button20.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub DataGridView_SAP_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_SAP.CellClick

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

    End Sub

    Private Sub TextBox_codice_SAP_RICERCA_TextChanged(sender As Object, e As EventArgs) Handles TextBox_codice_SAP_RICERCA.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    ' ──────────────────────────────────────────────────────────────────
    '  FUNZIONI FILTRO — logica asterischi
    '  *testo    → LIKE 'testo%'   (inizia con)
    '  testo*    → LIKE '%testo'   (finisce con)
    '  *testo*   → LIKE '%testo%'  (contiene, = senza asterischi)
    '  t1*t2     → LIKE '%t1%' AND LIKE '%t2%'  (multi-termine)
    '  **t1**t2  → idem con doppio separatore
    ' ──────────────────────────────────────────────────────────────────

    ''' <summary>Condizione WHERE SQL Server per un singolo campo.</summary>
    Private Function CostruisciCondizioneSAP(campo As String, testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then Return ""
        Dim t As String = testo.Trim()

        ' Inizia con: un solo * all'inizio, nessun altro *
        If t.StartsWith("*") AndAlso t.IndexOf("*"c, 1) < 0 Then
            Return " and " & campo & " LIKE '" & t.Substring(1) & "%'"
        End If
        ' Finisce con: un solo * alla fine
        If t.EndsWith("*") AndAlso t.LastIndexOf("*"c, t.Length - 2) < 0 Then
            Return " and " & campo & " LIKE '%" & t.Substring(0, t.Length - 1) & "'"
        End If
        ' Multi-termine o contiene
        Dim terms = t.Split({"*"c}, StringSplitOptions.RemoveEmptyEntries)
        Return String.Join("", terms.Where(Function(x) x.Trim() <> "").Select(Function(x) " and " & campo & " LIKE '%" & x.Trim() & "%'"))
    End Function

    ''' <summary>Condizione WHERE SQL Server con OR su più campi (es. BRB).</summary>
    Private Function CostruisciCondizioneOrSAP(campi As String(), testo As String) As String
        If String.IsNullOrWhiteSpace(testo) OrElse campi.Length = 0 Then Return ""
        Dim t As String = testo.Trim()

        Dim OredCond = Function(likeExpr As String) As String
                           Return " and (" & String.Join(" or ", campi.Select(Function(c) c & " LIKE '" & likeExpr & "'")) & ")"
                       End Function

        If t.StartsWith("*") AndAlso t.IndexOf("*"c, 1) < 0 Then Return OredCond(t.Substring(1) & "%")
        If t.EndsWith("*") AndAlso t.LastIndexOf("*"c, t.Length - 2) < 0 Then Return OredCond("%" & t.Substring(0, t.Length - 1))

        Dim terms = t.Split({"*"c}, StringSplitOptions.RemoveEmptyEntries)
        Return String.Join("", terms.Where(Function(x) x.Trim() <> "").Select(Function(x) OredCond("%" & x.Trim() & "%")))
    End Function

    ''' <summary>Condizione WHERE AS400/OPENQUERY per un singolo campo (escape '' per le virgolette).</summary>
    Private Function CostruisciCondizioneAS400(campo As String, testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then Return ""
        Dim t As String = testo.Trim().ToUpper()

        If t.StartsWith("*") AndAlso t.IndexOf("*"c, 1) < 0 Then
            Return " and upper(" & campo & ") LIKE ''" & t.Substring(1) & "%''"
        End If
        If t.EndsWith("*") AndAlso t.LastIndexOf("*"c, t.Length - 2) < 0 Then
            Return " and upper(" & campo & ") LIKE ''%" & t.Substring(0, t.Length - 1) & "''"
        End If
        Dim terms = t.Split({"*"c}, StringSplitOptions.RemoveEmptyEntries)
        Return String.Join("", terms.Where(Function(x) x.Trim() <> "").Select(Function(x) " and upper(" & campo & ") LIKE ''%" & x.Trim() & "%''"))
    End Function

    ' ──────────────────────────────────────────────────────────────────
    '  PANNELLO AI
    ' ──────────────────────────────────────────────────────────────────

    Private Sub BtnAiInvia_Click(sender As Object, e As EventArgs) Handles BtnAiInvia.Click
        InviaMessaggioAI()
    End Sub

    Private Sub TxtAiInput_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtAiInput.KeyDown
        ' Enter invia, Shift+Enter va a capo
        If e.KeyCode = Keys.Enter AndAlso Not e.Shift Then
            e.SuppressKeyPress = True
            InviaMessaggioAI()
        End If
    End Sub

    Private Async Sub InviaMessaggioAI()
        Dim domanda = TxtAiInput.Text.Trim()
        If String.IsNullOrEmpty(domanda) Then Return

        AppendChat("Tu:  " & domanda, False)
        TxtAiInput.Clear()
        BtnAiInvia.Enabled = False
        BtnAiInvia.Text = "..."

        Try
            Dim risposta = Await ChiediAIAsync(domanda)
            AppendChat("AI:  " & risposta, True)
        Catch ex As Exception
            AppendChat("[Errore] " & ex.Message, True)
        Finally
            BtnAiInvia.Enabled = True
            BtnAiInvia.Text = "Invia  (Invio)"
        End Try
    End Sub

    ''' <summary>
    ''' Chiama l'API Anthropic, interpreta la risposta JSON e gestisce l'azione.
    ''' Restituisce il testo "messaggio" da mostrare in chat.
    ''' Se azione="cerca" precompila i filtri ed esegue la ricerca.
    ''' </summary>
    Private Async Function ChiediAIAsync(domanda As String) As Threading.Tasks.Task(Of String)
        Dim keyPath = IO.Path.Combine(Application.StartupPath, "anthropic_key.txt")
        If Not IO.File.Exists(keyPath) Then
            Return "Chiave API non trovata. Crea il file 'anthropic_key.txt' nella cartella dell'applicazione."
        End If
        Dim apiKey = IO.File.ReadAllText(keyPath).Trim()

        _aiHistory.Add(New JObject From {{"role", "user"}, {"content", domanda}})

        Dim requestBody = New JObject From {
            {"model", "claude-sonnet-4-6"},
            {"max_tokens", 1500},
            {"system", AI_SYSTEM_PROMPT},
            {"messages", _aiHistory}
        }

        Using client As New HttpClient()
            client.DefaultRequestHeaders.Add("x-api-key", apiKey)
            client.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01")
            client.Timeout = TimeSpan.FromSeconds(60)

            Dim content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
            Dim response = Await client.PostAsync("https://api.anthropic.com/v1/messages", content)
            Dim body = Await response.Content.ReadAsStringAsync()

            If Not response.IsSuccessStatusCode Then
                Throw New Exception("Errore API " & CInt(response.StatusCode) & ": " & body)
            End If

            Dim apiParsed = JObject.Parse(body)
            Dim tokenVal = apiParsed.SelectToken("content[0].text")
            Dim rawTesto As String = If(tokenVal IsNot Nothing, tokenVal.ToString().Trim(), "(risposta vuota)")

            _aiHistory.Add(New JObject From {{"role", "assistant"}, {"content", rawTesto}})

            ' Interpreta il JSON di risposta del modello
            Return InterpretaRispostaAI(rawTesto)
        End Using
    End Function

    ''' <summary>
    ''' Analizza il JSON restituito dal modello.
    ''' Se azione=cerca, precompila i filtri e avvia la ricerca (su UI thread).
    ''' Restituisce il testo messaggio da mostrare.
    ''' </summary>
    Private Function InterpretaRispostaAI(rawTesto As String) As String
        Try
            ' Rimuove eventuali blocchi ```json ... ``` lasciati dal modello
            Dim jsonPulito = rawTesto
            Dim mdStart = jsonPulito.IndexOf("{")
            Dim mdEnd = jsonPulito.LastIndexOf("}")
            If mdStart >= 0 AndAlso mdEnd > mdStart Then
                jsonPulito = jsonPulito.Substring(mdStart, mdEnd - mdStart + 1)
            End If

            Dim risposta = JObject.Parse(jsonPulito)
            Dim azione = If(risposta("azione") IsNot Nothing, risposta("azione").ToString().ToLower(), "info")
            Dim messaggio = If(risposta("messaggio") IsNot Nothing, risposta("messaggio").ToString(), rawTesto)

            If azione = "cerca" AndAlso risposta("filtri") IsNot Nothing Then
                ' Esegui su UI thread (questa funzione è già su UI thread grazie ad Await)
                ApplicaFiltriAI(CType(risposta("filtri"), JObject))
                Return messaggio & vbCrLf & "[Ricerca avviata — vedi risultati nella griglia]"
            End If

            Return messaggio

        Catch ex As Exception
            ' Se il JSON è malformato, mostra il testo grezzo
            Return rawTesto
        End Try
    End Function

    ''' <summary>
    ''' Precompila i campi filtro dall'oggetto JSON "filtri" e avvia la ricerca.
    ''' </summary>
    Private Sub ApplicaFiltriAI(filtri As JObject)
        ' Pulisce tutti i filtri
        TextBox_codice_SAP_RICERCA.Text = ""
        TextBox_Descrizione_ricerca.Text = ""
        TextBox3.Text = ""
        TextBox_disegno_SAP.Text = ""
        TextBox6.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        ComboBox4.Text = ""
        ComboBox6.Text = ""

        ' Applica i valori ricevuti dall'AI (solo quelli non vuoti)
        Dim prendi = Function(chiave As String) As String
                         Dim tok = filtri(chiave)
                         Return If(tok IsNot Nothing, tok.ToString().Trim(), "")
                     End Function

        TextBox_codice_SAP_RICERCA.Text = prendi("codice")
        TextBox_Descrizione_ricerca.Text = prendi("descrizione")
        TextBox3.Text = prendi("desc_supp")
        TextBox_disegno_SAP.Text = prendi("disegno")
        TextBox2.Text = prendi("catalogo")
        TextBox4.Text = prendi("ubicazione")

        ' ComboBox (testo libero per ricerca parziale)
        Dim prod = prendi("produttore")
        If prod <> "" Then ComboBox4.Text = prod

        Dim forn = prendi("fornitore")
        If forn <> "" Then
            ComboBox6.Text = forn
            Dim idx = ComboBox6.FindString(forn)
            If idx >= 0 Then ComboBox6.SelectedIndex = idx
        End If

        Dim grp = prendi("gruppo")
        If grp <> "" Then
            Dim idxGrp = ComboBox1.FindString(grp)
            If idxGrp >= 0 Then
                ComboBox1.SelectedIndex = idxGrp
            Else
                ComboBox1.SelectedIndex = 0
            End If
        Else
            ComboBox1.SelectedIndex = 0
        End If

        ' Avvia la ricerca
        Button_CERCA.PerformClick()
    End Sub

    Private Sub AppendChat(testo As String, isAI As Boolean)
        RtbChat.SelectionStart = RtbChat.TextLength
        RtbChat.SelectionLength = 0
        If isAI Then
            RtbChat.SelectionColor = Color.FromArgb(22, 45, 84)
            RtbChat.SelectionFont = New Font(RtbChat.Font, FontStyle.Regular)
        Else
            RtbChat.SelectionColor = Color.FromArgb(50, 50, 50)
            RtbChat.SelectionFont = New Font(RtbChat.Font, FontStyle.Bold)
        End If
        RtbChat.AppendText(testo & vbCrLf & vbCrLf)
        RtbChat.SelectionStart = RtbChat.TextLength
        RtbChat.ScrollToCaret()
    End Sub

    Private Sub BtnAiPulisci_Click(sender As Object, e As EventArgs) Handles BtnAiPulisci.Click
        _aiHistory = New JArray()
        RtbChat.Clear()
    End Sub

    ' ──────────────────────────────────────────────────────────────────
    '  IMMAGINI NELLA GRIGLIA
    ' ──────────────────────────────────────────────────────────────────

    Private Sub CaricaImmagineRiga(rowIdx As Integer, disegno As String)
        If String.IsNullOrWhiteSpace(disegno) Then Return
        Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & disegno & ".PNG"
        If Not IO.File.Exists(percorso) Then Return
        Try
            Using fs As New IO.FileStream(percorso, IO.FileMode.Open, IO.FileAccess.Read)
                Using tmp As Image = Image.FromStream(fs)
                    DataGridView_SAP.Rows(rowIdx).Cells("Immagine").Value = New Bitmap(tmp)
                End Using
            End Using
        Catch
        End Try
    End Sub

    ' ──────────────────────────────────────────────────────────────────
    '  RICERCA PER IMMAGINE — paste e invio a Claude Vision
    ' ──────────────────────────────────────────────────────────────────

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        If keyData = (Keys.Control Or Keys.V) Then
            If _picRicercaImg IsNot Nothing AndAlso _picRicercaImg.Focused Then
                PicRicercaImg_IncollaClipboard()
                Return True
            End If
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Private Sub PicRicercaImg_Click(sender As Object, e As EventArgs)
        _picRicercaImg.Focus()
        PicRicercaImg_IncollaClipboard()
    End Sub

    Private Sub PicRicercaImg_IncollaClipboard()
        If Clipboard.ContainsImage() Then
            _picRicercaImg.Image = Clipboard.GetImage()
            _lblPastaHint.Visible = False
            _btnCercaImmagine.Enabled = True
        ElseIf Clipboard.ContainsData("PNG") OrElse Clipboard.ContainsData("Bitmap") Then
            _picRicercaImg.Image = DirectCast(Clipboard.GetData(DataFormats.Bitmap), Image)
            _lblPastaHint.Visible = False
            _btnCercaImmagine.Enabled = True
        Else
            AppendChat("AI:  Nessuna immagine negli appunti. Copia prima uno screenshot con Stamp o con lo strumento di cattura.", True)
        End If
    End Sub

    Private Async Sub BtnCercaImmagine_Click(sender As Object, e As EventArgs)
        If _picRicercaImg.Image Is Nothing Then Return
        _btnCercaImmagine.Enabled = False
        _btnCercaImmagine.Text = "..."
        AppendChat("Tu:  [ricerca per immagine]", False)
        Try
            ' Converti immagine in PNG base64
            Dim imgBase64 As String
            Using ms As New IO.MemoryStream()
                _picRicercaImg.Image.Save(ms, Imaging.ImageFormat.Png)
                imgBase64 = Convert.ToBase64String(ms.ToArray())
            End Using
            Dim risposta = Await ChiediAIConImmagineAsync(imgBase64)
            AppendChat("AI:  " & risposta, True)
        Catch ex As Exception
            AppendChat("[Errore] " & ex.Message, True)
        Finally
            _btnCercaImmagine.Enabled = True
            _btnCercaImmagine.Text = "Cerca per immagine"
        End Try
    End Sub

    Private Async Function ChiediAIConImmagineAsync(imgBase64 As String) As Threading.Tasks.Task(Of String)
        Dim keyPath = IO.Path.Combine(Application.StartupPath, "anthropic_key.txt")
        If Not IO.File.Exists(keyPath) Then Return "Chiave API non trovata."
        Dim apiKey = IO.File.ReadAllText(keyPath).Trim()

        Dim userContent = New JArray(
            New JObject From {
                {"type", "image"},
                {"source", New JObject From {
                    {"type", "base64"},
                    {"media_type", "image/png"},
                    {"data", imgBase64}
                }}
            },
            New JObject From {
                {"type", "text"},
                {"text", "Analizza questa immagine di un componente meccanico/articolo. Identifica di che tipo di pezzo si tratta e imposta i filtri per trovarlo nel database Tirelli. Usa descrizione, produttore, gruppo articoli e gli altri campi appropriati."}
            }
        )

        Dim requestBody = New JObject From {
            {"model", "claude-sonnet-4-6"},
            {"max_tokens", 1500},
            {"system", AI_SYSTEM_PROMPT},
            {"messages", New JArray(New JObject From {{"role", "user"}, {"content", userContent}})}
        }

        Using client As New HttpClient()
            client.DefaultRequestHeaders.Add("x-api-key", apiKey)
            client.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01")
            client.Timeout = TimeSpan.FromSeconds(60)
            Dim content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
            Dim response = Await client.PostAsync("https://api.anthropic.com/v1/messages", content)
            Dim body = Await response.Content.ReadAsStringAsync()
            If Not response.IsSuccessStatusCode Then
                Throw New Exception("Errore API " & CInt(response.StatusCode) & ": " & body)
            End If
            Dim apiParsed = JObject.Parse(body)
            Dim tokenVal = apiParsed.SelectToken("content[0].text")
            Dim rawTesto As String = If(tokenVal IsNot Nothing, tokenVal.ToString().Trim(), "(risposta vuota)")
            _aiHistory.Add(New JObject From {{"role", "user"}, {"content", userContent}})
            _aiHistory.Add(New JObject From {{"role", "assistant"}, {"content", rawTesto}})
            Return InterpretaRispostaAI(rawTesto)
        End Using
    End Function

    Private Sub BtnLegenda_Click(sender As Object, e As EventArgs) Handles BtnLegenda.Click
        MostraLegenda()
    End Sub

    Private Sub MostraLegenda()
        Dim msg As String =
            "LEGENDA CRITERI DI RICERCA" & vbCrLf &
            String.Empty.PadRight(40, "─"c) & vbCrLf & vbCrLf &
            "Sintassi nei campi di testo:" & vbCrLf & vbCrLf &
            "  parola          →  contiene 'parola'" & vbCrLf &
            "  *parola         →  inizia con 'parola'" & vbCrLf &
            "  parola*         →  finisce con 'parola'" & vbCrLf &
            "  *parola*        →  contiene 'parola'  (= senza *)" & vbCrLf &
            "  par1*par2       →  contiene 'par1'  E  contiene 'par2'" & vbCrLf &
            "  **p1**p2        →  contiene 'p1'  E  contiene 'p2'" & vbCrLf & vbCrLf &
            "Esempi:" & vbCrLf &
            "  *cuscinetto    →  LIKE 'cuscinetto%'" & vbCrLf &
            "  cuscinetto*    →  LIKE '%cuscinetto'" & vbCrLf &
            "  SKF*6205       →  LIKE '%SKF%' AND LIKE '%6205%'" & vbCrLf & vbCrLf &
            "Nota: le ricerche non distinguono maiuscole/minuscole."

        MessageBox.Show(msg, "Legenda criteri di ricerca",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Class
