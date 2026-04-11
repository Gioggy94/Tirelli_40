Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports System.Windows.Documents
Imports System.Diagnostics


Public Class UT
    Public Elenco_gruppi(10000) As String
    Public Elenco_fornitori(10000) As String
    Public Elenco_produttori(10000) As String
    Public Elenco_settori(1000) As String
    Public Elenco_paesi(1000) As String
    Public Elenco_codici_sap_dipendenti(1000) As String
    Public Elenco_codici_dipendenti(1000) As String
    Public Elenco_nome_SAP_dipendenti(1000) As String

    Private filtro_fam_disegno As String
    Private visualizzazione As String = "TIRELLI"
    Public SETTORE As String
    Public paese As String
    Public brand As String
    Public codice_BRB As String
    Private riga As Integer
    Private filtro_descrizione As String
    Public filtro_descrizione_supp As String
    Private filtro_disegno As String
    Private filtro_catalogo As String
    Private ubicazione_labelling As String = ""

    Sub inizializzazione_ut()
        Dim stopwatch As New Stopwatch()

        ' Misura il tempo di esecuzione di inserimento_gruppi
        stopwatch.Start()
        inserimento_gruppi(ComboBox_gruppo_articoli)
        inserimento_gruppi(ComboBox1)
        inserimento_gruppi(ComboBox9)
        stopwatch.Stop()
        Console.WriteLine("Tempo impiegato per inserimento_gruppi: " & stopwatch.Elapsed.TotalSeconds & " secondi")

        ' Misura il tempo di esecuzione di inserimento_produttore
        stopwatch.Reset()
        stopwatch.Start()
        inserimento_produttore()
        stopwatch.Stop()
        Console.WriteLine("Tempo impiegato per inserimento_produttore: " & stopwatch.Elapsed.TotalSeconds & " secondi")

        ' Misura il tempo di esecuzione di inserimento_settori
        stopwatch.Reset()
        stopwatch.Start()
        inserimento_settori()
        stopwatch.Stop()
        Console.WriteLine("Tempo impiegato per inserimento_settori: " & stopwatch.Elapsed.TotalSeconds & " secondi")

        ' Misura il tempo di esecuzione di inserimento_FORNITORI
        stopwatch.Reset()
        stopwatch.Start()
        inserimento_FORNITORI()
        stopwatch.Stop()
        Console.WriteLine("Tempo impiegato per inserimento_FORNITORI: " & stopwatch.Elapsed.TotalSeconds & " secondi")

        ' Misura il tempo di esecuzione di inserimento_paesi
        stopwatch.Reset()
        stopwatch.Start()
        inserimento_paesi()
        stopwatch.Stop()
        Console.WriteLine("Tempo impiegato per inserimento_paesi: " & stopwatch.Elapsed.TotalSeconds & " secondi")

        ' Misura il tempo di esecuzione di Inserimento_dipendenti
        stopwatch.Reset()
        stopwatch.Start()
        Inserimento_dipendenti()
        stopwatch.Stop()
        Console.WriteLine("Tempo impiegato per Inserimento_dipendenti: " & stopwatch.Elapsed.TotalSeconds & " secondi")
    End Sub


    Sub Inserimento_dipendenti()

        Combodipendenti.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.EMPID, case when T0.[USERID] is null then '' else t0.userid end as 'Codice dipendenti'
, T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 

, '' as  'position'
        FROM [TIRELLI_40].[dbo].OHEM T0 
left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code  

where t0.active='Y' order by T0.[lastName]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer = 0


        Do While cmd_SAP_reader.Read()

            Combodipendenti.Items.Add(cmd_SAP_reader("Nome"))
            Elenco_codici_sap_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            Elenco_codici_dipendenti(Indice) = cmd_SAP_reader("empid")

            Indice = Indice + 1


        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

        'If Homepage.UTENTE_NOME_SALVATO <> "" Then
        '    Combodipendenti.Text = Homepage.UTENTE_NOME_SALVATO
        'End If

    End Sub 'Inserisco le risorse nella combo box
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
            Cmd_SAP.CommandText = "select 
    trim(t10.code) as 'Codice'
    ,t10.des_code as 'Nome'
    ,'MANCA' AS 'Nome supp'
    ,trim(t10.disegno) as 'Disegno'
    ,t10.grup_art as 'Gruppo art'
    ,t10.prod_for as 'firmname'
    ,t10.codar_for as 'Suppcatnum'
    ,t10.desc_for as 'Cardname'
    ,t10.ubi_cOde as 'u_ubicazione'
    ,'manca' as 'u_famiglia_disegno'
    ,T10.onhand as 'onhand'
    ,9999 as 'WIP'
    ,T10.disp_tot as 'DISP'
    ,T10.costo_std as 'PRICE'
    ,'???' AS 'fROZENFOR'
from openquery(AS400,'
    select 
        T0.CODE, 
        T0.DES_CODE, 
        T0.DISEGNO, 
        T0.GRUP_ART, 
        T0.PROD_FOR, 
        T0.CODAR_FOR,
        T0.DESC_FOR, 
        T0.UBI_CODE,
        T0.COSTO_STD,
        coalesce(T1.TOT_QTA, 0) as onhand,
        coalesce(T1.TOT_DISP, 0) as disp_tot
    from S786FAD1.TIR90VIS.JGALART T0
    left join (
        select 
            CODART, 
            sum(QTA_MAG) as TOT_QTA,
            sum(QTA_DISP) as TOT_DISP
        from TIR90VIS.JGALMAG 
        group by CODART
    ) T1 ON T0.CODE = T1.CODART
where   upper(code) LIKE ''%" & Par_codice & "%''
	and upper(t0.des_code)  LIKE ''%" & par_descrizione & "%''
	and upper(t0.disegno)  LIKE ''%" & par_disegno & "%''
	and  upper(t0.grup_art)  LIKE ''%" & par_gruppo_art & "%''
	and upper(t0.prod_for)  LIKE ''%" & par_produttore & "%''
	and  upper(t0.codar_for) LIKE ''%" & par_cat_forn & "%''
	and  upper(t0.desc_for) LIKE ''%" & par_fornitore & "%''
	and upper(t0.ubi_cOde)  LIKE ''%" & par_ubicazione & "%''
'
)
as t10"

        End If

        Cmd_SAP_Reader = Cmd_SAP.ExecuteReader

        Do While Cmd_SAP_Reader.Read()

            DataGridView_SAP.Rows.Add(Cmd_SAP_Reader("Codice"), Cmd_SAP_Reader("Nome"), Cmd_SAP_Reader("Nome supp"), Cmd_SAP_Reader("u_famiglia_disegno"), Cmd_SAP_Reader("Disegno"), Cmd_SAP_Reader("Gruppo art"), Cmd_SAP_Reader("firmname"), Cmd_SAP_Reader("suppcatnum"), Cmd_SAP_Reader("cardname"), Cmd_SAP_Reader("u_ubicazione"), Cmd_SAP_Reader("onhand"), Cmd_SAP_Reader("wip"), Cmd_SAP_Reader("disp"), Cmd_SAP_Reader("price"), "???")
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
        cnn.Open()
        Dim Cmd_SAP As New SqlCommand
        Dim Cmd_SAP_Reader As SqlDataReader

        Cmd_SAP.Connection = cnn
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
        cnn.Close()
    End Sub





    Sub inserisci(par_utente_SAP As String, par_codice_articolo As String, par_descrizione_articolo As String, par_desc_supp As String, par_codice_disegno As String, par_gruppo_articoli As Integer, par_fornitore_preferito As String, par_catalogo_fornitore As String, par_produttore As String, par_tipo_montaggio As String, par_codice_BP As String, par_nome_bp As String, par_settore As String, par_paese As String, par_agente As String, par_brand As String, par_codice_BRB As String, par_phantom As String, par_costo As Decimal, par_ubicazione As String, par_approvvigionamento As String)


        insert_into_OITM(par_utente_SAP, par_codice_articolo, par_descrizione_articolo, par_desc_supp, par_codice_disegno, par_gruppo_articoli, par_fornitore_preferito, par_catalogo_fornitore, par_produttore, par_tipo_montaggio, par_codice_BP, par_nome_bp, par_settore, par_paese, par_agente, par_brand, par_codice_BRB, 0, par_phantom, par_ubicazione, par_approvvigionamento)



        Try


            ITM1(TextBox_codice_sap.Text, par_costo)
        Catch ex As Exception
            MsgBox("C'è un errore nella tabella ITM1")
        End Try

        Try

            OITW(TextBox_codice_sap.Text)
        Catch ex As Exception
            MsgBox("C'è un errore nella tabella OITW")
        End Try



    End Sub

    Sub inserisci_Nuovo_codice(par_utente_SAP As String, par_codice_articolo As String, par_descrizione_articolo As String, par_desc_supp As String, par_codice_disegno As String, par_gruppo_articoli As Integer, par_fornitore_preferito As String, par_catalogo_fornitore As String, par_produttore As String, par_tipo_montaggio As String, par_codice_BP As String, par_nome_bp As String, par_settore As String, par_paese As String, par_agente As String, par_brand As String, par_codice_BRB As String, par_revisione As String, par_phantom As String, par_costo As Integer, par_ubicazione As String, par_approvvigionamento As String)




        insert_into_OITM(par_utente_SAP, par_codice_articolo, Replace(par_descrizione_articolo, "'", " "), Replace(par_desc_supp, "'", " "), Replace(par_codice_disegno, "'", " "), par_gruppo_articoli, par_fornitore_preferito, Replace(par_catalogo_fornitore, "'", " "), Replace(par_produttore, "'", " "), par_tipo_montaggio, par_codice_BP, Replace(par_nome_bp, "'", " "), par_settore, Replace(par_paese, "'", " "), Replace(par_agente, "'", " "), par_brand, par_codice_BRB, par_revisione, par_phantom, par_ubicazione, par_approvvigionamento)



        Try


            ITM1(par_codice_articolo, par_costo)
        Catch ex As Exception
            MsgBox("C'è un errore nella tabella ITM1")
        End Try

        Try

            OITW(par_codice_articolo)
        Catch ex As Exception
            MsgBox("C'è un errore nella tabella OITW")
        End Try



    End Sub

    Sub insert_into_OITM(par_utente_SAP As String, par_codice_articolo As String, par_descrizione_articolo As String, par_desc_supp As String, par_codice_disegno As String, par_gruppo_articoli As Integer, par_fornitore_preferito As String, par_catalogo_fornitore As String, par_produttore As String, par_tipo_montaggio As String, par_codice_BP As String, par_nome_bp As String, par_settore As String, par_paese As String, par_agente As String, par_brand As String, par_codice_BRB As String, par_revisione As String, par_phantom As String, par_ubicazione As String, par_approvvigionamento As String)
        par_nome_bp = Replace(par_nome_bp, "'", " ")
        Dim par_gestione_a_magazzino As String = "Y"

        If par_phantom = "Y" Then
            par_gestione_a_magazzino = "N"

        End If
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO OITM (OITM.ITEMCODE, OITM.ItemName, OITM.FrgnName, OITM.ItmsGrpCod, OITM.CstGrpCode, OITM.VatGourpSa, OITM.CodeBars, OITM.VATLiable, OITM.PrchseItem, OITM.SellItem, OITM.InvntItem, OITM.OnHand, OITM.IsCommited, OITM.OnOrder, OITM.IncomeAcct, OITM.ExmptIncom, OITM.MaxLevel, OITM.DfltWH, OITM.CardCode, OITM.SuppCatNum, OITM.BuyUnitMsr, OITM.NumInBuy, OITM.ReorderQty, OITM.MinLevel, OITM.LstEvlPric, OITM.LstEvlDate, OITM.Canceled, OITM.SalUnitMsr, OITM.NumInSale, OITM.Consig, OITM.Counted, OITM.EvalSystem, OITM.PicturName, OITM.UserText, OITM.CommisPcnt, OITM.CommisSum, OITM.CommisGrp, OITM.TreeType, OITM.TreeQty, OITM.LastPurPrc, OITM.LastPurCur, OITM.LastPurDat, OITM.ExitPrice, OITM.ExitWH, OITM.AssetItem, OITM.WasCounted, OITM.ManSerNum, OITM.SHeight1, OITM.SHght1Unit, OITM.SHeight2, OITM.SHght2Unit, OITM.SWidth1, OITM.SWdth1Unit, OITM.SWidth2, OITM.SWdth2Unit, OITM.SLength1, OITM.SLen1Unit, OITM.Slength2, OITM.SLen2Unit, OITM.SVolume, OITM.SVolUnit, OITM.SWeight1, OITM.SWght1Unit, OITM.SWeight2, OITM.SWght2Unit, OITM.BHeight1, OITM.BHght1Unit, OITM.BHeight2, OITM.BHght2Unit, OITM.BWidth1, OITM.BWdth1Unit, OITM.BWidth2, OITM.BWdth2Unit, OITM.BLength1, OITM.BLen1Unit, OITM.Blength2, OITM.BLen2Unit, OITM.BVolume, OITM.BVolUnit, OITM.BWeight1, OITM.BWght1Unit, OITM.BWeight2, OITM.BWght2Unit, OITM.FixCurrCms, OITM.FirmCode, OITM.LstSalDate, OITM.QryGroup1, OITM.QryGroup2, OITM.QryGroup3, OITM.QryGroup4, OITM.QryGroup5, OITM.QryGroup6, OITM.QryGroup7, OITM.QryGroup8, OITM.QryGroup9, OITM.QryGroup10, OITM.QryGroup11, OITM.QryGroup12, OITM.QryGroup13, OITM.QryGroup14, OITM.QryGroup15, OITM.QryGroup16, OITM.QryGroup17, OITM.QryGroup18, OITM.QryGroup19, OITM.QryGroup20, OITM.QryGroup21, OITM.QryGroup22, OITM.QryGroup23, OITM.QryGroup24, OITM.QryGroup25, OITM.QryGroup26, OITM.QryGroup27, OITM.QryGroup28, OITM.QryGroup29, OITM.QryGroup30, OITM.QryGroup31, OITM.QryGroup32, OITM.QryGroup33, OITM.QryGroup34, OITM.QryGroup35, OITM.QryGroup36, OITM.QryGroup37, OITM.QryGroup38, OITM.QryGroup39, OITM.QryGroup40, OITM.QryGroup41, OITM.QryGroup42, OITM.QryGroup43, OITM.QryGroup44, OITM.QryGroup45, OITM.QryGroup46, OITM.QryGroup47, OITM.QryGroup48, OITM.QryGroup49, OITM.QryGroup50, OITM.QryGroup51, OITM.QryGroup52, OITM.QryGroup53, OITM.QryGroup54, OITM.QryGroup55, OITM.QryGroup56, OITM.QryGroup57, OITM.QryGroup58, OITM.QryGroup59, OITM.QryGroup60, OITM.QryGroup61, OITM.QryGroup62, OITM.QryGroup63, OITM.QryGroup64, OITM.CreateDate, OITM.UpdateDate, OITM.SalFactor1, OITM.SalFactor2, OITM.SalFactor3, OITM.SalFactor4, OITM.PurFactor1, OITM.PurFactor2, OITM.PurFactor3, OITM.PurFactor4, OITM.SalFormula, OITM.PurFormula, OITM.VatGroupPu, OITM.AvgPrice, OITM.PurPackMsr, OITM.PurPackUn, OITM.SalPackMsr, OITM.SalPackUn, OITM.ManBtchNum, OITM.ManOutOnly, OITM.validFor, OITM.validFrom, OITM.validTo, OITM.frozenFor, OITM.frozenFrom, OITM.frozenTo, OITM.ValidComm, OITM.FrozenComm, OITM.SWW, OITM.Deleted, OITM.ExpensAcct, OITM.FrgnInAcct, OITM.ShipType, OITM.GLMethod, OITM.ECInAcct, OITM.FrgnExpAcc, OITM.ECExpAcc, OITM.TaxType, OITM.ByWh, OITM.WTLiable, OITM.ItemType, OITM.WarrntTmpl, OITM.BaseUnit, OITM.CountryOrg, OITM.StockValue, OITM.Phantom, OITM.IssueMthd, OITM.FREE1, OITM.PricingPrc, OITM.MngMethod, OITM.ReorderPnt, OITM.InvntryUom, OITM.PlaningSys, OITM.PrcrmntMtd, OITM.OrdrIntrvl, OITM.OrdrMulti, OITM.MinOrdrQty, OITM.LeadTime, OITM.IndirctTax, OITM.TaxCodeAR, OITM.TaxCodeAP, OITM.OSvcCode, OITM.ISvcCode, OITM.ServiceGrp, OITM.NCMCode, OITM.MatType, OITM.MatGrp, OITM.ProductSrc, OITM.ServiceCtg, OITM.ItemClass, OITM.Excisable, OITM.ChapterID, OITM.NotifyASN, OITM.ProAssNum, OITM.AssblValue, OITM.DNFEntry, OITM.Spec, OITM.TaxCtg, OITM.Series, OITM.Number, OITM.FuelCode, OITM.BeverTblC, OITM.BeverGrpC, OITM.BeverTM, OITM.Attachment, OITM.AtcEntry, OITM.ToleranDay, OITM.UgpEntry, OITM.PUoMEntry, OITM.SUoMEntry, OITM.IUoMEntry, OITM.IssuePriBy, OITM.AssetClass, OITM.AssetGroup, OITM.InventryNo, OITM.Technician, OITM.Employee, OITM.Location, OITM.StatAsset, OITM.Cession, OITM.DeacAftUL, OITM.AsstStatus, OITM.CapDate, OITM.AcqDate, OITM.RetDate, OITM.GLPickMeth, OITM.NoDiscount, OITM.MgrByQty, OITM.AssetRmk1, OITM.AssetRmk2, OITM.AssetAmnt1, OITM.AssetAmnt2, OITM.DeprGroup, OITM.AssetSerNo, OITM.CntUnitMsr, OITM.NumInCnt, OITM.INUoMEntry, OITM.OneBOneRec, OITM.RuleCode, OITM.ScsCode, OITM.SpProdType, OITM.IWeight1, OITM.IWght1Unit, OITM.IWeight2, OITM.IWght2Unit, OITM.CompoWH, OITM.CreateTS, OITM.UpdateTS, OITM.VirtAstItm, OITM.SouVirAsst, OITM.InCostRoll, OITM.PrdStdCst, OITM.EnAstSeri, OITM.LinkRsc, OITM.OnHldPert, OITM.onHldLimt, OITM.PriceUnit, OITM.GSTRelevnt, OITM.SACEntry, OITM.GstTaxCtg, OITM.AssVal4WTR, OITM.ExcImpQUoM, OITM.ExcFixAmnt, OITM.ExcRate, OITM.SOIExc, OITM.TNVED, OITM.Imported, OITM.AutoBatch, OITM.CstmActing, OITM.StdItemId, OITM.CommClass, OITM.TaxCatCode, OITM.DataVers, OITM.NVECode, OITM.CESTCode, OITM.CtrSealQty, OITM.LegalText, OITM.U_UBIMAG, OITM.U_TEMPOMED, OITM.U_MODMAC, OITM.U_SEZIONE, OITM.U_PRG_TIR_MATERIALE, OITM.U_DESCFR, OITM.U_DESCING, OITM.U_UTILIZZ, OITM.U_ARTCES, OITM.U_BA_IsFA, OITM.U_BA_TypID, OITM.U_BA_NumID, OITM.U_BA_LVAFrom, OITM.U_BA_LVA, OITM.U_ULTVER, OITM.U_ULTPREZ, OITM.U_PrzPeso, OITM.U_PRG_AZS_DesAggArt, OITM.U_PRG_AZS_DesAgg2Art, OITM.U_PRG_AZS_CMM, OITM.U_PRG_AZS_ItmsGrp2Cod, OITM.U_PRG_AZS_GestComm, OITM.U_PRG_AZS_CreatedBy, OITM.U_PRG_AZS_Phantom, OITM.U_PRG_CLV_Tipo_Lav, OITM.U_PRG_CLV_Lav_EstAss, OITM.U_PRG_CLV_Grezzo, OITM.U_PRG_CLV_TipologiaLav, OITM.U_PRG_CLV_ArtAssLav, OITM.U_PRG_CLV_SemiLavComodo, OITM.U_PRG_CLV_MagPrePro, OITM.U_Anagrafica, OITM.U_Ricambio, OITM.U_Collaudo, OITM.U_DocumentazioneTecnica, OITM.U_CostoMateriale, OITM.U_CostoPrimo, OITM.U_Margine, OITM.U_ListinoMinimo, OITM.U_Famiglia, OITM.U_ErroreWip, OITM.U_Movimentazioni, OITM.U_TraferitoODP, OITM.U_Superlistino, OITM.U_Superlistino_Vecchio, OITM.U_Inventariato, OITM.U_Contato_da, OITM.U_Disegno, OITM.U_Indice_di_revisione, OITM.U_Tipo_macchina, OITM.U_Numero_formati, OITM.U_Gestione_magazzino, OITM.U_Macchina_standard, OITM.U_PRG_TIR_Rev, OITM.U_PRG_TIR_RevProd, OITM.U_PRG_TIR_RevProdDate, OITM.U_Tipo_montaggio, OITM.U_PRG_QLT_HasTC, OITM.U_Matrice_disegno, OITM.U_Storico_prezzi,oitm.usersign, OITM.U_FINAL_CUSTOMER_CODE, OITM.U_FINAL_CUSTOMER_NAME, oitm.u_sector, oitm.u_country_of_delivery,oitm.u_agent,oitm.u_brand, OITM.U_INSERT_DATE, OITM.U_CODICE_BRB,OITM.U_UBICAZIONE_LABELLING)

        SELECT '" & par_codice_articolo & "', '" & par_descrizione_articolo & "', '" & par_desc_supp & "', " & par_gruppo_articoli & ", T0.CstGrpCode, T0.VatGourpSa, T0.CodeBars, T0.VATLiable, CASE WHEN SUBSTRING('" & par_codice_articolo & "',1,1)='B' THEN 'N' ELSE T0.PrchseItem END, T0.SellItem, '" & par_gestione_a_magazzino & "', T0.OnHand, T0.IsCommited, T0.OnOrder, T0.IncomeAcct, T0.ExmptIncom, T0.MaxLevel, T0.DfltWH, '" & par_fornitore_preferito & "','" & par_catalogo_fornitore & "' , T0.BuyUnitMsr, T0.NumInBuy, T0.ReorderQty, T0.MinLevel, T0.LstEvlPric, T0.LstEvlDate, T0.Canceled, T0.SalUnitMsr, T0.NumInSale, T0.Consig, T0.Counted, T0.EvalSystem, T0.PicturName, T0.UserText, T0.CommisPcnt, T0.CommisSum, T0.CommisGrp, T0.TreeType, T0.TreeQty, T0.LastPurPrc, T0.LastPurCur, T0.LastPurDat, T0.ExitPrice, T0.ExitWH, T0.AssetItem, T0.WasCounted, T0.ManSerNum, T0.SHeight1, T0.SHght1Unit, T0.SHeight2, T0.SHght2Unit, T0.SWidth1, T0.SWdth1Unit, T0.SWidth2, T0.SWdth2Unit, T0.SLength1, T0.SLen1Unit, T0.Slength2, T0.SLen2Unit, T0.SVolume, T0.SVolUnit, T0.SWeight1, T0.SWght1Unit, T0.SWeight2, T0.SWght2Unit, T0.BHeight1, T0.BHght1Unit, T0.BHeight2, T0.BHght2Unit, T0.BWidth1, T0.BWdth1Unit, T0.BWidth2, T0.BWdth2Unit, T0.BLength1, T0.BLen1Unit, T0.Blength2, T0.BLen2Unit, T0.BVolume, T0.BVolUnit, T0.BWeight1, T0.BWght1Unit, T0.BWeight2, T0.BWght2Unit, T0.FixCurrCms, '" & par_produttore & "', T0.LstSalDate, T0.QryGroup1, T0.QryGroup2, T0.QryGroup3, T0.QryGroup4, T0.QryGroup5, T0.QryGroup6, T0.QryGroup7, T0.QryGroup8, T0.QryGroup9, T0.QryGroup10, T0.QryGroup11, T0.QryGroup12, T0.QryGroup13, T0.QryGroup14, T0.QryGroup15, T0.QryGroup16, T0.QryGroup17, T0.QryGroup18, T0.QryGroup19, T0.QryGroup20, T0.QryGroup21, T0.QryGroup22, T0.QryGroup23, T0.QryGroup24, T0.QryGroup25, T0.QryGroup26, T0.QryGroup27, T0.QryGroup28, T0.QryGroup29, T0.QryGroup30, T0.QryGroup31, T0.QryGroup32, T0.QryGroup33, T0.QryGroup34, T0.QryGroup35, T0.QryGroup36, T0.QryGroup37, T0.QryGroup38, T0.QryGroup39, T0.QryGroup40, T0.QryGroup41, T0.QryGroup42, T0.QryGroup43, T0.QryGroup44, T0.QryGroup45, T0.QryGroup46, T0.QryGroup47, T0.QryGroup48, T0.QryGroup49, T0.QryGroup50, T0.QryGroup51, T0.QryGroup52, T0.QryGroup53, T0.QryGroup54, T0.QryGroup55, T0.QryGroup56, T0.QryGroup57, T0.QryGroup58, T0.QryGroup59, T0.QryGroup60, T0.QryGroup61, T0.QryGroup62, T0.QryGroup63, T0.QryGroup64, getdate(), getdate(), T0.SalFactor1, T0.SalFactor2, T0.SalFactor3, T0.SalFactor4, T0.PurFactor1, T0.PurFactor2, T0.PurFactor3, T0.PurFactor4, T0.SalFormula, T0.PurFormula, T0.VatGroupPu, T0.AvgPrice, T0.PurPackMsr, T0.PurPackUn, T0.SalPackMsr, T0.SalPackUn, T0.ManBtchNum, T0.ManOutOnly, T0.validFor, T0.validFrom, T0.validTo, T0.frozenFor, T0.frozenFrom, T0.frozenTo, T0.ValidComm, T0.FrozenComm, T0.SWW, T0.Deleted, T0.ExpensAcct, T0.FrgnInAcct, T0.ShipType, T0.GLMethod, T0.ECInAcct, T0.FrgnExpAcc, T0.ECExpAcc, T0.TaxType, T0.ByWh, T0.WTLiable, T0.ItemType, T0.WarrntTmpl, T0.BaseUnit, T0.CountryOrg, T0.StockValue, '" & par_phantom & "', T0.IssueMthd, T0.FREE1, T0.PricingPrc, T0.MngMethod, T0.ReorderPnt, T0.InvntryUom, T0.PlaningSys, '" & par_approvvigionamento & "', T0.OrdrIntrvl, T0.OrdrMulti, T0.MinOrdrQty, T0.LeadTime, T0.IndirctTax, T0.TaxCodeAR, T0.TaxCodeAP, T0.OSvcCode, T0.ISvcCode, T0.ServiceGrp, T0.NCMCode, T0.MatType, T0.MatGrp, T0.ProductSrc, T0.ServiceCtg, T0.ItemClass, T0.Excisable, T0.ChapterID, T0.NotifyASN, T0.ProAssNum, T0.AssblValue, T0.DNFEntry, T0.Spec, T0.TaxCtg, '" & trova_serie_oitm(par_codice_articolo) & "', substring('" & par_codice_articolo & "',2,5), T0.FuelCode, T0.BeverTblC, T0.BeverGrpC, T0.BeverTM, T0.Attachment, T0.AtcEntry, T0.ToleranDay, T0.UgpEntry, T0.PUoMEntry, T0.SUoMEntry, T0.IUoMEntry, T0.IssuePriBy, T0.AssetClass, T0.AssetGroup, T0.InventryNo, T0.Technician, T0.Employee, T0.Location, T0.StatAsset, T0.Cession, T0.DeacAftUL, T0.AsstStatus, T0.CapDate, T0.AcqDate, T0.RetDate, T0.GLPickMeth, T0.NoDiscount, T0.MgrByQty, T0.AssetRmk1, T0.AssetRmk2, T0.AssetAmnt1, T0.AssetAmnt2, T0.DeprGroup, T0.AssetSerNo, T0.CntUnitMsr, T0.NumInCnt, T0.INUoMEntry, T0.OneBOneRec, T0.RuleCode, T0.ScsCode, T0.SpProdType, T0.IWeight1, T0.IWght1Unit, T0.IWeight2, T0.IWght2Unit, T0.CompoWH, cast(replace(convert(varchar, getdate(), 108),':','') as integer) , cast(replace(convert(varchar, getdate(), 108),':','') as integer), T0.VirtAstItm, T0.SouVirAsst, T0.InCostRoll, T0.PrdStdCst, T0.EnAstSeri, T0.LinkRsc, T0.OnHldPert, T0.onHldLimt, T0.PriceUnit, T0.GSTRelevnt, T0.SACEntry, T0.GstTaxCtg, T0.AssVal4WTR, T0.ExcImpQUoM, T0.ExcFixAmnt, T0.ExcRate, T0.SOIExc, T0.TNVED, T0.Imported, T0.AutoBatch, T0.CstmActing, T0.StdItemId, T0.CommClass, T0.TaxCatCode, T0.DataVers, T0.NVECode, T0.CESTCode, T0.CtrSealQty, T0.LegalText, T0.U_UBIMAG, T0.U_TEMPOMED, '4.0', T0.U_SEZIONE, T0.U_PRG_TIR_MATERIALE, T0.U_DESCFR, T0.U_DESCING, T0.U_UTILIZZ, T0.U_ARTCES, T0.U_BA_IsFA, T0.U_BA_TypID, T0.U_BA_NumID, T0.U_BA_LVAFrom, T0.U_BA_LVA, T0.U_ULTVER, T0.U_ULTPREZ, T0.U_PrzPeso, T0.U_PRG_AZS_DesAggArt, T0.U_PRG_AZS_DesAgg2Art, T0.U_PRG_AZS_CMM, T0.U_PRG_AZS_ItmsGrp2Cod, T0.U_PRG_AZS_GestComm, T0.U_PRG_AZS_CreatedBy, '" & par_phantom & "', T0.U_PRG_CLV_Tipo_Lav, T0.U_PRG_CLV_Lav_EstAss, T0.U_PRG_CLV_Grezzo, T0.U_PRG_CLV_TipologiaLav, T0.U_PRG_CLV_ArtAssLav, T0.U_PRG_CLV_SemiLavComodo, T0.U_PRG_CLV_MagPrePro, T0.U_Anagrafica, T0.U_Ricambio, T0.U_Collaudo, T0.U_DocumentazioneTecnica, T0.U_CostoMateriale, T0.U_CostoPrimo, T0.U_Margine, T0.U_ListinoMinimo, T0.U_Famiglia, T0.U_ErroreWip, T0.U_Movimentazioni, T0.U_TraferitoODP, T0.U_Superlistino, T0.U_Superlistino_Vecchio, T0.U_Inventariato, 'IMPORTATO',  '" & par_codice_disegno & "', T0.U_Indice_di_revisione, T0.U_Tipo_macchina, T0.U_Numero_formati, T0.U_Gestione_magazzino, T0.U_Macchina_standard, '" & par_revisione & "', T0.U_PRG_TIR_RevProd, T0.U_PRG_TIR_RevProdDate, '" & par_tipo_montaggio & "', T0.U_PRG_QLT_HasTC, T0.U_Matrice_disegno, T0.U_Storico_prezzi,'" & par_utente_SAP & "','" & par_codice_BP & "','" & par_nome_bp & "','" & par_settore & "','" & par_paese & "','" & par_agente & "','" & par_brand & "',getdate(),'" & par_codice_BRB & "',substring('" & par_ubicazione & "',1,10)  FROM OITM T0 WHERE T0.[ItemCode] ='D89403'"

        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()
        nnm1(par_codice_articolo)
    End Sub



    Sub ITM1(par_codice_Articolo As String, par_costo As Integer)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1



        Cmd_SAP.CommandText = "INSERT INTO ITM1 (ITM1.ItemCode, ITM1.PriceList, ITM1.Price, ITM1.Currency, ITM1.Ovrwritten, ITM1.Factor, ITM1.AddPrice1, ITM1.Currency1, ITM1.AddPrice2, ITM1.Currency2, ITM1.Ovrwrite1, ITM1.Ovrwrite2, ITM1.BasePLNum, ITM1.UomEntry, ITM1.PriceType)


 SELECT '" & par_codice_Articolo & "', T0.PriceList, 0, T0.Currency, T0.Ovrwritten, T0.Factor, T0.AddPrice1, T0.Currency1, T0.AddPrice2, T0.Currency2, T0.Ovrwrite1, T0.Ovrwrite2, T0.BasePLNum, T0.UomEntry, T0.PriceType FROM ITM1 T0 WHERE T0.[ItemCode] ='020646'"


        Cmd_SAP.ExecuteNonQuery()



        Cmd_SAP.CommandText = "update itm1 set itm1.price=" & par_costo & " from itm1 where itm1.itemcode='" & par_codice_Articolo & "' and itm1.pricelist=2"

        Cmd_SAP.ExecuteNonQuery()


        Cmd_SAP.CommandText = "update t1
        set t1.price=t0.price*2.55

        from oitm t2 
        inner join itm1 t0 on t0.itemcode=t2.itemcode
        inner join itm1 t1 on t0.itemcode=t1.itemcode

        where t0.pricelist=2 and t1.pricelist=3  AND substring(t0.itemcode,1,1)='C' and  T0.ITEMCODE='" & par_codice_Articolo & "'

        update t1
        set t1.price=t0.price*3.7
        from oitm t2 
        inner join itm1 t0 on t0.itemcode=t2.itemcode
        inner join itm1 t1 on t0.itemcode=t1.itemcode
        where t0.pricelist=2 and t1.pricelist=3  AND (substring(t0.itemcode,1,1)='D' or substring(t0.itemcode,1,1)='0') and  T0.ITEMCODE='" & par_codice_Articolo & "'


        update t1
        set t1.price=t0.price*146/100
        from oitm t2 
        inner join itm1 t0 on t0.itemcode=t2.itemcode
        inner join itm1 t1 on t0.itemcode=t1.itemcode
        where t0.pricelist=3 and t1.pricelist=4  and   T0.ITEMCODE='" & par_codice_Articolo & "'


        update t1
        set t1.price=t0.price*162.23/100
        from oitm t2 
        inner join itm1 t0 on t0.itemcode=t2.itemcode
        inner join itm1 t1 on t0.itemcode=t1.itemcode
        where t0.pricelist=3 and t1.pricelist=5  and  T0.ITEMCODE='" & par_codice_Articolo & "'


        update t1
        set t1.price=t0.price*124.30/100
        from oitm t2 
        inner join itm1 t0 on t0.itemcode=t2.itemcode
        inner join itm1 t1 on t0.itemcode=t1.itemcode
        where t0.pricelist=3 and t1.pricelist=10 and  T0.ITEMCODE='" & par_codice_Articolo & "'"

        Cmd_SAP.ExecuteNonQuery()




        Cnn1.Close()

    End Sub

    Sub OITW(par_codice_articolo As String)
        Dim CNN1 As New SqlConnection
        CNN1.ConnectionString = Homepage.sap_tirelli
        CNN1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = CNN1
        Cmd_SAP.CommandText = "INSERT INTO OITW (OITW.ItemCode, OITW.WhsCode, OITW.OnHand, OITW.IsCommited, OITW.OnOrder, OITW.Counted, OITW.WasCounted, OITW.MinStock, OITW.MaxStock, OITW.MinOrder, OITW.AvgPrice, OITW.Locked, OITW.BalInvntAc, OITW.SaleCostAc, OITW.TransferAc, OITW.RevenuesAc, OITW.VarianceAc, OITW.DecreasAc, OITW.IncreasAc, OITW.ReturnAc, OITW.ExpensesAc, OITW.EURevenuAc, OITW.EUExpensAc, OITW.FrRevenuAc, OITW.FrExpensAc, OITW.ExmptIncom, OITW.PriceDifAc, OITW.ExchangeAc, OITW.BalanceAcc, OITW.PurchaseAc, OITW.PAReturnAc, OITW.PurchOfsAc, OITW.ShpdGdsAct, OITW.VatRevAct, OITW.StockValue, OITW.DecresGlAc, OITW.IncresGlAc, OITW.StokRvlAct, OITW.StkOffsAct, OITW.WipAcct, OITW.WipVarAcct, OITW.CostRvlAct, OITW.CstOffsAct, OITW.ExpClrAct, OITW.ExpOfstAct, OITW.Object, OITW.logInstanc, OITW.createDate, OITW.userSign2, OITW.updateDate, OITW.ARCMAct, OITW.ARCMFrnAct, OITW.ARCMEUAct, OITW.ARCMExpAct, OITW.APCMAct, OITW.APCMFrnAct, OITW.APCMEUAct, OITW.RevRetAct, OITW.NegStckAct, OITW.StkInTnAct, OITW.PurBalAct, OITW.WhICenAct, OITW.WhOCenAct, OITW.WipOffset, OITW.StockOffst, OITW.DftBinAbs, OITW.DftBinEnfd, OITW.Freezed, OITW.FreezeDoc, OITW.FreeChrgSA, OITW.FreeChrgPU, OITW.IndEscala, OITW.CNJPMan)

                                SELECT '" & par_codice_articolo & "', T0.WhsCode, 0, 0, 0, T0.Counted, T0.WasCounted, T0.MinStock, T0.MaxStock, T0.MinOrder, T0.AvgPrice, T0.Locked, T0.BalInvntAc, T0.SaleCostAc, T0.TransferAc, T0.RevenuesAc, T0.VarianceAc, T0.DecreasAc, T0.IncreasAc, T0.ReturnAc, T0.ExpensesAc, T0.EURevenuAc, T0.EUExpensAc, T0.FrRevenuAc, T0.FrExpensAc, T0.ExmptIncom, T0.PriceDifAc, T0.ExchangeAc, T0.BalanceAcc, T0.PurchaseAc, T0.PAReturnAc, T0.PurchOfsAc, T0.ShpdGdsAct, T0.VatRevAct, T0.StockValue, T0.DecresGlAc, T0.IncresGlAc, T0.StokRvlAct, T0.StkOffsAct, T0.WipAcct, T0.WipVarAcct, T0.CostRvlAct, T0.CstOffsAct, T0.ExpClrAct, T0.ExpOfstAct, T0.Object, T0.logInstanc, getdate(), T0.userSign2, T0.updateDate, T0.ARCMAct, T0.ARCMFrnAct, T0.ARCMEUAct, T0.ARCMExpAct, T0.APCMAct, T0.APCMFrnAct, T0.APCMEUAct, T0.RevRetAct, T0.NegStckAct, T0.StkInTnAct, T0.PurBalAct, T0.WhICenAct, T0.WhOCenAct, T0.WipOffset, T0.StockOffst, T0.DftBinAbs, T0.DftBinEnfd, T0.Freezed, T0.FreezeDoc, T0.FreeChrgSA, T0.FreeChrgPU, T0.IndEscala, T0.CNJPMan FROM OITW T0 WHERE T0.[ItemCode] ='020646'"

        Cmd_SAP.ExecuteNonQuery()
        CNN1.Close()

    End Sub

    Sub nnm1(par_codice_articolo As String)
        If trova_serie_oitm(par_codice_articolo) = "2209" Or trova_serie_oitm(par_codice_articolo) = "2210" Or trova_serie_oitm(par_codice_articolo) = "2211" Then

            Dim CNN1 As New SqlConnection
            CNN1.ConnectionString = Homepage.sap_tirelli
            CNN1.Open()
            Dim Cmd_SAP As New SqlCommand
            Cmd_SAP.Connection = CNN1




            Cmd_SAP.CommandText = "update t11 set  t11.nextnumber=t10.max+1
from
(
SELECT T10.MAX
FROM
(
SELECT SUBSTRING (T0.itemcode,1,1) AS 'PRIMA',

max(case when substring('" & par_codice_articolo & "',1,1)='0' then SUBSTRING (t1.itemcode,2,6) 
when substring('" & par_codice_articolo & "',1,1)='D' then 
cast(substring(t0.itemcode,2,6) as integer)
else substring(T0.itemcode,2,6) end) AS 'MAX'
FROM OITM t0
left JOIN OITM T1 ON SUBSTRING(T1.ItemCode, 1, 1) = '0' and SUBSTRING(T0.ItemCode, 2, 5)<'80000' and t0.itemtype='I' and t0.itemcode=t1.itemcode
WHERE T0.[createDate] >= (CONVERT(DATETIME, '20190501', 112) ) and SUBSTRING (T0.itemcode,1,1)<> '1' and SUBSTRING (T0.itemcode,1,1)<> 'R' and SUBSTRING (T0.itemcode,1,1)<> 'B' 
group by SUBSTRING (T0.itemcode,1,1)
)
AS T10
WHERE SUBSTRING('" & par_codice_articolo & "',1,1)=T10.PRIMA
)
as t10 inner join nnm1 t11 on t11.series='" & trova_serie_oitm(par_codice_articolo) & "' and t11.objectcode='4'
"

            Cmd_SAP.ExecuteNonQuery()

            CNN1.Close()

        End If
    End Sub

    Public Function trova_serie_oitm(par_codice_articolo As String)

        Dim serie As Integer
        Dim Cnn4 As New SqlConnection
        Cnn4.ConnectionString = Homepage.sap_tirelli
        Cnn4.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn4



        CMD_SAP_1.CommandText = " select case
when substring('" & par_codice_articolo & "',1,1)='C' then 2209
when substring('" & par_codice_articolo & "',1,1)='D' then 2210 
when substring('" & par_codice_articolo & "',1,1)='0' then 2211 
else 1330 end as 'Series' 
"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader

        If cmd_SAP_reader_1.Read() Then


            serie = cmd_SAP_reader_1("Series")


        End If
        cmd_SAP_reader_1.Close()
        Cnn4.Close()

        Return serie
    End Function

    Public Function verifica_codice_attivo(par_codice_articolo As String)



        Dim risposta As String = "N"
        Dim Cnn4 As New SqlConnection
        Cnn4.ConnectionString = Homepage.sap_tirelli
        Cnn4.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn4



        CMD_SAP_1.CommandText = " select case 
when t0.frozenfor='Y' and t0.frozenfrom<=getdate() and t0.frozento>=getdate() then 'N' 
when t0.validfor='Y' and t0.validfrom<=getdate() and t0.validto>=getdate() then 'Y' 
when t0.validfor='Y' and t0.validfrom>getdate() OR t0.validto<=getdate() then 'N' 
when t0.validfor='Y' and t0.validfrom is null then 'Y'
when t0.FROZENFOR='N' and t0.validfrom is null then 'Y'

else 'N' end AS 'ATTIVO',
t0.frozenfor,t0.validfor, t0.validfrom, t0.validto, t0.frozenfrom, t0.frozento
from oitm t0 where t0.itemcode='" & par_codice_articolo & "' 
"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader

        If cmd_SAP_reader_1.Read() Then


            risposta = cmd_SAP_reader_1("Attivo")


        End If
        cmd_SAP_reader_1.Close()
        Cnn4.Close()

        Return risposta
    End Function


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
             CMD_SAP.CommandText = "SELECT T0.[ItmsGrpCod] AS 'Gruppo', T0.[ItmsGrpNam] as 'Nome gruppo' 
FROM OITB T0 
 "
        Else
             CMD_SAP.CommandText = "SELECT '' AS 'Gruppo', '' as 'Nome gruppo' 

 "
        End If


        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Indice = Indice + 1
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
            ComboBox3.Items.Clear()

            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli

            Cnn.Open()

            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader


            CMD_SAP.Connection = Cnn
            CMD_SAP.CommandText = "Select t0.firmcode, t0.firmname
from tirellisrldb.dbo.omrc t0 
ORDER BY T0.FIRMNAME"

            cmd_SAP_reader = CMD_SAP.ExecuteReader

            Dim Indice As Integer = 0

            Indice = Indice + 1
            Do While cmd_SAP_reader.Read()

                ComboBox3.Items.Add(cmd_SAP_reader("firmname"))
                ComboBox4.Items.Add(cmd_SAP_reader("firmname"))
                Elenco_produttori(Indice) = cmd_SAP_reader("firmcode")
                Indice = Indice + 1

            Loop
            cmd_SAP_reader.Close()
            Cnn.Close()

    End Sub

    Sub inserimento_settori()


        ComboBox2.Items.Clear()

        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.Code, T0.Name 
FROM tirellisrldb.[dbo].[@SETTORI]  T0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer = 0

        Indice = Indice + 1
        Do While cmd_SAP_reader.Read()

            ComboBox2.Items.Add(cmd_SAP_reader("code"))

            Elenco_settori(Indice) = cmd_SAP_reader("Name")
            Indice = Indice + 1

        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub inserimento_paesi()

        ComboBox7.Items.Clear()


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.Code, T0.Name, T0.U_NumCode 
FROM tirellisrldb.[dbo].[@BNCCRY]  T0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer = 0

        Indice = Indice + 1
        Do While cmd_SAP_reader.Read()

            ComboBox7.Items.Add(cmd_SAP_reader("code"))

            Elenco_paesi(Indice) = cmd_SAP_reader("Name")
            Indice = Indice + 1

        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub inserimento_FORNITORI()
        If Homepage.ERP_provenienza = "SAP" Then


            ComboBox5.Items.Clear()
            ComboBox6.Items.Clear()

            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli

            Cnn.Open()

            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader


            CMD_SAP.Connection = Cnn
            If Homepage.ERP_provenienza = "SAP" Then
                CMD_SAP.CommandText = "Select t0.CARDCODE, t0.CARDname
from OCRD t0 
WHERE T0.CARDTYPE='S' 
and t0.validfor='Y' 
ORDER BY T0.CARDNAME "
            Else

                CMD_SAP.CommandText = "select 
T10.CONTO AS 'Cardcode'
,t10.ds_conto as 'Cardname'


from openquery(AS400,'
SELECT *
FROM
S786FAD1.TIR90VIS.JGALACF
where clifor=''F''

order by ds_conto
'
)
as t10
"
            End If
            cmd_SAP_reader = CMD_SAP.ExecuteReader

            Dim Indice As Integer = 0


            Do While cmd_SAP_reader.Read()

                ComboBox5.Items.Add(cmd_SAP_reader("CARDname"))
                ComboBox6.Items.Add(cmd_SAP_reader("CARDname"))
                Elenco_fornitori(Indice) = cmd_SAP_reader("CARDcode")
                Indice = Indice + 1
            Loop
            cmd_SAP_reader.Close()
            Cnn.Close()
        End If
    End Sub



    Public Function Dammi_codice(par_prima_lettera As String)

        Dim codice As String
        Dim Cnn6 As New SqlConnection
        Cnn6.ConnectionString = Homepage.sap_tirelli
        Cnn6.Open()
        Dim Cmd_SAP As New SqlCommand
        Dim Cmd_SAP_Reader As SqlDataReader

        Cmd_SAP.Connection = Cnn6
        Cmd_SAP.CommandText = "SELECT TOP 1 
    CASE 
WHEN LEN(CONCAT(t10.prima_lettera, t10.numero)) = 2 THEN CONCAT(t10.prima_lettera, '0000', t10.numero)
WHEN LEN(CONCAT(t10.prima_lettera, t10.numero)) = 3 THEN CONCAT(t10.prima_lettera, '000', t10.numero)
WHEN LEN(CONCAT(t10.prima_lettera, t10.numero)) = 4 THEN CONCAT(t10.prima_lettera, '00', t10.numero)
        WHEN LEN(CONCAT(t10.prima_lettera, t10.numero)) = 5 THEN CONCAT(t10.prima_lettera, '0', t10.numero) 
        ELSE CONCAT(t10.prima_lettera, t10.numero) 
    END AS 'Ultimo_codice'
FROM
(
    SELECT 
        t0.numero, 
        COUNT(t1.itemcode) AS 'N', 
        '" & par_prima_lettera & "' AS 'Prima_lettera'
    FROM [Tirelli_40].[dbo].[Numeri_int_1000] t0 
    LEFT JOIN oitm t1 
        ON t0.numero = SUBSTRING(t1.itemcode, 2, 6) 
        AND t1.ItemType = 'I' 
        AND SUBSTRING(t1.itemcode, 1, 1) = '" & par_prima_lettera & "'
    GROUP BY t0.numero, SUBSTRING(t1.itemcode, 1, 1)
) AS t10
WHERE t10.N = 0
ORDER BY t10.numero"
        Cmd_SAP_Reader = Cmd_SAP.ExecuteReader

        If Cmd_SAP_Reader.Read = True Then

            codice = Cmd_SAP_Reader("Ultimo_codice")


        End If

        Cmd_SAP_Reader.Close()
        Cnn6.Close()

        Return codice
    End Function



    Private Sub Button3_Click(sender As Object, e As EventArgs)

        Me.Close()
    End Sub

    Sub check_codice(par_utente_sap As String, par_codice_sap As String, par_descrizione_codice As String, par_desc_supp As String, par_disegno As String, par_codice_brb As String, par_approvvigionamento As String)

        par_codice_brb = Replace(par_codice_brb, ",", " ")
        par_descrizione_codice = Replace(par_descrizione_codice, ",", " ")
        par_desc_supp = Replace(par_desc_supp, ",", " ")
        par_disegno = Replace(par_disegno, ",", " ")
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn
        If par_codice_brb = "" Then
            CMD_SAP.CommandText = "select t0.itemcode 
from oitm t0 where t0.itemcode='" & par_codice_sap & "'"
        Else
            CMD_SAP.CommandText = "select t0.itemcode 
from oitm t0 
where t0.itemcode='" & par_codice_sap & "' or t0.u_codice_brb ='" & par_codice_brb & "' or (coalesce(t0.u_disegno,'') ='" & par_disegno & "' and coalesce(t0.u_disegno,'')<>'') "
        End If


        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then
            Cnn.Close()
            MsgBox("Probabilmente il codice è già stato preso, riassegnare la prima lettera")
        Else
            Cnn.Close()
            If ComboBox5.SelectedIndex < 0 Then
                ComboBox5.SelectedIndex = 0
                Elenco_fornitori(ComboBox5.SelectedIndex) = ""

            End If

            If ComboBox3.SelectedIndex < 0 Then
                ComboBox3.SelectedIndex = 0
                Elenco_produttori(ComboBox3.SelectedIndex) = "-1"
            End If


            If ComboBox2.SelectedIndex = -1 Then
                SETTORE = ""

            Else
                SETTORE = Elenco_settori(ComboBox2.SelectedIndex)
            End If


            If ComboBox7.SelectedIndex = -1 Then
                paese = ""

            Else
                paese = Elenco_paesi(ComboBox7.SelectedIndex)
            End If


            If ComboBox8.SelectedIndex = -1 Then
                brand = ""

            Else
                brand = ComboBox8.Text
            End If
            inserisci(par_utente_sap, par_codice_sap, par_descrizione_codice, par_desc_supp, par_disegno, Elenco_gruppi(ComboBox_gruppo_articoli.SelectedIndex), Elenco_fornitori(ComboBox5.SelectedIndex), Replace(TextBox1.Text, ",", " "), Elenco_produttori(ComboBox3.SelectedIndex + 1), ComboBox_tipo_montaggio.Text, Replace(TextBox10.Text, ",", " "), Replace(TextBox5.Text, ",", " "), SETTORE, paese, Replace(TextBox8.Text, ",", " "), brand, par_codice_brb, ComboBox10.Text, 0, "", par_approvvigionamento)
            MsgBox("Codice inserito con successo")

        End If

        cmd_SAP_reader.Close()

    End Sub

    Public Function check_non_duplicazione_codici(par_codice_sap As String, par_disegno As String, par_codice_brb As String)
        Dim risultato As String = "N"
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "select t0.itemcode 
from oitm t0 
where t0.itemcode='" & par_codice_sap & "' or (coalesce(t0.u_codice_brb,'')='" & par_codice_brb & "' And coalesce(t0.u_codice_brb,'')<>'') or (coalesce(t0.u_disegno,'')='" & par_disegno & "' And coalesce(t0.u_disegno,'')<>'')"


        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then
            risultato = "Y"

        End If


        cmd_SAP_reader.Close()
        Cnn.Close()
        Return risultato
    End Function




    Private Sub Button1_Click(sender As Object, e As EventArgs)
        If Combodipendenti.Text = Nothing Then
            MsgBox("Utenticarsi prima di aggiungere un codice")
        Else
            If TextBox_codice_sap.Text <> "" And ComboBox_gruppo_articoli.Text <> "" Then

                check_codice(Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, TextBox_codice_sap.Text, TextBox_descrizione.Text, TextBox_DESC_SUPP.Text, TextBox_disegno.Text, TextBox12.Text, ComboBox11.Text)
                inserimento_risorsa()

                TextBox_codice_sap.Text = Nothing
                TextBox_disegno.Text = Nothing
                TextBox_descrizione.Text = Nothing
                TextBox_DESC_SUPP.Text = Nothing
                ComboBox_gruppo_articoli.Text = Nothing
                ComboBox_tipo_montaggio.Text = Nothing
                ComboBox3.Text = Nothing
                TextBox1.Text = Nothing
                ComboBox5.Text = Nothing

            Else

                MsgBox("Mancano dei campi fondamentali")

            End If
        End If
    End Sub

    Sub inserimento_risorsa()
        Dim CNN1 As New SqlConnection
        CNN1.ConnectionString = Homepage.sap_tirelli
        CNN1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = CNN1
        Cmd_SAP.CommandText = "Insert into oitw (itemcode,whscode) select t0.itemcode,'RIS' FROM OITM T0 WHERE  t0.itemcode='" & TextBox_codice_sap.Text & "'"

        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = "UPDATE T0 SET T0.DFLTWH ='RIS' FROM OITM T0 WHERE t0.itemcode='" & TextBox_codice_sap.Text & "' "

        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = "UPDATE T0 SET T0.OBJTYPE='290' FROM OITM T0 WHERE t0.itemcode='" & TextBox_codice_sap.Text & "' "

        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = "Update t0 set t0.linkrsc=t0.itemcode FROM OITM T0 WHERE t0.itemcode='" & TextBox_codice_sap.Text & "'"

        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = "update t0 set  t0.linkitm =t1.itemcode
from orsc t0 inner join oitm t1 on t0.visrescode=t1.itemcode where t1.itemcode='" & TextBox_codice_sap.Text & "'"

        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = "update t0 set  t0.invntitem='N'
from oitm t0 where t0.itemcode='" & TextBox_codice_sap.Text & "'"

        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = "UPDATE T0 SET T0.[ProductSrc]='0' from oitm t0 where t0.itemcode='" & TextBox_codice_sap.Text & "'"

        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = "UPDATE T0 SET T0.PrcrmntMtd='B' from oitm t0 where t0.itemcode='" & TextBox_codice_sap.Text & "'"

        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = "UPDATE T0 SET T0.EVALSYSTEM='S' from oitm t0 where t0.itemcode='" & TextBox_codice_sap.Text & "'"

        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = "UPDATE t1 set t1.price=1  from oitm t0 inner join itm1 t1 on t0.itemcode=t1.itemcode where t0.itemcode='" & TextBox_codice_sap.Text & "'"

        Cmd_SAP.ExecuteNonQuery()

        CNN1.Close()

    End Sub










    Private Sub UT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        inizializzazione_ut()
    End Sub





    Private Sub BRB_Click(sender As Object, e As EventArgs) Handles BRB.Enter

        visualizzazione = "BRB"
    End Sub

    Private Sub TIRELLI_Click(sender As Object, e As EventArgs) Handles Tirelli.Enter

        visualizzazione = "TIRELLI"
    End Sub



    Sub trova_dato_da_excel_pEr_importazionE(par_percorso_file As String, par_nome_foglio As String, par_riga_inizio As Integer, par_riga_fine As Integer)

        Dim itemcode As String
        Dim Descrizione As String
        Dim desc_supp As String
        Dim ubicazione As String

        Dim costo As String
        Dim codice_fornitore As String
        Dim nome_fornitore As String


        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True


        Do While par_riga_inizio <= par_riga_fine


            If Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value <> Nothing Then
                itemcode = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value
                Descrizione = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value
                desc_supp = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 3).value
                ubicazione = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value
                costo = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 6).value
                codice_fornitore = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 7).value
                nome_fornitore = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 8).value

                insert_into_codici_BRB(itemcode, Descrizione, desc_supp, "", codice_fornitore, nome_fornitore, ubicazione, costo)

            End If
            par_riga_inizio = par_riga_inizio + 1
        Loop


    End Sub


    Sub insert_into_codici_BRB(par_Codice_BRB As String, par_Descrizione_BRB As String, par_Descrizione_supp_BRB As String, par_Catalogo_fornitore As String, par_codice_fornitore As String, par_Fornitore As String, par_ubicazione As String, par_costo As String)

        par_Descrizione_BRB = Replace(par_Descrizione_BRB, "'", " ")
        par_Descrizione_supp_BRB = Replace(par_Descrizione_supp_BRB, "'", " ")
        par_Catalogo_fornitore = Replace(par_Catalogo_fornitore, "'", " ")
        par_Fornitore = Replace(par_Fornitore, "'", " ")
        par_costo = Replace(par_costo, ",", ".")
        par_ubicazione = Replace(par_ubicazione, "'", " ")

        Dim CNN1 As New SqlConnection
        CNN1.ConnectionString = Homepage.sap_tirelli
        CNN1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = CNN1
        Cmd_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].BRB_Codici
           ([Codice_BRB]
           ,[Descrizione_BRB]
           ,[Descrizione_supp_BRB]
           ,[Catalogo_fornitore]
,[codice_fornitore]
           ,[Fornitore]
           ,[Ubicazione]
           ,[Costo]
           ,[data_import])
     VALUES
           ('" & par_Codice_BRB & "'
           ,'" & par_Descrizione_BRB & "'
           ,'" & par_Descrizione_supp_BRB & "'
           ,'" & par_Catalogo_fornitore & "'
           ,'" & par_codice_fornitore & "'
           ,'" & par_Fornitore & "'
           ,'" & par_ubicazione & "'
           ,'" & par_costo & "'
,getdate())"

        Cmd_SAP.ExecuteNonQuery()
        CNN1.Close()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        trova_dato_da_excel_pEr_importazionE("C:\Users\giovannitirelli\Desktop\Business partner.xlsx", "Pulito", 25165, 31710)
    End Sub



    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs)

    End Sub



    Private Sub Button20_Click(sender As Object, e As EventArgs)
        Me.WindowState = FormWindowState.Minimized
    End Sub



    Private Sub Aggiungi_Click(sender As Object, e As EventArgs) Handles Aggiungi.Click
        If Combodipendenti.Text = Nothing Then
            MsgBox("Utenticarsi in basso prima di aggiungere un codice")

        Else

            If Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = 0 Then
                MsgBox("Non è associata alcuna licenza SAP all'utente selezionato")

            Else



                If TextBox_codice_sap.Text <> "" And ComboBox_gruppo_articoli.Text <> "" Then


                    If ComboBox_prima_lettera.Text = "W" Then


                        If TextBox10.Text = "" Then
                            MsgBox("Scegliere un codice business partner")
                        ElseIf TextBox5.Text = "" Then
                            MsgBox("Scegliere un nome business partner")


                        Else
                            check_codice(Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, TextBox_codice_sap.Text, TextBox_descrizione.Text, TextBox_DESC_SUPP.Text, TextBox_disegno.Text, TextBox12.Text, ComboBox11.Text)

                        End If

                    ElseIf ComboBox_prima_lettera.Text = "M" Then

                        If TextBox10.Text = "" Then
                            MsgBox("Scegliere un codice business partner")
                        ElseIf TextBox5.Text = "" Then
                            MsgBox("Scegliere un nome business partner")
                        ElseIf ComboBox2.SelectedIndex = -1 Then
                            MsgBox("Scegliere un settore del cliente")
                        ElseIf ComboBox7.SelectedIndex = -1 Then
                            MsgBox("Scegliere un paese di destinazione del cliente")



                        ElseIf ComboBox8.SelectedIndex = -1 Then
                            MsgBox("Scegliere un brand della macchina")

                        Else
                            check_codice(Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, TextBox_codice_sap.Text, TextBox_descrizione.Text, TextBox_DESC_SUPP.Text, TextBox_disegno.Text, TextBox12.Text, ComboBox11.Text)

                        End If
                    Else




                        check_codice(Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, TextBox_codice_sap.Text, TextBox_descrizione.Text, TextBox_DESC_SUPP.Text, TextBox_disegno.Text, TextBox12.Text, ComboBox11.Text)


                        TextBox_codice_sap.Text = Nothing
                        TextBox_disegno.Text = Nothing
                        TextBox_descrizione.Text = Nothing
                        TextBox_DESC_SUPP.Text = Nothing
                        ComboBox_gruppo_articoli.Text = Nothing
                        ComboBox_tipo_montaggio.Text = Nothing
                        ComboBox3.Text = Nothing
                        TextBox1.Text = Nothing
                        ComboBox5.Text = Nothing
                        TextBox12.Text = ""
                    End If
                Else

                    MsgBox("Mancano dei campi fondamentali")
                End If
            End If
        End If
    End Sub



    Private Sub Button_CERCA_Click_1(sender As Object, e As EventArgs) Handles Button_CERCA.Click
        If visualizzazione = "TIRELLI" Then
            cerca(TextBox_codice_SAP_RICERCA.Text.ToUpper, TextBox_Descrizione_ricerca.Text.ToUpper, TextBox3.Text.ToUpper, TextBox_disegno_SAP.Text.ToUpper, TextBox6.Text.ToUpper, ComboBox1.Text.ToUpper, ComboBox4.Text.ToUpper, TextBox2.Text.ToUpper, ComboBox6.Text.ToUpper, TextBox4.Text.ToUpper)
        Else
            cerca_BRB()
        End If
    End Sub

    Private Sub ComboBox_prima_lettera_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_prima_lettera.SelectedIndexChanged

        TextBox_codice_sap.Text = Dammi_codice(ComboBox_prima_lettera.Text)
        If ComboBox_prima_lettera.Text = "W" Then

            GroupBox5.Visible = False
            GroupBox9.Visible = False
            GroupBox8.Visible = False
            GroupBox11.Visible = False
            GroupBox17.Visible = False
            GroupBox18.Visible = False
            GroupBox19.Visible = False
            GroupBox20.Visible = False
            Group_box_disegno.Visible = False


        ElseIf ComboBox_prima_lettera.Text = "M" Then


            GroupBox5.Visible = False
            GroupBox9.Visible = False
            GroupBox8.Visible = False
            GroupBox11.Visible = False
            Group_box_disegno.Visible = False

        Else
            Panel33.Visible = False
        End If
    End Sub

    Private Sub Combodipendenti_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles Combodipendenti.SelectedIndexChanged
        Homepage.ID_SALVATO = Elenco_codici_dipendenti(Combodipendenti.SelectedIndex)
        ' Homepage.UTENTE_NOME_SALVATO = Combodipendenti.Text


        Homepage.Aggiorna_INI_COMPUTER()
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Business_partner.Provenienza = "UT"
        Business_partner.Show()
    End Sub

    Private Sub TextBox_disegno_SAP_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox_disegno_SAP.TextChanged
        If TextBox_disegno_SAP.Text = "" Or TextBox_disegno_SAP.Text = Nothing Then
            filtro_disegno = ""
        Else


            Dim separators() As Char = "*"
            Dim searchTerms As String() = TextBox_disegno_SAP.Text.Split(separators, StringSplitOptions.RemoveEmptyEntries)
            Dim conditions As New List(Of String)

            'Console.WriteLine(searchTerms)


            For Each term As String In searchTerms
                If Not String.IsNullOrWhiteSpace(term) Then
                    conditions.Add(" and T0.[u_disegno] LIKE '%" & term.Trim() & "%'")
                End If
            Next

            If conditions.Count > 0 Then
                Dim query As String = "" & String.Join("", conditions) & ""
                filtro_disegno = query

            Else

                filtro_disegno = " And T0.[u_disegno] Like '" & TextBox_disegno_SAP.Text & " %%' "



            End If


        End If
    End Sub

    Private Sub TextBox6_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If TextBox6.Text = "" Then
            filtro_fam_disegno = ""
        Else
            filtro_fam_disegno = " and t0.u_famiglia_disegno Like '%%" & TextBox6.Text & "%%'"
        End If
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




            End If
        End If
    End Sub

    Private Sub DataGridView_SAP_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_SAP.CellFormatting
        If DataGridView_SAP.Rows(e.RowIndex).Cells(columnName:="Attivo").Value = "N" Then


            DataGridView_SAP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightSlateGray
            '    ' Applica il font barrato
            '    Dim cellStyle As DataGridViewCellStyle = DataGridView_SAP.Rows(e.RowIndex).DefaultCellStyle
            '    Try
            '        cellStyle.Font = New Font(cellStyle.Font, FontStyle.Strikeout)
            '    Catch ex As Exception

            '    End Try


        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If ComboBox9.SelectedIndex < 0 Then
            MsgBox("Selezionare un gruppo articoli")

        Else
            If TextBox7.Text = "" Or TextBox11.Text = "" Then
                MsgBox("Selezionare un codice BRB")
            Else

                If ComboBox9.Text = "Articolo generico" Then
                    MsgBox("Selezionare un gruppo articolo diverso da articolo generico")

                Else
                    Dim nuovo_Codice_tirelli As String
                    nuovo_Codice_tirelli = Dammi_codice(TextBox11.Text)
                    inserisci_Nuovo_codice(Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, nuovo_Codice_tirelli, DataGridView1.Rows(riga).Cells(columnName:="Descrizione_").Value, DataGridView1.Rows(riga).Cells(columnName:="Desc_sup_").Value, "", Elenco_gruppi(ComboBox9.SelectedIndex), "", "", "-1", "", "", "", "", "", "", "", DataGridView1.Rows(riga).Cells(columnName:="Codice_").Value, 0, ComboBox10.Text, DataGridView1.Rows(riga).Cells(columnName:="costo_").Value, ubicazione_labelling, ComboBox11.Text)
                    MsgBox("Codice " & nuovo_Codice_tirelli & " creato con successo")
                    If visualizzazione = "TIRELLI" Then
                        cerca(TextBox_codice_SAP_RICERCA.Text, TextBox_Descrizione_ricerca.Text, TextBox3.Text, TextBox_disegno_SAP.Text, TextBox6.Text, ComboBox1.Text, ComboBox4.Text, TextBox2.Text, ComboBox6.Text, TextBox4.Text)
                    Else
                        cerca_BRB()
                    End If
                End If
            End If
        End If
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            TextBox7.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice_").Value

            codice_BRB = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice_").Value
            ComboBox9.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="gruppo_nome").Value
            TextBox11.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Prima_lettera_Tir").Value

            ubicazione_labelling = DataGridView1.Rows(e.RowIndex).Cells(columnName:="DataGridViewTextBoxColumn8").Value
            riga = e.RowIndex
            If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice_tirelli").Value = "" Then
                Button4.Visible = True
            Else
                Button4.Visible = False
            End If

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



            If TextBox_Descrizione_ricerca.Text = Nothing Or TextBox_Descrizione_ricerca.Text = "" Then
                filtro_descrizione = ""
            Else


                Dim separators() As Char = "*"
                Dim searchTerms As String() = TextBox_Descrizione_ricerca.Text.Split(separators, StringSplitOptions.RemoveEmptyEntries)
                Dim conditions As New List(Of String)

                'Console.WriteLine(searchTerms)

                For Each term As String In searchTerms
                    If Not String.IsNullOrWhiteSpace(term) Then
                        conditions.Add(" and t0.itemname LIKE '%" & term.Trim() & "%'")
                    End If
                Next

                If conditions.Count > 0 Then
                    Dim query As String = "" & String.Join("", conditions) & ""
                    filtro_descrizione = query
                    Console.WriteLine(filtro_descrizione)
                Else

                    filtro_descrizione = " And t0.itemname Like '" & TextBox_Descrizione_ricerca.Text & " %%' "


                    Console.WriteLine(filtro_descrizione)
                End If




            End If
        Else


            Dim separators() As Char = "*"
            Dim searchTerms As String() = TextBox_Descrizione_ricerca.Text.Split(separators, StringSplitOptions.RemoveEmptyEntries)
            Dim conditions As New List(Of String)

            'Console.WriteLine(searchTerms)

            For Each term As String In searchTerms
                If Not String.IsNullOrWhiteSpace(term) Then
                    conditions.Add(" and (t10.descrizione_brb LIKE '%" & term.Trim() & "%' or t10.itemname LIKE '%" & term.Trim() & "%')")
                End If
            Next

            If conditions.Count > 0 Then
                Dim query As String = "" & String.Join("", conditions) & ""
                filtro_descrizione = query
                Console.WriteLine(filtro_descrizione)
            Else

                filtro_descrizione = " And (t10.DESCRIZIONE_BRB Like '" & TextBox_Descrizione_ricerca.Text & " %%' or t10.itemname LIKE '%" & TextBox_Descrizione_ricerca.Text & "%') "


                Console.WriteLine(filtro_descrizione)
            End If



        End If



    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Or TextBox3.Text = Nothing Then
            filtro_descrizione_supp = ""
        Else


            Dim separators() As Char = "*"
            Dim searchTerms As String() = TextBox3.Text.Split(separators, StringSplitOptions.RemoveEmptyEntries)
            Dim conditions As New List(Of String)

            'Console.WriteLine(searchTerms)

            If visualizzazione = "TIRELLI" Then
                For Each term As String In searchTerms
                    If Not String.IsNullOrWhiteSpace(term) Then
                        conditions.Add(" and T0.[FrgnName] LIKE '%" & term.Trim() & "%'")
                    End If
                Next

                If conditions.Count > 0 Then
                    Dim query As String = "" & String.Join("", conditions) & ""
                    filtro_descrizione_supp = query

                Else

                    filtro_descrizione_supp = " And T0.[FrgnName] Like '" & TextBox3.Text & " %%' "



                End If
            Else
                For Each term As String In searchTerms
                    If Not String.IsNullOrWhiteSpace(term) Then
                        conditions.Add(" and (t10.descrizione_supp LIKE '%" & term.Trim() & "%' or t10.frgnname LIKE '%" & term.Trim() & "%') ")
                    End If
                Next

                If conditions.Count > 0 Then
                    Dim query As String = "" & String.Join("", conditions) & ""
                    filtro_descrizione_supp = query

                Else

                    filtro_descrizione_supp = " And (t10.descrizione_supp Like '" & TextBox3.Text & " %%' or t10.frgnname LIKE '%" & TextBox3.Text & "%')  "


                End If
            End If
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button20_Click_1(sender As Object, e As EventArgs) Handles Button20.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = "" Or TextBox2.Text = Nothing Then
            filtro_catalogo = ""
        Else


            Dim separators() As Char = "*"
            Dim searchTerms As String() = TextBox2.Text.Split(separators, StringSplitOptions.RemoveEmptyEntries)
            Dim conditions As New List(Of String)

            'Console.WriteLine(searchTerms)


            For Each term As String In searchTerms
                If Not String.IsNullOrWhiteSpace(term) Then
                    conditions.Add(" and T0.[suppcatnum] LIKE '%" & term.Trim() & "%'")
                End If
            Next

            If conditions.Count > 0 Then
                Dim query As String = "" & String.Join("", conditions) & ""
                filtro_catalogo = query

            Else

                filtro_catalogo = " And T0.[suppcatnum] Like '" & TextBox2.Text & " %%' "

            End If


        End If
    End Sub

    Private Sub DataGridView_SAP_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_SAP.CellClick

    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged

    End Sub

    Private Sub TextBox_codice_SAP_RICERCA_TextChanged(sender As Object, e As EventArgs) Handles TextBox_codice_SAP_RICERCA.TextChanged

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub
End Class