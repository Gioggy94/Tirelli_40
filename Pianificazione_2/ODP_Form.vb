
Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop
Imports Word = Microsoft.Office.Interop.Word





Public Class ODP_Form


    Public docnum_odp As String
    Public docentry_odp As Integer
    Public Elenco_stati_odp(1000) As String
    Public Elenco_produzione(1000) As String
    Public Elenco_fase(1000) As String
    Public elenco_stato_lav(1000) As String
    Public elenco_aggiorna_db(1000) As String
    Public Righe_cancellate(100) As Riga_cancellata
    Public num_righe_cancellate As Integer = 0
    Public stampa_etichetta As String
    Public percorso_documento As String



    Public testata_odp_itemcode As String
    Public testata_odp_prodname As String
    Public testata_odp_u_disegno As String
    Public testata_odp_data As String
    Public testata_odp_plannedqty As Integer
    Public testata_odp_resname As String
    Public testata_odp_u_produzione As String
    Public testata_odp_commessa As String
    Public testata_odp_max_righe As String
    Public testata_odp_docnum As String
    Public testata_odp_cardname As String
    Public testata_odp_U_Clientefinale As String
    Public testata_odp_docnum_oc As String
    Public testata_odp_cliente_eti As String
    Public testata_odp_Itemname_commessa As String
    Public testata_odp_warehouse As String
    Public magazzino_riga As String
    Public numerone As String

    Public oWord As Word.Application
    Public oDoc As Word.Document
    Public oTable As Word.Table
    Public visorder_selezionato As Integer
    Public DATAGRIDVIEW_odp_RIGA As Integer
    Public datagridview_odp_colonna As Integer
    Public itemcode_riga As String


    Public quantità_base_riga As String
    Public attrezzaggio_riga As String
    Public riga_selezionata As Integer
    Public contatore As Integer = 0

    Public max_linenum As Integer = 0
    Public max_visorder As Integer = 0

    Public c As Integer = 0
    Public id_modifica As Integer
    Private variabile_esistenza_odp_integrato As String

    Public Sel_Stampante As New PrintDialog

    Public Preview As New PrintPreviewDialog

    Public altezza_scontrino_odp As Integer = 700
    Public larghezza_scontrino_odp As Integer = 185
    Public altezza_scontrino_odp_labelling As Integer = 222

    Public preview_scontrino As Boolean = False
    '  Private operazione As String = "<>"

    Public codici_salvati(1000) As String
    Public contatore_codici As Integer = 0
    Public inizio As Boolean = True


    Public Structure Riga_cancellata
        Public Codice_riga As String
        Public linenum As Integer
        Public visorder As Integer

    End Structure



    Sub inizializza_form(par_docnum_odp As String)
        inizio = True
        Dim stopWatch As New Stopwatch()
        Dim ts As TimeSpan = stopWatch.Elapsed

        num_righe_cancellate = 0
        c = 0

        stopWatch.Start()



        formatta_form(par_docnum_odp)
        stopWatch.Stop()
        ts = stopWatch.Elapsed
        Console.WriteLine("sub informazioni_testata(): " & ts.TotalMilliseconds & " ms")
        stopWatch.Restart()
        riempi_datagridview_ODP(DataGridView_ODP, par_docnum_odp, TextBox6.Text)
        'new_riempi_datagridview_ODP(DataGridView_ODP, par_docnum_odp, TextBox6.Text, operazione, Homepage.Centro_di_costo)
        stopWatch.Stop()
        ts = stopWatch.Elapsed
        Console.WriteLine("sub riempi_datagridview_ODP(): " & ts.TotalMilliseconds & " ms")
        stopWatch.Restart()
        'esistenza_odp_integrato(par_docnum_odp)
        'If variabile_esistenza_odp_integrato = "Y" Then
        '    riempi_datagridview_ODP_integrato(par_docnum_odp)
        'End If

        stopWatch.Stop()
        ts = stopWatch.Elapsed
        Console.WriteLine("sub riempi_datagridview_ODP_integrato(): " & ts.TotalMilliseconds & " ms")
        stopWatch.Restart()
        TROVA_MAX_LINENUM_E_MAX_VISORDER(par_docnum_odp)
        stopWatch.Stop()
        ts = stopWatch.Elapsed
        Console.WriteLine("sub TROVA_MAX_LINENUM_E_MAX_VISORDER(): " & ts.TotalMilliseconds & " ms")
        stopWatch.Restart()
        controlli()
        stopWatch.Stop()
        ts = stopWatch.Elapsed
        Console.WriteLine("controlli(): " & ts.TotalMilliseconds & " ms")
        inizio = False
        Dashboard_MU_New.righe_ODP_macchine(DataGridView1, TextBox10.Text)

    End Sub

    Public Function ottieni_informazioni_odp(PAR_TIPO As String, PAR_DOCENTRY_ODP As Integer, par_docnum_odp As String)

        Dim filtro_odp As String

        If PAR_TIPO = "Numero" Then
            filtro_odp = " and T0.[DocNum]='" & par_docnum_odp & "'"
        ElseIf PAR_TIPO = "Docentry" Then
            filtro_odp = " and T0.[docentry]='" & PAR_DOCENTRY_ODP & "'"
        End If

        Dim dettagli As New Dettagliodp()

        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then


            CMD_SAP_2.CommandText = "SELECT t0.docentry, T0.[DocNum] AS 'docnum', T0.PLANNEDQTY
, case when t0.u_stato is null then '' else t0.u_stato end as 'u_stato',
t0.status AS 'Stato', t0.u_lavorazione as 'lavorazione', T0.[ItemCode] as 'Itemcode', substring(T1.[ItemName],1,32) as 'Itemname'
, case when T1.[U_Disegno] is null then '' else t1.u_disegno end as 'Disegno'
, case when T0.[U_PRG_AZS_Commessa] is null then '' else T0.[U_PRG_AZS_Commessa] end as 'Commessa'
,coalesce(t7.u_final_customer_name,coalesce(t0.u_utilizz,'')) as 'Cliente'
, case when T6.NAME is null then '' else T6.NAME end as 'Fase' 
, coalesce(T2.[ItmsGrpNam],'') as 'Gruppo articolo', case when T0.U_AGGIORNA_DB is null then '' else t0.u_aggiorna_db end as 'u_aggiorna_db', T0.POSTDATE, T0.startDATE, T0.DUEDATE, concat(t4.lastname,' ',t4.firstname) as 'Nome' ,t0.type, t0.warehouse, case when t0.u_utilizz is null then '' else t0.u_utilizz end as 'u_utilizz'
, coalesce(t5.DESCR,'') as 'Descr'
, CASE WHEN T0.U_PROGRESSIVO_COMMESSA IS NULL THEN '' ELSE T0.U_PROGRESSIVO_COMMESSA END AS 'U_PROGRESSIVO_COMMESSA'
,coalesce(t0.u_produzione,'') as 'u_produzione'
,coalesce(t0.cmpltqty,0) as 'Q_comp'
FROM OWOR T0 INNER JOIN OITM T1 ON T0.ITEMCODE=T1.ITEMCODE
left JOIN OITB T2 ON T1.[ItmsGrpCod] = T2.[ItmsGrpCod] 
left join orsc t3 on t3.visrescode =t0.u_fase
left join [TIRELLI_40].[dbo].OHEM t4 on t4.[userid]=t0.usersign
LEFT JOIN UFD1 T5 ON T5.tableid='OWOR' and T5.fieldid=2 AND T5.FLDVALUE=t0.u_produzione
LEFT JOIN [dbo].[@FASE] T6 ON T6.CODE=T0.U_fase
left join oitm t7 on t7.itemcode=t0.U_PRG_AZS_Commessa


WHERE 0=0 " & filtro_odp & ""

        Else

            CMD_SAP_2.CommandText = "select t10.numodp as 'Docentry'
, t10.numodp as 'Docnum'
, T10.QTA_PIA AS 'PLANNEDQTY'
, T10.QTA_RES AS 'Q_RESIDUA'
, T10.QTA_PIA-T10.QTA_RES as 'q_comp'
, '' AS 'U_STATO'
, T10.PIANIFICATO AS 'STATO'
, 0 AS 'Lavorazione'
, trim(t10.codart) as 'itemcode'
, t10.dscodart_odp as 'Itemname'
, trim(t10.disegno) as 'Disegno'
, trim(t10.cod_commessa) as 'Commessa'
, trim(t10.matricola) as 'Matricola'
,T10.DESC_COMMESSA
, trim(t10.cod_sottocommessa) as 'Sottocommessa'
, trim(t10.commessa) as 'Commessa_intero'
, t10.cliente as 'Cliente'
, '' as 'Fase'
, '' as 'Gruppo articolo'
, '' as 'u_aggiorna_db'
, CONVERT(DATETIME, CAST(t10.data_iniz AS CHAR(8)), 112) as 'startdate'
, CONVERT(DATETIME, CAST(t10.data_IMMISSIONE AS CHAR(8)), 112) as 'postdate'
, CONVERT(DATETIME, CAST(t10.data_scad AS CHAR(8)), 112) as 'duedate'
, '' as 'Nome'
, '' as 'Type'
, t10.mag_ver as 'warehouse'
, t10.cliente as 'u_utilizz'
, 'Manca' as 'Descr'
, t10.posizione as 'u_progressivo_commessa'
, '' as 'u_produzione'
, coalesce(t12.nome_baia, '') as 'Nome_baia'
FROM OPENQUERY([AS400], '
    SELECT t0.numodp, t0.qta_pia, t0.qta_res, t0.pianificato, t0.codart, t0.dscodart_odp, t0.disegno,
           t0.cod_commessa, t0.matricola,T1.ITEMNAME AS DESC_COMMESSA, t0.cod_sottocommessa,
           t0.commessa as commessa,
           t0.cliente, t0.data_iniz, t0.data_immissione, t0.data_scad,
           t0.mag_ver, t0.posizione
    FROM S786FAD1.TIR90VIS.JGALODP t0
	left join S786FAD1.TIR90VIS.JGALCOM T1 ON T1.matricola=t0.matricola
    where t0.numodp=''" & par_docnum_odp & "''
') AS t10
left join [Tirelli_40].[dbo].[Layout_CAP1] t11
    on t11.Commessa COLLATE SQL_Latin1_General_CP1_CI_AS = t10.matricola COLLATE SQL_Latin1_General_CP1_CI_AS
    and t11.Stato = 'O'
left join [Tirelli_40].[dbo].[Layout_CAP1_nomi] t12
    on t12.numero_baia = t11.Baia"

        End If
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        If cmd_SAP_reader_2.Read() = True Then
            dettagli.docnum = cmd_SAP_reader_2("docnum")
            dettagli.stato = cmd_SAP_reader_2("stato")
            dettagli.itemcode = cmd_SAP_reader_2("Itemcode")
            dettagli.Descrizione = cmd_SAP_reader_2("Itemname")
            dettagli.matricola = cmd_SAP_reader_2("matricola")
            dettagli.commessa = cmd_SAP_reader_2("commessa")
            dettagli.sottocommessa = cmd_SAP_reader_2("sottocommessa")
            dettagli.disegno = cmd_SAP_reader_2("Disegno")
            dettagli.fase = cmd_SAP_reader_2("Fase")
            dettagli.u_produzione = cmd_SAP_reader_2("u_produzione")
            dettagli.lavorazione = cmd_SAP_reader_2("Lavorazione")
            dettagli.u_stato = cmd_SAP_reader_2("u_stato")
            dettagli.quantità = Math.Round(cmd_SAP_reader_2("PLANNEDQTY"))
            dettagli.docentry = cmd_SAP_reader_2("docentry")
            dettagli.u_aggiorna_db = cmd_SAP_reader_2("u_aggiorna_db")
            dettagli.postdate = cmd_SAP_reader_2("POSTDATE")
            dettagli.startdate = cmd_SAP_reader_2("startDATE")
            dettagli.duedate = cmd_SAP_reader_2("DUEDATE")
            dettagli.nome = cmd_SAP_reader_2("NomE")
            dettagli.type = cmd_SAP_reader_2("Type")
            dettagli.warehouse = cmd_SAP_reader_2("warehouse")
            dettagli.u_utilizz = cmd_SAP_reader_2("u_utilizz")
            dettagli.descr = cmd_SAP_reader_2("DESCR")
            dettagli.numerone = cmd_SAP_reader_2("U_PROGRESSIVO_COMMESSA")
            dettagli.q_comp = cmd_SAP_reader_2("Q_comp")
            dettagli.Cliente = cmd_SAP_reader_2("Cliente")
            dettagli.nome_baia = cmd_SAP_reader_2("Nome_baia")
            dettagli.desc_commessa = cmd_SAP_reader_2("desc_commessa")
        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        Return dettagli
    End Function

    Sub formatta_form(par_docnum As String)
        ' 1. Esegui la funzione UNA SOLA VOLTA e salva il risultato in una variabile
        Dim info = ottieni_informazioni_odp("Numero", 0, par_docnum)

        ' 2. Verifica che l'oggetto non sia nullo (buona pratica per evitare crash)
        If info Is Nothing Then Exit Sub

        ' 3. Assegnazione valori ai controlli
        TextBox10.Text = info.docnum
        ComboBox1.Text = info.stato
        Button11.Text = info.itemcode
        TextBox3.Text = info.descrizione
        TextBox6.Text = info.commessa
        TextBox2.Text = info.sottocommessa
        TextBox7.Text = info.commessa


        Button5.Text = info.disegno
        TextBox13.Text = info.lavorazione
        ComboBox4.Text = info.u_stato
        TextBox4.Text = info.quantità
        docentry_odp = info.docentry
        ComboBox5.Text = info.u_aggiorna_db

        ' 4. Gestione date con controlli più puliti
        Try
            Data_ordine.Value = info.postdate
        Catch : End Try

        Try
            Data_inizio.Value = info.startdate
        Catch : End Try

        Data_scadenza.Value = info.duedate

        ' 5. Altri campi
        ComboBox6.Text = info.nome
        TextBox1.Text = info.type
        TextBox5.Text = info.warehouse
        TextBox11.Text = info.u_utilizz


        ' 6. Gestione variabile globale/di modulo e immagine
        numerone = info.numerone
        Label20.Text = numerone

        ' Uso il testo appena assegnato per coerenza
        Magazzino.visualizza_picture(info.disegno, PictureBox2)
    End Sub

    Sub controlli()
        If ComboBox1.Text = "L" Or ComboBox1.Text = "Rilasciato" Then
            ComboBox1.Enabled = False
        Else
            ComboBox1.Enabled = True
        End If
    End Sub

    Sub Inserimento_stati_odp_combobox()
        Dim indice As Integer = 0
        ComboBox1.Items.Clear()

        ComboBox1.Items.Add("P")
        Elenco_stati_odp(indice) = "P"
        indice = indice + 1
        ComboBox1.Items.Add("R")
        Elenco_stati_odp(indice) = "R"
        indice = indice + 1

    End Sub 'Inserisco le risorse nella combo box

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Hide()
    End Sub

    Sub riempi_datagridview_ODP(par_datagridview As DataGridView, par_docnum_odp As String, par_commessa As String)

        Dim filtro_risorse As String = ""

        par_datagridview.Rows.Clear()

            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_2 As New SqlCommand
            Dim cmd_SAP_reader_2 As SqlDataReader


            CMD_SAP_2.Connection = Cnn1



        If Homepage.ERP_provenienza = "SAP" Then

            CMD_SAP_2.CommandText = "declare @odp as integer
declare @commessa as varchar

set @odp='" & par_docnum_odp & "'
set @commessa='" & par_commessa & "'

select t1.itemcode, t1.ItemName,  coalesce(t2.u_disegno,'') as 'u_disegno' ,T1.[BaseQty], T1.[AdditQty], t1.PlannedQty, coalesce(t1.U_PRG_wip_Qtaspedita,0) as 'U_PRG_wip_Qtaspedita', t1.U_PRG_WIP_QtaDaTrasf, t1.wareHouse, case when t1.U_PRG_WIP_QtaDaTrasf=0 then 'OK' when a.giacenza>=t1.U_PRG_WIP_QtaDaTrasf then 'Trasferibile' when a.giacenza+a.CAP2>=t1.U_PRG_WIP_QtaDaTrasf then 'CAP2' when a.giacenza+a.CAP2+a.[CQ-Clavter]>=t1.U_PRG_WIP_QtaDaTrasf then 'CQ-Clavter' when a.giacenza+a.CAP2+a.[CQ-Clavter]+a.ordinato>=t1.U_PRG_WIP_QtaDaTrasf then 'IN APPROV' else'Da ordinare' end as 'Azione',

COALESCE(ODP_ESATTO.DOCNUM,b.odp) as 'ODP'
,COALESCE(ODP_ESATTO.u_progressivo_commessa,b.u_progressivo_commessa) as 'u_progressivo_commessa'
,coalesce(odp_esatto.U_PRG_AZS_Commessa,b.U_PRG_AZS_Commessa) as 'U_PRG_AZS_Commessa'
,coalesce(odp_esatto.u_produzione,b.U_PRODUZIONE) as 'U_produzione'
,coalesce(odp_esatto.duedate,b.Cons_odp) as 'cons_odp'
,b.oa,b.cardname,b.Shipdate,coalesce(odp_esatto.u_stato,coalesce(b.u_stato,'')) as 'u_stato', t1.linenum, t1.VisOrder, t1.U_Stato_lavorazione, t1.itemtype, case when t1.itemtype='4' then 'Articolo' else 'Risorsa' end as 'Tipo'
from owor t0 inner join wor1 t1 on t0.DocEntry=t1.docentry
left join oitm t2 on t2.itemcode=t1.itemcode
left join orsc t3 on t3.visrescode=t1.itemcode

inner join

(
select t30.docnum, t30.itemcode, t30.linenum, t30.Giacenza,t30.CAP2, t30.[CQ-Clavter], sum(case when t31.onorder is null then 0 else t31.onorder end) as 'Ordinato' , sum(case when t31.onhand is null then 0 else t31.onhand end) + sum(case when t31.onorder is null then 0 else t31.onorder end) -sum(case when t31.iscommited is null then 0 else t31.iscommited end)  as 'disponibile'
from
(
select t20.docnum, t20.itemcode, t20.linenum, t20.Giacenza,t20.CAP2, sum(case when t21.onhand is null then 0 else t21.onhand end) as 'CQ-Clavter'
from
(
select t10.docnum, t10.itemcode, t10.linenum, t10.Giacenza, case when t11.onhand is null then 0 else t11.onhand end as 'CAP2'
from
(
select t0.docnum, t1.itemcode, t1.linenum, sum(t2.onhand) as 'Giacenza'
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
inner join oitw t2 on t2.itemcode=t1.itemcode
inner join owhs t3 on t3.whscode=T2.[WhsCode] 
--and t3.location =13
where t0.docnum=@odp and t2.whscode<>'WIP' and t2.whscode<>'BWIP' and t2.whscode<>'CQ' and t2.whscode<>'CAP2' and t2.whscode<>'clavter'

group by t0.docnum, t1.itemcode, t1.linenum

)
as t10 left join oitw t11 on t11.itemcode=t10.itemcode and t11.whscode='CAP2'
)
as t20 left join oitw t21 on t21.itemcode=t20.itemcode and (t21.whscode= 'CQ' or t21.whscode='Clavter')
group by t20.docnum, t20.itemcode, t20.linenum, t20.Giacenza,t20.CAP2
)
as t30 left join oitw t31 on t31.itemcode=t30.itemcode
inner join owhs t32 on t32.whscode=T31.[WhsCode] 
--and t32.location =13
group by t30.docnum, t30.itemcode, t30.linenum, t30.Giacenza,t30.CAP2, t30.[CQ-Clavter]

) A on t0.docnum=a.docnum and t1.linenum=a.linenum

inner join 

(
select t40.docnum, t40.itemcode, t40.linenum, t40.ODP, t41.u_progressivo_commessa, t41.U_PRG_AZS_Commessa,t41.U_PRODUZIONE, case when substring(t41.u_produzione,1,3)='INT' then t41.U_Data_cons_MES else t41.DueDate end as 'Cons_odp',t41.u_stato, t42.docnum as 'OA',t42.cardname, t40.shipdate
from
(
select t30.docnum, t30.itemcode, t30.linenum, t30.ODP, t30.Shipdate,min(t31.docentry) as 'Docentry'
from
(
select t20.docnum, t20.itemcode, t20.linenum, t20.ODP, min(t21.ShipDate) as 'Shipdate'
from
(
select t10.docnum, t10.itemcode, t10.linenum, min(t10.ODP) as 'ODP'
from
(
select t0.docnum, t1.itemcode, t1.linenum, t2.docnum as 'ODP', min(case when substring(t2.u_produzione,1,3)='INT' then t2.U_Data_cons_MES else t2.duedate end)as 'Consegna'
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
left join owor t2 on t2.itemcode=t1.itemcode and (t2.status='P' or t2.status='R') 

where t0.docnum=@odp
group by t0.docnum, t1.itemcode, t1.linenum, t2.docnum
)
as t10 left join owor t11 on t11.docnum=t10.odp and t11.duedate=t10.consegna
group by t10.docnum, t10.itemcode, t10.linenum
)
as t20 left join por1 t21 on t21.OpenQty>0 and t20.ItemCode=t21.itemcode
group by t20.docnum, t20.itemcode, t20.linenum, t20.ODP
)
as t30 left join por1 t31 on t31.shipdate=t30.shipdate and t30.ItemCode=t31.itemcode and t31.OpenQty>0
group by t30.docnum, t30.itemcode, t30.linenum, t30.ODP, t30.Shipdate
)
as t40 left join owor t41 on t41.docnum=t40.odp
left join opor t42 on t42.docentry=t40.docentry

) B on t0.docnum=B.docnum and t1.linenum=B.linenum

left JOIN (
  select 
 t40.Docentry,
t41.itemcode,
t41.Duedate,

t41.PlannedQty,
t41.DocNum, 
t41.u_produzione,
t41.U_PRG_AZS_Commessa,
t41.u_progressivo_Commessa,
t41.u_stato
 from
 (
 select   min(t31.docentry) as 'Docentry'
--		t31.itemcode,
--		t31.Duedate,

--t31.PlannedQty
--,t31.DocNum, 
--t31.u_produzione,
--t31.U_PRG_AZS_Commessa,
--t31.u_progressivo_Commessa,
--t31.u_stato
 from
 (
 select min(t20.DocEntry) as 'docentry' ,t20.itemcode,t20.duedate
 from
 (
  SELECT 
        t11.docentry,
		t11.itemcode,
		min(t11.duedate) as 'Duedate'

    FROM 
        owor t11 INNER JOIN owhs t12 ON t12.WhsCode = T11.warehouse 

    WHERE 
        (t11.status='P' or t11.status='R') and t11.U_PRG_AZS_Commessa= 'M05594'
        group by t11.docentry,t11.itemcode
        )
		as t20 inner join owor t21 on t21.docentry=t20.docentry and t21.duedate=t20.duedate
		group by t20.itemcode,t20.duedate
		)
		as t30 inner join owor t31 on t30.docentry=t31.docentry

		
		)
		as t40 left join owor t41 on t41.docentry=t40.docentry

		) AS odp_esatto 
		on T1.[ItemCode] = odp_esatto.itemcode
AND t1.U_PRG_WIP_QtaDaTrasf > 0

and odp_esatto.PlannedQty>= t1.U_PRG_WIP_QtaDaTrasf




where t0.docnum=@odp   " & filtro_risorse & "
order by t1.visorder
"
        Else
            CMD_SAP_2.CommandText = "select 
trim(t10.codart) as 'itemcode'
,t10.des_code as 'itemname'
,trim(t10.disegno) as 'u_Disegno'
,0 as 'baseqty'
,0 as 'Additqty'
,t10.qtapia as 'Plannedqty'
,t10.qtatra as 'u_prg_wip_qtaspedita'
,t10.qtadatra as 'u_prg_wip_qtadatrasf'
,t10.codmag_im as 'warehouse'
,case when t10.qtadatra=0 then 'OK' else '?' end as 'Azione'
,null as 'odp'
,null as 'u_progressivo_commessa'
,null as 'u_prg_azs_commessa'
,null as 'u_produzione'
,null as 'cons_odp'
,null as 'oa'
,null as 'Cardname'
,null as 'shipdate'
,'' as 'u_stato'
,999 as 'linenum'
,999 as 'visorder'
,'?' as 'u_stato_lavorazione'
,4 as itemtype
,'Articolo' as 'Tipo'



FROM OPENQUERY(AS400, '
select  
    t0.codart,
    t1.des_code,
    t1.disegno,
    t0.qtapia,
    t0.qtatra,
    t0.qtadatra,
    t0.codmag_im,
    t0.commessa as commessa_imp,   -- 👈 alias

    t1.code,
    t1.commessa as commessa_art    -- 👈 alias
from
S786FAD1.TIR90VIS.JGALimp t0
 LEFT JOIN S786FAD1.TIR90VIS.JGALart t1 
        ON t0.codart = t1.code
where odp=''" & par_docnum_odp & "''

'
) as t10
"
        End If

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

            Dim i As Integer = 0
        Do While cmd_SAP_reader_2.Read()

            Dim img As Image = Nothing

            Dim codiceDisegno As String = If(IsDBNull(cmd_SAP_reader_2("u_Disegno")), "", cmd_SAP_reader_2("u_Disegno").ToString())
            Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & codiceDisegno & ".PNG"

            If codiceDisegno <> "" AndAlso File.Exists(percorso) Then
                Using fs As New FileStream(percorso, FileMode.Open, FileAccess.Read)
                    Using tmp As Image = Image.FromStream(fs)
                        img = New Bitmap(tmp) ' evita lock sul file
                    End Using
                End Using
            End If

            Dim idx As Integer = par_datagridview.Rows.Add(
        1,
        cmd_SAP_reader_2("Itemtype"),
        cmd_SAP_reader_2("tipo"),
        cmd_SAP_reader_2("itemcode"),
        img,
        cmd_SAP_reader_2("itemname"),
        cmd_SAP_reader_2("u_Disegno"),
        cmd_SAP_reader_2("baseqty"),
        cmd_SAP_reader_2("additqty"),
        cmd_SAP_reader_2("plannedqty"),
        cmd_SAP_reader_2("U_PRG_wip_Qtaspedita"),
        cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"),
        cmd_SAP_reader_2("wareHouse"),
        cmd_SAP_reader_2("Azione"),
        cmd_SAP_reader_2("ODP"),
        cmd_SAP_reader_2("u_progressivo_commessa"),
        cmd_SAP_reader_2("U_PRG_AZS_Commessa"),
        cmd_SAP_reader_2("u_produzione"),
        cmd_SAP_reader_2("Cons_ODP"),
        cmd_SAP_reader_2("u_stato"),
        cmd_SAP_reader_2("OA"),
        cmd_SAP_reader_2("Cardname"),
        cmd_SAP_reader_2("shipdate"),
        cmd_SAP_reader_2("linenum"),
        cmd_SAP_reader_2("visorder"),
        cmd_SAP_reader_2("u_Stato_lavorazione")
    )

            ' Gestione ReadOnly
            If Not IsDBNull(cmd_SAP_reader_2("U_PRG_wip_Qtaspedita")) AndAlso cmd_SAP_reader_2("U_PRG_wip_Qtaspedita") > 0 Then
                par_datagridview.Rows(idx).ReadOnly = True
                par_datagridview.Rows(idx).Cells("Q_base").ReadOnly = False
                par_datagridview.Rows(idx).Cells("Attrezzaggio").ReadOnly = False
            End If

            par_datagridview.Rows(idx).Cells("tipo").ReadOnly = True
            par_datagridview.Rows(idx).Cells("mag").ReadOnly = False

            ' Altezza riga
            If img Is Nothing Then
                par_datagridview.Rows(idx).Height = 35
            Else
                par_datagridview.Rows(idx).Height = 60
            End If

        Loop



        cmd_SAP_reader_2.Close()
            Cnn1.Close()

            par_datagridview.ClearSelection()

    End Sub

    Sub new_riempi_datagridview_ODP(par_datagridview As DataGridView, par_docnum_odp As Integer, par_commessa As String, par_operazione As String, par_centro_di_costo As String)
        Dim filtro_risorse As String = ""

        par_datagridview.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1



        CMD_SAP_2.CommandText = "
declare @odp as integer


set @odp=" & par_docnum_odp & "


SELECT t0.itemcode,T0.ITEMTYPE, t1.u_prg_azs_commessa
,case when t0.itemtype='4' then 'Articolo' else 'Risorsa' end as 'Tipo'
,t0.ItemName,  coalesce(t2.u_disegno,'') as 'u_disegno',T0.[BaseQty], T0.[AdditQty], t0.PlannedQty, t0.U_PRG_wip_Qtaspedita, t0.U_PRG_WIP_QtaDaTrasf, t0.wareHouse

from wor1 t0 inner join owor t1 on t0.DOCENTRY=t1.docentry
inner join oitm t2 on t2.itemcode=t0.itemcode

where t1.docnum=" & par_docnum_odp & ""

        'trova_giacenza(par_itemcode As String, PAR_OPERAZIONE As String, par_magazzino_diverso_1 As String, par_magazzino_diverso_2 As String, par_magazzino_diverso_3 As String, par_magazzino_diverso_4 As String, par_magazzino_diverso_5 As String, par_magazzino_diverso_6 As String)

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Dim i As Integer = 0
        Dim NUMERO_ODP As String = ""
        Dim azione As String = ""
        Dim Progressivo_commessa As String = ""
        Dim commessa As String = ""
        Dim u_produzione As String = ""
        Dim cons_odp As String
        Dim u_stato As String = ""
        Do While cmd_SAP_reader_2.Read()


            cons_odp = Nothing

            azione = ""

            Progressivo_commessa = ""
            u_produzione = ""
            u_stato = ""
            If cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf") <= 0 Then
                azione = "OK"
                NUMERO_ODP = ""
            ElseIf trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "", "WIP", "BWIP", "CQ", "CLAVTER", "BCLAVTER", "CLAVTER", "CAP2", "BCAP2") >= cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf") Then
                azione = "TRASFERIBILE"
                NUMERO_ODP = ""
            ElseIf trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "", "WIP", "BWIP", "CQ", "CLAVTER", "BCLAVTER", "CLAVTER", "CAP2", "BCAP2") + trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "CAP2", "", "", "", "", "", "", "", "") + trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "BCAP2", "", "", "", "", "", "", "", "") >= cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf") Then
                azione = "CAP2"
                NUMERO_ODP = ""
            ElseIf trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "", "WIP", "BWIP", "CQ", "CLAVTER", "BCLAVTER", "CLAVTER", "CAP2", "BCAP2") + trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "CAP2", "", "", "", "", "", "", "", "") + trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "BCAP2", "", "", "", "", "", "", "", "") + trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "CLAVTER", "", "", "", "", "", "", "", "") + trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "BCLAVTER", "", "", "", "", "", "", "", "") + trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "CQ", "", "", "", "", "", "", "", "") >= cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf") Then

                azione = "CQ-Clavter"
                NUMERO_ODP = ""
            ElseIf trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "", "WIP", "BWIP", "CQ", "CLAVTER", "BCLAVTER", "CLAVTER", "CAP2", "BCAP2") + trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "CAP2", "", "", "", "", "", "", "", "") + trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "BCAP2", "", "", "", "", "", "", "", "") + trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "CLAVTER", "", "", "", "", "", "", "", "") + trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "BCLAVTER", "", "", "", "", "", "", "", "") + trova_giacenza(cmd_SAP_reader_2("itemcode"), par_operazione, "CQ", "", "", "", "", "", "", "", "") + trova_ordinato(cmd_SAP_reader_2("ITEMCODE"), par_operazione) >= cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf") Then
                azione = "IN APPROV"

                NUMERO_ODP = trova_prima_Consegna_odp(cmd_SAP_reader_2("itemcode"), par_operazione, "", "", cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), cmd_SAP_reader_2("U_PRG_azs_commessa")).docnum
                Progressivo_commessa = trova_prima_Consegna_odp(cmd_SAP_reader_2("itemcode"), par_operazione, "", "", cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), cmd_SAP_reader_2("U_PRG_azs_commessa")).numerone
                u_produzione = trova_prima_Consegna_odp(cmd_SAP_reader_2("itemcode"), par_operazione, "", "", cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), cmd_SAP_reader_2("U_PRG_azs_commessa")).u_produzione
                u_stato = trova_prima_Consegna_odp(cmd_SAP_reader_2("itemcode"), par_operazione, "", "", cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), cmd_SAP_reader_2("U_PRG_azs_commessa")).u_stato
                If u_produzione = "INT" Or u_produzione = "B_INT" Then
                    cons_odp = trova_prima_Consegna_odp(cmd_SAP_reader_2("itemcode"), par_operazione, "", "", cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), cmd_SAP_reader_2("U_PRG_azs_commessa")).duedate
                End If
                If NUMERO_ODP = "" Then
                    NUMERO_ODP = trova_prima_Consegna_odp(cmd_SAP_reader_2("itemcode"), par_operazione, "", "", cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), "").docnum
                    commessa = trova_prima_Consegna_odp(cmd_SAP_reader_2("itemcode"), par_operazione, "", "", cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), "").commessa

                    u_produzione = trova_prima_Consegna_odp(cmd_SAP_reader_2("itemcode"), par_operazione, "", "", cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), "").u_produzione

                End If

                If u_produzione = "INT" Or u_produzione = "B_INT" Then
                    cons_odp = trova_prima_Consegna_odp(cmd_SAP_reader_2("itemcode"), par_operazione, "", "", cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), "").duedate
                End If

            Else
                azione = "Da ordinare"
                NUMERO_ODP = ""
            End If


            par_datagridview.Rows.Add(1, cmd_SAP_reader_2("Itemtype"), cmd_SAP_reader_2("tipo"), cmd_SAP_reader_2("itemcode"), cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("u_Disegno"), cmd_SAP_reader_2("baseqty"), cmd_SAP_reader_2("additqty"), cmd_SAP_reader_2("plannedqty"), cmd_SAP_reader_2("U_PRG_wip_Qtaspedita"), cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), cmd_SAP_reader_2("wareHouse"), azione, NUMERO_ODP, Progressivo_commessa, commessa, u_produzione, cons_odp, u_stato)

            If cmd_SAP_reader_2("U_PRG_wip_Qtaspedita") > 0 Then
                par_datagridview.Rows(i).ReadOnly = True
                par_datagridview.Rows(i).Cells(columnName:="Q_base").ReadOnly = False
                par_datagridview.Rows(i).Cells(columnName:="Attrezzaggio").ReadOnly = False
            End If
            par_datagridview.Rows(i).Cells(columnName:="tipo").ReadOnly = True
            i = i + 1


        Loop



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        par_datagridview.ClearSelection()

    End Sub

    Public Function trova_giacenza(par_itemcode As String, PAR_OPERAZIONE As String, par_magazzino_uguale_1 As String, par_magazzino_diverso_1 As String, par_magazzino_diverso_2 As String, par_magazzino_diverso_3 As String, par_magazzino_diverso_4 As String, par_magazzino_diverso_5 As String, par_magazzino_diverso_6 As String, par_magazzino_diverso_7 As String, par_magazzino_diverso_8 As String)
        Dim filtro_magazzino_uguale_1 As String = ""
        If par_magazzino_uguale_1 = "" Then
            filtro_magazzino_uguale_1 = ""
        Else
            filtro_magazzino_uguale_1 = " and t2.whscode = '" & par_magazzino_uguale_1 & "'"
        End If
        Dim giacenza As Decimal = 0

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1



        CMD_SAP_2.CommandText = "select  t0.itemcode,  sum(t2.onhand) as 'Giacenza'

FROM OITM T0
inner join oitw t2 on t2.itemcode=t0.itemcode
inner join owhs t3 on t3.whscode=T2.[WhsCode] and t3.location " & PAR_OPERAZIONE & "13
where t0.itemcode='" & par_itemcode & "' 
" & filtro_magazzino_uguale_1 & "
and t2.whscode<>'" & par_magazzino_diverso_1 & "' and t2.whscode<>'" & par_magazzino_diverso_2 & "' and t2.whscode<>'" & par_magazzino_diverso_3 & "' and t2.whscode<>'" & par_magazzino_diverso_4 & "' and t2.whscode<>'" & par_magazzino_diverso_5 & "' and t2.whscode<>'" & par_magazzino_diverso_6 & "' and t2.whscode<>'" & par_magazzino_diverso_7 & "' and t2.whscode<>'" & par_magazzino_diverso_8 & "'

group by t0.ITEMCODE"



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() Then
            giacenza = cmd_SAP_reader_2("giacenza")
            ' If cmd_SAP_reader_2("U_PRG_wip_Qtaspedita") > 0 Then
        Else
            giacenza = 0
        End If



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        Return giacenza

    End Function

    Public Function trova_ordinato(par_itemcode As String, PAR_OPERAZIONE As String)

        Dim ordinato As Decimal = 0

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1



        CMD_SAP_2.CommandText = "select  t0.itemcode,  sum(coalesce(t2.onorder,0)) as 'Ordinato'

FROM OITM T0
inner join oitw t2 on t2.itemcode=t0.itemcode
inner join owhs t3 on t3.whscode=T2.[WhsCode] and t3.location " & PAR_OPERAZIONE & "13
where t0.itemcode='" & par_itemcode & "' 

group by t0.ITEMCODE"



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() Then
            ordinato = cmd_SAP_reader_2("ordinato")
            ' If cmd_SAP_reader_2("U_PRG_wip_Qtaspedita") > 0 Then
        Else
            ordinato = 0
        End If



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        Return ordinato

    End Function

    Public Function trova_prima_Consegna_odp(par_itemcode As String, PAR_OPERAZIONE As String, par_produzione_1 As String, par_produzione_2 As String, par_quantità As String, par_commessa As String)
        par_quantità = Replace(par_quantità, ",", ".")

        Dim filtro_produzione_1 As String
        If par_produzione_1 = "" Then
            filtro_produzione_1 = ""
        Else
            filtro_produzione_1 = " and t0.u_produzione = '" & par_produzione_1 & "'"
        End If

        Dim filtro_produzione_2 As String
        If par_produzione_2 = "" Then
            filtro_produzione_2 = ""
        Else
            filtro_produzione_2 = " and t0.u_produzione = '" & par_produzione_1 & "'"
        End If


        Dim filtro_commessa As String
        If par_commessa = "" Then
            filtro_commessa = ""
        Else
            filtro_commessa = " and t0.u_prg_azs_commessa   Like '%%" & par_commessa & "%%'  "
        End If

        Dim PRIMA_DATA_ODP As New prima_data_odp()



        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1



        CMD_SAP_2.CommandText = "

select  t0.docnum as 'ODP', COALESCE(t0.U_Data_cons_MES,t0.duedate ) as 'Consegna', t0.u_produzione, t0.status, t0.u_prg_azs_Commessa
,coalesce(t0.u_progressivo_commessa,'') as 'u_progressivo_commessa'
,coalesce(t0.u_stato,'') as 'u_stato'

from owor t0 inner join owhs t1 on t0.warehouse=t1.whscode


where t0.ITEMCODE='" & par_itemcode & "' and (t0.status='P' OR t0.status='R')   and t0.plannedqty>= " & par_quantità & " and t1.location " & PAR_OPERAZIONE & " 13 " & filtro_commessa & filtro_produzione_1 & filtro_produzione_2 & "
ORDER BY COALESCE(t0.U_Data_cons_MES,t0.duedate )
"



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            PRIMA_DATA_ODP.docnum = cmd_SAP_reader_2("ODP")
            PRIMA_DATA_ODP.duedate = cmd_SAP_reader_2("consegna")
            PRIMA_DATA_ODP.u_produzione = cmd_SAP_reader_2("u_produzione")
            PRIMA_DATA_ODP.stato = cmd_SAP_reader_2("status")
            PRIMA_DATA_ODP.commessa = cmd_SAP_reader_2("u_prg_azs_Commessa")
            PRIMA_DATA_ODP.numerone = cmd_SAP_reader_2("u_progressivo_commessa")
            PRIMA_DATA_ODP.u_stato = cmd_SAP_reader_2("u_stato")

        End If



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        Return PRIMA_DATA_ODP

    End Function

    Public Function trova_prima_Consegna_oa(par_itemcode As String, PAR_OPERAZIONE As String, par_quantità As Decimal)


        Dim PRIMA_DATA_OA As New prima_data_oa()



        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1



        CMD_SAP_2.CommandText = "

select  t1.docnum as 'docnum', t1.cardname, t2.docduedate as 'Consegna',  t0.u_prg_azs_Commessa

from por1 t0 inner join inner join opor t1 on t0.docentry=t1.docentry
inner join owhs t2 on t0.whscode=t1.whscode


where t0.itemcode ='" & par_itemcode & "' and t0.openqty>0 and t0.openqty> " & par_quantità & " t0.ITEMCODE='" & par_itemcode & "'  and t2.location " & PAR_OPERAZIONE & " 13 
ORDER BY COALESCE(t2.U_Data_cons_MES,t2.duedate )
"



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            PRIMA_DATA_OA.docnum = cmd_SAP_reader_2("docnum")
            PRIMA_DATA_OA.duedate = cmd_SAP_reader_2("consegna")
            PRIMA_DATA_OA.cardname = cmd_SAP_reader_2("cardname")
        End If



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        Return PRIMA_DATA_OA

    End Function




    Sub esistenza_odp_integrato(par_docnum_odp As Integer)
        variabile_esistenza_odp_integrato = "N"
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1



        CMD_SAP_2.CommandText = "declare @odp as integer

set @odp=" & par_docnum_odp & "

select t0.itemcode from wor1 t0 where t0.U_ODP_integrato=@odp "

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() Then
            variabile_esistenza_odp_integrato = "Y"
        Else
            variabile_esistenza_odp_integrato = "N"
        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    '    Sub riempi_datagridview_ODP_integrato(par_docnum_odp As Integer)
    '        Dim filtro_risorse As String = ""
    '        If Mid(ComboBox2.Text, 1, 3) = "INT" Then
    '            filtro_risorse = "or t3.restype='L'"
    '        Else
    '            filtro_risorse = ""
    '        End If
    '        DataGridView1.Rows.Clear()

    '        Dim Cnn1 As New SqlConnection
    '        Cnn1.ConnectionString = Homepage.sap_tirelli
    '        Cnn1.Open()

    '        Dim CMD_SAP_2 As New SqlCommand
    '        Dim cmd_SAP_reader_2 As SqlDataReader


    '        CMD_SAP_2.Connection = Cnn1



    '        CMD_SAP_2.CommandText = "declare @odp as integer

    'set @odp=" & par_docnum_odp & "


    'select t1.itemcode, t1.ItemName,  t2.u_disegno,T1.[BaseQty], T1.[AdditQty], t1.PlannedQty, t1.U_PRG_wip_Qtaspedita, t1.U_PRG_WIP_QtaDaTrasf, t1.wareHouse, case when t1.U_PRG_WIP_QtaDaTrasf=0 then 'OK' when a.giacenza>=t1.U_PRG_WIP_QtaDaTrasf then 'Trasferibile' when a.giacenza+a.CAP2>=t1.U_PRG_WIP_QtaDaTrasf then 'CAP2' when a.giacenza+a.CAP2+a.[CQ-Clavter]>=t1.U_PRG_WIP_QtaDaTrasf then 'CQ-Clavter' when a.giacenza+a.CAP2+a.[CQ-Clavter]+a.ordinato>=t1.U_PRG_WIP_QtaDaTrasf then 'IN APPROV' else'Da ordinare' end as 'Azione',b.odp,b.U_PRG_AZS_Commessa,b.U_PRODUZIONE,b.Cons_odp,b.oa,b.cardname,b.Shipdate, t1.linenum, t1.VisOrder, t1.U_Stato_lavorazione, t1.itemtype, case when t1.itemtype='4' then 'Articolo' else 'Risorsa' end as 'Tipo'
    'from   wor1 t1 inner join owor t0  on t0.DocEntry=t1.docentry
    'left join oitm t2 on t2.itemcode=t1.itemcode
    'left join orsc t3 on t3.visrescode=t1.itemcode

    'inner join

    '(
    'select t30.docnum, t30.itemcode, t30.linenum, t30.Giacenza,t30.CAP2, t30.[CQ-Clavter], sum(case when t31.onorder is null then 0 else t31.onorder end) as 'Ordinato' , sum(case when t31.onhand is null then 0 else t31.onhand end) + sum(case when t31.onorder is null then 0 else t31.onorder end) -sum(case when t31.iscommited is null then 0 else t31.iscommited end)  as 'disponibile'
    'from
    '(
    'select t20.docnum, t20.itemcode, t20.linenum, t20.Giacenza,t20.CAP2, sum(case when t21.onhand is null then 0 else t21.onhand end) as 'CQ-Clavter'
    'from
    '(
    'select t10.docnum, t10.itemcode, t10.linenum, t10.Giacenza, case when t11.onhand is null then 0 else t11.onhand end as 'CAP2'
    'from
    '(
    'select t0.docnum, t1.itemcode, t1.linenum, sum(t2.onhand) as 'Giacenza'
    'from  wor1 t1  inner join owor t0 on t0.docentry=t1.docentry
    'inner join oitw t2 on t2.itemcode=t1.itemcode
    'where t1.U_ODP_integrato=@odp and t2.whscode<>'WIP' and t2.whscode<>'CQ' and t2.whscode<>'CAP2' and t2.whscode<>'clavter'

    'group by t0.docnum, t1.itemcode, t1.linenum

    ')
    'as t10 left join oitw t11 on t11.itemcode=t10.itemcode and t11.whscode='CAP2'
    ')
    'as t20 left join oitw t21 on t21.itemcode=t20.itemcode and (t21.whscode='CQ' or t21.whscode='Clavter')
    'group by t20.docnum, t20.itemcode, t20.linenum, t20.Giacenza,t20.CAP2
    ')
    'as t30 left join oitw t31 on t31.itemcode=t30.itemcode
    'group by t30.docnum, t30.itemcode, t30.linenum, t30.Giacenza,t30.CAP2, t30.[CQ-Clavter]

    ') A on t0.docnum=a.docnum and t1.linenum=a.linenum

    'inner join 

    '(
    'select t40.docnum, t40.itemcode, t40.linenum, t40.ODP, t41.U_PRG_AZS_Commessa,t41.U_PRODUZIONE, case when substring(t41.u_produzione,1,3)='INT' then t41.U_Data_cons_MES else t41.DueDate end as 'Cons_odp', t42.docnum as 'OA',t42.cardname, t40.shipdate
    'from
    '(
    'select t30.docnum, t30.itemcode, t30.linenum, t30.ODP, t30.Shipdate,min(t31.docentry) as 'Docentry'
    'from
    '(
    'select t20.docnum, t20.itemcode, t20.linenum, t20.ODP, min(t21.ShipDate) as 'Shipdate'
    'from
    '(
    'select t10.docnum, t10.itemcode, t10.linenum, min(t10.ODP) as 'ODP'
    'from
    '(
    'select t0.docnum, t1.itemcode, t1.linenum, t2.docnum as 'ODP', min(case when substring(t2.u_produzione,1,3)='INT' then t2.U_Data_cons_MES else t2.duedate end)as 'Consegna'
    'from  wor1 t1 inner join owor t0 on t0.docentry=t1.docentry
    'left join owor t2 on t2.itemcode=t1.itemcode and (t2.status='P' or t2.status='R') 

    'where t1.u_odp_integrato=@odp

    'group by t0.docnum, t1.itemcode, t1.linenum, t2.docnum
    ')
    'as t10 left join owor t11 on t11.docnum=t10.odp and t11.duedate=t10.consegna
    'group by t10.docnum, t10.itemcode, t10.linenum
    ')
    'as t20 left join por1 t21 on t21.OpenQty>0 and t20.ItemCode=t21.itemcode
    'group by t20.docnum, t20.itemcode, t20.linenum, t20.ODP
    ')
    'as t30 left join por1 t31 on t31.shipdate=t30.shipdate and t30.ItemCode=t31.itemcode and t31.OpenQty>0
    'group by t30.docnum, t30.itemcode, t30.linenum, t30.ODP, t30.Shipdate
    ')
    'as t40 left join owor t41 on t41.docnum=t40.odp
    'left join opor t42 on t42.docentry=t40.docentry

    ') B on t0.docnum=B.docnum and t1.linenum=B.linenum


    'where t1.U_ODP_integrato=@odp  and (substring(T1.[ITEMCODE],1,1)='0' or substring(T1.[ITEMCODE],1,1)='C' or substring(T1.[ITEMCODE],1,1)='D' or (substring(T1.[ITEMCODE],1,1)='R' and t3.restype='M')" & filtro_risorse & ")
    'order by t1.visorder


    '"

    '        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

    '        Dim i As Integer = 0
    '        Do While cmd_SAP_reader_2.Read()
    '            Panel11.Visible = True
    '            DataGridView1.Rows.Add(1, cmd_SAP_reader_2("Itemtype"), cmd_SAP_reader_2("tipo"), cmd_SAP_reader_2("itemcode"), cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("u_Disegno"), cmd_SAP_reader_2("baseqty"), cmd_SAP_reader_2("additqty"), cmd_SAP_reader_2("plannedqty"), cmd_SAP_reader_2("U_PRG_wip_Qtaspedita"), cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), cmd_SAP_reader_2("wareHouse"), cmd_SAP_reader_2("Azione"), cmd_SAP_reader_2("ODP"), cmd_SAP_reader_2("U_PRG_AZS_Commessa"), cmd_SAP_reader_2("u_produzione"), cmd_SAP_reader_2("Cons_ODP"), cmd_SAP_reader_2("OA"), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("shipdate"), cmd_SAP_reader_2("linenum"), cmd_SAP_reader_2("visorder"), cmd_SAP_reader_2("u_Stato_lavorazione"))

    '            i = i + 1
    '        Loop



    '        cmd_SAP_reader_2.Close()
    '        Cnn1.Close()

    '        DataGridView1.ClearSelection()

    '    End Sub





    Sub TROVA_MAX_LINENUM_E_MAX_VISORDER(PAR_DOCNUM_odp As Integer)
        If Homepage.ERP_provenienza = "SAP" Then


            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()


            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

            CMD_SAP.Connection = Cnn
            CMD_SAP.CommandText = "select coalesce(max(coalesce(t0.linenum,0))+1,0) as 'Max_linenum', coalesce(max(coalesce(t0.visorder,0))+1,0) as  'Max_visorder' 
from wor1 t0 inner join owor t1 on t0.docentry=t1.docentry
where t1.docnum=" & PAR_DOCNUM_odp & ""

            cmd_SAP_reader = CMD_SAP.ExecuteReader


            If cmd_SAP_reader.Read() Then

                max_linenum = cmd_SAP_reader("max_linenum")
                max_visorder = cmd_SAP_reader("Max_visorder")

            End If
            cmd_SAP_reader.Close()
            Cnn.Close()
        End If
    End Sub


    Sub Inserimento_items_combobox_stato_lav()
        If Homepage.ERP_provenienza = "SAP" Then


            ComboBox4.Items.Clear()
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()


            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

            CMD_SAP.Connection = Cnn
            CMD_SAP.CommandText = "SELECT *
From ufd1
where tableid='OWOR' and fieldid=76
order by indexid "

            cmd_SAP_reader = CMD_SAP.ExecuteReader

            Dim Indice As Integer
            Indice = 0
            Do While cmd_SAP_reader.Read()

                elenco_stato_lav(Indice) = cmd_SAP_reader("fldvalue")
                ComboBox4.Items.Add(cmd_SAP_reader("Descr"))
                Indice = Indice + 1
            Loop
            cmd_SAP_reader.Close()
            Cnn.Close()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click


        Dim new_Form_nuovo_ticket = New Form_nuovo_ticket

        new_Form_nuovo_ticket.Show()
        new_Form_nuovo_ticket.ComboBox2.Text = Homepage.business
        new_Form_nuovo_ticket.Inserimento_dipendenti()

        new_Form_nuovo_ticket.Administrator = 1
        new_Form_nuovo_ticket.Startup()
        new_Form_nuovo_ticket.Txt_Commessa.Text = TextBox6.Text
        new_Form_nuovo_ticket.Combo_Riferimenti.SelectedIndex = 1
        new_Form_nuovo_ticket.Txt_Nuovo_Riferimento.Text = docnum_odp
        new_Form_nuovo_ticket.Cmd_Aggiungi_Riferimento.PerformClick()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        stampa_etichetta = "no"


        If CheckBox1.Checked = True And CheckBox2.Checked = False Then




            percorso_documento = Homepage.PERCORSO_DOCUMENTO_ODP_ETICHETTA

            stampa_etichetta = "YES"
            Genera_ordine()

        ElseIf CheckBox2.Checked = True Then

            testata_odp(docnum_odp)
            Fun_Stampa()

        End If
        If CheckBox3.Checked = True Then
            testata_odp(docnum_odp)
            If testata_odp_u_disegno = "" Then
                MsgBox("Disegno non presente")
            Else

            End If

        End If

    End Sub

    Sub Genera_ordine()

        testata_odp(docnum_odp)


        oWord = CreateObject("Word.Application")

        oDoc = oWord.Documents.Add(percorso_documento)



        segnalibri_testata_odp()
        If stampa_etichetta = "YES" Then
            segnalibri_etichetta_cassetta()
        End If


        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("Tabella").Range, testata_odp_max_righe + 1, 7)

        oTable.Cell(1, 1).Range.Text = "COD"
        oTable.Cell(1, 2).Range.Text = "Descrizione"
        oTable.Cell(1, 3).Range.Text = "U.M."
        oTable.Cell(1, 4).Range.Text = "Q.TA"
        oTable.Cell(1, 5).Range.Text = "T"
        oTable.Cell(1, 6).Range.Text = "MAG"
        oTable.Cell(1, 7).Range.Text = "UBIC"


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        Dim i As Integer = 2

        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "Select T10.[ItemCode], T10.[ItemName], T10.[InvntryUom], T10.[PlannedQty], case when t10.u_prg_wip_qtaspedita is null then 0 else t10.u_prg_wip_qtaspedita end as 'u_prg_wip_qtaspedita' , T10.[wareHouse]

, t10.u_ubicazione 

, t10.Trasferibile
from
(
SELECT T1.[ItemCode], T2.[ItemName], case when T2.[InvntryUom] is null then '' else T2.[InvntryUom] end as 'Invntryuom', T1.[PlannedQty],  t1.u_prg_wip_qtaspedita  , T1.[wareHouse]

, coalesce( t2.u_ubicazione ,'') as 'U_ubicazione' 

, case when t3.onhand>=T1.[PlannedQty]- case when t1.u_prg_wip_qtaspedita is null then 0 else t1.u_prg_wip_qtaspedita end then 1 else 2 end as 'Trasferibile'

FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
inner join OITM T2 on t2.itemcode=t1.itemcode 
inner join oitw t3 on t3.itemcode=t1.itemcode and t3.whscode=T1.[wareHouse]
WHERE T0.[DocNum] ='" & docnum_odp & "' and t1.itemtype=4
)
as t10
order by t10.Trasferibile,T10.[wareHouse],t10.u_ubicazione, t10.itemcode"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            oTable.Cell(i, 1).Range.Text = cmd_SAP_reader_2("ItemCode")
            oTable.Cell(i, 2).Range.Text = cmd_SAP_reader_2("ItemName")
            oTable.Cell(i, 3).Range.Text = cmd_SAP_reader_2("InvntryUom")

            oTable.Cell(i, 4).Range.Text = FormatNumber(cmd_SAP_reader_2("PlannedQty"), 1, , , TriState.True)



            If cmd_SAP_reader_2("u_prg_wip_qtaspedita") = 0 Then
                oTable.Cell(i, 5).Range.Text = ""
            Else
                oTable.Cell(i, 5).Range.Text = FormatNumber(cmd_SAP_reader_2("u_prg_wip_qtaspedita"), 1, , , TriState.True)
            End If


            oTable.Cell(i, 6).Range.Text = cmd_SAP_reader_2("wareHouse")
            oTable.Cell(i, 7).Range.Text = cmd_SAP_reader_2("u_ubicazione")


            oTable.Cell(i, 2).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
            oTable.Rows(i).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter



            i = i + 1
        Loop
        cmd_SAP_reader_2.Close()


        Cnn1.Close()



        oTable.AutoFormat(ApplyColor:=False, ApplyBorders:=False)
        oTable.Rows.Item(1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Rows.Item(1).Range.Font.Bold = True
        'oTable.Columns.Item(1).Width = oWord.InchesToPoints(1)   'Change width of columns 1 & 2
        oTable.Rows(1).Cells.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        oTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        oTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        oTable.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)


        ' oWord.Visible = True
        ' oWord.ShowMe()



        oWord.PrintOut(Background:=True)





        oWord.Documents.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
        If oWord.Documents.Count = 0 Then
            oWord.Application.Quit()
        End If



    End Sub
    Sub testata_odp(par_docnum_odp)
        ' Una sola chiamata al database
        Dim info = ottieni_informazioni_odp("Numero", 0, par_docnum_odp)

        ' Assegna tutti i campi dall'oggetto già recuperato
        testata_odp_docnum = info.DOCNUM
        testata_odp_itemcode = info.itemcode
        testata_odp_prodname = info.descrizione
        testata_odp_u_disegno = info.disegno
        testata_odp_data = info.postdate
        testata_odp_plannedqty = info.quantità
        testata_odp_resname = info.fase
        testata_odp_u_produzione = info.u_produzione
        testata_odp_commessa = info.matricola
        testata_odp_max_righe = 0
        testata_odp_cardname = info.Cliente
        testata_odp_U_Clientefinale = info.Cliente
        testata_odp_docnum_oc = ""
        testata_odp_cliente_eti = info.Cliente
        testata_odp_Itemname_commessa = info.DESC_COMMESSA
        numerone = info.numerone
        testata_odp_warehouse = info.warehouse

    End Sub

    Sub segnalibri_testata_odp()

        oDoc.Bookmarks.Item("cod_f").Range.Text = testata_odp_itemcode
        oDoc.Bookmarks.Item("commessa").Range.Text = testata_odp_commessa
        oDoc.Bookmarks.Item("descrizione_f").Range.Text = testata_odp_prodname
        oDoc.Bookmarks.Item("data").Range.Text = testata_odp_data
        oDoc.Bookmarks.Item("disegno").Range.Text = testata_odp_u_disegno

        oDoc.Bookmarks.Item("odp").Range.Text = testata_odp_docnum
        oDoc.Bookmarks.Item("qta_f").Range.Text = Math.Round(testata_odp_plannedqty, 1)
        oDoc.Bookmarks.Item("Fase").Range.Text = testata_odp_resname
        oDoc.Bookmarks.Item("Produzione").Range.Text = testata_odp_u_produzione
        oDoc.Bookmarks.Item("Cliente").Range.Text = testata_odp_cardname
        oDoc.Bookmarks.Item("Utilizzatore").Range.Text = testata_odp_U_Clientefinale
        oDoc.Bookmarks.Item("OC").Range.Text = testata_odp_docnum_oc



    End Sub

    Sub segnalibri_etichetta_cassetta()
        oDoc.Bookmarks.Item("ODP_eti").Range.Text = docnum_odp
        oDoc.Bookmarks.Item("cod_eti").Range.Text = testata_odp_itemcode
        oDoc.Bookmarks.Item("fase_eti").Range.Text = testata_odp_resname
        oDoc.Bookmarks.Item("data_eti").Range.Text = testata_odp_data
        oDoc.Bookmarks.Item("Desc_eti").Range.Text = testata_odp_prodname
        oDoc.Bookmarks.Item("prod_eti").Range.Text = testata_odp_u_produzione

        oDoc.Bookmarks.Item("commessa_eti").Range.Text = testata_odp_commessa
        oDoc.Bookmarks.Item("cliente_eti").Range.Text = testata_odp_cliente_eti
        oDoc.Bookmarks.Item("oc_eti").Range.Text = testata_odp_docnum_oc
        oDoc.Bookmarks.Item("modello").Range.Text = testata_odp_Itemname_commessa
        oDoc.Bookmarks.Item("NUMERONE").Range.Text = numerone




    End Sub

    Private Sub Cmd_Materiale_Click(sender As Object, e As EventArgs) Handles Cmd_Materiale.Click
        Form_Richiesta_Materiale.Show()
        Form_Richiesta_Materiale.Owner = Me
        Form_Richiesta_Materiale.Txt_Commessa.Text = TextBox6.Text
        Form_Richiesta_Materiale.TXT_ODP.Text = docnum_odp
        Form_Richiesta_Materiale.Home_Lista()
        Me.Hide()
    End Sub





    Private Sub CancellaRigaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CancellaRigaToolStripMenuItem.Click
        If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Trasferito").Value > 0 Then
            MsgBox("Impossibile cancellare riga di un codice che risulta TRASFERITO")
        Else

            If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Presente").Value <> 0 Then
                Righe_cancellate(num_righe_cancellate).Codice_riga = DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Codice").Value
                Righe_cancellate(num_righe_cancellate).linenum = DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Linenum").Value
                Righe_cancellate(num_righe_cancellate).visorder = DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Visorder").Value
                num_righe_cancellate = num_righe_cancellate + 1
            End If
            DataGridView_ODP.Rows.RemoveAt(DATAGRIDVIEW_odp_RIGA)
        End If


    End Sub





    Private Sub DataGridView_ODP_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellClick
        If e.RowIndex >= 0 Then


            DataGridView_ODP.SelectionMode = DataGridViewSelectionMode.CellSelect
            visorder_selezionato = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Visorder").Value
            riga_selezionata = e.RowIndex

            If riga_selezionata = 0 Then
                Button3.Visible = False
            Else
                Button3.Visible = True
            End If

            If riga_selezionata >= DataGridView_ODP.Rows.Count - 2 Then
                Button4.Visible = False
            Else
                Button4.Visible = True
            End If

            If e.ColumnIndex = DataGridView_ODP.Columns.IndexOf(ODP) Then
                Dim new_form_odp_form = New ODP_Form
                new_form_odp_form.docnum_odp = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="ODP").Value
                new_form_odp_form.Show()
                new_form_odp_form.inizializza_form(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="ODP").Value)






            End If

            If e.ColumnIndex = DataGridView_ODP.Columns.IndexOf(Disegno) Then


                Magazzino.visualizza_disegno(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Disegno").Value)

            ElseIf e.ColumnIndex = DataGridView_ODP.Columns.IndexOf(img) Then

                Form_visualizza_picture.Show()
                Magazzino.visualizza_picture(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Disegno").Value, Form_visualizza_picture.PictureBox1)

            End If

        End If
    End Sub



    Private Sub DataGridView_ODP_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView_ODP.CellMouseDown
        If e.RowIndex >= 0 Then


            DATAGRIDVIEW_odp_RIGA = e.RowIndex
            datagridview_odp_colonna = e.ColumnIndex
            visorder_selezionato = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Visorder").Value
            DataGridView_ODP.ClearSelection()
            DataGridView_ODP.Rows(e.RowIndex).Cells(e.ColumnIndex).Selected = True
        End If

    End Sub

    Sub Inserimento_magazzini_righe()
        If Homepage.ERP_provenienza = "SAP" Then


            Dim Cnn As New SqlConnection

            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()



            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

            CMD_SAP.Connection = Cnn
            CMD_SAP.CommandText = "select t0.whscode, t0.whsname
from owhs t0
where t0.locked='N'"

            cmd_SAP_reader = CMD_SAP.ExecuteReader


            Do While cmd_SAP_reader.Read()
                ' MAG.Items.Add(cmd_SAP_reader("whscode"))
                'DataGridViewComboBoxColumn2.Items.Add(cmd_SAP_reader("whscode"))

            Loop
            cmd_SAP_reader.Close()
            Cnn.Close()
        End If
    End Sub



    Sub informazioni_articolo_riga(par_itemcode As String, par_riga As Integer)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "Select Case When T2.[VisResCode] Is null Then T0.objTYPE Else '290' end as 'objtype',
Case When T2.[VisResCode] Is null Then 'Articolo' Else 'Risorsa' end as 'Tipo_Articolo',
T0.[ItemName]

, case when T0.[DfltWH] is null then '01' else T0.[DfltWH] end as 'DfltWH', T1.[Price], T0.VALIDFOR 

FROM OITM T0  INNER JOIN ITM1 T1 ON T0.[ItemCode] = T1.[ItemCode] left join orsc t2 on T2.[VisResCode]=t0.itemcode
WHERE T0.[ItemCode] ='" & par_itemcode & "' AND  T1.[PriceList] =2"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then


            If cmd_SAP_reader("VALIDFOR") = "N" Then
                MsgBox("Articolo inattivo")
            Else

                DataGridView_ODP.Rows(par_riga).Cells(columnName:="Descrizione").Value = cmd_SAP_reader("ItemName")
                DataGridView_ODP.Rows(par_riga).Cells(columnName:="Itemtype").Value = cmd_SAP_reader("objTYPE")

                DataGridView_ODP.Rows(par_riga).Cells(columnName:="mag").Value = cmd_SAP_reader("DfltWH")


                DataGridView_ODP.Rows(par_riga).Cells(columnName:="tipo").Value = cmd_SAP_reader("Tipo_articolo")
                DataGridView_ODP.Rows(par_riga).Cells(columnName:="Q_base").Value = 1
                DataGridView_ODP.Rows(par_riga).Cells(columnName:="attrezzaggio").Value = 0
                DataGridView_ODP.Rows(par_riga).Cells(columnName:="Quantità").Value = TextBox4.Text
                DataGridView_ODP.Rows(par_riga).Cells(columnName:="Trasferito").Value = 0
                DataGridView_ODP.Rows(par_riga).Cells(columnName:="Da_trasf").Value = TextBox4.Text
                DataGridView_ODP.Rows(par_riga).Cells(columnName:="Stato").Value = "O"
                If DataGridView_ODP.Rows(par_riga).Cells(columnName:="linenum").Value >= 0 Then

                Else

                    If DataGridView_ODP.Rows(par_riga).Cells(columnName:="linenum").Value = Nothing Or DataGridView_ODP.Rows(par_riga).Cells(columnName:="linenum").Value = "" Then
                        DataGridView_ODP.Rows(par_riga).Cells(columnName:="linenum").Value = trova_max_linenum(TextBox10.Text)
                    End If
                End If



                'DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="linenum").Value = max_linenum
                'DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="visorder").Value = max_visorder

                max_linenum = max_linenum + 1
                max_visorder = max_visorder + 1

            End If
        Else
            MsgBox("Articolo non esistente")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()


    End Sub


    Public Function trova_max_linenum(par_docnum_odp As Integer)

        Dim linenum_max As Integer = 0
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "select max(coalesce(t1.linenum,0))+1 as 'Linenum'
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
where t0.docnum=" & par_docnum_odp & ""

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            linenum_max = If(IsDBNull(cmd_SAP_reader("Linenum")), 0, Convert.ToInt32(cmd_SAP_reader("Linenum")))


        End If
        cmd_SAP_reader.Close()
        Cnn.Close()

        Return linenum_max
    End Function



    Private Sub DataGridView_ODP_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellValueChanged
        If e.RowIndex >= 0 Then


            If e.ColumnIndex = DataGridView_ODP.Columns.IndexOf(Codice) And DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Codice").Value <> "" Then


                'Try
                itemcode_riga = UCase(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Codice").Value)


                ' Controllo duplicati
                For i As Integer = 0 To DataGridView_ODP.Rows.Count - 1
                    If i <> e.RowIndex Then ' Salta la riga corrente
                        Dim valoreCorrente As Object = DataGridView_ODP.Rows(i).Cells("Codice").Value
                        If Not IsDBNull(valoreCorrente) And valoreCorrente <> "" AndAlso UCase(valoreCorrente.ToString()) = itemcode_riga Then
                            MessageBox.Show("Il codice '" & itemcode_riga & "' è già presente in un'altra riga.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            DataGridView_ODP.Rows(e.RowIndex).Cells("Codice").Value = "" ' cancella valore duplicato
                            Return
                        End If
                    End If
                Next
                informazioni_articolo_riga(itemcode_riga, e.RowIndex)


                'Catch ex As Exception
                '    MsgBox("C'è un errore nell'articolo riga")
                'End Try

            ElseIf e.ColumnIndex = DataGridView_ODP.Columns.IndexOf(Q_base) Or e.ColumnIndex = DataGridView_ODP.Columns.IndexOf(Attrezzaggio) Then

                Try

                    If InStr(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Q_base").Value, ",") > 1 Then


                        quantità_base_riga = LSet(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Q_base").Value, InStr(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Q_base").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Q_base").Value), InStr(StrReverse(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Q_base").Value), ",") - 1))



                    Else
                        quantità_base_riga = Replace(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Q_base").Value, ",", ".")
                    End If



                    If InStr(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="attrezzaggio").Value, ",") > 1 Then


                        attrezzaggio_riga = LSet(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="attrezzaggio").Value, InStr(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="attrezzaggio").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Attrezzaggio").Value), InStr(StrReverse(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="attrezzaggio").Value), ",") - 1))

                    Else
                        attrezzaggio_riga = Replace(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="attrezzaggio").Value, ",", ".")
                    End If


                    Dim q_odp As String

                    If InStr(TextBox4.Text, ",") > 1 Then


                        q_odp = LSet(TextBox4.Text, InStr(TextBox4.Text, ",") - 1) & "." & StrReverse(LSet(StrReverse(TextBox4.Text), InStr(StrReverse(TextBox4.Text), ",") - 1))

                    Else
                        q_odp = Replace(TextBox4.Text, ",", ".")
                    End If

                    Dim trasferito As String

                    If InStr(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="trasferito").Value, ",") > 1 Then


                        trasferito = LSet(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="trasferito").Value, InStr(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="trasferito").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="trasferito").Value), InStr(StrReverse(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="trasferito").Value), ",") - 1))

                    Else
                        trasferito = Replace(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="trasferito").Value, ",", ".")
                    End If
                    Dim Cnn As New SqlConnection
                    Cnn.ConnectionString = Homepage.sap_tirelli
                    Cnn.Open()

                    Dim CMD_SAP As New SqlCommand
                    Dim cmd_SAP_reader As SqlDataReader
                    CMD_SAP.Connection = Cnn
                    '

                    CMD_SAP.CommandText = "SELECT (" & quantità_base_riga & " * " & q_odp & ")+" & attrezzaggio_riga & " as 'Totale',(" & quantità_base_riga & " * " & q_odp & ")+" & attrezzaggio_riga & "- " & trasferito & " as 'Da_trasferire' "

                    cmd_SAP_reader = CMD_SAP.ExecuteReader
                    If cmd_SAP_reader.Read() = True Then

                        DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Quantità").Value = cmd_SAP_reader("Totale")
                        DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Da_trasf").Value = cmd_SAP_reader("Da_trasferire")


                    End If
                    cmd_SAP_reader.Close()
                    Cnn.Close()



                Catch ex As Exception

                End Try

            End If
            'If DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Presente").Value = 1 Then

            '    DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Presente").Value = 2
            'ElseIf DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Presente").Value <> 2 Then

            '    DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Presente").Value = 0
            'End If
        End If

    End Sub



    'Sub aggiorna_odp()


    '    Dim password = InputBox("Inserire password")

    '    If ottieni_informazioni_odp("Numero", 0, TextBox10.Text).stato = "P" Or ottieni_informazioni_odp("Numero", 0, TextBox10.Text).stato = "R" Then

    '        ' If ottieni_informazioni_odp(TextBox10.Text).q_comp < ottieni_informazioni_odp(TextBox10.Text).quantità Then


    '        If UCase(password) = "-" Or UCase(password) = "." Then
    '            'inserisci_record_modifica_odp(Homepage.ID_SALVATO, TextBox10.Text)
    '            'aggiorna_odp_doc(TextBox10.Text, Elenco_stati_odp(ComboBox1.SelectedIndex), Replace(TextBox4.Text, ",", "."), TextBox13.Text, ComboBox4.Text, ComboBox5.Text, TextBox6.Text, TextBox11.Text, Elenco_produzione(ComboBox2.SelectedIndex), Elenco_fase(ComboBox3.SelectedIndex), Data_inizio.Value, Data_scadenza.Value)
    '            inserisci_record_modifica_odp(Homepage.ID_SALVATO, TextBox10.Text)
    '            ' aggiorna_odp_doc(TextBox10.Text, Elenco_stati_odp(ComboBox1.SelectedIndex), Replace(TextBox4.Text, ",", "."), TextBox13.Text, ComboBox4.Text, ComboBox5.Text, TextBox6.Text, TextBox11.Text, Elenco_produzione(ComboBox2.SelectedIndex), Elenco_fase(ComboBox3.SelectedIndex), Data_inizio.Value, Data_scadenza.Value)

    '            memorizza_codici(TextBox10.Text)
    '            cancella_righe_odp(TextBox10.Text)
    '            new_aggiornamento_righe()
    '            AWOR(TextBox10.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
    '            AWO1(TextBox10.Text)

    '            ripara_confermati_vecchi()

    '            MsgBox("Documento aggiornato con successo")

    '            formatta_form(docnum_odp)
    '            riempi_datagridview_ODP(DataGridView_ODP, docnum_odp, TextBox6.Text)


    '            controlli()
    '        Else
    '            MsgBox("Password errata")
    '        End If
    '        ' Else
    '        '  MsgBox("Quantità completata >= di quantità pianificata")
    '        'End If
    '    Else
    '        MsgBox("Non è possibile apportare modifiche ad un ODP chiuso")
    '    End If

    'End Sub

    Sub ripara_confermati_vecchi()
        Dim c As Integer = 0
        Do While c <= contatore_codici
            ripara_confermati(codici_salvati(c))
            ripara_confermati_PARTENDO_DA_OITW(codici_salvati(c))
            c += 1
        Loop

    End Sub

    Sub memorizza_codici(par_docnum)




        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select t1.itemcode
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
where t0.docnum='" & par_docnum & "'
group by t1.itemcode"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        contatore_codici = 0
        Do While cmd_SAP_reader_2.Read()



            If Not cmd_SAP_reader_2("itemcode") Is System.DBNull.Value Then
                codici_salvati(contatore_codici) = cmd_SAP_reader_2("itemcode")
                contatore_codici += 1

            End If

        Loop
        Cnn1.Close()



    End Sub



    Sub aggiorna_odp_doc(par_numero_odp As String, par_status As String, par_quantity As String, par_lavorazione As String, par_stato As String, par_aggiorna_db As String, par_commessa As String, par_utilizzatore As String, par_produzione As String, par_fase As String, par_data_inizio As String, par_data_fine As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "UPDATE owor SET STATUS='" & par_status & "', plannedqty = '" & par_quantity & "' 
, u_lavorazione = '" & par_lavorazione & "', u_stato='" & par_stato & "' , u_aggiorna_db='" & par_aggiorna_db & "', u_prg_azs_commessa='" & par_commessa & "'
, u_utilizz='" & par_utilizzatore & "', u_produzione='" & par_produzione & "'
, u_fase='" & par_fase & "',startdate = CONVERT(DATETIME, '" & par_data_inizio & "', 103) ,duedate = CONVERT(DATETIME, '" & par_data_fine & "', 103) 

where docnum='" & par_numero_odp & "' "

        CMD_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        ' Accetto solo l'immissione di numeri interi e decimali

        ' Recupero il codice ascii del tasto digitato
        ' il tasto digitato è memorizzato nella proprietà "KeyChar"
        ' dell'oggetto System.Windows.Forms.KeyPressEventArgs

        Dim KeyAscii As Short = Asc(e.KeyChar)

        ' In questo caso oltre a consentire numeri, tasto Canc
        ' e tasto BackSpace, devo consentire anche l'immissione
        ' del punto e della virgola
        If KeyAscii < 48 And KeyAscii <> 24 And KeyAscii <> 8 And e.KeyChar <> "." And e.KeyChar <> "," Then
            KeyAscii = 0
        ElseIf KeyAscii > 57 Then
            KeyAscii = 0
        End If

        ' Faccio in modo che se l'utente digita la virgola
        ' mi appaia il punto
        If e.KeyChar = "," Then
            KeyAscii = 46 ' 46 è il codice ascii del punto
        End If

        ' Il punto è si consentito
        ' ma non come primo carattere
        If TextBox4.TextLength = 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If

        ' ovviamente se c'è già un punto
        ' non è consentito digitarne altri
        If (KeyAscii = 46) And
            TextBox4.Text.IndexOf(".") > 0 Then
            KeyAscii = 0
        End If

        ' Reimposto il keychar
        e.KeyChar = Chr(KeyAscii)
    End Sub



    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles Data_scadenza.ValueChanged
        If Data_scadenza.Value.Date < Data_inizio.Value.Date Or Data_scadenza.Value.Date < Data_ordine.Value.Date Then
            MsgBox("Impossibile scegliere data di scandenza > di data di inizio o di data di creazione")
            Data_scadenza.Value = Data_inizio.Value
        End If
    End Sub





    Private Sub DataGridView_ODP_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_ODP.CellFormatting
        If e.RowIndex <= DataGridView_ODP.Rows.Count - 2 Then
            ' Il codice da eseguire se l'indice di riga è valido


            Dim azioneValue As String
            Dim statoValue As String
            Dim complValue As String


            If DataGridView_ODP.Rows(e.RowIndex).Cells("Azione").Value = Nothing Then
                azioneValue = ""
            Else
                azioneValue = DataGridView_ODP.Rows(e.RowIndex).Cells("Azione").Value.ToString().ToLower()
            End If

            If DataGridView_ODP.Rows(e.RowIndex).Cells("Stato").Value = Nothing Then
                statoValue = ""
            Else
                statoValue = DataGridView_ODP.Rows(e.RowIndex).Cells("Stato").Value.ToString()
            End If

            If DataGridView_ODP.Rows(e.RowIndex).Cells("compl").Value = Nothing Then
                complValue = ""
            Else
                complValue = DataGridView_ODP.Rows(e.RowIndex).Cells("compl").Value.ToString().ToLower()
            End If


            Select Case azioneValue
                Case "ok"
                    DataGridView_ODP.Rows(e.RowIndex).Cells("Azione").Style.BackColor = Color.Lime
                Case "trasferibile"
                    DataGridView_ODP.Rows(e.RowIndex).Cells("Azione").Style.BackColor = Color.Yellow

                Case "trasferibile/da ordinare"
                    DataGridView_ODP.Rows(e.RowIndex).Cells("Azione").Style.BackColor = Color.MediumSpringGreen
                Case "in approv/da ordinare"
                    DataGridView_ODP.Rows(e.RowIndex).Cells("Azione").Style.BackColor = Color.Orange
                Case "in approv"
                    DataGridView_ODP.Rows(e.RowIndex).Cells("Azione").Style.BackColor = Color.Khaki
                Case "da ordinare"
                    DataGridView_ODP.Rows(e.RowIndex).Cells("Azione").Style.BackColor = Color.Red
            End Select

            If statoValue = "C" Then
                DataGridView_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
            End If

            If complValue = "completato" Then
                DataGridView_ODP.Rows(e.RowIndex).Cells("compl").Style.BackColor = Color.Lime
            End If

            If CInt(DataGridView_ODP.Rows(e.RowIndex).Cells("da_trasf").Value) > 0 Then
                With DataGridView_ODP.Rows(e.RowIndex).Cells("da_trasf").Style
                    .Font = New Font(DataGridView_ODP.Font, FontStyle.Bold)
                    .ForeColor = Color.Black
                End With
            Else
                With DataGridView_ODP.Rows(e.RowIndex).Cells("da_trasf").Style
                    .Font = New Font(DataGridView_ODP.Font, FontStyle.Regular)
                    .ForeColor = Color.Black
                End With
            End If

        End If



    End Sub

    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles Data_inizio.ValueChanged
        If Data_scadenza.Value < Data_inizio.Value Then

            Data_scadenza.Value = Data_inizio.Value
        End If
    End Sub

    Sub new_aggiornamento_righe()
        contatore = 0
        Do While contatore <= DataGridView_ODP.Rows.Count - 2


            inserisci_riga_odp(DataGridView_ODP.Rows(contatore).Cells(columnName:="linenum").Value, DataGridView_ODP.Rows(contatore).Cells(columnName:="Codice").Value, contatore, DataGridView_ODP.Rows(contatore).Cells(columnName:="Trasferito").Value, DataGridView_ODP.Rows(contatore).Cells(columnName:="Da_trasf").Value)



            contatore = contatore + 1
        Loop
    End Sub


    Sub aggiornamento_righe()
        contatore = 0
        Do While contatore <= DataGridView_ODP.Rows.Count - 2


            If DataGridView_ODP.Rows(contatore).Cells(columnName:="Presente").Value = 0 Or DataGridView_ODP.Rows(contatore).Cells(columnName:="Presente").Value = 2 Then
                ' inserisci_riga_odp(DataGridView_ODP.Rows(contatore).Cells(columnName:="linenum").Value, DataGridView_ODP.Rows(contatore).Cells(columnName:="Codice").Value, contatore)


            End If
            contatore = contatore + 1
        Loop
    End Sub





    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim r_che_deve_salire = DataGridView_ODP.Rows(riga_selezionata)
        Dim r_che_deve_scendere = DataGridView_ODP.Rows(riga_selezionata - 1)
        DataGridView_ODP.Rows.Remove(r_che_deve_salire)
        DataGridView_ODP.Rows.Remove(r_che_deve_scendere)

        DataGridView_ODP.Rows.Insert(riga_selezionata - 1, r_che_deve_salire)
        DataGridView_ODP.Rows.Insert(riga_selezionata, r_che_deve_scendere)

        '
        DataGridView_ODP.ClearSelection()
        DataGridView_ODP.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView_ODP.Rows(riga_selezionata - 1).Selected = True

        Dim visorder_che_deve_salire As Integer
        Dim visorder_che_deve_scendere As Integer

        visorder_che_deve_salire = DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="visorder").Value
        visorder_che_deve_scendere = DataGridView_ODP.Rows(riga_selezionata - 1).Cells(columnName:="visorder").Value

        DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="visorder").Value = visorder_che_deve_scendere
        DataGridView_ODP.Rows(riga_selezionata - 1).Cells(columnName:="visorder").Value = visorder_che_deve_salire


        If DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="presente").Value = 1 Then

            DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="presente").Value = 2
        ElseIf DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="presente").Value <> 2 Then
            DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="presente").Value = 0
        End If

        If DataGridView_ODP.Rows(riga_selezionata - 1).Cells(columnName:="presente").Value = 1 Then

            DataGridView_ODP.Rows(riga_selezionata - 1).Cells(columnName:="presente").Value = 2
        ElseIf DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="presente").Value <> 2 Then
            DataGridView_ODP.Rows(riga_selezionata - 1).Cells(columnName:="presente").Value = 0
        End If



        riga_selezionata = riga_selezionata - 1

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim r_che_deve_salire = DataGridView_ODP.Rows(riga_selezionata + 1)
        Dim r_che_deve_scendere = DataGridView_ODP.Rows(riga_selezionata)
        DataGridView_ODP.Rows.Remove(r_che_deve_salire)
        DataGridView_ODP.Rows.Remove(r_che_deve_scendere)

        DataGridView_ODP.Rows.Insert(riga_selezionata, r_che_deve_salire)
        DataGridView_ODP.Rows.Insert(riga_selezionata + 1, r_che_deve_scendere)

        '
        DataGridView_ODP.ClearSelection()
        DataGridView_ODP.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView_ODP.Rows(riga_selezionata + 1).Selected = True

        Dim visorder_che_deve_salire As Integer
        Dim visorder_che_deve_scendere As Integer

        visorder_che_deve_salire = DataGridView_ODP.Rows(riga_selezionata + 1).Cells(columnName:="visorder").Value
        visorder_che_deve_scendere = DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="visorder").Value

        DataGridView_ODP.Rows(riga_selezionata + 1).Cells(columnName:="visorder").Value = visorder_che_deve_scendere
        DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="visorder").Value = visorder_che_deve_salire




        If DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="presente").Value = 1 Then

            DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="presente").Value = 2
        ElseIf DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="presente").Value <> 2 Then
            DataGridView_ODP.Rows(riga_selezionata).Cells(columnName:="presente").Value = 0
        End If

        If DataGridView_ODP.Rows(riga_selezionata + 1).Cells(columnName:="presente").Value = 1 Then

            DataGridView_ODP.Rows(riga_selezionata + 1).Cells(columnName:="presente").Value = 2
        ElseIf DataGridView_ODP.Rows(riga_selezionata + 1).Cells(columnName:="presente").Value <> 2 Then
            DataGridView_ODP.Rows(riga_selezionata + 1).Cells(columnName:="presente").Value = 0
        End If




        riga_selezionata = riga_selezionata + 1
    End Sub

    Sub inserisci_riga_odp(par_linenum As Integer, par_itemcode As String, par_contatore As Integer, par_trasferito As String, par_da_trasferire As String)

        ' itemcode_riga = DataGridView_ODP.Rows(contatore).Cells(columnName:="Codice").Value
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand
        Dim cmd_SAP_reader_7 As SqlDataReader
        CMD_SAP_7.Connection = Cnn

        CMD_SAP_7.CommandText = "SELECT top 1 T1.VALIDFOR AS 'Valido', t1.itemcode as 'Codice' FROM OITM T1 
WHERE T1.[itemcode]= '" & par_itemcode & "'"

        cmd_SAP_reader_7 = CMD_SAP_7.ExecuteReader
        If cmd_SAP_reader_7.Read() = True Then
            If cmd_SAP_reader_7("Valido") = "N" Then
                MsgBox("Il codice " & par_itemcode & " è inattivo ")
            Else


                inserisci_riga(TextBox10.Text, cmd_SAP_reader_7("Codice"), DataGridView_ODP.Rows(par_contatore).Cells(columnName:="Quantità").Value, DataGridView_ODP.Rows(par_contatore).Cells(columnName:="MAG").Value, DataGridView_ODP.Rows(par_contatore).Cells(columnName:="Itemtype").Value, DataGridView_ODP.Rows(par_contatore).Cells(columnName:="Attrezzaggio").Value, par_trasferito, par_da_trasferire, par_linenum)



                ripara_confermati(par_itemcode)

            End If

        Else
            MsgBox("Il codice " & itemcode_riga & " non esiste ")
        End If
        Cnn.Close()



    End Sub

    Sub inserisci_riga(par_docnum As Integer, par_itemcode As String, par_quantità_TOT As String, par_magazzino As String, par_itemtype As String, par_attrezzaggio As String, par_trasferito As String, par_da_trasferire As String, par_linenum As Integer)

        par_quantità_TOT = Replace(par_quantità_TOT, ",", ".")
        par_attrezzaggio = Replace(par_attrezzaggio, ",", ".")
        par_trasferito = Replace(par_trasferito, ",", ".")
        par_da_trasferire = Replace(par_da_trasferire, ",", ".")
        If par_linenum <= trova_max_linenum(par_docnum) - 1 Then
            par_linenum = trova_max_linenum(par_docnum)
        End If
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()
        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "
declare @maxvisorder as integer
declare @maxlinenum as integer

select @maxvisorder=coalesce(max(t1.visorder),0),@maxlinenum=coalesce(max(t1.linenum),0)
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
where t0.docnum=" & par_docnum & "


insert into WOR1 (WOR1.DocEntry, WOR1.LineNum, WOR1.ItemCode, WOR1.BaseQty, WOR1.PlannedQty, WOR1.IssuedQty, WOR1.IssueType, WOR1.wareHouse, WOR1.VisOrder, WOR1.WipActCode, WOR1.CompTotal, WOR1.OcrCode, WOR1.OcrCode2, WOR1.OcrCode3, WOR1.OcrCode4, WOR1.OcrCode5, WOR1.LocCode, WOR1.Project, WOR1.UomEntry, WOR1.UomCode, WOR1.ItemType, WOR1.AdditQty, WOR1.LineText, WOR1.PickStatus, WOR1.PickQty, WOR1.PickIdNo, WOR1.ReleaseQty, WOR1.ResAlloc, WOR1.StartDate, WOR1.EndDate, WOR1.StageId, WOR1.BaseQtyNum, WOR1.BaseQtyDen, WOR1.ReqDays, WOR1.RtCalcProp, WOR1.Status, WOR1.ItemName, WOR1.AlwProcDoc, WOR1.PoDocType, WOR1.PoDocNum, WOR1.PoDocEntry, WOR1.PoLineNum, WOR1.PoQuantity, WOR1.U_TEMPOME, WOR1.U_UBIMAG
, WOR1.U_Prezzolis, WOR1.U_CodDis, WOR1.U_PRG_AZS_Terzista, WOR1.U_PRG_AZS_StatoAv
, WOR1.U_PRG_AZS_PhanFat, WOR1.U_PRG_CLV_Ris_Orig, WOR1.U_PRG_CLV_Qta_Trasf,  WOR1.U_DisponiibileTOT
,WOR1.U_PRG_WIP_QtaSpedita, WOR1.U_PRG_WIP_QtaDaTrasf
, WOR1.U_Data_ora_inizio_fase, WOR1.U_Data_ora_fine_fase, WOR1.U_Dipendente, WOR1.U_Ordinato_TOT, WOR1.U_Confermato_TOT, WOR1.U_PRG_WIP_QtaRichMagAuto, WOR1.U_PRG_WMS_Exp, WOR1.U_PRG_WMS_ExpDate, WOR1.U_PRG_WMS_MdMovQty, WOR1.U_Stato_lavorazione)

SELECT t1.docentry, " & par_linenum & ",'" & par_itemcode & "' , " & par_quantità_TOT & "/case when t1.plannedqty = 0 then 1 else t1.plannedqty end , " & par_quantità_TOT & ",0,'B','" & par_magazzino & "',@maxvisorder+1,'',0,'','','','','','','','',''," & par_itemtype & "," & par_attrezzaggio & ",'','N',0,0,0,CASE WHEN " & par_itemtype & "='290' then 'F' else null end,T1.[STARTDate], T1.[DueDate],NULL,0,0,0,100,T1.STATUS,T3.itemname,'N','','','','',0,0,''
,T4.PRICE,'','','X'
,'',''
,0,0 ," & par_trasferito & "," & par_da_trasferire & " ,'','','',0,0,0,'N','',0,'O'
FROM OWOR T1
--inner join wor1 t2 on t1.docentry=t2.docentry
inner join oitm t3 on '" & par_itemcode & "' =t3.itemcode
inner join itm1 t4 on '" & par_itemcode & "' =t4.itemcode
WHERE (T1.STATUS ='P' OR T1.STATUS ='R') AND T1.DOCNUM=" & par_docnum & " and t4.pricelist=2 

group by t1.docentry, T1.PLANNEDQTY,T1.[STARTDate], T1.[DueDate],T1.STATUS,T3.itemname,T4.PRICE,t3.dfltwh"

        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()


    End Sub

    '    Sub aggiorna_riga(PAR_ITEMCODE As String, PAR_DOCNUM As Integer, PAR_VISORDER As Integer, PAR_LINENUM As Integer, PAR_QUANTITà As String, PAR_ATTREZZAGGIO As String, PAR_MAGAZZINO As String, PAR_ITEMTYPE As String, PAR_TRASFERITO As String, PAR_DA_TRASFERIRE As String, PAR_STATO As String)
    '        Dim Cnn3 As New SqlConnection
    '        Cnn3.ConnectionString = Homepage.sap_tirelli
    '        Cnn3.Open()

    '        Dim CMD_SAP_3 As New SqlCommand

    '        CMD_SAP_3.Connection = Cnn3

    '        CMD_SAP_3.CommandText = "UPDATE t2 
    'SET t2.visorder=9998

    'FROM OWOR T1 inner join wor1 t2 on t1.docentry=t2.docentry  
    'inner join oitm t3 on '" & PAR_ITEMCODE & "' =t3.itemcode
    'inner join itm1 t4 on '" & PAR_ITEMCODE & "' =t4.itemcode
    'WHERE (T1.STATUS ='P' OR T1.STATUS ='R') AND T1.DOCNUM=" & PAR_DOCNUM & " and t4.pricelist=2 and t2.visorder=" & PAR_VISORDER & ""

    '        CMD_SAP_3.ExecuteNonQuery()


    '        CMD_SAP_3.CommandText = "UPDATE t2 
    'SET t2.LineNum=" & PAR_LINENUM & ",
    't2.ItemCode= '" & PAR_ITEMCODE & "',
    ' t2.BaseQty= " & Replace(PAR_QUANTITà, ",", ".") & "/T2.PLANNEDQTY,
    't2.AdditQty = " & Replace(PAR_ATTREZZAGGIO, ",", ".") & ",
    't2.PlannedQty=" & Replace(PAR_QUANTITà, ",", ".") & ",
    't2.wareHouse='" & PAR_MAGAZZINO & "',
    ' t2.VisOrder=" & PAR_VISORDER & ",
    't2.ItemType=" & PAR_ITEMTYPE & ",
    ' t2.ResAlloc= CASE WHEN " & PAR_ITEMTYPE & "='290' then 'F' else null end,
    't2.U_PRG_WIP_QtaSpedita= CASE WHEN " & PAR_ITEMTYPE & "='290' THEN 0 ELSE " & Replace(PAR_TRASFERITO, ",", ".") & " END,
    't2.U_PRG_WIP_QtaDaTrasf= CASE WHEN " & PAR_ITEMTYPE & "='290' THEN 0 ELSE " & Replace(PAR_DA_TRASFERIRE, ",", ".") & " END,
    't2.U_Stato_lavorazione='" & PAR_STATO & "',
    't2.StartDate=T1.[STARTDate],
    't2.EndDate= T1.[DueDate],
    't2.Status=T1.STATUS,
    't2.ItemName=t3.itemname,
    't2.u_ubimag=t3.u_ubicazione,
    't2.u_prezzolis=t4.price


    'FROM OWOR T1 inner join wor1 t2 on t1.docentry=t2.docentry and t2.linenum=" & PAR_LINENUM & " 
    'inner join oitm t3 on '" & PAR_ITEMCODE & "' =t3.itemcode
    'inner join itm1 t4 on '" & PAR_ITEMCODE & "' =t4.itemcode
    'WHERE (T1.STATUS ='P' OR T1.STATUS ='R') AND T1.DOCNUM=" & PAR_DOCNUM & " and t4.pricelist=2
    '"

    '        CMD_SAP_3.ExecuteNonQuery()
    '        Cnn3.Close()


    '    End Sub



    Private Sub Button7_Click_1(sender As Object, e As EventArgs)
        Dim i As Integer = 0
        Do While i < num_righe_cancellate
            MsgBox(Righe_cancellate(i).Codice_riga)
            MsgBox(Righe_cancellate(i).visorder)
            MsgBox(Righe_cancellate(i).linenum)

            i = i + 1
        Loop

    End Sub


    Sub ripara_confermati(par_itemcode As String)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "update t41 set t41.iscommited=t40.import_confermati
from
(
Select t30.itemcode, sum(t30.IMPORT_CONFERMATI) as 'Import_confermati', T30.WHSCODE
from
(
SELECT t20.itemcode,T20.CONFERMATI, t20.MAG, T21.WHSCODE, CASE WHEN T21.WHSCODE=t20.MAG THEN T20.CONFERMATI ELSE 0 END AS 'IMPORT_CONFERMATI'
FROM
(
SELECT t10.itemcode,sum(t10.confermati) AS 'CONFERMATI', t10.MAG
FROM
(
SELECT T0.[ItemCode], SUM(T0.[PlannedQty]) AS 'CONFERMATI', T0.[wareHouse] AS 'mag'
 FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE (T1.[STATUS] ='P' OR  T1.[STATUS] ='R') AND T1.[CmpltQty]< T1.[PlannedQty] 
GROUP BY T0.[ItemCode],T0.[wareHouse]

UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.[WhsCode] 
FROM RDR1 T0  INNER JOIN ORDR T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T1.[DocStatus] ='O' AND T0.[OpenQty]>0
GROUP BY 
T0.[ItemCode],  T0.[WhsCode]
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.FROMWHSCOD 
FROM WTQ1 T0  INNER JOIN OWTQ T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T0.[OpenQty] >0 AND  T1.[DocStatus] ='O'
GROUP BY 
T0.[ItemCode], T0.FROMWHSCOD
)
AS T10
where t10.itemcode='" & par_itemcode & "'
group by t10.itemcode, t10.MAG
)
AS T20 LEFT JOIN OITW T21 ON T20.ITEMCODE=T21.ITEMCODE

)
as t30
where t30.itemcode='" & par_itemcode & "'
group by t30.itemcode, T30.WHSCODE
)
as t40 inner join oitw t41 on t41.itemcode=t40.itemcode and t41.whscode=t40.whscode
where t40.itemcode='" & par_itemcode & "'"

        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()

    End Sub

    Sub ripara_confermati_PARTENDO_DA_OITW(par_itemcode As String)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "UPDATE T21 SET T21.ISCOMMITED=T20.cONFERMATI_VERI
FROM
(
select T0.ITEMCODE,T0.WHSCODE, T0.ISCOMMITED, COALESCE(A.CONFERMATI,0) AS 'cONFERMATI_VERI'
from oitw t0

LEFT JOIN (
SELECT t10.itemcode,sum(t10.confermati) AS 'CONFERMATI', t10.MAG
FROM
(
SELECT T0.[ItemCode], SUM(T0.[PlannedQty]) AS 'CONFERMATI', T0.[wareHouse] AS 'mag'
 FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE (T1.[STATUS] ='P' OR  T1.[STATUS] ='R') AND T1.[CmpltQty]< T1.[PlannedQty] 
GROUP BY T0.[ItemCode],T0.[wareHouse]

UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.[WhsCode] 
FROM RDR1 T0  INNER JOIN ORDR T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T1.[DocStatus] ='O' AND T0.[OpenQty]>0
GROUP BY 
T0.[ItemCode],  T0.[WhsCode]
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.FROMWHSCOD 
FROM WTQ1 T0  INNER JOIN OWTQ T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T0.[OpenQty] >0 AND  T1.[DocStatus] ='O'
GROUP BY 
T0.[ItemCode], T0.FROMWHSCOD
)
AS T10
where t10.itemcode='" & par_itemcode & "'
group by t10.itemcode, t10.MAG
)
A ON A.ITEMCODE=T0.ITEMCODE AND T0.WHSCODE=a.MAG

where t0.itemcode='" & par_itemcode & "' and t0.iscommited>0
)
AS T20 INNER JOIN OITW T21 ON T21.ITEMCODE=T20.ItemCode AND T21.WhsCode=T20.WHSCODE"

        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()

    End Sub

    Sub AWOR(par_odp As String, par_utente As String)
        Try
            Using Cnn As New SqlConnection(Homepage.sap_tirelli)
                Cnn.Open()

                Dim query As String = "
                INSERT INTO AWOR (
                    DOCENTRY, UPDATEDATE, LOGINSTANC, DocNum, Series, ItemCode, Status, Type, PlannedQty, CmpltQty, 
                    RjctQty, PostDate, DueDate, OriginAbs, OriginNum, OriginType, UserSign, Comments, CloseDate, 
                    RlsDate, CardCode, Warehouse, Uom, LineDirty, JrnlMemo, TransId, CreateDate, Printed, OcrCode, 
                    PIndicator, OcrCode2, OcrCode3, OcrCode4, OcrCode5, SeqCode, Serial, SeriesStr, SubStr, Project, 
                    SupplCode, UomEntry, PickRmrk, SysCloseDt, SysCloseTm, CloseVerNm, StartDate, ObjType, ProdName, 
                    Priority, RouDatCalc, UpdAlloc, CreateTS, UpdateTS, VersionNum, AtcEntry, AsChild, LinkToObj, 
                    ProcItms, U_UTILIZZ, U_PRODUZIONE, U_Totcosto, U_PRG_AZS_Terzista, U_PRG_AZS_RdrLineNum, 
                    U_PRG_AZS_FromDate, U_PRG_AZS_FromHour, U_PRG_CLV_Fattibil, U_PRG_TOTCOSTO, U_STAMPATO, U_MATRIC, 
                    U_PRG_AZS_Commessa, U_Primadatadiconsegna, U_Consumomediomensile, U_Permag, U_LPONE, U_Inventario, 
                    U_ODPPadre, U_Distintabase, U_Aggiornaprezzo, U_Collaudatore, U_Elettrico, U_Assemblatore, 
                    U_Lavorazione, U_Lavorazione_in_corso, U_Lavoratore, U_Data_ora_inizio, U_Data_ora_fine, U_Disegno, 
                    U_Fase, U_Stato, U_PRG_AZS_CreatedBy, U_PRG_WMS_Exp, U_PRG_WMS_ExpDate, U_Data_cons_MES, 
                    U_Priorita_MES
                )
                SELECT 
                    t0.DOCENTRY, GETDATE(), COALESCE(MAX(t1.LOGINSTANC), 0) + 1, T0.DocNum, T0.Series, T0.ItemCode, 
                    T0.Status, T0.Type, T0.PlannedQty, T0.CmpltQty, T0.RjctQty, T0.PostDate, T0.DueDate, T0.OriginAbs, 
                    T0.OriginNum, T0.OriginType, @UserSign, T0.Comments, T0.CloseDate, T0.RlsDate, T0.CardCode, 
                    T0.Warehouse, T0.Uom, COALESCE(MAX(T1.LINEDIRTY), 0) + 1, T0.JrnlMemo, T0.TransId, T0.CreateDate, 
                    T0.Printed, T0.OcrCode, T0.PIndicator, T0.OcrCode2, T0.OcrCode3, T0.OcrCode4, T0.OcrCode5, T0.SeqCode, 
                    T0.Serial, T0.SeriesStr, T0.SubStr, T0.Project, T0.SupplCode, T0.UomEntry, T0.PickRmrk, T0.SysCloseDt, 
                    T0.SysCloseTm, T0.CloseVerNm, T0.StartDate, T0.ObjType, T0.ProdName, T0.Priority, T0.RouDatCalc, 
                    T0.UpdAlloc, T0.CreateTS, FORMAT(GETDATE(), 'HHmmss'), T0.VersionNum, T0.AtcEntry, T0.AsChild, 
                    T0.LinkToObj, T0.ProcItms, T0.U_UTILIZZ, T0.U_PRODUZIONE, T0.U_Totcosto, T0.U_PRG_AZS_Terzista, 
                    T0.U_PRG_AZS_RdrLineNum, T0.U_PRG_AZS_FromDate, T0.U_PRG_AZS_FromHour, T0.U_PRG_CLV_Fattibil, 
                    T0.U_PRG_TOTCOSTO, T0.U_STAMPATO, T0.U_MATRIC, T0.U_PRG_AZS_Commessa, T0.U_Primadatadiconsegna, 
                    T0.U_Consumomediomensile, T0.U_Permag, T0.U_LPONE, T0.U_Inventario, T0.U_ODPPadre, T0.U_Distintabase, 
                    T0.U_Aggiornaprezzo, T0.U_Collaudatore, T0.U_Elettrico, T0.U_Assemblatore, T0.U_Lavorazione, 
                    T0.U_Lavorazione_in_corso, T0.U_Lavoratore, T0.U_Data_ora_inizio, T0.U_Data_ora_fine, T0.U_Disegno, 
                    T0.U_Fase, T0.U_Stato, T0.U_PRG_AZS_CreatedBy, T0.U_PRG_WMS_Exp, T0.U_PRG_WMS_ExpDate, 
                    T0.U_Data_cons_MES, T0.U_Priorita_MES
                FROM OWOR T0
                LEFT JOIN AWOR T1 ON T0.DOCENTRY = T1.DOCENTRY
                WHERE T0.DOCNUM = @DocNum
                GROUP BY t0.DOCENTRY, T0.DocNum, T0.Series, T0.ItemCode, T0.Status, T0.Type, T0.PlannedQty, 
                         T0.CmpltQty, T0.RjctQty, T0.PostDate, T0.DueDate, T0.OriginAbs, T0.OriginNum, T0.OriginType, 
                         T0.Comments, T0.CloseDate, T0.RlsDate, T0.CardCode, T0.Warehouse, T0.Uom, T0.JrnlMemo, 
                         T0.TransId, T0.CreateDate, T0.Printed, T0.OcrCode, T0.PIndicator, T0.OcrCode2, T0.OcrCode3, 
                         T0.OcrCode4, T0.OcrCode5, T0.SeqCode, T0.Serial, T0.SeriesStr, T0.SubStr, T0.Project, 
                         T0.SupplCode, T0.UomEntry, T0.PickRmrk, T0.SysCloseDt, T0.SysCloseTm, T0.CloseVerNm, 
                         T0.StartDate, T0.ObjType, T0.ProdName, T0.Priority, T0.RouDatCalc, T0.UpdAlloc, T0.CreateTS, 
                         T0.VersionNum, T0.AtcEntry, T0.AsChild, T0.LinkToObj, T0.ProcItms, T0.U_UTILIZZ, T0.U_PRODUZIONE, 
                         T0.U_Totcosto, T0.U_PRG_AZS_Terzista, T0.U_PRG_AZS_RdrLineNum, T0.U_PRG_AZS_FromDate, 
                         T0.U_PRG_AZS_FromHour, T0.U_PRG_CLV_Fattibil, T0.U_PRG_TOTCOSTO, T0.U_STAMPATO, T0.U_MATRIC,t0.U_PRG_AZS_Commessa,T0.U_Primadatadiconsegna, 
                    T0.U_Consumomediomensile, T0.U_Permag, T0.U_LPONE, T0.U_Inventario, T0.U_ODPPadre, T0.U_Distintabase, 
                    T0.U_Aggiornaprezzo, T0.U_Collaudatore, T0.U_Elettrico, T0.U_Assemblatore, T0.U_Lavorazione, 
                    T0.U_Lavorazione_in_corso, T0.U_Lavoratore, T0.U_Data_ora_inizio, T0.U_Data_ora_fine, T0.U_Disegno, 
                    T0.U_Fase, T0.U_Stato, T0.U_PRG_AZS_CreatedBy, T0.U_PRG_WMS_Exp, T0.U_PRG_WMS_ExpDate, 
                    T0.U_Data_cons_MES, T0.U_Priorita_MES"

                Using Cmd_SAP As New SqlCommand(query, Cnn)
                    Cmd_SAP.Parameters.AddWithValue("@DocNum", par_odp)
                    Cmd_SAP.Parameters.AddWithValue("@UserSign", par_utente)
                    Cmd_SAP.ExecuteNonQuery()
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Errore: " & ex.Message)
        End Try
    End Sub

    Sub AWO1(par_docnum As Integer)


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "DECLARE @loginstanc AS INTEGER
DECLARE @docentry AS INTEGER

-- Ottenere il massimo LogInstanc tra OWOR e AWO1, quindi aggiungere 1
SELECT @loginstanc = COALESCE(MAX(LogInstanc), 0) 
FROM (
    SELECT MAX(COALESCE(AWOr.LogInstanc, 0)) AS LogInstanc FROM AWOr 
    WHERE AWOr.DocEntry IN (SELECT DocEntry FROM OWOR WHERE DOCNUM = '" & par_docnum & "')
    
    UNION ALL
    
    SELECT MAX(COALESCE(OWOR.LogInstanc, 0)) AS LogInstanc FROM OWOR 
    WHERE DOCNUM = '" & par_docnum & "'
) AS DerivedTable;

-- Ottenere il DocEntry associato al DOCNUM
SELECT @docentry = MAX(DocEntry) FROM OWOR WHERE DOCNUM = '" & par_docnum & "';

-- Eliminare eventuali record duplicati prima di inserire
DELETE FROM AWO1 WHERE DocEntry = @docentry AND LogInstanc = @loginstanc;

-- Inserire i nuovi dati con il LogInstanc aggiornato
INSERT INTO AWO1 (
    DOCENTRY, LineNum, ItemCode, BaseQty, PlannedQty, IssuedQty, IssueType, wareHouse, VisOrder, 
    WipActCode, CompTotal, OcrCode, OcrCode2, OcrCode3, OcrCode4, OcrCode5, LocCode, LogInstanc, 
    Project, UomEntry, UomCode, ItemType, AdditQty, LineText, PickStatus, PickQty, PickIdNo, 
    ReleaseQty, ResAlloc, StartDate, EndDate, StageId, BaseQtyNum, BaseQtyDen, ReqDays, RtCalcProp, 
    Status, ItemName, AlwProcDoc, PoDocType, PoDocNum, PoDocEntry, PoLineNum, PoQuantity, 
    U_TEMPOME, U_UBIMAG, U_Prezzolis, U_CodDis, U_PRG_AZS_Terzista, U_PRG_AZS_StatoAv, 
    U_PRG_AZS_PhanFat, U_PRG_CLV_Ris_Orig, U_PRG_CLV_Qta_Trasf, U_DisponiibileTOT, 
    U_PRG_WIP_QtaSpedita, U_PRG_WIP_QtaDaTrasf, U_Data_ora_inizio_fase, U_Data_ora_fine_fase, 
    U_Dipendente, U_Ordinato_TOT, U_Confermato_TOT, U_PRG_WIP_QtaRichMagAuto, U_PRG_WMS_Exp, 
    U_PRG_WMS_ExpDate, U_PRG_WMS_MdMovQty, U_Stato_lavorazione, U_Mag_01, U_Mag_fer
)
SELECT 
    T0.DocEntry, T1.LineNum, T1.ItemCode, T1.BaseQty, T1.PlannedQty, T1.IssuedQty, T1.IssueType, 
    T1.wareHouse, T1.VisOrder, T1.WipActCode, T1.CompTotal, T1.OcrCode, T1.OcrCode2, T1.OcrCode3, 
    T1.OcrCode4, T1.OcrCode5, T1.LocCode, @loginstanc, T1.Project, T1.UomEntry, T1.UomCode, 
    T1.ItemType, T1.AdditQty, T1.LineText, T1.PickStatus, T1.PickQty, T1.PickIdNo, T1.ReleaseQty, 
    T1.ResAlloc, T1.StartDate, T1.EndDate, T1.StageId, T1.BaseQtyNum, T1.BaseQtyDen, T1.ReqDays, 
    T1.RtCalcProp, T1.Status, T1.ItemName, T1.AlwProcDoc, T1.PoDocType, T1.PoDocNum, T1.PoDocEntry, 
    T1.PoLineNum, T1.PoQuantity, T1.U_TEMPOME, T1.U_UBIMAG, T1.U_Prezzolis, T1.U_CodDis, 
    T1.U_PRG_AZS_Terzista, T1.U_PRG_AZS_StatoAv, T1.U_PRG_AZS_PhanFat, T1.U_PRG_CLV_Ris_Orig, 
    T1.U_PRG_CLV_Qta_Trasf, T1.U_DisponiibileTOT, T1.U_PRG_WIP_QtaSpedita, T1.U_PRG_WIP_QtaDaTrasf, 
    T1.U_Data_ora_inizio_fase, T1.U_Data_ora_fine_fase, T1.U_Dipendente, T1.U_Ordinato_TOT, 
    T1.U_Confermato_TOT, T1.U_PRG_WIP_QtaRichMagAuto, T1.U_PRG_WMS_Exp, T1.U_PRG_WMS_ExpDate, 
    T1.U_PRG_WMS_MdMovQty, T1.U_Stato_lavorazione, T1.U_Mag_01, T1.U_Mag_fer
FROM OWOR T0 
INNER JOIN WOR1 T1 ON T0.DocEntry = T1.DocEntry
WHERE T0.DOCNUM = '" & par_docnum & "'
"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub

    Sub cancella_righe_odp(par_numero_odp As Integer)

        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "delete t1 from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry 
where t0.docnum=" & par_numero_odp & " "

        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()

    End Sub



    Sub cancella_righe_DB(par_numero_odp As Integer)
        c = 0
        Do While c < num_righe_cancellate
            Dim Cnn3 As New SqlConnection
            Cnn3.ConnectionString = Homepage.sap_tirelli
            Cnn3.Open()

            Dim CMD_SAP_3 As New SqlCommand

            CMD_SAP_3.Connection = Cnn3


            CMD_SAP_3.CommandText = "delete t1 from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry 
where t0.docnum=" & par_numero_odp & " and t1.itemcode='" & Righe_cancellate(c).Codice_riga & "' and  t1.linenum='" & Righe_cancellate(c).linenum & "'"

            CMD_SAP_3.ExecuteNonQuery()
            Cnn3.Close()
            c = c + 1
        Loop
        c = 0
    End Sub



    Private Sub DatiAnagraficiArticoloToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DatiAnagraficiArticoloToolStripMenuItem.Click

        Magazzino.Codice_SAP = DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Codice").Value

        ' Ripristina la finestra se è minimizzata
        If Magazzino.WindowState = FormWindowState.Minimized Then
            Magazzino.WindowState = FormWindowState.Normal
        End If

        ' Porta la finestra in primo piano
        Magazzino.BringToFront()
        Magazzino.Activate()

        ' Ripristina la finestra se è minimizzata
        If Magazzino.WindowState = FormWindowState.Minimized Then
            Magazzino.WindowState = FormWindowState.Normal
        End If

        ' Porta la finestra in primo piano
        Magazzino.BringToFront()
        Magazzino.Activate()

        Magazzino.Show()
        Magazzino.TextBox2.Text = Magazzino.Codice_SAP
        Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)

    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        'DataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect
        '  visorder_selezionato = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Visorder").Value
        riga_selezionata = e.RowIndex

        If riga_selezionata = 0 Then
            Button3.Visible = False
        Else
            Button3.Visible = True
        End If

        'If riga_selezionata >= DataGridView1.Rows.Count - 2 Then
        '    Button4.Visible = False
        'Else
        '    Button4.Visible = True
        'End If

        'If e.ColumnIndex = DataGridView1.Columns.IndexOf(ODP) Then
        '    Dim new_form_odp_form = New ODP_Form
        '    new_form_odp_form.docnum_odp = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ODP").Value
        '    new_form_odp_form.Show()
        '    new_form_odp_form.inizializza_form(DataGridView1.Rows(e.RowIndex).Cells(columnName:="ODP").Value)







        'End If

        'If e.ColumnIndex = DataGridView1.Columns.IndexOf(Disegno) Then


        '    Magazzino.visualizza_disegno(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disegno").Value)



        'End If

    End Sub

    Sub Trova_ID_modifiche_odp()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select max(id)+1 as 'ID' from [Tirelli_40].[dbo].[Modifiche_odp]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                id_modifica = cmd_SAP_reader_2("ID")
            Else
                id_modifica = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub


    Sub inserisci_record_modifica_odp(par_utente As Integer, par_numero_odp As Integer)
        Trova_ID_modifiche_odp()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "
DECLARE @id_modifica AS INTEGER
SELECT @id_modifica = coalesce(MAX(id),0) + 1 FROM [Tirelli_40].[dbo].[Modifiche_odp]

insert into [Tirelli_40].[dbo].[Modifiche_odp] (id,data,ora,id_utente,odp) 
values (@id_modifica,getdate(),convert(varchar, getdate(), 108),'" & par_utente & "','" & par_numero_odp & "')"

        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()
    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Magazzino.visualizza_disegno(Button5.Text)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Richiesta_trasferimento_materiale.Show()
        Richiesta_trasferimento_materiale.docnum_odp = docnum_odp
        Richiesta_trasferimento_materiale.docentry_odp = docentry_odp
        Richiesta_trasferimento_materiale.riempi_datagridview_rt(docnum_odp)

    End Sub



    'End Sub
    Sub Fun_Stampa()
        ' Dopo la rotazione 90°:
        ' - larghezza fisica del rotolo (315) = altezza logica del layout (H)
        ' - lunghezza carta (700) = larghezza logica del layout (W)
        larghezza_scontrino_odp = 700   ' lunghezza carta → diventa W logico
        altezza_scontrino_odp = 315     ' larghezza rotolo → diventa H logico

        Sel_Stampante.AllowSomePages = False
        Sel_Stampante.ShowHelp = False
        Sel_Stampante.Document = Scontrino

        ' PaperSize vuole (larghezza fisica, altezza fisica)
        ' larghezza fisica = 315 (80mm rotolo)
        ' altezza fisica   = 700 (lunghezza scontrino)
        Dim paperSize As New System.Drawing.Printing.PaperSize("Scontrino80mm", 210, 700)

        ' Margini a zero: altrimenti il default (100,100,100,100) = 1 pollice per lato
        ' su un'etichetta da 3.15" lascerebbe solo ~1.15" di area utile → etichetta troncata
        Scontrino.DefaultPageSettings.Margins = New System.Drawing.Printing.Margins(0, 0, 0, 0)
        Scontrino.OriginAtMargins = False

        If preview_scontrino = True Then
            If Homepage.Stampante_Selezionata = False Then
                Dim result As DialogResult = Sel_Stampante.ShowDialog()
                If result = DialogResult.OK Then
                    Homepage.Stampante_Selezionata = True
                    Scontrino.DefaultPageSettings.Landscape = False
                    Scontrino.DefaultPageSettings.PaperSize = paperSize
                    Dim previewDialog As New PrintPreviewDialog()
                    previewDialog.Document = Scontrino
                    previewDialog.ShowDialog()
                End If
            Else
                Scontrino.DefaultPageSettings.Landscape = False
                Scontrino.DefaultPageSettings.PaperSize = paperSize
                Dim previewDialog As New PrintPreviewDialog()
                previewDialog.Document = Scontrino
                previewDialog.ShowDialog()
            End If
        Else
            If Homepage.Stampante_Selezionata = False Then
                Dim result As DialogResult = Sel_Stampante.ShowDialog()
                If result = DialogResult.OK Then
                    Homepage.Stampante_Selezionata = True
                    Scontrino.DefaultPageSettings.Landscape = False
                    Scontrino.DefaultPageSettings.PaperSize = paperSize
                    Scontrino.Print()
                End If
            Else
                Scontrino.DefaultPageSettings.Landscape = False
                Scontrino.DefaultPageSettings.PaperSize = paperSize
                Scontrino.Print()
            End If
        End If
    End Sub

    Private Sub Scontrino_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles Scontrino.PrintPage

        Dim g As Graphics = e.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

        ' ============================================================
        ' ROTAZIONE 90° — foglio fisico 315 wide x 700 tall
        ' Dopo rotazione: piano logico 700 wide x 315 tall
        ' TranslateTransform deve usare l'ALTEZZA fisica (700), non la larghezza (315).
        ' Con 315 il contenuto oltre lx=315 finisce a py negativo (fuori pagina).
        ' ============================================================
        g.TranslateTransform(0, 700)
        g.RotateTransform(-90)

        Dim W As Integer = 700
        Dim H As Integer = 210   ' lato corto ridotto: era 315

        ' --- Penna e brush ---
        Dim penna As New Pen(Color.Black, 1.5)
        Dim brushNero As Brush = Brushes.Black
        Dim brushGrigio As Brush = New SolidBrush(Color.FromArgb(150, 150, 150))

        ' --- Font ---
        Dim fLbl As New Font("Calibri", 6, FontStyle.Italic)
        Dim fVal As New Font("Calibri", 10, FontStyle.Bold)
        Dim fValLg As New Font("Calibri", 14, FontStyle.Bold)
        Dim fValItalic As New Font("Calibri", 9, FontStyle.Italic)
        Dim fNumerone As New Font("Calibri", 38, FontStyle.Bold)
        Dim fOdp As New Font("Calibri", 11, FontStyle.Bold)

        ' --- StringFormat centrato (usato per i valori) ---
        Dim sfCenter As New StringFormat()
        sfCenter.Alignment = StringAlignment.Center
        sfCenter.LineAlignment = StringAlignment.Center

        ' ============================================================
        ' LAYOUT — 5 righe
        ' ============================================================
        Dim y0 As Integer = 0
        Dim y1 As Integer = 48
        Dim y2 As Integer = 84
        Dim y3 As Integer = 121
        Dim y4 As Integer = 153

        Dim h0 As Integer = 48
        Dim h1 As Integer = 36
        Dim h2 As Integer = 37
        Dim h3 As Integer = 32
        Dim h4 As Integer = H - y4   ' = 57

        Dim x1 As Integer = 180
        Dim x2 As Integer = 340
        Dim x3 As Integer = 500
        Dim xA As Integer = 380
        Dim xB As Integer = 540

        ' ============================================================
        ' BORDO ESTERNO
        ' ============================================================
        g.DrawRectangle(penna, 0, 0, W - 1, H - 1)

        ' ============================================================
        ' ROW 0 — COD | ODP | Commessa | Posizione
        ' ============================================================
        g.DrawLine(penna, 0, y1, W, y1)
        g.DrawLine(penna, x1, y0, x1, y1)
        g.DrawLine(penna, x2, y0, x2, y1)
        g.DrawLine(penna, x3, y0, x3, y1)

        g.DrawString("Cod. articolo", fLbl, brushGrigio, New RectangleF(2, y0 + 2, x1 - 4, 12), sfCenter)
        g.DrawString("COD: " & testata_odp_itemcode, fValLg, brushNero, New RectangleF(2, y0 + 14, x1 - 4, h0 - 16), sfCenter)

        g.DrawString("Ordine di produzione", fLbl, brushGrigio, New RectangleF(x1 + 2, y0 + 2, x2 - x1 - 4, 12), sfCenter)
        g.DrawString("ODP : " & testata_odp_docnum, fValLg, brushNero, New RectangleF(x1 + 2, y0 + 14, x2 - x1 - 4, h0 - 16), sfCenter)

        g.DrawString("Commessa / Matricola", fLbl, brushGrigio, New RectangleF(x2 + 2, y0 + 2, x3 - x2 - 4, 12), sfCenter)
        g.DrawString(testata_odp_commessa, fValLg, brushNero, New RectangleF(x2 + 2, y0 + 14, x3 - x2 - 4, h0 - 16), sfCenter)

        g.DrawString("Posizione", fLbl, brushGrigio, New RectangleF(x3 + 2, y0 + 2, W - x3 - 4, 12), sfCenter)
        g.DrawString(numerone, fValLg, brushNero, New RectangleF(x3 + 2, y0 + 14, W - x3 - 4, h0 - 16), sfCenter)

        ' ============================================================
        ' ROW 1 — Descrizione prodotto (larghezza piena)
        ' ============================================================
        g.DrawLine(penna, 0, y2, W, y2)
        g.DrawString("Descrizione prodotto", fLbl, brushGrigio, New RectangleF(4, y1 + 2, W - 8, 12))
        g.DrawString(testata_odp_prodname, fVal, brushNero, New RectangleF(2, y1 + 14, W - 4, h1 - 16), sfCenter)

        ' ============================================================
        ' ROW 2 — Cliente | [Numerone da xB]
        ' ============================================================
        g.DrawLine(penna, 0, y3, xB, y3)
        g.DrawLine(penna, xB, y2, xB, H - 1)   ' bordo sinistro cella numerone
        g.DrawLine(penna, xA, y2, xA, y3)

        g.DrawString("Cliente", fLbl, brushGrigio, New RectangleF(4, y2 + 2, xA - 8, 12))
        g.DrawString(testata_odp_cardname, fVal, brushNero, New RectangleF(2, y2 + 14, xA - 4, h2 - 16), sfCenter)

        ' ============================================================
        ' ROW 3 — Data/Ora | Mag Destinazione
        ' ============================================================
        g.DrawLine(penna, 0, y4, xB, y4)
        g.DrawLine(penna, xA, y3, xA, y4)

        g.DrawString("Data / Ora stampa", fLbl, brushGrigio, New RectangleF(4, y3 + 2, xA - 8, 12))
        g.DrawString(testata_odp_data & " " & TimeOfDay.ToString("HH:mm:ss"), fVal, brushNero, New RectangleF(2, y3 + 14, xA - 4, h3 - 16), sfCenter)

        g.DrawString("Mag Destinazione", fLbl, brushGrigio, New RectangleF(xA + 4, y3 + 2, xB - xA - 8, 12))
        g.DrawString(testata_odp_warehouse, fVal, brushNero, New RectangleF(xA + 2, y3 + 14, xB - xA - 4, h3 - 16), sfCenter)

        ' ============================================================
        ' NUMERONE — ROW 2 + ROW 3 + ROW 4 a destra
        ' ============================================================
        Dim numeroneStr As String
        Select Case testata_odp_u_produzione
            Case "EST" : numeroneStr = "E" & numerone
            Case "INT_SALD" : numeroneStr = "S" & numerone
            Case "B_INT" : numeroneStr = "B" & numerone
            Case "INT" : numeroneStr = "I" & numerone
            Case Else : numeroneStr = numerone
        End Select
        Dim rectNumerone As New RectangleF(xB + 2, y2 + 2, W - xB - 4, H - y2 - 4)
        Dim sfNumerone As New StringFormat()
        sfNumerone.Alignment = StringAlignment.Center
        sfNumerone.LineAlignment = StringAlignment.Center
        g.DrawString(numeroneStr, fNumerone, brushNero, rectNumerone, sfNumerone)

        ' ============================================================
        ' ROW 4 — Desc.Commessa | N°Pz
        ' ============================================================
        g.DrawLine(penna, xA, y4, xB, y4)

        g.DrawString("Descrizione commessa", fLbl, brushGrigio, New RectangleF(4, y4 + 2, xA - 8, 12))
        g.DrawString(testata_odp_Itemname_commessa, New Font("Calibri", 9, FontStyle.Italic), brushNero, New RectangleF(2, y4 + 14, xA - 4, h4 - 16), sfCenter)

        g.DrawString("N° pezzi pianificati", fLbl, brushGrigio, New RectangleF(xA + 4, y4 + 2, xB - xA - 8, 12))
        g.DrawString("N° Pz : " & testata_odp_plannedqty, fVal, brushNero, New RectangleF(xA + 2, y4 + 14, xB - xA - 4, h4 - 16), sfCenter)

        ' --- Pulizia ---
        penna.Dispose()
        fLbl.Dispose()
        fVal.Dispose()
        fValLg.Dispose()
        fValItalic.Dispose()
        fNumerone.Dispose()
        fOdp.Dispose()
        brushGrigio.Dispose()
        sfCenter.Dispose()
        sfNumerone.Dispose()

    End Sub







    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            preview_scontrino = True
        Else
            preview_scontrino = False
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Homepage.Stampante_Selezionata = False
    End Sub



    Private Sub TrasferimentoDiMagazzinoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TrasferimentoDiMagazzinoToolStripMenuItem.Click

        If ottieni_informazioni_odp("Numero", 0, docnum_odp).stato = "R" Then


            Trasferimento_magazzino.docentry_odp = docentry_odp
            Trasferimento_magazzino.docnum_odp = docnum_odp

            Trasferimento_magazzino.Text = "Trasferimento"
            Trasferimento_magazzino.inizializzazione_trasferimento(docentry_odp, 0, "Trasferimento", "ODP")

        Else
            MsgBox("Funzione disponibile solo per ordini rilasciati")
        End If

    End Sub

    Private Sub ODP_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'If Homepage.Centro_di_costo = "BRB01" Then
        '    operazione = "="
        'Else
        '    operazione = "<>"
        'End If

        Inserimento_stati_odp_combobox()
        ' Inserimento_items_combobox_produzione()
        ' Inserimento_items_combobox_fase()
        Inserimento_items_combobox_stato_lav()
        Inserimento_magazzini_righe()
    End Sub



    Private Sub Cmd_Indietro_Click(sender As Object, e As EventArgs) Handles Cmd_Indietro.Click
        inizializza_form((Integer.Parse(TextBox10.Text) - 1).ToString("00000000"))
    End Sub

    Private Sub Cmd_Avanti_Click(sender As Object, e As EventArgs) Handles Cmd_Avanti.Click
        inizializza_form((Integer.Parse(TextBox10.Text) + 1).ToString("00000000"))
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        inizializza_form(TextBox10.Text)
    End Sub

    Private Sub ResoDiMagazzinoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResoDiMagazzinoToolStripMenuItem.Click
        If ottieni_informazioni_odp("Numero", 0, docnum_odp).stato = "R" Then

            Trasferimento_magazzino.docentry_odp = docentry_odp
            Trasferimento_magazzino.docnum_odp = docnum_odp
            Dim magazzino_partenza As String = ""
            Dim magazzino_arrivo As String = ""
            If Form_Entrate_Merci.Trova_business_unit_magazzino(TextBox5.Text) <> 13 Then
                magazzino_partenza = "WIP"
                magazzino_arrivo = "01"
            Else
                magazzino_partenza = "BWIP"
                magazzino_arrivo = "B01"
            End If
            Trasferimento_magazzino.inizializzazione_trasferimento(docentry_odp, 0, "Reso", "ODP")
            Trasferimento_magazzino.Text = "Reso"
        Else
            MsgBox("Funzione disponibile solo per ordini rilasciati")
        End If
    End Sub

    Public Class Dettagliodp
        Public docnum As String
        Public Descrizione As String
        Public stato As String
        Public u_produzione As String
        Public fase As String
        Public commessa As String
        Public quantità As String
        Public itemcode As String
        Public disegno As String
        Public lavorazione As String
        Public u_stato As String
        Public docentry As Integer
        Public u_aggiorna_db As String
        Public postdate As String
        Public startdate As String
        Public duedate As String
        Public nome As String
        Public type As String
        Public warehouse As String
        Public u_utilizz As String
        Public descr As String
        Public numerone As String
        Public q_comp As Integer
        Public Cliente As String
        Public nome_baia As String
        Public Matricola As String
        Public sottocommessa As String
        Public desc_commessa As String


    End Class

    Public Class prima_data_odp
        Public docnum As String
        Public Descrizione As String
        Public stato As String
        Public u_produzione As String
        Public fase As String
        Public commessa As String
        Public quantità As String
        Public itemcode As String
        Public disegno As String
        Public lavorazione As String
        Public u_stato As String
        Public docentry As Integer
        Public u_aggiorna_db As String
        Public postdate As String
        Public startdate As String
        Public duedate As String
        Public nome As String
        Public type As String
        Public warehouse As String
        Public u_utilizz As String
        Public descr As String
        Public numerone As String

    End Class

    Public Class prima_data_oa
        Public docnum As Integer
        Public cardname As Integer
        Public duedate As String



    End Class

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Data_ordine_ValueChanged(sender As Object, e As EventArgs) Handles Data_ordine.ValueChanged

    End Sub





    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = DataGridView2.Columns.IndexOf(Itemcode) Then

                Magazzino.Codice_SAP = DataGridView2.Rows(e.RowIndex).Cells(columnName:="Itemcode").Value
                ' Ripristina la finestra se è minimizzata
                If Magazzino.WindowState = FormWindowState.Minimized Then
                    Magazzino.WindowState = FormWindowState.Normal
                End If

                ' Porta la finestra in primo piano
                Magazzino.BringToFront()
                Magazzino.Activate()

                Magazzino.Show()

                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)

            ElseIf e.ColumnIndex = DataGridView2.Columns.IndexOf(ODP) Then
                Dim new_form_odp_form = New ODP_Form
                new_form_odp_form.docnum_odp = DataGridView2.Rows(e.RowIndex).Cells(columnName:="ODP").Value
                new_form_odp_form.Show()
                new_form_odp_form.inizializza_form(DataGridView2.Rows(e.RowIndex).Cells(columnName:="ODP").Value)

            End If




            If e.ColumnIndex = DataGridView2.Columns.IndexOf(OA) Then


                Form_nuova_offerta.Show()

                Form_nuova_offerta.TextBox10.Text = DataGridView2.Rows(e.RowIndex).Cells(columnName:="OA").Value
                Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                Form_nuova_offerta.inizializzazione_form(DataGridView2.Rows(e.RowIndex).Cells(columnName:="OA").Value, "OPOR", "POR1", "Ordine_acquisto")


            End If




        End If

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        If DataGridView2.Rows(e.RowIndex).Cells(columnName:="Status").Value = "OK" Then


            DataGridView2.Rows(e.RowIndex).Cells(columnName:="Status").Style.BackColor = Color.Lime

        ElseIf DataGridView2.Rows(e.RowIndex).Cells(columnName:="Status").Value = "Trasferibile" Then

            DataGridView2.Rows(e.RowIndex).Cells(columnName:="Status").Style.BackColor = Color.Yellow

        ElseIf DataGridView2.Rows(e.RowIndex).Cells(columnName:="Status").Value = "Approvv" Then

            DataGridView2.Rows(e.RowIndex).Cells(columnName:="Status").Style.BackColor = Color.Orange

        ElseIf DataGridView2.Rows(e.RowIndex).Cells(columnName:="Status").Value = "Da_ordinare" Then

            DataGridView2.Rows(e.RowIndex).Cells(columnName:="Status").Style.BackColor = Color.OrangeRed


        End If

        If DataGridView2.Rows(e.RowIndex).Cells(columnName:="Disp").Value < 0 Then
            DataGridView2.Rows(e.RowIndex).Cells(columnName:="Disp").Style.ForeColor = Color.Red
        End If
    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter

        Dim operazione As String
        If Homepage.Centro_di_costo = "BRB01" Then
            operazione = "="
        Else
            operazione = "<>"
        End If
        DataGridView2.Rows.Clear()
        ODP_Tree.PULISCI_APPOGGIO(Homepage.ID_SALVATO, "ODP_TREE")
        form_Spare_Parts.dettaglio_ODP(DataGridView2, "", TextBox10.Text, 1, TextBox6.Text, 0, "ODP_TREE")
    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Enter

        Dashboard_MU_New.righe_ODP_macchine(DataGridView1, TextBox10.Text)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        'Distinta_base_form.Show()

        'Distinta_base_form.TextBox1.Text = Button11.Text

        Magazzino.Codice_SAP = Button11.Text

        ' Ripristina la finestra se è minimizzata
        If Magazzino.WindowState = FormWindowState.Minimized Then
            Magazzino.WindowState = FormWindowState.Normal
        End If

        ' Porta la finestra in primo piano
        Magazzino.BringToFront()
        Magazzino.Activate()

        ' Ripristina la finestra se è minimizzata
        If Magazzino.WindowState = FormWindowState.Minimized Then
            Magazzino.WindowState = FormWindowState.Normal
        End If

        ' Porta la finestra in primo piano
        Magazzino.BringToFront()
        Magazzino.Activate()

        Magazzino.Show()
        Magazzino.TextBox2.Text = Magazzino.Codice_SAP
        Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' aggiorna_odp()
    End Sub

    Private Sub DataGridView_ODP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellContentClick

    End Sub

    Private Sub DataGridView_ODP_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellLeave

    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        Dim par_datagridview As DataGridView = DataGridView_ODP
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

        ' Aggiungere dati dalla DataGridView al foglio di lavoro
        For row As Integer = 0 To par_datagridview.Rows.Count - 1
            For col As Integer = 0 To par_datagridview.Columns.Count - 1
                ' Imposta il formato della cella come testo
                excelWorksheet.Cells(row + 2, col + 1).NumberFormat = "@"
                ' Inserisce il valore nella cella
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


    Private Sub TextBox4_Leave(sender As Object, e As EventArgs) Handles TextBox4.Leave
        If inizio = False Then
            Dim annullaOperazione As Boolean = False
            Dim valoriOriginali As New Dictionary(Of Integer, Decimal) ' Per ripristinare in caso di errore
            Dim messaggioErrore As String = ""

            ' Primo ciclo: calcoliamo i nuovi valori senza modificarli
            For Each row As DataGridViewRow In DataGridView_ODP.Rows
                If Not row.IsNewRow Then
                    Dim quantitaAttuale As Decimal
                    Dim moltiplicatore As Decimal
                    Dim divisore As Decimal
                    Dim trasferito As Decimal

                    ' Verifica e conversione dei valori
                    If Decimal.TryParse(row.Cells("Quantità").Value.ToString(), quantitaAttuale) AndAlso
                       Decimal.TryParse(TextBox4.Text, moltiplicatore) AndAlso
                       Decimal.TryParse(ottieni_informazioni_odp("Numero", 0, TextBox10.Text).quantità.ToString(), divisore) AndAlso
                       Decimal.TryParse(row.Cells("Trasferito").Value.ToString(), trasferito) AndAlso
                       divisore <> 0 Then

                        ' Calcola nuova quantità
                        Dim nuovaQuantita As Decimal = (quantitaAttuale * moltiplicatore) / divisore
                        Dim nuovoDaTrasf As Decimal = nuovaQuantita - trasferito

                        ' Salviamo il valore originale della quantità
                        valoriOriginali(row.Index) = row.Cells("Quantità").Value

                        ' Se il valore di Da_trasf va sotto zero, annulliamo tutto
                        If nuovoDaTrasf < 0 Then
                            annullaOperazione = True
                            messaggioErrore = $"Errore alla riga {row.Index + 1}: 'Da_trasf' diventerebbe {nuovoDaTrasf}. Operazione annullata."

                            Exit For
                        End If
                    End If
                End If
            Next

            ' Se nessun valore è negativo, applichiamo le modifiche
            If Not annullaOperazione Then
                For Each row As DataGridViewRow In DataGridView_ODP.Rows
                    If Not row.IsNewRow Then
                        Dim quantitaAttuale As Decimal
                        Dim moltiplicatore As Decimal
                        Dim divisore As Decimal
                        Dim trasferito As Decimal

                        If Decimal.TryParse(row.Cells("Quantità").Value.ToString(), quantitaAttuale) AndAlso
                           Decimal.TryParse(TextBox4.Text, moltiplicatore) AndAlso
                           Decimal.TryParse(ottieni_informazioni_odp("Numero", 0, TextBox10.Text).quantità.ToString(), divisore) AndAlso
                           Decimal.TryParse(row.Cells("Trasferito").Value.ToString(), trasferito) AndAlso
                           divisore <> 0 Then

                            ' Calcola e assegna i nuovi valori
                            row.Cells("Quantità").Value = (quantitaAttuale * moltiplicatore) / divisore
                            row.Cells("Da_trasf").Value = row.Cells("Quantità").Value - trasferito
                            If row.Cells("Da_trasf").Value > 0 And row.Cells("MAg").Value = "WIP" Then
                                row.Cells("MAg").Value = "01"
                            End If
                        End If
                    End If
                Next
            Else
                ' Se c'è stato un errore, ripristiniamo i valori originali
                For Each row As DataGridViewRow In DataGridView_ODP.Rows
                    If valoriOriginali.ContainsKey(row.Index) Then
                        row.Cells("Quantità").Value = valoriOriginali(row.Index)
                    End If
                Next

                ' Mostra un messaggio di errore con il dettaglio della riga problematica
                MessageBox.Show(messaggioErrore, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                TextBox4.Text = ottieni_informazioni_odp("Numero", 0, TextBox10.Text).quantità.ToString()
            End If
        End If
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Try


            If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato_").Value = "S" Then
                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
            ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato_").Value = "O" Then
                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Yellow
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Function trova_esistenza_odp(par_docnum As String)
        Dim esiste As Boolean = False

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1



        CMD_SAP_2.CommandText = "select *
from
OPENQUERY([AS400], '
    SELECT *
    FROM S786FAD1.TIR90VIS.JGALODP
 where numodp=''" & par_docnum & "''
') AS t10"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            esiste = True
        Else
            esiste = False

        End If



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        Return esiste

    End Function
End Class