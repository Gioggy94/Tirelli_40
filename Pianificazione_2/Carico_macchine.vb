Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Runtime.InteropServices
Imports AxFOXITREADERLib
Imports Microsoft.Office.Interop
Public Class Carico_macchine
    Public delay_max As Integer
    Public data_consegna As Date
    Public lavorazione As Integer
    Private filtro_nesting As String

    Public Elenco_macchinari(1000) As String

    Public filtro_odp As String
    Public filtro_codice As String
    Public filtro_descrizione As String
    Public filtro_stato As String
    Public filtro_disegno As String
    Public filtro_commessa As String
    Public filtro_cliente As String
    Public filtro_mat_prima As String


    Public filtro_stato_completamento As String
    Public filtro_R00554 As String
    Public filtro_R00551 As String
    Public filtro_R00527 As String
    Public filtro_R00539 As String
    Public filtro_R00563 As String
    Public filtro_R00562 As String
    Public filtro_R00561 As String
    Public filtro_R00503 As String
    Public filtro_R00506 As String
    Public filtro_R00504 As String
    Public filtro_R00505 As String
    Public filtro_R00526 As String
    Public filtro_R00564 As String
    Public filtro_R00502 As String
    Public filtro_R00540 As String
    Public filtro_R00587 As String
    Public filtro_R00598 As String
    Public filtro_R00599 As String
    Public filtro_R00600 As String
    Private filtro_R00550 As String
    Private filtro_R00613 As String
    Private filtro_R00610 As String
    Private filtro_R00572 As String
    Public codice_pic As String

    Sub carico_macchine(par_odp As String, par_codice As String, par_descrizione As String, par_stato As String, par_disegno As String, par_commessa As String, par_cliente As String, par_materia_prima As String)

        Dim MAnodopera_tot As Integer = 0
        Dim Taglio_tot As Integer = 0
        Dim Tornio_tot As Integer = 0
        Dim Fresa_tot As Integer = 0

        DataGridView1.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand



        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then


            CMD_SAP_2.CommandText = "
declare @taglio as INTEGER
declare @tornio as INTEGER
declare @fresa as INTEGER

	
			
set @taglio = " & 999 & "
set @tornio = " & 999 & "
set @fresa = " & 999 & "

select *
from
(

select  t40.Seleziona, t40.docnum, t40.U_Lavorazione, t40.status, t40.postdate, t40.duedate,getdate()+(ROW_NUMBER() OVER(ORDER BY t40.priorità asc)/t50.ODP_chiusi_al_gg) as 'Cons', DATEDIFF(DD,T40.DUEDATE,GETDATE()) AS 'ANT/RIT',t40.stato_mat, t40.Arrivo_MAT,case when t62.MAt_prima is null then '' else t62.MAt_prima end as 'Mat_prima' , t40.Priorità, t40.ItemCode, t40.itemname, t40.Disegno, t40.[Famiglia disegno], t40.nesting, t40.disp,t40.stato_completamento,t40.PlannedQty,t40.U_PRG_AZS_Commessa,T40.WAREHOUSE, case when t61.U_Final_customer_name is null and t63.U_Final_customer_name is null then t64.custmrName when t61.U_Final_customer_name is null then  t63.U_Final_customer_name else t61.U_Final_customer_name  end as 'U_Final_customer_name'

,case when T60.R00554 is null then '' else T60.R00554 end as 'R00554'
,case when T60.R00550 is null then '' else T60.R00550 end as 'R00550'
,case when T60.R00551 is null then '' else T60.R00551 END AS 'R00551'
,case when T60.R00527 is null then '' else  T60.R00527 end as 'R00527'
,case when T60.R00539 is null then '' else  T60.R00539 end as 'R00539'
,case when T60.R00563 is null then '' else  T60.R00563 end as 'R00563'
,case when T60.R00562 is null then '' else  T60.R00562 end as 'R00562'
,case when T60.R00561 is null then '' else  T60.R00561 end as 'R00561'
,case when T60.R00503 is null then '' else  T60.R00503 end as 'R00503'
,case when T60.R00506 is null then '' else  T60.R00506 end as 'R00506'
,case when T60.R00504 is null then '' else  T60.R00504 end as 'R00504'
,case when T60.R00505 is null then '' else  T60.R00505 end as 'R00505'
,case when T60.R00526 is null then '' else  T60.R00526 end as 'R00526'
,case when T60.R00564 is null then '' else  T60.R00564 end as 'R00564'
,case when T60.R00502 is null then '' else  T60.R00502 end as 'R00502'
,case when T60.R00540 is null then '' else  T60.R00540 end as 'R00540'
,case when T60.R00572 is null then '' else  T60.R00572 end as 'R00572'
,case when T60.R00587 is null then '' else  T60.R00587 end as 'R00587'
,case when T60.R00598 is null then '' else  T60.R00598 end as 'R00598'
,case when T60.R00599 is null then '' else  T60.R00599 end as 'R00599'
,case when T60.R00600 is null then '' else  T60.R00600 end as 'R00600'
,case when T60.R00610 is null then '' else  T60.R00610 end as 'R00610'
,case when T60.R00611 is null then '' else  T60.R00611 end as 'R00611'
,case when T60.R00613 is null then '' else  T60.R00613 end as 'R00613'

from
(
select  t10.Seleziona, t10.docnum, coalesce(t10.U_Lavorazione,0) as 'U_lavorazione', t10.status, t10.postdate,T10.DUEDATE, t10.Cons, t10.[ant/rit],c.stato_mat, c.Arrivo_MAT,case when t10.u_lavorazione=0 then t10.Priorità else min(t10.priorità) OVER(PARTITION BY t10.u_lavorazione) end as'Priorità',t10.ItemCode, t10.itemname, t10.Disegno, t10.[Famiglia disegno], t10.nesting, B.disp,t10.stato_completamento, t10.PlannedQty,t10.U_PRG_AZS_Commessa,T10.WAREHOUSE
from
(

select 'sel' as 'Seleziona', t0.docnum, t0.PlannedQty, t0.u_lavorazione,t0.status, t0.PostDate,T0.DUEDATE, getdate() as 'Cons', 5 as 'ant/rit','MAT' as 'MAt','arrivoMAT' as 'ArrivoMAT',case when t0.U_Fase = 'urgenza_MFC' then 0 WHEN t0.U_Fase = 'urgenza_collaudo' then 1 else ROW_NUMBER() OVER(ORDER BY t0.Status DESC, T0.[duedate] ASC) end as 'Priorità', t0.ItemCode, t1.itemname, case when t1.u_disegno is null then '' else t1.u_disegno end  as 'Disegno',

case when t1.u_famiglia_disegno is null then '' else t1.u_famiglia_disegno end as 'Famiglia disegno',

COUNT(t1.u_famiglia_disegno) OVER(PARTITION BY t1.u_famiglia_disegno) as 'Nesting', t0.U_PRG_AZS_Commessa,T0.WAREHOUSE, 'DISP' as 'DISP',

case when t0.u_stato is null then '' else t0.u_stato end as 'stato_completamento'

from owor t0 left join oitm t1 on t0.itemcode=t1.itemcode 
where substring(t0.u_produzione,1,3)='INT' and (t0.Status='R' or t0.Status='P') 

)
as t10 left join 
(select t0.docnum, sum(t11.onhand-t11.iscommited+t11.onorder) as 'Disp' from owor t0 inner join oitw t11 on t11.itemcode=t0.itemcode where substring(t0.u_produzione,1,3)='INT' and (t0.status='P' or t0.status='R') group by t0.docnum) as B on B.DocNum=t10.docnum

left join (select t30.docnum,   case when t30.Da_trasferire is null or t30.da_trasferire=0 then 'OK' when t30.non_trasferibili is null or t30.non_trasferibili=0 then 'TRASF' else 'IN_APPR'  end as 'stato_mat', case when t30.Shipdate_oa_figlio is null then t30.Cons_odp_figlio when t30.Cons_odp_figlio is null then t30.Shipdate_oa_figlio end as 'Arrivo_MAT'
from
(
select t20.docnum, sum(case when t20.u_prg_wip_qtadatrasf>0 then 1 else 0 end) as 'Da_trasferire',sum(case when t20.u_prg_wip_qtadatrasf>t20.mag then 1 else 0 end) as 'non_Trasferibili', min(T20.Shipdate) as 'Shipdate_oa_figlio', min(T20.cons_odp) as 'Cons_odp_figlio'
from
(
select t10.docnum, t10.itemcode, t10.u_prg_wip_qtadatrasf, t10.Mag, min (t11.shipdate) as 'Shipdate', min (case when substring (t12.u_produzione,1,3)='INT' then t12.u_data_cons_mes else t12.duedate end ) as 'Cons_ODP'
from
(
select t0.docnum, t1.itemcode, t1.u_prg_wip_qtadatrasf, sum(case when t2.onhand is null then 0 else t2.onhand end) as 'Mag'
from owor t0 left join wor1 t1 on t0.docentry=t1.docentry
left join oitw t2 on t2.itemcode=t1.itemcode and  t1.u_prg_wip_qtadatrasf >0 and t2.whscode<>'WIP' and t2.whscode<>'Clavter'
where substring(t0.u_produzione,1,3)='INT' and (t0.Status='R' or t0.Status='P') and (SUBSTRING(t1.itemcode,1,1)='C' or SUBSTRING(t1.itemcode,1,1)='D' or SUBSTRING(t1.itemcode,1,1)='0')
group by t0.docnum, t1.itemcode, t1.u_prg_wip_qtadatrasf
)
as t10 left join por1 t11 on t11.itemcode=t10.itemcode and t10.mag<t10.u_prg_wip_qtadatrasf and t11.openqty>0
left join owor t12 on (t12.status ='P' or t12.status ='R') and t12.itemcode=t10.itemcode and t10.mag<t10.u_prg_wip_qtadatrasf
group by t10.docnum, t10.itemcode, t10.u_prg_wip_qtadatrasf, t10.Mag
)
as t20
group by t20.docnum
)
as t30
) as C on c.DocNum=t10.docnum

)
as t40
left JOIN 
(select t30.docnum
,case when t30.R00554=1 then 'O' when t30.R00554=1001 then 'C' else'' end as 'R00554'
,case when t30.R00550=1 then 'O' when t30.R00550=1001 then 'C' else'' end as 'R00550'
,case when t30.R00551=1 then 'O' when t30.R00551=1001 then 'C' else'' end as 'R00551'
,case when t30.R00527=1 then 'O' when t30.R00527=1001 then 'C' else'' end as 'R00527'
,case when t30.R00539=1 then 'O' when t30.R00539=1001 then 'C' else'' end as 'R00539'
,case when t30.R00563=1 then 'O' when t30.R00563=1001 then 'C'  else'' end as 'R00563'
,case when t30.R00562=1 then 'O' when t30.R00562=1001 then 'C' else'' end as 'R00562'
,case when t30.R00561=1 then 'O' when t30.R00561=1001 then 'C' else'' end as 'R00561'
,case when t30.R00503=1 then 'O' when t30.R00503=1001 then 'C' else'' end as 'R00503'
,case when t30.R00506=1 then 'O' when t30.R00506=1001 then 'C' else''  end as 'R00506'
,case when t30.R00504=1 then 'O' when t30.R00504=1001 then 'C' else'' end as 'R00504'
,case when t30.R00505=1 then 'O' when t30.R00505=1001 then 'C' else'' end as 'R00505'
,case when t30.R00526=1 then 'O' when t30.R00526=1001 then 'C' else'' end as 'R00526'
,case when t30.R00564=1 then 'O' when t30.R00564=1001 then 'C' else''  end as 'R00564'
,case when t30.R00502=1 then 'O' when t30.R00502=1001 then 'C' else'' end as 'R00502'
,case when t30.R00540=1 then 'O' when t30.R00540=1001 then 'C' else''  end as 'R00540'
,case when t30.R00572=1 then 'O' when t30.R00572=1001 then 'C' else''  end as 'R00572'
,case when t30.R00587=1 then 'O' when t30.R00587=1001 then 'C' else'' end as 'R00587'
,case when t30.R00598=1 then 'O' when t30.R00598=1001 then 'C' else'' end as 'R00598'
,case when t30.R00599=1 then 'O' when t30.R00599=1001 then 'C' else''  end as 'R00599'
,case when t30.R00600=1 then 'O' when t30.R00600=1001 then 'C' else'' end as 'R00600'
,case when t30.R00610=1 then 'O' when t30.R00610=1001 then 'C' else'' end as 'R00610'
,case when t30.R00611=1 then 'O' when t30.R00611=1001 then 'C' else'' end as 'R00611'
,case when t30.R00613=1 then 'O' when t30.R00613=1001 then 'C' else'' end as 'R00613'
from
(
select t20.docnum
,sum(case when t20.R00554='O' then 1 when t20.R00554='C' then 1001 else 0 end) as 'R00554'
,sum(case when t20.R00550='O' then 1 when t20.R00550='C' then 1001 else 0 end) as 'R00550'
,sum(case when t20.R00551='O' then 1 when t20.R00551='C' then 1001 else 0 end) as 'R00551'
,sum(case when t20.R00527='O' then 1 when t20.R00527='C' then 1001 else 0 end) as 'R00527'
,sum(case when t20.R00539='O' then 1 when t20.R00539='C' then 1001 else 0 end) as 'R00539'
,sum(case when t20.R00563='O' then 1 when t20.R00563='C' then 1001 else 0 end) as 'R00563'
,sum(case when t20.R00562='O' then 1 when t20.R00562='C' then 1001 else 0 end) as 'R00562'
,sum(case when t20.R00561='O' then 1 when t20.R00561='C' then 1001 else 0 end) as 'R00561'
,sum(case when t20.R00503='O' then 1 when t20.R00503='C' then 1001 else 0 end) as 'R00503'
,sum(case when t20.R00506='O' then 1 when t20.R00506='C' then 1001 else 0 end) as 'R00506'
,sum(case when t20.R00504='O' then 1 when t20.R00504='C' then 1001 else 0 end) as 'R00504'
,sum(case when t20.R00505='O' then 1 when t20.R00505='C' then 1001 else 0 end) as 'R00505'
,sum(case when t20.R00526='O' then 1 when t20.R00526='C' then 1001 else 0 end) as 'R00526'
,sum(case when t20.R00564='O' then 1 when t20.R00564='C' then 1001 else 0 end) as 'R00564'
,sum(case when t20.R00502='O' then 1 when t20.R00502='C' then 1001 else 0 end) as 'R00502'
,sum(case when t20.R00540='O' then 1 when t20.R00540='C' then 1001 else 0 end) as 'R00540'
,sum(case when t20.R00572='O' then 1 when t20.R00572='C' then 1001 else 0 end) as 'R00572'
,sum(case when t20.R00587='O' then 1 when t20.R00587='C' then 1001 else 0 end) as 'R00587'
,sum(case when t20.R00598='O' then 1 when t20.R00598='C' then 1001 else 0 end) as 'R00598'
,sum(case when t20.R00599='O' then 1 when t20.R00599='C' then 1001 else 0 end) as 'R00599'
,sum(case when t20.R00600='O' then 1 when t20.R00600='C' then 1001 else 0 end) as 'R00600'
,sum(case when t20.R00610='O' then 1 when t20.R00610='C' then 1001 else 0 end) as 'R00610'
,sum(case when t20.R00611='O' then 1 when t20.R00611='C' then 1001 else 0 end) as 'R00611'
,sum(case when t20.R00613='O' then 1 when t20.R00613='C' then 1001 else 0 end) as 'R00613'

from
(
select t10.docnum
,case when t10.itemcode='R00554' then t10.U_Stato_lavorazione else '' end as 'R00554'
,case when t10.itemcode='R00550' then t10.U_Stato_lavorazione else '' end as 'R00550'
,case when t10.itemcode='R00551' then t10.U_Stato_lavorazione else '' end as 'R00551'
,case when t10.itemcode='R00527' then t10.U_Stato_lavorazione else '' end as 'R00527'
,case when t10.itemcode='R00539' then t10.U_Stato_lavorazione else '' end as 'R00539'
,case when t10.itemcode='R00563' then t10.U_Stato_lavorazione else '' end as 'R00563'
,case when t10.itemcode='R00562' then t10.U_Stato_lavorazione else '' end as 'R00562'
,case when t10.itemcode='R00561' then t10.U_Stato_lavorazione else '' end as 'R00561'
,case when t10.itemcode='R00503' then t10.U_Stato_lavorazione else '' end as 'R00503'
,case when t10.itemcode='R00506' then t10.U_Stato_lavorazione else '' end as 'R00506'
,case when t10.itemcode='R00504' then t10.U_Stato_lavorazione else '' end as 'R00504'
,case when t10.itemcode='R00505' then t10.U_Stato_lavorazione else '' end as 'R00505'
,case when t10.itemcode='R00526' then t10.U_Stato_lavorazione else '' end as 'R00526'
,case when t10.itemcode='R00564' then t10.U_Stato_lavorazione else '' end as 'R00564'
,case when t10.itemcode='R00502' then t10.U_Stato_lavorazione else '' end as 'R00502'
,case when t10.itemcode='R00540' then t10.U_Stato_lavorazione else '' end as 'R00540'
,case when t10.itemcode='R00572' then t10.U_Stato_lavorazione else '' end as 'R00572'
,case when t10.itemcode='R00587' then t10.U_Stato_lavorazione else '' end as 'R00587'
,case when t10.itemcode='R00598' then t10.U_Stato_lavorazione else '' end as 'R00598'
,case when t10.itemcode='R00599' then t10.U_Stato_lavorazione else '' end as 'R00599'
,case when t10.itemcode='R00600' then t10.U_Stato_lavorazione else '' end as 'R00600'
,case when t10.itemcode='R00610' then t10.U_Stato_lavorazione else '' end as 'R00610'
,case when t10.itemcode='R00611' then t10.U_Stato_lavorazione else '' end as 'R00611'
,case when t10.itemcode='R00613' then t10.U_Stato_lavorazione else '' end as 'R00613'
from
(
select t0.docnum, t1.itemcode, t1.U_Stato_lavorazione
from owor t0 left join wor1 t1 on t0.docentry=t1.docentry
inner join orsc t2 on t2.VisResCode=t1.itemcode
where substring(t0.u_produzione,1,3)='INT' and (t0.status ='P' or t0.status='R') and t2.restype='M'
)
as t10

)
as t20
group by t20.docnum
)
as t30
)  T60 ON T60.DOCNUM=T40.DOCNUM

left join (select t0.docnum,t1.linenum, t1.ItemCode as 'MAt_prima'
from owor t0 inner join wor1 t1 on t0.DocEntry =t1.DocEntry
inner join (select t0.docnum, min (t1.linenum) as 'Linenum'
from owor t0 inner join wor1 t1 on t0.DocEntry=t1.DocEntry
where (substring(t1.ItemCode,1,1)='C' or substring(t1.ItemCode,1,1)='D') and substring(t0.u_produzione,1,3)='INT' and (t0.Status='P' or t0.Status='R')
group by t0.docnum) A on t0.DocNum=a.docnum and t1.LineNum = a.Linenum
where substring(t0.u_produzione,1,3)='INT' and (t0.Status='P' or t0.Status='R') and (substring(t1.ItemCode,1,1)='C' or substring(t1.ItemCode,1,1)='D')) t62 on t62.docnum=t40.DocNum

left join oitm t61 on t61.itemcode=t40.U_PRG_AZS_Commessa
left join oscl t64 on cast(t64.callid as varchar) = substring(t40.U_PRG_AZS_COMMESSA,4,999) and substring(t40.U_PRG_AZS_COMMESSA,1,3)='CDS'
left join oitm t63 on t63.itemCode=t64.itemcode

,
(select count(t0.docnum)/90 as 'ODP_chiusi_al_gg'
from owor t0
where substring(t0.u_produzione,1,3) ='INT' and t0.CloseDate>=getdate()-90 and t0.status='L') as t50



)
as t50
where 0=0 " & filtro_odp & " " & filtro_codice & " " & filtro_descrizione & " " & filtro_stato & " " & filtro_disegno & " " & filtro_commessa & " " & filtro_cliente & " " & filtro_mat_prima & " " & filtro_stato_completamento & " " & filtro_R00554 & " " & filtro_R00599 & "" & filtro_R00506 & "" & filtro_R00503 & "" & filtro_R00550 & "" & filtro_R00551 & "" & filtro_R00598 & "" & filtro_R00504 & "" & filtro_R00505 & "" & filtro_R00526 & "" & filtro_R00564 & "" & filtro_R00502 & "" & filtro_R00540 & "" & filtro_R00587 & "" & filtro_R00527 & "" & filtro_R00539 & "" & filtro_R00563 & "" & filtro_R00562 & "" & filtro_R00561 & " " & filtro_R00613 & filtro_R00610 & filtro_R00572 & filtro_R00600 & "
order by t50.Priorità, t50.duedate,t50.U_PRG_AZS_Commessa"

        Else
            CMD_SAP_2.CommandText = "SELECT 
    'false' AS Sel,
    t10.odp AS Docnum,
    999 AS U_lavorazione,
    t10.stato AS Status,

    CASE 
        WHEN data_iniz <> 0 
        THEN CONVERT(date, CAST(data_iniz AS char(8)), 112)
        ELSE NULL
    END AS postdate,

    CASE 
        WHEN t10.data_scad <> 0 
        THEN CONVERT(date, CAST(t10.data_scad AS char(8)), 112)
        ELSE NULL
    END AS duedate,

    CASE 
        WHEN t10.data_scad <> 0 
        THEN CONVERT(date, CAST(t10.data_scad AS char(8)), 112)
        ELSE NULL
    END AS Cons,

    999 AS [ANT/RIT],
    mate_av AS Stato_mat,

    CASE 
        WHEN data_cons <> 0 
        THEN CONVERT(date, CAST(data_cons AS char(8)), 112)
        ELSE NULL
    END AS Arrivo_Mat,

    tipomate AS Mat_prima,
    avanzamento AS Stato_mat_2,
    999 AS Priorità,
    TRIM(codart_odp) AS itemcode,
    dscodart_odp AS Itemname,
    disegno,
    famdis AS [Famiglia disegno],
    999 AS Nesting,
    999 AS Disp,
    '' AS stato_completamento,
    TRIM(matricola) AS u_prg_azs_commessa,
    qta_res AS plannedqty,
    dest AS warehouse,
    cliente AS u_final_customer_name

FROM OPENQUERY(AS400, '
    select 
        STATO,
        ODP,
        MATE_AV,
        MATRICOLA,
        DATA_INIZ,
        DATA_SCAD,
        CODART_ODP,
        DSCODART_ODP,
        DISEGNO,
        FAMDIS,
        QTA_ODP,
        QTA_RES,
        DEST,
        COMMESSA,
        CLIENTE,
        AVANZAMENTO,
        ORDACQ,
        DATA_CONS,
        COD_MAT,
        TIPOMATE,
        ALTEZZA,
        LARGHEZZA,
        LUNGHEZZA
    from S786FAD1.TIR90VIS.JGALodpmu
	where
   upper(codart_odp) LIKE ''%" & par_codice & "%''
	and upper(dscodart_odp)  LIKE ''%" & par_descrizione & "%''
and upper(dscodart_odp)  LIKE ''%" & par_descrizione & "%''
	and ODP  LIKE ''%" & par_odp & "%''
    and upper(stato)  LIKE ''%" & par_stato & "%''
    and upper(disegno)  LIKE ''%" & par_disegno & "%''
    and upper(matricola)  LIKE ''%" & par_commessa & "%''
    and upper(cliente)  LIKE ''%" & par_cliente & "%''
    and upper(tipomate)  LIKE ''%" & par_materia_prima & "%''
    order by data_scad
') AS t10

GROUP BY
    t10.odp,
    t10.stato,
    mate_av,
    tipomate,
    avanzamento,
    TRIM(codart_odp),
    dscodart_odp,
    disegno,
    famdis,
    TRIM(matricola),
    qta_res,
    dest,
    cliente,

    CASE 
        WHEN data_iniz <> 0 
        THEN CONVERT(date, CAST(data_iniz AS char(8)), 112)
        ELSE NULL
    END,

    CASE 
        WHEN t10.data_scad <> 0 
        THEN CONVERT(date, CAST(t10.data_scad AS char(8)), 112)
        ELSE NULL
    END,

    CASE 
        WHEN data_cons <> 0 
        THEN CONVERT(date, CAST(data_cons AS char(8)), 112)
        ELSE NULL
    END;
"
        End If
        ' Dichiarazione di cmd_SAP_reader_2 al di fuori del blocco Using
        Using cmd_SAP_reader_2 As SqlDataReader = CMD_SAP_2.ExecuteReader()
            Dim aggiornamenti As New List(Of String)()
            Dim riga As Integer = 0
            While cmd_SAP_reader_2.Read()
                ' Aggiungi i dati al DataGridView
                If Homepage.ERP_provenienza = "SAP" Then


                    DataGridView1.Rows.Add(False, cmd_SAP_reader_2("docnum"), cmd_SAP_reader_2("U_Lavorazione"),
                       cmd_SAP_reader_2("status"), cmd_SAP_reader_2("postdate"),
                       cmd_SAP_reader_2("Cons"), cmd_SAP_reader_2("duedate"),
                       cmd_SAP_reader_2("ANT/RIT"), cmd_SAP_reader_2("Mat_prima"),
                       cmd_SAP_reader_2("stato_mat"), cmd_SAP_reader_2("Arrivo_MAT"),
                       cmd_SAP_reader_2("Priorità"), cmd_SAP_reader_2("ItemCode"),
                       cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("Disegno"),
                       cmd_SAP_reader_2("Famiglia disegno"), cmd_SAP_reader_2("Nesting"),
                       cmd_SAP_reader_2("plannedqty"), cmd_SAP_reader_2("u_prg_azs_commessa"),
                       cmd_SAP_reader_2("U_Final_customer_name"), cmd_SAP_reader_2("warehouse"),
                       cmd_SAP_reader_2("Disp"), cmd_SAP_reader_2("Stato_completamento"),
                       cmd_SAP_reader_2("R00554"), cmd_SAP_reader_2("R00600"),
                       cmd_SAP_reader_2("R00550"), cmd_SAP_reader_2("R00551"),
                       cmd_SAP_reader_2("R00598"), cmd_SAP_reader_2("R00599"),
                       cmd_SAP_reader_2("R00527"), cmd_SAP_reader_2("R00539"),
                       cmd_SAP_reader_2("R00563"), cmd_SAP_reader_2("R00562"),
                       cmd_SAP_reader_2("R00561"), cmd_SAP_reader_2("R00613"),
                       cmd_SAP_reader_2("R00503"), cmd_SAP_reader_2("R00506"),
                       cmd_SAP_reader_2("R00504"), cmd_SAP_reader_2("R00505"),
                       cmd_SAP_reader_2("R00526"), cmd_SAP_reader_2("R00564"),
                       cmd_SAP_reader_2("R00502"), cmd_SAP_reader_2("R00540"),
                       cmd_SAP_reader_2("R00587"), cmd_SAP_reader_2("R00610"),
                       cmd_SAP_reader_2("R00572"))
                    ' Accumula l'aggiornamento in una lista
                    Dim priorita As Integer = cmd_SAP_reader_2("Priorità")
                    Dim docnum As Integer = cmd_SAP_reader_2("docnum")

                    ' Assicurati che 'Cons' sia trattato correttamente come una data
                    Dim cons As String = cmd_SAP_reader_2("Cons").ToString()

                    ' Usa CONVERT o CAST per formattare correttamente la data
                    aggiornamenti.Add($"UPDATE t101 SET t101.U_priorita_mes = {priorita}, t101.u_data_cons_mes = CONVERT(DATETIME, '{cons}', 103) FROM owor t101 WHERE t101.docnum = {docnum}")
                Else
                    DataGridView1.Rows.Add(False, cmd_SAP_reader_2("docnum"), cmd_SAP_reader_2("U_Lavorazione"),
                      cmd_SAP_reader_2("status"), cmd_SAP_reader_2("postdate"),
                      cmd_SAP_reader_2("Cons"), cmd_SAP_reader_2("duedate"),
                      cmd_SAP_reader_2("ANT/RIT"), cmd_SAP_reader_2("Mat_prima"),
                      cmd_SAP_reader_2("stato_mat"), cmd_SAP_reader_2("Arrivo_MAT"),
                      cmd_SAP_reader_2("Priorità"), cmd_SAP_reader_2("ItemCode"),
                      cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("Disegno"),
                      cmd_SAP_reader_2("Famiglia disegno"), cmd_SAP_reader_2("Nesting"),
                      cmd_SAP_reader_2("plannedqty"), cmd_SAP_reader_2("u_prg_azs_commessa"),
                      cmd_SAP_reader_2("U_Final_customer_name"), cmd_SAP_reader_2("warehouse"),
                      cmd_SAP_reader_2("Disp"), cmd_SAP_reader_2("Stato_completamento"))

                    '    colora_datagridview(DataGridView1, cmd_SAP_reader_2("docnum"), riga)
                End If
                riga += 1

            End While

            If Homepage.ERP_provenienza = "SAP" Then


                ' Esegui tutti gli aggiornamenti in un batch
                If aggiornamenti.Count > 0 Then
                    EseguiAggiornamentiBatch(aggiornamenti)
                End If
            Else

            End If
        End Using

        DataGridView1.ClearSelection()
        Cnn1.Close()

    End Sub

    Sub carico_macchine_NEW(par_odp As String,
                        par_codice As String,
                        par_descrizione As String,
                        par_stato As String,
                        par_disegno As String,
                        par_commessa As String,
                        par_cliente As String,
                        par_materia_prima As String, par_stato_materia_prima As String)

        DataGridView1.Rows.Clear()
        Dim contatore_ordini As Integer = 0
        Using Cnn1 As New SqlConnection(Homepage.sap_tirelli)
            Cnn1.Open()

            Using CMD As New SqlCommand()
                CMD.Connection = Cnn1

                CMD.CommandText = "
select *
from
(
SELECT 
    'false' AS Sel,
    t10.odp AS Docnum,
    999 AS U_lavorazione,
    t10.stato AS Status,

    CASE 
        WHEN data_iniz <> 0 THEN CONVERT(date, CAST(data_iniz AS char(8)), 112)
        ELSE NULL
    END AS postdate,

    CASE 
        WHEN t10.data_scad <> 0 THEN CONVERT(date, CAST(t10.data_scad AS char(8)), 112)
        ELSE NULL
    END AS duedate,

    CASE 
        WHEN t10.data_scad <> 0 THEN CONVERT(date, CAST(t10.data_scad AS char(8)), 112)
        ELSE NULL
    END AS Cons,

    999 AS [ANT/RIT],
    mate_av AS Stato_mat,

    CASE 
        WHEN data_cons <> 0 THEN CONVERT(date, CAST(data_cons AS char(8)), 112)
        ELSE NULL
    END AS Arrivo_Mat,

    tipomate AS Mat_prima,
    avanzamento AS Stato_mat_2,
    999 AS Priorità,
    TRIM(codart_odp) AS ItemCode,
    dscodart_odp AS Itemname,
    trim(disegno) as 'Disegno',
    famdis AS [Famiglia disegno],
    999 AS Nesting,
    999 AS Disp,
    '' AS stato_completamento,
    TRIM(matricola) AS u_prg_azs_commessa,
    qta_res AS plannedqty,
    dest AS warehouse,
    cliente AS u_final_customer_name,

     -- Qui raggruppiamo le attrezzature
     -- aggregazioni
    STRING_AGG(TRIM(attrezzatura), ', ') AS Attrezzatura,
	STRING_AGG(TRIM(cod_risorsa), ', ') AS cod_risorsa,
    STRING_AGG(CASE WHEN stato_fase IS NULL THEN '' ELSE stato_fase END, ', ') AS Stato_Fase

FROM OPENQUERY(AS400, '
    SELECT 
        STATO,
        ODP,
        MATE_AV,
        MATRICOLA,
        DATA_INIZ,
        DATA_SCAD,
        CODART_ODP,
        DSCODART_ODP,
        DISEGNO,
        FAMDIS,
        QTA_ODP,
        QTA_RES,
        DEST,
        COMMESSA,
        CLIENTE,
        AVANZAMENTO,
        ORDACQ,
        DATA_CONS,
        COD_MAT,
        TIPOMATE,
        ALTEZZA,
        LARGHEZZA,
        LUNGHEZZA,
		cod_risorsa,
        ATTREZZATURA,
		stato_fase
    FROM S786FAD1.TIR90VIS.JGALodpmu
    WHERE
 (dest=''TCP'' or dest=''MU'') and
        UPPER(codart_odp) LIKE ''%" & par_codice & "%''
        AND UPPER(dscodart_odp) LIKE ''%" & par_descrizione & "%''
        AND ODP LIKE ''%" & par_odp & "%''
        AND UPPER(stato) LIKE ''%" & par_stato & "%''
        AND UPPER(disegno) LIKE ''%" & par_disegno & "%''
        AND UPPER(coalesce(matricola,'''')) LIKE ''%" & par_commessa & "%''
        AND UPPER(cliente) LIKE ''%" & par_cliente & "%''
        AND UPPER(tipomate) LIKE ''%" & par_materia_prima & "%''
AND UPPER(mate_av) LIKE ''%" & par_stato_materia_prima & "%''
and tipo_macchina=''M''
ORDER BY ODP , CODE_FASE

') AS t10

GROUP BY
    t10.odp,
    t10.stato,
    mate_av,
    tipomate,
    avanzamento,
    TRIM(codart_odp),
    dscodart_odp,
    disegno,
    famdis,
    TRIM(matricola),
    qta_res,
    dest,
    cliente,
    CASE WHEN data_iniz <> 0 THEN CONVERT(date, CAST(data_iniz AS char(8)), 112) ELSE NULL END,
    CASE WHEN t10.data_scad <> 0 THEN CONVERT(date, CAST(t10.data_scad AS char(8)), 112) ELSE NULL END,
    CASE WHEN data_cons <> 0 THEN CONVERT(date, CAST(data_cons AS char(8)), 112) ELSE NULL END
)
as t20
where 0=0 " & filtro_stato_completamento & "

ORDER BY T20.duedate

"

                Using rd As SqlDataReader = CMD.ExecuteReader()



                    While rd.Read()

                        Dim img As Image = Nothing

                        Dim codiceDisegno As String = rd("Disegno").ToString()
                        Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & codiceDisegno & ".PNG"

                        If File.Exists(percorso) Then
                            Using fs As New FileStream(percorso, FileMode.Open, FileAccess.Read)
                                Using tmp As Image = Image.FromStream(fs)
                                    img = New Bitmap(tmp) ' evita lock sul file
                                End Using
                            End Using
                        End If
                        DataGridView1.Rows.Add(
                        False,
                        rd("docnum"),
                        rd("U_Lavorazione"),
                        rd("status"),
                        rd("postdate"),
                        rd("Cons"),
                        rd("duedate"),
                        rd("ANT/RIT"),
                        rd("Mat_prima"),
                        rd("stato_mat"),
                        rd("Arrivo_MAT"),
                        rd("Priorità"),
                        rd("ItemCode"),
                        img,
                        rd("itemname"),
                        rd("Disegno"),
                        rd("Famiglia disegno"),
                        rd("Nesting"),
                        rd("plannedqty"),
                        rd("u_prg_azs_commessa"),
                        rd("U_Final_customer_name"),
                        rd("warehouse"),
                        rd("Disp"),
                        rd("Attrezzatura"),
                        rd("cod_risorsa"),
                        rd("stato_fase")' 👈 COLONNA TECNICA (anche nascosta)
                    )
                        contatore_ordini += 1
                    End While
                End Using
            End Using
        End Using

        DataGridView1.ClearSelection()
        Label_n_ordini.Text = contatore_ordini
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) _
Handles DataGridView1.CellFormatting

        Dim row = DataGridView1.Rows(e.RowIndex)

        ColoraRisorsa(row, "800100", "Prog")
        ColoraRisorsa(row, "800200", "Prog_p")
        ColoraRisorsa(row, "100100", "T_man")
        ColoraRisorsa(row, "100200", "T_auto")
        ColoraRisorsa(row, "R00598", "Pant_P")
        ColoraRisorsa(row, "R00599", "Pant_A")
        ColoraRisorsa(row, "200100", "Tornio_")
        ColoraRisorsa(row, "200500", "Doosan")
        ColoraRisorsa(row, "200400", "Doosan_4")
        ColoraRisorsa(row, "200300", "Doos6")
        ColoraRisorsa(row, "200800", "Multi")
        ColoraRisorsa(row, "300600", "Haas_2")
        ColoraRisorsa(row, "300200", "Haas_3")
        ColoraRisorsa(row, "300100", "Haas_6")
        ColoraRisorsa(row, "300700", "Haas_5")
        ColoraRisorsa(row, "300300", "Famup")
        ColoraRisorsa(row, "300400", "Awea")
        ColoraRisorsa(row, "600100", "Stozza")
        ColoraRisorsa(row, "700100", "Sald")
        ColoraRisorsa(row, "500100", "Trap")
        ' Vulcanizzazione non ha codice Galileo definito
        ColoraRisorsa(row, "R00610", "Vulcanizzazione")
        ColoraRisorsa(row, "900100", "Finitura")

    End Sub

    Private Sub ColoraRisorsa(row As DataGridViewRow,
                          codiceRisorsa As String,
                          nomeColonna As String)

        Dim attRaw = row.Cells("Cod_risorsa").Value
        Dim statoRaw = row.Cells("stato_fase").Value

        If attRaw Is Nothing OrElse statoRaw Is Nothing Then Exit Sub

        Dim attrezzature = attRaw.ToString().
                        Split(","c).
                        Select(Function(x) x.Trim()).
                        ToList()

        Dim stati = statoRaw.ToString().
                Split(","c).
                Select(Function(x) x.Trim()).
                ToList()

        If attrezzature.Count <> stati.Count Then Exit Sub

        Dim idx = attrezzature.IndexOf(codiceRisorsa)
        If idx = -1 Then Exit Sub

        Dim stato = stati(idx)

        If stato = "S" Then
            row.Cells(nomeColonna).Style.BackColor = Color.LightGreen
        ElseIf String.IsNullOrWhiteSpace(stato) Then
            row.Cells(nomeColonna).Style.BackColor = Color.Yellow
        End If

    End Sub


    Private Sub DataGridView1_RowPrePaint(sender As Object, e As DataGridViewRowPrePaintEventArgs)


    End Sub

    Sub colora_datagridview(par_datagridview As DataGridView, par_odp As String, par_riga As Integer)



        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand



        CMD_SAP_2.Connection = Cnn1


        CMD_SAP_2.CommandText = "SELECT 
odp,
cod_risorsa
,code_fase
,desc_fase
,stato_fase
,trim(attrezzatura) as 'Attrezzatura'
,risorsa


FROM OPENQUERY(AS400, '
    select 
        *
    from S786FAD1.TIR90VIS.JGALodpmu
	where
   upper(codart_odp) LIKE ''%%''
	and upper(dscodart_odp)  LIKE ''%%''
and upper(dscodart_odp)  LIKE ''%%''
	and ODP  = ''" & par_odp & "''
    and upper(stato)  LIKE ''%%''
    and upper(disegno)  LIKE ''%%''
    and upper(matricola)  LIKE ''%%''
    and upper(cliente)  LIKE ''%%''
    and upper(tipomate)  LIKE ''%%''
    order by data_scad
') AS t10"

        ' Dichiarazione di cmd_SAP_reader_2 al di fuori del blocco Using
        Using cmd_SAP_reader_2 As SqlDataReader = CMD_SAP_2.ExecuteReader()
            Dim aggiornamenti As New List(Of String)()

            While cmd_SAP_reader_2.Read()

                If cmd_SAP_reader_2("attrezzatura") = "R00505" Then
                    If par_riga >= 0 Then
                        If cmd_SAP_reader_2("attrezzatura").ToString() = "R00505" Then
                            par_datagridview.Rows(par_riga).Cells(39).Style.BackColor = Color.LightGreen
                        End If
                    End If
                End If

            End While


        End Using


        Cnn1.Close()

    End Sub
    Sub EseguiAggiornamentiBatch(aggiornamenti As List(Of String))
        ' Filtra eventuali stringhe vuote
        Dim aggiornamentiValidi = aggiornamenti.Where(Function(x) Not String.IsNullOrWhiteSpace(x)).ToList()

        If aggiornamentiValidi.Count = 0 Then Exit Sub

        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()

            Dim comando As String = String.Join(";", aggiornamentiValidi)

            Using cmd_SAP As New SqlCommand(comando, CNN)
                cmd_SAP.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Sub Salva_priorita_singola(par_docnum As Integer, priorità_new As Integer)

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli

        CNN.Open()

        Dim Cmd_SAP As New SqlCommand


        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = CNN


        Cmd_SAP.CommandText = "update t101 set t101.U_priorita_mes=" & priorità_new & "
from owor t101 where t101.docnum=" & par_docnum & ""
        Cmd_SAP.ExecuteNonQuery()



        CNN.Close()
    End Sub



    Sub Carico_macchine_grafico_programmazione()
        If Homepage.ERP_provenienza = "SAP" Then


            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_1 As New SqlCommand
            Dim cmd_SAP_reader_1 As SqlDataReader
            CMD_SAP_1.Connection = Cnn1

            CMD_SAP_1.CommandText = " 
SELECT TOP 11 T3.[ItemCode],T3.ITEMNAME,  SUM(T1.[PlannedQty]/60) AS'Ore'
FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] inner join ORSC T2 on T2.[VisResCode] =t1.itemcode INNER JOIN OITM T3 ON T1.[ItemCode] = T3.[ItemCode] 

WHERE (T0.[Status] ='P' or  T0.[Status] ='R') and substring(t0.u_produzione,1,3)='INT' AND T2.RESTYPE='M'  AND T1.U_STATO_LAVORAZIONE='O' AND (t2.u_ordine=2 OR t2.u_ordine=3 )

GROUP BY T3.[ItemCode],T3.ITEMNAME,t2.u_ordine

order by t2.u_ordine DESC ,T3.ITEMNAME DESC "


            cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
            Do While cmd_SAP_reader_1.Read()
                Chart1.Series("Ore").BorderWidth = 1
                Chart_lavorazioni.Series("Ore").Points.AddXY(cmd_SAP_reader_1("ITEMNAME"), Math.Round(cmd_SAP_reader_1("Ore")))


            Loop

            Cnn1.Close()
        End If
    End Sub

    Sub Carico_macchine_grafico_torni()
        If Homepage.ERP_provenienza = "SAP" Then


            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_1 As New SqlCommand
            Dim cmd_SAP_reader_1 As SqlDataReader
            CMD_SAP_1.Connection = Cnn1

            CMD_SAP_1.CommandText = " 
SELECT T3.[ItemCode],T3.ITEMNAME,  SUM(T1.[PlannedQty]/60) AS'Ore'
FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] inner join ORSC T2 on T2.[VisResCode] =t1.itemcode INNER JOIN OITM T3 ON T1.[ItemCode] = T3.[ItemCode] 

WHERE (T0.[Status] ='P' or  T0.[Status] ='R') and substring(t0.u_produzione,1,3)='INT' AND T2.RESTYPE='M' and t2.u_ordine=2 AND  T1.U_STATO_LAVORAZIONE='O'

GROUP BY T3.[ItemCode],T3.ITEMNAME,t2.u_ordine

order by t2.u_ordine,T3.ITEMNAME"


            cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
            Do While cmd_SAP_reader_1.Read()

                Chart1.Series("Ore").Points.AddXY(cmd_SAP_reader_1("ITEMNAME"), Math.Round(cmd_SAP_reader_1("Ore")))


            Loop

            Cnn1.Close()
        End If
    End Sub

    Sub Carico_macchine_grafico_frese()
        If Homepage.ERP_provenienza = "SAP" Then
            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_1 As New SqlCommand
            Dim cmd_SAP_reader_1 As SqlDataReader
            CMD_SAP_1.Connection = Cnn1

            CMD_SAP_1.CommandText = " 
SELECT T3.[ItemCode],T3.ITEMNAME,  SUM(T1.[PlannedQty]/60) AS'Ore'
FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] inner join ORSC T2 on T2.[VisResCode] =t1.itemcode INNER JOIN OITM T3 ON T1.[ItemCode] = T3.[ItemCode] 

WHERE (T0.[Status] ='P' or  T0.[Status] ='R') and substring(t0.u_produzione,1,3)='INT' AND T2.RESTYPE='M' and t2.u_ordine=3 AND T1.U_STATO_LAVORAZIONE='O'

GROUP BY T3.[ItemCode],T3.ITEMNAME,t2.u_ordine

order by t2.u_ordine,T3.ITEMNAME"


            cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
            Do While cmd_SAP_reader_1.Read()

                Chart2.Series("Ore").Points.AddXY(cmd_SAP_reader_1("ITEMNAME"), Math.Round(cmd_SAP_reader_1("Ore")))


            Loop

            Cnn1.Close()
        End If
    End Sub

    Sub Carico_macchine_grafico_altro()
        If Homepage.ERP_provenienza = "SAP" Then
            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_1 As New SqlCommand
            Dim cmd_SAP_reader_1 As SqlDataReader
            CMD_SAP_1.Connection = Cnn1

            CMD_SAP_1.CommandText = " 
SELECT T3.[ItemCode],T3.ITEMNAME,  SUM(T1.[PlannedQty]/60) AS'Ore'
FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] inner join ORSC T2 on T2.[VisResCode] =t1.itemcode INNER JOIN OITM T3 ON T1.[ItemCode] = T3.[ItemCode] 

WHERE (T0.[Status] ='P' or  T0.[Status] ='R') and substring(t0.u_produzione,1,3)='INT' AND T2.RESTYPE='M' and t2.u_ordine=4 AND T1.U_STATO_LAVORAZIONE='O'

GROUP BY T3.[ItemCode],T3.ITEMNAME,t2.u_ordine

order by t2.u_ordine,T3.ITEMNAME"


            cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
            Do While cmd_SAP_reader_1.Read()

                Chart3.Series("Ore").Points.AddXY(cmd_SAP_reader_1("ITEMNAME"), Math.Round(cmd_SAP_reader_1("Ore")))

            Loop

            Cnn1.Close()
        End If
    End Sub



    Private Sub Button_rilascia_Click(sender As Object, e As EventArgs)
        If Homepage.ERP_provenienza = "SAP" Then
            Dim contatore As Integer = 0


            Do While contatore < DataGridView1.Rows.Count



                If DataGridView1.Rows(contatore).Cells(columnName:="seleziona").Value = True Then
                    If DataGridView1.Rows(contatore).Cells(columnName:="seleziona").Value = True Then
                        DataGridView1.Rows(contatore).Cells(columnName:="Stato").Value = "R"
                        DataGridView1.Rows(contatore).Cells(columnName:="Lavorazione_").Value = lavorazione
                        Dim CNN As New SqlConnection
                        CNN.ConnectionString = Homepage.sap_tirelli
                        CNN.Open()
                        Dim CMD_SAP As New SqlCommand
                        CMD_SAP.Connection = CNN

                        CMD_SAP.CommandText = "UPDATE owor SET STATUS='R' WHERE DOCNUM ='" & DataGridView1.Rows(contatore).Cells(columnName:="ODP_").Value & "'"
                        CMD_SAP.ExecuteNonQuery()
                        CNN.Close()

                    End If
                End If
                contatore = contatore + 1
            Loop
            contatore = 0
        End If
    End Sub

    Sub Max_numerazione()
        If Homepage.ERP_provenienza = "SAP" Then


            Dim CNN As New SqlConnection
            CNN.ConnectionString = Homepage.sap_tirelli
            CNN.Open()

            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader
            CMD_SAP.Connection = CNN

            CMD_SAP.CommandText = "SELECT max(T0.[U_Lavorazione]) as 'Lavorazione' FROM OWOR T0 where (t0.status='P' or t0.status='R') "

            cmd_SAP_reader = CMD_SAP.ExecuteReader

            If cmd_SAP_reader.Read() = True Then

                lavorazione = cmd_SAP_reader("Lavorazione") + 1

            End If

            CNN.Close()
            cmd_SAP_reader.Close()
        End If
    End Sub

    Sub Check_Lavorazioni_aperte_macchina()
        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.id as 'ID', t0.tipo_documento as 'Documento', t0.docnum as 'ODP', t1.itemcode as 'Itemcode', t2.itemname as 'Descrizione', case when t2.u_disegno is null then '' else t2.u_disegno end as 'Disegno', t1.plannedqty as 'quantita', T1.[U_PRG_AZS_Commessa] as 'Commessa', t4.firstname+' '+t4.lastname as 'Dipendente', t3.resname as 'Risorsa', t0.data as 'Data', t0.start as 'Start', t0.stop, t0.consuntivo
from manodopera t0 left join owor t1 on t0.docnum = t1.docnum
left join oitm t2 on t2.itemcode=t1.itemcode
inner join orsc t3 on t3.visrescode=t0.risorsa
inner join [TIRELLI_40].[dbo].ohem t4 on t4.empid=t0.dipendente
where (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='') AND t3.visrescode='" & Dashboard_MU_New.risorsa & "' AND t0.docnum<>'" & Dashboard_MU_New.docnum & "' "

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            MsgBox("Il macchinario " & cmd_SAP_reader_2("Risorsa") & " risulta aperto nell'ordine " & cmd_SAP_reader_2("ODP") & " da " & cmd_SAP_reader_2("Dipendente") & "  ")

        Else
            Check_attrezzaggio_lavorazione()
            Dashboard_MU_New.invia_Odp(Dashboard_MU_New.docnum)
            Dashboard_MU_New.righe_ODP_macchine(DataGridView1, Dashboard_MU_New.docnum)

        End If
        Cnn1.Close()
        cmd_SAP_reader_2.Close()

    End Sub

    Sub Check_attrezzaggio_lavorazione()
        Dim CNN2 As New SqlConnection
        CNN2.ConnectionString = Homepage.sap_tirelli
        CNN2.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = CNN2
        CMD_SAP_2.CommandText = "SELECT t0.id as 'id', t0.tipologia_lavorazione as 'Tipologia lavorazione' 
from manodopera t0 INNER join orsc t1 on t0.risorsa=t1.visrescode
 where t0.docnum=" & Dashboard_MU_New.docnum & " and (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='') and t1.[ResType]='M'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then

        End If
        CNN2.Close()
        cmd_SAP_reader_2.Close()

    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            Dim codiceDisegno As String = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disegno").Value
            Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & codiceDisegno & ".PNG"






            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Descrizione) Then

            End If




            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Commessa) Then
                Mostra.Hide()
                Homepage.commessa = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Commessa").Value

                Homepage.mostra_dashboard()

                Mostra.Show()


            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(ODP_) Then

                If DataGridView1.Rows(e.RowIndex).Cells(columnName:="ODP_").Value = 0 Then

                Else

                    ODP_Form.docnum_odp = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ODP_").Value
                    ODP_Form.Show()
                    ODP_Form.inizializza_form(DataGridView1.Rows(e.RowIndex).Cells(columnName:="ODP_").Value)



                End If


            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(Codice) Then

                Magazzino.Codice_SAP = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice").Value

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


            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(Cod_mat) Then

                Magazzino.Codice_SAP = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Cod_mat").Value

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
            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(Nesting) Then


                Dim i As Integer = 0
                Dim parola1 As String
                Dim parola2 As String
                parola1 = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Famiglia_Dis").Value
                Do While i < DataGridView1.RowCount
                    parola2 = DataGridView1.Rows(i).Cells(columnName:="Famiglia_Dis").Value
                    If parola1 = parola2 Then
                        DataGridView1.Rows(i).Visible = True
                    Else
                        DataGridView1.Rows(i).Visible = False
                    End If
                    i = i + 1
                Loop

            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(Disegno) Then
                Magazzino.visualizza_disegno(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disegno").Value)


            Else

                If File.Exists(percorso) Then

                    codice_pic = codiceDisegno
                    Magazzino.visualizza_picture(codice_pic, PictureBox2)

                End If

            End If




        End If
    End Sub








    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Sub filtra()
        Dim i = 0
        Dim parola1 As String
        Dim invisibile As Integer = 0
        Dim parola4 As String
        Dim parola9 As String
        Dim parola11 As String
        Dim parola10 As String
        Dim parola16 As String
        Dim filtro_cliente As String
        Dim filtro_mat_prima As String

        Do While i < DataGridView1.RowCount
            invisibile = 0
            Try

                parola1 = UCase(DataGridView1.Rows(i).Cells(columnName:="ODP_").Value)

                parola4 = UCase(DataGridView1.Rows(i).Cells(columnName:="Stato").Value)
                parola9 = UCase(DataGridView1.Rows(i).Cells(columnName:="Codice").Value)
                parola10 = UCase(DataGridView1.Rows(i).Cells(columnName:="Descrizione").Value)
                parola11 = UCase(DataGridView1.Rows(i).Cells(columnName:="Disegno").Value)
                parola16 = UCase(DataGridView1.Rows(i).Cells(columnName:="Commessa").Value)
                filtro_cliente = UCase(DataGridView1.Rows(i).Cells(columnName:="Cliente").Value)
                filtro_mat_prima = UCase(DataGridView1.Rows(i).Cells(columnName:="Cod_mat").Value)


                If parola1.Contains(UCase(TextBox5.Text)) Then
                    DataGridView1.Rows(i).Visible = True
                    If parola4.Contains(UCase(TextBox3.Text)) Then
                        DataGridView1.Rows(i).Visible = True


                        If parola9.Contains(UCase(TextBox1.Text)) Then
                            DataGridView1.Rows(i).Visible = True


                            If parola11.Contains(UCase(TextBox4.Text)) Then
                                DataGridView1.Rows(i).Visible = True

                                If parola10.Contains(UCase(TextBox2.Text)) Then
                                    DataGridView1.Rows(i).Visible = True

                                    If parola16.Contains(UCase(TextBox6.Text)) Then
                                        DataGridView1.Rows(i).Visible = True

                                        If filtro_cliente.Contains(UCase(TextBox14.Text)) Then
                                            DataGridView1.Rows(i).Visible = True

                                            If filtro_mat_prima.Contains(UCase(TextBox16.Text)) Then
                                                DataGridView1.Rows(i).Visible = True

                                                If CheckBox1.Checked = True Then
                                                    If DataGridView1.Rows(i).Cells(DataGridView1.Columns.IndexOf(Prog)).Value = "O" Or DataGridView1.Rows(i).Cells(DataGridView1.Columns.IndexOf(Prog)).Value = "C" Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If


                                                End If


                                                If CheckBox2.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="T_MAN").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="T_MAN").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If

                                                If CheckBox3.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="T_Auto").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="T_Auto").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If

                                                If CheckBox4.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Tornio_").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Tornio_").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If

                                                If CheckBox5.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Goodway").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Goodway").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If
                                                If CheckBox6.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Doosan").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Doosan").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If
                                                If CheckBox7.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Doosan_4").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Doosan_4").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If
                                                If CheckBox8.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Doos6").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Doos6").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If

                                                If CheckBox9.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Haas_2").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Haas_2").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If
                                                If CheckBox10.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Haas_3").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Haas_3").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If
                                                If CheckBox11.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Haas_5").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Haas_5").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If
                                                If CheckBox12.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Haas_6").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Haas_6").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If
                                                If CheckBox13.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Famup").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Famup").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If
                                                If CheckBox14.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Awea").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Awea").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If
                                                If CheckBox15.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Stozza").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Stozza").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If
                                                If CheckBox16.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Sald").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Sald").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If

                                                If CheckBox24.Checked = True Then
                                                    If (DataGridView1.Rows(i).Cells(columnName:="Multi").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Multi").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If

                                                If CheckBox17.Checked = True Then
                                                    If DataGridView1.Rows(i).Cells(columnName:="Completati").Value = "Completato" And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If

                                                If CheckBox18.Checked = True Then

                                                    If DataGridView1.Rows(i).Cells(columnName:="Prog").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Sald").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Stozza").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Awea").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Famup").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Haas_6").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Haas_5").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Haas_3").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Haas_2").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Doos6").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Doosan_4").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Doosan").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Goodway").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Tornio_").Value = "" And DataGridView1.Rows(i).Cells(columnName:="T_Auto").Value = "" And DataGridView1.Rows(i).Cells(columnName:="T_MAN").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Trap").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Fin").Value = "" And DataGridView1.Rows(i).Cells(columnName:="Pant_P").Value = "" And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If

                                                If CheckBox19.Checked = True Then

                                                    If (DataGridView1.Rows(i).Cells(columnName:="Trap").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Trap").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If


                                                If CheckBox21.Checked = True Then

                                                    If (DataGridView1.Rows(i).Cells(columnName:="Pant_P").Value = "O" Or DataGridView1.Rows(i).Cells(columnName:="Pant_P").Value = "C") And invisibile = 0 Then
                                                        DataGridView1.Rows(i).Visible = True

                                                    Else

                                                        DataGridView1.Rows(i).Visible = False
                                                        invisibile = invisibile + 1
                                                    End If

                                                End If



                                            Else
                                                DataGridView1.Rows(i).Visible = False
                                            End If

                                        Else
                                            DataGridView1.Rows(i).Visible = False


                                        End If


                                    Else
                                        DataGridView1.Rows(i).Visible = False

                                    End If

                                Else
                                    DataGridView1.Rows(i).Visible = False

                                End If


                            Else
                                DataGridView1.Rows(i).Visible = False

                            End If

                        Else
                            DataGridView1.Rows(i).Visible = False

                        End If


                    Else
                        DataGridView1.Rows(i).Visible = False

                    End If
                Else
                    DataGridView1.Rows(i).Visible = False

                End If



            Catch ex As Exception
                DataGridView1.Rows(i).Visible = False
            End Try
            i = i + 1
        Loop
    End Sub



    Sub Nesting_priorita()
        Dim CNN3 As New SqlConnection
        CNN3.ConnectionString = Homepage.sap_tirelli
        CNN3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = CNN3


        CMD_SAP_3.CommandText = "update t11 set t11.[U_Priorita_MES]=t10.[Priorita_min]
FROM
(
SELECT T0.[U_Lavorazione], COUNT(T0.[U_Lavorazione]) as 'Nesting' , MIN(T0.[U_Priorita_MES]) AS 'Priorita_min'
FROM OWOR T0 WHERE SUBSTRING(T0.[U_PRODUZIONE],1,3)='INT' and T0.[U_Lavorazione]<>'' and (T0.[Status]='P' or T0.[Status]='R')
GROUP BY T0.[U_Lavorazione]
)
as t10 inner join owor t11 on t11.[U_Lavorazione]=T10.[U_Lavorazione]
where t10.nesting>1 and (T11.[Status]='P' or T11.[Status]='R')"

        CMD_SAP_3.ExecuteNonQuery()
        CNN3.Close()



    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim contatore As Integer = 0


        Do While contatore < DataGridView1.Rows.Count
            If DataGridView1.Rows(contatore).Cells(0).Value = True Then

                Dim CNN As New SqlConnection
                CNN.ConnectionString = Homepage.sap_tirelli
                CNN.Open()
                Dim CMD_SAP As New SqlCommand
                CMD_SAP.Connection = CNN

                CMD_SAP.CommandText = "UPDATE owor set U_Lavorazione='" & DataGridView1.Rows(contatore).Cells(DataGridView1.Columns.IndexOf(Lavorazione_)).Value & "' WHERE DOCNUM ='" & DataGridView1.Rows(contatore).Cells(DataGridView1.Columns.IndexOf(ODP_)).Value & "'"
                CMD_SAP.ExecuteNonQuery()
                CNN.Close()

            End If

            contatore = contatore + 1
        Loop
        contatore = 1

        carico_macchine(TextBox5.Text.ToUpper, TextBox1.Text.ToUpper, TextBox2.Text.ToUpper, TextBox3.Text.ToUpper, TextBox4.Text.ToUpper, TextBox6.Text.ToUpper, TextBox14.Text.ToUpper, TextBox16.Text.ToUpper)

        backlog_ordini()
    End Sub

    Sub datagridview_carico_macchine()
        carico_macchine_NEW(TextBox5.Text.ToUpper, TextBox1.Text.ToUpper, TextBox2.Text.ToUpper, TextBox3.Text.ToUpper, TextBox4.Text.ToUpper, TextBox6.Text.ToUpper, TextBox14.Text.ToUpper, TextBox16.Text.ToUpper, TextBox17.Text.ToUpper)
    End Sub


    Private Sub Button1_Click_2(sender As Object, e As EventArgs)
        Max_numerazione()

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)
        filtra()
    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs)

        Dashboard_MU_New.ELIMINA_risorse_IN_ODP_COMPLETATI("")
        Dashboard_MU_New.carica_lavorazioni_in_ODP("")


        MsgBox("Manodopera caricata con successo")
    End Sub



    Sub Salva_priorita()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli

        CNN.Open()

        Dim Cmd_SAP As New SqlCommand


        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = CNN


        Cmd_SAP.CommandText = "update t101 set t101.U_priorita_mes=t100.Priorità, t101.u_data_cons_mes=t100.cons
from
(

select  t40.Seleziona, t40.docnum, t40.U_Lavorazione, t40.status, t40.postdate, getdate()+(ROW_NUMBER() OVER(ORDER BY t40.priorità asc)/t50.ODP_chiusi_al_gg) as 'Cons', DATEDIFF(DD,T40.DUEDATE,GETDATE()) AS 'ANT/RIT',t40.stato_mat, t40.Arrivo_MAT,t40.Priorità, t40.ItemCode, t40.itemname, t40.Disegno, t40.[Famiglia disegno], t40.nesting, t40.disp,t40.stato_completamento,

T60.R00554
,T60.R00550
,T60.R00527
,T60.R00539
,T60.R00563
,T60.R00562
,T60.R00561
,T60.R00503
,T60.R00506
,T60.R00504
,T60.R00505
,T60.R00526
,T60.R00564
,T60.R00502
,T60.R00540
,T60.R00587
,T60.R00572
,T60.R00598
,T60.R00599
,T60.R00600


from
(
select  t10.Seleziona, t10.docnum, t10.U_Lavorazione, t10.status, t10.postdate,T10.DUEDATE, t10.Cons, t10.[ant/rit],c.stato_mat, c.Arrivo_MAT,case when t10.u_lavorazione=0 then t10.Priorità else min(t10.priorità) OVER(PARTITION BY t10.u_lavorazione) end as'Priorità',t10.ItemCode, t10.itemname, t10.Disegno, t10.[Famiglia disegno], t10.nesting, B.disp,t10.stato_completamento
from
(

select 'sel' as 'Seleziona', t0.docnum, t0.u_lavorazione,t0.status, t0.PostDate,T0.DUEDATE, getdate() as 'Cons', 5 as 'ant/rit','MAT' as 'MAt','arrivoMAT' as 'ArrivoMAT',case when t0.U_Fase = 'urgenza_COLLAUDO' then 1 WHEN t0.U_Fase = 'urgenza_MFC' then 0 else ROW_NUMBER() OVER(ORDER BY t0.Status DESC, T0.[duedate] ASC) end as 'Priorità', t0.ItemCode, t1.itemname, case when t1.u_disegno is null then '' else t1.u_disegno end  as 'Disegno',

case when SUBSTRING (T1.U_Disegno,5,1) = '_'  then SUBSTRING (T1.U_Disegno,1,4)
else CASE when SUBSTRING (T1.U_Disegno,5,1) = '' then SUBSTRING (T1.U_Disegno,1,4)
ELSE SUBSTRING (T1.U_Disegno,1,5) end end as 'Famiglia disegno',

case when SUBSTRING (T1.U_Disegno,5,1) = '_'  then COUNT(SUBSTRING (T1.U_Disegno,1,4)) OVER(PARTITION BY SUBSTRING (T1.U_Disegno,1,4))
else CASE when SUBSTRING (T1.U_Disegno,5,1) = '' then COUNT(SUBSTRING (T1.U_Disegno,1,4)) OVER(PARTITION BY SUBSTRING (T1.U_Disegno,1,4))
ELSE COUNT(SUBSTRING (T1.U_Disegno,1,5)) OVER(PARTITION BY SUBSTRING (T1.U_Disegno,1,5)) end end  as 'Nesting', t0.U_PRG_AZS_Commessa, 'DISP' as 'DISP',

case when t0.u_stato is null then '' else t0.u_stato end as 'stato_completamento'

from owor t0 left join oitm t1 on t0.itemcode=t1.itemcode 
where substring(t0.u_produzione,1,3)='INT' and (t0.Status='R' or t0.Status='P') 

)
as t10 inner join 
(select t0.docnum, sum(t11.onhand-t11.iscommited+t11.onorder) as 'Disp' from owor t0 inner join oitw t11 on t11.itemcode=t0.itemcode where substring(t0.u_produzione,1,3)='INT' and (t0.status='P' or t0.status='R') group by t0.docnum) as B on B.DocNum=t10.docnum

inner join (select t30.docnum, t30.itemcode,  case when t30.Da_trasferire is null or t30.da_trasferire=0 then 'OK' when t30.non_trasferibili is null or t30.non_trasferibili=0 then 'TRASF' else 'IN_APPR'  end as 'stato_mat', case when t30.Shipdate_oa_figlio is null then t30.Cons_odp_figlio when t30.Cons_odp_figlio is null then t30.Shipdate_oa_figlio end as 'Arrivo_MAT'
from
(
select t20.docnum, t20.itemcode, sum(case when t20.u_prg_wip_qtadatrasf>0 then 1 else 0 end) as 'Da_trasferire',sum(case when t20.u_prg_wip_qtadatrasf>t20.mag then 1 else 0 end) as 'non_Trasferibili', min(T20.Shipdate) as 'Shipdate_oa_figlio', min(T20.cons_odp) as 'Cons_odp_figlio'
from
(
select t10.docnum, t10.itemcode, t10.u_prg_wip_qtadatrasf, t10.Mag, min (t11.shipdate) as 'Shipdate', min (case when substring (t12.u_produzione,1,3)='INT' then t12.u_data_cons_mes else t12.duedate end ) as 'Cons_ODP'
from
(
select t0.docnum, t1.itemcode, t1.u_prg_wip_qtadatrasf, sum(case when t2.onhand is null then 0 else t2.onhand end) as 'Mag'
from owor t0 left join wor1 t1 on t0.docentry=t1.docentry
left join oitw t2 on t2.itemcode=t1.itemcode and  t1.u_prg_wip_qtadatrasf >0 and t2.whscode<>'WIP' and t2.whscode<>'Clavter'
where substring(t0.u_produzione,1,3)='INT' and (t0.Status='R' or t0.Status='P') and (SUBSTRING(t1.itemcode,1,1)='C' or SUBSTRING(t1.itemcode,1,1)='D' or SUBSTRING(t1.itemcode,1,1)='0')
group by t0.docnum, t1.itemcode, t1.u_prg_wip_qtadatrasf
)
as t10 left join por1 t11 on t11.itemcode=t10.itemcode and t10.mag<t10.u_prg_wip_qtadatrasf and t11.openqty>0
left join owor t12 on (t12.status ='P' or t12.status ='R') and t12.itemcode=t10.itemcode and t10.mag<t10.u_prg_wip_qtadatrasf
group by t10.docnum, t10.itemcode, t10.u_prg_wip_qtadatrasf, t10.Mag
)
as t20
group by t20.docnum, t20.itemcode
)
as t30) as C on c.DocNum=t10.docnum

)
as t40
INNER JOIN 
(select t10.docnum
,case when t10.itemcode='R00554' then t10.U_Stato_lavorazione else '' end as 'R00554'
,case when t10.itemcode='R00550' then t10.U_Stato_lavorazione else '' end as 'R00550'
,case when t10.itemcode='R00551' then t10.U_Stato_lavorazione else '' end as 'R00551'
,case when t10.itemcode='R00527' then t10.U_Stato_lavorazione else '' end as 'R00527'
,case when t10.itemcode='R00539' then t10.U_Stato_lavorazione else '' end as 'R00539'
,case when t10.itemcode='R00563' then t10.U_Stato_lavorazione else '' end as 'R00563'
,case when t10.itemcode='R00562' then t10.U_Stato_lavorazione else '' end as 'R00562'
,case when t10.itemcode='R00561' then t10.U_Stato_lavorazione else '' end as 'R00561'
,case when t10.itemcode='R00503' then t10.U_Stato_lavorazione else '' end as 'R00503'
,case when t10.itemcode='R00506' then t10.U_Stato_lavorazione else '' end as 'R00506'
,case when t10.itemcode='R00504' then t10.U_Stato_lavorazione else '' end as 'R00504'
,case when t10.itemcode='R00505' then t10.U_Stato_lavorazione else '' end as 'R00505'
,case when t10.itemcode='R00526' then t10.U_Stato_lavorazione else '' end as 'R00526'
,case when t10.itemcode='R00564' then t10.U_Stato_lavorazione else '' end as 'R00564'
,case when t10.itemcode='R00502' then t10.U_Stato_lavorazione else '' end as 'R00502'
,case when t10.itemcode='R00540' then t10.U_Stato_lavorazione else '' end as 'R00540'
,case when t10.itemcode='R00587' then t10.U_Stato_lavorazione else '' end as 'R00587'
,case when t10.itemcode='R00572' then t10.U_Stato_lavorazione else '' end as 'R00572'
,case when t10.itemcode='R00598' then t10.U_Stato_lavorazione else '' end as 'R00598'
,case when t10.itemcode='R00599' then t10.U_Stato_lavorazione else '' end as 'R00599'
,case when t10.itemcode='R00600' then t10.U_Stato_lavorazione else '' end as 'R00600'
from
(
select t0.docnum, t1.itemcode, t1.U_Stato_lavorazione
from owor t0 left join wor1 t1 on t0.docentry=t1.docentry
inner join orsc t2 on t2.VisResCode=t1.itemcode
where substring(t0.u_produzione,1,3)='INT' and (t0.status ='P' or t0.status='R') and t2.restype='M'
)
as t10)  T60 ON T60.DOCNUM=T40.DOCNUM
,
(select count(t0.docnum)/90 as 'ODP_chiusi_al_gg'
from owor t0
where substring(t0.u_produzione,1,3) ='INT' and t0.CloseDate>=getdate()-90 and t0.status='L') as t50

)
as t100 left join owor t101 on t100.docnum=t101.docnum"
        Cmd_SAP.ExecuteNonQuery()



        CNN.Close()
    End Sub


    Sub Inserimento_macchine_CAM()

        ComboBox1.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()



        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = CNN
        CMD_SAP_docentry.CommandText = "SELECT T0.[macchina] FROM CAM T0 GROUP BY T0.MACCHINA"

        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_docentry_reader.Read()

            ComboBox1.Items.Add(cmd_SAP_docentry_reader("macchina"))
            Indice = Indice + 1
        Loop
        cmd_SAP_docentry_reader.Close()
        CNN.Close()


    End Sub 'Inserisco le risorse nella combo box

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter
        Inserimento_macchine_CAM()
    End Sub

    Private Sub TabPage5_Click(sender As Object, e As EventArgs) Handles TabPage5.Enter
        Numero_odp_al_giorno(DataGridView4, RichTextBox1.Text)
    End Sub

    Sub Numero_odp_al_giorno(par_datagridview As DataGridView, par_numero_giorni As Integer)

        ' Pulizia DataGridView
        par_datagridview.Rows.Clear()
        par_datagridview.Columns.Clear()

        ' Connessione
        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Cnn.Open()

            Using CMD_SAP As New SqlCommand()
                CMD_SAP.Connection = Cnn

                ' =========================
                ' QUERY SQL DINAMICA
                ' =========================
                CMD_SAP.CommandText = "
DECLARE @cols NVARCHAR(MAX);
DECLARE @query NVARCHAR(MAX);

-- colonne da ieri in avanti (escludendo domeniche)
SET @cols = '[Scad],' + STUFF((
    SELECT ',' + QUOTENAME(CONVERT(VARCHAR(10), DATEADD(DAY, n, CAST(GETDATE() AS DATE)), 120))
    FROM (
        SELECT TOP (" & par_numero_giorni & ") ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS n
        FROM master..spt_values
    ) x
    WHERE DATENAME(WEEKDAY, DATEADD(DAY, n, CAST(GETDATE() AS DATE))) <> 'Sunday'
    ORDER BY n
    FOR XML PATH(''), TYPE
).value('.', 'NVARCHAR(MAX)'), 1, 1, '');

SET @query = '
WITH Dati AS (
    SELECT 
        t1.itemcode AS Risorsa,
        t2.resname AS NomeRisorsa,
        T3.ResGrpNam,
        COUNT(*) AS NumODP,
        -- Se scad, scriviamo Scad, altrimenti la data
        CONVERT(VARCHAR(10),
            CASE 
                WHEN t0.DueDate < CAST(GETDATE() AS DATE) THEN CAST(''1900-01-01'' AS DATE)
                ELSE t0.DueDate 
            END, 120) AS DataPivot,
        t2.u_ordine
    FROM owor t0
    INNER JOIN wor1 t1 ON t0.docentry = t1.docentry
    INNER JOIN orsc t2 ON t2.VisResCode = t1.itemcode
    INNER JOIN ORSB T3 ON T2.ResGrpCod = T3.ResGrpCod
    WHERE (t0.status IN (''P'',''R'')) 
      AND SUBSTRING(t0.u_produzione,1,3) = ''INT'' 
      AND t2.restype = ''M'' 
      AND t1.u_stato_lavorazione=''O''
    GROUP BY 
        t1.itemcode, t2.resname, T3.ResGrpNam, t0.DueDate, t2.u_ordine
)
SELECT Risorsa, NomeRisorsa, ResGrpNam, ' + @cols + '
FROM (
    SELECT 
        Risorsa, NomeRisorsa, ResGrpNam, 
        CASE WHEN DataPivot = ''1900-01-01'' THEN ''Scad'' ELSE DataPivot END AS DataPivot,
        NumODP, u_ordine
    FROM Dati
) AS src
PIVOT (
    SUM(NumODP)
    FOR DataPivot IN (' + @cols + ')
) AS pv
ORDER BY u_ordine;
';

EXEC (@query);
"

                ' =========================
                ' LETTURA DATI
                ' =========================
                Using reader As SqlDataReader = CMD_SAP.ExecuteReader()

                    ' Crea colonne DataGridView dinamicamente
                    For i As Integer = 0 To reader.FieldCount - 1
                        Dim colName As String = reader.GetName(i)

                        ' Colonne date (indice >= 3)
                        If i >= 3 Then
                            Dim dt As Date
                            If Date.TryParse(colName, dt) Then
                                ' intestazione verticale: "Dom" in cima, "13/12" sotto
                                Dim headerText As String = dt.ToString("ddd") & vbCrLf & dt.ToString("dd/MM")
                                par_datagridview.Columns.Add(colName, headerText)

                                ' stile intestazione
                                par_datagridview.Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                                par_datagridview.Columns(i).HeaderCell.Style.WrapMode = DataGridViewTriState.True
                                par_datagridview.Columns(i).HeaderCell.Style.Font = New Font(par_datagridview.Font.FontFamily, 8, FontStyle.Bold)

                                ' colore rosso se sabato o domenica
                                If dt.DayOfWeek = DayOfWeek.Saturday OrElse dt.DayOfWeek = DayOfWeek.Sunday Then
                                    par_datagridview.Columns(i).HeaderCell.Style.BackColor = Color.Red
                                    par_datagridview.Columns(i).HeaderCell.Style.ForeColor = Color.White
                                End If

                                ' larghezza ridotta fissa
                                par_datagridview.Columns(i).Width = 40
                                par_datagridview.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            Else
                                ' Colonna "Scad"
                                par_datagridview.Columns.Add(colName, colName)
                                par_datagridview.Columns(i).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
                                par_datagridview.Columns(i).HeaderCell.Style.WrapMode = DataGridViewTriState.True
                                par_datagridview.Columns(i).HeaderCell.Style.Font = New Font(par_datagridview.Font.FontFamily, 8, FontStyle.Bold)
                                par_datagridview.Columns(i).Width = 40
                                par_datagridview.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                            End If
                        Else
                            ' prime 3 colonne rimangono uguali
                            par_datagridview.Columns.Add(colName, colName)
                        End If
                    Next

                    ' Congela prime 3 colonne
                    For i As Integer = 0 To 2
                        par_datagridview.Columns(i).Frozen = True
                    Next

                    ' Auto-size prime 3 colonne
                    For i As Integer = 0 To 2
                        par_datagridview.Columns(i).AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
                    Next

                    ' =========================
                    ' RIEMPIMENTO RIGHE
                    ' =========================
                    While reader.Read()
                        Dim r As Integer = par_datagridview.Rows.Add()

                        For c As Integer = 0 To reader.FieldCount - 1
                            Dim val As Object = If(IsDBNull(reader(c)), 0, reader(c))
                            par_datagridview.Rows(r).Cells(c).Value = val
                            par_datagridview.Rows(r).Cells(c).Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                            ' Colora celle se 1-2 giallo, >=3 rosso (solo colonne data)
                            If c >= 3 Then
                                If Convert.ToInt32(val) >= 3 Then
                                    par_datagridview.Rows(r).Cells(c).Style.BackColor = Color.Red
                                    par_datagridview.Rows(r).Cells(c).Style.ForeColor = Color.White
                                ElseIf Convert.ToInt32(val) >= 1 Then
                                    par_datagridview.Rows(r).Cells(c).Style.BackColor = Color.Yellow
                                End If
                            End If
                        Next
                    End While

                End Using
            End Using
        End Using

        ' Scorri all'ultima riga
        Try
            par_datagridview.FirstDisplayedScrollingRowIndex = par_datagridview.RowCount - 1
        Catch
        End Try

        par_datagridview.ClearSelection()

    End Sub


    Private Sub CheckBox17_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox17.CheckedChanged
        If CheckBox17.Checked = False Then
            filtro_stato_completamento = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_stato_completamento = " and t50.stato_completamento='Completato' "
            Else
                filtro_stato_completamento = " and TRIM(stato_fase) LIKE '%S' "
            End If

        End If
        datagridview_carico_macchine()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) _
    Handles CheckBox1.CheckedChanged
        ApplicaFiltri()
        'Dim nomeColonna As String = "Prog"
        'filtra_datagridview_per_risorsa(DataGridView1, nomeColonna, CheckBox1.Checked, Label_n_ordini)

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        ApplicaFiltri()
    End Sub

    Sub ApplicaFiltri()

        Dim contatore As Integer = 0

        For Each row As DataGridViewRow In DataGridView1.Rows

            If row.IsNewRow Then Continue For

            Dim visibile As Boolean = True

            ' Filtro CheckBox1 → colonna Prog
            If CheckBox1.Checked Then
                Dim cell = row.Cells("Prog")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            ' Filtro CheckBox2 → colonna Prog_p
            If CheckBox23.Checked Then
                Dim cell = row.Cells("Prog_p")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If


            If CheckBox2.Checked Then
                Dim cell = row.Cells("T_MAN")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox3.Checked Then
                Dim cell = row.Cells("T_auto")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox21.Checked Then
                Dim cell = row.Cells("Pant_p")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox20.Checked Then
                Dim cell = row.Cells("Pant_a")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox4.Checked Then
                Dim cell = row.Cells("Tornio_")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox5.Checked Then
                Dim cell = row.Cells("Goodway")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox6.Checked Then
                Dim cell = row.Cells("Doosan")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox7.Checked Then
                Dim cell = row.Cells("Doosan_4")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox8.Checked Then
                Dim cell = row.Cells("Doos6")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox24.Checked Then
                Dim cell = row.Cells("Multi")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox9.Checked Then
                Dim cell = row.Cells("Haas_2")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox10.Checked Then
                Dim cell = row.Cells("Haas_3")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox11.Checked Then
                Dim cell = row.Cells("Haas_5")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox12.Checked Then
                Dim cell = row.Cells("Haas_6")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox13.Checked Then
                Dim cell = row.Cells("Famup")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox14.Checked Then
                Dim cell = row.Cells("Awea")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox15.Checked Then
                Dim cell = row.Cells("Stozza")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox16.Checked Then
                Dim cell = row.Cells("Sald")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox19.Checked Then
                Dim cell = row.Cells("Trap")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox25.Checked Then
                Dim cell = row.Cells("Vulcanizzazione")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If

            If CheckBox22.Checked Then
                Dim cell = row.Cells("Finitura")
                visibile = visibile AndAlso
                (cell.Style.BackColor = Color.LightGreen OrElse
                 cell.Style.BackColor = Color.Yellow)
            End If





            row.Visible = visibile

            If visibile Then contatore += 1

        Next

        Label_n_ordini.Text = contatore.ToString()

    End Sub

    Sub filtra_datagridview_per_risorsa(
    par_datagridview As DataGridView,
    par_nome_colonna As String,
    par_mostra As Boolean,
    label_contatore As Label)

        Dim contatore As Integer = 0

        For Each row As DataGridViewRow In par_datagridview.Rows

            If row.IsNewRow Then Continue For

            Dim cell = row.Cells(par_nome_colonna)

            Dim isColorata As Boolean =
            cell.Style.BackColor = Color.LightGreen OrElse
            cell.Style.BackColor = Color.Yellow

            ' Se la checkbox è spuntata → mostra solo le righe colorate
            If par_mostra Then
                row.Visible = isColorata
            Else
                ' Checkbox non spuntata → mostra tutto
                row.Visible = True
            End If

            ' Conta solo le righe visibili
            If row.Visible Then
                contatore += 1
            End If

        Next

        label_contatore.Text = contatore.ToString()

    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox18_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox18.CheckedChanged
        filtra()
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox19_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox19.CheckedChanged
        ApplicaFiltri()
    End Sub





    Private Sub TextBox5_Leave(sender As Object, e As EventArgs) Handles TextBox5.Leave
        filtra()
    End Sub


    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        filtra()
    End Sub



    Private Sub TextBox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave
        filtra()
    End Sub





    Private Sub TextBox4_Leave(sender As Object, e As EventArgs) Handles TextBox4.Leave
        filtra()
    End Sub



    Private Sub TextBox6_Leave(sender As Object, e As EventArgs) Handles TextBox6.Leave
        filtra()
    End Sub



    Private Sub Button8_Click(sender As Object, e As EventArgs)
        Dim contatore As Integer = 0


        Do While contatore < DataGridView1.Rows.Count

            If DataGridView1.Rows(contatore).Cells(DataGridView1.Columns.IndexOf(Seleziona)).Value = True Then

                DataGridView1.Rows(contatore).Cells(columnName:="Lavorazione_").Value = lavorazione
                MsgBox(lavorazione)
                Dim CNN As New SqlConnection
                CNN.ConnectionString = Homepage.sap_tirelli
                CNN.Open()
                Dim CMD_SAP As New SqlCommand
                CMD_SAP.Connection = CNN

                CMD_SAP.CommandText = "UPDATE owor set U_LAVORAZIONE='" & lavorazione & "' WHERE DOCNUM ='" & DataGridView1.Rows(contatore).Cells(DataGridView1.Columns.IndexOf(ODP_)).Value & "'"
                CMD_SAP.ExecuteNonQuery()
                CNN.Close()


            End If
            contatore = contatore + 1
        Loop
        contatore = 0
    End Sub


    Sub Inserimento_risorse()
        ComboBox3.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()



        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = CNN
        CMD_SAP_docentry.CommandText = "SELECT T0.[VisResCode] AS 'Risorsa',T0.[ResGrpCod] as 'Tipo macchina', T0.[ResName] as 'Nome' FROM ORSC T0 WHERE T0.[ResType] ='M' ORDER BY t0.[resname]"

        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_docentry_reader.Read()
            Elenco_macchinari(Indice) = cmd_SAP_docentry_reader("Risorsa")

            ComboBox3.Items.Add(cmd_SAP_docentry_reader("Nome"))


            If cmd_SAP_docentry_reader("Risorsa") = "R00554" Then
                ComboBox3.SelectedIndex = Indice
            End If
            Indice = Indice + 1
        Loop
        cmd_SAP_docentry_reader.Close()
        CNN.Close()


    End Sub

    Sub start_statistiche()

        Inserimento_risorse()

        dettaglio_backlog_ordini()

    End Sub

    Private Sub TabPage4_Click(sender As Object, e As EventArgs) Handles TabPage4.Enter
        start_statistiche()

    End Sub

    Sub backlog_ordini()
        If Homepage.ERP_provenienza = "SAP" Then
            DataGridView_backlog.Rows.Clear()
            Dim CNN As New SqlConnection
            CNN.ConnectionString = Homepage.sap_tirelli
            CNN.Open()



            Dim CMD_SAP_docentry As New SqlCommand
            Dim cmd_SAP_docentry_reader As SqlDataReader

            CMD_SAP_docentry.Connection = CNN
            CMD_SAP_docentry.CommandText = "Select t10.u_produzione, sum(case when  t10.stato ='P' then 1 else 0 end) as 'P', sum(case when  t10.stato ='R' then 1 else 0 end) as 'R', sum(case when  t10.stato ='C' then 1 else 0 end) as 'C'
from
(
SELECT t0.u_produzione,CASE WHEN T0.U_STATO='Completato' then 'C' else T0.[Status] end as 'Stato' 
FROM OWOR T0 WHERE (T0.[Status] ='P' OR T0.[Status] ='R') AND  SUBSTRING(T0.[U_PRODUZIONE],1,3)='INT'
)
as t10

group by
t10.u_produzione"

            cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader


            Do While cmd_SAP_docentry_reader.Read()

                DataGridView_backlog.Rows.Add(cmd_SAP_docentry_reader("u_produzione"), cmd_SAP_docentry_reader("P"), cmd_SAP_docentry_reader("R"), cmd_SAP_docentry_reader("C"))

            Loop
            cmd_SAP_docentry_reader.Close()
            CNN.Close()

            DataGridView_backlog.ClearSelection()
        End If
    End Sub 'Inserisco le risorse nella combo box



    Sub dettaglio_backlog_ordini()

        DataGridView3.Rows.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()



        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = CNN
        CMD_SAP_docentry.CommandText = "Select t20.itemcode ,t22.itemname, sum(case when t20.Pronto = 'Primo' then 1 else 0 end) as 'Ordini_lavorabili', count(t20.pronto) as 'Ordini_TOT' , sum(case when t20.Pronto = 'Primo' then T20.[PlannedQty] else 0 end)/60 as 'Tempo_lavorabile(h)', sum(T20.[PlannedQty])/60 as 'Tempo_TOT(h)'
from
(
Select t10.docentry,T10.[ItemCode], T10.[PlannedQty], t10.u_stato_lavorazione, t10.linenum, min(t11.linenum) as 'Primo', case when t10.linenum=min(t11.linenum) then 'Primo' else 'Non_primo' end as 'Pronto'
from
(
SELECT t0.docentry,T1.[ItemCode], T1.[PlannedQty], t1.u_stato_lavorazione, t1.linenum 
FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
inner join orsc t2 on t2.visrescode=t1.itemcode and T2.[ResType] ='M' 

WHERE (T0.[Status] ='P' or  T0.[Status] ='R') and substring(T1.[ItemCode],1,1) = 'R'
)
as t10 left join wor1 t11 on t11.docentry=t10.docentry and t11.u_stato_lavorazione='O'
inner join orsc t12 on t12.visrescode=t11.itemcode and T12.[ResType] ='M' 

group by t10.docentry,T10.[ItemCode], T10.[PlannedQty], t10.u_stato_lavorazione, t10.linenum
)
as t20
left join orsc t21 on t21.visrescode=t20.itemcode
left join oitm t22 on t22.itemcode=t20.itemcode
where t20.u_stato_lavorazione='O'
group by t20.itemcode,t21.u_ordine, t22.itemname
order by t21.u_ordine, t20.itemcode"

        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader


        Do While cmd_SAP_docentry_reader.Read()

            DataGridView3.Rows.Add(cmd_SAP_docentry_reader("itemcode"), cmd_SAP_docentry_reader("itemname"), cmd_SAP_docentry_reader("Ordini_lavorabili"), cmd_SAP_docentry_reader("Ordini_TOT"), cmd_SAP_docentry_reader("Tempo_lavorabile(h)"), cmd_SAP_docentry_reader("Tempo_TOT(h)"))

        Loop
        cmd_SAP_docentry_reader.Close()
        CNN.Close()

        DataGridView3.ClearSelection()
    End Sub 'Inserisco le risorse nella combo box





    Sub start_carico_macchine()

        datagridview_carico_macchine()



        Max_numerazione()
        Carico_macchine_grafico_programmazione()
        Carico_macchine_grafico_torni()
        Carico_macchine_grafico_frese()
        Carico_macchine_grafico_altro()

        backlog_ordini()
    End Sub



    Private Sub CheckBox21_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox21.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox22_CheckedChanged(sender As Object, e As EventArgs)
        filtra()
    End Sub

    Private Sub CheckBox23_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox23.CheckedChanged
        ApplicaFiltri()
        'Dim nomeColonna As String = "Prog_p"
        'filtra_datagridview_per_risorsa(DataGridView1, nomeColonna, CheckBox23.Checked, Label_n_ordini)
    End Sub

    Private Sub CheckBox22_CheckedChanged_1(sender As Object, e As EventArgs)
        filtra()
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        If TextBox14.Text = Nothing Then
            filtro_cliente = ""
        Else
            filtro_cliente = " and t50.U_Final_customer_name    Like '%%" & TextBox14.Text & "%%'  "
        End If

        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

        If TextBox5.Text = Nothing Then
            filtro_odp = ""
        Else
            filtro_odp = " and t50.docnum    Like '%%" & TextBox5.Text & "%%'  "
        End If

        Timer1.Stop()
        Timer1.Start()


    End Sub



    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'datagridview_carico_macchine()

        Timer1.Stop()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = Nothing Then
            filtro_codice = ""
        Else
            filtro_codice = " and t50.itemcode  Like '%%" & TextBox1.Text & "%%'  "
        End If

        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = Nothing Then
            filtro_descrizione = ""
        Else
            filtro_descrizione = " and t50.itemname  Like '%%" & TextBox2.Text & "%%'  "
        End If

        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = Nothing Then
            filtro_stato = ""
        Else
            filtro_stato = " and t50.status  Like '%%" & TextBox3.Text & "%%'  "
        End If

        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = Nothing Then
            filtro_disegno = ""
        Else
            filtro_disegno = " and t50.disegno  Like '%%" & TextBox4.Text & "%%'  "
        End If

        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If TextBox6.Text = Nothing Then
            filtro_commessa = ""
        Else
            filtro_commessa = " and t50.U_PRG_AZS_Commessa  Like '%%" & TextBox6.Text & "%%'  "
        End If

        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        If TextBox16.Text = Nothing Then
            filtro_mat_prima = ""
        Else
            filtro_mat_prima = " and t50.mat_prima  Like '%%" & TextBox16.Text & "%%'  "
        End If

        Timer1.Stop()
        Timer1.Start()
    End Sub

    Private Sub CheckBox20_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox20.CheckedChanged
        ApplicaFiltri()
    End Sub


    Private Sub CheckBox24_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox24.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox22_CheckedChanged_2(sender As Object, e As EventArgs)

    End Sub

    Private Sub CheckBox25_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox25.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub CheckBox22_CheckedChanged_3(sender As Object, e As EventArgs) Handles CheckBox22.CheckedChanged
        ApplicaFiltri()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContextMenuStripChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContextMenuStripChanged

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        datagridview_carico_macchine()
    End Sub

    Private Sub TableLayoutPanel3_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel3.Paint

    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Numero_odp_al_giorno(DataGridView4, RichTextBox1.Text)
    End Sub

    Private Sub Carico_macchine_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        datagridview_carico_macchine()
    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click

        ExportVisibleColumnsToExcel(DataGridView1)
    End Sub

    Public Sub ExportVisibleColumnsToExcel(ByVal par_datagridview As DataGridView)

        Dim excelApp As New Excel.Application
        excelApp.Visible = True

        Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add()
        Dim excelWorksheet As Excel.Worksheet =
        CType(excelWorkbook.Worksheets(1), Excel.Worksheet)

        ' ===== Foglio 2: codici unici =====
        Dim excelWorksheet2 As Excel.Worksheet =
        CType(excelWorkbook.Worksheets.Add(After:=excelWorksheet), Excel.Worksheet)
        excelWorksheet2.Name = "Codici"

        Dim rowCount As Integer = par_datagridview.Rows.Count

        ' ===== colonne visibili =====
        Dim visibleColumns = par_datagridview.Columns _
        .Cast(Of DataGridViewColumn) _
        .Where(Function(c) c.Visible) _
        .ToList()

        Dim visibleColCount As Integer = visibleColumns.Count

        ' ===== intestazioni =====
        For colIndex As Integer = 0 To visibleColCount - 1
            excelWorksheet.Cells(1, colIndex + 1) = visibleColumns(colIndex).HeaderText
        Next

        ' ===== dimensioni celle =====
        Dim excelRowHeight As Double = 80
        Dim excelColWidth As Double = 20

        For i = 1 To visibleColCount
            CType(excelWorksheet.Columns(i), Excel.Range).ColumnWidth = excelColWidth
        Next

        excelWorksheet.Rows("1:" & (rowCount + 1)).RowHeight = excelRowHeight

        ' ===== gestione codici =====
        Dim codiciInseriti As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim nextRowFoglio2 As Integer = 1   ' colonna L

        ' ===== ciclo righe =====
        For row As Integer = 0 To rowCount - 1

            ' --- lettura ItemCode DIRETTA (anche se colonna nascosta) ---
            Dim codice As String = ""

            If par_datagridview.Columns.Contains("ItemCode") AndAlso
           par_datagridview.Rows(row).Cells("ItemCode").Value IsNot Nothing Then

                codice = par_datagridview.Rows(row).Cells("ItemCode").Value.ToString().Trim()

                If codice <> "" AndAlso Not codiciInseriti.Contains(codice) Then
                    codiciInseriti.Add(codice)
                    excelWorksheet2.Cells(nextRowFoglio2, 12).Value = codice ' colonna L
                    nextRowFoglio2 += 1
                End If
            End If

            ' --- export colonne visibili ---
            For colIndex As Integer = 0 To visibleColCount - 1

                Dim col = visibleColumns(colIndex)
                Dim value = par_datagridview.Rows(row).Cells(col.Index).Value
                Dim cell = CType(excelWorksheet.Cells(row + 2, colIndex + 1), Excel.Range)

                If TypeOf value Is Bitmap Then
                    Try
                        Dim originalImage As Bitmap = CType(value, Bitmap)

                        Dim scaleX = cell.Width / originalImage.Width
                        Dim scaleY = cell.Height / originalImage.Height
                        Dim scale = Math.Min(scaleX, scaleY)

                        Dim newWidth = CInt(originalImage.Width * scale)
                        Dim newHeight = CInt(originalImage.Height * scale)

                        Dim resizedImage As New Bitmap(newWidth, newHeight)
                        Using g = Graphics.FromImage(resizedImage)
                            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                            g.DrawImage(originalImage, 0, 0, newWidth, newHeight)
                        End Using

                        Clipboard.SetImage(resizedImage)
                        excelWorksheet.Paste(cell)

                        Dim pic = excelWorksheet.Pictures(excelWorksheet.Pictures.Count)
                        pic.Left = cell.Left + (cell.Width - newWidth) / 2
                        pic.Top = cell.Top + (cell.Height - newHeight) / 2
                        pic.Placement = Excel.XlPlacement.xlMoveAndSize

                    Catch ex As Exception
                        MessageBox.Show("Errore immagine: " & ex.Message)
                    End Try

                Else
                    If value IsNot Nothing Then
                        If IsNumeric(value) Then
                            cell.NumberFormat = "0"
                            cell.Value = CDbl(value)
                        Else
                            cell.NumberFormat = "@"
                            cell.Value = value.ToString()
                        End If
                    Else
                        cell.Value = ""
                    End If
                End If
            Next
        Next

        ' ===== allineamento =====
        Dim usedRange = excelWorksheet.Range(
        excelWorksheet.Cells(1, 1),
        excelWorksheet.Cells(rowCount + 1, visibleColCount))

        usedRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
        usedRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

        ' ===== salvataggio =====
        Dim sfd As New SaveFileDialog With {.Filter = "Excel (*.xlsx)|*.xlsx"}
        If sfd.ShowDialog() = DialogResult.OK Then
            excelWorkbook.SaveAs(sfd.FileName)
            MessageBox.Show("Esportazione completata", "OK",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        ' ===== cleanup =====
        ReleaseComObject(excelWorksheet2)
        ReleaseComObject(excelWorksheet)
        ReleaseComObject(excelWorkbook)
        ReleaseComObject(excelApp)

    End Sub

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Form_visualizza_picture.Show()
        Magazzino.visualizza_picture(codice_pic, Form_visualizza_picture.PictureBox1)
    End Sub

    't40.Seleziona, t40.docnum, t40.U_Lavorazione, t40.status, t40.postdate, t40.duedate,getdate()+(ROW_NUMBER() OVER(ORDER BY t40.priorità asc)/t50.ODP_chiusi_al_gg) as 'Cons', DATEDIFF(DD,T40.DUEDATE,GETDATE()) AS 'ANT/RIT',t40.stato_mat, t40.Arrivo_MAT,case when t62.MAt_prima is null then '' else t62.MAt_prima end as 'Mat_prima' , t40.Priorità, t40.ItemCode, t40.itemname, t40.Disegno, t40.[Famiglia disegno], t40.nesting, t40.disp,t40.stato_completamento,t40.PlannedQty,t40.U_PRG_AZS_Commessa, case when t61.U_Final_customer_name is null and t63.U_Final_customer_name is null then t64.custmrName when t61.U_Final_customer_name is null then  t63.U_Final_customer_name else t61.U_Final_customer_name  end as 'U_Final_customer_name'
End Class


