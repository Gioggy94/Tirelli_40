Imports System.Data.SqlClient
Imports Microsoft.Office.Interop


Public Class form_Spare_Parts
    ' Private operazione As String
    Private filtro_docnum As String
    Private filtro_docnum_text As String
    Private filtro_cardname_text As String
    Private filtro_causale_text As String
    Private filtro_stato_text As String
    Private filtro_rif_Cliente As String
    Private filtro_reparto As String
    Private filtro_famiglia As String
    Private N_ordine As Integer
    Private inizializzazione As Boolean = False
    Private filtro_owner As String
    Private filtro_matr_CDS As String
    ' Private divisione As String = "BRB01"

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Sub inizializza_form()
        inizializzazione = True
        DateTimePicker4.Value = DateAdd("d", -750, Today)
        DateTimePicker1.Value = DateAdd("d", 150, Today)
        inizializzazione = False

        lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)

    End Sub

    Sub lista_ordini_cliente_codice(par_datagridview As DataGridView, par_datetimepicker_inizio As DateTimePicker, par_datetimepicker_fine As DateTimePicker)

        Dim Cnn1 As New SqlConnection
        par_datagridview.Rows.Clear()
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader



        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "
select *
from
(
select 

    T30.[DocNum],
	min(
    case 
        when t30.A_Machinery > 0 then 'Machinery'
        when t30.B_Options > 0 then 'Options & Upgrades'
        when t30.C_Spare > 0 then 'Spare parts'
        when t30.D_Service > 0 then 'Service'
        when t30.E_Packaging > 0 then 'Packaging'
        else 'Others'
    end
) as 'Famiglia',
    T30.[CardCode], 
    T30.[CardName], 
t30.numatcard,
t30.Commento_interno,
    T30.[Final_bp], 
	t30.U_CausCons,
t30.owner,
coalesce(t30.u_matrcds,'') as 'u_matrcds',
		t30.U_Uffcompetenza,
    T30.[DocDate], 
    T30.[DocDueDate], 
	t30.n_articoli,
	t30.doctotal,
	case 
	when t30.Articoli_OK>=t30.n_articoli then 'Spedibile'
	when t30.Articoli_OK+t30.Articoli_trasferibile>=t30.n_articoli then 'Trasferibile'
when t30.Articoli_OK+t30.Articoli_trasferibile+t30.articoli_In_Accettazione>=t30.n_articoli then 'In_Accettazione'
when t30.Articoli_OK+t30.Articoli_trasferibile+t30.Articoli_Approvv>=t30.n_articoli and t30.ultimo_arrivo <=getdate()-1 then 'IN_approvv_Scaduto'
	when t30.Articoli_OK+t30.Articoli_trasferibile+t30.Articoli_Approvv>=t30.n_articoli then 'IN_approvv'
	when t30.Articoli_OK+t30.Articoli_trasferibile+t30.Articoli_Approvv+t30.Articoli_Da_ordinare>=t30.n_articoli then 'Da_ordinare'
	else 'Altro'
	end as 'Stato'
	,t30.ultimo_arrivo
from
(
select  
    T20.[DocNum], 

	sum(t20.A_Machinery) as 'A_Machinery'
	,sum(t20.B_Options) as 'B_Options'
     ,sum(T20.C_Spare) as 'C_Spare'
	 ,sum(t20.D_Service) as 'D_Service'
	 ,sum(t20.E_Packaging) as 'E_Packaging'
	 ,sum(t20.f_others) as 'F_others'
       


    ,T20.[CardCode], 
    T20.[CardName], 
	
t20.numatcard,
t20.Commento_interno,
    T20.[Final_bp], 
	t20.U_CausCons,
t20.owner,
		t20.U_Uffcompetenza,
		t20.u_matrcds,
    T20.[DocDate], 
    T20.[DocDueDate], 
	SUM(CASE WHEN T20.itemcode <>'L' THEN 1 ELSE 0 END) AS 'N_Articoli',
	t20.doctotal,
    SUM(CASE WHEN T20.itemcode <>'L' and T20.Stato = 'OK' THEN 1 ELSE 0 END) AS 'Articoli_OK',
    SUM(CASE WHEN T20.itemcode <>'L' and T20.Stato = 'Trasferibile' THEN 1 ELSE 0 END) AS 'Articoli_Trasferibile',
SUM(CASE WHEN T20.itemcode <>'L' and T20.Stato = 'In_Accettazione' THEN 1 ELSE 0 END) AS 'Articoli_In_Accettazione',
    SUM(CASE WHEN T20.itemcode <>'L' and T20.Stato = 'Approvv' THEN 1 ELSE 0 END) AS 'Articoli_Approvv',
    SUM(CASE WHEN T20.itemcode <>'L' and T20.Stato = 'Da_ordinare' THEN 1 ELSE 0 END) AS 'Articoli_Da_ordinare',
	SUM(CASE WHEN T20.itemcode <>'L' and T20.Stato = 'Altro' THEN 1 ELSE 0 END) AS 'Articoli_Altro'
	,case when
max(t20.cons_forn) is null then max(t20.cons_odp)
	when max(t20.cons_odp) is null then max(t20.cons_forn)
	when
	max(t20.cons_odp)>= max(t20.cons_forn) then max(t20.cons_odp) else max(t20.cons_forn) end  as 'Ultimo_arrivo'
from
(
SELECT   
   
    T10.[DocNum], 
	t10.A_Machinery, t10.B_Options,t10.C_Spare,t10.D_Service, t10.E_Packaging, t10.F_Others,
	   	  	t10.doctotal,
    T10.[CardCode], 
    T10.[CardName], 
	t10.u_matrcds,
t10.numatcard,
t10.Commento_interno,
    T10.[Final_bp], 
	t10.U_CausCons,
t10.owner,
		t10.U_Uffcompetenza,
    T10.[DocDate], 
    T10.[DocDueDate], 
    T10.[ItemCode], 
	t10.U_Codice_BRB,
    T10.[Dscription], 
    T10.whscode,
    T10.[OpenQty], 
    T10.LINETOTAL,
	case
	when T10.[U_Datrasferire]<=0 then 'OK'
	when T10.A_MAG>=T10.[U_Datrasferire] then 'Trasferibile'
when T10.A_MAG+t10.A_mag_01>=T10.[U_Datrasferire] then 'In_Accettazione'
when T10.A_MAG+t10.A_MAG_01+t10.A_MAG_Trattamento>=T10.[U_Datrasferire] then 'In_Trattamento'
	when T10.A_MAG+t10.A_MAG_01+t10.A_MAG_Trattamento+t10.Da_trattare>=T10.[U_Datrasferire] then 'Da_Trattare'
	when T10.A_MAG+t10.Ord>=T10.[U_Datrasferire] then 'Approvv'
	when t10.disp<0 then 'Da_ordinare'
	Else 'Altro'
	end as 'Stato',
    T10.[U_Trasferito], 
    T10.[U_Datrasferire] ,
    T10.A_MAG ,
	T10.ord,
    T10.Disp ,
	MinShipDates.docnum as 'OA',
	minshipdates.cardname as 'Forn',
	minshipdates.min_shipdate as 'Cons_forn',
	MinodpDates.docnum as 'ODP',
	MinodpDates.u_produzione as 'Prod',
MinodpDates.duedate as 'Cons_odp'
FROM
(
SELECT 
    T1.[DocNum], 
    T1.doctotal,
    T1.[CardCode], 
    T1.[CardName], 
    T1.numatcard,
    CAST(COALESCE(T1.U_Commento_Interno, '') AS VARCHAR) AS 'Commento_interno',
    T2.CardName AS 'Final_bp', 
    T1.U_CausCons,
    CONCAT(T6.LastName, ' ', T6.FirstName) AS 'Owner',
    T1.U_Uffcompetenza,
    T1.[DocDate], 
    T1.[DocDueDate], 
    T0.[ItemCode], 
    T5.U_Codice_BRB,
    T0.[Dscription], 
    T0.WhsCode,
    T0.[OpenQty], 
    T0.LineTotal, 
    T0.[U_Trasferito], 
    T0.[U_Datrasferire],
    T1.U_Matrcds,
    SUM(CASE WHEN T11.CODE = '01' THEN 1 ELSE 0 END) AS 'A_Machinery',
    SUM(CASE WHEN T11.CODE = '02' THEN 1 ELSE 0 END) AS 'B_Options',
    SUM(CASE WHEN T11.CODE = '03' THEN 1 ELSE 0 END) AS 'C_Spare',
    SUM(CASE WHEN T11.CODE = '04' THEN 1 ELSE 0 END) AS 'D_Service',
    SUM(CASE WHEN T11.CODE = '05' THEN 1 ELSE 0 END) AS 'E_Packaging',
    SUM(CASE WHEN T11.CODE = '06' THEN 1 ELSE 0 END) AS 'F_Others',
    (SELECT SUM(T3.OnHand) 
     FROM OITW T3 
     WHERE T3.ItemCode = T0.ItemCode AND T3.WhsCode NOT IN ('BWIP', 'WIP','01','16','Clavter')) AS 'A_MAG',

   (SELECT SUM(T3.OnHand) 
     FROM OITW T3 
     WHERE T3.ItemCode = T0.ItemCode AND T3.WhsCode = '01') AS 'A_MAG_01',
	 
			  (SELECT SUM(t3.onhand) FROM oitw t3 
INNER JOIN owhs t4 ON t4.WhsCode = T3.[WHSCODE] 
            WHERE t3.itemcode = T0.itemcode AND t3.whscode='Clavter' ) as 'A_MAG_Trattamento' ,
				  (SELECT SUM(t3.onhand) FROM oitw t3 
INNER JOIN owhs t4 ON t4.WhsCode = T3.[WHSCODE] 
            WHERE t3.itemcode = T0.itemcode AND t3.whscode='16' ) as 'Da_trattare' ,
    (SELECT SUM(T5.OnHand + T5.OnOrder - T5.IsCommited) 
     FROM OITW T5 
     WHERE T5.ItemCode = T0.ItemCode AND SUBSTRING(T5.ItemCode, 1, 1) <> 'L') AS 'Disp',
    (SELECT SUM(T6.OnOrder) 
     FROM OITW T6
     WHERE T6.ItemCode = T0.ItemCode AND SUBSTRING(T6.ItemCode, 1, 1) <> 'L') AS 'Ord'
FROM 
    RDR1 T0  
    INNER JOIN ORDR T1 ON T0.[DocEntry] = T1.[DocEntry]
    LEFT JOIN OCRD T2 ON T2.CardCode = T1.U_CodiceBP 
    LEFT JOIN OWHS T4 ON T4.WhsCode = T0.WhsCode
    INNER JOIN OITM T5 ON T5.ItemCode = T0.ItemCode
    LEFT JOIN [TIRELLI_40].[dbo].OHEM T6 ON T1.OwnerCode = T6.EmpID
    LEFT JOIN OACT T9 ON T9.[AcctCode] = T0.[AcctCode]
    LEFT JOIN [@FAMIGLIA_VENDITA] T11 ON T11.CODE = T9.U_FAMIGLIAVENDITA
WHERE    1 = 1         AND T0.[OpenQty] > 0 
 AND T1.[DocDueDate] >= @DataInizio
      AND T1.[DocDueDate] <= @DataFine


    
GROUP BY 
    T1.[DocNum], 
    T1.doctotal,
    T1.[CardCode], 
    T1.[CardName], 
    T1.numatcard,
    CAST(COALESCE(T1.U_Commento_Interno, '') AS VARCHAR),
    T2.CardName, 
    T1.U_CausCons,
    CONCAT(T6.LastName, ' ', T6.FirstName),
    T1.U_Uffcompetenza,
    T1.[DocDate], 
    T1.[DocDueDate], 
    T0.[ItemCode], 
    T5.U_Codice_BRB,
    T0.[Dscription], 
    T0.WhsCode,
    T0.[OpenQty], 
    T0.LineTotal, 
    T0.[U_Trasferito], 
    T0.[U_Datrasferire],
    T1.U_Matrcds


         

) AS T10

left JOIN (

    SELECT 
        t11.docentry,
        t11.itemcode,
        MIN(t11.shipdate) AS min_shipdate
		,t12.docnum
		,t12.CardName

    FROM 
        por1 t11 inner join opor t12 on t11.docentry=T12.docentry
		INNER JOIN owhs t13 ON t13.WhsCode = T11.whscode 
      WHERE 
        t11.OpenQty > 0
    GROUP BY 
        t11.docentry, t11.itemcode,t12.docnum
		,t12.CardName
) AS MinShipDates 
                    on T10.[ItemCode] = MinShipDates.itemcode
                    AND T10.[U_Datrasferire] > T10.A_MAG +T10.A_MAG_01 +t10.A_MAG_Trattamento+t10.Da_trattare
                    AND T10.ord+T10.A_MAG +T10.A_MAG_01 +t10.A_MAG_Trattamento+t10.Da_trattare >= T10.[U_Datrasferire]

					left JOIN (

    SELECT 
        t11.docentry,
        t11.itemcode,
        MIN(t11.duedate) AS duedate
		,t11.DocNum, 
		t11.u_produzione

    FROM 
        owor t11 INNER JOIN owhs t12 ON t12.WhsCode = T11.warehouse 
    WHERE 
        (t11.status='P' or t11.status='R') 
    GROUP BY 
       t11.docentry,
        t11.itemcode
		,t11.DocNum, 
		t11.u_produzione,
t11.duedate
) AS Minodpdates 
                    on T10.[ItemCode] = MinodpDates.itemcode
                    AND T10.[U_Datrasferire] > T10.A_MAG +T10.A_MAG_01
                    AND T10.ord+T10.A_MAG +T10.A_MAG_01 >= T10.[U_Datrasferire]
					)
					as t20

					group by 
    T20.[DocNum], 
    T20.[CardCode], 
    T20.[CardName], 
        t20.numatcard,
		t20.Commento_interno,
    T20.[Final_bp], 
	t20.U_CausCons,
        t20.owner,
		t20.U_Uffcompetenza,
    T20.[DocDate], 
    T20.[DocDueDate],
	t20.doctotal,
	t20.u_matrcds
		--,t20.Machinery, t20.Options,t20.Spare,t20.Service, t20.Packaging, t20.Others
	
	)
	as t30
	group by  T30.[DocNum],
	 T30.[CardCode], 
    T30.[CardName], 
t30.numatcard,
t30.Commento_interno,
    T30.[Final_bp], 
	t30.U_CausCons,
t30.owner,
coalesce(t30.u_matrcds,''),
t30.U_Uffcompetenza,
    T30.[DocDate], 
    T30.[DocDueDate], 
	t30.n_articoli,
	t30.doctotal,
	t30.Articoli_OK,
	t30.articoli_trasferibile,
t30.articoli_In_Accettazione,
	t30.Articoli_Approvv,
	t30.Ultimo_arrivo,
	t30.Articoli_Da_ordinare
)
as t40
where 0=0  " & filtro_docnum_text & " " & filtro_cardname_text & " " & filtro_causale_text & " " & filtro_stato_text & filtro_rif_Cliente & filtro_reparto & filtro_famiglia & filtro_owner & filtro_matr_CDS & "
order by T40.[DocDueDate]"

        ' 🔑 Parametri al posto delle date concatenate
        CMD_SAP_2.Parameters.Add("@DataInizio", SqlDbType.DateTime).Value = par_datetimepicker_inizio.Value.Date
        CMD_SAP_2.Parameters.Add("@DataFine", SqlDbType.DateTime).Value = par_datetimepicker_fine.Value.Date

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader_2("DocNum"),
                                     cmd_SAP_reader_2("famiglia"),
        cmd_SAP_reader_2("CardCode"),
        cmd_SAP_reader_2("CardName"),
        cmd_SAP_reader_2("Final_bp"),
        cmd_SAP_reader_2("u_matrcds"),
        cmd_SAP_reader_2("Numatcard"),
         cmd_SAP_reader_2("commento_interno"),
        cmd_SAP_reader_2("U_CausCons"),
        cmd_SAP_reader_2("Owner"),
        cmd_SAP_reader_2("U_Uffcompetenza"),
        cmd_SAP_reader_2("DocDate"),
        cmd_SAP_reader_2("DocDueDate"),
        cmd_SAP_reader_2("n_articoli"),
        cmd_SAP_reader_2("doctotal"),
        cmd_SAP_reader_2("Stato"),
        cmd_SAP_reader_2("ultimo_arrivo"))
        Loop



        Cnn1.Close()

    End Sub
    Sub dettaglio_codice(par_datagridview As DataGridView, par_filtro_docnum As String, par_docnum As Integer, par_tipo_appoggio As String)
        ODP_Tree.PULISCI_APPOGGIO(Homepage.ID_SALVATO, par_tipo_appoggio)
        par_datagridview.Rows.Clear()
        Dim contatore As Integer = 1

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader




        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "declare @docnum as integer


set @docnum='" & par_docnum & "'

SELECT   
    T10.ocrcode,
    T10.[DocNum], 
    T10.[CardCode], 
    T10.[CardName], 
    T10.[Final_bp], 
	t10.U_CausCons,
		t10.U_Uffcompetenza,
    T10.[DocDate], 
    T10.[DocDueDate], 
    T10.[ItemCode], 
	t10.U_Codice_BRB,
    T10.[Dscription], 
T10.[QryGroup20],
    T10.whscode,
    T10.[OpenQty], 
    T10.LINETOTAL,
	case
	when T10.[U_Datrasferire]<=0 then 'OK'
	when T10.A_MAG>=T10.[U_Datrasferire] then 'Trasferibile'
	when T10.A_MAG+t10.A_MAG_01>=T10.[U_Datrasferire] then 'In_Accettazione'
	when T10.A_MAG+t10.A_MAG_01+t10.A_MAG_Trattamento>=T10.[U_Datrasferire] then 'In_Trattamento'
	when T10.A_MAG+t10.A_MAG_01+t10.A_MAG_Trattamento+t10.Da_trattare>=T10.[U_Datrasferire] then 'Da_Trattare'
	when T10.A_MAG+t10.Ord>=T10.[U_Datrasferire] then 'Approvv'
	when t10.disp<0 then 'Da_ordinare'
ELSE ''
	end as 'Stato',
    T10.[U_Trasferito], 
    T10.[U_Datrasferire] ,
    T10.A_MAG ,
	T10.ord,
    T10.Disp ,
	MinShipDates.docnum as 'OA',
	coalesce(minshipdates.cardname,'') as 'Forn',
	minshipdates.min_shipdate as 'Cons_forn',
	coalesce(cast(coalesce(odp_esatto.docnum,MinodpDates.docnum) as varchar),'') as 'ODP',
	coalesce(odp_esatto.u_produzione,MinodpDates.u_produzione) as 'Prod',
coalesce(odp_esatto.duedate,MinodpDates.duedate) as 'Cons_odp'
,coalesce(odp_esatto.u_prg_azs_commessa,MinodpDates.u_prg_azs_commessa) as 'Comm_odp'
FROM
(
   SELECT 
        T0.ocrcode,
        T1.[DocNum], 
        T1.[CardCode], 
        T1.[CardName], 
        t2.cardname as 'Final_bp', 
		t1.U_CausCons,
		t1.U_Uffcompetenza,
        T1.[DocDate], 
        T1.[DocDueDate], 
        T0.[ItemCode], 
		t5.U_Codice_BRB,
        T0.[Dscription], 
T5.[QryGroup20],
        t0.whscode,
        T0.[OpenQty], 
        T0.LINETOTAL, 
        T0.[U_Trasferito], 
        T0.[U_Datrasferire] ,

        (SELECT SUM(t3.onhand) FROM oitw t3 
INNER JOIN owhs t4 ON t4.WhsCode = T3.[WHSCODE] 
            WHERE t3.itemcode = T0.itemcode AND t3.whscode <>'BWIP' AND t3.whscode <>'01' AND t3.whscode <>'WIP' AND t3.whscode <>'16' AND t3.whscode <>'Clavter') as 'A_MAG' ,
			
        (SELECT SUM(t3.onhand) FROM oitw t3 
INNER JOIN owhs t4 ON t4.WhsCode = T3.[WHSCODE] 
            WHERE t3.itemcode = T0.itemcode AND t3.whscode='01' ) as 'A_MAG_01' ,
			  (SELECT SUM(t3.onhand) FROM oitw t3 
INNER JOIN owhs t4 ON t4.WhsCode = T3.[WHSCODE] 
            WHERE t3.itemcode = T0.itemcode AND t3.whscode='Clavter' ) as 'A_MAG_Trattamento' ,
				  (SELECT SUM(t3.onhand) FROM oitw t3 
INNER JOIN owhs t4 ON t4.WhsCode = T3.[WHSCODE] 
            WHERE t3.itemcode = T0.itemcode AND t3.whscode='16' ) as 'Da_trattare' ,
        (SELECT SUM(t5.onhand+t5.onorder-t5.iscommited) FROM oitw t5 
INNER JOIN owhs t4 ON t4.WhsCode = T5.[WHSCODE] 
            WHERE t5.itemcode = T0.itemcode and substring(t5.itemcode,1,1)<>'L' ) as 'Disp' ,
        (SELECT SUM(t6.onorder) FROM oitw t6
INNER JOIN owhs t4 ON t4.WhsCode = T6.[WHSCODE] 
            WHERE t6.itemcode = T0.itemcode and substring(t6.itemcode,1,1)<>'L' ) as 'Ord' 

 FROM 
        RDR1 T0  
        INNER JOIN ORDR T1 ON T0.[DocEntry] = T1.[DocEntry]
        LEFT JOIN ocrd t2 ON t2.cardcode = T1.u_codicebp 
        INNER JOIN owhs t4 ON t4.WhsCode = T0.whscode 
		inner join oitm t5 on t5.itemcode=t0.itemcode
    WHERE 

         T0.[OpenQty] > 0 " & filtro_docnum & " and substring(t0.itemcode,1,1)<>'L'
) AS T10

left JOIN (

select t40.itemcode,
        t40.min_shipdate
		,t40.docentry
		, t40.Linenum
		,t43.docnum
		,t43.cardname
from
(
select t30.itemcode,
        t30.min_shipdate
		,t30.docentry
		, min(t31.linenum) as 'Linenum'
from
(
select t20.itemcode,
        t20.min_shipdate
		,min(t21.docentry) as 'docentry'
from
(
    SELECT 

        t11.itemcode,
        MIN(t11.shipdate) AS min_shipdate
		

    FROM 
        por1 t11 inner join opor t12 on t11.docentry=T12.docentry
		INNER JOIN owhs t13 ON t13.WhsCode = T11.whscode 
    WHERE 
        t11.OpenQty > 0
    GROUP BY 
        t11.itemcode
		)
		as t20 inner join por1 t21 on t21.shipdate=T20.min_shipdate and t21.itemcode=t20.itemcode and t21.OpenQty > 0
		INNER JOIN owhs t22 ON t22.WhsCode = T21.whscode 
		GROUP BY 
        t20.itemcode,
        t20.min_shipdate
		)
		as t30 inner join por1 t31 on t31.shipdate=T30.min_shipdate and t31.itemcode=t30.itemcode and t31.OpenQty > 0 and t31.docentry=t30.docentry
		INNER JOIN owhs t32 ON t32.WhsCode = T31.whscode 

		group by t30.itemcode,
        t30.min_shipdate
		,t30.docentry
		)
		as t40 inner join por1 t41 on t41.shipdate=T40.min_shipdate and t41.itemcode=t40.itemcode and t41.OpenQty > 0 and t41.docentry=t40.docentry
		INNER JOIN owhs t42 ON t42.WhsCode = T41.whscode 
		left join opor t43 on t43.docentry=t40.docentry
         group by 
		t40.itemcode,
        t40.min_shipdate
		,t40.docentry
		, t40.Linenum
		,t43.docnum
		,t43.cardname
) AS MinShipDates 
                    on T10.[ItemCode] = MinShipDates.itemcode
                    AND T10.[U_Datrasferire] > T10.A_MAG +T10.A_MAG_01 +t10.A_MAG_Trattamento+t10.Da_trattare
                    AND T10.ord+T10.A_MAG+T10.A_MAG_01 +t10.A_MAG_Trattamento+t10.Da_trattare  >= T10.[U_Datrasferire]

					left JOIN (


           SELECT T30.DOCENTRY,T30.ITEMCODE,T30.DUEDATE,T31.DOCNUM,T31.U_PRODUZIONE,T31.U_PRG_AZS_Commessa
	FROM
	(
	SELECT MIN(T21.DOCENTRY) AS 'DOCENTRY',  t20.itemcode, T20.duedate
	FROM
	(
	SELECT 
        
        t11.itemcode,
        MIN(t11.duedate) AS duedate
	

		

    FROM 
        owor t11 INNER JOIN owhs t12 ON t12.WhsCode = T11.warehouse 
         LEFT JOIN [TIRELLI_40].[DBO].APPOGGIO T1 ON T11.DOCNUM=T1.VALORE AND T1.TIPO='" & par_tipo_appoggio & "' AND T1.UTENTE=" & Homepage.ID_SALVATO & "
    WHERE 
        (t11.status='P' or t11.status='R') AND T1.VALORE IS NULL
    GROUP BY 

        t11.itemcode
      


)
AS T20 LEFT JOIN OWOR T21 ON T21.ITEMCODE=T20.ITEMCODE AND (t21.status='P' or t21.status='R')  AND T21.DUEDATE=T20.DUEDATE
INNER JOIN owhs t22 ON t22.WhsCode = T21.warehouse 
GROUP BY t20.itemcode, T20.duedate
)
AS T30
LEFT JOIN OWOR T31 ON T31.ITEMCODE=T30.ITEMCODE AND (t31.status='P' or t31.status='R')  AND T31.DUEDATE=T30.DUEDATE AND T31.DOCENTRY=T30.DOCENTRY
INNER JOIN owhs t32 ON t32.WhsCode = T31.warehouse 
) AS Minodpdates 
                    on T10.[ItemCode] = MinodpDates.itemcode
                    AND T10.[U_Datrasferire] > T10.A_MAG +T10.A_MAG_01 +t10.A_MAG_Trattamento+t10.Da_trattare
                    AND T10.ord+T10.A_MAG +T10.A_MAG_01+t10.A_MAG_Trattamento+t10.Da_trattare >= T10.[U_Datrasferire]

left JOIN (
        
select t20.docentry, t20.itemcode, t21.duedate ,t21.PlannedQty,
		t21.DocNum, 
		t21.u_produzione
        ,t21.u_prg_azs_commessa
from
(
    SELECT 
        min(t11.docentry) as 'docentry',
        t11.itemcode
    FROM 
        owor t11 INNER JOIN owhs t12 ON t12.WhsCode = T11.warehouse 
LEFT JOIN [TIRELLI_40].[DBO].APPOGGIO T1 ON T11.DOCNUM=T1.VALORE AND T1.TIPO='" & par_tipo_appoggio & "' AND T1.UTENTE=" & Homepage.ID_SALVATO & "
    WHERE 
        (t11.status='P' or t11.status='R') and  T1.VALORE IS NULL and (cast(substring(t11.U_PRG_AZS_Commessa,2,99) as varchar)=cast(@docnum as varchar) or cast(@docnum as varchar) = t11.OriginNum)
  group by t11.itemcode
)
as t20 left join owor t21 on t20.docentry=t21.docentry
) AS odp_esatto 
                    on T10.[ItemCode] = odp_esatto.itemcode
                    AND T10.[U_Datrasferire] > T10.A_MAG +t10.A_MAG_Trattamento+t10.Da_trattare
                    AND T10.ord+T10.A_MAG+t10.A_MAG_Trattamento+t10.Da_trattare  >= T10.[U_Datrasferire]
					and odp_esatto.PlannedQty>= T10.[U_Datrasferire]

"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()
            ODP_Tree.AGGIUNGI_RECORD_APPOGGIO(Homepage.ID_SALVATO, par_tipo_appoggio, cmd_SAP_reader_2("ODP"))
            par_datagridview.Rows.Add(
        cmd_SAP_reader_2("ocrcode"),
        contatore,
        cmd_SAP_reader_2("DocNum"),
        cmd_SAP_reader_2("CardCode"),
        cmd_SAP_reader_2("CardName"),
        cmd_SAP_reader_2("Final_bp"),
        cmd_SAP_reader_2("U_CausCons"),
        cmd_SAP_reader_2("U_Uffcompetenza"),
        cmd_SAP_reader_2("DocDate"),
        cmd_SAP_reader_2("DocDueDate"),
        cmd_SAP_reader_2("ItemCode"),
        cmd_SAP_reader_2("U_Codice_BRB"),
        cmd_SAP_reader_2("Dscription"),
        cmd_SAP_reader_2("whscode"),
        cmd_SAP_reader_2("OpenQty"),
        cmd_SAP_reader_2("LINETOTAL"),
        cmd_SAP_reader_2("Stato"),
        cmd_SAP_reader_2("U_Trasferito"),
        cmd_SAP_reader_2("U_Datrasferire"),
        cmd_SAP_reader_2("A_MAG"),
        cmd_SAP_reader_2("ord"),
        cmd_SAP_reader_2("Disp"),
        cmd_SAP_reader_2("OA"),   ' MinShipDates.docnum
        cmd_SAP_reader_2("Forn"), ' MinShipDates.cardname
        cmd_SAP_reader_2("Cons_forn"), ' MinShipDates.min_shipdate
        cmd_SAP_reader_2("ODP"), ' MinodpDates.docnum
        cmd_SAP_reader_2("Prod"),
        cmd_SAP_reader_2("Cons_ODP"),
        cmd_SAP_reader_2("comm_ODP"),
         cmd_SAP_reader_2("QryGroup20"))

            If cmd_SAP_reader_2("ODP") <> "" Then
                dettaglio_ODP(par_datagridview, par_filtro_docnum, cmd_SAP_reader_2("ODP"), contatore, cmd_SAP_reader_2("comm_ODP"), cmd_SAP_reader_2("DocNum"), par_tipo_appoggio)
            End If
            contatore = contatore + 1

        Loop


        Cnn1.Close()
        par_datagridview.ClearSelection()
    End Sub

    Sub dettaglio_ordine(par_richtextbox As RichTextBox, par_docnum As Integer, par_richtextbox2 As RichTextBox)


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader




        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "declare @docnum as integer


set @docnum='" & par_docnum & "'


SELECT   
    coalesce(t0.comments,'') as 'comments'
,COALESCE(T0.U_Commento_interno,'') as 'Commento_interno'
from ordr t0
where t0.docnum=@docnum"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() Then
            par_richtextbox.Text = cmd_SAP_reader_2("comments")
            par_richtextbox2.Text = cmd_SAP_reader_2("Commento_interno")


        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    Sub dettaglio_ODP(par_datagridview As DataGridView, par_filtro_docnum As String, par_odp As String, par_livello As String, par_commessa As String, par_docnum_oc As String, par_tipo_appoggio As String)

        Dim contatore As Integer = 1

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader




        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "
select 
t10.docnum,
t10.cardcode,
t10.CardName
,t10.Final_BP,
t10.u_produzione
,t10.Uff,
t10.postdate,
t10.duedate,
t10.itemcode,
t10.U_Codice_BRB,
t10.itemname,
T10.[QryGroup20],
T10.[wareHouse],
t10.PlannedQty,
t10.linetotal,
t10.U_PRG_WIP_QtaSpedita,
t10.U_PRG_WIP_QtaDaTrasf,
t10.A_MAG ,
       t10.Disp ,
	   t10.ord,
        case
when T10.[U_PRG_WIP_QtaDaTrasf]<=0 then 'OK'
when T10.A_MAG>=T10.[U_PRG_WIP_QtaDaTrasf] then 'Trasferibile'
when T10.A_MAG+t10.a_mag_01>=T10.[U_PRG_WIP_QtaDaTrasf] then 'In_Accettazione'
when T10.A_MAG+t10.A_MAG_01+t10.A_MAG_Trattamento>=T10.[U_PRG_WIP_QtaDaTrasf] then 'In_Trattamento'
	when T10.A_MAG+t10.A_MAG_01+t10.A_MAG_Trattamento+t10.Da_trattare>=T10.[U_PRG_WIP_QtaDaTrasf] then 'Da_Trattare'
when T10.A_MAG+t10.Ord>=T10.[U_PRG_WIP_QtaDaTrasf] then 'Approvv'
when t10.disp<0 then 'Da_ordinare'
else ''
end as 'Stato',
	MinShipDates.docnum as 'OA',
	coalesce(minshipdates.cardname,'') as 'Forn',
	minshipdates.min_shipdate as 'Cons_forn',
	coalesce(cast(coalesce(odp_esatto.docnum,MinodpDates.docnum) as varchar),'') as 'ODP',
	coalesce(odp_esatto.u_produzione,MinodpDates.u_produzione) as 'Prod',
coalesce(odp_esatto.duedate,MinodpDates.duedate) as 'Cons_odp'
,coalesce(odp_esatto.u_prg_azs_commessa,MinodpDates.u_prg_azs_commessa) as 'comm_odp'
from
(
select t0.docnum,
t0.cardcode,
t2.CardName
,t4.cardname as 'Final_BP',
t0.u_produzione
,'' as 'Uff',
t0.postdate,
t0.duedate,
t1.itemcode,
t5.U_Codice_BRB,
T5.[QryGroup20],
t5.itemname,
T1.[wareHouse],
t1.PlannedQty,
0 as 'linetotal',
t1.U_PRG_WIP_QtaSpedita,
t1.U_PRG_WIP_QtaDaTrasf,
        (SELECT SUM(t3.onhand) FROM oitw t3 
INNER JOIN owhs t4 ON t4.WhsCode = T3.[WHSCODE] 
            WHERE t3.itemcode = T1.itemcode AND t3.whscode <> 'BWIP' AND t3.whscode <> 'WIP' AND t3.whscode <> '01' AND t3.whscode <>'16' AND t3.whscode <>'Clavter' ) as 'A_MAG' ,
			     (SELECT SUM(t3.onhand) FROM oitw t3 
INNER JOIN owhs t4 ON t4.WhsCode = T3.[WHSCODE] 
            WHERE t3.itemcode = T1.itemcode  AND t3.whscode = '01' ) as 'A_MAG_01' ,
			  (SELECT SUM(t3.onhand) FROM oitw t3 
INNER JOIN owhs t4 ON t4.WhsCode = T3.[WHSCODE] 
            WHERE t3.itemcode = T0.itemcode AND t3.whscode='Clavter' ) as 'A_MAG_Trattamento' ,
				  (SELECT SUM(t3.onhand) FROM oitw t3 
INNER JOIN owhs t4 ON t4.WhsCode = T3.[WHSCODE] 
            WHERE t3.itemcode = T0.itemcode AND t3.whscode='16' ) as 'Da_trattare' ,
        (SELECT SUM(t5.onhand+t5.onorder-t5.iscommited) FROM oitw t5 
INNER JOIN owhs t4 ON t4.WhsCode = T5.[WHSCODE] 
            WHERE t5.itemcode = T1.itemcode and substring(t5.itemcode,1,1)<>'L' ) as 'Disp' ,
        (SELECT SUM(t6.onorder) FROM oitw t6
INNER JOIN owhs t4 ON t4.WhsCode = T6.[WHSCODE] 
            WHERE t6.itemcode = T1.itemcode and substring(t6.itemcode,1,1)<>'L' ) as 'Ord' 
from owor t0
inner join wor1 t1 on t0.docentry=t1.docentry
left join ocrd t2 on t2.cardcode=t0.cardcode
left join ordr t3 on t3.docnum =t0.OriginNum
left join ocrd t4 on t4.cardcode=t3.U_CodiceBP
left join oitm t5 on t5.ItemCode=t1.itemcode
where t0.docnum='" & par_odp & "' and substring(t1.itemcode,1,1)<>'L' and substring(t1.itemcode,1,1)<>'R'
)
as t10

left JOIN (

   SELECT T40.mIN_DOCENTRY,T40.ITEMCODE,T40.min_shipdate,t42.docnum
		,t42.CardName
FROM
(

SELECT T30.mIN_DOCENTRY,T30.ITEMCODE,T30.min_shipdate, MIN(T31.LINENUM) AS 'LINENUM'
FROM
(
SELECT MIN(T21.DOCENTRY) AS 'mIN_DOCENTRY'
, T20.ITEMCODE,T20.min_shipdate
FROM
(
    SELECT 
     
        t11.itemcode,
        MIN(t11.shipdate) AS min_shipdate
		

    FROM 
        por1 t11 
		INNER JOIN owhs t13 ON t13.WhsCode = T11.whscode 
    WHERE 
        t11.OpenQty > 0
    GROUP BY 
        t11.itemcode
		--,t12.docnum
		--,t12.CardName
		)
		AS T20 INNER JOIN POR1 T21 ON T21.SHIPDATE=T20.MIN_SHIPDATE  AND T21.OpenQty>0 AND T21.ITEMCODE=T20.ITEMCODE
		INNER JOIN owhs t23 ON t23.WhsCode = T21.whscode 
		GROUP BY  T20.ITEMCODE,T20.min_shipdate
		)
		AS T30 INNER JOIN POR1 T31 ON T31.SHIPDATE=T30.MIN_SHIPDATE AND T31.DOCENTRY=T30.MIN_DOCENTRY AND T31.OpenQty>0 AND T31.ITEMCODE=T30.ITEMCODE
INNER JOIN owhs t33 ON t33.WhsCode = T31.whscode 
GROUP BY T30.mIN_DOCENTRY,T30.ITEMCODE,T30.min_shipdate
)
AS T40 INNER JOIN POR1 T41 ON T41.SHIPDATE=T40.MIN_SHIPDATE AND T41.DOCENTRY=T40.MIN_DOCENTRY AND T41.OpenQty>0 AND T41.ITEMCODE=T40.ITEMCODE AND T41.LINENUM=T40.LINENUM
inner join opor t42 on t41.docentry=T42.docentry
INNER JOIN owhs t43 ON t43.WhsCode = T41.whscode 
) AS MinShipDates 
                    on T10.[ItemCode] = MinShipDates.itemcode
                    AND T10.[U_PRG_WIP_QtaDaTrasf] > T10.A_MAG +t10.A_MAG_Trattamento+t10.Da_trattare
                    AND T10.ord+T10.A_MAG +t10.A_MAG_Trattamento+t10.Da_trattare  >= T10.[U_PRG_WIP_QtaDaTrasf]

					left JOIN (

           SELECT T30.DOCENTRY,T30.ITEMCODE,T30.DUEDATE,T31.DOCNUM,T31.U_PRODUZIONE,T31.U_PRG_AZS_Commessa
	FROM
	(
	SELECT MIN(T21.DOCENTRY) AS 'DOCENTRY',  t20.itemcode, T20.duedate
	FROM
	(
	SELECT 
        
        t11.itemcode,
        MIN(t11.duedate) AS duedate
	

		

    FROM 
        owor t11 INNER JOIN owhs t12 ON t12.WhsCode = T11.warehouse 
LEFT JOIN [TIRELLI_40].[DBO].APPOGGIO T1 ON T11.DOCNUM=T1.VALORE AND T1.TIPO='" & par_tipo_appoggio & "' AND T1.UTENTE=" & Homepage.ID_SALVATO & "

    WHERE 
        (t11.status='P' or t11.status='R')  and  T1.VALORE IS NULL
    GROUP BY 

        t11.itemcode
      


)
AS T20 LEFT JOIN OWOR T21 ON T21.ITEMCODE=T20.ITEMCODE AND (t21.status='P' or t21.status='R')  AND T21.DUEDATE=T20.DUEDATE
INNER JOIN owhs t22 ON t22.WhsCode = T21.warehouse 
GROUP BY t20.itemcode, T20.duedate
)
AS T30
LEFT JOIN OWOR T31 ON T31.ITEMCODE=T30.ITEMCODE AND (t31.status='P' or t31.status='R')  AND T31.DUEDATE=T30.DUEDATE AND T31.DOCENTRY=T30.DOCENTRY
INNER JOIN owhs t32 ON t32.WhsCode = T31.warehouse
) AS Minodpdates 
                    on T10.[ItemCode] = MinodpDates.itemcode
                    AND T10.[U_PRG_WIP_QtaDaTrasf] > T10.A_MAG +T10.A_MAG_01 +t10.A_MAG_Trattamento+t10.Da_trattare
                    AND T10.ord+T10.A_MAG +T10.A_MAG_01  +t10.A_MAG_Trattamento+t10.Da_trattare>= T10.[U_PRG_WIP_QtaDaTrasf]

left JOIN (
        
select t20.docentry, t20.itemcode, t21.duedate ,t21.PlannedQty,
		t21.DocNum, 
		t21.u_produzione
        ,t21.u_prg_azs_commessa
from
(
    SELECT 
        min(t11.docentry) as 'docentry',
        t11.itemcode
 

		

    FROM 
        owor t11 INNER JOIN owhs t12 ON t12.WhsCode = T11.warehouse 
LEFT JOIN [TIRELLI_40].[DBO].APPOGGIO T1 ON T11.DOCNUM=T1.VALORE AND T1.TIPO='" & par_tipo_appoggio & "' AND T1.UTENTE=" & Homepage.ID_SALVATO & "

    WHERE 
        (t11.status='P' or t11.status='R') and  T1.VALORE IS NULL AND (cast(substring(t11.U_PRG_AZS_Commessa,2,99) as varchar)=cast('" & par_commessa & "' as varchar) or cast('" & par_docnum_oc & "' as varchar) = t11.OriginNum)
  group by t11.itemcode
)
as t20 left join owor t21 on t20.docentry=t21.docentry
) AS odp_esatto 
                    on T10.[ItemCode] = odp_esatto.itemcode
                    AND t10.U_PRG_WIP_QtaDaTrasf > T10.A_MAG  +t10.A_MAG_Trattamento+t10.Da_trattare
                    AND T10.ord+T10.A_MAG +t10.A_MAG_Trattamento+t10.Da_trattare >= t10.U_PRG_WIP_QtaDaTrasf
					and odp_esatto.PlannedQty>= t10.U_PRG_WIP_QtaDaTrasf"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()
            ODP_Tree.AGGIUNGI_RECORD_APPOGGIO(Homepage.ID_SALVATO, par_tipo_appoggio, cmd_SAP_reader_2("ODP"))
            par_datagridview.Rows.Add(
        "",
        par_livello & "-" & contatore,
        par_odp,
        cmd_SAP_reader_2("CardCode"),
        cmd_SAP_reader_2("CardName"),
        cmd_SAP_reader_2("Final_bp"),
        cmd_SAP_reader_2("u_produzione"),
        cmd_SAP_reader_2("Uff"),
        cmd_SAP_reader_2("postDate"),
        cmd_SAP_reader_2("DueDate"),
        cmd_SAP_reader_2("ItemCode"),
        cmd_SAP_reader_2("U_Codice_BRB"),
        cmd_SAP_reader_2("itemname"),
        cmd_SAP_reader_2("wareHouse"),
        cmd_SAP_reader_2("PlannedQty"),
        cmd_SAP_reader_2("LINETOTAL"),
        cmd_SAP_reader_2("Stato"),
        cmd_SAP_reader_2("U_PRG_WIP_QtaSpedita"),
        cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"),
        cmd_SAP_reader_2("A_MAG"),
        cmd_SAP_reader_2("ord"),
        cmd_SAP_reader_2("Disp"),
        cmd_SAP_reader_2("OA"),   ' MinShipDates.docnum
        cmd_SAP_reader_2("Forn"), ' MinShipDates.cardname
        cmd_SAP_reader_2("Cons_forn"), ' MinShipDates.min_shipdate
        cmd_SAP_reader_2("ODP"), ' MinodpDates.docnum
        cmd_SAP_reader_2("Prod"),
        cmd_SAP_reader_2("Cons_ODP"),
       cmd_SAP_reader_2("comm_ODP"),
       cmd_SAP_reader_2("QryGroup20"))


            If cmd_SAP_reader_2("ODP") <> "" Then
                dettaglio_ODP(par_datagridview, par_filtro_docnum, cmd_SAP_reader_2("ODP"), par_livello & "-" & contatore, cmd_SAP_reader_2("comm_ODP"), par_docnum_oc, par_tipo_appoggio)


            End If
            contatore = contatore + 1
        Loop


        Cnn1.Close()

    End Sub





    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)
    End Sub

    Private Sub form_Spare_Parts_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.BackColor = Homepage.colore_sfondo
        inizializza_check_listbox()
        inizializza_check_listbox2()
        inizializza_form()

    End Sub

    Sub inizializza_check_listbox()
        ' Aggiungi elementi alla CheckedListBox
        CheckedListBox1.Items.Add("DA DEFINIRE")
        CheckedListBox1.Items.Add("COMMERCIALE")
        CheckedListBox1.Items.Add("SERVICE")
        CheckedListBox1.Items.Add("SERVICE INTERVENTO")

        ' Seleziona automaticamente gli ultimi due elementi
        CheckedListBox1.SetItemChecked(1, True) ' Seleziona "SERVICE"
        CheckedListBox1.SetItemChecked(2, True)
        CheckedListBox1.SetItemChecked(3, True)
        CheckedListBox1.SetItemChecked(0, True) ' Seleziona "SERVICE INTERVENTO"
    End Sub


    Sub inizializza_check_listbox2()
        Dim par_checkbox As CheckedListBox = CheckedListBox2
        ' Aggiungi elementi alla CheckedListBox
        par_checkbox.Items.Add("Machinery")
        par_checkbox.Items.Add("Options & Upgrades")
        par_checkbox.Items.Add("Spare parts")
        par_checkbox.Items.Add("Service")
        par_checkbox.Items.Add("Packaging")
        par_checkbox.Items.Add("Others")


        ' Seleziona automaticamente gli ultimi due elementi
        par_checkbox.SetItemChecked(1, True) ' Seleziona "SERVICE"
        par_checkbox.SetItemChecked(2, True)
        par_checkbox.SetItemChecked(3, True)
        par_checkbox.SetItemChecked(4, True)
        par_checkbox.SetItemChecked(5, True)
        par_checkbox.SetItemChecked(0, True) ' Seleziona "SERVICE INTERVENTO"
    End Sub



    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick

        If e.RowIndex >= 0 Then

            N_ordine = DataGridView2.Rows(e.RowIndex).Cells(columnName:="Docnum_").Value

            If e.ColumnIndex = DataGridView2.Columns.IndexOf(Docnum_) Then


                Form_nuova_offerta.Show()

                Form_nuova_offerta.TextBox10.Text = DataGridView2.Rows(e.RowIndex).Cells(columnName:="Docnum_").Value
                Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                Form_nuova_offerta.inizializzazione_form(DataGridView2.Rows(e.RowIndex).Cells(columnName:="Docnum_").Value, "ORDR", "RDR1", DataGridView2.Rows(e.RowIndex).Cells(columnName:="Docnum_").Value)

            Else
                filtro_docnum = " and t1.docnum= " & DataGridView2.Rows(e.RowIndex).Cells(columnName:="Docnum_").Value
                dettaglio_ordine(RichTextBox1, DataGridView2.Rows(e.RowIndex).Cells(columnName:="Docnum_").Value, RichTextBox2)
                dettaglio_codice(DataGridView1, filtro_docnum, DataGridView2.Rows(e.RowIndex).Cells(columnName:="Docnum_").Value, "SPARE_TREE")

            End If



        End If
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Itemcode) Then

                Magazzino.Codice_SAP = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Itemcode").Value

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

            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(ODP) Then
                Dim new_form_odp_form = New ODP_Form
                new_form_odp_form.docnum_odp = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ODP").Value
                new_form_odp_form.Show()
                new_form_odp_form.inizializza_form(DataGridView1.Rows(e.RowIndex).Cells(columnName:="ODP").Value)

            End If





            If e.ColumnIndex = DataGridView1.Columns.IndexOf(OA) Then


                Form_nuova_offerta.Show()

                Form_nuova_offerta.TextBox10.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="OA").Value
                Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                Form_nuova_offerta.inizializzazione_form(DataGridView1.Rows(e.RowIndex).Cells(columnName:="OA").Value, "OPOR", "POR1", "Ordine_acquisto")




            End If




        End If

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Try
            If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Da_programmare").Value = "Y" Then


                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Orange


            End If
            If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "OK" Then


                DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Color.Lime

            ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "Trasferibile" Then

                DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Color.YellowGreen

            ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "In_Accettazione" Then

                DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Color.Yellow

            ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "In_Trattamento" Or DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "Da_Trattare" Then

                DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Color.Orange

            ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "Approvv" Then

                DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Color.DarkOrange

            ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "Da_ordinare" Then

                DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Color.OrangeRed


            End If

            If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disp").Value < 0 Then
                DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disp").Style.ForeColor = Color.Red
            End If
        Catch ex As Exception

        End Try
    End Sub

    ' Dizionario per assegnare colori univoci alla colonna "famiglia"
    Private FamigliaColori As New Dictionary(Of String, Color)
    Private Shared rnd As New Random() ' Generatore di colori casuali

    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        ' Try
        Dim dgv As DataGridView = DataGridView2
            Dim colonna As String = dgv.Columns(e.ColumnIndex).Name

            ' Colorazione della colonna "Stato_ord"
            If colonna = "Stato_ord" AndAlso e.Value IsNot Nothing Then
                Select Case e.Value.ToString()
                    Case "Spedibile"
                        e.CellStyle.BackColor = Color.Lime
                    Case "Trasferibile"
                        e.CellStyle.BackColor = Color.YellowGreen
                    Case "In_Accettazione"
                        e.CellStyle.BackColor = Color.Yellow
                Case "In_Trattamento"

                    e.CellStyle.BackColor = Color.Orange
                Case "Da_Trattare"
                    e.CellStyle.BackColor = Color.Orange
                Case "IN_approvv_Scaduto"
                        e.CellStyle.BackColor = Color.Gold

                    Case "IN_approvv"
                        e.CellStyle.BackColor = Color.DarkOrange
                    Case "Da_ordinare"
                        e.CellStyle.BackColor = Color.OrangeRed
                End Select
            End If



            ' Colorazione della colonna "Causale_"
            If colonna = "Causale_" AndAlso e.Value IsNot Nothing Then
                If e.Value.ToString() = "V" Then
                    e.CellStyle.BackColor = Color.Lime
                Else
                    e.CellStyle.BackColor = Color.OrangeRed
                End If
            End If

            ' Colorazione dinamica della colonna "famiglia"
            If colonna = "Famiglia" AndAlso e.Value IsNot Nothing Then
                Dim valore As String = e.Value.ToString().Trim()

                ' Se il valore non ha ancora un colore assegnato, generane uno nuovo
                If Not FamigliaColori.ContainsKey(valore) Then
                    FamigliaColori(valore) = GeneraColoreCasuale()
                End If

                ' Applica il colore alla cella
                e.CellStyle.BackColor = FamigliaColori(valore)
            End If

        'Catch ex As Exception
        ' Ignora errori su celle vuote o fuori intervallo
        'End Try
    End Sub

    ' Funzione per generare un colore casuale evitando colori troppo scuri
    Private Function GeneraColoreCasuale() As Color
        Return Color.FromArgb(255, rnd.Next(120, 255), rnd.Next(120, 255), rnd.Next(120, 255))
    End Function

    ' Rinfresca i colori dopo il caricamento dei dati
    Private Sub DataGridView2_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles DataGridView2.DataBindingComplete
        DataGridView2.Invalidate() ' Forza il refresh per assicurarsi che i colori vengano applicati correttamente
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            filtro_docnum_text = ""

        Else
            filtro_docnum_text = " And T40.[DocNum] LIKE '%" & TextBox1.Text & "%'"
        End If
        lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = "" Then
            filtro_cardname_text = ""

        Else
            filtro_cardname_text = " And (t40.[CardName] LIKE '%" & TextBox2.Text & "%' or t40.[Final_bp] LIKE '%" & TextBox2.Text & "%') "

        End If
        lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Then
            filtro_causale_text = ""

        Else
            filtro_causale_text = " And T40.[u_causcons] LIKE '%" & TextBox3.Text & "%'"
        End If
        lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = "" Then
            filtro_stato_text = ""

        Else
            filtro_stato_text = " And T40.[stato] LIKE '%" & TextBox4.Text & "%'"
        End If
        lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)
    End Sub


    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text = "" Then
            filtro_rif_Cliente = ""

        Else
            filtro_rif_Cliente = " And T40.[numatcard] LIKE '%" & TextBox5.Text & "%'"
        End If
        lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Layout_documenti.ComboBox1.SelectedIndex = 1

        Layout_documenti.TextBox1.Text = N_ordine
        Layout_documenti.Show()
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs)
        'If TextBox6.Text = "" Then
        '    filtro_reparto = ""

        'Else
        '    filtro_reparto = " And T40.[U_Uffcompetenza] LIKE '%" & TextBox6.Text & "%'"
        'End If
        'lista_ordini_cliente_codice(DataGridView2, operazione, ComboBox1.Text, DateTimePicker4, DateTimePicker1)
    End Sub

    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        If inizializzazione = False Then
            lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)
        End If

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If inizializzazione = False Then
            lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Scheda_commessa_Pianificazione.ExportVisibleColumnsToExcel(DataGridView2)
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

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        ' Definisci la variabile filtro_reporto come vuota

        Dim listaFiltri As New List(Of String) ' Lista per memorizzare i singoli filtri

        ' Cicla attraverso tutti gli elementi della CheckedListBox
        For i As Integer = 0 To CheckedListBox1.Items.Count - 1
            ' Verifica se l'elemento è selezionato
            If CheckedListBox1.GetItemChecked(i) Then
                ' Aggiungi il filtro per l'elemento selezionato
                listaFiltri.Add("t40.U_Uffcompetenza = '" & CheckedListBox1.Items(i).ToString() & "'")
            End If
        Next

        ' Se ci sono elementi selezionati, crea il filtro con OR
        If listaFiltri.Count > 0 Then
            filtro_reparto = " AND (" & String.Join(" OR ", listaFiltri) & ")"
        End If

        ' Mostra o usa il filtro generato
        ' MessageBox.Show(filtro_reparto)

        ' Puoi anche assegnare filtro_reporto a una variabile di livello superiore se necessario
        lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)
    End Sub

    Sub aggiorna_commento_interno(par_commento As String, par_docnum As Integer)
        par_commento = Replace(par_commento, "'", " ")
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE ORDR SET U_COMMENTO_INTERNO='" & par_commento & "' WHERE DOCNUM=" & par_docnum & ""
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        aggiorna_commento_interno(RichTextBox2.Text, N_ordine)
        MsgBox("Commento aggiornato con successo")

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.Text = "" Then
            filtro_owner = ""

        Else
            filtro_owner = " And t40.owner LIKE '%" & TextBox7.Text & "%' "

        End If
        lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)
    End Sub

    Private Sub TextBox6_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If TextBox6.Text = "" Then
            filtro_matr_CDS = ""

        Else
            filtro_matr_CDS = " And T40.[u_matrcds] LIKE '%" & TextBox6.Text & "%'"
        End If

        lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub CheckedListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox2.SelectedIndexChanged
        ' Definisci la variabile filtro_reporto come vuota

        Dim listaFiltri As New List(Of String) ' Lista per memorizzare i singoli filtri

        ' Cicla attraverso tutti gli elementi della CheckedListBox
        For i As Integer = 0 To CheckedListBox2.Items.Count - 1
            ' Verifica se l'elemento è selezionato
            If CheckedListBox2.GetItemChecked(i) Then
                ' Aggiungi il filtro per l'elemento selezionato
                listaFiltri.Add("t40.famiglia = '" & CheckedListBox2.Items(i).ToString() & "'")
            End If
        Next

        ' Se ci sono elementi selezionati, crea il filtro con OR
        If listaFiltri.Count > 0 Then
            filtro_famiglia = " AND (" & String.Join(" OR ", listaFiltri) & ")"
        End If

        ' Mostra o usa il filtro generato
        ' MessageBox.Show(filtro_reparto)

        ' Puoi anche assegnare filtro_reporto a una variabile di livello superiore se necessario
        lista_ordini_cliente_codice(DataGridView2, DateTimePicker4, DateTimePicker1)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Process.Start(Homepage.percorso_server & "00-Tirelli 4.0\File\Schede tecniche\scheda tecnica O&U.xlsx")
    End Sub
End Class