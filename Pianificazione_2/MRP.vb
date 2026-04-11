Imports System.ComponentModel
Imports System.Data
Imports Microsoft
Imports BrightIdeasSoftware
Imports Newtonsoft.Json.Linq
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Imports System.IO
Imports System.Net.Mail
Imports System.Collections

Imports System.Data.OleDb


Imports System.Windows.Forms
Imports TenTec.Windows.iGridLib

Imports ADGV
Imports System.Windows.Documents
Public Class MRP
    Dim dataTable As New DataTable()
    ' Private operazione As String
    Public id_selezionato As Integer = 0
    Public utente_selezionato As Integer = 0
    Public stato_selezionato As String = ""
    Private id_mrp As Integer
    '  Private regola As String
    Private N_record_max As Integer

    ' Public dataTableBindingSource As New BindingSource()
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub



    Private Sub MRP_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label4.Text = Homepage.ID_SALVATO

        riempi_datagridview_log(DataGridView1)
    End Sub



    Sub conta_record_max(par_connection_string As String, par_id_salvato As Integer, par_datagridview As DataGridView)



        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_connection_string
        Cnn_Tree.Open()

        Dim Cmd_Tree As New SqlCommand
        Cmd_Tree.Connection = Cnn_Tree

        ' Prima query per contare i record totali
        Cmd_Tree.CommandText = "
    SELECT COUNT(*) 
    FROM (
        SELECT 
            T0.ItemCode
        FROM OITW T0
        INNER JOIN OITM T1 ON T0.ItemCode = T1.ItemCode
        INNER JOIN OWHs T2 ON T2.WhsCode = T0.WhsCode
        GROUP BY 
            T0.ItemCode,
            T1.MINLEVEL,
            T1.PRCRMNTMTD,
            T1.DFLTWH,
            T1.U_GESTIONE_MAGAZZINO,
            T1.CreateDate,
            T1.[ItmsGrpCod]
        HAVING 
            SUM(T0.OnHand - T0.IsCommited + T0.OnOrder) < COALESCE(T1.MINLEVEL, 0)
            AND SUBSTRING(T0.ItemCode, 1, 1) IN ('0', 'C', 'D', 'M')
            AND (COALESCE(T1.DFLTWH, '') NOT IN ('03', 'B03'))
            AND NOT (
                CONVERT(DATETIME, T1.CreateDate, 120) >= CONVERT(DATETIME, '2025-01-09', 120)
                AND T1.[ItmsGrpCod] IN (183, 100)
            )
    ) AS SubQuery"

        N_record_max = Convert.ToInt32(Cmd_Tree.ExecuteScalar())

        ' Ora eseguiamo la query principale
        Cmd_Tree.CommandText = "
    SELECT
        0 AS LIV,
        T0.ItemCode,
        SUBSTRING(T0.ItemCode, 1, 1) AS 'Prima',
        SUM(CASE WHEN T0.WhsCode NOT IN ('WIP', 'BWIP') THEN T0.OnHand ELSE 0 END) AS MAG,
        SUM(T0.IsCommited) AS CONF,
        SUM(T0.OnOrder) AS ORD,
        COALESCE(T1.MINLEVEL, 0) AS 'Min',
        (SUM(T0.OnHand - T0.IsCommited + T0.OnOrder) - T1.MINLEVEL) AS DISP,
        T1.PRCRMNTMTD,
        T1.U_GESTIONE_MAGAZZINO,
        CONVERT(DATE, GETDATE()) AS 'OGGI'
    FROM
        OITW T0
    INNER JOIN OITM T1 ON T0.ItemCode = T1.ItemCode
    INNER JOIN OWHs T2 ON T2.WhsCode = T0.WhsCode
    GROUP BY
        T0.ItemCode,
        T1.MINLEVEL,
        T1.PRCRMNTMTD,
        T1.DFLTWH,
        T1.U_GESTIONE_MAGAZZINO,
        T1.CreateDate,
        T1.[ItmsGrpCod]
    HAVING
        SUM(T0.OnHand - T0.IsCommited + T0.OnOrder) < COALESCE(T1.MINLEVEL, 0)
        AND SUBSTRING(T0.ItemCode, 1, 1) IN ('0', 'C', 'D', 'M')
        AND (COALESCE(T1.DFLTWH, '') NOT IN ('03', 'B03'))
        AND NOT (
            CONVERT(DATETIME, T1.CreateDate, 120) >= CONVERT(DATETIME, '2025-01-09', 120)
            AND T1.[ItmsGrpCod] IN (183, 100)
        )
    ORDER BY T0.ItemCode
"

        Dim Reader_mrp As SqlDataReader = Cmd_Tree.ExecuteReader()
        Dim n_riga As Integer = 1

        Do While Reader_mrp.Read()
            If Reader_mrp("MIN") > 0 Then
                trova_stock(par_connection_string, par_id_salvato, Reader_mrp("ITEMCODE"), n_riga)
            Else
                trova_impegni(par_connection_string, par_id_salvato, Reader_mrp("ITEMCODE"), n_riga)
            End If

            BackgroundWorker1.ReportProgress(CInt((n_riga / N_record_max) * 100))
            n_riga += 1
        Loop

        Reader_mrp.Close()
        Cnn_Tree.Close()
    End Sub

    Sub calcola_MRP(par_connection_string As String, par_id_salvato As Integer)
        ' conta_record_max(par_connection_string, par_id_salvato, par_datagridview)

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_connection_string
        Cnn_Tree.Open()

        Dim Cmd_Tree As New SqlCommand
        Cmd_Tree.Connection = Cnn_Tree

        ' Prima query per contare i record totali e salvarli in n_record_max
        Cmd_Tree.CommandText = "
    SELECT COUNT(*) 
    FROM (
        SELECT 
            T0.ItemCode
        FROM OITW T0
        INNER JOIN OITM T1 ON T0.ItemCode = T1.ItemCode
        INNER JOIN OWHs T2 ON T2.WhsCode = T0.WhsCode
        GROUP BY 
            T0.ItemCode,
            T1.MINLEVEL,
            T1.PRCRMNTMTD,
            T1.DFLTWH,
            T1.U_GESTIONE_MAGAZZINO,
            T1.CreateDate,
            T1.[ItmsGrpCod]
      HAVING
    SUM(T0.OnHand - T0.IsCommited + T0.OnOrder) < COALESCE(T1.MINLEVEL, 0)
    AND SUBSTRING(T0.ItemCode, 1, 1) IN ('0', 'C', 'D', 'M')
    AND (COALESCE(T1.DFLTWH, '') NOT IN ('03', 'B03'))
    
    ) AS SubQuery"

        Dim n_record_max As Integer = Convert.ToInt32(Cmd_Tree.ExecuteScalar())

        ' Ora eseguiamo la query principale
        Cmd_Tree.CommandText = "
    SELECT
        0 AS LIV,
        T0.ItemCode,
        SUBSTRING(T0.ItemCode, 1, 1) AS 'Prima',
        SUM(CASE WHEN T0.WhsCode NOT IN ('WIP', 'BWIP') THEN T0.OnHand ELSE 0 END) AS MAG,
        SUM(T0.IsCommited) AS CONF,
        SUM(T0.OnOrder) AS ORD,
        COALESCE(T1.MINLEVEL, 0) AS 'Min',
        (SUM(T0.OnHand - T0.IsCommited + T0.OnOrder) - T1.MINLEVEL) AS DISP,
        T1.PRCRMNTMTD,
        T1.U_GESTIONE_MAGAZZINO,
        CONVERT(DATE, GETDATE()) AS 'OGGI'
    FROM
        OITW T0
    INNER JOIN OITM T1 ON T0.ItemCode = T1.ItemCode
    INNER JOIN OWHs T2 ON T2.WhsCode = T0.WhsCode
    GROUP BY
        T0.ItemCode,
        T1.MINLEVEL,
        T1.PRCRMNTMTD,
        T1.DFLTWH,
        T1.U_GESTIONE_MAGAZZINO,
        T1.CreateDate,
        T1.[ItmsGrpCod]
  HAVING
    SUM(T0.OnHand - T0.IsCommited + T0.OnOrder) < COALESCE(T1.MINLEVEL, 0)
    AND SUBSTRING(T0.ItemCode, 1, 1) IN ('0', 'C', 'D', 'M')
    AND (COALESCE(T1.DFLTWH, '') NOT IN ('03', 'B03'))
   
    ORDER BY T0.ItemCode
"

        Dim Reader_mrp As SqlDataReader = Cmd_Tree.ExecuteReader()
        Dim n_riga As Integer = 1

        Do While Reader_mrp.Read()

            If Reader_mrp("ITEMCODE") = "m05669" Then
                MsgBox("DRA")
            End If

            If Reader_mrp("MIN") > 0 Then
                trova_stock(par_connection_string, par_id_salvato, Reader_mrp("ITEMCODE"), n_riga)
            Else
                trova_impegni(par_connection_string, par_id_salvato, Reader_mrp("ITEMCODE"), n_riga)
            End If

            BackgroundWorker1.ReportProgress(CInt((n_riga / n_record_max) * 100))
            n_riga += 1
        Loop

        Reader_mrp.Close()
        Cnn_Tree.Close()
    End Sub

    Sub crea_tabella_MRP(par_connection_string As String, par_id_salvato As Integer)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = par_connection_string


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand



        Cmd_SAP.Connection = Cnn
        'Cmd_SAP.CommandText = "IF OBJECT_ID('[tirelli_40].dbo.MRP" & par_id_salvato & "') IS NOT NULL
        'BEGIN
        'DROP TABLE [tirelli_40].dbo.MRP" & par_id_salvato & ";

        'END


        'CREATE TABLE [tirelli_40].dbo.MRP" & par_id_salvato & " (
        'contatore INT IDENTITY(1,1),
        '    id varchar(255),
        '[Percorso] [varchar](255) NULL,
        'Motivo varchar(255),
        '   codice VARCHAR(255),
        '	quantity DECIMAL,
        'minimo DECIMAL,
        '	commessa VARCHAR(255),
        '	cliente VARCHAR(255),
        '	consegna date,
        '	ord_per_commessa DECIMAL,
        'ultimo_fornitore varchar(255),
        'mag decimal,
        'conf decimal,
        'ord decimal,
        'disp decimal,
        'inter decimal,
        'N_Ticket varchar(255),
        'FASE varchar(255),
        'Gestione_mag varchar(255),
        'Causale varchar(255),
        'Tipo_montaggio varchar(255)

        ')"

        'Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = "DELETE [tirelli_40].dbo.MRP" & par_id_salvato & ""

        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()


    End Sub

    Sub apri_lancio_mrp(PAR_UTENTE)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO [Tirelli_40].[dbo].[MRP_LANCIO]
           ([Data_INIZIO]
           ,[UTENTE]

           ,[STATO])
     VALUES
           (GETDATE()
           ," & PAR_UTENTE & "
           ,'CALCOLO')"

        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()


    End Sub

    Sub chiudi_lancio_mrp(PAR_ID As Integer)
        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand



        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "update [Tirelli_40].[dbo].[MRP_LANCIO] set [Data_fine] =GETDATE(), STATO ='OK'
WHERE ID_LANCIO ='" & PAR_ID & "'"

        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()


    End Sub

    Public Function INDIVIDUA_ID_MRP()

        Dim id As Integer = 0

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT MAX(ID_LANCIO) AS 'ID' FROM [Tirelli_40].[dbo].[MRP_LANCIO]
    
"
        Reader_mrp = Cmd_Tree.ExecuteReader()

        If Reader_mrp.Read() Then
            id = Reader_mrp("ID")
        End If



        Reader_mrp.Close()
        Cnn_Tree.Close()
        Return id
    End Function

    Sub insert_into_mrp_tempt(par_connection_string As String, par_id_salvato As Integer, PAR_id As String,
                              par_motivo As String,
   PAR_codice As String,
    PAR_quantity As String,
    PAR_Commessa As String,
    PAR_Cliente As String,
    PAR_Consegna As String,
    PAR_Ord_per_commessa As String,
    par_ultimo_fornitore As String,
    PAR_mag As String,
        PAR_conf As String,
        PAR_ord As String,
        PAR_disp As String,
     PAR_inter As String,
     par_n_ticket As String,
                              par_minimo As String,
                              PAR_FASE As String,
                              par_gestione_mag As String,
                              par_percorso As String,
                              par_causale As String,
                              par_tipo_montaggio As String)

        PAR_Ord_per_commessa = Replace(PAR_Ord_per_commessa, ",", ".")
        PAR_Commessa = Replace(PAR_Commessa, "'", " ")
        PAR_Cliente = Replace(PAR_Cliente, "'", " ")
        par_ultimo_fornitore = Replace(par_ultimo_fornitore, "'", " ")
        PAR_quantity = Replace(PAR_quantity, ",", ".")
        PAR_mag = Replace(PAR_mag, ",", ".")
        PAR_ord = Replace(PAR_ord, ",", ".")
        PAR_conf = Replace(PAR_conf, ",", ".")
        PAR_disp = Replace(PAR_disp, ",", ".")
        PAR_inter = Replace(PAR_inter, ",", ".")
        par_minimo = Replace(par_minimo, ",", ".")
        par_percorso = Replace(par_percorso, ",", ".")
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = par_connection_string



        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        ' manca da assegnare il valore par_docentry_rt a baseentry perchè devo trovarlo nel values

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO [TIRELLI_40].DBO.MRP" & par_id_salvato & "
( id ,
motivo,
   codice ,
	quantity ,
	commessa ,
	cliente ,
	consegna ,
	ord_per_commessa,
ultimo_fornitore,
mag,
ord,
conf,
disp,
inter,
N_ticket,
minimo,
FASE,
gestione_mag,
percorso,
causale,
tipo_montaggio

)
	VALUES
	('" & PAR_id & "',
    '" & par_motivo & "',
   '" & PAR_codice & "',
    '" & PAR_quantity & "',
    '" & PAR_Commessa & "',
    '" & PAR_Cliente & "',
  CONVERT(datetime,'" & PAR_Consegna & "', 103) ,
   '" & PAR_Ord_per_commessa & "',
'" & par_ultimo_fornitore & "',
'" & PAR_mag & "',
'" & PAR_ord & "',
'" & PAR_conf & "',
'" & PAR_disp & "',
'" & PAR_inter & "',
'" & par_n_ticket & "',
'" & par_minimo & "',
'" & PAR_FASE & "',
'" & par_gestione_mag & "',
'" & par_percorso & "',
'" & par_causale & "',
'" & par_tipo_montaggio & "'

)"

        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()


    End Sub

    Sub trova_stock(par_connection_string As String, par_id_salvato As Integer, par_codice As String, par_n_riga As Integer)

        Dim par_commessa As String = "STOCK"
        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_connection_string
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "


select t10.itemcode, t10.min, t10.disp,t10.PrcrmntMtd
, case when t10.Q_min>=-t10.disp+t10.min then t10.Q_min else -t10.disp+t10.min end as 'Q'
,t10.mag
,t10.conf,
t10.ord
,t10.disp
,T10.min
,t10.gestione_mag
,t10.tipo_montaggio
from
(
select t0.itemcode,

 coalesce(t0.minlevel,0)  as min
,coalesce(t0.minordrqty,0) as 'Q_min'
,sum(coalesce(case when t2.whscode<>'WIP' and t2.whscode<>'BWIP' then t2.onhand else 0 end,0)) as 'Mag'
,sum(coalesce(t2.iscommited,0)) as 'Conf'
,sum(coalesce(t2.onorder,0)) as 'Ord'
, sum(coalesce(t2.onhand,0)- coalesce(t2.iscommited,0)+coalesce(t2.onorder,0)) as 'Disp'

,t0.PrcrmntMtd
,coalesce(T0.U_GESTIONE_MAGAZZINO,'') as 'Gestione_mag'
,coalesce(t0.u_tipo_montaggio,'') as 'Tipo_montaggio'
from oitm t0 
left join oitb t1 on T0.[ItmsGrpCod]=T1.[ItmsGrpCod]
left join oitw t2 on t2.itemcode=t0.itemcode
INNER JOIN OWHs T3 ON T3.whscode = T2.WhsCode

where t0.itemcode='" & par_codice & "'
group by t0.itemcode,t0.itemname,t1.ItmsGrpNam,t0.u_disegno,t0.U_Codice_BRB,t0.u_tipo_montaggio

,t0.minlevel
,t0.minordrqty
,t0.PrcrmntMtd
, T0.U_GESTIONE_MAGAZZINO


)
as t10
order by t10.itemcode

"
        Reader_mrp = Cmd_Tree.ExecuteReader()

        If Reader_mrp.Read() Then


            insert_into_mrp_tempt(par_connection_string, par_id_salvato, par_n_riga,
                      "STOCK",
                      Reader_mrp("itemcode"),
                      Reader_mrp("q"),
                      par_commessa,
                      par_commessa,
                      trova_prima_data_stock(par_connection_string, Reader_mrp("itemcode")),
                       trova_ordinati_per_commessa(par_connection_string, par_codice, "STOCK"),
                      trova_ultimo_fornitore(par_connection_string, Reader_mrp("itemcode")),
                      Reader_mrp("mag"),
                      Reader_mrp("conf"),
                      Reader_mrp("ord"),
                      Reader_mrp("disp"),
                      0,
trova_ticket(par_connection_string, Reader_mrp("itemcode")),
Reader_mrp("min"),
"",
Reader_mrp("gestione_mag"), Reader_mrp("itemcode"), "STOCK", Reader_mrp("Tipo_montaggio"))




            If Reader_mrp("PrcrmntMtd") = "M" Then
                trova_figli(par_connection_string, par_id_salvato, Reader_mrp("itemcode"), par_n_riga, "STOCK", Reader_mrp("q"), par_commessa, par_commessa, trova_prima_data_stock(par_connection_string, Reader_mrp("itemcode")), "", Reader_mrp("itemcode"), "STOCK")
                ' trova_impegni(par_centro_di_costo, par_operazione, Reader_mrp("itemcode"), par_n_riga)

            End If
        End If
        Reader_mrp.Close()
        Cnn_Tree.Close()
    End Sub

    Sub trova_figli(par_connection_string As String, par_id_salvato As Integer, par_codice As String, par_n_riga As String, PAR_MOTIVO As String, par_q As String, par_commessa As String, par_cliente As String, PAR_Consegna As Date, PAR_FASE As String, par_percorso As String, par_causale As String)
        'If par_codice = "D86403" Then
        '    MsgBox("dra")
        'End If
        par_q = Replace(par_q, ",", ".")
        Dim n_riga_figlio As Integer = 1

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_connection_string
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "

select T0.CODE, " & par_q & " *T0.Quantity/t2.qauntity AS 'Quantity', t1.PrcrmntMtd, coalesce(t1.u_gestione_magazzino,'') as 'Gestione_mag'
, coalesce(t1.minlevel,0)  as 'Min'
, coalesce(t3.u_tipo_montaggio,'') as 'Tipo_montaggio'
from itt1 t0 inner join oitm t1 on t0.code=t1.itemcode
INNER JOIN oitt t2 on t2.code=t0.father
inner join oitm t3 on t3.itemcode=t0.code
where t0.father='" & par_codice & "' and t0.type='4' and substring(t0.code,1,1)<>'L'

order by T0.CODE
"
        Reader_mrp = Cmd_Tree.ExecuteReader()



        Do While Reader_mrp.Read()


            insert_into_mrp_tempt(par_connection_string, par_id_salvato, par_n_riga & "." & n_riga_figlio, PAR_MOTIVO,
   Reader_mrp("code"),
    Reader_mrp("Quantity"),
    par_commessa,
    par_cliente,
    PAR_Consegna,
    trova_ordinati_per_commessa(par_connection_string, Reader_mrp("code"), par_commessa),
trova_ultimo_fornitore(par_connection_string, Reader_mrp("code")),
trova_a_mag(par_connection_string, Reader_mrp("code")),
trova_conf(par_connection_string, Reader_mrp("code")),
trova_ORD(par_connection_string, Reader_mrp("code")),
trova_disp(par_connection_string, Reader_mrp("code")),
0,
trova_ticket(par_connection_string, Reader_mrp("code")),
Reader_mrp("min"),
PAR_FASE,
Reader_mrp("gestione_mag"),
par_percorso & " - " & Reader_mrp("code"), par_causale, Reader_mrp("Tipo_montaggio"))

            If Reader_mrp("PrcrmntMtd") = "M" Then
                trova_figli(par_connection_string, par_id_salvato, Reader_mrp("code"),
                            par_n_riga & "." & n_riga_figlio,
                            PAR_MOTIVO, Reader_mrp("Quantity"),
                            par_commessa, par_cliente, PAR_Consegna, PAR_FASE, par_percorso & " - " & Reader_mrp("code"), par_causale)
            End If
            n_riga_figlio += 1
        Loop
        Reader_mrp.Close()
        Cnn_Tree.Close()

    End Sub

    Public Function trova_prima_data_stock(par_conenction_string As String, par_codice As String)


        Dim prima_data As Date = Today

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_conenction_string
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "
declare @giorni as integer
set @giorni =5
select t30.itemcode, t30.min_consegna, 
    RIGHT('00' + CAST(DAY(t30.min_consegna) AS VARCHAR(2)), 2) + '/' +
    RIGHT('00' + CAST(MONTH(t30.min_consegna) AS VARCHAR(2)), 2) + '/' +
    CAST(YEAR(t30.min_consegna) AS VARCHAR(4)) AS Min_consegna_format
from
(
select t20.itemcode,  dbo.SubtractWorkingDays(t20.min_consegna, @giorni)  as 'Min_consegna'
from
(
select 
    t10.itemcode,
    MIN(t10.startdate) AS Min_consegna,

    RIGHT('00' + CAST(DAY(MIN(t10.startdate)) AS VARCHAR(2)), 2) + '/' +
    RIGHT('00' + CAST(MONTH(MIN(t10.startdate)) AS VARCHAR(2)), 2) + '/' +
    CAST(YEAR(MIN(t10.startdate)) AS VARCHAR(4)) AS Min_consegna_format
from
(
select t2.itemcode, t0.startdate
from owor t0 inner join owhs t1 on t0.warehouse=t1.whscode
inner join wor1 t2 on t2.docentry=t0.docentry
where (t0.status='R' or t0.status='P')  and t2.itemcode='" & par_codice & "' and t2.u_prg_wip_qtadatrasf>0
union all
select t1.itemcode, t0.docduedate
from ordr t0 
inner join rdr1 t1 on t1.docentry=t0.docentry
inner join owhs t2 on t2.whscode=t1.whscode
where t1.OpenQty>0 and t1.U_Datrasferire>0 and t1.itemcode='" & par_codice & "'
)
as t10
group by t10.itemcode
)
as t20
)
as t30
"
        Reader_mrp = Cmd_Tree.ExecuteReader()

        If Reader_mrp.Read() Then


            prima_data = Reader_mrp("Min_consegna_format")

        End If
        Reader_mrp.Close()
        Cnn_Tree.Close()
        Return prima_data
    End Function





    Sub trova_impegni(par_connection_string As String, par_id_salvato As Integer, par_codice As String, par_n_riga As Integer)

   

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_connection_string
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "declare @giorni as integer
set @giorni =5


select t20.Motivo, t20.itemcode,t20.U_PRG_AZS_Commessa, t20.causale,

t20.u_utilizz, min(t20.startdate) as 'Startdate', min(t20.u_fase) as 'U_fase', sum(t20.q) as 'Q', t20.PrcrmntMtd, t20.Gestione_mag,t20.min, min(t20.tipo_montaggio) as 'Tipo_montaggio'
from
(
select t10.Motivo, t10.itemcode,t10.U_PRG_AZS_Commessa, t10.causale,

t10.u_utilizz,


 dbo.SubtractWorkingDays(t10.startDate, @giorni) as startdate, t10.u_fase
,

t10.Q, t11.PrcrmntMtd, coalesce(t11.u_gestione_magazzino,'') as 'Gestione_mag' 
, coalesce(t11.minlevel,0)  as 'Min'
,coalesce(t11.u_tipo_montaggio,'') as 'Tipo_montaggio'
from
(
select 'ODP' as 'Motivo', t0.itemcode,t1.U_PRG_AZS_Commessa, coalesce(t7.u_causcons,'') as 'Causale',


coalesce(t1.u_utilizz,'') as 'U_utilizz'

,t1.startDate, sum(coalesce(t0.U_PRG_WIP_QtaDaTrasf,0)) as 'Q', coalesce(t3.NAME,'') AS 'U_FASE'
from wor1 t0
inner join owor t1 on t0.docentry=t1.docentry
inner join owhs t2 on t2.whscode=t0.wareHouse
LEFT JOIN [dbo].[@FASE]  T3 ON T1.[U_Fase] = T3.[Code]
left join rdr1 t4 on t4.itemcode=t1.U_PRG_AZS_Commessa and T4.[OpenQty]>0
left join ordr t5 on t5.docentry=t4.docentry
left join ocrd t6 on t6.cardcode=t5.u_codicebp
left join ordr t7 on (cast(substring(t1.U_PRG_AZS_Commessa,2,5) as varchar) =cast(t7.docnum as varchar)) or (cast(t1.U_PRG_AZS_Commessa as varchar) = CAST(T7.U_MATRCDS AS VARCHAR))

where t0.itemcode='" & par_codice & "' and (t1.status='P' or t1.status='R') and t0.U_PRG_WIP_QtaDaTrasf>0 
group by t0.itemcode,t1.U_PRG_AZS_Commessa,t1.startDate,t1.u_utilizz, t3.NAME,t6.cardname,t5.cardname,coalesce(t7.u_causcons,'')

union all

select 'OC',t0.itemcode,concat('_',t1.docnum), coalesce(t1.u_causcons,'') as 'u_causcons'

,

coalesce(t3.cardname,coalesce(t1.cardname,'')) , t1.docduedate, sum(coalesce(t0.U_datrasferire,0)) as 'Q',''
from rdr1 t0
inner join ordr t1 on t0.docentry=t1.docentry
inner join owhs t2 on t2.whscode=t0.whscode
left join ocrd t3 on t3.cardcode=t1.U_CodiceBP
where t0.itemcode='" & par_codice & "' and t0.OpenQty>0 and t0.U_Datrasferire>0 
group by t0.itemcode,t1.docnum,t1.docduedate,t3.cardname, t1.cardname,t3.cardname, coalesce(t1.u_causcons,'')
)
as t10 inner join oitm t11 on t10.itemcode=t11.itemcode
)
as t20
group by t20.Motivo, t20.itemcode,t20.U_PRG_AZS_Commessa, t20.causale,

t20.u_utilizz,t20.PrcrmntMtd, t20.Gestione_mag,t20.min
order by t20.itemcode
"
        Reader_mrp = Cmd_Tree.ExecuteReader()

        Do While Reader_mrp.Read()
            Dim valore1 As Decimal = Decimal.Parse(trova_ordinati_per_commessa(par_connection_string, par_codice, Reader_mrp("U_PRG_AZS_Commessa")))
            Dim valore2 As Decimal = Decimal.Parse(Reader_mrp("q"))
            If valore1 < valore2 Then


                insert_into_mrp_tempt(par_connection_string, par_id_salvato, par_n_riga,
                                Reader_mrp("Motivo"),
 Reader_mrp("itemcode"),
  Reader_mrp("q"),
  Reader_mrp("U_PRG_AZS_Commessa"),
    Reader_mrp("U_utilizz"),
  Reader_mrp("startDate"),
  trova_ordinati_per_commessa(par_connection_string, par_codice, Reader_mrp("U_PRG_AZS_Commessa")),
  trova_ultimo_fornitore(par_connection_string, par_codice),
  trova_a_mag(par_connection_string, par_codice),
trova_conf(par_connection_string, par_codice),
trova_ORD(par_connection_string, par_codice),
trova_disp(par_connection_string, par_codice),
0,
trova_ticket(par_connection_string, par_codice),
Reader_mrp("min"),
Reader_mrp("u_fase"),
Reader_mrp("gestione_mag"),
Reader_mrp("itemcode"),
Reader_mrp("causale"),
Reader_mrp("Tipo_montaggio"))

                If Reader_mrp("PrcrmntMtd") = "M" Then
                    trova_figli(par_connection_string, par_id_salvato, Reader_mrp("itemcode"), par_n_riga, Reader_mrp("Motivo"), Reader_mrp("q"), Reader_mrp("U_PRG_AZS_Commessa"), Reader_mrp("U_utilizz"), Reader_mrp("startDate"), Reader_mrp("u_fase"), Reader_mrp("itemcode"), Reader_mrp("causale"))
                End If
                par_n_riga += 1
            End If
        Loop


        Reader_mrp.Close()
        Cnn_Tree.Close()
    End Sub

    Public Function trova_ordinati_per_commessa(par_connection_string As String, par_codice As String, par_commessa As String)
        Dim ordinato_per_commessa As Decimal = 0

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_connection_string
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "select t10.itemcode, sum(t10.q) as 'Q'
from
(
select t0.itemcode, sum(t0.plannedqty) as 'q'

from owor t0
where t0.itemcode='" & par_codice & "' and t0.U_PRG_AZS_Commessa='" & par_commessa & "' and (t0.status='P' or t0.status='R')
group by t0.itemcode
union all
select t0.itemcode, sum(t0.opencreqty) as 'q'

from por1 t0
where t0.itemcode='" & par_codice & "' and t0.U_PRG_AZS_Commessa='" & par_commessa & "' and t0.opencreqty>0
group by t0.itemcode
)
as t10
group by t10.itemcode
"
        Reader_mrp = Cmd_Tree.ExecuteReader()

        If Reader_mrp.Read() Then

            ordinato_per_commessa = Reader_mrp("Q")
        Else
            ordinato_per_commessa = 0


        End If


        Reader_mrp.Close()
        Cnn_Tree.Close()
        ' Return Replace(ordinato_per_commessa, ",", ".")
        Return ordinato_per_commessa
    End Function

    Public Function trova_ticket(par_connection_string As String, par_codice As String)
        Dim n_ticket As String = ""

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_connection_string
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = " SELECT t1.codice_SAP, t0.Id_Ticket
from [TIRELLI_40].[DBO].coll_tickets t0 inner join [TIRELLI_40].[DBO].COLL_RIFERIMENTI t1 on t0.id_ticket=t1.rif_ticket

where (T0.MOTIVAZIONE=12 or T0.MOTIVAZIONE=7) and t0.aperto=1 and t1.codice_sap='" & par_codice & "' "
        Reader_mrp = Cmd_Tree.ExecuteReader()

        If Reader_mrp.Read() Then

            n_ticket = Reader_mrp("Id_Ticket")
        Else
            n_ticket = ""


        End If


        Reader_mrp.Close()
        Cnn_Tree.Close()
        Return n_ticket
    End Function

    Public Function trova_a_mag(par_connection_string As String, par_codice As String)
        Dim a_mag As String = 0

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_connection_string
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = " select sum(coalesce(case when t0.whscode<>'WIP' and t0.whscode<>'BWIP' then t0.onhand else 0 end,0)) as 'Mag'
from oitw t0 inner join owhs t1 on t0.whscode=t1.whscode
where t0.itemcode='" & par_codice & "' 


"
        Reader_mrp = Cmd_Tree.ExecuteReader()

        If Reader_mrp.Read() Then

            a_mag = Reader_mrp("mag")
        Else
            a_mag = 0


        End If


        Reader_mrp.Close()
        Cnn_Tree.Close()
        Return Replace(a_mag, ",", ".")
    End Function

    Public Function trova_ORD(par_connection_string As String, par_codice As String)
        Dim ord As String = 0

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_connection_string
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = " select sum(coalesce(t0.onorder,0)) as 'ord'
from oitw t0 inner join owhs t1 on t0.whscode=t1.whscode
where t0.itemcode='" & par_codice & "' 


"
        Reader_mrp = Cmd_Tree.ExecuteReader()

        If Reader_mrp.Read() Then

            ord = Reader_mrp("ord")
        Else
            ord = 0


        End If


        Reader_mrp.Close()
        Cnn_Tree.Close()
        Return Replace(ord, ",", ".")
    End Function

    Public Function trova_conf(par_connection_string As String, par_codice As String)
        Dim conf As String = 0

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_connection_string
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = " select sum(coalesce(t0.iscommited,0)) as 'conf'
from oitw t0 inner join owhs t1 on t0.whscode=t1.whscode
where t0.itemcode='" & par_codice & "' 


"
        Reader_mrp = Cmd_Tree.ExecuteReader()

        If Reader_mrp.Read() Then

            conf = Reader_mrp("conf")
        Else
            conf = 0


        End If


        Reader_mrp.Close()
        Cnn_Tree.Close()
        Return Replace(conf, ",", ".")
    End Function

    Public Function trova_disp(par_connection_string As String, par_codice As String)
        Dim disp As String = 0

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_connection_string
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = " select sum(coalesce(t0.onhand,0)+coalesce(t0.onorder,0)-coalesce(t0.iscommited,0)) as 'disp'
from oitw t0 inner join owhs t1 on t0.whscode=t1.whscode
where t0.itemcode='" & par_codice & "' 


"
        Reader_mrp = Cmd_Tree.ExecuteReader()

        If Reader_mrp.Read() Then

            disp = Reader_mrp("disp")
        Else
            disp = 0


        End If


        Reader_mrp.Close()
        Cnn_Tree.Close()
        Return Replace(disp, ",", ".")
    End Function

    '    Public Function trova_inter(par_connection_string As String, par_codice As String)


    '        Dim MAGAZZINO As String


    '        If par_operazione = "=" Then
    '            MAGAZZINO = "WIP"
    '        Else
    '            MAGAZZINO = "BWIP"
    '        End If
    '        Dim inter As String = 0

    '        Dim Cnn_Tree As New SqlConnection
    '        Cnn_Tree.ConnectionString = par_connection_string
    '        Cnn_Tree.Open()
    '        Dim Cmd_Tree As New SqlCommand
    '        Dim Reader_mrp As SqlDataReader
    '        Cmd_Tree.Connection = Cnn_Tree
    '        Cmd_Tree.CommandText = " select sum(coalesce(case when t0.whscode<>'WIP' and t0.whscode<>'BWIP' then t0.onhand else 0 end,0)) as 'inter'
    'from oitw t0 inner join owhs t1 on t0.whscode=t1.whscode
    'where t0.itemcode='" & par_codice & "' 
    'and coalesce(t1.location,0) " & par_operazione & "13
    'AND T0.WHSCODE<>'" & MAGAZZINO & "'


    '"
    '        Reader_mrp = Cmd_Tree.ExecuteReader()

    '        If Reader_mrp.Read() Then

    '            inter = Reader_mrp("inter")
    '        Else
    '            inter = 0

    '        End If


    '        Reader_mrp.Close()
    '        Cnn_Tree.Close()
    '        Return Replace(inter, ",", ".")
    '    End Function

    Public Function trova_ultimo_fornitore(par_connection_string As String, par_codice As String)
        Dim ultimo_fornitore As String = ""

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = par_connection_string
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "
select t11.docnum, t11.cardname
from
(
select max(t0.docentry) as 'MAX'
from por1 t0 where t0.itemcode='" & par_codice & "'
)
as t10 inner join opor t11 on t10.max=t11.docentry"
        Reader_mrp = Cmd_Tree.ExecuteReader()

        If Reader_mrp.Read() Then

            ultimo_fornitore = Reader_mrp("cardname")



        End If


        Reader_mrp.Close()
        Cnn_Tree.Close()
        Return ultimo_fornitore
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        apri_lancio_mrp(Homepage.ID_SALVATO)
        id_mrp = INDIVIDUA_ID_MRP()

        crea_tabella_MRP(Homepage.sap_tirelli, Homepage.ID_SALVATO)
        BackgroundWorker1.RunWorkerAsync(Homepage.sap_tirelli)

    End Sub

    Sub riempi_datagridview_log(par_datagridview As DataGridView)

        par_datagridview.Rows.Clear()
        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_mrp As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "select t11.[ID_LANCIO]
      ,t11.[Data_INIZIO]
      ,t11.[Data_FINE]
,DATEDIFF(MINUTE,  t11.[Data_FINE],GETDATE()) as 'Diff'
      ,t11.[UTENTE]
	  , concat(t12.lastname,' ', t12.firstname) as 'Utente_nome'
      ,t11.[BRAND]
      ,t11.[STATO]
from
(
select t0.utente, max(t0.id_lancio) as 'ID_lancio'
FROM [Tirelli_40].[dbo].[MRP_LANCIO] t0 left join [TIRELLI_40].[dbo].ohem t1 on t0.utente=t1.empid
where stato='OK'
group by t0.utente
)
as t10 inner join [Tirelli_40].[dbo].[MRP_LANCIO] t11 on t10.id_lancio=t11.id_lancio
left join [TIRELLI_40].[dbo].ohem t12 on t10.utente=t12.empid
ORDER BY T11.ID_LANCIO DESC
"
        Reader_mrp = Cmd_Tree.ExecuteReader()

        Do While Reader_mrp.Read()

            par_datagridview.Rows.Add(Reader_mrp("ID_LANCIO"), Reader_mrp("Data_INIZIO"), Reader_mrp("Data_FINE"), Reader_mrp("Diff"), Reader_mrp("utente"), Reader_mrp("utente_nome"), Reader_mrp("brand"), Reader_mrp("stato"))

        Loop

        Reader_mrp.Close()
        Cnn_Tree.Close()
        par_datagridview.ClearSelection()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            id_selezionato = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ID").Value
            utente_selezionato = DataGridView1.Rows(e.RowIndex).Cells(columnName:="utente_sap").Value
            Label2.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Ora_calcolo").Value

            stato_selezionato = DataGridView1.Rows(e.RowIndex).Cells(columnName:="stato").Value
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If id_selezionato = 0 Then
            MsgBox("Selezionare un MRP da caricare oppure lanciarne uno nuovo")
        ElseIf stato_selezionato <> "OK" Then
            MsgBox("è stato selezionato un MRP non calcolato pienamente. Selezionarne uno 'OK' oppure lanciarne uno nuovo")
        Else

            Process.Start(Homepage.percorso_acquisti & "\MRP" & utente_selezionato & ".xlsx")


        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        'If DataGridView1.Rows(e.RowIndex).Cells(columnName:="BRAND").Value = "BRB01" Then

        '    DataGridView1.Rows(e.RowIndex).Cells(columnName:="BRAND").Style.BackColor = Color.Yellow


        'Else
        '    DataGridView1.Rows(e.RowIndex).Cells(columnName:="BRAND").Style.BackColor = Color.LightBlue
        'End If

        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="STATO").Value = "OK" Then

            DataGridView1.Rows(e.RowIndex).Cells(columnName:="STATO").Style.BackColor = Color.Lime


        Else
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="STATO").Style.BackColor = Color.Yellow
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If id_selezionato = 0 Then
            MsgBox("Selezionare un MRP da caricare oppure lanciarne uno nuovo")
        ElseIf stato_selezionato <> "OK" Then
            MsgBox("è stato selezionato un MRP non calcolato pienamente. Selezionarne uno 'OK' oppure lanciarne uno nuovo")
        Else

            trova_dato_da_excel_pEr_importazionE(Homepage.percorso_acquisti & "\MRP" & utente_selezionato & ".xlsx", "MRP", 2)
        End If
    End Sub

    Sub trova_dato_da_excel_pEr_importazionE(par_percorso_file As String, par_nome_foglio As String, par_riga_inizio As Integer)

        Dim colonna1 As String
        Dim colonna2 As String
        Dim colonna3 As String
        Dim colonna4 As String
        Dim colonna5 As String
        Dim colonna6 As String
        Dim percorso As String




        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True


        Do While Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value <> ""


            If Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value = "D" Or Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value = "0" Then


                colonna1 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 9).value
                percorso = Homepage.percorso_disegni_generico & "PDF\"  & colonna1 & ".PDF"
                If File.Exists(percorso) Then

                    Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 10).value = "SI"

                Else

                    Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 10).value = "NO"
                End If

            End If
            par_riga_inizio = par_riga_inizio + 1
        Loop
        Beep()
        MsgBox("FINE CONTROLLO")


    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim connectionString As String = CStr(e.Argument)

        calcola_MRP(connectionString, Label4.Text)
        'calcola_MRP(DataGridView_MRP, ComboBox1.Text, operazione)


    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        ProgressBar1.Value = e.ProgressPercentage
        Label5.Text = e.ProgressPercentage & " % "
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Console.WriteLine(Homepage.percorso_acquisti & "\MRP" & Homepage.ID_SALVATO & ".xlsx")
        Process.Start(Homepage.percorso_acquisti & "\MRP" & Homepage.ID_SALVATO & ".xlsx")
        chiudi_lancio_mrp(id_mrp)
        riempi_datagridview_log(DataGridView1)
        Me.TopMost = True
        MsgBox("MRP CALCOLATO CON SUCCESSO")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub TableLayoutPanel3_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel3.Paint

    End Sub
End Class