'Imports System.IO
'Imports System.Net.Mail
'Imports System.Collections
'Imports System.Data
'Imports System.Data.SqlClient
'Imports System.Data.OleDb
'Imports Microsoft.Office.Interop
'Imports Word = Microsoft.Office.Interop.Word
'Imports System.Windows.Controls
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement
'Imports System.Windows.Shapes
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Button

'Imports System.Drawing
'Imports System.Drawing.Printing
'Imports System.Runtime.InteropServices

Imports Word = Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop


Imports BrightIdeasSoftware

Imports System.Windows.Forms
Imports TenTec.Windows.iGridLib
Imports System.ComponentModel
Imports ADGV
Imports System.Windows.Documents



Public Class Form_lotto_di_prelievo
    Public id_lotto As Integer
    Public codice_sap As String
    Public Elenco_dipendenti(1000) As String
    Private filtro_magazzino As String
    Private filtro_magazzino_oc As String
    Public Sel_Stampante As New PrintDialog

    Public altezza_scontrino_odp As Integer = 220
    Public larghezza_scontrino_odp As Integer = 185
    Public mag_destinazione As String
    Public tipo_documento As String = "ODP"
    Public numero_documento As String

    Public quantità_trasferimento As String

    Public preview_scontrino As String = "NO"



    Sub inizializzazione_lotto_di_prelievo()
        Txt_DocNum.Text = id_lotto
        Inserimento_dipendenti()
        RIEMPI_datagridview_documenti_lotto(id_lotto)
        RIEMPI_datagridview_documenti_lotto_oc(id_lotto, DataGridView3)
        info_testata_lotto_di_preliavo(id_lotto)
        RIEMPI_datagridview_CODICI_PRELIEVO(id_lotto)
        carica_magazzini(id_lotto)
    End Sub

    Sub RIEMPI_datagridview_documenti_lotto(id_lotto_prelievo As Integer)

        DataGridView4.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "declare @id_lotto_prelievo as integer
set @id_lotto_prelievo=' " & id_lotto_prelievo & "'


select t0.tipo_doc, t1.docnum, t1.itemcode, t1.ProdName, t2.U_Disegno, t1.status
, sum(case when t3.u_prg_wip_qtadatrasf > 0 and t3.wareHousE ='03' and t4.onhand>0 then 1 else 0 end) as 'Mag_03'
, sum(case when t3.u_prg_wip_qtadatrasf > 0 and t3.wareHousE ='SCA'and t4.onhand>0 then 1 else 0 end) as 'SCA'
, sum(case when t3.u_prg_wip_qtadatrasf > 0 and t3.wareHousE ='FERRETTO' and t4.onhand>0 then 1 else 0 end) as 'FERRETTO'
, sum(case when t3.u_prg_wip_qtadatrasf > 0 and t3.wareHousE ='BSCA' and t4.onhand>0 then 1 else 0 end) as 'BSCA'
, sum(case when t3.u_prg_wip_qtadatrasf > 0 then 1 else 0 end) as 'Da_trasf'
, sum(case when t3.u_prg_wip_qtadatrasf = 0 then 1 else 0 end) as 'TRASFERITO'
, case when t1.u_progressivo_commessa is null then 0 else t1.u_progressivo_commessa end as 'u_progressivo_commessa'
, case when t1.u_prg_azs_commessa is null then '' else t1.u_prg_azs_commessa end as 'u_prg_azs_commessa'

from 
[Tirelli_40].[dbo].[lotto_prelievo_riga] t0 
left join owor t1 on t0.docnum=t1.docnum and t0.tipo_DOC='ODP' 
left join oitm t2 on t2.itemcode=t1.itemcode

left join wor1 t3 on t3.docentry=t1.docentry AND (SUBSTRING(T3.ITEMCODE,1,1)='0' OR SUBSTRING(T3.ITEMCODE,1,1)='C' OR SUBSTRING(T3.ITEMCODE,1,1)='D' OR SUBSTRING(T3.ITEMCODE,1,1)='F')
inner join oitw t4 on t4.WhsCode=t3.wareHouse and t4.itemcode=t3.itemcode
where t0.id=@id_lotto_prelievo and (t1.status='P' or t1.status='R') 
group by t0.tipo_doc, t1.docnum, t1.itemcode, t1.ProdName, t2.U_Disegno, t1.status,t1.u_progressivo_commessa,t1.u_prg_azs_commessa
order by t1.u_prg_azs_commessa, t1.u_progressivo_commessa
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            DataGridView4.Rows.Add(cmd_SAP_reader_2("tipo_doc"), cmd_SAP_reader_2("docnum"), cmd_SAP_reader_2("u_prg_azs_commessa"), cmd_SAP_reader_2("u_progressivo_commessa"), cmd_SAP_reader_2("status"), cmd_SAP_reader_2("ITEMcode"), cmd_SAP_reader_2("prodname"), cmd_SAP_reader_2("u_disegno"), cmd_SAP_reader_2("Mag_03"), cmd_SAP_reader_2("SCA"), cmd_SAP_reader_2("Ferretto"), cmd_SAP_reader_2("BSCA"), cmd_SAP_reader_2("da_trasf"), cmd_SAP_reader_2("trasferito"))

        Loop



        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        DataGridView4.ClearSelection()
    End Sub

    Sub RIEMPI_datagridview_documenti_lotto_oc(id_lotto_prelievo As Integer, par_datagridview As DataGridView)

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "declare @id_lotto_prelievo as integer
set @id_lotto_prelievo='" & id_lotto_prelievo & "'

select t1.docentry,t1.docnum, t1.cardcode, t1.cardname
,coalesce(t3.cardcode,'') as 'Codice_CF', coalesce(t3.CardName,'') as 'Cliente_f'
,t1.u_causcons
, t1.CreateDate,t1.DocDueDate
, sum(case when t2.U_Datrasferire>0 then 1 else 0 end) as 'Codici_da_trasf'
,sum(case when t2.U_Datrasferire>0 and t4.onhand>0 then 1 else 0 end) as 'Codici_trasferibili'
,sum(case when t2.U_Datrasferire>0 and t5.onhand>0 then 1 else 0 end) as 'Ferretto'


from 
[Tirelli_40].[dbo].[lotto_prelievo_riga] t0 
left join ORDR t1 on t0.docnum=t1.docnum and t0.tipo_DOC='OC' 
left join rdr1 t2 on t2.docentry=t1.docentry
left join ocrd t3 on t3.cardcode=t1.U_codicebp
left join oitw t4 on t4.whscode=t2.WhsCode and t4.itemcode=t2.itemcode
left join oitw t5 on t5.whscode=t2.WhsCode and t5.itemcode=t2.itemcode and t5.WhsCode='FERRETTO'

where t0.id=@id_lotto_prelievo and t1.docstatus='O'
and t2.OpenCreQty>0 and (substring(t2.itemcode,1,1)='0' or substring(t2.itemcode,1,1)='C' or substring(t2.itemcode,1,1)='d' or substring(t2.itemcode,1,1)='F')
  group by t1.docentry,t1.docnum,t1.cardcode, t1.cardname,t3.cardcode , t3.CardName,t1.CreateDate,t1.DocDueDate,t1.u_causcons
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(False, cmd_SAP_reader_2("docentry"), cmd_SAP_reader_2("docnum"), cmd_SAP_reader_2("cardcode"), cmd_SAP_reader_2("cardname"), cmd_SAP_reader_2("codice_CF"), cmd_SAP_reader_2("cliente_f"), cmd_SAP_reader_2("u_causcons"), cmd_SAP_reader_2("createdate"), cmd_SAP_reader_2("docduedate"), cmd_SAP_reader_2("Codici_da_trasf"), cmd_SAP_reader_2("Codici_trasferibili"), cmd_SAP_reader_2("Ferretto"))

        Loop



        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        par_datagridview.ClearSelection()
    End Sub

    Sub RIEMPI_datagridview_CODICI_PRELIEVO(id_lotto_prelievo As Integer)
        Dim contatore As Integer = 0
        Dim contatore_trasferibili As Integer = 0
        Dim itemcode As String
        Dim docnum As Integer
        Dim tipo_doc As String

        DataGridView1.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        If Label1.Text = "MU" Then
            CMD_SAP_2.CommandText = "select T0.Tipo_doc,T1.DOCNUM,T2.ItemCode,t2.linenum,T3.ITEMNAME,
case when t3.u_disegno is null then '' else t3.u_disegno end as 'u_disegno',T2.PlannedQty,T2.U_PRG_WIP_QtaSpedita,T2.U_PRG_WIP_QtaDaTrasf,

case when t2.U_Qta_richiesta_wip is null then 0 else t2.U_Qta_richiesta_wip end as 'U_Qta_richiesta_wip' , T2.wareHouse ,
coalesce(case when COALESCE(T3.U_Ubicazione,'') ='' and coalesce(t3.u_ubicazione_labelling,'') <>'' then CONCAT('CAP3: ',t3.u_ubicazione_labelling) else T3.U_Ubicazione end,'') as 'U_ubicazione'
, CASE WHEN t5.itemcode <> '' then 'S' when T4.ONHAND>=T2.U_PRG_WIP_QtaDaTrasf THEN 'T' ELSE 'NT' END AS 'TRASFERIBILE' 
, T4.ONHAND AS 'Q_TRASF'
, case when t1.u_progressivo_commessa is null then 0 else t1.u_progressivo_commessa end as 'u_progressivo_commessa'
, case when t1.u_prg_azs_commessa is null then '' else t1.u_prg_azs_commessa end as 'u_prg_azs_commessa'

from [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 left join owor t1 on t1.docnum=t0.docnum and t0.tipo_doc='ODP' 
left join wor1 t2 on t2.docentry=t1.docentry
LEFT JOIN OITM T3 ON T3.ITEMCODE=T2.ITEMCODE
LEFT JOIN OITW T4 ON T4.ITEMCODE=T2.ITEMCODE AND T2.WAREHOUSE=T4.WhsCode
left join [Tirelli_40].[dbo].[Lotto_prelievo_skippati] t5 on t5.itemcode=t2.itemcode and t5.id_lotto=t0.id and t5.docnum=t1.docnum and t0.tipo_doc=t5.tipo_doc

left join
( select t0.itemcode, sum(t0.onhand) as'Q'
from oitw t0 where (t0.WhsCode='06' or t0.WhsCode='CAP2')
and t0.onhand >0
group by t0.itemcode) A on a.itemcode=t2.itemcode


where t0.id=" & id_lotto_prelievo & " AND substring(t2.itemcode,1,1) <> 'L' and T2.U_PRG_WIP_QtaDaTrasf>0 and  T2.PlannedQty> coalesce(t2.U_Qta_richiesta_wip,0)  " & filtro_magazzino & "

ORDER BY 
CASE WHEN t5.itemcode <> '' then 'S' when T4.ONHAND>=T2.U_PRG_WIP_QtaDaTrasf THEN 'T' ELSE 'NT' END desc,
T2.WAREHOUSE,
U_UBICAZIONE, t2.itemcode
--DA AGGIUNGERE PARTE DEGLI OC CON UNION ALL
"
        Else
            CMD_SAP_2.CommandText = "
SELECT *
FROM
(
select T0.Tipo_doc,T1.DOCNUM,T2.ItemCode,t2.linenum,T3.ITEMNAME,
case when t3.u_disegno is null then '' else t3.u_disegno end as 'u_disegno',T2.PlannedQty,T2.U_PRG_WIP_QtaSpedita,T2.U_PRG_WIP_QtaDaTrasf,t2.U_Qta_richiesta_wip, T2.wareHouse ,
coalesce(case when COALESCE(T3.U_Ubicazione,'') ='' and coalesce(t3.u_ubicazione_labelling,'') <>'' then CONCAT('CAP3: ',t3.u_ubicazione_labelling) else T3.U_Ubicazione end,'') as 'U_ubicazione'
, CASE WHEN t5.itemcode <> '' then 'S' when T4.ONHAND>0 and T2.U_PRG_WIP_QtaDaTrasf>0  THEN 'T' ELSE 'NT' END AS 'TRASFERIBILE' 
, T4.ONHAND AS 'Q_TRASF'
, case when t1.u_progressivo_commessa is null then 0 else t1.u_progressivo_commessa end as 'u_progressivo_commessa'
, case when t1.u_prg_azs_commessa is null then '' else t1.u_prg_azs_commessa end as 'u_prg_azs_commessa'

from [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 left join owor t1 on t1.docnum=t0.docnum and t0.tipo_doc='ODP' and (t1.status='R')
left join wor1 t2 on t2.docentry=t1.docentry
LEFT JOIN OITM T3 ON T3.ITEMCODE=T2.ITEMCODE
LEFT JOIN OITW T4 ON T4.ITEMCODE=T2.ITEMCODE AND T2.WAREHOUSE=T4.WhsCode
left join [Tirelli_40].[dbo].[Lotto_prelievo_skippati] t5 on t5.itemcode=t2.itemcode and t5.id_lotto=t0.id and t5.docnum=t1.docnum and t0.tipo_doc=t5.tipo_doc


where t0.id=" & id_lotto_prelievo & " AND substring(t2.itemcode,1,1) <> 'L' and T2.U_PRG_WIP_QtaDaTrasf>0 " & filtro_magazzino & "


union all

select T0.Tipo_doc,T1.DOCNUM,T2.ItemCode,t2.linenum,T3.ITEMNAME,
case when t3.u_disegno is null then '' else t3.u_disegno end as 'u_disegno',T2.OpenQty,T2.U_Trasferito,T2.U_Datrasferire,t2.U_Trasferito, T2.WhsCode ,
coalesce(case when COALESCE(T3.U_Ubicazione,'') ='' and coalesce(t3.u_ubicazione_labelling,'') <>'' then CONCAT('CAP3: ',t3.u_ubicazione_labelling) else T3.U_Ubicazione end,'') as 'U_ubicazione'
, CASE WHEN t5.itemcode <> '' then 'S' when T4.ONHAND>0 and T2.U_Datrasferire>0  THEN 'T' ELSE 'NT' END AS 'TRASFERIBILE' 
, T4.ONHAND AS 'Q_TRASF'
, 0 as 'u_progressivo_commessa'
, '' 'u_prg_azs_commessa'

from [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 left join ordr t1 on t1.docnum=t0.docnum and t0.tipo_doc='OC' and (t1.docstatus='O')
left join RDR1 t2 on t2.docentry=t1.docentry
LEFT JOIN OITM T3 ON T3.ITEMCODE=T2.ITEMCODE
LEFT JOIN OITW T4 ON T4.ITEMCODE=T2.ITEMCODE AND T2.WhsCode=T4.WhsCode
left join [Tirelli_40].[dbo].[Lotto_prelievo_skippati] t5 on t5.itemcode=t2.itemcode and t5.id_lotto=t0.id and t5.docnum=t1.docnum and t0.tipo_doc=t5.tipo_doc


where t0.id=" & id_lotto_prelievo & " AND substring(t2.itemcode,1,1) <> 'L' and T2.U_Datrasferire>0 " & filtro_magazzino_oc & "

)
AS T10

ORDER BY 
T10.TRASFERIBILE desc,
T10.WAREHOUSE,
T10.U_UBICAZIONE, t10.itemcode



"
        End If



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            If contatore = 0 Then
                itemcode = cmd_SAP_reader_2("itemcode")
                docnum = cmd_SAP_reader_2("DOCNUM")
                tipo_doc = cmd_SAP_reader_2("Tipo_doc")
            End If
            DataGridView1.Rows.Add(contatore, tipo_doc, cmd_SAP_reader_2("DOCNUM"), cmd_SAP_reader_2("u_prg_azs_commessa"), cmd_SAP_reader_2("u_progressivo_commessa"), cmd_SAP_reader_2("linenum"), cmd_SAP_reader_2("ItemCode"), cmd_SAP_reader_2("ITEMNAME"), cmd_SAP_reader_2("u_disegno"), cmd_SAP_reader_2("PlannedQty"), cmd_SAP_reader_2("U_PRG_WIP_QtaSpedita"), cmd_SAP_reader_2("U_Qta_richiesta_wip"), cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), cmd_SAP_reader_2("wareHouse"), cmd_SAP_reader_2("U_Ubicazione"), cmd_SAP_reader_2("TRASFERIBILE"), cmd_SAP_reader_2("Q_TRASF"))
            If cmd_SAP_reader_2("TRASFERIBILE") = "T" Then
                contatore_trasferibili = contatore_trasferibili + 1
            End If
            contatore = contatore + 1
        Loop


        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        DataGridView1.ClearSelection()
        codice_da_prelevare(id_lotto_prelievo, itemcode, tipo_doc, docnum)

        Label13.Text = contatore_trasferibili
    End Sub

    Sub RIEMPI_datagridview_prelievi_assiemati(id_lotto_prelievo As Integer, par_itemcode As String)

        DataGridView2.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "
select t10.docnum, t10.itemcode, t10.itemname, T10.U_PRG_WIP_QtaDaTrasf
from
(
select T0.Tipo_doc,T1.DOCNUM,T2.ItemCode,t2.linenum,T3.ITEMNAME,
case when t3.u_disegno is null then '' else t3.u_disegno end as 'u_disegno',T2.PlannedQty,T2.U_PRG_WIP_QtaSpedita,T2.U_PRG_WIP_QtaDaTrasf, T2.wareHouse , T3.U_Ubicazione 
, CASE WHEN t5.itemcode <> '' then 'S' when T4.ONHAND>=T2.U_PRG_WIP_QtaDaTrasf THEN 'T' ELSE 'NT' END AS 'TRASFERIBILE' 
, T4.ONHAND AS 'Q_TRASF'
from [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 left join owor t1 on t1.docnum=t0.docnum and t0.tipo_doc='ODP'
left join wor1 t2 on t2.docentry=t1.docentry
LEFT JOIN OITM T3 ON T3.ITEMCODE=T2.ITEMCODE
LEFT JOIN OITW T4 ON T4.ITEMCODE=T2.ITEMCODE AND T2.WAREHOUSE=T4.WhsCode
left join [Tirelli_40].[dbo].[Lotto_prelievo_skippati] t5 on t5.itemcode=t2.itemcode and t5.id_lotto=t0.id and t5.docnum=t1.docnum and t0.tipo_doc=t5.tipo_doc



where t0.id=" & id_lotto_prelievo & " AND T2.U_PRG_WIP_QtaDaTrasf>0 and t2.itemcode='" & par_itemcode & "'




--DA AGGIUNGERE PARTE DEGLI OC CON UNION ALL
)
as t10
order by T10.U_PRG_WIP_QtaDaTrasf
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            DataGridView2.Rows.Add(cmd_SAP_reader_2("ItemCode"), cmd_SAP_reader_2("ITEMNAME"), cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"))
        Loop


        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        DataGridView2.ClearSelection()


    End Sub

    Sub codice_da_prelevare(id_lotto_prelievo As Integer, par_itemcode As String, par_tipo_doc As String, par_docnum As Integer)
        Dim contatore As Integer = 0

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        If par_tipo_doc = "ODP" Then
            CMD_SAP_2.CommandText = "select T0.Tipo_doc,T1.DOCNUM,t1.status
, case when t1.u_prg_azs_commessa is null then '' else t1.u_prg_azs_commessa end as 'u_prg_azs_commessa', t2.linenum,T2.ItemCode,T3.ITEMNAME,
case when t3.u_disegno is null then '' else t3.u_disegno end as 'u_disegno',T2.PlannedQty,T2.U_PRG_WIP_QtaSpedita,T2.U_PRG_WIP_QtaDaTrasf, T2.wareHouse , 
coalesce(case when COALESCE(T3.U_Ubicazione,'')='' and coalesce(t3.u_ubicazione_labelling,'')<>''  then CONCAT('CAP3: ',t3.u_ubicazione_labelling) else T3.U_Ubicazione end,'') as 'U_ubicazione'
, CASE WHEN t5.itemcode <> '' then 'S' when T4.ONHAND>=T2.U_PRG_WIP_QtaDaTrasf THEN 'T' ELSE 'NT' END AS 'TRASFERIBILE' 
, T4.ONHAND AS 'Q_A_Mag'
,CASE WHEN t2.U_Qta_richiesta_wip is null then 0 else t2.U_Qta_richiesta_wip end as 'U_Qta_richiesta_wip' 
, case when t1.u_progressivo_commessa is null then 0 else t1.u_progressivo_commessa end as 'u_progressivo_commessa'
,t1.warehouse as 'MagDest'
,coalesce(t3.u_codice_brb,'') as 'Codice_BRB'
from [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 left join owor t1 on t1.docnum=t0.docnum and t0.tipo_doc='ODP'
left join wor1 t2 on t2.docentry=t1.docentry
LEFT JOIN OITM T3 ON T3.ITEMCODE=T2.ITEMCODE
LEFT JOIN OITW T4 ON T4.ITEMCODE=T2.ITEMCODE AND T2.WAREHOUSE=T4.WhsCode
left join [Tirelli_40].[dbo].[Lotto_prelievo_skippati] t5 on t5.itemcode=t2.itemcode and t5.id_lotto=t0.id and t5.docnum=t1.docnum and t0.tipo_doc=t5.tipo_doc

where t0.id=" & id_lotto_prelievo & " AND T2.U_PRG_WIP_QtaDaTrasf>0 and t2.itemcode='" & par_itemcode & "' and t0.tipo_doc ='" & par_tipo_doc & "' and t1.docnum ='" & par_docnum & "'

ORDER BY 
CASE WHEN t5.itemcode <> '' then 'S' when T4.ONHAND>=T2.U_PRG_WIP_QtaDaTrasf THEN 'T' ELSE 'NT' ENd DESC,
T2.WAREHOUSE,
U_UBICAZIONE


"

        ElseIf par_tipo_doc = "OC" Then
            CMD_SAP_2.CommandText = "select T0.Tipo_doc,T1.DOCNUM,t1.docstatus as 'Status'
, '' as 'u_prg_azs_commessa',
t2.linenum,T2.ItemCode,T3.ITEMNAME,
case when t3.u_disegno is null then '' else t3.u_disegno end as 'u_disegno'
,T2.openqty,T2.U_Trasferito as 'Q_TRASF',T2.U_Datrasferire as 'U_PRG_WIP_QtaDaTrasf', T2.WhsCode as 'Warehouse', 

coalesce(case when COALESCE(T3.U_Ubicazione,'')='' and coalesce(t3.u_ubicazione_labelling,'')<>'' then CONCAT('CAP3: ',t3.u_ubicazione_labelling) else T3.U_Ubicazione end,'') as 'U_ubicazione'
, CASE WHEN t5.itemcode <> '' then 'S' when T4.ONHAND>=T2.U_Datrasferire THEN 'T' ELSE 'NT' END AS 'TRASFERIBILE' 
, T4.ONHAND AS 'Q_A_Mag'
,0 as 'U_Qta_richiesta_wip' 
, 0 as 'u_progressivo_commessa'
,'' as 'MagDest'
,coalesce(t3.u_codice_brb,'') as 'Codice_BRB'

from [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 left join ordr t1 on t1.docnum=t0.docnum and t0.tipo_doc='OC'
left join rdr1 t2 on t2.docentry=t1.docentry
LEFT JOIN OITM T3 ON T3.ITEMCODE=T2.ITEMCODE
LEFT JOIN OITW T4 ON T4.ITEMCODE=T2.ITEMCODE AND T2.WhsCode=T4.WhsCode
left join [Tirelli_40].[dbo].[Lotto_prelievo_skippati] t5 on t5.itemcode=t2.itemcode and t5.id_lotto=t0.id and t5.docnum=t1.docnum and t0.tipo_doc=t5.tipo_doc

where t0.id=" & id_lotto_prelievo & " AND T2.U_Datrasferire>0 and t2.itemcode='" & par_itemcode & "' and t0.tipo_doc ='" & par_tipo_doc & "' and t1.docnum ='" & par_docnum & "'

ORDER BY 
CASE WHEN t5.itemcode <> '' then 'S' when T4.ONHAND>=T2.U_Datrasferire THEN 'T' ELSE 'NT' ENd DESC,
T2.whscode,
U_UBICAZIONE


"
        End If



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            If contatore = 0 Then
                codice_sap = cmd_SAP_reader_2("ITEMcode")
                Button1.Text = codice_sap
                Label2.Text = cmd_SAP_reader_2("itemname")
                Label16.Text = cmd_SAP_reader_2("Codice_BRB")
                Button2.Text = cmd_SAP_reader_2("u_disegno")
                Label6.Text = cmd_SAP_reader_2("warehouse")
                Label7.Text = cmd_SAP_reader_2("u_ubicazione")
                Label8.Text = Math.Round(cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), 2)
                Label9.Text = Math.Round(cmd_SAP_reader_2("Q_A_Mag"), 2)
                Label14.Text = Math.Round(cmd_SAP_reader_2("U_Qta_richiesta_wip"), 2)
                Button3.Text = cmd_SAP_reader_2("docnum")
                Label10.Text = cmd_SAP_reader_2("linenum")
                Label11.Text = cmd_SAP_reader_2("u_prg_azs_commessa")
                Label12.Text = cmd_SAP_reader_2("status")
                Label5.Text = cmd_SAP_reader_2("u_progressivo_commessa")
                Label15.Text = cmd_SAP_reader_2("MagDest")


                TextBox1.Text = minore(cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), cmd_SAP_reader_2("Q_A_Mag"))
                Label17.Text = par_tipo_doc

            End If

            contatore = contatore + 1
        Loop


        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        RIEMPI_datagridview_prelievi_assiemati(id_lotto_prelievo, par_itemcode)

    End Sub

    Public Function minore(par_primo_numero As Decimal, par_secondo_numero As Decimal)
        Dim risultato As Decimal

        If par_primo_numero <= par_secondo_numero Then
            risultato = par_primo_numero

        Else
            risultato = par_secondo_numero
        End If
        risultato = Math.Round(risultato, 3)
        Return risultato
    End Function



    Sub info_testata_lotto_di_preliavo(id_lotto_prelievo As Integer)


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "declare @id_lotto_prelievo as integer
set @id_lotto_prelievo=' " & id_lotto_prelievo & "'


select t0.id,t0.data,t0.ora,t0.dip,concat(t1.lastname,' ', t1.firstname) as 'Mag', t0.commento 

from [Tirelli_40].[dbo].[Lotto_prelievo_testata] t0 left join [TIRELLI_40].[dbo].OHEM t1 on t1.userid=t0.dip

where t0.id=@id_lotto_prelievo
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then

            Label3.Text = cmd_SAP_reader_2("mag")
            Label4.Text = cmd_SAP_reader_2("data")
            Label1.Text = cmd_SAP_reader_2("commento")


        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        If Label1.Text = "MU" Then
            Panel1.BackColor = Color.PaleGreen
            TableLayoutPanel1.BackColor = Color.PaleGreen
            Button12.Visible = True
            Button6.Visible = False
        Else
            Panel1.BackColor = Color.SteelBlue
            TableLayoutPanel1.BackColor = Color.SteelBlue
            Button12.Visible = False
            Button6.Visible = True
        End If
    End Sub

    Private Sub Cmd_Indietro_Click(sender As Object, e As EventArgs) Handles Cmd_Indietro.Click
        id_lotto = Int(id_lotto) - 1
        inizializzazione_lotto_di_prelievo()
    End Sub

    Private Sub Cmd_Avanti_Click(sender As Object, e As EventArgs) Handles Cmd_Avanti.Click
        id_lotto = Int(id_lotto) + 1
        inizializzazione_lotto_di_prelievo()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Public Sub ultimo_lotto()
        Dim Cnn_Entrate_Merci As New SqlConnection
        Dim Cmd_Entrate_Merci As New SqlCommand
        Dim Cmd_Entrate_Merci_Reader As SqlDataReader

        Cnn_Entrate_Merci.ConnectionString = Homepage.sap_tirelli
        Cnn_Entrate_Merci.Open()
        Cmd_Entrate_Merci.Connection = Cnn_Entrate_Merci


        Cmd_Entrate_Merci.CommandText = "select max(id) as 'MAX_docnum'
from [Tirelli_40].[dbo].[Lotto_prelievo_testata]"



        Cmd_Entrate_Merci_Reader = Cmd_Entrate_Merci.ExecuteReader
        Cmd_Entrate_Merci_Reader.Read()
        id_lotto = Cmd_Entrate_Merci_Reader("Max_docnum")
        Cnn_Entrate_Merci.Close()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click


        ' Itera all'indietro attraverso le righe selezionate nella DataGridView "datagridview1"
        For i As Integer = DataGridView4.SelectedRows.Count - 1 To 0 Step -1
            Dim row As DataGridViewRow = DataGridView4.SelectedRows(i)

            If MessageBox.Show($"Sei sicuro di voler togliere l'ODP '{row.Cells("Docnum").Value & " " & row.Cells("Desc").Value}' dal lotto di prelievo?", "Elimina file", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

                elimina_righe_lotto_prelievo(id_lotto, row.Cells("Docnum").Value, row.Cells("tipo").Value)
                DataGridView4.Rows.Remove(row)
                RIEMPI_datagridview_CODICI_PRELIEVO(id_lotto)
            End If


        Next
    End Sub

    Sub elimina_righe_lotto_prelievo(id_lotto_prelievo As Integer, par_Docnum_odp As Integer, par_tipo_doc As String)

        Dim Cnn As New SqlConnection

        For Each row As DataGridViewRow In DataGridView1.Rows
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()
            Dim CMD_SAP As New SqlCommand
            CMD_SAP.Connection = Cnn

            CMD_SAP.CommandText = "delete [Tirelli_40].[dbo].[lotto_prelievo_riga]
where id = " & id_lotto_prelievo & " and tipo_doc ='" & par_tipo_doc & "' and docnum = " & par_Docnum_odp & ""
            CMD_SAP.ExecuteNonQuery()
            Cnn.Close()


        Next

    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim par_datagridview As DataGridView = DataGridView1

        If e.RowIndex >= 0 Then


            codice_da_prelevare(id_lotto, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Itemcode_codice").Value, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Tipo_doc").Value, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Docnum_padre").Value)

            If e.ColumnIndex = par_datagridview.Columns.IndexOf(Docnum_padre) And par_datagridview.Rows(e.RowIndex).Cells(columnName:="Tipo_doc").Value = "ODP" Then

                ODP_Form.docnum_odp = par_datagridview.Rows(e.RowIndex).Cells(columnName:="docnum_padre").Value
                ODP_Form.Show()
                ODP_Form.inizializza_form(par_datagridview.Rows(e.RowIndex).Cells(columnName:="docnum_padre").Value)



            End If

            If e.ColumnIndex = par_datagridview.Columns.IndexOf(Itemcode_codice) Then

                Magazzino.Codice_SAP = par_datagridview.Rows(e.RowIndex).Cells(columnName:="Itemcode_codice").Value

                ' Ripristina la finestra se è minimizzata
                If Magazzino.WindowState = FormWindowState.Minimized Then
                    Magazzino.WindowState = FormWindowState.Normal
                End If

                ' Porta la finestra in primo piano
                Magazzino.BringToFront()
                Magazzino.Activate()
                Magazzino.Show()

                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(codice_sap)
            End If
        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Button6.Enabled = False

        Dim qta_a_mag As Decimal = Form_Entrate_Merci.giacenze_IN_magazzino(codice_sap, Label6.Text)
        Dim qta_da_trasferire_nel_doc As Decimal = Form_Entrate_Merci.quantita_da_trasferire_nel_documento(Label17.Text, codice_sap, Button3.Text, Button3.Text, Label10.Text)

        Dim qta_trasferimento As Decimal = minore(qta_a_mag, qta_da_trasferire_nel_doc)


        Label9.Text = Math.Round(qta_a_mag, 3)
        Label8.Text = Math.Round(qta_da_trasferire_nel_doc, 3)

        Dim magazzino_WIP As String

        If Label15.Text = "B02" Or Label15.Text = "BCAP2" Then
            magazzino_WIP = "BWIP"
        Else
            magazzino_WIP = "WIP"
        End If

        If qta_trasferimento <= 0 Then
            MsgBox("Non è possibile trasferire quantità <= 0")
        Else

            If Label17.Text = "ODP" Then



                If check_odp_rilasciato(Button3.Text) = "R" Then
                    If TextBox1.Text > qta_trasferimento Then
                        MsgBox("Ridurre la quantità in trasferimento")
                    Else

                        Form_Entrate_Merci.Trasferimento_in_WIP(Label17.Text, codice_sap, Button3.Text, 0, Replace(TextBox1.Text, ",", "."), Label6.Text, magazzino_WIP, Label10.Text, Label11.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, "Manuale", 0, 0, "Trasferimento")
                        If CheckBox1.Checked = True Then
                            Fun_Stampa(magazzino_WIP, Label17.Text, Button3.Text, codice_sap, Replace(TextBox1.Text, ",", "."))
                        End If
                        If Form_Entrate_Merci.variabile_controllo_trasferimento_1 = 0 Then
                            MsgBox("Codice " & codice_sap & vbCrLf & " Quantità " & Replace(TextBox1.Text, ",", ".") & vbCrLf & "Nell'ODP " & Button3.Text & vbCrLf & vbCrLf & " Trasferito con successo")
                            RIEMPI_datagridview_CODICI_PRELIEVO(id_lotto)
                        End If

                    End If
                ElseIf check_odp_rilasciato(Button3.Text) = "P" Then

                    Dim result As Integer
                    result = MessageBox.Show("L'ordine di produzione risulta PIANIFICATO, vuoi rilasciarlo?", "Confirmation", MessageBoxButtons.YesNo)
                    If result = DialogResult.Yes Then
                        Form_Entrate_Merci.cambia_stato(Button3.Text)

                        result = MessageBox.Show("Stampare l'ordine di produzione appena RILASCIATO?", "Confirmation", MessageBoxButtons.YesNo)
                        If result = DialogResult.Yes Then
                            Form_Entrate_Merci.stampa_odp_foglio(Button3.Text)
                        End If


                        RIEMPI_datagridview_CODICI_PRELIEVO(id_lotto)
                    End If


                End If
            Else

                If TextBox1.Text > qta_trasferimento Then
                    MsgBox("Ridurre la quantità in trasferimento")
                Else

                    Form_Entrate_Merci.Trasferimento_in_WIP(Label17.Text, codice_sap, Button3.Text, Button3.Text, Replace(TextBox1.Text, ",", "."), Label6.Text, magazzino_WIP, Label10.Text, Label11.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, "Manuale", 0, 0, "Trasferimento")
                    If CheckBox1.Checked = True Then
                        Fun_Stampa(magazzino_WIP, Label17.Text, Button3.Text, codice_sap, Replace(TextBox1.Text, ",", "."))
                    End If
                    If Form_Entrate_Merci.variabile_controllo_trasferimento_1 = 0 Then
                        MsgBox("Codice " & codice_sap & vbCrLf & " Quantità " & Replace(TextBox1.Text, ",", ".") & vbCrLf & "Nell'OC " & Button3.Text & vbCrLf & vbCrLf & " Trasferito con successo")
                        RIEMPI_datagridview_CODICI_PRELIEVO(id_lotto)
                    End If

                End If
            End If
        End If
        Button6.Enabled = True
    End Sub

    Private Sub TextBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox1.KeyPress
        ' Verifica se il tasto premuto non è un numero o una virgola
        If (Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "," AndAlso e.KeyChar <> ControlChars.Back) Then
            ' Impedisce la digitazione del tasto premuto
            e.Handled = True
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        salta_codice(id_lotto, codice_sap, Label17.Text, Button3.Text)
        RIEMPI_datagridview_CODICI_PRELIEVO(id_lotto)
    End Sub

    Sub salta_codice(id_lotto_prelievo As Integer, par_itemcode As String, Par_tipo_doc As String, par_docnum As Integer)

        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "
delete [Tirelli_40].[dbo].[Lotto_prelievo_skippati] where id_lotto=" & id_lotto_prelievo & " and itemcode='" & par_itemcode & "' and tipo_doc='" & Par_tipo_doc & "' and docnum='" & par_docnum & "'
insert into [Tirelli_40].[dbo].[Lotto_prelievo_skippati] (id_lotto, itemcode, tipo_doc,docnum) values (" & id_lotto_prelievo & ",'" & par_itemcode & "','" & Par_tipo_doc & "','" & par_docnum & "')"
        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()




    End Sub

    Sub togli_saltato_codice(id_lotto_prelievo As Integer, par_itemcode As String, Par_tipo_doc As String, par_docnum As Integer)


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "
delete [Tirelli_40].[dbo].[Lotto_prelievo_skippati] where id_lotto=" & id_lotto_prelievo & " and itemcode='" & par_itemcode & "' and tipo_doc='" & Par_tipo_doc & "' and docnum='" & par_docnum & "'"
        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()




    End Sub

    Function giacenze_a_magazzino_dato_il_magazzino(par_codice_sap As String, par_codice_magazzino As String)

        Dim qta_a_mag As Decimal
        Dim Cnn1 As New SqlConnection




        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.[WhsCode], CASE WHEN T0.[OnHand] is null then 0 else T0.[OnHand] END AS 'onhand' 

FROM OITW T0 WHERE  t0.itemcode='" & par_codice_sap & "'and T0.[WhsCode]='" & par_codice_magazzino & "' "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then


            qta_a_mag = cmd_SAP_reader_2("onhand")
        Else

            qta_a_mag = 0

        End If


        Cnn1.Close()
        Return qta_a_mag
    End Function

    Function check_da_trasferire_odp(par_codice_sap As String, par_odp As Integer, par_linenum As Integer)

        Dim qta_trasferibile As Decimal
        Dim Cnn1 As New SqlConnection
        Dim Cnn As New SqlConnection



        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T1.[ItemCode], case when t1.u_prg_wip_qtadatrasf is null then 0 else t1.u_prg_wip_qtadatrasf end as 'u_prg_wip_qtadatrasf'
FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T0.[DocNum] ='" & par_odp & "' and t1.linenum='" & par_linenum & "' and t1.itemcode ='" & par_codice_sap & "' "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then


            qta_trasferibile = cmd_SAP_reader_2("u_prg_wip_qtadatrasf")
        Else

            qta_trasferibile = 0

        End If


        Cnn.Close()
        Return qta_trasferibile
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click


        ODP_Form.docnum_odp = Button3.Text
        ODP_Form.Show()
        ODP_Form.inizializza_form(Button3.Text)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Magazzino.Codice_SAP = Button1.Text
        ' Ripristina la finestra se è minimizzata
        If Magazzino.WindowState = FormWindowState.Minimized Then
            Magazzino.WindowState = FormWindowState.Normal
        End If

        ' Porta la finestra in primo piano
        Magazzino.BringToFront()
        Magazzino.Activate()
        Magazzino.Show()

        Magazzino.TextBox2.Text = Magazzino.Codice_SAP
        Magazzino.OttieniDettagliAnagrafica(codice_sap)
    End Sub

    Private Sub Combodipendenti_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combodipendenti.SelectedIndexChanged
        ' Codicedip = Elenco_dipendenti(Combodipendenti.SelectedIndex)

        Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato = Elenco_dipendenti(Combodipendenti.SelectedIndex)

        Homepage.Aggiorna_INI_COMPUTER()
    End Sub

    Sub Inserimento_dipendenti()
        Dim Cnn As New SqlConnection
        Combodipendenti.Items.Clear()
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[userid] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 
left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
inner join [TIRELLI_40].[DBO].COLL_Reparti t2 on (t2.sap_id_reparto =t1.code or t2.sap_id_reparto_2 =t1.code)  
where t0.active='Y' AND T0.[userid] <>'' and t2.id_reparto='" & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto & "'  order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()

            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            Combodipendenti.Items.Add(cmd_SAP_reader("Nome"))

            If Elenco_dipendenti(Indice) = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato Then
                Combodipendenti.SelectedIndex = Indice
            End If

            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Function check_odp_rilasciato(PAR_NUMERO_ODP As String)

        Dim stato_odp As String
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.STATUS
        FROM OWOR T0 WHERE T0.DOCNUM='" & PAR_NUMERO_ODP & "' "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then

            stato_odp = cmd_SAP_reader("STATUS")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
        Return stato_odp
    End Function

    Private Sub Button8_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ' Itera all'indietro attraverso le righe selezionate nella DataGridView "datagridview1"
        For i As Integer = DataGridView4.SelectedRows.Count - 1 To 0 Step -1
            Dim row As DataGridViewRow = DataGridView4.SelectedRows(i)



            Form_Entrate_Merci.cambia_stato(row.Cells("Docnum").Value)





        Next
        RIEMPI_datagridview_documenti_lotto(id_lotto)
        RIEMPI_datagridview_documenti_lotto_oc(id_lotto, DataGridView3)

        RIEMPI_datagridview_CODICI_PRELIEVO(id_lotto)
    End Sub



    Private Sub DataGridView4_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView4.CellFormatting
        Try
            If DataGridView4.Rows(e.RowIndex).Cells(columnName:="status").Value = "R" Then



                DataGridView4.Rows(e.RowIndex).Cells(columnName:="status").Style.BackColor = Color.Lime
            ElseIf DataGridView4.Rows(e.RowIndex).Cells(columnName:="status").Value = "P" Then
                DataGridView4.Rows(e.RowIndex).Cells(columnName:="status").Style.BackColor = Color.Orange
            End If
        Catch ex As Exception

        End Try

    End Sub



    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        If e.RowIndex >= 0 Then


            If e.ColumnIndex = DataGridView4.Columns.IndexOf(Docnum) Then

                ODP_Form.docnum_odp = DataGridView4.Rows(e.RowIndex).Cells(columnName:="docnum").Value
                ODP_Form.Show()
                ODP_Form.inizializza_form(DataGridView4.Rows(e.RowIndex).Cells(columnName:="docnum").Value)



            ElseIf e.ColumnIndex = DataGridView4.Columns.IndexOf(itemcode) Then

                Magazzino.Codice_SAP = DataGridView4.Rows(e.RowIndex).Cells(columnName:="itemcode").Value

                ' Ripristina la finestra se è minimizzata
                If Magazzino.WindowState = FormWindowState.Minimized Then
                    Magazzino.WindowState = FormWindowState.Normal
                End If

                ' Porta la finestra in primo piano
                Magazzino.BringToFront()
                Magazzino.Activate()
                Magazzino.Show()

                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(codice_sap)

            ElseIf e.ColumnIndex = DataGridView4.Columns.IndexOf(DIS) Then


                Magazzino.visualizza_disegno(DataGridView4.Rows(e.RowIndex).Cells(columnName:="Dis").Value)


            End If


        End If
    End Sub



    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting


        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="itemcode_codice").Value = Button1.Text Then
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Aquamarine

        Else
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White

        End If




        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="trasferibile").Value = "T" Then



            DataGridView1.Rows(e.RowIndex).Cells(columnName:="trasferibile").Style.BackColor = Color.Lime
        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="trasferibile").Value = "S" Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="trasferibile").Style.BackColor = Color.Orange

        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="trasferibile").Value = "NT" Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="trasferibile").Style.BackColor = Color.Red
        End If




    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Magazzino.codice_disegno = Button2.Text
        Magazzino.visualizza_disegno(Magazzino.codice_disegno)
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        filtro_magazzino = "and "
        filtro_magazzino_oc = "and "
        For Each selectedItem As Object In CheckedListBox1.CheckedItems
            Dim selectedText As String = selectedItem.ToString()
            Dim mag_selezionato As String = selectedText

            '
            If filtro_magazzino = "and " Then
                filtro_magazzino = filtro_magazzino & "(T2.wareHouse= '" & mag_selezionato & "'"

                filtro_magazzino_oc = filtro_magazzino_oc & "(T2.whscode= '" & mag_selezionato & "'"
            Else
                filtro_magazzino = filtro_magazzino & " or T2.wareHouse= '" & mag_selezionato & "'"
                filtro_magazzino_oc = filtro_magazzino_oc & " or T2.whscode= '" & mag_selezionato & "'"
            End If
        Next
        filtro_magazzino = filtro_magazzino & ")"
        filtro_magazzino_oc = filtro_magazzino_oc & ")"
        RIEMPI_datagridview_CODICI_PRELIEVO(id_lotto)

    End Sub

    Sub carica_magazzini(id_lotto_prelievo As Integer)

        CheckedListBox1.Items.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = " 
 select *
 from
 (
 select T2.wareHouse 
from [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 left join owor t1 on t1.docnum=t0.docnum and t0.tipo_doc='ODP'
left join wor1 t2 on t2.docentry=t1.docentry

where t0.id=" & id_lotto_prelievo & " AND T2.U_PRG_WIP_QtaDaTrasf>0 
group by T2.wareHouse 

union all

 select T2.whscode

from [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 left join ordr t1 on t1.docnum=t0.docnum and t0.tipo_doc='OC'
left join rdr1 t2 on t2.docentry=t1.docentry

where t0.id=" & id_lotto_prelievo & " AND T2.U_datrasferire>0 
group by T2.whscode
)
as t10
group by t10.wareHouse
order by t10.wareHouse

"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            CheckedListBox1.Items.Add(cmd_SAP_reader_2("warehouse"), True)

        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    Private Sub Cmd_Cerca_Click(sender As Object, e As EventArgs) Handles Cmd_Cerca.Click
        id_lotto = Int(Txt_DocNum.Text)
        inizializzazione_lotto_di_prelievo()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        For Each riga As DataGridViewRow In DataGridView4.Rows


            If riga.Cells("tipo").Value = "ODP" Then

                Richiesta_trasferimento_materiale.riempi_datagridview_rt(riga.Cells("docnum").Value)
            End If

        Next
        Richiesta_trasferimento_materiale.Show()
    End Sub



    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click


        Ordine_di_produzione_lista.ID_lotto_di_prelievo = Ordine_di_produzione_lista.Trova_nuovo_lotto_di_prelievo()

        Ordine_di_produzione_lista.crea_testata_lotto_prelievo(Ordine_di_produzione_lista.ID_lotto_di_prelievo, "MU")
        trova_odp_con_trasferibili_per_mu()


        id_lotto = Ordine_di_produzione_lista.ID_lotto_di_prelievo
        MsgBox("Lotto di prelievo MU creato con successo")

        inizializzazione_lotto_di_prelievo()


    End Sub

    Sub trova_odp_con_trasferibili_per_mu()


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = " SELECT t0.docnum
FROM OWOR T0 INNER JOIN WOR1 T1 ON T0.DOCENTRY=T1.DOCENTRY
WHERE (T0.STATUS='P' OR T0.STATUS='R') AND T0.U_PRODUZIONE='INT' 
AND T1.U_PRG_WIP_QtaDaTrasf>0 and t1.ItemType=4 and t1.wareHouse<>'MUT'
group by t0.docnum"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            Ordine_di_produzione_lista.crea_righe_lotto_prelievo(Ordine_di_produzione_lista.ID_lotto_di_prelievo, cmd_SAP_reader_2("docnum"), "ODP")

        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    Function TRASFERIMENTO_AD_ALTRO_MAGAZZINO(par_codice_sap As String, par_magazzino_partenza As String, MAGAZZINO_DESTINAZIONE As String, par_quantità_da_trasferire As String, par_giacenza As String, par_tipo_documento As String, par_numero_documento As Integer, par_linenum_doc As Integer)

        If par_codice_sap = "" Then
            MsgBox("Selezionare un codice")

        Else

            Dim risultato_trasferimento As String

            ' Converte le stringhe in numeri
            Dim quantitaDaTrasferire As Double = Double.Parse(par_quantità_da_trasferire)
            Dim giacenza As Double = Double.Parse(par_giacenza)

            ' Effettua il confronto
            If quantitaDaTrasferire <= giacenza Then
                risultato_trasferimento = quantitaDaTrasferire
            Else
                risultato_trasferimento = giacenza
            End If

            If Form_Entrate_Merci.giacenze_IN_magazzino(par_codice_sap, par_magazzino_partenza) <= risultato_trasferimento Then
                risultato_trasferimento = Form_Entrate_Merci.giacenze_IN_magazzino(par_codice_sap, par_magazzino_partenza)

            End If


            Form_Entrate_Merci.Trasferisci_ad_altro_magazzino(par_codice_sap, Replace(risultato_trasferimento, ",", "."), 0, 0, par_magazzino_partenza, MAGAZZINO_DESTINAZIONE, par_linenum_doc, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, par_tipo_documento, par_numero_documento)

            Try
                Label9.Text = Label9.Text - risultato_trasferimento
            Catch ex As Exception

            End Try


            Form_Entrate_Merci.Quantità_trasferita_per_scontrino = Replace(risultato_trasferimento, ",", ".")
            Form_Entrate_Merci.CREA_SCONTRINO = "Y"

            MsgBox("Codice " & par_codice_sap & vbCrLf & " Quantità " & Replace(risultato_trasferimento, ",", ".") & vbCrLf & "Nel magazzino " & MAGAZZINO_DESTINAZIONE & vbCrLf & vbCrLf & " Trasferito con successo")




            RIEMPI_datagridview_CODICI_PRELIEVO(id_lotto)

        End If
    End Function

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Dim qta_a_mag As Decimal = Form_Entrate_Merci.giacenze_IN_magazzino(codice_sap, Label6.Text)
        Dim qta_da_trasferire_nel_doc As Decimal = Form_Entrate_Merci.quantita_da_trasferire_nel_documento("ODP", codice_sap, Button3.Text, 0, Label10.Text)

        Dim qta_trasferimento As Decimal = minore(qta_a_mag, qta_da_trasferire_nel_doc)


        Label9.Text = Math.Round(qta_a_mag, 2)
        Label8.Text = Math.Round(qta_da_trasferire_nel_doc, 2)

        If qta_trasferimento <= 0 Then
            MsgBox("Non è possibile trasferire quantità = 0")
        Else

            If TextBox1.Text > qta_trasferimento Then
                MsgBox("Quantità int Trasferimento > della quantità trasferibile")
            Else



                Magazzino.Documento = "ODP"

                TRASFERIMENTO_AD_ALTRO_MAGAZZINO(codice_sap, Label6.Text, "06", TextBox1.Text, Label9.Text, "ODP", Button3.Text, Label10.Text)
                If CheckBox1.Checked = True Then
                    Fun_Stampa("06", "ODP", Button3.Text, codice_sap, Replace(TextBox1.Text, ",", "."))
                End If

            End If
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        ' Itera all'indietro attraverso le righe selezionate nella DataGridView "datagridview1"
        If tipo_documento = "OC" Then


            For i As Integer = DataGridView3.SelectedRows.Count - 1 To 0 Step -1

                Dim row As DataGridViewRow = DataGridView3.SelectedRows(i)


                Form_nuova_offerta.Show()

                Form_nuova_offerta.TextBox10.Text = row.Cells("DataGridViewTextBoxColumn1").Value
                Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                Form_nuova_offerta.inizializzazione_form(row.Cells("DataGridViewTextBoxColumn1").Value, "ORDR", "RDR1", row.Cells("DataGridViewTextBoxColumn1").Value)
                Form_nuova_offerta.Fun_Stampa()
                'Layout_documenti.ComboBox1.SelectedIndex = 1

                'Layout_documenti.TextBox1.Text = row.Cells("DataGridViewTextBoxColumn1").Value
                'Layout_documenti.Show()
                'Layout_documenti.Button1.PerformClick()
            Next
        Else
            If RadioButton1.Checked = True Then


                For i As Integer = DataGridView4.SelectedRows.Count - 1 To 0 Step -1
                    Dim row As DataGridViewRow = DataGridView4.SelectedRows(i)


                    Form_Entrate_Merci.stampa_odp_foglio(row.Cells("Docnum").Value)
                Next

            Else
                For i As Integer = DataGridView4.SelectedRows.Count - 1 To 0 Step -1
                    Dim row As DataGridViewRow = DataGridView4.SelectedRows(i)

                    ODP_Form.testata_odp(row.Cells("Docnum").Value)
                    ODP_Form.Fun_Stampa()

                Next
            End If
        End If

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Homepage.Stampante_Selezionata = False

    End Sub



    Sub Fun_Stampa(par_mag_destinazione As String, par_tipo_documento As String, par_numero_documento As String, par_codice_sap As String, par_quantità As String)


        mag_destinazione = par_mag_destinazione
        tipo_documento = par_tipo_documento
        numero_documento = par_numero_documento
        codice_sap = par_codice_sap
        Magazzino.OttieniDettagliAnagrafica(codice_sap)

        quantità_trasferimento = par_quantità


        If par_tipo_documento = "ODP" Then

            ODP_Form.testata_odp(par_numero_documento)
        End If

        'Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", 185, 200)

        Sel_Stampante.AllowSomePages = False
        Sel_Stampante.ShowHelp = False
        Sel_Stampante.Document = Scontrino


        ' mettere qui la preview scontrino se si vorrà vedere
        If preview_scontrino = "SI" Then
            If Homepage.Stampante_Selezionata = False Then
                Dim result As DialogResult = Sel_Stampante.ShowDialog()
                If (result = DialogResult.OK) Then
                    Homepage.Stampante_Selezionata = True
                    ' Utilizza un PrintPreviewDialog per mostrare l'anteprima di stampa
                    Dim previewDialog As New PrintPreviewDialog()
                    Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", larghezza_scontrino_odp, altezza_scontrino_odp)
                    previewDialog.Document = Scontrino
                    result = previewDialog.ShowDialog()
                End If
            Else
                Dim previewDialog As New PrintPreviewDialog()
                Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", larghezza_scontrino_odp, altezza_scontrino_odp)
                previewDialog.Document = Scontrino
                Dim result As DialogResult = previewDialog.ShowDialog()
            End If
        Else
            If Homepage.Stampante_Selezionata = False Then
                Dim result As DialogResult = Sel_Stampante.ShowDialog()
                If (result = DialogResult.OK) Then
                    Homepage.Stampante_Selezionata = True
                    Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", larghezza_scontrino_odp, altezza_scontrino_odp)
                    ' Utilizza un PrintPreviewDialog per mostrare l'anteprima di stampa
                    Dim previewDialog As New PrintPreviewDialog()
                    previewDialog.Document = Scontrino
                    Scontrino.Print()
                End If
            Else
                Scontrino.Print()
            End If
        End If


    End Sub

    Private Sub Scontrino_PrintPage(sender As Object, e As PrintPageEventArgs) Handles Scontrino.PrintPage
        Dim Penna As New Pen(Color.Black)
        Dim Carattere_ODP As New Font("Calibri", 16, FontStyle.Bold)
        Dim Carattere_Matricola As New Font("Calibri", 25, FontStyle.Bold)

        Dim Carattere_numerone As New Font("Calibri", 60, FontStyle.Bold)
        Dim Carattere_Descrizione As New Font("Calibri", 14, FontStyle.Italic)
        Dim Carattere_Codice As New Font("Calibri", 22, FontStyle.Italic)
        Dim Carattere_Descrizione_Articolo As New Font("Calibri", 8, FontStyle.Italic)
        Dim Carattere_Qta As New Font("Calibri", 12, FontStyle.Bold)
        Dim Carattere_Ubicazione As New Font("Calibri", 12, FontStyle.Italic)
        Dim Carattere_posizione As New Font("Calibri", 16, FontStyle.Bold)
        Dim Carattere_Diciture As New Font("Calibri", 7, FontStyle.Italic)
        Dim Carattere_Diciture_secondarie As New Font("Calibri", 8, FontStyle.Italic)

        Dim unità_larghezza As Integer = larghezza_scontrino_odp / 2
        Dim unità_altezza As Integer = altezza_scontrino_odp / 4

        With e.Graphics
            .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
            .DrawRectangle(Penna, 1, 1, larghezza_scontrino_odp - 2, altezza_scontrino_odp - 5)

            .DrawRectangle(Penna, 1, 1, unità_larghezza, unità_altezza)

            .DrawString("Mag destinazione", Carattere_Diciture, Brushes.Black, 5, 5)

            .DrawString(mag_destinazione, Carattere_Matricola, Brushes.Black, unità_larghezza * 0.15, unità_altezza * 0.25)


            .DrawRectangle(Penna, unità_larghezza, 1, unità_larghezza, unità_altezza)

            .DrawString("N° documento", Carattere_Diciture, Brushes.Black, unità_larghezza + 4, 5)

            .DrawString(numero_documento, Carattere_ODP, Brushes.Black, unità_larghezza * 1.1, unità_altezza * 0.4)


            .DrawRectangle(Penna, 1, unità_altezza, unità_larghezza * 2, unità_altezza)

            .DrawString("Articolo", Carattere_Diciture, Brushes.Black, 5, unità_altezza + 4)

            .DrawString(codice_sap, Carattere_Codice, Brushes.Black, 5, unità_altezza * 1.1)

            .DrawString(Magazzino.OttieniDettagliAnagrafica(codice_sap).Descrizione, Carattere_Diciture_secondarie, Brushes.Black, 5, unità_altezza * 1.7)

            .DrawString(quantità_trasferimento & " Pz", Carattere_Diciture_secondarie, Brushes.Black, unità_larghezza * 1.4, unità_altezza * 1.4)


            .DrawRectangle(Penna, 1, unità_altezza * 2, unità_larghezza, unità_altezza)

            .DrawString("Commessa", Carattere_Diciture, Brushes.Black, 5, unità_altezza * 2 + 4)

            .DrawString(ODP_Form.testata_odp_commessa, Carattere_ODP, Brushes.Black, 5, unità_altezza * 2.25)

            .DrawRectangle(Penna, unità_larghezza, unità_altezza * 2, unità_larghezza, unità_altezza)

            .DrawString("Posizione", Carattere_Diciture, Brushes.Black, unità_larghezza + 4, unità_altezza * 2 + 4)

            .DrawString(ODP_Form.numerone, Carattere_ODP, Brushes.Black, unità_larghezza * 1.3, unità_altezza * 2.25)

            .DrawRectangle(Penna, 1, unità_altezza * 3, unità_larghezza * 2, unità_altezza)

            ' .DrawString("OCCCCCCCCCCCCCCC", Carattere_Codice, Brushes.Black, 5, 5)

            .DrawString(ODP_Form.testata_odp_prodname, Carattere_Diciture_secondarie, Brushes.Black, 5, unità_altezza * 3.15)


        End With
    End Sub

    Private Async Sub tabpage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Enter
        tipo_documento = "ODP"
    End Sub
    Private Async Sub tabpage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter
        tipo_documento = "OC"
    End Sub
    Private Sub TableLayoutPanel7_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel7.Paint

    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        For Each riga As DataGridViewRow In DataGridView1.SelectedRows
            ' Assicurati di sostituire "codice_sap" con il nome effettivo della colonna
            Dim itemCode As String = riga.Cells("Itemcode_codice").Value

            togli_saltato_codice(id_lotto, itemCode, Label17.Text, riga.Cells("Docnum_padre").Value)

        Next
        RIEMPI_datagridview_CODICI_PRELIEVO(id_lotto)
        MsgBox("Codici saltati selezionati riattivati con successo")

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        estrai_datagridview_in_excel(DataGridView4)
    End Sub

    Sub estrai_datagridview_in_excel(par_datagriview As DataGridView)
        ' Creare un'applicazione Excel
        Dim excelApp As New Excel.Application
        excelApp.Visible = True ' Mostrare Excel all'utente

        ' Creare un nuovo foglio di lavoro
        Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add
        Dim excelWorksheet As Excel.Worksheet = CType(excelWorkbook.Worksheets(1), Excel.Worksheet)

        ' Aggiungere intestazioni alla prima riga del foglio di lavoro (facoltativo)
        For col As Integer = 1 To par_datagriview.Columns.Count
            excelWorksheet.Cells(1, col) = par_datagriview.Columns(col - 1).HeaderText
        Next

        ' Formattare la colonna "codice" come testo
        Dim codiceColumnIndex As Integer = -1
        For col As Integer = 0 To par_datagriview.Columns.Count - 1
            If par_datagriview.Columns(col).Name.ToLower() = "codice" Then
                codiceColumnIndex = col + 1 ' Excel columns are 1-based
                Exit For
            End If
        Next
        If codiceColumnIndex > 0 Then
            excelWorksheet.Columns(codiceColumnIndex).NumberFormat = "@"
        End If

        ' Aggiungere dati dal DataGridView al foglio di lavoro
        For row As Integer = 0 To par_datagriview.Rows.Count - 1
            For col As Integer = 0 To par_datagriview.Columns.Count - 1
                excelWorksheet.Cells(row + 2, col + 1) = par_datagriview.Rows(row).Cells(col).Value
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

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        If ComboBox1.Text = "" Then
            MsgBox("Selezionare un utente")
        Else
            Me.Hide()
            Funzioni_utili.prelievo_a_ferretto_lotto_di_prelievo(Txt_DocNum.Text, ComboBox1.Text, ComboBox2.Text)
            Beep()
            MsgBox("Prelievo a FERRETTO lanciato con successo")
        End If

    End Sub

    Private Sub Form_lotto_di_prelievo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Funzioni_utili.Inserimento_postazioni(ComboBox1)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class