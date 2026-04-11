Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Form_Premontaggio
    Public ODP As String
    Public stato_lavorazione As String
    Public Elenco_dipendenti_MES(1000) As String
    Public Elenco_responsabili_Montaggio(1000) As String
    Public Elenco_responsabili_collaudo(1000) As String
    Public Elenco_risorse(1000) As String
    Public riga As Integer
    Public check_dipendente As String
    Public dt = New DataTable()
    Public id_ticket As String

    Public GRUPPI_PREASSEBLAGGIO_TOT As Integer = 0
    Public GRUPPI_MONTAGGIO_TOT As Integer = 0

    Public GRUPPI_PREASSEBLAGGIO_MANCANTI As Integer = 0
    Public GRUPPI_MONTAGGIO_MANCANTI As Integer = 0

    Public GRUPPI_PREASSEBLAGGIO_COMPLETATI As Integer = 0
    Public GRUPPI_MONTAGGIO_COMPLETATI As Integer = 0

    Private Sub DataGridView_ODP_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellClick

        If e.RowIndex >= 0 Then
            riga = e.RowIndex
            ODP = DataGridView_ODP.Rows(e.RowIndex).Cells(2).Value
            id_ticket = DataGridView_ODP.Rows(e.RowIndex).Cells(15).Value
            filtra_news()

        End If
        If e.ColumnIndex > 0 And e.RowIndex >= 0 Then
            If id_ticket > 0 Then
                Button12.Visible = True

            Else
                Button12.Visible = False
            End If

            If File.Exists(Homepage.percorso_DWF & DataGridView_ODP.Rows(riga).Cells(5).Value & ".iam.dwf") Then
                Button7.BackColor = Color.Lime
            Else
                Button7.BackColor = Color.Red
            End If

            If File.Exists(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_ODP.Rows(riga).Cells(5).Value & ".PDF") Then
                Button8.BackColor = Color.Lime
            Else
                Button8.BackColor = Color.Red
            End If

        End If
    End Sub

    Private Sub DataGridView_ODP_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_ODP.CellFormatting

        If DataGridView_ODP.Rows(e.RowIndex).Cells(14).Value > 0 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(14).Style.BackColor = Color.Aqua
        Else
            DataGridView_ODP.Rows(e.RowIndex).Cells(14).Value = Nothing
        End If

        If DataGridView_ODP.Rows(e.RowIndex).Cells(10).Value = 100 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(10).Style.BackColor = Color.Green
        ElseIf DataGridView_ODP.Rows(e.RowIndex).Cells(10).Value < 100 And DataGridView_ODP.Rows(e.RowIndex).Cells(10).Value > 90 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(10).Style.BackColor = Color.Yellow
        Else
            DataGridView_ODP.Rows(e.RowIndex).Cells(10).Style.BackColor = Color.Red
        End If

        If DataGridView_ODP.Rows(e.RowIndex).Cells(15).Value > 0 Then
            DataGridView_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightSlateGray
        Else
            If Not DataGridView_ODP.Rows(e.RowIndex).Cells(8).Value Is System.DBNull.Value Then
                If DataGridView_ODP.Rows(e.RowIndex).Cells(8).Value = "Completato" Then
                    DataGridView_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
                ElseIf DataGridView_ODP.Rows(e.RowIndex).Cells(8).Value = "In_esecuzione" Then
                    DataGridView_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Khaki
                Else
                    DataGridView_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
                End If
            End If
        End If


    End Sub

    Sub riempi_ODP()
        ODP_Form.DataGridView_ODP.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        If Dashboard_MU_New.MU = 0 Then


            CMD_SAP_2.CommandText = "Select t30.linenum, t30.visorder, t30.articolo as 'Articolo', t30.[Desc articolo] as 'Desc articolo', t30.Disegno as 'Disegno', t30.Quantita as 'Quantita', t30.Trasferito as 'Trasferito', t30.[Da trasferire] as 'Da trasferire' ,t30.azione as 'Azione' , t31.docnum as 'ODP', t31.[U_PRG_AZS_Commessa] as 'Commessa', t31.U_produzione as 'Reparto', CASE WHEN SUBSTRING(t31.U_produzione,1,3)='INT' THEN T31.[U_Data_cons_MES] ELSE t31.duedate END as 'Cons ODP', t30.OA as'OA',t30.Fornitore as 'Fornitore', t30.[Cons OA] as 'Cons OA'
from
(
Select t25.linenum, t25.visorder,  t25.articolo, t25.[Desc articolo], t25.Disegno, t25.Quantita, t25.Trasferito, t25.[Da trasferire],t25.azione, t25.[Cons ODP], min(t25.OA) as 'OA' , t25.Fornitore, t25.[Cons OA]
from
(
Select t20.linenum, t20.visorder,  t20.articolo, t20.[Desc articolo], t20.Disegno, t20.Quantita, t20.Trasferito, t20.[Da trasferire],t20.azione, min(case when t20.[Cons ODP] is null then '' else cast(t20.[Cons ODP] as date) end)  as 'Cons ODP', t22.docnum as 'OA' , t22.cardname as 'Fornitore', t21.shipdate as 'Cons OA'
from
(
Select t10.linenum, t10.visorder,t10.articolo, t10.[Desc articolo], t10.Disegno, t10.Quantita, t10.Trasferito, t10.[Da trasferire],t10.azione, t10.ODP, t10.[Cons ODP], t10.Commessa, t10.Reparto, min(t10.[Cons OA]) as 'Cons OA'
from
(

Select t100.linenum, t100.visorder, T100.Articolo, t100.[Desc articolo] , t100.Disegno, t100.Quantita,t100.Trasferito, t100.[Da trasferire], 
case when t100.[Da trasferire]=0 then 'OK' when (t100.giacenza>=t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 )  then 'Trasferibile/Da ordinare' when (t100.giacenza>=t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)>=0 )  then 'Trasferibile' when t100.giacenza<t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 and sum(t106.onorder)>=t100.[Da trasferire] then 'IN APPROV/DA ORDINARE' when t100.[Da trasferire]=0 then 'OK' when (t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)>=0 and t100.giacenza<t100.[Da trasferire]) then 'IN APPROV'   when sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 then 'Da ordinare' end as 'Azione', case when t100.[Da trasferire]=0 then '' else t102.docnum end as 'ODP', cast(case when t100.[Da trasferire]=0 then '' else cast(T102.[DueDate] as varchar) end as VARCHAR)as 'Cons ODP' , case when t100.[Da trasferire]=0 then '' else t102.U_PRG_AZS_commessa end as 'Commessa' ,case when t100.[Da trasferire]=0 then '' else  t102.U_produzione end as 'Reparto', case when t100.[Da trasferire]=0 then '' else t107.docnum  end  as 'OA', 
case when t100.[Da trasferire]=0 then '' else t107.cardname end as 'Fornitore', cast(case when t100.[Da trasferire]=0 then '' else cast(t103.[ShipDate] as varchar)  end as varchar) as 'Cons OA'


from
(
SELECT  t1.linenum, t1.visorder, T9.[ITEMCODE] as 'Articolo', t9.itemname as 'Desc articolo' , t9.u_disegno as 'Disegno', t1.plannedqty as 'Quantita',case when t1.U_prg_wip_qtaspedita is null then 0 else t1.U_prg_wip_qtaspedita end as 'Trasferito', t1.u_prg_wip_qtadatrasf as 'Da trasferire', sum (t20.onhand) as 'giacenza', t1.docentry

from wor1 t1 inner join owor t0 on t0.docentry=t1.docentry
inner join oitm t9 on t9.itemcode=t1.itemcode
inner join oitw t20 on t20.itemcode=t1.itemcode

WHERE t0.docnum='" & ODP & "' and t1.itemtype=4 and (substring(T9.[ITEMCODE],1,1)='0' or substring(T9.[ITEMCODE],1,1)='C' or substring(T9.[ITEMCODE],1,1)='D') and (t20.whscode='01' or t20.whscode='03' or t20.whscode='SCA' or t20.whscode='FERRETTO' or t20.whscode='MUT')

group by 
t1.linenum, t1.visorder, T9.[ITEMCODE] , t9.itemname  , t9.u_disegno , t1.plannedqty, t1.U_prg_wip_qtaspedita , t1.u_prg_wip_qtadatrasf, t1.docentry
)
as t100 left join wor1 t101 on t101.itemcode=t100.articolo and t101.docentry=t100.docentry and t100.linenum=t101.linenum
left join owor t102 on t101.itemcode=t102.itemcode and (T102.Status ='P' or T102.Status ='R' )
left join por1 t103 on t103.itemcode=t101.itemcode and t103.opencreqty >0
LEFT OUTER JOIN ITT1 T104 on T101.itemCode = T104.Father
left join oitw t105 on t105.itemcode=t104.code and t105.[WhsCode]='01'
left join oitw t106 on t106.itemcode=t101.itemcode
left join opor t107 on t107.docentry=t103.docentry

group by
 T100.[articolo], t100.trasferito, T100.[DESC articolo], t100.linenum, t100.visorder, t100.quantita,  t100.disegno, t100.giacenza,t100.[da trasferire], t102.docnum, T102.[DueDate],t102.U_PRG_AZS_commessa,t102.U_produzione,t107.docnum,t107.cardname,t103.[ShipDate]
)
as t10
group by t10.linenum, t10.visorder,t10.articolo, t10.[Desc articolo], t10.[Desc articolo], t10.Disegno, t10.Quantita, t10.Trasferito, t10.[Da trasferire],t10.azione, t10.ODP, t10.[Cons ODP], t10.Commessa, t10.Reparto
) 
as t20
left join por1 t21 on t21.itemcode=t20.articolo and t21.shipdate=t20.[Cons OA] and t21.opencreqty >0
left join opor t22 on t22.docentry=t21.docentry
group by
t20.linenum, t20.visorder, t20.articolo, t20.[Desc articolo], t20.Disegno, t20.Quantita, t20.Trasferito, t20.[Da trasferire],t20.azione,   t22.docnum ,t22.cardname, t21.shipdate
)
as t25
group by 
t25.linenum, t25.visorder,  t25.articolo, t25.[Desc articolo], t25.Disegno, t25.Quantita, t25.Trasferito, t25.[Da trasferire],t25.azione, t25.[Cons ODP] , t25.Fornitore, t25.[Cons OA]

)
as t30
left join owor t31 on t31.itemcode=t30.articolo and (T31.Status <> N'L' )  AND  (T31.Status <> N'C' ) and T31.[DueDate]=t30.[Cons ODP]
order by t30.linenum"


        Else

            CMD_SAP_2.CommandText = "Select t30.linenum, t30.visorder, t30.articolo as 'Articolo', t30.[Desc articolo] as 'Desc articolo', t30.Disegno as 'Disegno', t30.Quantita as 'Quantita', t30.Trasferito as 'Trasferito', t30.[Da trasferire] as 'Da trasferire' ,t30.azione as 'Azione' , t31.docnum as 'ODP', t31.[U_PRG_AZS_Commessa] as 'Commessa', t31.U_produzione as 'Reparto', CASE WHEN SUBSTRING(t31.U_produzione,1,3)='INT' THEN T31.[U_Data_cons_MES] ELSE t31.duedate END as 'Cons ODP', t30.OA as'OA',t30.Fornitore as 'Fornitore', t30.[Cons OA] as 'Cons OA',t30.u_stato_lavorazione
from
(
Select t20.linenum, t20.visorder, t20.articolo, t20.[Desc articolo], t20.Disegno, t20.Quantita, t20.Trasferito, t20.[Da trasferire],t20.azione, min(case when t20.[Cons ODP] is null then '' else cast(t20.[Cons ODP] as date) end) as 'Cons ODP', t22.docnum as 'OA' , t22.cardname as 'Fornitore', t21.shipdate as 'Cons OA',t20.u_stato_lavorazione
from
(
Select t10.linenum, t10.visorder, t10.articolo, t10.[Desc articolo], t10.Disegno, t10.Quantita, t10.Trasferito, t10.[Da trasferire],t10.azione, t10.ODP, t10.[Cons ODP], t10.Commessa, t10.Reparto, min(t10.[Cons OA]) as 'Cons OA',t10.u_stato_lavorazione
from
(
Select t100.linenum, t100.visorder, T100.Articolo, t100.[Desc articolo] , t100.Disegno, t100.Quantita,t100.Trasferito, t100.[Da trasferire], 
case when t100.[Da trasferire]=0 then 'OK' when (t100.giacenza>=t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 )  then 'Trasferibile/Da ordinare' when (t100.giacenza>=t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)>=0 )  then 'Trasferibile' when t100.giacenza<t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 and sum(t106.onorder)>=t100.[Da trasferire] then 'IN APPROV/DA ORDINARE' when t100.[Da trasferire]=0 then 'OK' when (t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)>=0 and t100.giacenza<t100.[Da trasferire]) then 'IN APPROV'   when sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 then 'Da ordinare' end as 'Azione', case when t100.[Da trasferire]=0 then '' else t102.docnum end as 'ODP', cast(case when t100.[Da trasferire]=0 then '' else cast(T102.[DueDate] as varchar) end as VARCHAR)as 'Cons ODP' , case when t100.[Da trasferire]=0 then '' else t102.U_PRG_AZS_commessa end as 'Commessa' ,case when t100.[Da trasferire]=0 then '' else  t102.U_produzione end as 'Reparto', case when t100.[Da trasferire]=0 then '' else t107.docnum  end  as 'OA', 
case when t100.[Da trasferire]=0 then '' else t107.cardname end as 'Fornitore', cast(case when t100.[Da trasferire]=0 then '' else cast(t103.[ShipDate] as varchar)  end as varchar) as 'Cons OA',t100.u_stato_lavorazione

from
(
SELECT  t1.linenum, t1.visorder,  T9.[ITEMCODE] as 'Articolo', case when t1.itemtype = -18 then CAST (t1.linetext as varchar) else t9.itemname end as 'Desc articolo' , t9.u_disegno as 'Disegno', t1.plannedqty as 'Quantita',case when t1.U_prg_wip_qtaspedita is null then 0 else t1.U_prg_wip_qtaspedita end as 'Trasferito', t1.u_prg_wip_qtadatrasf as 'Da trasferire', sum (t20.onhand) as 'giacenza', t1.docentry, t1.u_stato_lavorazione

from wor1 t1 inner join owor t0 on t0.docentry=t1.docentry
left join oitm t9 on t9.itemcode=t1.itemcode
left join oitw t20 on t20.itemcode=t1.itemcode

WHERE t0.docnum=" & ODP & " and  (t20.whscode='01' or t20.whscode='03' or t20.whscode='SCA' or t20.whscode='FERRETTO' or t20.whscode='MUT' OR t1.itemtype=290 OR t1.itemtype=-18)

group by 
t1.linenum, t1.itemtype, CAST (t1.linetext as varchar), t1.visorder, T9.[ITEMCODE] , t9.itemname  , t9.u_disegno , t1.plannedqty, t1.U_prg_wip_qtaspedita , t1.u_prg_wip_qtadatrasf, t1.docentry,t1.u_stato_lavorazione
)
as t100 left join wor1 t101 on t101.itemcode=t100.articolo and t101.docentry=t100.docentry and t100.linenum=t101.linenum
left join owor t102 on t101.itemcode=t102.itemcode and (T102.Status ='P' or T102.Status ='R' )
left join por1 t103 on t103.itemcode=t101.itemcode and t103.opencreqty >0
LEFT OUTER JOIN ITT1 T104 on T101.itemCode = T104.Father
left join oitw t105 on t105.itemcode=t104.code and t105.[WhsCode]='01'
left join oitw t106 on t106.itemcode=t101.itemcode
left join opor t107 on t107.docentry=t103.docentry

group by
 T100.[articolo], t100.trasferito, T100.[DESC articolo], t100.linenum, t100.visorder, t100.quantita,  t100.disegno, t100.giacenza,t100.[da trasferire], t102.docnum, T102.[DueDate],t102.U_PRG_AZS_commessa,t102.U_produzione,t107.docnum,t107.cardname,t103.[ShipDate],t100.u_stato_lavorazione
)
as t10
group by t10.linenum, t10.visorder, t10.articolo, t10.[Desc articolo], t10.[Desc articolo], t10.Disegno, t10.Quantita, t10.Trasferito, t10.[Da trasferire],t10.azione, t10.ODP, t10.[Cons ODP], t10.Commessa, t10.Reparto,t10.u_stato_lavorazione
) 
as t20
left join por1 t21 on t21.itemcode=t20.articolo and t21.shipdate=t20.[Cons OA] and t21.opencreqty >0
left join opor t22 on t22.docentry=t21.docentry
group by
t20.linenum, t20.visorder, t20.articolo, t20.[Desc articolo], t20.Disegno, t20.Quantita, t20.Trasferito, t20.[Da trasferire],t20.azione,   t22.docnum ,t22.cardname, t21.shipdate ,t20.u_stato_lavorazione
)
as t30
left join owor t31 on t31.itemcode=t30.articolo and (T31.Status <> N'L' )  AND  (T31.Status <> N'C' ) and T31.[DueDate]=t30.[Cons ODP]
order by t30.linenum"


        End If
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        If Dashboard_MU_New.mu = 0 Then


            Do While cmd_SAP_reader_2.Read()

                ODP_Form.DataGridView_ODP.Rows.Add(cmd_SAP_reader_2("Articolo"), cmd_SAP_reader_2("Desc articolo"), cmd_SAP_reader_2("Disegno"), Math.Round(cmd_SAP_reader_2("Quantita"), 2), Math.Round(cmd_SAP_reader_2("Trasferito"), 2), Math.Round(cmd_SAP_reader_2("Da trasferire"), 2), cmd_SAP_reader_2("Azione"), cmd_SAP_reader_2("ODP"), cmd_SAP_reader_2("Commessa"), cmd_SAP_reader_2("Reparto"), cmd_SAP_reader_2("Cons ODP"), cmd_SAP_reader_2("OA"), cmd_SAP_reader_2("Fornitore"), cmd_SAP_reader_2("Cons OA"), cmd_SAP_reader_2("linenum"), cmd_SAP_reader_2("visorder"))

            Loop

        Else
            Do While cmd_SAP_reader_2.Read()

                ODP_Form.DataGridView_ODP.Rows.Add(cmd_SAP_reader_2("Articolo"), cmd_SAP_reader_2("Desc articolo"), cmd_SAP_reader_2("Disegno"), cmd_SAP_reader_2("Quantita"), cmd_SAP_reader_2("Trasferito"), cmd_SAP_reader_2("Da trasferire"), cmd_SAP_reader_2("Azione"), cmd_SAP_reader_2("ODP"), cmd_SAP_reader_2("Commessa"), cmd_SAP_reader_2("Reparto"), cmd_SAP_reader_2("Cons ODP"), cmd_SAP_reader_2("OA"), cmd_SAP_reader_2("Fornitore"), cmd_SAP_reader_2("Cons OA"), cmd_SAP_reader_2("linenum"), cmd_SAP_reader_2("visorder"), cmd_SAP_reader_2("u_stato_lavorazione"))

            Loop
        End If

        cmd_SAP_reader_2.Close()
        cnn1.Close()

        ODP_Form.DataGridView_ODP.ClearSelection()

    End Sub

    Sub Cambia_stato_ODP()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand

        CMD_SAP_7.Connection = cnn

        CMD_SAP_7.CommandText = "update t0 set T0.u_stato= '" & stato_lavorazione & "' from owor t0 where t0.docnum= '" & ODP & "'"

        CMD_SAP_7.ExecuteNonQuery()

        cnn.Close()
    End Sub

    Sub formatta_form_7()


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.[DocNum] AS 'docnum', T0.PLANNEDQTY, case when t0.u_stato is null then '' else t0.u_stato end as 'u_stato', T0.STATUS AS 'Stato', t0.u_lavorazione as 'lavorazione', T0.[ItemCode] as 'Itemcode', T1.[ItemName] as 'Itemname', case when T1.[U_Disegno] is null then '' else t1.u_disegno end as 'Disegno', case when T0.[U_PRG_AZS_Commessa] is null then '' else T0.[U_PRG_AZS_Commessa] end as 'Commessa', case when T0.[U_Fase] is null then '' else t0.U_fase end as 'Fase' , T2.[ItmsGrpNam] as 'Gruppo articolo'

FROM OWOR T0 INNER JOIN OITM T1 ON T0.ITEMCODE=T1.ITEMCODE
INNER JOIN OITB T2 ON T1.[ItmsGrpCod] = T2.[ItmsGrpCod] 
WHERE T0.[DocNum] ='" & ODP & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() = True Then

            cmd_SAP_reader_2.Close()
        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()



    End Sub

    Sub formatta_form_8()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.[DocNum] AS 'docnum', T0.[ItemCode] as 'Itemcode', T1.[ItemName] as 'Itemname', case when T1.[U_Disegno] is null then '' else t1.u_disegno end as 'Disegno', T0.[U_PRG_AZS_Commessa] as 'Commessa', case when T0.[U_Fase] is null then '' else t0.U_fase end as 'Fase' , T2.[ItmsGrpNam] as 'Gruppo articolo'

FROM OWOR T0 INNER JOIN OITM T1 ON T0.ITEMCODE=T1.ITEMCODE
INNER JOIN OITB T2 ON T1.[ItmsGrpCod] = T2.[ItmsGrpCod] 
WHERE T0.[DocNum] ='" & DataGridView_ODP.Rows(riga).Cells(2).Value & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            Lavorazioni_MES_Premontaggio.Label_numero_ODP_F.Text = DataGridView_ODP.Rows(riga).Cells(2).Value
            Lavorazioni_MES_Premontaggio.Label_Codice_ODP_F.Text = cmd_SAP_reader_2("Itemcode")
            Lavorazioni_MES_Premontaggio.Label_descrizione.Text = cmd_SAP_reader_2("Itemname")
            Lavorazioni_MES_Premontaggio.Label_commessa_F.Text = cmd_SAP_reader_2("Commessa")
            Lavorazioni_MES_Premontaggio.Label_disegno_F.Text = cmd_SAP_reader_2("Disegno")
            Lavorazioni_MES_Premontaggio.Label_fase_F.Text = cmd_SAP_reader_2("Fase")
            Lavorazioni_MES_Premontaggio.Label_gruppo_articolo_F.Text = cmd_SAP_reader_2("Gruppo articolo")

            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub

    Sub inserimento_dipendenti_MES()

        Lavorazioni_MES_Premontaggio.ComboBox_dipendente.Items.Clear()
        Inventario.ComboBox_DIPENDENTE.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "select *
from
(
select '' as 'Codice dipendenti', '' as 'Nome', '' as 'Nome 2'
union all
SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code where t0.active='Y'
)
as t0
order by t0.nome"
        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Lavorazioni_MES_Premontaggio.Elenco_dipendenti_MES(Indice) = cmd_SAP_reader("Codice dipendenti")
            Lavorazioni_MES_Premontaggio.ComboBox_dipendente.Items.Add(cmd_SAP_reader("Nome"))
            Inventario.ComboBox_DIPENDENTE.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()
    End Sub

    Sub Inserimento_risorse_MES()
        Lavorazioni_MES_Premontaggio.ComboBox_risorse.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select t0.visrescode as 'Risorsa', t0.resname as 'Nome_risorsa'
from orsc t0
where t0.resgrpcod<>5 and t0.restype='L' and t0.validfor='Y' order by t0.resname"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_risorse(Indice) = cmd_SAP_reader("risorsa")
            Lavorazioni_MES_Premontaggio.ComboBox_risorse.Items.Add(cmd_SAP_reader("Nome_risorsa"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Sub Lavorazioni_aperte()
        Lavorazioni_MES_Premontaggio.DataGridView_lavorazioni.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        If Homepage.inter = 0 Then


            CMD_SAP_2.Connection = cnn1
            CMD_SAP_2.CommandText = "SELECT t0.ID as 'ID', t0.docnum as 'ODP', t0.tipo_documento as 'Tipo_documento', t3.itemcode as 'Itemcode',T3.PRODNAME as 'Descrizione', T3.[U_PRG_AZS_Commessa] as 'Commessa', case when t5.itemname is null then '' else t5.itemname end as 'Nome_commessa', case when t5.u_final_customer_name is null then '' else t5.u_final_customer_name end as 'Cliente', case when t4.u_disegno is null then '' else t4.u_disegno end as 'Disegno', t3.plannedqty as 'Quantita', T1.[LastName]+' '+T1.[FirstName] as 'Dipendente', t2.resname as 'Risorsa', t0.data as 'Data', t0.start as 'Start'
FROM MANODOPERA t0 inner join [TIRELLI_40].[dbo].ohem t1 on t1.[empID]=dipendente
inner join orsc t2 on t2.visrescode=t0.risorsa
LEFT JOIN OWOR T3 ON T3.DOCNUM=t0.docnum
left join oitm t4 on t4.itemcode=t3.itemcode
LEFT JOIN OITM T5 ON T5.ITEMCODE=T3.[U_PRG_AZS_Commessa]
where t0.docnum=" & DataGridView_ODP.Rows(riga).Cells(2).Value & " AND (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='')
order by T1.[LastName]+' '+T1.[FirstName], t0.data DESC "

            cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Else
            CMD_SAP_2.Connection = cnn1
            CMD_SAP_2.CommandText = "SELECT t0.ID as 'ID',t0.tipo_documento as 'Tipo_documento', t0.docnum as 'ODP', t3.itemcode as 'Itemcode',T3.PRODNAME as 'Descrizione', T3.[U_PRG_AZS_Commessa] as 'Commessa', case when t5.itemname is null then '' else t5.itemname end as 'Nome_commessa', case when t5.u_final_customer_name is null then '' else t5.u_final_customer_name end as 'Cliente', case when t4.u_disegno is null then '' else t4.u_disegno end as 'Disegno', t3.plannedqty as 'Quantita', T1.[LastName]+' '+T1.[FirstName] as 'Dipendente', t2.resname as 'Risorsa', t0.data as 'Data', t0.start as 'Start'
FROM MANODOPERA t0 inner join [TIRELLI_40].[dbo].ohem t1 on t1.[empID]=dipendente
inner join orsc t2 on t2.visrescode=t0.risorsa
LEFT JOIN OWOR T3 ON T3.DOCNUM=t0.docnum
left join oitm t4 on t4.itemcode=t3.itemcode
LEFT JOIN OITM T5 ON T5.ITEMCODE=T3.[U_PRG_AZS_Commessa]
where  (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='') ANd t0.start <> ''
order by T1.[LastName]+' '+T1.[FirstName], t0.data DESC"

            cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        End If

        Do While cmd_SAP_reader_2.Read()

            Lavorazioni_MES_Premontaggio.DataGridView_lavorazioni.Rows.Add(cmd_SAP_reader_2("ID"), cmd_SAP_reader_2("Tipo_documento"), cmd_SAP_reader_2("ODP"), cmd_SAP_reader_2("Itemcode"), cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("Disegno"), Math.Round(cmd_SAP_reader_2("Quantita"), 2), cmd_SAP_reader_2("Commessa"), cmd_SAP_reader_2("Nome_commessa"), cmd_SAP_reader_2("Cliente"), cmd_SAP_reader_2("Dipendente"), cmd_SAP_reader_2("Risorsa"), cmd_SAP_reader_2("Data"), cmd_SAP_reader_2("Start"))
        Loop


        cnn1.Close()

    End Sub

    Sub Check_Lavorazioni_aperte_dipendente(par_dipendente As String)

        Dim testo As String = ""
        Lavorazioni_MES_Premontaggio.DataGridView_lavorazioni.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.id as 'ID', t0.tipo_documento as 'Documento', t0.docnum as 'ODP', t1.itemcode as 'Itemcode', t2.itemname as 'Descrizione', case when t2.u_disegno is null then '' else t2.u_disegno end as 'Disegno', t1.plannedqty as 'quantita', T1.[U_PRG_AZS_Commessa] as 'Commessa', t4.lastname +' '+t4.firstname as 'Dipendente', t3.resname as 'Risorsa', t0.data as 'Data', t0.start as 'Start', t0.stop, t0.consuntivo
from manodopera t0 left join owor t1 on t0.docnum = t1.docnum
left join oitm t2 on t2.itemcode=t1.itemcode
inner join orsc t3 on t3.visrescode=t0.risorsa
left join [TIRELLI_40].[dbo].OHEM T4 ON T4.[empID]=T0.DIPENDENTE
where (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='') and t0.dipendente='" & par_dipendente & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            testo = "A"
            Lavorazioni_MES_Premontaggio.DataGridView_lavorazioni.Rows.Add(cmd_SAP_reader_2("ID"), cmd_SAP_reader_2("Documento"), cmd_SAP_reader_2("ODP"), cmd_SAP_reader_2("Itemcode"), cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("Disegno"), Math.Round(cmd_SAP_reader_2("Quantita"), 2), cmd_SAP_reader_2("Commessa"), cmd_SAP_reader_2("Dipendente"), cmd_SAP_reader_2("Risorsa"), cmd_SAP_reader_2("Data"), cmd_SAP_reader_2("Start"))
        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        If testo = "A" Then
            MsgBox("risultano le seguenti lavorazioni aperte su questo dipendente, devono essere concluse prima di poterne aprire delle altre")

            'Lavorazioni_MES.Button_start.Hide()
            Lavorazioni_MES_Premontaggio.Button_stop.Show()
            check_dipendente = "STOP"

        End If


        testo = ""

    End Sub

    Sub CHIUDI_lavorazione()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "SELECT t0.ID as 'ID', t0.docnum as 'ODP', t3.itemcode as 'Itemcode',T3.PRODNAME as 'Descrizione', T3.[U_PRG_AZS_Commessa] as 'Commessa', case when t4.u_disegno is null then '' else t4.u_disegno end as 'Disegno', t3.plannedqty as 'Quantita', T1.[firstName]+' '+T1.[lastName] as 'Dipendente', t2.resname as 'Risorsa', t0.data as 'Data', t0.start as 'Start'
FROM MANODOPERA t0 inner join [TIRELLI_40].[dbo].ohem t1 on t1.[empID]=dipendente
inner join orsc t2 on t2.visrescode=t0.risorsa
LEFT JOIN OWOR T3 ON T3.DOCNUM=t0.docnum
left join oitm t4 on t4.itemcode=t3.itemcode
where t0.docnum=" & ODP & " and (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='')"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If Not cmd_SAP_reader_2.Read() = True Then
            stato_lavorazione = ""
            Cambia_stato_ODP()

        End If

        cnn1.Close()

    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Lavorazioni_MES_Premontaggio.Button_start.Show()
        Lavorazioni_MES_Premontaggio.Button_stop.Hide()
        Lavorazioni_MES_Premontaggio.Show()
        Lavorazioni_MES_Premontaggio.DataGridView_lavorazioni.Rows.Clear()
        Lavorazioni_MES_Premontaggio.Owner = Me
        Me.Hide()
        Lavorazioni_MES_Premontaggio.ComboBox_risorse.Text = ""
        Lavorazioni_MES_Premontaggio.ComboBox_dipendente.Text = ""
        formatta_form_8()
        inserimento_dipendenti_MES()
        Inserimento_risorse_MES()
        Lavorazioni_MES_Premontaggio.Button_start.Show()
        Lavorazioni_MES_Premontaggio.GroupBox3.Show()
        Lavorazioni_MES_Premontaggio.GroupBox1.Show()
        Lavorazioni_MES_Premontaggio.GroupBox2.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Homepage.Form_precedente = 6
        Lavorazioni_MES_Premontaggio.Button_start.Hide()
        Lavorazioni_MES_Premontaggio.Button_stop.Show()
        Lavorazioni_MES_Premontaggio.Show()
        Me.Hide()
        Lavorazioni_MES_Premontaggio.ComboBox_risorse.Text = ""
        Lavorazioni_MES_Premontaggio.ComboBox_dipendente.Text = ""
        Homepage.inter = 0
        Lavorazioni_aperte()
        inserimento_dipendenti_MES()
        Inserimento_risorse_MES()
        Lavorazioni_MES_Premontaggio.GroupBox3.Show()
        Lavorazioni_MES_Premontaggio.GroupBox1.Hide()
        Lavorazioni_MES_Premontaggio.GroupBox2.Hide()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If DataGridView_ODP.Rows(riga).Cells(8).Value Is System.DBNull.Value Then
            stato_lavorazione = "Completato"
            Cambia_stato_ODP()
            DataGridView_ODP.Rows(riga).Cells(8).Value = "Completato"
        ElseIf DataGridView_ODP.Rows(riga).Cells(8).Value = "In_esecuzione" Then
            MsgBox("Risulta una lavorazione aperta, chiudere tutte le lavorazioni per poter completare l'operazione")

        Else
            stato_lavorazione = "Completato"
            Cambia_stato_ODP()
            DataGridView_ODP.Rows(riga).Cells(8).Value = "Completato"
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            If DataGridView_ODP.Rows(riga).Cells(8).Value = "Completato" Then
                stato_lavorazione = ""
                Cambia_stato_ODP()
                DataGridView_ODP.Rows(riga).Cells(8).Value = ""
            End If
        Catch ex As Exception

        End Try

    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)
        filtra_data()
    End Sub


    Sub filtra_data()
        Dim i = 0
        Do While i < DataGridView_ODP.RowCount
            Dim parola As String
            parola = UCase(DataGridView_ODP.Rows(i).Cells(6).Value)

            If parola.Contains(UCase(TextBox1.Text)) Then
                DataGridView_ODP.Rows(i).Visible = True

            Else
                DataGridView_ODP.Rows(i).Visible = False

            End If
            i = i + 1
        Loop
    End Sub

    Private Sub Button_commessa_Click(sender As Object, e As EventArgs)
        Mostra.Hide()
        Homepage.mostra_dashboard()
        Mostra.Owner = Me
        Mostra.Show()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)
        'Me.WindowState = FormWindowState.Normal

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub


    Sub filtra()
        Dim i = 0
        Dim parola0 As String
        Dim parola1 As String
        Dim parola2 As String
        Dim parola7 As String


        Do While i < DataGridView_ODP.RowCount
            Try

                parola0 = UCase(DataGridView_ODP.Rows(i).Cells(2).Value)
                parola1 = UCase(DataGridView_ODP.Rows(i).Cells(3).Value)
                parola2 = UCase(DataGridView_ODP.Rows(i).Cells(4).Value)
                parola7 = UCase(DataGridView_ODP.Rows(i).Cells(9).Value)


                If parola0.Contains(UCase(TextBox1.Text)) Then
                    DataGridView_ODP.Rows(i).Visible = True

                    If parola1.Contains(UCase(TextBox2.Text)) Then
                        DataGridView_ODP.Rows(i).Visible = True


                        If parola2.Contains(UCase(TextBox3.Text)) Then
                            DataGridView_ODP.Rows(i).Visible = True


                            If parola7.Contains(UCase(TextBox4.Text)) Then
                                DataGridView_ODP.Rows(i).Visible = True



                            Else
                                DataGridView_ODP.Rows(i).Visible = False

                            End If


                        Else
                            DataGridView_ODP.Rows(i).Visible = False

                        End If




                    Else
                        DataGridView_ODP.Rows(i).Visible = False

                    End If




                Else
                    DataGridView_ODP.Rows(i).Visible = False

                End If



            Catch ex As Exception
                DataGridView_ODP.Rows(i).Visible = False
            End Try
            i = i + 1
        Loop
    End Sub

    Sub filtra_news()

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        filtra()
    End Sub


    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If ODP = Nothing Then
            MsgBox("Scegliere un ordine di produzione")
        Else


            ODP_Form.docnum_odp = ODP
            ODP_Form.Show()
            ODP_Form.inizializza_form(ODP)



        End If

    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click
        If File.Exists(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_ODP.Rows(riga).Cells(3).Value & ".PDF") Then
            Process.Start(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_ODP.Rows(riga).Cells(3).Value & ".PDF")
        Else
            MsgBox("PDF non presente")
        End If
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        If File.Exists(Homepage.percorso_DWF & DataGridView_ODP.Rows(riga).Cells(3).Value & ".iam.dwf") Then
            Process.Start(Homepage.percorso_DWF & DataGridView_ODP.Rows(riga).Cells(3).Value & ".iam.dwf")
        Else
            MsgBox("3D non presente")
        End If
    End Sub



    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        FORM6.elenco_ODP_commessa(Pianificazione.commessa, FORM6.DataGridView_ODP)
        FORM6.news_materiale()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Form_nuovo_ticket.Show()
        Form_nuovo_ticket.Inserimento_dipendenti()
        Me.Hide()
        Form_nuovo_ticket.Owner = Me
        Form_nuovo_ticket.Reparto = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto
        Form_nuovo_ticket.Administrator = 0
        Form_nuovo_ticket.Startup()
        Form_nuovo_ticket.Txt_Commessa.Text = pianificazione.commessa
        Form_nuovo_ticket.ComboBox2.Text = Homepage.business
    End Sub

    Private Sub DataGridView_ODP_CellContextMenuStripChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellContextMenuStripChanged

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Form_Visualizza_Ticket.Txt_Id.Text = id_ticket
        Form_Visualizza_Ticket.Startup()
        Form_Visualizza_Ticket.Show()
        Form_Visualizza_Ticket.Owner = Me
        Me.Hide()
    End Sub




    Sub completamento_gruppi_preassemblaggio_assemblaggio()


        GRUPPI_PREASSEBLAGGIO_TOT = 0
        GRUPPI_MONTAGGIO_TOT = 0

        GRUPPI_PREASSEBLAGGIO_MANCANTI = 0
        GRUPPI_MONTAGGIO_MANCANTI = 0

        GRUPPI_PREASSEBLAGGIO_COMPLETATI = 0
        GRUPPI_MONTAGGIO_COMPLETATI = 0

        Lavorazioni_MES_Premontaggio.DataGridView_lavorazioni.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "SELECT case when t0.u_fase is null then '' else t0.u_fase end as 'u_fase', case when t0.u_stato is null then '' else t0.u_stato end as 'u_stato', count(T0.[DocNum]) as 'Numero_gruppi'
FROM OWOR T0 
WHERE T0.STATUS <>'C' AND t0.u_prg_azs_commessa='" & "M03411" & "' and (t0.u_produzione='ASSEMBL' or t0.u_produzione='EST')
group by t0.u_fase, t0.u_stato"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()


            If cmd_SAP_reader_2("u_fase") = "P01501" Then

                If cmd_SAP_reader_2("u_stato") = "Completato" Then
                    GRUPPI_PREASSEBLAGGIO_COMPLETATI = cmd_SAP_reader_2("Numero_gruppi")
                End If

                GRUPPI_PREASSEBLAGGIO_TOT = GRUPPI_PREASSEBLAGGIO_TOT + cmd_SAP_reader_2("Numero_gruppi")
            End If

            If cmd_SAP_reader_2("u_fase") = "P02001" Then

                If cmd_SAP_reader_2("u_stato") = "Completato" Then
                    GRUPPI_MONTAGGIO_COMPLETATI = cmd_SAP_reader_2("Numero_gruppi")
                End If
                GRUPPI_MONTAGGIO_TOT = GRUPPI_MONTAGGIO_TOT + cmd_SAP_reader_2("Numero_gruppi")
            End If

        Loop

        cmd_SAP_reader_2.Close()
        cnn1.Close()

        Commesse_MES.date_inizio_fine_commesse()


    End Sub

    Private Sub DataGridView_ODP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellContentClick

    End Sub
End Class