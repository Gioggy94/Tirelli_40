Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop


Public Class Commesse_magazzino
    Public magazzino As Integer = 0

    Sub Commesse_odp_aperte()
        DataGridView_commesse_odp.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = cnn1

        CMD_SAP_1.CommandText = " SELECT t20.commessa, t20.Descrizione,  T20.Cliente, T20.[Cliente finale], T20.Consegna, SUM(T20.Totale) AS 'Totale', sum(t20.[Trasferiti]) as 'Trasferiti', sum(t20.[da trasferire]) as 'da trasferire', sum(t20.trasferibile) as 'Trasferibile', sum(T20.[N PREM]) as 'N PREM', sum(T20.[N MONT]) as 'N MONT', sum(T20.[Trasferiti PREM]) as 'Trasferiti PREM', sum(T20.[Trasferiti MONT]) as 'Trasferiti MONT', sum(t20.[Trasferibile PREM]) as 'Trasferibile PREM', sum(t20.[Trasferibile MONT]) as 'Trasferibile MONT'
FROM
(
Select T10.N_ODP, t10.commessa, t1.itemname as 'Descrizione',  case when t3.cardname is null then T1.[U_Final_customer_name] else t3.cardname end  as 'Cliente', case when T3.[U_Clientefinale] is null then '' else T3.[U_Clientefinale] end as 'Cliente finale', T3.[DocDueDate]  as 'Consegna', sum(t10.[N]) as 'Totale', sum(t10.[Trasferiti]) as 'Trasferiti', sum(t10.[da trasferire]) as 'da trasferire', sum(t10.mag01)+sum(t10.magfer)+sum(t10.SCA)+sum(t10.mag03)+sum(t10.mut) as 'Trasferibile', T10.[N PREM], T10.[N MONT], T10.[Trasferiti PREM], T10.[Trasferiti MONT], case when t14.[U_Fase]='p01501' then sum(t10.mag01)+sum(t10.magfer)+sum(t10.SCA)+sum(t10.mag03)+sum(t10.mut) else 0 end as 'Trasferibile PREM', case when t14.[U_Fase]='p02001' then sum(t10.mag01)+sum(t10.magfer)+sum(t10.SCA)+sum(t10.mag03)+sum(t10.mut) else 0 end as 'Trasferibile MONT'
from
(

SELECT T0.[U_PRG_AZS_Commessa] AS 'Commessa', T0.[DocNum] as 'N_ODP', t0.status as 'Stato ODP', T0.[ItemCode] as 'Codice', T1.itemname as 'Descrizione', case when T1.[U_Disegno] is null then '' else T1.[U_Disegno] end as 'Disegno', T0.[PlannedQty] as 'Quantita',T0.U_stato as 'Stato', T0.[U_Fase] as 'Codice fase', T2.[Name] as 'Fase', sum(CASE WHEN substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C' then 1 else 0 end ) as 'N', sum(case when t3.U_prg_wip_qtadatrasf=0 and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C')  then 1 else 0 end) as 'trasferiti', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t4.dfltwh='01' and t3.U_prg_wip_qtadatrasf<=t5.onhand and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C') then 1 else 0 end) as 'MAG01', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t3.U_prg_wip_qtadatrasf>T5.ONHAND and t3.U_prg_wip_qtadatrasf<=t6.onhand and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C') then 1 else 0 end) as 'MAGFER', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t4.dfltwh='SCA' and t3.U_prg_wip_qtadatrasf<=t7.onhand and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C') then 1 else 0 end) as 'SCA', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t4.dfltwh='03' and t3.U_prg_wip_qtadatrasf<=t8.onhand and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C') then 1 else 0 end) as 'MAG03', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t4.dfltwh='MUT' and t3.U_prg_wip_qtadatrasf<=t9.onhand and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C') then 1 else 0 end) as 'MUT', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C') THEN 1 ELSE 0 END) AS 'Da trasferire',
case when T0.[U_Fase]='P01501' THEN sum(CASE WHEN substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C' then 1 else 0 end )  ELSE 0 END AS 'N PREM',
case when T0.[U_Fase]='P02001' THEN sum(CASE WHEN substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C' then 1 else 0 end )  ELSE 0 END AS 'N MONT',
case when T0.[U_Fase]='P01501' THEN sum(case when t3.U_prg_wip_qtadatrasf=0 and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C')  then 1 else 0 end)  ELSE 0 END AS 'Trasferiti PREM',
case when T0.[U_Fase]='P02001' THEN sum(case when t3.U_prg_wip_qtadatrasf=0 and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C')  then 1 else 0 end)  ELSE 0 END AS 'Trasferiti MONT'



FROM OWOR T0 inner join OITM T1 on t0.itemcode=t1.itemcode
 left JOIN [dbo].[@FASE]  T2 ON T0.[U_Fase] = T2.[Code] 
left join wor1 t3 on t3.docentry=t0.docentry

INNER JOIN OITM T4 ON T4.itemcode=t3.itemcode
LEFT JOIN OITW T5 ON T5.ITEMCODE=T3.ITEMCODE AND T5.WHSCODE='01'
LEFT JOIN OITW T6 ON T6.ITEMCODE=T3.ITEMCODE AND T6.WHSCODE='ferretto'
LEFT JOIN OITW T7 ON T7.ITEMCODE=T3.ITEMCODE AND T7.WHSCODE='SCA'
LEFT JOIN OITW T8 ON T8.ITEMCODE=T3.ITEMCODE AND T8.WHSCODE='03'
LEFT JOIN OITW T9 ON T9.ITEMCODE=T3.ITEMCODE AND T9.WHSCODE='MUT'
LEFT JOIN OWOR T10 ON T10.ITEMCODE=T3.ITEMCODE AND (T10.STATUS='P' OR T10.STATUS='R') and T10.[U_PRODUZIONE]='ASSEMBL'
WHERE (t0.status='P' or t0.status='R' ) and T0.[U_PRODUZIONE]='ASSEMBL' and t3.itemtype=4 AND T10.DOCNUM IS NULL
group by
T0.[DocNum] , T0.[ItemCode], T1.[U_Disegno] , T2.[Name], t1.itemname, T0.[PlannedQty], T0.[U_Fase], T0.U_stato,t0.status,T0.[U_PRG_AZS_Commessa]

)
as t10 left join oitm t1 on t10.commessa =t1.itemcode
left join rdr1 t2 on t1.itemcode=t2.itemcode and T2.[OpenQty]>0
left join ordr t3 on t3.docentry=t2.docentry and t3.docstatus='O'
LEFT JOIN OWOR T14 on t14.docnum=T10.N_ODP
group by N_ODP,t10.commessa, t1.itemname, t3.cardname,T1.[U_Final_customer_name],T3.[U_Clientefinale],T3.[DocDueDate], T10.[N PREM], T10.[N MONT], T10.[Trasferiti PREM], T10.[Trasferiti MONT],t14.[U_Fase]
)
AS T20
group by t20.commessa, t20.Descrizione,  T20.Cliente, T20.[Cliente finale], T20.Consegna
order by t20.commessa"

        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()
            Try
                DataGridView_commesse_odp.Rows.Add(cmd_SAP_reader_1("Commessa"), cmd_SAP_reader_1("Descrizione"), cmd_SAP_reader_1("Cliente"), cmd_SAP_reader_1("Cliente finale"), Format(cmd_SAP_reader_1("Consegna"), "dd/MM/yy"), cmd_SAP_reader_1("Trasferiti") / cmd_SAP_reader_1("Totale") * 100, (cmd_SAP_reader_1("Trasferiti") + cmd_SAP_reader_1("trasferibile")) / cmd_SAP_reader_1("Totale") * 100, cmd_SAP_reader_1("Trasferiti"), cmd_SAP_reader_1("da trasferire"), cmd_SAP_reader_1("trasferibile"), cmd_SAP_reader_1("Trasferiti PREM") / cmd_SAP_reader_1("N PREM") * 100, (cmd_SAP_reader_1("Trasferiti PREM") + cmd_SAP_reader_1("Trasferibile PREM")) / cmd_SAP_reader_1("N PREM") * 100, cmd_SAP_reader_1("Trasferiti MONT") / cmd_SAP_reader_1("N MONT") * 100, (cmd_SAP_reader_1("Trasferiti MONT") + cmd_SAP_reader_1("Trasferibile MONT")) / cmd_SAP_reader_1("N MONT") * 100)
            Catch ex As Exception
                DataGridView_commesse_odp.Rows.Add(cmd_SAP_reader_1("Commessa"), cmd_SAP_reader_1("Descrizione"), cmd_SAP_reader_1("Cliente"), cmd_SAP_reader_1("Cliente finale"), cmd_SAP_reader_1("Consegna"), cmd_SAP_reader_1("Trasferiti") / cmd_SAP_reader_1("Totale") * 100, (cmd_SAP_reader_1("Trasferiti") + cmd_SAP_reader_1("trasferibile")) / cmd_SAP_reader_1("Totale") * 100, cmd_SAP_reader_1("Trasferiti"), cmd_SAP_reader_1("da trasferire"), cmd_SAP_reader_1("trasferibile"), cmd_SAP_reader_1("Trasferiti PREM") / cmd_SAP_reader_1("N PREM") * 100, (cmd_SAP_reader_1("Trasferiti PREM") + cmd_SAP_reader_1("Trasferibile PREM")) / cmd_SAP_reader_1("N PREM") * 100, cmd_SAP_reader_1("Trasferiti MONT") / cmd_SAP_reader_1("N MONT") * 100)
            End Try
        Loop
        cnn1.Close()

    End Sub



    Private Sub DataGridView_commesse_odp_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_commesse_odp.CellContentClick
        If e.RowIndex >= 0 Then
            Pianificazione.commessa = DataGridView_commesse_odp.Rows(e.RowIndex).Cells(0).Value

            If e.ColumnIndex = 0 Then

                Trasferibili.Show()


                analisi_magazzino()
                elenco_ODP_commessa()

            End If
        End If
    End Sub

    Sub analisi_magazzino()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "SELECT T1.[ItemCode] as 'Commessa', T2.[ItemName] as 'Nome commessa', T0.[docnum] as 'OC', T0.[cardname] as 'cliente', case when t0.U_clientefinale is null then '' else t0.U_clientefinale end as 'Cliente finale',  T1.[ShipDate] as 'Consegna',  t3.stato as 'Stato' 
            FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.[DocEntry] = T1.[DocEntry] INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode]
            left join [Tirelli_40].[dbo].[PIANIFICAZIONE_COMMESSA] t3 on t3.commessa=t1.itemcode
            WHERE t1.itemcode= '" & pianificazione.commessa & "' and T0.DOCSTATUS='o'
            order by T1.[ItemCode]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Trasferibili.Button_commessa.Text = pianificazione.commessa

        If cmd_SAP_reader_2.Read() = True Then


            Trasferibili.Label_descrizione.Text = cmd_SAP_reader_2("Nome commessa")
            Trasferibili.Label_ordine_cliente.Text = cmd_SAP_reader_2("OC")
            Trasferibili.Label_cliente.Text = cmd_SAP_reader_2("Cliente")
            Trasferibili.Label_cliente_finale.Text = cmd_SAP_reader_2("Cliente finale")
            Trasferibili.Label_consegna.Text = cmd_SAP_reader_2("Consegna")

            cmd_SAP_reader_2.Close()
        End If
        cnn1.Close()



        FORM6.riga = Nothing
    End Sub


    Sub elenco_ODP_commessa()

        Trasferibili.DataGridView_ODP.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = cnn1

        CMD_SAP_1.CommandText = " Select t10.[N ODP] as 'N ODP', t10.[Stato ODP], t10.Codice as 'Codice', t10.Descrizione as 'Descrizione', t10.Disegno as 'Disegno', t10.quantita as 'Quantita', t10.stato as 'Stato', t10.fase as 'Fase', t10.N as 'N',t10.Trasferiti as 'Trasferiti', t10.[da trasferire],  t10.mag01,t10.magFER, t10.SCA, t10.MAG03, t10.MUT
from
(
SELECT T0.[DocNum] as 'N ODP', t0.status as 'Stato ODP', T0.[ItemCode] as 'Codice', T1.itemname as 'Descrizione', case when T1.[U_Disegno] is null then '' else T1.[U_Disegno] end as 'Disegno', T0.[PlannedQty] as 'Quantita',T0.U_stato as 'Stato', T0.[U_Fase] as 'Codice fase', T2.[Name] as 'Fase', sum(CASE WHEN substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C' then 1 else 0 end ) as 'N', sum(case when t3.U_prg_wip_qtadatrasf>0 and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C')  then 1 else 0 end) as 'Da trasferire', sum(case when t3.U_prg_wip_qtadatrasf=0 and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C')  then 1 else 0 end) as 'trasferiti', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t4.dfltwh='01' and t3.U_prg_wip_qtadatrasf<=t5.onhand then 1 else 0 end) as 'MAG01', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t3.U_prg_wip_qtadatrasf>T5.ONHAND and t3.U_prg_wip_qtadatrasf<=t6.onhand then 1 else 0 end) as 'MAGFER', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t4.dfltwh='SCA' and t3.U_prg_wip_qtadatrasf<=t7.onhand then 1 else 0 end) as 'SCA', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t4.dfltwh='03' and t3.U_prg_wip_qtadatrasf<=t8.onhand then 1 else 0 end) as 'MAG03', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t4.dfltwh='MUT' and t3.U_prg_wip_qtadatrasf<=t9.onhand then 1 else 0 end) as 'MUT'
FROM OWOR T0 inner join OITM T1 on t0.itemcode=t1.itemcode
 left JOIN [dbo].[@FASE]  T2 ON T0.[U_Fase] = T2.[Code] 
left join wor1 t3 on t3.docentry=t0.docentry
INNER JOIN OITM T4 ON T4.itemcode=t3.itemcode
LEFT JOIN OITW T5 ON T5.ITEMCODE=T3.ITEMCODE AND T5.WHSCODE='01'
LEFT JOIN OITW T6 ON T6.ITEMCODE=T3.ITEMCODE AND T6.WHSCODE='ferretto'
LEFT JOIN OITW T7 ON T7.ITEMCODE=T3.ITEMCODE AND T7.WHSCODE='SCA'
LEFT JOIN OITW T8 ON T8.ITEMCODE=T3.ITEMCODE AND T8.WHSCODE='03'
LEFT JOIN OITW T9 ON T9.ITEMCODE=T3.ITEMCODE AND T9.WHSCODE='MUT'
LEFT JOIN OWOR T10 ON T10.ITEMCODE=T3.ITEMCODE AND (T10.STATUS='P' OR T10.STATUS='R') and T10.[U_PRODUZIONE]='ASSEMBL'
WHERE T0.[U_PRG_AZS_Commessa] ='" & pianificazione.commessa & "'  and (t0.status='P' or t0.status='R') and T0.[U_PRODUZIONE]='ASSEMBL' and t3.itemtype=4 and (substring(t3.itemcode,1,1)='C' or substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='0') AND T10.DOCNUM IS NULL
group by
T0.[DocNum] , T0.[ItemCode], T1.[U_Disegno] , T2.[Name], t1.itemname, T0.[PlannedQty], T0.[U_Fase], T0.U_stato,t0.status
)
as t10"

        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()

            Trasferibili.DataGridView_ODP.Rows.Add(False, cmd_SAP_reader_1("N ODP"), cmd_SAP_reader_1("Codice"), cmd_SAP_reader_1("Descrizione"), cmd_SAP_reader_1("Disegno"), Math.Round(cmd_SAP_reader_1("Quantita")), cmd_SAP_reader_1("Stato ODP"), cmd_SAP_reader_1("Fase"), Math.Round(cmd_SAP_reader_1("trasferiti") / cmd_SAP_reader_1("N") * 100), Math.Round(cmd_SAP_reader_1("da trasferire")), (cmd_SAP_reader_1("MAG01") + cmd_SAP_reader_1("magFER") + cmd_SAP_reader_1("SCA") + cmd_SAP_reader_1("MAG03") + cmd_SAP_reader_1("MUT")) / cmd_SAP_reader_1("da trasferire") * 100, cmd_SAP_reader_1("MAG01"), cmd_SAP_reader_1("magFER"), cmd_SAP_reader_1("SCA"), cmd_SAP_reader_1("MAG03"), cmd_SAP_reader_1("MUT"))

        Loop
        cnn1.Close()
    End Sub

    Private Sub ComboBox_dipendente_SelectedIndexChanged(sender As Object, e As EventArgs)
        DataGridView_commesse_odp.Show()

    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click



        Me.Close()

    End Sub

    Sub filtra_commesse()
        Dim i = 0
        Do While i < DataGridView_commesse_odp.RowCount
            Try
                Dim parola As String


                parola = UCase(DataGridView_commesse_odp.Rows(i).Cells(0).Value)

                If parola.Contains(UCase(TextBox1.Text)) Then
                    DataGridView_commesse_odp.Rows(i).Visible = True

                Else
                    DataGridView_commesse_odp.Rows(i).Visible = False

                End If

            Catch ex As Exception
                DataGridView_commesse_odp.Rows(i).Visible = False
            End Try
            i = i + 1
        Loop
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

        filtra_commesse()

    End Sub

    Private Sub Button_CDS_Click(sender As Object, e As EventArgs) Handles Button_CDS.Click

    End Sub

    Private Sub Button_OC_Click(sender As Object, e As EventArgs) Handles Button_OC.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Commesse_magazzino_ODP.Show()

        Commesse_magazzino_ODP.Commesse_odp_aperte()
        Me.Hide()
    End Sub

    Private Sub DataGridView_commesse_odp_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_commesse_odp.CellClick

    End Sub
End Class