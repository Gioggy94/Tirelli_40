Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Public Class Commesse_magazzino_ODP


    Sub Commesse_odp_aperte()
        DataGridView_commesse_odp.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = cnn1

        CMD_SAP_1.CommandText = " Select t40.commessa,t40.descrizione, t40.cliente, t40.[cliente finale], t40.consegna, t40.n, t40.ok,t40.completabile, t40.N_prem, t40.ok_prem, t40.Completabile_PREM, t40.N_mont, t40.ok_mont, t40.completabile_mont 
from
(
Select T30.Commessa, t31.itemname as 'Descrizione', t33.cardname as 'Cliente', case when t33.u_clientefinale is null then t31.u_final_customer_name end as 'Cliente Finale', t33.docduedate as 'Consegna',
sum( t30.numeri ) as 'N',
sum(case when  stato_materiale ='OK' then t30.numeri else 0 end) as 'OK',
sum(case when  stato_materiale ='Completabile' then t30.numeri else 0 end) as 'Completabile',
sum(case when T30.[Codice fase] ='P01501'  then t30.numeri else 0 end) as 'N_PREM',
sum(case when T30.[Codice fase] ='P01501' and stato_materiale ='OK' then t30.numeri else 0 end) as 'OK_PREM',
sum(case when T30.[Codice fase] ='P01501' and stato_materiale ='Completabile' then t30.numeri else 0 end) as 'Completabile_PREM',
sum(case when T30.[Codice fase] ='P01501' and stato_materiale ='Incompleto' then t30.numeri else 0 end) as 'Incompleto_PREM',
sum(case when T30.[Codice fase] ='P02001'  then t30.numeri else 0 end) as 'N_MONT',
sum(case when T30.[Codice fase] ='P02001' and stato_materiale ='OK' then t30.numeri else 0 end) as 'OK_MONT',
sum(case when T30.[Codice fase] ='P02001' and stato_materiale ='Completabile' then t30.numeri else 0 end) as 'Completabile_MONT',
sum(case when T30.[Codice fase] ='P02001' and stato_materiale ='Incompleto' then t30.numeri else 0 end) as 'Incompleto_MONT',
sum(case when T30.[Codice fase] ='P04001'  then t30.numeri else 0 end) as 'N_COLLAUDO',
sum(case when T30.[Codice fase] ='P04001' and stato_materiale ='OK' then t30.numeri else 0 end) as 'OK_Collaudo',
sum(case when T30.[Codice fase] ='P04001' and stato_materiale ='Completabile' then t30.numeri else 0 end) as 'Completabile_Collaudo',
sum(case when T30.[Codice fase] ='P04001' and stato_materiale ='Incompleto' then t30.numeri else 0 end) as 'Incompleto_Collaudo'
from
(
Select  T20.Commessa, T20.[Codice fase], T20.Fase,   t20.Stato_materiale,count(T20.Commessa) as 'Numeri'
from
(
Select T10.Commessa, T10.[N_ODP], t10.[Stato ODP], t10.[Codice], T10.Descrizione, t10.Disegno, T10.Quantita,T10.Stato, T10.[Codice fase], T10.Fase, t10.[N], t10.trasferiti, t10.MAG01+ t10.[MAGFER]+ t10.SCA+ t10.MAG03+ t10.MUT as 'Trasferibili', t10.[Da trasferire], case when t10.trasferiti = t10.[N] then 'OK' when t10.trasferiti+t10.MAG01+ t10.[MAGFER]+ t10.SCA+ t10.MAG03+ t10.MUT=t10.[N] then 'Completabile' else 'Incompleto' end as 'Stato_materiale'
from
(
SELECT T0.[U_PRG_AZS_Commessa] AS 'Commessa', T0.[DocNum] as 'N_ODP', t0.status as 'Stato ODP', T0.[ItemCode] as 'Codice', T1.itemname as 'Descrizione', case when T1.[U_Disegno] is null then '' else T1.[U_Disegno] end as 'Disegno', T0.[PlannedQty] as 'Quantita',T0.U_stato as 'Stato', T0.[U_Fase] as 'Codice fase', T2.[Name] as 'Fase', sum(CASE WHEN substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C' then 1 else 0 end ) as 'N', sum(case when t3.U_prg_wip_qtadatrasf=0 and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C')  then 1 else 0 end) as 'trasferiti', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t4.dfltwh='01' and t3.U_prg_wip_qtadatrasf<=t5.onhand and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C') then 1 else 0 end) as 'MAG01', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t3.U_prg_wip_qtadatrasf>T5.ONHAND and t3.U_prg_wip_qtadatrasf<=t6.onhand and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C') then 1 else 0 end) as 'MAGFER', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t4.dfltwh='SCA' and t3.U_prg_wip_qtadatrasf<=t7.onhand and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C') then 1 else 0 end) as 'SCA', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t4.dfltwh='03' and t3.U_prg_wip_qtadatrasf<=t8.onhand and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C') then 1 else 0 end) as 'MAG03', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and t4.dfltwh='MUT' and t3.U_prg_wip_qtadatrasf<=t9.onhand and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C') then 1 else 0 end) as 'MUT', SUM(CASE WHEN t3.U_prg_wip_qtadatrasf>0 and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C') THEN 1 ELSE 0 END) AS 'Da trasferire'



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
WHERE (t0.status='P' or t0.status='R' ) and T0.[U_PRODUZIONE]='ASSEMBL' and t3.itemtype=4 AND T10.DOCNUM IS NULL and (substring(t3.itemcode,1,1)='0' or substring(t3.itemcode,1,1)='C' or substring(t3.itemcode,1,1)='D')
group by
T0.[DocNum] , T0.[ItemCode], T1.[U_Disegno] , T2.[Name], t1.itemname, T0.[PlannedQty], T0.[U_Fase], T0.U_stato,t0.status,T0.[U_PRG_AZS_Commessa]
)
as t10
)
as t20
group by 
  T20.[Codice fase], T20.Fase,  t20.Stato_materiale,T20.Commessa
)
as t30
 left join oitm t31 on t30.commessa =t31.itemcode
left join rdr1 t32 on t32.itemcode=t30.commessa and T32.[OpenQty]>0
left join ordr t33 on t33.docentry=t32.docentry and t33.docstatus='O'
group by T30.Commessa,t31.itemname, t33.cardname,  t33.u_clientefinale, t31.u_final_customer_name , t33.docduedate

)
as t40
where t40.n>0
order by t40.commessa"

        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()
            Try
                'DataGridView_commesse_odp.Rows.Add(cmd_SAP_reader_1("Commessa"), cmd_SAP_reader_1("Descrizione"), cmd_SAP_reader_1("Cliente"), cmd_SAP_reader_1("Cliente finale"), Format(cmd_SAP_reader_1("Consegna"), "dd/MM/yy"), cmd_SAP_reader_1("n"), cmd_SAP_reader_1("OK") / cmd_SAP_reader_1("N") * 100, (cmd_SAP_reader_1("OK") + cmd_SAP_reader_1("cOMPLETABILE")) / cmd_SAP_reader_1("N") * 100, cmd_SAP_reader_1("N_prem"), cmd_SAP_reader_1("OK_prem") / cmd_SAP_reader_1("N_prem") * 100, (cmd_SAP_reader_1("OK_prem") + cmd_SAP_reader_1("completabile_prem")) / cmd_SAP_reader_1("N_prem") * 100, cmd_SAP_reader_1("N_mont"), cmd_SAP_reader_1("OK_mont") / cmd_SAP_reader_1("N_mont") * 100, (cmd_SAP_reader_1("OK_mont") + cmd_SAP_reader_1("completabile_mont")) / cmd_SAP_reader_1("N_mont") * 100)
                DataGridView_commesse_odp.Rows.Add(cmd_SAP_reader_1("Commessa"), cmd_SAP_reader_1("Descrizione"), cmd_SAP_reader_1("Cliente"), cmd_SAP_reader_1("Cliente finale"), Format(cmd_SAP_reader_1("Consegna"), "dd/MM/yy"), cmd_SAP_reader_1("n"), cmd_SAP_reader_1("OK"), (cmd_SAP_reader_1("OK") + cmd_SAP_reader_1("cOMPLETABILE")), cmd_SAP_reader_1("N_prem"), cmd_SAP_reader_1("OK_prem"), (cmd_SAP_reader_1("OK_prem") + cmd_SAP_reader_1("completabile_prem")), cmd_SAP_reader_1("N_mont"), cmd_SAP_reader_1("OK_mont"), (cmd_SAP_reader_1("OK_mont") + cmd_SAP_reader_1("completabile_mont")))
            Catch ex As Exception

                DataGridView_commesse_odp.Rows.Add(cmd_SAP_reader_1("Commessa"), cmd_SAP_reader_1("Descrizione"), cmd_SAP_reader_1("Cliente"), cmd_SAP_reader_1("Cliente finale"), cmd_SAP_reader_1("Consegna"), cmd_SAP_reader_1("n"), cmd_SAP_reader_1("OK"), (cmd_SAP_reader_1("OK") + cmd_SAP_reader_1("cOMPLETABILE")), cmd_SAP_reader_1("N_prem"), cmd_SAP_reader_1("OK_prem"), (cmd_SAP_reader_1("OK_prem") + cmd_SAP_reader_1("completabile_prem")), cmd_SAP_reader_1("N_mont"), cmd_SAP_reader_1("OK_mont"), (cmd_SAP_reader_1("OK_mont") + cmd_SAP_reader_1("completabile_mont")))
            End Try
        Loop
        cnn1.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()

    End Sub



    Private Sub DataGridView_commesse_odp_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_commesse_odp.CellClick
        If e.RowIndex >= 0 Then
            pianificazione.commessa = DataGridView_commesse_odp.Rows(e.RowIndex).Cells(0).Value

            If e.ColumnIndex = 0 Then

                Commesse_magazzino.analisi_magazzino()
                Commesse_magazzino.elenco_ODP_commessa()
                Trasferibili.Owner = Me
                Trasferibili.Show()

                Me.Hide()
            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        filtra_commesse()
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

    Private Sub DataGridView_commesse_odp_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_commesse_odp.CellContentClick

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Commesse_magazzino.Commesse_odp_aperte()

        Commesse_magazzino.Show()

    End Sub
End Class