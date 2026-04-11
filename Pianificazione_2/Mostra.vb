Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib

Public Class Mostra
    Public premontaggio_consuntivo As Integer
    Public montaggio_consuntivo As Integer
    Private Sub Mostra_Click(sender As Object, e As EventArgs) Handles Me.Click
        Me.Hide()

    End Sub

    Sub lavorazioni_su_commessa()

        DataGridView_lavorazioni.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = cnn1



        CMD_SAP_1.CommandText = " SELECT T10.DIPENDENTE,T10.REPARTO, T10.LAVORAZIONE,SUM(T10.MINUTI)/60  AS 'Ore', min(t10.data) as 'Inizio',max(t10.data) as 'Fine'
FROM
(
SELECT  T1.LASTNAME +' ' + T1.FIRSTNAME AS 'Dipendente', t6.nAME as 'Reparto', T0.RISORSA, t2.resname as 'Lavorazione', case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end as 'Minuti', t0.data
FROM MANODOPERA t0 LEFT JOIN [TIRELLI_40].[dbo].OHEM T1 ON T1.CODE=T0.DIPENDENTE
left join orsc t2 on t2.visrescode=t0.risorsa
left join owor t3 on t3.docnum=t0.docnum and t0.tipo_documento='ODP'
LEFT JOIN OITM t4 ON T4.ITEMCODE=T3.ITEMCODE
left join oitm t5 on t5.itemcode=T3.[U_PRG_AZS_Commessa]
left join [TIRELLI_40].[dbo].oudp t6 on t1.dept=t6.code

where T3.[U_PRG_AZS_Commessa]='" & pianificazione.commessa & "' AND T1.DEPT<>5 and T2.ResGrpCod =1
)
AS T10 inner join orsc t11 on t10.risorsa=t11.visrescode

GROUP BY T10.DIPENDENTE,T10.REPARTO,T10.LAVORAZIONE, T10.RISORSA,t11.u_ordine
order by t11.u_ordine
"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Dim contatore As Integer = 0
        Do While cmd_SAP_reader_1.Read()

            DataGridView_lavorazioni.Rows.Add(cmd_SAP_reader_1("Dipendente"), cmd_SAP_reader_1("Reparto"), cmd_SAP_reader_1("Lavorazione"), cmd_SAP_reader_1("Ore"), cmd_SAP_reader_1("Inizio"), cmd_SAP_reader_1("Fine"))


        Loop
        cnn1.Close()
    End Sub

    Sub lavorazioni_per_reparto()

        Chart_lavorazioni.Series("Ore").Points.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = cnn1



        CMD_SAP_1.CommandText = " SELECT T10.LAVORAZIONE,SUM(T10.MINUTI)/60  AS 'Ore'
FROM
(
SELECT  T1.LASTNAME +' ' + T1.FIRSTNAME AS 'Dipendente', t6.nAME as 'Reparto', T0.RISORSA, CASE WHEN T0.RISORSA = 'R00574' OR T0.RISORSA ='R00573' OR T0.RISORSA='R00568' THEN 'Premontaggio' WHEN T0.RISORSA ='R00525' OR T0.RISORSA ='R00576' OR T0.RISORSA ='R00577' THEN 'Montaggio' ELSE t2.resname END as 'Lavorazione', case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end as 'Minuti', t0.data
FROM MANODOPERA t0 LEFT JOIN [TIRELLI_40].[dbo].OHEM T1 ON T1.CODE=T0.DIPENDENTE
left join orsc t2 on t2.visrescode=t0.risorsa
left join owor t3 on t3.docnum=t0.docnum and t0.tipo_documento='ODP'
LEFT JOIN OITM t4 ON T4.ITEMCODE=T3.ITEMCODE
left join oitm t5 on t5.itemcode=T3.[U_PRG_AZS_Commessa]
left join [TIRELLI_40].[dbo].oudp t6 on t1.dept=t6.code

where T3.[U_PRG_AZS_Commessa]='" & pianificazione.commessa & "' AND T1.DEPT <>5 and t2.ResGrpCod=1
)
AS T10 inner join orsc t11 on t10.risorsa=t11.visrescode

GROUP BY T10.LAVORAZIONE,t11.u_ordine
order by t11.u_ordine DESC
"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader

        Do While cmd_SAP_reader_1.Read()

            Chart_lavorazioni.Series("Ore").Points.AddXY(cmd_SAP_reader_1("Lavorazione"), cmd_SAP_reader_1("Ore"))
            If cmd_SAP_reader_1("Lavorazione") = "Premontaggio" Then

                premontaggio_consuntivo = Math.Round(cmd_SAP_reader_1("Ore"))

            End If

            If cmd_SAP_reader_1("Lavorazione") = "Montaggio" Then

                montaggio_consuntivo = cmd_SAP_reader_1("Ore")
            End If



        Loop
        cnn1.Close()
    End Sub



    Sub SITUAZIONE_MAGAZZINO()
        Chart1.Series("Codici").Points.Clear()
        Chart2.Series("Codici").Points.Clear()
        Chart3.Series("Codici").Points.Clear()
        Dim trasferito_premontaggio As Decimal
        Dim trasferito_montaggio As Decimal
        Dim trasferibile_premontaggio As Decimal
        Dim trasferibile_montaggio As Decimal
        Dim trasferito As Decimal
        Dim trasferibile As Decimal
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
WHERE (t0.status='P' or t0.status='R' ) and T0.[U_PRODUZIONE]='ASSEMBL' and t3.itemtype=4 AND T0.[U_PRG_AZS_Commessa]='" & pianificazione.commessa & "' AND T10.DOCNUM IS NULL and T4.[DfltWH]<>'03'
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
            If cmd_SAP_reader_1("Trasferiti PREM") = 0 And cmd_SAP_reader_1("N PREM") = 0 Then
                trasferito_premontaggio = 100
            Else
                trasferito_premontaggio = Math.Round(cmd_SAP_reader_1("Trasferiti PREM") / cmd_SAP_reader_1("N PREM") * 100, 1)
            End If

            If cmd_SAP_reader_1("Trasferiti MONT") = 0 And cmd_SAP_reader_1("N MONT") = 0 Then
                trasferito_montaggio = 100
            Else
                trasferito_montaggio = Math.Round(cmd_SAP_reader_1("Trasferiti MONT") / cmd_SAP_reader_1("N MONT") * 100, 1)
            End If


            If cmd_SAP_reader_1("Trasferiti PREM") = 0 And cmd_SAP_reader_1("N PREM") = 0 Then
                trasferibile_premontaggio = 0
            Else
                trasferibile_premontaggio = Math.Round((cmd_SAP_reader_1("Trasferiti PREM") + cmd_SAP_reader_1("Trasferibile PREM")) / cmd_SAP_reader_1("N PREM") * 100 - cmd_SAP_reader_1("Trasferiti PREM") / cmd_SAP_reader_1("N PREM") * 100, 1)
            End If

            If cmd_SAP_reader_1("Trasferiti MONT") = 0 And cmd_SAP_reader_1("N MONT") = 0 Then
                trasferibile_montaggio = 0
            Else
                trasferibile_montaggio = Math.Round((cmd_SAP_reader_1("Trasferiti MONT") + cmd_SAP_reader_1("Trasferibile MONT")) / cmd_SAP_reader_1("N MONT") * 100 - cmd_SAP_reader_1("Trasferiti MONT") / cmd_SAP_reader_1("N MONT") * 100, 1)
            End If
            If cmd_SAP_reader_1("Trasferiti") = 0 Then
                trasferito = 0
            Else
                trasferito = Math.Round(cmd_SAP_reader_1("Trasferiti") / cmd_SAP_reader_1("Totale") * 100, 1)
            End If
            If cmd_SAP_reader_1("Trasferiti") + cmd_SAP_reader_1("Trasferibile") = 0 Then
                trasferibile = 0
            Else
                trasferibile = Math.Round((cmd_SAP_reader_1("Trasferiti") + cmd_SAP_reader_1("Trasferibile")) / cmd_SAP_reader_1("Totale") * 100 - cmd_SAP_reader_1("Trasferiti") / cmd_SAP_reader_1("Totale") * 100, 1)
            End If

            Chart2.Series("Codici").Points.AddXY("Trasferito", trasferito_premontaggio)
            Chart2.Series("Codici").Points(0).Color = Color.Lime
            Chart2.Series("Codici").Points.AddXY("Trasferibile", trasferibile_premontaggio)
            Chart2.Series("Codici").Points(1).Color = Color.Yellow
            Chart2.Series("Codici").Points.AddXY("Mancante", 100 - trasferibile_premontaggio - trasferito_premontaggio)
            Chart2.Series("Codici").Points(2).Color = Color.OrangeRed


            Chart1.Series("Codici").Points.AddXY("Trasferito", trasferito_montaggio)
            Chart1.Series("Codici").Points(0).Color = Color.Lime
            Chart1.Series("Codici").Points.AddXY("Trasferibile", trasferibile_montaggio)
            Chart1.Series("Codici").Points(1).Color = Color.Yellow
            Chart1.Series("Codici").Points.AddXY("Mancante", 100 - trasferibile_montaggio - trasferito_montaggio)
            Chart1.Series("Codici").Points(2).Color = Color.OrangeRed

            Chart3.Series("Codici").Points.AddXY("Trasferito", trasferito)
            Chart3.Series("Codici").Points(0).Color = Color.Lime
            Chart3.Series("Codici").Points.AddXY("Trasferibile", trasferibile)
            Chart3.Series("Codici").Points(1).Color = Color.Yellow
            Chart3.Series("Codici").Points.AddXY("Mancante", 100 - trasferibile - trasferito)
            Chart3.Series("Codici").Points(2).Color = Color.OrangeRed

        Loop
        cnn1.Close()

    End Sub



    Sub Lavorazioni_aperte_mostra()
        Dim Cnn1 As New SqlConnection
        DataGridView_lavorazioni_aperte.Rows.Clear()
        cnn1.ConnectionString = Homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "SELECT T3.PRODNAME as 'Descrizione',T1.[LastName]+' '+T1.[FirstName] as 'Dipendente', T5.NAME as 'Reparto', t2.resname as 'Risorsa', t0.start as 'Start'
FROM MANODOPERA t0 inner join [TIRELLI_40].[dbo].ohem t1 on t1.[empID]=dipendente
inner join orsc t2 on t2.visrescode=t0.risorsa
LEFT JOIN OWOR T3 ON T3.DOCNUM=t0.docnum
left join oitm t4 on t4.itemcode=t3.itemcode
left join [TIRELLI_40].[dbo].oudp t5 on t1.dept=t5.code
where  (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='') and T3.[U_PRG_AZS_Commessa] ='" & Pianificazione.commessa & "' AND T1.DEPT<>5
order by t2.U_ordine, T1.[LastName]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()

            DataGridView_lavorazioni_aperte.Rows.Add(cmd_SAP_reader_2("Dipendente"), cmd_SAP_reader_2("Reparto"), cmd_SAP_reader_2("risorsa"), cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("Start"))
        Loop
        cmd_SAP_reader_2.Close()
        cnn1.Close()

    End Sub

    Sub Gantt()
        Dim contatore As Integer = 0
        Chart4.Series("Date").Points.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "SELECT substring(t0.risorsa,1,6) as 'Risorsa', t0.RISORSA_DESC AS 'Nome RIS',MIN(t0.Data_I) AS 'Data_I',max(t0.Data_f) as 'Data_F' 
from [Tirelli_40].[dbo].[PIANIFICAZIONE_OUTPUT] t0 
left join [Tirelli_40].[dbo].[PIANIFICAZIONE_COMMESSA] t1 on t0.commessa=t1.commessa 
LEFT JOIN [TIRELLI_40].[dbo].OHEM T2 ON T2.CODE=T0.dipendente
where SUBSTRING(t0.COMMESSA,1,6)='" & pianificazione.commessa & "' AND T0.LIVELLO=1
group by t0.RISORSA_DESC,t0.risorsa
order by 
MIN(t0.Data_I) DESC "

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()
            Chart4.Series("Date").Points.AddXY(cmd_SAP_reader_2("Nome RIS"), cmd_SAP_reader_2("Data_F"), cmd_SAP_reader_2("Data_I"))
            If cmd_SAP_reader_2("risorsa") = "P00001" Then
                Chart4.Series("Date").Points(contatore).Color = Color.Aquamarine
            ElseIf cmd_SAP_reader_2("risorsa") = "P01001" Then
                Chart4.Series("Date").Points(contatore).Color = Color.MistyRose
            ElseIf cmd_SAP_reader_2("risorsa") = "P01501" Then
                Chart4.Series("Date").Points(contatore).Color = Color.YellowGreen
            ElseIf cmd_SAP_reader_2("risorsa") = "P02001" Then
                Chart4.Series("Date").Points(contatore).Color = Color.LightBlue
            ElseIf cmd_SAP_reader_2("risorsa") = "P03001" Then
                Chart4.Series("Date").Points(contatore).Color = Color.Pink
            ElseIf cmd_SAP_reader_2("risorsa") = "P04001" Then
                Chart4.Series("Date").Points(contatore).Color = Color.Lime
            End If
            contatore = contatore + 1

        Loop
        cmd_SAP_reader_2.Close()
        cnn1.Close()

    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()

    End Sub

    Private Sub GroupBox8_Click(sender As Object, e As EventArgs) Handles GroupBox8.Click
        Analisi_riga_magazzino.TextBox7.Text = pianificazione.commessa
        Analisi_riga_magazzino.Materiale_mancante()
        Analisi_riga_magazzino.Owner = Me
        Analisi_riga_magazzino.Show()
        Me.Hide()
        Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa)
        Analisi_riga_magazzino.Button_commessa.Text = pianificazione.commessa
        Analisi_riga_magazzino.Label_descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Descrizione_commessa
        Analisi_riga_magazzino.Label_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_commessa
        Analisi_riga_magazzino.Label_cliente_finale.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_finale_commessa

    End Sub



    Private Sub Chart2_Click(sender As Object, e As EventArgs) Handles Chart2.Click



        Analisi_riga_magazzino.TextBox7.Text = pianificazione.commessa
        Analisi_riga_magazzino.Materiale_mancante()
        Analisi_riga_magazzino.Owner = Me
        Analisi_riga_magazzino.Show()
        Me.Hide()



        Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa)
        Analisi_riga_magazzino.Button_commessa.Text = pianificazione.commessa
        Analisi_riga_magazzino.Label_descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Descrizione_commessa
        Analisi_riga_magazzino.Label_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_commessa
        Analisi_riga_magazzino.Label_cliente_finale.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_finale_commessa
    End Sub

    Private Sub Chart1_Click(sender As Object, e As EventArgs) Handles Chart1.Click
        Analisi_riga_magazzino.TextBox7.Text = pianificazione.commessa
        Analisi_riga_magazzino.Materiale_mancante()
        Analisi_riga_magazzino.Owner = Me
        Analisi_riga_magazzino.Show()
        Me.Hide()
        Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa)
        Analisi_riga_magazzino.Button_commessa.Text = pianificazione.commessa
        Analisi_riga_magazzino.Label_descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Descrizione_commessa
        Analisi_riga_magazzino.Label_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_commessa
        Analisi_riga_magazzino.Label_cliente_finale.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_finale_commessa
    End Sub

    Private Sub Chart3_Click(sender As Object, e As EventArgs) Handles Chart3.Click
        Analisi_riga_magazzino.TextBox7.Text = pianificazione.commessa
        Analisi_riga_magazzino.Materiale_mancante()
        Analisi_riga_magazzino.Owner = Me
        Analisi_riga_magazzino.Show()
        Me.Hide()
        Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa)
        Analisi_riga_magazzino.Button_commessa.Text = pianificazione.commessa
        Analisi_riga_magazzino.Label_descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Descrizione_commessa
        Analisi_riga_magazzino.Label_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_commessa
        Analisi_riga_magazzino.Label_cliente_finale.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_finale_commessa
    End Sub



    Private Sub Chart5_Click(sender As Object, e As EventArgs)
        Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa)
        FORM6.Button_commessa.Text = pianificazione.commessa

        FORM6.Label_descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Descrizione_commessa
        FORM6.Label_ordine_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).ordine_cliente_commessa
        FORM6.Label_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_commessa
        FORM6.Label_cliente_finale.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_finale_commessa
        FORM6.Label_consegna.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Consegna_commessa


        'FORM6.completamento_gruppi_preassemblaggio_assemblaggio()
        FORM6.elenco_ODP_commessa(Pianificazione.commessa, FORM6.DataGridView_ODP)

        FORM6.riga = Nothing
        FORM6.Owner = Me
        FORM6.Show()
        Me.Hide()

    End Sub

    Private Sub Chart7_Click(sender As Object, e As EventArgs)
        Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa)
        FORM6.Button_commessa.Text = pianificazione.commessa

        FORM6.Label_descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Descrizione_commessa
        FORM6.Label_ordine_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).ordine_cliente_commessa
        FORM6.Label_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_commessa
        FORM6.Label_cliente_finale.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_finale_commessa
        FORM6.Label_consegna.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Consegna_commessa


        'FORM6.completamento_gruppi_preassemblaggio_assemblaggio()
        FORM6.elenco_ODP_commessa(Pianificazione.commessa, FORM6.DataGridView_ODP)

        FORM6.riga = Nothing
        FORM6.Owner = Me
        FORM6.Show()
        Me.Hide()

    End Sub

    Private Sub Chart8_Click(sender As Object, e As EventArgs)
        Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa)
        FORM6.Button_commessa.Text = pianificazione.commessa

        FORM6.Label_descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Descrizione_commessa
        FORM6.Label_ordine_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).ordine_cliente_commessa
        FORM6.Label_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_commessa
        FORM6.Label_cliente_finale.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_finale_commessa
        FORM6.Label_consegna.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Consegna_commessa


        'FORM6.completamento_gruppi_preassemblaggio_assemblaggio()
        FORM6.elenco_ODP_commessa(Pianificazione.commessa, FORM6.DataGridView_ODP)

        FORM6.riga = Nothing
        FORM6.Owner = Me
        FORM6.Show()
        Me.Hide()

    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click

        Materiale_mancante()
        DataGridView_materiale_mancante.Show()
        GroupBox13.Visible = True
        GroupBox17.Visible = True
    End Sub

    Sub Materiale_mancante()

        DataGridView_materiale_mancante.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "Select T30.DOCNUM, T30.ITEMCODE, T30.PRODNAME, t30.articolo as 'Articolo', t30.[Desc articolo] as 'Desc articolo', t30.Disegno as 'Disegno', T30.[ItmsGrpNam], t30.Quantita as 'Quantita', t30.Trasferito as 'Trasferito', t30.[Da trasferire] as 'Da trasferire' ,t30.azione as 'Azione' , t31.docnum as 'ODP', t31.[U_PRG_AZS_Commessa] as 'Commessa', t31.U_produzione as 'Reparto', t31.duedate as 'Cons ODP', t30.OA as'OA',t30.Fornitore as 'Fornitore', t30.[Cons OA] as 'Cons OA'
from
(
Select T20.DOCNUM, T20.ITEMCODE, T20.PRODNAME, t20.linenum, t20.articolo, t20.[Desc articolo], t20.Disegno, T20.[ItmsGrpNam], t20.Quantita, t20.Trasferito, t20.[Da trasferire],t20.azione, min(t20.[Cons ODP]) as 'Cons ODP', t22.docnum as 'OA' , t22.cardname as 'Fornitore', t21.shipdate as 'Cons OA'
from
(
Select T10.DOCNUM, T10.ITEMCODE, T10.PRODNAME, t10.linenum,t10.articolo, t10.[Desc articolo], t10.Disegno, T10.[ItmsGrpNam], t10.Quantita, t10.Trasferito, t10.[Da trasferire],t10.azione, t10.ODP, t10.[Cons ODP], t10.Commessa, t10.Reparto, min(t10.[Cons OA]) as 'Cons OA'
from
(

Select T100.DOCNUM, T100.ITEMCODE, T100.PRODNAME, t100.linenum, T100.Articolo, t100.[Desc articolo] , t100.Disegno, T100.[ItmsGrpNam], t100.Quantita,t100.Trasferito, t100.[Da trasferire], 
case when t100.[Da trasferire]=0 then 'OK' when (t100.giacenza>=t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 )  then 'Trasferibile/Da ordinare' when (t100.giacenza>=t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)>=0 )  then 'Trasferibile' when t100.[Da trasferire]=0 then 'OK' when (t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)>=0 and t100.giacenza<t100.[Da trasferire]) then 'IN APPROV'   when sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 then 'Da ordinare' end as 'Azione', case when t100.[Da trasferire]=0 then '' else t102.docnum end as 'ODP', cast(case when t100.[Da trasferire]=0 then '' else cast(T102.[DueDate] as varchar) end as VARCHAR)as 'Cons ODP' , case when t100.[Da trasferire]=0 then '' else t102.U_PRG_AZS_commessa end as 'Commessa' ,case when t100.[Da trasferire]=0 then '' else  t102.U_produzione end as 'Reparto', case when t100.[Da trasferire]=0 then '' else t107.docnum  end  as 'OA', 
case when t100.[Da trasferire]=0 then '' else t107.cardname end as 'Fornitore', cast(case when t100.[Da trasferire]=0 then '' else cast(t103.[ShipDate] as varchar)  end as varchar) as 'Cons OA'

from
(
SELECT  T0.DOCNUM, T0.ITEMCODE, T0.PRODNAME, t1.linenum, T9.[ITEMCODE] as 'Articolo', t9.itemname as 'Desc articolo' , t9.u_disegno as 'Disegno', T11.[ItmsGrpNam], t1.plannedqty as 'Quantita',case when t1.U_prg_wip_qtaspedita is null then 0 else t1.U_prg_wip_qtaspedita end as 'Trasferito', t1.u_prg_wip_qtadatrasf as 'Da trasferire', sum (t20.onhand) as 'giacenza', t1.docentry

from wor1 t1 inner join owor t0 on t0.docentry=t1.docentry
inner join oitm t9 on t9.itemcode=t1.itemcode
LEFT JOIN OITB T11 ON T9.[ItmsGrpCod] = T11.[ItmsGrpCod]
inner join oitw t20 on t20.itemcode=t1.itemcode
LEFT JOIN OWOR T10 ON T10.ITEMCODE=T1.ITEMCODE AND (T10.STATUS='P' OR T10.STATUS='R') and T10.[U_PRODUZIONE]='ASSEMBL'

WHERE T0.[U_PRG_AZS_Commessa]='" & pianificazione.commessa & "' and t1.itemtype=4 and (substring(T9.[ITEMCODE],1,1)='0' or substring(T9.[ITEMCODE],1,1)='C' or substring(T9.[ITEMCODE],1,1)='D') and (t20.whscode='01' or t20.whscode='03' or t20.whscode='SCA' or t20.whscode='FERRETTO') AND T10.DOCNUM IS NULL

group by 
T0.DOCNUM, T0.ITEMCODE, T0.PRODNAME, t1.linenum, T9.[ITEMCODE] , t9.itemname  , t9.u_disegno , T11.[ItmsGrpNam], t1.plannedqty, t1.U_prg_wip_qtaspedita , t1.u_prg_wip_qtadatrasf, t1.docentry
)
as t100 left join wor1 t101 on t101.itemcode=t100.articolo and t101.docentry=t100.docentry and t100.linenum=t101.linenum
left join owor t102 on t101.itemcode=t102.itemcode and (T102.Status ='P' or T102.Status ='R' )
left join por1 t103 on t103.itemcode=t101.itemcode and t103.opencreqty >0
LEFT OUTER JOIN ITT1 T104 on T101.itemCode = T104.Father
left join oitw t105 on t105.itemcode=t104.code and t105.[WhsCode]='01'
left join oitw t106 on t106.itemcode=t101.itemcode
left join opor t107 on t107.docentry=t103.docentry

group by
 T100.[articolo], t100.trasferito, T100.[DESC articolo], t100.linenum, t100.quantita,  t100.disegno, T100.[ItmsGrpNam], t100.giacenza,t100.[da trasferire], t102.docnum, T102.[DueDate],t102.U_PRG_AZS_commessa,t102.U_produzione,t107.docnum,t107.cardname,t103.[ShipDate],T100.DOCNUM, T100.ITEMCODE, T100.PRODNAME
)
as t10
group by T10.DOCNUM, T10.ITEMCODE, T10.PRODNAME, t10.linenum,t10.articolo, t10.[Desc articolo], t10.[Desc articolo], t10.Disegno, T10.[ItmsGrpNam], t10.Quantita, t10.Trasferito, t10.[Da trasferire],t10.azione, t10.ODP, t10.[Cons ODP], t10.Commessa, t10.Reparto
) 
as t20
left join por1 t21 on t21.itemcode=t20.articolo and t21.shipdate=t20.[Cons OA] and t21.opencreqty >0
left join opor t22 on t22.docentry=t21.docentry
group by
T20.DOCNUM, T20.ITEMCODE, T20.PRODNAME,t20.linenum, t20.articolo, t20.[Desc articolo], t20.Disegno, T20.[ItmsGrpNam], t20.Quantita, t20.Trasferito, t20.[Da trasferire],t20.azione,   t22.docnum ,t22.cardname, t21.shipdate 
)
as t30
left join owor t31 on t31.itemcode=t30.articolo and (T31.Status <> N'L' )  AND  (T31.Status <> N'C' ) and T31.[DueDate]=t30.[Cons ODP]
where t30.azione <>'OK'
order by t30.DOCNUM"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()

            DataGridView_materiale_mancante.Rows.Add(cmd_SAP_reader_2("Articolo"), cmd_SAP_reader_2("docnum"), cmd_SAP_reader_2("Azione"))
        Loop
        cmd_SAP_reader_2.Close()
        cnn1.Close()

    End Sub

    Sub filtra_CODICE()
        Dim i = 0
        Do While i < DataGridView_materiale_mancante.RowCount
            Dim parola As String
            parola = UCase(DataGridView_materiale_mancante.Rows(i).Cells(0).Value)

            If parola.Contains(UCase(TextBox1.Text)) Then
                DataGridView_materiale_mancante.Rows(i).Visible = True

            Else
                DataGridView_materiale_mancante.Rows(i).Visible = False

            End If
            i = i + 1
        Loop
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        filtra_CODICE()
    End Sub

    Sub filtra_odp()
        Dim i = 0
        Do While i < DataGridView_materiale_mancante.RowCount
            Dim parola_2 As String
            parola_2 = UCase(DataGridView_materiale_mancante.Rows(i).Cells(1).Value)

            If parola_2.Contains(UCase(TextBox2.Text)) Then
                DataGridView_materiale_mancante.Rows(i).Visible = True

            Else
                DataGridView_materiale_mancante.Rows(i).Visible = False

            End If
            i = i + 1
        Loop
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        filtra_odp()
    End Sub



    Sub CHART_AVANZAMENTO()
        Chart6.Series("Completato").Points.Clear()
        Chart6.Series("Mancante").Points.Clear()
        Chart6.Series("Completabile").Points.Clear()
        Chart5.Series("Completato").Points.Clear()
        Chart5.Series("Mancante").Points.Clear()
        Chart5.Series("Completabile").Points.Clear()
        ' Dashboard_pianificazione.tempi_standard_ODP_M_completati()
        'FORM6.completamento_gruppi_preassemblaggio_assemblaggio()


        'Chart6.Series("Completato").Points.AddXY("Preventivo", Dashboard_pianificazione.tempo_montaggio_ODP_M_completati * 8)
        'Chart5.Series("Completato").Points.AddXY("Preventivo", Dashboard_pianificazione.tempo_preass_ODP_M_completati * 8)

        Chart6.Series("Completato").Points.AddXY("Preventivo", FORM6.GRUPPI_MONTAGGIO_COMPLETATI)


        Chart5.Series("Completato").Points.AddXY("Preventivo", FORM6.GRUPPI_PREASSEBLAGGIO_COMPLETATI)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "SELECT T40.CAUSALE, COUNT(T40.[n ODP]) AS 'ODP'
FROM
(
SELECT t30.[N ODP], T30.[CODICE FASE], T30.FASE, CASE WHEN T30.[CODICE FASE] ='P01501' AND T30.STATO='Completato' then 'Premontaggio_completato' when T30.[CODICE FASE] ='P02001' AND T30.STATO='Completato' then 'Montaggio_completato'  WHEN T30.[CODICE FASE] ='P01501' and t30.trasferiti=t30.n then 'Premontaggio_completabile' WHEN T30.[CODICE FASE] ='P02001' and t30.trasferiti=t30.n then 'Montaggio_completabile' WHEN T30.[CODICE FASE] ='P01501' and t30.trasferiti<t30.n then 'Premontaggio_mancante' WHEN T30.[CODICE FASE] ='P02001' and t30.trasferiti<t30.n then 'Montaggio_mancante' ELSE '' end as 'Causale'
FROM
(
SELECT t20.[N ODP] , t20.[Stato ODP], t20.Codice , t20.Descrizione , t20.Disegno , t20.quantita , t20.stato , T20.[CODICE FASE], t20.fase , t20.N ,t20.Trasferiti ,  case when T20.PREM is null then 0 else t20.prem end as 'PREM' , case when T20.MONT is null then 0 else t20.mont end as 'MONT' ,T20.[ASS EL]
FROM
(

Select t10.[N ODP] as 'N ODP', t10.[Stato ODP], t10.Codice as 'Codice', t10.Descrizione as 'Descrizione', t10.Disegno as 'Disegno', t10.quantita as 'Quantita', t10.stato as 'Stato', t10.fase as 'Fase', T10.[Codice fase], t10.N as 'N',t10.Trasferiti as 'Trasferiti',  sum(CASE WHEN T10.[Codice fase] ='P01501' THEN t10.quantita* case when T11.[Code]='R00568' OR T11.CODE='R00525' then t11.quantity else 0 end END)   as 'PREM' , sum(CASE WHEN T10.[Codice fase] ='P02001' THEN t10.quantita* case when T11.[Code]='R00568' OR T11.CODE='R00525' then t11.quantity else 0 end END)   as 'MONT' ,sum(case when t11.code='R00530' then t11.quantity else 0 end) as 'ASS EL'
from
(
SELECT T0.[DocNum] as 'N ODP', t0.status as 'Stato ODP', T0.[ItemCode] as 'Codice', T1.itemname as 'Descrizione', case when T1.[U_Disegno] is null then '' else T1.[U_Disegno] end as 'Disegno', T0.[PlannedQty] as 'Quantita',case when T0.U_stato is null then '' else t0.u_stato end as 'Stato', T0.[U_Fase] as 'Codice fase', T2.[Name] as 'Fase', sum(CASE WHEN substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C' then 1 else 0 end ) as 'N', sum(case when t3.U_prg_wip_qtadatrasf=0 and (substring(t3.itemcode,1,1)='D' or substring(t3.itemcode,1,1)='C')  then 1 else 0 end) as 'trasferiti'
FROM OWOR T0 inner join OITM T1 on t0.itemcode=t1.itemcode
 left JOIN [dbo].[@FASE]  T2 ON T0.[U_Fase] = T2.[Code] 
left join wor1 t3 on t3.docentry=t0.docentry
WHERE T0.[U_PRG_AZS_Commessa] ='" & pianificazione.commessa & "'  and (t0.status='P' or t0.status='R' or t0.status='L') and T0.[U_PRODUZIONE]='ASSEMBL' and t3.itemtype=4
group by
T0.[DocNum] , T0.[ItemCode], T1.[U_Disegno] , T2.[Name], t1.itemname, T0.[PlannedQty], T0.[U_Fase], T0.U_stato,t0.status
)
as t10
left join itt1 t11 on t11.father=t10.codice

group by t10.[N ODP], t10.[Stato ODP], t10.Codice, t10.Descrizione, t10.Disegno, t10.quantita, t10.fase, t10.N,t10.Trasferiti, t10.[Codice Fase],t10.stato 
)
AS T20
)
AS T30
)
AS T40
GROUP BY 
T40.CAUSALE"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()
            If cmd_SAP_reader_2("Causale") = "Premontaggio_completato" Then

                Chart5.Series("Completato").Points.AddXY("Consuntivo", cmd_SAP_reader_2("ODP"))
            End If

            If cmd_SAP_reader_2("Causale") = "Montaggio_completato" Then

                Chart6.Series("Completato").Points.AddXY("Consuntivo", cmd_SAP_reader_2("ODP"))
            End If

            If cmd_SAP_reader_2("Causale") = "Premontaggio_completabile" Then

                Chart5.Series("Completabile").Points.AddXY("Preventivo", cmd_SAP_reader_2("ODP"))
            End If

            If cmd_SAP_reader_2("Causale") = "Montaggio_completabile" Then

                Chart6.Series("Completabile").Points.AddXY("Preventivo", cmd_SAP_reader_2("ODP"))
            End If

            If cmd_SAP_reader_2("Causale") = "Premontaggio_mancante" Then

                Chart5.Series("Mancante").Points.AddXY("Preventivo", cmd_SAP_reader_2("ODP"))
            End If

            If cmd_SAP_reader_2("Causale") = "Montaggio_mancante" Then

                Chart6.Series("Mancante").Points.AddXY("Preventivo", cmd_SAP_reader_2("ODP"))
            End If


            'Chart6.Series("Completabile").Points.AddXY("Preventivo", cmd_SAP_reader_2("MONTAGGIO_COMPLETABILE") / 60) Then
            'Chart6.Series("Mancante").Points.AddXY("Preventivo", cmd_SAP_reader_2("MONTAGGIO_MANCANTE") / 60)
            ' Chart5.Series("Completabile").Points.AddXY("Preventivo", cmd_SAP_reader_2("PREMONTAGGIO_COMPLETABILE") / 60)
            ' Chart5.Series("Mancante").Points.AddXY("Preventivo", cmd_SAP_reader_2("PREMONTAGGIO_MANCANTE") / 60)

        Loop
        cmd_SAP_reader_2.Close()
        cnn1.Close()



    End Sub



    Private Sub Chart6_Click(sender As Object, e As EventArgs) Handles Chart6.Click
        Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa)
        FORM6.Button_commessa.Text = pianificazione.commessa

        FORM6.Label_descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Descrizione_commessa
        FORM6.Label_ordine_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).ordine_cliente_commessa
        FORM6.Label_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_commessa
        FORM6.Label_cliente_finale.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_finale_commessa
        FORM6.Label_consegna.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Consegna_commessa


        'FORM6.completamento_gruppi_preassemblaggio_assemblaggio()
        FORM6.elenco_ODP_commessa(Pianificazione.commessa, FORM6.DataGridView_ODP)

        FORM6.riga = Nothing

        FORM6.Show()

    End Sub

    Sub tickets()
        Chart7.Series("Series1").Points.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1

        CMD_SAP_1.CommandText = " SELECT T1.Descrizione, count(T0.Aperto) as 'Ticket Aperti' from [TIRELLI_40].[DBO].coll_tickets T0 left join [TIRELLI_40].[DBO].COLL_Reparti T1 on T0.Destinatario=T1.ID_Reparto 
where T0.Aperto=1 and T0.Commessa='" & Pianificazione.commessa & "'
Group By T1.Descrizione
"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()

            Chart7.Series("Series1").Points.AddXY(cmd_SAP_reader_1("Descrizione"), cmd_SAP_reader_1("Ticket Aperti"))

        Loop
        cnn1.Close()
    End Sub

    Sub Collaudo()
        Chart8.Series("Collaudato").Points.Clear()
        Chart8.Series("Da Collaudare").Points.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1

        CMD_SAP_1.CommandText = " SELECT count(Commessa) as 'Num Formati', sum(case when Collaudato is NULL then 0 else Collaudato end) as 'Collaudati', count(Commessa)-sum(case when Collaudato is NULL then 0 else Collaudato end) as 'Da collaudare' FROM [TIRELLI_40].[DBO].COLL_Combinazioni where Commessa='" & Pianificazione.commessa & "'"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()

            Chart8.Series("Collaudato").Points.AddXY("Formati", cmd_SAP_reader_1("Collaudati"))
            Chart8.Series("Da Collaudare").Points.AddXY("Da Collaudare", cmd_SAP_reader_1("Da collaudare"))

        Loop
        cnn1.Close()
    End Sub

    Private Sub Chart5_Click_1(sender As Object, e As EventArgs) Handles Chart5.Click
        Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa)
        FORM6.Button_commessa.Text = pianificazione.commessa

        FORM6.Label_descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Descrizione_commessa
        FORM6.Label_ordine_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).ordine_cliente_commessa
        FORM6.Label_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_commessa
        FORM6.Label_cliente_finale.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_finale_commessa
        FORM6.Label_consegna.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Consegna_commessa


        'FORM6.completamento_gruppi_preassemblaggio_assemblaggio()
        FORM6.elenco_ODP_commessa(Pianificazione.commessa, FORM6.DataGridView_ODP)

        FORM6.riga = Nothing

        FORM6.Show()

    End Sub

    Sub leggi_chi_è_responsabile()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = cnn
        CMD_SAP_docentry.CommandText = "SELECT  case when t1.empid is null then '' else T1.[lastName] + ' ' + T1.[firstName] end AS 'Responsabile_montaggio' , case when t2.empid is null then '' else T2.[lastName] + ' ' + T2.[firstName] end AS 'Responsabile_collaudo'
         FROM OITM T0 left join [TIRELLI_40].[dbo].OHEM T1 ON T0.u_RESPONSABILE_montaggio =t1.empid
        left join [TIRELLI_40].[dbo].ohem t2 on t0.u_responsabile_collaudo =T2.empid WHERE T0.ITEMCODE='" & Pianificazione.commessa & "'"

        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader


        If cmd_SAP_docentry_reader.Read() Then
            Label1.Text = cmd_SAP_docentry_reader("Responsabile_montaggio")
            Label2.Text = cmd_SAP_docentry_reader("Responsabile_collaudo")

        End If
        cmd_SAP_docentry_reader.Close()
        cnn.Close()


    End Sub 'Inserisco le risorse nella combo box


End Class



