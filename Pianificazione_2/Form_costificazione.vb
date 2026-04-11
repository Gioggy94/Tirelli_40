Imports System.Data.SqlClient

Imports System.IO
Imports Microsoft.Office.Interop

Public Class Form_costificazione
    Public commessa As String
    Public riga As Integer = 1
    Public livello_esplosione As Integer = 0
    Public livello_max As Integer = 2
    Private filtro_itemcode As String
    Private filtro_doc As String
    Private filtro_itemname As String
    Private filtro_n_doc As String

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Txt_DocNum_TextChanged(sender As Object, e As EventArgs) Handles Txt_commessa.TextChanged
        commessa = Txt_commessa.Text
    End Sub

    Sub informazioni_anagrafiche(par_commessa As String)
        Label1.Text = Magazzino.OttieniDettagliAnagrafica(commessa).Descrizione

        Dim par_nome_documento As String = "Fattura"
        Dim par_intestazione_tabella As String = "OINV"
        Dim par_righe_tabella As String = "INV1"

        If Layout_documenti.OttieniDettaglidocumento_Articolo(commessa, par_nome_documento, par_intestazione_tabella, par_righe_tabella).N_DOC = 0 Then
            par_nome_documento = "Ordine Cliente"
            par_intestazione_tabella = "ORDR"
            par_righe_tabella = "RDR1"
        End If
        Button1.Text = Layout_documenti.OttieniDettaglidocumento_Articolo(commessa, par_nome_documento, par_intestazione_tabella, par_righe_tabella).N_DOC

        Label2.Text = Layout_documenti.OttieniDettaglidocumento_Articolo(commessa, par_nome_documento, par_intestazione_tabella, par_righe_tabella).cliente

        Label9.Text = Layout_documenti.OttieniDettaglidocumento_Articolo(commessa, par_nome_documento, par_intestazione_tabella, par_righe_tabella).cliente_Finale
        GroupBox4.Text = par_nome_documento

        Label3.Text = Layout_documenti.OttieniDettaglidocumento_Articolo(commessa, par_nome_documento, par_intestazione_tabella, par_righe_tabella).prezzo.ToString("C")
    End Sub

    Private Sub Cmd_Cerca_Click(sender As Object, e As EventArgs) Handles Cmd_Cerca.Click


        inizializza_form(commessa)
    End Sub

    Sub inizializza_form(par_commessa As String)
        informazioni_anagrafiche(par_commessa)

        Dim tabpage_componenti As TabPage
        tabpage_componenti = TabControl1.SelectedTab

        If tabpage_componenti Is TabPage1 Then
            componenti_datagridview(DataGridView4, par_commessa, Label5, Label4, filtro_doc, filtro_n_doc, filtro_itemcode, filtro_itemname)
        Else

        End If

        Dim tabpage_manodopera As TabPage
        tabpage_manodopera = TabControl2.SelectedTab

        If tabpage_manodopera Is Manodopera_raggruppata Then
            manodopera_raggruppata_datagridview(DataGridView1, par_commessa, Label6, Label7, Label8, Label10)
        ElseIf tabpage_manodopera Is Dettaglio_dipendente Then
            manodopera_raggruppata_per_dipendente_datagridview(DataGridView3, commessa, Label6, Label7, Label8)
        ElseIf tabpage_manodopera Is Tutti_log Then
            tutti_i_log_manodopera_datagridview(DataGridView2, commessa, Label6, Label7, Label8)
        End If

        Dim marginalità As Decimal

        marginalità = Label3.Text - Label4.Text - Label10.Text
        Dim marginalità_perc As Decimal


        Try
            marginalità_perc = marginalità / Label3.Text

            Label12.Text = marginalità_perc.ToString("P")




            Label11.Text = marginalità.ToString("C")
        Catch ex As Exception

        End Try



        Dim costo_tot As Decimal
        costo_tot = -(Label3.Text - Label4.Text - Label10.Text - Label3.Text)
        Label13.Text = costo_tot.ToString("C")
        If marginalità <> 0 Then
            If marginalità_perc > 0.4 Then
                Label11.ForeColor = Color.DarkGreen
                Label12.ForeColor = Color.DarkGreen
            ElseIf marginalità_perc >= 0 Then
                Label11.ForeColor = Color.DarkGoldenrod
                Label12.ForeColor = Color.DarkGoldenrod
            Else
                Label11.ForeColor = Color.Red
                Label12.ForeColor = Color.Red
            End If
        End If



    End Sub

    Private Sub Form_costificazione_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Txt_commessa.Text = commessa
        inizializza_form(commessa)
    End Sub

    Private Sub Cmd_Avanti_Click(sender As Object, e As EventArgs) Handles Cmd_Avanti.Click
        ' Legge il testo dal TextBox
        Dim input As String = Txt_commessa.Text

        ' Inizializza una stringa vuota per contenere i numeri
        Dim numeri As String = String.Empty

        ' Cicla attraverso ciascun carattere nella stringa di input
        For Each ch As Char In input
            ' Controlla se il carattere è un numero
            If Char.IsDigit(ch) Then
                ' Se sì, aggiungilo alla stringa dei numeri
                numeri &= ch
            End If
        Next
        numeri = numeri + 1

        Txt_commessa.Text = input.Substring(0, 2) & numeri
        inizializza_form(commessa)

    End Sub

    Private Sub Cmd_Indietro_Click(sender As Object, e As EventArgs) Handles Cmd_Indietro.Click
        ' Legge il testo dal TextBox
        Dim input As String = Txt_commessa.Text

        ' Inizializza una stringa vuota per contenere i numeri
        Dim numeri As String = String.Empty

        ' Cicla attraverso ciascun carattere nella stringa di input
        For Each ch As Char In input
            ' Controlla se il carattere è un numero
            If Char.IsDigit(ch) Then
                ' Se sì, aggiungilo alla stringa dei numeri
                numeri &= ch
            End If
        Next
        numeri = numeri - 1

        Txt_commessa.Text = input.Substring(0, 2) & numeri
        inizializza_form(commessa)
    End Sub

    Sub componenti_datagridview(par_datagridview As DataGridView, par_commessa As String, par_label_1 As Label, par_label_2 As Label, par_filtro_doc As String, par_filtro_n_doc As String, par_filtro_itemcode As String, par_filtro_itemname As String)

        Dim costo As Decimal = 0
        If par_filtro_doc = "" Then

        Else
            par_filtro_doc = " and t10.doc Like '%%" & par_filtro_doc & "%%' "
        End If
        If par_filtro_n_doc = "" Then

        Else
            par_filtro_n_doc = " and t10.docnum Like '%%" & par_filtro_n_doc & "%%' "
        End If

        If par_filtro_itemcode = "" Then

        Else
            par_filtro_itemcode = " and t10.itemcode Like '%%" & par_filtro_itemcode & "%%' "
        End If
        If par_filtro_itemname = "" Then

        Else
            par_filtro_itemname = " and t10.itemname Like '%%" & par_filtro_itemname & "%%' "
        End If
        par_datagridview.Rows.Clear()
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
select 'ODP' as 'DOC', t0.docnum, T0.PRODNAME , t0.U_PRG_AZS_Commessa
, t1.itemcode, t3.u_disegno, t3.u_codice_brb, t3.itemname, t4.ItmsGrpNam,t3.u_PRG_TIR_materiale,t1.PlannedQty
,t1.u_prg_wip_qtaspedita,
case when t1.U_Prezzolis = 0  OR t1.U_Prezzolis IS NULL then t5.price else t1.u_prezzolis end  as 'Costo_U'
, case when t1.U_Prezzolis is null OR t1.U_Prezzolis =0 then coalesce(t5.price,0) else coalesce(t1.u_prezzolis,0) end*t1.PlannedQty as 'Costo_Tot'
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
left join owor t2 on t2.itemcode=t1.itemcode and t0.U_PRG_AZS_Commessa=t2.U_PRG_AZS_Commessa AND T2.STATUS<>'C' AND (t2.u_produzione<='ASSEMBL' OR t2.u_produzione<='EST')
left join oitm t3 on t3.itemcode=t1.itemcode
left join oitb t4 on t4.ItmsGrpCod=t3.ItmsGrpCod
left join itm1 t5 on t5.itemcode=t1.itemcode

where t1.PlannedQty>=0 and t0.U_PRG_AZS_Commessa ='" & par_commessa & "' AND t0.status<>'C' and t2.docnum is null and t1.itemtype=4 and t5.pricelist=2 AND (t0.u_produzione='ASSEMBL' OR  t0.u_produzione='EST' )


UNION ALL

select 'OC',t1.docnum, T1.COMMENTS, t0.U_PRG_AZS_Commessa, t0.itemcode,t3.u_disegno, t3.u_codice_brb, t3.itemname, t4.ItmsGrpNam
,t3.u_PRG_TIR_materiale,t0.Quantity,t0.U_Trasferito,
case when t0.U_costo is null OR T0.U_COSTO=0 then t5.price else t0.u_costo end as 'Costo U'
,case when t0.u_costo is null OR T0.U_COSTO=0 then t5.price else t0.u_costo end*t0.Quantity as 'Costo_Tot'

from rdr1 t0 inner join ordr t1 on t0.docentry=t1.docentry
left join owor t2 on t2.itemcode=t0.itemcode and t0.U_PRG_AZS_Commessa=t2.U_PRG_AZS_Commessa  AND T2.STATUS<>'C' AND (t2.u_produzione='ASSEMBL' OR t2.u_produzione='EST')  
left join oitm t3 on t3.itemcode=t0.itemcode
left join oitb t4 on t4.ItmsGrpCod=t3.ItmsGrpCod
left join itm1 t5 on t5.itemcode=t0.itemcode
where t0.u_prg_azs_commessa='" & par_commessa & "' AND T1.CANCELED<>'Y' and t2.docnum is null and t0.itemtype=4 and t5.pricelist=2 and substring(t1.u_causcons,1,4)='COMP'
)
as t10
where 0=0 " & par_filtro_itemcode & par_filtro_itemname & par_filtro_doc & par_filtro_n_doc & "
order by t10.Costo_Tot DESC
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(cmd_SAP_reader_2("DOC"), cmd_SAP_reader_2("Docnum"), cmd_SAP_reader_2("PRODNAME"), cmd_SAP_reader_2("U_PRG_AZS_Commessa"), cmd_SAP_reader_2("itemcode"), cmd_SAP_reader_2("u_disegno"), cmd_SAP_reader_2("u_codice_brb"), cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("ItmsGrpNam"), cmd_SAP_reader_2("u_PRG_TIR_materiale"), cmd_SAP_reader_2("PlannedQty"), cmd_SAP_reader_2("u_prg_wip_qtaspedita"), cmd_SAP_reader_2("Costo_U"), cmd_SAP_reader_2("Costo_TOT"))

            costo = costo + cmd_SAP_reader_2("Costo_TOT")
        Loop

        Cnn1.Close()

        par_datagridview.ClearSelection()
        par_label_1.Text = costo.ToString("C")
        par_label_2.Text = costo.ToString("C")

    End Sub

    Sub manodopera_raggruppata_datagridview(par_datagridview As DataGridView, par_commessa As String, par_label1 As Label, par_label2 As Label, par_label3 As Label, par_label4 As Label)

        Dim costo_manodopera As Decimal = 0
        Dim costo_macchina As Decimal = 0
        Dim costo_tot As Decimal

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "
select T0.risorsa, T2.ResName,t2.restype
, SUM( case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end) AS 'Minuti', t3.price
, COALESCE(t3.price*SUM( case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end),0) as 'Costo_TOT'
from manodopera t0 
left join owor t1 on t0.tipo_documento='ODP' and t1.docnum=t0.docnum
INNER JOIN ORSC T2 ON T2.VisResCode=T0.risorsa 
left join itm1 t3 on t3.PriceList=2 and t3.itemcode=t0.risorsa

where (t1.U_PRG_AZS_Commessa='" & par_commessa & "' and t1.u_produzione='ASSEMBL') OR T0.COMMESSA='" & par_commessa & "'
group by T0.risorsa, T2.ResName,t3.price, t2.restype
order by t3.price*SUM( case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end) desc
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(cmd_SAP_reader_2("risorsa"), cmd_SAP_reader_2("resname"), cmd_SAP_reader_2("restype"), cmd_SAP_reader_2("Minuti"), cmd_SAP_reader_2("price"), cmd_SAP_reader_2("Costo_TOT"))
            If cmd_SAP_reader_2("restype") = "L" Then
                costo_manodopera = costo_manodopera + cmd_SAP_reader_2("Costo_TOT")
            ElseIf cmd_SAP_reader_2("restype") = "M" Then
                costo_macchina = costo_macchina + cmd_SAP_reader_2("Costo_TOT")
            End If

        Loop

        Cnn1.Close()

        par_datagridview.ClearSelection()
        costo_tot = costo_manodopera + costo_macchina
        par_label1.Text = costo_manodopera.ToString("C")
        par_label2.Text = costo_macchina.ToString("C")
        par_label3.Text = costo_tot.ToString("C")
        par_label4.Text = costo_tot.ToString("C")

    End Sub

    Sub manodopera_raggruppata_per_dipendente_datagridview(par_datagridview As DataGridView, par_commessa As String, par_label1 As Label, par_label2 As Label, par_label3 As Label)

        Dim costo_manodopera As Decimal = 0
        Dim costo_macchina As Decimal = 0
        Dim costo_tot As Decimal

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select T0.risorsa, T2.ResName,t2.restype,  T4.LASTNAME +' ' + T4.FIRSTNAME AS 'Dipendente', t5.nAME as 'Reparto'
, SUM( case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end) AS 'Minuti', t3.price
, t3.price*SUM( case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end) as 'Costo_TOT'
from manodopera t0 
left join owor t1 on t0.tipo_documento='ODP' and t1.docnum=t0.docnum
INNER JOIN ORSC T2 ON T2.VisResCode=T0.risorsa 
left join itm1 t3 on t3.PriceList=2 and t3.itemcode=t0.risorsa
LEFT JOIN [TIRELLI_40].[dbo].OHEM T4 ON T4.empid=T0.DIPENDENTE
left join [TIRELLI_40].[dbo].oudp t5 on t4.dept=t5.code

where (t1.U_PRG_AZS_Commessa='" & par_commessa & "' and t1.u_produzione='ASSEMBL') OR T0.COMMESSA='" & par_commessa & "'
group by T0.risorsa, T2.ResName,t3.price, t2.restype,T4.LASTNAME +' ' + T4.FIRSTNAME,t5.nAME
order by t3.price*SUM( case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end) desc
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(cmd_SAP_reader_2("risorsa"), cmd_SAP_reader_2("resname"), cmd_SAP_reader_2("restype"), cmd_SAP_reader_2("dipendente"), cmd_SAP_reader_2("reparto"), cmd_SAP_reader_2("Minuti"), cmd_SAP_reader_2("price"), cmd_SAP_reader_2("Costo_TOT"))
            If cmd_SAP_reader_2("restype") = "L" Then
                costo_manodopera = costo_manodopera + cmd_SAP_reader_2("Costo_TOT")
            ElseIf cmd_SAP_reader_2("restype") = "M" Then
                costo_macchina = costo_macchina + cmd_SAP_reader_2("Costo_TOT")
            End If

        Loop

        Cnn1.Close()

        par_datagridview.ClearSelection()
        costo_tot = costo_manodopera + costo_macchina
        par_label1.Text = costo_manodopera.ToString("C")
        par_label2.Text = costo_macchina.ToString("C")
        par_label3.Text = costo_tot.ToString("C")

    End Sub

    Private Sub Tutti_log_Click(sender As Object, e As EventArgs) Handles Tutti_log.Enter


        tutti_i_log_manodopera_datagridview(DataGridView2, commessa, Label6, Label7, Label8)


    End Sub

    Private Sub manodopera_raggruppata_Click(sender As Object, e As EventArgs) Handles Manodopera_raggruppata.Enter


        manodopera_raggruppata_datagridview(DataGridView1, commessa, Label6, Label7, Label8, Label10)


    End Sub

    Private Sub dettaglio_dipendente_Click(sender As Object, e As EventArgs) Handles Dettaglio_dipendente.Enter


        manodopera_raggruppata_per_dipendente_datagridview(DataGridView3, commessa, Label6, Label7, Label8)


    End Sub

    Sub tutti_i_log_manodopera_datagridview(par_datagridview As DataGridView, par_commessa As String, par_label1 As Label, par_label2 As Label, par_label3 As Label)

        Dim costo_manodopera As Decimal = 0
        Dim costo_macchina As Decimal = 0
        Dim costo_tot As Decimal

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT  T0.id, T0.tipo_documento, T0.docnum as 'N documento', T4.ITEMNAME AS 'Prodotto', T3.[U_PRG_AZS_Commessa] as 'commessa',  t5.itemname as 'Nome commessa', t5.u_final_customer_name as 'Cliente', T1.LASTNAME +' ' + T1.FIRSTNAME AS 'Dipendente', t6.nAME as 'Reparto',  T0.RISORSA AS 'Risorsa', t2.restype, t2.resname as 'Lavorazione', T0.data,T0.start,T0.stop,T0.consuntivo, case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end as 'Minuti', T0.combinazione, 

case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end * t7.price as 'Costo_TOT'
FROM MANODOPERA t0 LEFT JOIN [TIRELLI_40].[dbo].OHEM T1 ON T1.empid=T0.DIPENDENTE
inner join orsc t2 on t2.visrescode=t0.risorsa
left join owor t3 on t3.docnum=t0.docnum and t0.tipo_documento='ODP'
LEFT JOIN OITM t4 ON T4.ITEMCODE=T3.ITEMCODE
left join oitm t5 on t5.itemcode=T3.[U_PRG_AZS_Commessa]
left join [TIRELLI_40].[dbo].oudp t6 on t1.dept=t6.code
LEFT JOIN ITM1 T7 ON T7.ITEMCODE=T0.RISORSA AND T7.PRICELIST=2
where (t3.U_PRG_AZS_Commessa='" & par_commessa & "' and t3.u_produzione='ASSEMBL') OR T0.COMMESSA='" & par_commessa & "'

order by t0.id desc, t0.data DESC, T1.LASTNAME +' ' + T1.FIRSTNAME"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(cmd_SAP_reader_2("id"), cmd_SAP_reader_2("tipo_documento"), cmd_SAP_reader_2("N documento"), cmd_SAP_reader_2("commessa"), cmd_SAP_reader_2("Nome commessa"), cmd_SAP_reader_2("Cliente"), cmd_SAP_reader_2("dipendente"), cmd_SAP_reader_2("reparto"), cmd_SAP_reader_2("risorsa"), cmd_SAP_reader_2("lavorazione"), cmd_SAP_reader_2("data"), cmd_SAP_reader_2("start"), cmd_SAP_reader_2("stop"), cmd_SAP_reader_2("consuntivo"), cmd_SAP_reader_2("minuti"), cmd_SAP_reader_2("combinazione"))

            If cmd_SAP_reader_2("restype") = "L" Then
                costo_manodopera = costo_manodopera + cmd_SAP_reader_2("Costo_TOT")
            ElseIf cmd_SAP_reader_2("restype") = "M" Then
                costo_macchina = costo_macchina + cmd_SAP_reader_2("Costo_TOT")
            End If

        Loop

        Cnn1.Close()

        par_datagridview.ClearSelection()
        costo_tot = costo_manodopera + costo_macchina
        par_label1.Text = costo_manodopera.ToString("C")
        par_label2.Text = costo_macchina.ToString("C")
        par_label3.Text = costo_tot.ToString("C")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim par_nome_tabella As String
        Dim par_righe_tabella As String
        Dim par_nome_documento As String
        If GroupBox4.Text = "Ordine Cliente" Then
            par_nome_tabella = "ORDR"
            par_righe_tabella = "RDR1"
            par_nome_documento = "Ordine"

        ElseIf GroupBox4.Text = "Fattura" Then
            par_nome_tabella = "OINV"
            par_righe_tabella = "INV1"
            par_nome_documento = "FATTURA"
        End If
        Form_nuova_offerta.Show()

        Form_nuova_offerta.TextBox10.Text = Button1.Text
        Form_nuova_offerta.tipo_offerta = "Visualizzazione"
        Form_nuova_offerta.inizializzazione_form(Button1.Text, par_nome_tabella, par_righe_tabella, par_nome_documento)
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click


        Compila_Albero(Txt_commessa.Text, "Cost_Tree")


    End Sub

    Public Sub Compila_Albero(par_codice As String, PAR_TIPO_APPOGGIO As String)

        ' Declare Excel application object
        Dim excelApp As Object
        Dim workbook As Object
        Dim worksheet As Object

        ' Create a new instance of Excel
        excelApp = CreateObject("Excel.Application")
        excelApp.Visible = True ' Optional: Make Excel visible

        ' Add a new workbook
        workbook = excelApp.Workbooks.Add

        ' Get the first worksheet
        worksheet = workbook.Worksheets(1)




        ODP_Tree.PULISCI_APPOGGIO(Homepage.ID_SALVATO, PAR_TIPO_APPOGGIO)



        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT T0.[DocNum], T0.[ProdName], T0.[DocNum], T0.[U_PRG_AZS_Commessa],T0.[Status] ,T0.[Plannedqty] 
        FROM OWOR T0 
        LEFT JOIN [TIRELLI_40].[DBO].APPOGGIO T1 ON T0.DOCNUM=T1.VALORE AND T1.TIPO='" & PAR_TIPO_APPOGGIO & "' AND T1.UTENTE=" & Homepage.ID_SALVATO & "

        WHERE T0.[ItemCode] = '" & par_codice & "' AND (T0.[Status]<>'C') and (t0.u_produzione='ASSEMBL' or t0.u_produzione='EST')  AND T1.VALORE IS NULL"

        Reader_Tree = Cmd_Tree.ExecuteReader()
        'Dim r As Integer = 1
        Dim c As Integer = 1
        If Reader_Tree.Read() Then
            ODP_Tree.AGGIUNGI_RECORD_APPOGGIO(Homepage.ID_SALVATO, PAR_TIPO_APPOGGIO, Reader_Tree("DocNum"))
            ' Aggiunta di caselle di testo
            Dim shape As Object
            ' Prima casella di testo
            worksheet.Cells(riga, c).Value = "Percorso"
            c += 1
            worksheet.Cells(riga, c).Value = "livello"
            c += 1
            worksheet.Cells(riga, c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            worksheet.Cells(riga, c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            worksheet.Cells(riga, c).Value = "Nome"
            worksheet.Cells(riga, c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            worksheet.Cells(riga, c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            c += 1
            worksheet.Cells(riga, c).Value = "Q"
            worksheet.Cells(riga, c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            worksheet.Cells(riga, c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            c += 1
            worksheet.Cells(riga, c).Value = "P"
            worksheet.Cells(riga, c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            worksheet.Cells(riga, c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            c += 1
            worksheet.Cells(riga, c).Value = "TOT"
            worksheet.Cells(riga, c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            worksheet.Cells(riga, c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            c = 1
            riga += 1
            worksheet.Cells(riga, c).Value = 0
            worksheet.Cells(riga, c).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            worksheet.Cells(riga, c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            c += 1
            worksheet.Cells(riga, c).Value = 0
            worksheet.Cells(riga, c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            worksheet.Cells(riga, c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            c += 1
            worksheet.Cells(riga, c).Value = par_codice & " - " & Reader_Tree("ProdName")
            worksheet.Cells(riga, c).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            worksheet.Cells(riga, c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            c += 1
            worksheet.Cells(riga, c).Value = Reader_Tree("Plannedqty")
            worksheet.Cells(riga, c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            worksheet.Cells(riga, c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            c += 1

            Dim formula As String

            formula = "=SUM(" & worksheet.Cells(riga, c + 5).Address & ":" & worksheet.Cells(20000, c + 5).Address & ")"

            Try
                worksheet.Cells(riga, c).Formula = formula
            Catch ex As Exception
                worksheet.Cells(riga, c).Formula = 99999999
            End Try
            worksheet.Cells(riga, c).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            worksheet.Cells(riga, c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            c += 1
            worksheet.Cells(riga, c).Value = "=" & worksheet.Cells(riga, c - 2).address & "*" & worksheet.Cells(riga, c - 1).address & ""
            worksheet.Cells(riga, c).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            worksheet.Cells(riga, c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            c += 1
            riga += 1
            Trova_Figli(Reader_Tree("docnum"), par_codice, c, worksheet, PAR_TIPO_APPOGGIO, 1, 4, 0, 1)

        End If

        Cnn_Tree.Close()

        ' Optional: Save the workbook
        ' workbook.SaveAs "C:\path\to\your\folder\NewExcelFile.xlsx"

        ' Optional: Close Excel if you don't need it open
        ' workbook.Close
        ' excelApp.Quit

        ' Release objects
        worksheet = Nothing
        workbook = Nothing
        excelApp = Nothing
    End Sub

    Sub Trova_Figli(ODP As String, par_commessa As String, c As Integer, worksheet As Object, par_tipo_appoggio As String, par_riga_padre As Integer, par_colonna_padre As Integer, par_livello As Integer, par_percorso As String)

        Dim contatore As Integer = 0

        par_livello += 1

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT 	cast(T0.ItemCode as varchar) AS 'Cod. articolo'
        , T1.ItemName AS 'Descrizione articolo'
        , T2.DocNum AS 'Numero documento'
        , T2.ItemCode AS 'Cod. articolo', 
        	T0.baseQty AS 'Base'
        , T0.U_PRG_WIP_QtaDaTrasf as 'Da Trasferire'
        , T0.U_PRG_WIP_QtaSpedita as 'Trasferito'
,t3.price

        FROM  [dbo].[OITM] T1 
        INNER JOIN [dbo].[WOR1] T0 ON T1.ItemCode = T0.ItemCode
        INNER JOIN [dbo].[OWOR] T2 ON T2.DocEntry = T0.DocEntry 
inner join itm1 t3 on t3.itemcode=t0.itemcode and t3.pricelist=2


        WHERE T2.DocNum='" & ODP & "' AND (T2.U_PRODUZIONE='ASSEMBL' or T2.U_PRODUZIONE='EST') 

        ORDER BY T0.visorder,T1.ItemName,T2.DocNum,T2.Status"

        Reader_Tree = Cmd_Tree.ExecuteReader()

        worksheet.Columns(c).NumberFormat = "@"
        Do While Reader_Tree.Read()

            Dim Risultato As Risultato_ODP
            Risultato = Trova_ODP_Appropriato(Reader_Tree("Cod. articolo"), par_commessa, par_tipo_appoggio)
            contatore += 1
            worksheet.Cells(riga, 1).Value = par_percorso & "_" & contatore
            worksheet.Cells(riga, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            worksheet.Cells(riga, 1).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            worksheet.Cells(riga, 2).Value = par_livello
            worksheet.Cells(riga, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            worksheet.Cells(riga, 2).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            worksheet.Cells(riga, c).Value = Reader_Tree("Cod. articolo").ToString & " - " & Reader_Tree("Descrizione articolo")
            worksheet.Cells(riga, c).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            worksheet.Cells(riga, c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            worksheet.Cells(riga, c + 1).Value = Reader_Tree("Base")
            If Risultato.Num_ODP = "*" Then
                worksheet.Cells(riga, c + 2).Value = Reader_Tree("price")
            Else

            End If
            worksheet.Cells(riga, c + 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            worksheet.Cells(riga, c + 1).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            worksheet.Cells(riga, c + 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            worksheet.Cells(riga, c + 2).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            worksheet.Cells(riga, c + 3).Value = "=" & worksheet.Cells(riga, c + 1).address & "*" & worksheet.Cells(riga, c + 2).address & ""
            worksheet.Cells(riga, c + 3).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            worksheet.Cells(riga, c + 3).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            riga += 1

            If Risultato.Num_ODP <> "*" Then

                ODP_Tree.AGGIUNGI_RECORD_APPOGGIO(Homepage.ID_SALVATO, "ODP_TREE", Risultato.Num_ODP)

                Trova_Figli(Risultato.Num_ODP, par_commessa, c + 4, worksheet, par_tipo_appoggio, riga, c, par_livello, par_percorso & "_" & contatore)
                'riga += 1
            End If





        Loop

        Dim formula As String
        'formula = "=SUM(" & worksheet.Cells(riga, c + 5).Address & ":" & worksheet.Cells(riga + conta_Figli_albero(Risultato.Num_ODP, par_commessa, par_tipo_appoggio, 0), c + 5).Address & ")"

        'formula = "=SUM(" & worksheet.Cells(riga, c + 5).Address & ":" & worksheet.Cells(50, c + 5).Address & ")"
        formula = "=SUM(" & worksheet.Cells(par_riga_padre, c + 3).Address & ":" & worksheet.Cells(riga - 1, c + 3).Address & ")"


        Try
            worksheet.Cells(par_riga_padre - 1, c - 2).Formula = formula
        Catch ex As Exception
            worksheet.Cells(par_riga_padre - 1, c - 2).Formula = 99999999
        End Try


        Cnn_Tree.Close()


    End Sub

    Public Function conta_Figli_albero(ODP As String, par_commessa As String, par_tipo_appoggio As String, contatore As Integer)
        Dim figli As Integer = 0
        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT 	cast(T0.ItemCode as varchar) AS 'Cod. articolo'
        , T1.ItemName AS 'Descrizione articolo'
        , T2.DocNum AS 'Numero documento'
        , T2.ItemCode AS 'Cod. articolo', 
        	T0.PlannedQty AS 'Base'
        , T0.U_PRG_WIP_QtaDaTrasf as 'Da Trasferire'
        , T0.U_PRG_WIP_QtaSpedita as 'Trasferito'
,t3.price

        FROM  [dbo].[OITM] T1 
        INNER JOIN [dbo].[WOR1] T0 ON T1.ItemCode = T0.ItemCode
        INNER JOIN [dbo].[OWOR] T2 ON T2.DocEntry = T0.DocEntry 
inner join itm1 t3 on t3.itemcode=t0.itemcode and t3.pricelist=2


        WHERE T2.DocNum='" & ODP & "' AND T2.U_PRODUZIONE='ASSEMBL' 

        ORDER BY T0.visorder,T1.ItemName,T2.DocNum,T2.Status"

        Reader_Tree = Cmd_Tree.ExecuteReader()


        Do While Reader_Tree.Read()
            contatore += 1
            Dim Risultato As Risultato_ODP
            Risultato = Trova_ODP_Appropriato(Reader_Tree("Cod. articolo"), par_commessa, par_tipo_appoggio)


            If Risultato.Num_ODP <> "*" Then

                ODP_Tree.AGGIUNGI_RECORD_APPOGGIO(Homepage.ID_SALVATO, par_tipo_appoggio, Risultato.Num_ODP)

                conta_Figli_albero(Risultato.Num_ODP, par_commessa, par_tipo_appoggio, contatore)
                riga += 1
            End If





        Loop

        Cnn_Tree.Close()
        figli = contatore

        Return contatore

    End Function



    Private Function Trova_ODP_Appropriato(Codice As String, par_commessa As String, par_tipo_appoggio As String) As Risultato_ODP

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Dim Risultato As Risultato_ODP

        Cmd_Tree.CommandText = "SELECT 	T1.DocNum AS 'Numero documento', T0.ItemCode AS 'Cod. articolo',
T1.PlannedQty As 'Quantità', T0.ItemName AS 'Descrizione articolo', T1.U_PRODUZIONE AS 'Reparto', 
	T1.U_PRG_AZS_Commessa,T1.OriginNum, T1.U_UTILIZZ AS 'Rif a Cliente.'
, T1.Status AS 'Stato'
,coalescE(T1.[U_Progressivo_commessa],0) As 'Cassetta'
,coalescE(A.[ID],0) As 'lotto'
,coalesce(t1.u_disegno,'') as 'Disegno'
,count(t3.itemcode) as 'N_articoli'


FROM  [dbo].[OITM] T0 inner join [dbo].[OWOR] T1 on t0.itemcode=t1.itemcode 
left join (select  max(t0.id) as 'ID', t0.docnum 
from [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 
inner join owor t1 on t0.docnum=t1.docnum

where t1.itemcode='" & Codice & "'
group by t0.docnum ) A on A.docnum=t1.docnum
LEFT JOIN [TIRELLI_40].[DBO].APPOGGIO T2 ON T1.DOCNUM=T2.VALORE AND T2.TIPO='" & par_tipo_appoggio & "' AND T2.UTENTE=" & Homepage.ID_SALVATO & "
left join wor1 t3 on t3.docentry=t1.docentry

WHERE   (T1.Status <> N'C' )  AND T0.ItemCode='" & Codice & "' AND T1.U_PRG_AZS_Commessa='" & par_commessa & "' and t1.u_produzione='ASSEMBL' AND T2.VALORE IS NULL

group by T1.DocNum , T0.ItemCode ,
T1.PlannedQty, T0.ItemName, T1.U_PRODUZIONE , 
	T1.U_PRG_AZS_Commessa,T1.OriginNum, T1.U_UTILIZZ,T1.[U_Progressivo_commessa]
, T1.Status
,A.[ID]
,t1.u_disegno"

        Reader_Tree = Cmd_Tree.ExecuteReader()
        If Reader_Tree.Read() Then
            Risultato.Num_ODP = Reader_Tree("Numero documento")
            Risultato.Tipo = "Sul Carrello : "
            Risultato.Num_Cassetta = Reader_Tree("Cassetta")
            Risultato.qta = Reader_Tree("Quantità")
            Risultato.disegno = Reader_Tree("Disegno")
            Risultato.lotto_prelievo = Reader_Tree("lotto")
            Risultato.commessa = Reader_Tree("U_PRG_AZS_Commessa")
            Risultato.stato = Reader_Tree("stato")
            Risultato.num_articoli = Reader_Tree("N_articoli")

            Cnn_Tree.Close()
            Return Risultato
        End If

        Cnn_Tree.Close()
        Cnn_Tree.Open()
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT 	T1.DocNum AS 'Numero documento', T0.ItemCode AS 'Cod. articolo',T1.PlannedQty As 'Quantità', T0.ItemName AS 'Descrizione articolo', T0.CodeBars As 'Disegno', T1.U_PRODUZIONE AS 'Reparto', 
	T1.U_PRG_AZS_Commessa,T1.OriginNum, T1.U_UTILIZZ AS 'Rif a Cliente.', T1.Status AS 'Stato',case when T1.[U_Progressivo_commessa] is null then 0 else  T1.[U_Progressivo_commessa] end As 'Cassetta'

FROM  [dbo].[OITM] T0 inner join  [dbo].[OWOR] T1 on t0.itemcode=t1.itemcode
LEFT JOIN [TIRELLI_40].[DBO].APPOGGIO T2 ON T1.DOCNUM=T2.VALORE AND T2.TIPO='" & par_tipo_appoggio & "' AND T2.UTENTE=" & Homepage.ID_SALVATO & "

WHERE   (T1.Status <> N'C' ) AND T0.ItemCode='" & Codice & "' AND T1.U_PRG_AZS_Commessa='SCORTA' AND t1.u_produzione='ASSEMBL' AND T2.VALORE IS NULL"

        Reader_Tree = Cmd_Tree.ExecuteReader()
        If Reader_Tree.Read() Then
            Risultato.Num_ODP = Reader_Tree("Numero documento")
            Risultato.Tipo = "Premontaggio-Scorta : "
            Risultato.Num_Cassetta = Reader_Tree("Cassetta")
            Cnn_Tree.Close()
            Return Risultato
        End If
        Cnn_Tree.Close()

        Risultato.Num_ODP = "*"
        Risultato.Tipo = "Premontato"
        Risultato.Num_Cassetta = ""
        Cnn_Tree.Close()
        Return Risultato
    End Function

    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        Dim par_datagridview As DataGridView = DataGridView4

        If e.RowIndex >= 0 Then

            If e.ColumnIndex = par_datagridview.Columns.IndexOf(Itemcode) Then

                Magazzino.Codice_SAP = par_datagridview.Rows(e.RowIndex).Cells(columnName:="itemcode").Value

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



            ElseIf e.ColumnIndex = par_datagridview.Columns.IndexOf(N_DOC) And par_datagridview.Rows(e.RowIndex).Cells(columnName:="DOC").Value = "ODP" Then

                If par_datagridview.Rows(e.RowIndex).Cells(columnName:="N_DOC").Value = 0 Then

                Else

                    ODP_Form.docnum_odp = par_datagridview.Rows(e.RowIndex).Cells(columnName:="N_DOC").Value
                    ODP_Form.Show()
                    ODP_Form.inizializza_form(par_datagridview.Rows(e.RowIndex).Cells(columnName:="N_DOC").Value)



                End If

            ElseIf e.ColumnIndex = par_datagridview.Columns.IndexOf(N_DOC) And par_datagridview.Rows(e.RowIndex).Cells(columnName:="DOC").Value = "OC" Then


                Form_nuova_offerta.Show()

                Form_nuova_offerta.TextBox10.Text = par_datagridview.Rows(e.RowIndex).Cells(columnName:="N_DOC").Value
                Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                Form_nuova_offerta.inizializzazione_form(par_datagridview.Rows(e.RowIndex).Cells(columnName:="N_DOC").Value, "ORDR", "RDR1", par_datagridview.Rows(e.RowIndex).Cells(columnName:="N_DOC").Value)


            End If
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            filtro_doc = ""
        Else
            filtro_doc = TextBox1.Text
        End If
        componenti_datagridview(DataGridView4, Txt_commessa.Text, Label5, Label4, filtro_doc, filtro_n_doc, filtro_itemcode, filtro_itemname)
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Then
            filtro_itemcode = ""
        Else
            filtro_itemcode = TextBox3.Text
        End If
        componenti_datagridview(DataGridView4, Txt_commessa.Text, Label5, Label4, filtro_doc, filtro_n_doc, filtro_itemcode, filtro_itemname)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = "" Then
            filtro_itemname = ""
        Else
            filtro_itemname = TextBox4.Text
        End If
        componenti_datagridview(DataGridView4, Txt_commessa.Text, Label5, Label4, filtro_doc, filtro_n_doc, filtro_itemcode, filtro_itemname)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = "" Then
            filtro_n_doc = ""
        Else
            filtro_n_doc = TextBox2.Text
        End If
        componenti_datagridview(DataGridView4, Txt_commessa.Text, Label5, Label4, filtro_doc, filtro_n_doc, filtro_itemcode, filtro_itemname)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Scheda_commessa_Pianificazione.ExportVisibleColumnsToExcel(DataGridView4)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '   Scheda_commessa_Pianificazione.ExportVisibleColumnsToExcel(DataGridView1)
        Scheda_commessa_Pianificazione.ExportVisibleColumnsToExcel(DataGridView2)
    End Sub
End Class