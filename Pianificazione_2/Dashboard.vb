Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Drawing.Drawing2D
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop



Public Class Dashboard

    Public inizializzazione As Boolean = True


    Sub inizializzazione_form()
        inizializzazione = True
        DateTimePicker4.Value = Today.AddDays(-60)
        DateTimePicker1.Value = Today.AddDays(+1000)
        calcola_backlog()

        inizializzazione = False

    End Sub

    Sub calcola_backlog()
        backlog(DataGridView4, TextBox1.Text, TextBox2.Text, TextBox3.Text, TextBox5.Text, DateTimePicker4, DateTimePicker1, Aperte)
    End Sub

    Sub calcola_KPI_MATERIALE()
        kpi_MATERIALE(DataGridView3, Label1.Text)
    End Sub

    Sub kpi_MATERIALE(par_datagridview As DataGridView, par_commessa As String)



        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "declare @commessa as varchar (6)

set @commessa='" & par_commessa & "'


select t50.stato, t50.n
from
(
Select t40.stato, count(t40.itemcode) as 'N'
from
(
select t30.itemcode,
case 
when t30.DA_TRASF<=0 then 'OK'
WHEN t30.mag>=t30.DA_TRASF THEN 'TRASFERIBILE'
WHEN t30.mag+ t30.mag_est >=t30.DA_TRASF THEN 'MAG_EST'
WHEN t30.mag+ t30.mag_est +t30.In_ordine+t30.In_ordine_est >=t30.DA_TRASF then 'IN_APPROV'
ELSE 'DA_ORDINARE'

END AS 'STATO'
from
(
select t20.itemcode, t20.Da_trasf, t20.mag, t20.in_ordine, t20.disp , sum(t21.onhand) as'Mag_est', sum(t21.onorder) as 'In_ordine_est' , sum(t21.onhand)+sum(t21.onorder)-sum(t21.iscommited) as 'Disp_est'
from
(
select t10.itemcode, t10.Da_trasf, sum(t11.onhand) as'Mag', sum(t11.onorder) as 'In_ordine' , sum(t11.onhand)+sum(t11.onorder)-sum(t11.iscommited) as 'Disp'
from
(
SELECT T5.ITEMCODE, SUM(T5.DA_TRASF) AS 'DA_TRASF'
FROM
(
select t1.itemcode, sum(t1.u_prg_wip_qtadatrasf) as 'Da_trasf'
from
owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
inner join oitm t2 on t2.itemcode=t1.itemcode
LEFT JOIN OWOR T10 ON T10.ITEMCODE=T1.ITEMCODE AND (T10.status<>'C') and T10.[U_PRODUZIONE]='ASSEMBL'


where t0.U_PRG_AZS_Commessa=@commessa and (t0.status<>'C') and t10.itemcode is null and t2.itemtype='I'
AND (SUBSTRING(T1.ITEMCODE,1,1)='0' OR SUBSTRING(T1.ITEMCODE,1,1)='C' OR SUBSTRING(T1.ITEMCODE,1,1)='D')
group by t1.itemcode


union all

select t1.itemcode, sum(t1.U_Datrasferire) as 'Da_trasf'
from
 RDR1 t1 
inner join oitm t2 on t2.itemcode=t1.itemcode
LEFT JOIN OWOR T10 ON T10.ITEMCODE=T1.ITEMCODE AND (T10.STATUS='P' OR T10.STATUS='R') and T10.[U_PRODUZIONE]='ASSEMBL'


where t1.U_PRG_AZS_Commessa=@commessa and (T1.OpenQty>0) and t10.itemcode is null and t2.itemtype='I'
AND (SUBSTRING(T1.ITEMCODE,1,1)='0' OR SUBSTRING(T1.ITEMCODE,1,1)='C' OR SUBSTRING(T1.ITEMCODE,1,1)='D')
group by t1.itemcode
)
AS T5
GROUP BY T5.ITEMCODE

)
as t10 left join oitw t11 on t11.itemcode=t10.itemcode and t11.whscode<>'B03' and t11.whscode<>'03' and t11.whscode<>'06' and t11.whscode<>'b06' and t11.whscode<>'09' and t11.whscode<>'b09' and t11.whscode<>'Clavter' and t11.whscode<>'Bclavter' and t11.whscode<>'CQ'
group by t10.itemcode, t10.Da_trasf
)
as t20 left join oitw t21 on t21.itemcode=t20.itemcode and  (t21.whscode='B03' or t21.whscode='03' or t21.whscode='06' or t21.whscode='B06' or  t21.whscode='clavter' or  t21.whscode='Bclavter')

group by 
t20.itemcode, t20.Da_trasf, t20.mag, t20.in_ordine, t20.disp
)
as t30
)
as t40
group by t40.stato
)
as t50
order by t50.n desc
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        ' Dichiarazione delle variabili
        Dim totale As Integer = 0

        ' Prima passata per calcolare il totale di N
        Do While cmd_SAP_reader_2.Read()
            totale += Convert.ToInt32(cmd_SAP_reader_2("N"))
        Loop

        ' Riposiziona il DataReader all'inizio
        cmd_SAP_reader_2.Close()
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader()

        ' Seconda passata per aggiungere le righe con la percentuale
        Do While cmd_SAP_reader_2.Read()
            Dim stato As String = cmd_SAP_reader_2("Stato").ToString()
            Dim n As Integer = Convert.ToInt32(cmd_SAP_reader_2("N"))

            ' Calcola la percentuale
            Dim percentuale As Double = (n / totale) * 100

            ' Aggiungi i dati al DataGridView, con la percentuale nella terza colonna
            par_datagridview.Rows.Add(stato, n, Math.Round(percentuale, 2))
        Loop

        ' Chiudi le connessioni
        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        ' Optional: Rimuove la selezione delle righe
        ' par_datagridview.ClearSelection()
    End Sub

    Sub backlog(par_datagridview As DataGridView, par_oc As String, par_commessa As String, par_cliente As String, par_div As String, par_datetimepicker_inizio As DateTimePicker, par_datetimepicker_fine As DateTimePicker, par_radiobutton As RadioButton)

        Dim filtro_oc As String
        Dim filtro_commessa As String
        Dim filtro_cliente As String
        Dim filtro_div As String


        If par_oc = "" Then
            filtro_oc = ""
        Else
            filtro_oc = " and t40.[Order N°]='" & par_oc & "'"
        End If


        If par_commessa = "" Then
            filtro_commessa = ""
        Else
            filtro_commessa = " and t40.codice    Like '%%" & par_commessa & "%%'"
        End If

        If par_cliente = "" Then
            filtro_cliente = ""
        Else
            filtro_cliente = " and (t40.[Bp name] Like '%%" & par_cliente & "%%' or t40.[Final BP] Like '%%" & par_cliente & "%%')"
        End If


        If par_div = "" Then
            filtro_div = ""
        Else
            filtro_div = " and   t40.ocrcode Like '%%" & par_div & "%%' "
        End If

        Dim filtro_aperte

        If par_radiobutton.Checked = True Then
            filtro_aperte = " and T1.[OpenQty]> 0 And T0.[DocStatus] ='O'"

        Else

            filtro_aperte = ""
        End If

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "Select t40.[Order N°],t40.progetto,T40.U_CAUSCONS AS 'Causale', t40.[Codice], t40.[Descrizione articolo], T40.OcrCode,t40.[Insert date], t40.[Due date], t40.[Bp name], t40.[Final BP],t40.freetxt,T40.PM, t40.[Delivery country], t40.[Incoterms], t40.Settore, t40.[Total Open amount], T40.[Pricing list], t40.Discount, t40.Provvigione, t40.Multiplier, t40.[Totale ordine], (T40.[Acconto richiesto]+ sum(case when T40.[Richiesto stornato] is null then 0 else T40.[Richiesto stornato] end ))/case when t40.[Totale ordine] =0 then 1 else t40.[Totale ordine] end  as 'Acconto richiesto'
, (T40.[Acconto pagato]+ sum(case when t40.[Pagato stornato] is null then 0 else t40.[Pagato stornato] end))/case when t40.[Totale ordine] is null then 1 else case when t40.[Totale ordine] =0 then 1 else t40.[Totale ordine] end  end  as 'Acconto pagato',  T40.[PymntGroup]
,coalesce(t42.nome_baia,'') as 'Nome_baia'
from
(
SELECT t30.[Order N°],t30.progetto,T30.U_CAUSCONS, t30.[Codice], t30.[Descrizione articolo],t30.[Insert date], t30.[Due date], t30.[Bp name], t30.[Final BP],T30.PM, t30.[Delivery country], t30.[Incoterms], t30.[Data KOM], t30.Settore, t30.[Total Open amount], T30.[Pricing list], t30.Discount, t30.Provvigione, t30.Multiplier, t30.[Totale ordine], T30.[Acconto richiesto], T30.[Acconto pagato],-T33.DOCTOTAL AS 'Richiesto stornato ', -t33.paidtodate as 'Pagato stornato', T33.DOCNUM, T30.[PymntGroup],T30.OcrCode,t30.freetxt
FROM
(
SELECT t20.[Order N°],t20.progetto,T20.U_CAUSCONS, t20.[Codice], t20.[Descrizione articolo],t20.[Insert date], t20.[Due date], t20.[Bp name], t20.[Final BP],T20.PM, t20.[Delivery country], t20.[Incoterms], t20.[Data KOM], t20.Settore, t20.[Total Open amount], T20.[Pricing list], t20.Discount, t20.Provvigione, t20.Multiplier, t20.[Totale ordine], SUM(T20.[Acconto richiesto]) AS 'Acconto richiesto', SUM(t20.[Acconto pagato]) as 'Acconto pagato', T20.[PymntGroup], T20.OcrCode,t20.freetxt
FROM
(
Select t10.[Order N°],t10.progetto, T10.U_CAUSCONS, t10.[Codice], t10.[Descrizione articolo],  t10.[Insert date], t10.[Due date], t10.[Bp name], t10.[Final BP],coalesce(t10.gestore_progetto,T11.LASTNAME) AS'PM', t10.[Delivery country], t10.[Incoterms], t10.[Data KOM], t10.Settore, t10.[Total Open amount], T10.[Pricing list], t10.Discount, t10.Provv as 'Provvigione', t10.Multiplier, t10.tot_ordine AS 'Totale ordine', T13.DOCTOTAL AS 'Acconto richiesto', t13.paidtodate as 'Acconto pagato', T13.DOCNUM, T10.[PymntGroup], T10.OcrCode,t10.freetxt
from (
Select  T0.PM, T0.U_CAUSCONS, t0.[Order N°], t0.[Codice]
,substring(t0.[codice],1,1) as 'Prima lettera', t0.[descrizione articolo], t0.[Family],   t0.[insert date]
, case when month(t0.[insert date])>9 then year(t0.[insert date])+1 else year(t0.[insert date]) end as 'Fiscal year', t0.[due date], t0.[Data RDO], t0.[Data CM], t0.[Tirelli Owner], t0.[Tirelli salesman], t0.[Office], t0.[causale], t0.[bp code], t0.[bp name], t0.[final BP code], t0.[final BP]
, t0.[Arol branch], t0.[Arol salesman], t0.[OEM/Agent], t0.[Invoice country], t0.[Delivery country], t0.[Sector],  t0.[Incoterms], t0.[Data KOM], t0.settore,  t0.[Total Open amount] ,  t0.[Total amount], T0.[Pricing list], t0.discount, t0.Provv, T0.[Pricing list]*(1- t0.discount)*(1-T0.PROVV) as 'Multiplier', case when (T0.[Pricing list]*(1- t0.discount)) <>0 then  t0.[Total Open amount]/(T0.[Pricing list]*(1- t0.discount)) end as 'Cost', t0.tot_ordine, T0.[PymntGroup], T0.OcrCode,t0.freetxt, t0.progetto,t0.Gestore_progetto
from (
SELECT  T0.[DocNum] as 'Order N°', T0.OWNERCODE AS 'PM', T0.U_CausCons, T1.ITEMCODE AS 'Codice', T1.[Dscription] as 'descrizione articolo', T1.OcrCode, t10.name as 'Family', T0.[DocDate] as 'Insert date', case when T1.[ShipDate] is null then t0.docduedate else t1.shipdate end as 'Due date', T0.u_DATARDO AS 'Data RDO', t0.U_datacm as 'Data Cm', t7.lastname+' '+t7.firstname as 'Tirelli owner', t8.slpname as 'Tirelli salesman', T0.u_UFFCOMPETENZA AS 'Office', T0.u_CAUSCONS AS 'Causale', T0.[CardCode] as 'Bp Code', T0.[CardName] as 'Bp name', T0.u_CODICEBP AS 'Final BP code', t12.cardname as 'Final BP', t0.U_arolbranch as 'Arol branch', t4.name as 'Arol salesman',  T0.U_DISTRIBUTORE AS 'OEM/Agent', T6.name as  'Invoice country', T0.[U_Destinazione] as 'Delivery country', t0.U_settore as 'Sector', T0.U_PRG_AZS_INCOTERMS as 'Incoterms', T1.[U_DataKOM] as 'Data KOM', T0.[U_Settore] AS 'Settore',

  sum(case when t0.doctype='S' and t1.linestatus='O' then t1.linetotal else (t1.openqty*price)*((100 - case when t0.discprcnt is null then 0 else t0.discprcnt end)/100)/t0.docrate end)  as 'Total Open amount' , 

 sum(case when t0.doctype='S'  then t1.linetotal else (t1.quantity*price)*((100 - case when t0.discprcnt is null then 0 else t0.discprcnt end)/100)/t0.docrate end) as 'Total amount',

 sum(T1.LINETOTAL*((100-case when t0.discprcnt is null then 0 else t0.discprcnt end)/100))  AS 'Net total amount', T1.[U_coefficiente_vendita] as 'Pricing list', case when t0.doctype='S' then sum(t1.pricebefdi/case when t1.rate=0 then 1 else t1.rate end) else sum(t1.quantity*T1.[PriceBefDi]/case when t1.rate ='0' then 1 else t1.rate end) end as 'Total', t1.commission/100 as 'Provv',
1-(case when case when t0.doctype='S' then sum(t1.pricebefdi/case when t1.rate=0 then 1 else t1.rate end) else sum(t1.quantity*T1.[PriceBefDi]/case when t1.rate ='0' then 1 else t1.rate end) end = '0' then '0' else sum(T1.LINETOTAL*((100-case when t0.discprcnt is null then 0 else t0.discprcnt end)/100))/case when t0.doctype='S' then sum(t1.pricebefdi/case when t1.rate=0 then 1 else t1.rate end) else sum(t1.quantity*T1.[PriceBefDi]/case when t1.rate ='0' then 1 else t1.rate end) end  end) AS 'Discount', T0.DOCTOTAL AS 'Tot_ordine', T11.[PymntGroup], t1.freetxt, t13.docnum as 'Progetto'
,t14.lastname as 'Gestore_progetto'


FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.[DocEntry] = T1.[DocEntry]
INNER JOIN OCRD T2 ON T2.CARDCODE=T0.CARDCODE
LEFT JOIN OITM T3 ON T3.ITEMCODE=T1.ITEMCODE
LEFT JOIN [dbo].[@VENDITORI_AROL]  T4 ON T4.CODE=T0.u_VENDITOREAROL
left join ocry t6 on t6.code=t2.country
left join [TIRELLI_40].[dbo].OHEM T7 ON t7.empid=t0.ownercode
LEFT join OSLP T8 ON T8.slpcode =t0.slpcode
LEFT JOIN OACT T9 ON T9.[AcctCode]=T1.[AcctCode]
LEFT JOIN OCTG T11 ON T0.[GroupNum] = T11.[GroupNum]
left join [dbo].[@FAMIGLIA_VENDITA]  T10 on t10.code= t9.[U_FamigliaVendita]
left join ocrd t12 on t12.cardcode=t0.u_CODICEBP
left join opmg t13 on t13.absentry=t3.u_progetto
left join [TIRELLI_40].[dbo].ohem t14 on t14.empid=t13.owner
WHERE 0=0 and substring(t1.itemcode,1,1)='M' " & filtro_aperte & "   
--AND (T0.U_CausCons='V' OR T0.U_CausCons='UNCLEAN')
group by
T0.[DocNum], T1.ITEMCODE, T1.[Dscription], T0.[DocDate], T0.[DocDueDate] , T0.[CardCode], T0.[CardName], T0.u_CODICEBP ,T6.name, T0.[U_Destinazione], t7.firstname, t7.lastname, T0.U_PRG_AZS_INCOTERMS,  t0.U_arolbranch,  t4.name,  T1.[U_coefficiente_vendita], t0.doctype, t0.U_causcons, t8.slpname, t0.U_settore, T0.u_UFFCOMPETENZA, T0.U_DISTRIBUTORE, T0.u_DATARDO, t0.U_datacm, t10.name, t1.shipdate, T1.[U_DataKOM], T0.[U_Settore],t1.commission,T0.OWNERCODE,T0.DOCTOTAL , T11.[PymntGroup],t12.cardname, T1.OcrCode ,t1.freetxt, t13.docnum,t14.lastname
)
as t0

)
as t10
INNER JOIN [TIRELLI_40].[dbo].OHEM T11 ON T10.PM=T11.CODE
left join DPI1 T12 ON T12.BASEREF=T10.[Order N°]
LEFT JOIN ODPI T13 ON T13.DOCENTRY=T12.DOCENTRY
where (t10.[Prima lettera]='M' ) 
GROUP BY t10.[Order N°],t10.progetto,T10.U_CAUSCONS, t10.[Codice], t10.[Descrizione articolo],t10.[Insert date], t10.[Due date], t10.[Bp name], t10.[Final BP],T11.LASTNAME, t10.[Delivery country], t10.[Incoterms], t10.[Data KOM], t10.Settore, t10.[Total Open amount], T10.[Pricing list], t10.Discount, t10.Provv, t10.Multiplier, t10.tot_ordine , T13.DOCTOTAL, t13.paidtodate, T13.DOCNUM, T10.[PymntGroup],T10.OcrCode,t10.freetxt,t10.Gestore_progetto
)
AS T20
GROUP BY
t20.[Order N°],T20.U_CAUSCONS,t20.progetto, t20.[Codice], t20.[Descrizione articolo],t20.[Insert date], t20.[Due date], t20.[Bp name], t20.[Final BP],T20.PM, t20.[Delivery country], t20.[Incoterms], t20.[Data KOM], t20.Settore, t20.[Total Open amount], T20.[Pricing list], t20.Discount, t20.Provvigione, t20.Multiplier, t20.[Totale ordine], T20.[PymntGroup],T20.OcrCode,t20.freetxt

)
AS T30
left join DPI1 T12 ON T12.BASEREF=T30.[Order N°]
LEFT JOIN ODPI T13 ON T13.DOCENTRY=T12.DOCENTRY
left join RIN1 T32 ON T32.BASEentry=t13.docentry
LEFT JOIN ORIN T33 ON T33.DOCENTRY=T32.DOCENTRY
group by
t30.[Order N°],T30.U_CAUSCONS,t30.progetto, t30.[Codice], t30.[Descrizione articolo],t30.[Insert date], t30.[Due date], t30.[Bp name], t30.[Final BP],T30.PM, t30.[Delivery country], t30.[Incoterms], t30.[Data KOM], t30.Settore, t30.[Total Open amount], T30.[Pricing list], t30.Discount, t30.Provvigione, t30.Multiplier, t30.[Totale ordine], t30.[Acconto richiesto], t30.[Acconto pagato],-T33.DOCTOTAL , t33.paidtodate , T33.DOCNUM, T30.[PymntGroup], T30.OcrCode,t30.freetxt
)
as t40
left join [Tirelli_40].[dbo].[Layout_CAP1] t41 on t41.commessa=t40.codice and t41.stato='O'
left join [Tirelli_40].[dbo].[Layout_CAP1_nomi] t42 on t41.baia=t42.numero_baia
where t40.[Due date]>=CONVERT(DATETIME, '" & par_datetimepicker_inizio.Value & "', 103) and t40.[Due date]<=CONVERT(DATETIME, '" & par_datetimepicker_fine.Value & "', 103) " & filtro_commessa & filtro_cliente & filtro_div & " 
group by
 t40.[Order N°],t40.progetto,T40.U_CAUSCONS, t40.[Codice], t40.[Descrizione articolo],t40.[Insert date], t40.[Due date], t40.[Bp name], t40.[Final BP],T40.PM, t40.[Delivery country], t40.[Incoterms], t40.[Data KOM], t40.Settore, t40.[Total Open amount], T40.[Pricing list], t40.Discount, t40.Provvigione, t40.Multiplier, t40.[Totale ordine], T40.[Acconto richiesto], T40.[Acconto pagato], T40.[PymntGroup], T40.OcrCode,t40.freetxt ,coalesce(t42.nome_baia,'')
order by t40.[due date],t40.progetto, t40.codice"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader_2("Order N°"), cmd_SAP_reader_2("progetto"), cmd_SAP_reader_2("Causale"), cmd_SAP_reader_2("codice"), cmd_SAP_reader_2("Descrizione articolo"), cmd_SAP_reader_2("OcrCode"), cmd_SAP_reader_2("Insert date"), cmd_SAP_reader_2("Due date"), cmd_SAP_reader_2("Bp name"), cmd_SAP_reader_2("Final BP"), cmd_SAP_reader_2("Nome_baia"), cmd_SAP_reader_2("Total Open amount"))
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        ' par_datagridview.ClearSelection()
    End Sub

    Private Sub Dashboard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        inizializzazione_form()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs)
        Me.WindowState = FormWindowState.Minimized
    End Sub







    Sub dati_commessa(par_commessa As String, par_progetto As String)
        Label1.Text = par_commessa
        Label2.Text = Magazzino.OttieniDettagliAnagrafica(Label1.Text).u_progetto
        Label3.Text = Magazzino.OttieniDettagliAnagrafica(Label1.Text).Descrizione
        Label4.Text = Form_layout_CAP_1.check_baia_layout(Label1.Text).nome_baia
        Scheda_tecnica.riempi_datagridview_combinazioni(DataGridView1, par_commessa, Homepage.sap_tirelli)
        carica_appunti(par_commessa, "COMMESSA", DataGridView2, "", Homepage.ID_SALVATO)
        carica_appunti(par_progetto, "PROGETTO", DataGridView5, "", Homepage.ID_SALVATO)
        ' calcola_KPI_MATERIALE()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Commesse_MES.SCHEDA_COMMESSA(Label1.Text)


        Form_Scheda_Collaudi.inizializzazione_form(Label1.Text)
        Form_Scheda_Collaudi.Show()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        FORM6.Show()


        FORM6.inizializza_form(Label1.Text)
        FORM6.news_materiale()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click



        If Label1.Text >= "M04000" Then
            Scheda_tecnica.Show()
            Scheda_tecnica.BringToFront()
            Scheda_tecnica.inizializza_scheda_tecnica(Label1.Text)
        Else
            MsgBox("Non è minore")
        End If




    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged


    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        'calcola_backlog()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        '  calcola_backlog()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs)
        '  calcola_backlog()
    End Sub

    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        'If inizializzazione = False Then
        '    calcola_backlog()
        'End If

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        'If inizializzazione = False Then
        '    calcola_backlog()
        'End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub



    Private Sub DateTimePicker4_Leave(sender As Object, e As EventArgs) Handles DateTimePicker4.Leave

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        ' calcola_backlog()
    End Sub

    Private Sub Aperte_CheckedChanged(sender As Object, e As EventArgs) Handles Aperte.CheckedChanged
        'If inizializzazione = False Then
        '    calcola_backlog()
        'End If
    End Sub

    Private Sub Button20_Click_1(sender As Object, e As EventArgs) Handles Button20.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub


    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellClick

        dati_commessa(DataGridView4.Rows(e.RowIndex).Cells(columnName:="Matricola").Value, DataGridView4.Rows(e.RowIndex).Cells(columnName:="Prog").Value)

        '  calcola_KPI_MATERIALE()

        ' Logica per visualizzare il progetto
        If e.ColumnIndex = DataGridView4.Columns.IndexOf(Prog) Then
            Progetto.Show()
            Progetto.BringToFront()
            Progetto.absentry = DataGridView4.Rows(e.RowIndex).Cells(columnName:="Prog").Value
            Progetto.inizializza_progetto()
        End If
    End Sub


    ' Dizionario globale per mantenere l'associazione prog → colore
    Private progColors As New Dictionary(Of String, Color)
    Private availableColors As Color() = {Color.Red, Color.Blue, Color.Green, Color.Orange, Color.Purple, Color.Brown, Color.DarkCyan, Color.DarkMagenta}

    Private Sub DataGridView4_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView4.CellFormatting

        Dim par_datagridview As DataGridView = DataGridView4

        '--- Gestione colorazione per DIV
        If e.ColumnIndex = par_datagridview.Columns("DIV").Index AndAlso e.Value IsNot Nothing Then
            Select Case e.Value.ToString()
                Case "BRB01"
                    par_datagridview.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.Yellow
                Case "TIR01"
                    par_datagridview.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.Aqua
                Case "KTF01"
                    par_datagridview.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.Green
            End Select
        End If

        '--- Gestione colorazione per PROG
        If par_datagridview.Columns.Contains("prog") Then
            Dim progValue As String = par_datagridview.Rows(e.RowIndex).Cells("prog").Value.ToString()

            ' Se non c’è ancora un colore associato a questo prog, assegnane uno
            If Not progColors.ContainsKey(progValue) Then
                Dim nextColor As Color = availableColors(progColors.Count Mod availableColors.Length)
                progColors(progValue) = nextColor
            End If

            ' Applica il colore di ForeColor a tutta la riga
            par_datagridview.Rows(e.RowIndex).DefaultCellStyle.ForeColor = progColors(progValue)
        End If

    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Down Then
            ' Simulo il click del Button6
            Button6.PerformClick()
        ElseIf e.KeyCode = Keys.Up Then
            Button7.PerformClick()
        ElseIf e.KeyCode = Keys.Enter Then
            Button13.PerformClick()
        End If
    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim par_datagridview As DataGridView = DataGridView4
        Dim riga As Integer = par_datagridview.CurrentCell.RowIndex ' Ottenere l'indice della riga selezionata

        ' Verifica se esiste una riga successiva
        If riga < par_datagridview.Rows.Count - 1 Then
            ' Sposta la selezione alla riga successiva
            par_datagridview.CurrentCell = par_datagridview.Rows(riga + 1).Cells(par_datagridview.CurrentCell.ColumnIndex)

            ' Forza il trigger manuale dell'aggiornamento
            Call UpdateSelectedRowData(par_datagridview)
        Else
            MessageBox.Show("Non ci sono più righe successive.")
        End If
    End Sub

    Private Sub DataGridView4_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView4.SelectionChanged
        Dim par_datagridview As DataGridView = DataGridView4
        ' Chiama la funzione di aggiornamento dei dati
        Call UpdateSelectedRowData(par_datagridview)
    End Sub

    Private Sub UpdateSelectedRowData(par_datagridview As DataGridView)
        ' Verifica se c'è almeno una cella o riga selezionata
        If par_datagridview.CurrentRow IsNot Nothing Then
            ' Ottieni l'indice della riga selezionata
            Dim rigaSelezionata As Integer = par_datagridview.CurrentRow.Index

            ' Assicurati che l'indice sia valido e la colonna "Matricola" esista
            If rigaSelezionata >= 0 AndAlso par_datagridview.Columns.Contains("Matricola") Then
                ' Estrai il valore della cella nella colonna "Matricola"
                Dim matricolaValue As String = par_datagridview.Rows(rigaSelezionata).Cells("Matricola").Value.ToString()
                Dim ProgettoValue As String = par_datagridview.Rows(rigaSelezionata).Cells("Prog").Value.ToString()
                ' Chiama la funzione dati_commessa passando il valore estratto
                dati_commessa(matricolaValue, ProgettoValue)
            End If

            ' Gestisci la visibilità dei bottoni
            Button7.Visible = rigaSelezionata <> 0 ' Nasconde Button7 se è la prima riga
            If inizializzazione = False Then
                Button6.Visible = rigaSelezionata <> par_datagridview.Rows.Count - 1 ' Nasconde Button6 se è l'ultima riga
            End If
        End If
    End Sub




    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim par_datagridview As DataGridView = DataGridView4
        Dim riga As Integer = par_datagridview.CurrentCell.RowIndex ' Ottenere l'indice della riga selezionata

        ' Verifica se esiste una riga precedente
        If riga > 0 Then
            ' Sposta la selezione alla riga precedente
            par_datagridview.CurrentCell = par_datagridview.Rows(riga - 1).Cells(par_datagridview.CurrentCell.ColumnIndex)

            ' Forza il trigger manuale dell'aggiornamento
            Call UpdateSelectedRowData(par_datagridview)
        Else
            MessageBox.Show("Non ci sono più righe precedenti.")
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If ComboBox2.SelectedIndex < 0 Then
            MsgBox("Selezionare una destinazione dell'appunto")
            Return
        End If
        If ComboBox2.SelectedIndex = 0 Then
            inserisci_nuovo_appunto(Homepage.ID_SALVATO, ComboBox2.Text, Label1.Text, 0, Replace(RichTextBox2.Text, "'", " "), False, False, False)
        ElseIf ComboBox2.SelectedIndex = 1 Then
            inserisci_nuovo_appunto(Homepage.ID_SALVATO, ComboBox2.Text, Label2.Text, 0, Replace(RichTextBox2.Text, "'", " "), False, False, False)

        End If

        carica_appunti(Label1.Text, "COMMESSA", DataGridView2, "", Homepage.ID_SALVATO)
        carica_appunti(Label2.Text, "PROGETTO", DataGridView5, "", Homepage.ID_SALVATO)
        RichTextBox2.Text = ""
    End Sub

    Sub inserisci_nuovo_appunto(par_dipendente As Integer, par_tipo As String, par_commessa As String, par_n_progetto As Integer, par_contenuto As String, par_grassetto As Boolean, par_corsivo As Boolean, par_risolto As Boolean)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand
        CMD_SAP_5.Connection = CNN6
        CMD_SAP_5.CommandText = "INSERT INTO [TIRELLI_40].[DBO].Appunti_commesse
       ([Dipendente]
       ,[Data]
       ,[Ora]
,[tipo]
       ,[Commessa]
,[progetto]
       ,[Contenuto]
       ,[Grassetto]
       ,[Corsivo]
       ,[risolto])
 VALUES
       (@Dipendente
       ,getdate()
       ,@Ora
,@tipo
       ,@Commessa
,@progetto
       ,@Contenuto
       ,@Grassetto
       ,@Corsivo
       ,@risolto)"

        ' Aggiunta dei parametri
        CMD_SAP_5.Parameters.AddWithValue("@Dipendente", par_dipendente)

        CMD_SAP_5.Parameters.AddWithValue("@Ora", DateTime.Now.TimeOfDay) ' Ora attuale del sistema
        CMD_SAP_5.Parameters.AddWithValue("@Commessa", par_commessa)
        CMD_SAP_5.Parameters.AddWithValue("@Contenuto", par_contenuto)
        '  CMD_SAP_5.Parameters.AddWithValue("@Font", par_font)
        CMD_SAP_5.Parameters.AddWithValue("@Grassetto", par_grassetto)
        CMD_SAP_5.Parameters.AddWithValue("@Corsivo", par_corsivo)
        CMD_SAP_5.Parameters.AddWithValue("@risolto", par_risolto)
        CMD_SAP_5.Parameters.AddWithValue("@tipo", par_tipo)
        CMD_SAP_5.Parameters.AddWithValue("@progetto", par_n_progetto)

        CMD_SAP_5.ExecuteNonQuery()

        CNN6.Close()
    End Sub


    Sub carica_appunti(par_commessa As String, par_tipo As String, par_Datagridview As DataGridView, par_annulla_filtro As String, par_dipendente As Integer)

        Dim filtro_dipendente As String
        If par_dipendente = 0 Then
            filtro_dipendente = ""
        Else
            filtro_dipendente = " and t0.dipendente = " & par_dipendente & ""
        End If

        ' par_richtextbox.Text = ""
        par_Datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT  t0.[ID]
      ,t0.[Dipendente]
	  ,concat(t1.lastname, ' ', substring(t1.firstname,1,1)) as 'Nome'
      ,t0.[Data]
      ,t0.[Ora]
      ,t0.[Commessa]
,COALESCE(t2.[U_FINAL_cUSTOMER_NAME],'') AS 'U_FINAL_CUSTOMER_NAME'
      ,t0.[Contenuto]
      ,t0.[Font]
      ,t0.[Grassetto]
      ,t0.[Corsivo]
      ,t0.[risolto]
  FROM [TIRELLI_40].[DBO].Appunti_commesse t0 
  left join [TIRELLI_40].[dbo].ohem t1 on t0.dipendente=t1.empid
left join oitm t2 on t2.itemcode=t0.commessa
where (t0.commessa='" & par_commessa & "' or cast(t0.[Progetto] as varchar) ='" & par_commessa & "') and t0.[Tipo]='" & par_tipo & "' " & par_annulla_filtro & filtro_dipendente & " 
  order by t0.[Commessa],t0.[ID]
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            par_Datagridview.Rows.Add(cmd_SAP_reader_2("id"), cmd_SAP_reader_2("commessa"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("data"), cmd_SAP_reader_2("contenuto"), cmd_SAP_reader_2("Risolto"), cmd_SAP_reader_2("U_FINAL_CUSTOMER_NAME"))
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        par_Datagridview.ClearSelection()

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        modifica(DataGridView2, "ID", "Commento")
        carica_appunti(Label1.Text, "COMMESSA", DataGridView2, "", Homepage.ID_SALVATO)
        MsgBox("Aggiornato con successo")
    End Sub

    Sub modifica(par_datagridview As DataGridView, par_nome_id As String, par_nome_colonna_commento As String)
        ' Itera attraverso tutte le righe della DataGridView2
        For Each row As DataGridViewRow In par_datagridview.Rows
            ' Verifica che la riga non sia una riga vuota (ad esempio, la riga vuota in fondo)
            If Not row.IsNewRow Then
                ' Ottieni il valore della colonna "contenuto" per la riga corrente
                Dim contenuto As Object = row.Cells(par_nome_colonna_commento).Value
                aggiorna_appunti(row.Cells(par_nome_id).Value, row.Cells(par_nome_colonna_commento).Value)

                ' Esegui la logica desiderata con il valore della colonna "contenuto"
                ' MessageBox.Show("Contenuto: " & contenuto.ToString())
            End If
        Next
    End Sub

    Sub modifica_appunti_globali(par_datagridview As DataGridView)
        ' Itera attraverso tutte le righe della DataGridView2
        For Each row As DataGridViewRow In par_datagridview.Rows
            ' Verifica che la riga non sia una riga vuota (ad esempio, la riga vuota in fondo)
            If Not row.IsNewRow Then
                If row.Cells("ID_3").Value <> 0 Then


                    ' Ottieni il valore della colonna "contenuto" per la riga corrente
                    Dim contenuto As Object = row.Cells("commento__").Value
                    aggiorna_appunti(row.Cells("ID_3").Value, row.Cells("commento__").Value)

                    ' Esegui la logica desiderata con il valore della colonna "contenuto"
                    ' MessageBox.Show("Contenuto: " & contenuto.ToString())
                End If
            End If
        Next
    End Sub



    Private Sub CancellaRigaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CancellaRigaToolStripMenuItem.Click
        'If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Trasferito").Value > 0 Then
        '    MsgBox("Impossibile cancellare riga di un codice che risulta TRASFERITO")
        'Else
        Dim PAR_DATAGRIDVIEW As DataGridView
        PAR_DATAGRIDVIEW = DataGridView2

        ' Supponendo che COLONNAID sia il nome della colonna che vuoi usare per cancellare il commento
        Dim COLONNAID As String = "ID"
        Dim selectedRow As DataGridViewRow = PAR_DATAGRIDVIEW.CurrentRow

        ' Verifica che ci sia una riga selezionata prima di procedere
        If selectedRow IsNot Nothing Then
            ' Chiede conferma all'utente se vuole cancellare il commento
            Dim result As DialogResult = MessageBox.Show("Sei sicuro di voler cancellare il commento?" & vbCrLf & selectedRow.Cells("Commento").Value, "Conferma Cancellazione", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            ' Se l'utente conferma, procedi con la cancellazione
            If result = DialogResult.Yes Then
                ' Passa l'ID della riga alla funzione cancella_commento
                cancella_commento(selectedRow.Cells(COLONNAID).Value)

                ' Rimuovi la riga selezionata
                PAR_DATAGRIDVIEW.Rows.RemoveAt(selectedRow.Index)
            End If
        Else
            MessageBox.Show("Seleziona una riga prima di cancellarla.")
        End If

    End Sub

    Sub aggiorna_appunti(par_ID As Integer, par_commento As String)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand
        CMD_SAP_5.Connection = CNN6
        CMD_SAP_5.CommandText = "UPDATE [TIRELLI_40].[DBO].Appunti_commesse

       SET [Contenuto]=@Contenuto
WHERE ID=@ID"

        ' Aggiunta dei parametri
        CMD_SAP_5.Parameters.AddWithValue("@Contenuto", par_commento)
        CMD_SAP_5.Parameters.AddWithValue("@ID", par_ID)



        CMD_SAP_5.ExecuteNonQuery()

        CNN6.Close()
    End Sub

    Sub cancella_commento(par_ID As Integer)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand
        CMD_SAP_5.Connection = CNN6
        CMD_SAP_5.CommandText = "DELETE [TIRELLI_40].[DBO].dati_mancanti_progetto

      
WHERE ID=@ID"

        ' Aggiunta dei parametri

        CMD_SAP_5.Parameters.AddWithValue("@ID", par_ID)



        CMD_SAP_5.ExecuteNonQuery()

        CNN6.Close()
    End Sub

    Sub cambia_stato(par_ID As Integer)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand
        CMD_SAP_5.Connection = CNN6
        CMD_SAP_5.CommandText = "UPDATE [TIRELLI_40].[DBO].Appunti_commesse
SET  RISOLTO=CASE WHEN RISOLTO='FALSE' THEN 'TRUE' ELSE 'FALSE' END

      
WHERE ID=@ID"

        ' Aggiunta dei parametri

        CMD_SAP_5.Parameters.AddWithValue("@ID", par_ID)



        CMD_SAP_5.ExecuteNonQuery()

        CNN6.Close()
    End Sub



    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        ' Verifica se la colonna "stato" è presente (sostituisci "stato" con il nome corretto della colonna)
        Dim statoIndex As Integer = DataGridView2.Columns("stato").Index

        ' Verifica se siamo in una riga valida e non è una riga nuova
        If e.RowIndex >= 0 AndAlso Not DataGridView2.Rows(e.RowIndex).IsNewRow Then
            ' Controlla il valore della colonna "stato"
            Dim statoValue As Boolean = Convert.ToBoolean(DataGridView2.Rows(e.RowIndex).Cells(statoIndex).Value)

            ' Se il valore è "True", applica il font barrato a tutta la riga
            If statoValue Then
                For Each cell As DataGridViewCell In DataGridView2.Rows(e.RowIndex).Cells
                    cell.Style.Font = New Font(DataGridView2.Font, FontStyle.Strikeout)
                Next
            Else
                ' Rimuove il font barrato se non è "True"
                For Each cell As DataGridViewCell In DataGridView2.Rows(e.RowIndex).Cells
                    cell.Style.Font = New Font(DataGridView2.Font, FontStyle.Regular)
                Next
            End If
        End If
    End Sub

    Private Sub DatiAnagraficiArticoloToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DatiAnagraficiArticoloToolStripMenuItem.Click
        'If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Trasferito").Value > 0 Then
        '    MsgBox("Impossibile cancellare riga di un codice che risulta TRASFERITO")
        'Else
        Dim PAR_DATAGRIDVIEW As DataGridView
        PAR_DATAGRIDVIEW = DataGridView2

        ' Supponendo che COLONNAID sia il nome della colonna che vuoi usare per cancellare il commento
        Dim COLONNAID As String = "ID"
        Dim selectedRow As DataGridViewRow = PAR_DATAGRIDVIEW.CurrentRow

        ' Verifica che ci sia una riga selezionata prima di procedere
        If selectedRow IsNot Nothing Then
            ' Chiede conferma all'utente se vuole cancellare il commento

            cambia_stato(selectedRow.Cells(COLONNAID).Value)
            carica_appunti(Label1.Text, "COMMESSA", DataGridView2, "", Homepage.ID_SALVATO)


        Else
            MessageBox.Show("Seleziona una riga prima di CAMBIARNE STATO")
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        ' manda_mail(ComboBox1.Text, Homepage.ID_SALVATO)

        mail_progetti(ComboBox1.Text, Homepage.ID_SALVATO)
    End Sub











    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        carica_appunti(Label1.Text, "COMMESSA", DataGridView2, " or 0 = 0 ", Homepage.ID_SALVATO)
        Dim par_datagridview As DataGridView = DataGridView2
        ' Controlla se la colonna "commessa" esiste
        If par_datagridview.Columns.Contains("commessa") Then
            ' Inverti la visibilità della colonna "commessa"
            par_datagridview.Columns("commessa").Visible = Not par_datagridview.Columns("commessa").Visible
        End If

        ' Controlla se la colonna "cliente_finale" esiste
        If par_datagridview.Columns.Contains("cliente_") Then
            ' Inverti la visibilità della colonna "cliente_finale"
            par_datagridview.Columns("cliente_").Visible = Not par_datagridview.Columns("cliente_").Visible
        End If
    End Sub



    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        If DataGridView1.Rows(e.RowIndex).Cells("Collaudo").Value = 1 Then
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
        End If
    End Sub



    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Form_costificazione.commessa = Label1.Text
        Form_costificazione.Show()
    End Sub


    ' --- Sub principale per creare e inviare l'email dei progetti ---
    Public Sub mail_progetti(par_lista_distribuzione As String, par_id_dipendente As Integer)
        Dim filtro_Dipendente As String = If(par_id_dipendente = 0, "", " AND t0.dipendente = " & par_id_dipendente)
        Dim emailBody As New StringBuilder()
        Dim destinatariUnici As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        ' Variabili a livello di mail_progetti
        Dim colori As String() = {"blue", "green", "red", "purple", "orange"}
        Dim colorIndex As Integer = 0
        Dim coloriTag As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)


        Using Cnn1 As New SqlConnection(Homepage.sap_tirelli)
            Cnn1.Open()

            ' --- Query progetti ---
            Using CMD_SAP_2 As New SqlCommand("
            SELECT t20.progetto_commessa, t21.cardname,
                   COALESCE(t22.cardname,'') AS Cliente_finale,
                   t21.name, t20.data_min
            FROM (
                SELECT t10.progetto_commessa, MIN(t10.DOCDUEDATE) AS data_min
                FROM (
                    SELECT T0.ID, T0.tipo,
                           CAST(T0.Commessa AS VARCHAR) AS Commessa,
                           CAST(T0.progetto AS VARCHAR) AS Progetto,
                           CASE WHEN T0.tipo='Progetto'
                                THEN CAST(T0.progetto AS VARCHAR)
                                ELSE CAST(T2.u_progetto AS VARCHAR)
                           END AS Progetto_commessa,
                           CONCAT(T1.lastname, ' ', SUBSTRING(T1.firstname,1,1)) AS Nome,
                           T0.[Data], COALESCE(T2.itemname,'') AS Nome_macchina,
                           COALESCE(T2.[U_Final_customer_name],'') AS Cliente_finale,
                           T0.[Contenuto], T0.[risolto], A.DOCDUEDATE
                    FROM [TIRELLI_40].[DBO].Appunti_commesse T0
                    LEFT JOIN [TIRELLI_40].[dbo].OHEM T1 ON T0.dipendente=T1.empid
                    LEFT JOIN OITM T2 ON T2.itemcode=T0.commessa
                    LEFT JOIN (
                        SELECT T10.ITEMCODE, T12.DOCNUM, T12.DOCDUEDATE
                        FROM (
                            SELECT MIN(T1.DocEntry) AS DOCENTRY, T0.ITEMCODE
                            FROM RDR1 T0
                            INNER JOIN ORDR T1 ON T0.DocEntry=T1.DocEntry
                            WHERE T1.CANCELED<>'Y' AND SUBSTRING(T0.ITEMCODE,1,1)='M'
                            GROUP BY T0.ITEMCODE
                        ) T10
                        LEFT JOIN RDR1 T11 ON T11.DocEntry=T10.DocEntry AND T10.ITEMCODE=T11.ITEMCODE
                        LEFT JOIN ORDR T12 ON T12.DocEntry=T11.DocEntry
                    ) A ON A.ItemCode=T0.Commessa
                    WHERE 1=1 " & filtro_Dipendente & "
                ) t10
                GROUP BY t10.progetto_commessa
            ) t20
            LEFT JOIN OPMG t21 ON CAST(t20.Progetto_commessa AS VARCHAR)=CAST(T21.DocNum AS VARCHAR)
            LEFT JOIN OCRD t22 ON t22.CardCode=t21.U_Codice_cliente_finale
            ORDER BY t20.data_min", Cnn1)

                Using rdr As SqlDataReader = CMD_SAP_2.ExecuteReader()
                    Dim currentProgetto As String = String.Empty

                    While rdr.Read()
                        Dim progetto As String = "Progetto " & rdr("progetto_commessa").ToString() & " " &
                                             rdr("cardname").ToString() & " " &
                                             rdr("Cliente_finale").ToString() & " " &
                                             rdr("name").ToString()

                        If currentProgetto <> progetto Then
                            If currentProgetto <> String.Empty Then emailBody.AppendLine("<br />")
                            currentProgetto = progetto
                            emailBody.AppendLine("<b>" & progetto & "</b><br /><br />")
                        End If

                        ' Chiamata della sub
                        trova_appunti(rdr("progetto_commessa").ToString(), par_id_dipendente, emailBody, destinatariUnici, colori, colorIndex, coloriTag)
                    End While
                End Using
            End Using

            ' --- Aggiungi destinatari dalla lista di distribuzione ---
            Using CMD_Email As New SqlCommand("
            SELECT Mail
            FROM [Tirelli_40].[dbo].[Lista_distribuzione_riunioni]
            WHERE Nome_lista=@lista", Cnn1)

                CMD_Email.Parameters.AddWithValue("@lista", par_lista_distribuzione)

                Using rdrMail As SqlDataReader = CMD_Email.ExecuteReader()
                    While rdrMail.Read()
                        destinatariUnici.Add(rdrMail("Mail").ToString())
                    End While
                End Using
            End Using
        End Using  ' <-- qui chiudiamo l'uso della connessione

        ' --- Prepara lista finale destinatari ---
        Dim emailList As String = String.Join(";", destinatariUnici)

        ' --- Invia tramite Outlook ---
        Dim outlookApp As New Outlook.Application
        Dim mailItem As Outlook.MailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem)

        With mailItem
            .Subject = "Punti aperti"
            .HTMLBody = emailBody.ToString()
            .To = emailList
            .Display() ' oppure .Send()
        End With
    End Sub




    ' --- Sub per aggiungere contenuti e gestire i tag @ ---
    Sub trova_appunti(par_progetto As String, par_id_dipendente As Integer,
                  ByRef emailBody As StringBuilder,
                  ByRef destinatariUnici As HashSet(Of String),
                  colori() As String, ByRef colorIndex As Integer,
                  coloriTag As Dictionary(Of String, String))

        Dim filtro_Dipendente As String = If(par_id_dipendente = 0, "", " AND t0.dipendente = " & par_id_dipendente)

        Using Cnn1 As New SqlConnection(Homepage.sap_tirelli)
            Cnn1.Open()

            Using CMD_SAP_2 As New SqlCommand("
            SELECT T0.Commessa, COALESCE(T2.itemname,'') AS Nome_macchina, 
                   COALESCE(T2.[U_Final_customer_name],'') AS Cliente_finale,
                   T0.Contenuto, T0.risolto
            FROM [TIRELLI_40].[DBO].Appunti_commesse T0
            LEFT JOIN OITM T2 ON T2.itemcode=T0.commessa
            WHERE (CASE WHEN T0.tipo='Progetto' THEN CAST(T0.commessa AS VARCHAR) ELSE CAST(T2.u_progetto AS VARCHAR) END) = @progetto
            " & filtro_Dipendente & "
            ORDER BY T0.ID", Cnn1)

                CMD_SAP_2.Parameters.AddWithValue("@progetto", par_progetto)

                Using rdr As SqlDataReader = CMD_SAP_2.ExecuteReader()
                    Dim currentCommessa As String = String.Empty

                    While rdr.Read()
                        Dim commessa As String = rdr("Commessa").ToString() & " " &
                                             rdr("Nome_macchina").ToString() & " " &
                                             rdr("Cliente_finale").ToString()

                        If currentCommessa <> commessa Then
                            If currentCommessa <> String.Empty Then emailBody.AppendLine("<br />")
                            currentCommessa = commessa
                            '   emailBody.AppendLine("<b>" & commessa & "</b><br />")

                            emailBody.AppendLine("<b><div style='margin-left:30px;'>" & commessa & "<br /><br /></div> </b>")
                        End If

                        Dim contenuto As String = rdr("Contenuto").ToString()
                        If Convert.ToBoolean(rdr("risolto")) Then contenuto = "<s>" & contenuto & "</s>"

                        ' --- Gestione tag @ senza ByRef nella lambda ---
                        Dim regex As New Regex("@\S+")
                        Dim matches = regex.Matches(contenuto)
                        For Each match As Match In matches
                            Dim nome As String = match.Value.Substring(1).ToLower()
                            Dim indirizzo As String = nome & "@tirelli.net"
                            destinatariUnici.Add(indirizzo)

                            Dim coloreCorrente As String
                            If Not coloriTag.TryGetValue(nome, coloreCorrente) Then
                                coloreCorrente = colori(colorIndex)
                                coloriTag(nome) = coloreCorrente
                                colorIndex = (colorIndex + 1) Mod colori.Length
                            End If

                            contenuto = contenuto.Replace(match.Value,
                                    "<span style='color:" & coloreCorrente & "; font-weight:bold;'>" & match.Value & "</span>")
                        Next

                        emailBody.AppendLine("<div style='margin-left:60px;'>" & contenuto & "<br /><br /></div>")
                    End While
                End Using
            End Using
        End Using
    End Sub









    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        calcola_backlog()
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        'If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Trasferito").Value > 0 Then
        '    MsgBox("Impossibile cancellare riga di un codice che risulta TRASFERITO")
        'Else
        Dim PAR_DATAGRIDVIEW As DataGridView
        PAR_DATAGRIDVIEW = DataGridView5

        ' Supponendo che COLONNAID sia il nome della colonna che vuoi usare per cancellare il commento
        Dim COLONNAID As String = "ID_"
        Dim selectedRow As DataGridViewRow = PAR_DATAGRIDVIEW.CurrentRow

        ' Verifica che ci sia una riga selezionata prima di procedere
        If selectedRow IsNot Nothing Then
            ' Chiede conferma all'utente se vuole cancellare il commento
            Dim result As DialogResult = MessageBox.Show("Sei sicuro di voler cancellare il commento?" & vbCrLf & selectedRow.Cells("Commento_").Value, "Conferma Cancellazione", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            ' Se l'utente conferma, procedi con la cancellazione
            If result = DialogResult.Yes Then
                ' Passa l'ID della riga alla funzione cancella_commento
                cancella_commento(selectedRow.Cells(COLONNAID).Value)

                ' Rimuovi la riga selezionata
                PAR_DATAGRIDVIEW.Rows.RemoveAt(selectedRow.Index)
            End If
        Else
            MessageBox.Show("Seleziona una riga prima di cancellarla.")
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        'If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Trasferito").Value > 0 Then
        '    MsgBox("Impossibile cancellare riga di un codice che risulta TRASFERITO")
        'Else
        Dim PAR_DATAGRIDVIEW As DataGridView
        PAR_DATAGRIDVIEW = DataGridView5

        ' Supponendo che COLONNAID sia il nome della colonna che vuoi usare per cancellare il commento
        Dim COLONNAID As String = "ID_"
        Dim selectedRow As DataGridViewRow = PAR_DATAGRIDVIEW.CurrentRow

        ' Verifica che ci sia una riga selezionata prima di procedere
        If selectedRow IsNot Nothing Then
            ' Chiede conferma all'utente se vuole cancellare il commento

            cambia_stato(selectedRow.Cells(COLONNAID).Value)
            carica_appunti(Label2.Text, "PROGETTO", DataGridView5, "", Homepage.ID_SALVATO)


        Else
            MessageBox.Show("Seleziona una riga prima di CAMBIARNE STATO")
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        modifica(DataGridView5, "ID_", "Commento_")
        carica_appunti(Label4.Text, "PROGETTO", DataGridView5, "", Homepage.ID_SALVATO)
        MsgBox("Aggiornato con successo")
    End Sub

    Public Sub appunti_globali(par_id_dipendente As Integer)


        Dim par_datagridview As DataGridView = DataGridView6
        par_datagridview.Rows.Clear()
        Dim filtro_Dipendente As String = If(par_id_dipendente = 0, "", " AND t0.dipendente = " & par_id_dipendente)



        Using Cnn1 As New SqlConnection(Homepage.sap_tirelli)
            Cnn1.Open()

            ' --- Query progetti ---
            Using CMD_SAP_2 As New SqlCommand("
            SELECT t20.progetto_commessa, t21.cardname,

                   COALESCE(t22.cardname,'') AS Cliente_finale,
case when COALESCE(t22.cardname,'')='' then t21.cardname else COALESCE(t22.cardname,'') end as 'Cliente_def',
                   t21.name, t20.data_min
            FROM (
                SELECT t10.progetto_commessa, MIN(t10.DOCDUEDATE) AS data_min
                FROM (
                    SELECT T0.ID, T0.tipo,
                           CAST(T0.Commessa AS VARCHAR) AS Commessa,
                           CAST(T0.progetto AS VARCHAR) AS Progetto,
                           CASE WHEN T0.tipo='Progetto'
                                THEN CAST(T0.commessa AS VARCHAR)
                                ELSE CAST(T2.u_progetto AS VARCHAR)
                           END AS Progetto_commessa,
                           CONCAT(T1.lastname, ' ', SUBSTRING(T1.firstname,1,1)) AS Nome,
                           T0.[Data], COALESCE(T2.itemname,'') AS Nome_macchina,
                           COALESCE(T2.[U_Final_customer_name],'') AS Cliente_finale,
                           T0.[Contenuto], T0.[risolto], A.DOCDUEDATE
                    FROM [TIRELLI_40].[DBO].Appunti_commesse T0
                    LEFT JOIN [TIRELLI_40].[dbo].OHEM T1 ON T0.dipendente=T1.empid
                    LEFT JOIN OITM T2 ON T2.itemcode=T0.commessa
                    LEFT JOIN (
                        SELECT T10.ITEMCODE, T12.DOCNUM, T12.DOCDUEDATE
                        FROM (
                            SELECT MIN(T1.DocEntry) AS DOCENTRY, T0.ITEMCODE
                            FROM RDR1 T0
                            INNER JOIN ORDR T1 ON T0.DocEntry=T1.DocEntry
                            WHERE T1.CANCELED<>'Y' AND SUBSTRING(T0.ITEMCODE,1,1)='M'
                            GROUP BY T0.ITEMCODE
                        ) T10
                        LEFT JOIN RDR1 T11 ON T11.DocEntry=T10.DocEntry AND T10.ITEMCODE=T11.ITEMCODE
                        LEFT JOIN ORDR T12 ON T12.DocEntry=T11.DocEntry
                    ) A ON A.ItemCode=T0.Commessa
                    WHERE 1=1 " & filtro_Dipendente & "
                ) t10
                GROUP BY t10.progetto_commessa
            ) t20
            LEFT JOIN OPMG t21 ON CAST(t20.Progetto_commessa AS VARCHAR)=CAST(T21.DocNum AS VARCHAR)
            LEFT JOIN OCRD t22 ON t22.CardCode=t21.U_Codice_cliente_finale
            ORDER BY t20.data_min", Cnn1)

                Using rdr As SqlDataReader = CMD_SAP_2.ExecuteReader()
                    Dim currentProgetto As String = String.Empty

                    While rdr.Read()

                        par_datagridview.Rows.Add(0, rdr("progetto_commessa"), rdr("cliente_def"), rdr("name"))

                        appunti_per_commessa(par_id_dipendente, rdr("progetto_commessa"), par_datagridview)
                        'Dim progetto As String = "Progetto " & rdr("progetto_commessa").ToString() & " " &
                        '                     rdr("cardname").ToString() & " " &
                        '                     rdr("Cliente_finale").ToString() & " " &
                        '                     rdr("name").ToString()



                        ' Chiamata della sub
                        ' trova_appunti(rdr("progetto_commessa").ToString(), par_id_dipendente, emailBody, destinatariUnici, colori, colorIndex, coloriTag)
                    End While
                End Using
            End Using


        End Using  ' <-- qui chiudiamo l'uso della connessione


    End Sub

    Public Sub appunti_per_commessa(par_id_dipendente As Integer, par_progetto As String, par_datagridview As DataGridView)



        Dim filtro_Dipendente As String = If(par_id_dipendente = 0, "", " AND t0.dipendente = " & par_id_dipendente)



        Using Cnn1 As New SqlConnection(Homepage.sap_tirelli)
            Cnn1.Open()

            ' --- Query progetti ---
            Using CMD_SAP_2 As New SqlCommand("
            SELECT t0.id,T0.Commessa, COALESCE(T2.itemname,'') AS Nome_macchina, 
                   COALESCE(T2.[U_Final_customer_name],'') AS Cliente_finale,
                   T0.Contenuto, T0.risolto
,CONCAT(T3.lastname, ' ', SUBSTRING(T3.firstname,1,1)) AS dipendente,
t0.data

            FROM [TIRELLI_40].[DBO].Appunti_commesse T0
            LEFT JOIN OITM T2 ON T2.itemcode=T0.commessa
left join [TIRELLI_40].[dbo].ohem t3 on t3.empid=t0.dipendente
            WHERE (CASE WHEN T0.tipo='Progetto' THEN CAST(T0.commessa AS VARCHAR) ELSE CAST(T2.u_progetto AS VARCHAR) END) = '" & par_progetto & "'
            " & filtro_Dipendente & "
            ORDER BY T0.ID", Cnn1)

                Using rdr As SqlDataReader = CMD_SAP_2.ExecuteReader()
                    Dim currentProgetto As String = String.Empty

                    While rdr.Read()

                        par_datagridview.Rows.Add(rdr("ID"), 0, "", "", rdr("Commessa"), rdr("Nome_macchina"), rdr("dipendente"), rdr("data"), rdr("contenuto"), rdr("Risolto"))

                    End While
                End Using
            End Using


        End Using  ' <-- qui chiudiamo l'uso della connessione


    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        appunti_globali(Homepage.ID_SALVATO)
    End Sub

    Private Async Sub appunti_Click(sender As Object, e As EventArgs) Handles Appunti.Enter
        appunti_globali(Homepage.ID_SALVATO)
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub DataGridView6_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellContentClick

    End Sub

    Private Sub DataGridView6_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView6.CellFormatting
        Dim par_datagridview As DataGridView = DataGridView6

        ' Recupera indici colonne
        Dim statoIndex As Integer = par_datagridview.Columns("stato__").Index
        Dim id3Index As Integer = par_datagridview.Columns("ID_3").Index

        ' Verifica se siamo in una riga valida e non è una riga nuova
        If e.RowIndex >= 0 AndAlso Not par_datagridview.Rows(e.RowIndex).IsNewRow Then
            Dim riga As DataGridViewRow = par_datagridview.Rows(e.RowIndex)

            ' Controlla il valore della colonna "stato__"
            Dim statoValue As Boolean = Convert.ToBoolean(riga.Cells(statoIndex).Value)

            ' Controlla il valore della colonna "ID_3"
            Dim id3Value As Integer = Convert.ToInt32(riga.Cells(id3Index).Value)

            ' Font base: Regular o Strikeout
            Dim baseFontStyle As FontStyle = If(statoValue, FontStyle.Strikeout, FontStyle.Regular)

            If id3Value = 0 Then
                ' Riga arancione + grassetto (+ barrato se stato__ = True)
                For Each cell As DataGridViewCell In riga.Cells
                    cell.Style.BackColor = Color.Orange
                    cell.Style.Font = New Font(par_datagridview.Font, baseFontStyle Or FontStyle.Bold)
                Next
            Else
                ' Riga normale (solo barrata se stato__ = True)
                For Each cell As DataGridViewCell In riga.Cells
                    cell.Style.BackColor = par_datagridview.DefaultCellStyle.BackColor
                    cell.Style.Font = New Font(par_datagridview.Font, baseFontStyle)
                Next
            End If
        End If
    End Sub

    Private Sub DataGridView5_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellContentClick

    End Sub

    Private Sub DataGridView5_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView5.CellFormatting
        Dim par_datagridview As DataGridView = DataGridView5
        ' Verifica se la colonna "stato" è presente (sostituisci "stato" con il nome corretto della colonna)
        Dim statoIndex As Integer = par_datagridview.Columns("stato___").Index

        ' Verifica se siamo in una riga valida e non è una riga nuova
        If e.RowIndex >= 0 AndAlso Not par_datagridview.Rows(e.RowIndex).IsNewRow Then
            ' Controlla il valore della colonna "stato"
            Dim statoValue As Boolean = Convert.ToBoolean(par_datagridview.Rows(e.RowIndex).Cells(statoIndex).Value)

            ' Se il valore è "True", applica il font barrato a tutta la riga
            If statoValue Then
                For Each cell As DataGridViewCell In par_datagridview.Rows(e.RowIndex).Cells
                    cell.Style.Font = New Font(par_datagridview.Font, FontStyle.Strikeout)
                Next
            Else
                ' Rimuove il font barrato se non è "True"
                For Each cell As DataGridViewCell In par_datagridview.Rows(e.RowIndex).Cells
                    cell.Style.Font = New Font(par_datagridview.Font, FontStyle.Regular)
                Next
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        'If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Trasferito").Value > 0 Then
        '    MsgBox("Impossibile cancellare riga di un codice che risulta TRASFERITO")
        'Else
        Dim PAR_DATAGRIDVIEW As DataGridView
        PAR_DATAGRIDVIEW = DataGridView6

        ' Supponendo che COLONNAID sia il nome della colonna che vuoi usare per cancellare il commento
        Dim COLONNAID As String = "ID_3"
        Dim selectedRow As DataGridViewRow = PAR_DATAGRIDVIEW.CurrentRow

        ' Verifica che ci sia una riga selezionata prima di procedere
        If selectedRow IsNot Nothing Then
            ' Chiede conferma all'utente se vuole cancellare il commento
            Dim result As DialogResult = MessageBox.Show("Sei sicuro di voler cancellare il commento?" & vbCrLf & selectedRow.Cells("Commento__").Value, "Conferma Cancellazione", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            ' Se l'utente conferma, procedi con la cancellazione
            If result = DialogResult.Yes Then
                ' Passa l'ID della riga alla funzione cancella_commento
                cancella_commento(selectedRow.Cells(COLONNAID).Value)

                ' Rimuovi la riga selezionata
                PAR_DATAGRIDVIEW.Rows.RemoveAt(selectedRow.Index)
            End If
        Else
            MessageBox.Show("Seleziona una riga prima di cancellarla.")
        End If
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem4.Click
        'If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Trasferito").Value > 0 Then
        '    MsgBox("Impossibile cancellare riga di un codice che risulta TRASFERITO")
        'Else
        Dim PAR_DATAGRIDVIEW As DataGridView
        PAR_DATAGRIDVIEW = DataGridView6

        ' Supponendo che COLONNAID sia il nome della colonna che vuoi usare per cancellare il commento
        Dim COLONNAID As String = "ID_3"
        Dim selectedRow As DataGridViewRow = PAR_DATAGRIDVIEW.CurrentRow

        ' Verifica che ci sia una riga selezionata prima di procedere
        If selectedRow IsNot Nothing Then
            ' Chiede conferma all'utente se vuole cancellare il commento

            cambia_stato(selectedRow.Cells(COLONNAID).Value)
            appunti_globali(Homepage.ID_SALVATO)


        Else
            MessageBox.Show("Seleziona una riga prima di CAMBIARNE STATO")
        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        modifica_appunti_globali(DataGridView6)
        MsgBox("Aggiornato con successo")
    End Sub

    Private Sub ContextMenuStrip3_Opening(sender As Object, e As CancelEventArgs) Handles ContextMenuStrip3.Opening

    End Sub
End Class