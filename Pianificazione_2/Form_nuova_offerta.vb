Imports System.Data.SqlClient
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Tirelli.ODP_Form

Public Class Form_nuova_offerta
    Public tipo_offerta As String
    Public numero_documento As Integer
    Public Elenco_arol_branch(1000) As String
    Public Elenco_venditori_arol(1000) As String
    Public Elenco_settori(1000) As String
    Public Elenco_country(1000) As String
    Public Elenco_brand(1000) As String
    Public Elenco_hot(1000) As String
    Public Elenco_incoterms(1000) As String
    Public Elenco_causali(1000) As String
    Public Elenco_tipo_vendita(1000) As String
    Public Elenco_ufficio_competenza(1000) As String
    Public riga_selezionata As Integer
    Public Righe_cancellate(100) As Riga_cancellata
    Public num_righe_cancellate As Integer = 0
    Public visorder_selezionato As Integer
    Public itemcode_riga As String
    Public max_linenum As Integer = 0
    Public max_visorder As Integer = 0
    Public listino As Integer

    Public c As Integer = 0
    Public contatore As Integer

    Public quantità_riga As String
    Public prezzo_unitario_riga As String
    Public sconto_riga As String = 0
    Public iniziazione As Integer = 0

    Public tabella_intestazione As String
    Public tabella_righe As String
    Public tipo_documento As String
    Public docentry As Integer


    Public Sel_Stampante As New PrintDialog

    Public Preview As New PrintPreviewDialog

    Public altezza_scontrino As Integer = 700
    'Public altezza_scontrino As Integer = 300
    Public larghezza_scontrino As Integer = 185

    Private valore_15 As Integer
    Private valore_16 As Integer



    Public Structure Riga_cancellata
        Public Codice_riga As String
        Public linenum As Integer
        Public visorder As Integer

    End Structure

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Sub inizializzazione_form(par_numero_documento As Integer, par_tabella_intestazione As String, par_tabella_righe As String, par_tipo_documento As String)
        numero_documento = par_numero_documento
        tabella_intestazione = par_tabella_intestazione
        tabella_righe = par_tabella_righe
        tipo_documento = par_tipo_documento
        iniziazione = 0
        num_righe_cancellate = 0
        Inserimento_items_combobox_arolbranch()
        Inserimento_items_combobox_venditoriarol()
        Inserimento_items_combobox_settore()
        Inserimento_items_combobox_country()
        Inserimento_items_combobox_brand()
        Inserimento_items_combobox_UFFICIO_COMPETENZA()
        Inserimento_items_combobox_hot()
        Inserimento_items_combobox_incoterms()
        Inserimento_items_combobox_causale()
        Inserimento_items_combobox_tipo_vendita()
        intestazioni_offerta(par_numero_documento, par_tabella_intestazione)
        TROVA_MAX_LINENUM_E_MAX_VISORDER(par_numero_documento, par_tabella_intestazione, par_tabella_righe)
        riempi_datagridview_offerta(par_numero_documento, par_tabella_intestazione, par_tabella_righe)
        iniziazione = 1
        Label17.Text = par_tipo_documento
    End Sub

    Sub intestazioni_offerta(par_numero_documento As Integer, par_tabella_intestazione As String)
        Dim valuta As String
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1



        CMD_SAP_1.CommandText = "select 
t0.docentry,
t0.docnum
,t0.docdate
,t0.docduedate
,t0.docstatus
,t0.cardcode
,t0.cardname
,t2.lastname
, case when t1.cardcode is null then '' else t1.cardcode end as 'Cardcode_f'
, case when t1.cardname is null then '' else t1.cardname end as 'Cardname_F'
,t0.doccur
,t0.doctotal,
coalesce(t0.u_matrcds,'') as 'u_matrcds',
case when t0.doccur='USD' then t0.docrate else 1 end as 'docrate'
,case when t0.doccur='USD' then t0.doctotal*t0.docrate  end as 'Total_$'
, case when t3.name is null then '' else t3.name end as 'u_arolbranch'
, case when t4.name is null then '' else t4.name end as 'u_venditorearol'
,case when t0.u_settore is null then '' else t0.u_settore end as 'u_settore'
,case when t0.u_destinazione is null then '' else t0.u_destinazione end as 'U_destinazione'
,case when t0.u_brand is null then '' else t0.u_brand end as 'U_brand'
, case when t5.descr is null then '' else t5.descr end as 'ufficio_competenza' 
, case when t6.descr is null then '' else t6.descr end as 'u_hot' 

, case when t7.descr is null then '' else t7.descr end as 'incoterm' 
, case when t8.descr is null then '' else t8.descr end as 'causcons' 
, case when t9.descr is null then '' else t9.descr end as 'u_categoria' 


,case when t0.doccur='EUR' THEN (T0.DocTotal -t0.vatsum)/(1-  t0.discprcnt/100)  else (T0.DocTotalFC - t0.vatsum*t0.docrate)/(1-  t0.discprcnt/100) end AS 'Totale',

Case when t0.discprcnt is null then 0 else t0.discprcnt end as 'Sconto',
case when t0.doccur ='EUR' then t0.discsum else T0.DISCSUMFC end as 'Valore sconto',
case when t0.doccur='EUR' then t0.vatsum else t0.vatsum*t0.docrate end as 'IVA',
case when t0.doccur='EUR' then t0.paidtodate else t0.paidtodate*t0.docrate end as 'Importo_pagato',
case when t0.doccur='EUR' then t0.doctotal  else T0.DocTotalFC end as 'Totale netto',
case when t0.doccur='EUR' then t0.doctotal-t0.paidtodate else T0.DocTotalFC-t0.paidtodate*t0.docrate end as 'Saldo_in_scadenza'
, case when t0.comments is null then '' else t0.comments end as 'comments'
,coalesce(T0.U_Commento_interno,'') as 'Commento_interno'
,t10.listnum

from " & par_tabella_intestazione & " t0 left join ocrd t1 on t0.U_CodiceBP=t1.cardcode
left join [TIRELLI_40].[dbo].ohem t2 on t2.empid=t0.ownercode
left join [dbo].[@AROL_BRANCH]  T3 on t3.code=t0.u_arolbranch
left join [dbo].[@VENDITORI_AROL]  T4 on t4.code=t0.u_venditorearol
left join UFD1 T5 on t5.tableid='OQUT' and t5.fieldid=108 and t5.fldvalue=t0.u_uffcompetenza
left join UFD1 T6  on T6.tableid='OQUT'  and t6.fieldid=129 and t6.fldvalue=t0.u_hot
left join UFD1 T7 on T7.tableid='OQUT' and t7.fieldid=66 and t7.fldvalue=t0.u_prg_azs_incoterms
left join UFD1 T8 on T8.tableid='OQUT' and t8.fieldid=44 and t8.fldvalue=t0.u_causcons
left join UFD1 T9 on T9.tableid='OQUT' and t9.fieldid=153 and t9.fldvalue=t0.u_categoria
left join ocrd t10 on t10.cardcode=t0.cardcode

where t0.docnum= " & par_numero_documento & " 
"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        If cmd_SAP_reader_1.Read() Then
            TextBox1.Text = cmd_SAP_reader_1("cardcode")
            TextBox7.Text = cmd_SAP_reader_1("cardname")
            Data_ordine.Value = cmd_SAP_reader_1("docdate")
            DateTimePicker1.Value = cmd_SAP_reader_1("docduedate")
            TextBox17.Text = cmd_SAP_reader_1("u_matrcds")
            TextBox2.Text = cmd_SAP_reader_1("cardcode_f")
            TextBox8.Text = cmd_SAP_reader_1("Cardname_F")
            ComboBox1.Text = cmd_SAP_reader_1("doccur")
            TextBox3.Text = cmd_SAP_reader_1("docrate")
            ComboBox10.Text = cmd_SAP_reader_1("u_arolbranch")
            ComboBox11.Text = cmd_SAP_reader_1("u_venditorearol")
            ComboBox12.Text = cmd_SAP_reader_1("u_settore")
            ComboBox13.Text = cmd_SAP_reader_1("u_destinazione")
            ComboBox14.Text = cmd_SAP_reader_1("u_brand")
            TextBox6.Text = cmd_SAP_reader_1("docstatus")
            ComboBox6.Text = cmd_SAP_reader_1("ufficio_competenza")
            ComboBox7.Text = cmd_SAP_reader_1("u_hot")
            ComboBox8.Text = cmd_SAP_reader_1("incoterm")
            ComboBox2.Text = cmd_SAP_reader_1("causcons")
            ComboBox3.Text = cmd_SAP_reader_1("u_categoria")
            RichTextBox1.Text = cmd_SAP_reader_1("comments")
            RichTextBox2.Text = cmd_SAP_reader_1("commento_interno")
            listino = cmd_SAP_reader_1("listnum")
            docentry = cmd_SAP_reader_1("docentry")

            If cmd_SAP_reader_1("doccur") = "USD" Then
                valuta = "$"
            Else
                valuta = "€"
            End If


            TextBox4.Text = String.Format("{0:N2}", cmd_SAP_reader_1("Totale"))

            TextBox5.Text = String.Format("{0:N2}", cmd_SAP_reader_1("Sconto"))
            TextBox9.Text = String.Format("{0:N2}", cmd_SAP_reader_1("valore sconto"))

            TextBox12.Text = String.Format("{0:N2}", cmd_SAP_reader_1("Iva"))
            TextBox11.Text = String.Format("{0:N2}", cmd_SAP_reader_1("Totale netto"))

            TextBox13.Text = String.Format("{0:N2}", cmd_SAP_reader_1("Importo_pagato"))
            TextBox14.Text = String.Format("{0:N2}", cmd_SAP_reader_1("Saldo_in_scadenza"))



        End If

        Cnn1.Close()

    End Sub

    Sub Inserimento_items_combobox_arolbranch()

        ComboBox10.Items.Clear()
        Dim indice As Integer
        indice = 0
        ComboBox10.Items.Add("")
        indice = indice + 1
        ComboBox10.SelectedIndex = 0
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.Code, T0.Name FROM [dbo].[@AROL_BRANCH]  T0 "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            Elenco_arol_branch(indice) = cmd_SAP_reader("code")

            ComboBox10.Items.Add(cmd_SAP_reader("name"))

            indice = indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Sub Inserimento_items_combobox_venditoriarol()

        ComboBox11.Items.Clear()
        Dim indice As Integer
        indice = 0
        ComboBox11.Items.Add("")
        indice = indice + 1
        ComboBox11.SelectedIndex = 0

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.Code, T0.Name FROM [dbo].[@VENDITORI_AROL]  T0 "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            Elenco_venditori_arol(indice) = cmd_SAP_reader("code")

            ComboBox11.Items.Add(cmd_SAP_reader("name"))



            indice = indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Sub Inserimento_items_combobox_settore()
        ComboBox12.Items.Clear()
        Dim indice As Integer
        indice = 0
        ComboBox12.Items.Add("")
        indice = indice + 1
        ComboBox12.SelectedIndex = 0
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select t1.fldvalue,t1.descr from UFD1 T1
 WHERE T1.tableid='OQUT' AND FIELDID=124"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            Elenco_settori(indice) = cmd_SAP_reader("fldvalue")

            ComboBox12.Items.Add(cmd_SAP_reader("descr"))



            indice = indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Sub Inserimento_items_combobox_country()

        ComboBox13.Items.Clear()
        Dim indice As Integer
        indice = 0
        ComboBox13.Items.Add("")
        indice = indice + 1
        ComboBox13.SelectedIndex = 0

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.Code, T0.Name, T0.U_NumCode FROM [dbo].[@BNCCRY]  T0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            Elenco_country(indice) = cmd_SAP_reader("code")

            ComboBox13.Items.Add(cmd_SAP_reader("name"))



            indice = indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Sub Inserimento_items_combobox_brand()

        ComboBox14.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select t1.fldvalue,t1.descr from UFD1 T1
 WHERE T1.tableid='OQUT'  and t1.fieldid=163"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim indice As Integer
        indice = 0
        Do While cmd_SAP_reader.Read()

            Elenco_brand(indice) = cmd_SAP_reader("fldvalue")

            ComboBox14.Items.Add(cmd_SAP_reader("descr"))



            indice = indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Sub Inserimento_items_combobox_UFFICIO_COMPETENZA()

        ComboBox6.Items.Clear()
        Dim indice As Integer
        indice = 0
        ComboBox6.Items.Add("")
        indice = indice + 1
        ComboBox6.SelectedIndex = 0

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select  t1.fldvalue,t1.descr, t1.fieldid from UFD1 T1
 WHERE T1.tableid='OQUT'  and t1.fieldid=108"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            Elenco_ufficio_competenza(indice) = cmd_SAP_reader("fldvalue")

            ComboBox6.Items.Add(cmd_SAP_reader("descr"))

            indice = indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Sub Inserimento_items_combobox_hot()



        ComboBox7.Items.Clear()
        Dim indice As Integer
        indice = 0
        ComboBox11.Items.Add(0)
        indice = indice + 1
        ComboBox11.SelectedIndex = 0
        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select  t1.fldvalue,t1.descr, t1.fieldid from UFD1 T1
 WHERE T1.tableid='OQUT'  and t1.fieldid=129"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            Elenco_hot(indice) = cmd_SAP_reader("fldvalue")

            ComboBox7.Items.Add(cmd_SAP_reader("descr"))

            indice = indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Sub Inserimento_items_combobox_incoterms()

        ComboBox8.Items.Clear()
        Dim indice As Integer
        indice = 0
        ComboBox8.Items.Add("")
        indice = indice + 1
        ComboBox8.SelectedIndex = 0
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select  t1.fldvalue,t1.descr, t1.fieldid from UFD1 T1
 WHERE T1.tableid='OQUT' and t1.fieldid=66 "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            Elenco_incoterms(indice) = cmd_SAP_reader("fldvalue")

            ComboBox8.Items.Add(cmd_SAP_reader("descr"))

            indice = indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Sub Inserimento_items_combobox_causale()

        ComboBox2.Items.Clear()
        Dim indice As Integer
        indice = 0
        ComboBox2.Items.Add("")
        indice = indice + 1
        ComboBox2.SelectedIndex = 0


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select  t1.fldvalue,t1.descr, t1.fieldid from UFD1 T1
 WHERE T1.tableid='OQUT'   and t1.fieldid=44 "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            Elenco_causali(indice) = cmd_SAP_reader("fldvalue")

            ComboBox2.Items.Add(cmd_SAP_reader("descr"))

            indice = indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Sub Inserimento_items_combobox_tipo_vendita()

        ComboBox3.Items.Clear()
        Dim indice As Integer
        indice = 0
        ComboBox3.Items.Add("")
        indice = indice + 1
        ComboBox3.SelectedIndex = 0
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select  t1.fldvalue,t1.descr, t1.fieldid from UFD1 T1
 WHERE T1.tableid='OQUT'   and t1.fieldid=153 "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            Elenco_tipo_vendita(indice) = cmd_SAP_reader("fldvalue")

            ComboBox3.Items.Add(cmd_SAP_reader("descr"))

            indice = indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Private Sub Form_nuova_offerta_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox4.Text = Strings.FormatCurrency(12345.67, 2)
        TextBox9.Text = Strings.FormatCurrency(12345.67, 2)
        TextBox11.Text = Strings.FormatCurrency(12345.67, 2)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        aggiorna_off_doc()
        cancella_righe_DB()
        aggiornamento_righe()
        riempi_datagridview_offerta(numero_documento, tabella_intestazione, tabella_righe)
        MsgBox("Documento aggiornato con successo")
    End Sub


    Sub aggiorna_off_doc()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "update t0 set t0.cardcode='" & TextBox1.Text & "'
,t0.cardname='" & TextBox7.Text & "'
,t0.DocCur='" & ComboBox1.Text & "'
,t0.DocRate='" & Replace(TextBox3.Text, ",", ".") & "'
,t0.U_AROLBranch='" & Elenco_arol_branch(ComboBox10.SelectedIndex) & "'
,t0.U_VenditoreArol='" & Elenco_venditori_arol(ComboBox11.SelectedIndex) & "'
,t0.U_Settore='" & Elenco_settori(ComboBox12.SelectedIndex) & "'
,t0.U_Destinazione='" & Elenco_country(ComboBox13.SelectedIndex) & "'
, t0.U_Brand='" & Elenco_brand(ComboBox14.SelectedIndex) & "'
, t0.U_Uffcompetenza='" & Elenco_ufficio_competenza(ComboBox6.SelectedIndex) & "'
,t0.U_PRG_AZS_Incoterms='" & Elenco_incoterms(ComboBox8.SelectedIndex) & "'
,t0.u_hot='" & Elenco_hot(ComboBox7.SelectedIndex) & "'
,t0.U_CausCons='" & Elenco_causali(ComboBox2.SelectedIndex) & "'
,t0.U_Tipologia_vendita='" & Elenco_tipo_vendita(ComboBox3.SelectedIndex) & "'
,t0.Comments='" & RichTextBox1.Text & "'
,t0.DiscPrcnt='" & Replace(TextBox5.Text, ",", ".") & "'
,t0.DiscSum='" & Replace(TextBox9.Text, ",", ".") & "'
, t0.doctotal=" & Replace(Replace(TextBox11.Text, ".", ""), ",", ".") & "*" & Replace(TextBox3.Text, ",", ".") & "
,t0.u_codicebp='" & TextBox2.Text & "'
,t0.u_clientefinale='" & TextBox8.Text & "'

from OQUT t0
where t0.docnum='" & TextBox10.Text & "'"

        CMD_SAP.ExecuteNonQuery()



        cnn.Close()

    End Sub

    Sub riempi_datagridview_offerta(par_numero_documento As Integer, par_tabella_intestazione As String, par_tabella_righe As String)


        DataGridView_offerta.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1



        CMD_SAP_2.CommandText = "select T10.itemtype, t10.Tipo,T10.LineNum, t10.visorder, T10.ItemCode,t10.[Dscription],t10.u_disegno,  T10.Quantity,T10.Price, T10.Currency, T10.Rate, T10.DiscPrcnt, case when t10.rate <>0 then T10.LineTotal*t10.rate else t10.linetotal end as 'linetotal' ,T10.PriceBefDi, T10.U_PRG_AZS_Commessa, t10.freetxt, t10.leadtime,t10.u_al_mag_non_wip,t10.U_approvvigionamento_articolo
 ,  sum(case when t11.onhand is null then 0 else t11.onhand end -case when t11.iscommited is null then 0 else t11.iscommited end +case when t11.onorder is null then 0 else t11.onorder end) as 'disponibile',t10.u_descing,
 t10.ocrcode, t10.whscode, t10.U_Trasferito,t10.U_Datrasferire
from
(
SELECT  T1.ITEMTYPE, 'Articolo' as 'Tipo',T1.LineNum, t1.visorder, T1.ItemCode,t1.[Dscription],t1.u_disegno,  T1.Quantity,T1.Price, T1.Currency, T1.Rate, T1.DiscPrcnt, T1.LineTotal ,T1.PriceBefDi, T1.U_PRG_AZS_Commessa, t1.freetxt, t3.leadtime, t1.u_al_mag_non_wip,t1.U_approvvigionamento_articolo, t3.u_descing
,t1.ocrcode, t1.whscode, t1.U_Trasferito,t1.U_Datrasferire
 FROM " & par_tabella_intestazione & " T0  INNER JOIN " & par_tabella_righe & " T1 ON T0.[DocEntry] = T1.[DocEntry] 
left join oitm t3 on t3.itemcode=t1.itemcode
WHERE T0.[DocNum] ='" & par_numero_documento & "'
)
as t10 inner join oitw t11 on t11.itemcode=t10.itemcode
group by T10.itemtype, t10.Tipo,T10.LineNum, t10.visorder, T10.ItemCode,t10.[Dscription],t10.u_disegno,  T10.Quantity,T10.Price, T10.Currency, T10.Rate, T10.DiscPrcnt, T10.LineTotal ,T10.PriceBefDi, T10.U_PRG_AZS_Commessa, t10.freetxt, t10.leadtime,t10.u_al_mag_non_wip,t10.U_approvvigionamento_articolo,t10.u_descing, t10.ocrcode, t10.whscode, t10.U_Trasferito,t10.U_Datrasferire"



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()

            DataGridView_offerta.Rows.Add(1, cmd_SAP_reader_2("Itemtype"), cmd_SAP_reader_2("tipo"), cmd_SAP_reader_2("itemcode"), cmd_SAP_reader_2("Dscription"), cmd_SAP_reader_2("u_Descing"), cmd_SAP_reader_2("u_Disegno"), cmd_SAP_reader_2("Quantity"), cmd_SAP_reader_2("pricebefdi"), cmd_SAP_reader_2("DiscPrcnt"), cmd_SAP_reader_2("price"), cmd_SAP_reader_2("linetotal"), cmd_SAP_reader_2("leadtime"), cmd_SAP_reader_2("u_al_mag_non_wip"), cmd_SAP_reader_2("disponibile"), cmd_SAP_reader_2("u_approvvigionamento_articolo"), cmd_SAP_reader_2("U_PRG_AZS_Commessa"), cmd_SAP_reader_2("freetxt"), cmd_SAP_reader_2("linenum"), cmd_SAP_reader_2("visorder"), cmd_SAP_reader_2("ocrcode"), cmd_SAP_reader_2("whscode"), cmd_SAP_reader_2("U_Trasferito"), cmd_SAP_reader_2("U_Datrasferire"))

        Loop



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        DataGridView_offerta.ClearSelection()

    End Sub

    Private Sub DeleteRowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteRowToolStripMenuItem.Click
        If DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Presente").Value <> 0 Then
            Righe_cancellate(num_righe_cancellate).Codice_riga = DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Codice").Value
            Righe_cancellate(num_righe_cancellate).linenum = DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Linenum").Value
            Righe_cancellate(num_righe_cancellate).visorder = DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Visorder").Value
            num_righe_cancellate = num_righe_cancellate + 1
        End If
        DataGridView_offerta.Rows.RemoveAt(riga_selezionata)
        aggiorna_prezzo_totale()

    End Sub



    Private Sub DataGridView_offerta_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_offerta.CellClick

        If e.RowIndex >= 0 Then


            If e.ColumnIndex = DataGridView_offerta.Columns.IndexOf(Disegno) Then

                Magazzino.visualizza_disegno(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Disegno").Value)

            End If

            DataGridView_offerta.SelectionMode = DataGridViewSelectionMode.CellSelect
            visorder_selezionato = DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Visorder").Value
            riga_selezionata = e.RowIndex
        End If

    End Sub

    Private Sub ViewCodeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ViewCodeToolStripMenuItem.Click


        Magazzino.Codice_SAP = DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Codice").Value
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
    End Sub



    Private Sub DataGridView_offerta_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_offerta.CellValueChanged
        If e.RowIndex >= 0 Then


            If e.ColumnIndex = DataGridView_offerta.Columns.IndexOf(Codice) Then


                ' Try
                itemcode_riga = UCase(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Codice").Value)
                informazioni_articolo_riga()
                '  Catch ex As Exception
                ' MsgBox("C'è un errore nell'articolo riga")
                ' End Try

            ElseIf e.ColumnIndex = DataGridView_offerta.Columns.IndexOf(Quantità) Or e.ColumnIndex = DataGridView_offerta.Columns.IndexOf(Prezzo_unitario) Or e.ColumnIndex = DataGridView_offerta.Columns.IndexOf(Discount) Then




                If InStr(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Quantità").Value, ",") > 1 Then


                    quantità_riga = LSet(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Quantità").Value, InStr(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Quantità").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Quantità").Value), InStr(StrReverse(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Quantità").Value), ",") - 1))



                Else
                    quantità_riga = Replace(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Quantità").Value, ",", ".")
                End If



                If InStr(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Prezzo_unitario").Value, ",") > 1 Then


                    ' prezzo_unitario_riga = LSet(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="prezzo_unitario").Value, InStr(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="prezzo_unitario").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="prezzo_unitario").Value), InStr(StrReverse(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="prezzo_unitario").Value), ",") - 1)) * LSet(TextBox3.Text, InStr(TextBox3.Text, ",") - 1) & "." & StrReverse(LSet(StrReverse(TextBox3.Text), InStr(StrReverse(TextBox3.Text), ",") - 1))

                    prezzo_unitario_riga = LSet(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="prezzo_unitario").Value, InStr(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="prezzo_unitario").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="prezzo_unitario").Value), InStr(StrReverse(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="prezzo_unitario").Value), ",") - 1))


                Else
                    'prezzo_unitario_riga = Replace(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="prezzo_unitario").Value, ",", ".") * LSet(TextBox3.Text, InStr(TextBox3.Text, ",") - 1) & "." & StrReverse(LSet(StrReverse(TextBox3.Text), InStr(StrReverse(TextBox3.Text), ",") - 1))
                    prezzo_unitario_riga = Replace(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="prezzo_unitario").Value, ",", ".")



                End If


                If InStr(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Discount").Value, ",") > 1 Then


                    sconto_riga = LSet(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Discount").Value, InStr(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Discount").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Discount").Value), InStr(StrReverse(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Discount").Value), ",") - 1))

                Else
                    sconto_riga = Replace(DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Discount").Value, ",", ".")
                End If

                Dim rate As String


                rate = Replace(TextBox4.Text, ",", ".")



                Dim calcolo_sconto_riga As String

                If sconto_riga = Nothing Then
                    calcolo_sconto_riga = 0
                Else calcolo_sconto_riga = Replace((100 - sconto_riga) / 100, ",", ".")
                End If
                Dim Cnn2 As New SqlConnection
                Cnn2.ConnectionString = Homepage.sap_tirelli

                cnn2.Open()

                Dim CMD_SAP_2 As New SqlCommand
                Dim cmd_SAP_reader_2 As SqlDataReader
                CMD_SAP_2.Connection = cnn2

                If prezzo_unitario_riga = Nothing Then
                    prezzo_unitario_riga = 1
                End If


                CMD_SAP_2.CommandText = "SELECT  " & prezzo_unitario_riga & "*" & calcolo_sconto_riga & " As 'Prezzo_unitario_scontato', " & quantità_riga & " * " & prezzo_unitario_riga & "*" & calcolo_sconto_riga & " As 'Totale'  "

                cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
                If cmd_SAP_reader_2.Read() = True Then

                    DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="unit_price_bef_disc").Value = cmd_SAP_reader_2("Prezzo_unitario_scontato")
                    DataGridView_offerta.Rows(e.RowIndex).Cells(columnName:="Totale").Value = cmd_SAP_reader_2("Totale")



                End If
                cmd_SAP_reader_2.Close()
                cnn2.Close()





            End If
        End If

        aggiorna_prezzo_totale()





    End Sub

    Sub informazioni_articolo_riga()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "select t10.itemcode,t10.objtype,t10.tipo_articolo,t10.itemname,t10.DfltWH,t10.price,t10.validFor,t10.LeadTime,t10.u_descing,t10.disponibile, sum (t11.onhand) as 'Al_mag_non_wip'
from
(
Select t0.itemcode,Case When T2.[VisResCode] Is null Then T0.objTYPE Else '290' end as 'objtype',
Case When T2.[VisResCode] Is null Then 'Articolo' Else 'Risorsa' end as 'Tipo_Articolo',
T0.[ItemName], case when T0.[DfltWH] is null then '01' else T0.[DfltWH] end as 'DfltWH', T1.[Price]  as 'price', T0.VALIDFOR 
,t0.leadtime, t0.u_descing
, sum(case when t3.onhand is null then 0 else t3.onhand end -case when t3.iscommited is null then 0 else t3.iscommited end +case when t3.onorder is null then 0 else t3.onorder end) as 'disponibile'


FROM OITM T0  INNER JOIN ITM1 T1 ON T0.[ItemCode] = T1.[ItemCode] 
left join orsc t2 on T2.[VisResCode]=t0.itemcode
LEFT JOIN OITW T3 ON t3.itemcode=t0.itemcode
WHERE T0.[ItemCode] ='" & itemcode_riga & "' AND  T1.[PriceList] =" & listino & "
group by t0.itemcode,t2.VisResCode,t0.ObjType,t0.itemname,t0.DfltWH,t1.Price,t0.validFor,t0.LeadTime,t0.U_DESCING
)
as t10 left join oitw t11 on t10.itemcode=t11.itemcode and t11.whscode<>'WIP'
group by t10.itemcode,t10.objtype,t10.tipo_articolo,t10.itemname,t10.DfltWH,t10.price,t10.validFor,t10.LeadTime,t10.u_descing,t10.disponibile"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() = True Then


            If cmd_SAP_reader("VALIDFOR") = "N" Then
                MsgBox("Inactive code")
            Else

                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Descrizione").Value = cmd_SAP_reader("ItemName")
                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Itemtype").Value = cmd_SAP_reader("objTYPE")
                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="quantità").Value = 1


                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="prezzo_unitario").Value = cmd_SAP_reader("price") * TextBox3.Text

                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="discount").Value = 0
                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="unit_price_bef_disc").Value = cmd_SAP_reader("price") * TextBox3.Text
                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="totale").Value = cmd_SAP_reader("price") * TextBox3.Text
                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="lead_time").Value = cmd_SAP_reader("leadtime")
                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Eng_Description").Value = cmd_SAP_reader("u_descing")
                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Disp").Value = cmd_SAP_reader("disponibile")
                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="On_stock").Value = cmd_SAP_reader("al_mag_non_wip")


                If DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Presente").Value = 1 Then
                    DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Presente").Value = 2
                ElseIf DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Presente").Value <> 2 Then
                    DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Presente").Value = 0
                End If


                If DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Quantità").Value <= DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="On_stock").Value And DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Quantità").Value <= DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Disp").Value Then
                    DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Status").Value = "ON STOCK"
                ElseIf DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="lead_time").Value + 2 <= 5 Then
                    DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Status").Value = "5"
                ElseIf DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="lead_time").Value + 2 <= 10 Then
                    DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Status").Value = "10"
                ElseIf DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="lead_time").Value <= 30 Then
                    DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Status").Value = cmd_SAP_reader("leadtime")
                Else DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="Status").Value = "TBD"

                End If

                't1.quantity <= t1.U_Al_mag_non_wip And t1.quantity <= t1.U_Disponibile Then 'ON STOCK' when t1.U_Lead_time+2<=5 then cast(5 as varchar)  when t1.U_Lead_time+2<=10 then cast(10 as varchar) when t1.U_Lead_time<=30 then cast(round(t1.U_Lead_time, -1) as varchar) else 'TBD' end as 'Stato_approv'


                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="linenum").Value = max_linenum
                DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="visorder").Value = max_visorder


                max_linenum = max_linenum + 1
                max_visorder = max_visorder + 1
                aggiorna_prezzo_totale()
            End If
        Else
            MsgBox("Articolo non esistente")

        End If

        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub

    Sub TROVA_MAX_LINENUM_E_MAX_VISORDER(par_numero_documento As Integer, par_tabella_intestazione As String, par_tabella_righe As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "select max(t0.linenum)+1 as 'Max_linenum', max(t0.visorder)+1 as  'Max_visorder' 
from " & par_tabella_righe & " t0 inner join " & par_tabella_intestazione & " t1 on t0.docentry=t1.docentry
where t1.docnum=" & par_numero_documento & ""

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then

            max_linenum = cmd_SAP_reader("max_linenum")
            max_visorder = cmd_SAP_reader("Max_visorder")

        End If
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Dim r_che_deve_salire = DataGridView_offerta.Rows(riga_selezionata)
        Dim r_che_deve_scendere = DataGridView_offerta.Rows(riga_selezionata - 1)
        DataGridView_offerta.Rows.Remove(r_che_deve_salire)
        DataGridView_offerta.Rows.Remove(r_che_deve_scendere)

        DataGridView_offerta.Rows.Insert(riga_selezionata - 1, r_che_deve_salire)
        DataGridView_offerta.Rows.Insert(riga_selezionata, r_che_deve_scendere)

        '
        DataGridView_offerta.ClearSelection()
        DataGridView_offerta.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView_offerta.Rows(riga_selezionata - 1).Selected = True

        Dim visorder_che_deve_salire As Integer
        Dim visorder_che_deve_scendere As Integer

        visorder_che_deve_salire = DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="visorder").Value
        visorder_che_deve_scendere = DataGridView_offerta.Rows(riga_selezionata - 1).Cells(columnName:="visorder").Value

        DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="visorder").Value = visorder_che_deve_scendere
        DataGridView_offerta.Rows(riga_selezionata - 1).Cells(columnName:="visorder").Value = visorder_che_deve_salire


        If DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="presente").Value = 1 Then

            DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="presente").Value = 2
        ElseIf DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="presente").Value <> 2 Then
            DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="presente").Value = 0
        End If

        If DataGridView_offerta.Rows(riga_selezionata - 1).Cells(columnName:="presente").Value = 1 Then

            DataGridView_offerta.Rows(riga_selezionata - 1).Cells(columnName:="presente").Value = 2
        ElseIf DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="presente").Value <> 2 Then
            DataGridView_offerta.Rows(riga_selezionata - 1).Cells(columnName:="presente").Value = 0
        End If



        riga_selezionata = riga_selezionata - 1
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim r_che_deve_salire = DataGridView_offerta.Rows(riga_selezionata + 1)
        Dim r_che_deve_scendere = DataGridView_offerta.Rows(riga_selezionata)
        DataGridView_offerta.Rows.Remove(r_che_deve_salire)
        DataGridView_offerta.Rows.Remove(r_che_deve_scendere)

        DataGridView_offerta.Rows.Insert(riga_selezionata, r_che_deve_salire)
        DataGridView_offerta.Rows.Insert(riga_selezionata + 1, r_che_deve_scendere)

        '
        DataGridView_offerta.ClearSelection()
        DataGridView_offerta.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView_offerta.Rows(riga_selezionata + 1).Selected = True

        Dim visorder_che_deve_salire As Integer
        Dim visorder_che_deve_scendere As Integer

        visorder_che_deve_salire = DataGridView_offerta.Rows(riga_selezionata + 1).Cells(columnName:="visorder").Value
        visorder_che_deve_scendere = DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="visorder").Value

        DataGridView_offerta.Rows(riga_selezionata + 1).Cells(columnName:="visorder").Value = visorder_che_deve_scendere
        DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="visorder").Value = visorder_che_deve_salire




        If DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="presente").Value = 1 Then

            DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="presente").Value = 2
        ElseIf DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="presente").Value <> 2 Then
            DataGridView_offerta.Rows(riga_selezionata).Cells(columnName:="presente").Value = 0
        End If

        If DataGridView_offerta.Rows(riga_selezionata + 1).Cells(columnName:="presente").Value = 1 Then

            DataGridView_offerta.Rows(riga_selezionata + 1).Cells(columnName:="presente").Value = 2
        ElseIf DataGridView_offerta.Rows(riga_selezionata + 1).Cells(columnName:="presente").Value <> 2 Then
            DataGridView_offerta.Rows(riga_selezionata + 1).Cells(columnName:="presente").Value = 0
        End If




        riga_selezionata = riga_selezionata + 1
    End Sub

    Sub cancella_righe_DB()
        c = 0
        Do While c < num_righe_cancellate
            Dim Cnn3 As New SqlConnection
            Cnn3.ConnectionString = Homepage.sap_tirelli
            cnn3.Open()

            Dim CMD_SAP_3 As New SqlCommand

            CMD_SAP_3.Connection = cnn3


            CMD_SAP_3.CommandText = "delete t1 from oqut t0 inner join qut1 t1 on t0.docnum='" & TextBox10.Text & "' and t1.itemcode='" & Righe_cancellate(c).Codice_riga & "' and  t1.linenum='" & Righe_cancellate(c).linenum & "'"

            CMD_SAP_3.ExecuteNonQuery()
            cnn3.Close()
            c = c + 1
        Loop
        c = 0
    End Sub

    Sub aggiornamento_righe()
        contatore = 0
        Do While contatore <= DataGridView_offerta.Rows.Count - 2


            If DataGridView_offerta.Rows(contatore).Cells(columnName:="Presente").Value = 0 Or DataGridView_offerta.Rows(contatore).Cells(columnName:="Presente").Value = 2 Then
                inserisci_riga_offerta()


            End If
            contatore = contatore + 1
        Loop
    End Sub

    Sub inserisci_riga_offerta()

        itemcode_riga = DataGridView_offerta.Rows(contatore).Cells(columnName:="Codice").Value
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand
        Dim cmd_SAP_reader_7 As SqlDataReader
        CMD_SAP_7.Connection = cnn

        CMD_SAP_7.CommandText = "SELECT T1.VALIDFOR AS 'Valido', t1.itemcode as 'Codice' 

FROM OITM T1 WHERE T1.[itemcode]= '" & itemcode_riga & "'"

        cmd_SAP_reader_7 = CMD_SAP_7.ExecuteReader
        If cmd_SAP_reader_7.Read() = True Then
            If cmd_SAP_reader_7("Valido") = "N" Then
                MsgBox("Il codice " & itemcode_riga & " is inactive ")
            Else
                cnn.Close()
                If DataGridView_offerta.Rows(contatore).Cells(columnName:="Presente").Value = 0 Then
                    'DA INSERIRE
                    inserisci_riga()
                ElseIf DataGridView_offerta.Rows(contatore).Cells(columnName:="Presente").Value = 2 Then

                    'DA INSERIRE 

                    'aggiorna_riga()
                End If


            End If

        Else
            MsgBox("Il codice " & itemcode_riga & " non esiste ")
        End If
        cnn.Close()



    End Sub

    Sub inserisci_riga()
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = cnn3


        CMD_SAP_3.CommandText = "insert into qut1  (QUT1.DocEntry, QUT1.LineNum,QUT1.VISORDER, QUT1.TargetType,
 QUT1.BaseRef,
QUT1.BaseType, 
QUT1.LineStatus,
QUT1.ItemCode,
QUT1.Dscription,
QUT1.Quantity,
 QUT1.OpenQty,
QUT1.Price,
QUT1.Currency,
QUT1.Rate,
QUT1.DiscPrcnt,
QUT1.LineTotal,
QUT1.TotalFrgn,
QUT1.OpenSum,
QUT1.OpenSumFC,
QUT1.VendorNum,
QUT1.WhsCode,
QUT1.TreeType,
QUT1.AcctCode,
QUT1.TaxStatus,
QUT1.PriceBefDi,
QUT1.DocDate,
QUT1.OpenCreQty,
QUT1.UseBaseUn,
QUT1.BaseCard,
qut1.[NumPerMsr]
, QUT1.VATGROUP
,qut1.[InvQty]
,QUT1.[OpenInvQty]
,QUT1.[PcQuantity]
,QUT1.[PackQty]
,QUT1.[UomCode]
,QUT1.[unitMsr]

)

SELECT top 1
T1.DocEntry, " & DataGridView_offerta.Rows(contatore).Cells(columnName:="linenum").Value & ",
" & DataGridView_offerta.Rows(contatore).Cells(columnName:="VISORDER").Value & ",
-1,
0,
-1,
'O',
'" & DataGridView_offerta.Rows(contatore).Cells(columnName:="Codice").Value & "',
'" & DataGridView_offerta.Rows(contatore).Cells(columnName:="descrizione").Value & "',
" & DataGridView_offerta.Rows(contatore).Cells(columnName:="Quantità").Value & ",
" & DataGridView_offerta.Rows(contatore).Cells(columnName:="Quantità").Value & ",
" & Replace(DataGridView_offerta.Rows(contatore).Cells(columnName:="Prezzo_unitario").Value, ",", ".") & ",
'" & ComboBox1.Text & "',
case when t1.currency='USD' then '" & Replace(TextBox3.Text, ",", ".") & "' else 0 end,
" & Replace(DataGridView_offerta.Rows(contatore).Cells(columnName:="Discount").Value, ",", ".") & ",
" & Replace(DataGridView_offerta.Rows(contatore).Cells(columnName:="Totale").Value, ",", ".") & ",
case when t1.currency='EUR' then 0 else " & Replace(TextBox3.Text, ",", ".") & "*" & Replace(DataGridView_offerta.Rows(contatore).Cells(columnName:="Totale").Value, ",", ".") & " end,
" & Replace(DataGridView_offerta.Rows(contatore).Cells(columnName:="Totale").Value, ",", ".") & ",
case when t1.currency='EUR' then 0 else " & Replace(TextBox3.Text, ",", ".") & "*" & Replace(DataGridView_offerta.Rows(contatore).Cells(columnName:="Totale").Value, ",", ".") & " end,
T2.SUPPCATNUM,
CASE WHEN T2.DFLTWH IS NULL THEN '01' ELSE T2.DFLTWH END,
'N',
CASE WHEN T4.[GroupCode] ='100' THEN T3.REVENUESAC WHEN T4.[GroupCode] ='103' THEN T3.EURevenuAc ELSE T3.FrRevenuAc END,
'Y',
" & Replace(DataGridView_offerta.Rows(contatore).Cells(columnName:="Prezzo_unitario").Value, ",", ".") & ",
CONVERT(DATETIME, '" & Data_ordine.Value & "', 103),
1,
'N',
'" & TextBox1.Text & "',
1,
CASE WHEN T4.[GroupCode] ='100' THEN 'V22' WHEN T4.[GroupCode] ='103' THEN 'CB' ELSE 'N12' END,
" & DataGridView_offerta.Rows(contatore).Cells(columnName:="Quantità").Value & ",
" & DataGridView_offerta.Rows(contatore).Cells(columnName:="Quantità").Value & ",
" & DataGridView_offerta.Rows(contatore).Cells(columnName:="Quantità").Value & ",
" & DataGridView_offerta.Rows(contatore).Cells(columnName:="Quantità").Value & ",
'Manuale',
CASE WHEN T2.SALUNITMSR IS NULL THEN 'PZ' ELSE T2.SALUNITMSR END



FROM OQUT T0  INNER JOIN QUT1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
LEFT JOIN OITM T2 ON T2.ITEMCODE=T1.ITEMCODE
LEFT JOIN OITB T3 ON T3.ItmsGrpCod=T2.ItmsGrpCod
LEFT JOIN OCRD T4 ON T4.CARDCODE=T0.CARDCODE
WHERE T0.[DocNum] =" & TextBox10.Text & ""


        'Da inserire
        ', , QUT1.TotalSumSy, QUT1.OpenSumSys, QUT1.InvntSttus, QUT1.OcrCode, QUT1.Project, QUT1.CodeBars, QUT1.VatPrcnt, QUT1.VatGroup, QUT1.PriceAfVAT, QUT1.Height1, QUT1.Hght1Unit, QUT1.Height2, QUT1.Hght2Unit, QUT1.Width1, QUT1.Wdth1Unit, QUT1.Width2, QUT1.Wdth2Unit, QUT1.Length1, QUT1.Len1Unit, QUT1.length2, QUT1.Len2Unit, QUT1.Volume, QUT1.VolUnit, QUT1.Weight1, QUT1.Wght1Unit, QUT1.Weight2, QUT1.Wght2Unit, QUT1.Factor1, QUT1.Factor2, QUT1.Factor3, QUT1.Factor4, QUT1.PackQty, QUT1.UpdInvntry, QUT1.BaseDocNum, QUT1.BaseAtCard, QUT1.SWW, QUT1.VatSum, QUT1.VatSumFrgn, QUT1.VatSumSy, QUT1.FinncPriod, QUT1.ObjType, QUT1.BlockNum, QUT1.ImportLog, QUT1.DedVatSum, QUT1.DedVatSumF, QUT1.DedVatSumS, QUT1.IsAqcuistn, QUT1.DistribSum, QUT1.DstrbSumFC, QUT1.DstrbSumSC, QUT1.GrssProfit, QUT1.GrssProfSC, QUT1.GrssProfFC, QUT1.VisOrder, QUT1.INMPrice, QUT1.PoTrgNum, QUT1.PoTrgEntry, QUT1.DropShip, QUT1.PoLineNum, QUT1.Address, QUT1.TaxCode, QUT1.TaxType, QUT1.OrigItem, QUT1.BackOrdr, QUT1.FreeTxt, QUT1.PickStatus, QUT1.PickOty, QUT1.PickIdNo, QUT1.TrnsCode, QUT1.VatAppld, QUT1.VatAppldFC, QUT1.VatAppldSC, QUT1.BaseQty, QUT1.BaseOpnQty, QUT1.VatDscntPr, QUT1.WtLiable, QUT1.DeferrTax, QUT1.EquVatPer, QUT1.EquVatSum, QUT1.EquVatSumF, QUT1.EquVatSumS, QUT1.LineVat, QUT1.LineVatlF, QUT1.LineVatS, QUT1.unitMsr, QUT1.NumPerMsr, QUT1.CEECFlag, QUT1.ToStock, QUT1.ToDiff, QUT1.ExciseAmt, QUT1.TaxPerUnit, QUT1.TotInclTax, QUT1.CountryOrg, QUT1.StckDstSum, QUT1.ReleasQtty, QUT1.LineType, QUT1.TranType, QUT1.Text, QUT1.OwnerCode, QUT1.StockPrice, QUT1.ConsumeFCT, QUT1.LstByDsSum, QUT1.StckINMPr, QUT1.LstBINMPr, QUT1.StckDstFc, QUT1.StckDstSc, QUT1.LstByDsFc, QUT1.LstByDsSc, QUT1.StockSum, QUT1.StockSumFc, QUT1.StockSumSc, QUT1.StckSumApp, QUT1.StckAppFc, QUT1.StckAppSc, QUT1.ShipToCode, QUT1.ShipToDesc, QUT1.StckAppD, QUT1.StckAppDFC, QUT1.StckAppDSC, QUT1.BasePrice, QUT1.GTotal, QUT1.GTotalFC, QUT1.GTotalSC, QUT1.DistribExp, QUT1.DescOW, QUT1.DetailsOW, QUT1.GrossBase, QUT1.VatWoDpm, QUT1.VatWoDpmFc, QUT1.VatWoDpmSc, QUT1.CFOPCode, QUT1.CSTCode, QUT1.Usage, QUT1.TaxOnly, QUT1.WtCalced, QUT1.QtyToShip, QUT1.DelivrdQty, QUT1.OrderedQty, QUT1.CogsOcrCod, QUT1.CiOppLineN, QUT1.CogsAcct, QUT1.ChgAsmBoMW, QUT1.ActDelDate, QUT1.OcrCode2, QUT1.OcrCode3, QUT1.OcrCode4, QUT1.OcrCode5, QUT1.TaxDistSum, QUT1.TaxDistSFC, QUT1.TaxDistSSC, QUT1.PostTax, QUT1.Excisable, QUT1.AssblValue, QUT1.RG23APart1, QUT1.RG23APart2, QUT1.RG23CPart1, QUT1.RG23CPart2, QUT1.CogsOcrCo2, QUT1.CogsOcrCo3, QUT1.CogsOcrCo4, QUT1.CogsOcrCo5, QUT1.LnExcised, QUT1.LocCode, QUT1.StockValue, QUT1.GPTtlBasPr, QUT1.unitMsr2, QUT1.NumPerMsr2, QUT1.SpecPrice, QUT1.CSTfIPI, QUT1.CSTfPIS, QUT1.CSTfCOFINS, QUT1.ExLineNo, QUT1.isSrvCall, QUT1.PQTReqQty, QUT1.PQTReqDate, QUT1.PcDocType, QUT1.PcQuantity, QUT1.LinManClsd, QUT1.VatGrpSrc, QUT1.NoInvtryMv, QUT1.ActBaseEnt, QUT1.ActBaseLn, QUT1.ActBaseNum, QUT1.OpenRtnQty, QUT1.AgrNo, QUT1.AgrLnNum, QUT1.CredOrigin, QUT1.Surpluses, QUT1.DefBreak, QUT1.Shortages, QUT1.UomEntry, QUT1.UomEntry2, QUT1.UomCode, QUT1.UomCode2, QUT1.FromWhsCod, QUT1.NeedQty, QUT1.PartRetire, QUT1.RetireQty, QUT1.RetireAPC, QUT1.RetirAPCFC, QUT1.RetirAPCSC, QUT1.InvQty, QUT1.OpenInvQty, QUT1.EnSetCost, QUT1.RetCost, QUT1.Incoterms, QUT1.TransMod, QUT1.LineVendor, QUT1.DistribIS, QUT1.ISDistrb, QUT1.ISDistrbFC, QUT1.ISDistrbSC, QUT1.IsByPrdct, QUT1.ItemType, QUT1.PriceEdit, QUT1.PrntLnNum, QUT1.LinePoPrss, QUT1.FreeChrgBP, QUT1.TaxRelev, QUT1.LegalText, QUT1.ThirdParty, QUT1.LicTradNum, QUT1.InvQtyOnly, QUT1.UnencReasn, QUT1.ShipFromCo, QUT1.ShipFromDe, QUT1.FisrtBin, QUT1.AllocBinC, QUT1.ExpType, QUT1.ExpUUID, QUT1.ExpOpType, QUT1.DIOTNat, QUT1.MYFtype, QUT1.GPBefDisc, QUT1.ReturnRsn, QUT1.ReturnAct, QUT1.StgSeqNum, QUT1.StgEntry, QUT1.StgDesc, QUT1.ItmTaxType, QUT1.SacEntry, QUT1.NCMCode, QUT1.HsnEntry, QUT1.OriBAbsEnt, QUT1.OriBLinNum, QUT1.OriBDocTyp, QUT1.IsPrscGood, QUT1.IsCstmAct, QUT1.EncryptIV, QUT1.ExtTaxRate, QUT1.ExtTaxSum, QUT1.TaxAmtSrc, QUT1.ExtTaxSumF, QUT1.ExtTaxSumS, QUT1.StdItemId, QUT1.CommClass, QUT1.VatExEntry, QUT1.VatExLN, QUT1.NatOfTrans, QUT1.ISDtCryImp, QUT1.ISDtRgnImp, QUT1.ISOrCryExp, QUT1.ISOrRgnExp, QUT1.NVECode, QUT1.PoNum, QUT1.PoItmNum, QUT1.IndEscala, QUT1.CESTCode, QUT1.CtrSealQty, QUT1.CNJPMan, QUT1.UFFiscBene, QUT1.CUSplit, QUT1.LegalTIMD, QUT1.LegalTTCA, QUT1.LegalTW, QUT1.LegalTCD, QUT1.RevCharge, QUT1.U_BLD_LyID, QUT1.U_BLD_NCps, QUT1.U_O01FlagU, QUT1.U_O01ProAg, QUT1.U_O01ProCA, QUT1.U_O01ProCZ, QUT1.U_O01ProDI, QUT1.U_O01PrzGr, QUT1.U_O01ScoIm, QUT1.U_BnTrian, QUT1.U_Note, QUT1.U_TrasMgEM, QUT1.U_Totval, QUT1.U_BNIncTrm, QUT1.U_BNTrnMod, QUT1.U_TestoDOC, QUT1.U_QtySup, QUT1.U_PRG_AZS_OpDocEntry, QUT1.U_PRG_AZS_OpLineNum, QUT1.U_TpForn, QUT1.U_PRG_AZS_DescrAlt, QUT1.U_PRG_AZS_PrevMPS, QUT1.U_PRG_AZS_StatoComm, QUT1.U_Colli, QUT1.U_PRG_AZS_OcDocEntry, QUT1.U_PRG_AZS_OcDocNum, QUT1.U_PRG_AZS_OcLineNum, QUT1.U_PRG_AZS_OaDocEntry, QUT1.U_PRG_AZS_OaDocNum, QUT1.U_PRG_AZS_OaLineNum, QUT1.U_Datitecncompl, QUT1.U_UTdatainiz, QUT1.U_UTfineprog, QUT1.U_inizioassel, QUT1.U_Fineassel, QUT1.U_inizassmecc, QUT1.U_fineassmecc, QUT1.U_PRG_AZS_OpDocNum, QUT1.U_PRG_AZS_Commessa, QUT1.U_PRG_AZS_NumAtCard, QUT1.U_PRG_AZS_DataRic, QUT1.U_PRG_AZS_DataCon, QUT1.U_PRG_AZS_PrzProForma, QUT1.U_PRG_CLV_PrzPia, QUT1.U_PRG_CLV_PrzLav, QUT1.U_PRG_CVM_DocAssoc, QUT1.U_B1SYS_Discount, QUT1.U_B1SYS_Discount_FC, QUT1.U_B1SYS_Discount_SC, QUT1.U_B1SYS_DiscountVat, QUT1.U_B1SYS_DiscountVtFC, QUT1.U_B1SYS_DiscountVtSC, QUT1.U_Inizcol, QUT1.U_Finecol, QUT1.U_Fineapp, QUT1.U_mod_macchina, QUT1.U_Fine_app_MU, QUT1.U_Inizio_ass_EL, QUT1.U_Fine_ass_EL, QUT1.U_Inizioapprovvigionamento, QUT1.U_DataKOM, QUT1.U_PListinoAcqu, QUT1.U_Ultimoprezzodeterminato, QUT1.U_Migliorprezzo, QUT1.U_Migliorfornitore, QUT1.U_Trasferito, QUT1.U_Datrasferire, QUT1.U_Almag01, QUT1.U_AlmagCDS, QUT1.U_Opportunita, QUT1.U_Ubicazione, QUT1.U_O01Sc1, QUT1.U_O01Sc2, QUT1.U_O01Sc3, QUT1.U_O01Sc4, QUT1.U_O01Sc5, QUT1.U_O01Sc6, QUT1.U_Ricarico, QUT1.U_Prezzoarolbranch, QUT1.U_Commissione_agente, QUT1.U_Costo, QUT1.U_Data_scheda_tecnica, QUT1.U_Data_clean_order, QUT1.U_Disegno, QUT1.U_Produttore, QUT1.U_Revisione, QUT1.U_PRG_AZS_UbiDest, QUT1.U_PRG_AZS_PrjFather, QUT1.U_PRG_AZS_QtaEvasa, QUT1.U_PRG_WIP_QtaRichMagAuto, QUT1.U_PRG_QLT_QCDlnQty, QUT1.U_PRG_QLT_QCCntQty, QUT1.U_PRG_QLT_QCNCResE, QUT1.U_PRG_QLT_QCNCResM, QUT1.U_PRG_QLT_HasTC, QUT1.U_PRG_WMS_Exp, QUT1.U_PRG_WMS_ExpDate, QUT1.U_PRG_WMS_MdMovQty, QUT1.U_Coefficiente_vendita, QUT1.U_Gestito_Ferretto, QUT1.U_Mag_ferretto, QUT1.U_Fase, QUT1.U_Accettato, QUT1.U_Costo_aggiornato, QUT1.U_Lead_time, QUT1.U_Disponibile, QUT1.U_Al_mag_non_wip, QUT1.U_approvvigionamento_articolo, QUT1.U_Made_in, QUT1.U_Forzare_documento
        ' T1.VendorNum, T1.SerialNum, T1.WhsCode, T1.SlpCode, T1.Commission, T1.TreeType, T1.AcctCode, T1.TaxStatus, T1.GrossBuyPr, T1.PriceBefDi, T1.DocDate, T1.OpenCreQty, T1.UseBaseUn, T1.SubCatNum, T1.BaseCard, T1.TotalSumSy, T1.OpenSumSys, T1.InvntSttus, T1.OcrCode, T1.Project, T1.CodeBars, T1.VatPrcnt, T1.VatGroup, T1.PriceAfVAT, T1.Height1, T1.Hght1Unit, T1.Height2, T1.Hght2Unit, T1.Width1, T1.Wdth1Unit, T1.Width2, T1.Wdth2Unit, T1.Length1, T1.Len1Unit, T1.length2, T1.Len2Unit, T1.Volume, T1.VolUnit, T1.Weight1, T1.Wght1Unit, T1.Weight2, T1.Wght2Unit, T1.Factor1, T1.Factor2, T1.Factor3, T1.Factor4, T1.PackQty, T1.UpdInvntry, T1.BaseDocNum, T1.BaseAtCard, T1.SWW, T1.VatSum, T1.VatSumFrgn, T1.VatSumSy, T1.FinncPriod, T1.ObjType, T1.BlockNum, T1.ImportLog, T1.DedVatSum, T1.DedVatSumF, T1.DedVatSumS, T1.IsAqcuistn, T1.DistribSum, T1.DstrbSumFC, T1.DstrbSumSC, T1.GrssProfit, T1.GrssProfSC, T1.GrssProfFC, T1.VisOrder, T1.INMPrice, T1.PoTrgNum, T1.PoTrgEntry, T1.DropShip, T1.PoLineNum, T1.Address, T1.TaxCode, T1.TaxType, T1.OrigItem, T1.BackOrdr, T1.FreeTxt, T1.PickStatus, T1.PickOty, T1.PickIdNo, T1.TrnsCode, T1.VatAppld, T1.VatAppldFC, T1.VatAppldSC, T1.BaseQty, T1.BaseOpnQty, T1.VatDscntPr, T1.WtLiable, T1.DeferrTax, T1.EquVatPer, T1.EquVatSum, T1.EquVatSumF, T1.EquVatSumS, T1.LineVat, T1.LineVatlF, T1.LineVatS, T1.unitMsr, T1.NumPerMsr, T1.CEECFlag, T1.ToStock, T1.ToDiff, T1.ExciseAmt, T1.TaxPerUnit, T1.TotInclTax, T1.CountryOrg, T1.StckDstSum, T1.ReleasQtty, T1.LineType, T1.TranType, T1.Text, T1.OwnerCode, T1.StockPrice, T1.ConsumeFCT, T1.LstByDsSum, T1.StckINMPr, T1.LstBINMPr, T1.StckDstFc, T1.StckDstSc, T1.LstByDsFc, T1.LstByDsSc, T1.StockSum, T1.StockSumFc, T1.StockSumSc, T1.StckSumApp, T1.StckAppFc, T1.StckAppSc, T1.ShipToCode, T1.ShipToDesc, T1.StckAppD, T1.StckAppDFC, T1.StckAppDSC, T1.BasePrice, T1.GTotal, T1.GTotalFC, T1.GTotalSC, T1.DistribExp, T1.DescOW, T1.DetailsOW, T1.GrossBase, T1.VatWoDpm, T1.VatWoDpmFc, T1.VatWoDpmSc, T1.CFOPCode, T1.CSTCode, T1.Usage, T1.TaxOnly, T1.WtCalced, T1.QtyToShip, T1.DelivrdQty, T1.OrderedQty, T1.CogsOcrCod, T1.CiOppLineN, T1.CogsAcct, T1.ChgAsmBoMW, T1.ActDelDate, T1.OcrCode2, T1.OcrCode3, T1.OcrCode4, T1.OcrCode5, T1.TaxDistSum, T1.TaxDistSFC, T1.TaxDistSSC, T1.PostTax, T1.Excisable, T1.AssblValue, T1.RG23APart1, T1.RG23APart2, T1.RG23CPart1, T1.RG23CPart2, T1.CogsOcrCo2, T1.CogsOcrCo3, T1.CogsOcrCo4, T1.CogsOcrCo5, T1.LnExcised, T1.LocCode, T1.StockValue, T1.GPTtlBasPr, T1.unitMsr2, T1.NumPerMsr2, T1.SpecPrice, T1.CSTfIPI, T1.CSTfPIS, T1.CSTfCOFINS, T1.ExLineNo, T1.isSrvCall, T1.PQTReqQty, T1.PQTReqDate, T1.PcDocType, T1.PcQuantity, T1.LinManClsd, T1.VatGrpSrc, T1.NoInvtryMv, T1.ActBaseEnt, T1.ActBaseLn, T1.ActBaseNum, T1.OpenRtnQty, T1.AgrNo, T1.AgrLnNum, T1.CredOrigin, T1.Surpluses, T1.DefBreak, T1.Shortages, T1.UomEntry, T1.UomEntry2, T1.UomCode, T1.UomCode2, T1.FromWhsCod, T1.NeedQty, T1.PartRetire, T1.RetireQty, T1.RetireAPC, T1.RetirAPCFC, T1.RetirAPCSC, T1.InvQty, T1.OpenInvQty, T1.EnSetCost, T1.RetCost, T1.Incoterms, T1.TransMod, T1.LineVendor, T1.DistribIS, T1.ISDistrb, T1.ISDistrbFC, T1.ISDistrbSC, T1.IsByPrdct, T1.ItemType, T1.PriceEdit, T1.PrntLnNum, T1.LinePoPrss, T1.FreeChrgBP, T1.TaxRelev, T1.LegalText, T1.ThirdParty, T1.LicTradNum, T1.InvQtyOnly, T1.UnencReasn, T1.ShipFromCo, T1.ShipFromDe, T1.FisrtBin, T1.AllocBinC, T1.ExpType, T1.ExpUUID, T1.ExpOpType, T1.DIOTNat, T1.MYFtype, T1.GPBefDisc, T1.ReturnRsn, T1.ReturnAct, T1.StgSeqNum, T1.StgEntry, T1.StgDesc, T1.ItmTaxType, T1.SacEntry, T1.NCMCode, T1.HsnEntry, T1.OriBAbsEnt, T1.OriBLinNum, T1.OriBDocTyp, T1.IsPrscGood, T1.IsCstmAct, T1.EncryptIV, T1.ExtTaxRate, T1.ExtTaxSum, T1.TaxAmtSrc, T1.ExtTaxSumF, T1.ExtTaxSumS, T1.StdItemId, T1.CommClass, T1.VatExEntry, T1.VatExLN, T1.NatOfTrans, T1.ISDtCryImp, T1.ISDtRgnImp, T1.ISOrCryExp, T1.ISOrRgnExp, T1.NVECode, T1.PoNum, T1.PoItmNum, T1.IndEscala, T1.CESTCode, T1.CtrSealQty, T1.CNJPMan, T1.UFFiscBene, T1.CUSplit, T1.LegalTIMD, T1.LegalTTCA, T1.LegalTW, T1.LegalTCD, T1.RevCharge, T1.U_BLD_LyID, T1.U_BLD_NCps, T1.U_O01FlagU, T1.U_O01ProAg, T1.U_O01ProCA, T1.U_O01ProCZ, T1.U_O01ProDI, T1.U_O01PrzGr, T1.U_O01ScoIm, T1.U_BnTrian, T1.U_Note, T1.U_TrasMgEM, T1.U_Totval, T1.U_BNIncTrm, T1.U_BNTrnMod, T1.U_TestoDOC, T1.U_QtySup, T1.U_PRG_AZS_OpDocEntry, T1.U_PRG_AZS_OpLineNum, T1.U_TpForn, T1.U_PRG_AZS_DescrAlt, T1.U_PRG_AZS_PrevMPS, T1.U_PRG_AZS_StatoComm, T1.U_Colli, T1.U_PRG_AZS_OcDocEntry, T1.U_PRG_AZS_OcDocNum, T1.U_PRG_AZS_OcLineNum, T1.U_PRG_AZS_OaDocEntry, T1.U_PRG_AZS_OaDocNum, T1.U_PRG_AZS_OaLineNum, T1.U_Datitecncompl, T1.U_UTdatainiz, T1.U_UTfineprog, T1.U_inizioassel, T1.U_Fineassel, T1.U_inizassmecc, T1.U_fineassmecc, T1.U_PRG_AZS_OpDocNum, T1.U_PRG_AZS_Commessa, T1.U_PRG_AZS_NumAtCard, T1.U_PRG_AZS_DataRic, T1.U_PRG_AZS_DataCon, T1.U_PRG_AZS_PrzProForma, T1.U_PRG_CLV_PrzPia, T1.U_PRG_CLV_PrzLav, T1.U_PRG_CVM_DocAssoc, T1.U_B1SYS_Discount, T1.U_B1SYS_Discount_FC, T1.U_B1SYS_Discount_SC, T1.U_B1SYS_DiscountVat, T1.U_B1SYS_DiscountVtFC, T1.U_B1SYS_DiscountVtSC, T1.U_Inizcol, T1.U_Finecol, T1.U_Fineapp, T1.U_mod_macchina, T1.U_Fine_app_MU, T1.U_Inizio_ass_EL, T1.U_Fine_ass_EL, T1.U_Inizioapprovvigionamento, T1.U_DataKOM, T1.U_PListinoAcqu, T1.U_Ultimoprezzodeterminato, T1.U_Migliorprezzo, T1.U_Migliorfornitore, T1.U_Trasferito, T1.U_Datrasferire, T1.U_Almag01, T1.U_AlmagCDS, T1.U_Opportunita, T1.U_Ubicazione, T1.U_O01Sc1, T1.U_O01Sc2, T1.U_O01Sc3, T1.U_O01Sc4, T1.U_O01Sc5, T1.U_O01Sc6, T1.U_Ricarico, T1.U_Prezzoarolbranch, T1.U_Commissione_agente, T1.U_Costo, T1.U_Data_scheda_tecnica, T1.U_Data_clean_order, T1.U_Disegno, T1.U_Produttore, T1.U_Revisione, T1.U_PRG_AZS_UbiDest, T1.U_PRG_AZS_PrjFather, T1.U_PRG_AZS_QtaEvasa, T1.U_PRG_WIP_QtaRichMagAuto, T1.U_PRG_QLT_QCDlnQty, T1.U_PRG_QLT_QCCntQty, T1.U_PRG_QLT_QCNCResE, T1.U_PRG_QLT_QCNCResM, T1.U_PRG_QLT_HasTC, T1.U_PRG_WMS_Exp, T1.U_PRG_WMS_ExpDate, T1.U_PRG_WMS_MdMovQty, T1.U_Coefficiente_vendita, T1.U_Gestito_Ferretto, T1.U_Mag_ferretto, T1.U_Fase, T1.U_Accettato, T1.U_Costo_aggiornato, T1.U_Lead_time, T1.U_Disponibile, T1.U_Al_mag_non_wip, T1.U_approvvigionamento_articolo, T1.U_Made_in, T1.U_Forzare_documento


        CMD_SAP_3.ExecuteNonQuery()
        cnn3.Close()


    End Sub

    Sub aggiorna_prezzo_totale()
        If iniziazione = 1 Then


            Dim totale_doc As Double = 0
            Dim sconto_doc As Double = 0
            Dim colonna As Integer = DataGridView_offerta.Columns.IndexOf(Totale)

            For Each row As DataGridViewRow In DataGridView_offerta.Rows

                totale_doc += Convert.ToDouble(DataGridView_offerta.Rows(row.Index).Cells(colonna).Value)

            Next
            TextBox4.Text = String.Format("{0:N2}", totale_doc)

            sconto_doc = totale_doc * (Convert.ToDouble(TextBox5.Text) / 100)
            TextBox9.Text = sconto_doc
            TextBox11.Text = String.Format("{0:N2}", totale_doc - sconto_doc)
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Try
            TextBox9.Text = String.Format("{0:N2}", TextBox4.Text * TextBox5.Text / 100)
            TextBox11.Text = TextBox4.Text - TextBox9.Text
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Not (Char.IsDigit(e.KeyChar) Or e.KeyChar = "."c Or e.KeyChar = vbBack) Then

            e.Handled = True
        End If
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If Not (Char.IsDigit(e.KeyChar) Or e.KeyChar = "."c Or e.KeyChar = vbBack) Then

            e.Handled = True
        End If
    End Sub

    Private Sub TextBox5_Leave(sender As Object, e As EventArgs) Handles TextBox5.Leave

        Dim num As Double
        If Double.TryParse(TextBox5.Text, num) Then
            If num < 0 Or num > 100 Then
                MessageBox.Show("the discount must be between 0 e 100.")
                TextBox5.Focus()
            End If
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If RadioButton1.Checked = True Then
            Layout_documenti.Show()
            Layout_documenti.ComboBox1.SelectedIndex = 0
            Layout_documenti.TextBox1.Text = TextBox10.Text
        Else
            Fun_Stampa()

        End If


    End Sub

    Private Sub TrasferimentoDiMagazzinoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TrasferimentoDiMagazzinoToolStripMenuItem.Click
        If tabella_intestazione <> "ORDR" Then


            MsgBox("Funzione disponibile solo per ordini cliente")
        ElseIf TextBox6.Text <> "O" Then
            MsgBox("Lo stato dell'ordine deve essere aperto")

        Else

            Dim par_documento As String = ""
            If tabella_intestazione = "ORDR" Then
                par_documento = "OC"

            End If

            Trasferimento_magazzino.docentry_oc = docentry
            Trasferimento_magazzino.docnum_oc = TextBox10.Text
            Trasferimento_magazzino.inizializzazione_trasferimento(0, docentry, "Trasferimento", par_documento)

        End If
    End Sub

    Private Sub ContextMenuStrip2_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip2.Opening

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub Panel9_Paint(sender As Object, e As PaintEventArgs) Handles Panel9.Paint

    End Sub

    Private Sub ResoMagazzinoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResoMagazzinoToolStripMenuItem.Click
        If tabella_intestazione <> "ORDR" Then


            MsgBox("Funzione disponibile solo per ordini cliente")
        ElseIf TextBox6.Text <> "O" Then
            MsgBox("Lo stato dell'ordine deve essere aperto")

        Else

            Dim par_documento As String = ""
            If tabella_intestazione = "ORDR" Then
                par_documento = "OC"

            End If

            Trasferimento_magazzino.docentry_oc = docentry
            Trasferimento_magazzino.docnum_oc = TextBox10.Text
            Trasferimento_magazzino.inizializzazione_trasferimento(0, docentry, "Reso", par_documento)
        End If
    End Sub

    Private Sub DataGridView_offerta_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_offerta.CellContentClick

    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Sub Fun_Stampa()


        Dim preview_scontrino As Boolean

        If CheckBox1.Checked = True Then
            preview_scontrino = True
        Else
            preview_scontrino = False
        End If


        Sel_Stampante.AllowSomePages = False
        Sel_Stampante.ShowHelp = False
        Sel_Stampante.Document = Scontrino

        If preview_scontrino = True Then
            If Homepage.Stampante_Selezionata = False Then
                Dim result As DialogResult = Sel_Stampante.ShowDialog()
                If (result = DialogResult.OK) Then
                    Homepage.Stampante_Selezionata = True
                    ' Utilizza un PrintPreviewDialog per mostrare l'anteprima di stampa


                    Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", larghezza_scontrino, altezza_scontrino)
                    Dim previewDialog As New PrintPreviewDialog()
                    previewDialog.Document = Scontrino
                    result = previewDialog.ShowDialog()
                End If
            Else
                Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", larghezza_scontrino, altezza_scontrino)
                Dim previewDialog As New PrintPreviewDialog()
                previewDialog.Document = Scontrino
                Dim result As DialogResult = previewDialog.ShowDialog()
            End If
        Else
            If Homepage.Stampante_Selezionata = False Then
                Dim result As DialogResult = Sel_Stampante.ShowDialog()
                If (result = DialogResult.OK) Then
                    Homepage.Stampante_Selezionata = True
                    ' Utilizza un PrintPreviewDialog per mostrare l'anteprima di stampa
                    Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", larghezza_scontrino, altezza_scontrino)
                    Dim previewDialog As New PrintPreviewDialog()
                    previewDialog.Document = Scontrino
                    Scontrino.Print()
                End If
            Else
                Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", larghezza_scontrino, altezza_scontrino)
                Scontrino.Print()
            End If
        End If


    End Sub

    Private Sub Scontrino_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles Scontrino.PrintPage

        Dim Penna As New Pen(Color.Black)
        Dim Carattere_ODP As New Font("Calibri", 16, FontStyle.Bold)
        Dim Carattere_Matricola As New Font("Calibri", 25, FontStyle.Bold)
        Dim Carattere_numerone As New Font("Calibri", 60, FontStyle.Bold)
        Dim Carattere_Descrizione As New Font("Calibri", 20, FontStyle.Bold)
        Dim Carattere_Codice As New Font("Calibri", 22, FontStyle.Italic)
        Dim Carattere_Descrizione_Articolo As New Font("Calibri", 8, FontStyle.Italic)
        Dim Carattere_Qta As New Font("Calibri", 12, FontStyle.Bold)
        Dim Carattere_Ubicazione As New Font("Calibri", 12, FontStyle.Italic)
        Dim Carattere_posizione As New Font("Calibri", 16, FontStyle.Bold)
        Dim Carattere_Diciture As New Font("Calibri", 10, FontStyle.Italic)



        Dim Carattere_Matricola_labelling As New Font("Calibri", 25, FontStyle.Bold)
        Dim Carattere_numerone_labelling As New Font("Calibri", 70, FontStyle.Bold)
        Dim Carattere_descrizione_labelling As New Font("Calibri", 10, FontStyle.Italic)


        With e.Graphics
            ' Imposta la qualità grafica
            .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

            ' Salva lo stato grafico attuale
            Dim state2 As Drawing2D.GraphicsState = .Save()

            ' Imposta la rotazione per l'intero disegno (ruota di 90 gradi in senso antiorario)
            .RotateTransform(-90) ' Ruota tutto di 90 gradi
            .TranslateTransform(-altezza_scontrino, 0) ' Sposta l'area per evitare il taglio

            ' Imposta le dimensioni dei rettangoli e degli spazi
            Dim altezza_ret_titoli As Integer = altezza_scontrino / 6
            Dim larghezza_ret As Integer = larghezza_scontrino - 4

            ' Disegno dei rettangoli (disposti verticalmente)
            '.DrawRectangle(Penna, 2, 2, valore_15, larghezza_ret) ' Primo rettangolo (titoli)
            '.DrawRectangle(Penna, 2, altezza_ret_titoli + 5, altezza_ret_titoli * 2, larghezza_ret) ' Secondo rettangolo
            '.DrawRectangle(Penna, 2, altezza_ret_titoli * 3 + 10, altezza_ret_titoli * 2, larghezza_ret) ' Terzo rettangolo
            '.DrawRectangle(Penna, 2, altezza_ret_titoli * 5 + 15, altezza_ret_titoli, larghezza_ret) ' Quarto rettangolo

            ' Disegno del testo nei rettangoli
            ' Testo nella prima area
            .DrawString(TextBox10.Text, Carattere_numerone, Brushes.Black, 3, 3)

            ' Testo nella seconda area (nome_bp)
            .DrawString(Layout_documenti.trova_dettagli_documento(TextBox10.Text, "ORDR", "RDR1").nome_bp, Carattere_Descrizione, Brushes.Black, 220, 40)

            ' Testo nella seconda area (nome_final_bp)
            .DrawString(Layout_documenti.trova_dettagli_documento(TextBox10.Text, "ORDR", "RDR1").nome_final_bp, Carattere_Descrizione, Brushes.Black, 220, 90)

            ' Testo nella terza area (MATRCDS)
            .DrawString(Layout_documenti.trova_dettagli_documento(TextBox10.Text, "ORDR", "RDR1").MATRCDS, Carattere_Matricola_labelling, Brushes.Black, 500, 3)

            ' Testo nella quarta area (nome_final_bp o nome_bp)
            If Layout_documenti.trova_dettagli_documento(TextBox10.Text, "ORDR", "RDR1").nome_final_bp = "" Then
                .DrawString(Layout_documenti.trova_dettagli_documento(TextBox10.Text, "ORDR", "RDR1").nome_bp.Substring(0, 1), Carattere_numerone_labelling, Brushes.Black, 0, 70)
            Else
                .DrawString(Layout_documenti.trova_dettagli_documento(TextBox10.Text, "ORDR", "RDR1").nome_final_bp.Substring(0, 1), Carattere_numerone_labelling, Brushes.Black, 0, 70)
            End If


            .DrawString(Layout_documenti.trova_dettagli_documento(TextBox10.Text, "ORDR", "RDR1").causcons, Carattere_Matricola_labelling, Brushes.Black, 610, 140)

            .DrawString(Now, Carattere_Descrizione, Brushes.Black, 350, 140)

            ' Ripristina lo stato grafico originale
            .Restore(state2)
        End With






    End Sub

    Private Sub TableLayoutPanel4_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel4.Paint

    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        valore_15 = TextBox15.Text
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        valore_16 = TextBox16.Text
    End Sub
End Class