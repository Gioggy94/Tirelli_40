Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib
Imports Microsoft.Office.Interop.Word
Imports System.CodeDom
Imports MS.Internal.Xaml
Imports Tirelli.Magazzino
Imports Microsoft.Office.Interop.Excel
Imports Tirelli.ODP_Form
Imports Inventor


Public Class Layout_documenti
    Public nome_documento_SAP As String
    Public documento_SAP As String
    Public righe_SAP As String
    Public docnum As String
    Public Lingua As String
    Public Codice_BP As String
    Public Nome_BP As String
    Public Docdate As String
    Public Docduedate As String
    Public Contatto As String
    Public Incoterms As String
    Public Spedizione As String
    Public Vettore As String
    Public Pagamento As String
    Public PIVa As String
    Public IBAN As String
    Public Banca As String
    Public nota_apertura As String
    Public nota_chiusura As String
    Public Compilatore As String
    Public Destinatario_fattura As String
    Public Destinatario_spedizione As String
    Public Totale_DOC As String
    Public paidtodate As String
    Public Sconto_DOC As String
    Public garanzia As String = "N"
    Public alert As String

    Public Valore_sconto As String
    Public IVA As String
    Public Totale_netto_doc As String
    Public spese_di_nolo As String
    Public Valuta As String
    Public Riferimento_cliente As String
    Public peso_lordo As String
    Public peso_netto As String
    Public dimensioni As String
    Public codice_doganale As String
    Public Max_riga As Integer

    Public percorso_documento As String

    Public Totale_netto_pre_iva As String
    Public Presenza_sconto_righe As String
    Public Presenza_produttore As Integer = 0
    Public Presenza_catalogo_fornitore As Integer = 0
    Public Presenza_note As Integer = 0
    Public Riga_query As Integer

    Public oWord As Word.Application
    Public oDoc As Word.Document
    Public oTable As Word.Table

    Public parola_sconto As String
    Public parola_totale_netto As String
    Public osservazioni As String
    Public made_in As Boolean = False
    Public descrizione_temp As String
    Public codice_temp As String
    Public note As String
    Public approvvigionamento_articolo As String
    Public codice_KTF As Boolean = False

    Public percorso_documento_PDF As String
    Public e_mail_contatto As String

    Public prima_data_scadenza As String
    Public seconda_data_scadenza As String
    Public terza_data_scadenza As String
    Public quarta_data_scadenza As String

    Public prima_rata_importo As String
    Public seconda_rata_importo As String
    Public terza_rata_importo As String
    Public quarta_rata_importo As String

    Public swift As String
    Private codice_doganale_riga As Boolean = False
    Public Codice_sap As String
    Public disegno As String
    Public descrizione_sap As String
    Public N_pezzi_NC As String
    Public campo_definizione_NC As String
    Public descrizione_NC As String
    Public osservazioni_nc As String
    Public imputazione_nc As String

    Public richiesto_nc As String
    Public rilevato_nc As String

    Public OA_nc As String
    Public Data_OA_nc As String
    Public percorso_specifico As String = ""
    Public percorso_documento_acquisto_per_qualità As String = ""
    Public percorso_documento_nc_pdf As String = ""
    Private codice_BRB As Boolean = False
    Public P_iva As String
    Public righe_testo As Integer = 0

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = 0 Then
            nome_documento_SAP = "Off"
            documento_SAP = "OQUT"
            righe_SAP = "QUT1"
        ElseIf ComboBox1.SelectedIndex = 1 Then
            nome_documento_SAP = "Order"
            documento_SAP = "ORDR"
            righe_SAP = "RDR1"
        ElseIf ComboBox1.SelectedIndex = 2 Then
            nome_documento_SAP = "Consegna"
            documento_SAP = "ODLN"
            righe_SAP = "DLN1"
        ElseIf ComboBox1.SelectedIndex = 3 Then
            nome_documento_SAP = "Invoice"
            documento_SAP = "OINV"
            righe_SAP = "INV1"

        ElseIf ComboBox1.SelectedIndex = 4 Then
            nome_documento_SAP = "Richiesta_di_offerta"
            documento_SAP = "OPQT"
            righe_SAP = "PQT1"

        ElseIf ComboBox1.SelectedIndex = 5 Then
            nome_documento_SAP = "Ordine_acquisto"
            documento_SAP = "OPOR"
            righe_SAP = "POR1"

        ElseIf ComboBox1.SelectedIndex = 6 Then
            nome_documento_SAP = "Packing_list"
            documento_SAP = "ODLN"
            righe_SAP = "DLN1"

        ElseIf ComboBox1.SelectedIndex = 7 Then
            nome_documento_SAP = "Proforma"
            documento_SAP = "ORDR"
            righe_SAP = "RDR1"
        ElseIf ComboBox1.SelectedIndex = 8 Then

            nome_documento_SAP = "Non_conformità"



        End If
    End Sub

    Sub Informazioni_documento(par_docnum As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.DOCRATE, t0.docnum,t0.cardcode as 'Codice BP', t0.docduedate as 'Docduedate', case when t6.name is null then '' else t6.name end as 'Contatto', t0.cardname as 'Nome BP', T0.[DocDate] as 'Docdate',case when T0.u_prg_azs_incoterms is null then '' else t0.u_prg_azs_incoterms end as 'Incoterms',

case when T0.u_ACURA is null then '' else t0.u_acura end as 'Spedizione',case when t0.u_vettore is null then '' else t0.U_vettore end as 'Vettore' ,
case when T0.u_termini is null then '' else T0.u_termini end  as 'Pagamento', 
case when t0.lictradnum is null then '' else t0.lictradnum end as 'P.IVa', 
case when t3.bankname is null then '' else t3.bankname end as 'Banca', case when t1.hsbnkiban is null then '' else t1.hsbnkiban end as 'Iban',  T4.[lastName]+' ' +T4.[firstName] as 'Compilatore', case when T0.[Address] is null then '' else T0.[Address] end AS 'Destinatario_fattura',case when T0.[Address2] is null then '' else T0.[Address2] end AS 'Destinatario_spedizione', case when t7.name is null then 'Italian' else CASE WHEN t7.name <> 'English' and t7.name <> 'French' and t7.name<>'Italian' THEN 'English' else t7.name end end as 'Lingua',

case when t0.doccur='EUR' THEN (T0.DocTotal -t0.vatsum)/(1-  t0.discprcnt/100)  else (T0.DocTotalFC - t0.vatsum*t0.docrate)/(1-  t0.discprcnt/100) end AS 'Totale',

case when t0.doccur='EUR' THEN (T0.max1099)  else (T0.max1099 *t0.docrate) end AS 'max',

case when t0.doccur='EUR' THEN (T0.paidtodate)  else (T0.paidtodate *t0.docrate) end AS 'paidtodate',



Case when t0.discprcnt is null then 0 else t0.discprcnt end as 'Sconto', case when t0.doccur ='EUR' then t0.discsum else T0.DISCSUMFC end as 'Valore sconto',
case when t0.doccur='EUR' then t0.vatsum else t0.vatsum*t0.docrate end as 'IVA',
case when t0.doccur='EUR' then t0.doctotal  else T0.DocTotalFC end as 'Totale netto',

case when t0.doccur='EUR' then 'EUR' when t0.doccur='USD' then '$' ELSE '' end as 'Valuta', case when t0.numatcard is null then '' else t0.numatcard end as 'Riferimento_cliente', case when t0.u_pesolord is null then '' else t0.u_pesolord end as 'Peso_lordo',case when t0.u_pesonet is null then '' else t0.u_pesonet end as 'Peso_netto',case when t0.u_dimenimb is null then '' else t0.U_dimenimb end as 'Dimensioni', case when t0.U_coddog is null then '' else t0.U_coddog end as 'Codice_doganale', t0.comments
,CAST(t0.header AS VARCHAR(8000)) as 'HEader',CAST(t0.footer AS VARCHAR(8000)) as 'Footer'
,t0.u_01
,t0.u_02
,t0.u_03
,t0.u_04
,t0.u_001
,t0.u_002
,t0.u_003
,t0.u_004
,case when t1.hsbnkswift is null then '' else t1.hsbnkswift end as 'SWIFT'
,coalesce(t1.LicTradNum,'') as 'P_iva'

FROM " & documento_SAP & " T0 
inner join ocrd t1 on t1.cardcode=t0.cardcode
left join OCTG T2 on t2.groupnum=t0.groupnum
left join ODSC T3 ON T3.[BankCode]=t1.housebank 
left join [TIRELLI_40].[dbo].OHEM T4 on t4.empid=t0.ownercode
left join " & righe_SAP & " t5 on t5.docentry=t0.docentry
left join ocpr t6 on t6.cardcode = t0.cardcode AND T0.CNTCTCODE=T6.CNTCTCODE
left join olng t7 on t7.code= t0.langcode
WHERE T0.[DocNum] =" & par_docnum & "
        group by 
 CAST(t0.header AS VARCHAR(8000)),CAST(t0.footer AS VARCHAR(8000)),t0.u_termini,
t0.docnum,t0.cardcode ,T0.DISCSUMFC, t0.doccur, t0.docrate, t0.docduedate, t0.cardname ,T0.DocTotalFC, t6.name, T0.[DocDate],T0.u_prg_azs_incoterms,T0.u_ACURA,t0.u_vettore , T2.PymntGroup , t0.lictradnum , t3.bankname ,t1.hsbnkiban , t0.ownercode , T4.[lastName]+' ' +T4.[firstName] , t7.name, t0.discprcnt, t0.discsum, t0.vatsum, t0.doctotal, T0.[Address], t0.address2, t0.numatcard, t0.u_pesolord, t0.u_pesonet, t0.u_dimenimb, t0.U_coddog,t0.comments,t0.u_01
,t0.u_02
,t0.u_03
,t0.u_04
,t0.u_001
,t0.u_002
,t0.u_003
,t0.u_004,t1.hsbnkswift, t0.max1099, t0.paidtodate,t1.LicTradNum"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            Lingua = cmd_SAP_reader("Lingua")
            Codice_BP = cmd_SAP_reader("Codice BP")
            Nome_BP = cmd_SAP_reader("Nome BP")
            Docdate = cmd_SAP_reader("Docdate")
            Docduedate = cmd_SAP_reader("Docduedate")
            Contatto = cmd_SAP_reader("Contatto")
            Incoterms = cmd_SAP_reader("incoterms")
            Spedizione = cmd_SAP_reader("Spedizione")
            Vettore = cmd_SAP_reader("Vettore")
            Pagamento = cmd_SAP_reader("Pagamento")
            PIVa = cmd_SAP_reader("P.IVa")
            IBAN = cmd_SAP_reader("IBAN")
            Banca = cmd_SAP_reader("Banca")
            swift = cmd_SAP_reader("swift")
            P_iva = cmd_SAP_reader("P_iva")


            If Not cmd_SAP_reader("u_01") Is System.DBNull.Value Then
                prima_data_scadenza = cmd_SAP_reader("u_01")
            Else
                prima_data_scadenza = ""
            End If
            If Not cmd_SAP_reader("u_02") Is System.DBNull.Value Then
                seconda_data_scadenza = cmd_SAP_reader("u_02")
            Else
                seconda_data_scadenza = ""
            End If
            If Not cmd_SAP_reader("u_03") Is System.DBNull.Value Then
                terza_data_scadenza = cmd_SAP_reader("u_03")
            Else
                terza_data_scadenza = ""
            End If
            If Not cmd_SAP_reader("u_04") Is System.DBNull.Value Then
                quarta_data_scadenza = cmd_SAP_reader("u_04")
            Else
                quarta_data_scadenza = ""
            End If
            If Not cmd_SAP_reader("u_001") Is System.DBNull.Value Then
                prima_rata_importo = cmd_SAP_reader("u_001")
            Else
                prima_rata_importo = 0
            End If
            If Not cmd_SAP_reader("u_002") Is System.DBNull.Value Then
                seconda_rata_importo = cmd_SAP_reader("u_002")
            Else
                seconda_rata_importo = 0
            End If
            If Not cmd_SAP_reader("u_003") Is System.DBNull.Value Then
                terza_rata_importo = cmd_SAP_reader("u_003")
            Else
                terza_rata_importo = 0
            End If
            If Not cmd_SAP_reader("u_004") Is System.DBNull.Value Then
                quarta_rata_importo = cmd_SAP_reader("u_004")
            Else
                quarta_rata_importo = 0
            End If



            If Not cmd_SAP_reader("header") Is System.DBNull.Value Then
                nota_apertura = cmd_SAP_reader("header")
            Else
                nota_apertura = ""
            End If
            If Not cmd_SAP_reader("footer") Is System.DBNull.Value Then
                nota_chiusura = cmd_SAP_reader("footer")
            Else
                nota_chiusura = ""
            End If
            Compilatore = cmd_SAP_reader("Compilatore")
            Destinatario_fattura = cmd_SAP_reader("Destinatario_fattura")
            Destinatario_spedizione = cmd_SAP_reader("Destinatario_spedizione")
            paidtodate = cmd_SAP_reader("paidtodate")

            If nome_documento_SAP = "Invoice" Then
                Totale_DOC = cmd_SAP_reader("Max")

            Else
                Totale_DOC = cmd_SAP_reader("Totale")
            End If


            Sconto_DOC = cmd_SAP_reader("Sconto")
            If Not cmd_SAP_reader("Valore sconto") Is System.DBNull.Value Then
                Valore_sconto = cmd_SAP_reader("Valore sconto")
            Else
                Valore_sconto = 0
            End If
            IVA = cmd_SAP_reader("Iva")
            Totale_netto_doc = cmd_SAP_reader("Totale netto")
            Valuta = cmd_SAP_reader("Valuta")
            Riferimento_cliente = cmd_SAP_reader("Riferimento_cliente")
            peso_lordo = cmd_SAP_reader("Peso_lordo")
            peso_netto = cmd_SAP_reader("Peso_netto")
            dimensioni = cmd_SAP_reader("Dimensioni")
            codice_doganale = cmd_SAP_reader("codice_doganale")

            If Valuta = "EUR" Then
                Valuta = "€"
            End If

            If Not cmd_SAP_reader("comments") Is System.DBNull.Value Then
                osservazioni = cmd_SAP_reader("comments")
            Else
                osservazioni = ""
            End If


        Else
            MsgBox("La query Informazione documento offerta non sta funzionando")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Public Function trova_dettagli_documento(par_docnum As String, par_documento_sap As String, par_righe_sap As String)
        Dim dettagli As New Dettaglidocumento()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT t0.docentry, T0.DOCRATE, t0.docnum,t0.cardcode as 'Codice BP'
, t0.docduedate as 'Docduedate', case when t6.name is null then '' else t6.name end as 'Contatto'
, t0.cardname as 'Nome BP', T0.[DocDate] as 'Docdate',case when T0.u_prg_azs_incoterms is null then '' else t0.u_prg_azs_incoterms end as 'Incoterms',

case when T0.u_ACURA is null then '' else t0.u_acura end as 'Spedizione',case when t0.u_vettore is null then '' else t0.U_vettore end as 'Vettore' ,
case when T0.u_termini is null then '' else T0.u_termini end  as 'Pagamento', 
case when t0.lictradnum is null then '' else t0.lictradnum end as 'P.IVa', 
case when t3.bankname is null then '' else t3.bankname end as 'Banca', case when t1.hsbnkiban is null then '' else t1.hsbnkiban end as 'Iban',  T4.[lastName]+' ' +T4.[firstName] as 'Compilatore', case when T0.[Address] is null then '' else T0.[Address] end AS 'Destinatario_fattura',case when T0.[Address2] is null then '' else T0.[Address2] end AS 'Destinatario_spedizione', case when t7.name is null then 'Italian' else CASE WHEN t7.name <> 'English' and t7.name <> 'French' and t7.name<>'Italian' THEN 'English' else t7.name end end as 'Lingua',

case when t0.doccur='EUR' THEN (COALESCE(T0.DocTotal,0) -COALESCE(t0.vatsum,0))/(1-  COALESCE(t0.discprcnt,0)/100)  else (COALESCE(T0.DocTotalFC,0) - COALESCE(t0.vatsum,0)*COALESCE(t0.docrate,0))/(1-  COALESCE(t0.discprcnt,0)/100) end AS 'Totale',

case when t0.doccur='EUR' THEN (T0.max1099)  else (T0.max1099 *t0.docrate) end AS 'max',

case when t0.doccur='EUR' THEN (T0.paidtodate)  else (T0.paidtodate *t0.docrate) end AS 'paidtodate',



Case when t0.discprcnt is null then 0 else t0.discprcnt end as 'Sconto', case when t0.doccur ='EUR' then t0.discsum else T0.DISCSUMFC end as 'Valore sconto',
case when t0.doccur='EUR' then t0.vatsum else t0.vatsum*t0.docrate end as 'IVA',
case when t0.doccur='EUR' then t0.doctotal  else T0.DocTotalFC end as 'Totale netto',

case when t0.doccur='EUR' then 'EUR' when t0.doccur='USD' then '$' ELSE '' end as 'Valuta', case when t0.numatcard is null then '' else t0.numatcard end as 'Riferimento_cliente', case when t0.u_pesolord is null then '' else t0.u_pesolord end as 'Peso_lordo',case when t0.u_pesonet is null then '' else t0.u_pesonet end as 'Peso_netto',case when t0.u_dimenimb is null then '' else t0.U_dimenimb end as 'Dimensioni', case when t0.U_coddog is null then '' else t0.U_coddog end as 'Codice_doganale', t0.comments
,CAST(t0.header AS VARCHAR(8000)) as 'HEader',CAST(t0.footer AS VARCHAR(8000)) as 'Footer'
,t0.u_01
,t0.u_02
,t0.u_03
,t0.u_04
,t0.u_001
,t0.u_002
,t0.u_003
,t0.u_004
,case when t1.hsbnkswift is null then '' else t1.hsbnkswift end as 'SWIFT'
,coalesce(t1.LicTradNum,'') as 'P_iva'
, coalesce(CASE WHEN '" & par_documento_sap & "' ='OQUT' THEN t8.u_nomeinds
when '" & par_documento_sap & "' ='ORDR' THEN t9.u_nomeinds
ELSE '' END,'') AS 'Nome_indirizzo'
,coalesce(t0.shiptocode,'') as 'shiptocode'
,COALESCE(T0.U_MATRCDS,'') AS 'U_MATRCDS'
,COALESCE(T10.CARDNAME,'') AS 'Cliente_finale'
,COALESCE(T0.u_causcons,'') AS 'u_causcons'

FROM " & par_documento_sap & " T0 
inner join ocrd t1 on t1.cardcode=t0.cardcode
left join OCTG T2 on t2.groupnum=t0.groupnum
left join ODSC T3 ON T3.[BankCode]=t1.housebank 
left join [TIRELLI_40].[dbo].OHEM T4 on t4.empid=t0.ownercode
left join " & par_righe_sap & " t5 on t5.docentry=t0.docentry
left join ocpr t6 on t6.cardcode = t0.cardcode AND T0.CNTCTCODE=T6.CNTCTCODE
left join olng t7 on t7.code= t0.langcode
LEFT JOIN QUT12 t8 ON T8.DOCENTRY=T0.DOCENTRY and '" & par_documento_sap & "' ='OQUT'
LEFT JOIN RDR12 t9 ON T9.DOCENTRY=T0.DOCENTRY and '" & par_documento_sap & "' ='ORDR'
LEFT JOIN OCRD T10 ON T10.CARDCODE=T0.U_CODICEBP

WHERE T0.[DocNum] =" & par_docnum & "
        group by 
 CAST(t0.header AS VARCHAR(8000)),CAST(t0.footer AS VARCHAR(8000)),t0.u_termini,t0.shiptocode,
t0.docnum,t0.cardcode ,T0.DISCSUMFC, t0.doccur, t0.docrate, t0.docduedate, t0.cardname ,T0.DocTotalFC, t6.name, T0.[DocDate],T0.u_prg_azs_incoterms,T0.u_ACURA,t0.u_vettore , T2.PymntGroup , t0.lictradnum , t3.bankname ,t1.hsbnkiban , t0.ownercode , T4.[lastName]+' ' +T4.[firstName] , t7.name, t0.discprcnt, t0.discsum, t0.vatsum, t0.doctotal, T0.[Address], t0.address2, t0.numatcard, t0.u_pesolord, t0.u_pesonet, t0.u_dimenimb, t0.U_coddog,t0.comments,t0.u_01
,t0.u_02
,t0.u_03
,t0.u_04
,t0.u_001
,t0.u_002
,t0.u_003
,t0.u_004,t1.hsbnkswift, t0.max1099, t0.paidtodate,t1.LicTradNum, t8.u_nomeinds, t9.u_nomeinds,t0.docentry,T0.U_MATRCDS,t10.cardname, t0.u_causcons "

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then
            dettagli.docentry = cmd_SAP_reader("docentry")
            dettagli.Lingua = cmd_SAP_reader("Lingua")
            dettagli.Codice_BP = cmd_SAP_reader("Codice BP")
            dettagli.Nome_BP = cmd_SAP_reader("Nome BP")
            dettagli.Nome_final_BP = cmd_SAP_reader("Cliente_finale")
            dettagli.Docdate = cmd_SAP_reader("Docdate")
            dettagli.Docduedate = cmd_SAP_reader("Docduedate")
            dettagli.Contatto = cmd_SAP_reader("Contatto")
            dettagli.Incoterms = cmd_SAP_reader("incoterms")
            dettagli.Spedizione = cmd_SAP_reader("Spedizione")
            dettagli.Vettore = cmd_SAP_reader("Vettore")
            dettagli.Pagamento = cmd_SAP_reader("Pagamento")
            dettagli.PIVa = cmd_SAP_reader("P.IVa")
            dettagli.IBAN = cmd_SAP_reader("IBAN")
            dettagli.Banca = cmd_SAP_reader("Banca")
            dettagli.swift = cmd_SAP_reader("swift")
            dettagli.NOME_INDIRIZZO = cmd_SAP_reader("Nome_indirizzo")
            dettagli.P_iva = cmd_SAP_reader("P_iva")
            dettagli.shiptocode = cmd_SAP_reader("shiptocode")
            dettagli.causcons = cmd_SAP_reader("u_causcons")


            If Not cmd_SAP_reader("u_01") Is System.DBNull.Value Then
                dettagli.prima_data_scadenza = cmd_SAP_reader("u_01")
            Else
                dettagli.prima_data_scadenza = ""
            End If
            If Not cmd_SAP_reader("u_02") Is System.DBNull.Value Then
                dettagli.seconda_data_scadenza = cmd_SAP_reader("u_02")
            Else
                dettagli.seconda_data_scadenza = ""
            End If
            If Not cmd_SAP_reader("u_03") Is System.DBNull.Value Then
                dettagli.terza_data_scadenza = cmd_SAP_reader("u_03")
            Else
                dettagli.terza_data_scadenza = ""
            End If
            If Not cmd_SAP_reader("u_04") Is System.DBNull.Value Then
                dettagli.quarta_data_scadenza = cmd_SAP_reader("u_04")
            Else
                dettagli.quarta_data_scadenza = ""
            End If
            If Not cmd_SAP_reader("u_001") Is System.DBNull.Value Then
                dettagli.prima_rata_importo = cmd_SAP_reader("u_001")
            Else
                dettagli.prima_rata_importo = 0
            End If
            If Not cmd_SAP_reader("u_002") Is System.DBNull.Value Then
                dettagli.seconda_rata_importo = cmd_SAP_reader("u_002")
            Else
                dettagli.seconda_rata_importo = 0
            End If
            If Not cmd_SAP_reader("u_003") Is System.DBNull.Value Then
                dettagli.terza_rata_importo = cmd_SAP_reader("u_003")
            Else
                dettagli.terza_rata_importo = 0
            End If
            If Not cmd_SAP_reader("u_004") Is System.DBNull.Value Then
                dettagli.quarta_rata_importo = cmd_SAP_reader("u_004")
            Else
                dettagli.quarta_rata_importo = 0
            End If



            If Not cmd_SAP_reader("header") Is System.DBNull.Value Then
                dettagli.nota_apertura = cmd_SAP_reader("header")
            Else
                dettagli.nota_apertura = ""
            End If
            If Not cmd_SAP_reader("footer") Is System.DBNull.Value Then
                dettagli.nota_chiusura = cmd_SAP_reader("footer")
            Else
                dettagli.nota_chiusura = ""
            End If
            dettagli.Compilatore = cmd_SAP_reader("Compilatore")
            dettagli.Destinatario_fattura = cmd_SAP_reader("Destinatario_fattura").Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ")
            dettagli.Destinatario_spedizione = cmd_SAP_reader("Destinatario_spedizione").Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ")
            dettagli.paidtodate = cmd_SAP_reader("paidtodate")

            If nome_documento_SAP = "Invoice" Then
                dettagli.Totale_DOC = cmd_SAP_reader("Max")

            Else
                dettagli.Totale_DOC = cmd_SAP_reader("Totale")
            End If


            dettagli.Sconto_DOC = cmd_SAP_reader("Sconto")
            If Not cmd_SAP_reader("Valore sconto") Is System.DBNull.Value Then
                dettagli.Valore_sconto = cmd_SAP_reader("Valore sconto")
            Else
                dettagli.Valore_sconto = 0
            End If
            dettagli.IVA = cmd_SAP_reader("Iva")
            dettagli.Totale_netto_doc = cmd_SAP_reader("Totale netto")
            dettagli.Valuta = cmd_SAP_reader("Valuta")
            dettagli.Riferimento_cliente = cmd_SAP_reader("Riferimento_cliente")
            dettagli.peso_lordo = cmd_SAP_reader("Peso_lordo")
            dettagli.peso_netto = cmd_SAP_reader("Peso_netto")
            dettagli.dimensioni = cmd_SAP_reader("Dimensioni")
            dettagli.codice_doganale = cmd_SAP_reader("codice_doganale")
            dettagli.MATRCDS = cmd_SAP_reader("U_MATRCDS")

            If dettagli.Valuta = "EUR" Then
                dettagli.Valuta = "€"
            End If

            If Not cmd_SAP_reader("comments") Is System.DBNull.Value Then
                dettagli.osservazioni = cmd_SAP_reader("comments")
            Else
                dettagli.osservazioni = ""
            End If


        Else
            MsgBox("La query Informazione documento offerta non sta funzionando")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
        Return dettagli
    End Function

    Public Class Dettaglidocumento
        Public docentry As Integer
        Public docnum As String
        Public Lingua As String
        Public Codice_BP As String
        Public Nome_BP As String
        Public Nome_final_BP As String
        Public Docdate As String
        Public Docduedate As String
        Public Contatto As String
        Public Incoterms As String
        Public Spedizione As String
        Public Vettore As String
        Public Pagamento As String
        Public PIVa As String
        Public IBAN As String
        Public shiptocode As String
        Public Banca As String
        Public nota_apertura As String
        Public nota_chiusura As String
        Public Compilatore As String
        Public Destinatario_fattura As String
        Public Destinatario_spedizione As String
        Public Totale_DOC As String
        Public paidtodate As String
        Public Sconto_DOC As String
        Public garanzia As String = "N"
        Public alert As String
        Public Nome_indirizzo As String
        Public MATRCDS As String
        Public Causcons As String


        Public Valore_sconto As String
        Public IVA As String
        Public Totale_netto_doc As String
        Public spese_di_nolo As String
        Public Valuta As String
        Public Riferimento_cliente As String
        Public peso_lordo As String
        Public peso_netto As String
        Public dimensioni As String
        Public codice_doganale As String
        Public Max_riga As Integer
        Public swift As String
        Public P_iva As String
        Public prima_data_scadenza As String
        Public seconda_data_scadenza As String
        Public terza_data_scadenza As String
        Public quarta_data_scadenza As String
        Public prima_rata_importo As String
        Public seconda_rata_importo As String
        Public terza_rata_importo As String
        Public quarta_rata_importo As String
        Public osservazioni As String


    End Class

    Sub Informazioni_documento_acquisto(par_documento As Integer, par_documento_sap As String, par_righe_sap As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.DOCRATE, t0.docnum,t0.cardcode as 'Codice BP', t0.docduedate as 'Docduedate', case when t6.name is null then '' else t6.name end as 'Contatto', t0.cardname as 'Nome BP', T0.[DocDate] as 'Docdate',case when T0.u_prg_azs_incoterms is null then '' else t0.u_prg_azs_incoterms end as 'Incoterms',case when T0.u_ACURA is null then '' else t0.u_acura end as 'Spedizione',case when t0.u_vettore is null then '' else t0.U_vettore end as 'Vettore' , T2.PymntGroup as 'Pagamento', case when t0.lictradnum is null then '' else t0.lictradnum end as 'P.IVa',
case when t3.bankname is null then '' else t3.bankname end as 'Banca', case when t1.hsbnkiban is null then '' else t1.hsbnkiban end as 'Iban'
,  coalesce(T4.[lastName]+' ' +T4.[firstName],'') as 'Compilatore', case when T0.[Address] is null then '' else T0.[Address] end AS 'Destinatario_fattura',case when T0.[Address2] is null then '' else T0.[Address2] end AS 'Destinatario_spedizione', case when t7.name is null then 'Italian' else CASE WHEN t7.name <> 'English' and t7.name <> 'French' and t7.name<>'Italian' THEN 'English' else t7.name end end as 'Lingua',



case when t0.doccur='EUR' THEN (T0.DocTotal -t0.vatsum)/(1-  coalesce(t0.discprcnt,0)/100)  else (coalesce(T0.DocTotalFC,0) - coalesce(t0.vatsum,0)*coalesce(t0.docrate,0))/(1-  coalesce(t0.discprcnt,0)/100) end - coalesce(t0.totalexpns,0) AS 'Totale',



Case when t0.discprcnt is null then 0 else t0.discprcnt end as 'Sconto', case when t0.doccur ='EUR' then t0.discsum else T0.DISCSUMFC end as 'Valore sconto',
 t0.totalexpns,
case when t0.doccur='EUR' then t0.vatsum else t0.vatsum*t0.docrate end as 'IVA',
case when t0.doccur='EUR' then t0.doctotal  else T0.DocTotalFC end as 'Totale netto',

case when t0.doccur='EUR' then 'EUR' when t0.doccur='USD' then '$' ELSE '' end as 'Valuta', case when t0.numatcard is null then '' else t0.numatcard end as 'Riferimento_cliente', case when t0.u_pesolord is null then '' else t0.u_pesolord end as 'Peso_lordo',case when t0.u_pesonet is null then '' else t0.u_pesonet end as 'Peso_netto',case when t0.u_dimenimb is null then '' else t0.U_dimenimb end as 'Dimensioni', case when t0.U_coddog is null then '' else t0.U_coddog end as 'Codice_doganale', t0.comments, case when t6.e_maill is null then '' else t6.e_maill end as 'e_maill'
,case when t0.u_alert is null then '' else t0.u_alert end as 'Alert'  
,
CAST(t0.header AS VARCHAR(8000)) as 'HEader',CAST(t0.footer AS VARCHAR(8000)) as 'Footer'


FROM " & par_documento_sap & " T0 
inner join ocrd t1 on t1.cardcode=t0.cardcode
left join OCTG T2 on t2.groupnum=t0.groupnum
left join ODSC T3 ON T3.[BankCode]=t1.housebank 
left join [TIRELLI_40].[dbo].OHEM T4 on t4.empid=t0.ownercode
left join " & par_righe_sap & " t5 on t5.docentry=t0.docentry
left join ocpr t6 on t6.cardcode = t0.cardcode AND T0.CNTCTCODE=T6.CNTCTCODE
left join olng t7 on t7.code= t0.langcode
WHERE T0.[DocNum] =" & par_documento & "
         group by   
CAST(t0.header AS VARCHAR(8000)),CAST(t0.footer AS VARCHAR(8000)),
		t0.docnum,t0.cardcode ,T0.DISCSUMFC, t0.doccur, t0.docrate, t0.docduedate, t0.cardname ,T0.DocTotalFC, t6.name, T0.[DocDate],T0.u_prg_azs_incoterms,T0.u_ACURA,t0.u_vettore , T2.PymntGroup , t0.lictradnum , t3.bankname ,t1.hsbnkiban , t0.ownercode , T4.[lastName]+' ' +T4.[firstName] , t7.name, t0.discprcnt, t0.discsum, t0.vatsum, t0.doctotal, T0.[Address], t0.address2, t0.numatcard, t0.u_pesolord, t0.u_pesonet, t0.u_dimenimb, t0.U_coddog,t0.comments, t6.e_maill,t0.u_alert,  t0.totalexpns"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            Lingua = cmd_SAP_reader("Lingua")
            Codice_BP = cmd_SAP_reader("Codice BP")
            Nome_BP = cmd_SAP_reader("Nome BP")
            Docdate = cmd_SAP_reader("Docdate")
            Docduedate = cmd_SAP_reader("Docduedate")
            Contatto = cmd_SAP_reader("Contatto")
            Incoterms = cmd_SAP_reader("incoterms")
            Spedizione = cmd_SAP_reader("Spedizione")
            Vettore = cmd_SAP_reader("Vettore")
            Pagamento = cmd_SAP_reader("Pagamento")
            PIVa = cmd_SAP_reader("P.IVa")
            IBAN = cmd_SAP_reader("IBAN")
            Banca = cmd_SAP_reader("Banca")
            If Not cmd_SAP_reader("header") Is System.DBNull.Value Then
                nota_apertura = cmd_SAP_reader("header")
            Else
                nota_apertura = ""
            End If
            If Not cmd_SAP_reader("footer") Is System.DBNull.Value Then
                nota_chiusura = cmd_SAP_reader("footer")
            Else
                nota_chiusura = ""
            End If
            Compilatore = cmd_SAP_reader("Compilatore")
            Destinatario_fattura = cmd_SAP_reader("Destinatario_fattura")
            Destinatario_spedizione = cmd_SAP_reader("Destinatario_spedizione")
            Totale_DOC = cmd_SAP_reader("Totale")
            Sconto_DOC = cmd_SAP_reader("Sconto")
            e_mail_contatto = cmd_SAP_reader("e_maill")
            alert = cmd_SAP_reader("alert")
            spese_di_nolo = cmd_SAP_reader("totalexpns")

            If Not cmd_SAP_reader("Valore sconto") Is System.DBNull.Value Then
                Valore_sconto = cmd_SAP_reader("Valore sconto")
            Else
                Valore_sconto = 0
            End If
            IVA = cmd_SAP_reader("Iva")
            Totale_netto_doc = cmd_SAP_reader("Totale netto")
            Valuta = cmd_SAP_reader("Valuta")
            Riferimento_cliente = cmd_SAP_reader("Riferimento_cliente")
            peso_lordo = cmd_SAP_reader("Peso_lordo")
            peso_netto = cmd_SAP_reader("Peso_netto")
            dimensioni = cmd_SAP_reader("Dimensioni")
            codice_doganale = cmd_SAP_reader("codice_doganale")

            If Valuta = "EUR" Then
                Valuta = "€"
            End If

            If Not cmd_SAP_reader("comments") Is System.DBNull.Value Then
                osservazioni = cmd_SAP_reader("comments")
            Else
                osservazioni = ""
            End If


        Else
            MsgBox("La query Informazione documento acquisto non sta funzionando")
            Return
        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub Informazioni_NC(par_docnum As Integer)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = " 
select t10.id,t10.data,t10.cod,t10.itemname, t10.disegno, case when t10.odp is null then T15.DOCNUM else t10.odp end as 'ODP'
,coalesce(case when t10.oa is null then t14.docnum else t10.oa end,'') as 'OA'
,case when t14.docdate is null then '' else t14.DOCDATE end as 'dATA_OA'
,t10.emp,t10.emf,t10.ordine,t10.cod_fornitore,
coalesce(t10.Fornitore,'') as 'Fornitore',t10.imputazione,t10.zona_controllo,coalesce(t10.Dipendente,'') as 'Dipendente',t10.[Q.tà Contr.],t10.[Q.tà NC],t10.[Q.tà OK],t10.attività,t10.esito,t10.campo,t10.Descrizione_NC,t10.osservazioni,t10.peso,t10.stato,t10.richiesto,t10.rilevato,t10.concedente
, t10.resname, t10.[Codice operatore], t10.[operatore MU],t10.[Data lavorazione],t10.autocontrollo

from
(
 SELECT t0.ID, T0.DATA, T0.CODICE AS 'Cod',T1.ITEMNAME , T1.U_DISEGNO as 'Disegno', T0.ODP, T0.OA,
 case when t6.docnum is null then t0.emp else t6.docnum end as 'EMP',
 case when t5.docnum is null then t0.emf else t5.docnum end as 'EMF', 
 CASE WHEN ( T0.OA IS NULL OR T0.OA=0) THEN T0.ODP ELSE T0.OA END AS 'Ordine',
 CASE WHEN (T0.attività = 0 or t0.attività is null) THEN T3.CARDcode else t5.cardcode END AS 'Cod_Fornitore',
 CASE WHEN (T0.attività = 0 or t0.attività is null) THEN T3.CARDNAME else t5.cardname END AS 'Fornitore',
 
 T0.IMPUTAZIONE, T0.ZONA_CONTROLLO,  T2.[U_NAME] as 'Dipendente', T0.PZ_CONTR AS 'Q.tà Contr.', T0.PZ_NC AS 'Q.tà NC', T0.PZ_OK AS 'Q.tà OK', T0.Attività, T0.ESITO_AUTOCONTROLLO AS 'Esito', t0.campo_definizione_NC AS 'Campo', T0.DESCRIZIONE_NC as 'Descrizione_NC', T0.OSSERVAZIONI_NC as 'Osservazioni', T0.PESO_NC AS 'PESO', T0.STATO, t0.richiesto,t0.rilevato, t0.concedente, t0.autocontrollo
  ,t7.itemcode, t8.resname, t7.dipendente as 'Codice operatore', concat(t9.firstname,' ',t9.lastname) as 'operatore MU',t7.data as 'Data lavorazione'
FROM [TIRELLI_40].[DBO].cq_nuovo_controllo T0
INNER JOIN OITM T1 ON T1.ITEMCODE = T0.CODICE
LEFT JOIN OUSR T2 ON T0.OPERATORE=T2.USERID 
LEFT JOIN OPdn T3 ON T3.DOCNUM=T0.emf
left join oclg t4 on t4.ClgCode=t0.attività 
left join opdn t5 on t5.DocNum=t4.docnum and t4.doctype=20
left join oign t6 on t6.docnum=t4.docnum and t4.doctype=59
 left join [TIRELLI_40].[DBO].autocontrollo t7 on t7.id=t0.autocontrollo
left join orsc t8 on t8.visrescode=t7.itemcode
  left join [TIRELLI_40].[dbo].ohem t9 on t9.empid=t7.dipendente

where T0.id='" & par_docnum & "'
)
as t10
left join opdn t11 on t11.docnum=t10.emf
left join pdn1 t12 on t12.docentry=t11.docentry AND T10.COD=T12.ITEMCODE
LEFT JOIN POR1 T13 ON T13.DOCENTRY=T12.BASEENTRY AND T13.LINENUM=T12.BASELINE AND T13.ITEMCODE=T12.ITEMCODE
left join opor t14 on t14.docentry=t13.docentry
left join oign t15 on t15.docnum=t10.EMP
left join IGN1 t16 on t16.docentry=t15.docentry AND T10.COD=T16.ITEMCODE
LEFT JOIN OWOR T17 ON T17.DOCNUM=T16.BASEREF

order by t10.id desc"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then


            Compilatore = cmd_SAP_reader("dipendente")
            Nome_BP = cmd_SAP_reader("fornitore")
            Codice_sap = cmd_SAP_reader("Cod")
            disegno = cmd_SAP_reader("disegno")
            descrizione_sap = cmd_SAP_reader("itemname")
            N_pezzi_NC = cmd_SAP_reader("Q.tà NC")
            campo_definizione_NC = cmd_SAP_reader("Campo")
            descrizione_NC = cmd_SAP_reader("descrizione_NC")
            osservazioni_nc = cmd_SAP_reader("osservazioni")
            imputazione_nc = cmd_SAP_reader("IMPUTAZIONE")
            OA_nc = cmd_SAP_reader("OA")
            Data_OA_nc = cmd_SAP_reader("data_oa")
            'richiesto_nc = cmd_SAP_reader("richiesto")
            rilevato_nc = cmd_SAP_reader("rilevato")
            Lingua = "Italian"



        Else
            MsgBox("La query Informazione documento acquisto")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub Genera_documento(par_codice_brb As String)

        max_righe_documento()


        oWord = CreateObject("Word.Application")
        oDoc = oWord.Documents.Add("" & percorso_documento & "")

        oDoc.Bookmarks.Item("NomeBP").Range.Text = Nome_BP & vbCrLf & P_iva

        oDoc.Bookmarks.Item("Compilatore").Range.Text = Compilatore
        If nome_documento_SAP <> "Proforma" And nome_documento_SAP <> "Invoice" Then
            oDoc.Bookmarks.Item("Contatto").Range.Text = Contatto
        End If

        If nome_documento_SAP = "Proforma" Or nome_documento_SAP = "Invoice" Then
            oDoc.Bookmarks.Item("swift").Range.Text = swift
        End If

        oDoc.Bookmarks.Item("Docdate").Range.Text = Docdate
        If nome_documento_SAP <> "Packing_list" And nome_documento_SAP <> "Invoice" And nome_documento_SAP <> "Proforma" Then
            oDoc.Bookmarks.Item("Docduedate").Range.Text = Docduedate
        End If

        If nome_documento_SAP <> "Packing_list" Then

            oDoc.Bookmarks.Item("Totale_DOC").Range.Text = Valuta & " " & FormatNumber(Totale_DOC, 2, , , TriState.True)


            oDoc.Bookmarks.Item("Totale_Netto_doc").Range.Text = Valuta & " " & FormatNumber(Totale_netto_doc, 2, , , TriState.True)
            oDoc.Bookmarks.Item("Osservazioni").Range.Text = osservazioni
            oDoc.Bookmarks.Item("Iva").Range.Text = Valuta & " " & FormatNumber(IVA, 2, , , TriState.True)
            oDoc.Bookmarks.Item("Iban").Range.Text = IBAN
            oDoc.Bookmarks.Item("Pagamento").Range.Text = Pagamento
            oDoc.Bookmarks.Item("Banca").Range.Text = Banca
        End If
        Try
            oDoc.Bookmarks.Item("P_iva").Range.Text = P_iva
        Catch ex As Exception

        End Try

        If nome_documento_SAP = "Invoice" Then

            oDoc.Bookmarks.Item("paidtodate").Range.Text = Valuta & " " & FormatNumber(-Totale_DOC + Totale_netto_doc, 2, , , TriState.True)

        End If
        oDoc.Bookmarks.Item("Docnum").Range.Text = docnum


        oDoc.Bookmarks.Item("Incoterms").Range.Text = Incoterms
        oDoc.Bookmarks.Item("Destinatario_fattura").Range.Text = Destinatario_fattura
        oDoc.Bookmarks.Item("Destinatario_spedizione").Range.Text = Destinatario_spedizione



        If nome_documento_SAP = "Proforma" Or nome_documento_SAP = "Invoice" Then

            If prima_rata_importo <> "" And prima_rata_importo <> 0 Then

                oDoc.Bookmarks.Item("importo_prima_rata").Range.Text = Valuta & " " & FormatNumber(prima_rata_importo, 2, , , TriState.True)
            End If
            If seconda_rata_importo <> "" And seconda_rata_importo <> 0 Then
                oDoc.Bookmarks.Item("importo_seconda_rata").Range.Text = Valuta & " " & FormatNumber(seconda_rata_importo, 2, , , TriState.True)
            End If
            If terza_rata_importo <> "" And terza_rata_importo <> 0 Then
                oDoc.Bookmarks.Item("importo_terza_rata").Range.Text = Valuta & " " & FormatNumber(terza_rata_importo, 2, , , TriState.True)
            End If
            If quarta_rata_importo <> "" And quarta_rata_importo <> 0 Then
                oDoc.Bookmarks.Item("importo_quarta_rata").Range.Text = Valuta & " " & FormatNumber(quarta_rata_importo, 2, , , TriState.True)
            End If


            oDoc.Bookmarks.Item("Data_prima_rata").Range.Text = prima_data_scadenza
            oDoc.Bookmarks.Item("Data_seconda_rata").Range.Text = seconda_data_scadenza
            oDoc.Bookmarks.Item("Data_terza_rata").Range.Text = terza_data_scadenza

            oDoc.Bookmarks.Item("Data_quarta_rata").Range.Text = quarta_data_scadenza

        End If


        If documento_SAP = "ORDR" Or nome_documento_SAP = "proforma" Or nome_documento_SAP = "Invoice" Or nome_documento_SAP = "Off" Then
            oDoc.Bookmarks.Item("riferimento").Range.Text = Riferimento_cliente
        End If
        If Sconto_DOC > 0 Then

            If Lingua = "Italian" Then
                parola_sconto = "Sconto " & FormatNumber(Sconto_DOC, 2, , , TriState.True) & " %"
                parola_totale_netto = ""
            ElseIf Lingua = "French" Then
                parola_sconto = "Réduction " & FormatNumber(Sconto_DOC, 2, , , TriState.True) & " %"
                parola_totale_netto = ""
            Else
                parola_sconto = "Discount " & FormatNumber(Sconto_DOC, 2, , , TriState.True) & " %"
                parola_totale_netto = ""
            End If



            ' oDoc.Bookmarks.Item("Parola_Totale_netto").Range.Text = parola_totale_netto
            oDoc.Bookmarks.Item("Valore_Sconto").Range.Text = Valuta & " " & FormatNumber(Valore_sconto, 2, , , TriState.True)

            oDoc.Bookmarks.Item("parola_sconto").Range.Text = parola_sconto
        End If

        oDoc.Bookmarks.Item("Spedizione").Range.Text = Spedizione



        Totale_netto_pre_iva = Totale_DOC - Valore_sconto



        Verifica_Presenza_sconto_righe()
        Dim aggiunta_colonne As Integer = 0

        If Presenza_sconto_righe > 0 And nome_documento_SAP <> "Proforma" And nome_documento_SAP <> "Invoice" And nome_documento_SAP <> "Consegna" Then
            aggiunta_colonne = aggiunta_colonne + 1

        End If

        If documento_SAP = "OQUT" Then
            aggiunta_colonne = aggiunta_colonne + 1

        End If

        If made_in = True Then
            aggiunta_colonne = aggiunta_colonne + 1

        End If

        If codice_doganale_riga = True Then
            aggiunta_colonne = aggiunta_colonne + 1

        End If

        If note = "Y" Then
            aggiunta_colonne = aggiunta_colonne + 1

        End If

        If approvvigionamento_articolo = True Then
            aggiunta_colonne = aggiunta_colonne + 1

        End If
        If codice_KTF = True Then
            aggiunta_colonne = aggiunta_colonne + 1

        End If
        If par_codice_brb = True Then
            aggiunta_colonne = aggiunta_colonne + 1

        End If
        If nome_documento_SAP = "Packing_list" Then
            aggiunta_colonne = aggiunta_colonne + 2
        End If
        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("Tabella").Range, Max_riga + 2, 7 + aggiunta_colonne)
        Dim c As Integer
        c = 1
        If Lingua = "Italian" Then
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            oTable.Cell(1, c).Range.Text = "Riga"
            c = c + 1
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter

            oTable.Cell(1, c).Range.Text = "Codice"
            c = c + 1

            If par_codice_brb = True Then


                oTable.Cell(1, c).Range.Text = "BRB Code"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1


            End If
            oTable.Cell(1, c).Range.Text = "Descrizione"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            If nome_documento_SAP <> "Packing_list" Then
                oTable.Cell(1, c).Range.Text = "Commessa"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If

            oTable.Cell(1, c).Range.Text = "Quantità"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            If nome_documento_SAP <> "Packing_list" Then
                oTable.Cell(1, c).Range.Text = "Prezzo"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If
            If Presenza_sconto_righe > 0 And nome_documento_SAP <> "Proforma" And nome_documento_SAP <> "Invoice" And nome_documento_SAP <> "Packing_list" And nome_documento_SAP <> "Consegna" Then
                oTable.Cell(1, c).Range.Text = "Sconto %"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If
            If nome_documento_SAP <> "Packing_list" Then
                oTable.Cell(1, c).Range.Text = "Totale"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If

            If made_in = True Then
                oTable.Cell(1, c).Range.Text = "Made in"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If

            If codice_doganale_riga = True Then
                oTable.Cell(1, c).Range.Text = "Codice doganale"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If

            If note = "Y" Then
                oTable.Cell(1, c).Range.Text = "Note"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If


            If approvvigionamento_articolo = True Then



                If documento_SAP = "OQUT" Or documento_SAP = "ORDR" Then
                    oTable.Cell(1, c).Range.Text = "Lead Time"
                    oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    c = c + 1

                End If
            End If

            If codice_KTF = True Then


                oTable.Cell(1, c).Range.Text = "KTF Code"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1


            End If



        ElseIf Lingua = "French" Then
            oTable.Cell(1, c).Range.Text = "Rangée"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            oTable.Cell(1, c).Range.Text = "Code"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            If codice_BRB = True Then


                oTable.Cell(1, c).Range.Text = "BRB Code"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1


            End If
            oTable.Cell(1, c).Range.Text = "Description"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            If nome_documento_SAP <> "Packing_list" Then
                oTable.Cell(1, c).Range.Text = "Machine"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If
            oTable.Cell(1, c).Range.Text = "Quantité"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            If nome_documento_SAP <> "Packing_list" Then
                oTable.Cell(1, c).Range.Text = "Prix"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If
            If Presenza_sconto_righe > 0 And nome_documento_SAP <> "Proforma" And nome_documento_SAP <> "Invoice" And nome_documento_SAP <> "Packing_list" And nome_documento_SAP <> "Consegna" Then
                oTable.Cell(1, c).Range.Text = "Réduction %"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If
            If nome_documento_SAP <> "Packing_list" Then
                oTable.Cell(1, c).Range.Text = "Total"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If
            If made_in = True Then
                oTable.Cell(1, c).Range.Text = "Made in"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If

            If codice_doganale_riga = True Then
                oTable.Cell(1, c).Range.Text = "Custom code"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If

            If note = "Y" Then
                oTable.Cell(1, c).Range.Text = "Notes"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If

            If approvvigionamento_articolo = True Then


                If documento_SAP = "OQUT" Or documento_SAP = "ORDR" Then
                    oTable.Cell(1, c).Range.Text = "Lead Time"
                    oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    c = c + 1

                End If
            End If

            If codice_KTF = True Then


                oTable.Cell(1, c).Range.Text = "KTF Code"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1


            End If

        Else

            oTable.Cell(1, c).Range.Text = "Row"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            oTable.Cell(1, c).Range.Text = "Code"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            If par_codice_brb = True Then


                oTable.Cell(1, c).Range.Text = "BRB Code"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1


            End If
            oTable.Cell(1, c).Range.Text = "Description"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            If nome_documento_SAP <> "Packing_list" Then
                oTable.Cell(1, c).Range.Text = "Machine"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If
            oTable.Cell(1, c).Range.Text = "Quantity"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            If nome_documento_SAP <> "Packing_list" Then
                oTable.Cell(1, c).Range.Text = "Price"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If
            If Presenza_sconto_righe > 0 And nome_documento_SAP <> "Proforma" And nome_documento_SAP <> "Invoice" And nome_documento_SAP <> "Packing_list" And nome_documento_SAP <> "Consegna" Then
                oTable.Cell(1, c).Range.Text = "Discount %"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If
            If nome_documento_SAP <> "Packing_list" Then
                oTable.Cell(1, c).Range.Text = "Total"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If

            If made_in = True Then
                oTable.Cell(1, c).Range.Text = "Made in"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If

            If codice_doganale_riga = True Then
                oTable.Cell(1, c).Range.Text = "Custom code"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If

            If note = "Y" Then
                oTable.Cell(1, c).Range.Text = "Notes"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If

            If approvvigionamento_articolo = True Then


                If documento_SAP = "OQUT" Or documento_SAP = "ORDR" Then
                    oTable.Cell(1, c).Range.Text = "Lead Time"
                    Try
                        oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    Catch ex As Exception

                    End Try
                    c = c + 1


                End If
            End If

            If codice_KTF = True Then


                oTable.Cell(1, c).Range.Text = "KTF Code"
                oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1


            End If



        End If

        If nome_documento_SAP = "Packing_list" Then
            oTable.Cell(1, c).Range.Text = "Packaging type"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            oTable.Cell(1, c).Range.Text = "Dimension L (mms.)"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            oTable.Cell(1, c).Range.Text = "Dimension W (mms.)"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            oTable.Cell(1, c).Range.Text = "Dimension H (mms.)"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            oTable.Cell(1, c).Range.Text = "Gross / Net Weight per package (Kgs.)"
            oTable.Cell(1, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

        End If

        Informazioni_righe_offerta(par_codice_brb)



        oTable.AutoFormat(ApplyColor:=False, ApplyBorders:=True)

        oTable.Rows.Item(1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter


        oTable.Rows.Item(1).Range.Font.Bold = True
        'oTable.Columns.Item(1).Width = oWord.InchesToPoints(1)   'Change width of columns 1 & 2
        oTable.Rows(1).Cells.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        oTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle

        oTable.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)

        Dim variabile As String = Homepage.percorso_offerte_vendita & nome_documento_SAP & "_" & docnum & ".doc"
        oTable.Rows.Height = 30


        If percorso_specifico = "" Then
            oDoc.SaveAs(Homepage.percorso_offerte_vendita & nome_documento_SAP & "_" & docnum & ".doc")
        Else
            oDoc.SaveAs(percorso_specifico)
        End If

        oWord.Documents.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
        If oWord.Documents.Count = 0 Then
            oWord.Application.Quit()
        End If



        If percorso_specifico = "" Then
            ConvertWordToPDF(Homepage.percorso_offerte_vendita & nome_documento_SAP & "_" & docnum & ".doc")
        Else
            ConvertWordToPDF(percorso_specifico)
        End If
    End Sub

    Sub Genera_documento_ACQUISTO()

        max_righe_documento()


        oWord = CreateObject("Word.Application")
        oDoc = oWord.Documents.Add("" & percorso_documento & "")

        oDoc.Bookmarks.Item("NomeBP").Range.Text = Nome_BP
        oDoc.Bookmarks.Item("Banca").Range.Text = Banca
        oDoc.Bookmarks.Item("Compilatore").Range.Text = Compilatore
        oDoc.Bookmarks.Item("Contatto").Range.Text = Contatto
        oDoc.Bookmarks.Item("Docdate").Range.Text = Docdate
        oDoc.Bookmarks.Item("Docduedate").Range.Text = Docduedate
        oDoc.Bookmarks.Item("Docnum").Range.Text = docnum
        oDoc.Bookmarks.Item("Iban").Range.Text = IBAN
        oDoc.Bookmarks.Item("Pagamento").Range.Text = Pagamento
        oDoc.Bookmarks.Item("Incoterms").Range.Text = Incoterms
        oDoc.Bookmarks.Item("nota_apertura").Range.Text = nota_apertura
        oDoc.Bookmarks.Item("Nota_chiusura").Range.Text = nota_chiusura
        Try
            oDoc.Bookmarks.Item("Destinatario_spedizione").Range.Text = Destinatario_spedizione
        Catch ex As Exception

        End Try


        oDoc.Bookmarks.Item("Osservazioni").Range.Text = osservazioni
        oDoc.Bookmarks.Item("Alert").Range.Text = alert

        If Sconto_DOC > 0 Then
            If Lingua = "Italian" Then
                parola_sconto = "" & FormatNumber(Sconto_DOC, 2, , , TriState.True) & " %"
            ElseIf Lingua = "French" Then
                parola_sconto = "" & FormatNumber(Sconto_DOC, 2, , , TriState.True) & " %"
            Else
                parola_sconto = "" & FormatNumber(Sconto_DOC, 2, , , TriState.True) & " %"
            End If
        End If

        Totale_netto_pre_iva = Totale_DOC - Valore_sconto

        ' Recupera il range del segnalibro
        Dim rng As Word.Range = oDoc.Bookmarks.Item("tabella").Range

        ' Crea la tabella con 6 righe e 2 colonne
        Dim tbl As Word.Table = oDoc.Tables.Add(rng, 6, 2)
        tbl.Borders.Enable = True

        ' Applica stile "Table Grid Light" (se disponibile nel tuo template)
        Try
            tbl.Style = "Tabella griglia chiara" ' stile italiano
        Catch
            Try
                tbl.Style = "Table Grid Light" ' stile inglese
            Catch
                ' Se lo stile non esiste, nessuna azione
            End Try
        End Try

        ' Imposta larghezza metà documento e allineamento a destra
        tbl.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints
        tbl.PreferredWidth = 226.8 ' ~8 cm
        tbl.Rows.Alignment = Word.WdRowAlignment.wdAlignRowRight

        ' Disabilita lo spezzamento riga per riga
        For Each row As Word.Row In tbl.Rows
            row.AllowBreakAcrossPages = False
        Next

        ' Recupera il paragrafo che contiene la tabella e imposta KeepWithNext e KeepTogether
        Dim tblParagraph As Word.Paragraph = tbl.Range.Paragraphs.First
        With tblParagraph.Format
            .KeepWithNext = True
            .KeepTogether = True
        End With

        ' Inserisci dati
        Dim etichette() As String = {
    "Totale Documento",
    "Sconto",
    "Valore Sconto",
    "Costi di Nolo",
    "IVA",
    "Totale Netto"
}

        Dim valori() As String = {
    Valuta & " " & FormatNumber(Totale_DOC, 2, , , TriState.True),
    parola_sconto,
    Valuta & " " & FormatNumber(Valore_sconto, 2, , , TriState.True),
    Valuta & " " & FormatNumber(spese_di_nolo, 2, , , TriState.True),
    Valuta & " " & FormatNumber(IVA, 2, , , TriState.True),
    Valuta & " " & FormatNumber(Totale_netto_doc, 2, , , TriState.True)
}

        ' Inserisci le celle con formattazione
        For i As Integer = 0 To 5
            tbl.Cell(i + 1, 1).Range.Text = etichette(i)
            tbl.Cell(i + 1, 1).Range.Font.Bold = True
            tbl.Cell(i + 1, 2).Range.Text = valori(i)
            tbl.Cell(i + 1, 2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
        Next



        Dim aggiunta_colonne As Integer = 0
        Verifica_Presenza_produttore()


        If Presenza_produttore > 0 Then
            aggiunta_colonne = aggiunta_colonne + 1

        End If

        Verifica_Presenza_catalogo_fornitore()


        If Presenza_catalogo_fornitore > 0 Then
            aggiunta_colonne = aggiunta_colonne + 1

        End If

        Verifica_Presenza_note()


        If Presenza_note > 0 Then
            aggiunta_colonne = aggiunta_colonne + 1

        End If

        If Verifica_Presenza_disegno(documento_SAP, righe_SAP, docnum) = 1 Then
            aggiunta_colonne = aggiunta_colonne + 1
        End If


        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("Tabella").Range, Max_riga + 2, 9 + aggiunta_colonne)
        Dim c As Integer
        c = 1

        If Lingua = "Italian" Then
            oTable.Cell(1, c).Range.Text = "Riga"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "Codice"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "Descrizione"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "Desc supp"
            c = c + 1
            If Presenza_note > 0 Then
                oTable.Cell(1, c).Range.Text = "Note"
                c = c + 1
            End If
            If Verifica_Presenza_disegno(documento_SAP, righe_SAP, docnum) = 1 Then
                oTable.Cell(1, c).Range.Text = "Disegno"
                c = c + 1
            End If

            If Presenza_produttore > 0 Then
                oTable.Cell(1, c).Range.Text = "Produttore"
                c = c + 1
            End If


            If Presenza_catalogo_fornitore > 0 Then


                oTable.Cell(1, c).Range.Text = "Cod. Forn."
                c = c + 1
            End If
            oTable.Cell(1, c).Range.Text = "UM"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "Q.tà"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "€ Cad"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "€ TOT"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "Consegna"
            c = c + 1

        ElseIf Lingua = "French" Then
            oTable.Cell(1, c).Range.Text = "Riga"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "Codice"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "Descrizione"
            c = c + 1
            If Presenza_note > 0 Then
                oTable.Cell(1, c).Range.Text = "Note"
                c = c + 1
            End If
            oTable.Cell(1, c).Range.Text = "Disegno"
            c = c + 1

            If Presenza_produttore > 0 Then
                oTable.Cell(1, c).Range.Text = "Produttore"
                c = c + 1
            End If

            If Presenza_catalogo_fornitore > 0 Then


                oTable.Cell(1, c).Range.Text = "Cod. Forn."
                c = c + 1
            End If
            oTable.Cell(1, c).Range.Text = "UM"
            c = c + 1

            oTable.Cell(1, c).Range.Text = "Q.tà"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "€ Cad"
            c = c + 1

            oTable.Cell(1, c).Range.Text = "€ TOT"
            c = c + 1

            oTable.Cell(1, c).Range.Text = "Consegna"
            c = c + 1
        Else
            oTable.Cell(1, c).Range.Text = "Riga"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "Codice"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "Descrizione"
            c = c + 1
            If Presenza_note > 0 Then
                oTable.Cell(1, c).Range.Text = "Note"
                c = c + 1
            End If
            oTable.Cell(1, c).Range.Text = "Disegno"
            c = c + 1

            If Presenza_produttore > 0 Then
                oTable.Cell(1, c).Range.Text = "Produttore"
                c = c + 1
            End If

            If Presenza_catalogo_fornitore > 0 Then


                oTable.Cell(1, c).Range.Text = "Cod. Forn."
                c = c + 1
            End If
            oTable.Cell(1, c).Range.Text = "UM"
            c = c + 1

            oTable.Cell(1, c).Range.Text = "Q.tà"
            c = c + 1
            oTable.Cell(1, c).Range.Text = "€ Cad"
            c = c + 1

            oTable.Cell(1, c).Range.Text = "€ TOT"
            c = c + 1

            oTable.Cell(1, c).Range.Text = "Consegna"
            c = c + 1
        End If

        Informazioni_righe_acquisto()




        oTable.AutoFormat(ApplyColor:=False, ApplyBorders:=True)
        oTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        oTable.Rows.Item(1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Rows.Item(1).Range.Font.Bold = True

        ' Set individual cell borders
        For Each cell In oTable.Rows(1).Cells
            cell.Borders(Microsoft.Office.Interop.Word.WdBorderType.wdBorderBottom).LineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        Next cell

        ' Auto-fit the table to the window
        ' oTable.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)


        oTable.Rows.Height = 20

        If documento_SAP = "OPQT" Then
            oDoc.SaveAs(Homepage.percorso_acquisti & "RDO\RDO_" & docnum & ".doc")
        ElseIf documento_SAP = "OPOR" Then
            oDoc.SaveAs(Homepage.percorso_acquisti & "ODA\ODA_" & docnum & ".doc")

        End If


        oWord.Documents.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
        If oWord.Documents.Count = 0 Then
            oWord.Application.Quit()
        End If


        If documento_SAP = "OPQT" Then
            ConvertWordToPDF(Homepage.percorso_acquisti & "RDO\RDO_" & docnum & ".doc")

            percorso_documento_PDF = Homepage.percorso_acquisti & "RDO\RDO_" & docnum & ".PDF"
        ElseIf documento_SAP = "OPOR" Then
            ConvertWordToPDF(Homepage.percorso_acquisti & "ODA\ODA_" & docnum & ".doc")

            percorso_documento_PDF = Homepage.percorso_acquisti & "ODA\ODA_" & docnum & ".PDF"
            percorso_documento_acquisto_per_qualità = Homepage.percorso_acquisti & "ODA\ODA_" & docnum & ".PDF"
        End If




    End Sub

    Sub Genera_documento_NC(par_N_NC As String)




        oWord = CreateObject("Word.Application")
        oDoc = oWord.Documents.Add("" & percorso_documento & "")

        oDoc.Bookmarks.Item("N_nc").Range.Text = par_N_NC
        oDoc.Bookmarks.Item("Data_nc").Range.Text = Today
        oDoc.Bookmarks.Item("rilevatore_nc").Range.Text = Compilatore
        oDoc.Bookmarks.Item("Fornitore").Range.Text = Nome_BP
        oDoc.Bookmarks.Item("Codice_articolo").Range.Text = Codice_sap
        oDoc.Bookmarks.Item("descrizione_articolo").Range.Text = descrizione_sap
        oDoc.Bookmarks.Item("N_disegno").Range.Text = disegno
        oDoc.Bookmarks.Item("quantità_nc").Range.Text = N_pezzi_NC
        oDoc.Bookmarks.Item("definizione_nc").Range.Text = campo_definizione_NC
        oDoc.Bookmarks.Item("a_imputazione_nc").Range.Text = descrizione_NC

        oDoc.Bookmarks.Item("richiesto").Range.Text = osservazioni_nc
        oDoc.Bookmarks.Item("quota_rilevata").Range.Text = rilevato_nc

        oDoc.Bookmarks.Item("data_oa").Range.Text = Data_OA_nc
        oDoc.Bookmarks.Item("N_ordine").Range.Text = OA_nc






        '  oDoc.Bookmarks.Item("Banca").Range.Text = Banca



        'MsgBox(Homepage.PERCORSO_QUALITA & "NC_" & par_N_NC & ".docx")

        Dim percorso_file_new As String

        percorso_file_new = Homepage.PERCORSO_QUALITA & "NC_" & par_N_NC


        oDoc.SaveAs(percorso_file_new & ".docx")



        oWord.Documents.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
        If oWord.Documents.Count = 0 Then
            oWord.Application.Quit()
        End If



        ConvertWordToPDF(percorso_file_new & ".docx")


        Console.WriteLine(percorso_file_new & ".PDF")
        percorso_documento_nc_pdf = percorso_file_new & ".PDF"




    End Sub

    Private Sub ConvertWordToPDF(filename As String)
        Dim wordApplication As New Microsoft.Office.Interop.Word.Application
        Dim wordDocument As Microsoft.Office.Interop.Word.Document = Nothing
        Dim outputFilename As String

        Try
            wordDocument = wordApplication.Documents.Open(filename)
            outputFilename = System.IO.Path.ChangeExtension(filename, "pdf")

            If Not wordDocument Is Nothing Then

                If documento_SAP = "OPOR" Then

                    wordDocument.ExportAsFixedFormat(outputFilename, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, False, Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, 0, 0, Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent, True, True, Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, True, True, False)
                ElseIf documento_SAP = "OPQT" Then
                    wordDocument.ExportAsFixedFormat(outputFilename, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, False, Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, 0, 0, Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent, True, True, Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, True, True, False)
                Else
                    wordDocument.ExportAsFixedFormat(outputFilename, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, True, Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, 0, 0, Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent, True, True, Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, True, True, False)


                End If


            End If
        Catch ex As Exception
            'TODO: handle exception
        Finally
            If Not wordDocument Is Nothing Then
                wordDocument.Close(False)
                wordDocument = Nothing
            End If

            If Not wordApplication Is Nothing Then
                wordApplication.Quit()
                wordApplication = Nothing
            End If
        End Try

    End Sub

    Sub Informazioni_righe_acquisto()

        Dim r As Integer = 2

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "
declare @lingua as varchar(20)
declare @docnum as integer

set @lingua ='" & Lingua & "'
set @docnum= " & docnum & "

SELECT case when T1.[ItemCode] is null then '' else t1.itemcode end as 'Itemcode', T1.[Dscription],
coalesce(t2.frgnname,'') as 'frgnname',
T1.FREETXT,T1.U_DISEGNO,T3.FIRMNAME,T2.SUPPCATNUM,T2.BUYUNITMSR,T1.QUANTITY, T1.PRICE,T1.LINETOTAL, case when '" & documento_SAP & "' ='OPOR' then T1.SHIPDATE else t1.pqtreqdate end as 'shipdate'
FROM " & documento_SAP & " T0  INNER JOIN " & righe_SAP & " T1 ON T0.[DocEntry] = T1.[DocEntry]
LEFT JOIN OITM T2 ON T2.ITEMCODE=T1.ITEMCODE
LEFT JOIN OMRC T3 ON T3.FIRMCODE=T2.FIRMCODE WHERE T0.[DocNum] =@docnum 
        
ORDER BY T1.VISORDER"


        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Dim c As Integer
        Do While cmd_SAP_reader.Read()
            c = 1

            oTable.Cell(r, c).Range.Text = r - 1
            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            c = c + 1
            oTable.Cell(r, c).Range.Text = cmd_SAP_reader("ItemCode")
            c = c + 1



            If Not cmd_SAP_reader("Dscription") Is System.DBNull.Value Then
                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("Dscription")
            Else
                oTable.Cell(r, c).Range.Text = ""

            End If

            With oTable.Cell(r, c).Range.Font
                .Name = "Arial"
                .Bold = 1



                If Len(oTable.Cell(r, c).Range.Text) < 15 Then
                    .Size = 8
                ElseIf Len(oTable.Cell(r, c).Range.Text) < 25 Then
                    .Size = 7
                Else
                    .Size = 6
                End If


            End With


            c = c + 1
            oTable.Cell(r, c).Range.Text = cmd_SAP_reader("frgnname")
            With oTable.Cell(r, c).Range.Font
                .Name = "Arial"
                .Bold = 1



                If Len(oTable.Cell(r, c).Range.Text) < 15 Then
                    .Size = 8
                ElseIf Len(oTable.Cell(r, c).Range.Text) < 25 Then
                    .Size = 7
                Else
                    .Size = 6
                End If


            End With

            c = c + 1
            If Presenza_note > 0 Then

                If Not cmd_SAP_reader("freetxt") Is System.DBNull.Value Then
                    oTable.Cell(r, c).Range.Text = cmd_SAP_reader("freetxt")
                Else
                    oTable.Cell(r, c).Range.Text = ""

                End If
                c = c + 1
            End If
            If Verifica_Presenza_disegno(documento_SAP, righe_SAP, docnum) = 1 Then
                If Not cmd_SAP_reader("u_disegno") Is System.DBNull.Value Then
                    oTable.Cell(r, c).Range.Text = cmd_SAP_reader("u_disegno")
                Else
                    oTable.Cell(r, c).Range.Text = ""

                End If
                c = c + 1
            End If


            If Presenza_produttore > 0 Then
                If Not cmd_SAP_reader("firmname") Is System.DBNull.Value Then
                    oTable.Cell(r, c).Range.Text = cmd_SAP_reader("firmname")
                Else
                    oTable.Cell(r, c).Range.Text = ""
                End If
                c = c + 1

            End If

            If Presenza_catalogo_fornitore > 0 Then


                If Not cmd_SAP_reader("suppcatnum") Is System.DBNull.Value Then
                    oTable.Cell(r, c).Range.Text = cmd_SAP_reader("suppcatnum")
                Else
                    oTable.Cell(r, c).Range.Text = ""
                End If
                c = c + 1
            End If

            If Not cmd_SAP_reader("BUYUNITMSR") Is System.DBNull.Value Then
                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("BUYUNITMSR")
            Else
                oTable.Cell(r, c).Range.Text = ""
            End If

            c = c + 1

            oTable.Cell(r, c).Range.Text = FormatNumber(cmd_SAP_reader("QUANTITY"), 2, , , TriState.True)

            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter

            c = c + 1

            If cmd_SAP_reader("Price") = 0 Then
                oTable.Cell(r, c).Range.Text = ""
            Else
                ' oTable.Cell(r, c).Range.Text = Valuta & " " & FormatNumber(cmd_SAP_reader("Price"), 2, , , TriState.True)
                oTable.Cell(r, c).Range.Text = FormatNumber(cmd_SAP_reader("Price"), 2, , , TriState.True)

            End If


            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
            c = c + 1

            If cmd_SAP_reader("linetotal") = 0 Then
                oTable.Cell(r, c).Range.Text = ""
            Else
                ' oTable.Cell(r, c).Range.Text = Valuta & " " & FormatNumber(cmd_SAP_reader("linetotal"), 2, , , TriState.True)
                oTable.Cell(r, c).Range.Text = FormatNumber(cmd_SAP_reader("linetotal"), 2, , , TriState.True)

            End If


            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
            c = c + 1

            If Not cmd_SAP_reader("shipdate") Is System.DBNull.Value Then
                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("shipdate")
            Else
                oTable.Cell(r, c).Range.Text = ""
            End If

            c = c + 1

            r = r + 1


        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub ' Informazioni righe offerta



    Sub Informazioni_righe_offerta(par_codice_brb As String)


        Dim r As Integer = 2
        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "
declare @lingua as varchar(20)
declare @docnum as integer

set @lingua ='" & Lingua & "'
set @docnum= " & docnum & "

SELECT T1.[ItemCode] as 'Codice',
coalesce(t6.substitute,'') as 'Codice_bp',
CASE WHEN T1.ITEMCODE ='F99999' THEN case when (T1.DSCRIPTION ='' or t1.Dscription is null) then t1.u_note else T1.DSCRIPTION end  ELSE
CASE WHEN @lingua='Italian' then
	
 T1.[Dscription] 

	when @lingua='French' then 
		 t2.U_DESCFR
		
		else  t2.U_DESCING  END END  as 'Descrizione',
case when T1.u_prg_azs_commessa is null then '' else T1.u_prg_azs_commessa end as 'Commessa',
RTRIM(CAST(CAST(T1.[Quantity] AS DECIMAL(10, 2)) AS FLOAT)) AS 'Quantità',
CASE WHEN T1.[DiscPrcnt] >0 THEN 
CASE WHEN T0.DocCur<>'EUR' AND T1.CURRENCY ='EUR' THEN T1.[Pricebefdi]*t0.docrate else
T1.[Pricebefdi] end
ELSE 
CASE WHEN T0.DocCur<>'EUR' AND T1.CURRENCY ='EUR' THEN
T1.[Pricebefdi]*t0.docrate else
T1.[Pricebefdi] end
END  as 'Prezzo unitario',
coalesce(CASE WHEN T0.DocCur<>'EUR' AND T1.CURRENCY ='EUR' THEN t1.price*t0.docrate else t1.price end,0)  as 'Prezzo_dopo_sconto',
case when T1.[DiscPrcnt] is null then 0 else T1.[DiscPrcnt] end  as 'Sconto',
case when t0.DOCCUR='EUR' then t1.linetotal else t1.linetotal*t0.docrate end as'Totale',
case when t1.u_approvvigionamento_articolo is null then '' else t1.u_approvvigionamento_articolo end as 'u_approvvigionamento_articolo',
case when t2.u_codice_ktf is null then '' else t2.u_codice_ktf end as 'Codice_KTF',
coalesce(t2.u_codice_brb,'') as 'Codice_brb',
case when t3.[ISOriCntry] is null then '' else t3.[ISOriCntry] end as'Made_in',
case when T4.Code is null then '' else T4.Code end as'Codice_doganale',
case when t1.freetxt is null then '' else t1.freetxt end as'Note',
case when t2.itemname is null then '' else t2.itemname end as 'Desc_ITA'
,t1.linenum 
,t0.docentry
, COALESCE(T5.LINETEXT,'') AS 'RIGA_TESTO'

FROM " & documento_SAP & " T0  INNER JOIN " & righe_SAP & " T1 ON T0.[DocEntry] = T1.[DocEntry] 
left join oitm t2 on t2.itemcode=t1.itemcode
left join itm10 t3 on t3.itemcode=t1.itemcode
left join ODCI T4 on t4.AbsEntry=t3.ISCommCode
left join QUT10 T5 ON T5.DOCENTRY=T0.DOCENTRY AND '" & documento_SAP & "' ='OQUT' AND T0.OBJTYPE='23' AND T5.AFTLINENUM=T1.visorder-1
left join oscn t6 on t6.itemcode=t1.itemcode and t6.cardcode=t0.cardcode

       
        WHERE T0.[DocNum] =@docnum 
ORDER BY T1.VISORDER"


        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Dim c As Integer


        Do While cmd_SAP_reader.Read()

            If cmd_SAP_reader("RIGA_TESTO") <> "" Then
                'Dim newRow As Word.Row = oTable.Rows.Add()

                '' Imposta il testo nella cella della nuova riga
                'newRow.Cells(1).Range.Text = cmd_SAP_reader("RIGA_TESTO")

                '' Unisci tutte le celle della nuova riga
                'newRow.Cells.Merge()


                ' Aggiungi una nuova riga alla tabella
                oTable.Rows.Add()
                ' Imposta il testo nella cella della riga corrente
                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("RIGA_TESTO")
                ' Unisci tutte le celle della riga corrente
                oTable.Rows(r).Cells.Merge()
                ' Incrementa il contatore della riga
                oTable.Rows(r).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter

                r += 1

            End If
            c = 1


            oTable.Cell(r, c).Range.Text = r - 1
            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1
            If cmd_SAP_reader("Codice_BP") <> "" Then
                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("Codice_bp") & vbCrLf & cmd_SAP_reader("Codice")
            Else
                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("Codice")
            End If


            oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1

            If par_codice_brb = True Then



                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("codice_BRB")
                oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If

            If Not cmd_SAP_reader("Descrizione") Is System.DBNull.Value Then
                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("Descrizione")
                oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            Else

                descrizione_temp = InputBox("Il codice " & cmd_SAP_reader("Codice") & " " & cmd_SAP_reader("Desc_ita") & " ha descrizione nulla, inserire la descrizione che si vuole visualizzare")

                oTable.Cell(r, c).Range.Text = UCase(descrizione_temp)
                oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                codice_temp = cmd_SAP_reader("Codice")
                aggiorna_traduzione_codice(codice_temp, descrizione_temp, Lingua)
            End If

            With oTable.Cell(r, c).Range.Font
                .Name = "Arial"
                .Bold = 1


                If Len(oTable.Cell(r, c).Range.Text) < 35 Then
                    .Size = 9
                ElseIf Len(oTable.Cell(r, c).Range.Text) < 40 Then
                    .Size = 8
                Else
                    .Size = 7
                End If


            End With
            c = c + 1
            If nome_documento_SAP <> "Packing_list" Then
                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("Commessa")
                oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1
            End If

            oTable.Cell(r, c).Range.Text = cmd_SAP_reader("quantità")
            oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
            oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
            c = c + 1
            If nome_documento_SAP <> "Proforma" And nome_documento_SAP <> "Packing_list" And nome_documento_SAP <> "Consegna" Then
                oTable.Cell(r, c).Range.Text = Valuta & " " & FormatNumber(cmd_SAP_reader("Prezzo unitario"), 2, , , TriState.True)
                oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1
            ElseIf nome_documento_SAP = "Packing_list" Then

            Else
                oTable.Cell(r, c).Range.Text = Valuta & " " & FormatNumber(cmd_SAP_reader("Prezzo_dopo_sconto"), 2, , , TriState.True)
                oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1
            End If



            If Presenza_sconto_righe > 0 And nome_documento_SAP <> "Proforma" And nome_documento_SAP <> "Invoice" And nome_documento_SAP <> "Packing_list" And nome_documento_SAP <> "Consegna" Then
                oTable.Cell(r, c).Range.Text = FormatPercent(cmd_SAP_reader("sconto") / 100)
                oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1
            End If
            If nome_documento_SAP <> "Packing_list" Then
                oTable.Cell(r, c).Range.Text = Valuta & " " & FormatNumber(cmd_SAP_reader("Totale"), 2, , , TriState.True)
                oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1
            End If

            If made_in = True Then
                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("made_in")
                oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1
            End If

            If codice_doganale_riga = True Then
                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("codice_doganale")
                oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1
            End If

            If note = "Y" Then
                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("note")
                oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1
            End If

            If approvvigionamento_articolo = True Then


                If documento_SAP = "OQUT" Or documento_SAP = "ORDR" Then
                    oTable.Cell(r, c).Range.Text = cmd_SAP_reader("u_approvvigionamento_articolo")
                    oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
                    oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    c = c + 1
                End If
            End If

            If codice_KTF = True Then



                oTable.Cell(r, c).Range.Text = cmd_SAP_reader("codice_KTF")
                oTable.Cell(r, c).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
                oTable.Cell(r, c).VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                c = c + 1

            End If



            r = r + 1


        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub ' Informazioni righe offerta

    Function riga_di_testo(par_docentry As Integer, par_linenum As Integer)


        Dim stringa As String = ""
        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "SELECT T0.DocEntry, T0.LineSeq, T0.AftLineNum, T0.OrderNum, T0.LineType, T0.LineText, T0.ObjType
FROM QUT10 T0

where t0.docentry=" & par_docentry & " and T0.AftLineNum = " & par_linenum & "-1"


        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Dim c As Integer


        If cmd_SAP_reader.Read() Then
            stringa = cmd_SAP_reader("LineText")
        Else
            stringa = ""

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
        Return stringa
    End Function ' Informazioni righe offerta


    Sub max_righe_documento()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT t20.max_linenum + SUM(CASE WHEN T2.LINETEXT IS NULL THEN 0 ELSE 1 END) + SUM(CASE WHEN T3.LINETEXT IS NULL THEN 0 ELSE 1 END) as 'Max_riga'
from
(
SELECT max(T1.[visorder]) as 'MAX_linenum', t0.docentry

        FROM " & documento_SAP & " T0  INNER JOIN " & righe_SAP & " T1 ON T0.[DocEntry] = T1.[DocEntry] 
        
        WHERE T0.[DocNum] =" & docnum & "
        group by t0.docentry
)
AS T20 left join qut10 t2 on t2.docentry=t20.docentry and '" & righe_SAP & "' ='QUT1'
         left join RDR10 t3 on t3.docentry=t20.docentry and '" & righe_SAP & "' ='RDR1'
group by t20.MAX_linenum"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then
            Max_riga = cmd_SAP_reader("Max_riga")

            cmd_SAP_reader.Close()

        Else
            MsgBox("La query max_righe_documento non sta funzionando")
        End If
        Cnn.Close()

    End Sub 'Individuo numero di righe 'Verifico presenza di sconto nel documento

    Sub trova_word_base(par_lingua As String, par_documento_sap As String, par_garanzia As String, par_nome_documento_sap As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "
declare @lingua as varchar (20)
declare @Documento as varchar (20)
declare @garanzia as varchar (20)
declare @tipo_documento as varchar (20)
declare @azienda as varchar(10)
 
set @lingua ='" & par_lingua & "'
set @Documento ='" & par_documento_sap & "'
set @Garanzia ='" & par_garanzia & "'
set @tipo_documento ='" & par_nome_documento_sap & "'


select top 1 t0.percorso
from [Tirelli_40].[dbo].[Percorso_layout_documenti] t0
where t0.lingua=@lingua and t0.documento=@Documento and t0.Garanzia=@garanzia and t0.tipo_documento=@tipo_documento and t0.active='Y'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then
            percorso_documento = cmd_SAP_reader("Percorso")



        Else
            MsgBox("La query trova word base non sta funzionando")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Individuo numero di righe 'Verifico presenza di sconto nel documento



    Sub prepara_etichette_excel(par_percorso_file As String, par_nome_foglio As String, par_lingua As String, par_documento As String, par_numero_documento As String, par_testata_sap As String, par_righe_sap As String)
        Dim appExcel As New Excel.Application
        Dim workbook As Excel.Workbook

        Try
            ' Apri il file Excel
            workbook = appExcel.Workbooks.Open(par_percorso_file)
            'appExcel.Visible = True
            Dim colonna_traduzioni As Integer
            If par_lingua = "Italian" Then
                colonna_traduzioni = 1
            ElseIf par_lingua = "English" Then
                colonna_traduzioni = 2
            ElseIf par_lingua = "French" Then
                colonna_traduzioni = 3

            ElseIf par_lingua = "Spanish" Then
                colonna_traduzioni = 4

            Else
                colonna_traduzioni = 2
            End If

            Dim testo_base As String
            Dim testo_grassetto As String
            Dim cella As Excel.Range

            'cliente


            Dim testo_destinatario As String

            Dim lunghezza_base As Long
            Dim lunghezza_grassetto As Long

            ' Valori iniziali
            testo_base = workbook.Sheets("Traduzioni").Cells(2, colonna_traduzioni).Value
            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).nome_bp
            testo_destinatario = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).destinatario_fattura

            ' Determina lunghezze
            lunghezza_base = Len(testo_base) + 2 ' Include ": " e vbCrLf
            lunghezza_grassetto = Len(testo_grassetto)

            ' Imposta il valore della cella
            cella = workbook.Sheets(par_nome_foglio).Range("B3")
            cella.Value = testo_base & ": " & vbCrLf & testo_grassetto & vbCrLf & testo_destinatario

            ' Applica il grassetto solo alla parte desiderata
            cella.Characters(Start:=lunghezza_base + 1, Length:=lunghezza_grassetto + 2).Font.Bold = True
            cella.Characters(Start:=lunghezza_base + 1, Length:=lunghezza_grassetto + 2).Font.Color = RGB(30, 30, 30)
            ''indirizzo fatturazione

            'testo_base = workbook.Sheets("Traduzioni").Cells(3, colonna_traduzioni).Value
            'testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).destinatario_fattura
            'cella = workbook.Sheets(par_nome_foglio).Range("B4")
            'cella.Value = testo_base & ":" & vbCrLf & testo_grassetto
            'cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

            'indirizzo destinazione

            testo_base = workbook.Sheets("Traduzioni").Cells(4, colonna_traduzioni).Value
            If trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).nome_indirizzo = "" Then
                testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).destinatario_spedizione
            Else
                'testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).nome_indirizzo  & " - " & trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).destinatario_spedizione
                testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).shiptocode & " - " & trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).destinatario_spedizione
            End If

            cella = workbook.Sheets(par_nome_foglio).Range("B4")
            cella.Value = testo_base & ":" & vbCrLf & testo_grassetto
            cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True


            'indirizzo Rif cliente

            testo_base = workbook.Sheets("Traduzioni").Cells(5, colonna_traduzioni).Value
            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).riferimento_cliente
            cella = workbook.Sheets(par_nome_foglio).Range("E4")
            cella.Value = testo_base & ":" & vbCrLf & testo_grassetto
            cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

            'Documento 


            If par_documento = "Invoice" Then
                workbook.Sheets(par_nome_foglio).Range("E3").Value = workbook.Sheets("Traduzioni").Cells(6, colonna_traduzioni).Value
                testo_base = workbook.Sheets("Traduzioni").Cells(6, colonna_traduzioni).Value
            ElseIf par_documento = "Order" Then
                testo_base = workbook.Sheets("Traduzioni").Cells(8, colonna_traduzioni).Value

            ElseIf par_documento = "Off" Then
                testo_base = workbook.Sheets("Traduzioni").Cells(9, colonna_traduzioni).Value
            Else
                testo_base = ""
            End If

            testo_grassetto = par_numero_documento
            cella = workbook.Sheets(par_nome_foglio).Range("E3")
            cella.Value = testo_base & ": " & testo_grassetto
            cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

            'Data

            testo_base = workbook.Sheets("Traduzioni").Cells(7, colonna_traduzioni).Value
            testo_grassetto = Format(Now, "dd/MM/yyyy")
            cella = workbook.Sheets(par_nome_foglio).Range("I3")
            cella.Value = testo_base & ": " & testo_grassetto
            cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

            ' Data consegna


            If par_documento = "Order" Then
                testo_base = workbook.Sheets("Traduzioni").Cells(10, colonna_traduzioni).Value
            ElseIf par_documento = "Off" Then
                testo_base = workbook.Sheets("Traduzioni").Cells(38, colonna_traduzioni).Value
            Else
                testo_base = ""
            End If

            If par_documento = "Order" Or par_documento = "Off" Then

                testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).docduedate
                cella = workbook.Sheets(par_nome_foglio).Range("I4")
                cella.Value = testo_base & ": " & testo_grassetto
                cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True
            End If


            ' incoterms

            testo_base = workbook.Sheets("Traduzioni").Cells(11, colonna_traduzioni).Value
            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).Incoterms
            cella = workbook.Sheets(par_nome_foglio).Range("B7")
            cella.Value = testo_base & ": " & testo_grassetto
            cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

            'Spedizione

            testo_base = workbook.Sheets("Traduzioni").Cells(12, colonna_traduzioni).Value
            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).Spedizione
            cella = workbook.Sheets(par_nome_foglio).Range("B8")
            cella.Value = testo_base & ": " & testo_grassetto
            cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

            'Vettore

            testo_base = workbook.Sheets("Traduzioni").Cells(13, colonna_traduzioni).Value
            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).vettore
            cella = workbook.Sheets(par_nome_foglio).Range("B9")
            cella.Value = testo_base & ": " & testo_grassetto
            cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

            'Pagamento

            testo_base = workbook.Sheets("Traduzioni").Cells(14, colonna_traduzioni).Value
            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).pagamento
            cella = workbook.Sheets(par_nome_foglio).Range("B10")
            cella.Value = testo_base & ": " & testo_grassetto
            cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

            'Banca

            testo_base = workbook.Sheets("Traduzioni").Cells(15, colonna_traduzioni).Value
            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).Banca
            cella = workbook.Sheets(par_nome_foglio).Range("F7")
            cella.Value = testo_base & ": " & testo_grassetto
            cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

            'Iban
            cella = workbook.Sheets(par_nome_foglio).Range("f8")
            testo_base = workbook.Sheets("Traduzioni").Cells(16, colonna_traduzioni).Value
            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).IBAN
            cella.Value = testo_base & ": " & testo_grassetto
            cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

            'Ns contatto

            testo_base = workbook.Sheets("Traduzioni").Cells(17, colonna_traduzioni).Value
            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).Compilatore
            cella = workbook.Sheets(par_nome_foglio).Range("f9")
            cella.Value = testo_base & ": " & testo_grassetto
            cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

            'Vs contatto

            testo_base = workbook.Sheets("Traduzioni").Cells(18, colonna_traduzioni).Value
            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).Contatto
            cella = workbook.Sheets(par_nome_foglio).Range("f10")
            cella.Value = testo_base & ": " & testo_grassetto
            cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

            Dim RIGA_SOTTO As Integer = 16

            'Subtotale


            testo_base = workbook.Sheets("Traduzioni").Cells(19, colonna_traduzioni).Value
            'testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).IVA
            cella = workbook.Sheets(par_nome_foglio).Range("G" & RIGA_SOTTO)
            cella.Value = testo_base & ": "
            workbook.Sheets(par_nome_foglio).Range("G17").Font.Bold = True


            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).totale_doc
            cella = workbook.Sheets(par_nome_foglio).Range("I" & RIGA_SOTTO)
            cella.Value = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).valuta & " " & FormatNumber(testo_grassetto, 2, , , TriState.True)
            workbook.Sheets(par_nome_foglio).Range("I" & RIGA_SOTTO).Font.Bold = True

            'Sconto

            If trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).sconto_doc > 0 Then



                testo_base = workbook.Sheets("Traduzioni").Cells(22, colonna_traduzioni).Value
                testo_grassetto = FormatNumber(trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).sconto_doc, 2, , , TriState.True) & " %"
                cella = workbook.Sheets(par_nome_foglio).Range("G" & RIGA_SOTTO + 1)
                cella.Value = testo_base & ":  " & testo_grassetto
                workbook.Sheets(par_nome_foglio).Range("G" & RIGA_SOTTO + 1).Font.Bold = True


                testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).valore_sconto
                cella = workbook.Sheets(par_nome_foglio).Range("I" & RIGA_SOTTO + 1)
                cella.Value = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).valuta & " " & FormatNumber(testo_grassetto, 2, , , TriState.True)
                workbook.Sheets(par_nome_foglio).Range("I" & RIGA_SOTTO + 1).Font.Bold = True

            End If


            'IVA

            testo_base = workbook.Sheets("Traduzioni").Cells(20, colonna_traduzioni).Value
            'testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).IVA
            cella = workbook.Sheets(par_nome_foglio).Range("G" & RIGA_SOTTO + 2)
            cella.Value = testo_base & ": "
            workbook.Sheets(par_nome_foglio).Range("G" & RIGA_SOTTO + 2).Font.Bold = True


            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).IVA
            cella = workbook.Sheets(par_nome_foglio).Range("I" & RIGA_SOTTO + 2)
            cella.Value = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).valuta & " " & FormatNumber(testo_grassetto, 2, , , TriState.True)
            workbook.Sheets(par_nome_foglio).Range("I" & RIGA_SOTTO + 2).Font.Bold = True

            'TOTALE

            testo_base = workbook.Sheets("Traduzioni").Cells(21, colonna_traduzioni).Value
            cella = workbook.Sheets(par_nome_foglio).Range("G" & RIGA_SOTTO + 3)
            cella.Value = testo_base & ": "
            workbook.Sheets(par_nome_foglio).Range("G" & RIGA_SOTTO + 3).Font.Bold = True


            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).totale_netto_doc
            cella = workbook.Sheets(par_nome_foglio).Range("I" & RIGA_SOTTO + 3)
            cella.Value = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).valuta & " " & FormatNumber(testo_grassetto, 2, , , TriState.True)
            workbook.Sheets(par_nome_foglio).Range("I" & RIGA_SOTTO + 3).Font.Bold = True

            'Lead time
            workbook.Sheets(par_nome_foglio).Range("B" & RIGA_SOTTO).Value = workbook.Sheets("Traduzioni").Cells(24, colonna_traduzioni).Value
            testo_base = workbook.Sheets("Traduzioni").Cells(24, colonna_traduzioni).Value

            cella = workbook.Sheets(par_nome_foglio).Range("B" & RIGA_SOTTO)
            cella.Value = testo_base
            'cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True
            'workbook.Sheets(par_nome_foglio).Range("H19").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            'workbook.Sheets(par_nome_foglio).Range("H19").VerticalAlignment = Excel.XlVAlign.xlVAlignTop

            'Note
            workbook.Sheets(par_nome_foglio).Range("B" & RIGA_SOTTO + 2).Value = workbook.Sheets("Traduzioni").Cells(23, colonna_traduzioni).Value
            testo_base = workbook.Sheets("Traduzioni").Cells(23, colonna_traduzioni).Value
            testo_grassetto = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).osservazioni
            cella = workbook.Sheets(par_nome_foglio).Range("B" & RIGA_SOTTO + 2)
            cella.Value = testo_base & ": " & testo_grassetto
            'cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True
            ' workbook.Sheets(par_nome_foglio).Range("H16").HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            ' workbook.Sheets(par_nome_foglio).Range("H16").VerticalAlignment = Excel.XlVAlign.xlVAlignTop


            'Riga
            cella = workbook.Sheets(par_nome_foglio).Range("B" & RIGA_SOTTO - 4)
            testo_base = workbook.Sheets("Traduzioni").Cells(25, colonna_traduzioni).Value

            cella.Value = testo_base
            cella.Font.Bold = True
            workbook.Sheets(par_nome_foglio).Range("B" & RIGA_SOTTO - 4).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            workbook.Sheets(par_nome_foglio).Range("B" & RIGA_SOTTO - 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            'Codice
            cella = workbook.Sheets(par_nome_foglio).Range("C" & RIGA_SOTTO - 4)
            testo_base = workbook.Sheets("Traduzioni").Cells(26, colonna_traduzioni).Value

            cella.Value = testo_base
            cella.Font.Bold = True
            workbook.Sheets(par_nome_foglio).Range("C" & RIGA_SOTTO - 4).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            workbook.Sheets(par_nome_foglio).Range("C" & RIGA_SOTTO - 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter


            'Descrizione
            cella = workbook.Sheets(par_nome_foglio).Range("D" & RIGA_SOTTO - 4)
            testo_base = workbook.Sheets("Traduzioni").Cells(28, colonna_traduzioni).Value

            cella.Value = testo_base
            cella.Font.Bold = True
            workbook.Sheets(par_nome_foglio).Range("D" & RIGA_SOTTO - 4).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            workbook.Sheets(par_nome_foglio).Range("D" & RIGA_SOTTO - 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter



            'Matricola
            cella = workbook.Sheets(par_nome_foglio).Range("E" & RIGA_SOTTO - 4)
            testo_base = workbook.Sheets("Traduzioni").Cells(29, colonna_traduzioni).Value

            cella.Value = testo_base
            cella.Font.Bold = True
            workbook.Sheets(par_nome_foglio).Range("E" & RIGA_SOTTO - 4).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            workbook.Sheets(par_nome_foglio).Range("E" & RIGA_SOTTO - 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            'Quantità
            cella = workbook.Sheets(par_nome_foglio).Range("F" & RIGA_SOTTO - 4)
            testo_base = workbook.Sheets("Traduzioni").Cells(30, colonna_traduzioni).Value

            cella.Value = testo_base
            cella.Font.Bold = True
            workbook.Sheets(par_nome_foglio).Range("F" & RIGA_SOTTO - 4).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            workbook.Sheets(par_nome_foglio).Range("F" & RIGA_SOTTO - 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            'Prezzo U
            cella = workbook.Sheets(par_nome_foglio).Range("G" & RIGA_SOTTO - 4)
            testo_base = workbook.Sheets("Traduzioni").Cells(31, colonna_traduzioni).Value

            cella.Value = testo_base
            cella.Font.Bold = True
            workbook.Sheets(par_nome_foglio).Range("G" & RIGA_SOTTO - 4).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            workbook.Sheets(par_nome_foglio).Range("G" & RIGA_SOTTO - 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            'Sconto
            cella = workbook.Sheets(par_nome_foglio).Range("H" & RIGA_SOTTO - 4)
            testo_base = workbook.Sheets("Traduzioni").Cells(32, colonna_traduzioni).Value

            cella.Value = testo_base
            cella.Font.Bold = True
            workbook.Sheets(par_nome_foglio).Range("H" & RIGA_SOTTO - 4).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            workbook.Sheets(par_nome_foglio).Range("H" & RIGA_SOTTO - 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            'Totale
            cella = workbook.Sheets(par_nome_foglio).Range("I" & RIGA_SOTTO - 4)
            testo_base = workbook.Sheets("Traduzioni").Cells(33, colonna_traduzioni).Value

            cella.Value = testo_base
            cella.Font.Bold = True
            workbook.Sheets(par_nome_foglio).Range("I" & RIGA_SOTTO - 4).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            workbook.Sheets(par_nome_foglio).Range("I" & RIGA_SOTTO - 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            'Cod KTF
            cella = workbook.Sheets(par_nome_foglio).Range("J" & RIGA_SOTTO - 4)
            testo_base = workbook.Sheets("Traduzioni").Cells(34, colonna_traduzioni).Value

            cella.Value = testo_base
            cella.Font.Bold = True
            workbook.Sheets(par_nome_foglio).Range("J" & RIGA_SOTTO - 4).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            workbook.Sheets(par_nome_foglio).Range("J" & RIGA_SOTTO - 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            'Origine merce
            cella = workbook.Sheets(par_nome_foglio).Range("K" & RIGA_SOTTO - 4)
            testo_base = workbook.Sheets("Traduzioni").Cells(35, colonna_traduzioni).Value

            cella.Value = testo_base
            cella.Font.Bold = True
            workbook.Sheets(par_nome_foglio).Range("K" & RIGA_SOTTO - 4).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            workbook.Sheets(par_nome_foglio).Range("K" & RIGA_SOTTO - 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            'Cod doganale
            cella = workbook.Sheets(par_nome_foglio).Range("L" & RIGA_SOTTO - 4)
            testo_base = workbook.Sheets("Traduzioni").Cells(36, colonna_traduzioni).Value

            cella.Value = testo_base
            cella.Font.Bold = True
            workbook.Sheets(par_nome_foglio).Range("L" & RIGA_SOTTO - 4).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter
            workbook.Sheets(par_nome_foglio).Range("L" & RIGA_SOTTO - 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

            'Lead time
            cella = workbook.Sheets(par_nome_foglio).Range("M" & RIGA_SOTTO - 4)
            testo_base = workbook.Sheets("Traduzioni").Cells(37, colonna_traduzioni).Value

            cella.Value = testo_base
            cella.Font.Bold = True
            workbook.Sheets(par_nome_foglio).Range("M" & RIGA_SOTTO - 4).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            workbook.Sheets(par_nome_foglio).Range("M" & RIGA_SOTTO - 4).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter




            ' Chiama la funzione "righe" passando il workbook aperto
            righe(workbook, par_nome_foglio, 13, 2, par_lingua, par_documento, par_numero_documento, par_testata_sap, par_righe_sap, trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).valuta, codice_KTF, made_in, codice_doganale_riga, approvvigionamento_articolo, codice_BRB)


            'Esporta i fogli specifici In PDF, escludendo il foglio "traduzioni"
            Dim gtc_da_usare As String = ""


            If trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).LINGUA = "Italian" Then
                gtc_da_usare = "GTC ITA"
                'ElseIf trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).LINGUA = "Spanish" Then
                '    gtc_da_usare = "GTC SPA"
            ElseIf trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).LINGUA = "French" Then
                gtc_da_usare = "GTC FR"
            Else
                gtc_da_usare = "GTC ENG"
            End If

            Dim pdfFilePath As String = EsportaFogliInPDF(workbook, par_nome_foglio, gtc_da_usare, par_percorso_file)

            ' Apri il PDF creato
            If Not String.IsNullOrEmpty(pdfFilePath) Then
                Process.Start(pdfFilePath)
            Else
                MsgBox("Errore nell'esportazione del PDF.")
            End If

        Catch ex As Exception
            MsgBox("Errore: " & ex.Message)

        Finally
            '  Chiudi il workbook e l'applicazione Excel
            If workbook IsNot Nothing Then workbook.Close(False)
            appExcel.Quit()

            ' Rilascia le risorse
            ReleaseObject(workbook)
            ReleaseObject(appExcel)
        End Try
    End Sub



    Sub Apri_file_Excel(par_percorso_file As String)
        Dim appExcel As New Excel.Application
        Dim workbook As Excel.Workbook = Nothing

        Try
            ' Apri il file Excel
            workbook = appExcel.Workbooks.Open(par_percorso_file)
            appExcel.Visible = True

            ' 🔹 Aggiorna tutte le QueryTables nei fogli
            For Each ws As Excel.Worksheet In workbook.Sheets
                For Each qt As Excel.QueryTable In ws.QueryTables
                    qt.Refresh(False)
                Next
            Next

            ' 🔹 Aggiorna tutte le connessioni dati (se presenti)
            For Each conn In workbook.Connections
                conn.OLEDBConnection.BackgroundQuery = False ' Aspetta il completamento
                conn.Refresh()
            Next

            ' 🔹 Aspetta che Excel abbia completato i calcoli
            Do While appExcel.CalculationState <> Excel.XlCalculationState.xlDone
                System.Threading.Thread.Sleep(500) ' Attesa attiva ogni 500ms
            Loop

            ' Ora esporta in PDF
            Dim pdfFilePath As String = Esporta_in_pdf(workbook, par_percorso_file)

            ' Apri il PDF creato
            If Not String.IsNullOrEmpty(pdfFilePath) Then
                Process.Start(pdfFilePath)
            Else
                MsgBox("Errore nell'esportazione del PDF.")
            End If

        Catch ex As Exception
            MsgBox("Errore: " & ex.Message)

        Finally
            ' Chiudi il workbook e l'applicazione Excel
            If workbook IsNot Nothing Then workbook.Close(False)
            appExcel.Quit()

            ' Rilascia le risorse
            ReleaseObject(workbook)
            ReleaseObject(appExcel)
        End Try
    End Sub

    Function Esporta_in_pdf(workbook As Excel.Workbook, par_percorso_file As String) As String
        Dim pdfFilePath As String = ""

        Try
            ' Creiamo una lista di fogli visibili
            Dim fogliDaStampare As New List(Of Excel.Worksheet)

            For Each ws As Excel.Worksheet In workbook.Sheets
                If ws.Visible = Excel.XlSheetVisibility.xlSheetVisible Then
                    ' Adatta l'area di stampa al contenuto
                    ws.PageSetup.Zoom = False
                    ws.PageSetup.FitToPagesWide = 1
                    ws.PageSetup.FitToPagesTall = False
                    ws.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait

                    ' Imposta l'area di stampa solo se non è già definita
                    If ws.UsedRange.Rows.Count > 0 And ws.UsedRange.Columns.Count > 0 Then
                        ws.PageSetup.PrintArea = ws.UsedRange.Address
                    End If

                    fogliDaStampare.Add(ws)
                End If
            Next

            ' Se non ci sono fogli visibili, interrompe l'esecuzione
            If fogliDaStampare.Count = 0 Then
                MsgBox("Errore: Nessun foglio visibile da esportare.")
                Return ""
            End If

            ' Determina il nome del file PDF
            Dim basePdfFilePath As String = "C:\Users\giovannitirelli\Desktop\Prova.pdf"
            Dim version As Integer = 1

            pdfFilePath = basePdfFilePath
            While IO.File.Exists(pdfFilePath)
                pdfFilePath = "C:\Users\giovannitirelli\Desktop\Prova_" & version.ToString("D2") & ".pdf"
                version += 1
            End While

            ' Esporta tutti i fogli visibili in un unico PDF
            workbook.ExportAsFixedFormat(Type:=Excel.XlFixedFormatType.xlTypePDF,
                                     Filename:=pdfFilePath,
                                     Quality:=Excel.XlFixedFormatQuality.xlQualityStandard,
                                     IncludeDocProperties:=True,
                                     IgnorePrintAreas:=False,
                                     OpenAfterPublish:=True)

        Catch ex As Exception
            MsgBox("Errore durante l'esportazione in PDF: " & ex.Message)
        End Try

        Return pdfFilePath
    End Function



    Sub righe(workbook As Excel.Workbook, par_nome_foglio As String, par_prima_riga As Integer, par_prima_colonna As Integer, par_lingua As String, par_documento As String, par_numero_documento As String, par_testata_sap As String, par_righe_sap As String, par_valuta As String, par_codice_ktf As Boolean, par_origine_merce As Boolean, par_codice_doganale As Boolean, par_lead_time As Boolean, par_codice_brb As Boolean)

        ' Controlla se il foglio esiste
        Dim sheet As Excel.Worksheet = Nothing
        For Each ws As Excel.Worksheet In workbook.Sheets
            If ws.Name = par_nome_foglio Then
                sheet = ws
                Exit For
            End If
        Next

        If sheet Is Nothing Then
            MsgBox("Il foglio specificato non esiste: " & par_nome_foglio)
            Exit Sub
        End If


        Dim startRow As Integer = par_prima_riga
        Dim startColumn As Integer = par_prima_colonna
        Dim ultima_colonna As Integer = 13
        Dim randomGen As New Random()

        Dim presenza_sconto As Boolean = False

        ' Inserisce una nuova riga prima di iniziare a scrivere
        Dim newRow As Excel.Range = sheet.Rows(startRow + 1)




        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "
declare @lingua as varchar(20)
declare @docnum as integer

set @lingua ='" & par_lingua & "'
set @docnum= " & par_numero_documento & "

SELECT cast(T1.[ItemCode] as varchar) as 'Codice',
coalesce(t6.substitute,'') as 'Codice_bp',
coalesce(CASE WHEN T1.ITEMCODE ='F99999' THEN case when (T1.DSCRIPTION ='' or t1.Dscription is null) then t1.u_note else T1.DSCRIPTION end  ELSE
CASE WHEN @lingua='Italian' then
	
 T1.[Dscription] 

	when @lingua='French' then 
		 t2.U_DESCFR
		
		else  t2.U_DESCING  END END,'')  as 'Descrizione',
case when T1.u_prg_azs_commessa is null then '' else T1.u_prg_azs_commessa end as 'Commessa',
RTRIM(CAST(CAST(T1.[Quantity] AS DECIMAL(10, 2)) AS FLOAT)) AS 'Quantità',
CASE WHEN T1.[DiscPrcnt] >0 THEN 
CASE WHEN T0.DocCur<>'EUR' AND T1.CURRENCY ='EUR' THEN T1.[Pricebefdi]*t0.docrate else
T1.[Pricebefdi] end
ELSE 
CASE WHEN T0.DocCur<>'EUR' AND T1.CURRENCY ='EUR' THEN
T1.[Pricebefdi]*t0.docrate else
T1.[Pricebefdi] end
END  as 'Prezzo unitario',
coalesce(CASE WHEN T0.DocCur<>'EUR' AND T1.CURRENCY ='EUR' THEN t1.price*t0.docrate else t1.price end,0)  as 'Prezzo_dopo_sconto',
case when T1.[DiscPrcnt] is null then 0 else T1.[DiscPrcnt] end  as 'Sconto',
case when t0.DOCCUR='EUR' then t1.linetotal else t1.linetotal*t0.docrate end as'Totale',
case when t1.u_approvvigionamento_articolo is null then '' else t1.u_approvvigionamento_articolo end as 'u_approvvigionamento_articolo',
case when t2.u_codice_ktf is null then '' else t2.u_codice_ktf end as 'Codice_KTF',
coalesce(t2.u_codice_brb,'') as 'Codice_brb',
case when t3.[ISOriCntry] is null then '' else t3.[ISOriCntry] end as'Made_in',
case when T4.Code is null then '' else T4.Code end as'Codice_doganale',
case when t1.freetxt is null then '' else t1.freetxt end as'Note',
case when t2.itemname is null then '' else t2.itemname end as 'Desc_ITA'
,t1.linenum 
,t0.docentry
, COALESCE(T5.LINETEXT,'') AS 'RIGA_TESTO'


FROM " & par_testata_sap & " T0  INNER JOIN " & par_righe_sap & " T1 ON T0.[DocEntry] = T1.[DocEntry] 
left join oitm t2 on t2.itemcode=t1.itemcode
left join itm10 t3 on t3.itemcode=t1.itemcode
left join ODCI T4 on t4.AbsEntry=t3.ISCommCode
left join QUT10 T5 ON T5.DOCENTRY=T0.DOCENTRY AND '" & par_testata_sap & "' ='OQUT' AND T0.OBJTYPE='23' AND T5.AFTLINENUM=T1.visorder-1

left join oscn t6 on t6.itemcode=t1.itemcode and t6.cardcode=t0.cardcode

       
        WHERE T0.[DocNum] =@docnum 
ORDER BY T1.VISORDER"

        cmd_SAP_reader = CMD_SAP.ExecuteReader



        Dim conta_riga As Integer = 0
        Dim conta_colonna As Integer = 0


        verifica_presenza_testi(workbook, par_nome_foglio, -1, startRow + conta_riga, trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).docentry, par_righe_sap, startColumn, ultima_colonna)
        conta_riga = conta_riga + righe_testo
        If righe_testo = 1 Then
            newRow.Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown)
        End If
        righe_testo = 0

        Do While cmd_SAP_reader.Read() = True

            conta_colonna = 0

            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = conta_riga + 1
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            conta_colonna += 1
            If cmd_SAP_reader("codice_BP") <> "" Then
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = cmd_SAP_reader("Codice_bp") & vbCrLf & cmd_SAP_reader("Codice")
            Else
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = cmd_SAP_reader("Codice")
            End If

            If par_codice_brb <> False Then
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value & vbCrLf & cmd_SAP_reader("Codice_BRB")
            End If

            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            'conta_colonna += 1
            'If par_codice_brb = True Then
            '    sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = cmd_SAP_reader("Codice_BRB")
            '    sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
            '    sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            'End If


            conta_colonna += 1

            If cmd_SAP_reader("Descrizione") = "" Then
                Dim descrizione_temporanea As String
                descrizione_temporanea = InputBox("Il codice " & cmd_SAP_reader("Codice") & " " & cmd_SAP_reader("Desc_ita") & vbCrLf & " ha descrizione in lingua nulla, inserire la descrizione che si vuole visualizzare")
                If cmd_SAP_reader("Note") <> "" Then
                    sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = descrizione_temporanea & vbCrLf & cmd_SAP_reader("Note")
                Else
                    sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = descrizione_temporanea
                End If

                aggiorna_traduzione_codice(cmd_SAP_reader("Codice"), descrizione_temporanea, par_lingua)
            Else
                If cmd_SAP_reader("Note") <> "" Then
                    sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = cmd_SAP_reader("Descrizione") & vbCrLf & cmd_SAP_reader("Note")
                Else
                    sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = cmd_SAP_reader("Descrizione")
                End If


            End If

            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter




            conta_colonna += 1
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = cmd_SAP_reader("Commessa")
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            conta_colonna += 1
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = cmd_SAP_reader("Quantità")
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            conta_colonna += 1

            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = par_valuta & " " & FormatNumber(cmd_SAP_reader("Prezzo unitario"), 2, , , TriState.True)
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            conta_colonna += 1


            If cmd_SAP_reader("sconto") > 0 Then
                presenza_sconto = True
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = FormatPercent(cmd_SAP_reader("sconto") / 100)
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End If


            conta_colonna += 1

            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = par_valuta & " " & FormatNumber(cmd_SAP_reader("Totale"), 2, , , TriState.True)
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            conta_colonna += 1
            If par_codice_ktf = True Then
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = cmd_SAP_reader("Codice_KTF")
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End If

            conta_colonna += 1
            If par_origine_merce = True Then
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = cmd_SAP_reader("MAde_IN")
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End If

            conta_colonna += 1
            If par_codice_doganale = True Then
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = cmd_SAP_reader("Codice_doganale")
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End If
            conta_colonna += 1
            If par_lead_time = True Then
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).Value = cmd_SAP_reader("u_approvvigionamento_articolo")
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                sheet.Cells(startRow + (conta_riga), startColumn + conta_colonna).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
            End If

            Dim colonna_da_nascondere As Integer
            Dim colonna_di_appoggio As Integer

            If par_codice_brb = False Then
                colonna_da_nascondere = 4
                colonna_di_appoggio = 5


            End If

            sheet.Rows(startRow + (conta_riga)).EntireRow.AutoFit()

            conta_colonna += 1


            conta_riga += 1

            verifica_presenza_testi(workbook, par_nome_foglio, cmd_SAP_reader("linenum"), startRow + conta_riga, cmd_SAP_reader("docentry"), par_righe_sap, startColumn, ultima_colonna)
            conta_riga = conta_riga + righe_testo
            righe_testo = 0
            newRow.Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown)
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()



        ' Aggiungi bordi a tutte le celle e al perimetro della tabella
        Dim rangeTable As Excel.Range = sheet.Range(sheet.Cells(startRow, startColumn), sheet.Cells(startRow + conta_riga - 1, startColumn + conta_colonna - 1))

        ' Bordi esterni
        With rangeTable.Borders(Excel.XlBordersIndex.xlEdgeLeft)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With rangeTable.Borders(Excel.XlBordersIndex.xlEdgeTop)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With rangeTable.Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With rangeTable.Borders(Excel.XlBordersIndex.xlEdgeRight)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
        End With

        ' Bordi interni
        With rangeTable.Borders(Excel.XlBordersIndex.xlInsideHorizontal)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With rangeTable.Borders(Excel.XlBordersIndex.xlInsideVertical)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        AggiungiRigheSeNecessario(conta_riga, sheet, startRow)
        nascondi_colonne(workbook, par_nome_foglio, startRow, conta_riga, par_codice_brb, presenza_sconto, par_codice_ktf, par_origine_merce, par_codice_doganale, par_lead_time)


    End Sub

    Sub verifica_presenza_testi(workbook As Excel.Workbook, par_nome_foglio As String, par_linenum As Integer, par_riga As Integer, par_DOCENTRY As Integer, par_righe_sap As String, par_prima_colonna As Integer, par_ultima_colonna As Integer)

        righe_testo = 0
        Dim tabella_testo As String = ""

        If par_righe_sap = "QUT1" Then
            tabella_testo = "QUT10"
        ElseIf par_righe_sap = "RDR1" Then
            tabella_testo = "RDR10"
        ElseIf par_righe_sap = "DLN1" Then
            tabella_testo = "DLN10"
        ElseIf par_righe_sap = "INV1" Then
            tabella_testo = "INV10"
        End If


        ' Controlla se il foglio esiste
        Dim sheet As Excel.Worksheet = Nothing
        For Each ws As Excel.Worksheet In workbook.Sheets
            If ws.Name = par_nome_foglio Then
                sheet = ws
                Exit For
            End If
        Next

        If sheet Is Nothing Then
            MsgBox("Il foglio specificato non esiste: " & par_nome_foglio)
            Exit Sub
        End If



        Dim randomGen As New Random()

        Dim presenza_sconto As Boolean = False




        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.DocEntry, T0.LineSeq, T0.AftLineNum, T0.OrderNum, T0.LineType, T0.LineText, T0.ObjType 
FROM " & tabella_testo & " T0 
WHERE T0.[DocEntry] = " & par_DOCENTRY & " and T0.AftLineNum= " & par_linenum & " "

        cmd_SAP_reader = CMD_SAP.ExecuteReader



        Dim conta_riga As Integer = 0
        Dim conta_colonna As Integer = 0
        If cmd_SAP_reader.Read() = True Then

            ' Inserisce una nuova riga prima di iniziare a scrivere
            Dim newRow As Excel.Range = sheet.Rows(par_riga)
            newRow.Insert(Shift:=Excel.XlInsertShiftDirection.xlShiftDown)
            sheet.Cells(par_riga, par_prima_colonna).Value = cmd_SAP_reader("LineText")
            ' Unisci orizzontalmente le celle D (colonna startColumn + 2) ed E (colonna startColumn + 3)
            Dim rangeToMerge As Excel.Range = sheet.Range(sheet.Cells(par_riga, par_prima_colonna), sheet.Cells(par_riga, par_ultima_colonna))
            rangeToMerge.Merge()

            ' Imposta il testo in grassetto
            rangeToMerge.Font.Bold = True
            rangeToMerge.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft ' Centra il contenuto
            rangeToMerge.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter ' Centra verticalmente il contenuto
            righe_testo += 1
        End If

        sheet.Rows(par_riga).EntireRow.AutoFit()
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub
    Sub AggiungiRigheSeNecessario(ByRef conta_riga As Integer, sheet As Excel.Worksheet, riga_iniziale As Integer)
        ' Se conta_riga è inferiore a 10, aggiungi 10 righe vuote
        If conta_riga < 10 Then
            For i As Integer = 1 To 11 - conta_riga
                AggiungiRiga(sheet, conta_riga, riga_iniziale)
            Next
        End If
    End Sub

    Sub AggiungiRiga(sheet As Excel.Worksheet, ByRef conta_riga As Integer, riga_iniziale As Integer)
        ' Aggiungi una riga vuota sotto conta_riga
        sheet.Rows(riga_iniziale + conta_riga + 2).Insert(Excel.XlInsertShiftDirection.xlShiftDown)
        conta_riga += 1 ' Aggiorna conta_riga dopo aver aggiunto la riga
    End Sub


    Sub nascondi_colonne(workbook As Excel.Workbook, par_nome_foglio As String, par_prima_riga As Integer, par_ultima_riga As Integer, par_codice_BRB As Boolean, par_sconto As Boolean, par_cod_KTF As Boolean, par_origine_merce As Boolean, par_cod_doganale As Boolean, par_lead_time As Boolean)
        Try
            ' Controlla se il foglio esiste
            Dim sheet As Excel.Worksheet = Nothing
            For Each ws As Excel.Worksheet In workbook.Sheets
                If ws.Name = par_nome_foglio Then
                    sheet = ws
                    Exit For
                End If
            Next

            If sheet Is Nothing Then
                MsgBox("Il foglio specificato non esiste: " & par_nome_foglio)
                Exit Sub
            End If


            ' Espande la colonna E e sovrascrive i valori della colonna D
            ' For i As Integer = par_prima_riga - 1 To par_prima_riga + par_ultima_riga - 1

            '  Next
            If par_lead_time = False Then
                sheet.Columns(13).delete
            End If
            If par_cod_doganale = False Then
                sheet.Columns(12).delete
            End If
            If par_origine_merce = False Then
                sheet.Columns(11).delete
            End If

            If par_cod_KTF = False Then

                sheet.Columns(10).delete
            End If
            If par_sconto = False Then
                sheet.Columns(8).delete

            End If



            ' Opzionale: AutoFit per assicurarti che i dati siano visibili
            sheet.Columns(4).EntireColumn.AutoFit()

            ' sheet.Rows(3).EntireRow.
            ' sheet.Rows(4).EntireRow.AutoFit()
            'sheet.Rows(5).EntireRow.AutoFit()
            sheet.Columns(7).EntireColumn.AutoFit()
            sheet.Columns(8).EntireColumn.AutoFit()
            sheet.Columns(9).EntireColumn.AutoFit()



        Catch ex As Exception
            MsgBox("Errore: " & ex.Message)
        End Try
    End Sub





    Function EsportaFogliInPDF(workbook As Excel.Workbook, par_nome_foglio As String, foglioAggiuntivo As String, par_percorso_file As String) As String
        Dim fogliDaStampare As New List(Of Excel.Worksheet)
        Dim pdfFilePath As String = ""
        Dim sheetGTCENG As Excel.Worksheet = Nothing
        Dim sheetGTCFR As Excel.Worksheet = Nothing

        Try
            ' Nascondi le schede "GTC ENG" e "GTC FR"
            If foglioAggiuntivo = "GTC ITA" Then
                sheetGTCENG = workbook.Sheets("GTC ENG")
                sheetGTCFR = workbook.Sheets("GTC FR")
            ElseIf foglioAggiuntivo = "GTC ENG" Then
                sheetGTCENG = workbook.Sheets("GTC ITA")
                sheetGTCFR = workbook.Sheets("GTC FR")
            ElseIf foglioAggiuntivo = "GTC FR" Then
                sheetGTCENG = workbook.Sheets("GTC ENG")
                sheetGTCFR = workbook.Sheets("GTC ITA")
            End If

            sheetGTCENG.Visible = Excel.XlSheetVisibility.xlSheetHidden
            sheetGTCFR.Visible = Excel.XlSheetVisibility.xlSheetHidden

            ' Aggiungi i fogli da includere nel PDF (escludendo "traduzioni")
            Dim sheet1 As Excel.Worksheet = workbook.Sheets(par_nome_foglio)
            Dim sheet2 As Excel.Worksheet = workbook.Sheets(foglioAggiuntivo)

            fogliDaStampare.Add(sheet1)
            fogliDaStampare.Add(sheet2)

            ' Determina il nome del file PDF
            Dim basePdfFilePath As String = Homepage.percorso_offerte_vendita & nome_documento_SAP & "_" & docnum & ".pdf"
            Dim version As Integer = 1

            pdfFilePath = basePdfFilePath
            While IO.File.Exists(pdfFilePath)
                pdfFilePath = Homepage.percorso_offerte_vendita & nome_documento_SAP & "_" & docnum & "_" & version.ToString("D2") & ".pdf"
                version += 1
            End While

            ' Esporta i fogli selezionati in PDF
            workbook.ExportAsFixedFormat(Type:=Excel.XlFixedFormatType.xlTypePDF,
                                      Filename:=pdfFilePath,
                                      Quality:=Excel.XlFixedFormatQuality.xlQualityStandard,
                                      IncludeDocProperties:=True,
                                      IgnorePrintAreas:=False,
                                      OpenAfterPublish:=True) ' Apri il PDF dopo l'esportazione

            ' Determina il nome del file Excel
            Dim excelFilePath As String = Homepage.percorso_offerte_vendita & nome_documento_SAP & "_" & docnum & ".xlsx"
            version = 1
            While IO.File.Exists(excelFilePath)
                excelFilePath = Homepage.percorso_offerte_vendita & nome_documento_SAP & "_" & docnum & "_" & version.ToString("D2") & ".xlsx"
                version += 1
            End While

            ' Salva l'Excel nel percorso specificato
            workbook.SaveAs(excelFilePath)

        Catch ex As Exception
            MsgBox("Errore durante l'esportazione in PDF: " & ex.Message)
        Finally
            ' Ripristina la visibilità delle schede "GTC ENG" e "GTC FR"
            If sheetGTCENG IsNot Nothing Then sheetGTCENG.Visible = Excel.XlSheetVisibility.xlSheetVisible
            If sheetGTCFR IsNot Nothing Then sheetGTCFR.Visible = Excel.XlSheetVisibility.xlSheetVisible

            ' Rilascia le risorse
            For Each sheet As Excel.Worksheet In fogliDaStampare
                ReleaseObject(sheet)
            Next
        End Try

        Return pdfFilePath
    End Function




    ' Funzione per rilasciare le risorse COM


    ' Funzione per rilasciare le risorse COM
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.SelectedIndex < 0 Then
            MsgBox("Selezionare un documento")
        Else


            If nome_documento_SAP = "Ordine_acquisto" Then
                Informazioni_documento_acquisto(docnum, documento_SAP, righe_SAP)
                trova_word_base(Lingua, documento_SAP, garanzia, nome_documento_SAP)
                Genera_documento_ACQUISTO()
            ElseIf nome_documento_SAP = "Richiesta_di_offerta" Then
                Informazioni_documento_acquisto(docnum, documento_SAP, righe_SAP)
                trova_word_base(Lingua, documento_SAP, garanzia, nome_documento_SAP)
                Genera_documento_ACQUISTO()

            ElseIf nome_documento_SAP = "Non_conformità" Then

                Informazioni_NC(docnum)
                trova_word_base(Lingua, documento_SAP, garanzia, nome_documento_SAP)
                Genera_documento_NC(docnum)

            ElseIf nome_documento_SAP = "Order" Or nome_documento_SAP = "Off" Then

                prepara_etichette_excel(Homepage.percorso_server & "00-Tirelli 4.0\T4.0vb\Eseguibili\Layout documenti\Documento_base_5.xlsx", "Documento", trova_dettagli_documento(docnum, documento_SAP, righe_SAP).Lingua, nome_documento_SAP, docnum, documento_SAP, righe_SAP)
                ' prepara_etichette_excel("C:\Users\giovannitirelli\Desktop\B.xlsx", "Documento", trova_dettagli_documento(docnum, documento_SAP, righe_SAP).Lingua, nome_documento_SAP, docnum, documento_SAP, righe_SAP)

            Else

                Informazioni_documento(docnum)
                trova_word_base(Lingua, documento_SAP, garanzia, nome_documento_SAP)
                Genera_documento(codice_BRB)

            End If


        End If


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        docnum = TextBox1.Text
    End Sub



    Sub Verifica_Presenza_sconto_righe()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "Select sum(Case When t0.discprcnt>0 Then 1 Else 0 End) As 'Presenza Sconto' 
    FROM " & righe_SAP & " T0  INNER JOIN " & documento_SAP & " T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T1.[DocNum]= " & docnum & ""

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then
            Presenza_sconto_righe = cmd_SAP_reader("Presenza Sconto")



        Else
            MsgBox("La query max righe_offerta non sta funzionando")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Verifico presenza di sconto nelle righe

    Sub Verifica_Presenza_produttore()
        Presenza_produttore = 0
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T1.[ItemCode], T1.[Dscription],T1.FREETXT,T1.U_DISEGNO,T3.FIRMNAME,T2.SUPPCATNUM,T2.BUYUNITMSR,T1.QUANTITY,T1.PRICE,T1.LINETOTAL, T1.SHIPDATE 
FROM " & documento_SAP & " T0  INNER JOIN " & righe_SAP & " T1 ON T0.[DocEntry] = T1.[DocEntry]
LEFT JOIN OITM T2 ON T2.ITEMCODE=T1.ITEMCODE
LEFT JOIN OMRC T3 ON T3.FIRMCODE=T2.FIRMCODE WHERE T0.[DocNum] =" & docnum & " and t3.firmname<>'-' "

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then
            Presenza_produttore = 1



        Else
            Presenza_produttore = 0

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Verifico presenza di sconto nelle righe

    Sub Verifica_Presenza_catalogo_fornitore()
        Presenza_catalogo_fornitore = 0
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T1.[ItemCode], T1.[Dscription],T1.FREETXT,T1.U_DISEGNO,T3.FIRMNAME,T2.SUPPCATNUM,T2.BUYUNITMSR,T1.QUANTITY,T1.PRICE,T1.LINETOTAL, T1.SHIPDATE 
FROM " & documento_SAP & " T0  INNER JOIN " & righe_SAP & " T1 ON T0.[DocEntry] = T1.[DocEntry]
LEFT JOIN OITM T2 ON T2.ITEMCODE=T1.ITEMCODE
LEFT JOIN OMRC T3 ON T3.FIRMCODE=T2.FIRMCODE WHERE T0.[DocNum] =" & docnum & " and T2.SUPPCATNUM<>'' "

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then
            Presenza_catalogo_fornitore = 1



        Else
            Presenza_catalogo_fornitore = 0

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Verifico presenza di sconto nelle righe

    Sub Verifica_Presenza_note()
        Presenza_note = 0
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T1.[ItemCode], T1.[Dscription],T1.FREETXT,T1.U_DISEGNO,T3.FIRMNAME,T2.SUPPCATNUM,T2.BUYUNITMSR,T1.QUANTITY,T1.PRICE,T1.LINETOTAL, T1.SHIPDATE 
FROM " & documento_SAP & " T0  INNER JOIN " & righe_SAP & " T1 ON T0.[DocEntry] = T1.[DocEntry]
LEFT JOIN OITM T2 ON T2.ITEMCODE=T1.ITEMCODE
LEFT JOIN OMRC T3 ON T3.FIRMCODE=T2.FIRMCODE WHERE T0.[DocNum] =" & docnum & " and T1.freetxt<>null "

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then
            Presenza_note = 1



        Else
            Presenza_note = 0

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Verifico presenza di sconto nelle righe

    Public Function Verifica_Presenza_disegno(par_documento_sap As String, par_righe_sap As String, par_docnum As String)
        Dim Presenza_disegno = 0
        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT T1.[ItemCode], T1.[Dscription],T1.FREETXT,T1.U_DISEGNO,T3.FIRMNAME,T2.SUPPCATNUM,T2.BUYUNITMSR,T1.QUANTITY,T1.PRICE,T1.LINETOTAL, T1.SHIPDATE 
FROM " & par_documento_sap & " T0  INNER JOIN " & par_righe_sap & " T1 ON T0.[DocEntry] = T1.[DocEntry]
LEFT JOIN OITM T2 ON T2.ITEMCODE=T1.ITEMCODE
LEFT JOIN OMRC T3 ON T3.FIRMCODE=T2.FIRMCODE 
WHERE T0.[DocNum] =" & par_docnum & " and COALESCE(T1.U_DISEGNO,'')<>''"

        Reader_Tree = Cmd_Tree.ExecuteReader()
        If Reader_Tree.Read() = True Then
            Presenza_disegno = 1



        Else
            Presenza_disegno = 0

        End If
        Reader_Tree.Close()
        Cnn_Tree.Close()
        Return Presenza_disegno
    End Function 'Verifico presenza di sconto nelle righe



    Sub aggiorna_traduzione_codice(par_codice_sap As String, par_descrizione As String, par_lingua As String)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        If par_lingua = "Italian" Then

        ElseIf par_lingua = "French" Then
            CMD_SAP_3.CommandText = "UPDATE T0 SET T0.u_descfr = '" & par_descrizione & "' from oitm t0 where t0.itemcode='" & par_codice_sap & "' "
        Else
            CMD_SAP_3.CommandText = "UPDATE T0 SET T0.u_descing = '" & par_descrizione & "' from oitm t0 where t0.itemcode='" & par_codice_sap & "' "
        End If


        CMD_SAP_3.ExecuteNonQuery()



        Cnn3.Close()


    End Sub




    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            garanzia = "Y"
        Else
            garanzia = "N"
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            made_in = True
        Else
            made_in = False
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If CheckBox3.Checked = True Then
            note = "Y"
        Else
            note = "N"
        End If
    End Sub



    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            approvvigionamento_articolo = True
        Else
            approvvigionamento_articolo = False
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If CheckBox5.Checked = True Then
            codice_KTF = True
        Else
            codice_KTF = False
        End If
    End Sub



    Private Sub Layout_documenti_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LinkLabel1.Text = Homepage.percorso_offerte_vendita
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start(LinkLabel1.Text)
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        If CheckBox6.Checked = True Then
            codice_doganale_riga = True
        Else
            codice_doganale_riga = False
        End If
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        If CheckBox7.Checked = True Then
            codice_BRB = True
        Else
            codice_BRB = False
        End If
    End Sub

    Public Function OttieniDettaglidocumento_Articolo(par_Codice_SAP As String, par_nome_documento As String, par_tabella_intestazione As String, par_tabella_righe As String) As Documento_

        Dim dettagli_ As New Documento_

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = " SELECT t12.docnum, t11.itemcode, t11.linetotal, t12.cardcode, t12.cardname, coalesce(t13.cardname,'') as 'Cliente_F'
FROM
(
select max (t0.docentry) as 'Docentry'
from " & par_tabella_righe & " T0 WHERE T0.ITEMCODE='" & par_Codice_SAP & "'
)
AS T10 INNER JOIN " & par_tabella_righe & " T11 on t11.docentry=t10.docentry and t11.itemcode='" & par_Codice_SAP & "'
inner join " & par_tabella_intestazione & " t12 on t12.docentry=t10.docentry
left join ocrd t13 on t13.cardcode=t12.u_codicebp
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then
            dettagli_.documento = par_nome_documento
            dettagli_.N_DOC = cmd_SAP_reader("docnum")
            dettagli_.prezzo = cmd_SAP_reader("linetotal")
            dettagli_.cliente = cmd_SAP_reader("cardname")
            dettagli_.cliente_Finale = cmd_SAP_reader("Cliente_F")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
        Return dettagli_
    End Function

    Public Class Documento_
        Public documento As String = ""
        Public cliente As String = ""
        Public cliente_Finale As String = ""
        Public N_DOC As Integer = 0
        Public prezzo As Decimal = 0


    End Class


End Class