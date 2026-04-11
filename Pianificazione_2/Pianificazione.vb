Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports Npgsql

Public Class Pianificazione

    'Public Cnn As New SqlConnection

    Public Cnn2 As New SqlConnection
    Public Cnn3 As New SqlConnection
    Public Cnn4 As New SqlConnection
    Public cnn5 As New SqlConnection
    Public cnn6 As New SqlConnection
    ' Public cnn7 As New SqlConnection

    Public Cn As New OleDb.OleDbConnection
    Public reader As OleDb.OleDbDataReader
    Public reader2 As OleDb.OleDbDataReader
    Public Codice As String

    Public Consegna As String
    'Public commessa As String
    Public risorsa As String
    Public nome_risorsa As String
    Public estrazione As String
    Public contatore As Integer = 5
    Public data_inizio As String
    Public data_fine As String
    Public Nome_commessa As String
    Public Ordine_cliente As String
    Public Cliente As String
    Public Fase As String
    Public Cliente_finale As String
    Public commessa_aperta As String
    Public commessa_appoggio As String
    Public ordinatore As Integer = 0
    Public stato As String
    Public KPI As String
    '  Public SAP As String
    Public SAP_4LIFE As String
    ' Public SAP_TIRELLI As String
    Public SAP_PROVA As String

    Public ConfigDictionary As New Dictionary(Of String, String)

    Public ACCESS As String = ""

    Public Excel As Excel.Application
    'Public DATABASE_MU_4_0 = "Data Source=srvtirapp01;Initial Catalog=DbOverOneNew;Persist Security Info=True;User ID=OverOneReader;Password=ReaderOvermach2018!"
    Public DATABASE_MU_4_0 = "Data Source=srvtirapp01;Initial Catalog=DbOverOneNew;Persist Security Info=True;User ID=sa;Password=123B1Admin"


    'lancio odp
    Public docentry As String
    Public docnum As String
    Dim Series As String
    Dim pindicator As String
    Dim Versionnum As String
    Dim JrnlMemo As String
    Dim itemcode As String
    Dim quantità As Integer
    Dim itemname As String
    Dim magazzino As String
    Dim suppcatnum As String
    Dim produzione As String
    Public ItemcodeDB As String
    Public DescrizioneDB As String
    Public QuantitàDB As String
    Public VisorderDB As Integer
    Public maxvisorder As Integer
    Public MagazzinoDB As String
    Public TypeDB As Integer
    Public AttrezzaggioDB As String
    Public TestoDB As String
    Public Resallocdb As String
    Public UomcodeDB As String
    Public UomentryDB As String
    Public BP_code As String
    Public DatrasferireDB As String
    Private RIGA As Integer
    Public sender_mail As String = "report@tirelli.net"
    Public Password_Mail As String = "Ras70773"
    Public ultimo_progressivo_commessa As Integer

    Public blocco As String
    Public stringa_blocco As String

    Public object_type As String
    Public transaction_type As String
    Public num_of_cols_in_key As Integer
    Public list_of_key_cols_tab_del As String
    Public list_of_cols_val_tab_del As String
    Private filtro_divisione As String
    Public commessa As String



    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Commesse_aperte()
        Me.BackColor = Homepage.colore_sfondo



    End Sub


    Sub Commesse_aperte()
        DataGridView_Pianificazione.Rows.Clear()
        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1

        CMD_SAP_1.CommandText = " select  t10.commessa  , t10.Docentry,t10.KPI,


            coalesce(t14.itemname ,t10.[Nome commessa]) As 'Nome commessa',
            coalesce(t11.docnum,t10.oc) as 'ODC',
            substring(coalesce(t11.cardname,t10.Cliente),1,25) As 'Cliente',
coalesce(t13.cardname,'') as 'Cliente Finale',
            coalesce(T12.[ShipDate],t10.consegna) as 'Consegna'
, coalesce(t12.ocrcode,'') as 'ocrcode'

from
(
SELECT t0.commessa 'Commessa' ,t0.descrizione as 'Nome commessa', t0.consegna as 'Consegna', t0.OC as 'OC'
, max(t1.docentry) as 'Docentry'
, min(t1.linenum) as 'Linenum'

,coalesce(t0.cliente,'') as 'Cliente',
coalesce(t0.kpi,'No') as 'KPI'

from [Tirelli_40].[dbo].pianificazione_commessa t0 
left join rdr1 t1 on t1.itemcode=t0.commessa
where t0.stato<>'C'
group by t0.commessa  ,t0.descrizione, t0.consegna , t0.OC , 
t0.cliente,
t0.kpi
)
as t10 left join ordr t11 on t11.docentry=t10.docentry
left join rdr1 t12 on t12.itemcode=t10.commessa and t10.linenum=t12.linenum
left join ocrd t13 on t13.CardCode=t11.U_CodiceBP
left join oitm t14 on t14.itemcode=t12.itemcode
where 0=0 " & filtro_divisione & "
order by t10.commessa "


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()

            DataGridView_Pianificazione.Rows.Add(cmd_SAP_reader_1("Commessa"), cmd_SAP_reader_1("Nome commessa"), cmd_SAP_reader_1("Cliente"), cmd_SAP_reader_1("Cliente finale"), cmd_SAP_reader_1("ocrcode"), cmd_SAP_reader_1("consegna"), cmd_SAP_reader_1("KPI"))




        Loop
        Cnn1.Close()

    End Sub






    Sub aggiorna_commesse_aperte()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T1.[ItemCode] as 'Commessa', T2.[ItemName] as 'Nome commessa', T0.[DocNum] as 'Ordine cliente', T1.[ShipDate] as 'Consegna' 
FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] WHERE substring( T2.[ItemCode] ,1,1)='M' and  T1.[OpenQty] >0 and t0.docstatus='O'
order by T1.[ItemCode]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()


            Nome_commessa = cmd_SAP_reader("Nome commessa")
            Ordine_cliente = cmd_SAP_reader("Ordine cliente")
            Consegna = cmd_SAP_reader("Consegna")
            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_1 As New SqlCommand
            Dim cmd_SAP_reader_1 As SqlDataReader
            CMD_SAP_1.Connection = Cnn1

            CMD_SAP_1.CommandText = "SELECT t0.ID from [Tirelli_40].[dbo].Pianificazione_commessa t0 where t0.commessa='" & cmd_SAP_reader("Commessa") & "' "

            cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader

            If cmd_SAP_reader_1.Read() = True Then
                Cnn2.ConnectionString = Homepage.sap_tirelli
                Cnn2.Open()

                CMD_SAP.Connection = Cnn2
                CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].pianificazione_commessa SET pianificazione_commessa.[Stato]='O', pianificazione_commessa.[Descrizione]='" & Nome_commessa & "', pianificazione_commessa.[OC]='" & Ordine_cliente & "', pianificazione_commessa.consegna=CONVERT(DATETIME, '" & Consegna & "',103) where pianificazione_commessa.commessa= '" & cmd_SAP_reader("Commessa") & "'"


                CMD_SAP.ExecuteNonQuery()
                Cnn2.Close()
            Else
                Cnn2.ConnectionString = Homepage.sap_tirelli
                Cnn2.Open()
                CMD_SAP.Connection = Cnn2
                CMD_SAP.CommandText = "insert into [Tirelli_40].[dbo].pianificazione_commessa ([commessa],[Descrizione],[Stato],[OC], consegna,GESTIONE) VALUES ('" & cmd_SAP_reader("Commessa") & "',substring('" & Nome_commessa & "',1,40),'O','" & Ordine_cliente & "',CONVERT(DATETIME, '" & Consegna & "',103),'MACCHINE')"
                CMD_SAP.ExecuteNonQuery()
                Cnn2.Close()

            End If



            Cnn1.Close()
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()


        Commesse_aperte()

    End Sub

    Private Sub Button_aggiorna_commesse_aperte_Click(sender As Object, e As EventArgs) Handles Button_aggiorna_commesse_aperte.Click
        aggiorna_commesse_aperte()
    End Sub

    Sub aggiorna_CDS()
        Dim Cnn As New SqlConnection
        Dim Cnn1 As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT 'ORDINE CLIENTE' AS 'Documento', T0.[DocNum] , '' as 'Offerta' ,  t2.callid as 'CDS',T2.SUBJECT, t4.itemcode as 'Matricola', t4.itemname as 'Matricola ',t0.cardcode, T0.[CardName], T0.U_CODICEBP,T0.U_CLIENTEFINALE, T0.[DocDate], T0.[DocDueDate] AS 'Data consegna', T0.u_PRIMADATACONSEGNA, (T0.DocTotal-T0.PaidSys+T0.VatPaid-T0.VatSum) AS 'Totale aperto', T0.u_PRG_AZS_STATOCOMM, T0.U_uffcompetenza, T0.[U_Commento], T0.[U_CausCons] as 'Vendita/Garanzia', T0.[U_ELABORATORE],T0.[U_DatainizioProg] AS 'Inizio UT', T0.[U_DataprevfineUT] AS 'Fine UT', t0.U_inizioass, t0.u_fineass, T0.[U_Datacollaudo] AS 'Collaudo', T0.[U_DataFatcliente] AS 'Data FAT', T0.[U_Datalayout] As 'Approv Lay-out' 

FROM ORDR T0, OCLG T1 
left join oscl t2 on t2.callid=T1.[parentId]
LEFT JOIN OQUT T3 ON T3.DOCNUM='0'
left join oitm t4 on t4.itemcode=t2.itemcode
WHERE T0.DocStatus = N'o' AND T0.[DocNum] = T1.[DocNum] AND  T1.[DocType] =17 AND T1.[parentId] >1"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()


            Nome_commessa = cmd_SAP_reader("SUBJECT")
            Ordine_cliente = cmd_SAP_reader("Documento")
            Consegna = cmd_SAP_reader("Data consegna")

            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_1 As New SqlCommand
            Dim cmd_SAP_reader_1 As SqlDataReader
            CMD_SAP_1.Connection = Cnn1

            CMD_SAP_1.CommandText = "SELECT t0.ID from [Tirelli_40].[dbo].Pianificazione_commessa t0 where t0.commessa='" & "CDS" & cmd_SAP_reader("CDS") & "' AND T0.OC='" & Ordine_cliente & "' "

            cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader

            If cmd_SAP_reader_1.Read() = True Then
                Cnn2.ConnectionString = Homepage.sap_tirelli
                Cnn2.Open()

                CMD_SAP.Connection = Cnn2
                CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].pianificazione_commessa Set pianificazione_commessa.[Stato]='O', pianificazione_commessa.[Descrizione]='" & Nome_commessa & "', pianificazione_commessa.[OC]='" & Ordine_cliente & "', pianificazione_commessa.consegna=CONVERT(DATETIME, '" & Consegna & "',103) where pianificazione_commessa.commessa= '" & "CDS" & cmd_SAP_reader("CDS") & "'"


                CMD_SAP.ExecuteNonQuery()
                Cnn2.Close()
            Else
                Cnn2.ConnectionString = Homepage.sap_tirelli
                Cnn2.Open()
                CMD_SAP.Connection = Cnn2
                CMD_SAP.CommandText = "insert into [Tirelli_40].[dbo].pianificazione_commessa ([commessa],[Descrizione],[Stato],[OC], consegna,GESTIONE) VALUES ('" & "CDS" & cmd_SAP_reader("CDS") & "','" & Nome_commessa & "','O','" & Ordine_cliente & "',CONVERT(DATETIME, '" & Consegna & "',103),'CDS')"
                CMD_SAP.ExecuteNonQuery()
                Cnn2.Close()

            End If



            Cnn1.Close()
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()




    End Sub



    Sub Chiudi_commessa_M()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].Pianificazione_commessa
        SET Pianificazione_commessa.stato='C'
        WHERE Pianificazione_commessa.commessa='" & commessa & "'"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

        Commesse_aperte()

    End Sub


    Sub elimina_commessa_M()
        Dim Cnn As New SqlConnection
        Cnn.Open()
        Cnn.ConnectionString = Homepage.sap_tirelli
        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "delete [Tirelli_40].[dbo].Pianificazione_commessa
        
        WHERE Pianificazione_commessa.commessa='" & commessa & "'"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

        Commesse_aperte()

    End Sub

    Sub Pulizia_output()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "DELETE [TIRELLI_40].[DBO].PIANIFICAZIONE_OUTPUT"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub

    Sub Pianificazione_output()
        Pulizia_output()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1

        CMD_SAP_1.CommandText = "select t20.commessa , coalesce(t23.itemname,t20.[nome commessa]) as 'Nome commessa', coalesce(t22.shipdate,t20.Consegna) as 'Consegna',
coalesce(t24.cardname,coalesce(t21.cardname,t20.Cliente)) as 'cliente', t20.[Min Data_i], t20.[Max Data_f] , coalesce(t22.ocrcode,'') as 'Div'
from
(
select 
t10.commessa , t10.[nome commessa], t10.Consegna,
t10.Cliente, coalesce(t10.[Min Data_i],'') as 'Min data_i', coalesce(t10.[Max Data_f],'') as 'MAx data_f' , max(t11.docentry) as 'Docentry', min(t11.linenum) as 'Linenum'
from
(
select t0.commessa as 'Commessa', t0.descrizione as 'Nome commessa', t0.consegna as 'Consegna',
coalesce(t0.cliente,'') as 'Cliente'
, min(t1.data_i ) as 'Min Data_i', max(t1.data_f) as 'Max Data_f' 
from [Tirelli_40].[dbo].pianificazione_commessa t0 left join [TIRELLI_40].[DBO].pianificazione t1 on t0.commessa=t1.commessa
where t0.stato<>'C' 
group by t0.commessa , t0.descrizione, t0.consegna , t0.oc, t0.cliente

)
as t10 left join rdr1 t11 on t10.commessa=t11.itemcode
group by t10.commessa , t10.[nome commessa], t10.Consegna,
t10.Cliente, t10.[Min Data_i], t10.[Max Data_f]
)
as t20
left join ordr t21 on t21.docentry=t20.docentry
left join rdr1 t22 on t22.docentry=t20.docentry and t22.linenum=t20.linenum
left join oitm t23 on t23.itemcode=t20.commessa
left join ocrd t24 on t24.cardcode=t21.u_codicebp
order by t20.consegna"
        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()

            inserisci_commessa_pianificazione_output(cmd_SAP_reader_1("commessa"), cmd_SAP_reader_1("Nome commessa"), cmd_SAP_reader_1("cliente"), cmd_SAP_reader_1("Min Data_i"), cmd_SAP_reader_1("Max Data_f"), cmd_SAP_reader_1("consegna"), cmd_SAP_reader_1("Div"))

            inserisci_risorse(cmd_SAP_reader_1("commessa"), cmd_SAP_reader_1("Nome commessa"), cmd_SAP_reader_1("cliente"), cmd_SAP_reader_1("consegna"), cmd_SAP_reader_1("Div"))

        Loop
        cmd_SAP_reader_1.Close()
        Cnn1.Close()

    End Sub

    Sub inserisci_risorse(par_commessa As String, par_nome_commessa As String, par_cliente As String, par_consegna As String, PAR_DIV As String)

        Cnn2.ConnectionString = Homepage.sap_tirelli
        Cnn2.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        CMD_SAP_2.Connection = Cnn2


        CMD_SAP_2.CommandText = "select t0.id as 'ID', t0.risorsa as 'Risorsa', t0.data_i as 'Data_i', t0.data_f as 'Data_F', t1.resname as 'Nome_risorsa', case when t0.dipendente is null then '' else t0.dipendente end as 'Dipendente', case when t0.attivita is null then '' else t0.attivita end as 'Attivita', case when t0.unita is null then 0 else t0.unita end as 'Unita' 
From [TIRELLI_40].[DBO].pianificazione t0 inner Join orsc t1 on t0.risorsa=t1.visrescode
Where t0.commessa ='" & par_commessa & "'
order by  t0.[risorsa],t0.[Data_I], t0.id"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()


            inserisci_risorse_pianificazione_output(par_commessa, par_nome_commessa, par_cliente, cmd_SAP_reader_2("risorsa"), cmd_SAP_reader_2("Nome_risorsa"), cmd_SAP_reader_2("Dipendente"), cmd_SAP_reader_2("attivita"), cmd_SAP_reader_2("Data_I"), cmd_SAP_reader_2("Data_F"), par_consegna, cmd_SAP_reader_2("unita"), PAR_DIV)
        Loop
        cmd_SAP_reader_2.Close()
        Cnn2.Close()
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button_aggiorna_excel.Click
        commessa_appoggio = commessa
        Pianificazione_output()
        commessa = commessa_appoggio
        Beep()
        MsgBox("Excel aggiornato")

    End Sub



    Sub inserisci_commessa_pianificazione_output(par_commessa As String, par_nome_commessa As String, par_cliente As String, par_data_inizio As String, par_data_fine As String, par_consegna As String, par_DIV As String)

        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()
        Dim CMD_SAP_3 As New SqlCommand
        CMD_SAP_3.Connection = Cnn3
        CMD_SAP_3.CommandText = "INSERT INTO [TIRELLI_40].[DBO].Pianificazione_output (Pianificazione_output.id,Pianificazione_output.[Livello],Pianificazione_output.[commessa],Pianificazione_output.[descrizione],Pianificazione_output.Cliente,Pianificazione_output.DIV, Pianificazione_output.[Data_I], Pianificazione_output.[Data_F], Pianificazione_output.attivita, Pianificazione_output.risorsa, Pianificazione_output.risorsa_desc, Pianificazione_output.dipendente, pianificazione_output.consegna) 
                                                                        Values (" & ordinatore & ",'0','" & par_commessa & "',substring('" & par_nome_commessa & "',1,70),'" & par_cliente & "','" & par_DIV & "',CONVERT(DATETIME, '" & par_data_inizio & "',103), CONVERT(DATETIME, '" & par_data_fine & "',103), '','','','', CONVERT(DATETIME, '" & par_consegna & "',103)) "
        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()
        ordinatore = ordinatore + 1



    End Sub

    Sub inserisci_risorse_pianificazione_output(par_commessa As String, par_nome_commessa As String, par_cliente As String, par_risorsa As String, par_nome_risorsa As String, par_dipendente As String, par_attivita As String, par_data_inizio As String, par_data_fine As String, par_consegna As String, par_unita As Integer, PAR_DIV As String)
        risorsa = Mid(risorsa, 1, 6)

        Cnn4.ConnectionString = Homepage.sap_tirelli
        Cnn4.Open()
        Dim CMD_SAP_4 As New SqlCommand
        CMD_SAP_4.Connection = Cnn4
        CMD_SAP_4.CommandText = "INSERT INTO [TIRELLI_40].[DBO].Pianificazione_output (Pianificazione_output.id, Pianificazione_output.[Livello],Pianificazione_output.[commessa],Pianificazione_output.[descrizione],Pianificazione_output.Cliente,Pianificazione_output.DIV,Pianificazione_output.[Risorsa],Pianificazione_output.[Risorsa_desc],Pianificazione_output.[Dipendente],Pianificazione_output.[Attivita],Pianificazione_output.[Data_I],Pianificazione_output.[Data_F], Pianificazione_output.consegna,PIANIFICAZIONE_OUTPUT.unita) 
                                                                        Values(" & ordinatore & ",'1','" & par_commessa & "','" & par_nome_commessa & "','" & par_cliente & "','" & PAR_DIV & "','" & par_risorsa & "','" & par_nome_risorsa & "', '" & par_dipendente & "','" & par_attivita & "',CONVERT(DATETIME, '" & par_data_inizio & "',103),CONVERT(DATETIME, '" & par_data_fine & "',103),CONVERT(DATETIME, '" & par_consegna & "',103),'" & par_unita & "') "
        CMD_SAP_4.ExecuteNonQuery()
        Cnn4.Close()
        ordinatore = ordinatore + 1


    End Sub




    Private Sub Chiudi_commessa_Click(sender As Object, e As EventArgs)
        Chiudi_commessa_M()
    End Sub


    Private Sub Button_apri_commessa_Click(sender As Object, e As EventArgs) Handles Button_apri_commessa.Click
        Form109.Show()
    End Sub







    Sub lancio_odp()

        ' Try

        Dim contatore_excel As Integer = 2
        Dim lunghezza_excel As Integer = 200
        Dim N_tipologie As Integer = 3
        Dim contatore_colonne As Integer = 2
        risorsa = "P00001"
        Dim testo As String = ""
        'Excel
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(Homepage.percorso_acquisti & "\Lancio_ODP.xlsx")
        Excel.Visible = True



        Do While contatore_excel <= lunghezza_excel


            If Excel.Sheets("Lancio_ODP").Cells(contatore_excel, 1).value <> Nothing Then
                itemcode = Excel.Sheets("Lancio_ODP").Cells(contatore_excel, 1).value
                produzione = Excel.Sheets("Lancio_ODP").Cells(contatore_excel, 2).value
                Fase = Excel.Sheets("Lancio_ODP").Cells(contatore_excel, 3).value
                data_inizio = Excel.Sheets("Lancio_ODP").Cells(contatore_excel, 4).value
                data_fine = Excel.Sheets("Lancio_ODP").Cells(contatore_excel, 5).value
                commessa = Excel.Sheets("Lancio_ODP").Cells(contatore_excel, 6).value
                Cliente = Excel.Sheets("Lancio_ODP").Cells(contatore_excel, 7).value
                quantità = Excel.Sheets("Lancio_ODP").Cells(contatore_excel, 8).value
                Ordine_cliente = Excel.Sheets("Lancio_ODP").Cells(contatore_excel, 9).value
                BP_code = Excel.Sheets("Lancio_ODP").Cells(contatore_excel, 10).value

                Dim Cnn As New SqlConnection
                Cnn.ConnectionString = Homepage.sap_tirelli
                Cnn.Open()

                Dim CMD_SAP_7 As New SqlCommand
                Dim cmd_SAP_reader_7 As SqlDataReader
                CMD_SAP_7.Connection = Cnn

                CMD_SAP_7.CommandText = "SELECT case when t2.validfor is null then '' else t2.validfor end as 'validita_padre', t2.u_gestione_magazzino, case when T1.VALIDFOR is null then '' else t1.validfor end AS 'Valido', t1.itemcode as 'Codice' FROM ITT1 T0 INNER JOIN OITM T1 ON T0.CODE=T1.ITEMCODE inner join oitm t2 on t2.itemcode=t0.father WHERE T0.[Father]= '" & itemcode & "' AND (T1.VALIDFOR='N' or t2.validfor='N' or t2.u_gestione_magazzino='ESAURIMENTO')"

                cmd_SAP_reader_7 = CMD_SAP_7.ExecuteReader

                If cmd_SAP_reader_7.Read() = True Then

                    If cmd_SAP_reader_7("Valido") = "N" Then
                        MsgBox("La distinta " & itemcode & " contiene l'articolo " & cmd_SAP_reader_7("Codice") & " inattivo")
                    End If

                    If cmd_SAP_reader_7("validita_padre") = "N" Then
                        MsgBox("Il codice padre " & itemcode & " risulta inattivo")
                    End If

                    If cmd_SAP_reader_7("u_gestione_magazzino") = "N" Then
                        MsgBox("Il codice padre " & itemcode & " risulta in esaurimento, non è più possibile ordinarlo (guardare gestione magazzino)")
                    End If



                Else
                    blocco = ""
                    stringa_blocco = ""

                    If produzione = "ASSEMBL" Then
                        check_materia_prima_in_ordine_assemblaggio()
                    End If
                    check_non_duplicazione_impegni_con_anticipi()

                    check_che_non_ci_siano_codici_doppi_nelle_righe()

                    If blocco = "Y" Then
                        MsgBox(stringa_blocco)
                    Else
                        Cnn2.ConnectionString = Homepage.sap_tirelli
                        Cnn2.Open()

                        Dim CMD_SAP_2 As New SqlCommand
                        Dim cmd_SAP_reader_2 As SqlDataReader


                        CMD_SAP_2.Connection = Cnn2
                        CMD_SAP_2.CommandText = "SELECT T0.[ItemName] as 'Itemname', CASE WHEN T0.DFLTWH IS NULL THEN '01' ELSE T0.DFLTWH END as 'Magazzino', CASE WHEN t0.suppcatnum is null then'' else t0.suppcatnum end as 'suppcatnum' 
                        From OITM T0 inner join oitt t1 on t0.itemcode=t1.code where t0.itemcode='" & itemcode & "'"

                        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


                        If cmd_SAP_reader_2.Read() = True Then
                            itemname = cmd_SAP_reader_2("Itemname")
                            magazzino = cmd_SAP_reader_2("Magazzino")
                            suppcatnum = cmd_SAP_reader_2("suppcatnum")
                            cmd_SAP_reader_2.Close()
                            If produzione = "ASSEMBL" Then
                                magazzino = "02"
                            ElseIf produzione = "INT" Or produzione = "INT_SALD" Then
                                magazzino = "CAP2"
                            End If
                            max_docentry_docnum()
                            If produzione = "ASSEMBL" Then
                                progressivo_commessa(commessa)
                            Else
                                ultimo_progressivo_commessa = 0
                            End If
                            Dim Cnn1 As New SqlConnection
                            Cnn1.ConnectionString = Homepage.sap_tirelli
                            Cnn1.Open()

                            Dim CMD_SAP_1 As New SqlCommand

                            CMD_SAP_1.Connection = Cnn1
                            CMD_SAP_1.CommandText = "INSERT INTO OWOR (OWOR.DOCENTRY,OWOR.DOCNUM,OWOR.ITEMCODE,OWOR.PRODNAME, owor.plannedqty, OWOR.[U_PRG_AZS_Commessa], owor.U_fase, OWOR.U_UTILIZZ,OWOR.STARTDATE,OWOR.DUEDATE,owor.cmpltqty,owor.rjctqty,owor.postdate,owor.warehouse,owor.uom,owor.JrnlMemo,owor.pindicator,owor.uomentry,owor.updalloc,owor.versionnum, owor.U_produzione,owor.series, owor.usersign, OWOR.originnum, owor.cardcode,owor.U_Progressivo_commessa) 
                    values (" & docentry & "+1,'" & docnum & "'+1, '" & itemcode & "','" & itemname & "','" & quantità & "','" & commessa & "','" & Fase & "','" & Cliente & "',CONVERT(DATETIME, '" & data_inizio & "',103),CONVERT(DATETIME, '" & data_fine & "',103),0,0,getdate(),'" & magazzino & "','" & suppcatnum & "','" & JrnlMemo & "','" & pindicator & "','-1','C','" & Versionnum & "','" & produzione & "','" & Series & "',50, '" & Ordine_cliente & "', '" & BP_code & "'," & ultimo_progressivo_commessa & " )"
                            CMD_SAP_1.ExecuteNonQuery()

                            Cnn1.Close()
                            IMPORTA_DB()
                        Else
                            MsgBox(" manca distinta base per " & itemcode & "")
                            cmd_SAP_reader_2.Close()

                        End If
                        Cnn2.Close()

                    End If





                End If
                Cnn.Close()
            End If

            contatore_excel = contatore_excel + 1

        Loop
        max_docentry_docnum()
        aggiusta_ORDINATO()
        aggiusta_CONFERMATO()
        aggiusta_righe_odp_confermato_tot_ordinato_tot()
        MsgBox("Importazione avvenuta con successo")
        '  Catch ex As Exception

        '  max_docentry_docnum()
        '  End Try


    End Sub

    Sub max_docentry_docnum()
        INFO_odp_precedente()
        cnn6.ConnectionString = Homepage.sap_tirelli
        cnn6.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn6

        CMD_SAP.CommandText = "select max(t0.docentry) as 'Docentry',max(t0.docnum) as 'Docnum' from owor t0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            docentry = cmd_SAP_reader("docentry")
            docnum = cmd_SAP_reader("docnum")

            cmd_SAP_reader.Close()
        End If
        cnn6.Close()
        AGGIUSTA_NUMERATORE()

    End Sub

    Sub check_materia_prima_in_ordine_assemblaggio()


        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn3

        CMD_SAP.CommandText = "SELECT t2.itemcode,t2.itemname, t2.ItmsGrpCod 
From oitt t0 inner join itt1 t1 on t1.father=t0.code
inner join oitm t2 on t2.itemcode=t1.code
where t0.code='" & itemcode & "' and t2.ItmsGrpCod=121 and substring(t2.itemcode,1,1)='C'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then
            blocco = "Y"
            stringa_blocco = " Nella distinta " & itemcode & " è presente il codice " & cmd_SAP_reader("itemcode") & " materia prima. non può essere presente in ODP Assemblaggio"


            cmd_SAP_reader.Close()
        End If
        Cnn3.Close()


    End Sub

    Sub check_non_duplicazione_impegni_con_anticipi()


        cnn6.ConnectionString = Homepage.sap_tirelli
        cnn6.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn6

        CMD_SAP.CommandText = "select t0.father, t0.Code, a.docnum
from itt1 t0 inner join 

(select t1.itemcode, t0.docnum from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry where (t0.status='P' or t0.status='R') and t0.U_PRG_AZS_Commessa='" & commessa & "' and t0.prodname Like 'ANTICIP%%') A on A.itemcode=t0.code

where t0.father='" & itemcode & "' and t0.type=4"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then
            blocco = "Y"
            stringa_blocco = " La distinta " & cmd_SAP_reader("father") & " contiene il codice " & cmd_SAP_reader("Code") & " già contenuto nell ordine di anticipo " & cmd_SAP_reader("docnum") & " "


            cmd_SAP_reader.Close()
        End If
        cnn6.Close()


    End Sub

    Sub check_che_non_ci_siano_codici_doppi_nelle_righe()


        cnn6.ConnectionString = Homepage.sap_tirelli
        cnn6.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn6

        CMD_SAP.CommandText = "select t10.code,t10.n
from
(
select t0.code, count(t0.code) as 'N'
from
itt1 t0 where t0.father='" & itemcode & "' and t0.type=4
group by t0.code
)
as t10
where t10.n>1
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then
            blocco = "Y"
            stringa_blocco = " La distinta " & itemcode & " contiene il codice " & cmd_SAP_reader("Code") & " in due differenti righe, unirle in una unica e poi lanciare l ODP nuovamente "


            cmd_SAP_reader.Close()
        End If
        cnn6.Close()


    End Sub

    Sub progressivo_commessa(par_COMMESSA)


        cnn6.ConnectionString = Homepage.sap_tirelli
        cnn6.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn6

        CMD_SAP.CommandText = "SELECT MAX(CASE WHEN T0.U_Progressivo_commessa IS NULL THEN 0 ELSE T0.U_PROGRESSIVO_COMMESSA END) AS 'Progressivo_commessa'
FROM OWOR T0
WHERE T0.U_PRG_AZS_Commessa='" & par_COMMESSA & "'
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() = True Then

            If Not cmd_SAP_reader("Progressivo_commessa") Is System.DBNull.Value Then
                ultimo_progressivo_commessa = cmd_SAP_reader("Progressivo_commessa") + 1
            Else
                ultimo_progressivo_commessa = 1
            End If



            cmd_SAP_reader.Close()
        End If
        cnn6.Close()


    End Sub

    Sub tentativo_check_stored_procedure()



        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "EXEC SBO_SP_TransactionNotification 
@object_type = '" & object_type & "'
, @Transaction_type = '" & transaction_type & "'
, @num_of_cols_in_key = " & num_of_cols_in_key & "
,@list_of_key_cols_tab_del ='" & list_of_key_cols_tab_del & "'
,@list_of_cols_val_tab_del = '" & list_of_cols_val_tab_del & "
   
GO
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then
            MsgBox(cmd_SAP_reader(1))


            cmd_SAP_reader.Close()
        End If
        Cnn.Close()


    End Sub

    Sub AGGIUSTA_NUMERATORE()
        INFO_odp_precedente()
        cnn5.ConnectionString = Homepage.sap_tirelli
        cnn5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = cnn5
        CMD_SAP_5.CommandText = "UPDATE ONNM SET AUTOKEY ='" & docentry & "'+1 WHERE OBJECTCODE='202'"
        CMD_SAP_5.ExecuteNonQuery()


        CMD_SAP_5.CommandText = "Update NNM1 SET NEXTNUMBER='" & docnum & "'+1 WHERE OBJECTCODE='202' And SERIES=" & Series & ""
        CMD_SAP_5.ExecuteNonQuery()

        cnn5.Close()
    End Sub

    Sub INFO_odp_precedente()

        cnn6.ConnectionString = Homepage.sap_tirelli
        cnn6.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn6

        CMD_SAP.CommandText = "select t0.series as 'series',t0.pindicator as 'pindicator',case when t0.versionnum is null then '' else t0.versionnum end as 'Versionnum', case when t0.JrnlMemo is null then '' else t0.JrnlMemo end as 'JrnlMemo' from owor t0 where t0.docentry=" & docentry & "-1"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            Series = cmd_SAP_reader("series")
            pindicator = cmd_SAP_reader("pindicator")
            Versionnum = cmd_SAP_reader("versionnum")
            JrnlMemo = cmd_SAP_reader("JrnlMemo")
            cmd_SAP_reader.Close()
        End If
        cnn6.Close()
    End Sub

    Sub IMPORTA_DB()
        maxvisorder = 0
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand
        Dim cmd_SAP_reader_3 As SqlDataReader
        CMD_SAP_3.Connection = Cnn3

        CMD_SAP_3.CommandText = "SELECT CASE when t0.code is null then '' else T0.[Code] end as 'ItemcodeDB', CASE WHEN t1.itemname IS NULL THEN'' ELSE T1.ITEMNAME END as 'DescrizioneDB', T0.[Quantity]/T2.[Qauntity] as 'QuantitàDB', T0.[VisOrder] as 'VisorderDB', case when t0.warehouse is null then '' else T0.[Warehouse] end as 'MagazzinoDB', T0.Type as 'TypeDB', case when T0.AddQuantit is null then 0 else T0.AddQuantit end as 'AttrezzaggioDB', case when T0.LineText is null then '' else t0.linetext end as 'TestoDB', case when t0.type=4 then 'null' else 'F' end as'Resallocdb', case when t0.type=4 then '-1' else '' end as'UOMENTRYDB', case when t0.type=4 then 'Manuale' else '' end as'UomcodeDB', case when (substring(t0.code,1,1)='0' or substring(t0.code,1,1)='C' or substring(t0.code,1,1)='D' or substring(t0.code,1,1)='F') then t0.quantity else 0 end as 'DatrasferireDB' 
from itt1 t0 left join oitm t1 on t0.code=t1.itemcode
left join oitt t2 on t0.father=t2.code
 WHERE T0.[Father] ='" & itemcode & "'"

        cmd_SAP_reader_3 = CMD_SAP_3.ExecuteReader

        Do While cmd_SAP_reader_3.Read()

            ItemcodeDB = cmd_SAP_reader_3("ItemcodeDB")
            DescrizioneDB = cmd_SAP_reader_3("DescrizioneDB")
            QuantitàDB = cmd_SAP_reader_3("QuantitàDB")
            VisorderDB = cmd_SAP_reader_3("VisorderDB")
            If VisorderDB > maxvisorder Then
                maxvisorder = VisorderDB
            End If
            MagazzinoDB = cmd_SAP_reader_3("MagazzinoDB")
            TypeDB = cmd_SAP_reader_3("TypeDB")
            AttrezzaggioDB = cmd_SAP_reader_3("AttrezzaggioDB")
            TestoDB = cmd_SAP_reader_3("TestoDB")
            If cmd_SAP_reader_3("resallocdb") = "null" Then
                Resallocdb = Nothing
            Else
                Resallocdb = cmd_SAP_reader_3("resallocdb")
            End If
            UomentryDB = cmd_SAP_reader_3("UomentryDB")
            UomcodeDB = cmd_SAP_reader_3("UomcodeDB")
            DatrasferireDB = cmd_SAP_reader_3("DatrasferireDB")

            QuantitàDB = Replace(QuantitàDB, ",", ".")
            AttrezzaggioDB = Replace(AttrezzaggioDB, ",", ".")
            DatrasferireDB = Replace(DatrasferireDB, ",", ".")
            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_1 As New SqlCommand

            CMD_SAP_1.Connection = Cnn1

            If TypeDB = 4 Then

                CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType,  wor1.[StartDate], wor1.[EndDate], wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf])
                                            VALUES(" & docentry & "+1, " & VisorderDB & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & VisorderDB & ", " & QuantitàDB & " ,'" & AttrezzaggioDB & " ', " & AttrezzaggioDB & " + (" & QuantitàDB & " * " & quantità & "),'" & MagazzinoDB & "'," & TypeDB & ",'B', CONVERT(DATETIME, '" & data_inizio & "',103), CONVERT(DATETIME, '" & data_fine & "',103),0,0,0," & QuantitàDB & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & DatrasferireDB & " * " & quantità & ")"

            ElseIf TypeDB = -18 Then

                CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType, wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf])
                                            VALUES(" & docentry & "+1, " & VisorderDB & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & VisorderDB & ", " & QuantitàDB & " ,'" & AttrezzaggioDB & " ', " & AttrezzaggioDB & " + (" & QuantitàDB & " * " & quantità & "),'" & MagazzinoDB & "'," & TypeDB & ",'B',0,0,0," & QuantitàDB & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & DatrasferireDB & " * " & quantità & ")"

            Else
                CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType,  wor1.[StartDate], wor1.[EndDate], wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf],WOR1.RESALLOC)
                                            VALUES(" & docentry & "+1, " & VisorderDB & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & VisorderDB & ", " & QuantitàDB & " ,'" & AttrezzaggioDB & " ', " & AttrezzaggioDB & " + (" & QuantitàDB & " * " & quantità & "),'" & MagazzinoDB & "'," & TypeDB & ",'B', CONVERT(DATETIME, '" & data_inizio & "',103), CONVERT(DATETIME, '" & data_fine & "',103),0,0,0," & QuantitàDB & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & DatrasferireDB & " * " & quantità & ",'" & Resallocdb & "')"

            End If

            CMD_SAP_1.ExecuteNonQuery()

            Cnn1.Close()

        Loop

        sostituisci_pseudoarticoli()


        aggiusta_risorse()
        aggiusta_testi()

        cmd_SAP_reader_3.Close()
        Cnn3.Close()





    End Sub

    Sub aggiusta_risorse()
        Cnn4.ConnectionString = Homepage.sap_tirelli
        Cnn4.Open()

        Dim CMD_SAP_4 As New SqlCommand

        CMD_SAP_4.Connection = Cnn4

        CMD_SAP_4.CommandText = "update wor1 set wor1.resalloc='F' where wor1.docentry=" & docentry & "+1 and wor1.itemtype=290"

        CMD_SAP_4.ExecuteNonQuery()

        Cnn4.Close()

    End Sub

    Sub aggiusta_testi()
        cnn5.ConnectionString = Homepage.sap_tirelli
        cnn5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = cnn5

        CMD_SAP_5.CommandText = "update wor1 set wor1.issuetype='', wor1.startdate= null, wor1.enddate= null where wor1.docentry=" & docentry & " and wor1.itemtype='-18'"

        CMD_SAP_5.ExecuteNonQuery()

        cnn5.Close()

    End Sub

    Sub aggiusta_ORDINATO()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP_6 As New SqlCommand

        CMD_SAP_6.Connection = Cnn

        CMD_SAP_6.CommandText = "update t41 set t41.onorder=t40.ORDINATI
from
(
SELECT t10.itemcode,sum(t10.Ordinati) AS 'ORDINATI', t10.MAG
FROM
(
SELECT T0.ITEMCODE, 0 AS 'ORDINATI', T0.WHSCODE AS 'MAG'
FROM OITW T0
WHERE T0.ONORDER>0 OR T0.ONORDER<0
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[PlannedQty]) AS 'CONFERMATI', T0.[wareHouse] AS 'mag'
FROM OWOR T0   WHERE (T0.[STATUS] ='P' OR  T0.[STATUS] ='R')
GROUP BY T0.[ItemCode],T0.[wareHouse]

UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.[WhsCode] 
FROM POR1 T0  INNER JOIN OPOR T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T1.[DocStatus] ='O' AND T0.[OpenQty]>0
GROUP BY 
T0.[ItemCode],  T0.[WhsCode]
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.WHSCODE
FROM WTQ1 T0  INNER JOIN OWTQ T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T0.[OpenQty] >0 AND  T1.[DocStatus] ='O'
GROUP BY 
T0.[ItemCode], T0.WHSCODE
)
AS T10
group by t10.itemcode, t10.MAG
)
as t40 inner join oitw t41 on t41.itemcode=t40.itemcode and t41.whscode=t40.MAG"

        CMD_SAP_6.ExecuteNonQuery()

        Cnn.Close()
    End Sub

    Sub aggiusta_righe_odp_confermato_tot_ordinato_tot()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP_6 As New SqlCommand

        CMD_SAP_6.Connection = Cnn

        CMD_SAP_6.CommandText = "UPDATE T23 SET T23.U_CONFERMATO_TOT=T20.CONF, T23.U_ORDINATO_TOT=T20.ORD, T23.U_DISPONIIBILETOT=T20.MAG-T20.CONF+T20.ORD

FROM
(
SELECT T10.ITEMCODE,SUM(T11.ONHAND) AS 'MAG', SUM(T11.ISCOMMITED) AS 'CONF', SUM(T11.ONORDER) AS 'ORD'
FROM
(
SELECT T0.[ItemCode] 
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE (T1.[Status] ='P' OR  T1.[Status] ='R') AND T0.ITEMTYPE=4 
GROUP BY T0.[ItemCode]
)
AS T10 INNER JOIN OITW T11 ON T11.ITEMCODE=T10.ITEMCODE
GROUP BY T10.ITEMCODE
)
AS T20 INNER JOIN WOR1 T21 ON T20.ITEMCODE=T21.ITEMCODE
INNER JOIN OWOR T22 ON T22.DOCENTRY=T21.DOCENTRY
INNER JOIN WOR1 T23 ON T23.DOCENTRY=T22.DOCENTRY AND T23.ITEMCODE=T20.ITEMCODE

WHERE T22.STATUS='P' OR T22.STATUS='R'"

        CMD_SAP_6.ExecuteNonQuery()

        Cnn.Close()
    End Sub

    Sub aggiusta_CONFERMATO()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand

        CMD_SAP_7.Connection = Cnn

        CMD_SAP_7.CommandText = "update t41 set t41.iscommited=t40.confermati
from
(

SELECT t10.itemcode,sum(t10.confermati) AS 'CONFERMATI', t10.MAG
FROM
(
SELECT T0.ITEMCODE, 0 AS 'CONFERMATI', T0.WHSCODE AS 'MAG'
FROM OITW T0
WHERE T0.ISCOMMITED>0 OR T0.ISCOMMITED<0
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[PlannedQty]) AS 'CONFERMATI', T0.[wareHouse] AS 'mag'
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE (T1.[STATUS] ='P' OR  T1.[STATUS] ='R') AND T1.[CmpltQty]< T1.[PlannedQty] 
GROUP BY T0.[ItemCode],T0.[wareHouse]

UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.[WhsCode] 
FROM RDR1 T0  INNER JOIN ORDR T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T1.[DocStatus] ='O' AND T0.[OpenQty]>0
GROUP BY 
T0.[ItemCode],  T0.[WhsCode]
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.FROMWHSCOD 
FROM WTQ1 T0  INNER JOIN OWTQ T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T0.[OpenQty] >0 AND  T1.[DocStatus] ='O'
GROUP BY 
T0.[ItemCode], T0.FROMWHSCOD
)
AS T10
group by t10.itemcode, t10.MAG

)
as t40 inner join oitw t41 on t41.itemcode=t40.itemcode and t41.whscode=t40.mag"

        CMD_SAP_7.ExecuteNonQuery()

        Cnn.Close()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)

        Commesse_MES.Show()
        Me.Hide()
        Commesse_MES.Commesse_odp_aperte(Commesse_MES.DataGridView_commesse, Commesse_MES.TextBox_commessa.Text, Commesse_MES.TextBox1.Text, Commesse_MES.TextBox2.Text, Commesse_MES.CheckBox1.Checked)



    End Sub





    Sub aggiorna_date_ODP()
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand
        Dim cmd_SAP_reader_3 As SqlDataReader
        CMD_SAP_3.Connection = Cnn3

        CMD_SAP_3.CommandText = "SELECT commessa as 'Commessa', SUBSTRING(risorsa,1,6) as 'Risorsa', case when risorsa= 'P05001' then consegna else min(data_I) end as 'Data Inizio'
from
pianificazione_output
where risorsa <>''
group by commessa, risorsa, consegna"

        cmd_SAP_reader_3 = CMD_SAP_3.ExecuteReader

        Do While cmd_SAP_reader_3.Read()

            commessa = cmd_SAP_reader_3("Commessa")
            risorsa = cmd_SAP_reader_3("Risorsa")
            data_inizio = cmd_SAP_reader_3("Data Inizio")
            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_1 As New SqlCommand

            CMD_SAP_1.Connection = Cnn1
            If risorsa = "P04001" Then
                meno_7_giorni_lavorativi()
            Else
                meno_3_giorni_lavorativi()
            End If


            CMD_SAP_1.CommandText = "UPDATE OWOR
set OWOR.STARTDATE=CONVERT(DATETIME, '" & data_inizio & "', 103), OWOR.DUEDATE=CONVERT(DATETIME, '" & data_inizio & "', 103)
WHERE OWOR.[U_PRG_AZS_Commessa]='" & commessa & "' AND OWOR.[U_Fase]='" & risorsa & "' and (OWOR.[Status]='R' or OWOR.[Status]='P') and OWOR.[u_PRODUZIONE]='ASSEMBL'

UPDATE WOR1
SET WOR1.STARTDATE=CONVERT(DATETIME, '" & data_inizio & "', 103), WOR1.ENDDATE=CONVERT(DATETIME, '" & data_inizio & "', 103)
FROM WOR1 INNER JOIN OWOR ON OWOR.DOCENTRY=WOR1.DOCENTRY
WHERE OWOR.[U_PRG_AZS_Commessa]='" & commessa & "' AND OWOR.[U_Fase]='" & risorsa & "' and (OWOR.[Status]='R' or OWOR.[Status]='P') AND OWOR.[u_PRODUZIONE]='ASSEMBL'"


            CMD_SAP_1.ExecuteNonQuery()

            Cnn1.Close()

        Loop

        cmd_SAP_reader_3.Close()
        Cnn3.Close()
        MsgBox("Date aggiornate con successo")
    End Sub

    Sub meno_3_giorni_lavorativi()
        Dim Giorni_LAv = -3


        data_inizio = Dashboard_pianificazione.AddWorkDays(data_inizio, Giorni_LAv)


    End Sub

    Sub meno_7_giorni_lavorativi()
        Dim Giorni_LAv = -7


        data_inizio = Dashboard_pianificazione.AddWorkDays(data_inizio, Giorni_LAv)


    End Sub

    Private Sub Button1_Click_3(sender As Object, e As EventArgs)
        System.IO.File.OpenRead("C:\Users\GioTirelli.TIRELLISRL\Desktop\5.pdf")
    End Sub



    Private Sub DataGridView_pianificazione_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_Pianificazione.CellFormatting


        If DataGridView_Pianificazione.Rows(e.RowIndex).Cells(columnName:="KPI_").Value = "SI" Then
            DataGridView_Pianificazione.Rows(e.RowIndex).Cells(columnName:="KPI_").Style.BackColor = Color.Green
        Else
            DataGridView_Pianificazione.Rows(e.RowIndex).Cells(columnName:="KPI_").Style.BackColor = Color.Red

        End If

        If DataGridView_Pianificazione.Rows(e.RowIndex).Cells(columnName:="DIV").Value = "BRB01" Then
            DataGridView_Pianificazione.Rows(e.RowIndex).Cells(columnName:="DIV").Style.BackColor = Color.Yellow
        ElseIf DataGridView_Pianificazione.Rows(e.RowIndex).Cells(columnName:="DIV").Value = "KTF01" Then
            DataGridView_Pianificazione.Rows(e.RowIndex).Cells(columnName:="DIV").Style.BackColor = Color.Green

        ElseIf DataGridView_Pianificazione.Rows(e.RowIndex).Cells(columnName:="DIV").Value = "TIR01" Then
            DataGridView_Pianificazione.Rows(e.RowIndex).Cells(columnName:="DIV").Style.BackColor = Color.Blue
            DataGridView_Pianificazione.Rows(e.RowIndex).Cells(columnName:="DIV").Style.ForeColor = Color.White

        End If
    End Sub
    Sub sostituisci_pseudoarticoli()

        Dim NEW_VISORSDER = maxvisorder + 1
        cnn6.ConnectionString = Homepage.sap_tirelli
        cnn6.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn6

        CMD_SAP.CommandText = "SELECT  CASE when T3.CODE is null then '' else t3.code end as 'ItemcodeDB' , t4.itemname as 'ItemnameDB', T3.QUANTITY as 'QuantitàDB' ,T0.PLANNEDQTY AS 'Quantità pseudo', T3.[VisOrder] as 'VisorderDB', case when t3.warehouse is null then '' else T3.[Warehouse] end as 'MagazzinoDB', T3.Type as 'TypeDB', case when T3.AddQuantit is null then 0 else T3.AddQuantit end as 'AttrezzaggioDB', case when T3.LineText is null then '' else t3.linetext end as 'TestoDB', case when t3.type=4 then 'null' else 'F' end as'Resallocdb', case when t3.type=4 then '-1' else '' end as'UOMENTRYDB', case when t3.type=4 then 'Manuale' else '' end as'UomcodeDB', case when (substring(t3.code,1,1)='0' or substring(t3.code,1,1)='C' or substring(t3.code,1,1)='D' or substring(t3.code,1,1)='F') then t3.quantity else 0 end as 'DatrasferireDB' 
FROM WOR1 T0 inner join OITM T1 on t0.itemcode=t1.itemcode 
INNER JOIN ITT1 T3 ON T3.FATHER=T0.ITEMCODE
INNER JOIN OITM T4 ON T4.ITEMCODE=T3.CODE
WHERE T0.[DOCENTRY] =" & docentry & "+1 and t1.phantom='Y'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()


            If cmd_SAP_reader("resallocdb") = "null" Then
                Resallocdb = Nothing
            Else
                Resallocdb = cmd_SAP_reader("resallocdb")
            End If

            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_1 As New SqlCommand

            CMD_SAP_1.Connection = Cnn1

            If TypeDB = 4 Then

                CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, WOR1.ITEMNAME, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty],WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType,  wor1.[StartDate], wor1.[EndDate], wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext,WOR1.[U_PRG_WIP_QtaDaTrasf])  
                                                    VALUES(" & docentry & "+1,'" & NEW_VISORSDER & "','" & cmd_SAP_reader("ItemcodeDB") & "','" & cmd_SAP_reader("ItemnameDB") & "','" & NEW_VISORSDER & "', " & Replace(cmd_SAP_reader("QuantitàDB"), ",", ".") & ",'" & Replace(cmd_SAP_reader("AttrezzaggioDB"), ",", ".") & " '," & Replace(cmd_SAP_reader("AttrezzaggioDB"), ",", ".") & " + (" & Replace(cmd_SAP_reader("QuantitàDB"), ",", ".") & "* " & Replace(cmd_SAP_reader("Quantità pseudo"), ",", ".") & "),'" & cmd_SAP_reader("MagazzinoDB") & "'," & cmd_SAP_reader("TypeDB") & ",'B', CONVERT(DATETIME, '" & data_inizio & "',103), CONVERT(DATETIME, '" & data_fine & "',103),0,0,0," & Replace(cmd_SAP_reader("QuantitàDB"), ",", ".") & ",1,0,' " & cmd_SAP_reader("UomentryDB") & "', '" & cmd_SAP_reader("UomcodeDB") & "', '" & cmd_SAP_reader("TestoDB") & "'," & Replace(cmd_SAP_reader("DatrasferireDB"), ",", ".") & " * " & Replace(cmd_SAP_reader("Quantità pseudo"), ",", ".") & ")"


            Else
                CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty],WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType,  wor1.[StartDate], wor1.[EndDate], wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext,WOR1.[U_PRG_WIP_QtaDaTrasf],wor1.resalloc)  
                                                    VALUES(" & docentry & "+1,'" & NEW_VISORSDER & "','" & cmd_SAP_reader("ItemcodeDB") & "','" & NEW_VISORSDER & "', " & Replace(cmd_SAP_reader("QuantitàDB"), ",", ".") & ",'" & Replace(cmd_SAP_reader("AttrezzaggioDB"), ",", ".") & " '," & Replace(cmd_SAP_reader("AttrezzaggioDB"), ",", ".") & " + (" & Replace(cmd_SAP_reader("QuantitàDB"), ",", ".") & "* " & Replace(cmd_SAP_reader("Quantità pseudo"), ",", ".") & "),'" & cmd_SAP_reader("MagazzinoDB") & "'," & cmd_SAP_reader("TypeDB") & ",'B', CONVERT(DATETIME, '" & data_inizio & "',103), CONVERT(DATETIME, '" & data_fine & "',103),0,0,0," & Replace(cmd_SAP_reader("QuantitàDB"), ",", ".") & ",1,0,' " & cmd_SAP_reader("UomentryDB") & "', '" & cmd_SAP_reader("UomcodeDB") & "', '" & cmd_SAP_reader("TestoDB") & "'," & Replace(cmd_SAP_reader("DatrasferireDB"), ",", ".") & " * " & Replace(cmd_SAP_reader("Quantità pseudo"), ",", ".") & ",'" & Resallocdb & "')"


            End If

            CMD_SAP_1.ExecuteNonQuery()

            Cnn1.Close()
            If TypeDB = 4 Then
                Dim Cnn As New SqlConnection
                Cnn.ConnectionString = Homepage.sap_tirelli
                Cnn.Open()

                Dim CMD_SAP_7 As New SqlCommand

                CMD_SAP_7.Connection = Cnn

                CMD_SAP_7.CommandText = "update t0 set T0.[IsCommited]=T0.[IsCommited]+" & Replace(cmd_SAP_reader("AttrezzaggioDB"), ",", ".") & " + (" & Replace(cmd_SAP_reader("QuantitàDB"), ",", ".") & "* " & Replace(cmd_SAP_reader("Quantità pseudo"), ",", ".") & ") from oitw t0 where t0.itemcode='" & cmd_SAP_reader("ItemcodeDB") & "' and t0.whscode='" & cmd_SAP_reader("MagazzinoDB") & "'"

                CMD_SAP_7.ExecuteNonQuery()

                Cnn.Close()

            End If
            NEW_VISORSDER = NEW_VISORSDER + 1
        Loop

        Cnn4.ConnectionString = Homepage.sap_tirelli
        Cnn4.Open()

        Dim CMD_SAP_4 As New SqlCommand

        CMD_SAP_4.Connection = Cnn4

        CMD_SAP_4.CommandText = "DELETE WOR1
FROM WOR1 T0  
inner join oitm t2 on t2.itemcode=t0.itemcode
WHERE T0.[DocEntry]=" & docentry & "+1 and t2.phantom='Y'"


        CMD_SAP_4.ExecuteNonQuery()

        Cnn4.Close()

        cnn6.Close()


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Chiudi_commessa_M()

    End Sub

    Private Sub DataGridView_Pianificazione_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_Pianificazione.CellClick
        If e.RowIndex >= 0 Then
            commessa = Trim(DataGridView_Pianificazione.Rows(e.RowIndex).Cells(columnName:="Commessa_Tab").Value)
            Dashboard_pianificazione.commessa = Trim(DataGridView_Pianificazione.Rows(e.RowIndex).Cells(columnName:="Commessa_Tab").Value)
            RIGA = e.RowIndex


            If e.ColumnIndex = DataGridView_Pianificazione.Columns.IndexOf(Pian) Then




                Dim Cnn1 As New SqlConnection
                Cnn1.ConnectionString = Homepage.sap_tirelli
                Cnn1.Open()

                Dim CMD_SAP_1 As New SqlCommand
                Dim cmd_SAP_reader_1 As SqlDataReader


                CMD_SAP_1.Connection = Cnn1
                CMD_SAP_1.CommandText = "SELECT T1.[ItemCode] as 'Commessa', T2.[ItemName] as 'Nome commessa', T0.[docnum] as 'OC', T0.[cardname] as 'cliente', case when t0.U_clientefinale is null then '' else t0.U_clientefinale end as 'Cliente finale',  T1.[ShipDate] as 'Consegna',  t3.stato as 'Stato', case when t3.kpi is null then 'NO' else t3.kpi end as 'KPI' 
            FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.[DocEntry] = T1.[DocEntry] INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode]
            left join [Tirelli_40].[dbo].pianificazione_commessa t3 on t3.commessa=t1.itemcode
            WHERE t1.itemcode= '" & commessa & "' and T0.DOCSTATUS='o'
            order by T1.[ItemCode]"

                cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader


                If cmd_SAP_reader_1.Read() = True Then

                    Nome_commessa = cmd_SAP_reader_1("Nome commessa")
                    Consegna = cmd_SAP_reader_1("Consegna")
                    Cliente = cmd_SAP_reader_1("Cliente")
                    Cliente_finale = cmd_SAP_reader_1("Cliente finale")
                    Consegna = cmd_SAP_reader_1("Consegna")
                    Ordine_cliente = cmd_SAP_reader_1("OC")
                    stato = cmd_SAP_reader_1("Stato")
                    KPI = cmd_SAP_reader_1("KPI")

                    cmd_SAP_reader_1.Close()

                Else

                    Cnn2.ConnectionString = Homepage.sap_tirelli
                    Cnn2.Open()

                    Dim CMD_SAP_2 As New SqlCommand
                    Dim cmd_SAP_reader_2 As SqlDataReader
                    CMD_SAP_2.Connection = Cnn2

                    CMD_SAP_2.CommandText = "SELECT  t0.commessa as 'Commessa', t0.descrizione as 'Nome commessa', t0.consegna as 'Consegna', case when t0.oc is null then '' else t0.oc end as 'OC', case when t0.cliente is null then '' else t0.cliente end  as 'Cliente', min(case when t1.data_i is null then '' else t1.data_i end) as 'Min Data_i', max(case when t1.data_f is null then '' else t1.data_f end) as 'Max Data_f', t0.stato, case when t0.kpi is null then '' else t0.kpi end as 'KPI'
from [Tirelli_40].[dbo].pianificazione_commessa t0 left join [Tirelli_40].[dbo].pianificazione t1 on t0.commessa=t1.commessa
where  t0.commessa = '" & commessa & "'
group by t0.commessa , t0.descrizione, t0.consegna , t0.oc, t0.cliente, t0.stato, t0.kpi
order by t0.consegna, t0.oc"

                    cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
                    If cmd_SAP_reader_2.Read() = True Then

                        Nome_commessa = cmd_SAP_reader_2("Nome commessa")
                        Consegna = cmd_SAP_reader_2("Consegna")
                        Cliente = cmd_SAP_reader_2("Cliente")
                        Cliente_finale = ""
                        Consegna = cmd_SAP_reader_2("Consegna")
                        Ordine_cliente = cmd_SAP_reader_2("OC")
                        stato = cmd_SAP_reader_2("stato")
                        KPI = cmd_SAP_reader_2("KPI")

                    End If
                    cmd_SAP_reader_2.Close()
                    Cnn2.Close()

                End If

                Cnn1.Close()



                Dashboard_pianificazione.Show()
                Dashboard_pianificazione.iniziazione_form()

                risorsa = Dashboard_pianificazione.risorsa_appoggio
                Dashboard_pianificazione.ComboBox_dipendente.Text = Nothing

            End If
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Close()
    End Sub

    Sub filtra()
        Dim i = 0
        Dim parola0 As String
        Dim parola1 As String
        Dim parola2 As String


        Do While i < DataGridView_Pianificazione.RowCount

            Try

                parola0 = UCase(DataGridView_Pianificazione.Rows(i).Cells(0).Value)
                parola1 = UCase(DataGridView_Pianificazione.Rows(i).Cells(1).Value)
                parola2 = UCase(DataGridView_Pianificazione.Rows(i).Cells(2).Value)


                If parola0.Contains(UCase(TextBox8.Text)) Then
                    DataGridView_Pianificazione.Rows(i).Visible = True
                    If parola1.Contains(UCase(TextBox7.Text)) Then
                        DataGridView_Pianificazione.Rows(i).Visible = True


                        If parola2.Contains(UCase(TextBox5.Text)) Then
                            DataGridView_Pianificazione.Rows(i).Visible = True

                        Else
                            DataGridView_Pianificazione.Rows(i).Visible = False

                        End If


                    Else
                        DataGridView_Pianificazione.Rows(i).Visible = False

                    End If

                Else
                    DataGridView_Pianificazione.Rows(i).Visible = False

                End If

            Catch ex As Exception
                DataGridView_Pianificazione.Rows(i).Visible = False
            End Try
            i = i + 1
        Loop
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        filtra()
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        filtra()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        filtra()
    End Sub




    Private Sub Button1_Click_4(sender As Object, e As EventArgs) Handles Button1.Click
        elimina_commessa_M()
        MsgBox("Commessa eliminata")
    End Sub



    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = "" Then
            filtro_divisione = ""
        Else
            filtro_divisione = " and t12.ocrcode   Like '%%" & TextBox4.Text & "%%'  "
        End If
        Commesse_aperte()
    End Sub

    Public Sub date_jpm_commessa(par_commessa As String, par_datagridview As DataGridView)
        par_datagridview.Rows.Clear()


        Dim connString As String = Homepage.JPM_TIRELLI
        Using conn As New NpgsqlConnection(connString)
            conn.Open()
            ' Esegui le query qui
            Dim STRINGA_QUERY As String = ""

            Dim cmd As New NpgsqlCommand(STRINGA_QUERY, conn)
            Dim reader As NpgsqlDataReader = cmd.ExecuteReader()


            Do While reader.Read()

                par_datagridview.Rows.Add(reader("ATT_WBS"), reader("ATT_DES"), reader("ATT_DTPIAINI"), reader("ATT_DTPIAFIN"), reader("DURATAORI"))

            Loop




            reader.Close()
            conn.Close()
        End Using




    End Sub

End Class
