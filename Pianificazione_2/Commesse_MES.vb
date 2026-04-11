Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports Tirelli.Magazzino


Public Class Commesse_MES
    Public OC As String
    Public CDS_ As String










    Public riga As Integer

    Public CONDIZIONI_COMMESSE As String = "  substring(T0.[U_PRG_AZS_Commessa],1,1)='M' and (T0.[Status] ='P' or  T0.[Status] ='R' OR T0.[CloseDate]>=getdate()-60) or T0.[U_PRG_AZS_Commessa]='M05505'  or T0.[U_PRG_AZS_Commessa]='M05506'  or T0.[U_PRG_AZS_Commessa]='M05507'  or T0.[U_PRG_AZS_Commessa]='M05508'  or T0.[U_PRG_AZS_Commessa]='M05509'  or T0.[U_PRG_AZS_Commessa]='M05510'  or T0.[U_PRG_AZS_Commessa]='M05511'  or T0.[U_PRG_AZS_Commessa]='M05512'  or T0.[U_PRG_AZS_Commessa]='M05513' "
    Public Property Filtro_docnum As String
    Public Property Filtro_riferimento As String
    Public Property Filtro_causcons As String
    Public Property Filtro_azione As String
    Public Property Filtro_cliente_f As String
    Public Property Filtro_cliente As String
    Public Property Filtro_commento As String
    Public Property Filtro_cds As String



    Public n_record As Integer = 50

    Sub Commesse_odp_aperte(par_datagridview As DataGridView, par_itemcode As String, par_itemname As String, par_cliente As String, par_solo_M As Boolean)
        Dim filtro_solo_m As String

        If par_solo_M = True Then
            filtro_solo_m = " and substring(t0.matricola,1,1)=''M'' "
        Else
            filtro_solo_m = ""
        End If
        Dim Filtro_commessa_datagrid As String
        Dim Filtro_descrizione_datagrid As String
        Dim Filtro_cliente_datagrid As String

        If par_itemcode = "" Then
            Filtro_commessa_datagrid = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then
                Filtro_commessa_datagrid = " AND t10.commessa LIKE '%" & par_itemcode & "%' "

            End If

        End If

        If par_itemname = "" Then
            Filtro_descrizione_datagrid = ""
        Else

            Filtro_descrizione_datagrid = " And t10.descrizione Like '%%" & par_itemname & "%%' "
        End If


        If par_cliente = "" Then
            Filtro_cliente_datagrid = ""
        Else

            Filtro_cliente_datagrid = "  And (t10.cliente Like '%%" & par_cliente & "%%'  or t10.cliente_finale Like '%%" & par_cliente & "%%') "
        End If


        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_1.CommandText = " 
select *
from
(
SELECT distinct(T0.[U_PRG_AZS_Commessa]) as 'Commessa'
 , t1.itemname as 'Descrizione'
 , t3.cardcode as 'Codice_bp'
 , case when t3.cardname is null then T1.[U_Final_customer_name] else t3.cardname end  as 'Cliente'
 , case when t4.cardcode is null then '' else t4.cardcode end  as 'Codice cliente finale'
 , case when T4.[cardname] is null then '' else T4.[cardname] end as 'Cliente_finale'
 , T3.[DocDueDate]  as 'Consegna'

FROM OWOR T0 left join oitm t1 on t0.[U_PRG_AZS_Commessa] =t1.itemcode

				LEFT JOIN (SELECT distinct(T0.[U_PRG_AZS_Commessa]) as 'Commessa', max(t2.docentry) as 'Docentry'

				FROM OWOR T0 left join oitm t1 on t0.[U_PRG_AZS_Commessa] =t1.itemcode
				left join rdr1 t2 on t1.itemcode=t2.itemcode 
				left join ordr t3 on t3.docentry=t2.docentry
				WHERE    substring(T0.[U_PRG_AZS_Commessa],1,1)='M' and (T0.[Status] ='P' or  T0.[Status] ='R' OR T0.[CloseDate]>=getdate()-60) and (t0.u_produzione='ASSEMBL' OR T0.U_PRODUZIONE='EST') and t3.CANCELED<>'Y'
				group by 
				T0.[U_PRG_AZS_Commessa]) A ON T0.[U_PRG_AZS_Commessa]=a.Commessa

				left join  rdr1 t2 on T0.[U_PRG_AZS_Commessa]=t2.itemcode and T2.[docentry]=a.Docentry
				left join ordr t3 on t3.docentry=t2.docentry
				left join ocrd t4 on t4.cardcode=t3.U_CodiceBP


WHERE    " & CONDIZIONI_COMMESSE & " and (t0.u_produzione='ASSEMBL' OR T0.U_PRODUZIONE='EST') 
group by 
T0.[U_PRG_AZS_Commessa], t1.itemname , t3.cardcode, t3.cardname, T4.[cardname], T3.[DocDueDate],T1.[U_Final_customer_name], t4.cardcode
)
as t10
where 0 = 0 " & Filtro_cliente_datagrid & " " & Filtro_commessa_datagrid & " " & Filtro_descrizione_datagrid & " 
order by
 T10.commessa"
        Else

            CMD_SAP_1.CommandText = "SELECT top 100 trim(t10.matricola) as 'Commessa'
, t10.itemname as 'Descrizione', t10.desc_supp
, T10.DSCLI_FATT as 'Cliente'
, T10.CLI_FATT as 'Codice_cliente',
        t10.codice_finale as 'Cliente_finale'
		, t10.itemcode as 'absentry',
        trim(t10.itemcode) as 'Numero_progetto',
		T10.NAME_progetto AS 'DESC_PROGETTO',
		'' as 'Nome_stato',
        '' as 'Livello_rischio_totale', '' as 'Name',
        t10.pm as 'CODICE_PM'
		,t10.DESC_pm as 'PM'
		, T10.DSNAZ_FINALE as u_country_of_delivery,
        t10.brand AS 'CODICE_BRAND',
		T10.DESC_BRAND AS 'BRAND',
		'' as 'Baia'
		, '' as 'Zona'

		,T10.NOME_STATO AS 'STATO_COMMESSA'
FROM OPENQUERY(AS400, '
    SELECT *
    FROM TIR90VIS.JGALCOM t0
    WHERE 

       UPPER(t0.matricola) LIKE ''%%" & par_itemcode.ToUpper() & "%%''
      AND UPPER(t0.itemname) LIKE ''%%" & par_itemname.ToUpper() & "%%''
AND (UPPER(t0.codice_finale) LIKE ''%%" & par_cliente.ToUpper() & "%%'' 
OR UPPER(t0.dscli_fatt) LIKE ''%%" & par_cliente.ToUpper() & "%%'') 
           

AND T0.NOME_STATO=''Aperta'' " & filtro_solo_m & "
       

ORDER BY t0.matricola DESC

limit 100  
') T10"
        End If



        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()

            par_datagridview.Rows.Add(cmd_SAP_reader_1("Commessa"), cmd_SAP_reader_1("Descrizione"), cmd_SAP_reader_1("Cliente"), cmd_SAP_reader_1("Cliente_finale"))

        Loop
        Cnn1.Close()
        par_datagridview.ClearSelection()
    End Sub




    Sub Commesse_odp_cds_aperte()
        DataGridView_commesse.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1

        CMD_SAP_1.CommandText = " SELECT distinct(T0.[U_PRG_AZS_Commessa]) as 'Commessa', t1.itemname as 'Descrizione', t3.cardname as 'Cliente', case when T3.[U_Clientefinale] is null then '' else T3.[U_Clientefinale] end as 'Cliente finale', T3.[DocDueDate] as 'Consegna'

FROM OWOR T0 left join oitm t1 on t0.[U_PRG_AZS_Commessa] =t1.itemcode
left join rdr1 t2 on t1.itemcode=t2.itemcode and T2.[OpenQty]>0
left join ordr t3 on t3.docentry=t2.docentry and t3.docstatus='O'

WHERE (T0.[Status] ='P' or  T0.[Status] ='R' or T0.[CloseDate]>=getdate()-30) AND substring(T0.[U_PRG_AZS_Commessa],1,1)<>'M'
group by 
T0.[U_PRG_AZS_Commessa], t1.itemname , t3.cardname, T3.[U_Clientefinale], T3.[DocDueDate]

order by
 T0.[U_PRG_AZS_Commessa]"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()
            Try
                DataGridView_commesse.Rows.Add(cmd_SAP_reader_1("Commessa"), cmd_SAP_reader_1("Descrizione"), cmd_SAP_reader_1("Cliente"), cmd_SAP_reader_1("Cliente finale"), Format(cmd_SAP_reader_1("Consegna"), "dd/MM/yy"))
            Catch ex As Exception
                DataGridView_commesse.Rows.Add(cmd_SAP_reader_1("Commessa"), cmd_SAP_reader_1("Descrizione"), cmd_SAP_reader_1("Cliente"), cmd_SAP_reader_1("Cliente finale"), cmd_SAP_reader_1("Consegna"))
            End Try
        Loop
        Cnn1.Close()

    End Sub

    Sub Commesse_OC_aperte()
        MsgBox("è stato tolto")
    End Sub


    Public Function SCHEDA_COMMESSA(par_Codice_commessa As String) As Dettaglicommessa

        Dim dettagli As New Dettaglicommessa()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_2.CommandText = "SELECT t0.docnum, T1.[ItemCode] as 'Commessa', T2.[ItemName] as 'Nome commessa', case when (t2.itemname Like '%%Riempitrice%%' OR t2.itemname Like '%%MONOBLOCCO%%')  then 'Y' else 'N' end as 'Riempimento',  T0.[docnum] as 'OC', T0.[cardname] as 'cliente', case when t4.CARDNAME is null then '' else t4.CARDNAME end as 'Cliente finale',  T1.[ShipDate] as 'Consegna',  t3.stato as 'Stato' ,CAST(T1.[ShipDate]-getdate() AS INTEGER) as 'Giorni alla consegna', case when T2.U_COUNTRY_OF_DELIVERY is null then '' else T2.U_COUNTRY_OF_DELIVERY end  AS 'Destinazione', T0.CARDCODE AS 'CODICE_BP', CASE WHEN T0.U_CODICEBP IS NULL THEN '' ELSE T0.U_CODICEBP END AS 'CODICE_FINAL_BP'
            FROM ORDR T0  INNER JOIN RDR1 T1 ON T0.[DocEntry] = T1.[DocEntry] INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode]
            left join [TIRELLI_40].[DBO].[pianificazione_commessa] t3 on t3.commessa=t1.itemcode
LEFT JOIN OCRD T4 ON T4.CARDCODE=T0.U_CODICEBP
            WHERE t1.itemcode= '" & par_Codice_commessa & "' 
            order by t0.docnum DESC"
        Else
            CMD_SAP_2.CommandText = "

			SELECT top 100 
			t10.itemcode as 'docnum'
			,trim(t10.matricola) as 'Commessa'
			, t10.itemname as 'Nome commessa'
			
			,'N' as 'Riempimento'
			,t10.itemcode as 'OC'
			, T10.DSCLI_FATT as 'Cliente'
			  ,t10.codice_finale as 'Cliente finale'
			  
			,t10.DATA_CONSEGNA AS 'Consegna'
			,t10.nome_stato AS 'Stato'
			,999 as 'Giorni alla consegna'
			,T10.DSNAZ_FINALE as 'Destinazione'
			

, T10.CLI_FATT as 'Codice_BP',
t10.codice_finale as 'CODICE_FINAL_BP'
      
		,t10.itemcode as 'absentry',
        trim(t10.itemcode) as 'Numero_progetto',
		T10.NAME_progetto AS 'DESC_PROGETTO',
		'' as 'Nome_stato',
        '' as 'Livello_rischio_totale', '' as 'Name',
        t10.pm as 'CODICE_PM'
		,t10.DESC_pm as 'PM'
	
        ,t10.brand AS 'CODICE_BRAND',
		T10.DESC_BRAND AS 'BRAND',
		'' as 'Baia'
		, '' as 'Zona'
		
		,T10.NOME_STATO AS 'STATO_COMMESSA'
FROM OPENQUERY(AS400, '
    
	SELECT *
    FROM TIR90VIS.JGALCOM t0
    WHERE 
t0.matricola=''" & par_Codice_commessa & "''
         
ORDER BY t0.matricola DESC

limit 100  
') T10"
        End If


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() = True Then

            dettagli.Descrizione_commessa = cmd_SAP_reader_2("Nome commessa")
            dettagli.ordine_cliente_commessa = cmd_SAP_reader_2("OC")
            dettagli.Cliente_commessa = cmd_SAP_reader_2("Cliente")
            dettagli.Cliente_finale_commessa = cmd_SAP_reader_2("Cliente finale")
            'dettagli.Consegna_commessa = cmd_SAP_reader_2("Consegna")
            dettagli.Giorni_alla_consegna = cmd_SAP_reader_2("Giorni alla consegna")
            dettagli.codice_cliente = cmd_SAP_reader_2("Codice_bp")
            dettagli.codice_cliente_finale = cmd_SAP_reader_2("Codice_final_bp")


            If cmd_SAP_reader_2("Destinazione") = Nothing Then
                dettagli.destinazione = ""
            Else
                dettagli.destinazione = UCase(cmd_SAP_reader_2("Destinazione"))
            End If

            'If cmd_SAP_reader_2("Riempimento") = "Y" Then
            '    FORM6.Button9.Enabled = True
            'Else
            '    FORM6.Button9.Enabled = False
            'End If


            cmd_SAP_reader_2.Close()
        Else


        End If
        Cnn1.Close()
        Return dettagli
    End Function





    Sub date_inizio_fine_commesse_mostra()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.COMMESSA AS 'COMM',t0.DESCRIZIONE AS 'DESCR',t0.CLIENTE,t0.CLIENTE_F,t0.CONSEGNA, 

min(case when t0.RISORSA ='P00001' and t0.attivita='ANT' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz ANT_UT', max(case when t0.RISORSA ='P00001'and t0.attivita='ANT' then t0.data_f else 0 end) as 'Fine ANT_UT', 
min(case when t0.RISORSA ='P00001' and t0.attivita='Prog' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz Prog_UT', max(case when t0.RISORSA ='P00001'and t0.attivita='Prog' then t0.data_f else 0 end) as 'Fine Prog_UT',
min(case when t0.RISORSA ='P00001' and t0.attivita='Formati' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz Form_UT', max(case when t0.RISORSA ='P00001'and t0.attivita='Formati' then t0.data_f else 0 end) as 'Fine Form_UT',
min(case when t0.RISORSA ='P01001' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz APPROV', max(case when t0.RISORSA ='P01001' then t0.data_f else 0 end) as 'Fine APPROV', 
min(case when t0.RISORSA ='P01501' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz PREM', max(case when t0.RISORSA ='P01501' then t0.data_f else 0 end) as 'Fine PREM',
min(case when t0.RISORSA ='P02001' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) 'Iniz MEC_MONT', max(case when t0.RISORSA ='P02001' then t0.data_f else 0 end) as 'Fine MEC_MONT',
min(case when t0.RISORSA ='P03001' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz ELE_MONT', max(case when t0.RISORSA ='P03001' then t0.data_f else 0 end) as 'Fine ELE_MONT',
min(case when t0.RISORSA ='P03001' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz ELE_MONT', max(case when t0.RISORSA ='P03001' then t0.data_f else 0 end) as 'Fine ELE_MONT',
min(case when t0.RISORSA ='P04001' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz COLL', max(case when t0.RISORSA ='P04001' then t0.data_f else 0 end) as 'Fine COLL',
min(case when t0.RISORSA ='P05001' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as '1a CONS'

from [Tirelli_40].[dbo].[PIANIFICAZIONE_OUTPUT] t0 

where t0.livello=1 and t0.commessa='" & Pianificazione.commessa & "'
group by t0.LIVELLO,t0.COMMESSA ,t0.DESCRIZIONE,t0.CLIENTE,t0.CLIENTE_F,t0.CONSEGNA
order by t0.CONSEGNA"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If cmd_SAP_reader_2("Iniz PREM") = "29/09/9999" Then
                ' Mostra.Label_data_inizio_prem.Text = "Non pianificato"
            Else
                ' Mostra.Label_data_inizio_prem.Text = cmd_SAP_reader_2("Iniz PREM")
            End If

            If cmd_SAP_reader_2("Fine PREM") = "01/01/1900" Then
                '  Mostra.Label_data_fine_prem.Text = "Non pianificato"
            Else
                '  Mostra.Label_data_fine_prem.Text = cmd_SAP_reader_2("Fine PREM")
            End If
            If cmd_SAP_reader_2("Iniz MEC_MONT") = "29/09/9999" Then
                '   Mostra.Label_data_inizio_mont.Text = "Non pianificato"
            Else
                '   Mostra.Label_data_inizio_mont.Text = cmd_SAP_reader_2("Iniz MEC_MONT")
            End If
            If cmd_SAP_reader_2("Fine MEC_MONT") = "01/01/1900" Then
                '  Mostra.Label_data_fine_mont.Text = "Non pianificato"
            Else
                '  Mostra.Label_data_fine_mont.Text = cmd_SAP_reader_2("Fine MEC_MONT")
            End If

            '  If cmd_SAP_reader_2("Iniz ELE_MONT") = "29/09/9999" Then
            '  Mostra.Label_data_inizio_el.Text = "Non pianificato"
            '   Else
            '  Mostra.Label_data_inizio_el.Text = cmd_SAP_reader_2("Iniz ELE_MONT")
            'End If

            'If cmd_SAP_reader_2("Fine ELE_MONT") = "01/01/1900" Then

            'Mostra.Label_data_fine_el.Text = ""
            'Else
            ' Mostra.Label_data_fine_el.Text = cmd_SAP_reader_2("Fine ELE_MONT")
            ' If

            ' If cmd_SAP_reader_2("Iniz COLL") = "29/09/9999" Then
            ' Mostra.Label_data_inizio_collaudo.Text = "Non pianificato"
            '  Else
            '  Mostra.Label_data_inizio_collaudo.Text = cmd_SAP_reader_2("Iniz COLL")
            ' End If
            ' If cmd_SAP_reader_2("Fine COLL") = "01/01/1900" Then
            '  Mostra.Label_data_fine_collaudo.Text = ""
            ' Else
            ' Mostra.Label_data_fine_collaudo.Text = cmd_SAP_reader_2("Fine COLL")
        End If

        cmd_SAP_reader_2.Close()
        'End If
        Cnn1.Close()
    End Sub

    Sub date_inizio_fine_commesse()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.COMMESSA AS 'COMM',t0.DESCRIZIONE AS 'DESCR',t0.CLIENTE,t0.CONSEGNA, 

min(case when t0.RISORSA ='P00001' and t0.attivita='ANT' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz ANT_UT', max(case when t0.RISORSA ='P00001'and t0.attivita='ANT' then t0.data_f else 0 end) as 'Fine ANT_UT', 
min(case when t0.RISORSA ='P00001' and t0.attivita='Prog' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz Prog_UT', max(case when t0.RISORSA ='P00001'and t0.attivita='Prog' then t0.data_f else 0 end) as 'Fine Prog_UT',
min(case when t0.RISORSA ='P00001' and t0.attivita='Formati' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz Form_UT', max(case when t0.RISORSA ='P00001'and t0.attivita='Formati' then t0.data_f else 0 end) as 'Fine Form_UT',
min(case when t0.RISORSA ='P01001' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz APPROV', max(case when t0.RISORSA ='P01001' then t0.data_f else 0 end) as 'Fine APPROV', 
min(case when t0.RISORSA ='P01501' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz PREM', max(case when t0.RISORSA ='P01501' then t0.data_f else 0 end) as 'Fine PREM',
min(case when t0.RISORSA ='P02001' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) 'Iniz MEC_MONT', max(case when t0.RISORSA ='P02001' then t0.data_f else 0 end) as 'Fine MEC_MONT',
min(case when t0.RISORSA ='P03001' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz ELE_MONT', max(case when t0.RISORSA ='P03001' then t0.data_f else 0 end) as 'Fine ELE_MONT',
min(case when t0.RISORSA ='P04001' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as 'Iniz COLL', max(case when t0.RISORSA ='P04001' then t0.data_f else 0 end) as 'Fine COLL',
min(case when t0.RISORSA ='P05001' then t0.data_i else CONVERT(DATETIME, '99990929', 112) end) as '1a CONS'

from [TIRELLI_40].DBO.[pianificazione_output] t0 

where t0.livello=1 and t0.commessa='" & Pianificazione.commessa & "'
group by t0.LIVELLO,t0.COMMESSA ,t0.DESCRIZIONE,t0.CLIENTE,t0.CONSEGNA
order by t0.CONSEGNA"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If cmd_SAP_reader_2("Iniz PREM") = "29/09/9999" Then
                FORM6.Label12.Text = "Non pianificato"
            Else
                FORM6.Label12.Text = cmd_SAP_reader_2("Iniz PREM")
            End If

            If cmd_SAP_reader_2("Fine PREM") = "01/01/1900" Then
                FORM6.Label11.Text = "Non pianificato"
            Else
                FORM6.Label11.Text = cmd_SAP_reader_2("Fine PREM")
            End If
            If cmd_SAP_reader_2("Iniz MEC_MONT") = "29/09/9999" Then
                FORM6.Label6.Text = "Non pianificato"
            Else
                FORM6.Label6.Text = cmd_SAP_reader_2("Iniz MEC_MONT")
            End If
            If cmd_SAP_reader_2("Fine MEC_MONT") = "01/01/1900" Then
                FORM6.Label9.Text = "Non pianificato"
            Else
                FORM6.Label9.Text = cmd_SAP_reader_2("Fine MEC_MONT")
            End If

            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub




    Private Sub ComboBox_dipendente_SelectedIndexChanged(sender As Object, e As EventArgs)
        DataGridView_commesse.Show()

    End Sub

    Private Sub Button_CDS_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        DataGridView_commesse.Hide()
        'DataGridView_CDS.Show()
        Commesse_OC_aperte()
    End Sub


    'Private Sub DataGridView_OC_CellClick(sender As Object, e As DataGridViewCellEventArgs)
    '    If e.RowIndex >= 0 Then
    '        riga = e.RowIndex
    '        OC = DataGridView_CDS.Rows(e.RowIndex).Cells(columnName:="N_DOC").Value
    '        Try
    '            CDS_ = DataGridView_CDS.Rows(e.RowIndex).Cells(columnName:="CDS").Value
    '        Catch ex As Exception
    '            CDS_ = "999999999"
    '        End Try


    '        If e.ColumnIndex = DataGridView_CDS.Columns.IndexOf(N_DOC) Then

    '            Materiale_CDS.inserimento_reparti()
    '            Materiale_CDS.Testata_ordine()
    '            Materiale_CDS.righe_ordine()
    '            Materiale_CDS.Button_commessa.Text = "OC " & OC
    '            Materiale_CDS.Label_cliente.Text = DataGridView_CDS.Rows(e.RowIndex).Cells(columnName:="DataGridViewTextBoxColumn2").Value
    '            Try
    '                Materiale_CDS.Label_cliente_finale.Text = DataGridView_CDS.Rows(e.RowIndex).Cells(columnName:="Cliente_Finale").Value
    '            Catch ex As Exception
    '                Materiale_CDS.Label_cliente_finale.Text = ""
    '            End Try

    '            Materiale_CDS.DataGridView_riga_OC.Show()

    '            Materiale_CDS.Show()

    '        End If

    '        If e.ColumnIndex = DataGridView_CDS.Columns.IndexOf(CDS) Then

    '            Materiale_CDS.DataGridView_riga_OC.Hide()

    '            Materiale_CDS.Button_commessa.Text = "CDS" & CDS_
    '            Materiale_CDS.Label_cliente.Text = DataGridView_CDS.Rows(e.RowIndex).Cells(columnName:="DataGridViewTextBoxColumn2").Value
    '            Materiale_CDS.Label_cliente_finale.Text = DataGridView_CDS.Rows(e.RowIndex).Cells(columnName:="Cliente_Finale").Value
    '            Materiale_CDS.Show()
    '            Materiale_CDS.Owner = Me

    '        End If




    '    End If
    'End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub


    Sub leggi_ini()
        Dim File_INI_Stream As StreamReader
        Dim Str_Lettura As String
        If File.Exists("C:\Program Files\MES.INI") Then
            File_INI_Stream = My.Computer.FileSystem.OpenTextFileReader("C:\Program Files\MES.INI")
            Str_Lettura = File_INI_Stream.ReadLine
            Do While Not Str_Lettura Is Nothing

                If Str_Lettura.StartsWith("[Mostra_macchina]=") Then
                    Pianificazione.commessa = Str_Lettura.Remove(0, 17)
                End If


                Str_Lettura = File_INI_Stream.ReadLine
            Loop
            File_INI_Stream.Close()
        Else
            File.Create("C:\Program Files\MES.INI")

        End If

    End Sub

    '    Sub riempi_CDS()

    '        DataGridView_CDS.Rows.Clear()

    '        Dim Cnn1 As New SqlConnection
    '        Cnn1.ConnectionString = Homepage.sap_tirelli
    '        Cnn1.Open()

    '        Dim CMD_SAP_1 As New SqlCommand
    '        Dim cmd_SAP_reader_1 As SqlDataReader
    '        CMD_SAP_1.Connection = Cnn1



    '        CMD_SAP_1.CommandText = "
    'SELECT TOP " & n_record & " *
    'FROM
    '(
    'SELECT   t10.docnum, t10.callid, t10.itemcode, t10.itemname,t10.Cliente, t10.Cliente_F, t10.docdate,t10.DocDueDate,t10.doctotal, t10.U_Uffcompetenza,t10.U_Commento,t10.u_causcons,t10.lastname,t10.Comments,
    'case when c.N_CODICI is null then 0 else c.N_codici end as 'N_codici'
    ',case when c.OK is null then 0 else c.OK end as 'OK'
    ',case when c.Trasferibili is null then 0 else c.Trasferibili end as 'Trasferibili'
    ',case when c.CQ_TRATTAMENTO is null then 0 else c.CQ_TRATTAMENTO end as 'CQ_TRATTAMENTO'
    ',case when c.IN_APPROV_SCADUTO is null then 0 else c.IN_APPROV_SCADUTO end as 'IN_APPROV_SCADUTO'
    ',case when c.IN_APPROV_FUTURO is null then 0 else c.IN_APPROV_FUTURO end as 'IN_APPROV_FUTURO'
    ',case when c.DA_ORDINARE is null then 0 else c.DA_ORDINARE end as 'DA_ORDINARE'

    ',case when c.azione ='Spedibile' and 
    'T10.u_peson<> 0 and T10.u_pesoL<>0 AND T10.u_PRG_AZS_DIMIMB IS NOT NULL THEN 'IN_ATTESA_DOC'
    'when c.azione ='Spedibile' then 'DA_IMBALLARE' else case when c.azione is null then '' else c.azione end end as 'Azione'

    'FROM
    '(
    'select t0.docnum, t0.u_peson,T0.u_pesoL, t0.u_PRG_AZS_DIMIMB,
    't2.callid, t3.itemcode, t3.itemname,t0.cardname as 'Cliente', t4.cardname as 'Cliente_F', t0.docdate,t0.DocDueDate,t0.doctotal, t0.U_Uffcompetenza,t0.U_Commento,t0.u_causcons,t5.lastname,t0.Comments, COUNT(T6.ITEMCODE) AS 'MACCHINE'
    'from ordr t0 
    'LEFT JOIN OCLG T1 ON T1.[DocNum] = T0.[DocNum] AND   T1.[DocType] =17 AND T1.[parentId] >1
    'left join oscl t2 on t2.callid=T1.[parentId]
    'LEFT JOIN OITM T3 on T3.ITEMCODE = T2.itemcode
    'LEFT JOIN OCRD T4 ON T4.CARDCODE=T0.U_CodiceBP
    'left join ohem t5 on t5.code=t0.ownercode
    'left join rdr1 t6 on t6.docentry =t0.docentry and substring(t6.itemcode,1,1)='M'
    'where t0.DocStatus='O'
    'GROUP BY t0.docnum,t0.u_peson,T0.u_pesoL, t0.u_PRG_AZS_DIMIMB, t2.callid, t3.itemcode, t3.itemname,t0.cardname , t4.cardname , t0.docdate,t0.DocDueDate,t0.doctotal, t0.U_Uffcompetenza,t0.U_Commento,t0.u_causcons,t5.lastname,t0.Comments
    ')
    'AS T10


    'left join 


    '(

    'select T80.OC, case when T80.N_CODICI is null then 0 else t80.n_codici end as 'N_codici',
    'case when T80.OK is null then 0 else t80.ok end as 'OK',
    'case when T80.Trasferibili is null then 0 else t80.trasferibili end as 'Trasferibili',
    'case when T80.CQ_TRATTAMENTO is null then 0 else t80.CQ_TRATTAMENTO end as 'CQ_TRATTAMENTO',
    'case when T80.IN_APPROV_SCADUTO is null then 0 else t80.IN_APPROV_SCADUTO end as 'IN_APPROV_SCADUTO',
    'case when T80.IN_APPROV_FUTURO is null then 0 else t80.IN_APPROV_FUTURO end as 'IN_APPROV_FUTURO',
    'case when T80.DA_ORDINARE is null then 0 else t80.DA_ORDINARE end as 'DA_ORDINARE',
    'CASE WHEN T80.N_CODICI= T80.OK THEN 'SPEDIBILE'
    'WHEN  T80.OK+T80.Trasferibili =T80.N_CODICI THEN 'TRASFERIBILE'
    'WHEN  T80.OK+T80.Trasferibili+T80.CQ_TRATTAMENTO =T80.N_CODICI THEN 'IN_TRATTAMENTO_CQ'
    'WHEN  T80.OK+T80.Trasferibili+T80.CQ_TRATTAMENTO+T80.IN_APPROV_SCADUTO =T80.N_CODICI THEN 'IN_APPROV_SCADUTO'
    'WHEN  T80.OK+T80.Trasferibili+T80.CQ_TRATTAMENTO+T80.IN_APPROV_SCADUTO+T80.IN_APPROV_FUTURO =T80.N_CODICI THEN 'In_APPROV_FUTURO'
    'else 'DA_ORDINARE'
    'END AS 'AZIONE'
    'from
    '(
    'select t70.oc, Count(concat(t70.itemcode,t70.linenum))-sum(case when t70.azione_n =10 then 1 else 0 end ) as 'N_codici'
    ',sum(case when t70.azione_n =1 then 1 else 0 end )as 'OK'
    ',sum(case when t70.azione_n =2 then 1 else 0 end )as 'Trasferibili'
    ',sum(case when t70.azione_n =3 or t70.azione_n =4 or t70.azione_n =12 or t70.azione_n =13 then 1 else 0 end )as 'CQ_Trattamento'
    ',sum(case when t70.azione_n =5 or t70.azione_n =7 then 1 else 0 end )as 'In_approv_scaduto'
    ',sum(case when t70.azione_n =6 or t70.azione_n =8 then 1 else 0 end )as 'In_approv_futuro'
    ',sum(case when t70.azione_n =9 then 1 else 0 end )as 'Da_ordinare'

    'from
    '(
    'select case when t60.Tipo =1 then 'Materiale OC' else concat('Materiale per ODP',t61.prodname) end as 'Fonte', t60.tipo,t60.DOC, t60.oc,t60.ODP_FONTE, t60.CDS,t60.itemcode,t60.itemname, t60.whscode, t60.u_disegno,t60.ItmsGrpNam , t60.linenum, t60.openqty, t60.U_Datrasferire,t60.Giacenza,t60.CQ, t60.Clavter, t60.Ordinato,  t60.Disp,
    't60.Azione,
    't60.ODP,t60.[Cons ODP],t60.U_PRODUZIONE, t60.OA,t60.Fornitore,t60.ShipDate
    ',
    'case when  t60.Azione='OK' then 1 
    'when  t60.Azione='Trasferibile' then 2
    'when  t60.Azione='CQ' then 3
    'when  t60.Azione='Clavter' then 4
    'when  t60.Azione='06' then 12
    'when  t60.Azione='16' then 13
    'when  t60.Azione='In_approv_scaduto' then 5
    'when  t60.Azione='In_approv_futuro' then 6
    'when  t60.Azione='In_approv_scaduto/Da_Ordinare' then 7
    'when  t60.Azione='In_approv_futuro/Da_Ordinare' then 8
    'when  t60.Azione='Da_ordinare' then 9
    'when  t60.Azione='Assembl' then 10
    'when  t60.Azione='?' then 11
    'end as 'Azione_N'
    ', t61.itemcode as 'Codice_ODP', t61.prodname

    'from
    '(

    'select t50.Tipo,t50.DOC, t50.oc,t50.OC_dell_ODP,t50.odp as 'ODP_fonte', t50.CDS,t50.itemcode,t51.itemname, t50.whscode, t51.u_disegno,t52.ItmsGrpNam , t50.linenum, t50.openqty, t50.U_Datrasferire,t50.Giacenza,t50.CQ, t50.Clavter,T50.[06],T50.[16], t50.Ordinato,  t50.Disp,
    'case when t50.azione='In_approv' and a.u_produzione='ASSEMBL' then 'Assembl'

    'when t50.azione='In_approv' and case when a.DueDate is null then b.ShipDate when b.ShipDate is null then a.DueDate when  a.DueDate <=b.ShipDate then a.DueDate else b.ShipDate end <= getdate() then 'In_approv_scaduto'
    'when t50.azione='In_approv' and case when a.DueDate is null then b.ShipDate when b.ShipDate is null then a.DueDate when  a.DueDate <=b.ShipDate then a.DueDate else b.ShipDate end > getdate() then 'In_approv_futuro'
    'when t50.azione='In_approv/Da_Ordinare' and case when a.DueDate is null then b.ShipDate when b.ShipDate is null then a.DueDate when  a.DueDate <=b.ShipDate then a.DueDate else b.ShipDate end <= getdate() then 'In_approv_scaduto/Da_Ordinare'
    'when t50.azione='In_approv/Da_Ordinare' and case when a.DueDate is null then b.ShipDate when b.ShipDate is null then a.DueDate when  a.DueDate <=b.ShipDate then a.DueDate else b.ShipDate end > getdate() then 'In_approv_futuro/Da_Ordinare'
    'else t50.azione end as 'Azione',

    'A.Docnum as 'ODP',a.DueDate as 'Cons ODP',a.U_PRODUZIONE, b.Docnum as 'OA',b.CardName as 'Fornitore',b.ShipDate


    'from
    '(
    'select t40.Tipo,t40.DOC, t40.oc,t40.OC_dell_ODP,t40.odp, t40.CDS,t40.itemcode,t40.linenum, t40.openqty, t40.U_Datrasferire,t40.Giacenza,t40.CQ, t40.Clavter,T40.[06],T40.[16], t40.Ordinato,  t40.Disp,
    'case when t40.U_Datrasferire =0 or t40.U_Datrasferire is null then 'OK'
    'when t40.Giacenza>=t40.U_Datrasferire then 'Trasferibile'
    'when t40.Giacenza+t40.cq>=t40.U_Datrasferire then 'CQ'
    'when t40.Giacenza+t40.clavter>=t40.U_Datrasferire then 'Clavter'
    'when t40.Giacenza+t40.[06]>=t40.U_Datrasferire then '06'
    'when t40.Giacenza+t40.[16]>=t40.U_Datrasferire then '16'
    'when t40.Giacenza+t40.ordinato>= t40.U_Datrasferire and t40.Disp>= 0 then 'In_approv'
    'when t40.Giacenza+t40.ordinato>= t40.U_Datrasferire and t40.Disp< 0 then 'In_approv/Da_Ordinare'
    'when t40.Giacenza+t40.ordinato<=t40.U_Datrasferire and t40.Disp< 0 then 'Da_ordinare'
    'else '?'
    'end as 'Azione', t40.whscode
    'from
    '(
    'select t30.Tipo,t30.DOC, t30.oc,t30.OC_dell_ODP,t30.odp, t30.CDS,t30.itemcode,t30.linenum, t30.openqty, t30.U_Datrasferire,t30.Giacenza,case when t32.onhand is null then 0 else t32.onhand end as 'CQ', case when t33.onhand is null then 0 else t33.onhand end as'Clavter', case when t34.onhand is null then 0 else t34.onhand end as'06', case when t35.onhand is null then 0 else t35.onhand end as'16', sum(case when t31.onorder is null then 0 else t31.onorder end) as 'Ordinato',  t30.Disp, t30.whscode

    'from
    '(
    'select 
    't20.Tipo,t20.DOC, t20.oc,t20.OC_dell_ODP,t20.odp, t20.CDS,t20.itemcode,t20.linenum, t20.openqty, t20.U_Datrasferire,t20.Giacenza, sum(case when t21.onhand is null then 0 else t21.onhand end -case when t21.iscommited is null then 0 else t21.iscommited end + case when t21.onorder is null then 0 else t21.onorder end) as 'Disp', t20.whscode

    'from
    '(
    'select 
    't10.Tipo,t10.DOC, t10.OC,t10.OC_dell_ODP,t10.odp, t10.CDS,t10.itemcode,t10.linenum, t10.openqty, t10.U_Datrasferire,sum(t11.onhand) as 'Giacenza', t10.whscode

    'from
    '(

    'select 1 as 'Tipo','OC' as 'DOC', t0.docnum as 'OC','' as 'ODP','' as 'OC_dell_ODP', cast(t3.callid as varchar) as 'CDS',t1.itemcode,t1.linenum, t1.openqty, t1.U_Datrasferire, t1.whscode

    'from ordr t0 inner join rdr1 t1 on t0.docentry=t1.docentry
    'left join oclg t2 on t2.DocNum=t0.docnum and t2.DocType =17 AND T2.[parentId] >1
    'left join oscl t3 on t3.callid=T2.[parentId]

    'WHERE T0.DocStatus = N'o'  and (substring(t1.itemcode,1,1)='0' or substring(t1.itemcode,1,1)='C' or substring(t1.itemcode,1,1)='D')  

    'union all

    'select 2,'ODP',T0.ORIGINNUM as 'OC', t0.docnum as 'ODP','' as 'OC_dell_ODP',cast(substring(t0.u_prg_azs_commessa,4,99) as varchar), t1.itemcode, t1.linenum,t1.PlannedQty, t1.U_PRG_WIP_QtaDaTrasf, t1.wareHouse
    'from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
    'where (t0.Status= 'p' or t0.Status= 'r') and T0.ORIGINNUM<>0


    ')
    'as t10 left join oitw t11 on t10.itemcode=t11.itemcode and t11.whscode<>'WIP' and t11.whscode<>'CQ' and t11.whscode<>'Clavter' and t11.whscode<>'06' and t11.whscode<>'16'
    'group by t10.Tipo,t10.DOC, t10.oc,t10.OC_dell_ODP, t10.CDS,t10.itemcode,t10.linenum, t10.openqty, t10.U_Datrasferire,t10.odp, t10.whscode
    ')
    'as t20 left join oitw t21 on t20.itemcode=t21.itemcode

    'group by t20.Tipo,t20.DOC, t20.oc,t20.OC_dell_ODP, t20.CDS,t20.itemcode,t20.linenum, t20.openqty, t20.U_Datrasferire,t20.Giacenza,t20.odp, t20.whscode
    ')
    'as t30 left join oitw t31 on t31.itemcode=t30.itemcode
    'left join oitw t32 on t32.itemcode=t30.itemcode and t32.whscode='CQ'
    'left join oitw t33 on t33.itemcode=t30.itemcode and t33.whscode='Clavter'
    'left join oitw t34 on t34.itemcode=t30.itemcode and t34.whscode='06'
    'left join oitw t35 on t35.itemcode=t30.itemcode and t35.whscode='16'
    'group by t30.Tipo,t30.DOC, t30.oc,t30.OC_dell_ODP, t30.CDS,t30.itemcode,t30.linenum, t30.openqty, t30.U_Datrasferire,t30.Giacenza,  t30.Disp, t32.onhand, t33.onhand,t30.odp, t30.whscode, T34.ONHAND, T35.ONHAND
    ')
    'as t40
    ')
    'as t50

    'left join 
    '(
    'select t10.itemcode, min(t11.docnum) as 'Docnum', t11.duedate, t11.U_PRODUZIONE
    'from 
    '(
    'select t0.itemcode,  min(t0.DueDate) as 'Min_data_odp'
    'from owor t0
    'where (t0.status='P' or t0.status='R')
    'group by t0.itemcode
    ')
    'as t10 left join owor t11 on t11.itemcode=t10.itemcode and (t11.status='P' or t11.status='R') and t10.Min_data_odp=t11.duedate
    'group by t11.duedate,t10.itemcode, t11.U_PRODUZIONE
    ')
    'A on a.itemcode=t50.itemcode and t50.azione   Like '%%IN_APPROV%%' 

    'left join 
    '(
    'select t10.itemcode, min(t12.docnum) as 'Docnum', t11.shipdate, t12.cardname
    'from 
    '(
    'select t0.itemcode,  min(t0.shipDate) as 'Min_data_Oa'
    'from por1 t0
    'where (t0.OpenQty>0)
    'group by t0.itemcode
    ')
    'as t10 left join por1 t11 on t11.itemcode=t10.itemcode and t11.OpenQty>0 and t10.Min_data_oa=t11.shipdate
    'left join opor t12 on t12.docentry=t11.docentry
    'group by t11.shipdate,t10.itemcode,t12.cardname
    ')
    'B on b.itemcode=t50.itemcode and t50.azione   Like '%%IN_APPROV%%'
    'LEFT JOIN OITM t51 on t51.itemcode=t50.itemcode
    'left join oitb t52 on t52.ItmsGrpCod=t51.ItmsGrpCod


    ')
    'as t60 LEFT JOIN OWOR T61 ON T61.DOCNUM=T60.ODP_FONTE AND T60.ODP_FONTE<>0
    ')
    'as t70
    'group by t70.oc

    ')
    'as t80
    ')
    'C on C.oc=t10.DocNum

    'WHERE T10.MACCHINE=0
    ')
    'AS
    'T20

    'WHERE 0=0 " & Filtro_docnum & Filtro_cds & Filtro_cliente & Filtro_cliente_f & Filtro_commento & Filtro_causcons & Filtro_riferimento & Filtro_azione & "

    'ORDER BY T20.DOCDUEDATE"


    '        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
    '        Do While cmd_SAP_reader_1.Read()

    '            DataGridView_CDS.Rows.Add(cmd_SAP_reader_1("DocNum"), cmd_SAP_reader_1("Callid"), cmd_SAP_reader_1("itemcode"), cmd_SAP_reader_1("itemname"), cmd_SAP_reader_1("Cliente"), cmd_SAP_reader_1("Cliente_F"), cmd_SAP_reader_1("Docdate"), cmd_SAP_reader_1("Docduedate"), cmd_SAP_reader_1("doctotal"), cmd_SAP_reader_1("U_uffcompetenza"), cmd_SAP_reader_1("U_Commento"), cmd_SAP_reader_1("u_causcons"), cmd_SAP_reader_1("lastname"), cmd_SAP_reader_1("COMMENTS"), cmd_SAP_reader_1("azione"), cmd_SAP_reader_1("N_codici"), cmd_SAP_reader_1("OK"), cmd_SAP_reader_1("Trasferibili"), cmd_SAP_reader_1("Cq_trattamento"), cmd_SAP_reader_1("in_approv_scaduto"), cmd_SAP_reader_1("in_approv_futuro"), cmd_SAP_reader_1("da_ordinare"))
    '        Loop

    '        Cnn1.Close()
    '        DataGridView_CDS.ClearSelection()
    '    End Sub





    Sub filtra_commessa()
        Dim i = 0


        Do While i < DataGridView_commesse.RowCount

            Try



                If UCase(DataGridView_commesse.Rows(i).Cells(columnName:="Commessa_Tab").Value).Contains(UCase(TextBox_commessa.Text)) Then
                    DataGridView_commesse.Rows(i).Visible = True

                    If UCase(DataGridView_commesse.Rows(i).Cells(columnName:="Descrizione").Value).Contains(UCase(TextBox1.Text)) Then
                        DataGridView_commesse.Rows(i).Visible = True









                    Else
                        DataGridView_commesse.Rows(i).Visible = False

                    End If

                Else
                    DataGridView_commesse.Rows(i).Visible = False
                End If
            Catch ex As Exception
                DataGridView_commesse.Rows(i).Visible = False
            End Try
            i = i + 1
        Loop
    End Sub

    'Sub filtra_OC()
    '    Dim parola0 As String
    '    Dim parola1 As String
    '    Dim parola4 As String
    '    Dim parola5 As String
    '    Dim parola6 As String
    '    Dim i = 0
    '    Do While i < DataGridView_CDS.RowCount

    '        parola0 = DataGridView_CDS.Rows(i).Cells(0).Value
    '        parola1 = DataGridView_CDS.Rows(i).Cells(1).Value
    '        parola4 = UCase(DataGridView_CDS.Rows(i).Cells(4).Value)
    '        parola5 = UCase(DataGridView_CDS.Rows(i).Cells(5).Value)
    '        parola6 = UCase(DataGridView_CDS.Rows(i).Cells(columnName:="Commento").Value)


    '        Try

    '            If parola0.Contains(TextBox_OC.Text) Then
    '                DataGridView_CDS.Rows(i).Visible = True

    '                If parola1.Contains(TextBox_CDS.Text) Then
    '                    DataGridView_CDS.Rows(i).Visible = True


    '                    If parola4.Contains(UCase(TextBox_cliente.Text)) Then
    '                        DataGridView_CDS.Rows(i).Visible = True



    '                        If UCase(DataGridView_CDS.Rows(i).Cells(columnName:="PM").Value).Contains(UCase(TextBox4.Text)) Then
    '                                DataGridView_CDS.Rows(i).Visible = True

    '                                If UCase(DataGridView_CDS.Rows(i).Cells(columnName:="Stato").Value).Contains(UCase(TextBox5.Text)) Then
    '                                    DataGridView_CDS.Rows(i).Visible = True

    '                                    If UCase(DataGridView_CDS.Rows(i).Cells(columnName:="Causale").Value).Contains(UCase(TextBox6.Text)) Then
    '                                        DataGridView_CDS.Rows(i).Visible = True

    '                                        If UCase(DataGridView_CDS.Rows(i).Cells(columnName:="Commento").Value).Contains(UCase(TextBox7.Text)) Then
    '                                            DataGridView_CDS.Rows(i).Visible = True
    '                                        Else
    '                                            DataGridView_CDS.Rows(i).Visible = False
    '                                        End If


    '                                    Else
    '                                        DataGridView_CDS.Rows(i).Visible = False
    '                                    End If

    '                                Else
    '                                    DataGridView_CDS.Rows(i).Visible = False
    '                                End If

    '                            Else
    '                                DataGridView_CDS.Rows(i).Visible = False

    '                            End If



    '                    Else
    '                        DataGridView_CDS.Rows(i).Visible = False

    '                    End If


    '                Else
    '                    DataGridView_CDS.Rows(i).Visible = False

    '                End If

    '            Else
    '                DataGridView_CDS.Rows(i).Visible = False

    '            End If

    '        Catch ex As Exception
    '            DataGridView_CDS.Rows(i).Visible = False
    '        End Try
    '        i = i + 1
    '    Loop

    'End Sub



    'Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs)
    '    If TextBox_OC.Text = "" Then
    '        Filtro_docnum = ""
    '    Else
    '        Filtro_docnum = "And t20.docnum   Like '%%" & TextBox_OC.Text & "%%'  "
    '    End If
    'End Sub
    'Private Sub TextBox_ocKeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        riempi_CDS()

    '    End If
    'End Sub


    'Private Sub TextBox_CDS_TextChanged(sender As Object, e As EventArgs)
    '    If TextBox_CDS.Text = "" Then
    '        Filtro_cds = ""
    '    Else
    '        Filtro_cds = " And t20.callid  Like '%%" & TextBox_CDS.Text & "%%'  "
    '    End If
    'End Sub

    'Private Sub TextBox_cliente_TextChanged(sender As Object, e As EventArgs)
    '    If TextBox_cliente.Text = "" Then
    '        Filtro_cliente = ""
    '    Else
    '        Filtro_cliente = " And (t20.cliente  Like '%%" & TextBox_cliente.Text & "%%' or t20.cliente_f  Like '%%" & TextBox_cliente.Text & "%%')  "
    '    End If
    'End Sub



    Private Sub Button1_Click_1(sender As Object, e As EventArgs)
        Commesse_odp_cds_aperte()
    End Sub



    Private Sub TabPage1_Click(sender As Object, e As EventArgs)
        DataGridView_commesse.Show()
        GroupBox_COMMESSE.Show()
        'TableLayoutPanel4.Show()

    End Sub

    'Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
    '    If TextBox1.Text = "" Then
    '        Filtro_descrizione_datagrid = ""
    '    Else

    '        Filtro_descrizione_datagrid = " And t10.descrizione Like '%%" & TextBox1.Text & "%%' "
    '    End If
    '    Commesse_odp_aperte()
    'End Sub



    'Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged


    '    If TextBox2.Text = "" Then
    '        Filtro_cliente_datagrid = ""
    '    Else

    '        Filtro_cliente_datagrid = "  And (t10.cliente Like '%%" & TextBox2.Text & "%%'  or t10.cliente_finale Like '%%" & TextBox2.Text & "%%') "
    '    End If
    '    Commesse_odp_aperte()

    'End Sub



    'Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs)
    '    If TextBox4.Text = "" Then
    '        Filtro_riferimento = ""
    '    Else
    '        Filtro_riferimento = " And t20.lastname  Like '%%" & TextBox4.Text & "%%'  "
    '    End If
    'End Sub

    'Private Sub TextBox5_TextChanged_1(sender As Object, e As EventArgs)
    '    If TextBox5.Text = "" Then
    '        Filtro_azione = ""
    '    Else
    '        Filtro_azione = " And t20.azione  Like '%%" & TextBox5.Text & "%%'  "
    '    End If
    'End Sub

    'Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs)
    '    If TextBox6.Text = "" Then
    '        Filtro_causcons = ""
    '    Else
    '        Filtro_causcons = " And t20.u_causcons  Like '%%" & TextBox6.Text & "%%'  "
    '    End If
    'End Sub



    'Private Sub DataGridView_CDS_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)

    '    If DataGridView_CDS.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "IN_ATTESA_DOC" Then

    '        DataGridView_CDS.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
    '    ElseIf DataGridView_CDS.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "DA_IMBALLARE" Then

    '        DataGridView_CDS.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.YellowGreen

    '    ElseIf DataGridView_CDS.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "TRASFERIBILE" Then

    '        DataGridView_CDS.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Yellow

    '    End If




    '    If DataGridView_CDS.Rows(e.RowIndex).Cells(columnName:="Causale").Value = "V" Then

    '        DataGridView_CDS.Rows(e.RowIndex).Cells(columnName:="Causale").Style.BackColor = Color.Green


    '    Else
    '        DataGridView_CDS.Rows(e.RowIndex).Cells(columnName:="Causale").Style.BackColor = Color.Red

    '    End If




    ' End Sub

    Sub KPI_ordine()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1

        CMD_SAP_1.CommandText = " Select t110.azione, count(t110.articolo) as 'Articoli'
from
(
Select  t100.linenum, T100.[WhsCode], T100.Articolo, t100.[Desc articolo] , t100.Disegno, T100.[ItmsGrpNam], t100.Quantita,t100.Trasferito, t100.[Da trasferire], 
case when t100.[Da trasferire]=0 then 'OK' when (t100.giacenza>=t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 )  then 'Trasferibile/Da ordinare' when (t100.giacenza>=t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)>=0 )  then 'Trasferibile' when t100.[Da trasferire]=0 then 'OK' when (t100.[Da trasferire]>0 and sum(case when t106.onhand is null then 0 else t106.onhand end - case when T106.[IsCommited] is null then 0 else  T106.[IsCommited] end +case when t106.onorder is null then 0 else t106.onorder end)>=0 and t100.giacenza<t100.[Da trasferire]) then 'IN APPROV'   when sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 then 'Da ordinare' end as 'Azione', case when t100.[Da trasferire]=0 then '' else t102.docnum end as 'ODP', cast(case when t100.[Da trasferire]=0 then '' else cast(T102.[DueDate] as varchar) end as VARCHAR)as 'Cons ODP' , case when t100.[Da trasferire]=0 then '' else t102.U_PRG_AZS_commessa end as 'Commessa' ,case when t100.[Da trasferire]=0 then '' else  t102.U_produzione end as 'Reparto', case when t100.[Da trasferire]=0 then '' else t107.docnum  end  as 'OA', 
case when t100.[Da trasferire]=0 then '' else t107.cardname end as 'Fornitore', cast(case when t100.[Da trasferire]=0 then '' else cast(t103.[ShipDate] as varchar)  end as varchar) as 'Cons OA'

from
(
SELECT 'OC' AS 'Documento',t1.linenum, T1.[WhsCode], T9.[ITEMCODE] as 'Articolo', t9.itemname as 'Desc articolo' , case when t9.u_disegno is null then '' else t9.u_disegno end as 'Disegno', T11.[ItmsGrpNam], t1.openqty as 'Quantita',case when t1.U_trasferito is null then 0 else t1.U_trasferito end as 'Trasferito', t1.u_DATRASFERIRE as 'Da trasferire', sum (t20.onhand) as 'giacenza', t1.docentry

from rdr1 t1 inner join ordr t0 on t0.docentry=t1.docentry
inner join oitm t9 on t9.itemcode=t1.itemcode
LEFT JOIN OITB T11 ON T9.[ItmsGrpCod] = T11.[ItmsGrpCod]
inner join oitw t20 on t20.itemcode=t1.itemcode

WHERE T0.[docnum]='""' and t1.itemtype=4 and (substring(T9.[ITEMCODE],1,1)='0' or substring(T9.[ITEMCODE],1,1)='C' or substring(T9.[ITEMCODE],1,1)='D') and (t20.whscode='01' or t20.whscode='03' or t20.whscode='SCA' or t20.whscode='FERRETTO' or t20.whscode='MUT') 

group by 
T0.DOCNUM,  t1.linenum, T1.[WhsCode], T9.[ITEMCODE] , t9.itemname  , t9.u_disegno , T11.[ItmsGrpNam], t1.openqty, t1.U_trasferito , t1.u_DATRASFERIRE, t1.docentry
)
as t100 
left join owor t102 on t100.articolo=t102.itemcode and (T102.Status ='P' or T102.Status ='R' )
left join por1 t103 on t103.itemcode=t100.articolo and t103.opencreqty >0
LEFT OUTER JOIN ITT1 T104 on t100.articolo = T104.Father
left join oitw t105 on t105.itemcode=t104.code and t105.[WhsCode]='01'
left join oitw t106 on t106.itemcode=t100.articolo
left join opor t107 on t107.docentry=t103.docentry

group by
 T100.[articolo], t100.trasferito, T100.[DESC articolo], t100.linenum, T100.[WhsCode], t100.quantita,  t100.disegno, T100.[ItmsGrpNam], t100.giacenza,t100.[da trasferire], t102.docnum, T102.[DueDate],t102.U_PRG_AZS_commessa,t102.U_produzione,t107.docnum,t107.cardname,t103.[ShipDate]
)
as t110
group by t110.azione"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()



        Loop
        Cnn1.Close()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()

    End Sub


    Private Sub Button1_Click_2(sender As Object, e As EventArgs) Handles Button1.Click
        Button4.Visible = True
        Button1.Visible = False
        CONDIZIONI_COMMESSE = "  substring(T0.[U_PRG_AZS_Commessa],1,2)<>'M0' and (T0.[Status] ='P' or  T0.[Status] ='R') "
        Commesse_odp_aperte(DataGridView_commesse, TextBox_commessa.Text, TextBox1.Text, TextBox2.Text, CheckBox1.Checked)
    End Sub



    Private Sub Cmd_Premontaggio_Click(sender As Object, e As EventArgs) Handles Cmd_Premontaggio.Click
        Form_Premontaggio.Owner = Me
        Form_Premontaggio.Show()
        'SCHEDA_COMMESSA()
        'FORM6.Button_commessa.Text = pianificazione.commessa
        'FORM6.Label_descrizione.Text = DESCRIZIONE_commessa
        'FORM6.Label_ordine_cliente.Text = Ordine_cliente_commessa
        'FORM6.Label_cliente.Text = Cliente_commessa
        'FORM6.Label_cliente_finale.Text = Cliente_finale_commessa
        'FORM6.Label_consegna.Text = Consegna_commessa
        Form_Premontaggio.completamento_gruppi_preassemblaggio_assemblaggio()
        elenco_ODP_Premontaggio()
        'news_materiale()
        'FORM6.Inserimento_responsabile_montaggio()
        'FORM6.Inserimento_responsabile_collaudo()
        'FORM6.leggi_chi_è_responsabile()

        'FORM6.riga = Nothing

    End Sub

    Sub elenco_ODP_Premontaggio()
        Dim Data_Ordine As Date
        Form_Premontaggio.DataGridView_ODP.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1

        CMD_SAP_1.CommandText = " SELECT T40.StartDate, T40.U_PRG_AZS_Commessa,t40.[N ODP] , t40.[Stato ODP], t40.Codice , t40.Descrizione , t40.Disegno , t40.quantita , t40.stato ,T40.[Codice fase], t40.fase , t40.N ,t40.Trasferiti ,  T40.PREM, T40.MONT ,T40.[ASS EL], T40.NEWS,t40.tipo,case when MAX(t40.Id_Ticket) is null then '' else MAX(t40.Id_Ticket) end AS 'ID_TICKET', case when t42.descrizione is null then '' else t42.descrizione end as 'Reparto_ticket'
FROM
(
	SELECT T30.StartDate, T30.U_PRG_AZS_Commessa,t30.[N ODP] , t30.[Stato ODP], t30.Codice , t30.Descrizione , t30.Disegno , t30.quantita , t30.stato ,T30.[Codice fase], t30.fase , t30.N ,t30.Trasferiti ,  T30.PREM, T30.MONT ,T30.[ASS EL], T30.NEWS,t30.tipo,MAX(t32.Id_Ticket) AS 'ID_TICKET'
	FROM
	(
		SELECT T20.StartDate, T20.U_PRG_AZS_Commessa,t20.[N ODP] , t20.[Stato ODP], t20.Codice , t20.Descrizione , t20.Disegno , t20.quantita , t20.stato ,T20.[Codice fase], t20.fase , t20.N ,t20.Trasferiti ,  case when T20.PREM is null then 0 else t20.prem end as 'PREM' , case when T20.MONT is null then 0 else t20.mont end as 'MONT' ,T20.[ASS EL], SUM(CASE WHEN T22.DOCDATE>GETDATE()-2 THEN 1 ELSE 0 END)  AS 'NEWS',t20.tipo
		FROM
		(

			Select T10.StartDate, T10.U_PRG_AZS_Commessa,T10.DOCENTRY,t10.[N ODP] as 'N ODP', t10.[Stato ODP], t10.Codice as 'Codice', t10.Descrizione as 'Descrizione', t10.Disegno as 'Disegno', t10.quantita as 'Quantita', t10.stato as 'Stato', t10.fase as 'Fase', T10.[Codice fase], t10.N as 'N',t10.Trasferiti as 'Trasferiti',  sum(CASE WHEN T10.[Codice fase] ='P01501' THEN t10.quantita* case when T11.[Code]='R00568' OR T11.CODE='R00525' then t11.quantity else 0 end END)   as 'PREM' , sum(CASE WHEN T10.[Codice fase] ='P02001' THEN t10.quantita* case when T11.[Code]='R00568' OR T11.CODE='R00525' then t11.quantity else 0 end END)   as 'MONT' ,sum(case when t11.code='R00530' then t11.quantity else 0 end) as 'ASS EL', t10.tipo
			from
			(
				SELECT T0.StartDate,T0.U_PRG_AZS_Commessa,T0.DOCENTRY,T0.[DocNum] as 'N ODP', t0.status as 'Stato ODP', T0.[ItemCode] as 'Codice', T1.itemname as 'Descrizione', case when T1.[U_Disegno] is null then '' else T1.[U_Disegno] end as 'Disegno', T0.[PlannedQty] as 'Quantita',T0.U_stato as 'Stato', T0.[U_Fase] as 'Codice fase', T2.[Name] as 'Fase', sum(CASE WHEN T3.ITEMTYPE='4' then 1 else 0 end ) as 'N', sum(case when t3.U_prg_wip_qtadatrasf=0 and T3.ITEMTYPE='4'  then 1 else 0 end) as 'trasferiti', SUBSTRING(T0.u_PRODUZIONE,1,1) AS 'Tipo'
				FROM OWOR T0 inner join OITM T1 on t0.itemcode=t1.itemcode
				left JOIN [dbo].[@FASE]  T2 ON T0.[U_Fase] = T2.[Code] 
				left join wor1 t3 on t3.docentry=t0.docentry
				LEFT JOIN OWOR T10 ON T10.ITEMCODE=T3.ITEMCODE AND (T10.STATUS='P' OR T10.STATUS='R') and (T10.[U_PRODUZIONE]='ASSEMBL')
				WHERE (T0.U_PRG_AZS_Commessa LIKE 'M0%' OR T0.U_PRG_AZS_Commessa LIKE 'MAGAZ%' OR T0.U_PRG_AZS_Commessa LIKE 'SCORT%') AND NOT T0.[ItemCode] LIKE 'M%' AND NOT T0.[ItemCode] LIKE 'F%' AND t0.U_stato<>'Completato' and (t0.status='P' or t0.status='R') and (T0.[U_PRODUZIONE]='ASSEMBL' or T0.[U_PRODUZIONE]='EST') and t3.itemtype=4 and substring(t3.itemcode,1,1)<>'L' AND T10.DOCNUM IS NULL
				group by
				T0.StartDate,T0.U_PRG_AZS_Commessa, T0.DOCENTRY, T0.[DocNum] , T0.[ItemCode], T1.[U_Disegno] , T2.[Name], t1.itemname, T0.[PlannedQty], T0.[U_Fase], T0.U_stato,t0.status, t0.U_produzione
			)
			as t10
			left join itt1 t11 on t11.father=t10.codice
			group by T10.StartDate,T10.U_PRG_AZS_Commessa,T10.DOCENTRY, t10.[N ODP], t10.[Stato ODP], t10.Codice, t10.Descrizione, t10.Disegno, t10.quantita, t10.fase, t10.N,t10.Trasferiti, t10.[Codice Fase],t10.stato ,t10.tipo
		)
		AS T20
		LEFT JOIN WTR1 T21 ON T21.U_PRG_AZS_OPDOCENTRY=T20.DOCENTRY AND T21.WHSCODE='WIP'
		LEFT JOIN OWTR T22 ON T22.DOCENTRY=T21.DOCENTRY
		GROUP BY T20.StartDate, T20.U_PRG_AZS_Commessa,t20.[N ODP] , t20.[Stato ODP], t20.Codice , t20.Descrizione , t20.Disegno , t20.quantita , t20.stato , t20.fase , t20.N ,t20.Trasferiti, T20.PREM,T20.MONT, T20.[ASS EL],T20.[Codice Fase],t20.tipo
	)
	AS T30 LEFT JOIN [TIRELLI_40].[DBO].COLL_RIFERIMENTI T31 ON T31.Tipo_Codice='Ordine' and t31.Codice_SAP=t30.[N ODP]
	left join [TIRELLI_40].[DBO].coll_tickets t32 on t32.Id_Ticket=t31.Rif_Ticket and t32.aperto=1
	group by T30.StartDate, T30.U_PRG_AZS_Commessa,t30.[N ODP] , t30.[Stato ODP], t30.Codice , t30.Descrizione , t30.Disegno , t30.quantita , t30.stato ,T30.[Codice fase], t30.fase , t30.N ,t30.Trasferiti ,  T30.PREM, T30.MONT ,T30.[ASS EL], T30.NEWS,t30.tipo
)
AS T40 left join [TIRELLI_40].[DBO].coll_tickets t41 on t41.Id_Ticket=t40.Id_Ticket
left join [TIRELLI_40].[DBO].COLL_Reparti t42 on t42.Id_Reparto=t41.Destinatario
group by T40.StartDate, T40.U_PRG_AZS_Commessa,t40.[N ODP] , t40.[Stato ODP], t40.Codice , t40.Descrizione , t40.Disegno , t40.quantita , t40.stato ,T40.[Codice fase], t40.fase , t40.N ,t40.Trasferiti ,  T40.PREM, T40.MONT ,T40.[ASS EL], T40.NEWS,t40.tipo,t42.descrizione
order by T40.StartDate, T40.U_PRG_AZS_Commessa,T40.TIPO, t40.[Codice Fase],  t40.[N ODP]
"





        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()
            Data_Ordine = cmd_SAP_reader_1("StartDate")
            Form_Premontaggio.DataGridView_ODP.Rows.Add(Data_Ordine.ToString("dd/MM/yyyy"), cmd_SAP_reader_1("U_PRG_AZS_Commessa"), cmd_SAP_reader_1("N ODP"), cmd_SAP_reader_1("Codice"), cmd_SAP_reader_1("Descrizione"), cmd_SAP_reader_1("Disegno"), "", Math.Round(cmd_SAP_reader_1("Quantita")), cmd_SAP_reader_1("Stato"), cmd_SAP_reader_1("Fase"), cmd_SAP_reader_1("trasferiti") / cmd_SAP_reader_1("N") * 100, cmd_SAP_reader_1("PREM") + cmd_SAP_reader_1("MONT"), cmd_SAP_reader_1("ASS EL"), cmd_SAP_reader_1("TIPO"), cmd_SAP_reader_1("NEWS"), cmd_SAP_reader_1("id_Ticket"), cmd_SAP_reader_1("Reparto_ticket"))

        Loop
        Cnn1.Close()
        Form_Premontaggio.DataGridView_ODP.ClearSelection()
    End Sub






    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Button4.Visible = False
        Button1.Visible = True
        CONDIZIONI_COMMESSE = " substring(T0.[U_PRG_AZS_Commessa],1,1)='M' and (T0.[Status] ='P' or  T0.[Status] ='R' OR T0.[CloseDate]>=getdate()-60)"
        Commesse_odp_aperte(DataGridView_commesse, TextBox_commessa.Text, TextBox1.Text, TextBox2.Text, CheckBox1.Checked)
    End Sub

    Private Sub DataGridView_commesse_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridView_CDS_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click


        FORM6.Show()
        'SCHEDA_COMMESSA(Pianificazione.commessa)
        FORM6.inizializza_form(Pianificazione.commessa)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        SCHEDA_COMMESSA(Pianificazione.commessa)

        Form_Scheda_Collaudi.inizializzazione_form(Pianificazione.commessa)
        Form_Scheda_Collaudi.Show()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        If Pianificazione.commessa >= "M04000" Then
            Scheda_tecnica.Show()
            Scheda_tecnica.BringToFront()
            Scheda_tecnica.inizializza_scheda_tecnica(Pianificazione.commessa)

        Else

            Scheda_commessa_documentazione.inizializzazione = 0
            Scheda_commessa_documentazione.carico_iniziale = 0

            Scheda_commessa_documentazione.Azzera_campi()

            Scheda_commessa_documentazione.commessa = Pianificazione.commessa


            'Scheda_commessa_documentazione.codice_bp_campione = DataGridView.Rows(e.RowIndex).Cells(columnName:="Codice_cliente").Value





            Scheda_commessa_Pianificazione.layout_scheda_tecnica()


            Scheda_commessa_documentazione.compila_anagrafica(Pianificazione.commessa)


            Scheda_commessa_documentazione.Inserimento_dipendenti()

            Scheda_commessa_documentazione.COMPILA_RECORD_INIZIALI()

            '  Scheda_commessa_documentazione.Rischio_effettivo()
            Scheda_commessa_documentazione.Ultimo_aggiornamento()
            Scheda_commessa_documentazione.riempi_datagridview_combinazioni()
            Scheda_commessa_documentazione.riempi_datagridview_campioni()
            Scheda_commessa_documentazione.cerca_file()
            Scheda_commessa_documentazione.Show()



            Scheda_commessa_documentazione.carico_iniziale = 1
            Scheda_commessa_documentazione.inizializzazione = 1
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Mostra.Hide()
        Homepage.mostra_dashboard()
        Mostra.Owner = Me
        Mostra.Show()
    End Sub

    'Private Sub tabpage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter
    '    GroupBox_COMMESSE.Hide()
    '    TableLayoutPanel4.Show()
    '    riempi_CDS()


    'End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        ODP_Tree.Show()
        ODP_Tree.Lbl_Matricola.Text = Pianificazione.commessa
        ODP_Tree.Compila_Albero(Pianificazione.commessa, "ODP_TREE", ODP_Tree.CheckBox3.Checked)
    End Sub

    'Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs)
    '    If TextBox7.Text = "" Then
    '        Filtro_commento = ""
    '    Else
    '        Filtro_commento = "And t20.u_commento  Like '%%" & TextBox7.Text & "%%'  "
    '    End If
    'End Sub



    'Private Sub TextBox_CDS_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        riempi_CDS()

    '    End If
    'End Sub

    'Private Sub TextBox_cliente_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        riempi_CDS()

    '    End If
    'End Sub

    'Private Sub TextBox_Cliente_f_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        riempi_CDS()

    '    End If
    'End Sub

    'Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        riempi_CDS()

    '    End If
    'End Sub

    'Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        riempi_CDS()

    '    End If
    'End Sub

    'Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        riempi_CDS()

    '    End If
    'End Sub

    'Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs)
    '    If e.KeyCode = Keys.Enter Then
    '        riempi_CDS()

    '    End If
    'End Sub

    'Private Sub Button10_Click(sender As Object, e As EventArgs)
    '    n_record = TextBox8.Text
    '    riempi_CDS()
    'End Sub

    Private Sub TextBox_commessa_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_commessa.KeyDown
        If e.KeyCode = Keys.Enter Then
            ApplicaFiltro(TextBox_commessa.Text, TextBox1.Text, TextBox2.Text, CheckBox1.Checked)
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            ApplicaFiltro(TextBox_commessa.Text, TextBox1.Text, TextBox2.Text, CheckBox1.Checked)
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            ApplicaFiltro(TextBox_commessa.Text, TextBox1.Text, TextBox2.Text, CheckBox1.Checked)
        End If
    End Sub

    Private Sub ApplicaFiltro(par_commessa As String, par_descrizione As String, parcliente As String, par_solo_M As Boolean)


        Commesse_odp_aperte(DataGridView_commesse, par_commessa, par_descrizione, parcliente, par_solo_M)
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub DataGridView_commesse_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)

    End Sub

    Public Class Dettaglicommessa
        Public Descrizione_commessa As String
        Public ordine_cliente_commessa As String
        Public Cliente_commessa As String
        Public Cliente_finale_commessa As String
        Public Consegna_commessa As Date
        Public Giorni_alla_consegna As Integer
        Public codice_cliente As String
        Public codice_cliente_finale As String
        Public destinazione As String
        Public riempimento As String



    End Class

    Private Sub TextBox_commessa_TextChanged(sender As Object, e As EventArgs) Handles TextBox_commessa.TextChanged

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        ApplicaFiltro(TextBox_commessa.Text, TextBox1.Text, TextBox2.Text, CheckBox1.Checked)
    End Sub

    Private Sub DataGridView_commesse_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_commesse.CellContentClick

    End Sub

    Private Sub DataGridView_commesse_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_commesse.CellClick
        If e.RowIndex >= 0 Then


            Pianificazione.commessa = DataGridView_commesse.Rows(e.RowIndex).Cells(columnName:="Commessa_tab").Value
            Homepage.Aggiorna_INI_COMPUTER()



        End If
    End Sub
End Class