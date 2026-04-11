Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Runtime.InteropServices
Imports FOXITREADERLib
Imports AxFOXITREADERLib

Public Class Form_Entrate_Merci

	Public codice_magazzino As String
	Public Codice_SAP As String
	'Public Qta_Mag As Decimal
	Public qta_trasferibile As Decimal
	Public Riga As Integer
	Public Documento As String
	Public Num_ODP As String
	Public Num_OC As String
	'Public Qta_max_Trasferibile As String
	Public Preview As New PrintPreviewDialog
	Public Sel_Stampante As New PrintDialog
	Public Stampante_Selezionata As Boolean
	Public Stampa_ODP As String
	Public Stampa_Tipo As String
	Public Stampa_Matricola As String
	Public Stampa_Descrizione As String
	Public Stampa_Descrizione_2 As String
	Public Stampa_Codice As String
	Public Stampa_Descrizione_Articolo As String
	Public Stampa_Descrizione_Articolo_2 As String
	Public Stampa_Qta As STRING
	Public Stampa_Ubicazione_Macchina As String
	Public Stampa_progressivo_commessa As String
	Public Stampa_Codice_Ordinazione As String
	Private nuovo_stato As String
	Public Risposta_BP_Trattamento As String
	Public riga_datagrid As Integer
	'Public codicedip As Integer
	Private variabile_controllo_trasferimento As Integer
	Private commento_controllo_trasferimento As String

	Public variabile_controllo_trasferimento_1 As Integer
	Private commento_controllo_trasferimento_1 As String
	Private ripetizione As String
	Public Soggetto_controllo_bp As String

	Private magazzino_destinazione As String
	Public Quantità_trasferita_per_scontrino As String
	Public CREA_SCONTRINO As String
	'Public preview_scontrino As Boolean = False


	Private codice_precedente As String

	Private quantità_predente As String
	Private numero_odp_precedente As Integer
	Private numero_oc_precedente As Integer
	Private giacenze_in_magazzino_precedente As Decimal
	Private tipo_entrata As String = "EM"
	Public trasferimento_eseguito = "NO"

	Public inizializzazione_form As Boolean = True
	Public Codice_BP_finale As String = ""
	Private id_em As Integer = 0
	Private Elenco_priorità_pacchi(1000) As String


	Public Sub Aggiorna()

		ComboBox1.Text = codice_magazzino

		Dim Cnn_Entrate_Merci As New SqlConnection
		Dim Cmd_Entrate_Merci As New SqlCommand
		Dim Cmd_Entrate_Merci_Reader As SqlDataReader

		Cnn_Entrate_Merci.ConnectionString = Homepage.sap_tirelli
		Cnn_Entrate_Merci.Open()
		Cmd_Entrate_Merci.Connection = Cnn_Entrate_Merci

		If RadioButton1.Checked = True Then
			Cmd_Entrate_Merci.CommandText = "SELECT MAX(DocNum) as 'Max_docnum' FROM OPDN"
		ElseIf RadioButton2.Checked = True Then
			Cmd_Entrate_Merci.CommandText = "SELECT MAX(DocNum) as 'Max_docnum' FROM OWTR"
		ElseIf RadioButton3.Checked = True Then
			Cmd_Entrate_Merci.CommandText = "SELECT MAX(DocNum) as 'Max_docnum' FROM OIGN"
		ElseIf RadioButton4.Checked = True Then
			Cmd_Entrate_Merci.CommandText = "SELECT MAX(DocNum) as 'Max_docnum' FROM OIQR"
		End If


		Cmd_Entrate_Merci_Reader = Cmd_Entrate_Merci.ExecuteReader
		Cmd_Entrate_Merci_Reader.Read()
		Txt_DocNum.Text = Cmd_Entrate_Merci_Reader("Max_docnum")
		Cnn_Entrate_Merci.Close()

		If RadioButton1.Checked = True Then
			Aggiorna_EM(Txt_DocNum.Text)
		ElseIf RadioButton2.Checked = True Then
			Aggiorna_trasferimento()

		ElseIf RadioButton3.Checked = True Then
			Aggiorna_EMP()
		ElseIf RadioButton4.Checked = True Then
			Aggiorna_CS(Txt_DocNum.Text)
		End If




	End Sub

	Public Sub Aggiorna_EM(par_numero_documento As Integer)

		Dim Cnn_Entrate_Merci As New SqlConnection
		Dim Cmd_Entrate_Merci As New SqlCommand
		Dim Cmd_Entrate_Merci_Reader As SqlDataReader
		Dim Txt_Ubicazione As String
		Txt_Fornitore.Text = ""
		Label6.Text = ""
		Soggetto_controllo_bp = "N"
		'Intestazione Entrata Merci

		Cnn_Entrate_Merci.ConnectionString = Homepage.sap_tirelli
		Cnn_Entrate_Merci.Open()
		Cmd_Entrate_Merci.Connection = Cnn_Entrate_Merci

		Cmd_Entrate_Merci.CommandText = "SELECT T0.[DocNum], T0.[DocDate] as 'Data',t0.cardcode, T0.[CardName] as 'Fornitore', 
CASE WHEN T1.[U_PRG_QLT_HasTC] IS NULL THEN 'N' ELSE T1.[U_PRG_QLT_HasTC] END AS 'Soggetto_Collaudo'

FROM OPDN T0 left join OCRD T1 ON T1.CARDCODE=T0.CARDCODE
WHERE T0.[DocNum] =" & par_numero_documento & ""
		Cmd_Entrate_Merci_Reader = Cmd_Entrate_Merci.ExecuteReader
		If Cmd_Entrate_Merci_Reader.Read() Then
			Txt_Fornitore.Text = Cmd_Entrate_Merci_Reader("Fornitore")
			Label6.Text = Cmd_Entrate_Merci_Reader("Data").ToString
			Soggetto_controllo_bp = Cmd_Entrate_Merci_Reader("Soggetto_Collaudo")
		End If
		Cnn_Entrate_Merci.Close()

		'Compilazione Tabella'
		Cnn_Entrate_Merci.ConnectionString = Homepage.sap_tirelli
		Cnn_Entrate_Merci.Open()
		Cmd_Entrate_Merci.Connection = Cnn_Entrate_Merci
		Cmd_Entrate_Merci.CommandText = "SELECT T0.[DocNum], T0.[DocDate] as 'Data', T0.[CardName] as 'Fornitore', T1.[ItemCode] as 'Codice'
, T1.[Dscription] as 'Descrizione'
,coalesce(t2.u_disegno,'') as 'U_disegno', sum(T1.[Quantity]) as 'Qta', T2.[U_PRG_TIR_Trattamento] as 'Trattamento',T2.[QryGroup10] as 'Ferretto',CASE WHEN T2.[U_Ubicazione] is NULL then 'XXX' else T2.[U_Ubicazione] end as 'Ubicazione', COALESCE(T3.ONHAND,0) as 'Mag_accettazione' 
FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] 
INNER JOIN OITW T3 ON T2.[ItemCode] = T3.[ItemCode]  
WHERE T3.[WhsCode] ='" & codice_magazzino & "' and t1.whscode='" & codice_magazzino & "'  AND  T0.[DocNum] =" & par_numero_documento & " AND  T1.ITEMCODE   Like '%%" & TextBox2.Text & "%%' and coalesce(t2.u_disegno,'') Like '%%" & TextBox5.Text & "%%'   and T3.ONHAND>0
group by T0.[DocNum], T0.[DocDate] , T0.[CardName] , T1.[ItemCode] , T1.[Dscription] ,coalesce(t2.u_disegno,'') ,T2.[U_PRG_TIR_Trattamento],T2.[QryGroup10],CASE WHEN T2.[U_Ubicazione] is NULL then 'XXX' else T2.[U_Ubicazione] end,T3.[OnHand]
ORDER BY T1.ITEMCODE"
		Cmd_Entrate_Merci_Reader = Cmd_Entrate_Merci.ExecuteReader
		DataGrid_EM.Rows.Clear()
		Dim conta_riga As Integer = 0
		Do While Cmd_Entrate_Merci_Reader.Read()
			If Cmd_Entrate_Merci_Reader("Ferretto") = "Y" Then
				Txt_Ubicazione = "FER"
			Else
				Txt_Ubicazione = Cmd_Entrate_Merci_Reader("Ubicazione")
			End If

			DataGrid_EM.Rows.Add(Cmd_Entrate_Merci_Reader("Codice"), Cmd_Entrate_Merci_Reader("Descrizione"), Cmd_Entrate_Merci_Reader("U_disegno"), Cmd_Entrate_Merci_Reader("Trattamento"), Math.Round(Cmd_Entrate_Merci_Reader("Qta"), 3), Math.Round(Cmd_Entrate_Merci_Reader("Mag_accettazione"), 3), Txt_Ubicazione, "Promemoria Refilling")


			conta_riga = conta_riga + 1
		Loop
		Cnn_Entrate_Merci.Close()
	End Sub

	Public Sub Aggiorna_EMP()

		Dim Cnn_Entrate_Merci As New SqlConnection
		Dim Cmd_Entrate_Merci As New SqlCommand
		Dim Cmd_Entrate_Merci_Reader As SqlDataReader
		Dim Txt_Ubicazione As String
		Txt_Fornitore.Text = ""
		Label6.Text = ""
		Soggetto_controllo_bp = "N"
		'Intestazione Entrata Merci

		Cnn_Entrate_Merci.ConnectionString = Homepage.sap_tirelli
		Cnn_Entrate_Merci.Open()
		Cmd_Entrate_Merci.Connection = Cnn_Entrate_Merci

		Cmd_Entrate_Merci.CommandText = "SELECT T0.[DocNum], T0.[DocDate] as 'Data'


FROM OIGN T0 left join OCRD T1 ON T1.CARDCODE=T0.CARDCODE
WHERE T0.[DocNum] =" & Txt_DocNum.Text & ""
		Cmd_Entrate_Merci_Reader = Cmd_Entrate_Merci.ExecuteReader

		If Cmd_Entrate_Merci_Reader.Read() Then
			'	Txt_Fornitore.Text = Cmd_Entrate_Merci_Reader("Fornitore")
			Label6.Text = Cmd_Entrate_Merci_Reader("Data").ToString
			'	Soggetto_controllo_bp = Cmd_Entrate_Merci_Reader("Soggetto_Collaudo")
		End If
		Cnn_Entrate_Merci.Close()

		'Compilazione Tabella'
		Cnn_Entrate_Merci.ConnectionString = Homepage.sap_tirelli
		Cnn_Entrate_Merci.Open()
		Cmd_Entrate_Merci.Connection = Cnn_Entrate_Merci
		Cmd_Entrate_Merci.CommandText = "SELECT T0.[DocNum], T0.[DocDate] as 'Data', T0.[CardName] as 'Fornitore', T1.[ItemCode] as 'Codice', T1.[Dscription] as 'Descrizione'
,coalesce(t2.u_disegno,'') as 'U_disegno', sum(T1.[Quantity]) as 'Qta', T2.[U_PRG_TIR_Trattamento] as 'Trattamento',T2.[QryGroup10] as 'Ferretto',CASE WHEN T2.[U_Ubicazione] is NULL then 'XXX' else T2.[U_Ubicazione] end as 'Ubicazione', COALESCE(T3.ONHAND,0) as 'Mag_accettazione' 
FROM OIGN T0  INNER JOIN IGN1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] 
INNER JOIN OITW T3 ON T2.[ItemCode] = T3.[ItemCode]  
WHERE T3.[WhsCode] ='" & codice_magazzino & "' and t1.whscode='" & codice_magazzino & "' AND  T1.ITEMCODE   Like '%%" & TextBox2.Text & "%%' and coalesce(t2.u_disegno,'') Like '%%" & TextBox5.Text & "%%'  AND  T0.[DocNum] =" & Txt_DocNum.Text & " and T3.ONHAND>0
group by T0.[DocNum], T0.[DocDate] , T0.[CardName] , T1.[ItemCode] , T1.[Dscription] ,coalesce(t2.u_disegno,'') ,T2.[U_PRG_TIR_Trattamento],T2.[QryGroup10],CASE WHEN T2.[U_Ubicazione] is NULL then 'XXX' else T2.[U_Ubicazione] end,T3.[OnHand]
ORDER BY T1.ITEMCODE"
		Cmd_Entrate_Merci_Reader = Cmd_Entrate_Merci.ExecuteReader
		DataGrid_EM.Rows.Clear()
		Dim conta_riga As Integer = 0
		Do While Cmd_Entrate_Merci_Reader.Read()
			If Cmd_Entrate_Merci_Reader("Ferretto") = "Y" Then
				Txt_Ubicazione = "FER"
			Else
				Txt_Ubicazione = Cmd_Entrate_Merci_Reader("Ubicazione")
			End If

			DataGrid_EM.Rows.Add(Cmd_Entrate_Merci_Reader("Codice"), Cmd_Entrate_Merci_Reader("Descrizione"), Cmd_Entrate_Merci_Reader("U_disegno"), Cmd_Entrate_Merci_Reader("Trattamento"), Math.Round(Cmd_Entrate_Merci_Reader("Qta"), 3), Math.Round(Cmd_Entrate_Merci_Reader("Mag_accettazione"), 3), Txt_Ubicazione, "Promemoria Refilling")


			conta_riga = conta_riga + 1
		Loop
		Cnn_Entrate_Merci.Close()
	End Sub

	Public Sub Aggiorna_trasferimento()
		'Codice_SAP = ""
		Dim Cnn_Entrate_Merci As New SqlConnection
		Dim Cmd_Entrate_Merci As New SqlCommand
		Dim Cmd_Entrate_Merci_Reader As SqlDataReader
		Dim Txt_Ubicazione As String
		Txt_Fornitore.Text = ""
		Label6.Text = ""
		'Intestazione Entrata Merci

		Cnn_Entrate_Merci.ConnectionString = Homepage.sap_tirelli
		Cnn_Entrate_Merci.Open()
		Cmd_Entrate_Merci.Connection = Cnn_Entrate_Merci
		Cmd_Entrate_Merci.CommandText = "SELECT T0.[DocNum], T0.[DocDate] as 'Data', T0.[filler] as 'Fornitore'FROM OWTR T0  WHERE T0.[DocNum] =" & Txt_DocNum.Text
		Cmd_Entrate_Merci_Reader = Cmd_Entrate_Merci.ExecuteReader
		If Cmd_Entrate_Merci_Reader.Read() Then
			Txt_Fornitore.Text = Cmd_Entrate_Merci_Reader("Fornitore")
			Label6.Text = Cmd_Entrate_Merci_Reader("Data").ToString
		End If
		Cnn_Entrate_Merci.Close()

		'Compilazione Tabella'
		Cnn_Entrate_Merci.ConnectionString = Homepage.sap_tirelli
		Cnn_Entrate_Merci.Open()
		Cmd_Entrate_Merci.Connection = Cnn_Entrate_Merci
		Cmd_Entrate_Merci.CommandText = "SELECT T0.[DocNum], T0.[DocDate] as 'Data', T0.[CardName] as 'Fornitore', T1.[ItemCode] as 'Codice', T1.[Dscription] as 'Descrizione'
,coalesce(t2.u_disegno,'') as 'U_disegno', T1.[Quantity] as 'Qta', T2.[U_PRG_TIR_Trattamento] as 'Trattamento',T2.[QryGroup10] as 'Ferretto',CASE WHEN T2.[U_Ubicazione] is NULL then 'XXX' else T2.[U_Ubicazione] end as 'Ubicazione', T3.[OnHand] as 'Mag_accettazione' 

FROM owtr T0  INNER JOIN wtr1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] 
INNER JOIN OITW T3 ON T2.[ItemCode] = T3.[ItemCode] 
WHERE T3.[WhsCode] ='" & codice_magazzino & " ' and t1.whscode='" & codice_magazzino & "' AND T3.[OnHand] >0 AND  T1.ITEMCODE   Like '%%" & TextBox2.Text & "%%' and coalesce(t2.u_disegno,'') Like '%%" & TextBox5.Text & "%%' AND  T0.[DocNum] =" & Txt_DocNum.Text & "
ORDER BY T1.ITEMCODE
"
		Cmd_Entrate_Merci_Reader = Cmd_Entrate_Merci.ExecuteReader
		DataGrid_EM.Rows.Clear()
		Do While Cmd_Entrate_Merci_Reader.Read()
			If Cmd_Entrate_Merci_Reader("Ferretto") = "Y" Then
				Txt_Ubicazione = "FER"
			Else
				Txt_Ubicazione = Cmd_Entrate_Merci_Reader("Ubicazione")
			End If
			DataGrid_EM.Rows.Add(Cmd_Entrate_Merci_Reader("Codice"), Cmd_Entrate_Merci_Reader("Descrizione"), Cmd_Entrate_Merci_Reader("U_disegno"), Cmd_Entrate_Merci_Reader("Trattamento"), Math.Round(Cmd_Entrate_Merci_Reader("Qta"), 3), Math.Round(Cmd_Entrate_Merci_Reader("Mag_accettazione"), 3), Txt_Ubicazione, "Promemoria Refilling")
		Loop
		Cnn_Entrate_Merci.Close()
	End Sub

	Public Sub Aggiorna_CS(PAR_NUMERO_DOCUMENTO As Integer)
		'Codice_SAP = ""
		Dim Cnn_Entrate_Merci As New SqlConnection
		Dim Cmd_Entrate_Merci As New SqlCommand
		Dim Cmd_Entrate_Merci_Reader As SqlDataReader
		Dim Txt_Ubicazione As String
		Txt_Fornitore.Text = ""
		Label6.Text = ""
		'Intestazione Entrata Merci

		Cnn_Entrate_Merci.ConnectionString = Homepage.sap_tirelli
		Cnn_Entrate_Merci.Open()
		Cmd_Entrate_Merci.Connection = Cnn_Entrate_Merci
		Cmd_Entrate_Merci.CommandText = "SELECT T0.[DocNum], T0.[DocDate] as 'Data'
FROM OIQR T0  WHERE T0.[DocNum] =" & PAR_NUMERO_DOCUMENTO
		Cmd_Entrate_Merci_Reader = Cmd_Entrate_Merci.ExecuteReader
		If Cmd_Entrate_Merci_Reader.Read() Then

			Label6.Text = Cmd_Entrate_Merci_Reader("Data").ToString
		End If
		Cnn_Entrate_Merci.Close()

		'Compilazione Tabella'
		Cnn_Entrate_Merci.ConnectionString = Homepage.sap_tirelli
		Cnn_Entrate_Merci.Open()
		Cmd_Entrate_Merci.Connection = Cnn_Entrate_Merci
		Cmd_Entrate_Merci.CommandText = "SELECT T0.[DocNum], T0.[DocDate] as 'Data', '' as 'Fornitore', T1.[ItemCode] as 'Codice', T2.[ITEMNAME] as 'Descrizione'
,coalesce(t2.u_disegno,'') as 'U_disegno'
, T1.[Quantity] as 'Qta', T2.[U_PRG_TIR_Trattamento] as 'Trattamento',T2.[QryGroup10] as 'Ferretto'
,CASE WHEN T2.[U_Ubicazione] is NULL then 'XXX' else T2.[U_Ubicazione] end as 'Ubicazione', T3.[OnHand] as 'Mag_accettazione' 

FROM OIQR T0  INNER JOIN IQR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] 
INNER JOIN OITW T3 ON T2.[ItemCode] = T3.[ItemCode] 
 
WHERE T3.[WhsCode] ='" & codice_magazzino & " ' and t1.whscode='" & codice_magazzino & "' AND T3.[OnHand] >0 AND  T1.ITEMCODE   Like '%%" & TextBox2.Text & "%%' and coalesce(t2.u_disegno,'') Like '%%" & TextBox5.Text & "%%' AND  T0.[DocNum] =" & Txt_DocNum.Text & "
ORDER BY T1.ITEMCODE
"
		Cmd_Entrate_Merci_Reader = Cmd_Entrate_Merci.ExecuteReader
		DataGrid_EM.Rows.Clear()
		Do While Cmd_Entrate_Merci_Reader.Read()
			If Cmd_Entrate_Merci_Reader("Ferretto") = "Y" Then
				Txt_Ubicazione = "FER"
			Else
				Txt_Ubicazione = Cmd_Entrate_Merci_Reader("Ubicazione")
			End If
			DataGrid_EM.Rows.Add(Cmd_Entrate_Merci_Reader("Codice"), Cmd_Entrate_Merci_Reader("Descrizione"), Cmd_Entrate_Merci_Reader("U_disegno"), Cmd_Entrate_Merci_Reader("Trattamento"), Math.Round(Cmd_Entrate_Merci_Reader("Qta"), 3), Math.Round(Cmd_Entrate_Merci_Reader("Mag_accettazione"), 3), Txt_Ubicazione, "Promemoria Refilling")


		Loop
		Cnn_Entrate_Merci.Close()
	End Sub

	Private Sub Cmd_Indietro_Click(sender As Object, e As EventArgs) Handles Cmd_Indietro.Click
		If inizializzazione_form = False Then


			Txt_DocNum.Text = Int(Txt_DocNum.Text) - 1
			If RadioButton1.Checked = True Then
				Aggiorna_EM(Txt_DocNum.Text)
			ElseIf RadioButton2.Checked = True Then
				Aggiorna_trasferimento()
			ElseIf RadioButton3.Checked = True Then
				Aggiorna_EMP()
			ElseIf RadioButton4.Checked = True Then
				Aggiorna_CS(Txt_DocNum.Text)
			End If
		End If
	End Sub

	Private Sub Cmd_Avanti_Click(sender As Object, e As EventArgs) Handles Cmd_Avanti.Click
		If inizializzazione_form = False Then
			Txt_DocNum.Text = Int(Txt_DocNum.Text) + 1
			If RadioButton1.Checked = True Then
				Aggiorna_EM(Txt_DocNum.Text)
			ElseIf RadioButton2.Checked = True Then
				Aggiorna_trasferimento()
			ElseIf RadioButton3.Checked = True Then
				Aggiorna_EMP()
			ElseIf RadioButton4.Checked = True Then
				Aggiorna_CS(Txt_DocNum.Text)
			End If
		End If
	End Sub

	Sub trasferito(par_codice_sap As String, par_datagridview As DataGridView)
		Dim Cnn1 As New SqlConnection
		par_datagridview.Rows.Clear()

		Cnn1.ConnectionString = Homepage.sap_tirelli
		'MsgBox(Stringa_Connessione_SAP)
		Cnn1.Open()


		Dim CMD_SAP_2 As New SqlCommand
		Dim cmd_SAP_reader_2 As SqlDataReader


		CMD_SAP_2.Connection = Cnn1
		CMD_SAP_2.CommandText = "declare @giorni_per_prelievo as integer
set @giorni_per_prelievo=10

Select t0.[Documento] , T0.[ODP], t0.[OC],t0.[Codice], T0.ITEMNAME,t0.[Q.tà pianificata], T0.[Trasferito], case when t0.[Da trasferire] is null then 0 else t0.[Da trasferire] end as 'Da trasferire',
case when T0.[U_PRG_WIP_QtaRichMagAuto] is null then 0 else T0.[U_PRG_WIP_QtaRichMagAuto] end as 'U_PRG_WIP_QtaRichMagAuto' ,
t0.U_qta_richiesta_wip,
coalesce(t0.[U_prg_azs_commessa],'') as 'U_prg_azs_commessa', coalesce(t0.[U_utilizz],'') as 'u_utilizz',
coalesce(t2.nome_baia,'') as 'Baia',
t0.[status] , t0.[u_PRODUZIONE],T0.LINENUM, t0.resname, t0.startdate, t0.progressivo_commessa
,T0.DIV

 
from (
SELECT 'ODP' as 'Documento', T1.[DocNum] as 'ODP', '' as 'OC',T2.[ItemCode] as 'Codice', T1.PRODNAME AS 'ITEMNAME', T0.[PlannedQty] as 'Q.tà pianificata',
CASE WHEN T0.[U_PRG_WIP_QtaSpedita] is null then 0 else T0.[U_PRG_WIP_QtaSpedita] end AS 'Trasferito',
CASE WHEN T1.U_PRODUZIONE='INT' THEN
CASE 
	WHEN COALESCE(T0.[U_PRG_WIP_QtaSpedita],0)>COALESCE(T0.[U_qta_richiesta_wip],0) THEN T0.[PlannedQty]-COALESCE(T0.[U_PRG_WIP_QtaDaTrasf],0) 
	ELSE T0.[PlannedQty] -COALESCE(T0.[U_qta_richiesta_wip],0) END
ELSE
T0.[U_PRG_WIP_QtaDaTrasf] END as 'Da trasferire', T0.[U_PRG_WIP_QtaRichMagAuto],
case when T0.[U_PRG_WIP_QtaRichMagAuto] is null then 0 else T0.[U_qta_richiesta_wip] end as 'U_qta_richiesta_wip',
t1.U_prg_azs_commessa, t1.U_utilizz, t1.status , T1.u_PRODUZIONE,T0.LINENUM, t3.resname, t1.startdate,
case when t1.u_progressivo_commessa is null then 0 else t1.u_progressivo_commessa end as 'Progressivo_commessa'
,case when coalesce(t4.location,'')='13' then 'BRB01' ELSE 'TIR01' END AS 'DIV'


FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] 
inner join oitm t2 on t2.itemcode= t0.itemcode
left join orsc t3 on t3.visrescode=t1.u_fase
left join owhs t4 on t4.whscode=t0.wareHouse



WHERE T0.[ItemCode] = '" & par_codice_sap & "' AND  (T1.Status <> N'L' )  AND  (T1.Status <> N'C' )


union all


SELECT 'OC','',T1.DOCNUM AS 'OC',T0.[ItemCode], '',  T0.[OpenQty], T0.[U_Trasferito] as 'Totale trasferito', T0.[U_Datrasferire] as 'Da trasferire' ,T0.[U_PRG_WIP_QtaRichMagAuto],
0, T1.U_MATRCDS AS 'CDS', T1.CARDNAME,'','' ,T0.LINENUM,'', t1.docduedate
,0 as 'Progressivo_commessa'
,COALESCE(T0.OcrCode,'') AS 'DIV'
FROM RDR1 T0  INNER JOIN ORDR T1 ON T0.[DocEntry] = T1.[DocEntry]


WHERE T1.DocStatus = N'o' and T0.[OpenCreQty] >0 and T0.[ItemCode] ='" & par_codice_sap & "'
) as t0
left join [tirelli_40].dbo.Layout_CAP1 t1 on t1.Commessa =t0.U_PRG_AZS_Commessa and t1.Stato='O'
left join [tirelli_40].dbo.Layout_CAP1_nomi t2 on t2.numero_baia=t1.baia

group by
t0.[Documento] ,coalesce(t2.nome_baia,''), T0.[ODP], t0.[OC],t0.[Codice], T0.ITEMNAME,t0.[Q.tà pianificata], T0.[Trasferito], t0.[Da trasferire], t0.[U_prg_azs_commessa], t0.[U_utilizz], t0.[status] , t0.[u_PRODUZIONE],T0.LINENUM,t0.resname , t0.startdate,T0.[U_PRG_WIP_QtaRichMagAuto],t0.U_qta_richiesta_wip,t0.progressivo_commessa,T0.DIV
order by t0.STARTdate"

		cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

		Do While cmd_SAP_reader_2.Read()

			If cmd_SAP_reader_2("Da trasferire") > 0 Then
				par_datagridview.Rows.Add(cmd_SAP_reader_2("documento"), cmd_SAP_reader_2("ODP"), cmd_SAP_reader_2("ITEMNAME"), cmd_SAP_reader_2("Progressivo_commessa"), cmd_SAP_reader_2("OC"), cmd_SAP_reader_2("Q.tà pianificata"), cmd_SAP_reader_2("Trasferito"), cmd_SAP_reader_2("Da trasferire"), cmd_SAP_reader_2("u_prg_wip_qtarichmagauto"), cmd_SAP_reader_2("u_qta_richiesta_wip"), cmd_SAP_reader_2("U_prg_azs_commessa"), cmd_SAP_reader_2("U_utilizz"), cmd_SAP_reader_2("Baia"), cmd_SAP_reader_2("DIV"), cmd_SAP_reader_2("status"), cmd_SAP_reader_2("LINENUM"), cmd_SAP_reader_2("u_produzione"), cmd_SAP_reader_2("Resname"), 0, "Trasferisci", "")
			End If
		Loop

		cmd_SAP_reader_2.Close()
		Cnn1.Close()
		par_datagridview.ClearSelection()
	End Sub

	Private Sub Cmd_Cerca_Click(sender As Object, e As EventArgs) Handles Cmd_Cerca.Click
		If inizializzazione_form = False Then


			If RadioButton1.Checked = True Then
				Aggiorna_EM(Txt_DocNum.Text)
			ElseIf RadioButton2.Checked = True Then
				Aggiorna_trasferimento()
			ElseIf RadioButton3.Checked = True Then
				Aggiorna_EMP()
			ElseIf RadioButton4.Checked = True Then
				Aggiorna_CS(Txt_DocNum.Text)
			End If
		End If



	End Sub

	Sub giacenze_magazzino(par_codice_sap As String)
		Dim Cnn1 As New SqlConnection

		Dim magazzino_tot As Decimal
		Dim Confermato_tot As Decimal
		Dim ordinato_tot As Decimal
		Dim disponibile As Decimal

		DataGridView_magazzino.Rows.Clear()
		Cnn1.ConnectionString = Homepage.sap_tirelli
		Cnn1.Open()


		Dim CMD_SAP_2 As New SqlCommand
		Dim cmd_SAP_reader_2 As SqlDataReader


		CMD_SAP_2.Connection = Cnn1
		CMD_SAP_2.CommandText = "SELECT T0.[WhsCode], CASE WHEN T0.[OnHand] is null then 0 else T0.[OnHand] END AS 'onhand' , case when T0.[IsCommited] is null then 0 else T0.[IsCommited] end as 'iscommited' , case when T0.[OnOrder] is null then 0 else T0.[OnOrder] end as 'onorder'  
FROM OITW T0 WHERE (T0.[OnHand]>0 or t0.iscommited>0 or t0.onorder>0) and t0.itemcode='" & par_codice_sap & "'"


		cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

		Do While cmd_SAP_reader_2.Read()


			DataGridView_magazzino.Rows.Add(cmd_SAP_reader_2("whscode"), cmd_SAP_reader_2("onhand"), cmd_SAP_reader_2("iscommited"), cmd_SAP_reader_2("onorder"))

		Loop

		cmd_SAP_reader_2.Close()
		Cnn1.Close()
		Dim Cnn As New SqlConnection
		Cnn.Close()
		Cnn.ConnectionString = Homepage.sap_tirelli

		Cnn.Open()

		Dim CMD_SAP As New SqlCommand
		Dim cmd_SAP_reader As SqlDataReader


		CMD_SAP.Connection = Cnn
		CMD_SAP.CommandText = " select  sum(case when T0.[OnHand] is null then 0 else T0.[OnHand] end ) as 'Magazzino_TOT', sum(case when T0.[iscoMMited] is null then 0 else T0.[iscoMMited] end) as 'Confermato_TOT', sum(case when T0.[onorder] is null then 0 else T0.[onorder] end) as 'ordinato_TOT',  sum(case when T0.[OnHand] is null then 0 else T0.[OnHand] end-case when T0.[iscoMMited] is null then 0 else T0.[iscoMMited] end+case when T0.[onorder] is null then 0 else T0.[onorder] end) as 'Disponibile'
FROM OITW T0 WHERE (T0.[OnHand]>0 or t0.iscommited>0 or t0.onorder>0) and t0.itemcode='" & Codice_SAP & "'"
		cmd_SAP_reader = CMD_SAP.ExecuteReader

		If cmd_SAP_reader.Read() = True Then

			If Not cmd_SAP_reader("Magazzino_TOT") Is System.DBNull.Value Then
				magazzino_tot = cmd_SAP_reader("Magazzino_TOT")
			Else
				magazzino_tot = 0
			End If

			If Not cmd_SAP_reader("Confermato_TOT") Is System.DBNull.Value Then
				Confermato_tot = cmd_SAP_reader("Confermato_TOT")
			Else
				Confermato_tot = 0
			End If

			If Not cmd_SAP_reader("ordinato_TOT") Is System.DBNull.Value Then
				ordinato_tot = cmd_SAP_reader("ordinato_TOT")
			Else
				ordinato_tot = 0
			End If

			If Not cmd_SAP_reader("disponibile") Is System.DBNull.Value Then
				disponibile = cmd_SAP_reader("disponibile")
			Else
				disponibile = 0
			End If


		Else
			magazzino_tot = 0
			Confermato_tot = 0
			ordinato_tot = 0
			disponibile = 0

		End If
		cmd_SAP_reader.Close()
		Cnn.Close()
		DataGridView_magazzino.Rows.Add("TOTALE", magazzino_tot, Confermato_tot, ordinato_tot, disponibile)
	End Sub

	Public Function giacenze_IN_magazzino(par_codice_sap As String, par_codice_magazzino As String)

		Dim qta_a_mag As Decimal = 0


		Dim Cnn1 As New SqlConnection


		Cnn1.ConnectionString = Homepage.sap_tirelli
		Cnn1.Open()


		Dim CMD_SAP_2 As New SqlCommand
		Dim cmd_SAP_reader_2 As SqlDataReader


		CMD_SAP_2.Connection = Cnn1
		CMD_SAP_2.CommandText = "SELECT T0.[WhsCode], CASE WHEN T0.[OnHand] is null then 0 else T0.[OnHand] END AS 'onhand' 

FROM OITW T0 WHERE  t0.itemcode='" & par_codice_sap & "'and T0.[WhsCode]='" & par_codice_magazzino & "' "


		cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

		If cmd_SAP_reader_2.Read() Then

			qta_a_mag = cmd_SAP_reader_2("onhand")

		End If

		cmd_SAP_reader_2.Close()

		Cnn1.Close()

		Return qta_a_mag
	End Function

	Public Function Trova_business_unit_magazzino(par_codice_magazzino As String)

		Dim location As Integer


		Dim Cnn1 As New SqlConnection


		Cnn1.ConnectionString = Homepage.sap_tirelli
		Cnn1.Open()


		Dim CMD_SAP_2 As New SqlCommand
		Dim cmd_SAP_reader_2 As SqlDataReader


		CMD_SAP_2.Connection = Cnn1
		CMD_SAP_2.CommandText = "select coalesce(t0.location,0) as 'Locazione'
from owhs t0
where t0.whscode='" & par_codice_magazzino & "' "


		cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

		If cmd_SAP_reader_2.Read() Then

			location = cmd_SAP_reader_2("locazione")

		End If

		cmd_SAP_reader_2.Close()

		Cnn1.Close()

		Return location
	End Function

	Public Function Verifica_odp_anticipo_materiale(par_numero_odp As String, par_magazzino_destinazione As String)

		Dim verifica As Integer = 0


		Dim Cnn1 As New SqlConnection


		Cnn1.ConnectionString = Homepage.sap_tirelli
		Cnn1.Open()


		Dim CMD_SAP_2 As New SqlCommand
		Dim cmd_SAP_reader_2 As SqlDataReader


		CMD_SAP_2.Connection = Cnn1
		CMD_SAP_2.CommandText = "select t1.itemcode, T1.DOCNUM
from OWOR t1
where t1.prodname Like '%%ANTICIP%%' and t1.docnum ='" & par_numero_odp & "' and ('" & par_magazzino_destinazione & "' ='WIP' OR '" & par_magazzino_destinazione & "'='BWIP') "


		cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

		If cmd_SAP_reader_2.Read() Then

			verifica = 1
		Else
			verifica = 0
		End If

		cmd_SAP_reader_2.Close()

		Cnn1.Close()

		Return verifica
	End Function

	Public Function quantita_da_trasferire_nel_documento(par_tipo_documento As String, par_codice_sap As String, par_numero_odp As Integer, par_numero_oc As Integer, par_linenum As Integer)

		Dim qta_da_trasferire As Decimal = 0


		Dim Cnn1 As New SqlConnection


		Cnn1.ConnectionString = Homepage.sap_tirelli
		Cnn1.Open()


		Dim CMD_SAP_2 As New SqlCommand
		Dim cmd_SAP_reader_2 As SqlDataReader


		CMD_SAP_2.Connection = Cnn1

		If par_tipo_documento = "ODP" Then

			CMD_SAP_2.CommandText = "SELECT T0.[docnum],
coalesce(t1.U_PRG_WIP_QTADATRASF,0) AS 'da_trasferire'

from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
where t0.docnum='" & par_numero_odp & "' and t1.linenum ='" & par_linenum & "' and t1.itemcode ='" & par_codice_sap & "'"


		ElseIf par_tipo_documento = "OC" Then
			CMD_SAP_2.CommandText = "SELECT T0.[docnum],
coalesce(t1.U_datrasferire,0) AS 'da_trasferire'

from ordr t0 inner join rdr1 t1 on t0.docentry=t1.docentry
where t0.docnum='" & par_numero_oc & "' and t1.linenum ='" & par_linenum & "' and t1.itemcode ='" & par_codice_sap & "'"

		End If



		cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

		If cmd_SAP_reader_2.Read() Then

			qta_da_trasferire = cmd_SAP_reader_2("da_trasferire")

		End If

		cmd_SAP_reader_2.Close()

		Cnn1.Close()

		Return qta_da_trasferire
	End Function

	Public Function quantita_entrata_con_em(par_numero_doc As Integer, par_codice_sap As String, par_codice_magazzino As String, par_tipo_entrata As String)

		Dim qta_entrata_cont_em As Decimal = 0


		Dim Cnn1 As New SqlConnection


		Cnn1.ConnectionString = Homepage.sap_tirelli
		Cnn1.Open()


		Dim CMD_SAP_2 As New SqlCommand
		Dim cmd_SAP_reader_2 As SqlDataReader


		CMD_SAP_2.Connection = Cnn1


		If par_tipo_entrata = "EM" Then
			CMD_SAP_2.CommandText = "SELECT sum(coalesce(T1.[Quantity],0)) as 'Qta'


FROM OPDN T0  INNER JOIN PDN1 T1 ON T0.[DocEntry] = T1.[DocEntry] 

WHERE T1.[WhsCode] ='" & par_codice_magazzino & "' and t1.itemcode='" & par_codice_sap & "'  AND  T0.[DocNum] =" & par_numero_doc & " 

"

		ElseIf par_tipo_entrata = "TRASF" Then
			CMD_SAP_2.CommandText = "SELECT sum(coalesce(T1.[Quantity],0)) as 'Qta'


FROM OWTR T0  INNER JOIN WTR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 

WHERE T1.[WhsCode] ='" & par_codice_magazzino & "' and t1.itemcode='" & par_codice_sap & "'  AND  T0.[DocNum] =" & par_numero_doc & " 
"
		ElseIf par_tipo_entrata = "EMP" Then

			CMD_SAP_2.CommandText = "SELECT sum(coalesce(T1.[Quantity],0)) as 'Qta'


FROM OIGN T0  INNER JOIN IGN1 T1 ON T0.[DocEntry] = T1.[DocEntry] 

WHERE T1.[WhsCode] ='" & par_codice_magazzino & "' and t1.itemcode='" & par_codice_sap & "'  AND  T0.[DocNum] =" & par_numero_doc & " 
"

		ElseIf par_tipo_entrata = "CS" Then

			CMD_SAP_2.CommandText = "SELECT sum(coalesce(T1.[Quantity],0)) as 'Qta'


FROM  OIQR T0  INNER JOIN Iqr1 T1 ON T0.[DocEntry] = T1.[DocEntry] 

WHERE T1.[WhsCode] ='" & par_codice_magazzino & "' and t1.itemcode='" & par_codice_sap & "'  AND  T0.[DocNum] =" & par_numero_doc & " 
"

		Else

			CMD_SAP_2.CommandText = "Select 999999 as 'Qta'"

		End If


		cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

		If cmd_SAP_reader_2.Read() Then

			qta_entrata_cont_em = cmd_SAP_reader_2("qta")

		End If

		cmd_SAP_reader_2.Close()

		Cnn1.Close()

		Return qta_entrata_cont_em
	End Function





	Public Function controllo_RIchiesta_trasferimento_ferretto(par_tipo_documento As String, par_numero_odp As Integer, par_numero_oc As Integer, par_linenum As Integer, par_codice_sap As String)


		Dim qta_richiesta_mag_auto As Decimal = 0

		Dim Cnn1 As New SqlConnection


		Cnn1.ConnectionString = Homepage.sap_tirelli
		Cnn1.Open()


		Dim CMD_SAP_2 As New SqlCommand
		Dim cmd_SAP_reader_2 As SqlDataReader


		CMD_SAP_2.Connection = Cnn1

		If par_tipo_documento = "ODP" Then

			CMD_SAP_2.CommandText = "SELECT T0.[docnum],
coalesce(T1.u_prg_wip_qtarichmagauto, 0)+SUM(COALESCE(T2.QUANTITY,0)) AS 'RICHIESTA_ferretto'
From owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
left JOIN WTQ1 T2 ON T2.LINESTATUS='O' AND T2.ItEMCODE=T1.ITEMCODE AND T2.U_PRG_AZS_OPDOCENTRY=T0.DOCENTRY
where t0.docnum='" & par_numero_odp & "' and t1.linenum ='" & par_linenum & "' and t1.itemcode ='" & par_codice_sap & "'
GROUP BY T0.[docnum],coalesce(T1.u_prg_wip_qtarichmagauto, 0)"


		ElseIf par_tipo_documento = "OC" Then
			CMD_SAP_2.CommandText = "SELECT T0.[docnum],
coalesce(T1.u_prg_wip_qtarichmagauto,0)+SUM(COALESCE(T2.QUANTITY,0)) AS 'RICHIESTA_ferretto'

from ordr t0 inner join rdr1 t1 on t0.docentry=t1.docentry
left JOIN WTQ1 T2 ON T2.LINESTATUS='O' AND T2.ItEMCODE=T1.ITEMCODE AND T2.U_PRG_AZS_OcDOCENTRY=T0.DOCENTRY
where t0.docnum='" & par_numero_oc & "' and t1.linenum ='" & par_linenum & "' and t1.itemcode ='" & par_codice_sap & "'
GROUP BY T0.[docnum],coalesce(T1.u_prg_wip_qtarichmagauto,0)"

		End If



		cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

		If cmd_SAP_reader_2.Read() Then

			qta_richiesta_mag_auto = cmd_SAP_reader_2("RICHIESTA_ferretto")

		End If

		cmd_SAP_reader_2.Close()

		Cnn1.Close()

		Return qta_richiesta_mag_auto

	End Function

	Sub controllo_RT_Ferretto_aperte_per_trasferimento_ad_altro_magazzino()
		variabile_controllo_trasferimento = 0
		commento_controllo_trasferimento = ""

		If DataGridView2.Rows(Riga).Cells(columnName:="Richiesta_ferretto").Value > 0 Then
			variabile_controllo_trasferimento = variabile_controllo_trasferimento + 1
			commento_controllo_trasferimento = "Il codice è in richiesta di prelievo a Ferretto"
		End If


	End Sub

	Sub dettaglio_backlog_ordini()
		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()



		Dim CMD_SAP_docentry As New SqlCommand
		Dim cmd_SAP_docentry_reader As SqlDataReader

		CMD_SAP_docentry.Connection = Cnn
		CMD_SAP_docentry.CommandText = ""

		cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader


		If cmd_SAP_docentry_reader.Read() Then

		End If
		cmd_SAP_docentry_reader.Close()
		Cnn.Close()


	End Sub 'Inserisco le risorse nella combo box


	Sub Fun_Stampa(par_stampa_tipo As String, par_preview_scontrino As Boolean, par_stampante_selezionata As Boolean, par_scontrino As PrintDocument, par_magazzino_destinazione As String, par_ubicazione As String, par_codice_sap As String)
		Codice_SAP = par_codice_sap
		magazzino_destinazione = par_magazzino_destinazione
		Dim altezza_scontrino As Integer
		If par_stampa_tipo <> "Trasferimento interno" Then
			altezza_scontrino = 215
		Else
			altezza_scontrino = 155
		End If

		'preview_scontrino = True

		If par_preview_scontrino = True Then
			If par_stampante_selezionata = False Then
				Sel_Stampante.AllowSomePages = False
				Sel_Stampante.ShowHelp = False
				Sel_Stampante.Document = Scontrino

				' Utilizza un PrintPreviewDialog per mostrare l'anteprima di stampa
				Dim previewDialog As New PrintPreviewDialog()
				previewDialog.Document = Scontrino

				Dim result As DialogResult = previewDialog.ShowDialog()

				If (result = DialogResult.OK) Then
					par_stampante_selezionata = True

					' Ora la stampante è selezionata, puoi chiamare Scontrino.Print()
					par_scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", 185, altezza_scontrino)
					par_scontrino.Print()
				End If
			Else
				' Se la stampante è già stata selezionata in precedenza, stampa direttamente
				par_scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", 185, altezza_scontrino)
				par_scontrino.Print()
			End If

		Else
			If par_stampante_selezionata = False Then
				Sel_Stampante.AllowSomePages = False
				Sel_Stampante.ShowHelp = False
				Sel_Stampante.Document = Scontrino
				Dim result As DialogResult = Sel_Stampante.ShowDialog()
				If (result = DialogResult.OK) Then
					Stampante_Selezionata = True
					' Utilizza un PrintPreviewDialog per mostrare l'anteprima di stampa
					par_scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", 185, altezza_scontrino)
					Dim previewDialog As New PrintPreviewDialog()
					previewDialog.Document = Scontrino
					par_scontrino.Print()
				End If
			Else
				par_scontrino.Print()
			End If
		End If


	End Sub




	Private Sub Scontrino_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles Scontrino.PrintPage

		Dim Penna As New Pen(Color.Black)
		Dim Carattere_ODP As New Font("Calibri", 16, FontStyle.Bold)
		Dim Carattere_Matricola As New Font("Calibri", 14, FontStyle.Bold)
		Dim Carattere_Descrizione As New Font("Calibri", 8, FontStyle.Italic)
		Dim Carattere_Codice As New Font("Calibri", 22, FontStyle.Italic)
		Dim Carattere_Descrizione_Articolo As New Font("Calibri", 8, FontStyle.Italic)
		Dim Carattere_Qta As New Font("Calibri", 12, FontStyle.Bold)
		Dim Carattere_Ubicazione As New Font("Calibri", 12, FontStyle.Italic)
		Dim Carattere_posizione As New Font("Calibri", 15, FontStyle.Bold)
		Dim Carattere_Diciture As New Font("Calibri", 6, FontStyle.Italic)



		With e.Graphics
			.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
			.DrawRectangle(Penna, 1, 1, 183, 189)
			'.DrawLine(Penna, 1, 220, 227, 220)

			If Stampa_Tipo <> "Trasferimento interno" Then
				' ODP
				.DrawRectangle(Penna, 90, 3, 92, 35)
				.DrawString(Stampa_Tipo, Carattere_Diciture, Brushes.Black, 92, 5)
				.DrawString(Stampa_ODP, Carattere_Matricola, Brushes.Black, Valore_centro(90, 92, Stampa_ODP, Carattere_Matricola, e, "X"), Valore_centro(3, 35, Stampa_ODP, Carattere_Matricola, e, "Y") + 5)


				'Matricola
				.DrawRectangle(Penna, 3, 114, 84, 40)
				.DrawString("Matricola", Carattere_Diciture, Brushes.Black, 6, 116)
				.DrawString(Stampa_Matricola, Carattere_Matricola, Brushes.Black, 6, 121)
				.DrawString(ODP_Form.ottieni_informazioni_odp("Numero", 0, Stampa_ODP).cliente, Carattere_Diciture, Brushes.Black, 6, 141)
				.DrawString(Stampa_Ubicazione_Macchina, Carattere_Matricola, Brushes.Black, Valore_centro(3, 84, Stampa_Ubicazione_Macchina, Carattere_Matricola, e, "X"), 135)

				'Posizione
				.DrawRectangle(Penna, 90, 114, 92, 40)
				.DrawString("Posizione", Carattere_Diciture, Brushes.Black, 93, 116)





				If ODP_Form.ottieni_informazioni_odp("Numero", 0, Stampa_ODP).u_produzione = "EST" Then
					.DrawString("E" & Stampa_progressivo_commessa, Carattere_posizione, Brushes.Black, Valore_centro(90, 92, Stampa_progressivo_commessa, Carattere_posizione, e, "X") - 35, Valore_centro(114, 40, Stampa_progressivo_commessa, Carattere_posizione, e, "Y") + 5)

				ElseIf ODP_Form.ottieni_informazioni_odp("Numero", 0, Stampa_ODP).u_produzione = "INT_SALD" Then
					.DrawString("S" & Stampa_progressivo_commessa, Carattere_posizione, Brushes.Black, Valore_centro(90, 92, Stampa_progressivo_commessa, Carattere_posizione, e, "X") - 35, Valore_centro(114, 40, Stampa_progressivo_commessa, Carattere_posizione, e, "Y") + 5)

				ElseIf ODP_Form.ottieni_informazioni_odp("Numero", 0, Stampa_ODP).u_produzione = "B_INT" Then
					.DrawString("B" & Stampa_progressivo_commessa, Carattere_posizione, Brushes.Black, Valore_centro(90, 92, Stampa_progressivo_commessa, Carattere_posizione, e, "X") - 35, Valore_centro(114, 40, Stampa_progressivo_commessa, Carattere_posizione, e, "Y") + 5)

				ElseIf ODP_Form.ottieni_informazioni_odp("Numero", 0, Stampa_ODP).u_produzione = "INT" Then
					.DrawString("I" & Stampa_progressivo_commessa, Carattere_posizione, Brushes.Black, Valore_centro(90, 92, Stampa_progressivo_commessa, Carattere_posizione, e, "X") - 35, Valore_centro(114, 40, Stampa_progressivo_commessa, Carattere_posizione, e, "Y") + 5)
				Else

					.DrawString(Stampa_progressivo_commessa, Carattere_posizione, Brushes.Black, Valore_centro(90, 92, Stampa_progressivo_commessa, Carattere_posizione, e, "X") - 35, Valore_centro(114, 40, Stampa_progressivo_commessa, Carattere_posizione, e, "Y") + 5)
				End If

				If ODP_Form.ottieni_informazioni_odp("Numero", 0, Stampa_ODP).fase = "QUADRO ELETTRICO" Then
					'And Form_layout_CAP_1.check_baia_layout(Stampa_Matricola).zona = "Magazzino" Then
					.DrawString("QE", Carattere_posizione, Brushes.Black, Valore_centro(90, 92, Stampa_progressivo_commessa, Carattere_posizione, e, "X"), Valore_centro(114, 40, Stampa_progressivo_commessa, Carattere_posizione, e, "Y") + 5)

				Else
					.DrawString(Form_layout_CAP_1.check_baia_layout(Stampa_Matricola).nome_baia, Carattere_posizione, Brushes.Black, Valore_centro(90, 92, Stampa_progressivo_commessa, Carattere_posizione, e, "X"), Valore_centro(114, 40, Stampa_progressivo_commessa, Carattere_posizione, e, "Y") + 5)
				End If




				'Descrizione_odp
				If Stampa_Descrizione = "INT_SALD" Or Stampa_Descrizione = "FORMATI_ATTREZZATURE" Then

						.DrawString(Stampa_Descrizione, Carattere_Qta, Brushes.Black, 7, 154)
					Else
						.DrawString(Stampa_Descrizione, Carattere_Descrizione, Brushes.Black, 7, 154)
					End If

					.DrawString(Stampa_Descrizione_2, Carattere_Descrizione, Brushes.Black, 7, 164)

				ElseIf magazzino_destinazione = "SCA" Or magazzino_destinazione = "BSCA" Or magazzino_destinazione = "03" Or magazzino_destinazione = "B03" Then
					.DrawRectangle(Penna, 90, 3, 92, 35)
					.DrawString("Ubicazione", Carattere_Diciture, Brushes.Black, 92, 5)
					.DrawString(Magazzino.OttieniDettagliAnagrafica(Codice_SAP).Ubicazione, Carattere_Matricola, Brushes.Black, Valore_centro(90, 92, Magazzino.OttieniDettagliAnagrafica(Codice_SAP).Ubicazione, Carattere_Matricola, e, "X"), Valore_centro(3, 35, Stampa_ODP, Carattere_Matricola, e, "Y") + 5)

				ElseIf magazzino_destinazione = "CQ" Then
					.DrawRectangle(Penna, 90, 3, 92, 35)
				Dim documento_scontrino As String
				If RadioButton1.Checked = True Then
					documento_scontrino = "EM"
				Else
					documento_scontrino = "TR"
				End If
				.DrawString(documento_scontrino & " " & Txt_DocNum.Text, Carattere_Diciture, Brushes.Black, 92, 5)
				.DrawString(Microsoft.VisualBasic.Left(Txt_Fornitore.Text, 28), Carattere_Diciture, Brushes.Black, 92, 15)

			End If



			'Articolo
			.DrawRectangle(Penna, 3, 40, 179, 71)
			.DrawString("Articolo", Carattere_Diciture, Brushes.Black, 4, 42)
			.DrawString(Stampa_Codice, Carattere_Codice, Brushes.Black, 4, 44)


			.DrawString(Stampa_Descrizione_Articolo, Carattere_Descrizione_Articolo, Brushes.Black, 4, 74)
			.DrawString(Stampa_Descrizione_Articolo_2, Carattere_Descrizione_Articolo, Brushes.Black, 4, 84)
			.DrawString(Stampa_Codice_Ordinazione, Carattere_Descrizione_Articolo, Brushes.Black, 4, 94)
			Dim numero_Qta As Decimal
			If Decimal.TryParse(Stampa_Qta.Replace(".", ","), numero_Qta) Then
				Stampa_Qta = Math.Round(numero_Qta).ToString("0")
			End If

			.DrawString(Stampa_Qta & " PZ", Carattere_Qta, Brushes.Black, 115, 42)


			'magazzino di destinazione
			.DrawRectangle(Penna, 3, 3, 85, 35)

			.DrawString("MAG Destinazione", Carattere_Diciture, Brushes.Black, 5, 5)

			'.DrawString(magazzino_destinazione.Substring(0, Math.Min(magazzino_destinazione.Length, 3)), Carattere_Matricola, Brushes.Black, Valore_centro(3, 85), Valore_centro(3, 35))
			.DrawString(magazzino_destinazione.Substring(0, Math.Min(magazzino_destinazione.Length, 4)), Carattere_Matricola, Brushes.Black, Valore_centro(3, 85, magazzino_destinazione.Substring(0, Math.Min(magazzino_destinazione.Length, 4)), Carattere_Matricola, e, "X"), Valore_centro(3, 35, magazzino_destinazione, Carattere_Matricola, e, "Y") + 5)


			If magazzino_destinazione = "15" Then


				'N° Ferretto
				.DrawRectangle(Penna, 3, 114, 84, 40)
				.DrawString("N° Fer", Carattere_Diciture, Brushes.Black, 6, 116)

				.DrawString(cassetto_del_codice(Stampa_Codice).numero_magazzino, Carattere_Matricola, Brushes.Black, 6, 121)

				'N° Cassetto
				.DrawRectangle(Penna, 90, 114, 92, 40)
				.DrawString("N° CAS", Carattere_Diciture, Brushes.Black, 93, 116)

				.DrawString(cassetto_del_codice(Stampa_Codice).numero_cassetto, Carattere_posizione, Brushes.Black, Valore_centro(90, 92, cassetto_del_codice(Stampa_Codice).numero_cassetto, Carattere_posizione, e, "X"), Valore_centro(114, 40, cassetto_del_codice(Stampa_Codice).numero_cassetto, Carattere_posizione, e, "Y") + 5)


			End If


			Dim Carattere_Data As New Font("Calibri", 7, FontStyle.Regular)
			Dim DataOggi As String = DateTime.Now.ToString("dd/MM/yy HH:mm")

			' Disegna la data e l'ora in basso a destra
			e.Graphics.DrawString(DataOggi, Carattere_Data, Brushes.Black, 120, 179)



			' Ottieni il nome del dipendente e troncalo a 12 caratteri
			Dim NomeDipendente As String = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).COGNOME & " " & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).NOME
			Dim NomeTroncato As String = NomeDipendente.Substring(0, Math.Min(12, NomeDipendente.Length))

			' Disegna il nome troncato in basso a destra
			e.Graphics.DrawString(NomeTroncato, Carattere_Data, Brushes.Black, 130, 169)

		End With
	End Sub

	Function Valore_centro(par_x_iniziale As Integer, par_lunghezza_x As Integer, par_contenuto As String, par_font_size As Font, e As System.Drawing.Printing.PrintPageEventArgs, par_asse As String)


		Dim stringSize As SizeF = e.Graphics.MeasureString(par_contenuto, par_font_size)
		Dim centro As Integer
		If par_asse = "X" Then
			centro = (par_x_iniziale + par_x_iniziale + par_lunghezza_x) / 2 - stringSize.Width / 2
		Else
			centro = (par_x_iniziale + par_x_iniziale + par_lunghezza_x) / 2 - stringSize.Height / 2
		End If

		Return centro
	End Function




	Private Sub Form_Entrate_Merci_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		MisuraTempo(AddressOf azzera_form, "azzera_form")
		MisuraTempo(Sub() N_codici_in_mag(codice_magazzino), "N_codici_in_mag")
		'MisuraTempo(AddressOf PErmanenza_media, "PErmanenza_media")
		MisuraTempo(AddressOf Aggiorna, "Aggiorna")
		riempi_datagridview_coda_em(DataGridView1, TextBox3.Text, TextBox4.Text)
		inserimento_tipologia_pacchi(ComboBox2)
	End Sub

	Private Sub MisuraTempo(action As Action, nomeSub As String)
		Dim stopwatch As New Stopwatch()
		stopwatch.Start()

		action.Invoke()

		stopwatch.Stop()
		Console.WriteLine($"{nomeSub} eseguita in {stopwatch.ElapsedMilliseconds} ms")
	End Sub

	Sub azzera_form()
		ComboBox1.SelectedIndex = 0


		Stampante_Selezionata = False
		Stampa_ODP = ""
		Stampa_Matricola = ""
		Stampa_Descrizione = ""
		Stampa_Descrizione_2 = ""
		Stampa_Codice = ""
		Stampa_Descrizione_Articolo = ""
		Stampa_Descrizione_Articolo_2 = ""
		Stampa_Qta = 0
		Stampa_Ubicazione_Macchina = ""
		Stampa_Codice_Ordinazione = ""
		Stampa_Tipo = "Ordine di Produzione"
	End Sub



	Private Function Get_Ubicazione_Matricola(Matricola As String) As String
		Dim Cnn_Matricola As New SqlConnection
		Dim Cmd_Matricola As New SqlCommand
		Dim Cmd_Matricola_Reader As SqlDataReader
		Dim Ubicazione As String

		Cnn_Matricola.ConnectionString = Homepage.sap_tirelli
		Cnn_Matricola.Open()
		Cmd_Matricola.Connection = Cnn_Matricola
		Cmd_Matricola.CommandText = "Select ItemCode
,case when U_Ubicazione is null then '' else U_Ubicazione end as 'Ubicazione' 
FROM OITM WHERE ItemCode = '" & Matricola & "'"
		Cmd_Matricola_Reader = Cmd_Matricola.ExecuteReader
		If Cmd_Matricola_Reader.Read() Then
			Ubicazione = Cmd_Matricola_Reader("Ubicazione")
		Else
			Ubicazione = "MAG"
		End If
		Cnn_Matricola.Close()
		Return Ubicazione
	End Function



	Sub N_codici_in_mag(par_magazzino_partenza As String)
		Dim Cnn_Matricola As New SqlConnection
		Dim Cmd_Matricola As New SqlCommand
		Dim Cmd_Matricola_Reader As SqlDataReader

		Cnn_Matricola.ConnectionString = Homepage.sap_tirelli
		Cnn_Matricola.Open()
		Cmd_Matricola.Connection = Cnn_Matricola
		Cmd_Matricola.CommandText = "declare @magazzino as varchar(50)
set @magazzino='" & par_magazzino_partenza & "'
SELECT COUNT(T0.ITEMCODE) AS 'N'
FROM OITW T0
WHERE T0.ONHAND >0 AND T0.WHSCODE=@magazzino and (substring(t0.itemcode,1,1)='0' or substring(t0.itemcode,1,1)='C' or substring(t0.itemcode,1,1)='D')"
		Cmd_Matricola_Reader = Cmd_Matricola.ExecuteReader
		If Cmd_Matricola_Reader.Read() Then
			Label1.Text = Cmd_Matricola_Reader("N")
		Else
			Label1.Text = 0

		End If
		Cnn_Matricola.Close()

	End Sub

	Sub PErmanenza_media()
		'		Dim Cnn_Matricola As New SqlConnection
		'		Dim Cmd_Matricola As New SqlCommand
		'		Dim Cmd_Matricola_Reader As SqlDataReader

		'		Cnn_Matricola.ConnectionString = Homepage.sap_tirelli
		'		Cnn_Matricola.Open()
		'		Cmd_Matricola.Connection = Cnn_Matricola
		'		Cmd_Matricola.CommandText = "declare @magazzino as varchar(50)
		'set @magazzino='" & codice_magazzino & "'

		'select cast(AVG(t10.giacente_da) as decimal ) as 'Giorni_medi'
		'from
		'(
		'select t0.itemcode, max(t0.docdate) as 'Ultima entrata', dbo.WorkDaysBetweenDates(max(t0.docdate),getdate()) as 'Giacente_da'
		'from oivl t0 inner join oitw t1 on t0.itemcode=t1.itemcode and t0.loccode=t1.WhsCode and t0.InQty>0 and t0.loccode=@magazzino
		'WHERE T1.ONHAND >0 AND T1.WHSCODE=@magazzino and t0.loccode=@magazzino and t0.InQty>0 and (substring(t0.itemcode,1,1)='0' or substring(t0.itemcode,1,1)='C' or substring(t0.itemcode,1,1)='D')

		'group by t0.itemcode
		')
		'as t10"
		'		Cmd_Matricola_Reader = Cmd_Matricola.ExecuteReader
		'		If Cmd_Matricola_Reader.Read() Then

		'			If Not Cmd_Matricola_Reader("Giorni_medi") Is System.DBNull.Value Then
		'				Label2.Text = Cmd_Matricola_Reader("Giorni_medi")
		'			Else
		'				Label2.Text = 0
		'			End If


		'		Else
		'			Label2.Text = 0

		'		End If
		'		Cnn_Matricola.Close()

	End Sub

	Private Function Get_Codice_Ordinazione(Codice As String) As String
		Dim Cnn_Codice As New SqlConnection
		Dim Cmd_Codice As New SqlCommand
		Dim Cmd_Codice_Reader As SqlDataReader
		Dim Codice_Ordinazione As String

		Cnn_Codice.ConnectionString = Homepage.sap_tirelli
		Cnn_Codice.Open()
		Cmd_Codice.Connection = Cnn_Codice
		Cmd_Codice.CommandText = "SELECT T0.[ItemCode], T0.[ItemName], T0.[SuppCatNum] as 'Codice', T1.[FirmName] as 'Fornitore' FROM OITM T0  INNER JOIN OMRC T1 ON T0.[FirmCode] = T1.[FirmCode] WHERE T0.[ItemCode] = '" & Codice & "'"
		Cmd_Codice_Reader = Cmd_Codice.ExecuteReader
		Cmd_Codice_Reader.Read()
		If Cmd_Codice_Reader.Read = True Then
			Codice_Ordinazione = Cmd_Codice_Reader("Fornitore") & " - " & Cmd_Codice_Reader("Codice")
		Else
			Codice_Ordinazione = ""
		End If

		Cnn_Codice.Close()
		Return Codice_Ordinazione
	End Function



	Public Sub Trasferimento_in_WIP(Documento As String, Codice As String, numero_odp As Integer, numero_oc As Integer, Qta As String, Mag_Partenza As String, Mag_Destinazione As String, linenum_ODP As String, Matricola As String, par_utente_sap As String, Par_automatismo As String, par_docentry_rt As Integer, par_riga As Integer, par_tipo_trasferimento As String)
		trasferimento_eseguito = "NO"
		If Mag_Partenza = Mag_Destinazione Then
			MsgBox("Il magazzino di partenza non può essere = al magazzino destinazione")
			Return
		End If
		If Codice = codice_precedente Then

		Else
			quantità_predente = 0
		End If

		If Documento = "ODP" Then
			numero_oc_precedente = 0
			If numero_odp = numero_odp_precedente And Codice = codice_precedente Then
				MsgBox("Non è possibile fare due trasferimenti di fila sullo stesso ordine di produzione. Chiudere il programma e riaprirlo")
				Return
			End If

		ElseIf Documento = "OC" Then
			numero_odp_precedente = 0
			If numero_oc = numero_oc_precedente And Codice = codice_precedente Then
				MsgBox("Non è possibile fare due trasferimenti di fila sullo stesso ordine cliente. Chiudere il programma e riaprirlo")
				Return
			End If
		End If
		Dim new_form_Magazzino = New Magazzino
		new_form_Magazzino.ripristino_giacenze_corrette(Codice)
		Qta = Replace(Qta, ",", ".")
		If check_giacenza_per_trasferimento(Codice, Mag_Partenza, Qta) = 1 Then

			MsgBox("Si sta cercando di trasferire una quantità maggiore della giacenza OITW")


			Return

		End If

		If check_giacenza_per_trasferimento_trasferimenti_oivl(Codice, Mag_Partenza, Qta) = 1 Then
			MsgBox("Si sta cercando di trasferire una quantità maggiore della giacenza Oivl")
			Return
		End If

		If Magazzino.OttieniDettagliAnagrafica(Codice).attivo = "N" Then
			MsgBox("L'articolo " & Codice & " Risulta inattivo")
		End If



		new_form_Magazzino.Codice_SAP = Codice

		Dim par_ref_1 As String
		If Documento = "ODP" Then
			par_ref_1 = numero_odp
		ElseIf Documento = "OC" Then
			par_ref_1 = numero_oc
		End If
		Dim ultimo_docentry As Integer
		ultimo_docentry = new_form_Magazzino.DOCENTRY_Trasferimenti()
		If par_riga = 0 Then

			new_form_Magazzino.Inserisci_documento_trasferimento(ultimo_docentry, new_form_Magazzino.DOCNUM_Trasferimenti(), Documento, numero_odp, numero_oc, Qta, Mag_Partenza, Mag_Destinazione, par_utente_sap, par_docentry_rt, Replace(Magazzino.OttieniDettagliAnagrafica(Codice).Prezzo_listino_acquisto, ",", "."), par_tipo_trasferimento, par_ref_1)
			new_form_Magazzino.aggiorna_NNM1_trasferimento()
			new_form_Magazzino.AGGIUSTA_docentry()
		End If

		new_form_Magazzino.Inserisci_righe_trasferimento(ultimo_docentry + 1, Documento, numero_odp, numero_oc, Codice, Qta, Mag_Partenza, Mag_Destinazione, linenum_ODP, Replace(Magazzino.OttieniDettagliAnagrafica(Codice).Prezzo_listino_acquisto, ",", "."), Magazzino.OttieniDettagliAnagrafica(Codice).Descrizione, par_riga)


		Try
			new_form_Magazzino.Trova_NUMERATORE_OIVL()
			new_form_Magazzino.Trova_message_id()
		Catch ex As Exception
			delete_owtr_wtr1(ultimo_docentry + 1)
			new_form_Magazzino.AGGIUSTA_docentry()
			Return
		End Try

		Dim messageid_oilm As Integer = new_form_Magazzino.MESSAGEID + 1
		Dim messageid_oilm_1 As Integer = new_form_Magazzino.MESSAGEID + 2
		Try

			new_form_Magazzino.OILM(Documento, Codice, Qta, Mag_Partenza, Mag_Destinazione, par_tipo_trasferimento, Matricola, par_utente_sap, Replace(Magazzino.OttieniDettagliAnagrafica(Codice).Prezzo_listino_acquisto, ",", "."), Magazzino.OttieniDettagliAnagrafica(Codice).Descrizione, par_ref_1)

		Catch ex As Exception
			delete_owtr_wtr1(ultimo_docentry + 1)
			new_form_Magazzino.AGGIUSTA_docentry()
			Return
		End Try

		Try
			new_form_Magazzino.OIVL_IVL1_OIVK(Codice, Qta, Mag_Partenza, Mag_Destinazione, par_utente_sap, Replace(Magazzino.OttieniDettagliAnagrafica(Codice).Prezzo_listino_acquisto, ",", "."))

		Catch ex As Exception
			delete_owtr_wtr1(ultimo_docentry + 1)
			new_form_Magazzino.AGGIUSTA_docentry()
			delete_oilm(messageid_oilm)
			delete_oilm(messageid_oilm_1)
			Return
		End Try

		'new_form_Magazzino.IVL1(Qta)
		'new_form_Magazzino.OIVK()
		new_form_Magazzino.Aggiusta_numeratore_messageid()
		If par_riga = 0 Then
			new_form_Magazzino.aggiusta_Numeratore_OIVL()
		End If

		new_form_Magazzino.aggiorna_OITW(Qta, Mag_Partenza, Mag_Destinazione, Codice)
		new_form_Magazzino.aggiorna_da_trasferire(Documento, Qta, linenum_ODP, numero_odp, numero_oc, par_tipo_trasferimento)

		new_form_Magazzino.metto_wip_nel_magazzino_riga(Documento, Mag_Destinazione, linenum_ODP, numero_odp, numero_oc)



		Acquisti.aggiusta_CONFERMATO(Codice)
		If par_riga = 0 Then
			If Documento = "ODP" Then
				new_form_Magazzino.AWOR(numero_odp, linenum_ODP, par_utente_sap)
			End If
		End If

		If par_riga = 0 Then
			numero_odp_precedente = numero_odp
			numero_oc_precedente = numero_oc
			codice_precedente = Codice
			quantità_predente = Qta
		End If
		trasferimento_eseguito = "SI"
	End Sub

	Sub delete_owtr_wtr1(par_docentry As Integer)
		Dim CNN As New SqlConnection
		CNN.ConnectionString = Homepage.sap_tirelli
		CNN.Open()
		Dim CMD_SAP As New SqlCommand
		CMD_SAP.Connection = CNN

		CMD_SAP.CommandText = "delete owtr where docentry='" & par_docentry & "'
delete wtr1 where docentry='" & par_docentry & "'"
		CMD_SAP.ExecuteNonQuery()
		CNN.Close()
	End Sub

	Sub delete_oilm(par_MESSAGEID As Integer)
		Dim CNN As New SqlConnection
		CNN.ConnectionString = Homepage.sap_tirelli
		CNN.Open()
		Dim CMD_SAP As New SqlCommand
		CMD_SAP.Connection = CNN

		CMD_SAP.CommandText = "DELETE OILM WHERE OILM.MessageID=" & par_MESSAGEID & ""
		CMD_SAP.ExecuteNonQuery()
		CNN.Close()
	End Sub

	Public Function check_giacenza_per_trasferimento(par_itemcode As String, par_magazzino As String, par_quantità As String)


		par_quantità = Replace(par_quantità, ",", ".")
		Dim errore As Integer = 0

		Dim Cnn_Matricola As New SqlConnection
		Dim Cmd_Matricola As New SqlCommand
		Dim Cmd_Matricola_Reader As SqlDataReader

		Cnn_Matricola.ConnectionString = Homepage.sap_tirelli
		Cnn_Matricola.Open()
		Cmd_Matricola.Connection = Cnn_Matricola
		Cmd_Matricola.CommandText = "select t0.itemcode
,coalesce(t1.onhand,0) as 'onhand',
case when " & par_quantità & "> coalesce(t1.onhand,0)  then 1 else 0 end as 'Errore'
from oitm t0 inner join oitw t1 on t0.itemcode=t1.itemcode and t1.whscode='" & par_magazzino & "'
WHERE T0.ITEMCODE='" & par_itemcode & "'"
		Cmd_Matricola_Reader = Cmd_Matricola.ExecuteReader
		If Cmd_Matricola_Reader.Read() Then

			If Cmd_Matricola_Reader("errore") = 1 Then

				errore = 1
			Else
				errore = 0


			End If
		Else
			errore = 1

		End If
		Cmd_Matricola_Reader.Close()
		Cnn_Matricola.Close()
		Return errore
	End Function

	Public Function Trova_info_odp(par_docnum As Integer)

		Dim magazzino_destinazione_odp As String = ""



		Dim Cnn_Matricola As New SqlConnection
		Dim Cmd_Matricola As New SqlCommand
		Dim Cmd_Matricola_Reader As SqlDataReader

		Cnn_Matricola.ConnectionString = Homepage.sap_tirelli
		Cnn_Matricola.Open()
		Cmd_Matricola.Connection = Cnn_Matricola
		Cmd_Matricola.CommandText = "select t0.warehouse 
from owor t0
where t0.docnum=" & par_docnum & ""

		Cmd_Matricola_Reader = Cmd_Matricola.ExecuteReader
		If Cmd_Matricola_Reader.Read() Then



			magazzino_destinazione_odp = Cmd_Matricola_Reader("warehouse")



		End If
		Cmd_Matricola_Reader.Close()
		Cnn_Matricola.Close()
		Return magazzino_destinazione_odp
	End Function

	Public Function Trova_info_RIGA_OC(par_docnum As Integer, PAR_LINENUM As Integer)

		Dim Regola_distribuzione As String = ""



		Dim Cnn_Matricola As New SqlConnection
		Dim Cmd_Matricola As New SqlCommand
		Dim Cmd_Matricola_Reader As SqlDataReader

		Cnn_Matricola.ConnectionString = Homepage.sap_tirelli
		Cnn_Matricola.Open()
		Cmd_Matricola.Connection = Cnn_Matricola
		Cmd_Matricola.CommandText = "select t1.ocrcode
from ordr t0 inner join rdr1 t1 on t0.docentry=t1.docentry
where t0.docnum=" & par_docnum & " and t1.linenum =" & PAR_LINENUM & ""

		Cmd_Matricola_Reader = Cmd_Matricola.ExecuteReader
		If Cmd_Matricola_Reader.Read() Then



			Regola_distribuzione = Cmd_Matricola_Reader("ocrcode")



		End If
		Cmd_Matricola_Reader.Close()
		Cnn_Matricola.Close()
		Return Regola_distribuzione
	End Function



	Public Function check_giacenza_per_trasferimento_trasferimenti_oivl(par_itemcode As String, par_magazzino As String, par_quantità As String)

		par_quantità = Replace(par_quantità, ",", ".")
		Dim errore As Integer = 0

		Dim Cnn_Matricola As New SqlConnection
		Dim Cmd_Matricola As New SqlCommand
		Dim Cmd_Matricola_Reader As SqlDataReader

		Cnn_Matricola.ConnectionString = Homepage.sap_tirelli
		Cnn_Matricola.Open()
		Cmd_Matricola.Connection = Cnn_Matricola
		Cmd_Matricola.CommandText = "
select itemcode, loccode,case when " & par_quantità & " > sum(coalesce(inqty,0)-coalesce(outqty,0) ) then 1 else 0 end as 'Errore'
from oivl
where loccode='" & par_magazzino & "' and itemcode='" & par_itemcode & "'
group by itemcode, loccode"
		Cmd_Matricola_Reader = Cmd_Matricola.ExecuteReader
		If Cmd_Matricola_Reader.Read() Then



			If Cmd_Matricola_Reader("errore") = 1 Then

				errore = 1
			Else
				errore = 0


			End If

		End If
		Cmd_Matricola_Reader.Close()
		Cnn_Matricola.Close()
		Return errore
	End Function

	Sub INSERISCI_LOG_ERRORI_FERRETTO(PAR_ITEMCODE As String, docentry_rt As Integer, par_motivo As String)


		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()
		Dim CMD_SAP As New SqlCommand
		CMD_SAP.Connection = Cnn

		CMD_SAP.CommandText = "insert into LOG_ERRORI_FERRETTO (ID,ITEMCODE,DOCENTRY_RT,MOTIVO) VALUES (1, '" & PAR_ITEMCODE & "'," & docentry_rt & ",'" & par_motivo & "'"
		CMD_SAP.ExecuteNonQuery()
		Cnn.Close()




	End Sub

	Sub Trasferisci_ad_altro_magazzino(Codice As String, Qta As String, numero_odp As Integer, numero_oc As Integer, Mag_Partenza As String, Mag_Destinazione As String, linenum_odp As String, par_utente_sap As String, par_tipo_documento As String, par_numero_documento As Integer)

		If Codice = codice_precedente Then

		Else
			quantità_predente = 0
		End If
		Dim new_form_Magazzino = New Magazzino
		new_form_Magazzino.ripristino_giacenze_corrette(Codice)
		If check_giacenza_per_trasferimento(Codice, Mag_Partenza, Qta) = 1 Then

			MsgBox("Si sta cercando di trasferire una quantità maggiore della giacenza OITW")


			Return

		End If

		If check_giacenza_per_trasferimento_trasferimenti_oivl(Codice, Mag_Partenza, Qta) = 1 Then
			MsgBox("Si sta cercando di trasferire una quantità maggiore della giacenza Ovl")
			Return
		End If


		new_form_Magazzino.Codice_SAP = Codice
		'new_form_Magazzino.Trova_serie()
		new_form_Magazzino.Trova_PERIODO_contabile()
		new_form_Magazzino.DOCENTRY_Trasferimenti()

		If Mag_Destinazione = "06" And par_tipo_documento = "ODP" Then
			new_form_Magazzino.aggiorna_qta_richiesta_per_wip(par_tipo_documento, par_numero_documento, Qta, linenum_odp)

		ElseIf magazzino_destinazione = "06" And par_tipo_documento = "" Then

			MsgBox("Trasferire il pezzo riferendosi all'odp")
			Return

		End If


		new_form_Magazzino.DOCNUM_Trasferimenti()
		'new_form_Magazzino.dettagli_anagrafica(Codice)
		'new_form_Magazzino.stringa_trasferimento = "Trasferimento a magazzino " & Mag_Destinazione
		Dim ULTIMO_DOCENTRY As Integer
		ULTIMO_DOCENTRY = Magazzino.DOCENTRY_Trasferimenti()
		new_form_Magazzino.Inserisci_documento_trasferimento(ULTIMO_DOCENTRY, Magazzino.DOCNUM_Trasferimenti(), Documento, numero_odp, numero_oc, Qta, Mag_Partenza, Mag_Destinazione, par_utente_sap, 0, Magazzino.OttieniDettagliAnagrafica(Codice).Prezzo_listino_acquisto, "Trasferimento a magazzino " & Mag_Destinazione, "")
		new_form_Magazzino.Inserisci_righe_trasferimento(ULTIMO_DOCENTRY + 1, Documento, numero_odp, numero_oc, Codice, Qta, Mag_Partenza, Mag_Destinazione, linenum_odp, Magazzino.OttieniDettagliAnagrafica(Codice).Prezzo_listino_acquisto, Magazzino.OttieniDettagliAnagrafica(Codice).Descrizione, 0)
		new_form_Magazzino.aggiorna_NNM1_trasferimento()
		new_form_Magazzino.AGGIUSTA_docentry()

		'new_form_Magazzino.Business_partner_della_commessa()
		new_form_Magazzino.Trova_NUMERATORE_OIVL()
		new_form_Magazzino.Trova_message_id()
		new_form_Magazzino.OILM(Documento, Codice, Qta, Mag_Partenza, Mag_Destinazione, "Trasferimento", "", par_utente_sap, Replace(Magazzino.OttieniDettagliAnagrafica(Codice).Prezzo_listino_acquisto, ",", "."), Magazzino.OttieniDettagliAnagrafica(Codice).Descrizione, "")
		new_form_Magazzino.OIVL_IVL1_OIVK(Codice, Qta, Mag_Partenza, Mag_Destinazione, par_utente_sap, Replace(Magazzino.OttieniDettagliAnagrafica(Codice).Prezzo_listino_acquisto, ",", "."))
		'new_form_Magazzino.IVL1(Qta)
		'new_form_Magazzino.OIVK()
		new_form_Magazzino.aggiusta_Numeratore_OIVL()
		new_form_Magazzino.Aggiusta_numeratore_messageid()



		new_form_Magazzino.aggiorna_OITW(Qta, Mag_Partenza, Mag_Destinazione, Codice)



		Acquisti.aggiusta_CONFERMATO(Codice)

		codice_precedente = Codice
		quantità_predente = Qta

	End Sub



	Private Sub Txt_Codice_TextChanged(sender As Object, e As EventArgs)
		Magazzino.OttieniDettagliAnagrafica(Button7.Text)
	End Sub

	Sub cambia_stato(par_docnum_odp As String)



		nuovo_stato = "R"
		Try
			DataGridView2.Rows(Riga).Cells(columnName:="Stato").Value = nuovo_stato
		Catch ex As Exception

		End Try

		Dim Cnn As New SqlConnection

		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()
		Dim CMD_SAP As New SqlCommand
		CMD_SAP.Connection = Cnn

		CMD_SAP.CommandText = "UPDATE owor SET STATUS='" & nuovo_stato & "' WHERE DOCNUM ='" & par_docnum_odp & "'"
		CMD_SAP.ExecuteNonQuery()
		Cnn.Close()




	End Sub


	Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
		Me.Close()
	End Sub



	Sub stampa_odp_foglio(docnum_odp As String)
		ODP_Form.docnum_odp = docnum_odp
		ODP_Form.stampa_etichetta = "no"


		ODP_Form.percorso_documento = Homepage.PERCORSO_DOCUMENTO_ODP_ETICHETTA
		ODP_Form.stampa_etichetta = "YES"
		ODP_Form.Genera_ordine()



	End Sub

	Private Sub DataGrid_EM_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGrid_EM.CellClick

		If e.RowIndex >= 0 Then
			riga_datagrid = e.RowIndex
			Codice_SAP = DataGrid_EM.Rows(e.RowIndex).Cells(columnName:="Codice").Value


			If e.ColumnIndex = DataGrid_EM.Columns.IndexOf(Codice) Then
				Button7.Text = Codice_SAP
				Button5.Text = DataGrid_EM.Rows(e.RowIndex).Cells(columnName:="Cmd_Disegno").Value

				'ControllaEVisualizzaPDF(CheckBox2, Codice_SAP, WebBrowser1)

				'Magazzino.dettagli_anagrafica(Codice_SAP)
				Txt_Descrizione.Text = Magazzino.OttieniDettagliAnagrafica(Codice_SAP).Descrizione
				Label4.Text = Magazzino.OttieniDettagliAnagrafica(Codice_SAP).Descrizione_SUP
				Label5.Text = Magazzino.OttieniDettagliAnagrafica(Codice_SAP).Catalogo
				Label3.Text = Magazzino.OttieniDettagliAnagrafica(Codice_SAP).Ubicazione

				Txt_Magazzino_01.Text = giacenze_IN_magazzino(Codice_SAP, codice_magazzino)



				trasferito(Codice_SAP, DataGridView2)
				giacenze_magazzino(Codice_SAP)
				Txt_q_em.Text = DataGrid_EM.Rows(riga_datagrid).Cells(columnName:="Qta").Value
				'giacenze_IN_magazzino_PER_TRASFERIMENTO(Codice_SAP, codice_magazzino)

				Txt_q_em.Text = DataGrid_EM.Rows(riga_datagrid).Cells(columnName:="Qta").Value

				Txt_Magazzino_01.Text = Math.Round(giacenze_IN_magazzino(Codice_SAP, codice_magazzino), 3)

				Try
					If Decimal.Parse(Txt_q_em.Text) <= Decimal.Parse(Txt_Magazzino_01.Text) Then

						Txt_trasferibile.Text = Txt_q_em.Text
					Else

						Txt_trasferibile.Text = Txt_Magazzino_01.Text
					End If
				Catch ex As Exception
					Txt_trasferibile.Text = 0
				End Try



			ElseIf e.ColumnIndex = DataGrid_EM.Columns.IndexOf(Cmd_Disegno) And DataGrid_EM.Rows(e.RowIndex).Cells(columnName:="Cmd_Disegno").Value.ToString.Length > 0 Then
				'Visualizza Disegno
				Dim num_foglio As Integer = 1
				If File.Exists(Homepage.percorso_disegni_generico & "PDF\" & DataGrid_EM.Rows(e.RowIndex).Cells(2).Value.ToString & ".PDF") Then
					Process.Start(Homepage.percorso_disegni_generico & "PDF\" & DataGrid_EM.Rows(e.RowIndex).Cells(2).Value.ToString & ".PDF")
				ElseIf File.Exists(Homepage.percorso_disegni_generico & "PDF\" & DataGrid_EM.Rows(e.RowIndex).Cells(2).Value.ToString & "_foglio_" & num_foglio & ".PDF") Then
					Process.Start(Homepage.percorso_disegni_generico & "PDF\" & DataGrid_EM.Rows(e.RowIndex).Cells(2).Value.ToString & "_foglio_" & num_foglio & ".PDF")
				Else
					MsgBox("PDF non trovato")
				End If



			ElseIf e.ColumnIndex = DataGrid_EM.Columns.IndexOf(Stampa_Codice_EM) Then
				Stampa_Matricola = "------"
				Stampa_Descrizione = "----------------------------"
				Stampa_Descrizione_2 = "----------------------------"
				Stampa_Codice = DataGrid_EM.Rows(e.RowIndex).Cells(columnName:="Codice").Value
				Stampa_Descrizione_Articolo = Microsoft.VisualBasic.Left(DataGrid_EM.Rows(e.RowIndex).Cells(columnName:="Descrizione").Value, 28)
				Stampa_Descrizione_Articolo_2 = Microsoft.VisualBasic.Mid(DataGrid_EM.Rows(e.RowIndex).Cells(columnName:="descrizione").Value, 29, 56)
				Stampa_Qta = DataGrid_EM.Rows(e.RowIndex).Cells(columnName:="In_mag_accettazione").Value
				Stampa_Ubicazione_Macchina = DataGrid_EM.Rows(e.RowIndex).Cells(columnName:="Ubicazione").Value

				If Stampa_Ubicazione_Macchina = "FER" Then
					Stampa_ODP = "FERRETTO"
				Else
					If Stampa_Ubicazione_Macchina = "XXX" Then
						Stampa_ODP = "UBICARE"
					Else
						Stampa_ODP = "SCAFFALE"
					End If
				End If
				Stampa_Codice_Ordinazione = Get_Codice_Ordinazione(Stampa_Codice)
				Stampa_Tipo = "Refilling"
				Fun_Stampa(Stampa_Tipo, CheckBox1.Checked, Stampante_Selezionata, Scontrino, "", "", Stampa_Codice)
			End If

			If Magazzino.OttieniDettagliAnagrafica(Codice_SAP).Soggetto_collaudo = "Y" And RadioButton1.Checked = True And codice_magazzino <> "B01" Then

				MsgBox("ARTICOLO SOGGETTO A CONTROLLO TRASFERIRLO IN CQ")
				DataGridView2.Visible = False
				Button1.Visible = False
				Button2.Visible = False

				Button8.Visible = False
				Button9.Visible = False
				Button3.Visible = True

			ElseIf Magazzino.OttieniDettagliAnagrafica(Codice_SAP).Soggetto_collaudo = "Y" And RadioButton1.Checked = False And (Txt_Fornitore.Text = "CAP2" Or Txt_Fornitore.Text = "09") And codice_magazzino <> "B01" Then

				MsgBox("ARTICOLO SOGGETTO A CONTROLLO TRASFERIRLO IN CQ _1")
				DataGridView2.Visible = False
				Button1.Visible = False
				Button2.Visible = False

				Button8.Visible = False
				Button9.Visible = False
				Button3.Visible = True


				'	Inserire controllo per i trasferimenti proventienti da Clavter ed il fornitore è soggetto a 


			ElseIf Soggetto_controllo_bp = "Y" And RadioButton1.Checked = True And codice_magazzino <> "B01" Then
				MsgBox("BUSINESS PARTNER SOGGETTO A CONTROLLO TRASFERIRLO IN CQ_ 2")
				DataGridView2.Visible = False
				Button1.Visible = False
				Button2.Visible = False

				Button8.Visible = False
				Button9.Visible = False
				Button3.Visible = True
			Else
				DataGridView2.Visible = True
				Button1.Visible = True
				Button2.Visible = True

				Button8.Visible = True
				Button9.Visible = True
			End If

		End If
	End Sub

	Public Async Sub ControllaEVisualizzaPDF(par_checkbox As CheckBox, par_disegno As String, par_web_browser As WebBrowser)
		If par_checkbox.Checked Then
			Await VisualizzaPDFInBackground(Homepage.percorso_disegni_generico, par_disegno, par_web_browser)
		End If
	End Sub

	Public Async Function VisualizzaPDFInBackground(par_percorso_base As String, par_disegno As String, par_web_browser As WebBrowser) As Task
		' Componi il percorso del file PDF
		Dim pdfPath As String = par_percorso_base & "PDF\" & par_disegno & ".PDF"

		' Verifica l'esistenza del file in modo asincrono
		Dim fileExists As Boolean = Await Task.Run(Function() File.Exists(pdfPath))

		If fileExists Then
			' Aggiungi i parametri per nascondere barra degli strumenti e pannelli laterali
			pdfPath &= "#toolbar=0&zoom=100&navpanes=0"

			' Naviga al PDF nel WebBrowser sul thread principale
			' Questa operazione potrebbe bloccarsi se il WebBrowser è occupato
			Await Task.Run(Sub()
							   ' Questa parte non deve eseguire interazioni dirette con l'UI
							   ' Esegui l'operazione di navigazione sul thread principale
							   Invoke(Sub() par_web_browser.Navigate(pdfPath))
						   End Sub)
		End If
	End Function




	Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click

		Dim par_magazzino_destinazione As String
		'If Homepage.Centro_di_costo = "BRB01" Then
		'	par_magazzino_destinazione = "B16"
		'Else

		'	par_magazzino_destinazione = "16"
		'End If

		par_magazzino_destinazione = "16"

		trasferimento_altro_magazzino_DEF(Codice_SAP, codice_magazzino, tipo_entrata, Txt_DocNum.Text, CheckBox1.Checked, par_magazzino_destinazione)

		aggiorna_datagridivew_codici(trasferimento_eseguito)





	End Sub

	Sub TRASFERIMENTO_AD_ALTRO_MAGAZZINO(par_codice_sap As String, par_magazzino_partenza As String, MAGAZZINO_DESTINAZIONE As String, par_quantità As Decimal, par_giacenza_01 As String, par_quantità_trasferita_scontrino As String)


		If par_codice_sap = "" Then
			MsgBox("Selezionare un codice")

		Else
			Dim result As Integer

			result = MessageBox.Show("Mandare il codice " & par_codice_sap & " al magazzino  " & MAGAZZINO_DESTINAZIONE, "Confirmation", MessageBoxButtons.YesNo)
			If result = DialogResult.Yes Then





				Trasferisci_ad_altro_magazzino(par_codice_sap, Replace(par_quantità, ",", "."), 0, 0, par_magazzino_partenza, MAGAZZINO_DESTINAZIONE, "", Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, "", 0)
				par_giacenza_01 = par_giacenza_01 - par_quantità
				par_quantità_trasferita_scontrino = Replace(par_quantità, ",", ".")
				CREA_SCONTRINO = "Y"

				giacenze_magazzino(par_codice_sap)

				N_codici_in_mag(par_magazzino_partenza)
				MsgBox("Trasferimento di " & Replace(par_quantità, ",", ".") & " in " & MAGAZZINO_DESTINAZIONE & " Effettuato con successo")
				trasferimento_eseguito = "SI"
			Else
				trasferimento_eseguito = "NO"
				MsgBox("nulla")


			End If

		End If

	End Sub



	Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
		If inizializzazione_form = False Then
			Aggiorna()
			TableLayoutPanel1.BackColor = Color.SteelBlue
		End If

		If RadioButton1.Checked = True Then
			tipo_entrata = "EM"
		End If

	End Sub

	Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
		If inizializzazione_form = False Then
			Aggiorna()
			TableLayoutPanel1.BackColor = Color.YellowGreen
		End If

		If RadioButton2.Checked = True Then
			tipo_entrata = "TRASF"
		End If
	End Sub

	Sub compila_scontrino(par_articolo As String, par_descrizione_articolo As String, par_quantità As String, par_commessa As String)


		Stampa_ODP = ""
		Stampa_Matricola = ""
		Stampa_Descrizione = Microsoft.VisualBasic.Left("", 28)
		If Magazzino.OttieniDettagliAnagrafica(par_articolo).codice_brb = "" Then
			Stampa_Descrizione_2 = Microsoft.VisualBasic.Mid("", 29, 66)
		Else
			Stampa_Descrizione_2 = Magazzino.OttieniDettagliAnagrafica(par_articolo).codice_brb

		End If

		Stampa_Codice = par_articolo
		Stampa_Descrizione_Articolo = Microsoft.VisualBasic.Left(par_descrizione_articolo, 28)
		Stampa_Descrizione_Articolo_2 = Microsoft.VisualBasic.Mid(par_descrizione_articolo, 29, 56)

		Stampa_Qta = par_quantità



		Stampa_Ubicazione_Macchina = Get_Ubicazione_Matricola(par_commessa)

		Stampa_Codice_Ordinazione = Get_Codice_Ordinazione(par_articolo)
	End Sub


	Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
		Dim par_magazzino_destinazione As String
		' Controllo sulla DataGridView2
		Dim countMaggioreDiZero As Integer = 0

		For Each row As DataGridViewRow In DataGridView2.Rows
			If Not row.IsNewRow Then
				Dim valore As Integer
				If Integer.TryParse(row.Cells("da_tras").Value?.ToString(), valore) AndAlso valore > 0 Then
					countMaggioreDiZero += 1
				End If
			End If
		Next

		' Se c'è un solo record con valore > 0, chiedi conferma all'utente
		If countMaggioreDiZero = 1 Then
			Dim result As DialogResult = MessageBox.Show("Esiste la possibilità di trasferire a WIP l'articolo, è fortemente sconsigliato proseguire con questa operazione. Proseguire?",
													 "Attenzione", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

			If result = DialogResult.No Then
				Exit Sub
			End If
		End If


		par_magazzino_destinazione = "SCA"

		trasferimento_altro_magazzino_DEF(Codice_SAP, codice_magazzino, tipo_entrata, Txt_DocNum.Text, CheckBox1.Checked, par_magazzino_destinazione)

		aggiorna_datagridivew_codici(trasferimento_eseguito)

	End Sub

	Sub aggiorna_datagridivew_codici(par_trasferimento_eseguito As String)
		Select Case True
			Case RadioButton1.Checked
				Aggiorna_EM(Txt_DocNum.Text)
			Case RadioButton2.Checked
				Aggiorna_trasferimento()
			Case RadioButton3.Checked
				Aggiorna_EMP()
			Case RadioButton4.Checked
				Aggiorna_CS(Txt_DocNum.Text)
		End Select
	End Sub

	Sub trasferimento_altro_magazzino_DEF(par_codice_sap As String, par_magazzino_da As String, par_tipo_entrata As String, par_numero_documento As String, par_preview_scontrino As Boolean, par_magazzino_destinazione As String)
		CREA_SCONTRINO = "Y"
		Stampa_Tipo = "Trasferimento interno"


		Dim qta_a_mag As Decimal = giacenze_IN_magazzino(par_codice_sap, par_magazzino_da)

		Dim qta_entrata_con_em As Decimal = quantita_entrata_con_em(par_numero_documento, par_codice_sap, par_magazzino_da, par_tipo_entrata)
		Dim qta_trasferimento As Decimal = Form_lotto_di_prelievo.minore(qta_a_mag, qta_entrata_con_em)
		qta_trasferimento = Form_lotto_di_prelievo.minore(qta_trasferimento, qta_entrata_con_em)




		If (Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Ubicazione = "" Or Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Ubicazione = Nothing) And (par_magazzino_destinazione = "SCA" Or par_magazzino_destinazione = "BSCA" Or par_magazzino_destinazione = "B03" Or par_magazzino_destinazione = "03") Then
			MsgBox("Non è possibile trasferire a " & par_magazzino_destinazione & " senza che il codice abbia UBICAZIONE")
		Else
			'giacenze_IN_magazzino_PER_TRASFERIMENTO(Codice_SAP, magazzino_destinazione)
			TRASFERIMENTO_AD_ALTRO_MAGAZZINO(par_codice_sap, par_magazzino_da, par_magazzino_destinazione, qta_trasferimento, giacenze_IN_magazzino(par_codice_sap, par_magazzino_da), qta_trasferimento)
			If trasferimento_eseguito = "SI" Then


				If CREA_SCONTRINO = "Y" Then
					compila_scontrino(par_codice_sap, Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, qta_trasferimento, "")
					Fun_Stampa(Stampa_Tipo, par_preview_scontrino, Stampante_Selezionata, Scontrino, par_magazzino_destinazione, Magazzino.OttieniDettagliAnagrafica(Codice_SAP).Ubicazione, par_codice_sap)
				End If
			End If
		End If

	End Sub

	Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
		Dim par_magazzino_destinazione As String

		par_magazzino_destinazione = "CQ"
		trasferimento_altro_magazzino_DEF(Codice_SAP, codice_magazzino, tipo_entrata, Txt_DocNum.Text, CheckBox1.Checked, par_magazzino_destinazione)
		aggiorna_datagridivew_codici(trasferimento_eseguito)



	End Sub





	Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

		Magazzino.nuovo_valore_string = InputBox("Inserire nuova ubicazione")
		Magazzino.cambiare_gestione_ubicazione(Codice_SAP, Magazzino.nuovo_valore_string, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
		Magazzino.OttieniDettagliAnagrafica(Codice_SAP)
		Label3.Text = Magazzino.OttieniDettagliAnagrafica(Codice_SAP).Ubicazione

	End Sub

	Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

		Magazzino.Codice_SAP = Button7.Text
		' Ripristina la finestra se è minimizzata
		If Magazzino.WindowState = FormWindowState.Minimized Then
			Magazzino.WindowState = FormWindowState.Normal
		End If

		' Porta la finestra in primo piano
		Magazzino.BringToFront()
		Magazzino.Activate()
		Magazzino.Show()

		Magazzino.TextBox2.Text = Magazzino.Codice_SAP
		Magazzino.OttieniDettagliAnagrafica(Codice_SAP)
	End Sub



	Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
		Dim par_magazzino_destinazione As String
		'If Homepage.Centro_di_costo = "BRB01" Then
		'	par_magazzino_destinazione = "B03"

		'Else
		'	par_magazzino_destinazione = "03"
		'End If
		par_magazzino_destinazione = "03"

		trasferimento_altro_magazzino_DEF(Codice_SAP, codice_magazzino, tipo_entrata, Txt_DocNum.Text, CheckBox1.Checked, par_magazzino_destinazione)
		aggiorna_datagridivew_codici(trasferimento_eseguito)
	End Sub

	Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
		Dim par_magazzino_destinazione As String
		par_magazzino_destinazione = "15"

		' Controllo sulla DataGridView2
		Dim countMaggioreDiZero As Integer = 0

		For Each row As DataGridViewRow In DataGridView2.Rows
			If Not row.IsNewRow Then

				If row.Cells("da_tras").Value > 0 Then
					countMaggioreDiZero += 1
				End If
			End If
		Next

		If Magazzino.OttieniDettagliAnagrafica(Codice_SAP).Gestito_a_ferretto <> "Y" Then
			MsgBox("Articolo " & Codice_SAP & " non gestito a Ferretto")
			Return

		End If

		' Se c'è un solo record con valore > 0, chiedi conferma all'utente
		If countMaggioreDiZero = 1 Then
			Dim result As DialogResult = MessageBox.Show("Esiste la possibilità di trasferire a WIP l'articolo, è fortemente sconsigliato proseguire con questa operazione. Proseguire?",
													 "Attenzione", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

			If result = DialogResult.No Then
				Exit Sub
			End If
		End If


		' Esegui il trasferimento
		trasferimento_altro_magazzino_DEF(Codice_SAP, codice_magazzino, tipo_entrata, Txt_DocNum.Text, CheckBox1.Checked, par_magazzino_destinazione)
		aggiorna_datagridivew_codici(trasferimento_eseguito)
	End Sub

	Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
		'If CheckBox1.Checked = True Then
		'	preview_scontrino = True
		'Else
		'	preview_scontrino = False
		'End If
	End Sub



	Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
		codice_magazzino = ComboBox1.Text

		If ComboBox1.SelectedIndex = 0 Then
			Button2.Text = "SCA"
			Button3.Text = "CQ"
			Button8.Text = "03"

			Button1.Text = "Conto lavoro 16"
		ElseIf ComboBox1.SelectedIndex = 1 Then

			Button2.Text = "BSCA"
			Button1.Text = "Conto lavoro B16"
			Button8.Text = "B03"
			Button9.Visible = False

		End If

	End Sub

	Private Sub DataGridView_trasferito_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

	End Sub

	Private Sub DataGrid_EM_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGrid_EM.CellContentClick

	End Sub

	Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
		If inizializzazione_form = False Then
			Aggiorna()
			TableLayoutPanel1.BackColor = Color.Orange
		End If

		If RadioButton3.Checked = True Then
			tipo_entrata = "EMP"
		End If
	End Sub





	Private Sub BackgroundWorker_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker.RunWorkerCompleted

	End Sub

	Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
		Magazzino.codice_disegno = Button5.Text
		Magazzino.visualizza_disegno(Button5.Text)
	End Sub





	Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick

		Dim linenum As Integer = 0
		Dim commessa_destinazione As String = ""
		Dim par_datagridview As DataGridView = DataGridView2
		Dim par_documento As String

		If e.RowIndex >= 0 Then
			Riga = e.RowIndex
			linenum = par_datagridview.Rows(e.RowIndex).Cells("linenum").Value
			par_documento = par_datagridview.Rows(e.RowIndex).Cells("DOC").Value

			If e.ColumnIndex = par_datagridview.Columns.IndexOf(Trasferisci) Then
				' Esegui il trasferimento WIP
				trasferimento_wip(
					par_documento, linenum, Codice_SAP,
					par_datagridview.Rows(e.RowIndex).Cells("REP").Value,
					Homepage.Centro_di_costo, codice_magazzino,
					par_datagridview.Rows(e.RowIndex).Cells("ODP").Value,
					par_datagridview.Rows(e.RowIndex).Cells("OC").Value,
					par_datagridview.Rows(e.RowIndex).Cells("Comm").Value,
					par_datagridview.Rows(e.RowIndex).Cells("Cliente").Value,
					par_datagridview.Rows(e.RowIndex).Cells("Trasferisci").Value,
					par_datagridview.Rows(e.RowIndex).Cells("Stato").Value,
					Txt_DocNum.Text, tipo_entrata,
					par_datagridview, Riga, CheckBox1.Checked
				)

				' Aggiorna in base alla selezione dell'utente
				Select Case True
					Case RadioButton1.Checked
						Aggiorna_EM(Txt_DocNum.Text)
					Case RadioButton2.Checked
						Aggiorna_trasferimento()
					Case RadioButton3.Checked
						Aggiorna_EMP()
					Case RadioButton4.Checked
						Aggiorna_CS(Txt_DocNum.Text)
				End Select

				' Aggiornamento dati magazzino
				N_codici_in_mag(codice_magazzino)
				trasferito(Codice_SAP, par_datagridview)
				giacenze_magazzino(Codice_SAP)

				' Gestione quantità magazzino con gestione errori
				Try
					Txt_q_em.Text = DataGrid_EM.Rows(riga_datagrid).Cells("Qta").Value
				Catch ex As Exception
					Txt_q_em.Text = 0
				End Try

				Txt_Magazzino_01.Text = Math.Round(giacenze_IN_magazzino(Codice_SAP, codice_magazzino), 3)

			ElseIf e.ColumnIndex = par_datagridview.Columns.IndexOf(Stampa) AndAlso
				   par_datagridview.Rows(e.RowIndex).Cells("doc").Value = "ODP" Then
				' Esegui stampa se il documento è ODP
				stampa_da_datagridview(
					par_datagridview.Rows(e.RowIndex).Cells("ODP").Value,
					Codice_SAP,
					par_datagridview.Rows(e.RowIndex).Cells("Col_Trasferito").Value,
					CheckBox1.Checked
				)
			End If
		End If





	End Sub

	Sub stampa_da_datagridview(par_numero_odp As String, par_codice_sap As String, par_quantità As String, par_preview_scontrino As Boolean)

		Stampa_ODP = par_numero_odp
		Stampa_Matricola = ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).commessa
		If ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).u_produzione = "INT_SALD" Then
			Stampa_Descrizione = "SALDATO"
			Stampa_Descrizione_2 = ""
		ElseIf ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).FASE = "FORMATI_ATTREZZATURE" Then
			Stampa_Descrizione = "FORMATO"
			Stampa_Descrizione_2 = ""
		Else
			Stampa_Descrizione = Microsoft.VisualBasic.Left(ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).descrizione, 28)
			Stampa_Descrizione_2 = Microsoft.VisualBasic.Mid(ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).descrizione, 29, 66)
		End If

		Stampa_Codice = par_codice_sap
		Stampa_Descrizione_Articolo = Microsoft.VisualBasic.Left(Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, 28)
		Stampa_Descrizione_Articolo_2 = Microsoft.VisualBasic.Mid(Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, 29, 56)
		Stampa_Qta = par_quantità
		Stampa_Ubicazione_Macchina = Get_Ubicazione_Matricola(ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).commessa)
		Stampa_Codice_Ordinazione = Get_Codice_Ordinazione(par_codice_sap)
		Stampa_progressivo_commessa = ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).numerone
		Stampa_Tipo = "Ordine di Produzione"
		Fun_Stampa(Stampa_Tipo, par_preview_scontrino, Stampante_Selezionata, Scontrino, magazzino_destinazione, "", par_codice_sap)

	End Sub

	Sub trasferimento_wip(par_documento As String, par_linenum As Integer, par_codice_sap As String, par_reparto As String, par_centro_di_costo As String, par_magazzino_partenza As String, par_numero_odp As String, par_numero_oc As String, par_commessa As String, par_cliente As String, par_valore_trasferimento_datagridview As String, par_stato_odp As String, par_numero_em As String, par_tipo_entrata As String, par_datagridview2 As DataGridView, par_riga_datagridview As Integer, par_preview_Scontrino As Boolean)

		Dim par_magazzino_destinazione As String

		If par_documento = "ODP" Then


			If par_reparto = "INT" Then


				par_magazzino_destinazione = "06"



			ElseIf par_reparto = "ASSEMBL" Or par_reparto = "EST" Or par_reparto = "INT_SALD" Or par_reparto = "B_INT" Then


				par_magazzino_destinazione = "WIP"


			Else MsgBox("NON è POSSIBILE EFFETTUARE TRASFERIMENTI PER ODP CON PRODUZIONE " & par_reparto)

				Return
			End If
			Num_ODP = par_numero_odp
			Num_OC = 0



			If Verifica_odp_anticipo_materiale(par_numero_odp, par_magazzino_destinazione) = 1 Then
				MsgBox("Non è possibile fare WIPPATA per ordini ANTICIPO MATERIALE, riporre il materiale a magazzino")
				Return
			End If

			If par_numero_odp = numero_odp_precedente And Num_ODP <> 0 And par_codice_sap = codice_precedente Then
				MsgBox("Non è possibile fare 2 wippate consecutive per lo stesso ordine di produzione")
				Return
			End If
		Else


			par_magazzino_destinazione = "WIP"

			Num_OC = par_numero_oc
			Num_ODP = 0
			par_commessa = par_cliente

			If Num_OC = numero_oc_precedente And Num_OC <> 0 And par_codice_sap = codice_precedente Then
				MsgBox("Non è possibile fare 2 wippate consecutive per lo stesso ordine cliente")
				Return
			End If
		End If



		If par_codice_sap = codice_precedente And giacenze_IN_magazzino(par_codice_sap, par_magazzino_partenza) <> giacenze_in_magazzino_precedente - quantità_predente Then
			MsgBox("Non è stato scaricato correttamente il magazzino nel trasferimento precedente di questo codice. Controllare.")
			Return
		End If

		If par_valore_trasferimento_datagridview = "Trasferisci" Then
			Dim qta_a_mag As Decimal = giacenze_IN_magazzino(par_codice_sap, par_magazzino_partenza)
			Dim qta_da_trasferire_nel_doc As Decimal = quantita_da_trasferire_nel_documento(par_documento, par_codice_sap, par_numero_odp, par_numero_oc, par_linenum)
			Dim qta_entrata_con_em As Decimal = quantita_entrata_con_em(par_numero_em, par_codice_sap, par_magazzino_partenza, par_tipo_entrata)
			Dim qta_trasferimento As Decimal = Form_lotto_di_prelievo.minore(qta_a_mag, qta_da_trasferire_nel_doc)
			qta_trasferimento = Form_lotto_di_prelievo.minore(qta_trasferimento, qta_entrata_con_em)
			If par_documento = "ODP" Then



				If controllo_RIchiesta_trasferimento_ferretto(par_documento, par_numero_odp, par_numero_oc, par_linenum, par_codice_sap) = 0 Then


					If qta_trasferimento > 0 Then

						If par_commessa.StartsWith("M") And Form_layout_CAP_1.check_baia_layout(par_commessa).numero_baia = 0 Then
							Dim result_test As Integer

							result_test = MessageBox.Show("La commessa " & par_commessa & " non risulta piazzata nella dashbard layout. è consigliato PRIMA piazzare la commessa in PRE e poi proseguire con la wippata. Vuoi wippare lo stesso? ", "Confirmation", MessageBoxButtons.YesNo)
							If result_test = DialogResult.Yes Then

							Else
								Return
							End If
						End If
						If par_stato_odp = "P" Then
							Dim result As Integer
							result = MessageBox.Show("L'ordine di produzione risulta PIANIFICATO, vuoi rilasciarlo?", "Confirmation", MessageBoxButtons.YesNo)
							If result = DialogResult.Yes Then
								cambia_stato(par_numero_odp)
								trasferito(par_codice_sap, par_datagridview2)
								result = MessageBox.Show("Stampare l'ordine di produzione appena RILASCIATO?", "Confirmation", MessageBoxButtons.YesNo)
								If result = DialogResult.Yes Then

									ODP_Form.testata_odp(par_numero_odp)
									ODP_Form.Fun_Stampa()

								End If
							Else
								MsgBox("Non è possibile trasferire per ODP pianificato")
								Return

							End If
						End If


						If check_giacenza_per_trasferimento(par_codice_sap, par_magazzino_partenza, qta_trasferimento) = 1 Then


							Beep()
							MsgBox("ERRORE Si sta cercando di trasferire una quantità maggiore della giacenza OITW")


							Return

						End If

						If check_giacenza_per_trasferimento_trasferimenti_oivl(Codice_SAP, codice_magazzino, qta_trasferimento) = 1 Then
							Beep()
							MsgBox("ERRORE Si sta cercando di trasferire una quantità maggiore della giacenza Ovl")
							Return
						End If


						par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="stampa").Value = "Stampa"


						par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="Col_Trasferito").Value = Math.Round(CDbl(qta_trasferimento), 3)

						par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="tras").Value = Math.Round(CDbl(qta_trasferimento), 3) + par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="Tras").Value

						par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="DA_TRAS").Value = par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="Q").Value - qta_trasferimento


						giacenze_in_magazzino_precedente = giacenze_IN_magazzino(par_codice_sap, par_magazzino_partenza)
						Trasferimento_in_WIP(par_documento, par_codice_sap, par_numero_odp, par_numero_oc, Replace(qta_trasferimento, ",", "."), par_magazzino_partenza, par_magazzino_destinazione, par_linenum, par_commessa, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, "Manuale", 0, 0, "Trasferimento")

						quantità_predente = qta_trasferimento

						numero_odp_precedente = par_numero_odp
						numero_oc_precedente = par_numero_oc
						codice_precedente = par_codice_sap
						If trasferimento_eseguito = "SI" Then
							Stampa_ODP = par_numero_odp
							Stampa_Matricola = par_commessa


							If ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).u_produzione = "INT_SALD" Then
								Stampa_Descrizione = "SALDATO"
								Stampa_Descrizione_2 = ""
							ElseIf ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).FASE = "FORMATI_ATTREZZATURE" Then
								Stampa_Descrizione = "FORMATO"
								Stampa_Descrizione_2 = ""
							Else
								Stampa_Descrizione = Microsoft.VisualBasic.Left(ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).descrizione, 28)
								If Magazzino.OttieniDettagliAnagrafica(par_codice_sap).codice_brb = "" Then
									Stampa_Descrizione_2 = Microsoft.VisualBasic.Mid(ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).descrizione, 29, 66)
								Else
									Stampa_Descrizione_2 = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).codice_brb

								End If
							End If


							Stampa_Codice = par_codice_sap
							Stampa_Descrizione_Articolo = Microsoft.VisualBasic.Left(Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, 28)
							Stampa_Descrizione_Articolo_2 = Microsoft.VisualBasic.Mid(Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, 29, 56)
							Stampa_Qta = qta_trasferimento
							Stampa_Ubicazione_Macchina = Get_Ubicazione_Matricola(par_commessa)
							Stampa_Codice_Ordinazione = Get_Codice_Ordinazione(par_codice_sap)
							Stampa_progressivo_commessa = ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_odp).numerone
							Stampa_Tipo = "Ordine di Produzione"
							Fun_Stampa(Stampa_Tipo, CheckBox1.Checked, Stampante_Selezionata, Scontrino, par_magazzino_destinazione, "", par_codice_sap)
						End If


						Acquisti.aggiusta_CONFERMATO(par_codice_sap)


						Magazzino.AWOR(par_numero_odp, par_linenum, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)



						trasferito(par_codice_sap, par_datagridview2)




						MsgBox("Trasferimento di " & qta_trasferimento & " in " & par_magazzino_destinazione & " Effettuato con successo")



					Else
						MsgBox("Non risulta ci sia materiale trasferibile")
					End If

				Else
					MsgBox("il codice risulta In richiesta trasferimento da Ferretto")

				End If





			ElseIf par_documento = "OC" Then


				If qta_trasferimento > 0 Then

					par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="Trasferisci").Value = ""
					par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="stampa").Value = "Stampa"


					par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="Col_Trasferito").Value = Math.Round(CDbl(qta_trasferimento), 3)
					par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="tras").Value = Math.Round(CDbl(qta_trasferimento), 3) + par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="Tras").Value

					par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="DA_TRAS").Value = par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="Q").Value - qta_trasferimento



					Trasferimento_in_WIP(par_documento, par_codice_sap, par_numero_odp, par_numero_oc, Replace(qta_trasferimento, ",", "."), par_magazzino_partenza, par_magazzino_destinazione, par_linenum, par_commessa, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, "Manuale", 0, 0, "Trasferimento")

					numero_odp_precedente = par_numero_odp
					numero_oc_precedente = par_numero_oc
					codice_precedente = par_codice_sap

					Stampa_ODP = par_numero_oc
					Stampa_Matricola = Microsoft.VisualBasic.Left(par_commessa, 10)
					Stampa_Descrizione = Microsoft.VisualBasic.Left(par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="ODP_name").Value, 28)
					Stampa_Descrizione_2 = Microsoft.VisualBasic.Mid(par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="ODP_name").Value, 29, 66)
					Stampa_Codice = par_codice_sap
					Stampa_Descrizione_Articolo = Microsoft.VisualBasic.Left(Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, 28)
					Stampa_Descrizione_Articolo_2 = Microsoft.VisualBasic.Mid(Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, 29, 56)
					Stampa_Qta = par_datagridview2.Rows(par_riga_datagridview).Cells(columnName:="Col_Trasferito").Value
					Stampa_Codice_Ordinazione = Get_Codice_Ordinazione(par_codice_sap)
					Stampa_Tipo = "OC"
					Fun_Stampa(Stampa_Tipo, par_preview_Scontrino, Stampante_Selezionata, Scontrino, par_magazzino_destinazione, "", par_codice_sap)

					Acquisti.aggiusta_CONFERMATO(Codice_SAP)



					MsgBox("Trasferimento di " & qta_trasferimento & " in " & magazzino_destinazione & " Effettuato con successo")


				Else
					MsgBox("Non risulta quantità trasferibile")
				End If
			End If

		End If
	End Sub




	Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
		Dim divValue As String = DataGridView2.Rows(e.RowIndex).Cells("DIV").Value.ToString()
		Select Case divValue
			Case "BRB01"
				DataGridView2.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.Yellow
			Case "TIR01"
				DataGridView2.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.LightBlue
			Case "KTF01"
				DataGridView2.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.Green
		End Select

		'If DataGridView2.Rows(e.RowIndex).Cells(columnName:="Baia").Value <=> "Prelievo" Then
		'	DataGridView2.Rows(e.RowIndex).DefaultCellStyle.Font = New Font(DataGridView2.DefaultCellStyle.Font, FontStyle.Bold)
		'End If
	End Sub



	Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
		If inizializzazione_form = False Then
			Aggiorna()
			TableLayoutPanel1.BackColor = Color.Orange
		End If

		If RadioButton4.Checked = True Then
			tipo_entrata = "CS"
		End If
	End Sub



	Private Sub Cmd_Ripeti_Ultima_Stampa_Click(sender As Object, e As EventArgs) Handles Cmd_Ripeti_Ultima_Stampa.Click
		Fun_Stampa(Stampa_Tipo, CheckBox1.Checked, Stampante_Selezionata, Scontrino, "", "", Stampa_Codice)
	End Sub

	Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

	End Sub

	Private Sub DataGridView2_CellContextMenuStripChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContextMenuStripChanged

	End Sub

	Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
		If Codice_BP_finale = "" Then
			MsgBox("Selezionare un fornitore")
			Return
		End If

		If ComboBox2.SelectedIndex < 0 Then
			MsgBox("Selezionare tipo di pacco")
			Return
		End If

		inserisci_nuovo_arrivo(Codice_BP_finale, TextBox1.Text, Homepage.ID_SALVATO, Elenco_priorità_pacchi(ComboBox2.SelectedIndex))
		riempi_datagridview_coda_em(DataGridView1, TextBox3.Text, TextBox4.Text)
		TextBox1.Text = ""
		MsgBox("Coda aggiornata")
	End Sub

	Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
		Business_partner.Show()
		Business_partner.Provenienza = "Form_entrata_merce"
	End Sub

	Sub inserisci_nuovo_arrivo(Cardcode As String, par_ddt As String, Utente As Integer, par_tipo_pacco As String)


		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()
		Dim CMD_SAP As New SqlCommand
		CMD_SAP.Connection = Cnn

		CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[dbo].[Entrata_merce_coda]
           ([N°]
           ,[Cardcode]
,[DDT]
,tipo_pacco
           ,[Data_i]
           ,[Utente]
           ,[Stato])
     VALUES
           (" & trova_n_ingresso() & "
           ,'" & Cardcode & "'
,'" & par_ddt & "'
,'" & par_tipo_pacco & "'
           ,getdate()
           ,'" & Utente & "'
           ,'O')
          "
		CMD_SAP.ExecuteNonQuery()
		Cnn.Close()

	End Sub

	Public Function trova_n_ingresso()
		Dim n_max As Integer = 0
		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli

		Cnn.Open()

		Dim CMD_SAP As New SqlCommand
		Dim cmd_SAP_reader As SqlDataReader


		CMD_SAP.Connection = Cnn

		CMD_SAP.CommandText = "Select coalesce(max(coalesce([N°],0)),0) as 'MAx'
from [Tirelli_40].[dbo].[Entrata_merce_coda] 
t0 where t0.stato='O'"


		cmd_SAP_reader = CMD_SAP.ExecuteReader

		If cmd_SAP_reader.Read() Then
			n_max = cmd_SAP_reader("max") + 1
		Else
			n_max = cmd_SAP_reader("max") = 1
		End If
		cmd_SAP_reader.Close()
		Cnn.Close()
		Return n_max
	End Function

	Sub riempi_datagridview_coda_em(par_datagridview As DataGridView, par_forn As String, par_ddt As String)
		Dim filtro_forn As String = ""
		If par_forn = "" Then
			filtro_forn = ""
		Else
			filtro_forn = " and t10.cardname   Like '%%" & par_forn & "%%' "

		End If

		Dim filtro_ddt As String = ""
		If par_ddt = "" Then
			filtro_ddt = ""
		Else
			filtro_ddt = " and t10.ddt   Like '%%" & par_ddt & "%%'  "

		End If
		Dim contatore As Integer = 0
		par_datagridview.Rows.Clear()
		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()

		Dim CMD_SAP As New SqlCommand
		Dim cmd_SAP_reader As SqlDataReader

		CMD_SAP.Connection = Cnn

		CMD_SAP.CommandText = "select *
from
(
SELECT  t0.[ID]
      ,t0.[N°]
      ,t0.[Cardcode]
	  , coalesce(t1.cardname,'') as 'Cardname'
      ,t0.[DDT]
      ,t0.[Data_i]
      ,t0.[Utente]
      ,t0.[Data_fi]
      ,t0.[Utente_fine]
      ,t0.[Stato]
	  , coalesce(count(t2.id_em),0) as 'Prio'
,coalesce(t3.priorità,0) as 'Prio_pacco'
,t0.tipo_pacco

  FROM [Tirelli_40].[dbo].[Entrata_merce_coda] t0 
  left join ocrd t1 on t0.cardcode=t1.cardcode
  left join [Tirelli_40].[dbo].[Entrate_merce_prio] t2 on t2.id_em=t0.id
left join [Tirelli_40].[dbo].[Entrata_merce_tipologia_pacchi] t3 on t3.abbreviazione=t0.tipo_pacco
where t0.stato='O' or t0.stato='E'
  group by t0.[ID]
      ,t0.[N°]
      ,t0.[Cardcode]
	  , coalesce(t1.cardname,'') 
      ,t0.[DDT]
      ,t0.[Data_i]
      ,t0.[Utente]
      ,t0.[Data_fi]
      ,t0.[Utente_fine]
      ,t0.[Stato]
,t0.tipo_pacco
,t3.priorità
	  )
	  as t10
where 0= 0 " & filtro_forn & filtro_ddt & "
	  order by t10.prio DESC,t10.prio_pacco, t10.[N°]
"




		cmd_SAP_reader = CMD_SAP.ExecuteReader

		Do While cmd_SAP_reader.Read()


			contatore += 1
			par_datagridview.Rows.Add(cmd_SAP_reader("ID"), cmd_SAP_reader("N°"), cmd_SAP_reader("cardname"), cmd_SAP_reader("DDT"), cmd_SAP_reader("Data_i"), cmd_SAP_reader("tipo_pacco"), cmd_SAP_reader("prio"), cmd_SAP_reader("Stato"))


		Loop
		cmd_SAP_reader.Close()
		Cnn.Close()
		par_datagridview.ClearSelection()
		Label8.Text = contatore

	End Sub

	Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
		If id_em = 0 Then
			MsgBox("Selezionare un'EM")
			Return
		End If
		dai_priorità_em(id_em, Homepage.ID_SALVATO)
		riempi_datagridview_coda_em(DataGridView1, TextBox3.Text, TextBox4.Text)
		MsgBox("Priorità data con successo")
	End Sub

	Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
		If id_em = 0 Then
			MsgBox("Selezionare un'EM")
			Return
		End If
		completa_em(id_em, Homepage.ID_SALVATO)
		riempi_datagridview_coda_em(DataGridView1, TextBox3.Text, TextBox4.Text)
		TextBox1.Text = ""
		MsgBox("EM Segnata come completata")

	End Sub



	Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
		Dim par_datagridview As DataGridView = DataGridView1
		If e.RowIndex >= 0 Then
			Label7.Text = par_datagridview.Rows(e.RowIndex).Cells(columnName:="Forn").Value
			TextBox1.Text = par_datagridview.Rows(e.RowIndex).Cells(columnName:="DDT").Value
			Label2.Text = par_datagridview.Rows(e.RowIndex).Cells(columnName:="N").Value
			id_em = par_datagridview.Rows(e.RowIndex).Cells(columnName:="ID").Value
		End If
	End Sub

	Sub inserimento_tipologia_pacchi(par_combobox As ComboBox)


		par_combobox.Items.Clear()
		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()

		Dim CMD_SAP As New SqlCommand
		Dim cmd_SAP_reader As SqlDataReader

		CMD_SAP.Connection = Cnn
		CMD_SAP.CommandText = "SELECT T0.[Abbreviazione] , T0.[nome] , T0.[Priorità] 
        FROM [tirelli_40].dbo.entrata_merce_tipologia_pacchi t0
order by T0.[nome]
"

		cmd_SAP_reader = CMD_SAP.ExecuteReader

		Dim Indice As Integer
		Indice = 0
		Do While cmd_SAP_reader.Read()
			Elenco_priorità_pacchi(Indice) = cmd_SAP_reader("Abbreviazione")
			par_combobox.Items.Add(cmd_SAP_reader("Abbreviazione"))
			Indice = Indice + 1
		Loop
		cmd_SAP_reader.Close()
		Cnn.Close()



	End Sub 'Inserisco le risorse nella combo box


	Sub dai_priorità_em(id_em As Integer, Utente As Integer)


		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()
		Dim CMD_SAP As New SqlCommand
		CMD_SAP.Connection = Cnn

		CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[dbo].[Entrate_merce_prio]
           ([id_em]
           ,[Utente]
           ,[Data])
     VALUES
           (" & id_em & "
           ,'" & Utente & "'
           ,GETDATE())"
		CMD_SAP.ExecuteNonQuery()
		Cnn.Close()

	End Sub

	Sub completa_em(id_em As Integer, Utente As Integer)


		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()
		Dim CMD_SAP As New SqlCommand
		CMD_SAP.Connection = Cnn

		CMD_SAP.CommandText = "UPDATE [TIRELLI_40].[dbo].[Entrata_merce_coda]

      SET
      [Data_fi] = GETDATE()
      ,[Utente_fine] = '" & Utente & "'
      ,[Stato] = 'C'
 WHERE [ID]='" & id_em & "'"


		CMD_SAP.ExecuteNonQuery()
		Cnn.Close()

	End Sub

	Sub em_em(id_em As Integer, Utente As Integer)


		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()
		Dim CMD_SAP As New SqlCommand
		CMD_SAP.Connection = Cnn

		CMD_SAP.CommandText = "UPDATE [TIRELLI_40].[dbo].[Entrata_merce_coda]

      SET
      [Data_fi] = GETDATE()
      ,[Utente_fine] = '" & Utente & "'
      ,[Stato] = 'E'
 WHERE [ID]='" & id_em & "'"


		CMD_SAP.ExecuteNonQuery()
		Cnn.Close()

	End Sub

	Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

	End Sub

	Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
		riempi_datagridview_coda_em(DataGridView1, TextBox3.Text, TextBox4.Text)
	End Sub

	Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
		riempi_datagridview_coda_em(DataGridView1, TextBox3.Text, TextBox4.Text)
	End Sub

	Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
		Dim par_datagridview As DataGridView = DataGridView1
		If par_datagridview.Rows(e.RowIndex).Cells(columnName:="Stato_EM").Value = "E" Then
			par_datagridview.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.GreenYellow
		End If

		If par_datagridview.Rows(e.RowIndex).Cells(columnName:="Prio").Value = 1 Then

			par_datagridview.Rows(e.RowIndex).Cells(columnName:="Prio").Style.BackColor = Color.Yellow


		ElseIf par_datagridview.Rows(e.RowIndex).Cells(columnName:="Prio").Value > 1 Then
			par_datagridview.Rows(e.RowIndex).Cells(columnName:="Prio").Style.BackColor = Color.Red
		End If

		If par_datagridview.Rows(e.RowIndex).Cells(columnName:="Tipo").Value = "B" Then

			par_datagridview.Rows(e.RowIndex).Cells(columnName:="Tipo").Style.ForeColor = Color.White
			par_datagridview.Rows(e.RowIndex).Cells(columnName:="Tipo").Style.BackColor = Color.Blue


		ElseIf par_datagridview.Rows(e.RowIndex).Cells(columnName:="Tipo").Value = "P" Then
			par_datagridview.Rows(e.RowIndex).Cells(columnName:="Tipo").Style.BackColor = Color.LightBlue


		ElseIf par_datagridview.Rows(e.RowIndex).Cells(columnName:="Tipo").Value = "C" Then
			par_datagridview.Rows(e.RowIndex).Cells(columnName:="Tipo").Style.BackColor = Color.Pink
		End If


	End Sub

	Private Sub RadioButton4_Click(sender As Object, e As EventArgs) Handles RadioButton4.Click

	End Sub

	Private Sub CancellaRigaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CancellaRigaToolStripMenuItem.Click
		'If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Trasferito").Value > 0 Then
		'    MsgBox("Impossibile cancellare riga di un codice che risulta TRASFERITO")
		'Else
		Dim PAR_DATAGRIDVIEW As DataGridView
		PAR_DATAGRIDVIEW = DataGridView2

		' Supponendo che COLONNAID sia il nome della colonna che vuoi usare per cancellare il commento
		Dim COLONNAID As String = "ID"
		Dim COLONNAddt As String = "DDT"
		Dim COLONNA_tipo_pacco As String = "Tipo"

		Dim selectedRow As DataGridViewRow = PAR_DATAGRIDVIEW.CurrentRow

		' Verifica che ci sia una riga selezionata prima di procedere
		If selectedRow IsNot Nothing Then
			' Chiede conferma all'utente se vuole cancellare il commento
			Dim result As DialogResult = MessageBox.Show("Sei sicuro di voler aggiornare questa riga dalla coda?" & vbCrLf & selectedRow.Cells("Forn").Value, "Conferma Aggiornamento", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

			' Se l'utente conferma, procedi con la cancellazione
			If result = DialogResult.Yes Then
				' Passa l'ID della riga alla funzione cancella_commento
				aggiorna_riga_coda(selectedRow.Cells(COLONNAID).Value, selectedRow.Cells(COLONNAddt).Value, selectedRow.Cells(COLONNA_tipo_pacco).Value)

				' Rimuovi la riga selezionata
				PAR_DATAGRIDVIEW.Rows.RemoveAt(selectedRow.Index)
			End If
		Else
			MessageBox.Show("Seleziona una riga prima di cancellarla.")
		End If
	End Sub

	Private Sub DatiAnagraficiArticoloToolStripMenuItem_Click(sender As Object, e As EventArgs)

	End Sub

	Private Sub EliminaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EliminaToolStripMenuItem.Click
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
			Dim result As DialogResult = MessageBox.Show("Sei sicuro di voler cancellare questa riga dalla coda?" & vbCrLf & selectedRow.Cells("Forn").Value, "Conferma Cancellazione", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

			' Se l'utente conferma, procedi con la cancellazione
			If result = DialogResult.Yes Then
				' Passa l'ID della riga alla funzione cancella_commento
				cancella_riga_coda(selectedRow.Cells(COLONNAID).Value)

				' Rimuovi la riga selezionata
				PAR_DATAGRIDVIEW.Rows.RemoveAt(selectedRow.Index)
			End If
		Else
			MessageBox.Show("Seleziona una riga prima di cancellarla.")
		End If
	End Sub

	Sub cancella_riga_coda(par_ID As Integer)
		Dim CNN6 As New SqlConnection
		CNN6.ConnectionString = Homepage.sap_tirelli
		CNN6.Open()

		Dim CMD_SAP_5 As New SqlCommand
		CMD_SAP_5.Connection = CNN6
		CMD_SAP_5.CommandText = "DELETE FROM [tirelli_40].[dbo].[Entrata_merce_coda]
    
WHERE ID=@ID"



		CMD_SAP_5.Parameters.AddWithValue("@ID", par_ID)



		CMD_SAP_5.ExecuteNonQuery()

		CNN6.Close()
	End Sub


	Sub aggiorna_riga_coda(par_ID As Integer, par_ddt As String, par_tipo_pacco As String)
		Dim CNN6 As New SqlConnection
		CNN6.ConnectionString = Homepage.sap_tirelli
		CNN6.Open()

		Dim CMD_SAP_5 As New SqlCommand
		CMD_SAP_5.Connection = CNN6
		CMD_SAP_5.CommandText = "Update [tirelli_40].[dbo].[Entrata_merce_coda]

set [DDT] =@ddt, tipo_pacco=@tipo_pacco
    
WHERE ID=@ID"



		CMD_SAP_5.Parameters.AddWithValue("@ID", par_ID)
		CMD_SAP_5.Parameters.AddWithValue("@ddt", par_ddt)
		CMD_SAP_5.Parameters.AddWithValue("@tipo_pacco", par_tipo_pacco)



		CMD_SAP_5.ExecuteNonQuery()

		CNN6.Close()
	End Sub

	Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
		riempi_datagridview_coda_em(DataGridView1, TextBox3.Text, TextBox4.Text)
	End Sub

	Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
		If id_em = 0 Then
			MsgBox("Selezionare un'EM")
			Return
		End If
		em_em(id_em, Homepage.ID_SALVATO)
		riempi_datagridview_coda_em(DataGridView1, TextBox3.Text, TextBox4.Text)
		TextBox1.Text = ""
		MsgBox("EM segnalata")
	End Sub

	Public Function cassetto_del_codice(par_codice As String)
		Dim dettagli As New Dettagli_cassetto_codice()

		dettagli.numero_cassetto = 0
		dettagli.numero_magazzino = 0


		Dim CNN As New SqlConnection
		CNN.ConnectionString = Homepage.sap_tirelli
		CNN.Open()


		Dim CMD_SAP As New SqlCommand
		Dim cmd_SAP_reader As SqlDataReader

		CMD_SAP.Connection = CNN
		CMD_SAP.CommandText = "SELECT  [Codice]
      ,[Magazzino]
      ,[Cassetto]
  FROM [Tirelli_40].[dbo].[Cassetto_codici]
where codice='" & par_codice & "'
"

		cmd_SAP_reader = CMD_SAP.ExecuteReader

		If cmd_SAP_reader.Read() Then


			dettagli.numero_cassetto = cmd_SAP_reader("Cassetto")
			dettagli.numero_magazzino = cmd_SAP_reader("Magazzino")


		End If
		cmd_SAP_reader.Close()
		CNN.Close()


		Return dettagli
	End Function

	Public Class Dettagli_cassetto_codice
		Public numero_cassetto As Integer
		Public numero_magazzino As Integer


	End Class

	Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

		aggiorna_tabella()

	End Sub

	Sub aggiorna_tabella()

		If RadioButton1.Checked = True Then
			Aggiorna_EM(Txt_DocNum.Text)
		ElseIf RadioButton2.Checked = True Then
			Aggiorna_trasferimento()
		ElseIf RadioButton3.Checked = True Then
			Aggiorna_EMP()
		ElseIf RadioButton4.Checked = True Then
			Aggiorna_CS(Txt_DocNum.Text)
		End If
	End Sub

	Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
		aggiorna_tabella()
	End Sub

	Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
		Timer1.Start()
		Button16.Visible = False
		Button17.Visible = True
		TextBox6.Text = Trova_max_prelievo()
		TextBox6.BackColor = Color.Green

	End Sub

	Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
		Timer1.Stop()
		Button16.Visible = True
		Button17.Visible = False
		TextBox6.BackColor = Color.Red
	End Sub


	Sub stampa_wippate_ferretto(par_id_max As Integer)
		Timer1.Stop()
		TextBox6.BackColor = Color.Red
		Dim Cnn1 As New SqlConnection


		Cnn1.ConnectionString = Homepage.sap_tirelli
		'MsgBox(Stringa_Connessione_SAP)
		Cnn1.Open()


		Dim CMD_SAP_2 As New SqlCommand
		Dim cmd_SAP_reader_2 As SqlDataReader


		CMD_SAP_2.Connection = Cnn1
		CMD_SAP_2.CommandText = "SELECT TOP (1)
SUBSTRING(LISTNUMBER,5,5) AS 'Docentry_rt'
,t2.docnum as 'Docnum_RT'
,coalesce(t5.docnum,0) as 'Docnum_ODP'
,coalesce(t6.docnum,0) as 'Docnum_OC'
,t0.[id]
      ,t0.[recordStatus]
      ,t0.[recordWritingDate]
      ,t0.[recordImportationDate]
      ,t0.[plantId]
      ,t0.[response]
      ,t0.[listType]
      ,t0.[listNumber]
      ,t0.[lineNumber]
      ,t0.[item]
      ,t0.[batch]
      ,t0.[serialNumber]
      ,t0.[requestedQty]
      ,t0.[processedQty]
      ,t0.[errorCause]
      ,t0.[wmsGenerated]
      ,t0.[auxText01]
      ,t0.[auxText02]
      ,t0.[auxText03]


	  
  FROM [FGWmsErp].[dbo].[LISTS_RESULT] t0
  LEFT JOIN [TIRELLISRLDB].[dbo].[wtq1] t1 on t1.LineNum= t0.lineNumber  and SUBSTRING(t0.LISTNUMBER,5,5)=t1.docentry
  left join [TIRELLISRLDB].[dbo].owtq t2 on t2.docentry=t1.docentry
  left join [TIRELLISRLDB].[dbo].wor1 t3 on t3.docentry=t1.U_PRG_AZS_OpDocEntry
  left join [TIRELLISRLDB].[dbo].rdr1 t4 on t4.docentry=t1.U_PRG_AZS_OcDocEntry
  left join [TIRELLISRLDB].[dbo].owor t5 on t3.docentry=t5.docentry
  left join [TIRELLISRLDB].[dbo].ordr t6 on t4.docentry=t6.docentry
left join [Tirelli_40].[dbo].[Stampa_Ferretto] t7 on t7.id_ferretto=t0.id
  where response =1 AND AUXTEXT02='WIP' and coalesce(t7.id_ferretto,'')='' and t0.processedqty >0 and t0.id>" & par_id_max & "
  order by id 
"

		cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

		If cmd_SAP_reader_2.Read() Then
			Dim tipo_doc As String = "ODP"

			If cmd_SAP_reader_2("Docnum_ODP") = 0 Then
				tipo_doc = "OC"
			Else
				tipo_doc = "ODP"
			End If


			stampa_per_wip(tipo_doc, cmd_SAP_reader_2("item"), cmd_SAP_reader_2("Docnum_ODP"), cmd_SAP_reader_2("processedqty"), cmd_SAP_reader_2("auxtext02"))
			INSERISCI_LOG_stampa_FERRETTO(cmd_SAP_reader_2("id"), cmd_SAP_reader_2("item"), cmd_SAP_reader_2("listnumber"), "W")

		End If

		cmd_SAP_reader_2.Close()
		Cnn1.Close()
		Timer1.Start()
		TextBox6.BackColor = Color.Green
	End Sub

	Public Function Trova_max_prelievo()

		Dim id_prelievo As Integer = 0
		Dim Cnn1 As New SqlConnection
		Cnn1.ConnectionString = Homepage.sap_tirelli
		Cnn1.Open()

		Dim CMD_SAP_2 As New SqlCommand
		Dim cmd_SAP_reader_2 As SqlDataReader


		CMD_SAP_2.Connection = Cnn1
		CMD_SAP_2.CommandText = "SELECT TOP 1
max([id]) as 'Max'


	  
  FROM [FGWmsErp].[dbo].[LISTS_RESULT] t0
  LEFT JOIN [TIRELLISRLDB].[dbo].[wtq1] t1 on t1.LineNum= t0.lineNumber  and SUBSTRING(t0.LISTNUMBER,5,5)=t1.docentry
  left join [TIRELLISRLDB].[dbo].owtq t2 on t2.docentry=t1.docentry
  left join [TIRELLISRLDB].[dbo].wor1 t3 on t3.docentry=t1.U_PRG_AZS_OpDocEntry
  left join [TIRELLISRLDB].[dbo].rdr1 t4 on t4.docentry=t1.U_PRG_AZS_OcDocEntry
  left join [TIRELLISRLDB].[dbo].owor t5 on t3.docentry=t5.docentry
  left join [TIRELLISRLDB].[dbo].ordr t6 on t4.docentry=t6.docentry
  where response =1 AND AUXTEXT02='WIP'
    "

		cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

		If cmd_SAP_reader_2.Read() Then

			id_prelievo = cmd_SAP_reader_2("Max")
		End If

		cmd_SAP_reader_2.Close()
		Cnn1.Close()
		Return id_prelievo
	End Function

	Sub stampa_per_wip(par_tipo_doc As String, par_codice_sap As String, par_numero_doc As Integer, qta_trasferimento As String, par_magazzino_destinazione As String)
		If par_tipo_doc = "ODP" Then
			Stampa_ODP = par_numero_doc
			Stampa_Matricola = ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_doc).commessa


			If ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_doc).u_produzione = "INT_SALD" Then
				Stampa_Descrizione = "SALDATO"
				Stampa_Descrizione_2 = ""
			ElseIf ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_doc).FASE = "FORMATI_ATTREZZATURE" Then
				Stampa_Descrizione = "FORMATO"
				Stampa_Descrizione_2 = ""
			Else
				Stampa_Descrizione = Microsoft.VisualBasic.Left(ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_doc).descrizione, 28)
				If Magazzino.OttieniDettagliAnagrafica(par_codice_sap).codice_brb = "" Then
					Stampa_Descrizione_2 = Microsoft.VisualBasic.Mid(ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_doc).descrizione, 29, 66)
				Else
					Stampa_Descrizione_2 = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).codice_brb

				End If

			End If


			Stampa_Codice = par_codice_sap
			Stampa_Descrizione_Articolo = Microsoft.VisualBasic.Left(Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, 28)
			Stampa_Descrizione_Articolo_2 = Microsoft.VisualBasic.Mid(Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, 29, 56)
			Stampa_Qta = qta_trasferimento
			Stampa_Ubicazione_Macchina = Get_Ubicazione_Matricola(Stampa_Matricola)
			Stampa_Codice_Ordinazione = Get_Codice_Ordinazione(par_codice_sap)
			Stampa_progressivo_commessa = ODP_Form.ottieni_informazioni_odp("Numero", 0, par_numero_doc).numerone
			Stampa_Tipo = "Ordine di Produzione"
			If Trasferimento_magazzino.CheckBox1.Checked = False Then

				Fun_Stampa(Stampa_Tipo, CheckBox1.Checked, Stampante_Selezionata, Scontrino, par_magazzino_destinazione, "", par_codice_sap)
			End If
		Else
			Stampa_ODP = par_numero_doc
			Stampa_Matricola = Microsoft.VisualBasic.Left("_" & par_numero_doc, 10)
			Stampa_Descrizione = Microsoft.VisualBasic.Left(Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, 28)
			Stampa_Descrizione_2 = Microsoft.VisualBasic.Left(Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, 28)
			Stampa_Codice = par_codice_sap
			Stampa_Descrizione_Articolo = Microsoft.VisualBasic.Left(Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, 28)
			Stampa_Descrizione_Articolo_2 = Microsoft.VisualBasic.Mid(Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione, 29, 56)
			Stampa_Qta = qta_trasferimento
			Stampa_Codice_Ordinazione = Get_Codice_Ordinazione(par_codice_sap)
			Stampa_Tipo = "OC"
			If Trasferimento_magazzino.CheckBox1.Checked = False Then
				Fun_Stampa(Stampa_Tipo, CheckBox1.Checked, Stampante_Selezionata, Scontrino, par_magazzino_destinazione, "", par_codice_sap)
			End If
		End If

	End Sub

	Sub INSERISCI_LOG_stampa_FERRETTO(PAR_id As Integer, par_codice As String, par_rt As String, par_stato As String)


		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()
		Dim CMD_SAP As New SqlCommand
		CMD_SAP.Connection = Cnn

		CMD_SAP.CommandText = "INSERT INTO [tirelli_40].[dbo].[Stampa_Ferretto]
           ([id_ferretto]
           ,[Codice]
           ,[RT]
           ,[Stato])
     VALUES
           (" & PAR_id & "
           ,'" & par_codice & "'
           ,'" & par_rt & "'
           ,'" & par_stato & "')"
		CMD_SAP.ExecuteNonQuery()
		Cnn.Close()




	End Sub

	Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

		stampa_wippate_ferretto(TextBox6.Text - 1)


	End Sub

	Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click

	End Sub
End Class