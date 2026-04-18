Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Scheda_commessa_Pianificazione
    Private _rigaDxClick As Integer = -1
    Private WithEvents _contextMenuDGV As New ContextMenuStrip()
    Private WithEvents _menuStatoProgetto As New ToolStripMenuItem("Stato commesse - per Progetto")
    Private WithEvents _menuStatoMatricola As New ToolStripMenuItem("Stato commesse - per Matricola")
    Private WithEvents _menuStatoSottocommessa As New ToolStripMenuItem("Stato commesse - per Sottocommessa")

    Public filtro_cliente_f As String = ""
    Public campione As String
    Public dato_1_min As String
    Public dato_1_max As String
    Private filtro_numero_progetto As String
    Private filtro_cliente_progetto As String
    Public filtro_stato_progetto As String
    Public filtro_stato_rev_progetto As String

    Private filtro_nome_progetto As String
    Private filtro_PM As String
    Public filtro_n_progetto As String

    Public filtro_nome_progetto_commessa As String
    Public tipo_campione As String = 100
    Public filtro_tipo_campione = ""
    Public filtro_cliente_campione As String = ""
    Public filtro_nome_campione As String = ""
    Public filtro_codice_sap_campione As String = ""
    Public filtro_id_Campione As String
    Public filtro_desc_sup As String
    Public Codice_commessa As String
    Private filtro_brand As String
    Public filtro_acq As String
    Public inizializzazione As Boolean = True

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Sub carica_commesse(par_datagridview As DataGridView, par_itemcode As String, par_itemname As String, par_cliente As String, par_n_progetto As String, par_filtro_nome_progetto_commessa As String, par_destinazione As String, par_desc_sup As String, par_filtro_brand As String, par_filtro_baia As String)
        Dim par_filtro_n_progetto As String = ""
        Dim filtro_destinazione As String = ""
        If Homepage.ERP_provenienza = "SAP" Then
            If par_n_progetto = "" Then
                par_filtro_n_progetto = ""
            Else
                par_filtro_n_progetto = " and t20.Numero_progetto=" & par_n_progetto & ""
            End If


            If par_destinazione = "" Then
                filtro_destinazione = ""
            Else
                filtro_destinazione = "AND t20.u_country_of_delivery Like '%%" & TextBox16.Text & "%%'"
            End If

        Else

            If par_n_progetto = "" Then
                par_filtro_n_progetto = ""
            Else

                '                par_filtro_n_progetto = "  And (t0.Numero_progetto = concat(''PJ'',''" & par_n_progetto & "'') OR t0.Numero_progetto = ''" & par_n_progetto & "'')"
                par_filtro_n_progetto = " AND (t0.itemcode LIKE '%" & par_n_progetto & "%' OR t0.itemcode = '" & par_n_progetto & "')"

            End If




        End If

        Dim contatore As Integer = 0
        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_2.CommandText = "
            Select  top 100 t30.itemcode, t30.itemname, t30.desc_supp, t30.cliente, t30.cliente_finale,
t30.codice_cliente,
t30.numero_progetto,t30.absentry,t30.name,t30.pm,t30.u_country_of_delivery
,t30.brand,t30.baia,t30.zona, coalesce(t32.ordine,999) as 'Ordine'
, coalesce(t32.nome,'') as 'Nome_stato'
, t33.livello_rischio_totale
,'' AS 'Codice_cliente_finale'
from
(
select t20.itemcode, t20.itemname, t20.desc_supp, t20.cliente, t20.cliente_finale,
t20.codice_cliente,
t20.numero_progetto,t20.absentry,t20.name,t20.pm,t20.u_country_of_delivery
,t20.brand,t20.baia,t20.zona
, max(coalesce(t21.rev,0)) as 'rev_max'
from
(
Select t10.itemcode, t13.itemname,
coalesce(t13.frgnname,'') as 'Desc_supp',
  COALESCE(t14.cardname, t20.CARDNAME) As 'Cliente'
, COALESCE(t15.cardname,coalesce(t20.CARDNAME,t13.u_final_customer_name)) AS 'Cliente_finale'
,  case when t15.cardCODE is null AND T14.CARDCODE IS NULL then T13.U_FINAL_CUSTOMER_CODE WHEN T15.CARDCODE IS NULL THEN T14.CARDCODE ELSE T15.CARDCODE end AS 'Codice_cliente'
, t16.docnum as 'Numero_progetto' ,t16.[AbsEntry], t16.name
,concat(t17.lastname,' ' , t17.firstname) as 'PM'
, coalesce(t12.u_destinazione,t13.u_country_of_delivery) as 'u_country_of_delivery', coalesce(t13.u_brand,'') as 'Brand'
,coalesce(t19.nome_baia,'') as 'baia'
,coalesce(t19.[Zona],'') as 'Zona'
from
(
Select t7.itemcode, max(t0.docentry) As 'Docentry'
From oitm t7 left Join rdr1 t0 on t7.itemcode=t0.itemcode
Left Join ordr t1 on t1.docentry=t0.docentry And T1.CANCELED='N'

                        where substring(t7.itemcode, 1, 1) ='M' 
group by t7.itemcode
)
as t10 left join rdr1 t11 on t11.itemcode = t10.itemcode And t11.docentry =t10.docentry
Left Join ordr t12 on t12.docentry=t11.docentry
Left Join oitm t13 on t13.itemcode=t10.itemcode
Left Join ocrd t14 on t14.cardcode=t12.cardcode
Left Join ocrd t15 on t15.cardcode=t12.U_CodiceBP
Left Join opmg t16 on t16.[AbsEntry]=t13.u_progetto
Left Join [TIRELLI_40].[dbo].ohem t17 on t17.empid=t16.owner
Left Join [Tirelli_40].[dbo].[Layout_CAP1] t18 on t18.commessa=t10.itemcode And T18.STATO='O'
                        Left Join [Tirelli_40].[dbo].[Layout_CAP1_nomi] T19 ON T19.NUMERO_baia =t18.baia
Left Join OCRD T20 ON T20.CARDCODE=T13.U_Final_customer_Code

)
as t20
left join [Tirelli_40].[dbo].[Scheda_Tecnica_valori_progetto] t21 on t21.n_progetto=t20.numero_progetto

where T20.ITEMcode Like '%%" & par_itemcode & "%%' and T20.ITEMNAME Like '%%" & par_itemname & "%%'  AND t20.desc_supp  Like '%%" & par_desc_sup & "%%' 
And (COALESCE(t20.cliente,'') Like '%%" & par_cliente & "%%' or COALESCE(t20.cliente_finale,'') Like '%%" & par_cliente & "%%') and t20.baia Like '%%" & par_filtro_baia & "%%'  " & par_filtro_n_progetto & filtro_nome_progetto_commessa & filtro_destinazione & filtro_desc_sup & filtro_brand & "

group by t20.itemcode, t20.itemname, t20.desc_supp, t20.cliente, t20.cliente_finale,
t20.codice_cliente,
t20.numero_progetto,t20.absentry,t20.name,t20.pm,t20.u_country_of_delivery
,t20.brand,t20.baia,t20.zona
)
as t30
left join [Tirelli_40].[dbo].[Scheda_tecnica_revisioni_progetto] t31 on t31.n_progetto=t30.numero_progetto and t31.Numero=t30.rev_max
LEFT JOIN [Tirelli_40].[dbo].[Scheda_Tecnica_stato_progetto] T32 ON T32.ID=T31.STATO
left join [Tirelli_40].[dbo].[Scheda_Tecnica_valori_progetto] t33 on t33.n_progetto=t30.numero_progetto and t33.rev=t30.rev_max
order by t30.itemcode DESc"

        Else
            CMD_SAP_2.CommandText =
"SELECT top 100 
trim(t10.matricola) as 'Itemcode', t10.itemname, t10.desc_supp
, T10.DSCLI_FATT as 'Cliente'
, T10.CLI_FATT as 'Codice_cliente'
,t10.codice_cliente as 'Codice_cliente_finale'
        ,t10.codice_finale as 'Cliente_finale'
		, t10.itemcode as 'absentry',
        trim(t10.itemcode) as 'Numero_progetto',
		T10.NAME_progetto AS 'DESC_PROGETTO',
		coalesce(t14.nome,'') as 'Nome_stato',
        coalesce(t13.[livello_rischio_totale],'') as 'Livello_rischio_totale', '' as 'Name',
        t10.pm as 'CODICE_PM'
		,t10.DESC_pm as 'PM'
		, T10.DSNAZ_FINALE as u_country_of_delivery,
        t10.brand AS 'CODICE_BRAND',
		trim(T10.DESC_BRAND) AS 'BRAND',
		coalesce(t12.Nome_Baia,'') as 'Baia'
		,coalesce(t12.Zona,'') as 'Zona'
		,DATA_CONSEGNA
		,T10.NOME_STATO AS 'STATO_COMMESSA'
FROM OPENQUERY(AS400, '
    SELECT *
    FROM TIR90VIS.JGALCOM t0
    WHERE 
t0.matricola<>'''' and
       UPPER(t0.matricola) LIKE ''%%" & par_itemcode.ToUpper() & "%%'' and substring(t0.matricola,1,1)=''M'' 
      AND UPPER(t0.itemname) LIKE ''%%" & par_itemname.ToUpper() & "%%''
AND (UPPER(t0.codice_finale) LIKE ''%%" & par_cliente.ToUpper() & "%%'' 
OR UPPER(t0.dscli_fatt) LIKE ''%%" & par_cliente.ToUpper() & "%%'') 
            AND UPPER(t0.itemcode) LIKE ''%%" & par_n_progetto & "%%'' 
 AND UPPER(t0.desc_supp) LIKE ''%%" & par_desc_sup & "%%'' 
and UPPER(T0.DSNAZ_FINALE) LIKE ''%%" & par_destinazione & "%%'' 
--AND LEFT(UPPER(t0.itemcode),2) <> ''TZ''
       

ORDER BY t0.matricola DESC

limit 100  
') T10 LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1] t11
     ON t11.[Stato]<>'P'
and TRIM(t10.matricola) COLLATE SQL_Latin1_General_CP1_CI_AS = t11.commessa COLLATE SQL_Latin1_General_CP1_CI_AS
	left join [Tirelli_40].[dbo].[Layout_CAP1_nomi] t12 on t12.NUMERO_BAIA=t11.Baia
	left join [Tirelli_40].[dbo].[Scheda_Tecnica_valori_progetto_last] t13 on t13.[Codice_progetto] COLLATE SQL_Latin1_General_CP1_CI_AS=t10.numero_progetto COLLATE SQL_Latin1_General_CP1_CI_AS
	left join [Tirelli_40].[dbo].[Scheda_Tecnica_stato_progetto] t14 on t14.ordine =t13.stato"
        End If


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader_2("itemcode"),
                                      cmd_SAP_reader_2("itemname"),
                                      cmd_SAP_reader_2("desc_supp"),
                                      cmd_SAP_reader_2("cliente"),
                                      cmd_SAP_reader_2("Cliente_finale"),
                                      cmd_SAP_reader_2("Codice_cliente"),
                                      cmd_SAP_reader_2("Codice_cliente_finale"),
                                      cmd_SAP_reader_2("absentry"),
                                      cmd_SAP_reader_2("numero_progetto"),
                                      cmd_SAP_reader_2("nome_stato"),
                                      cmd_SAP_reader_2("livello_rischio_totale"),
                                      cmd_SAP_reader_2("name"),
                                      cmd_SAP_reader_2("PM"),
                                      cmd_SAP_reader_2("u_country_of_delivery"),
                                      cmd_SAP_reader_2("Brand"),
                                      cmd_SAP_reader_2("Baia"),
                                      cmd_SAP_reader_2("Zona"))
            contatore += 1
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        par_datagridview.ClearSelection()
        Label1.Text = contatore
    End Sub

    Private Async Sub tabpage15_Click(sender As Object, e As EventArgs) Handles dati_mancanti.Enter
        appunti_globali(Homepage.ID_SALVATO)
    End Sub

    Public Sub appunti_globali(par_id_dipendente As Integer)


        Dim par_datagridview As DataGridView = DataGridView13
        par_datagridview.Rows.Clear()
        Dim filtro_Dipendente As String = If(par_id_dipendente = 0, "", " AND t0.dipendente = " & par_id_dipendente)



        Using Cnn1 As New SqlConnection(Homepage.sap_tirelli)
            Cnn1.Open()

            ' --- Query progetti ---
            Using CMD_SAP_2 As New SqlCommand("
              select t30.progetto_commessa, t30.cardname,
          t30.Cliente_finale,
          t30.name, t30.data_min
		  ,t30.Rev_Max , t31.Stato, t32.Nome as 'Nome_stato', t32.Ordine
  from
  (
  SELECT t20.progetto_commessa, t21.cardname,
          COALESCE(t22.cardname,t21.cardname) AS Cliente_finale,
          t21.name, t20.data_min
		  ,max(t23.numero) as 'Rev_Max'
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
           FROM [TIRELLI_40].[DBO].dati_mancanti_progetto T0
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
           WHERE 1=1 
       ) t10
       GROUP BY t10.progetto_commessa
   ) t20
   LEFT JOIN OPMG t21 ON CAST(t20.Progetto_commessa AS VARCHAR)=CAST(T21.DocNum AS VARCHAR)
   LEFT JOIN OCRD t22 ON t22.CardCode=t21.U_Codice_cliente_finale
   left join [Tirelli_40].[dbo].[Scheda_tecnica_revisioni_progetto] t23 on t23.n_progetto=t20.Progetto_commessa
  group by 
  t20.progetto_commessa, t21.cardname,
          COALESCE(t22.cardname,t21.cardname) ,
          t21.name, t20.data_min
		  )
		  as t30
		  left join [Tirelli_40].[dbo].[Scheda_tecnica_revisioni_progetto] t31 on t31.n_progetto=t30.Progetto_commessa and t30.rev_max=t31.Numero
		  left join [Tirelli_40].[dbo].[Scheda_Tecnica_stato_progetto] t32 on t32.ID=t31.Stato
  ORDER BY t32.ordine, t30.data_min", Cnn1)

                Using rdr As SqlDataReader = CMD_SAP_2.ExecuteReader()
                    Dim currentProgetto As String = String.Empty

                    While rdr.Read()

                        par_datagridview.Rows.Add(0, rdr("progetto_commessa"), rdr("Ordine"), rdr("Nome_stato"), rdr("cliente_finale"), rdr("name"))

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

            FROM [TIRELLI_40].[DBO].[Dati_mancanti_progetto] T0
            LEFT JOIN OITM T2 ON T2.itemcode=T0.commessa
left join [TIRELLI_40].[dbo].ohem t3 on t3.empid=t0.dipendente
            WHERE (CASE WHEN T0.tipo='Progetto' THEN CAST(T0.progetto AS VARCHAR) ELSE CAST(T2.u_progetto AS VARCHAR) END) = '" & par_progetto & "'
            
            ORDER BY T0.ID", Cnn1)

                Using rdr As SqlDataReader = CMD_SAP_2.ExecuteReader()
                    Dim currentProgetto As String = String.Empty

                    While rdr.Read()

                        par_datagridview.Rows.Add(rdr("ID"), "", "", "", "", "", rdr("Commessa"), rdr("Nome_macchina"), rdr("dipendente"), rdr("data"), rdr("contenuto"), rdr("Risolto"))

                    End While
                End Using
            End Using


        End Using  ' <-- qui chiudiamo l'uso della connessione


    End Sub




    Private Sub Scheda_commessa_Pianificazione_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inizializzazione = True
        filtro_commesse()   '  Costruzione_datagridview_progetti()
        DateTimePicker4.Value = DateAdd("d", -30, Today)
        inizializzazione = False

        _contextMenuDGV.Items.AddRange(New ToolStripItem() {_menuStatoProgetto, _menuStatoMatricola, _menuStatoSottocommessa})
        DataGridView.ContextMenuStrip = _contextMenuDGV
    End Sub

    Private Sub DataGridView_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView.CellMouseDown
        If e.Button = MouseButtons.Right AndAlso e.RowIndex >= 0 Then
            _rigaDxClick = e.RowIndex
            DataGridView.ClearSelection()
            DataGridView.Rows(e.RowIndex).Selected = True
        End If
    End Sub

    Private Sub _menuStatoProgetto_Click(sender As Object, e As EventArgs) Handles _menuStatoProgetto.Click
        If _rigaDxClick < 0 Then Return
        Dim progetto As String = ""
        Dim cell = DataGridView.Rows(_rigaDxClick).Cells("codice_Progetto")
        If cell.Value IsNot Nothing Then progetto = cell.Value.ToString().Trim()
        Dim frm As New Form_stato_commesse()
        frm.TextBox1.Text = progetto
        frm.Show()
    End Sub

    Private Sub _menuStatoMatricola_Click(sender As Object, e As EventArgs) Handles _menuStatoMatricola.Click
        If _rigaDxClick < 0 Then Return
        Dim matricola As String = ""
        Dim cell = DataGridView.Rows(_rigaDxClick).Cells("Commessa")
        If cell.Value IsNot Nothing Then matricola = cell.Value.ToString().Trim()
        Dim frm As New Form_stato_commesse()
        frm.TextBox2.Text = matricola
        frm.Show()
    End Sub

    Private Sub _menuStatoSottocommessa_Click(sender As Object, e As EventArgs) Handles _menuStatoSottocommessa.Click
        If _rigaDxClick < 0 Then Return
        Dim frm As New Form_stato_commesse()
        frm.Show()
        frm.TextBox15.Focus()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        filtro_commesse()
    End Sub

    Sub filtro_commesse()
        carica_commesse(DataGridView, TextBox1.Text.ToUpper, TextBox2.Text.ToUpper, TextBox4.Text.ToUpper, TextBox14.Text, filtro_nome_progetto_commessa, TextBox16.Text.ToUpper, TextBox7.Text.ToUpper, TextBox8.Text, "")

    End Sub



    Private Async Sub DataGridView_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView.CellClick
        If e.RowIndex >= 0 Then

            Codice_commessa = DataGridView.Rows(e.RowIndex).Cells(columnName:="Commessa").Value

            If e.ColumnIndex = DataGridView.Columns.IndexOf(Commessa) Then
                If DataGridView.Rows(e.RowIndex).Cells(columnName:="Commessa").Value >= "M04000" Then
                    Scheda_tecnica.Close()
                    Scheda_tecnica.Show()
                    Scheda_tecnica.BringToFront()
                    ' Chiamata asincrona corretta
                    Await Scheda_tecnica.inizializza_scheda_tecnica(Codice_commessa)

                    Scheda_tecnica.codice_bp_campione = DataGridView.Rows(e.RowIndex).Cells(columnName:="Codice_cliente").Value
                        Scheda_tecnica.bp_code_galileo = DataGridView.Rows(e.RowIndex).Cells(columnName:="Codice_cliente").Value
                    Scheda_tecnica.final_bp_code_galileo = DataGridView.Rows(e.RowIndex).Cells(columnName:="Codice_cliente_finale_").Value


                Else


                    Scheda_commessa_documentazione.Close()
                    Scheda_commessa_documentazione.inizializzazione = 0
                    Scheda_commessa_documentazione.carico_iniziale = 0

                    Scheda_commessa_documentazione.Azzera_campi()

                    Scheda_commessa_documentazione.commessa = DataGridView.Rows(e.RowIndex).Cells(columnName:="Commessa").Value

                    Try
                        Scheda_commessa_documentazione.codice_bp_campione = DataGridView.Rows(e.RowIndex).Cells(columnName:="Codice_cliente").Value
                    Catch ex As Exception

                    End Try



                    'Scheda_commessa_documentazione.codice_bp_campione = DataGridView.Rows(e.RowIndex).Cells(0).Value


                    layout_scheda_tecnica()


                    Scheda_commessa_documentazione.compila_anagrafica(DataGridView.Rows(e.RowIndex).Cells(columnName:="Commessa").Value)


                    Scheda_commessa_documentazione.Inserimento_dipendenti()

                    Scheda_commessa_documentazione.COMPILA_RECORD_INIZIALI()

                    '  Scheda_commessa_documentazione.Rischio_effettivo()
                    Scheda_commessa_documentazione.Ultimo_aggiornamento()

                    Dim percorso_sap As String = Homepage.sap_tirelli
                    Dim percorso_immagini As String = Homepage.Percorso_immagini


                    Scheda_tecnica.riempi_datagridview_campioni(Scheda_commessa_documentazione.DataGridView3, Scheda_commessa_documentazione.codice_bp_campione, Scheda_commessa_documentazione.bp_code, Scheda_commessa_documentazione.final_bp_code, percorso_immagini, percorso_sap)

                    Scheda_tecnica.riempi_datagridview_combinazioni(Scheda_commessa_documentazione.DataGridView1, DataGridView.Rows(e.RowIndex).Cells(columnName:="Commessa").Value, Homepage.sap_tirelli)
                    'Scheda_commessa_documentazione.riempi_datagridview_combinazioni()
                    ' Scheda_commessa_documentazione.riempi_datagridview_campioni()
                    Scheda_commessa_documentazione.cerca_file()
                    Scheda_commessa_documentazione.Show()



                    Scheda_commessa_documentazione.carico_iniziale = 1
                    Scheda_commessa_documentazione.inizializzazione = 1
                    'Attendibilità_info_popup.Close()
                End If
            ElseIf e.ColumnIndex = DataGridView.Columns.IndexOf(codice_Progetto) Then

                Progetto.Show()
                Progetto.BringToFront()
                Dim valore As String = DataGridView.Rows(e.RowIndex).Cells("absentry_progetto").Value.ToString()
                Dim numero As Integer = Integer.Parse(System.Text.RegularExpressions.Regex.Match(valore, "\d+").Value)
                Progetto.codice_progetto = valore
                Progetto.absentry = numero
                Progetto.inizializza_progetto()
            End If

        End If
    End Sub

    Private Async Sub dati_mancanti_Click(sender As Object, e As EventArgs) Handles dati_mancanti.Enter
        appunti_globali(Homepage.ID_SALVATO)
    End Sub

    Sub layout_scheda_tecnica()

        Scheda_commessa_documentazione.GroupBox34.Visible = False
        Scheda_commessa_documentazione.GroupBox35.Visible = False
        Scheda_commessa_documentazione.GroupBox37.Visible = False
        Scheda_commessa_documentazione.GroupBox38.Visible = False
        Scheda_commessa_documentazione.GroupBox39.Visible = False
        Scheda_commessa_documentazione.GroupBox40.Visible = False
        Scheda_commessa_documentazione.GroupBox44.Visible = False


        Scheda_commessa_documentazione.GroupBox42.Visible = False
        Scheda_commessa_documentazione.GroupBox43.Visible = False

        Scheda_commessa_documentazione.GroupBox46.Visible = False
        Scheda_commessa_documentazione.GroupBox47.Visible = False

        Scheda_commessa_documentazione.GroupBox58.Visible = False
        Scheda_commessa_documentazione.GroupBox59.Visible = False
        Scheda_commessa_documentazione.GroupBox60.Visible = False
    End Sub



    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave

        filtro_commesse()
    End Sub

    Private Sub TextBox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave
        filtro_commesse()
    End Sub

    Private Sub TextBox4_Leave(sender As Object, e As EventArgs) Handles TextBox4.Leave

    End Sub

    Private Sub TextBox3_Leave(sender As Object, e As EventArgs)

        filtro_commesse()
    End Sub




    Sub riempi_datagridview_campioni_new(par_tipo_campione As Integer, par_numero_risultati As Integer, par_datagridview_flaconi As DataGridView, par_datagridview_tappi As DataGridView, par_datagridview_sottotappi As DataGridView, par_datagridview_pompette As DataGridView, par_datagridview_etichette As DataGridView, par_datagridview_trigger As DataGridView, par_datagridview_prodotto As DataGridView, par_datagridview_film As DataGridView, par_datagridview_copritappo As DataGridView, par_datagridview_scatole As DataGridView)

        Dim par_datagridview As DataGridView

        If par_tipo_campione = 100 Then
            par_datagridview = par_datagridview_flaconi
        ElseIf par_tipo_campione = 101 Then
            par_datagridview = par_datagridview_tappi
        ElseIf par_tipo_campione = 102 Then
            par_datagridview = par_datagridview_sottotappi
        ElseIf par_tipo_campione = 103 Then
            par_datagridview = par_datagridview_pompette
        ElseIf par_tipo_campione = 104 Then
            par_datagridview = par_datagridview_etichette
        ElseIf par_tipo_campione = 105 Then
            par_datagridview = par_datagridview_trigger
        ElseIf par_tipo_campione = 106 Then
            par_datagridview = par_datagridview_prodotto
        ElseIf par_tipo_campione = 107 Then
            par_datagridview = par_datagridview_film
        ElseIf par_tipo_campione = 108 Then
            par_datagridview = par_datagridview_copritappo
        ElseIf par_tipo_campione = 109 Then
            par_datagridview = par_datagridview_scatole
        End If

        par_datagridview.Rows.Clear()


        Dim Cnn6 As New SqlConnection
        Cnn6.ConnectionString = Homepage.sap_tirelli
        Cnn6.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn6
        CMD_SAP_2.CommandText = "Select TOP " & par_numero_risultati & "     T0.[Id_Campione],
Case WHEN T0.[Codice_BP] Is NULL THEN '' ELSE T0.[Codice_BP] END AS 'Codice_BP',
    Case WHEN T0.[Nome] Is NULL THEN '' ELSE T0.[Nome] END AS 'Nome',
    Case WHEN T0.[Descrizione] Is NULL THEN '' ELSE T0.[Descrizione] END AS 'Descrizione',
    Case WHEN T0.[Codice_SAP] Is NULL THEN '' ELSE T0.[Codice_SAP] END AS 'Codice_SAP',
    Case WHEN T0.[Tipo_Campione] Is NULL THEN '' ELSE T0.[Tipo_Campione] END AS 'Tipo_Campione',
    Case WHEN (t0.immagine Is null Or t0.immagine ='' ) then 'N_A.JPG' ELSE t0.immagine END AS 'immagine'


 ,CASE WHEN t1.[Altezza] Is NULL THEN 0 ELSE t1.[Altezza] END AS 'Altezza_t1',
    Case WHEN t1.[Larghezza] Is NULL THEN 0 ELSE t1.[Larghezza] END AS 'Larghezza_t1',
    Case WHEN t1.[Profondita] Is NULL THEN 0 ELSE t1.[Profondita] END AS 'Profondita_t1',
    Case WHEN t1.[Diametro_Interno] Is NULL THEN 0 ELSE t1.[Diametro_Interno] END AS 'Diametro_Interno_t1',
  Case WHEN t1.[Diametro_Esterno] Is NULL THEN 0 ELSE t1.[Diametro_Esterno] END AS 'Diametro_Esterno_t1',
Case WHEN t1.[Volume] Is NULL THEN 0 ELSE t1.[Volume] END AS 'Volume_t1'


,CASE WHEN t1.[Spazio_Testa] Is NULL THEN 0 ELSE t1.[Spazio_Testa] END AS 'Spazio_Testa_t1',
Case WHEN t1.[Materiale] Is NULL THEN '' ELSE t1.[Materiale] END AS 'Materiale_t1',
Case WHEN t1.[Forma] Is NULL THEN '' ELSE t1.[Forma] END AS 'Forma_t1',
Case WHEN t1.[Sezione] Is NULL THEN '' ELSE t1.[Sezione] END AS 'Sezione_t1'


,CASE WHEN t1.[Superficie] Is NULL THEN '' ELSE t1.[Superficie] END AS 'Superficie_t1',
Case WHEN t1.[Produttore] Is NULL THEN '' ELSE t1.[Produttore] END AS 'Produttore_t1'

,CASE WHEN t1.[Codice_Produttore] Is NULL THEN '' ELSE t1.[Codice_Produttore] END AS 'Codice_Produttore_t1'

,CASE WHEN t1.[Collo_Centrato] Is NULL THEN '' ELSE t1.[Collo_Centrato]  END AS 'Collo_Centrato_t1'


,CASE WHEN t1.[Tipo_Tappo] Is NULL THEN '' ELSE t1.[Tipo_Tappo] END AS 'Tipo_Tappo_t1'


,CASE WHEN t1.[Filettatura] Is NULL THEN '' ELSE t1.[Filettatura] END AS 'Filettatura_t1',
Case WHEN t1.[Diametro_Esterno_Fil] Is NULL THEN 0 ELSE t1.[Diametro_Esterno_Fil] END AS 'Diametro_Esterno_Fil_t1',
Case WHEN t1.[Passo] Is NULL THEN 0 ELSE t1.[Passo] END AS 'Passo_t1',
Case WHEN t1.[Num_Principi] Is NULL THEN 0 ELSE t1.[Num_Principi] END AS 'Num_Principi_t1'

,CASE WHEN t2.[Altezza] Is NULL THEN 0 ELSE t2.[Altezza] END AS 'Altezza_t2',
Case WHEN t2.[Larghezza] Is NULL THEN 0 ELSE t2.[Larghezza] END AS 'Larghezza_t2',
Case WHEN t2.[Profondità] Is NULL THEN 0 ELSE t2.[Profondità] END AS 'Profondità_t2',
Case WHEN t2.[Diametro_Interno] Is NULL THEN 0 ELSE t2.[Diametro_Interno] END AS 'Diametro_Interno_t2',
Case WHEN t2.[Fissaggio] Is NULL THEN '' ELSE t2.[Fissaggio] END AS 'Fissaggio_t2',
Case WHEN t2.[Forma] Is NULL THEN '' ELSE t2.[Forma] END AS 'Forma_t2',
Case WHEN t2.[Materiale] Is NULL THEN '' ELSE t2.[Materiale] END AS 'Materiale_t2',
Case WHEN t2.[Superficie] Is NULL THEN '' ELSE t2.[Superficie] END AS 'Superficie_t2',
Case WHEN t2.[Produttore] Is NULL THEN '' ELSE t2.[Produttore] END AS 'Produttore_t2',
Case WHEN t2.[Codice_Produttore] Is NULL THEN '' ELSE t2.[Codice_Produttore] END AS 'Codice_Produttore_t2'

-- Campi t3
,CASE WHEN t3.[Altezza] Is NULL THEN 0 ELSE t3.[Altezza] END AS 'Altezza_t3',
Case WHEN t3.[Larghezza] Is NULL THEN 0 ELSE t3.[Larghezza] END AS 'Larghezza_t3',
Case WHEN t3.[Profondità] Is NULL THEN 0 ELSE t3.[Profondità] END AS 'Profondità_t3',
Case WHEN t3.[Diametro_Interno] Is NULL THEN 0 ELSE t3.[Diametro_Interno] END AS 'Diametro_Interno_t3',
Case WHEN t3.[Vite_Pressione] Is NULL THEN '' ELSE t3.[Vite_Pressione] END AS 'Vite_Pressione_t3',
Case WHEN t3.[Forma] Is NULL THEN '' ELSE t3.[Forma] END AS 'Forma_t3',
Case WHEN t3.[Materiale] Is NULL THEN '' ELSE t3.[Materiale] END AS 'Materiale_t3',


   Case WHEN t4.[A] Is NULL THEN 0 ELSE t4.[A] END AS 'A_t4',
Case WHEN t4.[B] Is NULL THEN 0 ELSE t4.[B] END AS 'B_t4',
Case WHEN t4.[C] Is NULL THEN 0 ELSE t4.[C] END AS 'C_t4',
Case WHEN t4.[D] Is NULL THEN 0 ELSE t4.[D] END AS 'D_t4',
Case WHEN t4.[Quota_A] Is NULL THEN 0 ELSE t4.[Quota_A] END AS 'Quota_A_t4',
Case WHEN t4.[Quota_B] Is NULL THEN 0 ELSE t4.[Quota_B] END AS 'Quota_B_t4',
Case WHEN t4.[Quota_C] Is NULL THEN 0 ELSE t4.[Quota_C] END AS 'Quota_C_t4',
Case WHEN t4.[Quota_D] Is NULL THEN 0 ELSE t4.[Quota_D] END AS 'Quota_D_t4',
Case WHEN t4.[Quota_E] Is NULL THEN 0 ELSE t4.[Quota_E] END AS 'Quota_E_t4',
Case WHEN t4.[Quota_F] Is NULL THEN 0 ELSE t4.[Quota_F] END AS 'Quota_F_t4',
Case WHEN t4.[Quota_L] Is NULL THEN 0 ELSE t4.[Quota_L] END AS 'Quota_L_t4',
Case WHEN t4.[SP] Is NULL THEN 0 ELSE t4.[SP] END AS 'SP_t4',
Case WHEN t4.[Materiale] Is NULL THEN '' ELSE t4.[Materiale] END AS 'Materiale_t4',
Case WHEN t4.[Tipologia] Is NULL THEN '' ELSE t4.[Tipologia] END AS 'Tipologia_t4',
Case WHEN t4.[Superficie] Is NULL THEN '' ELSE t4.[Superficie] END AS 'Superficie_t4',
Case WHEN t4.[Produttore] Is NULL THEN '' ELSE t4.[Produttore] END AS 'Produttore_t4',
Case WHEN t4.[cod_produttore] Is NULL THEN '' ELSE t4.[cod_produttore] END AS 'Codice_Produttore_t4',
Case WHEN t4.[Fissaggio] Is NULL THEN '' ELSE t4.[Fissaggio] END AS 'Fissaggio_t4',
Case WHEN t4.[Ghiera] Is NULL THEN '' ELSE t4.[Ghiera] END AS 'Ghiera_t4',
Case WHEN t4.[Copritappo] Is NULL THEN '' ELSE t4.[Copritappo] END AS 'Copritappo_t4'

-- Campi t5
,CASE WHEN t5.[Altezza] Is NULL THEN 0 ELSE t5.[Altezza] END AS 'Altezza_t5',
Case WHEN t5.[Larghezza] Is NULL THEN 0 ELSE t5.[Larghezza] END AS 'Larghezza_t5',
Case WHEN t5.[Trasparenza] Is NULL THEN '' ELSE t5.[Trasparenza] END AS 'Trasparenza_t5',
Case WHEN t5.[Forma] Is NULL THEN '' ELSE t5.[Forma] END AS 'Forma_t5',
Case WHEN t5.[Diametro_Esterno_Bobina] Is NULL THEN 0 ELSE t5.[Diametro_Esterno_Bobina] END AS 'Diametro_Esterno_Bobina_t5',
Case WHEN t5.[Diametro_Interno_Bobina] Is NULL THEN 0 ELSE t5.[Diametro_Interno_Bobina] END AS 'Diametro_Interno_Bobina_t5',
Case WHEN t5.[Avvolgimento_Bobina] Is NULL THEN '' ELSE t5.[Avvolgimento_Bobina] END AS 'Avvolgimento_Bobina_t5',
Case WHEN t5.[Materiale] Is NULL THEN '' ELSE t5.[Materiale] END AS 'Materiale_t5',

Case WHEN t6.[A] Is NULL THEN 0 ELSE t6.[A] END AS 'A_t6',
Case WHEN t6.[B] Is NULL THEN 0 ELSE t6.[B] END AS 'B_t6',
Case WHEN t6.[Quota_S] Is NULL THEN 0 ELSE t6.[Quota_S] END AS 'Quota_S_t6',
Case WHEN t6.[Quota_H] Is NULL THEN 0 ELSE t6.[Quota_H] END AS 'Quota_H_t6',
Case WHEN t6.[Quota_L] Is NULL THEN 0 ELSE t6.[Quota_L] END AS 'Quota_L_t6',
Case WHEN t6.[Quota_W] Is NULL THEN 0 ELSE t6.[Quota_W] END AS 'Quota_W_t6',
Case WHEN t6.[Quota_V] Is NULL THEN 0 ELSE t6.[Quota_V] END AS 'Quota_V_t6',
Case WHEN t6.[Pressione/Vite] Is NULL THEN '' ELSE t6.[Pressione/Vite] END AS 'Pressione/Vite_t6',
Case WHEN t6.[Produttore] Is NULL THEN '' ELSE t6.[Produttore] END AS 'Produttore_t6',
Case WHEN t6.[Codice_produttore] Is NULL THEN '' ELSE t6.[Codice_produttore] END AS 'Codice_Produttore_t6',
Case WHEN t6.[Materiale] Is NULL THEN '' ELSE t6.[Materiale] END AS 'Materiale_t6',
Case WHEN t6.[SP] Is NULL THEN 0 ELSE t6.[SP] END AS 'SP_t6',
Case WHEN t6.[T] Is NULL THEN 0 ELSE t6.[T] END AS 'T_t6',
Case WHEN t6.[Fissaggio] Is NULL THEN '' ELSE t6.[Fissaggio] END AS 'Fissaggio_t6',
Case WHEN t6.[Ghiera] Is NULL THEN '' ELSE t6.[Ghiera] END AS 'Ghiera_t6',
Case WHEN t6.[Grileltto] Is NULL THEN '' ELSE t6.[Grileltto] END AS 'Grileltto_t6',
Case WHEN t6.[Protezione] Is NULL THEN '' ELSE t6.[Protezione] END AS 'Protezione_t6',
Case WHEN t6.[Note] Is NULL THEN '' ELSE t6.[Note] END AS 'Note_t6',
Case WHEN t6.[Cannuccia] Is NULL THEN '' ELSE t6.[Cannuccia] END AS 'Cannuccia_t6',

-- Campi t7
Case WHEN t7.[Densita] Is NULL THEN 0 ELSE t7.[Densita] END AS 'Densita_t7',
Case WHEN t7.[Viscosita_Dinamica] Is NULL THEN 0 ELSE t7.[Viscosita_Dinamica] END AS 'Viscosita_Dinamica_t7',
Case WHEN t7.[Conducibilita_Elettrica] Is NULL THEN 0 ELSE t7.[Conducibilita_Elettrica] END AS 'Conducibilita_Elettrica_t7',
Case WHEN t7.[Categoria] Is NULL THEN '' ELSE t7.[Categoria] END AS 'Categoria_t7',
Case WHEN t7.[Infiammabile] Is NULL THEN '' ELSE t7.[Infiammabile] END AS 'Infiammabile_t7',
Case WHEN t7.[Nome_Commerciale] Is NULL THEN '' ELSE t7.[Nome_Commerciale] END AS 'Nome_Commerciale_t7',
Case WHEN t7.[Viscosità_Cinematica] Is NULL THEN 0 ELSE t7.[Viscosità_Cinematica] END AS 'Viscosità_Cinematica_t7',
Case WHEN t7.[Corrosivo] Is NULL THEN '' ELSE t7.[Corrosivo] END AS 'Corrosivo_t7',
Case WHEN t7.[Nocivo/Tossico] Is NULL THEN '' ELSE t7.[Nocivo/Tossico] END AS 'Nocivo/Tossico_t7',
Case WHEN t7.[Note] Is NULL THEN '' ELSE t7.[Note] END AS 'Note_t7',

-- Campi t8
Case WHEN t8.[Larghezza] Is NULL THEN 0 ELSE t8.[Larghezza] END AS 'Larghezza_t8',
Case WHEN t8.[Diametro_Fulcro] Is NULL THEN 0 ELSE t8.[Diametro_Fulcro] END AS 'Diametro_Fulcro_t8',
Case WHEN t8.[Materiale] Is NULL THEN '' ELSE t8.[Materiale] END AS 'Materiale_t8',
Case WHEN t8.[Temperatura_Saldatura] Is NULL THEN 0 ELSE t8.[Temperatura_Saldatura] END AS 'Temperatura_Saldatura_t8',
Case WHEN t8.[Diametro_Esterno] Is NULL THEN 0 ELSE t8.[Diametro_Esterno] END AS 'Diametro_Esterno_t8',

 Case WHEN t9.[Altezza] Is NULL THEN 0 ELSE t9.[Altezza] END AS 'Altezza_t9',
Case WHEN t9.[Larghezza] Is NULL THEN 0 ELSE t9.[Larghezza] END AS 'Larghezza_t9',
Case WHEN t9.[Profondità] Is NULL THEN 0 ELSE t9.[Profondità] END AS 'Profondità_t9',
Case WHEN t9.[Diametro_Interno] Is NULL THEN 0 ELSE t9.[Diametro_Interno] END AS 'Diametro_Interno_t9',
Case WHEN t9.[Fissaggio] Is NULL THEN '' ELSE t9.[Fissaggio] END AS 'Fissaggio_t9',
Case WHEN t9.[Forma] Is NULL THEN '' ELSE t9.[Forma] END AS 'Forma_t9',
Case WHEN t9.[Materiale] Is NULL THEN '' ELSE t9.[Materiale] END AS 'Materiale_t9',
Case WHEN t9.[Superficie] Is NULL THEN '' ELSE t9.[Superficie] END AS 'Superficie_t9',
Case WHEN t9.[Produttore] Is NULL THEN '' ELSE t9.[Produttore] END AS 'Produttore_t9',
Case WHEN t9.[Codice_produttore] Is NULL THEN '' ELSE t9.[Codice_produttore] END AS 'Codice_Produttore_t9',

-- Campi t10
Case WHEN t10.iniziale_sigla Is NULL THEN '' ELSE t10.iniziale_sigla END AS 'iniziale_sigla',

-- Campi t11
Case WHEN t11.onhand Is NULL THEN 0 ELSE cast(t11.onhand as integer) END AS 'onhand',
Case WHEN t11.u_ubicazione Is NULL THEN '' ELSE t11.u_ubicazione END AS 'u_ubicazione',

-- Campi t12
Case WHEN t12.cardname Is NULL THEN '' ELSE t12.cardname END AS 'cardname',

t13.immagine_descrizione

,T14.CARDNAME
,case when t15.cardcode Is null then '' else t15.cardcode end as 'codice_bp_principale'
,case when t15.cardname Is null then '' else t15.cardname end as 'Cliente_principale'

From [TIRELLI_40].[dbo].[coll_campioni] AS T0
Left Join [TIRELLI_40].[dbo].[coll_campioni_flaconi] t1 on t0.id_campione=t1.codice_campione And t0.tipo_campione=100
Left Join [TIRELLI_40].[dbo].[coll_campioni_tappi] t2 on t0.id_campione=t2.codice_campione And t0.tipo_campione=101
Left Join [TIRELLI_40].[dbo].[Coll_campioni_sottotappi] t3 on t0.id_campione=t3.codice_campione And t0.tipo_campione=102
Left Join [TIRELLI_40].[dbo].[Coll_campioni_pompette] t4 on t0.id_campione=t4.codice_campione And t0.tipo_campione=103
Left Join [TIRELLI_40].[dbo].[Coll_campioni_etichette] t5 on t0.id_campione=t5.codice_campione And t0.tipo_campione=104
Left Join [TIRELLI_40].[dbo].[Coll_campioni_trigger] t6 on t0.id_campione=t6.codice_campione And t0.tipo_campione=105
Left Join [TIRELLI_40].[dbo].[Coll_campioni_prodotti] t7 on t0.id_campione=t7.codice_campione And t0.tipo_campione=106
Left Join [TIRELLI_40].[dbo].[Coll_campioni_film] t8 on t0.id_campione=t8.codice_campione And t0.tipo_campione=107
Left Join [TIRELLI_40].[dbo].[Coll_campioni_copritappi] t9 on t0.id_campione=t9.codice_campione And t0.tipo_campione=108
Left Join [TIRELLI_40].[dbo].[COLL_Tipo_Campione] t10 on t10.Id_Tipo_Campione=t0.Tipo_Campione
Left Join [TIRELLISRLDB].[dbo].oitm t11 on t11.itemcode=t0.codice_sap
Left Join [TIRELLISRLDB].[dbo].ocrd t12 on t12.cardcode=t0.codice_BP
Left Join [TIRELLI_40].[dbo].coll_tipo_campione t13 on t13.id_tipo_campione=t0.tipo_campione
Left Join [TIRELLISRLDB].[dbo].OCRD T14 ON cast(T14.CARDCODE as varchar) = cast(t0.codice_bp as varchar)
Left Join [TIRELLISRLDB].[dbo].ocrd t15 on cast(t15.u_bp_riferimento as varchar)=cast(t14.cardcode as varchar)
Left  Join  [TIRELLI_40].dbo.[COLL_TIPO_CAMPIONE] t16 on t0.TIPO_campione= T16.ID_TIPO_CAMPIONE


WHERE T0.[Tipo_Campione] = " & par_tipo_campione & " " & filtro_cliente_campione & " " & filtro_id_Campione & " " & filtro_nome_campione & "" & filtro_codice_sap_campione & "

        order by T14.CARDNAME , t10.INIZIALE_SIGLA, cast(T0.NOME As Integer)"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If par_tipo_campione = 100 Then
            Do While cmd_SAP_reader_2.Read()
                Try


                    par_datagridview.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("codice_sap"), cmd_SAP_reader_2("u_ubicazione"), cmd_SAP_reader_2("onhand"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("cliente_principale"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("codice_bp_principale"))
                Catch ex As Exception
                    par_datagridview.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("codice_sap"), cmd_SAP_reader_2("u_ubicazione"), cmd_SAP_reader_2("onhand"), Image.FromFile(Homepage.Percorso_immagini & "N_A.JPG"), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("cliente_principale"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("codice_bp_principale"))
                End Try
            Loop
        ElseIf par_tipo_campione = 101 Then
            Do While cmd_SAP_reader_2.Read()

                par_datagridview.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("codice_sap"), cmd_SAP_reader_2("u_ubicazione"), cmd_SAP_reader_2("onhand"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("cliente_principale"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("codice_bp_principale"))

            Loop
        ElseIf par_tipo_campione = 102 Then
            Do While cmd_SAP_reader_2.Read()

                par_datagridview.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("codice_sap"), cmd_SAP_reader_2("u_ubicazione"), cmd_SAP_reader_2("onhand"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("cliente_principale"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("codice_bp_principale"))

            Loop
        ElseIf par_tipo_campione = 103 Then
            Do While cmd_SAP_reader_2.Read()

                par_datagridview.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("codice_sap"), cmd_SAP_reader_2("u_ubicazione"), cmd_SAP_reader_2("onhand"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("cliente_principale"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("codice_bp_principale"))

            Loop
        ElseIf par_tipo_campione = 104 Then
            Do While cmd_SAP_reader_2.Read()

                par_datagridview.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("codice_sap"), cmd_SAP_reader_2("u_ubicazione"), cmd_SAP_reader_2("onhand"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("cliente_principale"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("codice_bp_principale"))

            Loop
        ElseIf par_tipo_campione = 105 Then
            Do While cmd_SAP_reader_2.Read()
                Try
                    par_datagridview.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("codice_sap"), cmd_SAP_reader_2("u_ubicazione"), cmd_SAP_reader_2("onhand"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("cliente_principale"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("codice_bp_principale"))

                Catch ex As Exception

                    Form_campione_visualizza.id_campione = cmd_SAP_reader_2("id_campione")
                    Form_campione_visualizza.Show()
                    Form_campione_visualizza.inizializza_form()

                End Try

            Loop
        ElseIf par_tipo_campione = 106 Then
            Do While cmd_SAP_reader_2.Read()

                par_datagridview.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("codice_sap"), cmd_SAP_reader_2("u_ubicazione"), cmd_SAP_reader_2("onhand"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("cliente_principale"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("codice_bp_principale"))

            Loop
        ElseIf par_tipo_campione = 107 Then
            Do While cmd_SAP_reader_2.Read()

                par_datagridview.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("codice_sap"), cmd_SAP_reader_2("u_ubicazione"), cmd_SAP_reader_2("onhand"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("cliente_principale"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("codice_bp_principale"))

            Loop

        ElseIf par_tipo_campione = 108 Then
            Do While cmd_SAP_reader_2.Read()
                Try
                    par_datagridview.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("codice_sap"), cmd_SAP_reader_2("u_ubicazione"), cmd_SAP_reader_2("onhand"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("cliente_principale"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("codice_bp_principale"))

                Catch ex As Exception


                    Form_campione_visualizza.id_campione = cmd_SAP_reader_2("id_campione")
                    Form_campione_visualizza.Show()
                    Form_campione_visualizza.inizializza_form()

                End Try

            Loop

        ElseIf par_tipo_campione = 109 Then
            Do While cmd_SAP_reader_2.Read()
                Try
                    par_datagridview.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("codice_sap"), cmd_SAP_reader_2("u_ubicazione"), cmd_SAP_reader_2("onhand"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("cliente_principale"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("codice_bp_principale"))

                Catch ex As Exception


                    Form_campione_visualizza.id_campione = cmd_SAP_reader_2("id_campione")
                    Form_campione_visualizza.Show()
                    Form_campione_visualizza.inizializza_form()

                End Try

            Loop
        End If



        cmd_SAP_reader_2.Close()
        Cnn6.Close()
    End Sub









    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If TextBox6.Text = "" Then
            filtro_cliente_campione = ""
        Else
            filtro_cliente_campione = "And (t12.cardname   Like '%%" & TextBox6.Text & "%%' or t15.cardname Like '%%" & TextBox6.Text & "%%')"
        End If

    End Sub


    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs)
        filtro_commesse()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        filtro_commesse()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        'If TextBox4.Text = Nothing Then
        '    filtro_cliente_f = ""
        'Else
        '    filtro_cliente_f = TextBox4.Text
        'End If
        filtro_commesse()
    End Sub



    Private Sub tabpage14_Click(sender As Object, e As EventArgs) Handles TabPage14.Enter



        riempi_datagridview_fatturate(DataGridView12, TextBox20.Text, TextBox19.Text, TextBox14.Text, TextBox23.Text, TextBox25.Text, DateTimePicker4, DateTimePicker1)


    End Sub

    '    Sub Costruzione_datagridview_progetti()



    '        Dim i As Integer = 4
    '        Dim Cnn1 As New SqlConnection
    '        Cnn1.ConnectionString = Homepage.sap_tirelli
    '        Cnn1.Open()


    '        Dim CMD_SAP_2 As New SqlCommand
    '        Dim cmd_SAP_reader_2 As SqlDataReader


    '        CMD_SAP_2.Connection = Cnn1
    '        CMD_SAP_2.CommandText = "SELECT *
    'FROM [TIRELLI_40].[DBO].[Requisiti_progetto] 
    'where active='Y' and riepilogo='Y' and documento='Progetto' order by id"

    '        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



    '        Do While cmd_SAP_reader_2.Read()
    '            Dim column As New DataGridViewTextBoxColumn()
    '            column.Name = cmd_SAP_reader_2("codice_requisito")
    '            column.HeaderText = cmd_SAP_reader_2("Nome requisito")
    '            Try
    '                DataGridView10.Columns.Add(column)
    '            Catch ex As Exception
    '                Console.WriteLine(i)
    '            End Try

    '            i = i + 1
    '        Loop


    '        cmd_SAP_reader_2.Close()
    '        Cnn1.Close()
    '    End Sub

    Sub riempi_datagridview_progetti(PAR_DATAGRIDVIEW As DataGridView)

        PAR_DATAGRIDVIEW.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select T0.ABSENTRY,T0.DOCNUM,coalesce(A.NOME,'') as 'Stato',coalesce(A.tipo,''),
coalesce(a.ordine,9999) as 'Ordine', T0.CARDCODE, T0.CARDNAME, T0.U_CODICE_CLIENTE_FINALE, T0.U_CLIENTE_FINALE,t0.name, CONCAT(T1.LASTNAME,' ',T1.FIRSTNAME) AS 'PM', T2.SLPNAME, coalesce(CONCAT(T3.LASTNAME,' ',T3.FIRSTNAME),'') AS 'Resp_acq'
from opmg T0 LEFT JOIN [TIRELLI_40].[dbo].OHEM T1 ON T0.OWNER=T1.EMPID
LEFT JOIN OSLP T2 ON T2.SLPCODE=T0.EMPLOYEE
left join [TIRELLI_40].[dbo].ohem t3 on t3.empid=t0.U_Resp_acquisti

left join
(
select t10.n_progetto, t12.NOME,t12.tipo, t12.ordine
from
(
select n_progetto, max(rev) as 'Rev_max'
from
[Tirelli_40].[dbo].[Scheda_Tecnica_valori_progetto]
group by n_progetto
)

as t10
left join [Tirelli_40].[dbo].[Scheda_tecnica_revisioni_progetto] t11 on t11.n_progetto=t10.n_progetto and t11.Numero=t10.rev_max
LEFT JOIN [Tirelli_40].[dbo].[Scheda_Tecnica_stato_progetto] T12 ON T12.ID=T11.STATO
) A on A.n_progetto=t0.docnum


where T0.series>=2075 AND T0.STATUS='S' " & filtro_numero_progetto & filtro_stato_progetto & filtro_stato_rev_progetto & " " & filtro_cliente_progetto & "  " & filtro_nome_progetto & " " & filtro_PM & filtro_acq & "  
order by coalesce(a.ordine,9999), t0.absentry DESC"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader_2.Read()

            PAR_DATAGRIDVIEW.Rows.Add(cmd_SAP_reader_2("absentry"), cmd_SAP_reader_2("docnum"), cmd_SAP_reader_2("cardcode"), cmd_SAP_reader_2("Stato"), cmd_SAP_reader_2("cardname"), cmd_SAP_reader_2("U_CODICE_CLIENTE_FINALE"), cmd_SAP_reader_2("U_CLIENTE_FINALE"), cmd_SAP_reader_2("name"), cmd_SAP_reader_2("PM"), cmd_SAP_reader_2("Resp_acq"), cmd_SAP_reader_2("Slpname"))

        Loop


        cmd_SAP_reader_2.Close()
        Cnn1.Close()
    End Sub

    Sub riempi_datagridview_fatturate(par_datagridview As DataGridView, par_commessa As String, par_descrizione As String, par_cliente As String, par_progetto As String, par_brand As String, par_datetimepicker_inizio As DateTimePicker, par_datetimepicker_fine As DateTimePicker)

        Dim filtro_commessa As String
        Dim filtro_descrizione As String
        Dim filtro_cliente As String
        Dim filtro_progetto As String
        Dim filtro_brand As String

        If par_commessa <> "" Then
            filtro_commessa = " AND T10.[ItemCode] LIKE '%" & par_commessa & "%'"
        Else
            filtro_commessa = ""
        End If

        If par_descrizione <> "" Then
            filtro_descrizione = " AND T10.[ItemName] LIKE '%" & par_descrizione & "%'"
        Else
            filtro_descrizione = ""
        End If

        If par_cliente <> "" Then
            filtro_cliente = " AND (T10.[CardName] LIKE '%" & par_cliente & "%' or T10.U_CLIENTEFINALE LIKE '%" & par_cliente & "%') "
        Else
            filtro_cliente = ""
        End If

        If par_progetto <> "" Then
            filtro_progetto = " AND t10.u_progetto = '" & par_progetto & "'"
        Else
            filtro_progetto = ""
        End If

        If par_brand <> "" Then
            filtro_brand = " AND COALESCE(t10.u_brand,'') LIKE '%" & par_brand & "%'"
        Else
            filtro_brand = ""
        End If


        par_datagridview.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select  T10.[DocNum], T10.[DocDate],t10.[Fiscal year], T10.[ItemCode], T10.[ItemName],t10.u_progetto, t10.u_brand, T10.[CardName], T10.U_CLIENTEFINALE,t10.country, t10.[Tirelli salesman],T10.name, t10.[Total_amount], T10.[Pricing list], t10.Discount, t10.Provv, (T10.[Pricing list]*(1- t10.discount))*(1-t10.provv) as 'MultiplierF', case when (T10.[Pricing list]*(1- t10.discount)) <>0 then  t10.[Total_amount]/(T10.[Pricing list]*(1- t10.discount)) end as 'Costo preventivato'
, coalesce(a.Costo_componenti,0) as 'Costo_componenti'
,t10.[Total_amount]-coalesce(a.Costo_componenti,0) as 'GM',
case when t10.[Total_amount]=0 then 0 else (t10.[Total_amount]-coalesce(a.Costo_componenti,0))/t10.[Total_amount] end as 'GM_perc'

, coalesce(c.valore,0) as 'Manodopera'
,t10.[Total_amount]-coalesce(a.Costo_componenti,0)-coalesce(c.valore,0) as 'Margin' 
,case when t10.[Total_amount]=0 then 0 else (t10.[Total_amount]-coalesce(a.Costo_componenti,0)-coalesce(c.valore,0))/t10.[Total_amount] end as 'M_perc'
from
(
SELECT T0.[DocNum], T0.[DocDate],case when month(T0.[DocDate])>9 then year(T0.[DocDate])+1 else year(T0.[DocDate]) end as 'Fiscal year', T1.[ItemCode]
, T2.[ItemName], coalesce(t2.U_brand,case when t1.ocrcode='TIR01' THEN 'TIRELLI' WHEN T1.OCRCODE='KTF01' THEN 'KTF' WHEN T1.OCRCODE='BRB01' THEN 'BRB' ELSE'' END) AS 'U_BRAND' , t2.u_progetto,   T0.[CardName], T0.U_CLIENTEFINALE, t4.slpname as 'Tirelli salesman',T7.name
, case when t0.U_causcons='V' then sum(case when t0.doctype='S'  then t1.linetotal else (t1.quantity*price)*((100 - case when t0.discprcnt is null then 0 else t0.discprcnt end)/100)/t0.docrate end) else 0 end as 'Total_amount',sum(T1.LINETOTAL*((100-case when t0.discprcnt is null then 0 else t0.discprcnt end)/100))  AS 'Net total amount', T1.[U_coefficiente_vendita] as 'Pricing list', case when t0.doctype='S' then sum(t1.pricebefdi/case when t1.rate=0 then 1 else t1.rate end) else sum(t1.quantity*T1.[PriceBefDi]/case when t1.rate ='0' then 1 else t1.rate end) end as 'Total', 
1-(case when case when t0.doctype='S' then sum(t1.pricebefdi/case when t1.rate=0 then 1 else t1.rate end) else sum(t1.quantity*T1.[PriceBefDi]/case when t1.rate ='0' then 1 else t1.rate end) end = '0' then '0' else sum(T1.LINETOTAL*((100-case when t0.discprcnt is null then 0 else t0.discprcnt end)/100))/case when t0.doctype='S' then sum(t1.pricebefdi/case when t1.rate=0 then 1 else t1.rate end) else sum(t1.quantity*T1.[PriceBefDi]/case when t1.rate ='0' then 1 else t1.rate end) end  end) AS 'Discount'
, t1.commission/100 as 'Provv', coalesce(t6.country,'') as 'Country'
from
(
select t0.itemcode, max(t0.docentry) as 'Max'
from inv1 t0
group by t0.itemcode
)
as t9999

INNER JOIN INV1 T1 ON T9999.[max] = T1.[DocEntry] and t9999.itemcode=t1.itemcode
inner join oinv t0 on t0.docentry=t1.docentry
INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] 
left join [TIRELLI_40].[dbo].OHEM T3 ON t3.empid=t0.ownercode
LEFT join OSLP T4 ON T4.slpcode =t0.slpcode
INNER JOIN OCRD T6 ON T6.CARDCODE=T0.CARDCODE
left join ocry t7 on t7.code=t6.country


WHERE SUBSTRING( T1.[ItemCode] ,1,1)='M' 

group by T0.[DocNum], T0.[DocDate],T1.[ItemCode],T2.[ItemName],t2.u_brand, t1.ocrcode,  t2.u_progetto,t6.country,   T0.[CardName], T0.U_CLIENTEFINALE, t4.slpname,t0.U_causcons, t1.u_coefficiente_vendita, t0.doctype,t1.commission, t7.name
) as t10

left join owor t11 on t11.itemcode=t10.itemcode and t11.postdate<=T10.[DocDate] AND T11.STATUS<>'C'

inner join (select t0.itemcode, max(t0.docentry) as docentry from owor t0 where substring(t0.itemcode,1,1)='M'   AND T0.STATUS<>'C' group by t0.itemcode) B on b.docentry=t11.docentry

left join (
select t20.U_PRG_AZS_Commessa, sum(t20.costo_componenti) as 'Costo_componenti'
from
(
select t10.U_PRG_AZS_Commessa, sum(t10.costo_tot) as 'Costo_componenti'
from
(
select t0.U_PRG_AZS_Commessa, t1.itemcode, t3.itemname, t4.ItmsGrpNam,t3.u_PRG_TIR_materiale,t1.PlannedQty,t1.u_prg_wip_qtaspedita, case when t1.U_Prezzolis is null then t5.price else t1.u_prezzolis end as 'Costo U'
, case when t1.U_Prezzolis is null then t5.price else t1.u_prezzolis end*t1.PlannedQty as 'Costo_Tot'
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
left join owor t2 on t2.itemcode=t1.itemcode and t0.U_PRG_AZS_Commessa=t2.U_PRG_AZS_Commessa and t2.status<>'C'
left join oitm t3 on t3.itemcode=t1.itemcode
left join oitb t4 on t4.ItmsGrpCod=t3.ItmsGrpCod
left join itm1 t5 on t5.itemcode=t1.itemcode
where t1.PlannedQty>=0 and substring(t0.U_PRG_AZS_Commessa,1,1) ='M' AND t0.status<>'C' and t2.docnum is null and t1.itemtype=4 and t5.pricelist=2
)
as t10
group by t10.U_PRG_AZS_Commessa

union all

select  t0.U_PRG_AZS_Commessa
,sum(case when t0.u_costo is null OR T0.U_COSTO=0 then t5.price else t0.u_costo end*t0.Quantity) as 'Costo_Tot'

from rdr1 t0 inner join ordr t1 on t0.docentry=t1.docentry
left join owor t2 on t2.itemcode=t0.itemcode and t0.U_PRG_AZS_Commessa=t2.U_PRG_AZS_Commessa  AND T2.STATUS<>'C'
left join oitm t3 on t3.itemcode=t0.itemcode
left join oitb t4 on t4.ItmsGrpCod=t3.ItmsGrpCod
left join itm1 t5 on t5.itemcode=t0.itemcode
where substring(t0.U_PRG_AZS_Commessa,1,1) ='M' AND T1.CANCELED<>'Y' and t2.docnum is null and t0.itemtype=4 and t5.pricelist=2 and substring(t1.u_causcons,1,4)='COMP'

group by t0.U_PRG_AZS_Commessa
)
as t20

group by t20.U_PRG_AZS_Commessa
) A on A.U_PRG_AZS_Commessa=t10.ItemCode

left join
(
select t10.commessa, sum(t10.minuti*t10.price) as 'Valore'
from
(
SELECT  T3.[U_PRG_AZS_Commessa] as 'commessa',  
case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end as 'Minuti'
,  T7.PRICE
FROM MANODOPERA t0 LEFT JOIN [TIRELLI_40].[dbo].OHEM T1 ON T1.empid=T0.DIPENDENTE
left join orsc t2 on t2.visrescode=t0.risorsa
left join owor t3 on t3.docnum=t0.docnum and t0.tipo_documento='ODP'
LEFT JOIN OITM t4 ON T4.ITEMCODE=T3.ITEMCODE
left join oitm t5 on t5.itemcode=T3.[U_PRG_AZS_Commessa]
left join [TIRELLI_40].[dbo].oudp t6 on t1.dept=t6.code
LEFT JOIN ITM1 T7 ON T7.ITEMCODE=T0.RISORSA AND T7.PRICELIST=2

where substring(T3.[U_PRG_AZS_Commessa],1,1)='M'

)
as t10
group by t10.commessa
) C on c.commessa=t10.ItemCode
where 0=0  and t10.docdate>=CONVERT(DATETIME, '" & par_datetimepicker_inizio.Value & "', 103) and t10.docdate<=CONVERT(DATETIME, '" & par_datetimepicker_fine.Value & "', 103)" & filtro_commessa & filtro_descrizione & filtro_cliente & filtro_brand & filtro_progetto & "
order by T10.[DocDate] DESC, t10.docnum DESC"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Dim totale_amount As Decimal = 0
        Dim totale_costo_componenti As Decimal = 0
        Dim contatore As Integer = 0
        Dim gm_total As Integer = 0
        Dim manodopera_tot As Integer = 0
        Dim Margin_total As Integer = 0


        Do While cmd_SAP_reader_2.Read()
            Dim totalAmount As Decimal = Convert.ToDecimal(cmd_SAP_reader_2("Total_amount"))
            Dim costoComponenti As Decimal = Convert.ToDecimal(cmd_SAP_reader_2("Costo_componenti"))
            Dim GM As Decimal = Convert.ToDecimal(cmd_SAP_reader_2("GM"))
            Dim manodopera As Decimal = Convert.ToDecimal(cmd_SAP_reader_2("Manodopera"))
            Dim margin As Decimal = Convert.ToDecimal(cmd_SAP_reader_2("Margin"))

            totale_amount += totalAmount
            totale_costo_componenti += costoComponenti
            gm_total += GM
            manodopera_tot += manodopera
            Margin_total += margin

            par_datagridview.Rows.Add(
            cmd_SAP_reader_2("DocNum"),
            cmd_SAP_reader_2("DocDate"),
            cmd_SAP_reader_2("Fiscal year"),
            cmd_SAP_reader_2("ItemCode"),
            cmd_SAP_reader_2("ItemName"),
            cmd_SAP_reader_2("u_progetto"),
            cmd_SAP_reader_2("u_brand"),
            cmd_SAP_reader_2("CardName"),
            cmd_SAP_reader_2("U_CLIENTEFINALE"),
            cmd_SAP_reader_2("country"),
            totalAmount,
            costoComponenti,
            cmd_SAP_reader_2("GM"),
            cmd_SAP_reader_2("GM_perc"),
            cmd_SAP_reader_2("Manodopera"),
            cmd_SAP_reader_2("Margin"),
            cmd_SAP_reader_2("M_perc"))
            contatore += 1
        Loop

        Label2.Text = totale_amount.ToString("N0")
        Label3.Text = totale_costo_componenti.ToString("N0")
        Try
            Label4.Text = ((totale_amount - totale_costo_componenti) / totale_amount).ToString("P2")
        Catch ex As Exception

        End Try
        Label5.Text = gm_total.ToString("N0")

        Label6.Text = contatore

        Label7.Text = manodopera_tot.ToString("N0")
        Label8.Text = Margin_total.ToString("N0")

        Try
            Label9.Text = ((totale_amount - totale_costo_componenti - manodopera_tot) / totale_amount).ToString("P2")
        Catch ex As Exception

        End Try


        cmd_SAP_reader_2.Close()
        Cnn1.Close()
    End Sub





    Private Sub DataGridView10_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub











    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        'If TextBox14.Text = Nothing Then
        '    filtro_n_progetto = ""
        'Else
        '    filtro_n_progetto = "AND t20.docnum = '" & TextBox14.Text & "'"
        'End If
        filtro_commesse()
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        If TextBox15.Text = Nothing Then
            filtro_nome_progetto_commessa = ""
        Else
            filtro_nome_progetto_commessa = "AND t20.name Like '%%" & TextBox15.Text & "%%'"
        End If
        filtro_commesse()
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged

        filtro_commesse()
    End Sub



    Public Function GetFirstPart(fullText As String) As String
        Dim splittedText() As String = fullText.Split(" "c)
        If splittedText.Length > 0 Then
            Return splittedText(0)
        Else
            Return ""
        End If
    End Function

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
        If TextBox17.Text = "" Then
            filtro_nome_campione = ""
        Else
            filtro_nome_campione = "and (t16.INIZIALE_SIGLA + T0.NOME   Like '%%" & TextBox17.Text & "%%')"


        End If

    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged
        If TextBox18.Text = "" Then
            filtro_codice_sap_campione = ""
        Else
            filtro_codice_sap_campione = "and (t0.codice_sap Like '%%" & TextBox18.Text & "%%')"


        End If

    End Sub





    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form_nuovo_campione.Show()
        Form_nuovo_campione.inizializza_form()

    End Sub



    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub TextBox3_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Then
            filtro_id_Campione = ""
        Else
            filtro_id_Campione = " and t0 = '" & TextBox3.Text & "' "
        End If

    End Sub





    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.RowIndex >= 0 Then




            Form_campione_visualizza.id_campione = DataGridView2.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn2").Value
            Form_campione_visualizza.Show()
            Form_campione_visualizza.inizializza_form()

        End If
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 1 Then


            tipo_campione = TabControl3.SelectedIndex + 100
            riempi_datagridview_campioni_new(tipo_campione, TextBox5.Text, DataGridView1, DataGridView2, DataGridView4, DataGridView5, DataGridView6, DataGridView7, DataGridView8, DataGridView9, DataGridView11, DataGridView3)
        End If
    End Sub

    Private Sub TabControl3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl3.SelectedIndexChanged
        tipo_campione = TabControl3.SelectedIndex + 100
        riempi_datagridview_campioni_new(tipo_campione, TextBox5.Text, DataGridView1, DataGridView2, DataGridView4, DataGridView5, DataGridView6, DataGridView7, DataGridView8, DataGridView9, DataGridView11, DataGridView3)

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then




            Form_campione_visualizza.id_campione = DataGridView1.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn1").Value
            Form_campione_visualizza.Show()
            Form_campione_visualizza.inizializza_form()

        End If
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        If e.RowIndex >= 0 Then




            Form_campione_visualizza.id_campione = DataGridView4.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn3").Value
            Form_campione_visualizza.Show()
            Form_campione_visualizza.inizializza_form()

        End If
    End Sub

    Private Sub DataGridView5_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellContentClick

    End Sub

    Private Sub DataGridView5_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellClick
        If e.RowIndex >= 0 Then




            Form_campione_visualizza.id_campione = DataGridView5.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn4").Value
            Form_campione_visualizza.Show()
            Form_campione_visualizza.inizializza_form()

        End If
    End Sub

    Private Sub DataGridView6_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellContentClick

    End Sub

    Private Sub DataGridView6_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellClick
        If e.RowIndex >= 0 Then




            Form_campione_visualizza.id_campione = DataGridView6.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn5").Value
            Form_campione_visualizza.Show()
            Form_campione_visualizza.inizializza_form()

        End If
    End Sub

    Private Sub DataGridView7_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView7.CellContentClick

    End Sub

    Private Sub DataGridView7_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView7.CellClick
        If e.RowIndex >= 0 Then




            Form_campione_visualizza.id_campione = DataGridView7.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn6").Value
            Form_campione_visualizza.Show()
            Form_campione_visualizza.inizializza_form()

        End If
    End Sub

    Private Sub DataGridView8_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView8.CellContentClick

    End Sub

    Private Sub DataGridView8_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView8.CellClick
        If e.RowIndex >= 0 Then




            Form_campione_visualizza.id_campione = DataGridView8.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn7").Value
            Form_campione_visualizza.Show()
            Form_campione_visualizza.inizializza_form()

        End If
    End Sub

    Private Sub DataGridView9_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView9.CellClick
        If e.RowIndex >= 0 Then




            Form_campione_visualizza.id_campione = DataGridView9.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn8").Value
            Form_campione_visualizza.Show()
            Form_campione_visualizza.inizializza_form()

        End If
    End Sub



    Private Sub DataGridView11_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView11.CellClick
        If e.RowIndex >= 0 Then




            Form_campione_visualizza.id_campione = DataGridView11.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn9").Value
            Form_campione_visualizza.Show()
            Form_campione_visualizza.inizializza_form()

        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If TextBox3.Text = "" Then
            riempi_datagridview_campioni_new(tipo_campione, TextBox5.Text, DataGridView1, DataGridView2, DataGridView4, DataGridView5, DataGridView6, DataGridView7, DataGridView8, DataGridView9, DataGridView11, DataGridView3)
        Else
            trova_tipo_campione(TextBox3.Text)
            TabControl3.SelectedIndex = tipo_campione - 100
            riempi_datagridview_campioni_new(tipo_campione, TextBox5.Text, DataGridView1, DataGridView2, DataGridView4, DataGridView5, DataGridView6, DataGridView7, DataGridView8, DataGridView9, DataGridView11, DataGridView3)
        End If



    End Sub

    Sub trova_tipo_campione(par_id_Campione As Integer)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT [Tipo_Campione]
FROM [TIRELLISRLDB].[dbo].[coll_campioni] where id_Campione=" & par_id_Campione & " "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        If cmd_SAP_reader_2.Read() Then

            tipo_campione = cmd_SAP_reader_2("tipo_campione")

        End If


        cmd_SAP_reader_2.Close()
        Cnn1.Close()
    End Sub

    Private Sub DataGridView_CellBorderStyleChanged(sender As Object, e As EventArgs) Handles DataGridView.CellBorderStyleChanged

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Albero.Show()
        Albero.commessa = Codice_commessa
        Albero.TextBox1.Text = Albero.commessa
        Albero.inizializza_albero(Codice_commessa)

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.Text = Nothing Then
            filtro_desc_sup = ""
        Else
            filtro_desc_sup = " AND t20.desc_supp  Like '%%" & TextBox7.Text & "%%' "
        End If
        filtro_commesse()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        CREA_CARTELLO(Homepage.percorso_server & "00-Tirelli 4.0\File\Vari\CARTELLO COMMESSA A 3.xlsx", "Cartello", Codice_commessa)
        Beep()
        MsgBox("Cartello creato con successo")
    End Sub


    Sub CREA_CARTELLO(par_percorso_file As String, par_nome_foglio As String, par_commessa As String)

        Dim VELOCITA As Integer = 0

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.ITEMCODE, T0.ITEMNAME, substring(COALESCE(T3.CARDNAME,COALESCE(T2.CARDNAME,T0.U_FINAL_CUSTOMER_NAME)),1,25) AS 'CLIENTE', coalesce(t4.velocita,0) as 'Velocita'
,coalesce(t2.u_destinazione,t0.u_country_of_delivery) as 'u_destinazione'
    FROM [TIRELLISRLDB].[DBO].OITM T0
    LEFT JOIN (SELECT MAX(T1.DOCENTRY) AS 'DOCENTRY' 
FROM [TIRELLISRLDB].[DBO].RDR1 T1 WHERE T1.ITEMCODE='" & par_commessa & "') A ON 1=1
    LEFT JOIN [TIRELLISRLDB].[DBO].ORDR T2 ON T2.DOCENTRY=A.DOCENTRY
    LEFT JOIN [TIRELLISRLDB].[DBO].OCRD T3 ON T3.CARDCODE=T2.U_CodiceBP
    left join (select max(t0.numero) as 'Numero' from [TIRELLI_40].[dbo].[Scheda_tecnica_revisioni] t0 where t0.commessa='" & par_commessa & "') B ON 1=1
    left join [TIRELLI_40].[dbo].[Scheda_Tecnica_valori] t4 on t4.rev=B.numero and t4.commessa='" & par_commessa & "'

    WHERE T0.ITEMCODE='" & par_commessa & "' "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True

        Dim colonna1 As String
        Dim colonna2 As String
        Dim colonna3 As String
        Dim colonna4 As String
        Dim colonna5 As String
        Dim colonna6 As String

        If cmd_SAP_reader_2.Read() Then

            Excel.Sheets(par_nome_foglio).Cells(10, 1).value = cmd_SAP_reader_2("itemcode")
            Excel.Sheets(par_nome_foglio).Cells(17, 1).value = cmd_SAP_reader_2("itemname")
            Excel.Sheets(par_nome_foglio).Cells(32, 1).value = cmd_SAP_reader_2("cliente")

            VELOCITA = cmd_SAP_reader_2("velocita")
            Excel.Sheets(par_nome_foglio).Cells(32, 9).value = cmd_SAP_reader_2("u_destinazione")
            ' Inserire l'immagine nella cella (40, 40)
            Try


                Excel.Sheets(par_nome_foglio).Pictures.Insert("\\tirfs01\tirelli\00-Tirelli 4.0\Immagini\Flags\" & cmd_SAP_reader_2("u_destinazione") & ".png").Select
                With Excel.Selection
                    '.ShapeRange.LockAspectRatio = msoTrue

                    .ShapeRange.Height = Excel.Sheets(par_nome_foglio).Cells(32, 14).Height * 0.9
                    .ShapeRange.Width = Excel.Sheets(par_nome_foglio).Cells(32, 14).Width * 0.9
                    .Top = Excel.Sheets(par_nome_foglio).Cells(32, 14).Top
                    .Left = Excel.Sheets(par_nome_foglio).Cells(32, 14).Left
                End With
            Catch ex As Exception

                MsgBox("Bandiera " & cmd_SAP_reader_2("u_destinazione") & " non presente. Inserire in \\tirfs01\tirelli\00-Tirelli 4.0\Immagini\Flags\")
            End Try

        End If


        cmd_SAP_reader_2.Close()
        Cnn1.Close()




        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT t0.id_combinazione,t0.vel_richiesta, t0.campione_1, t11.INIZIALE_SIGLA + T1.NOME   as 'Nome_1',t1.immagine as 'Immagine_1', t0.campione_2, t12.INIZIALE_SIGLA + T2.NOME  as 'Nome_2', t2.immagine as 'Immagine_2', t0.campione_3,t13.INIZIALE_SIGLA + T3.NOME  as 'Nome_3',t3.immagine as 'immagine_3', t0.campione_4,t14.INIZIALE_SIGLA + T4.NOME  as 'Nome_4',t4.immagine as 'immagine_4', t0.campione_5,t15.INIZIALE_SIGLA + T5.NOME  as 'Nome_5',t5.immagine as 'immagine_5', t0.campione_6,t16.INIZIALE_SIGLA + T6.NOME  as 'Nome_6' ,t6.immagine as 'immagine_6', t0.campione_7, t17.INIZIALE_SIGLA + T7.NOME  as 'Nome_7',t7.immagine as 'immagine_7', t0.campione_8,t18.INIZIALE_SIGLA + T8.NOME  as 'Nome_8',t8.immagine as 'immagine_8', t0.campione_9,t19.INIZIALE_SIGLA + T9.NOME  as 'Nome_9',t9.immagine as 'immagine_9', t0.campione_10
,t20.INIZIALE_SIGLA + T10.NOME  as 'Nome_10',t10.immagine as 'immagine_10'

FROM [TIRELLI_40].[dbo].COLL_COMBINAZIONI t0
left join [TIRELLI_40].[dbo].coll_campioni t1 on t0.campione_1=t1.id_campione
left join [TIRELLI_40].[dbo].coll_campioni t2 on t0.campione_2=t2.id_campione
left join [TIRELLI_40].[dbo].coll_campioni t3 on t0.campione_3=t3.id_campione
left join [TIRELLI_40].[dbo].coll_campioni t4 on t0.campione_4=t4.id_campione
left join [TIRELLI_40].[dbo].coll_campioni t5 on t0.campione_5=t5.id_campione
left join [TIRELLI_40].[dbo].coll_campioni t6 on t0.campione_6=t6.id_campione
left join [TIRELLI_40].[dbo].coll_campioni t7 on t0.campione_7=t7.id_campione
left join [TIRELLI_40].[dbo].coll_campioni t8 on t0.campione_8=t8.id_campione
left join [TIRELLI_40].[dbo].coll_campioni t9 on t0.campione_9=t9.id_campione
left join [TIRELLI_40].[dbo].coll_campioni t10 on t0.campione_10=t10.id_campione

left join [TIRELLI_40].[dbo].COLL_TIPO_CAMPIONE t11 on t1.TIPO_campione= T11.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[dbo].COLL_TIPO_CAMPIONE t12 on t2.TIPO_campione= T12.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[dbo].COLL_TIPO_CAMPIONE t13 on t3.TIPO_campione= T13.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[dbo].COLL_TIPO_CAMPIONE t14 on t4.TIPO_campione= T14.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[dbo].COLL_TIPO_CAMPIONE t15 on t5.TIPO_campione= T15.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[dbo].COLL_TIPO_CAMPIONE t16 on t6.TIPO_campione= T16.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[dbo].COLL_TIPO_CAMPIONE t17 on t7.TIPO_campione= T17.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[dbo].COLL_TIPO_CAMPIONE t18 on t8.TIPO_campione= T18.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[dbo].COLL_TIPO_CAMPIONE t19 on t9.TIPO_campione= T19.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[dbo].COLL_TIPO_CAMPIONE t20 on t10.TIPO_campione= T20.ID_TIPO_CAMPIONE

where t0.commessa='" & par_commessa & "'
order by

t11.INIZIALE_SIGLA ,  cast(substring(T1.NOME,1,99) as integer) "


        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim colonna As Integer = 1

        Do While cmd_SAP_reader.Read() And colonna < 8

            ' Imposta i valori nelle celle del foglio Excel
            Excel.Sheets(par_nome_foglio).Cells(23, colonna).Value = cmd_SAP_reader("nome_1")
            Excel.Sheets(par_nome_foglio).Cells(24, colonna).Value = cmd_SAP_reader("nome_2")
            Excel.Sheets(par_nome_foglio).Cells(25, colonna).Value = cmd_SAP_reader("nome_3")
            Excel.Sheets(par_nome_foglio).Cells(29, colonna).Value = cmd_SAP_reader("vel_richiesta")
            If cmd_SAP_reader("vel_richiesta") >= VELOCITA Then
                VELOCITA = cmd_SAP_reader("vel_richiesta")
            End If

            Excel.Sheets(par_nome_foglio).Cells(22, colonna).Value = colonna


            colonna += 1
        Loop


        cmd_SAP_reader.Close()
        Cnn.Close()

        If VELOCITA <> 0 Then
            Excel.Sheets(par_nome_foglio).Cells(24, 11).value = VELOCITA
        End If



    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        If TextBox8.Text = Nothing Then
            filtro_brand = ""
        Else
            filtro_brand = "AND COALESCE(t20.brand,'') Like '%%" & TextBox8.Text & "%%' "
        End If
        filtro_commesse()
    End Sub


    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        If e.RowIndex >= 0 Then


            Form_campione_visualizza.id_campione = DataGridView3.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn10").Value
            Form_campione_visualizza.Show()
            Form_campione_visualizza.inizializza_form()

        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        'MsgBox("Funzione non più disponibile")
        'Return
        Form_costificazione.commessa = Codice_commessa
        Form_costificazione.Show()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Commesse_MES.SCHEDA_COMMESSA(Codice_commessa)
        Form_Scheda_Collaudi.Lbl_Commessa.Text = Codice_commessa
        Form_Scheda_Collaudi.inizializzazione_form(Codice_commessa)
        Form_Scheda_Collaudi.Show()


    End Sub

    Private Sub DataGridView12_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView12.CellContentClick

    End Sub

    Private Sub DataGridView12_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView12.CellClick
        If e.RowIndex >= 0 Then

            Codice_commessa = DataGridView12.Rows(e.RowIndex).Cells(columnName:="Macchina").Value
        End If

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Form_costificazione.commessa = Codice_commessa
        Form_costificazione.Show()
    End Sub

    Private Sub DataGridView12_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView12.CellFormatting
        ' Formattazione della colonna Gross_margin
        If DataGridView12.Columns(e.ColumnIndex).Name = "Gross_margin_perc" AndAlso e.Value IsNot Nothing Then
            Dim margin As Double
            If Double.TryParse(e.Value.ToString(), margin) Then
                Select Case margin
                    Case > 0.6
                        e.CellStyle.BackColor = Color.Green
                        e.CellStyle.ForeColor = Color.White
                    Case 0.4 To 0.6
                        e.CellStyle.BackColor = Color.Yellow
                        e.CellStyle.ForeColor = Color.Black
                    Case Else
                        e.CellStyle.BackColor = Color.Red
                        e.CellStyle.ForeColor = Color.White
                End Select
            End If
        End If

        ' Formattazione della colonna PV (Più alto è il valore, più scuro il verde)
        If DataGridView12.Columns(e.ColumnIndex).Name = "PV" AndAlso e.Value IsNot Nothing Then
            Dim pvValue As Double
            If Double.TryParse(e.Value.ToString(), pvValue) Then
                ' Normalizza il valore tra 0 e 255 per il colore verde
                ' Supponiamo che PV vada da 0 a 1000, puoi cambiare il range in base ai tuoi dati
                Dim maxPV As Double = 500000
                Dim minPV As Double = 0
                Dim greenIntensity As Integer = CInt(255 - ((pvValue - minPV) / (maxPV - minPV) * 200))

                ' Limita il valore tra 55 e 255 per evitare colori troppo scuri o chiari
                greenIntensity = Math.Max(55, Math.Min(255, greenIntensity))

                ' Imposta il colore con più verde se il valore è alto
                e.CellStyle.BackColor = Color.FromArgb(0, greenIntensity, 0)
                e.CellStyle.ForeColor = Color.White
            End If
        End If

        If DataGridView12.Columns(e.ColumnIndex).Name = "Gross_margin" AndAlso e.Value IsNot Nothing Then
            Dim pvValue As Double
            If Double.TryParse(e.Value.ToString(), pvValue) Then
                Dim maxPV As Double = 200000
                Dim minPV As Double = 0
                Dim greenIntensity As Integer
                Dim redIntensity As Integer

                If pvValue >= 0 Then
                    ' Scala di verde
                    greenIntensity = CInt(255 - ((pvValue - minPV) / (maxPV - minPV) * 200))
                    greenIntensity = Math.Max(55, Math.Min(255, greenIntensity))
                    e.CellStyle.BackColor = Color.FromArgb(0, greenIntensity, 0)
                Else
                    ' Scala di rosso per valori negativi (da -100000 a 0)
                    Dim minNegPV As Double = -100000
                    redIntensity = CInt(255 - ((pvValue - minNegPV) / (minPV - minNegPV) * 200))
                    redIntensity = Math.Max(55, Math.Min(255, redIntensity))
                    e.CellStyle.BackColor = Color.FromArgb(redIntensity, 0, 0)
                End If

                e.CellStyle.ForeColor = Color.White
            End If
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click


        riempi_datagridview_fatturate(DataGridView12, TextBox20.Text, TextBox19.Text, TextBox24.Text, TextBox23.Text, TextBox25.Text, DateTimePicker4, DateTimePicker1)

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click

        ExportVisibleColumnsToExcel(DataGridView12)
    End Sub

    Public Sub ExportVisibleColumnsToExcel(ByVal par_datagridview As DataGridView)
        Dim excelApp As New Excel.Application
        excelApp.Visible = True

        Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add()
        Dim excelWorksheet As Excel.Worksheet = CType(excelWorkbook.Worksheets(1), Excel.Worksheet)
        Dim rowCount As Integer = par_datagridview.Rows.Count

        Dim visibleColumns = par_datagridview.Columns.Cast(Of DataGridViewColumn).Where(Function(c) c.Visible).ToList()
        Dim visibleColCount As Integer = visibleColumns.Count

        ' Intestazioni
        For colIndex As Integer = 0 To visibleColCount - 1
            excelWorksheet.Cells(1, colIndex + 1) = visibleColumns(colIndex).HeaderText
        Next

        ' Imposta altezza righe e larghezza colonne
        Dim excelRowHeight As Double = 80
        Dim excelColWidth As Double = 20 ' circa 130 pixel
        For i = 1 To visibleColCount
            CType(excelWorksheet.Columns(i), Excel.Range).ColumnWidth = excelColWidth
        Next
        excelWorksheet.Rows("1:" & (rowCount + 1)).RowHeight = excelRowHeight

        ' Inserimento dati
        ' Inserimento dati
        For row As Integer = 0 To rowCount - 1
            For colIndex As Integer = 0 To visibleColCount - 1
                Dim col = visibleColumns(colIndex)
                Dim value = par_datagridview.Rows(row).Cells(col.Index).Value
                Dim cell = CType(excelWorksheet.Cells(row + 2, colIndex + 1), Excel.Range)

                If TypeOf value Is Bitmap Then
                    Try
                        Dim originalImage As Bitmap = CType(value, Bitmap)
                        Dim targetCellWidth As Integer = CInt(cell.Width)
                        Dim targetCellHeight As Integer = CInt(cell.Height)

                        Dim scaleX As Double = targetCellWidth / originalImage.Width
                        Dim scaleY As Double = targetCellHeight / originalImage.Height
                        Dim scale As Double = Math.Min(scaleX, scaleY)

                        Dim newWidth As Integer = CInt(originalImage.Width * scale)
                        Dim newHeight As Integer = CInt(originalImage.Height * scale)

                        Dim resizedImage As New Bitmap(newWidth, newHeight)
                        Using g As Graphics = Graphics.FromImage(resizedImage)
                            g.Clear(Color.White)
                            g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
                            g.DrawImage(originalImage, 0, 0, newWidth, newHeight)
                        End Using

                        Clipboard.SetImage(resizedImage)
                        excelWorksheet.Activate()
                        cell.Select()
                        Application.DoEvents()
                        Threading.Thread.Sleep(100)

                        excelWorksheet.Paste()

                        Dim pastedPicture = excelWorksheet.Pictures(excelWorksheet.Pictures.Count)
                        With pastedPicture
                            .Left = cell.Left + (cell.Width - newWidth) / 2
                            .Top = cell.Top + (cell.Height - newHeight) / 2
                            .Width = newWidth
                            .Height = newHeight
                            .Placement = Excel.XlPlacement.xlMoveAndSize
                        End With
                    Catch ex As Exception
                        MessageBox.Show("Errore durante l'inserimento immagine: " & ex.Message)
                    End Try

                Else
                    If value IsNot Nothing Then
                        ' Se la colonna è "ItemCode" → forza come testo
                        If col.HeaderText.ToLower().Contains("itemcode") Then
                            cell.NumberFormat = "@"
                            cell.Value = "'" & value.ToString()
                        ElseIf IsNumeric(value) Then
                            ' Per i numeri → imposta formato numerico corretto
                            cell.NumberFormat = "0.00" ' o "0" se vuoi senza decimali
                            cell.Value = CDbl(value)
                        Else
                            ' Testo normale
                            cell.NumberFormat = "@"
                            cell.Value = value.ToString()
                        End If
                    Else
                        cell.Value = ""
                    End If
                End If
            Next
        Next

        ' Allineamento celle
        Dim usedRange As Excel.Range = excelWorksheet.Range(
        excelWorksheet.Cells(1, 1),
        excelWorksheet.Cells(rowCount + 1, visibleColCount)
    )
        With usedRange
            .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
            .VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
        End With

        ' Salvataggio opzionale
        Dim saveFileDialog As New SaveFileDialog With {
        .Filter = "Excel Workbook (*.xlsx)|*.xlsx"
    }

        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            excelWorkbook.SaveAs(saveFileDialog.FileName)
            MessageBox.Show("Esportazione completata con successo!", "Successo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        ' Rilascio risorse COM
        ReleaseComObject(excelWorksheet)
        ReleaseComObject(excelWorkbook)
        ReleaseComObject(excelApp)
    End Sub

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub







    Private Sub DataGridView_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView.CellFormatting

        Dim PAR_DATAGRIDVIEW As DataGridView = DataGridView
        Dim nome_Colonna_stato_pogetto As String = "Stato_prog"

        If PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:=nome_Colonna_stato_pogetto).Value = "REV P" Then

            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:=nome_Colonna_stato_pogetto).Style.BackColor = Color.Purple
            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:=nome_Colonna_stato_pogetto).Style.ForeColor = Color.White

        ElseIf PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:=nome_Colonna_stato_pogetto).Value = "REV 0" Then

            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:=nome_Colonna_stato_pogetto).Style.BackColor = Color.Yellow

        ElseIf PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:=nome_Colonna_stato_pogetto).Value = "REV A" Then

            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:=nome_Colonna_stato_pogetto).Style.BackColor = Color.Lime

        ElseIf PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:=nome_Colonna_stato_pogetto).Value = "SOSPESO" Then

            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:=nome_Colonna_stato_pogetto).Style.BackColor = Color.Aqua

        ElseIf PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:=nome_Colonna_stato_pogetto).Value = "CHIUSO" Then

            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:=nome_Colonna_stato_pogetto).Style.BackColor = Color.Gray
        Else

            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:=nome_Colonna_stato_pogetto).Style.BackColor = Nothing
        End If



        ' Controlliamo se la colonna è quella giusta
        If DataGridView.Columns(e.ColumnIndex).Name = "Brand" AndAlso e.Value IsNot Nothing Then
            Select Case e.Value.ToString()
                Case "BRB"
                    e.CellStyle.BackColor = Color.Yellow
                Case "KTF"
                    e.CellStyle.BackColor = Color.Green
                Case "GHERRI"
                    e.CellStyle.BackColor = Color.Orange
                Case "TIRELLI"
                    e.CellStyle.BackColor = Color.DarkBlue
                    e.CellStyle.ForeColor = Color.White ' Per migliorare la visibilità del testo
            End Select
        End If
    End Sub



    Private Sub CancellaPuntoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CancellaPuntoToolStripMenuItem.Click
        'If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Trasferito").Value > 0 Then
        '    MsgBox("Impossibile cancellare riga di un codice che risulta TRASFERITO")
        'Else
        Dim PAR_DATAGRIDVIEW As DataGridView
        PAR_DATAGRIDVIEW = DataGridView13

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

    Private Sub CompletatoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CompletatoToolStripMenuItem.Click
        'If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Trasferito").Value > 0 Then
        '    MsgBox("Impossibile cancellare riga di un codice che risulta TRASFERITO")
        'Else
        Dim PAR_DATAGRIDVIEW As DataGridView
        PAR_DATAGRIDVIEW = DataGridView13

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

    Sub cambia_stato(par_ID As Integer)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand
        CMD_SAP_5.Connection = CNN6
        CMD_SAP_5.CommandText = "UPDATE [TIRELLI_40].[DBO].dati_mancanti_progetto
SET  RISOLTO=CASE WHEN RISOLTO='FALSE' THEN 'TRUE' ELSE 'FALSE' END

      
WHERE ID=@ID"

        ' Aggiunta dei parametri

        CMD_SAP_5.Parameters.AddWithValue("@ID", par_ID)



        CMD_SAP_5.ExecuteNonQuery()

        CNN6.Close()
    End Sub

    Private Sub DataGridView13_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView13.CellContentClick


    End Sub

    Private Sub DataGridView13_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView13.CellFormatting
        Dim par_datagridview As DataGridView = DataGridView13
        ' Verifica se la colonna "stato" è presente (sostituisci "stato" con il nome corretto della colonna)
        Dim statoIndex As Integer = par_datagridview.Columns("stato__").Index

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



        If par_datagridview.Rows(e.RowIndex).Cells(columnName:="Nome_Stato").Value = "REV P" Then

            par_datagridview.Rows(e.RowIndex).Cells(columnName:="Nome_Stato").Style.BackColor = Color.Purple
            par_datagridview.Rows(e.RowIndex).Cells(columnName:="Nome_Stato").Style.ForeColor = Color.White

        ElseIf par_datagridview.Rows(e.RowIndex).Cells(columnName:="Nome_Stato").Value = "REV 0" Then

            par_datagridview.Rows(e.RowIndex).Cells(columnName:="Nome_Stato").Style.BackColor = Color.Yellow

        ElseIf par_datagridview.Rows(e.RowIndex).Cells(columnName:="Nome_Stato").Value = "REV A" Then

            par_datagridview.Rows(e.RowIndex).Cells(columnName:="Nome_Stato").Style.BackColor = Color.Lime

        ElseIf par_datagridview.Rows(e.RowIndex).Cells(columnName:="Nome_Stato").Value = "SOSPESO" Then

            par_datagridview.Rows(e.RowIndex).Cells(columnName:="Nome_Stato").Style.BackColor = Color.Aqua

        ElseIf par_datagridview.Rows(e.RowIndex).Cells(columnName:="Nome_Stato").Value = "CHIUSO" Then

            par_datagridview.Rows(e.RowIndex).Cells(columnName:="Nome_Stato").Style.BackColor = Color.Gray
        Else

            par_datagridview.Rows(e.RowIndex).Cells(columnName:="Nome_Stato").Style.BackColor = Nothing
        End If
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        modifica_appunti_globali(DataGridView13)
        MsgBox("Aggiornato con successo")
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

    Sub aggiorna_appunti(par_ID As Integer, par_commento As String)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand
        CMD_SAP_5.Connection = CNN6
        CMD_SAP_5.CommandText = "UPDATE [TIRELLI_40].[DBO].dati_mancanti_progetto

       SET [Contenuto]=@Contenuto
WHERE ID=@ID"

        ' Aggiunta dei parametri
        CMD_SAP_5.Parameters.AddWithValue("@Contenuto", par_commento)
        CMD_SAP_5.Parameters.AddWithValue("@ID", par_ID)

        CMD_SAP_5.ExecuteNonQuery()

        CNN6.Close()
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        appunti_globali(Homepage.ID_SALVATO)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Progetto.mail_dati_mancanti("")
    End Sub

    Private Sub DataGridView13_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView13.CellClick
        If e.RowIndex >= 0 Then



            If e.ColumnIndex = DataGridView13.Columns.IndexOf(Num_prog) Then
                Progetto.Show()
                Progetto.BringToFront()
                Progetto.absentry = DataGridView13.Rows(e.RowIndex).Cells(columnName:="Num_prog").Value
                Progetto.inizializza_progetto()


            End If
        End If
    End Sub

    Private Sub DataGridView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView.CellContentClick

    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        ExportVisibleColumnsToExcel(DataGridView)
    End Sub
End Class