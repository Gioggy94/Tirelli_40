Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Linq.Expressions

Imports System.Threading.Tasks
Imports Npgsql.Internal.TypeHandlers.NetworkHandlers

Public Class Inventario
    Public riga As Integer
    Public Codice_SAP As String
    Public Descrizione As String
    Public Descrizione_SUP As String
    Public osservazioni As String
    Public disegno As String
    Public Diff As String
    Public inventariato As String

    Public filtro_order_by As String
    Private filtro_ubicazione As String
    Private id_inv As Integer
    Public blocco_denominatore As Boolean = False
    Private Denominatore As Integer
    Private n_numeratore As Object
    Public Elenco_dipendenti(1000) As String

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button_inserisci_Click(sender As Object, e As EventArgs) Handles Button_inserisci.Click
        If ComboBox_DIPENDENTE.SelectedIndex < 0 Then
            MsgBox("Selezionare Un dipendente")
            Return
        End If

        If TextBox17.Text = "" Then
            MsgBox("Indicare il magazzino")
            Return
        End If

        If TextBox3.Text = Nothing Then
            MsgBox("codice scorretto")
            Return
        End If



        If Magazzino.OttieniDettagliAnagrafica(TextBox2.Text).Descrizione = "" Then
            MsgBox("Codice non esistente")
            Return
        End If



        inserisci_conteggio_inventario(TextBox6.Text, Elenco_dipendenti(ComboBox_DIPENDENTE.SelectedIndex), TextBox2.Text, TextBox17.Text, TextBox1.Text, TextBox3.Text, "SI")
        inserimento_datagridview_conteggio_inventario()


        inventario_AUTOMATICO_new(TextBox17.Text, TextBox1.Text, TextBox6.Text, TextBox14.Text)



    End Sub
    Sub inserisci_conteggio_inventario(par_codice_inventario As String, par_dipendente As String, par_itemcode As String, par_magazzino As String, par_ubicazione As String, par_quantity As String, par_inventariato As String)
        par_quantity = Replace(par_quantity, ",", ".")

        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()
        Dim CMD_SAP_3 As New SqlCommand
        CMD_SAP_3.Connection = Cnn3
        CMD_SAP_3.CommandText = "INSERT INTO [dbo].[INVENTARIO]
           ([Codice_inventario]
           ,[Dipendente]
           ,[Itemcode]
           ,[Magazzino]
           ,[ubicazione]
           ,[quantity]
           ,[data]
           ,[inventariato])
     VALUES
           ('" & par_codice_inventario & "'
           ,'" & par_dipendente & "'
           ,'" & par_itemcode & "'
           ,'" & par_magazzino & "'
           ,'" & par_ubicazione & "'
           ,'" & par_quantity & "'
           ,getdate()
           ,'" & par_inventariato & "')"
        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()

    End Sub

    Sub calcola_database(par_data_inventario As String, par_percentuale As String)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()
        Dim CMD_SAP_3 As New SqlCommand
        CMD_SAP_3.Connection = Cnn3
        CMD_SAP_3.CommandText = "
DECLARE @data_inventario DATE;
DECLARE @listino INT;

SET @listino = 2;
SET @data_inventario = '" & par_data_inventario & "'
delete [Tirelli_40].[dbo].inventario_tempo
INSERT INTO [Tirelli_40].[dbo].INVENTARIO_tempo (itemcode, WhsCode,  Giacenza_oggi, Movimentazioni, Giacenza_alla_data, Price, Valore_alla_data)
SELECT TOP " & par_percentuale & " PERCENT t10.itemcode, t10.WhsCode, 
       SUM(t10.Giacenza_oggi) AS 'Giacenza_oggi', 
       SUM(t10.Movimentazioni) AS 'Movimentazioni', 
       SUM(t10.Giacenza_oggi) - SUM(t10.Movimentazioni) AS 'Giacenza_alla_data', 
       t13.Price, 
       (SUM(t10.Giacenza_oggi) - SUM(t10.Movimentazioni)) * t13.Price AS 'Valore_alla_data'
FROM (
    SELECT t0.itemcode, t0.WhsCode, 
           CASE WHEN t0.OnHand IS NULL THEN 0 ELSE t0.OnHand END AS 'Giacenza_oggi', 
           0 AS 'Movimentazioni'
    FROM OITW T0 
    INNER JOIN OWHS T1 ON T0.WHSCODE = T1.WhsCode
    WHERE ((T0.[ItemCode] LIKE N'[C]%' ) OR 
           (T0.[ItemCode] LIKE N'[D]%' ) OR 
           (T0.[ItemCode] LIKE N'[M]%' ) OR 
           (T0.[ItemCode] LIKE N'[0]%' )) 
      AND t0.whscode <> '03' 
      AND t0.whscode <> 'B03' 
 AND t0.whscode <> 'WIP' 
      AND t0.whscode <> 'BWIP' 
      AND t0.OnHand <> 0
    UNION ALL
    SELECT t5.itemcode, t5.loccode, t5.Giacenza_oggi, t5.Movimentazioni
    FROM (
        SELECT t0.itemcode, t0.loccode, 0 AS 'Giacenza_oggi', 
               SUM(CASE WHEN t0.InQty IS NULL THEN 0 ELSE t0.InQty END - 
                   CASE WHEN t0.OutQty IS NULL THEN 0 ELSE t0.OutQty END) AS 'Movimentazioni'
        FROM oivl t0
        LEFT JOIN [OILM] T1 ON T1.[MessageID] = T0.[MessageID]
        WHERE t1.TaxDate > CONVERT(DATETIME, @data_inventario, 112) 
          AND ((T0.[ItemCode] LIKE N'[C]%' ) OR 
               (T0.[ItemCode] LIKE N'[D]%' ) OR 
               (T0.[ItemCode] LIKE N'[M]%' ) OR 
               (T0.[ItemCode] LIKE N'[0]%' )) 
          AND t0.loccode <> '03' 
          AND t0.loccode <> 'B03'
        GROUP BY t0.itemcode, t0.loccode
    ) AS t5 
    WHERE t5.Movimentazioni <> 0
) AS t10
INNER JOIN oitm t11 ON t11.itemcode = t10.itemcode
INNER JOIN OWHS T12 ON T10.WhsCode = T12.WhsCode
INNER JOIN itm1 t13 ON t13.itemcode = t10.ItemCode
WHERE t13.PriceList = @listino
GROUP BY t10.itemcode, t10.WhsCode, t11.itemname, T12.WHSNAME, t13.Price
HAVING SUM(t10.Giacenza_oggi) - SUM(t10.Movimentazioni) <> 0
ORDER BY (SUM(t10.Giacenza_oggi) - SUM(t10.Movimentazioni)) * t13.Price DESC "
        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()

    End Sub

    Sub inserisci_ubicazione(par_codice_dipendente As String)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()
        Dim CMD_SAP_3 As New SqlCommand
        CMD_SAP_3.Connection = Cnn3
        CMD_SAP_3.CommandText = ""
        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()

    End Sub

    Sub salta(par_codice_dipendente As String)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()
        Dim CMD_SAP_3 As New SqlCommand
        CMD_SAP_3.Connection = Cnn3
        CMD_SAP_3.CommandText = ""
        CMD_SAP_3.ExecuteNonQuery()
        Cnn3.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_DIPENDENTE.SelectedIndexChanged
        ' Dashboard_pianificazione.Dipendente = Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_DIPENDENTE.SelectedIndex)
    End Sub






    Public Sub ShowTemporaryMessage(message As String)
        Dim tempForm As New FormMessage() ' Crea una nuova istanza del form personalizzato
        tempForm.SetMessage(message) ' Imposta il messaggio da mostrare
        tempForm.Show() ' Mostra il form
        ' Imposta un timer per chiudere automaticamente il form dopo 2 secondi
        Dim t As New Timer()
        t.Interval = TextBox15.Text
        AddHandler t.Tick, Sub(sender, e)
                               tempForm.Close()
                               t.Stop()
                           End Sub
        t.Start()
    End Sub

    Sub numeratore(par_magazzino As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "select count(T0.itemcode) as 'N'
from [Tirelli_40].[dbo].[INVENTARIO] T0 INNER JOIN [Tirelli_40].[dbo].Inventario_tempo T1 ON T0.ITEMCODE=T1.ITEMCODE AND T0.magazzino=t1.whscode
where T0.magazzino ='" & par_magazzino & "'"


        cmd_SAP_reader = CMD_SAP.ExecuteReader



        If cmd_SAP_reader.Read() Then
            n_numeratore = cmd_SAP_reader("N")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub


    Sub denominatore_function(par_magazzino As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "select count(itemcode) as 'N'
from [Tirelli_40].[dbo].inventario_tempo t0 
where t0.whscode ='" & par_magazzino & "'"


        cmd_SAP_reader = CMD_SAP.ExecuteReader



        If cmd_SAP_reader.Read() Then
            Denominatore = cmd_SAP_reader("N")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub inserimento_datagridview_conteggio_inventario()
        DataGridView1.Rows.Clear()

        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Try
                Cnn.Open()

                ' Nota: Usiamo LEFT JOIN con OHEM per avere Nome e Cognome del dipendente
                Dim sql As String = "SELECT " &
                "t0.id, " &
                "COALESCE(t1.lastName, '') + ' ' + COALESCE(t1.firstName, '') AS Dipendente, " &
                "t0.Itemcode, " &
                "t0.quantity, " &
                "t0.data, " &
                "t0.ubicazione AS Cassetto, " & ' Mappato su ubicazione
                "t0.Magazzino, " &
                "t0.inventariato " &
                "FROM [Tirelli_40].[dbo].[INVENTARIO] t0 " &
                "LEFT JOIN [Tirelli_40].[dbo].[OHEM] t1 ON t1.empID = t0.Dipendente " &
                "WHERE t0.Itemcode LIKE @ricerca " &
                "ORDER BY id DESC"

                Using CMD_SAP As New SqlCommand(sql, Cnn)
                    ' Il parametro previene errori di sintassi nella ricerca
                    CMD_SAP.Parameters.AddWithValue("@ricerca", "%" & TextBox4.Text & "%")

                    Using cmd_SAP_reader As SqlDataReader = CMD_SAP.ExecuteReader()
                        Do While cmd_SAP_reader.Read()
                            ' Assicurati che l'ordine delle colonne nella DataGridView 
                            ' corrisponda a questo elenco:
                            DataGridView1.Rows.Add(
                            cmd_SAP_reader("id"),
                            cmd_SAP_reader("Itemcode"),
                            cmd_SAP_reader("quantity"),
                            cmd_SAP_reader("Dipendente"),
                            cmd_SAP_reader("data"),
                            cmd_SAP_reader("Cassetto"),
                            cmd_SAP_reader("Magazzino"),
                            cmd_SAP_reader("inventariato")
                        )
                        Loop
                    End Using
                End Using

            Catch ex As Exception
                MessageBox.Show("Errore durante il caricamento dell'inventario: " & ex.Message)
            End Try
        End Using
    End Sub

    Sub inserimento_datagridview_ubicazione()
        DataGridView3.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = cnn
        If TextBox10.Text = Nothing Then
            CMD_SAP.CommandText = "Select  T0.ID AS 'ID', t1.lastname+' '+t1.firstname as 'Dipendente', t0.itemcode, t0.quantity,t0.data,t0.ora, t0.cassetto, t0.magazzino,t0.anno
from [Tirelli_40].[dbo].[INVENTARIO] t0 left join [TIRELLI_40].[dbo].ohem t1 on t1.code = t0.dipendente
where t0.itemcode like '%%" & TextBox9.Text & "%%' 
order by t0.data DESC, t0.ora DESC"
        Else
            CMD_SAP.CommandText = "Select T0.ID AS 'ID', t1.lastname+' '+t1.firstname as 'Dipendente', t0.itemcode, t0.quantity,t0.data,t0.ora, t0.cassetto, t0.magazzino,t0.anno
from [Tirelli_40].[dbo].[INVENTARIO] t0 left join [TIRELLI_40].[dbo].ohem t1 on t1.code = t0.dipendente
where t0.itemcode like '%%" & TextBox9.Text & "%%' and t0.cassetto = '" & TextBox10.Text & "'
order by t0.data DESC, t0.ora DESC"

        End If
        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()
            DataGridView3.Rows.Add(cmd_SAP_reader("ID"), cmd_SAP_reader("itemcode"), cmd_SAP_reader("quantity"), cmd_SAP_reader("dipendente"), cmd_SAP_reader("data"), cmd_SAP_reader("ora"), cmd_SAP_reader("Cassetto"), cmd_SAP_reader("magazzino"), cmd_SAP_reader("anno"))
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()
    End Sub

    Sub inserimento_dipendenti_MES()


        ComboBox_DIPENDENTE.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "select *
from
(
select '' as 'Codice dipendenti', '' as 'Nome', '' as 'Nome 2'
union all
SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code where t0.active='Y'
)
as t0
order by t0.nome"
        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Lavorazioni_MES.Elenco_dipendenti_MES(Indice) = cmd_SAP_reader("Codice dipendenti")
            Lavorazioni_MES.ComboBox_dipendente.Items.Add(cmd_SAP_reader("Nome"))
            ComboBox_DIPENDENTE.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        inserimento_datagridview_conteggio_inventario()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        inserimento_datagridview_conteggio_inventario()
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "select t0.itemname, case when t0.frgnname is null then '' else t0.frgnname end as 'frgnname' , case when t0.u_disegno is null then '' else t0.U_disegno end as 'U_disegno', case when t0.validcomm is null then '' else t0.validcomm end as 'validcomm'  from oitm t0 where t0.itemcode= '" & TextBox2.Text & "'"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then
            Label1.Text = cmd_SAP_reader("itemname")
            Label2.Text = cmd_SAP_reader("frgnname")
            Label3.Text = cmd_SAP_reader("validcomm")
            Button2.Text = cmd_SAP_reader("u_disegno")

        End If
        cmd_SAP_reader.Close()
        cnn.Close()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Codice_SAP = TextBox2.Text

        'Try
        dettagli_anagrafica()


        Label1.Text = Descrizione
        Label2.Text = Descrizione_SUP
        Label3.Text = osservazioni

    End Sub

    Sub dettagli_anagrafica()
        ' Usiamo "Using" per garantire la chiusura della connessione anche in caso di errori
        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Try
                Cnn.Open()

                Using CMD_SAP As New SqlCommand()
                    CMD_SAP.Connection = Cnn

                    ' 1. Determiniamo la Query in base alla provenienza (come nella sub start_magazzino)
                    If Homepage.ERP_provenienza = "SAP" Then
                        CMD_SAP.CommandText = "SELECT " &
                        "COALESCE(t0.itemcode, '') AS itemcode, " &
                        "COALESCE(t0.itemname, '') AS itemname, " &
                        "COALESCE(t0.frgnname, '') AS frgnname, " &
                        "COALESCE(t0.u_disegno, '') AS u_disegno, " &
                        "COALESCE(t0.validcomm, '') AS validcomm, " &
                        "COALESCE(t0.u_ubicazione, '') AS u_ubicazione, " &
                        "COALESCE(t0.u_codice_brb, '') AS Codice_brb " &
                        "FROM oitm t0 WHERE t0.itemcode = @codice"
                        CMD_SAP.Parameters.AddWithValue("@codice", Codice_SAP)
                    Else
                        ' Caso AS400 (adattato per restituire gli stessi nomi colonna di SAP per compatibilità)
                        CMD_SAP.CommandText = String.Format("SELECT " &
                        "trim(CODE) AS itemcode, " &
                        "DES_CODE AS itemname, " &
                        "'' AS frgnname, " &
                        "'' AS u_disegno, " &
                        "'' AS validcomm, " &
                        "'' AS u_ubicazione, " &
                        "'' AS Codice_brb " &
                        "FROM OPENQUERY(AS400, 'SELECT * FROM S786FAD1.TIR90VIS.JGALART WHERE code = ''{0}''') T10",
                        Codice_SAP.Replace("'", "''"))
                    End If

                    ' 2. Esecuzione e lettura
                    Using cmd_SAP_reader As SqlDataReader = CMD_SAP.ExecuteReader()
                        If cmd_SAP_reader.Read() Then
                            ' Assegnazione ai controlli e alle variabili
                            Button11.Text = cmd_SAP_reader("itemcode").ToString()
                            Descrizione = cmd_SAP_reader("itemname").ToString()
                            Descrizione_SUP = cmd_SAP_reader("frgnname").ToString()
                            osservazioni = cmd_SAP_reader("validcomm").ToString()
                            disegno = cmd_SAP_reader("u_disegno").ToString()

                            TextBox1.Text = cmd_SAP_reader("u_ubicazione").ToString()
                            TextBox12.Text = cmd_SAP_reader("Codice_brb").ToString()
                            Button2.Text = cmd_SAP_reader("u_disegno").ToString()

                            ' Aggiornamento Grid e Feedback
                            Magazzino.giacenze_magazzino(DataGridView_magazzino, Codice_SAP)
                            Beep()
                        End If
                    End Using
                End Using

            Catch ex As Exception
                MessageBox.Show("Errore nel caricamento dettagli: " & ex.Message)
            End Try
        End Using
    End Sub

    Sub Inserimento_dipendenti(par_combobox As ComboBox)


        par_combobox.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim Indice As Integer
        Indice = 0



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 
left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
inner join [TIRELLI_40].[DBO].COLL_Reparti t2 on (t2.id_reparto =t0.u_reparto_tickets)   
where t0.active='Y' 
order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            par_combobox.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

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
            pdfPath &= "#toolbar=0&zoom=50&navpanes=0"

            ' Esegui l'operazione di navigazione sul thread dell'UI
            Await Task.Run(Sub()
                               par_web_browser.Invoke(Sub()
                                                          par_web_browser.Show()
                                                          par_web_browser.Navigate(pdfPath)
                                                      End Sub)
                           End Sub)
        Else
            ' Nascondi il WebBrowser sul thread dell'UI
            par_web_browser.Invoke(Sub() par_web_browser.Hide())
        End If
    End Function


    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        inserimento_datagridview_conteggio_inventario()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If File.Exists(Homepage.percorso_disegni_generico & "PDF\" & Button2.Text & ".PDF") Then
            Process.Start(Homepage.percorso_disegni_generico & "PDF\" & Button2.Text & ".PDF")
        Else
            MsgBox("PDF non presente")
        End If



        ''Verifica se il file esiste e poi visualizzalo nel WebBrowser
        ''If File.Exists(Homepage.percorso_disegni_generico & "PDF\" & Button2.Text & ".PDF") Then
        ''    WebBrowser1.Navigate(Homepage.percorso_disegni_generico & "PDF\" & Button2.Text & ".PDF")
        ''Else
        ''    MsgBox("PDF non presente")
        ''End If
    End Sub

    Sub trova_ID()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "SELECT 'o', max(case when t0.id is null then 0 else t0.id end )+1 as 'ID' from [Tirelli_40].[dbo].[INVENTARIO] t0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then

            If Not cmd_SAP_reader("ID") Is System.DBNull.Value Then
                id_inv = cmd_SAP_reader("ID")
            Else
                id_inv = 1
            End If
        Else
            id_inv = 1
        End If

        cnn.Close()
        cmd_SAP_reader.Close()


    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            riga = e.RowIndex
        End If
    End Sub


    Sub elimina_record_conteggio()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand

        CMD_SAP_7.Connection = Cnn

        CMD_SAP_7.CommandText = "DELETE [Tirelli_40].[dbo].[INVENTARIO] WHERE ID='" & DataGridView1.Rows(riga).Cells(0).Value & "'"

        CMD_SAP_7.ExecuteNonQuery()

        cnn.Close()
    End Sub

    Sub elimina_record_ubicazione()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand

        CMD_SAP_7.Connection = Cnn

        CMD_SAP_7.CommandText = "DELETE [Tirelli_40].[dbo].[INVENTARIO] WHERE ID='" & DataGridView3.Rows(riga).Cells(0).Value & "'"

        CMD_SAP_7.ExecuteNonQuery()

        cnn.Close()
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        If DataGridView3.Rows(riga).Cells(0).Value >= 1 Then
            elimina_record_ubicazione()
            inserimento_datagridview_ubicazione()
        End If
    End Sub

    'Sub magazzini()
    '    Dim Cnn2 As New SqlConnection
    '    ComboBox1.Items.Clear()
    '    cnn2.ConnectionString = Homepage.sap_tirelli
    '    cnn2.Open()

    '    Dim CMD_SAP_2 As New SqlCommand
    '    Dim cmd_SAP_reader_2 As SqlDataReader

    '    CMD_SAP_2.Connection = cnn2
    '    CMD_SAP_2.CommandText = "SELECT t0.whscode from OWHS t0"

    '    cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


    '    Do While cmd_SAP_reader_2.Read()

    '        ComboBox1.Items.Add(cmd_SAP_reader_2("WHSCODE"))

    '    Loop
    '    cmd_SAP_reader_2.Close()
    '    cnn2.Close()


    'End Sub



    Sub inventario_AUTOMATICO_new(PAR_MAGAZZINO As String, Par_ubicazione As String, par_codice_inventario As String, par_percentuale As String)

        Dim Cnn4 As New SqlConnection
        Cnn4.ConnectionString = Homepage.sap_tirelli

        Cnn4.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn4

        CMD_SAP.CommandText = "SELECT TOP 1 t0.[Codice_inventario]
      ,t0.[Itemcode]
      ,t0.[Mag]
      ,t0.[ubicazione]
      ,t0.[quantity]
      ,t0.[Valore]
      ,t0.[data]
  FROM [Tirelli_40].[dbo].[Da_inventariare] t0
  left join [Tirelli_40].[dbo].[INVENTARIO] t1 on t0.itemcode=t1.itemcode 
  where t0.mag<>'T03' and t0.mag Like '%%" & PAR_MAGAZZINO & "%%' 
  and t0.ubicazione Like '%%" & Par_ubicazione & "%%'
and t0.[Codice_inventario]Like '%%" & par_codice_inventario & "%%' and t1.itemcode is null
  order by t0.valore desc
"


        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            TextBox2.Text = cmd_SAP_reader("itemcode")
            Label4.Text = "Giacenza alla data nel mag = " & cmd_SAP_reader("quantity")
            '  Label5.Text = "Movimentazioni da inizio inventario = " & cmd_SAP_reader("Movimentazioni")
            Label9.Text = cmd_SAP_reader("Valore")
            TextBox12.Text = cmd_SAP_reader("ubicazione")
            Button2.Text = Magazzino.OttieniDettagliAnagrafica(cmd_SAP_reader("itemcode")).Disegno
            Magazzino.visualizza_picture(Button2.Text, PictureBox2)
        Else
            TextBox2.Text = ""
            Label4.Text = ""
            Label5.Text = ""
            Label9.Text = ""
            MsgBox("Codici finiti con questi criteri")
        End If

        cmd_SAP_reader.Close()
        Cnn4.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If ComboBox_DIPENDENTE.SelectedIndex < 0 Then
            MsgBox("Selezionare Un dipendente")
            Return
        End If

        If TextBox17.Text = "" Then
            MsgBox("Indicare il magazzino")
            Return
        End If

        If TextBox3.Text = Nothing Then
            MsgBox("codice scorretto")
            Return
        End If



        If Magazzino.OttieniDettagliAnagrafica(TextBox2.Text).Descrizione = "" Then
            MsgBox("Codice non esistente")
            Return
        End If



        inserisci_conteggio_inventario(TextBox6.Text, Elenco_dipendenti(ComboBox_DIPENDENTE.SelectedIndex), TextBox2.Text, TextBox17.Text, TextBox1.Text, TextBox3.Text, "NO")

        inserimento_datagridview_conteggio_inventario()




    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then

            inventario_AUTOMATICO_new(TextBox17.Text, TextBox1.Text, TextBox6.Text, TextBox14.Text)


        End If

    End Sub


    Sub giacenze_magazzino()
        Dim Cnn1 As New SqlConnection
        DataGridView_magazzino.Rows.Clear()
        cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "SELECT T0.[WhsCode], CASE WHEN T0.[OnHand] is null then 0 else T0.[OnHand] END AS 'onhand' , case when T0.[IsCommited] is null then 0 else T0.[IsCommited] end as 'iscommited' , case when T0.[OnOrder] is null then 0 else T0.[OnOrder] end as 'onorder'  FROM OITW T0 WHERE (T0.[OnHand]>0 or t0.iscommited>0 or t0.onorder>0) and t0.itemcode='" & Codice_SAP & "'"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()


            DataGridView_magazzino.Rows.Add(cmd_SAP_reader_2("whscode"), cmd_SAP_reader_2("onhand"), cmd_SAP_reader_2("iscommited"), cmd_SAP_reader_2("onorder"))
        Loop


        cnn1.Close()


    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        Magazzino.Codice_SAP = TextBox8.Text
        Magazzino.OttieniDettagliAnagrafica(TextBox8.Text)
        Label6.Text = Magazzino.OttieniDettagliAnagrafica(TextBox8.Text).Descrizione
        Label7.Text = Magazzino.OttieniDettagliAnagrafica(TextBox8.Text).Descrizione_SUP
        Label8.Text = Magazzino.OttieniDettagliAnagrafica(TextBox8.Text).Osservazioni
        Button8.Text = Magazzino.OttieniDettagliAnagrafica(TextBox8.Text).Disegno
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Magazzino.OttieniDettagliAnagrafica(TextBox8.Text).Test = "NO" Then
            MsgBox("Codice articolo non esistente")

        Else
            trova_ID()
            inserisci_ubicazione(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_DIPENDENTE.SelectedIndex))
            inserimento_datagridview_ubicazione()
            TextBox8.Text = ""
        End If
    End Sub



    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick
        If e.RowIndex >= 0 Then
            riga = e.RowIndex
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If riga >= 0 Then
            If DataGridView1.Rows(riga).Cells(0).Value >= 1 Then
                elimina_record_conteggio()
                inserimento_datagridview_conteggio_inventario()
            End If
        Else
            MsgBox("Selezionare codice da eliminare")
        End If



        riga = -1
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        inserimento_datagridview_ubicazione()

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        inserimento_datagridview_ubicazione()
    End Sub



    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Me.Close()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            filtro_order_by = " ORDER BY t0.[Valore] DESC "


        Else


            filtro_order_by = " ORDER BY t0.[u_ubicazione] "

        End If
        inventario_AUTOMATICO_new(TextBox17.Text, TextBox1.Text, TextBox6.Text, TextBox14.Text)

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.Text = "" Then
            filtro_ubicazione = ""
        Else

            If Homepage.Centro_di_costo = "BRB01" Then
                filtro_ubicazione = " AND t1.u_ubicazione_labelling LIKE '" & TextBox7.Text & "%' "
            Else

                filtro_ubicazione = " AND t1.u_ubicazione LIKE '" & TextBox7.Text & "%' "
            End If


        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Inventario_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Inserimento_dipendenti(ComboBox_DIPENDENTE)
        If Homepage.Centro_di_costo = "BRB01" Then
            filtro_order_by = "ORDER BY t1.[u_ubicazione_labelling]"
        Else

            filtro_order_by = "ORDER BY t1.[u_ubicazione]"
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Magazzino.Codice_SAP = Button11.Text

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

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        inventario_AUTOMATICO_new(TextBox17.Text, TextBox1.Text, TextBox6.Text, TextBox14.Text)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs)
        denominatore_function(TextBox17.Text)
        blocco_denominatore = False
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        'calcola_database(TextBox11.Text, TextBox14.Text)
        'MsgBox("Database calcolato con successo")
        If TextBox16.Text = "" Then
            MsgBox("Dare un nome all'inventario che si sta importanto")
            Return
        End If
        trova_dato_da_excel_pEr_importazionE("C:\Users\giovannitirelli\Desktop\Inventario.xlsx", "Foglio1", 2, TextBox16.Text)
        MsgBox("Inventario importato con successo")
    End Sub


    Sub trova_dato_da_excel_pEr_importazionE(par_percorso_file As String, par_nome_foglio As String, par_riga_inizio As Integer, par_nome_inventario As String)

        Dim colonna1 As String
        Dim colonna2 As String
        Dim colonna3 As String
        Dim colonna4 As String
        Dim colonna5 As String
        Dim colonna6 As String

        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True

        delete_codici(par_nome_inventario)
        Do While Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value <> ""


            If Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value <> Nothing Then
                colonna1 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value
                colonna2 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value
                colonna3 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 3).value
                colonna4 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value
                colonna5 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 5).value
                colonna6 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 6).value

                ' crea_codice_articolo(colonna1, colonna2, colonna3, colonna4, colonna5, colonna6)
                insert_into_codici(colonna1, colonna2, colonna3, colonna4, colonna5, colonna6, par_nome_inventario)
                'Update(colonna1, colonna2, colonna3, colonna4, colonna5)
            End If
            par_riga_inizio = par_riga_inizio + 1
        Loop
        Beep()
        MsgBox("Importazione effettuata con successo")


    End Sub

    Sub delete_codici(par_nome_inventario As String)




        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1
        Cmd_SAP.CommandText = "
delete [Tirelli_40].[dbo].[Da_inventariare] where [Codice_inventario]='" & par_nome_inventario & "'"
        Cmd_SAP.ExecuteNonQuery()


        Cnn1.Close()

    End Sub
    Sub insert_into_codici(par_Colonna_1 As String, par_Colonna_2 As String, par_Colonna_3 As String, par_Colonna_4 As String, par_colonna_5 As String, par_colonna_6 As String, par_nome_inventario As String)

        par_Colonna_1 = Replace(par_Colonna_1, "'", " ")
        par_Colonna_2 = Replace(par_Colonna_2, "'", " ")
        par_Colonna_3 = Replace(par_Colonna_3, "'", " ")
        par_Colonna_4 = Replace(par_Colonna_4, "'", " ")
        par_colonna_5 = Replace(par_colonna_5, ",", ".")
        par_colonna_6 = Replace(par_colonna_6, ",", ".")


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1


        Cmd_SAP.CommandText = "
INSERT INTO [Tirelli_40].[dbo].[Da_inventariare]
           ([Codice_inventario]
           ,[Itemcode]
           ,[Mag]
           ,[ubicazione]
           ,[quantity]
           ,[Valore]
           ,[data])
     VALUES
           ('" & par_nome_inventario & "'
,'" & par_Colonna_1 & "'
           ,'" & par_Colonna_2 & "'
           ,'" & par_Colonna_3 & "'
,'" & par_Colonna_4 & "'
,'" & par_colonna_5 & "'
,getdate())
          "

        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()

    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged

    End Sub
End Class