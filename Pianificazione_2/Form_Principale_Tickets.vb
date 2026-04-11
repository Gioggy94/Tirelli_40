Imports System.IO
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop


Public Class Pianificazione_Tickets

    Public Administrator As Integer

    Public Elenco_Reparti(1000) As String
    Public Reparti_Caricati As Integer
    ' Public Mittente_Mail As String
    Public Password_Mail As String
    Public Ultima_Riga As String
    'Public Business As String
    Public CODICE_REPARTO As Integer
    Public filtro_reparto As String
    Public filtro_reparto_task As String
    Public filtro_commessa As String
    Public filtro_id As String
    Public filtro_id_padre As String
    Public filtro_cliente As String
    Public filtro_mittente_padre As String
    Public filtro_business As String
    Public filtro_utente_padre As String
    Public filtro_utente As String
    Public filtro_riunione As String
    Public filtro_articolo As String
    Public filtro_assegnato As String
    Public riga As Integer
    Public status_1 As String = "t0.aperto=1"
    Public status_2 As String = "t10.aperto=1"
    Public Declare Ansi Function ExtractIconEx Lib "Shell32.dll" _
    (ByVal lpszFile As String,
    ByVal nIconIndex As Integer, ByVal phIconLarge As IntPtr(),
    ByVal phIconSmall As IntPtr(), ByVal nIcons As Integer) _
    As Integer

    Public form_visualizzato As String = "Ticket"

    Public variabile_iniziazione As Integer = 0
    Public filtro_contenuto As String


    Private Sub Pianificazione_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
    End Sub

    Sub inizializzazione_form()

        variabile_iniziazione = 1
        Lbl_Nome_Reparto.Text = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).nome_reparto
        Carica_Reparti()
        riempi_tickets(DataGridView1)
        variabile_iniziazione = 1
    End Sub






    Private Sub Cmd_Cambia_Click(sender As Object, e As EventArgs) Handles Cmd_Cambia.Click
        Form_Cambia_Reparto.Show()

    End Sub


    Private Sub Cmd_Nuovo_Click(sender As Object, e As EventArgs) Handles Cmd_Nuovo.Click
        Ultima_Riga = ""
        Form_nuovo_ticket.Show()
        Form_nuovo_ticket.Inserimento_dipendenti()

        Form_nuovo_ticket.Reparto = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto
        Form_nuovo_ticket.Administrator = 1
        Form_nuovo_ticket.Startup()
        Form_nuovo_ticket.ComboBox2.Text = Homepage.business

    End Sub

    Private Function Reparto(Num_Reparto As Integer) As String
        Return Elenco_Reparti(Num_Reparto)
    End Function

    Private Sub Carica_Reparti()




        Dim Cnn_Reparto As New SqlConnection
        Cnn_Reparto.ConnectionString = Homepage.sap_tirelli
        Cnn_Reparto.Open()
        Dim Cmd_Reparto As New SqlCommand
        Dim Reader_Reparto As SqlDataReader

        Cmd_Reparto.Connection = Cnn_Reparto
        Cmd_Reparto.CommandText = "SELECT Id_Reparto,Descrizione 
FROM [TIRELLI_40].[DBO].COLL_Reparti
WHERE active ='Y' "
        Reader_Reparto = Cmd_Reparto.ExecuteReader()


        Do While Reader_Reparto.Read()
            Elenco_Reparti(Reader_Reparto("Id_Reparto")) = Reader_Reparto("Descrizione")
        Loop
        Cnn_Reparto.Close()
    End Sub


    Private Sub Cmd_Esci_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub


    Sub riempi_tickets(par_datagridview As DataGridView)

        Dim contatore As Integer = 0

        If RadioButton4.Checked = True Then
            filtro_reparto = "and t0.destinatario= '" & CODICE_REPARTO & "'"

        Else
            filtro_reparto = ""
        End If
        par_datagridview.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        CMD_SAP_2.CommandText = "

SELECT *
from
(
select
t0.[Id_Ticket]
,coalesce(t5.[Descrizione_Motivo],'') as 'Descrizione_Motivo'
      ,t0.[Commessa], 
case WHEN t12.itemname is not null then t12.itemname when t6.itemname is null then '' else t6.itemname end as 'Itemname',
	  case when substring(t0.COMMESSA,1,3)='CDS' THEN case when t12.U_Final_customer_name is null then t11.custmrName else t12.U_Final_customer_name end when t7.cardname is null then t6.u_final_customer_name else t7.cardname end as 'Cliente'
      ,t0.[Data_Creazione]
	  , DATEDIFF(day,t0.[Data_Creazione], getdate()) as 'giorni'
      ,t0.[Data_Chiusura]
      ,t0.[Data_Prevista_Chiusura]
      ,t0.[Aperto]
      ,t0.[Descrizione] 
	  ,t4.Descrizione as 'Mittente_padre'
      ,t1.[descrizione] as 'Mittente'
      ,t2.[descrizione] as 'Destinatario'
      ,t0.[Immagine]
      ,t0.[Id_Padre]
      ,t0.[Business]
, t0.oggetto
      ,t0.[Utente], concat(t9.firstname,' ', t9.lastname) as 'Nome_utente'
, concat(t10.firstname,' ', t10.lastname) as 'Utente_padre'
, case when t0.assegnato is null then '' else concat(t8.firstname,' ', t8.lastname) end as 'Assegnato'
      ,t0.[Data_chiusura_totale], case when t0.aperto =1 then 'Y' else 'N' end as 'stato'
,coalesce(t0.tpr,'') as 'TPR'
,coalesce(t0.riunione,'') as 'Riunione'

from  [TIRELLI_40].[DBO].coll_tickets t0 
  left join [TIRELLI_40].[DBO].COLL_Reparti t1 on t1.Id_Reparto=t0.Mittente
  left join [TIRELLI_40].[DBO].COLL_Reparti t2 on t2.Id_Reparto=t0.destinatario
  left join [TIRELLI_40].[DBO].COLL_Tickets t3 on t3.Id_Ticket= t0.id_padre
  left join [TIRELLI_40].[DBO].COLL_Reparti t4 on t4.Id_Reparto=t3.Mittente
  left join [TIRELLI_40].[DBO].COLL_motivazione t5 on t5.Id_Motivo = t0.Motivazione
LEFT JOIN [TIRELLISRLDB].[DBO].oitm t6 on t6.itemcode=t0.[Commessa]
left join [TIRELLISRLDB].[DBO].ocrd t7 on t7.cardcode=t6.u_final_customer_code
left join [TIRELLI_40].[DBO].ohem t8 on t8.empid=t0.assegnato
left join [TIRELLI_40].[DBO].ohem t9 on t9.empid=t0.utente
left join [TIRELLI_40].[DBO].ohem t10 on t10.empid=t3.utente
left join [TIRELLISRLDB].[DBO].oscl t11 on cast(t11.callid as varchar) = CAST(substring(t0.COMMESSA,4,999) AS VARCHAR) and substring(t0.COMMESSA,1,3)='CDS'
left join [TIRELLISRLDB].[DBO].oitm t12 on t12.itemcode=t11.itemcode
" & filtro_articolo & "
left join
(select t0.[Id_Padre], max(t0.[Id_Ticket]) as 'Ticket_max' from [TIRELLI_40].[DBO].coll_tickets t0 group by t0.[Id_Padre] ) a on t0.[Id_Ticket]=a.[Ticket_max]

 where " & status_1 & " " & filtro_reparto & " " & filtro_commessa & " " & filtro_id & " " & filtro_id_padre & "
)
as t10

 where " & status_2 & " " & filtro_cliente & " " & filtro_mittente_padre & " " & filtro_business & " " & filtro_utente_padre & " " & filtro_utente & filtro_riunione & filtro_contenuto & filtro_assegnato & " 

  order by t10.giorni DESC"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader_2.Read()
            Try


                par_datagridview.Rows.Add(cmd_SAP_reader_2("Id_Ticket"), cmd_SAP_reader_2("Id_padre"), cmd_SAP_reader_2("Descrizione_Motivo"), cmd_SAP_reader_2("Commessa"), cmd_SAP_reader_2("Itemname"), cmd_SAP_reader_2("Cliente"), cmd_SAP_reader_2("Business"), cmd_SAP_reader_2("Mittente_padre"), cmd_SAP_reader_2("Mittente"), cmd_SAP_reader_2("Destinatario"), cmd_SAP_reader_2("Assegnato"), cmd_SAP_reader_2("Data_creazione"), cmd_SAP_reader_2("Giorni"), cmd_SAP_reader_2("Stato"), cmd_SAP_reader_2("Oggetto"), cmd_SAP_reader_2("Riunione"), cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("TPR"))
            Catch ex As Exception

            End Try
            contatore += 1
        Loop

        Label1.Text = contatore

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        par_datagridview.ClearSelection()

    End Sub



    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)
        If variabile_iniziazione = 0 Then

        Else
            If form_visualizzato = "Ticket" Then
                riempi_tickets(DataGridView1)

            ElseIf form_visualizzato = "task" Then
                riempi_tasks()
            End If
        End If


    End Sub



    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting

        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Causale").Value = "Richiesta di Miglioria" Then
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Gray

        End If
        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Open").Value = "Y" Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="Open").Style.BackColor = Color.OrangeRed
        Else
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="Open").Style.BackColor = Color.Lime

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = Nothing Then
            filtro_commessa = ""
        Else

            filtro_commessa = "and t0.commessa Like '%%" & TextBox1.Text & "%%'"
        End If
        riempi_tickets(DataGridView1)
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        If e.RowIndex >= 0 Then
            riga = e.RowIndex

            If e.ColumnIndex = DataGridView1.Columns.IndexOf(ID) Then

                Dim new_form_visualizza_ticket = New Form_Visualizza_Ticket

                new_form_visualizza_ticket.Show()



                new_form_visualizza_ticket.Show()
                new_form_visualizza_ticket.Inserimento_dipendenti()

                new_form_visualizza_ticket.Reparto = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto
                new_form_visualizza_ticket.Administrator = 1
                new_form_visualizza_ticket.Txt_Id.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ID").Value
                new_form_visualizza_ticket.Startup()

            End If
        End If

    End Sub

    Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
        Form_Visualizza_Ticket.Show()
        Form_Visualizza_Ticket.Inserimento_dipendenti()

        Form_Visualizza_Ticket.Reparto = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto
        Form_Visualizza_Ticket.Administrator = 1
        Form_Visualizza_Ticket.Txt_Id.Text = DataGridView1.Rows(riga).Cells(columnName:="ID").Value
        Form_Visualizza_Ticket.Startup()

    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "[]" Then

            Me.WindowState = FormWindowState.Maximized
            Button1.Text = "Riduci"
        ElseIf Button1.Text = "Riduci" Then
            Me.WindowState = FormWindowState.Normal
            Button1.Text = "[]"
        End If

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = Nothing Then
            filtro_cliente = ""
        Else
            filtro_cliente = "and t10.cliente Like '%%" & TextBox2.Text & "%%'"
        End If
        riempi_tickets(DataGridView1)
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        status_1 = "t0.aperto=1"
        status_2 = "t10.aperto=1"
        If variabile_iniziazione = 0 Then

        Else
            riempi_tickets(DataGridView1)
        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        status_1 = "t0.aperto=0"
        status_2 = "t10.aperto=0"
        If variabile_iniziazione = 0 Then
        Else
            riempi_tickets(DataGridView1)
        End If

    End Sub



    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = Nothing Then
            filtro_mittente_padre = ""
        Else
            filtro_mittente_padre = "and t10.mittente_padre Like '%%" & TextBox3.Text & "%%'"
        End If
        riempi_tickets(DataGridView1)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = Nothing Then
            filtro_id = ""
        Else

            filtro_id = "and t0.[Id_Ticket] Like '%%" & TextBox4.Text & "%%'"
        End If
        riempi_tickets(DataGridView1)
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text = Nothing Then
            filtro_business = ""
        Else
            filtro_business = "and t10.business Like '%%" & TextBox5.Text & "%%'"
        End If
        riempi_tickets(DataGridView1)
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If TextBox6.Text = Nothing Then
            filtro_id_padre = ""
        Else

            filtro_id_padre = "and t0.[Id_padre] Like '%%" & TextBox6.Text & "%%'"
        End If
        riempi_tickets(DataGridView1)
    End Sub



    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.Text = Nothing Then
            filtro_utente_padre = ""
        Else
            filtro_utente_padre = "and t10.Utente_padre Like '%%" & TextBox7.Text & "%%'"
        End If
        riempi_tickets(DataGridView1)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Process.Start("\\tirfs01\00-Responsible\KPI\Analisi tickets.xlsx")
    End Sub

    Sub riempi_tasks()

        DataGridView2.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        CMD_SAP_2.CommandText = "select t0.id,t0.oc, t1.cardname
, CASE WHEN T5.CARDNAME IS NULL THEN '' ELSE t5.cardname end as 'Cliente_finale'
, t0.task, t2.Nome_task, t0.reparto,t3.Descrizione,t0.riferimento
, t4.riferimento as 'Nome_riferimento', t0.giorni,t0.stato,t0.linenum
, t0.data_inizio, t0.data_fine, t0.id_link, t0.Data_chiusura_task, t0.Ora_chiusura_task 

from [Tirelli_40].[dbo].[Pianificazione_CDS] t0 inner join ordr t1 on t0.oc=t1.docnum
left join [Tirelli_40].[dbo].[Pianificazione_CDS_TASK] t2 on t2.id =t0.task
left join [TIRELLI_40].[DBO].COLL_Reparti t3 on t0.reparto=t3.Id_Reparto

  left join [Tirelli_40].[dbo].[Pianificazione_CDS_Riferimenti] t4 on t0.Riferimento=t4.id
left join ocrd t5 on t5.cardcode=t1.u_CODICEBP

inner join

(select t0.oc, min(t0.linenum) as 'linenum'
from [Tirelli_40].[dbo].[Pianificazione_CDS] t0 where t0.stato='P'
group by t0.oc) A on A.linenum=t0.Linenum and t0.oc=a.oc where 0=0 " & filtro_reparto_task & "

"



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader_2.Read()

            DataGridView2.Rows.Add(cmd_SAP_reader_2("Id"), cmd_SAP_reader_2("OC"), cmd_SAP_reader_2("cardname"), cmd_SAP_reader_2("Cliente_finale"), cmd_SAP_reader_2("Task"), cmd_SAP_reader_2("Nome_task"), cmd_SAP_reader_2("Reparto"), cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("riferimento"), cmd_SAP_reader_2("Nome_riferimento"), cmd_SAP_reader_2("Giorni"), cmd_SAP_reader_2("Stato"), cmd_SAP_reader_2("Linenum"), cmd_SAP_reader_2("Data_inizio"), cmd_SAP_reader_2("Data_fine"), cmd_SAP_reader_2("Id_link"), cmd_SAP_reader_2("Data_chiusura_task"), cmd_SAP_reader_2("Ora_chiusura_task"))

        Loop



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        DataGridView2.ClearSelection()

    End Sub

    Private Sub tabpage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Enter

        riempi_tickets(DataGridView1)
        form_visualizzato = "Ticket"

    End Sub

    Private Sub tabpage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter
        riempi_tasks()
        form_visualizzato ="task"

    End Sub



    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.RowIndex >= 0 Then
            riga = e.RowIndex

            If e.ColumnIndex = DataGridView2.Columns.IndexOf(ID_) Then
                Task_visualizza.id_task = DataGridView2.Rows(e.RowIndex).Cells(columnName:="ID_").Value

                Task_visualizza.Show()


            End If
        End If
    End Sub



    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        If TextBox8.Text = Nothing Then
            filtro_utente = ""
        Else
            filtro_utente = "and t10.nome_Utente Like '%%" & TextBox8.Text & "%%'"
        End If
        riempi_tickets(DataGridView1)
    End Sub



    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        If variabile_iniziazione = 0 Then

        Else
            If form_visualizzato = "Ticket" Then
                riempi_tickets(DataGridView1)

            ElseIf form_visualizzato = "task" Then
                riempi_tasks()
            End If
        End If
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form_impostazioni_ticket.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        status_1 = "(t0.aperto=1 or t0.aperto=0)"
        status_2 = "(t10.aperto=1 or t10.aperto=0)"
        If variabile_iniziazione = 0 Then

        Else
            riempi_tickets(DataGridView1)
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dim par_datagridview As DataGridView = DataGridView1
        ' Creare un'applicazione Excel
        Dim excelApp As New Excel.Application
        excelApp.Visible = True ' Mostrare Excel all'utente

        ' Creare un nuovo foglio di lavoro
        Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add
        Dim excelWorksheet As Excel.Worksheet = CType(excelWorkbook.Worksheets(1), Excel.Worksheet)

        ' Aggiungere intestazioni alla prima riga del foglio di lavoro (facoltativo)
        For col As Integer = 1 To par_datagridview.Columns.Count
            excelWorksheet.Cells(1, col) = par_datagridview.Columns(col - 1).HeaderText
        Next

        ' Aggiungere dati alla DataGridView al foglio di lavoro
        For row As Integer = 0 To par_datagridview.Rows.Count - 1
            For col As Integer = 0 To par_datagridview.Columns.Count - 1
                excelWorksheet.Cells(row + 2, col + 1) = par_datagridview.Rows(row).Cells(col).Value
            Next
        Next

        ' Salvare il file Excel
        Dim saveFileDialog As New SaveFileDialog
        saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            excelWorkbook.SaveAs(saveFileDialog.FileName)
            MessageBox.Show("Esportazione completata con successo!", "Successo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        ' Chiudere Excel
        excelApp.Quit()
        ReleaseComObject(excelApp)
    End Sub

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        If TextBox9.Text = Nothing Then
            filtro_riunione = ""
        Else
            filtro_riunione = "and coalesce(t10.riunione,'') Like '%%" & TextBox9.Text & "%%'"
        End If
        riempi_tickets(DataGridView1)
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        If TextBox10.Text = Nothing Then
            filtro_articolo = ""
        Else
            filtro_articolo = " inner join [Tirelli_40].[dbo].[COLL_Riferimenti] t13 on t13.codice_sap Like '%%" & TextBox10.Text & "%%' and t13.Rif_Ticket=t0.[Id_Ticket] "
        End If
        riempi_tickets(DataGridView1)
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged
        If RichTextBox1.Text = Nothing Then
            filtro_contenuto = ""
        Else
            filtro_contenuto = "and t10.descrizione Like '%%" & RichTextBox1.Text & "%%'"
        End If

        riempi_tickets(DataGridView1)
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        If TextBox11.Text = Nothing Then
            filtro_assegnato = ""
        Else
            filtro_assegnato = " and t10.Assegnato Like '%%" & TextBox11.Text & "%%' "
        End If
        riempi_tickets(DataGridView1)
    End Sub
End Class
