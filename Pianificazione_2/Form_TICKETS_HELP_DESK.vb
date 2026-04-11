Imports System.Data.SqlClient
Imports System.DirectoryServices.ActiveDirectory
Imports System.IO
Imports System.Windows.Controls

Imports System.Net.Mail
Imports System.Collections
Imports System.Data

Imports System.Data.OleDb

Public Class Form_TICKETS_HELP_DESK
    Public Elenco_Reparti(1000) As Integer
    Public Elenco_dipendenti(1000) As String
    Public codice_bp As String
    Public Codice_BP_JG_selezionato As String
    Public Immagine_Caricata As Integer = 0
    Private id_ticket As Integer
    Private stato_ticket As String
    Private numero_revisione As Integer

    Public Sub Startup()

        Inserimento_dipendenti()



    End Sub

    Sub select_ticket(par_id_ticket As Integer)


        Dim Cnn_Reparti As New SqlConnection

        Cnn_Reparti.ConnectionString = homepage.sap_tirelli
        Cnn_Reparti.Open()

        Dim Cmd_Reparti As New SqlCommand
        Dim Reader_Reparti As SqlDataReader

        Cmd_Reparti.Connection = Cnn_Reparti
        Cmd_Reparti.CommandText = "SELECT t0.[ID]
      ,t0.[Id_Ticket]
      ,t0.[Commessa]
      ,t0.[Codice_cliente]
	  ,coalesce(t1.cardname,'') as 'Cardname'
      ,t0.[Data_Creazione]
      ,t0.[Data_Chiusura]
      ,t0.[stato]
      ,t0.[Descrizione]
      ,t0.[Mittente]
	  ,concat(t2.lastname,' ',t2.firstname) as 'Nome_mittente'
      ,t0.[destinatario]
	   ,concat(t3.lastname,' ',t3.firstname) as 'Nome_destinatario'
      ,t0.[Immagine]
      ,t0.[file]
      ,t0.[tipo_problema]
      ,t0.[causale]
      ,t0.[n_revisione]
  FROM [TIRELLI_40].[dbo].[Help_Desk_Tickets] t0
  left join ocrd t1 on t0.Codice_cliente=t1.cardcode
  left join [TIRELLI_40].[DBO].ohem t2 on t2.empid=t0.Mittente
  left join [TIRELLI_40].[DBO].ohem t3 on t3.empid=t0.destinatario
where t0.id_ticket=" & par_id_ticket & ""
        Reader_Reparti = Cmd_Reparti.ExecuteReader()



        If Reader_Reparti.Read() Then
            Label3.Text = Reader_Reparti("Id_ticket")
            id_ticket = Reader_Reparti("Id_ticket")
            TextBox1.Text = Reader_Reparti("Commessa")
            Label1.Text = Reader_Reparti("cardname")
            codice_bp = Reader_Reparti("Codice_cliente")
            Label4.Text = Reader_Reparti("Data_Creazione")

            If Reader_Reparti("stato") = "R" Then
                RadioButton1.Checked = True
            ElseIf Reader_Reparti("stato") = "N" Then
                RadioButton2.Checked = True
            ElseIf Reader_Reparti("stato") = "I" Then
                RadioButton3.Checked = True
            End If

            RichTextBox1.Text = Reader_Reparti("descrizione")
            ComboBox1.Text = Reader_Reparti("Nome_mittente")

            If Reader_Reparti("Immagine").ToString.Length > 0 Then
                Picture_ticket.SizeMode = PictureBoxSizeMode.Zoom
                Dim MyImage As Bitmap

                Try
                    MyImage = New Bitmap(Homepage.Percorso_Immagini_TICKETS_HELPDESK & Reader_Reparti("Immagine").ToString)
                    Picture_ticket.Image = MyImage
                    Immagine_Caricata = 1
                Catch ex As Exception
                    MessageBox.Show("Errore nel caricamento dell'immagine: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

            ComboBox2.Text = Reader_Reparti("tipo_problema")
            ComboBox3.Text = Reader_Reparti("Causale")


        End If
        Reader_Reparti.Close()
        Cnn_Reparti.Close()
        riempi_datagridview_crediti(codice_bp)
        select_files(par_id_ticket)
    End Sub

    Sub select_files(par_id_ticket As Integer)

        DataGridView_files.Rows.Clear()
        Dim Cnn_Reparti As New SqlConnection

        Cnn_Reparti.ConnectionString = homepage.sap_tirelli
        Cnn_Reparti.Open()

        Dim Cmd_Reparti As New SqlCommand
        Dim Reader_Reparti As SqlDataReader

        Cmd_Reparti.Connection = Cnn_Reparti
        Cmd_Reparti.CommandText = "SELECT TOP (1000) t0.[ID]
      ,t0.[Id_Ticket]
      ,t0.[file_name]
  FROM [TIRELLI_40].[dbo].[Help_Desk_Tickets_files] t0
where t0.id_ticket=" & par_id_ticket & ""
        Reader_Reparti = Cmd_Reparti.ExecuteReader()



        Do While Reader_Reparti.Read()

            DataGridView_files.Rows.Add(Reader_Reparti("file_name"), Homepage.Percorso_FILE_TICKETS_HELPDESK & Reader_Reparti("file_name"))



        Loop
        Reader_Reparti.Close()
        Cnn_Reparti.Close()
    End Sub



    Private Sub Form_TICKETS_HELP_DESK_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Startup()
        Timer1.Start()
        ' Abilita il trascinamento nella ListView

    End Sub

    Sub trova_ID()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "SELECT  max(coalesce(t0.id_TICKET,0)) +1 as 'ID_ticket' 
from [TIRELLISRLDB].[dbo].[Help_Desk_Tickets] t0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then

            If Not cmd_SAP_reader("ID_ticket") Is System.DBNull.Value Then
                id_ticket = cmd_SAP_reader("ID_ticket")
            Else
                id_ticket = 1
            End If
        Else
            id_ticket = 1
        End If

        cnn.Close()
        cmd_SAP_reader.Close()


    End Sub

    Private Sub aggiorna_Ticket_helpdesk()


        Dim Stringa_Immagine As String
        If Immagine_Caricata = 1 Then
            Stringa_Immagine = "Ticket_" & id_ticket & ".jpg"
            Try
                If File.Exists(Homepage.Percorso_Immagini_TICKETS_HELPDESK & Stringa_Immagine) Then
                    File.Delete(Homepage.Percorso_Immagini_TICKETS_HELPDESK & Stringa_Immagine)
                End If
            Catch ex As Exception

            End Try
            Try
                Picture_ticket.Image.Save(Homepage.Percorso_Immagini_TICKETS_HELPDESK & Stringa_Immagine)
            Catch ex As Exception

            End Try


        Else
            Stringa_Immagine = ""
        End If




        RichTextBox2.Text = RichTextBox1.Text & vbCrLf & Now & " | " & ComboBox1.Text & vbCrLf & RichTextBox2.Text
        RichTextBox2.Text = Replace(RichTextBox2.Text, "'", "''")

        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "UPDATE [dbo].[Help_Desk_Tickets]
SET
    [Commessa] = '" & TextBox1.Text & "',
    [Codice_cliente] = '" & codice_bp & "',
    [stato] = '" & stato_ticket & "',
    [data_chiusura] = case when '" & stato_ticket & "' ='R' and [data_chiusura] is null then getdate() end,
    [Descrizione] = '" & RichTextBox2.Text & "',
    [Mittente] = " & Elenco_dipendenti(ComboBox1.SelectedIndex) & ",
    [Immagine] = '" & Stringa_Immagine & "',
    [tipo_problema] = '" & ComboBox2.Text & "',
    [causale] = '" & ComboBox3.Text & "',
    [N_REVISIONE] = " & numero_revisione & "
WHERE [Id_Ticket] = " & id_ticket & ""

        Cmd_Ticket.ExecuteNonQuery()
        Cnn_Ticket.Close()




    End Sub

    Private Sub Nuovo_Ticket_helpdesk()
        numero_revisione = 0
        trova_ID()
        Dim Stringa_Immagine As String
        If Immagine_Caricata = 1 Then
            Stringa_Immagine = "Ticket_" & id_ticket & ".jpg"
            If File.Exists(Homepage.Percorso_Immagini_TICKETS_HELPDESK & Stringa_Immagine) Then
                File.Delete(Homepage.Percorso_Immagini_TICKETS_HELPDESK & Stringa_Immagine)
            End If

            Picture_ticket.Image.Save(Homepage.Percorso_Immagini_TICKETS_HELPDESK & Stringa_Immagine)
        Else
            Stringa_Immagine = ""
        End If




        RichTextBox2.Text = Now & " | " & ComboBox1.Text & vbCrLf & RichTextBox2.Text
        RichTextBox2.Text = Replace(RichTextBox2.Text, "'", "''")

        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "INSERT INTO [dbo].[Help_Desk_Tickets]
           ([Id_Ticket]
           ,[Commessa]
           ,[Codice_cliente]
           ,[Data_Creazione]

           ,[stato]
           ,[Descrizione]
           ,[Mittente]

           ,[Immagine]
           ,[tipo_problema]
           ,[causale]
,N_REVISIONE)

  VALUES
           (" & id_ticket & "
           ,'" & TextBox1.Text & "'
           ,'" & codice_bp & "'
           ,getdate()
           
           ,'" & stato_ticket & "'
           ,'" & RichTextBox2.Text & "'
           ," & Elenco_dipendenti(ComboBox1.SelectedIndex) & "

           ,'" & Stringa_Immagine & "'
           ,'" & ComboBox2.Text & "'
           ,'" & ComboBox3.Text & "'
," & numero_revisione & ")"

        Cmd_Ticket.ExecuteNonQuery()
        Cnn_Ticket.Close()




    End Sub

    Sub Inserimento_dipendenti()

        ComboBox1.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim Indice As Integer
        Indice = 0



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[DBO].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code
inner join [tirelli_40].dbo.coll_reparti t2 on (t2.sap_id_reparto =t1.code or t2.sap_id_reparto_2 =t1.code)  
where t0.active='Y' AND (T0.POSITION<>3 OR T0.POSITION IS NULL) 
order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            ComboBox1.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Private Sub Combo_Mittente_SelectedIndexChanged(sender As Object, e As EventArgs)
        Inserimento_dipendenti()
    End Sub



    Private Sub Cmd_Incolla_Click(sender As Object, e As EventArgs) Handles Cmd_Incolla.Click
        Picture_ticket.SizeMode = PictureBoxSizeMode.Zoom
        Picture_ticket.Image = Clipboard.GetImage
        If Picture_ticket.Image IsNot Nothing Then
            Immagine_Caricata = 1
        End If
    End Sub

    Private Sub Cmd_Cancella_Immagine_Click(sender As Object, e As EventArgs) Handles Cmd_Cancella_Immagine.Click
        Immagine_Caricata = 0
        Picture_ticket.Image = Nothing
    End Sub

    Private Sub Cmd_Zoom_Click(sender As Object, e As EventArgs) Handles Cmd_Zoom.Click
        Form_Zoom.Show()
        Form_Zoom.Picture_Zoom.Image = Picture_ticket.Image
        Form_Zoom.Owner = Me
        Me.Hide()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Button6.Text = "Aggiorna" Then
            If ComboBox1.SelectedIndex < 0 Then
                MsgBox("Selezionare un utente")
            ElseIf codice_bp = "" Then
                MsgBox("Selezionare un cliente")
            ElseIf TextBox1.Text = "" Then
                MsgBox("Selezionare una commessa")
            ElseIf ComboBox2.SelectedIndex < 0 Then
                MsgBox("Selezionare un tipo di problema")
            ElseIf ComboBox3.SelectedIndex < 0 Then
                MsgBox("Selezionare una causale")
            Else
                aggiorna_Ticket_helpdesk()
                copia_file()
                MsgBox("Ticket aggiornato con successo")
            End If

        ElseIf Button6.Text = "Nuovo" Then

            If ComboBox1.SelectedIndex < 0 Then
                MsgBox("Selezionare un utente")
            ElseIf codice_bp = "" Then
                MsgBox("Selezionare un cliente")
            ElseIf TextBox1.Text = "" Then
                MsgBox("Selezionare una commessa")
            ElseIf ComboBox2.SelectedIndex < 0 Then
                MsgBox("Selezionare un tipo di problema")
            ElseIf ComboBox3.SelectedIndex < 0 Then
                MsgBox("Selezionare una causale")
            Else

                Nuovo_Ticket_helpdesk()
                copia_file()
                MsgBox("Ticket creato con successo")
            End If




        End If
        Form_tickets_help_desk_tabella.startup()
        Me.Close()

    End Sub

    Private Sub copia_file()
        ' Percorso di destinazione
        Dim destinationFolder As String = Homepage.Percorso_FILE_TICKETS_HELPDESK

        ' Assicurati che ci siano righe nella DataGridView
        If DataGridView_files.Rows.Count > 0 Then
            ' Itera attraverso le righe della DataGridView
            For Each row As DataGridViewRow In DataGridView_files.Rows
                ' Ottieni il percorso del file dalla colonna "percorso"
                Dim filePath As String = row.Cells("percorso").Value.ToString()

                ' Verifica se il file esiste prima di copiarlo
                If File.Exists(filePath) Then
                    ' Ottieni il nome del file dalla colonna "nome file"
                    Dim fileName As String = "Ticket_" & id_ticket & "_" & row.Cells("Nome_file").Value.ToString()

                    ' Crea il percorso di destinazione sul desktop
                    Dim destinationPath As String = Path.Combine(destinationFolder, fileName)

                    ' Copia il file
                    File.Copy(filePath, destinationPath, True)
                    inserisci_file_in_db(fileName, id_ticket)
                Else
                    ' Gestisci il caso in cui il file non esiste
                    MessageBox.Show($"Il file non esiste: {filePath}", "File non trovato", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            Next

            ' Messaggio di conferma
            '  MessageBox.Show("Files copiati nella cartella di destinazione.", "Successo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            '  MessageBox.Show("La DataGridView è vuota.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    Sub inserisci_file_in_db(par_nome_file As String, par_id_ticket As Integer)
        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "

delete [dbo].[Help_Desk_Tickets_files] where id_ticket=" & par_id_ticket & "

INSERT INTO [dbo].[Help_Desk_Tickets_files]
           ([Id_Ticket]
           ,file_name)

  VALUES
           (" & par_id_ticket & ", '" & par_nome_file & "')"

        Cmd_Ticket.ExecuteNonQuery()
        Cnn_Ticket.Close()




    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            stato_ticket = "R"
        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            stato_ticket = "N"
        End If

    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then
            stato_ticket = "I"
        End If

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Close()
    End Sub




    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Create an OpenFileDialog instance
        Dim openFileDialog1 As New OpenFileDialog()

        ' Set the title of the dialog
        openFileDialog1.Title = "Choose a file"

        ' Set the initial directory (optional)
        openFileDialog1.InitialDirectory = "C:\"

        ' Set the filter for the file types you want to allow
        openFileDialog1.Filter = "All Files (*.*)|*.*|Text Files (*.txt)|*.txt"

        ' Show the dialog and check if the user selected a file
        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            ' Get the selected file path
            Dim filePath As String = openFileDialog1.FileName

            ' Get the file name
            Dim fileName As String = Path.GetFileName(filePath)

            ' Add the file name and file path to DataGridView_files
            DataGridView_files.Rows.Add(fileName, filePath)
        End If
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Business_partner.Show()


        Business_partner.Provenienza = "Help_desk_tickets_BP"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Assicurati che ci sia almeno una riga selezionata nella DataGridView
        If DataGridView_files.SelectedRows.Count > 0 Then
            ' Ottieni l'indice della riga selezionata
            Dim selectedRowIndex As Integer = DataGridView_files.SelectedRows(0).Index

            ' Ottieni il nome del file dalla colonna 0 (prima colonna)
            Dim fileName As String = DataGridView_files.Rows(selectedRowIndex).Cells(0).Value.ToString()

            ' Ottieni il percorso del file dalla colonna 1 (seconda colonna)
            Dim filePath As String = DataGridView_files.Rows(selectedRowIndex).Cells(1).Value.ToString()

            ' Chiedi all'utente conferma per cancellare
            Dim result As DialogResult = MessageBox.Show($"Sei sicuro di voler cancellare il file '{fileName}'?", "Conferma cancellazione", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            ' Se l'utente conferma, cancella la riga dalla DataGridView
            If result = DialogResult.Yes Then
                DataGridView_files.Rows.RemoveAt(selectedRowIndex)
                MessageBox.Show("File cancellato con successo.", "Cancellazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Else
            MessageBox.Show("Seleziona una riga da cancellare.", "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub



    Private Sub DataGridView_files_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_files.CellClick
        If e.RowIndex >= 0 Then
            Try
                Process.Start(DataGridView_files.Rows(e.RowIndex).Cells(columnName:="percorso").Value)
            Catch ex As Exception
                MsgBox("File non presente")
            End Try

        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim dataInserita As DateTime
        If DateTime.TryParse(Label4.Text, dataInserita) Then
            ' Calcola la differenza tra la data corrente e la data inserita
            Dim differenza As TimeSpan = DateTime.Now - dataInserita

            ' Visualizza la differenza nei vari formati
            Label2.Text = differenza.Days & " gg, " & differenza.Hours & " hh, " & differenza.Minutes & " mm, " & differenza.Seconds & " ss."
        Else
            ' Se la conversione della data non è riuscita
            Label2.Text = "-"
        End If
    End Sub

    Sub riempi_datagridview_crediti(par_codice_bp)

        DataGridView1.Rows.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "
select *
from 
(
SELECT t0.docentry, t0.docnum as 'Fattura',  T0.[DocDate] as 'Data Fatt', t0.docduedate as 'Data Scad',DATEDIFF(DAY, t0.docduedate, GETDATE()) as 'Overdue', T0.[CardCode] as 'BP code', T0.[CardName] as 'BP Name',t11.cardcode as 'Final_BP_code', t11.cardname as 'Final BP', T2.[Indicator] as 'Year',  t0.U_uffcompetenza as 'Department' ,t9.balance,   t5.slpname as 'Salesman', t10.name ,  T0.[DocTotal] as 'Total', T0.[PaidToDate] as 'Paid',t0.U_aggiustamentofattura as'Adjustment', T0.[DocTotal]-T0.[PaidToDate]-case when t0.U_aggiustamentofattura is null then '0' else t0.U_aggiustamentofattura end as 'Credit',   T0.[GroupNum], T3.[PymntGroup], t0.u_settore as 'Settore'
FROM OINV T0
INNER JOIN NNM1 T2 ON T0.[Series] = T2.[Series] 
inner join octg t3 on T0.[GroupNum]= t3.[GroupNum]

inner join OSLP T5 ON T5.slpcode =t0.slpcode

LEFT JOIN OCRD T9 ON T9.[CardCode] = T0.[CardCode]
INNER join OCRY t10 on t10.code = t9.country
left join ocrd t11 on t11.cardcode=t0.u_codicebp


WHERE T0.[docDate] >= (CONVERT(DATETIME, '20141001', 112) ) and T0.[DocTotal]-T0.[PaidToDate] >0 and t0.docentry<>'5848'and t0.docentry<>'6447' and t0.docentry<>'5199' and t0.docentry <>'7925' and t0.docentry<>'6882'and t0.docentry<>'7785'and t0.docentry<>'7528'and t0.docentry<>'7168' and t0.docentry<>'7426' and t0.docentry<>'7573' and t0.docentry<>'7932' and t0.docentry<>'8436'
)
as t10

where 0=0 and t10.[BP code]='" & par_codice_bp & "'

order by t10.overdue DESC 
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()
            DataGridView1.Rows.Add(False, cmd_SAP_reader("docentry"), cmd_SAP_reader("Fattura"), cmd_SAP_reader("Data Fatt"), cmd_SAP_reader("Data Scad"), cmd_SAP_reader("Overdue"), cmd_SAP_reader("BP code"), cmd_SAP_reader("BP Name"), cmd_SAP_reader("Final_BP_code"), cmd_SAP_reader("Final BP"), cmd_SAP_reader("Year"), cmd_SAP_reader("Department"), cmd_SAP_reader("Balance"), cmd_SAP_reader("Salesman"), cmd_SAP_reader("name"), cmd_SAP_reader("Total"), cmd_SAP_reader("Paid"), cmd_SAP_reader("Adjustment"), cmd_SAP_reader("Credit"), cmd_SAP_reader("PymntGroup"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()




    End Sub
End Class