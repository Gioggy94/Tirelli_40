Imports System.Data.SqlClient
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form_tickets_help_desk_tabella
    Private Sub Form_tickets_help_desk_tabella_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        startup()

    End Sub

    Sub startup()
        riempi_tickets()
        riempi_tickets_chiusi()
    End Sub

    Private Sub Cmd_Nuovo_Click(sender As Object, e As EventArgs) Handles Cmd_Nuovo.Click
        Form_TICKETS_HELP_DESK.Show()
        Form_TICKETS_HELP_DESK.Button6.Text = "Nuovo"
    End Sub

    Sub riempi_tickets()


        DataGridView1.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1

        CMD_SAP_2.CommandText = "SELECT TOP (100) [ID]
      ,t0.[Id_Ticket]
      ,t0.[Commessa]
      ,t0.[Codice_cliente]
,t1.cardname
      ,t0.[Data_Creazione]
      ,t0.[Data_Chiusura]
      ,t0.[stato]
      ,t0.[Descrizione]
      ,t0.[Mittente]
, concat(t2.lastname,' ',t2.firstname) as 'Nome_mittente'
      ,t0.[destinatario]
      ,t0.[Immagine]
      ,t0.[tipo_problema]
      ,t0.[causale]
      ,t0.[n_revisione]
  FROM [Tirelli_40].[dbo].[Help_Desk_Tickets] t0
left join ocrd t1 on t0.codice_cliente=t1.cardcode
left join [TIRELLI_40].[DBO].ohem t2 on t2.empid=t0.mittente
  where t0.stato='N' or t0.stato='I'
order by id_ticket
"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Dim percorso_immagine As String
        Do While cmd_SAP_reader_2.Read()
            percorso_immagine = Homepage.Percorso_Immagini_TICKETS_HELPDESK & cmd_SAP_reader_2("Immagine")
            If File.Exists(percorso_immagine) Then
            Else
                percorso_immagine = Homepage.Percorso_Immagini_TICKETS & "Bianco.jpg"
            End If
            ' Load the image from file path
            Dim image As Image = Image.FromFile(percorso_immagine)

            DataGridView1.Rows.Add(
                cmd_SAP_reader_2("Id"),
        cmd_SAP_reader_2("Id_Ticket"),
        cmd_SAP_reader_2("Commessa"),
        cmd_SAP_reader_2("Codice_cliente"),
        cmd_SAP_reader_2("cardname"),
        cmd_SAP_reader_2("Data_Creazione"),
        cmd_SAP_reader_2("stato"),
        cmd_SAP_reader_2("Descrizione"),
        cmd_SAP_reader_2("Mittente"),
        cmd_SAP_reader_2("Nome_mittente"),
        cmd_SAP_reader_2("destinatario"),
        image,
        cmd_SAP_reader_2("tipo_problema"),
        cmd_SAP_reader_2("causale"),
        cmd_SAP_reader_2("n_revisione")
    )

        Loop



        cmd_SAP_reader_2.Close()
        cnn1.Close()

        DataGridView1.ClearSelection()

    End Sub

    Sub riempi_tickets_chiusi()


        DataGridView2.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1

        CMD_SAP_2.CommandText = "SELECT TOP (10) [ID]
      ,t0.[Id_Ticket]
      ,t0.[Commessa]
      ,t0.[Codice_cliente]
,t1.cardname
      ,t0.[Data_Creazione]
      ,t0.[Data_Chiusura]
      ,t0.[stato]
      ,t0.[Descrizione]
      ,t0.[Mittente]
, concat(t2.lastname,' ',t2.firstname) as 'Nome_mittente'
      ,t0.[destinatario]
      ,t0.[Immagine]
      ,t0.[tipo_problema]
      ,t0.[causale]
      ,t0.[n_revisione]
  FROM [Tirelli_40].[dbo].[Help_Desk_Tickets] t0
left join ocrd t1 on t0.codice_cliente=t1.cardcode
left join [TIRELLI_40].[DBO].ohem t2 on t2.empid=t0.mittente
  where t0.stato='R'
order by t0.[Data_Chiusura] desc
"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Dim percorso_immagine As String
        Do While cmd_SAP_reader_2.Read()

            percorso_immagine = Homepage.Percorso_Immagini_TICKETS & "Bianco.jpg"

            If File.Exists(Homepage.Percorso_Immagini_TICKETS_HELPDESK & cmd_SAP_reader_2("Immagine")) Then
                percorso_immagine = Homepage.Percorso_Immagini_TICKETS_HELPDESK & cmd_SAP_reader_2("Immagine")

            End If


            ' Load the image from file path
            Dim image As Image = Image.FromFile(percorso_immagine)

            DataGridView2.Rows.Add(
                cmd_SAP_reader_2("Id"),
        cmd_SAP_reader_2("Id_Ticket"),
        cmd_SAP_reader_2("Commessa"),
        cmd_SAP_reader_2("Codice_cliente"),
        cmd_SAP_reader_2("cardname"),
        cmd_SAP_reader_2("Data_Creazione"),
        cmd_SAP_reader_2("Data_chiusura"),
        cmd_SAP_reader_2("stato"),
        cmd_SAP_reader_2("Descrizione"),
        cmd_SAP_reader_2("Mittente"),
        cmd_SAP_reader_2("Nome_mittente"),
        cmd_SAP_reader_2("destinatario"),
        image,
        cmd_SAP_reader_2("tipo_problema"),
        cmd_SAP_reader_2("causale"),
        cmd_SAP_reader_2("n_revisione"))

        Loop



        cmd_SAP_reader_2.Close()
        cnn1.Close()

        DataGridView2.ClearSelection()

    End Sub

    ' Function to convert byte array to Image
    Function ByteArrayToImage(ByVal byteArrayIn As Byte()) As Image
        Dim ms As New MemoryStream(byteArrayIn)
        Dim returnImage As Image = Image.FromStream(ms)
        Return returnImage
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        riempi_tickets()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Id_ticket) Then

                Form_TICKETS_HELP_DESK.Show()
                Form_TICKETS_HELP_DESK.select_ticket(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Id_ticket").Value)
                Form_TICKETS_HELP_DESK.Button6.Text = "Aggiorna"

                'If File.Exists(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disegno").Value & ".PDF") Then
                '    AxFoxitCtl1.OpenFile(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disegno").Value & ".PDF")

                '    AxFoxitCtl1.Show()
                'Else

                '    AxFoxitCtl1.Hide()
                'End If
            End If
        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = DataGridView2.Columns.IndexOf(DataGridViewButtonColumn2) Then

                Form_TICKETS_HELP_DESK.Show()
                Form_TICKETS_HELP_DESK.select_ticket(DataGridView2.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn2").Value)
                Form_TICKETS_HELP_DESK.Button6.Text = "Aggiorna"


            End If
        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting

        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="STATO").Value = "N" Then



            DataGridView1.Rows(e.RowIndex).Cells(columnName:="STATO").Style.BackColor = Color.OrangeRed

        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="STATO").Value = "I" Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="STATO").Style.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        If DataGridView2.Rows(e.RowIndex).Cells(columnName:="DataGridViewTextBoxColumn5").Value = "R" Then



            DataGridView2.Rows(e.RowIndex).Cells(columnName:="DataGridViewTextBoxColumn5").Style.BackColor = Color.Lime

        End If
    End Sub
End Class