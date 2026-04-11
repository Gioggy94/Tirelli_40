Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb


Public Class Form_Richiesta_Materiale

    Public Stringa_Connessione_SAP As String
    Public Livello_Scelta = 0
    Public Elenco_Riferimenti(1000) As String
    Public Num_Riferimenti As Integer
    Public Elenco_Reparti(1000) As Integer
    Public Num_Reparti As Integer
    Public Elenco_Dipendenti(1000) As String
    Public Consumabili As Integer
    Public Consumabili_Qta As Integer
    Public Consumabili_Minimo As Integer
    Public Categoria As String

    Private Sub Form_Richiesta_Materiale_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Larghezza_Colonna_Immagine As Integer
        Dim Larghezza_Colonna_Descrizione As Integer
        Dim Larghezza_Colonna_Codice As Integer
        Dim Larghezza_Colonna_Categoria As Integer
        Dim Larghezza_Colonna_Taglio As Integer
        Dim Larghezza_Colonna_Qta As Integer
        Dim Larghezza_Colonna_Minimo As Integer

        Larghezza_Colonna_Immagine = DataGrid_Materiale.Width * 20 / 100
        Larghezza_Colonna_Descrizione = DataGrid_Materiale.Width * 50 / 100
        Larghezza_Colonna_Codice = DataGrid_Materiale.Width * 10 / 100
        If TXT_ODP.Text.Length > 0 Then
            Larghezza_Colonna_Categoria = DataGrid_Materiale.Width * 10 / 100
            Larghezza_Colonna_Taglio = DataGrid_Materiale.Width * 10 / 100
        Else
            Larghezza_Colonna_Categoria = 0
            Larghezza_Colonna_Taglio = 0
        End If


        Larghezza_Colonna_Qta = DataGrid_Materiale.Width * 10 / 100
        Larghezza_Colonna_Minimo = DataGrid_Materiale.Width * 10 / 100

        Dim myFont As System.Drawing.Font
        myFont = New System.Drawing.Font("Arial", 17, FontStyle.Bold) ' Or FontStyle.Italic)

        Dim Col_Immagine_Materiale As New DataGridViewImageColumn
        Col_Immagine_Materiale.HeaderText = ""
        Col_Immagine_Materiale.ImageLayout = DataGridViewImageCellLayout.Zoom
        Col_Immagine_Materiale.Width = Larghezza_Colonna_Immagine
        DataGrid_Materiale.Columns.Add(Col_Immagine_Materiale)

        Dim Col_Descrizione As New DataGridViewTextBoxColumn
        Col_Descrizione.HeaderText = "Set Commessa"
        Col_Descrizione.Width = Larghezza_Colonna_Descrizione
        Col_Descrizione.CellTemplate.Style.Font = myFont
        DataGrid_Materiale.Columns.Add(Col_Descrizione)

        Dim Col_Codice As New DataGridViewTextBoxColumn
        Col_Codice.HeaderText = "Codice"
        Col_Codice.Width = Larghezza_Colonna_Codice
        Col_Codice.CellTemplate.Style.Font = myFont
        DataGrid_Materiale.Columns.Add(Col_Codice)

        Dim Col_Categoria As New DataGridViewTextBoxColumn
        Col_Categoria.HeaderText = "Categoria"
        Col_Categoria.Width = Larghezza_Colonna_Categoria
        Col_Categoria.CellTemplate.Style.Font = myFont
        DataGrid_Materiale.Columns.Add(Col_Categoria)

        Dim Col_Taglio As New DataGridViewTextBoxColumn
        Col_Taglio.HeaderText = "Taglio"
        Col_Taglio.Width = Larghezza_Colonna_Taglio
        Col_Taglio.CellTemplate.Style.Font = myFont
        DataGrid_Materiale.Columns.Add(Col_Taglio)

        Dim Col_Qta As New DataGridViewTextBoxColumn
        Col_Qta.HeaderText = "Qta"
        Col_Qta.Width = Larghezza_Colonna_Qta
        Col_Qta.CellTemplate.Style.Font = myFont
        DataGrid_Materiale.Columns.Add(Col_Qta)

        Dim Col_Minimo As New DataGridViewTextBoxColumn
        Col_Minimo.HeaderText = "Minimo"
        Col_Minimo.Width = Larghezza_Colonna_Minimo
        Col_Minimo.CellTemplate.Style.Font = myFont
        DataGrid_Materiale.Columns.Add(Col_Minimo)

        Num_Riferimenti = 0
        List_Materiale.Items.Clear()
        Cmd_Invia.Enabled = False
        Txt_Codice.Text = ""
        Txt_Descrizione.Text = ""
        Txt_Codice.Enabled = False
        Txt_Descrizione.Enabled = False
        TXT_ODP.Enabled = False
        Txt_Commessa.Enabled = False

        Dim Indice As Integer
        Dim Indice_Combo As Integer
        Dim Cnn_Reparti As New SqlConnection

        Cnn_Reparti.ConnectionString = homepage.sap_tirelli
        Cnn_Reparti.Open()

        Dim Cmd_Reparti As New SqlCommand
        Dim Reader_Reparti As SqlDataReader

        Cmd_Reparti.Connection = Cnn_Reparti
        Cmd_Reparti.CommandText = "SELECT * FROM [TIRELLI_40].[DBO].COLL_Reparti ORDER BY Descrizione"
        Reader_Reparti = Cmd_Reparti.ExecuteReader()
        Indice = 0
        Indice_Combo = -1
        Combo_Mittente.Items.Clear()

        Do While Reader_Reparti.Read()
            Elenco_Reparti(Indice) = Reader_Reparti("Id_Reparto")
            Combo_Mittente.Items.Add(Reader_Reparti("Descrizione"))
            Indice = Indice + 1
        Loop
        Num_Reparti = Indice
        Cnn_Reparti.Close()
    End Sub

    Public Sub Home_Lista()
        DataGrid_Materiale.Rows.Clear()
        Dim Cnn_Materiale As New SqlConnection
        Dim Indice As Integer
        Cnn_Materiale.ConnectionString = homepage.sap_tirelli
        Cnn_Materiale.Open()

        Dim Cmd_Materiale As New SqlCommand
        Dim Cmd_Materiale_Reader As SqlDataReader
        Indice = 0
        Cmd_Materiale.Connection = Cnn_Materiale
        If TXT_ODP.Text.Length > 0 Then
            Cmd_Materiale.CommandText = " SELECT * FROM [TIRELLI_40].[DBO].COLL_materiale WHERE Codice='0'"
            Consumabili = 0
            Cmd_Aggiungi.Text = "Aggiungi"
        Else
            Cmd_Materiale.CommandText = " SELECT * FROM [TIRELLI_40].[DBO].COLL_materiale WHERE Codice='1'"
            Consumabili = 1
            Cmd_Aggiungi.Text = "Preleva"
        End If
        Cmd_Materiale_Reader = Cmd_Materiale.ExecuteReader
        Do While Cmd_Materiale_Reader.Read()
            DataGrid_Materiale.Rows.Add()
            Try
                DataGrid_Materiale.Rows(Indice).Cells(0).Value = Image.FromFile(Cmd_Materiale_Reader("Immagine"))
            Catch ex As Exception
            End Try
            DataGrid_Materiale.Rows(Indice).Cells(1).Value = Cmd_Materiale_Reader("Descrizione")
            DataGrid_Materiale.Rows(Indice).Cells(2).Value = ""
            DataGrid_Materiale.Rows(Indice).Cells(3).Value = Cmd_Materiale_Reader("Categoria")
            DataGrid_Materiale.Rows(Indice).Cells(4).Value = Cmd_Materiale_Reader("Taglio")
            DataGrid_Materiale.Rows(Indice).Cells(5).Value = Cmd_Materiale_Reader("Qta")
            DataGrid_Materiale.Rows(Indice).Cells(6).Value = Cmd_Materiale_Reader("Minimo")
            DataGrid_Materiale.Rows(Indice).Height = 100
            Indice = Indice + 1
        Loop
        Cnn_Materiale.Close()
        Livello_Scelta = 0
        Txt_Lunghezza.Enabled = False
        Txt_Qta.Enabled = False
        Txt_Codice.Text = ""
        Txt_Descrizione.Text = ""
        Txt_Codice.Enabled = False
        Txt_Descrizione.Enabled = False
        Cmd_Aggiungi.Enabled = False
    End Sub

    Private Sub List_Materiale_DoubleClick(sender As Object, e As EventArgs) Handles List_Materiale.DoubleClick
        If List_Materiale.SelectedIndex >= 0 Then
            If MsgBox("Eliminare il riferimento " & Elenco_Riferimenti(List_Materiale.SelectedIndex) & " - " & Elenco_Riferimenti(List_Materiale.SelectedIndex), vbYesNo, "Eliminare Materiale?") = vbYes Then
                Dim i As Integer
                For i = List_Materiale.SelectedIndex To Num_Riferimenti - 1 Step 1
                    Elenco_Riferimenti(i) = Elenco_Riferimenti(i + 1)
                Next
                Num_Riferimenti = Num_Riferimenti - 1
                Compila_Lista_Materiale()
            End If
        End If
    End Sub

    Private Sub Compila_Lista_Materiale()
        List_Materiale.Items.Clear()
        Dim i As Integer
        For i = 0 To Num_Riferimenti - 1 Step 1
            List_Materiale.Items.Add(Elenco_Riferimenti(i))
        Next
        If Num_Riferimenti = 0 Then
            Cmd_Invia.Enabled = False
        End If
    End Sub

    Private Sub DataGrid_Materiale_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGrid_Materiale.CellContentClick
        If e.RowIndex >= 0 Then
            If DataGrid_Materiale.Rows(e.RowIndex).Cells(2).Value.ToString.Length = 0 Then
                Dim Cnn_Materiale As New SqlConnection
                Dim Indice As Integer
                Cnn_Materiale.ConnectionString = homepage.sap_tirelli
                Cnn_Materiale.Open()

                Dim Cmd_Materiale As New SqlCommand
                Dim Cmd_Materiale_Reader As SqlDataReader
                Indice = 0
                Cmd_Materiale.Connection = Cnn_Materiale

                Cmd_Materiale.CommandText = " SELECT * FROM [TIRELLI_40].[DBO].COLL_materiale WHERE (NOT Codice='0') AND (NOT Codice='1') AND Categoria=" & DataGrid_Materiale.Rows(e.RowIndex).Cells(3).Value.ToString & " ORDER BY Descrizione"

                DataGrid_Materiale.Rows.Clear()
                    Cmd_Materiale_Reader = Cmd_Materiale.ExecuteReader
                Do While Cmd_Materiale_Reader.Read()
                    DataGrid_Materiale.Rows.Add()
                    Try
                        DataGrid_Materiale.Rows(Indice).Cells(0).Value = Image.FromFile(Cmd_Materiale_Reader("Immagine"))
                    Catch ex As Exception
                    End Try
                    DataGrid_Materiale.Rows(Indice).Cells(1).Value = Cmd_Materiale_Reader("Descrizione")
                    DataGrid_Materiale.Rows(Indice).Cells(2).Value = Cmd_Materiale_Reader("Codice")
                    DataGrid_Materiale.Rows(Indice).Cells(3).Value = Cmd_Materiale_Reader("Categoria")
                    DataGrid_Materiale.Rows(Indice).Cells(4).Value = Cmd_Materiale_Reader("Taglio")
                    DataGrid_Materiale.Rows(Indice).Cells(5).Value = Cmd_Materiale_Reader("Qta")
                    DataGrid_Materiale.Rows(Indice).Cells(6).Value = Cmd_Materiale_Reader("Minimo")
                    DataGrid_Materiale.Rows(Indice).Height = 100
                    Indice = Indice + 1
                Loop
                Cnn_Materiale.Close()
                Livello_Scelta = 1
                Txt_Lunghezza.Enabled = False
                Txt_Qta.Enabled = False
            Else
                If DataGrid_Materiale.Rows(e.RowIndex).Cells(2).Value.ToString.Length > 0 Then
                    Categoria = DataGrid_Materiale.Rows(e.RowIndex).Cells(3).Value.ToString
                    Txt_Codice.Text = DataGrid_Materiale.Rows(e.RowIndex).Cells(2).Value.ToString
                    Txt_Descrizione.Text = DataGrid_Materiale.Rows(e.RowIndex).Cells(1).Value.ToString
                    Txt_Qta.Text = ""
                    Txt_Lunghezza.Text = ""
                    If DataGrid_Materiale.Rows(e.RowIndex).Cells(4).Value.ToString = "1" Then
                        Txt_Lunghezza.Enabled = True
                    Else
                        Txt_Lunghezza.Enabled = False
                    End If

                    Txt_Qta.Enabled = True
                    Cmd_Aggiungi.Enabled = True
                    Consumabili_Qta = Val(DataGrid_Materiale.Rows(e.RowIndex).Cells(5).Value.ToString)
                    Consumabili_Minimo = Val(DataGrid_Materiale.Rows(e.RowIndex).Cells(6).Value.ToString)
                End If
            End If
        End If
    End Sub

    Private Sub Cmd_Home_Click(sender As Object, e As EventArgs) Handles Cmd_Home.Click
        Home_Lista()
    End Sub

    Private Sub Cmd_Aggiungi_Click(sender As Object, e As EventArgs) Handles Cmd_Aggiungi.Click
        If Combo_Utente.Text.Length > 1 Then
            If Txt_Lunghezza.Enabled And Txt_Lunghezza.Text = "" Then
                MsgBox("Inserire la Lunghezza")
            Else
                If Txt_Qta.Text = "" Then
                    MsgBox("Inserire la Quantità")
                Else
                    If Consumabili = 0 Then
                        If Txt_Lunghezza.Enabled Then
                            Elenco_Riferimenti(Num_Riferimenti) = "- " & Txt_Qta.Text & "Pz L=" & Txt_Lunghezza.Text & " - " & Txt_Codice.Text & " - " & Txt_Descrizione.Text
                            Num_Riferimenti = Num_Riferimenti + 1
                            Compila_Lista_Materiale()

                            Txt_Lunghezza.Enabled = False
                            Txt_Qta.Enabled = False
                            Txt_Codice.Text = ""
                            Txt_Descrizione.Text = ""
                            Txt_Lunghezza.Text = ""
                            Txt_Qta.Text = ""
                            Txt_Codice.Enabled = False
                            Txt_Descrizione.Enabled = False
                            Cmd_Aggiungi.Enabled = False
                            Cmd_Invia.Enabled = True
                        Else
                            Elenco_Riferimenti(Num_Riferimenti) = "- " & Txt_Qta.Text & "Pz - " & Txt_Codice.Text & " - " & Txt_Descrizione.Text
                            Num_Riferimenti = Num_Riferimenti + 1
                            Compila_Lista_Materiale()
                            Txt_Lunghezza.Enabled = False
                            Txt_Qta.Enabled = False
                            Txt_Codice.Text = ""
                            Txt_Descrizione.Text = ""
                            Txt_Lunghezza.Text = ""
                            Txt_Qta.Text = ""
                            Txt_Codice.Enabled = False
                            Txt_Descrizione.Enabled = False
                            Cmd_Aggiungi.Enabled = False
                            Cmd_Invia.Enabled = True
                        End If
                    Else
                        ' Aggiorno il Database dei Consumabili

                        Consumabili_Qta = Consumabili_Qta - Val(Txt_Qta.Text)
                        Dim Cnn_Consumabili As New SqlConnection
                        Cnn_Consumabili.ConnectionString = homepage.sap_tirelli
                        Cnn_Consumabili.Open()
                        Dim Cmd_Consumabili As New SqlCommand
                        Cmd_Consumabili.Connection = Cnn_Consumabili
                        Cmd_Consumabili.CommandText = "UPDATE [TIRELLI_40].[DBO].COLL_materiale
                                  SET Qta='" & Consumabili_Qta & "'
                                  WHERE Codice='" & Txt_Codice.Text & "'"
                        Cmd_Consumabili.ExecuteNonQuery()
                        Cnn_Consumabili.Close()

                        ' Controlla il Minimo

                        If Consumabili_Qta < Consumabili_Minimo And Consumabili_Minimo > 0 Then
                            Invia_Mail_Acquisto()
                        End If

                        ' Aggiorno il LOG

                        Dim Descrizione As String

                        Descrizione = Val(Txt_Qta.Text) & " Pz - " & Txt_Codice.Text & " - " & Txt_Descrizione.Text

                        Dim Cnn_Log_Materiale As New SqlConnection
                        Cnn_Log_Materiale.ConnectionString = homepage.sap_tirelli
                        Cnn_Log_Materiale.Open()
                        Dim Cmd_Log_Materiale As New SqlCommand
                        Cmd_Log_Materiale.Connection = Cnn_Log_Materiale
                        Cmd_Log_Materiale.CommandText = "INSERT INTO [TIRELLI_40].[DBO].COLL_log_materiale
                                                (Data,Utente,Descrizione)
                                                VALUES('" & Now.ToString("yyyyMMdd") & "'
                                                , '" & Combo_Utente.Text & "'
                                                , '" & Descrizione & "'
                                                )"
                        Cmd_Log_Materiale.ExecuteNonQuery()
                        Cnn_Log_Materiale.Close()

                        Txt_Lunghezza.Enabled = False
                        Txt_Qta.Enabled = False
                        Txt_Codice.Text = ""
                        Txt_Descrizione.Text = ""
                        Txt_Lunghezza.Text = ""
                        Txt_Qta.Text = ""
                        Txt_Codice.Enabled = False
                        Txt_Descrizione.Enabled = False
                        Cmd_Aggiungi.Enabled = False

                        ' Aggiorno Sub Categoria

                        Dim Cnn_Materiale As New SqlConnection
                        Dim Indice As Integer
                        Cnn_Materiale.ConnectionString = homepage.sap_tirelli
                        Cnn_Materiale.Open()

                        Dim Cmd_Materiale As New SqlCommand
                        Dim Cmd_Materiale_Reader As SqlDataReader
                        Indice = 0
                        Cmd_Materiale.Connection = Cnn_Materiale

                        Cmd_Materiale.CommandText = " SELECT * FROM [TIRELLI_40].[DBO].COLL_materiale WHERE (NOT Codice='0') AND (NOT Codice='1') AND Categoria='" & Categoria & "' ORDER BY Descrizione"

                        DataGrid_Materiale.Rows.Clear()
                        Cmd_Materiale_Reader = Cmd_Materiale.ExecuteReader
                        Do While Cmd_Materiale_Reader.Read()
                            DataGrid_Materiale.Rows.Add()
                            Try
                                DataGrid_Materiale.Rows(Indice).Cells(0).Value = Image.FromFile(Cmd_Materiale_Reader("Immagine"))
                            Catch ex As Exception
                            End Try
                            DataGrid_Materiale.Rows(Indice).Cells(1).Value = Cmd_Materiale_Reader("Descrizione")
                            DataGrid_Materiale.Rows(Indice).Cells(2).Value = Cmd_Materiale_Reader("Codice")
                            DataGrid_Materiale.Rows(Indice).Cells(3).Value = Cmd_Materiale_Reader("Categoria")
                            DataGrid_Materiale.Rows(Indice).Cells(4).Value = Cmd_Materiale_Reader("Taglio")
                            DataGrid_Materiale.Rows(Indice).Cells(5).Value = Cmd_Materiale_Reader("Qta")
                            DataGrid_Materiale.Rows(Indice).Cells(6).Value = Cmd_Materiale_Reader("Minimo")
                            DataGrid_Materiale.Rows(Indice).Height = 100
                            Indice = Indice + 1
                        Loop
                        Cnn_Materiale.Close()
                        Livello_Scelta = 1
                        Txt_Lunghezza.Enabled = False
                        Txt_Qta.Enabled = False

                    End If
                End If
            End If
        Else
            MsgBox("Selezionare il proprio nome nella lista Utente")
        End If
    End Sub

    Sub Inserimento_dipendenti()
        Combo_Utente.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[DBO].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code inner join [TIRELLI_40].[DBO].COLL_Reparti t2 on (t2.sap_id_reparto =t1.code or t2.sap_id_reparto_2 =t1.code)   where t0.active='Y' AND (T0.POSITION<>3 OR T0.POSITION IS NULL) and t2.id_reparto='" & Elenco_Reparti(Combo_Mittente.SelectedIndex) & "'  order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_Dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            Combo_Utente.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Private Sub Combo_Mittente_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Mittente.SelectedIndexChanged
        Inserimento_dipendenti()
    End Sub

    Private Sub Cmd_Annulla_Click(sender As Object, e As EventArgs) Handles Cmd_Annulla.Click
        Owner.Show()
        Me.Close()
    End Sub

    Public Sub Invia_Mail_Acquisto()
        Dim Testo_Mail As String
        Testo_Mail = "<BODY><H3>Richiesta di Acquisto Automatica da Totem</h3><P>"
        Testo_Mail = Testo_Mail & "<BR>Si richiede l'acquisto dell'articolo"
        Testo_Mail = Testo_Mail & "<BR>Codice : " & Txt_Codice.Text
        Testo_Mail = Testo_Mail & "<BR>Descrizione : " & Txt_Descrizione.Text
        Testo_Mail = Testo_Mail & "<BR>Qta : " & Consumabili_Minimo

        Testo_Mail = Testo_Mail & "<BR><BR>Questo è un messaggio automatico. Non rispondere a questa mail"


        Testo_Mail = Testo_Mail & "</P></BODY>"

        Dim mySmtp As New SmtpClient
        Dim myMail As New MailMessage()
        mySmtp.UseDefaultCredentials = False
        mySmtp.Credentials = New Net.NetworkCredential("report@tirelli.net", "Ras70773")
        mySmtp.Host = "smtp.office365.com"
        mySmtp.Port = 587
        mySmtp.EnableSsl = True
        mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network


        myMail = New MailMessage()
        myMail.From = New MailAddress("report@tirelli.net")
        myMail.To.Add("nicolaravenotti@tirelli.net")
        myMail.Bcc.Add("report@tirelli.net")
        myMail.Subject = "Richiesta di Acquisto di Materiali Consumabili"
        myMail.IsBodyHtml = True
        myMail.Body = Testo_Mail
        Try
            mySmtp.Send(myMail)
        Catch ex As Exception
            MsgBox("Errore Invio Mail" & ex.ToString)
        End Try
        MsgBox("Mail Inviata")
    End Sub

    Private Sub Cmd_Invia_Click(sender As Object, e As EventArgs) Handles Cmd_Invia.Click
        Nuovo_Ticket()
    End Sub

    Private Function Nuovo_ID() As Integer
        Dim Cnn_Ticket As New SqlConnection
        Dim Risultato As Integer

        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader

        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "SELECT MAX(Id_Ticket) As 'Massimo' FROM [TIRELLI_40].[DBO].coll_tickets"
        Reader_Ticket = Cmd_Ticket.ExecuteReader()
        Reader_Ticket.Read()
        If Not DBNull.Value.Equals(Reader_Ticket("Massimo")) Then
            Risultato = Reader_Ticket("Massimo") + 1
        Else
            Risultato = 1
        End If
        Cnn_Ticket.Close()
        Return Risultato
    End Function

    Private Sub Nuovo_Ticket()
        'Inserimento Ticket

        Dim Id As Integer
        Id = Nuovo_ID()
        Dim Stringa_Immagine As String
        Stringa_Immagine = ""

        Dim Descrizione As String

        Descrizione = Combo_Utente.Text & " " & vbCrLf & "Richiesta Taglio Materiale per ODP n." & TXT_ODP.Text & vbCrLf

        Dim i As Integer
        For i = 0 To Num_Riferimenti - 1 Step 1
            Descrizione = Descrizione & Elenco_Riferimenti(i) & vbCrLf
        Next
        If Num_Riferimenti = 0 Then
            Cmd_Invia.Enabled = False
        End If

        Dim Data_Prevista As Date



        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "INSERT INTO [TIRELLI_40].[DBO].coll_tickets
                                                (Id_Ticket,Commessa,Data_Creazione,Data_Chiusura,Data_Prevista_Chiusura,
                                                Aperto,Descrizione,Mittente,Destinatario,Immagine,Motivazione,Id_Padre, BUSINESS,utente)
                                                VALUES(" & Id & "
                                                , '" & Txt_Commessa.Text.ToUpper & "'
                                                , '" & Now.ToString("yyyy-MM-dd") & "'
                                                , '" & Data_Prevista.ToString("yyyy-MM-dd") & "'
                                                , '" & Data_Prevista.ToString("yyyy-MM-dd") & "'
                                                , 1
                                                , '" & Descrizione & "'
                                                , " & Elenco_Reparti(Combo_Mittente.SelectedIndex) & "
                                                , " & "5" & "
                                                , '" & Stringa_Immagine & "'
                                                , " & 8 & "
                                                , " & Id & ", '" & "CONTINUING" & "', '" & Elenco_Dipendenti(Combo_Utente.SelectedIndex) & "'
                                                )"
        Cmd_Ticket.ExecuteNonQuery()
        Cnn_Ticket.Close()





        'MsgBox("Ticket Inserito Con Successo")
        Invia_Mail(Id)

        Me.Close()
    End Sub


    Private Function Cerca_Reparto(id As Integer) As String
        Dim Cnn_Reparto As New SqlConnection
        Cnn_Reparto.ConnectionString = homepage.sap_tirelli
        Cnn_Reparto.Open()
        Dim Cmd_Reparto As New SqlCommand
        Dim Reader_Reparto As SqlDataReader
        Dim Risultato As String

        Cmd_Reparto.Connection = Cnn_Reparto
        Cmd_Reparto.CommandText = "SELECT Descrizione FROM [TIRELLI_40].[DBO].COLL_Reparti WHERE Id_Reparto=" & id
        Reader_Reparto = Cmd_Reparto.ExecuteReader()
        Reader_Reparto.Read()
        Risultato = Reader_Reparto("Descrizione")
        Cnn_Reparto.Close()
        Return Risultato
    End Function

    Public Sub Invia_Mail(id As Integer)
        Dim Cnn_Ticket As New SqlConnection

        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()

        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader

        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "SELECT * FROM [TIRELLI_40].[DBO].coll_tickets,[TIRELLI_40].[DBO].COLL_reparti,

[TIRELLI_40].[DBO].COLL_motivazione
WHERE Id_Motivo=Motivazione AND Destinatario=Id_Reparto AND Id_Ticket=" & id
        Reader_Ticket = Cmd_Ticket.ExecuteReader()
        Reader_Ticket.Read()

        Dim Data_Creazione As Date
        Data_Creazione = Reader_Ticket("Data_Creazione")

        Dim Testo_Mail As String
        Testo_Mail = "<BODY><H3>Nuovo Ticket</h3><P>"
        Testo_Mail = Testo_Mail & "Hai ricevuto un nuovo ticket in riferimento alla commessa " & Reader_Ticket("Commessa")
        Testo_Mail = Testo_Mail & "<BR><BR>Data di Creazione : " & Data_Creazione.ToString("dd/MM/yyyy")
        Testo_Mail = Testo_Mail & "<BR>Mittente : " & Cerca_Reparto(Reader_Ticket("Mittente"))
        Testo_Mail = Testo_Mail & "<BR>Commessa : " & Reader_Ticket("Commessa")
        Testo_Mail = Testo_Mail & "<BR>Descrizione : " & Reader_Ticket("Descrizione")
        Testo_Mail = Testo_Mail & "<BR>Tipologia : " & Reader_Ticket("Descrizione_Motivo")


        Testo_Mail = Testo_Mail & "<BR><BR>Utilizzare l'applicazione <a href='" & Homepage.percorso_server & " TIRELLI\00-Tirelli 4.0\T4.0vb\Eseguibili\Tirelli 4.0.exe'>Tickets</a> per consultare l'elenco dei Tickets aperti e per poter inoltrare la risposta"
        Testo_Mail = Testo_Mail & "<BR>Questo è un messaggio automatico. Non rispondere a questa mail"


        Testo_Mail = Testo_Mail & "</P></BODY>"

        Dim mySmtp As New SmtpClient
        Dim myMail As New MailMessage()
        mySmtp.UseDefaultCredentials = False
        mySmtp.Credentials = New Net.NetworkCredential(Homepage.Mittente_Mail, Pianificazione_Tickets.Password_Mail)
        mySmtp.Host = "smtp.office365.com"
        mySmtp.Port = 25
        mySmtp.EnableSsl = True
        mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network


        myMail = New MailMessage()
        myMail.From = New MailAddress("report@tirelli.net")
        myMail.To.Add(Reader_Ticket("Mail_1"))
        If Reader_Ticket("Mail_2").ToString.Length > 1 Then
            myMail.To.Add(Reader_Ticket("Mail_2"))
        End If
        myMail.Bcc.Add("report@tirelli.net")
        myMail.Subject = Reader_Ticket("Commessa") & " - Inserimento Nuovo Ticket per " & Reader_Ticket("Descrizione_Motivo")
        myMail.IsBodyHtml = True
        myMail.Body = Testo_Mail

        Try
            mySmtp.Send(myMail)
        Catch ex As Exception
            MsgBox("Errore Invio Mail" & ex.ToString)
        End Try
        Cnn_Ticket.Close()
    End Sub


End Class