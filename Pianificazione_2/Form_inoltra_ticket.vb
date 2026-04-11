Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb



Public Class Form_Inoltra_Ticket
    Public Structure Riferimento
        Public Rif As String
        Public Descrizione As String
        Public Tipo As String
    End Structure

    Public Elenco_Reparti(1000) As Integer
    Public Num_Reparti As Integer
    Public Elenco_Motivi(1000) As Integer

    Public Num_Motivi As Integer
    Public Stringa_Connessione_SAP As String
    Public Reparto As Integer
    Public Administrator As Integer
    Public Data_Creazione As Date
    Public Data_Chiusura As Date
    Public Data_Prevista As Date
    Public Elenco_Riferimenti(1000) As Riferimento
    Public Num_Riferimenti As Integer
    Public Immagine_Caricata As Integer
    Public Aperto As Integer
    Public Num_Mittente As Integer
    Public Num_Destinatario As Integer
    Public Num_Motivo As Integer
    Public Ticket_Aperto As Integer
    Public Num_Mittente_Padre As Integer
    Public Elenco_dipendenti(1000) As String
    Public TPR As String



    Public Sub Startup()
        Stringa_Connessione_SAP = homepage.sap_tirelli


        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader

        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "SELECT case when riunione is null then '' else riunione end as 'riunione', destinatario,motivazione,commessa,descrizione,id_padre,id_ticket,case when business is null then '' else business end as 'business', immagine, oggetto 
, case when tpr ='Y' then 'Y' else 'N' end as 'TPR'  
FROM [TIRELLI_40].[DBO].coll_tickets WHERE Id_Ticket=" & Txt_Id.Text
        Reader_Ticket = Cmd_Ticket.ExecuteReader()
        Reader_Ticket.Read()
        Num_Mittente = Reader_Ticket("Destinatario")
        Num_Motivo = Reader_Ticket("Motivazione")
        Txt_Commessa.Text = Reader_Ticket("Commessa")
        Data_Creazione = Today
        Txt_Data_Creazione.Text = Data_Creazione.ToString("dd/MM/yyyy")
        Ticket_Aperto = 1
        Txt_Descrizione.Text = Reader_Ticket("Descrizione")
        Txt_Id_Padre.Text = Reader_Ticket("Id_Padre")
        Txt_Id_Prec.Text = Reader_Ticket("Id_Ticket")
        ComboBox2.Text = Reader_Ticket("business")
        ComboBox3.Text = Reader_Ticket("riunione")

        If Reader_Ticket("TPR") = "N" Then
            CheckBox1.Checked = False
        Else
            CheckBox1.Checked = True
        End If

        If Not DBNull.Value.Equals(Reader_Ticket("Oggetto")) Then
            TextBox2.Text = Reader_Ticket("Oggetto")
        Else
            TextBox2.Text = Nothing
        End If

        If Reader_Ticket("Immagine").ToString.Length > 0 Then
            Picture_Campione.SizeMode = PictureBoxSizeMode.Zoom
            Dim MyImage As Bitmap
            Try
                MyImage = New Bitmap(Homepage.Percorso_Immagini_TICKETS & Reader_Ticket("Immagine").ToString)
            Catch ex As Exception
            End Try
            Picture_Campione.Image = CType(MyImage, Image)
            Immagine_Caricata = 1
        End If
        Cnn_Ticket.Close()

        Cnn_Ticket.Open()
        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "SELECT COLL_Tickets.Mittente, COLL_Reparti.Descrizione 
FROM [TIRELLI_40].[DBO].coll_tickets
,[TIRELLI_40].[DBO].COLL_Reparti 
WHERE Id_Ticket=" & Txt_Id_Padre.Text & " 
AND Coll_Tickets.Mittente=COLL_Reparti.Id_Reparto"
        Reader_Ticket = Cmd_Ticket.ExecuteReader()
        Reader_Ticket.Read()
        Num_Mittente_Padre = Reader_Ticket("Mittente")
        Lbl_Mittente_Padre.Text = Reader_Ticket("Descrizione")
        Cnn_Ticket.Close()

        Dim Indice As Integer
        Dim Indice_Mittente As Integer
        '  Dim Indice_Destinatario As Integer
        Dim Cnn_Reparti As New SqlConnection

        Cnn_Reparti.ConnectionString = Homepage.sap_tirelli
        Cnn_Reparti.Open()

        Dim Cmd_Reparti As New SqlCommand
        Dim Reader_Reparti As SqlDataReader

        Cmd_Reparti.Connection = Cnn_Reparti
        Cmd_Reparti.CommandText = "SELECT * 
FROM [TIRELLI_40].[DBO].COLL_Reparti
ORDER BY Descrizione"
        Reader_Reparti = Cmd_Reparti.ExecuteReader()
        Indice = 0
        Indice_Mittente = 0
        ' Indice_Destinatario = -1
        Combo_Mittente.Items.Clear()
        'Combo_Destinatario.Items.Clear()

        Do While Reader_Reparti.Read()
            Elenco_Reparti(Indice) = Reader_Reparti("Id_Reparto")
            Combo_Mittente.Items.Add(Reader_Reparti("Descrizione"))
            '    Combo_Destinatario.Items.Add(Reader_Reparti("Descrizione"))
            If Num_Mittente = Elenco_Reparti(Indice) Then
                Indice_Mittente = Indice
            End If
            Indice = Indice + 1
        Loop
        Num_Reparti = Indice
        Cnn_Reparti.Close()


        '        Dim Cnn_Motivi As New SqlConnection
        '        Cnn_Motivi.ConnectionString = Homepage.sap_tirelli
        '        Cnn_Motivi.Open()
        '        Dim Cmd_Motivi As New SqlCommand
        '        Dim Reader_Motivi As SqlDataReader
        '        Dim Indice_Motivo As Integer

        '        Cmd_Motivi.Connection = Cnn_Motivi
        '        Cmd_Motivi.CommandText = "SELECT * FROM 
        '[TIRELLI_40].[DBO].COLL_Motivazione 
        'where active='Y'

        'ORDER BY Descrizione_Motivo"
        '        Reader_Motivi = Cmd_Motivi.ExecuteReader()
        '        Indice = 0
        '        Indice_Motivo = 0
        '        Combo_Motivazione.Items.Clear()

        '        Do While Reader_Motivi.Read()
        '            Elenco_Motivi(Indice) = Reader_Motivi("Id_Motivo")
        '            Combo_Motivazione.Items.Add(Reader_Motivi("Descrizione_Motivo"))
        '            If Num_Motivo = Elenco_Motivi(Indice) Then
        '                Indice_Motivo = Indice
        '            End If
        '            Indice = Indice + 1
        '        Loop
        '        Num_Motivi = Indice
        '        Cnn_Motivi.Close()
        '        Combo_Motivazione.SelectedIndex = Indice_Motivo
        Combo_Mittente.SelectedIndex = Indice_Mittente

        '        Combo_Destinatario.SelectedIndex = Indice_Destinatario

        Txt_Data_Creazione.Enabled = False

        Txt_Data_Chiusura.Enabled = False
        Txt_Id_Padre.Enabled = False
        Txt_Id.Enabled = False
        Combo_Mittente.Enabled = False
        Combo_Destinatario.Enabled = True
        ' Combo_Motivazione.Enabled = False

        Txt_Commessa.Enabled = False
        Txt_Id.Text = ""
        Compila_Lista_Riferimenti_Startup()
    End Sub

    Private Sub Compila_Lista_Riferimenti_Startup()
        ListBox_Riferimenti.Items.Clear()
        Num_Riferimenti = 0

        Dim Cnn_Riferimenti As New SqlConnection
        Cnn_Riferimenti.ConnectionString = homepage.sap_tirelli
        Cnn_Riferimenti.Open()
        Dim Cmd_Riferimenti As New SqlCommand
        Dim Reader_Riferimenti As SqlDataReader
        Cmd_Riferimenti.Connection = Cnn_Riferimenti
        Cmd_Riferimenti.CommandText = "SELECT * FROM 
[TIRELLI_40].DBO.COLL_Riferimenti WHERE Rif_Ticket=" & Txt_Id_Prec.Text
        Reader_Riferimenti = Cmd_Riferimenti.ExecuteReader()
        Do While Reader_Riferimenti.Read()
            Elenco_Riferimenti(Num_Riferimenti).Rif = Reader_Riferimenti("Codice_SAP")
            Elenco_Riferimenti(Num_Riferimenti).Tipo = Reader_Riferimenti("Tipo_Codice")
            Elenco_Riferimenti(Num_Riferimenti).Descrizione = Descrivi_Codice(Elenco_Riferimenti(Num_Riferimenti).Rif, Elenco_Riferimenti(Num_Riferimenti).Tipo)
            Num_Riferimenti = Num_Riferimenti + 1
        Loop
        Cnn_Riferimenti.Close()

        Dim i As Integer
        For i = 0 To Num_Riferimenti - 1 Step 1
            ListBox_Riferimenti.Items.Add(Elenco_Riferimenti(i).Tipo & " - " & Elenco_Riferimenti(i).Rif & " - " & Elenco_Riferimenti(i).Descrizione)
        Next
    End Sub

    Private Function Descrivi_Codice(Codice As String, Tipo As String) As String
        Dim Risultato As String
        If Tipo = "Articolo" Then ' CODICE ARTICOLO
            Dim Cnn_Codice As New SqlConnection
            Cnn_Codice.ConnectionString = homepage.sap_tirelli
            Cnn_Codice.Open()
            Dim Cmd_Codice As New SqlCommand
            Dim Reader_Codice As SqlDataReader
            Cmd_Codice.Connection = Cnn_Codice
            If Homepage.ERP_provenienza = "SAP" Then
                Cmd_Codice.CommandText = "SELECT T0.[ItemCode], T0.[ItemName], SUM(T1.[OnHand]) as 'Al Magazzino', SUM(T1.[IsCommited]) as 'Impegnato', SUM(T1.[OnOrder]) as 'In Ordine' 
                                          FROM OITM T0  INNER JOIN OITW T1 ON T0.[ItemCode] = T1.[ItemCode] 
                                          WHERE T0.[ItemCode] ='" & Codice & "' 
                                          GROUP BY T0.[ItemCode], T0.[ItemName]"
            Else
                Cmd_Codice.CommandText = "  Select 

Trim(CODE) As 'itemCODE',

DES_CODE AS 'itemname',

CHECK_DB AS 'Code'

From OPENQUERY(AS400, '
    SELECT *
    From S786FAD1.TIR90VIS.JGALART
    Where code =  ''" & Codice & "''
') T10"
            End If

            Reader_Codice = Cmd_Codice.ExecuteReader()
            If Reader_Codice.Read() Then
                Risultato = Reader_Codice("ItemName")
            End If
            Cnn_Codice.Close()
        End If
        If Tipo = "Ordine" Then ' ORDINE DI PRODUZIONE
            Dim Cnn_Codice As New SqlConnection
            Cnn_Codice.ConnectionString = homepage.sap_tirelli
            Cnn_Codice.Open()
            Dim Cmd_Codice As New SqlCommand
            Dim Reader_Codice As SqlDataReader
            Cmd_Codice.Connection = Cnn_Codice
            If Homepage.ERP_provenienza = "SAP" Then
                Cmd_Codice.CommandText = "SELECT T0.[DocNum] as 'Ordine', T0.[ItemCode] as 'Codice', T0.[ProdName] as 'Descrizione', sum(T1.[U_PRG_WIP_QtaSpedita]) as 'Totale Trasferiti', sum(T1.[U_PRG_WIP_QtaDaTrasf]) as 'Totale da Trasferire', 
                                            CASE 
                                            WHEN T0.[Status]='P' THEN 'Pianificato' 
                                            WHEN T0.[Status]='C' THEN 'Stornato'
                                            WHEN T0.[Status]='L' THEN 'Chiuso'
                                            WHEN T0.[Status]='R' THEN 'Rilasciato'
                                            ELSE T0.[Status] END as 'Stato' 
                                            FROM OWOR T0 INNER JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
                                            WHERE T0.[DocNum] = '" & Codice & "' 
                                            GROUP BY T0.[Status], T0.[DocNum], T0.[ItemCode], T0.[ProdName]"


            Else
                Cmd_Codice.CommandText = " select 'Manca descrizione' as Descrizione'"
            End If

            Reader_Codice = Cmd_Codice.ExecuteReader()
                If Reader_Codice.Read() Then
                    Risultato = Reader_Codice("Descrizione")
                End If
                Cnn_Codice.Close()
            End If
            Return Risultato
    End Function

    Private Sub Compila_Lista_Riferimenti()
        ListBox_Riferimenti.Items.Clear()
        Dim i As Integer
        For i = 0 To Num_Riferimenti - 1 Step 1
            ListBox_Riferimenti.Items.Add(Elenco_Riferimenti(i).Tipo & " - " & Elenco_Riferimenti(i).Rif & " - " & Elenco_Riferimenti(i).Descrizione)
        Next
    End Sub


    Private Sub ListBox_Riferimenti_DoubleClick(sender As Object, e As EventArgs)
        If ListBox_Riferimenti.SelectedIndex >= 0 Then
            If MsgBox("Eliminare il riferimento " & Elenco_Riferimenti(ListBox_Riferimenti.SelectedIndex).Rif & " - " & Elenco_Riferimenti(ListBox_Riferimenti.SelectedIndex).Descrizione, vbYesNo, "Eliminare Riferimento") = vbYes Then
                Dim i As Integer
                For i = ListBox_Riferimenti.SelectedIndex To Num_Riferimenti - 1 Step 1
                    Elenco_Riferimenti(i).Rif = Elenco_Riferimenti(i + 1).Rif
                    Elenco_Riferimenti(i).Tipo = Elenco_Riferimenti(i + 1).Tipo
                    Elenco_Riferimenti(i).Descrizione = Elenco_Riferimenti(i + 1).Descrizione
                Next
                Num_Riferimenti = Num_Riferimenti - 1
                Compila_Lista_Riferimenti()
            End If
        End If
    End Sub


    Private Sub Cmd_Incolla_Click(sender As Object, e As EventArgs)
        Picture_Campione.SizeMode = PictureBoxSizeMode.Zoom
        Picture_Campione.Image = Clipboard.GetImage
        If Picture_Campione.Image IsNot Nothing Then
            Immagine_Caricata = 1
        Else
            MsgBox("Non caricata")
        End If
    End Sub

    Private Sub Cmd_Zoom_Click(sender As Object, e As EventArgs)
        Form_Zoom.Show()
        Form_Zoom.Picture_Zoom.Image = Picture_Campione.Image
        Form_Zoom.Owner = Me
        Me.Hide()
    End Sub



    Private Function Nuovo_ID() As Integer
        Dim Cnn_Ticket As New SqlConnection
        Dim Risultato As Integer

        Cnn_Ticket.ConnectionString = Homepage.sap_tirelli
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


    Private Sub Cmd_Aggiungi_Riferimento_Click(sender As Object, e As EventArgs)
        If Txt_Nuovo_Riferimento.Text.Length > 0 Then
            If Combo_Riferimenti.SelectedIndex = 0 Then ' CODICE ARTICOLO

                Dim Cnn_Codice As New SqlConnection
                Cnn_Codice.ConnectionString = Homepage.sap_tirelli
                Cnn_Codice.Open()
                Dim Cmd_Codice As New SqlCommand
                Dim Reader_Codice As SqlDataReader
                Cmd_Codice.Connection = Cnn_Codice
                Cmd_Codice.CommandText = "SELECT T0.[ItemCode], T0.[ItemName], SUM(T1.[OnHand]) as 'Al Magazzino', SUM(T1.[IsCommited]) as 'Impegnato', SUM(T1.[OnOrder]) as 'In Ordine' 
                                          FROM OITM T0  INNER JOIN OITW T1 ON T0.[ItemCode] = T1.[ItemCode] 
                                          WHERE T0.[ItemCode] ='" & Txt_Nuovo_Riferimento.Text & "' 
                                          GROUP BY T0.[ItemCode], T0.[ItemName]"
                Reader_Codice = Cmd_Codice.ExecuteReader()
                If Reader_Codice.Read() Then
                    Elenco_Riferimenti(Num_Riferimenti).Rif = Reader_Codice("ItemCode")
                    Elenco_Riferimenti(Num_Riferimenti).Descrizione = Reader_Codice("ItemName")
                    Elenco_Riferimenti(Num_Riferimenti).Tipo = "Articolo"
                    Num_Riferimenti = Num_Riferimenti + 1
                    Txt_Nuovo_Riferimento.Text = ""
                Else
                    MsgBox("Articolo Inesistente")
                    Txt_Nuovo_Riferimento.Text = ""
                End If
                Cnn_Codice.Close()
            End If
            If Combo_Riferimenti.SelectedIndex = 1 Then ' ORDINE DI PRODUZIONE
                Dim Cnn_Codice As New SqlConnection
                Cnn_Codice.ConnectionString = Homepage.sap_tirelli
                Cnn_Codice.Open()
                Dim Cmd_Codice As New SqlCommand
                Dim Reader_Codice As SqlDataReader
                Cmd_Codice.Connection = Cnn_Codice
                Cmd_Codice.CommandText = "SELECT T0.[DocNum] as 'Ordine', T0.[ItemCode] as 'Codice', T0.[ProdName] as 'Descrizione', sum(T1.[U_PRG_WIP_QtaSpedita]) as 'Totale Trasferiti', sum(T1.[U_PRG_WIP_QtaDaTrasf]) as 'Totale da Trasferire', 
                                            CASE 
                                            WHEN T0.[Status]='P' THEN 'Pianificato' 
                                            WHEN T0.[Status]='C' THEN 'Stornato'
                                            WHEN T0.[Status]='L' THEN 'Chiuso'
                                            WHEN T0.[Status]='R' THEN 'Rilasciato'
                                            ELSE T0.[Status] END as 'Stato' 
                                            FROM OWOR T0 INNER JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
                                            WHERE T0.[DocNum] = '" & Txt_Nuovo_Riferimento.Text & "' 
                                            GROUP BY T0.[Status], T0.[DocNum], T0.[ItemCode], T0.[ProdName]"
                Reader_Codice = Cmd_Codice.ExecuteReader()
                If Reader_Codice.Read() Then
                    Elenco_Riferimenti(Num_Riferimenti).Rif = Reader_Codice("Ordine")
                    Elenco_Riferimenti(Num_Riferimenti).Descrizione = Reader_Codice("Codice") & " - " & Reader_Codice("Descrizione")
                    Elenco_Riferimenti(Num_Riferimenti).Tipo = "Ordine"
                    Num_Riferimenti = Num_Riferimenti + 1
                    Txt_Nuovo_Riferimento.Text = ""
                Else
                    MsgBox("Articolo Inesistente")
                    Txt_Nuovo_Riferimento.Text = ""
                End If
                Cnn_Codice.Close()
            End If
        End If
        Compila_Lista_Riferimenti()
    End Sub


    Private Function Nuovo_ID_Riferimento() As Integer
        Dim Cnn_Ticket As New SqlConnection
        Dim Risultato As Integer

        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader

        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "SELECT MAX(Id_Riferimento) As 'Massimo' FROM 
[TIRELLI_40].DBO.COLL_Riferimenti"
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

    Private Sub Form_Inoltra_Ticket_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo

        Combo_Riferimenti.Items.Clear()
        Combo_Riferimenti.Items.Add("Articolo")
        Combo_Riferimenti.Items.Add("Ordine di Produzione")
        Form_nuovo_ticket.riempi_combobox_descrizioni_NC(ComboBox5)
        Combo_Riferimenti.SelectedIndex = 0
        Txt_Id.Text = ""
        'Me.CancelButton = Cmd_Annulla

    End Sub





    Private Sub Cmd_Inoltra_Click(sender As Object, e As EventArgs) Handles Cmd_Inoltra.Click
        If Combo_Destinatario.SelectedIndex < 0 Then

            MsgBox("Selezionare un Destinatario")
            Return

        End If

        If Combo_Motivazione.SelectedIndex < 0 Then
            MsgBox("Selezionare una motivazione")
            Return
        End If


        If ComboBox1.SelectedIndex < 0 Then

            MsgBox("Selezionare un utente mittente")
            Return
        End If


        If Txt_Descrizione.Text.Length < 1 Then


            MsgBox("Aggiungere una Descrizione")
            Return
        End If

        If TextBox3.Text = "" And (Elenco_Reparti(Combo_Destinatario.SelectedIndex) = 1 Or Elenco_Reparti(Combo_Destinatario.SelectedIndex) = 17 Or Elenco_Reparti(Combo_Destinatario.SelectedIndex) = 25 Or Elenco_Reparti(Combo_Destinatario.SelectedIndex) = 26) And (Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 5 Or Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 6) Then


            MsgBox("Indicare quanto tempo è stato impiegato per risolvere questa fase del Ticket")

            Return
        End If

        If TextBox4.Text = "" And (Elenco_Reparti(Combo_Mittente.SelectedIndex) = 4 Or Elenco_Reparti(Combo_Mittente.SelectedIndex) = 5) Then

            ' Se il mittente è 4 e il destinatario è 5, non bloccare
            If Elenco_Reparti(Combo_Mittente.SelectedIndex) = 4 And Elenco_Reparti(Combo_Destinatario.SelectedIndex) = 5 Then
                ' Salta il blocco
            Else
                MsgBox("Indicare il costo del materiale che è stato impiegato per risolvere questa fase del Ticket")
                Return
            End If

        End If

        'If Elenco_Motivi(Combo_Motivazione.SelectedIndex) = Nothing Or Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 0 Or Elenco_Motivi(Combo_Motivazione.SelectedIndex) = "" Then
        '    MsgBox("Selezionare una motivazione valida")
        '    Return
        'End If


        If Txt_Commessa.Text.Length = 0 Then



            If MsgBox("Nessuna Commessa Indicata. Continuare?", MsgBoxStyle.YesNo) = vbYes Then
                Txt_Commessa.Text = "Varie"
                Nuovo_Ticket_Inoltrato()

                MsgBox("Messaggio inoltrato con successo")
            Else
                Return
            End If




        Else
            Nuovo_Ticket_Inoltrato()

            MsgBox("Messaggio inoltrato con successo")

        End If



    End Sub

    Private Sub Chiudi_Ticket_Precedente()

        Dim tempo As Integer
        Dim costo As Integer

        If TextBox3.Text = "" Then
            tempo = 0

        Else
            tempo = TextBox3.Text
        End If

        If TextBox4.Text = "" Then
            costo = 0

        Else
            costo = TextBox4.Text
        End If


        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = Homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "Update [TIRELLI_40].[DBO].coll_tickets
                                  SET Aperto=0,Data_Chiusura=getdate() , tempo = " & tempo & " 
                                  , Costo = " & costo & " 
                                  WHERE Id_Ticket ='" & Txt_Id_Prec.Text & "'"
        Cmd_Ticket.ExecuteNonQuery()
        Cnn_Ticket.Close()
    End Sub


    Private Sub Nuovo_Ticket_Inoltrato()
        'Inserimento Ticket
        Dim new_text As String
        Txt_Id.Text = Nuovo_ID()
        Dim Stringa_Immagine As String
        If Immagine_Caricata = 1 Then
            Stringa_Immagine = "Ticket_" & Txt_Id.Text & ".jpg"

            ' If File.Exists(Homepage.Percorso_Immagini_TICKETS & Stringa_Immagine) Then
            Picture_Campione.Image.Save(Homepage.Percorso_Immagini_TICKETS & Stringa_Immagine)
            '  End If


        Else
                Stringa_Immagine = ""
        End If


        Txt_Descrizione.Text = Replace(Txt_Descrizione.Text, "'", " ")
        new_text = Txt_Descrizione.Text & " " & vbCrLf & vbCrLf & " " & ComboBox1.Text & " " & Now & vbCrLf & " " & TextBox1.Text
        new_text = Replace(new_text, "'", "")

        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Cmd_Ticket.Connection = Cnn_Ticket
        Dim par_descrizione_nc As Integer
        If ComboBox5.SelectedIndex = -1 Then
            par_descrizione_nc = 0
        Else
            par_descrizione_nc = Form_nuovo_ticket.Elenco_descrizione_nc(ComboBox5.SelectedIndex)
        End If

        Cmd_Ticket.CommandText = "INSERT INTO [TIRELLI_40].[DBO].coll_tickets
                                                (Id_Ticket,Commessa,Data_Creazione,Data_Chiusura,Data_Prevista_Chiusura,
                                                Aperto,Descrizione,Mittente,Destinatario,Immagine,Motivazione,Id_Padre,business, utente, riunione, oggetto,Tpr,descrizione_nc
)
                                                VALUES(" & Txt_Id.Text & "
                                                , '" & Txt_Commessa.Text & "'
                                                , '" & Data_Creazione.ToString("yyyy-MM-dd") & "'
                                                , '" & Data_Prevista.ToString("yyyy-MM-dd") & "'
                                                , '" & Data_Chiusura.ToString("yyyy-MM-dd") & "'
                                                , 1
                                                , '" & new_text & "'
                                                , " & Elenco_Reparti(Combo_Mittente.SelectedIndex) & "
                                                , " & Form_nuovo_ticket.Elenco_Reparti_destinatario(Combo_Destinatario.SelectedIndex) & "
                                                , '" & Stringa_Immagine & "'
                                                , " & Elenco_Motivi(Combo_Motivazione.SelectedIndex) & "
                                                , " & Txt_Id_Padre.Text & ",'" & ComboBox2.Text & "', '" & Elenco_dipendenti(ComboBox1.SelectedIndex) & "',
                                                '" & ComboBox3.Text & "','" & TextBox2.Text & "','" & TPR & "'," & par_descrizione_nc & "
                                               )"
        Cmd_Ticket.ExecuteNonQuery()
        Cnn_Ticket.Close()


        Dim Cnn_Riferimenti As New SqlConnection
        Cnn_Riferimenti.ConnectionString = homepage.sap_tirelli
        Cnn_Riferimenti.Open()
        Dim Cmd_Riferimenti As New SqlCommand
        Cmd_Riferimenti.Connection = Cnn_Riferimenti
        Dim i As Integer
        If Num_Riferimenti > 0 Then
            For i = 0 To Num_Riferimenti - 1 Step 1
                Cmd_Riferimenti.CommandText = "INSERT INTO [TIRELLI_40].DBO.COLL_Riferimenti
                                                (Id_Riferimento,Rif_Ticket,Codice_Sap,Tipo_Codice)
                                                VALUES(" & Nuovo_ID_Riferimento() & "
                                                , " & Txt_Id.Text & "
                                                , '" & Elenco_Riferimenti(i).Rif & "'
                                                , '" & Elenco_Riferimenti(i).Tipo & "'
                                                )"
                Cmd_Riferimenti.ExecuteNonQuery()
            Next
        End If

        Cnn_Riferimenti.Close()

        Chiudi_Ticket_Precedente()
        'MsgBox("Ticket Inoltrato Con Successo")
        Invia_Mail(Txt_Id.Text)
        Pianificazione_Tickets.Show()
        Pianificazione_Tickets.riempi_tickets(Pianificazione_Tickets.DataGridView1)
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
        Cmd_Reparto.CommandText = "SELECT Descrizione FROM 
[TIRELLI_40].[DBO].COLL_Reparti WHERE Id_Reparto=" & id
        Reader_Reparto = Cmd_Reparto.ExecuteReader()
        Reader_Reparto.Read()
        Risultato = Reader_Reparto("Descrizione")
        Cnn_Reparto.Close()
        Return Risultato
    End Function

    Public Sub Invia_Mail(id As Integer)
        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = Homepage.sap_tirelli
        Cnn_Ticket.Open()

        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader

        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "SELECT t0.Data_Creazione, t0.Commessa, " &
    "COALESCE(t3.ItemName, '') AS 'Descrizione_commessa', " &
    "COALESCE(t3.U_Final_customer_name, '') AS 'Cliente', " &
    "T4.Descrizione AS 'Reparto_mittente', t0.Descrizione, " &
    "t2.Descrizione_Motivo, t1.Mail_1, t1.Mail_2, t1.Mail_3 " &
    "FROM [TIRELLI_40].[DBO].coll_tickets t0 " &
    "INNER JOIN [TIRELLI_40].[DBO].COLL_Reparti t1 ON t0.Destinatario = t1.Id_Reparto " &
    "INNER JOIN [TIRELLI_40].[DBO].COLL_Motivazione t2 ON t2.Id_Motivo = t0.Motivazione " &
    "LEFT JOIN [TIRELLISRLDB].[DBO].OITM t3 ON t3.ItemCode = t0.Commessa " &
    "INNER JOIN [TIRELLI_40].[DBO].COLL_Reparti T4 ON T4.Id_Reparto = T0.Mittente " &
    "WHERE Id_Ticket = @IdTicket"

        Cmd_Ticket.Parameters.AddWithValue("@IdTicket", id)

        Reader_Ticket = Cmd_Ticket.ExecuteReader()
        If Reader_Ticket.Read() Then
            Dim Data_Creazione As Date = Reader_Ticket("Data_Creazione")

            ' Gestione delle interruzioni di riga nella descrizione
            Dim Descrizione As String = Reader_Ticket("Descrizione").ToString().Replace(vbCrLf, "<br>").Replace(vbCr, "<br>").Replace(vbLf, "<br>")

            Dim Testo_Mail As String = "<BODY>" &
        "<H3 style='color:#0056b3;'>📌 Nuovo Ticket N° " & id & "</H3>" &
        "<P>Hai ricevuto un nuovo ticket in riferimento alla commessa:<br>" &
        "<strong>🛠 Commessa:</strong> " & Reader_Ticket("Commessa") & " - " &
        Reader_Ticket("Descrizione_commessa") & " - " & Reader_Ticket("Cliente") & "</P>" &
        "<P><strong>📅 Data di Creazione:</strong> " & Data_Creazione.ToString("dd/MM/yyyy") & "<br>" &
        "<strong>📍 Mittente:</strong> " & Reader_Ticket("Reparto_mittente") & "<br>" &
        "<strong>📄 Descrizione:</strong> " & Descrizione & "<br>" &
        "<strong>📂 Tipologia:</strong> " & Reader_Ticket("Descrizione_Motivo") & "</P>"

            ' Aggiunta dell'elenco riferimenti, se presente
            If ListBox_Riferimenti.Items.Count > 0 Then
                Testo_Mail &= "<P><strong>📌 Elenco dei Riferimenti:</strong><br><ul>"
                For Each item In ListBox_Riferimenti.Items
                    Testo_Mail &= "<li>" & item.ToString() & "</li>"
                Next
                Testo_Mail &= "</ul></P>"
            End If

            ' Aggiunta delle note finali
            Testo_Mail &= "<P>🔍 Per maggiori dettagli, utilizzare l'applicazione <strong style='color:blue;'>Tirelli 4.0</strong> per consultare l'elenco dei ticket aperti e per inoltrare o chiudere la risposta.</P>" &
        "<P style='color:red;'><em>⚠️ Questo è un messaggio automatico. Non rispondere a questa mail.</em></P>" &
        "</BODY>"

            Using mySmtp As New SmtpClient("tirelli-net.mail.protection.outlook.com", 25)
                mySmtp.UseDefaultCredentials = False
                mySmtp.Credentials = New Net.NetworkCredential(Homepage.Mittente_Mail, Pianificazione_Tickets.Password_Mail)
                mySmtp.EnableSsl = True
                mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network

                Using myMail As New MailMessage()
                    myMail.From = New MailAddress(Homepage.Mittente_Mail)
                    myMail.To.Add(Reader_Ticket("Mail_1"))
                    If Not String.IsNullOrEmpty(Reader_Ticket("Mail_2").ToString()) Then myMail.To.Add(Reader_Ticket("Mail_2"))
                    If Not String.IsNullOrEmpty(Reader_Ticket("Mail_3").ToString()) Then myMail.To.Add(Reader_Ticket("Mail_3"))
                    myMail.Bcc.Add("report@tirelli.net")

                    myMail.Subject = "Nuovo ticket per " & Reader_Ticket("Commessa") & " " & Reader_Ticket("Descrizione_commessa") & " " & Reader_Ticket("Cliente")
                    myMail.IsBodyHtml = True
                    myMail.Body = Testo_Mail

                    Try
                        mySmtp.Send(myMail)
                    Catch ex As Exception
                        MsgBox("Errore Invio Mail: " & ex.ToString)
                    End Try
                End Using
            End Using
        End If

        Reader_Ticket.Close()
        Cnn_Ticket.Close()
    End Sub

    Private Sub Cmd_Annulla_Click(sender As Object, e As EventArgs)
        Pianificazione_Tickets.Show()
        Pianificazione_Tickets.riempi_tickets(Pianificazione_Tickets.DataGridView1)
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
        Me.Hide()
    End Sub

    Sub Inserimento_dipendenti()
        Dim reparto As Integer
        If Combo_Mittente.SelectedIndex = -1 Then
            reparto = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto
        Else
            reparto = Elenco_Reparti(Combo_Mittente.SelectedIndex)
        End If

        ComboBox1.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code inner join [TIRELLI_40].[DBO].coll_reparti t2 on (t2.sap_id_reparto =t1.code or t2.sap_id_reparto_2 =t1.code)   where t0.active='Y' and t2.id_reparto='" & reparto & "'  order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            ComboBox1.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub 'Inserisco le risorse nella combo box



    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TPR = "Y"
            Me.BackColor = Color.Wheat
        Else
            TPR = "N"
            Me.BackColor = Homepage.colore_sfondo
        End If
    End Sub
    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress, TextBox4.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "," AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True ' Blocca il carattere
        ElseIf e.KeyChar = "," AndAlso DirectCast(sender, TextBox).Text.Contains(",") Then
            e.Handled = True ' Blocca se c'è già una virgola
        End If
    End Sub

    Private Sub Combo_Mittente_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Mittente.SelectedIndexChanged

        If (Elenco_Reparti(Combo_Mittente.SelectedIndex) = 1 Or Elenco_Reparti(Combo_Mittente.SelectedIndex) = 17 Or Elenco_Reparti(Combo_Mittente.SelectedIndex) = 25 Or Elenco_Reparti(Combo_Mittente.SelectedIndex) = 26) Then
            GroupBox20.Visible = True
        Else
            GroupBox20.Visible = False
        End If
        'And (Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 5 Or Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 6)

        If Elenco_Reparti(Combo_Mittente.SelectedIndex) <> 4 And Elenco_Reparti(Combo_Mittente.SelectedIndex) <> 5 Then
            GroupBox21.Visible = False
        Else
            GroupBox21.Visible = True
        End If
        Form_nuovo_ticket.riempi_combobox_destinatario(Combo_Destinatario, Elenco_Reparti(Combo_Mittente.SelectedIndex), "Inoltra")
    End Sub

    Private Sub TableLayoutPanel13_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel13.Paint

    End Sub

    Private Sub Combo_Motivazione_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Motivazione.SelectedIndexChanged
        If Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 6 Then
            GroupBox22.Visible = True
        Else

            GroupBox22.Visible = False
            ComboBox5.SelectedIndex = -1
        End If
    End Sub

    Private Sub Combo_Destinatario_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Destinatario.SelectedIndexChanged
        '  Form_nuovo_ticket.riempi_combobox_causali(Combo_Motivazione, Elenco_Reparti(Combo_Mittente.SelectedIndex), Form_nuovo_ticket.Elenco_Reparti_destinatario(Combo_Destinatario.SelectedIndex), "Inoltra")
    End Sub


    Sub riempi_combobox_causali(par_combobox As ComboBox, par_causale As Integer)


        Dim Indice As Integer

        Dim Cnn_Motivi As New SqlConnection
        Cnn_Motivi.ConnectionString = Homepage.sap_tirelli
        Cnn_Motivi.Open()
        Dim Cmd_Motivi As New SqlCommand
        Dim Reader_Motivi As SqlDataReader

        Cmd_Motivi.Connection = Cnn_Motivi

        Cmd_Motivi.CommandText = "
SELECT t0.Id_motivo,[Descrizione_Motivo]
FROM [TIRELLI_40].[DBO].coll_motivazione t0


where t0.active='Y'

"



        Reader_Motivi = Cmd_Motivi.ExecuteReader()

        par_combobox.Items.Clear()

        Do While Reader_Motivi.Read()
            Elenco_Motivi(Indice) = Reader_Motivi("Id_Motivo")
            par_combobox.Items.Add(Reader_Motivi("Descrizione_Motivo"))
            If par_causale = Reader_Motivi("Id_Motivo") Then
                par_combobox.Text = Reader_Motivi("Descrizione_Motivo")
            End If
            Indice = Indice + 1
        Loop
        Num_Motivi = Indice
        Cnn_Motivi.Close()



    End Sub

    Private Sub TableLayoutPanel16_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel16.Paint

    End Sub

    Private Sub Cmd_Incolla_Click_1(sender As Object, e As EventArgs) Handles Cmd_Incolla.Click
        Picture_Campione.SizeMode = PictureBoxSizeMode.Zoom
        Picture_Campione.Image = Clipboard.GetImage
        If Picture_Campione.Image IsNot Nothing Then
            Immagine_Caricata = 1
        End If
    End Sub

    Private Sub Cmd_Cancella_Immagine_Click(sender As Object, e As EventArgs) Handles Cmd_Cancella_Immagine.Click
        Immagine_Caricata = 0
        Picture_Campione.Image = Nothing
    End Sub

    Private Sub Cmd_Zoom_Click_1(sender As Object, e As EventArgs) Handles Cmd_Zoom.Click
        Form_Zoom.Show()
        Form_Zoom.Picture_Zoom.Image = Picture_Campione.Image
        Form_Zoom.Owner = Me
        Me.Hide()
    End Sub

    Private Sub Cmd_Apri_Immagine_Click(sender As Object, e As EventArgs) Handles Cmd_Apri_Immagine.Click

    End Sub
End Class