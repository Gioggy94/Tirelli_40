Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class Form_Visualizza_Ticket
    Public id As Integer
    Public riga_datagridview As Integer = -1
    ' Public id_riferimento_canc As Integer
    ' Public Codice_canc As String
    ' Public documento_canc As String
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
    Public Elenco_dipendenti_assegnato(1000) As String
    Public codice_utente_assegnato As String
    Public TPR As String


    Private Sub Form_Visualizza_Ticket_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto = 6 Or Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto = 10 Then
            Button4.Visible = True

        Else
            Button4.Visible = False
        End If

        Me.CancelButton = Cmd_Annulla
        Calendario_Prevista.Visible = False

        Combo_Riferimenti.Items.Clear()
        Combo_Riferimenti.Items.Add("Articolo")
        Combo_Riferimenti.Items.Add("Ordine di Produzione")
        Combo_Riferimenti.Items.Add("Ordine cliente")
        Combo_Riferimenti.Items.Add("CDS")
        Form_nuovo_ticket.riempi_combobox_descrizioni_NC(ComboBox5)
        Combo_Riferimenti.SelectedIndex = 0
        Me.BackColor = Homepage.colore_sfondo



    End Sub

    Sub riempi_combo_motivazioni()
        Dim indice As Integer
        Dim Cnn_Motivi As New SqlConnection
        Cnn_Motivi.ConnectionString = Homepage.sap_tirelli
        Cnn_Motivi.Open()
        Dim Cmd_Motivi As New SqlCommand
        Dim Reader_Motivi As SqlDataReader
        Dim Indice_Motivo As Integer

        Cmd_Motivi.Connection = Cnn_Motivi
        Cmd_Motivi.CommandText = "SELECT *
        FROM [TIRELLI_40].[DBO].COLL_Motivazione 
        --where active='Y'
        ORDER BY Descrizione_Motivo"
        Reader_Motivi = Cmd_Motivi.ExecuteReader()
        Indice = 0
        Indice_Motivo = 0
        Combo_Motivazione.Items.Clear()

        Do While Reader_Motivi.Read()
            Elenco_Motivi(Indice) = Reader_Motivi("Id_Motivo")
            Combo_Motivazione.Items.Add(Reader_Motivi("Descrizione_Motivo"))
            If Num_Motivo = Elenco_Motivi(Indice) Then
                Indice_Motivo = Indice
            End If
            Indice = Indice + 1
        Loop
        Num_Motivi = Indice
        Cnn_Motivi.Close()
        If Indice_Motivo > 0 Then
            Combo_Motivazione.SelectedIndex = Indice_Motivo
        End If
    End Sub

    Public Function Ticket_assegnato_A(par_id_ticket As Integer)

        Dim assegnato As Integer = 0

        Dim Cnn_Motivi As New SqlConnection
        Cnn_Motivi.ConnectionString = Homepage.sap_tirelli
        Cnn_Motivi.Open()
        Dim Cmd_Motivi As New SqlCommand
        Dim Reader_Motivi As SqlDataReader


        Cmd_Motivi.Connection = Cnn_Motivi
        Cmd_Motivi.CommandText = "SELECT TOP (1000) [Id_Ticket]
      ,[Commessa]
      ,[Data_Creazione]
      ,[Data_Chiusura]
      ,[Data_Prevista_Chiusura]
      ,[Aperto]
      ,[Descrizione]
      ,[Mittente]
      ,[Destinatario]
      ,[Immagine]
      ,[Motivazione]
      ,[Id_Padre]
      ,[Business]
      ,[Utente]
      ,[Data_chiusura_totale]
      ,case when coalesce([Assegnato],0) ='' then 0 else coalesce([Assegnato],0) end  as 'Assegnato'
      ,[Riunione]
      ,[Oggetto]
      ,[TPR]
      ,[Tempo]
      ,[Costo]
      ,[Descrizione_NC]
      ,[Chiuditore]
  FROM [Tirelli_40].[dbo].[COLL_Tickets]
where id_ticket=" & par_id_ticket & ""

        Reader_Motivi = Cmd_Motivi.ExecuteReader()


        If Reader_Motivi.Read() Then
            assegnato = Reader_Motivi("Assegnato")

        End If

        Cnn_Motivi.Close()
        Return assegnato
    End Function


    Public Sub Startup()
        riempi_combo_motivazioni()
        Cmd_Inoltra.Enabled = False
        Stringa_Connessione_SAP = Homepage.sap_tirelli


        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = Homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader

        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "SELECT t0.MITTENTE, t0.DESTINATARIO, t0.MOTIVAZIONE, t0.COMMESSA, t0.DATA_CREAZIONE,t0.DATA_PREVISTA_cHIUSURA,t0.DATA_CHIUSURA, CASE WHEN t0.BUSINESS IS NULL THEN '' ELSE t0.BUSINESS END AS 'BUSINESS', t0.APERTO, t0.DESCRIZIONE, t0.ID_PADRE, t0.IMMAGINE,  t0.assegnato, case when t0.riunione is null then '' else t0.riunione end as 'Riunione', t0.oggetto
, case when t0.tpr='Y' then 'Y' else 'N' end as 'TPR'  
,coalesce(t0.tempo,0) as 'Tempo'
,coalesce(t0.costo,0) as 'Costo'
, coalesce(t2.descrizione_motivo,'') as 'Descrizione_motivo'
, coalesce(t3.descrizione_motivo,'') as 'Descrizione_motivo_ticket'
,t0.aperto
,coalesce(concat(t4.lastname,' ',t4.firstname),'') as 'Chiuditore'
,t0.data_chiusura_totale

FROM [TIRELLI_40].[DBO].coll_tickets t0 

left join [TIRELLI_40].[dbo].ohem t1 on t0.assegnato=t1.empid 
left join [Tirelli_40].[dbo].[COLL_Motivazione_Descrizione_NC] t2 on t2.id_motivo=t0.descrizione_nc
left join [Tirelli_40].[dbo].coll_motivazione t3 on t3.id_motivo=t0.motivazione
left join [TIRELLI_40].[dbo].ohem t4 on t0.chiuditore=t4.empid 

WHERE t0.Id_Ticket=" & Txt_Id.Text
        Reader_Ticket = Cmd_Ticket.ExecuteReader()
        Reader_Ticket.Read()
        Num_Mittente = Reader_Ticket("Mittente")
        Num_Destinatario = Reader_Ticket("Destinatario")
        Num_Motivo = Reader_Ticket("Motivazione")
        Txt_Commessa.Text = Reader_Ticket("Commessa")
        Data_Creazione = Reader_Ticket("Data_Creazione")
        Txt_Data_Creazione.Text = Data_Creazione.ToString("dd/MM/yyyy")
        Data_Prevista = Reader_Ticket("Data_Prevista_Chiusura")
        TextBox1.Text = Reader_Ticket("Business")
        ComboBox3.Text = Reader_Ticket("Riunione")
        Combo_Motivazione.Text = Reader_Ticket("Descrizione_motivo_ticket")
        If Reader_Ticket("aperto") = 0 Then
            Label2.Text = "Chiuso"
            GroupBox6.BackColor = Color.Red
        Else
            Label2.Text = "Aperto"
            GroupBox6.BackColor = Color.Lime
        End If

        If Reader_Ticket("Descrizione_motivo") <> "" Then
            ComboBox5.Text = Reader_Ticket("Descrizione_motivo")
        End If


        If Reader_Ticket("Tempo") > 0 Then
            TextBox3.Text = Reader_Ticket("Tempo")
        End If

        If Reader_Ticket("Costo") > 0 Then
            TextBox4.Text = Reader_Ticket("Costo")
        End If

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


        If Not DBNull.Value.Equals(Reader_Ticket("assegnato")) Then
            codice_utente_assegnato = Reader_Ticket("Assegnato")
        Else
            codice_utente_assegnato = Nothing
        End If

        Data_Chiusura = Reader_Ticket("Data_Chiusura")



        Ticket_Aperto = Reader_Ticket("Aperto")
        Txt_Descrizione.Text = Reader_Ticket("Descrizione")
        Txt_Id_Padre.Text = Reader_Ticket("Id_Padre")
        Label3.Text = Reader_Ticket("chiuditore")
        Try
            Label4.Text = Reader_Ticket("data_chiusura_totale")
        Catch ex As Exception

        End Try



        If Reader_Ticket("Immagine").ToString.Length > 0 Then
            Picture_Campione.SizeMode = PictureBoxSizeMode.Zoom
            Dim MyImage As Bitmap

            Console.WriteLine(Homepage.Percorso_Immagini_TICKETS & Reader_Ticket("Immagine").ToString)

            Try
                MyImage = New Bitmap(Homepage.Percorso_Immagini_TICKETS & Reader_Ticket("Immagine").ToString)
            Catch ex As Exception
            End Try
            Picture_Campione.Image = CType(MyImage, Image)
            Immagine_Caricata = 1
        End If
        Cnn_Ticket.Close()

        Dim Indice As Integer
        Dim Indice_Mittente As Integer
        Dim Indice_Destinatario As Integer

        Dim Cnn_Reparti As New SqlConnection

        Cnn_Reparti.ConnectionString = Homepage.sap_tirelli
        Cnn_Reparti.Open()

        Dim Cmd_Reparti As New SqlCommand
        Dim Reader_Reparti As SqlDataReader

        Cmd_Reparti.Connection = Cnn_Reparti
        Cmd_Reparti.CommandText = "SELECT * FROM [TIRELLI_40].[DBO].COLL_Reparti ORDER BY Descrizione"
        Reader_Reparti = Cmd_Reparti.ExecuteReader()
        Indice = 0
        Indice_Mittente = 0
        Indice_Destinatario = 0
        Combo_Mittente.Items.Clear()
        Combo_Destinatario.Items.Clear()

        Do While Reader_Reparti.Read()
            Elenco_Reparti(Indice) = Reader_Reparti("Id_Reparto")
            Combo_Mittente.Items.Add(Reader_Reparti("Descrizione"))
            Combo_Destinatario.Items.Add(Reader_Reparti("Descrizione"))
            'If Reader_Reparti("Fittizio") = 1 Then
            'Administrator = 1
            'End If
            If Num_Mittente = Elenco_Reparti(Indice) Then
                Indice_Mittente = Indice
            End If
            If Num_Destinatario = Elenco_Reparti(Indice) Then
                Indice_Destinatario = Indice
            End If
            Indice = Indice + 1
        Loop
        Num_Reparti = Indice
        Cnn_Reparti.Close()




        Combo_Mittente.SelectedIndex = Indice_Mittente
        Combo_Destinatario.SelectedIndex = Indice_Destinatario

        If Ticket_Aperto = 1 And (Administrator = 1 Or Reparto = Num_Destinatario) Then
            Cmd_Inoltra.Enabled = True

            Txt_Descrizione.Enabled = True

        Else



            Txt_Descrizione.Enabled = True

        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''
        Cnn_Ticket.Open()
        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "SELECT COLL_Tickets.Mittente, COLL_Reparti.Descrizione 
FROM [TIRELLI_40].[DBO].coll_tickets,[TIRELLI_40].[DBO].COLL_Reparti 
WHERE Id_Ticket=" & Txt_Id_Padre.Text & " AND Coll_Tickets.Mittente=COLL_Reparti.Id_Reparto"
        Reader_Ticket = Cmd_Ticket.ExecuteReader()
        Reader_Ticket.Read()
        Num_Mittente_Padre = Reader_Ticket("Mittente")
        Lbl_Mittente_Padre.Text = Reader_Ticket("Descrizione")
        Cnn_Ticket.Close()

        If Ticket_Aperto = 1 And (Administrator = 1 Or Reparto = Num_Mittente_Padre) Then
            Cmd_Chiudi.Enabled = True
        Else
            Cmd_Chiudi.Enabled = False
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''

        Txt_Data_Creazione.Enabled = False


        Combo_Destinatario.Enabled = False

        Txt_Id_Padre.Enabled = False
        Txt_Id.Enabled = False
        Combo_Mittente.Enabled = False
        'If Administrator = 1 Then
        '    Txt_Commessa.Enabled = True
        '    Cmd_Aggiorna_Ticket.Enabled = True
        'Else
        '    Txt_Commessa.Enabled = False
        'End If
        Compila_datagridview_Riferimenti()
        'Compila_Lista_Riferimenti()
        Inserimento_dipendenti_assegnato()
    End Sub



    Private Sub Compila_datagridview_Riferimenti()
        DataGridView.Rows.Clear()


        Dim Cnn_Riferimenti As New SqlConnection
        Cnn_Riferimenti.ConnectionString = homepage.sap_tirelli
        Cnn_Riferimenti.Open()
        Dim Cmd_Riferimenti As New SqlCommand
        Dim Reader_Riferimenti As SqlDataReader
        Cmd_Riferimenti.Connection = Cnn_Riferimenti

        Dim codice_sap As String = ""
        Dim numero_ODP As String = ""
        Dim descrizione As String = ""

        If Homepage.ERP_provenienza = "SAP" Then


            Cmd_Riferimenti.CommandText = "SELECT t0.Id_Riferimento,t0.rif_ticket,t0.tipo_codice,case when  t0.tipo_codice='Ordine di produzione' then '' else t1.itemcode end as 'Itemcode',case when t0.tipo_codice='Ordine di produzione' then t3.itemname when t0.tipo_codice='Articolo' then t1.itemname end as'Itemname' , case when t0.tipo_codice ='Ordine di produzione' then t0.codice_sap end as 'ODP'

FROM [TIRELLI_40].[DBO].COLL_Riferimenti t0 left join oitm t1 on t0.Codice_SAP=t1.itemcode 
left join owor t2 on cast(t2.docnum as varchar)=t0.codice_sap  and t0.tipo_codice ='Ordine di produzione'
left join oitm t3 on t3.itemcode=t2.itemcode
WHERE Rif_Ticket=" & Txt_Id.Text & ""
        Else
            Cmd_Riferimenti.CommandText = "SELECT TOP (1000) [Id_Riferimento]
      ,[Rif_Ticket]
      ,[Codice_SAP]
      ,[Tipo_Codice]
,'' as 'Itemname'
  FROM [Tirelli_40].[dbo].[COLL_Riferimenti]
where Rif_Ticket=" & Txt_Id.Text & ""

        End If
        Reader_Riferimenti = Cmd_Riferimenti.ExecuteReader()
        Do While Reader_Riferimenti.Read()
            If Reader_Riferimenti("Tipo_Codice") = "Ordine di produzione" Then
                codice_sap = ""
                numero_ODP = Reader_Riferimenti("Codice_sap")
                descrizione = Magazzino.OttieniDettagliAnagrafica(Reader_Riferimenti("Codice_sap")).Descrizione
            Else
                codice_sap = Reader_Riferimenti("Codice_sap")
                numero_ODP = ""
                descrizione = Magazzino.OttieniDettagliAnagrafica(Reader_Riferimenti("Codice_sap")).Descrizione
            End If
            DataGridView.Rows.Add(Reader_Riferimenti("Id_Riferimento"), Reader_Riferimenti("Tipo_Codice"), numero_ODP, codice_sap, Reader_Riferimenti("itemname"))

        Loop
        Cnn_Riferimenti.Close()
        DataGridView.ClearSelection()
        riga_datagridview = -1

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
            Cmd_Codice.CommandText = "SELECT T0.[ItemCode], T0.[ItemName], SUM(T1.[OnHand]) as 'Al Magazzino', SUM(T1.[IsCommited]) as 'Impegnato', SUM(T1.[OnOrder]) as 'In Ordine' 
                                          FROM [TirelliSRLDB].[dbo].OITM T0  INNER JOIN OITW T1 ON T0.[ItemCode] = T1.[ItemCode] 
                                          WHERE T0.[ItemCode] ='" & Codice & "' 
                                          GROUP BY T0.[ItemCode], T0.[ItemName]"
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
            Reader_Codice = Cmd_Codice.ExecuteReader()
            If Reader_Codice.Read() Then
                Risultato = Reader_Codice("Descrizione")
            End If
            Cnn_Codice.Close()
        End If
        Return Risultato
    End Function



    Private Sub Cmd_Zoom_Click(sender As Object, e As EventArgs) Handles Cmd_Zoom.Click
        Form_Zoom.Show()
        Form_Zoom.Picture_Zoom.Image = Picture_Campione.Image
        Form_Zoom.Owner = Me
        Me.Hide()
    End Sub

    Private Sub Cmd_Annulla_Click(sender As Object, e As EventArgs) Handles Cmd_Annulla.Click

        Me.Close()
    End Sub

    Private Sub Cmd_Chiudi_Click(sender As Object, e As EventArgs) Handles Cmd_Chiudi.Click

        If TextBox3.Text = "" And (Elenco_Reparti(Combo_Destinatario.SelectedIndex) = 1 Or Elenco_Reparti(Combo_Destinatario.SelectedIndex) = 17 Or Elenco_Reparti(Combo_Destinatario.SelectedIndex) = 25 Or Elenco_Reparti(Combo_Destinatario.SelectedIndex) = 26) And (Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 5 Or Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 6) Then


            MsgBox("Indicare quanto tempo è stato impiegato per risolvere questa fase del Ticket")

            Return
        Else
            TextBox3.Text = 0

        End If

        If TextBox4.Text = "" And (Elenco_Reparti(Combo_Destinatario.SelectedIndex) = 4 Or Elenco_Reparti(Combo_Destinatario.SelectedIndex) = 5) Then


            MsgBox("Indicare il costo del materiale che è stato impiegato per risolvere questa fase del Ticket")

            Return
        Else
            TextBox4.Text = 0
        End If

        Dim stringa_assegnato As String

        If ComboBox2.SelectedIndex < 0 Then
            stringa_assegnato = Nothing
        Else
            stringa_assegnato = ", assegnato='" & Elenco_dipendenti_assegnato(ComboBox2.SelectedIndex) & "'"
        End If


        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Txt_Descrizione.Text = Replace(Txt_Descrizione.Text, "'", "''")
        Dim Cmd_Ticket As New SqlCommand
        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "UPDATE [TIRELLI_40].[DBO].coll_tickets
                                  SET Data_Prevista_Chiusura='" & Data_Prevista.ToString("yyyy-MM-dd") & "',
                                  Commessa='" & Txt_Commessa.Text & "',
                                  riunione='" & ComboBox3.Text & "',
                                  Descrizione='" & Txt_Descrizione.Text & "'
                                  ,tempo ='" & TextBox3.Text & "'
                                  ,costo ='" & TextBox4.Text & "'
                                  , MOTIVAZIONE = '" & Elenco_Motivi(Combo_Motivazione.SelectedIndex) & "' " & stringa_assegnato & "
                                 ,chiuditore= " & Homepage.ID_SALVATO & "

                                  WHERE Id_Ticket='" & Txt_Id.Text & "'
                                
                                  UPDATE [TIRELLI_40].[DBO].coll_tickets
                                  SET  MOTIVAZIONE = '" & Elenco_Motivi(Combo_Motivazione.SelectedIndex) & "'
                                  ,chiuditore= " & Homepage.ID_SALVATO & "
                                  WHERE Id_Ticket='" & Txt_Id_Padre.Text & "'"
        Cmd_Ticket.ExecuteNonQuery()


        Cnn_Ticket.Close()



        Data_Chiusura = Today


        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()

        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "UPDATE [TIRELLI_40].[DBO].coll_tickets
                                  SET Aperto=0,Data_Chiusura=getdate() 
                                  WHERE Id_ticket='" & Txt_Id.Text & "'"

        Cmd_Ticket.ExecuteNonQuery()

        Cmd_Ticket.CommandText = "UPDATE [TIRELLI_40].[DBO].coll_tickets
                                  SET Aperto=0,Data_Chiusura_totale=getdate() 
                                  WHERE Id_padre='" & Txt_Id_Padre.Text & "'"

        Cmd_Ticket.ExecuteNonQuery()
        Cnn_Ticket.Close()
        Pianificazione_Tickets.Show()
        Pianificazione_Tickets.riempi_tickets(Pianificazione_Tickets.DataGridView1)
        MsgBox("Ticket chiuso con successo")
        Me.Close()
    End Sub

    Private Sub Cmd_Inoltra_Click(sender As Object, e As EventArgs) Handles Cmd_Inoltra.Click
        Form_Inoltra_Ticket.Show()
        Form_Inoltra_Ticket.Owner = Me
        Form_Inoltra_Ticket.Inserimento_dipendenti()
        Form_Inoltra_Ticket.Txt_Id.Text = Txt_Id.Text
        Form_Inoltra_Ticket.Startup()
        If Combo_Motivazione.SelectedIndex >= 0 Then
            Form_Inoltra_Ticket.riempi_combobox_causali(Form_Inoltra_Ticket.Combo_Motivazione, Elenco_Motivi(Combo_Motivazione.SelectedIndex))

        End If
        Me.Hide()
    End Sub

    Private Sub Cmd_Data_Prevista_Click(sender As Object, e As EventArgs)
        Calendario_Prevista.Visible = True
    End Sub



    Private Sub Cmd_Aggiorna_Ticket_Click(sender As Object, e As EventArgs) Handles Cmd_Aggiorna_Ticket.Click
        Dim invia_mail_ As Boolean = False
        If TextBox3.Text = "" Then
            TextBox3.Text = 0
        End If
        If TextBox4.Text = "" Then
            TextBox4.Text = 0
        End If
        Dim stringa_assegnato As String

        If ComboBox2.SelectedIndex < 0 Then
            stringa_assegnato = Nothing
        Else
            If Ticket_assegnato_A(Txt_Id.Text) <> Elenco_dipendenti_assegnato(ComboBox2.SelectedIndex) Then
                invia_mail_ = True
            End If

            stringa_assegnato = ", assegnato='" & Elenco_dipendenti_assegnato(ComboBox2.SelectedIndex) & "'"
        End If

        Txt_Descrizione.Text = Replace(Txt_Descrizione.Text, "'", "''")

        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Txt_Descrizione.Text = Replace(Txt_Descrizione.Text, "'", "''")
        Dim Cmd_Ticket As New SqlCommand
        Cmd_Ticket.Connection = Cnn_Ticket

        Dim par_descrizione_nc As Integer
        If ComboBox5.SelectedIndex = -1 Then
            par_descrizione_nc = 0
        Else
            par_descrizione_nc = Form_nuovo_ticket.Elenco_descrizione_nc(ComboBox5.SelectedIndex)
        End If

        Cmd_Ticket.CommandText = "UPDATE [TIRELLI_40].[DBO].coll_tickets
                                  SET Data_Prevista_Chiusura='" & Data_Prevista.ToString("yyyy-MM-dd") & "',
                                  Commessa='" & Txt_Commessa.Text & "',
                                  business='" & TextBox1.Text & "',
                                  oggetto='" & TextBox2.Text & "',
                                  riunione='" & ComboBox3.Text & "',
                                   tempo ='" & TextBox3.Text & "'
                                  ,costo ='" & TextBox4.Text & "'
                                  ,Descrizione='" & Txt_Descrizione.Text & "', MOTIVAZIONE = '" & Elenco_Motivi(Combo_Motivazione.SelectedIndex) & "' " & stringa_assegnato & "
                                  ,TPR='" & TPR & "'
                                  , descrizione_nc ='" & par_descrizione_nc & "'
                                  WHERE Id_Ticket='" & Txt_Id.Text & "'
                                  UPDATE [TIRELLI_40].[DBO].coll_tickets
                                  SET  MOTIVAZIONE = '" & Elenco_Motivi(Combo_Motivazione.SelectedIndex) & "'
                                  WHERE Id_Ticket='" & Txt_Id_Padre.Text & "'"
        Cmd_Ticket.ExecuteNonQuery()


        Cnn_Ticket.Close()
        If invia_mail_ = True Then
            Invia_Mail(Txt_Id.Text)
        End If
        Pianificazione_Tickets.Show()
        Pianificazione_Tickets.riempi_tickets(Pianificazione_Tickets.DataGridView1)
        Me.Close()
    End Sub

    Public Sub Invia_Mail(id As Integer)
        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = Homepage.sap_tirelli
        Cnn_Ticket.Open()

        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader

        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "SELECT t0.Data_Creazione, t0.Commessa, 
        COALESCE(t3.ItemName, '') AS 'Descrizione_commessa', 
    COALESCE(t3.U_Final_customer_name, '') AS 'Cliente', 
    T4.Descrizione AS 'Reparto_mittente', t0.Descrizione, 
    t2.Descrizione_Motivo, t1.Mail_1, t1.Mail_2, t1.Mail_3 
,coalesce(t5.email,'report@tirelli.net') as 'Email'

    FROM [TIRELLI_40].[DBO].coll_tickets t0 
    INNER JOIN [TIRELLI_40].[DBO].COLL_Reparti t1 ON t0.Destinatario = t1.Id_Reparto 
    INNER JOIN [TIRELLI_40].[DBO].COLL_Motivazione t2 ON t2.Id_Motivo = t0.Motivazione 
    LEFT JOIN [TirelliSRLDB].[dbo].OITM t3 ON t3.ItemCode = t0.Commessa 
    INNER JOIN [TIRELLI_40].[DBO].COLL_Reparti T4 ON T4.Id_Reparto = T0.Mittente 
left join [TIRELLI_40].[dbo].OHEM T5 ON T5.EMPID=T0.ASSEGNATO
    WHERE Id_Ticket = @IdTicket"

        Cmd_Ticket.Parameters.AddWithValue("@IdTicket", id)

        Reader_Ticket = Cmd_Ticket.ExecuteReader()
        If Reader_Ticket.Read() Then
            Dim Data_Creazione As Date = Reader_Ticket("Data_Creazione")

            ' Gestione delle interruzioni di riga nella descrizione
            Dim Descrizione As String = Reader_Ticket("Descrizione").ToString().Replace(vbCrLf, "<br>").Replace(vbCr, "<br>").Replace(vbLf, "<br>")

            Dim Testo_Mail As String = "<BODY>" &
        "<H3 style='color:#0056b3;'>📌 Nuovo Ticket N° " & id & "</H3>" &
        "<P>Ti è stato assegnato un nuovo ticket in riferimento alla commessa:<br>" &
        "<strong>🛠 Commessa:</strong> " & Reader_Ticket("Commessa") & " - " &
        Reader_Ticket("Descrizione_commessa") & " - " & Reader_Ticket("Cliente") & "</P>" &
        "<P><strong>📅 Data di Creazione:</strong> " & Data_Creazione.ToString("dd/MM/yyyy") & "<br>" &
        "<strong>📍 Mittente:</strong> " & Reader_Ticket("Reparto_mittente") & "<br>" &
        "<strong>📄 Descrizione:</strong> " & Descrizione & "<br>" &
        "<strong>📂 Tipologia:</strong> " & Reader_Ticket("Descrizione_Motivo") & "</P>"

            ' Aggiunta dell'elenco riferimenti, se presente
            'If ListBox_Riferimenti.Items.Count > 0 Then
            '    Testo_Mail &= "<P><strong>📌 Elenco dei Riferimenti:</strong><br><ul>"
            '    For Each item In ListBox_Riferimenti.Items
            '        Testo_Mail &= "<li>" & item.ToString() & "</li>"
            '    Next
            '    Testo_Mail &= "</ul></P>"
            'End If

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
                    myMail.To.Add(Reader_Ticket("Email"))
                    '  If Not String.IsNullOrEmpty(Reader_Ticket("Mail_2").ToString()) Then myMail.To.Add(Reader_Ticket("Mail_2"))
                    '  If Not String.IsNullOrEmpty(Reader_Ticket("Mail_3").ToString()) Then myMail.To.Add(Reader_Ticket("Mail_3"))
                    myMail.Bcc.Add("report@tirelli.net")

                    myMail.Subject = "Ticket assegnato per " & Reader_Ticket("Commessa") & " " & Reader_Ticket("Descrizione_commessa") & " " & Reader_Ticket("Cliente")
                    myMail.IsBodyHtml = True
                    myMail.Body = Testo_Mail

                    Try
                        mySmtp.Send(myMail)
                    Catch ex As Exception
                        ' MsgBox("Errore Invio Mail: " & ex.ToString)
                    End Try
                End Using
            End Using
        End If

        Reader_Ticket.Close()
        Cnn_Ticket.Close()
    End Sub



    'Private Sub Cmd_Mail_Click(sender As Object, e As EventArgs)
    '    Invia_Mail(Txt_Id.Text)
    'End Sub

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

    Public Sub Invia_Mail_old(id As Integer)
        Dim Cnn_Ticket As New SqlConnection

        Cnn_Ticket.ConnectionString = Homepage.sap_tirelli
        Cnn_Ticket.Open()

        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader

        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "SELECT  t0.[Id_Ticket]
      ,t0.[Commessa]
      ,t0.[Data_Creazione]
      ,t0.[Data_Chiusura]
      ,t0.[Data_Prevista_Chiusura]
      ,t0.[Aperto]
      ,t0.[Descrizione]
      ,t0.[Mittente]
      ,t0.[Destinatario]
      ,t0.[Immagine]
      ,t0.[Motivazione]
      ,t0.[Id_Padre]
      ,t0.[Business]
      ,t0.[Utente]
      ,t0.[Data_chiusura_totale]
      ,t0.[Assegnato]
      ,t0.[Riunione]
      ,t0.[Oggetto]
      ,t0.[TPR]
      ,t0.[Tempo]
      ,t0.[Costo]
      ,t0.[Descrizione_NC]
      ,t0.[Chiuditore]
	  ,coalesce(t1.email,'report@tirelli.net') as 'Email'
  FROM [Tirelli_40].[dbo].[COLL_Tickets] t0 
  left join [TIRELLI_40].[dbo].[ohem] t1 on t0.assegnato=t1.empid
 where t0.Id_Ticket=" & id
        Reader_Ticket = Cmd_Ticket.ExecuteReader()
        Reader_Ticket.Read()

        Dim Data_Creazione As Date
        Data_Creazione = Reader_Ticket("Data_Creazione")

        Dim Testo_Mail As String
        Testo_Mail = "<BODY><H3>Nuovo Ticket</h3><P>"
        Testo_Mail = Testo_Mail & "Ti è stato assegnato un nuovo ticket in riferimento alla commessa " & Reader_Ticket("Commessa")
        Testo_Mail = Testo_Mail & "<BR><BR>Data di Creazione : " & Data_Creazione.ToString("dd/MM/yyyy")
        Testo_Mail = Testo_Mail & "<BR>Mittente : " & Cerca_Reparto(Reader_Ticket("Mittente"))
        Testo_Mail = Testo_Mail & "<BR>Business : " & Reader_Ticket("Business")
        Testo_Mail = Testo_Mail & "<BR>Commessa : " & Reader_Ticket("Commessa")
        Testo_Mail = Testo_Mail & "<BR>Descrizione : " & Reader_Ticket("Descrizione")
        '-- '       Testo_Mail = Testo_Mail & "<BR>Tipologia : " & Reader_Ticket("Descrizione_Motivo")


        If DataGridView.Rows.Count > 0 Then
            Testo_Mail = Testo_Mail & "<BR><BR>Elenco dei Riferimenti :"
            Dim i As Integer
            For i = 0 To DataGridView.Rows.Count - 1
                Testo_Mail = Testo_Mail & "<BR> - ODP " & DataGridView.Rows(i).Cells(columnName:="ODP").Value & " Codice " & DataGridView.Rows(i).Cells(columnName:="Codice").Value & " Descrizione " & DataGridView.Rows(i).Cells(columnName:="DESC").Value
            Next
        End If
        Testo_Mail = Testo_Mail & "<BR><BR> Utilizzare l'applicazione <a href='" & Homepage.percorso_server & "\Tirelli 4.0.exe'>Tickets</a> per consultare l'elenco dei Tickets aperti e per poter inoltrare la risposta"
        Testo_Mail = Testo_Mail & "<BR>Questo è un messaggio automatico. Non rispondere a questa mail"


        Testo_Mail = Testo_Mail & "</P></BODY>"

        Dim mySmtp As New SmtpClient
        Dim myMail As New MailMessage()
        mySmtp.UseDefaultCredentials = False
        mySmtp.Credentials = New Net.NetworkCredential(Homepage.Mittente_Mail, Pianificazione_Tickets.Password_Mail)
        mySmtp.Host = "smtp.office365.com"
        mySmtp.Port = 587
        mySmtp.EnableSsl = True
        mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network


        myMail = New MailMessage()
        myMail.From = New MailAddress(Homepage.Mittente_Mail)
        myMail.To.Add(Reader_Ticket("eMail"))
        'If Reader_Ticket("Mail_2").ToString.Length > 1 Then
        '    myMail.To.Add(Reader_Ticket("Mail_2"))
        'End If
        myMail.Bcc.Add("report@tirelli.net")
        myMail.Subject = Reader_Ticket("Commessa") & " - Inserimento Nuovo Ticket per "
        myMail.IsBodyHtml = True
        myMail.Body = Testo_Mail

        Try
            mySmtp.Send(myMail)
        Catch ex As Exception
            MsgBox("Errore Invio Mail" & ex.ToString)
        End Try
        Cnn_Ticket.Close()
    End Sub

    Sub Inserimento_dipendenti()
        ComboBox1.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 
left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
inner join 
[TIRELLI_40].[DBO].coll_reparti t2 on (t2.sap_id_reparto =t1.code or t2.sap_id_reparto_2 =t1.code)  
where t0.active='Y' and t2.id_reparto='" & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto & "'  order by T0.[lastName] + ' ' + T0.[firstName] "

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

    Sub Inserimento_dipendenti_assegnato()

        ComboBox2.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 
left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code
inner join [TIRELLI_40].[DBO].coll_reparti t2 on t2.id_reparto=t0.u_reparto_tickets

where t0.active='Y' 
--and ((t2.id_reparto='" & Elenco_Reparti(Combo_Destinatario.SelectedIndex) & "' ) or (t2.id_reparto='1' or t2.id_reparto='17' or t2.id_reparto='25'))
order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti_assegnato(Indice) = cmd_SAP_reader("Codice dipendenti")
            ComboBox2.Items.Add(cmd_SAP_reader("Nome"))
            If codice_utente_assegnato = cmd_SAP_reader("Codice dipendenti").ToString Then
                ComboBox2.SelectedIndex = Indice
            End If
            Indice = Indice + 1
        Loop


        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub 'Inserisco le risorse nella combo box





    Private Sub DataGridView_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView.CellClick



        If e.RowIndex >= 0 Then
            riga_datagridview = e.RowIndex
            '= DataGridView.Rows(e.RowIndex).Cells(columnName:="id_riferimento").Value
            'Codice_canc = DataGridView.Rows(e.RowIndex).Cells(columnName:="Codice").Value
            'documento_canc = DataGridView.Rows(e.RowIndex).Cells(columnName:="Doc").Value

            Try
                If e.ColumnIndex = DataGridView.Columns.IndexOf(Codice) Then
                    Magazzino.Codice_SAP = DataGridView.Rows(e.RowIndex).Cells(columnName:="Codice").Value

                    Dim new_form_magazzino = New Magazzino
                    new_form_magazzino.Show()

                    new_form_magazzino.TextBox2.Text = Magazzino.Codice_SAP
                    new_form_magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)

                End If

            Catch ex As Exception

            End Try
            Try



                If e.ColumnIndex = DataGridView.Columns.IndexOf(ODP) Then





                    ODP_Form.docnum_odp = DataGridView.Rows(e.RowIndex).Cells(columnName:="ODP").Value
                    ODP_Form.Show()
                    ODP_Form.inizializza_form(DataGridView.Rows(e.RowIndex).Cells(columnName:="ODP").Value)


                End If

            Catch ex As Exception

            End Try




        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Txt_Nuovo_Riferimento.Text.Length > 0 Then
            If Combo_Riferimenti.SelectedIndex = 0 Then ' CODICE ARTICOLO

                Dim Cnn_Codice As New SqlConnection
                Cnn_Codice.ConnectionString = homepage.sap_tirelli
                Cnn_Codice.Open()
                Dim Cmd_Codice As New SqlCommand
                Dim Reader_Codice As SqlDataReader
                Cmd_Codice.Connection = Cnn_Codice
                Cmd_Codice.CommandText = "SELECT T0.[ItemCode], T0.[ItemName], SUM(T1.[OnHand]) as 'Al Magazzino', SUM(T1.[IsCommited]) as 'Impegnato', SUM(T1.[OnOrder]) as 'In Ordine' 
                                          FROM [TirelliSRLDB].[dbo].OITM T0  
INNER JOIN [TirelliSRLDB].[dbo].OITW T1 ON T0.[ItemCode] = T1.[ItemCode] 
                                          WHERE T0.[ItemCode] ='" & Txt_Nuovo_Riferimento.Text & "' 
                                          GROUP BY T0.[ItemCode], T0.[ItemName]"
                Reader_Codice = Cmd_Codice.ExecuteReader()

                If Reader_Codice.Read() Then
                    inserisci_riferimento()


                Else
                    MsgBox("Articolo Inesistente")
                    Txt_Nuovo_Riferimento.Text = ""
                End If
                Cnn_Codice.Close()
            End If
            If Combo_Riferimenti.SelectedIndex = 1 Then ' ORDINE DI PRODUZIONE
                Dim Cnn_Codice As New SqlConnection
                Cnn_Codice.ConnectionString = homepage.sap_tirelli
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
                                            WHERE cast(T0.[DocNum] as varchar) = '" & Txt_Nuovo_Riferimento.Text & "' 
                                            GROUP BY T0.[Status], T0.[DocNum], T0.[ItemCode], T0.[ProdName]"
                Reader_Codice = Cmd_Codice.ExecuteReader()
                If Reader_Codice.Read() Then
                    inserisci_riferimento()
                Else
                    MsgBox("Articolo Inesistente")
                    Txt_Nuovo_Riferimento.Text = ""
                End If
                Cnn_Codice.Close()
            End If
        End If
        Compila_datagridview_Riferimenti()
        riga_datagridview = -1
    End Sub

    Sub Trova_ID()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select case when max(id_riferimento)+1 is null then 1 else max(id_riferimento)+1 end as 'ID' 
from [TIRELLI_40].[DBO].coll_riferimenti"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                id = cmd_SAP_reader_2("ID")
            Else
                id = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        cnn1.Close()
    End Sub

    Sub inserisci_riferimento()
        Trova_ID()



        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Txt_Descrizione.Text = Replace(Txt_Descrizione.Text, "'", "''")
        Dim Cmd_Ticket As New SqlCommand
        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "insert into 
[TIRELLI_40].[DBO].coll_riferimenti (Id_Riferimento, Rif_Ticket,Codice_SAP,Tipo_Codice) values (" & id & ", '" & Txt_Id.Text & "','" & Txt_Nuovo_Riferimento.Text & "','" & Combo_Riferimenti.Text & "') "
        Cmd_Ticket.ExecuteNonQuery()


        Cnn_Ticket.Close()

    End Sub

    Sub elimina_riferimento()


        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Txt_Descrizione.Text = Replace(Txt_Descrizione.Text, "'", "''")
        Dim Cmd_Ticket As New SqlCommand
        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "delete [TIRELLI_40].[DBO].coll_riferimenti where Id_Riferimento =" & DataGridView.Rows(riga_datagridview).Cells(columnName:="Id_riferimento").Value & ""
        Cmd_Ticket.ExecuteNonQuery()


        Cnn_Ticket.Close()

    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If riga_datagridview = -1 Then
            MsgBox("Selezionare un riferimento")
        Else
            Dim Question
            Question = MsgBox("Sei sicuro di voler eliminare  " & DataGridView.Rows(riga_datagridview).Cells(columnName:="DOC").Value & " " & DataGridView.Rows(riga_datagridview).Cells(columnName:="Codice").Value & "  ?", vbYesNo)
            If Question = vbYes Then
                elimina_riferimento()
                Compila_datagridview_Riferimenti()

            End If
        End If


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

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
        If Combo_Motivazione.SelectedIndex >= 0 Then


            If (Elenco_Reparti(Combo_Mittente.SelectedIndex) = 1 Or Elenco_Reparti(Combo_Mittente.SelectedIndex) = 17 Or Elenco_Reparti(Combo_Mittente.SelectedIndex) = 25 Or Elenco_Reparti(Combo_Mittente.SelectedIndex) = 26) And (Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 5 Or Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 6) Then
                GroupBox20.Visible = True
            Else
                GroupBox20.Visible = False
            End If


        End If
    End Sub


    Private Sub Combo_Motivazione_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Motivazione.SelectedIndexChanged
        If Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 6 Then
            GroupBox22.Visible = True
        Else

            GroupBox22.Visible = False
            ComboBox5.SelectedIndex = -1
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If MessageBox.Show($"Sei sicuro di voler eliminare per sempre questo ticket? ATTENZIONE così eliminerai anche tutti i ticket incatenati a questo precedentemente ", "Elimina Ticket", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then


            elimina_ticket(Txt_Id_Padre.Text)
            MsgBox("Ticket eliminato con successo")
            Pianificazione_Tickets.riempi_tickets(Pianificazione.DataGridView1)




        End If
    End Sub

    Sub elimina_ticket(par_id_ticket As Integer)
        Dim Cnn_Campioni As New SqlConnection
        Cnn_Campioni.ConnectionString = Homepage.sap_tirelli
        Cnn_Campioni.Open()

        Dim Cmd_Campioni As New SqlCommand
        Cmd_Campioni.Connection = Cnn_Campioni
        Cmd_Campioni.CommandText = "delete [Tirelli_40].[DBO].COLL_tickets

WHERE Id_padre = " & par_id_ticket & ""

        Cmd_Campioni.ExecuteNonQuery()




        Cmd_Campioni.ExecuteNonQuery()


        Cnn_Campioni.Close()
    End Sub

    Private Sub Combo_Destinatario_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Destinatario.SelectedIndexChanged
        If Combo_Destinatario.SelectedIndex > 0 Then


            If Elenco_Reparti(Combo_Destinatario.SelectedIndex) <> 4 And Elenco_Reparti(Combo_Destinatario.SelectedIndex) <> 5 Then
                GroupBox21.Visible = False
            Else
                GroupBox21.Visible = True
            End If
        End If
    End Sub

    Private Sub GroupBox15_Enter(sender As Object, e As EventArgs) Handles GroupBox15.Enter

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
End Class