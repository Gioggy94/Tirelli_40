Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb


Public Class Form_nuovo_ticket

    Public Structure Riferimento
        Public Rif As String
        Public Descrizione As String
        Public Tipo As String
    End Structure

    Public Elenco_Reparti(1000) As Integer
    Public Elenco_Reparti_destinatario(1000) As Integer
    Public Elenco_descrizione_nc(1000) As Integer
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
    Public Elenco_dipendenti(1000) As String
    Public Elenco_dipendenti_assegnato(1000) As String
    Public TPR As String


    Private Sub Cmd_Annulla_Click(sender As Object, e As EventArgs)


        Me.Close()
    End Sub

    Private Sub Form_Nuovo_Ticket_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.BackColor = Homepage.colore_sfondo
        ApplicaStile()
        Inserimento_dipendenti()
        Inserimento_dipendenti_assegnato()

        Combo_Riferimenti.Items.Clear()
        Combo_Riferimenti.Items.Add("Articolo")
        Combo_Riferimenti.Items.Add("Ordine di Produzione")
        Combo_Riferimenti.Items.Add("Ordine cliente")
        Combo_Riferimenti.Items.Add("CDS")
        Combo_Riferimenti.SelectedIndex = 0
        Txt_Id.Text = ""

    End Sub

    Private Sub ApplicaStile()
        Dim navy As Color = Color.FromArgb(22, 45, 84)
        Dim navyHover As Color = Color.FromArgb(30, 63, 122)
        Dim navyDark As Color = Color.FromArgb(10, 26, 55)
        Dim fontUI As String = "Segoe UI"

        Me.Font = New Font(fontUI, 9)

        ' Pulsante Inserisci — navy bold
        Cmd_Inserisci.BackColor = navy
        Cmd_Inserisci.ForeColor = Color.White
        Cmd_Inserisci.FlatStyle = FlatStyle.Flat
        Cmd_Inserisci.FlatAppearance.BorderSize = 0
        Cmd_Inserisci.FlatAppearance.MouseOverBackColor = navyHover
        Cmd_Inserisci.FlatAppearance.MouseDownBackColor = navyDark
        Cmd_Inserisci.Font = New Font(fontUI, 14, FontStyle.Bold)

        ' Pulsante Annulla / Chiudi
        For Each btn As Button In New Button() {Button3}
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 1
            btn.Font = New Font(fontUI, 9)
        Next

        ' Altri pulsanti (zoom, incolla, aggiungi rif.)
        For Each btn As Button In New Button() {Cmd_Zoom, Cmd_Incolla, Cmd_Aggiungi_Riferimento}
            btn.FlatStyle = FlatStyle.Flat
            btn.FlatAppearance.BorderSize = 1
            btn.Font = New Font(fontUI, 9)
        Next

        ' GroupBox header color
        For Each gb As GroupBox In New GroupBox() {GroupBox1, GroupBox2, GroupBox3, GroupBox4,
                                                    GroupBox8, GroupBox9, GroupBox10,
                                                    grpProgetto, grpSottocommessa}
            gb.Font = New Font(fontUI, 8.5F, FontStyle.Bold)
            gb.ForeColor = navy
        Next

        ' Label avviso
        lblAvvisoMatricola.Font = New Font(fontUI, 9, FontStyle.Bold)
        lblAvvisoMatricola.ForeColor = Color.DarkRed

        ' Panel avviso
        pnlAvviso.BackColor = Color.FromArgb(255, 243, 205)

        ' Txt_Commessa — font bold evidenziato
        Txt_Commessa.Font = New Font(fontUI, 14, FontStyle.Bold)

        ' Combo motivazione — font grande leggibile
        Combo_Motivazione.Font = New Font(fontUI, 13)
    End Sub

    Public Sub Startup()



        riempi_combobox_causali(Combo_Motivazione, 0, 0, "Nuovo")
        riempi_combobox_descrizioni_NC(ComboBox5)
        riempi_combobox_mittente(Combo_Mittente)
        riempi_combobox_destinatario(Combo_Destinatario, 0, "Nuovo")



        Combo_Destinatario.SelectedItem = 0
        Data_Creazione = Today
        Txt_Data_Creazione.Enabled = False

        Txt_Data_Creazione.Text = Data_Creazione.ToString(“dd/MM/yyyy”)

        Txt_Id_Padre.Enabled = False
        Txt_Id.Enabled = False
        Num_Riferimenti = 0

    End Sub

    Sub riempi_combobox_causali(par_combobox As ComboBox, par_reparto_mittente As Integer, par_reparto_destinatario As Integer, par_tipo_ticket As String)

        If par_reparto_mittente = 0 Then
            par_combobox.Enabled = False
        Else
            par_combobox.Enabled = True
        End If



        Dim Indice As Integer

        Dim Cnn_Motivi As New SqlConnection
        Cnn_Motivi.ConnectionString = Homepage.sap_tirelli
        Cnn_Motivi.Open()
        Dim Cmd_Motivi As New SqlCommand
        Dim Reader_Motivi As SqlDataReader

        Cmd_Motivi.Connection = Cnn_Motivi
        If par_tipo_ticket = "Inoltra" Then
            Cmd_Motivi.CommandText = "select *
from
(
SELECT t0.Id_causale
FROM [TIRELLI_40].[DBO].COLL_Reparti_autorizzazioni t0
where t0.active='Y'
group by t0.Id_causale
)
 as t10 inner join [TIRELLI_40].[DBO].coll_motivazione t11 on t11.id_motivo=t10.Id_causale and t11.active='Y'
order by t11.Descrizione_Motivo
"
        Else
            Cmd_Motivi.CommandText = "select *
from
(
SELECT t0.Id_causale
FROM [TIRELLI_40].[DBO].COLL_Reparti_autorizzazioni t0

where t0.Id_Reparto_mittente=" & par_reparto_mittente & " and t0.Id_Reparto_destinatario=" & par_reparto_destinatario & " and t0.tipo_ticket='" & par_tipo_ticket & "'

group by t0.Id_causale
)
 as t10 inner join [TIRELLI_40].[DBO].coll_motivazione t11 on t11.id_motivo=t10.Id_causale and t11.active='Y'
order by t11.Descrizione_Motivo
"
        End If


        Reader_Motivi = Cmd_Motivi.ExecuteReader()
        Indice = 0
        par_combobox.Items.Clear()

        Do While Reader_Motivi.Read()
            Elenco_Motivi(Indice) = Reader_Motivi("Id_Motivo")
            par_combobox.Items.Add(Reader_Motivi("Descrizione_Motivo"))

            Indice = Indice + 1
        Loop
        Num_Motivi = Indice
        Cnn_Motivi.Close()
        par_combobox.SelectedIndex = -1


    End Sub

    Sub riempi_combobox_mittente(par_combobox As ComboBox)




        Dim Indice As Integer
        Dim Indice_Combo As Integer
        Dim Cnn_Reparti As New SqlConnection

        Cnn_Reparti.ConnectionString = Homepage.sap_tirelli
        Cnn_Reparti.Open()

        Dim Cmd_Reparti As New SqlCommand
        Dim Reader_Reparti As SqlDataReader

        Cmd_Reparti.Connection = Cnn_Reparti
        Cmd_Reparti.CommandText = "SELECT * 
FROM [TIRELLI_40].[DBO].COLL_Reparti
where 0=0 
ORDER BY Descrizione"
        Reader_Reparti = Cmd_Reparti.ExecuteReader()
        Indice = 0
        Indice_Combo = -1
        par_combobox.Items.Clear()


        Do While Reader_Reparti.Read()
            Elenco_Reparti(Indice) = Reader_Reparti("Id_Reparto")

            par_combobox.Items.Add(Reader_Reparti("Descrizione"))


            If Reparto = Elenco_Reparti(Indice) Then
                Indice_Combo = Indice
            End If
            Indice = Indice + 1
        Loop
        Num_Reparti = Indice
        Cnn_Reparti.Close()
    End Sub

    Sub riempi_combobox_descrizioni_NC(par_combobox As ComboBox)


        Dim Indice As Integer
        Dim Indice_Combo As Integer
        Dim Cnn_Reparti As New SqlConnection

        Cnn_Reparti.ConnectionString = Homepage.sap_tirelli
        Cnn_Reparti.Open()

        Dim Cmd_Reparti As New SqlCommand
        Dim Reader_Reparti As SqlDataReader

        Cmd_Reparti.Connection = Cnn_Reparti
        Cmd_Reparti.CommandText = "SELECT  [Id_Motivo]
      ,[Descrizione_Motivo]
      ,[Active]
  FROM [Tirelli_40].[dbo].[COLL_Motivazione_Descrizione_NC]
where [Active]='Y'
order by [Descrizione_Motivo]
"
        Reader_Reparti = Cmd_Reparti.ExecuteReader()
        Indice = 0
        par_combobox.Items.Clear()


        Do While Reader_Reparti.Read()

            Elenco_descrizione_nc(Indice) = Reader_Reparti("Id_Motivo")

            par_combobox.Items.Add(Reader_Reparti("Descrizione_Motivo"))


            If Reparto = Elenco_Reparti(Indice) Then
                Indice_Combo = Indice
            End If
            Indice = Indice + 1
        Loop
        Num_Reparti = Indice
        Cnn_Reparti.Close()
    End Sub

    Sub riempi_combobox_destinatario(par_combobox As ComboBox, par_reparto_mittente As Integer, par_tipo_ticket As String)

        If par_reparto_mittente = 0 Then
            par_combobox.Enabled = False
        Else
            par_combobox.Enabled = True
        End If


        Dim Indice As Integer
        ' Dim Indice_Combo As Integer
        Dim Cnn_Reparti As New SqlConnection

        Cnn_Reparti.ConnectionString = Homepage.sap_tirelli
        Cnn_Reparti.Open()

        Dim Cmd_Reparti As New SqlCommand
        Dim Reader_Reparti As SqlDataReader

        Cmd_Reparti.Connection = Cnn_Reparti
        If par_reparto_mittente = 0 Then
            Cmd_Reparti.CommandText = "SELECT * 
FROM [TIRELLI_40].[DBO].COLL_Reparti

where 0=0 
ORDER BY Descrizione"
        Else
            Cmd_Reparti.CommandText = "select *
from
(
SELECT t0.Id_Reparto_destinatario
FROM [TIRELLI_40].[DBO].COLL_Reparti_autorizzazioni t0

where t0.Id_Reparto_mittente=" & par_reparto_mittente & " and t0.tipo_ticket='" & par_tipo_ticket & "'

group by t0.Id_Reparto_destinatario
)
 as t10 inner join [TIRELLI_40].[DBO].COLL_Reparti t11 on t11.Id_Reparto=t10.Id_Reparto_destinatario
order by t11.Descrizione

"
        End If
        Reader_Reparti = Cmd_Reparti.ExecuteReader()
        Indice = 0
        ' Indice_Combo = -1
        par_combobox.Items.Clear()


        Do While Reader_Reparti.Read()

            If par_reparto_mittente > 0 Then


                par_combobox.Items.Add(Reader_Reparti("Descrizione"))

                Elenco_Reparti_destinatario(Indice) = Reader_Reparti("Id_Reparto")
                ' If Reparto = Elenco_Reparti(Indice) Then
                ' Indice_Combo = Indice
                '  End If
                Indice = Indice + 1
            End If
        Loop




        Num_Reparti = Indice
        Cnn_Reparti.Close()







    End Sub





    Private Sub Cmd_Aggiungi_Riferimento_Click(sender As Object, e As EventArgs) Handles Cmd_Aggiungi_Riferimento.Click
        If Txt_Nuovo_Riferimento.Text.Length > 0 Then
            If Combo_Riferimenti.SelectedIndex = 0 Then ' CODICE ARTICOLO

                Dim Cnn_Codice As New SqlConnection
                Cnn_Codice.ConnectionString = homepage.sap_tirelli
                Cnn_Codice.Open()
                Dim Cmd_Codice As New SqlCommand
                Dim Reader_Codice As SqlDataReader
                Cmd_Codice.Connection = Cnn_Codice
                If Homepage.ERP_provenienza = "SAP" Then
                    Cmd_Codice.CommandText = "SELECT T0.[ItemCode], T0.[ItemName], SUM(T1.[OnHand]) as 'Al Magazzino', SUM(T1.[IsCommited]) as 'Impegnato', SUM(T1.[OnOrder]) as 'In Ordine' 
                                          FROM [TIRELLISRLDB].[DBO].OITM T0  INNER JOIN OITW T1 ON T0.[ItemCode] = T1.[ItemCode] 
                                          WHERE T0.[ItemCode] ='" & Txt_Nuovo_Riferimento.Text & "' 
                                          GROUP BY T0.[ItemCode], T0.[ItemName]"
                Else
                    Cmd_Codice.CommandText = "SELECT   

CODE AS itemCODE,

DES_CODE AS itemname,

CHECK_DB AS 'Code'


FROM OPENQUERY(AS400, '
    SELECT *
    FROM S786FAD1.TIR90VIS.JGALART
    WHERE code =  ''" & Txt_Nuovo_Riferimento.Text & "''
') T10 "
                End If

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
                If Homepage.ERP_provenienza = "SAP" Then
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
                Else
                    Cmd_Codice.CommandText = "											select 
											t10.numodp as 'Ordine'
											,t10.codart as 'Codice'
											,t10.dscodart_odp as 'Descrizione'
											,*
											FROM OPENQUERY(AS400, '
    SELECT *
    FROM TIR90VIS.JGALODP t0
	where numodp=''" & Txt_Nuovo_Riferimento.Text & "''

') T10"
                End If

                Reader_Codice = Cmd_Codice.ExecuteReader()
                If Reader_Codice.Read() Then
                    Elenco_Riferimenti(Num_Riferimenti).Rif = Reader_Codice("Ordine")
                    Elenco_Riferimenti(Num_Riferimenti).Descrizione = Reader_Codice("Codice") & " - " & Reader_Codice("Descrizione")
                    Elenco_Riferimenti(Num_Riferimenti).Tipo = "Ordine di produzione"
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

    Private Sub Compila_Lista_Riferimenti()
        ListBox_Riferimenti.Items.Clear()
        Dim i As Integer
        For i = 0 To Num_Riferimenti - 1 Step 1
            ListBox_Riferimenti.Items.Add(Elenco_Riferimenti(i).Tipo & " - " & Elenco_Riferimenti(i).Rif & " - " & Elenco_Riferimenti(i).Descrizione)
        Next
    End Sub

    Private Sub ListBox_Riferimenti_DoubleClick(sender As Object, e As EventArgs) Handles ListBox_Riferimenti.DoubleClick
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

    Private Sub Cmd_Apri_Immagine_Click(sender As Object, e As EventArgs) Handles Cmd_Apri_Immagine.Click
        'Dim openFileDialog1 As New OpenFileDialog()
        'openFileDialog1.InitialDirectory = "c:\"
        'openFileDialog1.Filter = "File Immagine|*.jpg"
        'openFileDialog1.FilterIndex = 1
        'openFileDialog1.RestoreDirectory = True
        'If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
        '    Try
        '        Picture_Campione.Image.Dispose()
        '    Catch ex As Exception

        '    End Try
        '    Picture_Campione.SizeMode = PictureBoxSizeMode.Zoom
        '    Dim MyImage As Bitmap
        '    Try
        '        MyImage = New Bitmap(openFileDialog1.FileName)
        '    Catch ex As Exception
        '        MsgBox("Impossibile Aprire l'Immagine Selezionata")
        '        Return
        '    End Try

        '    Picture_Campione.Image = CType(MyImage, Image)
        '    Immagine_Caricata = 1
        'End If
    End Sub

    Private Sub Cmd_Incolla_Click(sender As Object, e As EventArgs) Handles Cmd_Incolla.Click
        Picture_Campione.SizeMode = PictureBoxSizeMode.Zoom
        Picture_Campione.Image = Clipboard.GetImage
        If Picture_Campione.Image IsNot Nothing Then
            Immagine_Caricata = 1
        End If
    End Sub

    Private Sub Cmd_Zoom_Click(sender As Object, e As EventArgs) Handles Cmd_Zoom.Click
        Form_Zoom.Show()
        Form_Zoom.Picture_Zoom.Image = Picture_Campione.Image
        Form_Zoom.Owner = Me
        Me.Hide()
    End Sub



    Private Sub Cmd_Inserisci_Click(sender As Object, e As EventArgs) Handles Cmd_Inserisci.Click

        'If Elenco_Motivi(Combo_Motivazione.SelectedIndex) = Nothing Or Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 0 Or Elenco_Motivi(Combo_Motivazione.SelectedIndex) = "" Then
        '    MsgBox("Selezionare una motivazione valida")
        '    Return
        'End If


        If GroupBox5.Visible = True And ComboBox5.SelectedIndex < 0 Then
            MsgBox("Selezionare una descrizione NC")
            Return
        End If

        If Combo_Mittente.SelectedIndex < 0 Then

            MsgBox("Selezionare un reparto mittente")


        Else
            If ComboBox1.SelectedIndex < 0 Then

                MsgBox("Selezionare un utente mittente")


            Else


                If Combo_Destinatario.SelectedIndex < 0 Then

                    MsgBox("Selezionare un Destinatario")


                Else
                    If Combo_Motivazione.SelectedIndex < 0 Then


                        MsgBox("Selezionare una motivazione")

                    Else


                        If Txt_Descrizione.Text.Length < 1 Then



                            MsgBox("Aggiungere una Descrizione")


                        Else
                            If Txt_Commessa.Text.Length < 6 Then

                                MsgBox("Indicare la commessa con almeno 6 caratteri ")

                            Else

                                ' ── Controllo matricola / progetto+sottocommessa ──
                                Dim matricolaValida As Boolean = VerificaMatricola()
                                If Not matricolaValida Then
                                    ' Matricola non trovata: verifica che progetto E sottocommessa siano compilati
                                    If txtProgetto.Text.Trim() = "" OrElse txtSottocommessa.Text.Trim() = "" Then
                                        MsgBox("La matricola non è stata trovata in AS400." & vbCrLf &
                                               "Compilare i campi Progetto e Sottocommessa, o indicare una matricola corretta.",
                                               MsgBoxStyle.Exclamation, "Verifica commessa")
                                        txtProgetto.Focus()
                                        GoTo FineInserisci
                                    End If
                                End If
                                ' ─────────────────────────────────────────────────

                                If TextBox1.Text = Nothing And ComboBox2.Text = "HELP_DESK" Then

                                    MsgBox("Indicare un oggetto  ")

                                Else
                                    Nuovo_Ticket()

                                    MsgBox("Ticket inviato")

                                    Txt_Descrizione.Text = ""
                                End If







                            End If
                        End If
                    End If
                End If
            End If
        End If

FineInserisci:
    End Sub


    Private Sub Nuovo_Ticket()
        'Inserimento Ticket
        Txt_Id.Text = Nuovo_ID()
        Txt_Id_Padre.Text = Txt_Id.Text
        Dim Stringa_Immagine As String
        If Immagine_Caricata = 1 Then
            Stringa_Immagine = "Ticket_" & Txt_Id.Text & ".jpg"
            Picture_Campione.Image.Save(Homepage.Percorso_Immagini_TICKETS & Stringa_Immagine)
        Else
            Stringa_Immagine = ""
        End If

        Dim assegnatario As String

        assegnatario = Nothing

        If ComboBox3.SelectedIndex < 0 Then
            assegnatario = Nothing
        Else
            assegnatario = Elenco_dipendenti_assegnato(ComboBox3.SelectedIndex)
        End If


        Txt_Descrizione.Text = ComboBox1.Text & " " & Now & vbCrLf & Txt_Descrizione.Text
        Txt_Descrizione.Text = Replace(Txt_Descrizione.Text, "'", "''")

        Dim Cnn_Ticket As New SqlConnection
        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Cmd_Ticket.Connection = Cnn_Ticket
        Dim par_descrizione_nc As Integer
        If ComboBox5.SelectedIndex = -1 Then
            par_descrizione_nc = 0
        Else
            par_descrizione_nc = Elenco_descrizione_nc(ComboBox5.SelectedIndex)
        End If


        Cmd_Ticket.CommandText = "INSERT INTO [TIRELLI_40].[DBO].coll_tickets
                                                (Id_Ticket,Commessa,Data_Creazione,Data_Chiusura,Data_Prevista_Chiusura,
                                                Aperto,Descrizione,Mittente,Destinatario,Immagine,Motivazione,Id_Padre, BUSINESS,utente,Assegnato, riunione, oggetto,tpr,descrizione_nc)
                                                VALUES(" & Txt_Id.Text & "
                                                , '" & Txt_Commessa.Text.ToUpper & "'
                                                , '" & Data_Creazione.ToString("yyyy-MM-dd") & "'
                                                , '" & Data_Prevista.ToString("yyyy-MM-dd") & "'
                                                , '" & Data_Chiusura.ToString("yyyy-MM-dd") & "'
                                                , 1
                                                , '" & Txt_Descrizione.Text & "'
                                                , " & Elenco_Reparti(Combo_Mittente.SelectedIndex) & "
                                                , " & Elenco_Reparti_destinatario(Combo_Destinatario.SelectedIndex) & "
                                                , '" & Stringa_Immagine & "'
                                                , " & Elenco_Motivi(Combo_Motivazione.SelectedIndex) & "
                                                , " & Txt_Id_Padre.Text & ", '" & ComboBox2.Text & "', '" & Elenco_dipendenti(ComboBox1.SelectedIndex) & "'
                                                ,'" & assegnatario & "'
                                                ,'" & ComboBox4.Text & "'
                                                ,'" & TextBox1.Text & "'
,'" & TPR & "'
," & par_descrizione_nc & "
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
                Cmd_Riferimenti.CommandText = "INSERT INTO [TIRELLI_40].[DBO].COLL_RIFERIMENTI
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


        'MsgBox("Ticket Inserito Con Successo")
        Invia_Mail(Txt_Id.Text)

        'Me.Close()
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

        Try
            Cnn_Ticket.ConnectionString = Homepage.sap_tirelli
            Cnn_Ticket.Open()

            Dim Cmd_Ticket As New SqlCommand
            Dim Reader_Ticket As SqlDataReader

            Cmd_Ticket.Connection = Cnn_Ticket
            Cmd_Ticket.CommandText = "SELECT t0.Data_Creazione,t0.Commessa,coalesce(t0.IMMAGINE,'') as 'Immagine',

case when t3.ItemName is null then '' else t3.itemname end as 'Descrizione_commessa' " &
        ", case when t3.U_Final_customer_name is null then '' else t3.u_final_customer_name end as 'Cliente' 
        , T4.Descrizione AS 'Reparto_mittente',t0.descrizione,t2.Descrizione_Motivo,t1.Mail_1,t1.Mail_2,t1.Mail_3 " &
        "FROM [TIRELLI_40].[DBO].coll_tickets t0 " &
        "INNER JOIN [TIRELLI_40].[DBO].COLL_Reparti t1 ON t0.Destinatario=t1.Id_Reparto " &
        "INNER JOIN [TIRELLI_40].[DBO].COLL_motivazione t2 ON t2.Id_Motivo=t0.motivazione " &
        "LEFT JOIN [TIRELLISRLDB].[DBO].oitm t3 ON t3.itemcode=t0.Commessa " &
        "INNER JOIN [TIRELLI_40].[DBO].COLL_Reparti T4 ON T4.Id_Reparto=T0.Mittente " &
        "WHERE Id_Ticket=@id"

            Cmd_Ticket.Parameters.AddWithValue("@id", id)

            Reader_Ticket = Cmd_Ticket.ExecuteReader()
            If Reader_Ticket.Read() Then
                Dim Data_Creazione As Date = Reader_Ticket("Data_Creazione")
                Dim Descrizione As String = Reader_Ticket("Descrizione").ToString().Replace(vbCrLf, "<br>").Replace(vbLf, "<br>")

                Dim Testo_Mail As String = "<BODY>" &
            "<H3 style='color:#0056b3;'>📌 Nuovo Ticket N° " & id & "</H3>" &
            "<P>Hai ricevuto un nuovo ticket in riferimento alla commessa: <br>" &
            "<strong>🛠 Commessa:</strong> " & Reader_Ticket("Commessa") & " - " & Reader_Ticket("Descrizione_commessa") & " - " & Reader_Ticket("Cliente") & "</P>" &
            "<P><strong>📅 Data di Creazione:</strong> " & Data_Creazione.ToString("dd/MM/yyyy") & "<br>" &
            "<strong>📍 Mittente:</strong> " & Reader_Ticket("Reparto_mittente") & "<br>" &
     "<strong>📄 Descrizione:</strong> " & Descrizione & "<br>" &
            "<strong>📂 Tipologia:</strong> " & Reader_Ticket("Descrizione_Motivo") & "</P>"

                If ListBox_Riferimenti.Items.Count > 0 Then
                    Testo_Mail &= "<P><strong>📌 Elenco dei Riferimenti:</strong><br><ul>"
                    For Each item In ListBox_Riferimenti.Items
                        Testo_Mail &= "<li>" & item.ToString() & "</li>"
                    Next
                    Testo_Mail &= "</ul></P>"
                End If

                Testo_Mail &= "<P>🔍 Per maggiori dettagli, utilizzare l'applicazione <strong style='color:blue;'>Tirelli 4.0</strong> per controllare l'elenco dei ticket aperti e inoltrare o chiudere la risposta.</P>" &
            "<P style='color:red;'><em>⚠️ Questo è un messaggio automatico. Non rispondere a questa mail.</em></P>" &
            "</BODY>"

                Dim mySmtp As New SmtpClient("tirelli-net.mail.protection.outlook.com")
                mySmtp.UseDefaultCredentials = False
                mySmtp.Credentials = New Net.NetworkCredential(Homepage.Mittente_Mail, Pianificazione_Tickets.Password_Mail)
                mySmtp.Port = 25
                mySmtp.EnableSsl = True
                mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network

                Dim myMail As New MailMessage()
                myMail.From = New MailAddress(Homepage.Mittente_Mail)
                myMail.To.Add(Reader_Ticket("Mail_1"))
                If Not String.IsNullOrEmpty(Reader_Ticket("Mail_2").ToString()) Then
                    myMail.To.Add(Reader_Ticket("Mail_2"))
                End If
                If Not String.IsNullOrEmpty(Reader_Ticket("Mail_3").ToString()) Then
                    myMail.To.Add(Reader_Ticket("Mail_3"))
                End If
                myMail.Bcc.Add("report@tirelli.net")
                myMail.Subject = "Nuovo ticket per " & Reader_Ticket("Commessa") & " " & Reader_Ticket("Descrizione_commessa") & " " & Reader_Ticket("Cliente")
                myMail.IsBodyHtml = True
                myMail.Body = Testo_Mail

                Try
                    mySmtp.Send(myMail)
                Catch ex As Exception
                    MsgBox("Errore Invio Mail: " & ex.Message)
                End Try
            End If
        Catch ex As Exception
            MsgBox("Errore nella connessione al database: " & ex.Message)
        Finally
            If Cnn_Ticket.State = ConnectionState.Open Then Cnn_Ticket.Close()
        End Try
    End Sub



    Private Sub Cmd_Cancella_Immagine_Click(sender As Object, e As EventArgs) Handles Cmd_Cancella_Immagine.Click
        Immagine_Caricata = 0
        Picture_Campione.Image = Nothing

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


    Private Function Nuovo_ID_Riferimento() As Integer
        Dim Cnn_Ticket As New SqlConnection
        Dim Risultato As Integer

        Cnn_Ticket.ConnectionString = homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader

        Cmd_Ticket.Connection = Cnn_Ticket
        Cmd_Ticket.CommandText = "SELECT MAX(Id_Riferimento) As 'Massimo' FROM [TIRELLI_40].[DBO].COLL_RIFERIMENTI"
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

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Close()
    End Sub

    Sub Inserimento_dipendenti()

        Dim reparto As Integer
        If Combo_Mittente.SelectedIndex = -1 Then
            reparto = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto
        Else
            reparto = Elenco_Reparti(Combo_Mittente.SelectedIndex)
        End If
        ComboBox1.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim Indice As Integer
        Indice = 0

        If reparto = 14 Then
            Elenco_dipendenti(Indice) = 78
            ComboBox1.Items.Add("Cattabriga Denis")
            Indice = Indice + 1
        End If

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 
left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
inner join [TIRELLI_40].[DBO].COLL_Reparti t2 on (t2.id_reparto =t0.u_reparto_tickets)   
where t0.active='Y' and t2.id_reparto='" & reparto & "'  
order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            ComboBox1.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Sub Inserimento_dipendenti_assegnato()
        Dim reparto As String
        If Combo_Destinatario.SelectedIndex = -1 Then
            reparto = Nothing
        Else
            reparto = Elenco_Reparti_destinatario(Combo_Destinatario.SelectedIndex)
        End If



        ComboBox3.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 
left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
inner join [TIRELLI_40].[DBO].COLL_Reparti t2 on (t2.sap_id_reparto =t1.code or t2.sap_id_reparto_2 =t1.code)   where t0.active='Y' 
and t2.id_reparto='" & reparto & "'  order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti_assegnato(Indice) = cmd_SAP_reader("Codice dipendenti")
            ComboBox3.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub 'Inserisco le risorse nella combo box



    Private Sub Combo_Destinatario_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Destinatario.SelectedIndexChanged
        Inserimento_dipendenti_assegnato()
        riempi_combobox_causali(Combo_Motivazione, Elenco_Reparti(Combo_Mittente.SelectedIndex), Elenco_Reparti_destinatario(Combo_Destinatario.SelectedIndex), "Nuovo")

    End Sub

    Private Sub Combo_Mittente_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Mittente.SelectedIndexChanged
        Inserimento_dipendenti()
        riempi_combobox_destinatario(Combo_Destinatario, Elenco_Reparti(Combo_Mittente.SelectedIndex), "Nuovo")
        Dim id_reparto_destinatario As Integer = 0
        If Combo_Destinatario.SelectedIndex < 0 Then
            id_reparto_destinatario = 0
        Else
            id_reparto_destinatario = Elenco_Reparti_destinatario(Combo_Destinatario.SelectedIndex)
        End If
        riempi_combobox_causali(Combo_Motivazione, Elenco_Reparti(Combo_Mittente.SelectedIndex), id_reparto_destinatario, "Nuovo")
    End Sub



    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Homepage.business = ComboBox2.Text
        Homepage.Aggiorna_INI_COMPUTER()
    End Sub



    Private Sub Txt_Commessa_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Txt_Commessa.KeyPress
        If e.KeyChar = " " Then
            e.Handled = True
        End If
    End Sub

    ' ─────────────────────────────────────────────────────────────────
    '  VERIFICA MATRICOLA — chiamata su Leave del campo Txt_Commessa
    ' ─────────────────────────────────────────────────────────────────

    Private Sub Txt_Commessa_Leave(sender As Object, e As EventArgs) Handles Txt_Commessa.Leave
        VerificaMatricola()
    End Sub

    ''' <summary>
    ''' Verifica che la matricola esista in JGALCOM (AS400).
    ''' Se non trovata mostra il pannello con i campi Progetto e Sottocommessa.
    ''' Restituisce True se la matricola è valida, False altrimenti.
    ''' </summary>
    Private Function VerificaMatricola() As Boolean
        Dim matricola As String = Txt_Commessa.Text.Trim().ToUpper()

        If matricola.Length < 6 Then
            pnlAvviso.Visible = False
            Txt_Commessa.BackColor = SystemColors.Window
            Return False
        End If

        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Using cmd As New SqlCommand()
                    cmd.Connection = cnn
                    cmd.CommandTimeout = 30
                    ' Stessa logica di Scheda_commessa_form_principale — JGALCOM su AS400
                    cmd.CommandText = String.Format(
                        "SELECT TOP 1 trim(t10.matricola) as matricola " &
                        "FROM OPENQUERY(AS400, 'SELECT matricola FROM TIR90VIS.JGALCOM " &
                        "WHERE trim(matricola) = ''{0}''') T10",
                        matricola.Replace("'", "''"))

                    Dim trovata As Boolean = False
                    Using rd As SqlDataReader = cmd.ExecuteReader()
                        trovata = rd.Read()
                    End Using

                    If trovata Then
                        pnlAvviso.Visible = False
                        Txt_Commessa.BackColor = Color.FromArgb(220, 255, 220)
                        Return True
                    Else
                        lblAvvisoMatricola.Text = "⚠  Matricola """ & matricola & """ non trovata in AS400 — compilare Progetto e Sottocommessa, o indica matricola corretta."
                        pnlAvviso.Visible = True
                        Txt_Commessa.BackColor = Color.FromArgb(255, 220, 180)
                        txtProgetto.Focus()
                        Return False
                    End If
                End Using
            End Using
        Catch ex As Exception
            MsgBox("Impossibile verificare la matricola su AS400: " & ex.Message, MsgBoxStyle.Exclamation)
            Return False
        End Try
    End Function

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TPR = "Y"
            Me.BackColor = Color.Wheat
        Else
            TPR = "N"
            Me.BackColor = Homepage.colore_sfondo
        End If
    End Sub



    Private Sub Combo_Motivazione_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Motivazione.SelectedIndexChanged
        If Elenco_Motivi(Combo_Motivazione.SelectedIndex) = 6 Then
            GroupBox5.Visible = True
        Else

            GroupBox5.Visible = False
            ComboBox5.SelectedIndex = -1
        End If
    End Sub
End Class