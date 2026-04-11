Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Drawing.Text
Imports System.Drawing.Printing
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form_Scheda_Collaudi
    Public Elenco_Tipo_Macchine(1000) As Integer
    Public Num_Tipo_Macchine As Integer
    Public Stringa_Connessione As String
    Public Stringa_Connessione_SAP As String
    Public Minuti_Impianto As Integer
    Public Minuti_Collaudo_Totale As Integer

    Public Codice_BP As Integer
    Public Elenco_BP(10000) As Integer
    Public Elenco_Campioni(10000) As Integer
    Private Elenco_Elementi(10) As Tab_Combinazione
    Public Elenco_Dipendenti(1000) As Integer
    Public Elenco_Combinazioni(1000) As Integer
    Public Num_Dipendenti As Integer
    Public Num_Elementi As Integer
    Public Num_Combinazioni As Integer = 0
    Public Codici_Campioni(100, 10) As Integer
    Public id_combinazione_selezionato As Integer
    Private num_collaudati As Integer

    Private Sub Form_Scheda_Collaudi_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo

        Dim Larghezza_Colonna_Immagine As Integer
        Dim Larghezza_Colonna_Testo As Integer
        Dim Larghezza_Colonna_Bottone As Integer

        Dim myFont As System.Drawing.Font
        myFont = New System.Drawing.Font("Webdings", 22) 'FontStyle.Bold Or FontStyle.Italic)

        'Larghezza_Colonna_Immagine = DataGrid_Combinazione.Width * 15 / 100
        'Larghezza_Colonna_Testo = DataGrid_Combinazione.Width * 7 / 100
        'Larghezza_Colonna_Bottone = DataGrid_Combinazione.Width * 5 / 100
        Dim Col_ID As New DataGridViewTextBoxColumn
        Col_ID.HeaderText = "ID"
        Col_ID.Width = Larghezza_Colonna_Testo / 2
        'DataGrid_Combinazione.Columns.Add(Col_ID)
        Dim Col_Vel As New DataGridViewTextBoxColumn
        Col_Vel.HeaderText = "Velocità"
        Col_Vel.Width = Larghezza_Colonna_Testo / 2
        'DataGrid_Combinazione.Columns.Add(Col_Vel)
        Dim Col_Bottone_Lavorazione As New DataGridViewButtonColumn
        Col_Bottone_Lavorazione.HeaderText = "Lavorazione"
        Col_Bottone_Lavorazione.Width = Larghezza_Colonna_Bottone
        Col_Bottone_Lavorazione.CellTemplate.Style.Font = myFont
        'DataGrid_Combinazione.Columns.Add(Col_Bottone_Lavorazione)
        Dim Col_Bottone_Modifica As New DataGridViewButtonColumn
        Col_Bottone_Modifica.HeaderText = "Aggiorna Dati"
        Col_Bottone_Modifica.Width = Larghezza_Colonna_Bottone
        Col_Bottone_Modifica.CellTemplate.Style.Font = myFont
        'DataGrid_Combinazione.Columns.Add(Col_Bottone_Modifica)
        Dim Col_Bottone_Video As New DataGridViewButtonColumn
        Col_Bottone_Video.HeaderText = "Carica Video"
        Col_Bottone_Video.Width = Larghezza_Colonna_Bottone
        Col_Bottone_Video.CellTemplate.Style.Font = myFont
        'DataGrid_Combinazione.Columns.Add(Col_Bottone_Video)
        Dim Col_Elem_Immagine(10) As DataGridViewImageColumn
        Dim Col_Elem_Testo(10) As DataGridViewTextBoxColumn
        Dim i As Integer
        i = 0
        For i = 0 To 9 Step 1
            Col_Elem_Immagine(i) = New DataGridViewImageColumn
            Col_Elem_Immagine(i).ImageLayout = DataGridViewImageCellLayout.Zoom
            Col_Elem_Immagine(i).HeaderText = "Elemento " & i + 1
            Col_Elem_Immagine(i).Width = Larghezza_Colonna_Immagine
            Col_Elem_Immagine(i).DefaultCellStyle.NullValue = Nothing
            'DataGrid_Combinazione.Columns.Add(Col_Elem_Immagine(i))
            Col_Elem_Testo(i) = New DataGridViewTextBoxColumn
            Col_Elem_Testo(i).Width = Larghezza_Colonna_Testo
            Col_Elem_Testo(i).HeaderText = "Elemento " & i + 1
            'DataGrid_Combinazione.Columns.Add(Col_Elem_Testo(i))
        Next
        'Aggiorna_Lista_Combinazioni()


    End Sub

    Sub inizializzazione_form(par_commessa As String)
        odp_macchina()
        Scheda_commessa_documentazione.commessa = par_commessa

        Lbl_Commessa.Text = par_commessa
        Lbl_Descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(par_commessa).Descrizione_commessa
        Lbl_Cod_Cliente.Text = Commesse_MES.SCHEDA_COMMESSA(par_commessa).codice_cliente
        Lbl_Cliente.Text = Commesse_MES.SCHEDA_COMMESSA(par_commessa).Cliente_commessa
        Lbl_Cod_Cliente_Finale.Text = Commesse_MES.SCHEDA_COMMESSA(par_commessa).codice_cliente_finale
        Lbl_Cliente_Finale.Text = Commesse_MES.SCHEDA_COMMESSA(par_commessa).Cliente_finale_commessa
        Lbl_Consegna.Text = Commesse_MES.SCHEDA_COMMESSA(par_commessa).Consegna_commessa
        Scheda_commessa_documentazione.compila_anagrafica(par_commessa)
        LinkLabel1.Text = Scheda_commessa_documentazione.cartella_macchina
        mostra_video()
        Scheda_tecnica.riempi_datagridview_combinazioni(DataGridView1, par_commessa, Homepage.sap_tirelli)
        avanzamento_collaudo(Scheda_tecnica.Ottieni_numero_combinazioni(par_commessa).Numero_combinazioni, Scheda_tecnica.Ottieni_numero_combinazioni(par_commessa).Numero_collaudati)

    End Sub

    Sub mostra_video()
        Try


            ' Verifica se rootDirectory esiste
            Dim rootDirectoryPath As String = Homepage.percorso_cartelle_macchine & Scheda_commessa_documentazione.cartella_macchina & "\video"

            ' Cancella tutti i nodi precedenti nel TreeView
            TreeView1.Nodes.Clear()

            ' Aggiunge il nodo radice al TreeView
            Dim rootDirectory As New DirectoryInfo(rootDirectoryPath)

            Dim rootNode As New TreeNode(rootDirectory.Name)
            rootNode.Tag = rootDirectory
            TreeView1.Nodes.Add(rootNode)

            ' Aggiunge tutti i nodi figli del nodo radice
            AddDirectories(rootNode)
            Addfiles(rootNode)
            ' Espande tutti i nodi del TreeView
            TreeView1.ExpandAll()

            ' Abilita il trascinamento dei file sulla TreeView
            TreeView1.AllowDrop = True
        Catch ex As Exception

        End Try
    End Sub

    Public Sub AddDirectories(parentNode As TreeNode)
        Try


            Dim parentDirectory As DirectoryInfo = TryCast(parentNode.Tag, DirectoryInfo)

            ' Aggiunge tutte le cartelle come nodi figli
            For Each directory As DirectoryInfo In parentDirectory.GetDirectories()
                Dim directoryNode As New TreeNode(directory.Name)
                directoryNode.Tag = directory

                '' Aggiunge l'icona della cartella alla ImageList
                'Dim folderIcon As Icon = Icon.ExtractAssociatedIcon("C:\Windows\System32\shell32.dll")
                'ImageList1.Images.Add("folder", folderIcon)

                '' Imposta l'icona della cartella sulla chiave "folder" nell'ImageList
                'directoryNode.ImageKey = "folder"

                If ImageList1.Images.Count > 0 AndAlso Not ImageList1.Images.ContainsKey("folder") Then
                    ' Imposta l'immagine della cartella sulla chiave "folder" nell'ImageList
                    directoryNode.ImageKey = "folder"
                End If

                parentNode.Nodes.Add(directoryNode)

                ' Aggiunge tutti i file come nodi figli della cartella
                For Each file As FileInfo In directory.GetFiles()
                    Dim fileNode As New TreeNode(file.Name & " (" & file.LastWriteTime.ToString() & ")")
                    fileNode.Tag = file

                    ' Ottiene l'icona del file
                    Dim fileIcon As Icon = SystemIcons.WinLogo
                    Dim filepath As String
                    filepath = file.FullName

                    filepath = Replace(filepath, "\\tirfs01\Tirelli", "T:")
                    Try
                        fileIcon = Icon.ExtractAssociatedIcon(filepath)
                    Catch ex As Exception

                    End Try


                    ' Aggiunge l'icona alla ImageList e imposta la proprietà ImageKey del nodo file
                    If Not ImageList1.Images.ContainsKey(file.Extension) Then
                        ImageList1.Images.Add(file.Extension, fileIcon)
                    End If
                    fileNode.ImageKey = file.Extension

                    directoryNode.Nodes.Add(fileNode)
                Next
                ' Ricorsivamente aggiunge tutti i nodi figli della cartella
                AddDirectories(directoryNode)

            Next
        Catch ex As Exception

        End Try

    End Sub

    Public Sub Addfiles(parentNode As TreeNode)
        Dim parentDirectory As DirectoryInfo = TryCast(parentNode.Tag, DirectoryInfo)
        Try
            For Each file As FileInfo In parentDirectory.GetFiles()
                Dim fileNode As New TreeNode(file.Name & " (" & file.LastWriteTime.ToString() & ")")
                fileNode.Tag = file

                ' Ottiene l'icona del file
                Dim fileIcon As Icon = SystemIcons.WinLogo
                Dim filepath As String
                filepath = file.FullName


                Try
                    fileIcon = Icon.ExtractAssociatedIcon(filepath)
                Catch ex As Exception

                End Try


                ' Aggiunge l'icona alla ImageList e imposta la proprietà ImageKey del nodo file
                If Not ImageList1.Images.ContainsKey(file.Extension) Then
                    ImageList1.Images.Add(file.Extension, fileIcon)
                End If
                fileNode.ImageKey = file.Extension

                parentNode.Nodes.Add(fileNode)
            Next
        Catch ex As Exception

        End Try
        ' Aggiunge anche i file direttamente presenti nella cartella padre




    End Sub

    Sub avanzamento_collaudo(par_num_combinazioni As Integer, par_num_collaudati As Integer)
        If par_num_combinazioni > 0 Then
            Barra_Collaudi.Value = par_num_collaudati * 100 / par_num_combinazioni
            Lbl_Progressione.Text = par_num_collaudati & " su " & par_num_combinazioni & " - " & Barra_Collaudi.Value & "%"
        Else
            Barra_Collaudi.Value = 0
            Lbl_Progressione.Text = "0%"
        End If

    End Sub
    Private Sub Cmd_Esci_Click(sender As Object, e As EventArgs) Handles Cmd_Esci.Click

        Me.Close()
    End Sub



    Private Function Get_Titolo(ID As Integer) As String

        Dim Cnn_Campione As New SqlConnection
        Dim Risultato As String

        Cnn_Campione.ConnectionString = Homepage.sap_tirelli
        Cnn_Campione.Open()

        Dim Cmd_Campione As New SqlCommand
        Dim Cmd_Campione_Reader As SqlDataReader

        Cmd_Campione.Connection = Cnn_Campione
        Cmd_Campione.CommandText = " SELECT COLL_Campioni.ID_Campione, COLL_Campioni.Nome, COLL_Campioni.Immagine, COLL_Tipo_Campione.Iniziale_Sigla, COLL_Tipo_Campione.Descrizione as 'Descrizione_Campione' 
FROM [TIRELLI_40].[DBO].coll_campioni, [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE WHERE COLL_Campioni.Tipo_Campione=COLL_Tipo_Campione.Id_Tipo_Campione AND Id_Campione=" & ID
        Cmd_Campione_Reader = Cmd_Campione.ExecuteReader
        If Cmd_Campione_Reader.Read() Then
            Risultato = Cmd_Campione_Reader("Iniziale_Sigla") & Cmd_Campione_Reader("Nome")
        End If

        Cnn_Campione.Close()
        Return Risultato
    End Function

    Private Function Get_Immagine(ID As Integer) As String
        Dim Cnn_Campione As New SqlConnection
        Dim Risultato As String

        Cnn_Campione.ConnectionString = Homepage.sap_tirelli
        Cnn_Campione.Open()

        Dim Cmd_Campione As New SqlCommand
        Dim Cmd_Campione_Reader As SqlDataReader

        Cmd_Campione.Connection = Cnn_Campione
        Cmd_Campione.CommandText = " SELECT * FROM [TIRELLI_40].[DBO].coll_campioni WHERE Id_Campione=" & ID
        Cmd_Campione_Reader = Cmd_Campione.ExecuteReader
        If Cmd_Campione_Reader.Read() Then
            Risultato = Cmd_Campione_Reader("Immagine")
        End If
        Cnn_Campione.Close()
        Return Risultato
    End Function



    Private Sub Cmd_Chiudi_Manodopera_Click(sender As Object, e As EventArgs)

        Form_Inserisci_Lavorazione.Show()
        Form_Inserisci_Lavorazione.Lbl_ODP.Text = "------------"
        Form_Inserisci_Lavorazione.Lbl_Combinazione.Text = "---"
        Form_Inserisci_Lavorazione.Chiusura = 1

    End Sub



    Private Sub Cmd_Manodopera_Generica_Click(sender As Object, e As EventArgs) Handles Cmd_Manodopera_Generica.Click

        Form_Inserisci_Lavorazione.Show()
        Form_Inserisci_Lavorazione.Lbl_Combinazione.Text = 0
        Form_Inserisci_Lavorazione.Lbl_ODP.Text = Button1.Text

    End Sub

    Private Sub Cmd_Tickets_Click(sender As Object, e As EventArgs)
        Pianificazione_Tickets.Show()


    End Sub








    Public Function Titolo_Video(Combinazione As Integer) As String
        Dim Cnn_Combinazioni As New SqlConnection
        Dim Titolo As String

        Cnn_Combinazioni.ConnectionString = Homepage.sap_tirelli
        Cnn_Combinazioni.Open()

        Dim Cmd_Combinazioni As New SqlCommand
        Dim Cmd_Combinazioni_Reader As SqlDataReader

        Cmd_Combinazioni.Connection = Cnn_Combinazioni
        Cmd_Combinazioni.CommandText = "SELECT * FROM [TIRELLI_40].[DBO].COLL_Combinazioni WHERE Id_Combinazione='" & Combinazione & "'"
        Cmd_Combinazioni_Reader = Cmd_Combinazioni.ExecuteReader

        Titolo = ""

        If Cmd_Combinazioni_Reader.Read() Then

            If (Cmd_Combinazioni_Reader("Campione_1")) > 0 Then
                Titolo = Titolo & "_" & Get_Titolo(Cmd_Combinazioni_Reader("Campione_1"))
            End If
            If (Cmd_Combinazioni_Reader("Campione_2")) > 0 Then
                Titolo = Titolo & "_" & Get_Titolo(Cmd_Combinazioni_Reader("Campione_2"))
            End If
            If (Cmd_Combinazioni_Reader("Campione_3")) > 0 Then
                Titolo = Titolo & "_" & Get_Titolo(Cmd_Combinazioni_Reader("Campione_3"))
            End If
            If (Cmd_Combinazioni_Reader("Campione_4")) > 0 Then
                Titolo = Titolo & "_" & Get_Titolo(Cmd_Combinazioni_Reader("Campione_4"))
            End If
            If (Cmd_Combinazioni_Reader("Campione_5")) > 0 Then
                Titolo = Titolo & "_" & Get_Titolo(Cmd_Combinazioni_Reader("Campione_5"))
            End If
            If (Cmd_Combinazioni_Reader("Campione_6")) > 0 Then
                Titolo = Titolo & "_" & Get_Titolo(Cmd_Combinazioni_Reader("Campione_6"))
            End If
            If (Cmd_Combinazioni_Reader("Campione_7")) > 0 Then
                Titolo = Titolo & "_" & Get_Titolo(Cmd_Combinazioni_Reader("Campione_7"))
            End If
            If (Cmd_Combinazioni_Reader("Campione_8")) > 0 Then
                Titolo = Titolo & "_" & Get_Titolo(Cmd_Combinazioni_Reader("Campione_8"))
            End If
            If (Cmd_Combinazioni_Reader("Campione_9")) > 0 Then
                Titolo = Titolo & "_" & Get_Titolo(Cmd_Combinazioni_Reader("Campione_9"))
            End If
            If (Cmd_Combinazioni_Reader("Campione_10")) > 0 Then
                Titolo = Titolo & "_" & Get_Titolo(Cmd_Combinazioni_Reader("Campione_10"))
            End If
        End If
        Cnn_Combinazioni.Close()
        Return Titolo
    End Function


    Sub odp_macchina()
        If Homepage.ERP_provenienza = "SAP" Then


            Dim Cnn_Combinazioni As New SqlConnection


            Cnn_Combinazioni.ConnectionString = Homepage.sap_tirelli
            Cnn_Combinazioni.Open()

            Dim Cmd_Combinazioni As New SqlCommand
            Dim Cmd_Combinazioni_Reader As SqlDataReader

            Cmd_Combinazioni.Connection = Cnn_Combinazioni
            Cmd_Combinazioni.CommandText = "SELECT max(t0.docnum) as 'Docnum' 
FROM OWOR T0 where t0.itemcode= '" & Lbl_Commessa.Text & "' and  t0.status <>'C'"
            Cmd_Combinazioni_Reader = Cmd_Combinazioni.ExecuteReader



            If Cmd_Combinazioni_Reader.Read() Then



                If Not Cmd_Combinazioni_Reader("docnum") Is System.DBNull.Value Then
                    Button1.Text = Cmd_Combinazioni_Reader("docnum")
                Else
                    Button1.Text = ""
                End If
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click



        ODP_Form.docnum_odp = Button1.Text
        ODP_Form.Show()
        ODP_Form.inizializza_form(Button1.Text)



    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start(Homepage.percorso_cartelle_macchine & LinkLabel1.Text)
    End Sub



    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Rows(e.RowIndex).Cells("Collaudo").Value = 1 Then
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
        End If
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            ID_combinazione_selezionato = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ID_combinazione").Value
            Label1.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Numero").Value
            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_1) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_2) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_3) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_4) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_5) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_6) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_7) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_8) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_9) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_10) Then


                Form_campione_visualizza.id_campione = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex - 2).Value
                Form_campione_visualizza.Show()
                Form_campione_visualizza.inizializza_form()

            End If




        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If id_combinazione_selezionato = Nothing Then
            MsgBox("Selezionare una combinazione")
        Else
            Modifica_Scheda_Combinazione.Show()
            Modifica_Scheda_Combinazione.Lbl_ID.Text = id_combinazione_selezionato
            Modifica_Scheda_Combinazione.Aggiorna_Scheda_Combinazione(id_combinazione_selezionato)
            If Lbl_Descrizione.Text.ToLower().Contains("riempitrice") OrElse
   Lbl_Descrizione.Text.ToLower().Contains("monoblocco") Then

                Modifica_Scheda_Combinazione.GroupBox3.Visible = True
            Else
                Modifica_Scheda_Combinazione.GroupBox3.Visible = False
            End If


        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If id_combinazione_selezionato = Nothing Then
            MsgBox("Selezionare una combinazione")
        Else
            Form_Carica_Video.Show()
            Form_Carica_Video.Commessa = Lbl_Commessa.Text
            Form_Carica_Video.Cartella = LinkLabel1.Text
            Form_Carica_Video.Combinazione = Titolo_Video(id_combinazione_selezionato)

            Form_Carica_Video.Aggiorna_Video()
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        mostra_video()
    End Sub



    Private Sub TreeView1_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick
        ' Verifica se il nodo selezionato è una directory
        If TypeOf e.Node.Tag Is FileInfo Then
            ' Apri il file con l'applicazione predefinita
            Dim file As FileInfo = DirectCast(e.Node.Tag, FileInfo)
            Process.Start(file.FullName)
        ElseIf TypeOf e.Node.Tag Is DirectoryInfo Then
            ' Se il nodo selezionato è una directory, apri la cartella
            Dim directory As DirectoryInfo = DirectCast(e.Node.Tag, DirectoryInfo)
            Process.Start("explorer.exe", directory.FullName)
        End If


    End Sub
End Class