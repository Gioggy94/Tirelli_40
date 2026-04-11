Imports Microsoft.Office.Interop
Imports System.Data.SqlClient
Imports Tirelli.ODP_Form

Public Class Form_Codici_vendita

    Public id As Integer
    Public immagine_Caricata As Integer = 0
    Public percorso_immagine As String
    Private Sub Form_codici_vendita_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Sales.Inserimento_sottogruppi_eti(ComboBox3)
    End Sub

    Sub inizializza_form(par_codice As String)
        TextBox1.Text = par_codice
        compila_anagrafica(par_codice)
        presenza_distinte(par_codice)
    End Sub

    Sub compila_anagrafica(par_codice As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT TOP (1000) [ID]
      ,[Codice]
      ,[Tipo_macchina]
      ,[Descrizione]
      ,coalesce([Costo],0) as 'Costo'
      ,[ultima_revisione]
      ,coalesce([Note],'') as 'Note'
      ,[Active]
      ,[Sottogruppo]
      ,[ADE]
      ,[STATIC]
      ,[HOT]
      ,[FLEX]
      ,[EU]
      ,[USA]
      ,coalesce([Costo_Materiale],0) as 'Costo_materiale'
      ,coalesce([Costificato],'N') as 'Costificato'
,coalesce(immagine,'') as 'Immagine'
  FROM [Tirelli_40].[dbo].[Superlistino_codici] where [Codice] ='" & par_codice & "'"
        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            id = cmd_SAP_reader("id")
            TextBox5.Text = cmd_SAP_reader("Tipo_macchina")

            TextBox1.Text = cmd_SAP_reader("Codice")
            TextBox2.Text = cmd_SAP_reader("Descrizione")
            TextBox3.Text = cmd_SAP_reader("Costo_Materiale")
            TextBox4.Text = cmd_SAP_reader("Costo")
            ComboBox1.Text = cmd_SAP_reader("active")
            DateTimePicker1.Value = cmd_SAP_reader("ultima_revisione")
            If cmd_SAP_reader("EU") = "Y" Then
                CheckBox5.Checked = True
            Else
                CheckBox5.Checked = False
            End If

            If cmd_SAP_reader("USA") = "Y" Then
                CheckBox6.Checked = True
            Else
                CheckBox6.Checked = False
            End If
            RichTextBox1.Text = cmd_SAP_reader("Note")

            ComboBox2.Text = cmd_SAP_reader("Costificato")
            ComboBox3.Text = cmd_SAP_reader("Sottogruppo")

            If cmd_SAP_reader("Tipo_macchina") = "Etichettatrice rotativa" Then


                If cmd_SAP_reader("Ade") = "Y" Then
                    CheckBox1.Checked = True
                Else
                    CheckBox1.Checked = False
                End If
                If cmd_SAP_reader("Static") = "Y" Then
                    CheckBox2.Checked = True
                Else
                    CheckBox2.Checked = False
                End If
                If cmd_SAP_reader("Hot") = "Y" Then
                    CheckBox3.Checked = True
                Else
                    CheckBox3.Checked = False
                End If
                If cmd_SAP_reader("Flex") = "Y" Then
                    CheckBox4.Checked = True
                Else
                    CheckBox4.Checked = False
                End If
            End If

            If cmd_SAP_reader("Immagine").ToString.Length > 0 Then
                    Picture_Campione.SizeMode = PictureBoxSizeMode.Zoom
                    Dim MyImage As Bitmap

                    Console.WriteLine(Homepage.Percorso_Immagini_TICKETS & cmd_SAP_reader("Immagine").ToString)

                    Try
                        MyImage = New Bitmap(Homepage.Percorso_Immagini_TICKETS & cmd_SAP_reader("Immagine").ToString)
                        percorso_immagine = Homepage.Percorso_Immagini_TICKETS & cmd_SAP_reader("Immagine").ToString
                    Catch ex As Exception
                    End Try
                    Picture_Campione.Image = CType(MyImage, Image)
                    immagine_Caricata = 1
                Else
                    Picture_Campione.Image = Nothing
                    immagine_Caricata = 0
                End If
            End If
            cmd_SAP_reader.Close()
        Cnn.Close()


    End Sub

    Public Function trova_immagine_codice(par_codice As String)

        Dim immagine As String = ""

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT 
coalesce(immagine,'') as 'Immagine'
  FROM [Tirelli_40].[dbo].[Superlistino_codici] where [Codice] ='" & par_codice & "'"
        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then



            immagine = cmd_SAP_reader("Immagine")
        Else
            immagine = ""

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()

        Return immagine

    End Function


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        aggiorna_anagrafica(TextBox1.Text)
        Sales.filtra_datagridview()
        MessageBox.Show("Anagrafica aggiornata con successo!", "Conferma", MessageBoxButtons.OK, MessageBoxIcon.Information)
        inizializza_form(TextBox1.Text)
    End Sub

    Sub aggiorna_anagrafica(par_codice As String)
        Dim Stringa_Immagine As String
        Dim immaginePath As String = Homepage.Percorso_Immagini_TICKETS & "Codice_vendita_" & TextBox1.Text & ".jpg"
        Dim versione As Integer = 1

        ' Controlla se esiste già un file con lo stesso nome
        While System.IO.File.Exists(immaginePath)
            ' Se il file esiste, incrementa la versione
            immaginePath = Homepage.Percorso_Immagini_TICKETS & "Codice_vendita_" & TextBox1.Text & "_" & versione.ToString() & ".jpg"
            versione += 1
        End While

        ' Ora salva l'immagine con il nuovo nome
        If immagine_Caricata = 1 Then
            Try
                ' Salva l'immagine
                Picture_Campione.Image.Save(immaginePath)

                ' Imposta la stringa immagine con il nuovo nome del file
                Stringa_Immagine = "Codice_vendita_" & TextBox1.Text & If(versione > 1, "_" & (versione - 1).ToString(), "") & ".jpg"
            Catch ex As System.Runtime.InteropServices.ExternalException
                MsgBox("Errore durante il salvataggio dell'immagine: " & ex.Message)
            End Try
        Else
            Stringa_Immagine = ""
        End If

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        ' 1. Verifica se esiste un id associato al Codice
        Dim checkCmd As New SqlCommand("SELECT [id] 
FROM [Tirelli_40].[dbo].[Superlistino_codici] WHERE [Codice] = @Codice", Cnn)
        checkCmd.Parameters.AddWithValue("@Codice", par_codice)

        Dim existingId As Object = checkCmd.ExecuteScalar()

        If existingId Is Nothing Then
            ' Gestisci l'errore se il Codice non esiste
            MsgBox("Codice non trovato. Ne verrà creato uno nuovo")
            existingId = Trova_ID_codice()
        End If

        ' 2. DELETE (rimuovi il record con il Codice specificato)
        Dim deleteCmd As New SqlCommand("DELETE FROM [Tirelli_40].[dbo].[Superlistino_codici] 
WHERE [Codice] = @Codice", Cnn)
        deleteCmd.Parameters.AddWithValue("@Codice", par_codice)
        deleteCmd.ExecuteNonQuery()

        ' 3. INSERT con l'id esistente
        Dim insertCmd As New SqlCommand("
        INSERT INTO [Tirelli_40].[dbo].[Superlistino_codici] (
            [id], [Codice], [Tipo_macchina], [Descrizione], [Costo], [ultima_revisione], [Note], [Active],
            [Sottogruppo], [ADE], [STATIC], [HOT], [FLEX], [EU], [USA], [Costo_Materiale], [Costificato], immagine
        ) VALUES (
            @id, @Codice, @Tipo_macchina, @Descrizione, @Costo, @UltimaRevisione, @Note, @Active,
            @Sottogruppo, @ADE, @STATIC, @HOT, @FLEX, @EU, @USA, @Costo_Materiale, @Costificato , @immagine
        )", Cnn)

        ' Aggiungi il parametro id preso dal SELECT
        insertCmd.Parameters.AddWithValue("@id", existingId)

        ' Parametri da controlli del form
        insertCmd.Parameters.AddWithValue("@Codice", TextBox1.Text)
        insertCmd.Parameters.AddWithValue("@Tipo_macchina", TextBox5.Text) ' Inserisci se hai il campo nel form
        insertCmd.Parameters.AddWithValue("@Descrizione", TextBox2.Text)
        insertCmd.Parameters.AddWithValue("@Costo", TextBox4.Text)
        insertCmd.Parameters.AddWithValue("@UltimaRevisione", DateTimePicker1.Value)
        insertCmd.Parameters.AddWithValue("@Note", RichTextBox1.Text)
        insertCmd.Parameters.AddWithValue("@Active", ComboBox1.Text) ' o 0, o altro valore se serve logica
        insertCmd.Parameters.AddWithValue("@Sottogruppo", ComboBox3.Text)
        insertCmd.Parameters.AddWithValue("@ADE", If(CheckBox1.Checked, "Y", "N"))
        insertCmd.Parameters.AddWithValue("@STATIC", If(CheckBox2.Checked, "Y", "N"))
        insertCmd.Parameters.AddWithValue("@HOT", If(CheckBox3.Checked, "Y", "N"))
        insertCmd.Parameters.AddWithValue("@FLEX", If(CheckBox4.Checked, "Y", "N"))
        insertCmd.Parameters.AddWithValue("@EU", If(CheckBox5.Checked, "Y", "N"))
        insertCmd.Parameters.AddWithValue("@USA", If(CheckBox6.Checked, "Y", "N"))
        insertCmd.Parameters.AddWithValue("@Costo_Materiale", TextBox3.Text)
        insertCmd.Parameters.AddWithValue("@Costificato", ComboBox2.Text)
        insertCmd.Parameters.AddWithValue("@immagine", Stringa_Immagine)

        insertCmd.ExecuteNonQuery()

        Cnn.Close()
    End Sub

    Public Function Trova_ID_codice()
        Dim id As Integer = 0
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "Select max(id)+1 As 'ID'
from  [Tirelli_40].[dbo].[Superlistino_codici]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                id = cmd_SAP_reader_2("ID")
            Else
                id = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
        Return id
    End Function

    Public Function Trova_codice_libero()
        Dim codice As String = ""
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT 
  'Y' + RIGHT('00000' + CAST(CAST(SUBSTRING(MAX(codice), 2, 5) AS INT) + 1 AS VARCHAR), 5) AS Codice
FROM 
  [Tirelli_40].[dbo].[Superlistino_codici]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("Codice") Is System.DBNull.Value Then
               codice= cmd_SAP_reader_2("Codice")
            Else
                Codice= 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
        Return Codice
    End Function
    Sub elimina_anagrafica(par_codice As String)

        Dim risposta As DialogResult
        risposta = MessageBox.Show("Vuoi eliminare il codice = " & par_codice & "?", "Conferma eliminazione", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If risposta = DialogResult.Yes Then
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()

            Dim deleteCmd As New SqlCommand("DELETE FROM [Tirelli_40].[dbo].[Superlistino_codici] WHERE [Codice] = @Codice", Cnn)
            deleteCmd.Parameters.AddWithValue("@Codice", par_codice)
            deleteCmd.ExecuteNonQuery()

            Dim deleteCmd_1 As New SqlCommand("DELETE  FROM [Tirelli_40].[dbo].[Superlistino_Distinte] WHERE [Figlio] = @Codice", Cnn)
            deleteCmd_1.Parameters.AddWithValue("@Codice", par_codice)
            deleteCmd_1.ExecuteNonQuery()

            Cnn.Close()

            MessageBox.Show("Codice eliminato con successo.", "Eliminazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Eliminazione annullata.", "Operazione annullata", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        elimina_anagrafica(TextBox1.Text)
    End Sub

    Sub presenza_distinte(par_codice_sap As String)
        Dim par_datagridview As DataGridView = DataGridView4
        par_datagridview.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "SELECT  t0.[Padre], coalesce(t1.Nome,'') as 'Nome'
      ,t0.[Figlio]
      ,t0.[Optional]
      ,t0.[N_figlio]
      ,t0.[Q]
      ,t0.[Note]
      ,t0.[Ultima_revisione]
  FROM [Tirelli_40].[dbo].[Superlistino_Distinte] t0 
left join [Tirelli_40].[dbo].[Modelli_macchine] t1 on t0.padre =t1.Codice
where [Figlio] ='" & par_codice_sap & "' order by [Padre],[N_figlio],[Optional]  "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader("Padre"), cmd_SAP_reader("Nome"), cmd_SAP_reader("q"), cmd_SAP_reader("Optional"))
        Loop

        cmd_SAP_reader.Close()
        Cnn.Close()

        Try
            par_datagridview.FirstDisplayedScrollingRowIndex = par_datagridview.RowCount
        Catch ex As Exception

        End Try


    End Sub 'Inserisco le risorse nella combo box

    Private Sub Button3_Click(sender As Object, e As EventArgs) 
        inizializza_form(TextBox1.Text)
    End Sub

    Private Sub Cmd_Zoom_Click(sender As Object, e As EventArgs) Handles Cmd_Zoom.Click
        Form_Zoom.Show()
        Form_Zoom.Picture_Zoom.Image = Picture_Campione.Image

        Me.Hide()
    End Sub

    Private Sub Cmd_Cancella_Immagine_Click(sender As Object, e As EventArgs) Handles Cmd_Cancella_Immagine.Click
        immagine_Caricata = 0
        Picture_Campione.Image = Nothing
    End Sub

    Private Sub Cmd_Incolla_Click(sender As Object, e As EventArgs) Handles Cmd_Incolla.Click
        Picture_Campione.SizeMode = PictureBoxSizeMode.Zoom
        Picture_Campione.Image = Clipboard.GetImage
        If Picture_Campione.Image IsNot Nothing Then
            immagine_Caricata = 1
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        If e.RowIndex >= 0 Then
            Dim par_datagridview As DataGridView = DataGridView4
            If e.ColumnIndex = par_datagridview.Columns.IndexOf(Padre) Then

                Form_distinta_vendita.Show()
                Form_distinta_vendita.inizializza_form(par_datagridview.Rows(e.RowIndex).Cells(columnName:="Padre").Value)
            End If
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim Par_Combobox As ComboBox = ComboBox2
        If Par_Combobox.Text = "Y" Then
            Par_Combobox.BackColor = Color.Lime
        ElseIf Par_Combobox.Text = "S" Then
            Par_Combobox.BackColor = Color.Yellow
        Else
            Par_Combobox.BackColor = Color.Red
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim Par_Combobox As ComboBox = ComboBox1
        If Par_Combobox.Text = "Y" Then
            Par_Combobox.BackColor = Color.Lime
        ElseIf Par_Combobox.Text = "S" Then
            Par_Combobox.BackColor = Color.Yellow
        Else
            Par_Combobox.BackColor = Color.Red
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        inizializza_form(TextBox1.Text)
    End Sub


    Public Function trova_dettagli_id_codice(par_id As String)
        Dim dettagli As New Dettagli_codice_vendita()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT TOP 1 [ID]
      ,[Codice]
      ,[Tipo_macchina]
      ,[Descrizione]
      ,[Costo]
      ,[ultima_revisione]
      ,[Note]
      ,[Active]
      ,[Sottogruppo]
      ,[ADE]
      ,[STATIC]
      ,[HOT]
      ,[FLEX]
      ,[EU]
      ,[USA]
      ,[Costo_Materiale]
      ,[Costificato]
      ,coalesce([immagine],'') as 'Immagine'
  FROM [Tirelli_40].[dbo].[Superlistino_codici]
where id=" & par_id & ""
        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            dettagli.id = cmd_SAP_reader("id")
            dettagli.codice = cmd_SAP_reader("codice")
            dettagli.Descrizione = cmd_SAP_reader("Descrizione")
            dettagli.immagine = cmd_SAP_reader("immagine")
        End If
        cmd_SAP_reader.Close()
        Cnn.Close()

        Return dettagli
    End Function

    Private Sub Cmd_Indietro_Click(sender As Object, e As EventArgs) Handles Cmd_Indietro.Click
        inizializza_form(trova_dettagli_id_codice(id - 1).codice)
    End Sub

    Private Sub Cmd_Avanti_Click(sender As Object, e As EventArgs) Handles Cmd_Avanti.Click
        inizializza_form(trova_dettagli_id_codice(id + 1).codice)
    End Sub

    Public Class Dettagli_codice_vendita
        Public id As Integer
        Public codice As String
        Public Descrizione As String
        Public immagine As String



    End Class

    Private Sub Picture_Campione_Click(sender As Object, e As EventArgs) Handles Picture_Campione.Click
        Process.Start(percorso_immagine)
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = Trova_codice_libero()
    End Sub
End Class