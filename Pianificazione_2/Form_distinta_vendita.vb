Imports System.Data.SqlClient
Imports System.Globalization
Imports System.IO
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Tirelli.ODP_Form

Public Class Form_distinta_vendita
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '' Mostra il messaggio di conferma
        'Dim risposta As DialogResult = MessageBox.Show("Per salvare assicurati di aver rimosso tutti i filtri premendo il tasto rimuovi filtri. Vuoi proseguire?",
        '                                        "Conferma Salvataggio", ' Titolo del MessageBox
        '                                        MessageBoxButtons.YesNo, ' Pulsanti (Yes/No)
        '                                        MessageBoxIcon.Question) ' Icona (domanda)

        'If risposta = DialogResult.Yes Then
        If Button11.Enabled = False Then


            '  Try
            ' Salvataggio anagrafica e distinte
            salva_anagrafica(TextBox1.Text)
            Salva_distinta(DataGridView4, TextBox1.Text, "N")
            Salva_distinta(DataGridView1, TextBox1.Text, "Y")

            ' Messaggio di successo
            MsgBox("Salvataggio effettuato con successo.", MsgBoxStyle.Information)
            '  Catch ex As Exception
            ' Gestione errori durante il salvataggio
            ' MsgBox("Si è verificato un errore durante il salvataggio: " & ex.Message, MsgBoxStyle.Critical)
            '  End Try
            ' End If
        Else
            MsgBox("Premere rimuovi filtri prima di salvare")
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        inizializza_form(TextBox1.Text)
    End Sub

    Sub inizializza_form(par_codice As String)
        Button10.Enabled = True
        TextBox1.Text = par_codice
        compila_anagrafica(par_codice)
        compila_distinta(DataGridView4, par_codice, "N", False, TextBox5.Text, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, Label3)
        compila_distinta(DataGridView1, par_codice, "Y", False, TextBox5.Text, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, Label3)
        Button11.Enabled = False
        'compila_distinta(par_codice)
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
      ,[Tipo]
      ,[Nome]
      ,[Serie]
      ,[Diametro_primitivo]
      ,[N_piattelli]
      ,COALESCE([N_baie],0) AS 'N_baie'
,COALESCE(ADE,'N') AS 'ADE'
,COALESCE(STATIC,'N') AS 'STATIC'
,COALESCE(HOT,'N') AS 'HOT'
,coalesce(made_in,'') as 'Produttore'
,coalesce(Note,'') as 'Note'

  FROM [Tirelli_40].[dbo].[Modelli_macchine]
 where [Codice] ='" & par_codice & "'"
        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            Label1.Text = cmd_SAP_reader("ID")
            TextBox1.Text = cmd_SAP_reader("Codice")
            TextBox2.Text = cmd_SAP_reader("nome")
            TextBox5.Text = cmd_SAP_reader("Tipo_macchina")
            TextBox6.Text = cmd_SAP_reader("Tipo")
            ComboBox1.Text = cmd_SAP_reader("Serie")
            TextBox3.Text = cmd_SAP_reader("Diametro_primitivo")
            TextBox4.Text = cmd_SAP_reader("N_piattelli")
            TextBox7.Text = cmd_SAP_reader("N_baie")
            TextBox8.Text = cmd_SAP_reader("Produttore")
            RichTextBox1.Text = cmd_SAP_reader("Note")
            If cmd_SAP_reader("ADE") = "Y" Then
                CheckBox7.Checked = True

            End If
            If cmd_SAP_reader("STATIC") = "Y" Then
                CheckBox8.Checked = True

            End If
            If cmd_SAP_reader("HOT") = "Y" Then
                CheckBox9.Checked = True

            End If


        End If
        cmd_SAP_reader.Close()
        Cnn.Close()


    End Sub

    Public Function trova_macchina(par_id As String)
        Dim codice As String = ""
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT TOP 1 
      coalesce([Codice],'') as 'Codice'
     

  FROM [Tirelli_40].[dbo].[Modelli_macchine]
 where [id] ='" & par_id & "'"
        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            codice = cmd_SAP_reader("codice")
        End If
        cmd_SAP_reader.Close()
        Cnn.Close()

        Return codice
    End Function

    Sub salva_anagrafica(par_codice As String)
        Dim ADE As String = "N"
        Dim STATIC_ As String = "N"
        Dim HOT As String = "N"
        If CheckBox7.Checked = True Then
            ADE = "Y"
        End If
        If CheckBox8.Checked = True Then
            STATIC_ = "Y"
        End If
        If CheckBox9.Checked = True Then
            HOT = "Y"
        End If


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Try
            Cnn.Open()

            ' Prima verifichiamo se esiste già un record con lo stesso Codice
            Dim CMD_Select As New SqlCommand("SELECT [ID] FROM [Tirelli_40].[dbo].[Modelli_macchine] WHERE [Codice] = @Codice", Cnn)
            CMD_Select.Parameters.AddWithValue("@Codice", par_codice)

            Dim existingId As Object = CMD_Select.ExecuteScalar()

            ' Se esiste, utilizziamo l'ID esistente, altrimenti creiamo un nuovo ID
            Dim recordId As Integer
            If existingId IsNot Nothing Then
                ' Se esiste un ID, usiamolo
                recordId = Convert.ToInt32(existingId)
            Else
                ' Se non esiste, dobbiamo ottenere il primo ID libero
                Dim CMD_MaxId As New SqlCommand("SELECT ISNULL(MAX([ID]), 0) + 1 FROM [Tirelli_40].[dbo].[Modelli_macchine]", Cnn)
                recordId = Convert.ToInt32(CMD_MaxId.ExecuteScalar())
            End If

            ' Prima cancelliamo il record esistente
            Dim CMD_Delete As New SqlCommand("DELETE FROM [Tirelli_40].[dbo].[Modelli_macchine] WHERE [Codice] = @Codice", Cnn)
            CMD_Delete.Parameters.AddWithValue("@Codice", par_codice)
            CMD_Delete.ExecuteNonQuery()

            ' Poi inseriamo il nuovo record
            Dim CMD_Insert As New SqlCommand("
        INSERT INTO [Tirelli_40].[dbo].[Modelli_macchine]
        ([ID], [Codice],[made_in], [Tipo_macchina], [Tipo], [Nome], [Serie], [Diametro_primitivo], [N_piattelli], [N_baie],ADE,STATIC,HOT, note)
        VALUES (@ID, @Codice, @made_in, @Tipo_macchina, @Tipo, @Nome, @Serie, @Diametro_primitivo, @N_piattelli, @N_baie,@ADE,@STATIC,@HOT, @note)", Cnn)

            Dim valore_note As String = ""
            If RichTextBox1.Text = "" Then
                valore_note = ""
            Else
                valore_note = Replace(RichTextBox1.Text, "'", " ")
                valore_note = Replace(valore_note, ",", ".")
            End If



            ' Aggiunta dei parametri
            CMD_Insert.Parameters.AddWithValue("@ID", recordId)
            CMD_Insert.Parameters.AddWithValue("@Codice", TextBox1.Text)
            CMD_Insert.Parameters.AddWithValue("@made_in", TextBox8.Text)
            CMD_Insert.Parameters.AddWithValue("@Nome", TextBox2.Text)
            CMD_Insert.Parameters.AddWithValue("@Tipo_macchina", TextBox5.Text)
            CMD_Insert.Parameters.AddWithValue("@Tipo", TextBox6.Text)
            CMD_Insert.Parameters.AddWithValue("@Serie", ComboBox1.Text)
            CMD_Insert.Parameters.AddWithValue("@Diametro_primitivo", Val(TextBox3.Text))
            CMD_Insert.Parameters.AddWithValue("@N_piattelli", Val(TextBox4.Text))
            CMD_Insert.Parameters.AddWithValue("@N_baie", Val(TextBox7.Text))
            CMD_Insert.Parameters.AddWithValue("@ADE", ADE)
            CMD_Insert.Parameters.AddWithValue("@STATIC", STATIC_)
            CMD_Insert.Parameters.AddWithValue("@HOT", HOT)
            CMD_Insert.Parameters.AddWithValue("@note", valore_note)


            CMD_Insert.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Errore durante il salvataggio: " & ex.Message)
        Finally
            Cnn.Close()
        End Try

    End Sub

    Sub Inserimento_serie_macchina(par_combobox As ComboBox)

        par_combobox.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT 
      [Serie]
      
  FROM [Tirelli_40].[dbo].[Modelli_macchine]
  group by [Serie]
  order by serie"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        par_combobox.Items.Add("")
        Do While cmd_SAP_reader.Read()

            par_combobox.Items.Add(cmd_SAP_reader("Serie"))

        Loop




        cmd_SAP_reader.Close()
        Cnn.Close()


    End Sub 'Inserisco le risorse nella combo box

    Sub compila_distinta(par_datagridview As DataGridView, par_codice As String, par_optional As String, par_filtro As Boolean, par_tipo_macchina As String, par_EU As Boolean, par_usa As Boolean, par_Ade As Boolean, par_static As Boolean, par_hot As Boolean, par_flex As Boolean, par_label_costo As Label)

        Dim filtro_eu, filtro_usa, filtro_ade, filtro_static, filtro_hot, filtro_flex As String
        If par_filtro = True Then
            ' Gestisci i filtri
            filtro_eu = If(par_EU, " and t1.eu='Y' ", "")
            filtro_usa = If(par_usa, " and t1.usa='Y' ", "")

            If par_tipo_macchina = "Etichettatrice rotativa" Then
                filtro_ade = If(par_Ade, " and t1.ade='Y' ", "")
                filtro_hot = If(par_Ade, " and t1.hot='Y' ", "")
                filtro_static = If(par_Ade, " and t1.static='Y' ", "")
            Else
                filtro_ade = ""
                filtro_hot = ""
                filtro_static = ""
            End If


        Else
            filtro_eu = ""
            filtro_usa = ""
            filtro_ade = ""
            filtro_hot = ""
            filtro_static = ""
        End If



        par_datagridview.Rows.Clear()

        ' Connessione al database
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT t0.[Padre], t1.id, t0.[Figlio], t1.Descrizione, t0.[Optional], t0.[N_figlio], t0.[Q], t1.Costo_Materiale, t1.Costo, t0.molt, t1.costificato, t0.[Q]*t0.molt*t1.Costo as 'Costo_Tot', t1.[Note], t0.[Ultima_revisione], coalesce(t1.immagine,'') as 'Immagine' 
                           ,[ADE]
      ,[STATIC]
      ,[HOT]
      ,[FLEX]
      ,[EU]
      ,[USA]
FROM [Tirelli_40].[dbo].[Superlistino_Distinte] t0 " &
                          "LEFT JOIN [Tirelli_40].[dbo].[Superlistino_codici] t1 ON t0.Figlio = t1.Codice " &
                          "WHERE t0.Padre = '" & par_codice & "' AND t0.[Optional] = '" & par_optional & "' " & filtro_eu & filtro_usa & filtro_ade & filtro_static & filtro_hot & filtro_flex & " " &
                          "ORDER BY t0.Optional, t0.[N_figlio]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim sommaCostoTotale As Decimal = 0
        Dim sommaPrezzoTotale As Decimal = 0
        Dim percorso_immagine As String
        Dim indexPictureBox As Integer = 0 ' Indice per le PictureBox
        If par_optional = "N" Then


            ' Rimuovi tutte le PictureBox esistenti nella FlowLayoutPanel5
            For i As Integer = FlowLayoutPanel5.Controls.Count - 1 To 0 Step -1
                Dim ctrl As Control = FlowLayoutPanel5.Controls(i)
                If TypeOf ctrl Is PictureBox Then
                    FlowLayoutPanel5.Controls.RemoveAt(i)
                    ctrl.Dispose()
                End If
            Next
        End If



        Do While cmd_SAP_reader.Read()

            ' Imposta il percorso dell'immagine
            percorso_immagine = Homepage.Percorso_Immagini_TICKETS & "Bianco.jpg" ' Immagine di default
            If Not String.IsNullOrEmpty(cmd_SAP_reader("Immagine").ToString()) Then
                Dim immagineFile As String = Homepage.Percorso_Immagini_TICKETS & cmd_SAP_reader("Immagine")
                If File.Exists(immagineFile) Then
                    percorso_immagine = immagineFile ' Usa l'immagine se esiste
                End If

            End If

            ' Controlla se l'immagine esiste prima di aggiungerla
            If File.Exists(percorso_immagine) Then
                ' Carica l'immagine
                Dim image As Image = Image.FromFile(percorso_immagine)

                ' Imposta l'altezza massima desiderata
                Dim maxHeight As Integer = 35
                Dim scaleFactor As Double = maxHeight / image.Height
                Dim newWidth As Integer = CInt(image.Width * scaleFactor)
                Dim newSize As New Size(newWidth, maxHeight)

                ' Crea l'immagine ridimensionata
                Dim smallImage As New Bitmap(image, newSize)

                If par_optional = "N" Then


                    If percorso_immagine <> Homepage.Percorso_Immagini_TICKETS & "Bianco.jpg" Then



                        'Moltiplica per la quantità e aggiungi le immagini
                        Dim qty As Integer = Convert.ToInt32(cmd_SAP_reader("Q"))

                        For i As Integer = 1 To qty
                            Dim picBox As New PictureBox()
                            picBox.Image = smallImage
                            picBox.Size = newSize
                            picBox.SizeMode = PictureBoxSizeMode.StretchImage
                            picBox.Margin = New Padding(5) ' Evita la sovrapposizione
                            FlowLayoutPanel5.Controls.Add(picBox)
                            indexPictureBox += 1
                        Next



                    End If

                End If
                ' Aggiungi i dati alla DataGridView
                Dim costoTot As Decimal = 0
                If Not IsDBNull(cmd_SAP_reader("Costo_Tot")) Then
                    costoTot = Convert.ToDecimal(cmd_SAP_reader("Costo_Tot"))
                    sommaCostoTotale += costoTot
                End If

                par_datagridview.Rows.Add(
                cmd_SAP_reader("id"),
                cmd_SAP_reader("Figlio"),
                cmd_SAP_reader("Descrizione"),
                cmd_SAP_reader("Q"),
                cmd_SAP_reader("Costo"),
                cmd_SAP_reader("Molt"),
                cmd_SAP_reader("Costo_Tot"),
                cmd_SAP_reader("Costificato"),
                cmd_SAP_reader("Note"),
                smallImage,
                  cmd_SAP_reader("EU"),
      cmd_SAP_reader("USA"),
                 cmd_SAP_reader("ADE"),
      cmd_SAP_reader("STATIC"),
      cmd_SAP_reader("HOT"),
      cmd_SAP_reader("FLEX"))
            End If

        Loop

        cmd_SAP_reader.Close()
        Cnn.Close()

        Try
            par_datagridview.FirstDisplayedScrollingRowIndex = par_datagridview.RowCount
        Catch ex As Exception
        End Try

        par_datagridview.ClearSelection()

        If par_optional = "N" Then
            ' Mostra la somma nella Label3 in formato valuta
            par_label_costo.Text = sommaCostoTotale.ToString("C0") ' ad esempio €1.234,56
        End If

    End Sub





    Sub Salva_distinta(par_datagridview As DataGridView, par_codice As String, par_optional As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        'Try
        Cnn.Open()

        ' 1. DELETE esistente
        Dim CMD_Delete As New SqlCommand("DELETE FROM [Tirelli_40].[dbo].[Superlistino_Distinte] WHERE [Padre] = @Padre AND [Optional] = @Optional", Cnn)
        CMD_Delete.Parameters.AddWithValue("@Padre", par_codice)
        CMD_Delete.Parameters.AddWithValue("@Optional", par_optional)
        CMD_Delete.ExecuteNonQuery()

        ' 2. INSERT delle righe da DataGridView
        For i As Integer = 0 To par_datagridview.Rows.Count - 1
            Dim riga As DataGridViewRow = par_datagridview.Rows(i)
            If riga.IsNewRow Then Continue For ' salta l'ultima riga vuota

            Dim CMD_Insert As New SqlCommand("
                INSERT INTO [Tirelli_40].[dbo].[Superlistino_Distinte]
                ([Padre], [Figlio], [Optional], [N_figlio], [Q], [Ultima_revisione], MOLT)
                VALUES (@Padre, @Figlio, @Optional, @N_figlio, @Q, @Ultima_revisione, @MOLT)", Cnn)

            CMD_Insert.Parameters.AddWithValue("@Padre", par_codice)
            CMD_Insert.Parameters.AddWithValue("@Figlio", riga.Cells(1).Value.ToString()) ' Figlio
            CMD_Insert.Parameters.AddWithValue("@Optional", par_optional)
            CMD_Insert.Parameters.AddWithValue("@N_figlio", i + 1) ' posizione
            CMD_Insert.Parameters.AddWithValue("@Q", Convert.ToDecimal(riga.Cells(3).Value)) ' Q (quantità)
            CMD_Insert.Parameters.AddWithValue("@Ultima_revisione", DateTime.Now) ' puoi cambiare se serve
            CMD_Insert.Parameters.AddWithValue("@MOLT", Convert.ToDecimal(riga.Cells(5).Value))

            CMD_Insert.ExecuteNonQuery()
        Next

        'Catch ex As Exception
        '  MsgBox("Errore durante il salvataggio distinta: " & ex.Message)
        'Finally
        Cnn.Close()
        'End Try

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        If e.RowIndex >= 0 Then
            Dim par_datagridview As DataGridView = DataGridView4
            If e.ColumnIndex = par_datagridview.Columns.IndexOf(Img) Then

                Process.Start(Homepage.Percorso_Immagini_TICKETS & Form_Codici_vendita.trova_dettagli_id_codice(par_datagridview.Rows(e.RowIndex).Cells(columnName:="id_codice").Value).immagine)

                '        Form_Codici_vendita.Show()
                '        Form_Codici_vendita.inizializza_form(par_datagridview.Rows(e.RowIndex).Cells(columnName:="Codice").Value)
            End If
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim par_datagridview As DataGridView = DataGridView4

        If par_datagridview.CurrentRow Is Nothing Then Exit Sub

        Dim rigaSelezionata As Integer = par_datagridview.CurrentRow.Index

        ' Sposta la riga su
        Distinta_base_form.SpostaRigaSu(par_datagridview, rigaSelezionata)

        ' Aggiorna la selezione alla nuova posizione
        rigaSelezionata -= 1
        If rigaSelezionata >= 0 Then
            par_datagridview.ClearSelection()
            par_datagridview.Rows(rigaSelezionata).Selected = True
            par_datagridview.CurrentCell = par_datagridview.Rows(rigaSelezionata).Cells(1)
        End If

        ' Mostra/nasconde i bottoni in base alla posizione
        If rigaSelezionata = 0 Then
            Button6.Visible = False
        End If

        Button5.Visible = True ' Nel dubbio, se è visibile puoi tornare a spostare giù
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim par_datagridview As DataGridView = DataGridView4

        If par_datagridview.CurrentRow Is Nothing Then Exit Sub

        Dim rigaSelezionata As Integer = par_datagridview.CurrentRow.Index

        ' Sposta la riga giù
        Distinta_base_form.SpostaRigaGiù(par_datagridview, rigaSelezionata)

        ' Aggiorna la selezione alla nuova posizione
        rigaSelezionata += 1
        If rigaSelezionata < par_datagridview.Rows.Count Then
            par_datagridview.ClearSelection()
            par_datagridview.Rows(rigaSelezionata).Selected = True
            par_datagridview.CurrentCell = par_datagridview.Rows(rigaSelezionata).Cells(1)
        End If

        ' Mostra/nasconde i bottoni in base alla posizione
        If rigaSelezionata = par_datagridview.RowCount - 2 Then
            Button5.Visible = False

        End If
        Button6.Visible = True
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim par_datagridview As DataGridView = DataGridView1

        If par_datagridview.CurrentRow Is Nothing Then Exit Sub

        Dim rigaSelezionata As Integer = par_datagridview.CurrentRow.Index

        ' Sposta la riga su
        Distinta_base_form.SpostaRigaSu(par_datagridview, rigaSelezionata)

        ' Aggiorna la selezione alla nuova posizione
        rigaSelezionata -= 1
        If rigaSelezionata >= 0 Then
            par_datagridview.ClearSelection()
            par_datagridview.Rows(rigaSelezionata).Selected = True
            par_datagridview.CurrentCell = par_datagridview.Rows(rigaSelezionata).Cells(1)
        End If

        ' Mostra/nasconde i bottoni in base alla posizione
        If rigaSelezionata = 0 Then
            Button8.Visible = False
        End If

        Button7.Visible = True ' Nel dubbio, se è visibile puoi tornare a spostare giù
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim par_datagridview As DataGridView = DataGridView1

        If par_datagridview.CurrentRow Is Nothing Then Exit Sub

        Dim rigaSelezionata As Integer = par_datagridview.CurrentRow.Index

        ' Sposta la riga giù
        Distinta_base_form.SpostaRigaGiù(par_datagridview, rigaSelezionata)

        ' Aggiorna la selezione alla nuova posizione
        rigaSelezionata += 1
        If rigaSelezionata < par_datagridview.Rows.Count Then
            par_datagridview.ClearSelection()
            par_datagridview.Rows(rigaSelezionata).Selected = True
            par_datagridview.CurrentCell = par_datagridview.Rows(rigaSelezionata).Cells(1)
        End If

        ' Mostra/nasconde i bottoni in base alla posizione
        If rigaSelezionata = par_datagridview.RowCount - 2 Then
            Button7.Visible = False

        End If
        Button8.Visible = True
    End Sub

    Private Sub Form_distinta_vendita_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Inserimento_serie_macchina(ComboBox1)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        elimina_macchina(TextBox1.Text)
    End Sub

    Sub elimina_macchina(par_codice As String)

        Dim risposta As DialogResult
        risposta = MessageBox.Show("Vuoi eliminare la macchina = " & par_codice & "?", "Conferma eliminazione", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If risposta = DialogResult.Yes Then
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()

            ' Prima elimino la distinta
            Dim deleteDistintaCmd As New SqlCommand("DELETE FROM [Tirelli_40].[dbo].[Superlistino_Distinte] WHERE [Padre] = @Codice", Cnn)
            deleteDistintaCmd.Parameters.AddWithValue("@Codice", par_codice)
            deleteDistintaCmd.ExecuteNonQuery()

            ' Poi elimino l'anagrafica macchina
            Dim deleteMacchinaCmd As New SqlCommand("DELETE FROM [Tirelli_40].[dbo].[Modelli_macchine] WHERE [Codice] = @Codice", Cnn)
            deleteMacchinaCmd.Parameters.AddWithValue("@Codice", par_codice)
            deleteMacchinaCmd.ExecuteNonQuery()

            Cnn.Close()

            MessageBox.Show("Codice eliminato con successo.", "Eliminazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Eliminazione annullata.", "Operazione annullata", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Public Function trova_tipo_macchina(par_codice_macchina As String)

        Dim tipo_macchina As String = ""
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT  [id]
      ,[Codice]
      ,[Tipo_macchina]
      ,[Tipo]
      ,[Nome]
      ,[Serie]
      ,[Diametro_primitivo]
      ,[N_piattelli]
      ,[N_baie]
  FROM [Tirelli_40].[dbo].[Modelli_macchine] where codice='" & par_codice_macchina & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            tipo_macchina = cmd_SAP_reader("Tipo_macchina")
        End If

        cmd_SAP_reader.Close()
        Cnn.Close()


        Return tipo_macchina
    End Function

    Sub elimina_distinta(par_codice As String)

        Dim risposta As DialogResult
        risposta = MessageBox.Show("Vuoi eliminare la macchina = " & par_codice & "?", "Conferma eliminazione", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If risposta = DialogResult.Yes Then
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()

            ' Prima elimino la distinta
            Dim deleteDistintaCmd As New SqlCommand("DELETE FROM [Tirelli_40].[dbo].[Superlistino_Distinte] WHERE [Padre] = @Codice", Cnn)
            deleteDistintaCmd.Parameters.AddWithValue("@Codice", par_codice)
            deleteDistintaCmd.ExecuteNonQuery()



            Cnn.Close()

            MessageBox.Show("Codice eliminato con successo.", "Eliminazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            MessageBox.Show("Eliminazione annullata.", "Operazione annullata", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        elimina_distinta(TextBox1.Text)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        compila_distinta(DataGridView4, TextBox1.Text, "N", True, TextBox5.Text, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, Label3)
        compila_distinta(DataGridView1, TextBox1.Text, "Y", True, TextBox5.Text, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, Label3)
        Button11.Enabled = True
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        compila_distinta(DataGridView4, TextBox1.Text, "N", False, TextBox5.Text, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, Label3)
        compila_distinta(DataGridView1, TextBox1.Text, "Y", False, TextBox5.Text, CheckBox1.Checked, CheckBox2.Checked, CheckBox3.Checked, CheckBox4.Checked, CheckBox5.Checked, CheckBox6.Checked, Label3)
        Button11.Enabled = False
    End Sub

    Private Sub DataGridView4_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView4.CellFormatting
        ' Formattazione colonna "Costificato"
        Dim par_datagrdiview As DataGridView = DataGridView4
        If par_datagrdiview.Columns(e.ColumnIndex).Name = "Costificato" Or par_datagrdiview.Columns(e.ColumnIndex).Name = "Ade" Or par_datagrdiview.Columns(e.ColumnIndex).Name = "Eu" Or par_datagrdiview.Columns(e.ColumnIndex).Name = "Usa" Or par_datagrdiview.Columns(e.ColumnIndex).Name = "Sta" Or par_datagrdiview.Columns(e.ColumnIndex).Name = "HOT" Or par_datagrdiview.Columns(e.ColumnIndex).Name = "FLEX" Then
            If e.Value IsNot Nothing AndAlso Not IsDBNull(e.Value) Then
                Select Case e.Value.ToString().ToUpper()
                    Case "Y"
                        e.CellStyle.BackColor = Color.LightGreen
                    Case "S"
                        e.CellStyle.BackColor = Color.Yellow
                    Case Else
                        e.CellStyle.BackColor = Color.LightCoral
                End Select
            Else
                e.CellStyle.BackColor = Color.LightCoral
            End If
        End If



        ' Formattazione colonna "Costo_tot"
        If par_datagrdiview.Columns(e.ColumnIndex).Name = "Costo_Tot" Then
            If e.Value IsNot Nothing AndAlso IsNumeric(e.Value) Then
                ' Calcolo del massimo dinamico
                Dim maxValore As Decimal = 0D
                For Each row As DataGridViewRow In DataGridView4.Rows
                    If Not row.IsNewRow AndAlso IsNumeric(row.Cells("Costo_tot").Value) Then
                        Dim val As Decimal = Convert.ToDecimal(row.Cells("Costo_tot").Value)
                        If val > maxValore Then
                            maxValore = val
                        End If
                    End If
                Next

                ' Evita divisione per zero
                If maxValore = 0 Then maxValore = 1

                Dim valore As Decimal = Convert.ToDecimal(e.Value)
                Dim intensita As Integer = Math.Min(255, CInt((valore / maxValore) * 255))

                ' Dal verde (basso) al rosso (alto)
                e.CellStyle.BackColor = Color.FromArgb(255, 255 - intensita, intensita)
            End If
        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Dim par_datagridview As DataGridView = DataGridView1
        If par_datagridview.Columns(e.ColumnIndex).Name = "Costificato_" Or par_datagridview.Columns(e.ColumnIndex).Name = "ADE_" Or par_datagridview.Columns(e.ColumnIndex).Name = "EU_" Or par_datagridview.Columns(e.ColumnIndex).Name = "USA_" Or par_datagridview.Columns(e.ColumnIndex).Name = "STA_" Or par_datagridview.Columns(e.ColumnIndex).Name = "HOT_" Or par_datagridview.Columns(e.ColumnIndex).Name = "FLEX_" Then
            If e.Value IsNot Nothing AndAlso Not IsDBNull(e.Value) Then
                Select Case e.Value.ToString().ToUpper()
                    Case "Y"
                        e.CellStyle.BackColor = Color.LightGreen
                    Case "S"
                        e.CellStyle.BackColor = Color.Yellow
                    Case Else
                        e.CellStyle.BackColor = Color.LightCoral
                End Select
            Else
                ' Se il valore è nullo o DBNull, lo consideriamo come "altrimenti"
                e.CellStyle.BackColor = Color.LightCoral
            End If
        End If

        ' Formattazione colonna "Costo_tot"
        If par_datagridview.Columns(e.ColumnIndex).Name = "Costo_Tot" Then
            If e.Value IsNot Nothing AndAlso IsNumeric(e.Value) Then
                ' Calcolo del massimo dinamico
                Dim maxValore As Decimal = 0
                For Each row As DataGridViewRow In par_datagridview.Rows
                    If Not row.IsNewRow AndAlso IsNumeric(row.Cells("Costo_tot_").Value) Then
                        Dim val As Decimal = Convert.ToDecimal(row.Cells("Costo_to_t").Value)
                        If val > maxValore Then
                            maxValore = val
                        End If
                    End If
                Next

                ' Evita divisione per zero
                If maxValore = 0 Then maxValore = 1

                Dim valore As Decimal = Convert.ToDecimal(e.Value)
                Dim intensita As Integer = Math.Min(255, CInt((valore / maxValore) * 255))

                ' Dal verde (basso) al rosso (alto)
                e.CellStyle.BackColor = Color.FromArgb(255, 255 - intensita, intensita)
            End If
        End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub ApriCodiceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ApriCodiceToolStripMenuItem.Click
        Dim currentRow As DataGridViewRow = Nothing

        If DataGridView4.CurrentRow IsNot Nothing Then
            currentRow = DataGridView4.CurrentRow

            If currentRow.Cells("Codice").Value IsNot Nothing Then
                Dim nuovaFinestra As New Form_Codici_vendita()

                nuovaFinestra.Show()
                nuovaFinestra.inizializza_form(currentRow.Cells("Codice").Value.ToString())
            End If
        End If
    End Sub

    Private Sub ApriCodiceToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ApriCodiceToolStripMenuItem1.Click
        Dim currentRow As DataGridViewRow = Nothing

        If DataGridView1.CurrentRow IsNot Nothing Then
            currentRow = DataGridView1.CurrentRow

            If currentRow.Cells("Codice_").Value IsNot Nothing Then
                Dim nuovaFinestra As New Form_Codici_vendita()

                nuovaFinestra.Show()
                nuovaFinestra.inizializza_form(currentRow.Cells("Codice_").Value.ToString())
            End If
        End If





    End Sub

    Private Sub EliminaRigaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EliminaRigaToolStripMenuItem.Click
        If DataGridView4.CurrentRow IsNot Nothing Then
            DataGridView4.Rows.Remove(DataGridView4.CurrentRow)
        End If
    End Sub

    Private Sub EliminaRigaToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles EliminaRigaToolStripMenuItem1.Click
        If DataGridView1.CurrentRow IsNot Nothing Then
            DataGridView1.Rows.Remove(DataGridView1.CurrentRow)
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

    End Sub

    Private Sub DataGridView4_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellValueChanged
        Dim par_datagridview As DataGridView = DataGridView4
        Dim itemcode As String
        If e.RowIndex >= 0 Then

            itemcode = UCase(par_datagridview.Rows(e.RowIndex).Cells(columnName:="Codice").Value)
            If e.ColumnIndex = par_datagridview.Columns.IndexOf(Codice) Then


                'Try


                par_datagridview.Rows(e.RowIndex).Cells(columnName:="Desc").Value = ottieni_informazioni_codice_vendita(itemcode).descrizione
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="Q").Value = 1
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="costo_u").Value = ottieni_informazioni_codice_vendita(itemcode).costo_U
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="Molt").Value = 1
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="costo_tot").Value = ottieni_informazioni_codice_vendita(itemcode).costo_U
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="Costificato").Value = ottieni_informazioni_codice_vendita(itemcode).costificato
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="Note").Value = ottieni_informazioni_codice_vendita(itemcode).note

                Dim percorso_immagine As String


                percorso_immagine = Homepage.Percorso_Immagini_TICKETS & "Bianco.jpg"
                Dim TEMPO = Homepage.Percorso_Immagini_TICKETS & ottieni_informazioni_codice_vendita(itemcode).immagine
                If File.Exists(Homepage.Percorso_Immagini_TICKETS & ottieni_informazioni_codice_vendita(itemcode).immagine) Then
                    percorso_immagine = Homepage.Percorso_Immagini_TICKETS & ottieni_informazioni_codice_vendita(itemcode).immagine

                End If
                ' Load the image from file path
                Dim image As Image = Image.FromFile(percorso_immagine)
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="img").Value = image

            ElseIf e.ColumnIndex = par_datagridview.Columns.IndexOf(Q) Then

                'Dim itemcode As String = par_datagridview.Rows(e.RowIndex).Cells("itemcode").Value.ToString()
                Dim costo_U As Decimal = ottieni_informazioni_codice_vendita(itemcode).costo_U
                Dim Q As Decimal = Convert.ToDecimal(par_datagridview.Rows(e.RowIndex).Cells("Q").Value)
                Dim MoltValore As Decimal
                ' Leggi il valore di Molt con cultura italiana
                If par_datagridview.Rows(e.RowIndex).Cells("Molt").Value = Nothing Then
                    par_datagridview.Rows(e.RowIndex).Cells("Molt").Value = 1
                End If


                Dim MoltString As String = Replace(par_datagridview.Rows(e.RowIndex).Cells("Molt").Value.ToString(), ".", ",")
                    MoltValore = Decimal.Parse(MoltString, New CultureInfo("it-IT"))


                par_datagridview.Rows(e.RowIndex).Cells("costo_tot").Value = costo_U * Q * MoltValore
            ElseIf e.ColumnIndex = par_datagridview.Columns.IndexOf(Molt) Then

                'Dim itemcode As String = par_datagridview.Rows(e.RowIndex).Cells("itemcode").Value.ToString()
                Dim costo_U As Decimal = ottieni_informazioni_codice_vendita(itemcode).costo_U
                Dim Q As Decimal = Convert.ToDecimal(par_datagridview.Rows(e.RowIndex).Cells("Q").Value)

                ' Leggi il valore di Molt con cultura italiana
                Dim MoltString As String = Replace(par_datagridview.Rows(e.RowIndex).Cells("Molt").Value.ToString(), ".", ",")
                Dim MoltValore As Decimal = Decimal.Parse(MoltString, New CultureInfo("it-IT"))

                par_datagridview.Rows(e.RowIndex).Cells("costo_tot").Value = costo_U * Q * MoltValore

            End If
        End If

        ' Calcola il totale della colonna "costo_tot_"
        Dim totale As Decimal = 0
        For Each row As DataGridViewRow In par_datagridview.Rows
            If Not row.IsNewRow AndAlso IsNumeric(row.Cells("costo_tot").Value) Then
                totale += Convert.ToDecimal(row.Cells("costo_tot").Value)
            End If
        Next

        Label3.Text = totale.ToString("N2") ' Formattato con due decimali

    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        Dim par_datagridview As DataGridView = DataGridView1
        Dim itemcode As String

        If e.RowIndex >= 0 Then
            itemcode = UCase(par_datagridview.Rows(e.RowIndex).Cells("Codice_").Value)

            If e.ColumnIndex = par_datagridview.Columns("Codice_").Index Then
                par_datagridview.Rows(e.RowIndex).Cells("Desc_").Value = ottieni_informazioni_codice_vendita(itemcode).descrizione
                par_datagridview.Rows(e.RowIndex).Cells("Q_").Value = 1
                par_datagridview.Rows(e.RowIndex).Cells("costo_u_").Value = ottieni_informazioni_codice_vendita(itemcode).costo_U
                par_datagridview.Rows(e.RowIndex).Cells("Molt_").Value = 1
                par_datagridview.Rows(e.RowIndex).Cells("costo_tot_").Value = ottieni_informazioni_codice_vendita(itemcode).costo_U
                par_datagridview.Rows(e.RowIndex).Cells("Costificato_").Value = ottieni_informazioni_codice_vendita(itemcode).costificato
                par_datagridview.Rows(e.RowIndex).Cells("Note_").Value = ottieni_informazioni_codice_vendita(itemcode).note

                Dim percorso_immagine As String = Homepage.Percorso_Immagini_TICKETS & "Bianco.jpg"
                Dim immagine_path As String = Homepage.Percorso_Immagini_TICKETS & ottieni_informazioni_codice_vendita(itemcode).immagine

                If File.Exists(immagine_path) Then
                    percorso_immagine = immagine_path
                End If

                Dim image As Image = Image.FromFile(percorso_immagine)
                par_datagridview.Rows(e.RowIndex).Cells("img_").Value = image

            ElseIf e.ColumnIndex = par_datagridview.Columns("Q_").Index Then
                Dim quantita As Decimal = Convert.ToDecimal(par_datagridview.Rows(e.RowIndex).Cells("Q_").Value)
                Dim MoltString As String = Replace(par_datagridview.Rows(e.RowIndex).Cells("Molt").Value.ToString(), ".", ",")
                par_datagridview.Rows(e.RowIndex).Cells("costo_tot_").Value = ottieni_informazioni_codice_vendita(itemcode).costo_U * quantita * MoltString






            ElseIf e.ColumnIndex = par_datagridview.Columns("Molt_").Index Then
                Dim quantita As Decimal = Convert.ToDecimal(par_datagridview.Rows(e.RowIndex).Cells("Q_").Value)
                Dim MoltString As String = Replace(par_datagridview.Rows(e.RowIndex).Cells("Molt_").Value.ToString(), ".", ",")
                par_datagridview.Rows(e.RowIndex).Cells("costo_tot_").Value = ottieni_informazioni_codice_vendita(itemcode).costo_U * quantita * MoltString
            End If
        End If


    End Sub






    Public Function ottieni_informazioni_codice_vendita(PAR_CODICE As String)

        Dim dettagli As New Dettagli_CODICE()

        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT  [ID]
      ,[Codice]
      ,[Tipo_macchina]
      ,[Descrizione]
      ,[Costo]
      ,[ultima_revisione]
      ,COALESCE([Note],'') AS 'Note'
      ,[Active]
      ,[Sottogruppo]
      ,[ADE]
      ,[STATIC]
      ,[HOT]
      ,[FLEX]
      ,[EU]
      ,[USA]
      ,[Costo_Materiale]
      ,COALESCE([Costificato],'N') as 'Costificato'
      ,coalesce([immagine],'') as 'Immagine'
  FROM [Tirelli_40].[dbo].[Superlistino_codici]
  where codice='" & PAR_CODICE & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() = True Then
            ' dettagli.Descrizione = cmd_SAP_reader("itemname")
            dettagli.id = cmd_SAP_reader_2("id")
            dettagli.Descrizione = cmd_SAP_reader_2("Descrizione")
            dettagli.Costo_u = cmd_SAP_reader_2("Costo")
            dettagli.costificato = cmd_SAP_reader_2("Costificato")
            dettagli.note = cmd_SAP_reader_2("note")
            dettagli.immagine = cmd_SAP_reader_2("Immagine")



        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()





        Return dettagli
    End Function

    Public Class Dettagli_CODICE
        Public id As Integer
        Public Descrizione As String
        Public Costo_u As String
        Public costificato As String
        Public note As String
        Public immagine As String

    End Class

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        inizializza_form(TextBox1.Text)
    End Sub

    Private Sub Cmd_Indietro_Click(sender As Object, e As EventArgs) Handles Cmd_Indietro.Click
        inizializza_form(trova_macchina(Label1.Text - 1))
    End Sub

    Private Sub Cmd_Avanti_Click(sender As Object, e As EventArgs) Handles Cmd_Avanti.Click
        inizializza_form(trova_macchina(Label1.Text + 1))
    End Sub
End Class