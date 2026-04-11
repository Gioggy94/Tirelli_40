Imports System.Data.SqlClient

Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Tirelli.ODP_Form

Public Class ucConfiguratoreModulo
    Public Property Posizione As Integer
    Public Property Codice_macchina As String
        Get
            Return lblcodice.Text
        End Get
        Set(value As String)
            lblcodice.Text = value
        End Set
    End Property

    Public listino As Decimal
    Public Eu_ As String
    Public Usa_ As String
    Public Ade_ As String
    Public Static_ As String
    Public HOT_ As String
    Public FLEX_ As String
    Public valuta As String





    ' Aggiunge righe alla DataGridView interna (DataGridView4)
    Public Sub AggiungiRigheADatagrid(par_codice_macchina As String, par_optional As String, par_filtro As Boolean, par_EU As Boolean, par_usa As Boolean, par_Ade As Boolean, par_static As Boolean, par_hot As Boolean, par_flex As Boolean, par_listino As Decimal, par_valuta As String, par_cambio As Decimal)
        compila_distinta(DataGridView4, par_codice_macchina, par_optional, par_filtro, par_EU, par_usa, par_Ade, par_static, par_hot, Label2, Label4, Label3, par_listino, par_valuta, par_cambio)
        compila_distinta(DataGridView1, par_codice_macchina, "Y", par_filtro, par_EU, par_usa, par_Ade, par_static, par_hot, Label2, Label4, Label3, par_listino, par_valuta, par_cambio)
    End Sub

    'Public Property Titolo As String
    '    Get
    '        Return textbox_tit.Text
    '    End Get
    '    Set(value As String)
    '        lbltitolo.Text = value
    '    End Set
    'End Property

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
      ,[N_baie]

  FROM [Tirelli_40].[dbo].[Modelli_macchine]
 where [Codice] ='" & par_codice & "'"
        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            ' Label1.Text = cmd_SAP_reader("ID")
            ' TextBox1.Text = cmd_SAP_reader("Codice")
            TextBox_titolo_macchina.Text = cmd_SAP_reader("nome")
            txtQuantita.Text = 1
            '  TextBox5.Text = cmd_SAP_reader("Tipo_macchina")
            '  TextBox6.Text = cmd_SAP_reader("Tipo")
            '  ComboBox1.Text = cmd_SAP_reader("Serie")
            '  TextBox3.Text = cmd_SAP_reader("Diametro_primitivo")
            '  TextBox4.Text = cmd_SAP_reader("N_piattelli")
            '  TextBox7.Text = cmd_SAP_reader("N_baie")


        End If
        cmd_SAP_reader.Close()
        Cnn.Close()


    End Sub



    Public Property Figli As New List(Of ucConfiguratoreModulo)

    Public Event ConfigurazioneCambiata()

    Private Sub cmbOpzioni_SelectedIndexChanged(sender As Object, e As EventArgs)
        ' Cancella eventuali figli precedenti

        Figli.Clear()

        ' Esempio: aggiunge un figlio solo se la scelta è "Motore A"


        RaiseEvent ConfigurazioneCambiata()
    End Sub

    Public Function GetConfigurazione() As List(Of ConfigurazioneModulo)
        Dim lista As New List(Of ConfigurazioneModulo)

        Dim conf As New ConfigurazioneModulo With {
            .Titolo = Me.Codice_macchina}

        lista.Add(conf)

        For Each f As ucConfiguratoreModulo In Figli
            lista.AddRange(f.GetConfigurazione())
        Next

        Return lista
    End Function

    Private Sub NotificaCambio()
        RaiseEvent ConfigurazioneCambiata()
    End Sub

    Public Class ConfigurazioneModulo
        Public Property Titolo As String

        Public Property Quantita As Integer
    End Class

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Rimuove l'usercontrol dal suo contenitore (es. FlowLayoutPanel)
        Me.Parent.Controls.Remove(Me)
        For i As Integer = Form_configuratore_vendita.DataGridView4.Rows.Count - 1 To 0 Step -1
            Dim row As DataGridViewRow = Form_configuratore_vendita.DataGridView4.Rows(i)
            If Not row.IsNewRow Then
                If row.Cells("N").Value.ToString() = Label1.Text Then
                    Form_configuratore_vendita.DataGridView4.Rows.RemoveAt(i)
                End If
            End If
        Next
    End Sub

    Public Sub compila_distinta(par_datagridview As DataGridView, par_codice As String, par_optional As String, par_filtro As Boolean, par_EU As Boolean, par_usa As Boolean, par_Ade As Boolean, par_static As Boolean, par_hot As Boolean, par_label_costo As Label, par_label_prezzo As Label, par_label_moltiplicatore As Label, par_listino As Decimal, par_valuta As String, par_cambio As Decimal)

        If par_valuta = "E" Then
            For Each col As DataGridViewColumn In par_datagridview.Columns
                If col.Name.ToLower().StartsWith("prezzo_usd") Then
                    col.Visible = False
                End If
            Next
        End If
        Dim filtro_eu, filtro_usa, filtro_ade, filtro_static, filtro_hot, filtro_flex As String
        If par_filtro = True Then
            ' Gestisci i filtri
            filtro_eu = If(par_EU, " and t1.eu='Y' ", "")
            filtro_usa = If(par_usa, " and t1.usa='Y' ", "")

        Else
            filtro_eu = ""
            filtro_usa = ""

        End If

        If Form_distinta_vendita.trova_tipo_macchina(par_codice) = "ETICHETTATRICE BRB" Then
            If par_filtro = True Then
                filtro_ade = If(par_Ade, " and t1.ade='Y' ", "")
                filtro_static = If(par_static, " and t1.static='Y' ", "")
                filtro_hot = If(par_hot, " and t1.hot='Y' ", "")

            Else
                filtro_ade = ""
                filtro_static = ""
                filtro_hot = ""

            End If
        End If
        Dim stringa_opt As String
        If par_optional = "Y" Then
            stringa_opt = "O"
        Else
            stringa_opt = "S"
        End If
        par_datagridview.Rows.Clear()

        ' Connessione al database
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT t0.[Padre], t1.id, t0.[Figlio], t1.Descrizione, t0.[Optional], t0.[N_figlio], t0.[Q], t1.Costo_Materiale, t1.Costo, t1.costificato, t0.[Q]*t1.Costo as 'Costo_Tot',
t0.molt as 'Moltiplicatore',

t1.[Note], t0.[Ultima_revisione], coalesce(t1.immagine,'') as 'Immagine' 
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


                ' Aggiungi i dati alla DataGridView
                Dim costoTot As Decimal = 0
                If Not IsDBNull(cmd_SAP_reader("Costo_Tot")) Then
                    costoTot = Convert.ToDecimal(cmd_SAP_reader("Costo_Tot"))
                    sommaCostoTotale += costoTot
                End If

                Dim prezzoTot As Decimal = 0
                If Not IsDBNull(cmd_SAP_reader("Costo_Tot") * par_listino) Then
                    prezzoTot = Convert.ToDecimal(cmd_SAP_reader("Costo_Tot") * par_listino)
                    sommaPrezzoTotale += prezzoTot
                End If

                par_datagridview.Rows.Add(
                cmd_SAP_reader("id"),
                cmd_SAP_reader("Figlio"),
                cmd_SAP_reader("Descrizione"),
                cmd_SAP_reader("Q"),
                cmd_SAP_reader("Costo"),
                cmd_SAP_reader("Costo_Tot"),
                par_listino,
                cmd_SAP_reader("Costo_Tot") * par_listino,
                cmd_SAP_reader("Costo_Tot") * par_listino * par_cambio,
                cmd_SAP_reader("Costificato"),
                cmd_SAP_reader("Note"),
                smallImage,
                stringa_opt,
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

            If par_valuta = "E" Then
                par_label_costo.Text = sommaCostoTotale.ToString("C0", Globalization.CultureInfo.GetCultureInfo("it-IT")) ' ad esempio €1.234
                par_label_prezzo.Text = sommaPrezzoTotale.ToString("C0", Globalization.CultureInfo.GetCultureInfo("it-IT"))
            ElseIf par_valuta = "$" Then
                par_label_costo.Text = (sommaCostoTotale * par_cambio).ToString("C0", Globalization.CultureInfo.GetCultureInfo("en-US")) ' es. $1,234
                par_label_prezzo.Text = (sommaPrezzoTotale * par_cambio).ToString("C0", Globalization.CultureInfo.GetCultureInfo("en-US"))
            Else
                ' Valuta non gestita, uso simbolo generico e separatori standard
                par_label_costo.Text = par_valuta & " " & (sommaCostoTotale * par_cambio).ToString("N0")
                par_label_prezzo.Text = par_valuta & " " & (sommaPrezzoTotale * par_cambio).ToString("N0")
            End If
            Try
                par_label_moltiplicatore.Text = sommaPrezzoTotale / sommaCostoTotale
            Catch ex As Exception

            End Try
        End If
        If par_optional = "N" Then
            If par_valuta = "E" Then
                Form_configuratore_vendita.DataGridView4.Columns("Prezzo_usd").Visible = False
            Else
                Form_configuratore_vendita.DataGridView4.Columns("Prezzo_usd").Visible = True
            End If
            Form_configuratore_vendita.DataGridView4.Rows.Add(Label1.Text, TextBox_titolo_macchina.Text, txtQuantita.Text, par_label_moltiplicatore.Text, sommaPrezzoTotale.ToString("C0"), sommaPrezzoTotale * par_cambio)

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
            Button6.Visible = True
        End If
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
        If par_datagrdiview.Columns(e.ColumnIndex).Name = "Prezzo" Then
            If e.Value IsNot Nothing AndAlso IsNumeric(e.Value) Then
                ' Calcolo del massimo dinamico
                Dim maxValore As Decimal = 0D
                For Each row As DataGridViewRow In DataGridView4.Rows
                    If Not row.IsNewRow AndAlso IsNumeric(row.Cells("Prezzo").Value) Then
                        Dim val As Decimal = Convert.ToDecimal(row.Cells("Prezzo").Value)
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

        If par_datagrdiview.Columns(e.ColumnIndex).Name = "Prezzo_usd" Then
            If e.Value IsNot Nothing AndAlso IsNumeric(e.Value) Then
                ' Calcolo del massimo dinamico
                Dim maxValore As Decimal = 0D
                For Each row As DataGridViewRow In DataGridView4.Rows
                    If Not row.IsNewRow AndAlso IsNumeric(row.Cells("Prezzo_usd").Value) Then
                        Dim val As Decimal = Convert.ToDecimal(row.Cells("Prezzo_usd").Value)
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

        If par_datagrdiview.Rows(e.RowIndex).Cells(columnName:="Tipo").Value = "S" Then
            par_datagrdiview.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.Black
        ElseIf par_datagrdiview.Rows(e.RowIndex).Cells(columnName:="Tipo").Value = "O" Then
            par_datagrdiview.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.DarkGreen
        ElseIf par_datagrdiview.Rows(e.RowIndex).Cells(columnName:="Tipo").Value = "N" Then
            par_datagrdiview.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.DARKOrange
        End If



        If par_datagrdiview.Columns(e.ColumnIndex).Name = "Prezzo_usd" AndAlso IsNumeric(e.Value) Then
            Dim valore As Decimal = Convert.ToDecimal(e.Value)
            e.Value = valore.ToString("$#,0", Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            e.FormattingApplied = True
        End If


    End Sub



    Private Sub ApriCodiceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ApriCodiceToolStripMenuItem.Click
        Dim currentRow As DataGridViewRow = Nothing

        If DataGridView4.CurrentRow IsNot Nothing Then
            currentRow = DataGridView4.CurrentRow

            If currentRow.Cells("Codice").Value IsNot Nothing Then
                Dim nuovaFinestra As New Form_Codici_vendita()
                nuovaFinestra.TopMost = True
                nuovaFinestra.Show()
                nuovaFinestra.inizializza_form(currentRow.Cells("Codice").Value.ToString())
            End If
        End If
    End Sub

    Private Sub EliminaRigaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EliminaRigaToolStripMenuItem.Click
        Dim par_datagridview As DataGridView = DataGridView4
        If par_datagridview.CurrentRow IsNot Nothing Then
            par_datagridview.Rows.Remove(DataGridView4.CurrentRow)
        End If
        ' Calcola il totale della colonna "costo_tot_"
        Dim totale_costo As Decimal = 0
        Dim totale_prezzo As Decimal = 0
        For Each row As DataGridViewRow In par_datagridview.Rows
            If Not row.IsNewRow AndAlso IsNumeric(row.Cells("costo_tot").Value) Then
                totale_costo += Convert.ToDecimal(row.Cells("costo_tot").Value)
            End If

            If Not row.IsNewRow AndAlso IsNumeric(row.Cells("prezzo").Value) Then
                totale_prezzo += Convert.ToDecimal(row.Cells("prezzo").Value)
            End If
        Next

        Label2.Text = totale_costo.ToString("N2") ' Formattato con due decimali
        Label4.Text = totale_prezzo.ToString("N2")
        Try
            Label3.Text = (totale_prezzo / totale_costo).ToString("0.00")
        Catch ex As Exception

        End Try

        For Each row As DataGridViewRow In Form_configuratore_vendita.DataGridView4.Rows
            If Not row.IsNewRow Then
                If row.Cells("N").Value.ToString() = Label1.Text Then

                    row.Cells("Molt").Value = (totale_prezzo / totale_costo).ToString("0.00")
                    row.Cells("q").Value = txtQuantita.Text
                    row.Cells("prezzo").Value = txtQuantita.Text * totale_prezzo.ToString("N2")
                End If
            End If
        Next

    End Sub

    Private Sub DataGridView4_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellValueChanged
        Dim par_datagridview As DataGridView = DataGridView4
        Dim itemcode As String
        If e.RowIndex >= 0 Then

            itemcode = UCase(par_datagridview.Rows(e.RowIndex).Cells(columnName:="Codice").Value)
            If e.ColumnIndex = par_datagridview.Columns.IndexOf(Codice) Then


                'Try



                par_datagridview.Rows(e.RowIndex).Cells(columnName:="Desc").Value = Form_distinta_vendita.ottieni_informazioni_codice_vendita(itemcode).descrizione
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="Q").Value = 1
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="costo_u").Value = Form_distinta_vendita.ottieni_informazioni_codice_vendita(itemcode).costo_U
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="costo_tot").Value = Form_distinta_vendita.ottieni_informazioni_codice_vendita(itemcode).costo_U
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="molt").Value = 1
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="prezzo").Value = par_datagridview.Rows(e.RowIndex).Cells(columnName:="costo_tot").Value * par_datagridview.Rows(e.RowIndex).Cells(columnName:="molt").Value

                par_datagridview.Rows(e.RowIndex).Cells(columnName:="Costificato").Value = Form_distinta_vendita.ottieni_informazioni_codice_vendita(itemcode).costificato
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="Note").Value = Form_distinta_vendita.ottieni_informazioni_codice_vendita(itemcode).note
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="Tipo").Value = "N"
                Dim percorso_immagine As String


                percorso_immagine = Homepage.Percorso_Immagini_TICKETS & "Bianco.jpg"
                Dim TEMPO = Homepage.Percorso_Immagini_TICKETS & Form_distinta_vendita.ottieni_informazioni_codice_vendita(itemcode).immagine
                If File.Exists(Homepage.Percorso_Immagini_TICKETS & Form_distinta_vendita.ottieni_informazioni_codice_vendita(itemcode).immagine) Then
                    percorso_immagine = Homepage.Percorso_Immagini_TICKETS & Form_distinta_vendita.ottieni_informazioni_codice_vendita(itemcode).immagine

                End If
                ' Load the image from file path
                Dim image As Image = Image.FromFile(percorso_immagine)



                ' Imposta l'altezza massima desiderata
                Dim maxHeight As Integer = 35
                Dim scaleFactor As Double = maxHeight / image.Height
                Dim newWidth As Integer = CInt(image.Width * scaleFactor)
                Dim newSize As New Size(newWidth, maxHeight)

                ' Crea l'immagine ridimensionata
                Dim smallImage As New Bitmap(image, newSize)


                par_datagridview.Rows(e.RowIndex).Cells(columnName:="img").Value = smallImage

            ElseIf e.ColumnIndex = par_datagridview.Columns.IndexOf(Q) Then
                If par_datagridview.Rows(e.RowIndex).Cells("Molt").Value <> Nothing Then
                    Dim MoltString As String = Replace(par_datagridview.Rows(e.RowIndex).Cells("Molt").Value.ToString(), ".", ",")
                    par_datagridview.Rows(e.RowIndex).Cells(columnName:="costo_tot").Value = Form_distinta_vendita.ottieni_informazioni_codice_vendita(itemcode).costo_U * par_datagridview.Rows(e.RowIndex).Cells(columnName:="Q").Value

                    par_datagridview.Rows(e.RowIndex).Cells(columnName:="PREZZO").Value = par_datagridview.Rows(e.RowIndex).Cells(columnName:="costo_tot").Value * MoltString

                End If



            ElseIf e.ColumnIndex = par_datagridview.Columns.IndexOf(Molt) Then
                Dim MoltString As String = Replace(par_datagridview.Rows(e.RowIndex).Cells("Molt").Value.ToString(), ".", ",")
                par_datagridview.Rows(e.RowIndex).Cells(columnName:="PREZZO").Value = par_datagridview.Rows(e.RowIndex).Cells(columnName:="costo_tot").Value * MoltString
            End If
        End If

        ' Calcola il totale della colonna "costo_tot_"
        Dim totale_costo As Decimal = 0
        Dim totale_prezzo As Decimal = 0
        For Each row As DataGridViewRow In par_datagridview.Rows
            If Not row.IsNewRow AndAlso IsNumeric(row.Cells("costo_tot").Value) Then
                totale_costo += Convert.ToDecimal(row.Cells("costo_tot").Value)
            End If

            If Not row.IsNewRow AndAlso IsNumeric(row.Cells("prezzo").Value) Then
                totale_prezzo += Convert.ToDecimal(row.Cells("prezzo").Value)
            End If
        Next

        Label2.Text = totale_costo.ToString("N2") ' Formattato con due decimali
        Label4.Text = totale_prezzo.ToString("N2")
        Try
            Label3.Text = (totale_prezzo / totale_costo).ToString("0.00")
        Catch ex As Exception

        End Try

        For Each row As DataGridViewRow In Form_configuratore_vendita.DataGridView4.Rows
            If Not row.IsNewRow Then
                If row.Cells("N").Value.ToString() = Label1.Text Then

                    row.Cells("Molt").Value = (totale_prezzo / totale_costo).ToString("0.00")
                    row.Cells("q").Value = txtQuantita.Text
                    row.Cells("prezzo").Value = txtQuantita.Text * totale_prezzo.ToString("N2")
                End If
            End If
        Next
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub ContextMenuStrip2_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs)

    End Sub

    Private Sub TrasferisciInMacchinaBaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TrasferisciInMacchinaBaseToolStripMenuItem.Click
        Dim par_datagridview As DataGridView = DataGridView1
        Dim currentRow As DataGridViewRow = Nothing

        If par_datagridview.CurrentRow IsNot Nothing Then
            currentRow = par_datagridview.CurrentRow

            If currentRow.Cells("Codice_").Value IsNot Nothing Then
                ' Prendo il valore del codice selezionato
                Dim codiceSelezionato As String = currentRow.Cells("Codice_").Value.ToString()
                Dim q_Selezionato As String = currentRow.Cells("Q_").Value.ToString()

                ' Ottengo le informazioni del codice
                Dim info = Form_distinta_vendita.ottieni_informazioni_codice_vendita(codiceSelezionato)

                ' Preparo il percorso immagine
                Dim percorso_immagine As String = Homepage.Percorso_Immagini_TICKETS & "Bianco.jpg"
                Dim percorsoTentativo As String = Homepage.Percorso_Immagini_TICKETS & info.immagine

                If File.Exists(percorsoTentativo) Then
                    percorso_immagine = percorsoTentativo
                End If

                ' Carico e ridimensiono l'immagine
                Dim image As Image = Image.FromFile(percorso_immagine)
                Dim maxHeight As Integer = 35
                Dim scaleFactor As Double = maxHeight / image.Height
                Dim newWidth As Integer = CInt(image.Width * scaleFactor)
                Dim newSize As New Size(newWidth, maxHeight)
                Dim smallImage As New Bitmap(image, newSize)

                ' Aggiungo direttamente la riga completa
                DataGridView4.Rows.Add(info.id,
                codiceSelezionato,                ' Codice
                info.descrizione,                  ' Desc
                q_Selezionato,                                 ' Q
                info.costo_U,                      ' costo_u
                q_Selezionato * info.costo_U,                  ' costo_tot (Q * costo_u)
                listino,                                 ' molt
                q_Selezionato * info.costo_U * listino,                               ' prezzo (se fisso 999)
                info.costificato,                  ' Costificato
                info.note,                         ' Note
                smallImage,                         ' img
                "O"
            )
            End If
        End If


        par_datagridview = DataGridView4
        ' Calcola il totale della colonna "costo_tot_"
        Dim totale_costo As Decimal = 0
        Dim totale_prezzo As Decimal = 0
        For Each row As DataGridViewRow In par_datagridview.Rows
            If Not row.IsNewRow AndAlso IsNumeric(row.Cells("costo_tot").Value) Then
                totale_costo += Convert.ToDecimal(row.Cells("costo_tot").Value)
            End If

            If Not row.IsNewRow AndAlso IsNumeric(row.Cells("prezzo").Value) Then
                totale_prezzo += Convert.ToDecimal(row.Cells("prezzo").Value)
            End If
        Next

        Label2.Text = totale_costo.ToString("N2") ' Formattato con due decimali
        Label4.Text = totale_prezzo.ToString("N2")
        Try
            Label3.Text = (totale_prezzo / totale_costo).ToString("0.00")
        Catch ex As Exception

        End Try

        For Each row As DataGridViewRow In Form_configuratore_vendita.DataGridView4.Rows
            If Not row.IsNewRow Then
                If row.Cells("N").Value.ToString() = Label1.Text Then

                    row.Cells("Molt").Value = (totale_prezzo / totale_costo).ToString("0.00")
                    row.Cells("q").Value = txtQuantita.Text
                    row.Cells("prezzo").Value = txtQuantita.Text * totale_prezzo.ToString("N2")
                End If
            End If
        Next
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub TextBox_titolo_macchina_TextChanged(sender As Object, e As EventArgs) Handles TextBox_titolo_macchina.TextChanged
        For i As Integer = Form_configuratore_vendita.DataGridView4.Rows.Count - 1 To 0 Step -1
            Dim row As DataGridViewRow = Form_configuratore_vendita.DataGridView4.Rows(i)
            If Not row.IsNewRow Then
                If row.Cells("N").Value.ToString() = Label1.Text Then

                    row.Cells("Desc").Value = TextBox_titolo_macchina.Text
                End If
            End If
        Next
    End Sub

    Private Sub txtQuantita_TextChanged(sender As Object, e As EventArgs) Handles txtQuantita.TextChanged
        For i As Integer = Form_configuratore_vendita.DataGridView4.Rows.Count - 1 To 0 Step -1
            Dim row As DataGridViewRow = Form_configuratore_vendita.DataGridView4.Rows(i)
            If Not row.IsNewRow Then
                If row.Cells("N").Value.ToString() = Label1.Text Then

                    row.Cells("Q").Value = txtQuantita.Text
                    row.Cells("Prezzo").Value = txtQuantita.Text * Label4.Text

                End If
            End If
        Next
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Dim par_datagridview As DataGridView = DataGridView1
        If par_datagridview.Columns(e.ColumnIndex).Name = "Prezzo_usd_" AndAlso IsNumeric(e.Value) Then
            Dim valore As Decimal = Convert.ToDecimal(e.Value)
            e.Value = valore.ToString("$#,0", Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            e.FormattingApplied = True
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub
End Class
