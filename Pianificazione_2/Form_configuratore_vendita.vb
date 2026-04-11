Imports System.Data.SqlClient
Imports stdole
Imports Tirelli.ucConfiguratoreModulo

Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Tirelli.ODP_Form
Imports MS.Internal.Xaml

Public Class Form_configuratore_vendita

    Public N_opportunità As Integer
    Public numero_ultima_revisione As Integer
    Public valuta As String



    Public Sub inizializzazione_modulo(par_codice As String, par_optional As String, par_filtro As Boolean, par_EU As Boolean, par_usa As Boolean, par_Ade As Boolean, par_static As Boolean, par_hot As Boolean, par_flex As Boolean, par_listino As Decimal)
        ' Crea un nuovo modulo
        Dim moduloBase As New ucConfiguratoreModulo With {
        .Codice_macchina = par_codice
    }

        ' Collega l'evento click
        AddHandler moduloBase.Click, AddressOf Modulo_Click

        ' Aggiungi al contenitore visivo
        panelModuli.Controls.Add(moduloBase)

        moduloBase.listino = par_listino
        ' Non serve SetChildIndex: così lo aggiunge sempre in fondo
        moduloBase.Posizione = panelModuli.Controls.GetChildIndex(moduloBase) + 1
        moduloBase.Label1.Text = moduloBase.Posizione
        ' Imposta il Dock a Top per impilarlo automaticamente

        moduloBase.Width = panelModuli.Width - 15
        '  moduloBase.Dock = DockStyle.Top
        ' Aggiungi le righe
        moduloBase.compila_anagrafica(par_codice)

        moduloBase.Eu_ = If(par_EU, "Y", "N")
        moduloBase.Usa_ = If(par_usa, "Y", "N")
        moduloBase.Ade_ = If(par_Ade, "Y", "N")
        moduloBase.Static_ = If(par_static, "Y", "N")
        moduloBase.HOT_ = If(par_hot, "Y", "N")
        moduloBase.FLEX_ = If(par_flex, "Y", "N")
        moduloBase.RadioButton1.Checked = par_EU
        moduloBase.RadioButton2.Checked = par_usa

        moduloBase.AggiungiRigheADatagrid(par_codice, par_optional, par_filtro, par_EU, par_usa, par_Ade, par_static, par_hot, par_flex, par_listino, valuta, Label3.Text)
    End Sub

    Private Sub Modulo_Click(sender As Object, e As EventArgs)
        Dim moduloCliccato As ucConfiguratoreModulo = DirectCast(sender, ucConfiguratoreModulo)
        MsgBox("Hai cliccato sul modulo con codice: " & moduloCliccato.Codice_macchina)
    End Sub


    Public Sub aggiorna_posizioni_moduli()
        For i As Integer = 0 To panelModuli.Controls.Count - 1
            Dim ctrl = panelModuli.Controls(i)
            If TypeOf ctrl Is ucConfiguratoreModulo Then
                CType(ctrl, ucConfiguratoreModulo).Posizione = i
            End If
        Next
    End Sub


    Private Sub btnConferma_Click(sender As Object, e As EventArgs)
        ' Raccogli le configurazioni da tutti i moduli
        Dim configurazioniTotali As New List(Of ConfigurazioneModulo)

        For Each ctrl As Control In panelModuli.Controls
            If TypeOf ctrl Is ucConfiguratoreModulo Then
                Dim modulo As ucConfiguratoreModulo = DirectCast(ctrl, ucConfiguratoreModulo)
                configurazioniTotali.AddRange(modulo.GetConfigurazione())
            End If
        Next

        ' Esempio: stampa a video la configurazione
        Dim riepilogo As String = ""
        For Each conf In configurazioniTotali
            riepilogo &= $"Titolo: {conf.Titolo} - Scelta:'' - Q.tà: {conf.Quantita}" & vbCrLf
        Next

        MessageBox.Show(riepilogo, "Configurazione finale")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        inserisci_numero_nuova_revisione(N_opportunità, "STD", RichTextBox1.Text, valuta, Label3.Text)
        salva_costo(N_opportunità, numero_ultima_revisione + 1)

        MsgBox("Costo macchina aggiornato con successo")

    End Sub

    Sub inserisci_numero_nuova_revisione(par_opportunità As Integer, par_tipo As String, par_note As String, par_valuta As String, par_cambio As Decimal)

        par_cambio = Replace(par_cambio, ",", ".")
        trova_ultima_revisione(par_opportunità)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "

INSERT INTO [Tirelli_40].[dbo].[Superlistino_log_costificazioni_testata]
           ([Tipo]
      ,[Opportunità]
      ,[REV]
      ,[utente]
      ,[Data]
      ,[ora]
,note
,valuta
,cambio)
     VALUES
           ('" & par_tipo & "'
           ," & par_opportunità & "
           ," & numero_ultima_revisione & "+1
           ," & Homepage.ID_SALVATO & "
,getdate()
           ,convert(varchar, getdate(), 108)
,'" & par_note & "'
,'" & par_valuta & "',
'" & par_cambio & "')"


        CMD_SAP_3.ExecuteNonQuery()

        Cnn3.Close()


    End Sub



    Sub trova_ultima_revisione(par_numero_opportunità As Integer)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "
               Select  coalesce(t11.REV,0) as 'Ultima_rev'
,t11.utente, coalesce(CONCAT(T12.LASTNAME,' ',T12.FIRSTNAME),'-') as 'Nome_utente'
, t11.data,t11.ora
from
(
SELECT MAX(t0.id) as 'Ultimo_id'
     
  FROM [Tirelli_40].[dbo].[Superlistino_log_costificazioni_testata] t0
where t0.Opportunità ='" & par_numero_opportunità & "'

)
as t10 left join [Tirelli_40].[dbo].[Superlistino_log_costificazioni_testata] t11 on t10.ultimo_id=t11.id
left join [TIRELLI_40].[dbo].ohem t12 on t12.empid=t11.utente

"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            numero_ultima_revisione = cmd_SAP_reader("ultima_Rev")

            If Not cmd_SAP_reader("Data") Is System.DBNull.Value Then
                Label8.Text = cmd_SAP_reader("Data") & " | " & cmd_SAP_reader("ORA")
            Else
                Label8.Text = "-"
            End If


            Label9.Text = cmd_SAP_reader("Nome_utente")

        Else
            numero_ultima_revisione = -1
            Label8.Text = "-"
            Label9.Text = "-"
        End If

        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Sub salva_costo(par_N_opp As Integer, par_n_rev As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()



        ' Prima cancelliamo il record esistente
        Dim CMD_Delete As New SqlCommand("DELETE FROM [Tirelli_40].[dbo].[Superlistino_Log_costificazioni] 
WHERE [N_opp] = @N_opp and [N_rev] =@N_rev ", Cnn)
        CMD_Delete.Parameters.AddWithValue("@N_opp", par_N_opp)
        CMD_Delete.Parameters.AddWithValue("@N_rev", par_N_opp)
        CMD_Delete.ExecuteNonQuery()


        ' Cicla su tutti i controlli nel panel
        Dim conta_elementi As Integer = 1
        For Each ctrl As Control In panelModuli.Controls
            ' Se è un modulo valido

            If TypeOf ctrl Is ucConfiguratoreModulo Then
                Dim modulo As ucConfiguratoreModulo = CType(ctrl, ucConfiguratoreModulo)

                Dim codicePadre As String = modulo.Codice_macchina
                Dim Nome_Padre As String = modulo.TextBox_titolo_macchina.Text
                Dim quantity_Padre As Integer = modulo.txtQuantita.Text

                Dim CMD_Insert_0 As New SqlCommand("
                    INSERT INTO [Tirelli_40].[dbo].[Superlistino_Log_costificazioni_padri]
                        ([N_opp]
           ,[N_rev]
           ,[N_elemento]
,listino
           ,[Padre]
           ,[Nome_Padre]
           ,[Q_padre]
           ,[EU]
           ,[USA]
           ,[ADE]
           ,[Static]
           ,[Hot]
           ,[Flex])
                    VALUES 
                        (@N_opp, @N_rev, @N_elemento,@listino, @Padre, @Nome_padre,@Q_padre,@EU,@USA,@ADE,@Static,@Hot,@Flex)", Cnn)

                ' Aggiunta dei parametri
                CMD_Insert_0.Parameters.AddWithValue("@N_opp", par_N_opp)
                CMD_Insert_0.Parameters.AddWithValue("@N_rev", par_n_rev)
                CMD_Insert_0.Parameters.AddWithValue("@N_elemento", conta_elementi)
                CMD_Insert_0.Parameters.AddWithValue("@Listino", modulo.listino)
                CMD_Insert_0.Parameters.AddWithValue("@Padre", codicePadre)
                CMD_Insert_0.Parameters.AddWithValue("@Nome_padre", Nome_Padre)
                CMD_Insert_0.Parameters.AddWithValue("@Q_padre", quantity_Padre)
                CMD_Insert_0.Parameters.AddWithValue("@EU", modulo.Eu_)
                CMD_Insert_0.Parameters.AddWithValue("@USA", modulo.Usa_)
                CMD_Insert_0.Parameters.AddWithValue("@ADE", modulo.Ade_)
                CMD_Insert_0.Parameters.AddWithValue("@Static", modulo.Static_)
                CMD_Insert_0.Parameters.AddWithValue("@Hot", modulo.HOT_)
                CMD_Insert_0.Parameters.AddWithValue("@Flex", modulo.FLEX_)



                CMD_Insert_0.ExecuteNonQuery()


                If modulo.DataGridView4 IsNot Nothing Then
                    Dim conta_riga As Integer = 1

                    For Each riga As DataGridViewRow In modulo.DataGridView4.Rows
                        If Not riga.IsNewRow AndAlso riga.Cells.Count > 0 Then
                            Dim valore As Object = riga.Cells("Codice").Value
                            Dim descrizione As Object = riga.Cells("Desc").Value
                            Dim quantity As Decimal = riga.Cells("Q").Value
                            Dim costo_u As Decimal = riga.Cells("Costo_u").Value
                            Dim val_tot As Decimal = riga.Cells("prezzo").Value
                            Dim Tipo As String = riga.Cells("tipo").Value

                            Dim nota As String = Replace(riga.Cells("Note").Value, "'", "")

                            If nota = Nothing Then nota = ""
                            If valore IsNot Nothing Then
                                Debug.Print("Codice Padre: " & codicePadre)
                                Debug.Print("  - " & valore.ToString())

                                ' Inserimento del record
                                Dim CMD_Insert As New SqlCommand("
                    INSERT INTO [Tirelli_40].[dbo].[Superlistino_Log_costificazioni]
                        ([N_opp], [N_rev], [N_elemento], [Padre], [Nome_padre],[q_padre], [Riga], [Codice], [Descrizione], [Quantity], [Costo], [Val_tot],[tipo], [Note])
                    VALUES 
                        (@N_opp, @N_rev, @N_elemento, @Padre, @Nome_padre,@Q_padre, @Riga, @Codice, @Descrizione, @Quantity, @costo, @Val_tot,@tipo, @Note)", Cnn)

                                ' Aggiunta dei parametri
                                CMD_Insert.Parameters.AddWithValue("@N_opp", par_N_opp)
                                CMD_Insert.Parameters.AddWithValue("@N_rev", par_n_rev)
                                CMD_Insert.Parameters.AddWithValue("@N_elemento", conta_elementi)
                                CMD_Insert.Parameters.AddWithValue("@Padre", codicePadre)
                                CMD_Insert.Parameters.AddWithValue("@Nome_padre", Nome_Padre)
                                CMD_Insert.Parameters.AddWithValue("@Q_padre", quantity_Padre)
                                CMD_Insert.Parameters.AddWithValue("@Riga", conta_riga)
                                CMD_Insert.Parameters.AddWithValue("@Codice", valore.ToString())
                                CMD_Insert.Parameters.AddWithValue("@Descrizione", If(descrizione IsNot Nothing, descrizione.ToString(), ""))
                                CMD_Insert.Parameters.AddWithValue("@Quantity", quantity)
                                CMD_Insert.Parameters.AddWithValue("@Costo", costo_u)
                                CMD_Insert.Parameters.AddWithValue("@Val_tot", val_tot)
                                CMD_Insert.Parameters.AddWithValue("@Tipo", Tipo)
                                CMD_Insert.Parameters.AddWithValue("@Note", nota)


                                CMD_Insert.ExecuteNonQuery()
                                conta_riga += 1
                            End If
                        End If
                    Next
                End If
            End If
            conta_elementi += 1
        Next






        Cnn.Close()


    End Sub

    Private Sub Cmd_Cerca_Click(sender As Object, e As EventArgs) Handles Cmd_Cerca.Click
        N_opportunità = Txt_DocNum.Text
    End Sub

    Private Sub Txt_DocNum_TextChanged(sender As Object, e As EventArgs) Handles Txt_DocNum.TextChanged
        N_opportunità = Txt_DocNum.Text
    End Sub

    Public Sub informazioni_testata(par_N_opp As String, par_n_rev As String)
        panelModuli.Controls.Clear()

        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Cnn.Open()

            Dim query As String = "SELECT TOP (1000) t0.[ID]
      ,t0.[Tipo]
      ,t0.[Opportunità]
,coalesce(T1.cardname,'') as 'Cardname'
      ,t0.[REV]
      ,t0.[Valuta]
      ,t0.[Cambio]
      ,t0.[utente]
,coalesce(concat(t2.lastname,' ',t2.firstname),'') as 'Nome_utente'
      ,t0.[Data]
      ,t0.[ora]
      ,t0.[Note]
  FROM [Tirelli_40].[dbo].[Superlistino_log_costificazioni_testata] t0
left join oopr t1 on t1.opprid=t0.[Opportunità]
left join [TIRELLI_40].[dbo].ohem t2 on t2.empid=t0.utente

        WHERE t0.[Opportunità] = @N_opp AND t0.[REV] = @N_rev
        
        "

            Dim cmd As New SqlCommand(query, Cnn)
            cmd.Parameters.AddWithValue("@N_opp", par_N_opp)
            cmd.Parameters.AddWithValue("@N_rev", par_n_rev)

            Dim reader As SqlDataReader = cmd.ExecuteReader()

            If reader.Read() Then
                valuta = reader("Valuta")
                If valuta = "E" Then
                    RadioButton1.Checked = True
                    RadioButton2.Checked = False


                Else
                    RadioButton2.Checked = True
                    RadioButton1.Checked = False


                End If

                Txt_DocNum.Text = reader("Opportunità")
                TextBox1.Text = reader("REV")
                Label2.Text = reader("Cardname")
                Label5.Text = reader("nome_utente")
                Label6.Text = reader("Data") & " | " & reader("ora")
                RichTextBox1.Text = reader("Note")

            End If

            reader.Close()




        End Using

    End Sub

    Public Sub informazioni_testata_nuova(par_N_opp As String)
        panelModuli.Controls.Clear()

        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Cnn.Open()

            Dim query As String = "SELECT 
      
coalesce(T1.cardname,'') as 'Cardname'
   ,coalesce(concat(t2.firstname,' ', t2.lastname),'') as 'Nome_utente'
  FROM oopr t1
left join [TIRELLI_40].[dbo].ohem t2 on t2.empid=" & Homepage.ID_SALVATO & "

        WHERE t1.opprid = @N_opp 
        
        "

            Dim cmd As New SqlCommand(query, Cnn)
            cmd.Parameters.AddWithValue("@N_opp", par_N_opp)


            Dim reader As SqlDataReader = cmd.ExecuteReader()

            If reader.Read() Then




                Label2.Text = reader("Cardname")
                Label5.Text = reader("nome_utente")
                Label6.Text = Now


            End If

            reader.Close()




        End Using

    End Sub

    Public Sub ricrea_form_da_database_new(par_N_opp As String, par_n_rev As String)
        informazioni_testata(par_N_opp, par_n_rev)
        panelModuli.Controls.Clear()

        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Cnn.Open()

            Dim query As String = "SELECT TOP (1000) [ID]
      ,[N_opp]
      ,[N_rev]
      ,[N_elemento]
,listino
      ,[Padre]
      ,[Nome_Padre]
      ,[Q_padre]
      ,[EU]
      ,[USA]
      ,[ADE]
      ,[Static]
      ,[Hot]
      ,[Flex]
  FROM [Tirelli_40].[dbo].[Superlistino_Log_costificazioni_padri] t0

        WHERE t0.N_opp = @N_opp AND t0.N_rev = @N_rev
        ORDER BY t0.N_elemento DESC
        "

            Dim cmd As New SqlCommand(query, Cnn)
            cmd.Parameters.AddWithValue("@N_opp", par_N_opp)
            cmd.Parameters.AddWithValue("@N_rev", par_n_rev)

            Dim reader As SqlDataReader = cmd.ExecuteReader()

            ' Dati raggruppati per modulo (Padre)
            Dim moduliDict As New Dictionary(Of String, ucConfiguratoreModulo)

            While reader.Read()
                Dim padre As String = reader("Padre").ToString()
                Dim Nome_padre As String = reader("Nome_Padre").ToString()



                ' Se non esiste ancora un modulo per questo padre, lo creiamo
                If Not moduliDict.ContainsKey(padre) Then
                    Dim modulo As New ucConfiguratoreModulo With {
                    .Codice_macchina = padre
                }
                    modulo.TextBox_titolo_macchina.Text = Nome_padre
                    modulo.txtQuantita.Text = reader("Q_Padre")
                    modulo.Label1.Text = reader("N_elemento")
                    modulo.Eu_ = reader("EU")
                    modulo.Usa_ = reader("USA")
                    modulo.Ade_ = reader("ADE")
                    modulo.Static_ = reader("Static")
                    modulo.HOT_ = reader("Hot")
                    modulo.FLEX_ = reader("Flex")

                    If modulo.Eu_ = "Y" Then
                        modulo.RadioButton1.Checked = True
                    Else
                        modulo.RadioButton1.Checked = False
                    End If

                    If modulo.Usa_ = "Y" Then
                        modulo.RadioButton2.Checked = True
                    Else
                        modulo.RadioButton2.Checked = False
                    End If
                    Dim ade_boolean As Boolean
                    Dim hot_boolean As Boolean
                    Dim static_boolean As Boolean
                    Dim flex_boolean As Boolean

                    ' Esegui il controllo per ciascuna variabile
                    If modulo.Ade_ = "Y" Then
                        ade_boolean = True
                    Else
                        ade_boolean = False
                    End If

                    If modulo.HOT_ = "Y" Then
                        hot_boolean = True
                    Else
                        hot_boolean = False
                    End If

                    If modulo.FLEX_ = "Y" Then
                        flex_boolean = True
                    Else
                        flex_boolean = False
                    End If

                    If modulo.Static_ = "Y" Then
                        static_boolean = True
                    Else
                        static_boolean = False
                    End If



                    riempi_datagridview_da_modulo(modulo, par_N_opp, par_n_rev, reader("N_elemento"), modulo.Label2, modulo.Label3, modulo.Label4, valuta, Label3.Text)
                    '   AdattaAltezzaDataGridView(DataGridView4)
                    modulo.compila_distinta(modulo.DataGridView1, padre, "Y", True, modulo.RadioButton1.Checked, modulo.RadioButton2.Checked, ade_boolean, static_boolean, hot_boolean, modulo.Label2, modulo.Label4, modulo.Label3, reader("Listino"), valuta, Label3.Text)
                    moduliDict.Add(padre, modulo)



                End If


            End While

            reader.Close()


            ' Aggiunge i moduli al form
            For Each modulo In moduliDict.Values
                panelModuli.Controls.Add(modulo)
                modulo.Width = panelModuli.Width - 15
                panelModuli.Controls.SetChildIndex(modulo, 0)
            Next
        End Using
    End Sub

    Sub riempi_datagridview_da_modulo(par_ucConfiguratoreModulo As ucConfiguratoreModulo, par_n_opp As Integer, par_n_rev As Integer, par_n_elemento As Integer, par_label_costo As Label, par_label_molt As Label, par_label_prezzo As Label, par_valuta As String, par_cambio As String)
        DataGridView4.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim Valore_costo_tot As Decimal = 0
        Dim Valore_prezzo_tot As Decimal = 0

        Dim Valore_costo_tot_usd As Decimal = 0
        Dim Valore_prezzo_tot_usd As Decimal = 0


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "
        SELECT 
            t1.id,
            t0.[Codice],
            t0.[Descrizione],
            t0.Quantity,
            t1.Costo,
            t0.quantity * t1.costo AS Costo_tot,
            t0.[Val_tot],
            t1.Costificato,
            t0.Note, t0.tipo,
            COALESCE(t1.Immagine, '') AS Immagine
        FROM [Tirelli_40].[dbo].[Superlistino_Log_costificazioni] t0 
        LEFT JOIN [Tirelli_40].[dbo].Superlistino_codici t1 ON t0.Codice = t1.Codice
        WHERE t0.N_opp = " & par_n_opp & " AND t0.N_rev = " & par_n_rev & " AND t0.N_elemento = " & par_n_elemento & "
        ORDER by t0.Riga

"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()
            Dim id As String = cmd_SAP_reader("ID").ToString()
            Dim codice As String = cmd_SAP_reader("Codice").ToString()
            Dim descrizione As String = cmd_SAP_reader("Descrizione").ToString()
            Dim qty As Integer = Convert.ToInt32(cmd_SAP_reader("Quantity"))
            Dim costo As Decimal = Convert.ToDecimal(cmd_SAP_reader("Costo"))
            Dim costoTot As Decimal = Convert.ToDecimal(cmd_SAP_reader("Costo_tot"))
            Dim costoTot_usd As Decimal = Convert.ToDecimal(cmd_SAP_reader("Costo_tot") * par_cambio)
            Dim prezzo_tot As Decimal = Convert.ToDecimal(cmd_SAP_reader("Val_tot"))
            Dim prezzo_tot_usd As Decimal = Convert.ToDecimal(cmd_SAP_reader("Val_tot") * par_cambio)
            Dim costificato As String = cmd_SAP_reader("Costificato").ToString()
            Dim note As String = cmd_SAP_reader("Note").ToString()
            Dim tipo As String = cmd_SAP_reader("Tipo").ToString()
            Valore_costo_tot += costoTot
            Valore_prezzo_tot += prezzo_tot
            Valore_costo_tot_usd += costoTot_usd
            Valore_prezzo_tot_usd += prezzo_tot_usd
            Dim immagineFile As String = Homepage.Percorso_Immagini_TICKETS & cmd_SAP_reader("Immagine").ToString()
            If Not File.Exists(immagineFile) Then
                immagineFile = Homepage.Percorso_Immagini_TICKETS & "Bianco.jpg"
            End If

            ' Carica l'immagine ridimensionata
            Dim image As Image = Image.FromFile(immagineFile)
            Dim maxHeight As Integer = 60
            Dim scaleFactor As Double = maxHeight / image.Height
            Dim newWidth As Integer = CInt(image.Width * scaleFactor)
            Dim smallImage As New Bitmap(image, New Size(newWidth, maxHeight))



            ' Aggiunge la riga al modulo corrispondente
            With par_ucConfiguratoreModulo.DataGridView4
                .Rows.Add(id, codice, descrizione, qty, costo, costoTot, If(costoTot <> 0, prezzo_tot / costoTot, 0), prezzo_tot, prezzo_tot_usd, costificato, note, smallImage, tipo)
            End With


        Loop
        If par_valuta = "E" Then


            par_label_costo.Text = Valore_costo_tot.ToString("C0")

            par_label_prezzo.Text = Valore_prezzo_tot.ToString("C0")
        Else
            par_label_costo.Text = Valore_costo_tot_usd.ToString("C0", Globalization.CultureInfo.GetCultureInfo("en-US"))

            par_label_prezzo.Text = Valore_prezzo_tot_usd.ToString("C0", Globalization.CultureInfo.GetCultureInfo("en-US"))
        End If
        Try
            par_label_molt.Text = (Valore_prezzo_tot / Valore_costo_tot).ToString("N2")
        Catch ex As Exception

        End Try
        cmd_SAP_reader.Close()
        Cnn.Close()
        DataGridView4.Rows.Insert(0, par_n_elemento, par_ucConfiguratoreModulo.TextBox_titolo_macchina.Text, par_ucConfiguratoreModulo.txtQuantita.Text, par_label_molt.Text, Valore_prezzo_tot.ToString("C0"), Valore_prezzo_tot_usd.ToString("C0", Globalization.CultureInfo.GetCultureInfo("en-US")))
    End Sub



    Private Sub AdattaAltezzaDataGridView(dgv As DataGridView)
        Dim rigaAltezza As Integer = dgv.RowTemplate.Height
        Dim intestazioneAltezza As Integer = dgv.ColumnHeadersHeight
        Dim righeVisibili As Integer = dgv.RowCount

        dgv.Height = intestazioneAltezza + (rigaAltezza * righeVisibili) + 2 ' +2 per margini
    End Sub

    Public Sub ricrea_form_da_database(par_N_opp As String, par_n_rev As String)
        panelModuli.Controls.Clear()

        Using Cnn As New SqlConnection(Homepage.sap_tirelli)
            Cnn.Open()

            Dim query As String = "
        SELECT 
            t0.[Padre],
            t0.[Nome_Padre],
            t1.id,
            t0.[Codice],
            t0.[Descrizione],
            t0.Quantity,
            t1.Costo,
            t0.quantity * t1.costo AS Costo_tot,
            t0.[Val_tot],
            t1.Costificato,
            t0.Note, t0.tipo,
            COALESCE(t1.Immagine, '') AS Immagine
        FROM [Tirelli_40].[dbo].[Superlistino_Log_costificazioni] t0 
        LEFT JOIN [Tirelli_40].[dbo].Superlistino_codici t1 ON t0.Codice = t1.Codice
        WHERE t0.N_opp = @N_opp AND t0.N_rev = @N_rev
        ORDER BY t0.N_elemento DESC, t0.Riga
        "

            Dim cmd As New SqlCommand(query, Cnn)
            cmd.Parameters.AddWithValue("@N_opp", par_N_opp)
            cmd.Parameters.AddWithValue("@N_rev", par_n_rev)

            Dim reader As SqlDataReader = cmd.ExecuteReader()

            ' Dati raggruppati per modulo (Padre)
            Dim moduliDict As New Dictionary(Of String, ucConfiguratoreModulo)
            Dim costoPerPadre As New Dictionary(Of String, Decimal)
            Dim prezzoPerPadre As New Dictionary(Of String, Decimal)

            While reader.Read()
                Dim padre As String = reader("Padre").ToString()
                Dim Nome_padre As String = reader("Nome_Padre").ToString()
                Dim id As String = reader("ID").ToString()
                Dim codice As String = reader("Codice").ToString()
                Dim descrizione As String = reader("Descrizione").ToString()
                Dim qty As Integer = Convert.ToInt32(reader("Quantity"))
                Dim costo As Decimal = Convert.ToDecimal(reader("Costo"))
                Dim costoTot As Decimal = Convert.ToDecimal(reader("Costo_tot"))
                Dim prezzo_tot As Decimal = Convert.ToDecimal(reader("Val_tot"))
                Dim costificato As String = reader("Costificato").ToString()
                Dim note As String = reader("Note").ToString()
                Dim tipo As String = reader("Tipo").ToString()

                Dim immagineFile As String = Homepage.Percorso_Immagini_TICKETS & reader("Immagine").ToString()
                If Not File.Exists(immagineFile) Then
                    immagineFile = Homepage.Percorso_Immagini_TICKETS & "Bianco.jpg"
                End If

                ' Carica l'immagine ridimensionata
                Dim image As Image = Image.FromFile(immagineFile)
                Dim maxHeight As Integer = 60
                Dim scaleFactor As Double = maxHeight / image.Height
                Dim newWidth As Integer = CInt(image.Width * scaleFactor)
                Dim smallImage As New Bitmap(image, New Size(newWidth, maxHeight))

                ' Se non esiste ancora un modulo per questo padre, lo creiamo
                If Not moduliDict.ContainsKey(padre) Then
                    Dim modulo As New ucConfiguratoreModulo With {
                    .Codice_macchina = padre
                }
                    modulo.TextBox_titolo_macchina.Text = Nome_padre
                    modulo.txtQuantita.Text = qty
                    moduliDict.Add(padre, modulo)
                End If

                ' Aggiunge la riga al modulo corrispondente
                With moduliDict(padre).DataGridView4
                    .Rows.Add(id, codice, descrizione, qty, costo, costoTot, If(costoTot <> 0, prezzo_tot / costoTot, 0), prezzo_tot, costificato, note, smallImage, tipo)
                End With

                ' Aggiorna i totali per modulo
                If Not costoPerPadre.ContainsKey(padre) Then
                    costoPerPadre(padre) = 0
                    prezzoPerPadre(padre) = 0
                End If
                costoPerPadre(padre) += costoTot
                prezzoPerPadre(padre) += prezzo_tot
            End While

            reader.Close()

            ' Ora aggiorna le label per ogni modulo
            For Each kvp In moduliDict
                Dim padre As String = kvp.Key
                Dim modulo As ucConfiguratoreModulo = kvp.Value

                If costoPerPadre.ContainsKey(padre) Then
                    modulo.Label2.Text = costoPerPadre(padre).ToString("N2")
                End If
                If prezzoPerPadre.ContainsKey(padre) Then
                    modulo.Label4.Text = prezzoPerPadre(padre).ToString("N2")
                End If

                Try
                    If costoPerPadre(padre) <> 0 Then
                        modulo.Label3.Text = (prezzoPerPadre(padre) / costoPerPadre(padre)).ToString("0.00")
                    Else
                        modulo.Label3.Text = "0.00"
                    End If
                Catch ex As Exception
                    modulo.Label3.Text = "0.00"
                End Try
            Next

            ' Aggiunge i moduli al form
            For Each modulo In moduliDict.Values
                panelModuli.Controls.Add(modulo)
                modulo.Width = panelModuli.Width - 15
                panelModuli.Controls.SetChildIndex(modulo, 0)
            Next
        End Using
    End Sub



    Private Sub panelModuli_SizeChanged(sender As Object, e As EventArgs) Handles panelModuli.SizeChanged
        If panelModuli.Controls.Count = 0 Then Exit Sub

        For Each ctrl As Control In panelModuli.Controls
            If TypeOf ctrl Is ucConfiguratoreModulo Then
                ctrl.Width = panelModuli.Width - 15
            End If
        Next
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub Form_configuratore_vendita_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        AddHandler DataGridView4.RowsAdded, AddressOf DataGridView4_RigheModificate
        AddHandler DataGridView4.RowsRemoved, AddressOf DataGridView4_RigheModificate
        AddHandler DataGridView4.DataSourceChanged, AddressOf DataGridView4_RigheModificate

    End Sub

    Private Sub DataGridView4_RigheModificate(sender As Object, e As EventArgs)
        AdattaAltezzaDataGridView(DataGridView4)
    End Sub

    Private Sub DataGridView4_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellValueChanged
        ' Verifica che la colonna modificata sia "prezzo"
        If DataGridView4.Columns(e.ColumnIndex).Name = "Prezzo" Then
            AggiornaTotalePrezzi()
        End If
    End Sub
    Private Sub DataGridView4_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridView4.RowsAdded
        AggiornaTotalePrezzi()
    End Sub

    Private Sub DataGridView4_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles DataGridView4.RowsRemoved
        AggiornaTotalePrezzi()
    End Sub

    Private Sub AggiornaTotalePrezzi()
        Dim totale As Decimal = 0
        Dim totale_usd As Decimal = 0

        For Each row As DataGridViewRow In DataGridView4.Rows
            If Not row.IsNewRow Then
                Dim valore As Object = row.Cells("prezzo").Value
                Dim valore_usd As Object = row.Cells("prezzo_usd").Value
                If valore IsNot Nothing Then
                    Dim testoPulito As String = valore.ToString()
                    Dim testoPulito_usd As String = valore_usd.ToString()

                    ' Rimuove simboli di valuta, spazi e punti (separatori di migliaia)
                    testoPulito = testoPulito.Replace("€", "").Replace(" ", "").Replace(".", "").Replace(",", ".")
                    testoPulito_usd = testoPulito_usd.Replace("$", "").Replace(" ", "").Replace(".", "").Replace(",", ".")

                    Dim prezzoDecimal As Decimal
                    Dim prezzoDecimal_usd As Decimal
                    If Decimal.TryParse(testoPulito, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, prezzoDecimal) Then
                        totale += prezzoDecimal
                    End If

                    If Decimal.TryParse(testoPulito_usd, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, prezzoDecimal_usd) Then
                        totale_usd += prezzoDecimal_usd
                    End If
                End If
            End If
        Next

        Label1.Text = totale.ToString("N2") & " €"
        If valuta = "$" Then
            Label4.Text = totale_usd.ToString("N2") & " $"
        End If

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            Label3.Text = 1
            valuta = "E"
            DataGridView4.Columns("Prezzo_usd").Visible = False
            GroupBox9.Visible = False
        Else
            Label3.Text = 1.25
            valuta = "$"
            DataGridView4.Columns("Prezzo_usd").Visible = True
            Label4.Text = (Val(Label1.Text) * Val(Label3.Text)).ToString("C2", Globalization.CultureInfo.CreateSpecificCulture("en-US"))
            GroupBox9.Visible = True
        End If

        For Each row As DataGridViewRow In DataGridView4.Rows
            If Not row.IsNewRow Then
                Dim cellVal As Object = row.Cells("Prezzo").Value
                If cellVal IsNot Nothing Then
                    ' Rimuove simboli di valuta e spazi
                    Dim rawText As String = cellVal.ToString().Replace("€", "").Replace("$", "").Replace(" ", "").Trim()

                    Dim prezzoOriginale As Decimal
                    If Decimal.TryParse(rawText, prezzoOriginale) Then
                        Dim moltiplicatore As Decimal = Decimal.Parse(Label3.Text)
                        row.Cells("Prezzo_usd").Value = prezzoOriginale * moltiplicatore
                    Else
                        MsgBox("Valore non numerico: " & rawText)
                    End If
                End If
            End If
        Next

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton1.Checked = True Then
            Label3.Text = 1
            valuta = "E"
        Else
            Label3.Text = 1.25
            valuta = "$"
        End If
    End Sub
End Class

