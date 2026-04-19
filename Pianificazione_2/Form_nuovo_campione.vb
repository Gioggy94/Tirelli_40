Imports System.Data.SqlClient
Imports System.Reflection.Emit
Imports System.Windows.Documents


Public Class Form_nuovo_campione

    Public Codice_BP_selezionato As String
    Public Codice_BP_JG_selezionato As String

    Public Codice_BP As String
    Public Codice_BP_finale As String
    Public Elenco_Tipo_Campioni(1000) As Integer
    Public Id_Campione As Integer
    Public Immagine_Caricata As Integer
    Public Blocco_univocità As String
    Public nome_bp_selezionato As String
    Private blocco_scheda As Integer
    Public numero_combinazioni As Integer
    Private num_collaudati As Integer

    Private Sub Form_nuovo_campione_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        riempi_combobox_tipo_campione()

    End Sub

    Sub inizializza_form()
        GroupBox2.Visible = False
        Label1.Text = "-"
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Business_partner.Show()
        Business_partner.Provenienza = "Form_nuovo_campione"
    End Sub

    Sub riempi_combobox_tipo_campione()
        Combo_tipo_campione.Items.Clear()
        Dim Indice As Integer

        Dim Cnn_Tipo As New SqlConnection


        Cnn_Tipo.ConnectionString = Homepage.sap_tirelli
        Cnn_Tipo.Open()

        Dim Cmd_Tipo As New SqlCommand
        Dim Cmd_Tipo_Reader As SqlDataReader

        Indice = 0
        Cmd_Tipo.Connection = Cnn_Tipo
        Cmd_Tipo.CommandText = " SELECT * FROM [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE
ORDER BY Id_Tipo_Campione"
        Cmd_Tipo_Reader = Cmd_Tipo.ExecuteReader
        Combo_tipo_campione.Items.Add("")
        Indice = Indice + 1
        Do While Cmd_Tipo_Reader.Read()
            Combo_tipo_campione.Items.Add(Cmd_Tipo_Reader("Descrizione"))
            Elenco_Tipo_Campioni(Indice) = Cmd_Tipo_Reader("Id_Tipo_Campione")
            Indice = Indice + 1
        Loop

        Cmd_Tipo_Reader.Close()
        Cnn_Tipo.Close()

    End Sub

    Private Sub Combo_tipo_campione_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_tipo_campione.SelectedIndexChanged

        If Combo_tipo_campione.SelectedIndex = 1 Then
            TabControl1.SelectedTab = Flacone


        ElseIf Combo_tipo_campione.SelectedIndex = 2 Then
            TabControl1.SelectedTab = Tappo

        ElseIf Combo_tipo_campione.SelectedIndex = 3 Then
            TabControl1.SelectedTab = Sottotappo
        ElseIf Combo_tipo_campione.SelectedIndex = 4 Then
            TabControl1.SelectedTab = Pompetta

        ElseIf Combo_tipo_campione.SelectedIndex = 5 Then
            TabControl1.SelectedTab = Etichetta


        ElseIf Combo_tipo_campione.SelectedIndex = 6 Then
            TabControl1.SelectedTab = Trigger

        ElseIf Combo_tipo_campione.SelectedIndex = 7 Then
            TabControl1.SelectedTab = Prodotto
        ElseIf Combo_tipo_campione.SelectedIndex = 8 Then

            TabControl1.SelectedTab = Film
        ElseIf Combo_tipo_campione.SelectedIndex = 9 Then
            TabControl1.SelectedTab = Copritappo

        ElseIf Combo_tipo_campione.SelectedIndex = 10 Then
            TabControl1.SelectedTab = Scatola


        End If
        If Combo_tipo_campione.SelectedIndex > 0 Then
            TabControl1.Visible = True
        Else
            TabControl1.Visible = False
        End If





        Dim Cnn_Tipo As New SqlConnection
        Cnn_Tipo.ConnectionString = Homepage.sap_tirelli
        Cnn_Tipo.Open()

        Dim Cmd_Tipo As New SqlCommand
        Dim Cmd_Tipo_Reader As SqlDataReader

        Cmd_Tipo.Connection = Cnn_Tipo
        Cmd_Tipo.CommandText = " SELECT * FROM [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE WHERE Id_Tipo_Campione=" & Combo_tipo_campione.SelectedIndex + 99 & ""
        Cmd_Tipo_Reader = Cmd_Tipo.ExecuteReader
        Cmd_Tipo_Reader.Read()

        Txt_Sigla.Text = Cmd_Tipo_Reader("Iniziale_Sigla")

        If Cmd_Tipo_Reader("Immagine_Descrizione").ToString.Length > 1 Then
            Dim MyImage As Bitmap
            img_descrizione.SizeMode = PictureBoxSizeMode.Zoom
            Try
                MyImage = New Bitmap(Homepage.Percorso_immagini & Cmd_Tipo_Reader("Immagine_Descrizione").ToString)
            Catch ex As Exception
                MsgBox("Impossibile Aprire l'Immagine d'esempio Selezionata")
            End Try
            img_descrizione.Image = CType(MyImage, Image)
        Else
            img_descrizione.Image = Nothing
        End If

        Cmd_Tipo_Reader.Close()
        Cnn_Tipo.Close()

        If Label1.Text <> "-" Then
            TableLayoutPanel7.Visible = True

            Picture_Campione.Visible = True
            TableLayoutPanel6.Visible = True
        End If
    End Sub

    Private Sub flacone_Click(sender As Object, e As EventArgs) Handles Flacone.Enter

        Combo_tipo_campione.SelectedIndex = 1


    End Sub

    Private Sub Scatola_Click(sender As Object, e As EventArgs) Handles Scatola.Enter

        Combo_tipo_campione.SelectedIndex = 10


    End Sub
    Private Sub tappo_Click(sender As Object, e As EventArgs) Handles Tappo.Enter

        Combo_tipo_campione.SelectedIndex = 2


    End Sub
    Private Sub sottotappo_Click(sender As Object, e As EventArgs) Handles Sottotappo.Enter

        Combo_tipo_campione.SelectedIndex = 3


    End Sub
    Private Sub pompetta_Click(sender As Object, e As EventArgs) Handles Pompetta.Enter

        Combo_tipo_campione.SelectedIndex = 4


    End Sub
    Private Sub etichetta_Click(sender As Object, e As EventArgs) Handles Etichetta.Enter

        Combo_tipo_campione.SelectedIndex = 5


    End Sub
    Private Sub trigger_Click(sender As Object, e As EventArgs) Handles Trigger.Enter

        Combo_tipo_campione.SelectedIndex = 6


    End Sub
    Private Sub prodotto_Click(sender As Object, e As EventArgs) Handles Prodotto.Enter

        Combo_tipo_campione.SelectedIndex = 7


    End Sub
    Private Sub film_Click(sender As Object, e As EventArgs) Handles Film.Enter

        Combo_tipo_campione.SelectedIndex = 8


    End Sub
    Private Sub copritappo_Click(sender As Object, e As EventArgs) Handles Copritappo.Enter

        Combo_tipo_campione.SelectedIndex = 9


    End Sub


    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Label1_TextChanged(sender As Object, e As EventArgs) Handles Label1.TextChanged
        If Label1.Text = "-" Then
            GroupBox2.Visible = False
            TableLayoutPanel7.Visible = False
            TabControl1.Visible = False
            Picture_Campione.Visible = False
            TableLayoutPanel6.Visible = False


        Else
            GroupBox2.Visible = True
            If Combo_tipo_campione.SelectedIndex > 0 Then

                TableLayoutPanel7.Visible = True

                Picture_Campione.Visible = True
                TableLayoutPanel6.Visible = True

            End If

            riempi_datagridview_campioni(Codice_BP_selezionato, Codice_BP)
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click



        If Txt_nome.Text.Length = 0 Then
                MsgBox("Inserire un Nome Campione")
            Else

            If Codice_BP_JG_selezionato = Nothing Or Codice_BP_JG_selezionato = "" Then
                MsgBox("Selezionare un codice business partner")
            Else



                check_univocità_campione(Txt_nome.Text, Elenco_Tipo_Campioni(Combo_tipo_campione.SelectedIndex), Codice_BP_selezionato, Codice_BP_JG_selezionato, Id_Campione)

                If Blocco_univocità = "Y" Then
                        MsgBox("Questo campione per questo cliente risulta già")

                    Else
                        trova_id_campione()

                        If Combo_tipo_campione.SelectedIndex = 1 Then

                            inserisci_flacone(Id_Campione)

                        ElseIf Combo_tipo_campione.SelectedIndex = 2 Then
                            inserisci_tappo(Id_Campione)
                        ElseIf Combo_tipo_campione.SelectedIndex = 3 Then

                            inserisci_sottotappo(Id_Campione)

                        ElseIf Combo_tipo_campione.SelectedIndex = 4 Then
                            inserisci_pompetta(Id_Campione)
                        ElseIf Combo_tipo_campione.SelectedIndex = 5 Then
                            inserisci_etichetta(Id_Campione)
                        ElseIf Combo_tipo_campione.SelectedIndex = 6 Then
                            inserisci_trigger(Id_Campione)
                        ElseIf Combo_tipo_campione.SelectedIndex = 7 Then
                            inserisci_prodotto(Id_Campione)
                        ElseIf Combo_tipo_campione.SelectedIndex = 8 Then
                            inserisci_film(Id_Campione)
                        ElseIf Combo_tipo_campione.SelectedIndex = 9 Then
                            inserisci_copritappo(Id_Campione)
                        ElseIf Combo_tipo_campione.SelectedIndex = 10 Then
                            inserisci_Scatola(Id_Campione)
                        End If
                        inserisci_dati_generici_campione(Id_Campione)

                        MsgBox("Campione inserito con successo")
                        pulisci_form()
                    Scheda_tecnica.riempi_datagridview_campioni(DataGridView3, Codice_BP_selezionato, Codice_BP, "", Homepage.Percorso_immagini, Homepage.sap_tirelli)
                    Scheda_tecnica.riempi_datagridview_campioni(Scheda_tecnica.DataGridView3, Scheda_tecnica.codice_bp_campione, Scheda_tecnica.bp_code, Scheda_tecnica.final_bp_code, Homepage.Percorso_immagini, Homepage.sap_tirelli)
                    Scheda_tecnica.riempi_datagridview_combinazioni(Scheda_tecnica.DataGridView1, Scheda_tecnica.codice_commessa, Homepage.sap_tirelli)
                End If
                End If


            End If


    End Sub

    Sub pulisci_form()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox22.Text = ""
        TextBox22.Text = ""
        TextBox23.Text = ""
        TextBox24.Text = ""
        TextBox25.Text = ""
        TextBox26.Text = ""
        TextBox27.Text = ""
        TextBox28.Text = ""
        TextBox29.Text = ""
        TextBox30.Text = ""
        TextBox31.Text = ""
        TextBox32.Text = ""
        TextBox33.Text = ""
        TextBox34.Text = ""
        TextBox36.Text = ""
        TextBox37.Text = ""
        TextBox38.Text = ""
        TextBox39.Text = ""
        TextBox40.Text = ""
        TextBox41.Text = ""
        TextBox42.Text = ""
        TextBox43.Text = ""
        TextBox44.Text = ""
        TextBox45.Text = ""
        TextBox46.Text = ""
        TextBox47.Text = ""
        TextBox48.Text = ""
        TextBox49.Text = ""
        TextBox50.Text = ""
        TextBox51.Text = ""
        TextBox52.Text = ""
        TextBox53.Text = ""
        TextBox54.Text = ""
        TextBox55.Text = ""
        TextBox56.Text = ""
        TextBox57.Text = ""
        TextBox58.Text = ""
        TextBox59.Text = ""
        TextBox60.Text = ""
        TextBox61.Text = ""
        TextBox62.Text = ""
        TextBox63.Text = ""
        TextBox64.Text = ""
        TextBox65.Text = ""
        TextBox66.Text = ""
        TextBox67.Text = ""
        TextBox68.Text = ""
        TextBox69.Text = ""
        TextBox70.Text = ""
        TextBox71.Text = ""
        TextBox72.Text = ""
        TextBox73.Text = ""
        TextBox74.Text = ""
        TextBox75.Text = ""
        TextBox76.Text = ""
        TextBox77.Text = ""
        TextBox78.Text = ""
        TextBox79.Text = ""
        TextBox80.Text = ""
        TextBox81.Text = ""
        TextBox82.Text = ""
        TextBox83.Text = ""
        TextBox84.Text = ""
        TextBox85.Text = ""
        TextBox86.Text = ""
        TextBox87.Text = ""
        TextBox88.Text = ""
        TextBox89.Text = ""
        TextBox90.Text = ""
        TextBox91.Text = ""
        TextBox92.Text = ""
        TextBox93.Text = ""
        TextBox94.Text = ""
        TextBox95.Text = ""
        TextBox96.Text = ""
        TextBox97.Text = ""
        TextBox98.Text = ""
        TextBox99.Text = ""
        TextBox100.Text = ""
        TextBox101.Text = ""
        TextBox102.Text = ""
        TextBox103.Text = ""
        TextBox104.Text = ""
        TextBox105.Text = ""
        TextBox106.Text = ""
        TextBox107.Text = ""
        TextBox108.Text = ""
        TextBox109.Text = ""
        TextBox110.Text = ""
        TextBox111.Text = ""
        TextBox112.Text = ""



        ComboBox6.SelectedIndex = 0
        ComboBox1.SelectedIndex = 0
        ComboBox2.SelectedIndex = 0
        ComboBox3.SelectedIndex = 0
        ComboBox4.SelectedIndex = 0
        ComboBox5.SelectedIndex = 0
    End Sub

    Sub inserisci_dati_generici_campione(par_id_campione)
        Dim Cnn_Campioni As New SqlConnection
        Cnn_Campioni.ConnectionString = Homepage.sap_tirelli
        Cnn_Campioni.Open()

        Dim Cmd_Campioni As New SqlCommand
        Dim Stringa_Immagine As String



        If Immagine_Caricata = 1 Then
            '
            'Stringa_Immagine = Percorso_Immagine & Txt_Codice_BP.Text & "_" & Txt_Sigla.Text & Txt_Nome.Text & "_" & Txt_Id_Campione.Text & ".jpg"
            Stringa_Immagine = par_id_campione & ".jpg"
            Dim contenuto_immagine As String
            contenuto_immagine = Homepage.Percorso_immagini & Stringa_Immagine
            Picture_Campione.Image.Save(contenuto_immagine)
        Else
            Stringa_Immagine = ""
        End If
        trova_id_campione()
        Cmd_Campioni.Connection = Cnn_Campioni
        Cmd_Campioni.CommandText = "INSERT INTO [TIRELLI_40].[DBO].coll_campioni (Id_Campione,Codice_BP,CODICE_BP_GALILEO,Nome,Descrizione,Tipo_Campione,Immagine,insertdate,updatedate,ownerid,note)
VALUES (" & par_id_campione & ",'" & Codice_BP_selezionato & "','" & Codice_BP_JG_selezionato & "','" & Txt_nome.Text & "','" & Replace(Txt_descrizione.Text, "'", " ") & "'," & Elenco_Tipo_Campioni(Combo_tipo_campione.SelectedIndex) & ",'" & Stringa_Immagine & "',getdate(),getdate(),'" & Homepage.ID_SALVATO & "', '" & RichTextBox1.Text & "')"

        Cmd_Campioni.ExecuteNonQuery()
        Cnn_Campioni.Close()


        Txt_nome.Text = ""

        Txt_descrizione.Text = ""


        Picture_Campione.Image = Nothing
        Immagine_Caricata = 0



    End Sub

    Public Sub check_univocità_campione(par_nome As String, par_tipo As Integer, par_cardcode As String, par_cardcode_jgal As String, par_id_campione As String)
        Blocco_univocità = "N"
        Dim Cnn_BP As New SqlConnection
        Cnn_BP.ConnectionString = Homepage.sap_tirelli
        Cnn_BP.Open()

        Dim Cmd_BP As New SqlCommand
        Dim Cmd_BP_Reader As SqlDataReader

        Cmd_BP.Connection = Cnn_BP
        Cmd_BP.CommandText = " SELECT id_campione
 [Id_Campione]
     
  FROM [TIRELLI_40].[DBO].[coll_campioni]

  where nome='" & par_nome & "' and tipo_campione='" & par_tipo & "' 
  and (codice_bp='" & par_cardcode & "' or codice_bp_galileo= '" & par_cardcode_jgal & "') and id_campione <> '" & par_id_campione & "'"

        Cmd_BP_Reader = Cmd_BP.ExecuteReader
        If Cmd_BP_Reader.Read() Then

            Blocco_univocità = "Y"
        Else
            Blocco_univocità = "N"


        End If
        Cnn_BP.Close()


    End Sub

    Sub trova_id_campione()


        Dim Cnn_Campioni As New SqlConnection
        Cnn_Campioni.ConnectionString = Homepage.sap_tirelli
        Cnn_Campioni.Open()

        Dim Cmd_Campioni As New SqlCommand
        Dim Cmd_Campioni_Reader As SqlDataReader

        Cmd_Campioni.Connection = Cnn_Campioni
        Cmd_Campioni.CommandText = "SELECT MIN(t1.id_campione + 1) AS 'N'
FROM [TIRELLI_40].[DBO].coll_campioni t1
WHERE NOT EXISTS (
    SELECT 1
    FROM [TIRELLI_40].[DBO].coll_campioni t2
    WHERE t2.id_campione = t1.id_campione + 1
)
AND t1.id_campione >= 4038 "
        Cmd_Campioni_Reader = Cmd_Campioni.ExecuteReader

        If Cmd_Campioni_Reader.Read() Then

            Id_Campione = Cmd_Campioni_Reader("N")

        End If
        Cmd_Campioni_Reader.Close()
        Cnn_Campioni.Close()



    End Sub

    Sub inserisci_flacone(par_id_campione As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].coll_campioni_flaconi 
(codice_campione, altezza, larghezza, profondita, diametro_interno, diametro_esterno, volume, spazio_testa, materiale, forma, sezione, superficie, produttore, codice_produttore, collo_centrato, tipo_tappo, filettatura, diametro_esterno_fil, passo, num_principi) 
VALUES (" & par_id_campione & ", " & If(String.IsNullOrEmpty(TextBox3.Text), "NULL", TextBox3.Text) & ", " & If(String.IsNullOrEmpty(TextBox4.Text), "NULL", TextBox4.Text) & ", " & If(String.IsNullOrEmpty(TextBox5.Text), "NULL", TextBox5.Text) & ", " & If(String.IsNullOrEmpty(TextBox6.Text), "NULL", TextBox6.Text) & ", " & If(String.IsNullOrEmpty(TextBox7.Text), "NULL", TextBox7.Text) & ", " & If(String.IsNullOrEmpty(TextBox8.Text), "NULL", TextBox8.Text) & ", " & If(String.IsNullOrEmpty(TextBox9.Text), "NULL", TextBox9.Text) & ", " & If(String.IsNullOrEmpty(TextBox10.Text), "NULL", "'" & TextBox10.Text & "'") & ", " & If(String.IsNullOrEmpty(TextBox11.Text), "NULL", "'" & TextBox11.Text & "'") & ", " & If(String.IsNullOrEmpty(TextBox12.Text), "NULL", "'" & TextBox12.Text & "'") & ", " & If(String.IsNullOrEmpty(TextBox13.Text), "NULL", "'" & TextBox13.Text & "'") & ", " & If(String.IsNullOrEmpty(TextBox14.Text), "NULL", "'" & TextBox14.Text & "'") & ", " & If(String.IsNullOrEmpty(TextBox15.Text), "NULL", "'" & TextBox15.Text & "'") & ", " & If(String.IsNullOrEmpty(TextBox16.Text), "NULL", "'" & TextBox16.Text & "'") & ", " & If(String.IsNullOrEmpty(TextBox19.Text), "NULL", "'" & TextBox19.Text & "'") & ", " & If(String.IsNullOrEmpty(TextBox20.Text), "NULL", "'" & TextBox20.Text & "'") & ", " & If(String.IsNullOrEmpty(TextBox21.Text), "NULL", TextBox21.Text) & ", " & If(String.IsNullOrEmpty(TextBox22.Text), "NULL", TextBox22.Text) & ", " & If(String.IsNullOrEmpty(TextBox23.Text), "NULL", TextBox23.Text) & ")"


        CMD_SAP.ExecuteNonQuery()





        cnn.Close()
    End Sub

    Sub inserisci_scatola(par_id_campione As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].coll_campioni_scatole " &
    "(codice_campione, altezza, larghezza, profondita, volume, materiale, forma, sezione, superficie, produttore, codice_produttore) " &
    "VALUES (" &
    par_id_campione & ", " &
    If(String.IsNullOrEmpty(TextBox112.Text), "NULL", "'" & TextBox112.Text & "'") & ", " &
    If(String.IsNullOrEmpty(TextBox111.Text), "NULL", "'" & TextBox111.Text & "'") & ", " &
    If(String.IsNullOrEmpty(TextBox110.Text), "NULL", "'" & TextBox110.Text & "'") & ", " &
    If(String.IsNullOrEmpty(TextBox109.Text), "NULL", "'" & TextBox109.Text & "'") & ", " &
    If(String.IsNullOrEmpty(TextBox108.Text), "NULL", "'" & TextBox108.Text & "'") & ", " &
    If(String.IsNullOrEmpty(TextBox107.Text), "NULL", "'" & TextBox107.Text & "'") & ", " &
    If(String.IsNullOrEmpty(TextBox106.Text), "NULL", "'" & TextBox106.Text & "'") & ", " &
    If(String.IsNullOrEmpty(TextBox105.Text), "NULL", "'" & TextBox105.Text & "'") & ", " &
    If(String.IsNullOrEmpty(TextBox104.Text), "NULL", "'" & TextBox104.Text & "'") & ", " &
    If(String.IsNullOrEmpty(TextBox35.Text), "NULL", "'" & TextBox35.Text & "'") & ")"


        CMD_SAP.ExecuteNonQuery()





        Cnn.Close()
    End Sub

    Sub inserisci_tappo(par_id_campione As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].coll_campioni_tappi
(codice_campione, altezza, larghezza, profondità, diametro_interno, Fissaggio, Forma, Materiale, Superficie, Produttore, Codice_produttore) 
VALUES (" & par_id_campione & ",
" & If(String.IsNullOrEmpty(TextBox17.Text), "NULL", TextBox17.Text) & ",
" & If(String.IsNullOrEmpty(TextBox18.Text), "NULL", TextBox18.Text) & ",
" & If(String.IsNullOrEmpty(TextBox1.Text), "NULL", TextBox1.Text) & ",
" & If(String.IsNullOrEmpty(TextBox2.Text), "NULL", TextBox2.Text) & ",
" & If(String.IsNullOrEmpty(TextBox24.Text), "NULL", "'" & TextBox24.Text & "'") & ",
" & If(String.IsNullOrEmpty(TextBox25.Text), "NULL", "'" & TextBox25.Text & "'") & ",
" & If(String.IsNullOrEmpty(TextBox26.Text), "NULL", "'" & TextBox26.Text & "'") & ",
" & If(String.IsNullOrEmpty(TextBox27.Text), "NULL", "'" & TextBox27.Text & "'") & ",
" & If(String.IsNullOrEmpty(TextBox28.Text), "NULL", "'" & TextBox28.Text & "'") & ",
" & If(String.IsNullOrEmpty(TextBox29.Text), "NULL", "'" & TextBox29.Text & "'") & ")"


        CMD_SAP.ExecuteNonQuery()


        cnn.Close()
    End Sub

    Sub inserisci_sottotappo(par_id_campione As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].coll_campioni_sottotappi
(codice_campione, altezza, larghezza, profondità, diametro_interno, [vite_pressione], Forma, Materiale) 
VALUES (" & par_id_campione & ",
" & If(String.IsNullOrEmpty(TextBox39.Text), "NULL", TextBox39.Text) & ",
" & If(String.IsNullOrEmpty(TextBox38.Text), "NULL", TextBox38.Text) & ",
" & If(String.IsNullOrEmpty(TextBox37.Text), "NULL", TextBox37.Text) & ",
" & If(String.IsNullOrEmpty(TextBox36.Text), "NULL", TextBox36.Text) & ",
'" & ComboBox6.Text & "',
" & If(String.IsNullOrEmpty(TextBox34.Text), "NULL", "'" & TextBox34.Text & "'") & ",
" & If(String.IsNullOrEmpty(TextBox33.Text), "NULL", "'" & TextBox33.Text & "'") & ")"


        CMD_SAP.ExecuteNonQuery()


        cnn.Close()
    End Sub

    Sub inserisci_pompetta(par_id_campione As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].[coll_campioni_pompette]
           ([codice_campione]
           ,[A]
           ,[B]
           ,[C]
           ,[D]
           ,[Quota_A]
           ,[Quota_B]
           ,[Quota_C]
           ,[Quota_D]
           ,[Quota_E]
           ,[Quota_F]
           ,[Quota_L]
           ,[SP]
           ,[Materiale]
           ,[Tipologia]
           ,[Superficie]
           ,[Produttore]
           ,[cod_produttore]
           ,[Fissaggio]
           ,[Ghiera]
           ,[Copritappo])
     VALUES
           (" & par_id_campione & "
           ," & If(String.IsNullOrEmpty(TextBox30.Text), "NULL", TextBox30.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox31.Text), "NULL", TextBox31.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox32.Text), "NULL", TextBox32.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox40.Text), "NULL", TextBox40.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox41.Text), "NULL", TextBox41.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox42.Text), "NULL", TextBox42.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox43.Text), "NULL", TextBox43.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox44.Text), "NULL", TextBox44.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox45.Text), "NULL", TextBox45.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox46.Text), "NULL", TextBox46.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox47.Text), "NULL", TextBox47.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox48.Text), "NULL", TextBox48.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox49.Text), "NULL", "'" & TextBox49.Text & "'") & "
           ," & If(String.IsNullOrEmpty(TextBox50.Text), "NULL", "'" & TextBox50.Text & "'") & "
           ," & If(String.IsNullOrEmpty(TextBox51.Text), "NULL", "'" & TextBox51.Text & "'") & "
           ," & If(String.IsNullOrEmpty(TextBox52.Text), "NULL", "'" & TextBox52.Text & "'") & "
           ," & If(String.IsNullOrEmpty(TextBox53.Text), "NULL", "'" & TextBox53.Text & "'") & "
           ," & If(String.IsNullOrEmpty(TextBox54.Text), "NULL", "'" & TextBox54.Text & "'") & "
           ," & If(String.IsNullOrEmpty(TextBox55.Text), "NULL", "'" & TextBox55.Text & "'") & "
           ," & If(String.IsNullOrEmpty(TextBox56.Text), "NULL", "'" & TextBox56.Text & "'") & ")"


        CMD_SAP.ExecuteNonQuery()


        cnn.Close()
    End Sub

    Sub inserisci_etichetta(par_id_campione As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].[coll_campioni_etichette]
           ([codice_campione]
           ,[Altezza]
           ,[larghezza]
           ,[Trasparenza]
           ,[forma]
           ,[diametro_esterno_bobina]
           ,[diametro_interno_bobina]
           ,[Avvolgimento_bobina]
           ,[materiale])
     VALUES
           (" & par_id_campione & "
           ," & If(String.IsNullOrEmpty(TextBox75.Text), "NULL", TextBox75.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox74.Text), "NULL", TextBox74.Text) & "
           ,'" & ComboBox1.Text & "'
           ," & If(String.IsNullOrEmpty(TextBox72.Text), "NULL", "'" & TextBox72.Text & "'") & "
           ," & If(String.IsNullOrEmpty(TextBox71.Text), "NULL", TextBox71.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox70.Text), "NULL", TextBox70.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox69.Text), "NULL", "'" & TextBox69.Text & "'") & "
           ," & If(String.IsNullOrEmpty(TextBox68.Text), "NULL", "'" & TextBox68.Text & "'") & ")"


        CMD_SAP.ExecuteNonQuery()


        cnn.Close()
    End Sub

    Sub inserisci_trigger(par_id_campione As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].[coll_campioni_trigger]
           ([codice_campione]
           ,[A]
           ,[B]
           ,[Quota_S]
           ,[Quota_H]
           ,[Quota_L]
           ,[Quota_W]
           ,[Quota_V]
           ,[Pressione/Vite]
           ,[Produttore]
           ,[Codice_produttore]
           ,[Materiale]
           ,[SP]
           ,[T]
           ,[Fissaggio]
           ,[Ghiera]
           ,[Grileltto]
           ,[Protezione]
           ,[Note]
           ,[Cannuccia])
     VALUES
           (" & par_id_campione & "
           ," & If(String.IsNullOrEmpty(TextBox78.Text), "NULL", TextBox78.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox79.Text), "NULL", TextBox79.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox80.Text), "NULL", TextBox80.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox81.Text), "NULL", TextBox81.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox82.Text), "NULL", TextBox82.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox83.Text), "NULL", TextBox83.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox58.Text), "NULL", TextBox58.Text) & "
           ,'" & ComboBox2.Text & "'
," & If(String.IsNullOrEmpty(TextBox76.Text), "NULL", "'" & TextBox76.Text & "'") & "
," & If(String.IsNullOrEmpty(TextBox73.Text), "NULL", "'" & TextBox73.Text & "'") & "
," & If(String.IsNullOrEmpty(TextBox67.Text), "NULL", "'" & TextBox67.Text & "'") & "
," & If(String.IsNullOrEmpty(TextBox66.Text), "NULL", TextBox66.Text) & "
," & If(String.IsNullOrEmpty(TextBox65.Text), "NULL", TextBox65.Text) & "
," & If(String.IsNullOrEmpty(TextBox64.Text), "NULL", "'" & TextBox64.Text & "'") & "
," & If(String.IsNullOrEmpty(TextBox63.Text), "NULL", "'" & TextBox63.Text & "'") & "
," & If(String.IsNullOrEmpty(TextBox62.Text), "NULL", "'" & TextBox62.Text & "'") & "
," & If(String.IsNullOrEmpty(TextBox61.Text), "NULL", "'" & TextBox61.Text & "'") & "
," & If(String.IsNullOrEmpty(TextBox60.Text), "NULL", "'" & TextBox60.Text & "'") & "
," & If(String.IsNullOrEmpty(TextBox59.Text), "NULL", "'" & TextBox59.Text & "'") & ")
"


        CMD_SAP.ExecuteNonQuery()


        cnn.Close()
    End Sub

    Sub inserisci_prodotto(par_id_campione As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].[coll_campioni_prodotti]
           ([codice_campione]
           ,[densita]
           ,[viscosita_dinamica]
           ,[conducibilita_elettrica]
           ,[categoria]
           ,[infiammabile]
           ,[nome_commerciale]
           ,[viscosità_cinematica]
           ,[Corrosivo]
           ,[Nocivo/tossico]
           ,[Note])
     VALUES
           (" & par_id_campione & "
           ," & If(String.IsNullOrEmpty(TextBox57.Text), "NULL", TextBox57.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox77.Text), "NULL", TextBox77.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox100.Text), "NULL", TextBox100.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox99.Text), "NULL", "'" & TextBox99.Text & "'") & "
           ,'" & ComboBox3.Text & "'
           ," & If(String.IsNullOrEmpty(TextBox84.Text), "NULL", "'" & TextBox84.Text & "'") & "
           ," & If(String.IsNullOrEmpty(TextBox97.Text), "NULL", TextBox97.Text) & "
           ,'" & ComboBox4.Text & "'
,'" & ComboBox5.Text & "'
," & If(String.IsNullOrEmpty(TextBox87.Text), "NULL", "'" & TextBox87.Text & "'") & ")"


        CMD_SAP.ExecuteNonQuery()


        cnn.Close()
    End Sub



    Sub inserisci_film(par_id_campione As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].[coll_campioni_film]
           ([codice_campione]
           ,[larghezza]
           ,[diametro_fulcro]
           ,[materiale]
           ,[temperatura_saldatura]
           ,[Diametro_esterno])
     VALUES
           (" & par_id_campione & "
           ," & If(String.IsNullOrEmpty(TextBox89.Text), "NULL", TextBox89.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox90.Text), "NULL", TextBox90.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox85.Text), "NULL", "'" & TextBox85.Text & "'") & "
           ," & If(String.IsNullOrEmpty(TextBox88.Text), "NULL", TextBox88.Text) & "
           ," & If(String.IsNullOrEmpty(TextBox91.Text), "NULL", TextBox91.Text) & ")"


        CMD_SAP.ExecuteNonQuery()


        cnn.Close()
    End Sub

    Sub inserisci_copritappo(par_id_campione As Integer)


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn



        CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].[coll_campioni_copritappi]
(codice_campione, altezza, larghezza, profondità, diametro_interno, Fissaggio, Forma, Materiale, Superficie, Produttore, Codice_produttore) 
VALUES (" & par_id_campione & ",
" & If(String.IsNullOrEmpty(TextBox103.Text), "NULL", TextBox103.Text) & ",
" & If(String.IsNullOrEmpty(TextBox102.Text), "NULL", TextBox102.Text) & ",
" & If(String.IsNullOrEmpty(TextBox101.Text), "NULL", TextBox101.Text) & ",
" & If(String.IsNullOrEmpty(TextBox98.Text), "NULL", TextBox98.Text) & ",
" & If(String.IsNullOrEmpty(TextBox96.Text), "NULL", "'" & TextBox96.Text & "'") & ",
" & If(String.IsNullOrEmpty(TextBox95.Text), "NULL", "'" & TextBox95.Text & "'") & ",
" & If(String.IsNullOrEmpty(TextBox94.Text), "NULL", "'" & TextBox94.Text & "'") & ",
" & If(String.IsNullOrEmpty(TextBox93.Text), "NULL", "'" & TextBox93.Text & "'") & ",
" & If(String.IsNullOrEmpty(TextBox92.Text), "NULL", "'" & TextBox92.Text & "'") & ",
" & If(String.IsNullOrEmpty(TextBox86.Text), "NULL", "'" & TextBox86.Text & "'") & ")"


        CMD_SAP.ExecuteNonQuery()


        cnn.Close()
    End Sub

    Sub riempi_datagridview_campioni(par_codice_bp As String, par_codice_bp_finale As String)
        Dim Cnn1 As New SqlConnection
        DataGridView3.Rows.Clear()
        DataGridView3.Columns(columnName:="Campione_").Visible = False
        DataGridView3.Columns(columnName:="Tipo_").Visible = False
        DataGridView3.Columns(columnName:="Nome_").Visible = False
        DataGridView3.Columns(columnName:="immagine_").Visible = False
        DataGridView3.Columns(columnName:="dato_6").Visible = False
        DataGridView3.Columns(columnName:="desc").Visible = False



        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.ID_CAMPIONE,  t1.INIZIALE_SIGLA + T0.NOME as 'Nome', case when (t0.immagine is null or t0.immagine ='') then 'N_A.JPG' else t0.immagine end as 'immagine', t1.descrizione as 'Tipo'
, t0.Dato_6, t0.descrizione
from [TIRELLI_40].[DBO].coll_campioni t0 left  join  [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t1 on t0.TIPO_campione= T1.ID_TIPO_CAMPIONE
where (t0.codice_bp=cast('" & par_codice_bp & "' as integer) or t0.codice_bp=cast('" & par_codice_bp & "' as integer) or  t0.codice_bp=cast('" & par_codice_bp_finale & "' as integer))

order by

t1.INIZIALE_SIGLA ,  cast(substring(T0.NOME,1,99) as integer)
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        ' Dim contatore As Integer = 0
        Dim i As Integer = 0
        Do While cmd_SAP_reader_2.Read()


            DataGridView3.Columns(columnName:="Tipo_").Visible = True
            DataGridView3.Columns(columnName:="Nome_").Visible = True
            DataGridView3.Columns(columnName:="immagine_").Visible = True
            DataGridView3.Columns(columnName:="dato_6").Visible = True
            DataGridView3.Columns(columnName:="desc").Visible = True


            Dim MyImage As Bitmap

            Try
                MyImage = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine"))
            Catch ex As Exception
                MyImage = Image.FromFile(Homepage.Percorso_immagini & "N_A.JPG")

            End Try

            DataGridView3.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("Tipo"), MyImage, cmd_SAP_reader_2("dato_6"), cmd_SAP_reader_2("descrizione")) 'Image.FromFile(cmd_SAP_reader_2("immagine"))

            i = i + 1
        Loop


        cmd_SAP_reader_2.Close()
        Cnn1.Close()


        DataGridView3.ClearSelection()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Picture_Campione.SizeMode = PictureBoxSizeMode.Zoom
        Picture_Campione.Image = Clipboard.GetImage
        If Picture_Campione.Image IsNot Nothing Then
            Immagine_Caricata = 1
        End If
    End Sub



    Private Sub TextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress, Txt_nome.KeyPress, TextBox3.KeyPress, TextBox4.KeyPress, TextBox6.KeyPress, TextBox7.KeyPress, TextBox8.KeyPress, TextBox9.KeyPress, TextBox21.KeyPress, TextBox22.KeyPress, TextBox23.KeyPress, TextBox17.KeyPress, TextBox18.KeyPress, TextBox1.KeyPress, TextBox2.KeyPress, TextBox39.KeyPress, TextBox38.KeyPress, TextBox37.KeyPress, TextBox36.KeyPress, TextBox30.KeyPress, TextBox31.KeyPress, TextBox32.KeyPress, TextBox40.KeyPress, TextBox41.KeyPress, TextBox42.KeyPress, TextBox43.KeyPress, TextBox44.KeyPress, TextBox45.KeyPress, TextBox46.KeyPress, TextBox47.KeyPress, TextBox48.KeyPress, TextBox75.KeyPress, TextBox74.KeyPress, TextBox70.KeyPress, TextBox78.KeyPress, TextBox79.KeyPress, TextBox80.KeyPress, TextBox81.KeyPress, TextBox82.KeyPress, TextBox83.KeyPress, TextBox58.KeyPress, TextBox66.KeyPress, TextBox65.KeyPress, TextBox57.KeyPress, TextBox77.KeyPress, TextBox100.KeyPress, TextBox97.KeyPress, TextBox89.KeyPress, TextBox88.KeyPress, TextBox91.KeyPress, TextBox103.KeyPress, TextBox102.KeyPress, TextBox101.KeyPress, TextBox98.KeyPress, TextBox71.KeyPress, TextBox112.KeyPress, TextBox111.KeyPress, TextBox110.KeyPress, TextBox109.KeyPress
        Dim currentTextBox As TextBox = DirectCast(sender, TextBox)

        If currentTextBox Is Txt_nome Then
            ' Logica per Txt_nome: consente solo numeri interi
            If (Not Char.IsDigit(e.KeyChar)) AndAlso (e.KeyChar <> ControlChars.Back) Then
                e.Handled = True
            End If
        Else
            ' Logica per tutti gli altri TextBox: consente solo numeri interi, il punto decimale e il tasto Backspace
            If (Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "." AndAlso e.KeyChar <> ControlChars.Back) Then
                e.Handled = True
            End If

            ' Controlla che ci sia solo un punto decimale
            If e.KeyChar = "." AndAlso currentTextBox.Text.Contains(".") Then
                e.Handled = True
            End If
        End If
    End Sub

    'Private Sub TextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Txt_nome.KeyPress
    '    ' Consenti solo numeri interi
    '    If (Not Char.IsDigit(e.KeyChar)) AndAlso (e.KeyChar <> ControlChars.Back) Then
    '        e.Handled = True
    '    End If
    'End Sub



    Public Sub trova_bp_dato_id_campione(par_id_campione As Integer)
        Dim Cnn_BP As New SqlConnection
        Cnn_BP.ConnectionString = Homepage.sap_tirelli
        Cnn_BP.Open()

        Dim Cmd_BP As New SqlCommand
        Dim Cmd_BP_Reader As SqlDataReader

        Cmd_BP.Connection = Cnn_BP
        Cmd_BP.CommandText = " SELECT t0.codice_bp, t1.cardname
FROM [TIRELLI_40].[DBO].coll_campioni t0 left join ocrd t1 on t0.cardcode=t1.cardname
where t0=" & par_id_campione & ""

        Cmd_BP_Reader = Cmd_BP.ExecuteReader
        If Cmd_BP_Reader.Read() Then

            Codice_BP_selezionato = Cmd_BP_Reader("codice_bp")
            nome_bp_selezionato = Cmd_BP_Reader("cardname")


        End If
        Cnn_BP.Close()


    End Sub



    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then



            '   Form_Inserisci_Campioni.trova_bp_dato_id_campione(DataGridView3.Rows(e.RowIndex).Cells(columnName:="Campione_").Value)
            Form_campione_visualizza.id_campione = DataGridView3.Rows(e.RowIndex).Cells(columnName:="Campione_").Value
            Form_campione_visualizza.Show()
            Form_campione_visualizza.inizializza_form()





        End If
    End Sub

    Private Sub Txt_nome_TextChanged(sender As Object, e As EventArgs) Handles Txt_nome.TextChanged

    End Sub
End Class