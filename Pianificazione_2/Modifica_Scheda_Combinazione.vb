Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class Modifica_Scheda_Combinazione
    Private Elenco_Elementi(10) As Tab_Combinazione
    Public Elenco_Dipendenti(1000) As Integer
    Public Elenco_responsabili(1000) As Integer
    Public Elenco_Combinazioni(1000) As Integer
    Public Num_Dipendenti As Integer
    Public Num_Dipendenti_responsabili As Integer
    Public numero_combinazioni As Integer

    Public Num_Elementi As Integer
    Public Num_Combinazioni As Integer
    Private num_collaudati As Integer



    Private Function Get_Info_Campione(ID As Integer) As Tab_Combinazione
        Dim Cnn_Campione As New SqlConnection
        Dim Risultato As Tab_Combinazione

        Cnn_Campione.ConnectionString = Homepage.sap_tirelli
        Cnn_Campione.Open()

        Dim Cmd_Campione As New SqlCommand
        Dim Cmd_Campione_Reader As SqlDataReader

        Cmd_Campione.Connection = Cnn_Campione
        Cmd_Campione.CommandText = " SELECT COLL_Campioni.ID_Campione, COLL_Campioni.Nome, COLL_Campioni.Immagine, COLL_Tipo_Campione.Iniziale_Sigla, COLL_Tipo_Campione.Descrizione as 'Descrizione_Campione' 
FROM [TIRELLI_40].[DBO].coll_campioni, COLL_Tipo_Campione WHERE COLL_Campioni.Tipo_Campione=COLL_Tipo_Campione.Id_Tipo_Campione AND Id_Campione=" & ID
        Cmd_Campione_Reader = Cmd_Campione.ExecuteReader
        If Cmd_Campione_Reader.Read() Then
            Risultato.Id_Campione = Cmd_Campione_Reader("ID_Campione")
            Risultato.Nome = Cmd_Campione_Reader("Descrizione_Campione") & " - " & Cmd_Campione_Reader("Iniziale_Sigla") & Cmd_Campione_Reader("Nome")
            Risultato.Automatico = False
            Risultato.Immagine = Cmd_Campione_Reader("Immagine")
        End If

        Cnn_Campione.Close()
        Return Risultato
    End Function

    Public Sub Aggiorna_Scheda_Combinazione(par_id_combinazione As Integer)

        Form_Nuova_combinazione.info_combinazioni(DataGridView2, par_id_combinazione)

        Dim Id As Integer
        Id = Val(par_id_combinazione)

        Dim Cnn_Combinazioni As New SqlConnection
        Cnn_Combinazioni.ConnectionString = Homepage.sap_tirelli
        Cnn_Combinazioni.Open()

        Dim Cmd_Combinazioni As New SqlCommand
        Dim Cmd_Combinazioni_Reader As SqlDataReader

        Cmd_Combinazioni.Connection = Cnn_Combinazioni
        Cmd_Combinazioni.CommandText = "SELECT [Id_Combinazione]
      ,[Commessa]
      ,[Campione_1]
      ,[Campione_2]
      ,[Campione_3]
      ,[Campione_4]
      ,[Campione_5]
      ,[Campione_6]
      ,[Campione_7]
      ,[Campione_8]
      ,[Campione_9]
      ,[Campione_10]
      ,[Automatico_1]
      ,[Automatico_2]
      ,[Automatico_3]
      ,[Automatico_4]
      ,[Automatico_5]
      ,[Automatico_6]
      ,[Automatico_7]
      ,[Automatico_8]
      ,[Automatico_9]
      ,[Automatico_10]
      ,[Vel_Effettiva]
      ,[Vel_Richiesta]
      ,coalesce([Video],0) as 'Video'
      ,coalesce([Firma_Collaudo],0) as 'Firma_collaudo'
      ,[Ricetta]
      ,coalesce([Num_Campioni],0) as 'Num_campioni'
      ,coalesce([Collaudato],0) as 'Collaudato'
      ,[Note]
      ,[Efficienza]
      ,[Estetica_prodotto]
      ,[Tempo_cambio_formato_richiesto]
      ,[Tempo_cambio_formato_effettivo]
      ,coalesce([Firma responsabile],0) as 'Firma_responsabile'

      ,[Insertdate]
      ,[updatedate]
      ,[ownerid]
      ,[Numero_combinazione]
,coalesce(collaudato_con,'') as 'Collaudato_con'
FROM [TIRELLI_40].[DBO].COLL_Combinazioni WHERE Id_Combinazione='" & Id.ToString & "'"
        Cmd_Combinazioni_Reader = Cmd_Combinazioni.ExecuteReader


        If Cmd_Combinazioni_Reader.Read() Then
            Lbl_ID.Text = Cmd_Combinazioni_Reader("Id_Combinazione").ToString
            Txt_Vel_Effettiva.Text = Cmd_Combinazioni_Reader("Vel_Effettiva").ToString
            Txt_Vel_Richiesta.Text = Cmd_Combinazioni_Reader("Vel_Richiesta").ToString
            Txt_Note.Text = Cmd_Combinazioni_Reader("Note").ToString
            Txt_Ricetta.Text = Cmd_Combinazioni_Reader("Ricetta").ToString
            ComboBox2.Text = Cmd_Combinazioni_Reader("collaudato_con").ToString
            Num_Elementi = Cmd_Combinazioni_Reader("Num_Campioni")

            If Num_Elementi >= 1 Then
                Elenco_Elementi(0) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_1"))
                Elenco_Elementi(0).Id_Campione = Cmd_Combinazioni_Reader("Campione_1")
                Elenco_Elementi(0).Automatico = Cmd_Combinazioni_Reader("Automatico_1")
            End If
            If Num_Elementi >= 2 Then
                Elenco_Elementi(1) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_2"))
                Elenco_Elementi(1).Id_Campione = Cmd_Combinazioni_Reader("Campione_2")
                Elenco_Elementi(1).Automatico = Cmd_Combinazioni_Reader("Automatico_2")
            End If
            If Num_Elementi >= 3 Then
                Elenco_Elementi(2) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_3"))
                Elenco_Elementi(2).Id_Campione = Cmd_Combinazioni_Reader("Campione_3")
                Elenco_Elementi(2).Automatico = Cmd_Combinazioni_Reader("Automatico_3")
            End If
            If Num_Elementi >= 4 Then
                Elenco_Elementi(3) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_4"))
                Elenco_Elementi(3).Id_Campione = Cmd_Combinazioni_Reader("Campione_4")
                Elenco_Elementi(3).Automatico = Cmd_Combinazioni_Reader("Automatico_4")
            End If
            If Num_Elementi >= 5 Then
                Elenco_Elementi(4) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_5"))
                Elenco_Elementi(4).Id_Campione = Cmd_Combinazioni_Reader("Campione_5")
                Elenco_Elementi(4).Automatico = Cmd_Combinazioni_Reader("Automatico_5")
            End If
            If Num_Elementi >= 6 Then
                Elenco_Elementi(5) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_6"))
                Elenco_Elementi(5).Id_Campione = Cmd_Combinazioni_Reader("Campione_6")
                Elenco_Elementi(5).Automatico = Cmd_Combinazioni_Reader("Automatico_6")
            End If
            If Num_Elementi >= 7 Then
                Elenco_Elementi(6) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_7"))
                Elenco_Elementi(6).Id_Campione = Cmd_Combinazioni_Reader("Campione_7")
                Elenco_Elementi(6).Automatico = Cmd_Combinazioni_Reader("Automatico_7")
            End If
            If Num_Elementi >= 8 Then
                Elenco_Elementi(7) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_8"))
                Elenco_Elementi(7).Id_Campione = Cmd_Combinazioni_Reader("Campione_8")
                Elenco_Elementi(7).Automatico = Cmd_Combinazioni_Reader("Automatico_8")
            End If
            If Num_Elementi >= 9 Then
                Elenco_Elementi(8) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_9"))
                Elenco_Elementi(8).Id_Campione = Cmd_Combinazioni_Reader("Campione_9")
                Elenco_Elementi(8).Automatico = Cmd_Combinazioni_Reader("Automatico_9")
            End If
            If Num_Elementi >= 10 Then
                Elenco_Elementi(9) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_10"))
                Elenco_Elementi(9).Id_Campione = Cmd_Combinazioni_Reader("Campione_10")
                Elenco_Elementi(9).Automatico = Cmd_Combinazioni_Reader("Automatico_10")
            End If

            Compila_Combo_Dipendenti()
            Compila_Combo_Dipendenti_responsabile()

            If Cmd_Combinazioni_Reader("Firma_Collaudo") Then
                Dim i As Integer
                Dim res As Integer
                res = 0
                For i = 0 To Num_Dipendenti - 1 Step 1
                    If Elenco_Dipendenti(i) = Cmd_Combinazioni_Reader("Firma_Collaudo") Then
                        res = i
                    End If
                Next
                Combo_Dipendenti.SelectedIndex = res
            End If

            If Not DBNull.Value.Equals(Cmd_Combinazioni_Reader("Firma_responsabile")) Then
                Dim i As Integer
                Dim res As Integer
                res = 0
                For i = 0 To Num_Dipendenti_responsabili - 1 Step 1
                    If Elenco_responsabili(i) = Cmd_Combinazioni_Reader("Firma_responsabile") Then
                        res = i
                    End If
                Next
                ComboBox1.SelectedIndex = res
            End If



            If Cmd_Combinazioni_Reader("Video") > 0 Then
                Check_Video.Checked = True
            Else
                Check_Video.Checked = False
            End If
            If Cmd_Combinazioni_Reader("Collaudato") > 0 Then
                Check_Collaudato.Checked = True
            Else
                Check_Collaudato.Checked = False
            End If


        End If
        Cnn_Combinazioni.Close()
    End Sub

    Private Sub Compila_Combo_Dipendenti()
        Dim Indice As Integer

        Dim Cnn_Dipendenti As New SqlConnection

        Combo_Dipendenti.Items.Clear()

        Cnn_Dipendenti.ConnectionString = Homepage.sap_tirelli
        Cnn_Dipendenti.Open()

        Dim Cmd_Dipendenti As New SqlCommand
        Dim Reader_Dipendenti As SqlDataReader

        Cmd_Dipendenti.Connection = Cnn_Dipendenti
        Cmd_Dipendenti.CommandText = "select * FROM [TIRELLI_40].[dbo].OHEM WHERE active='Y' ORDER BY lastName, firstName"

        Reader_Dipendenti = Cmd_Dipendenti.ExecuteReader()
        Indice = 0
        Combo_Dipendenti.Items.Clear()

        Do While Reader_Dipendenti.Read()
            Elenco_Dipendenti(Indice) = Reader_Dipendenti("empID")
            Combo_Dipendenti.Items.Add(Reader_Dipendenti("lastName") & " " & Reader_Dipendenti("firstName"))
            Indice = Indice + 1
        Loop
        Num_Dipendenti = Indice

        Cnn_Dipendenti.Close()
    End Sub


    Private Sub Compila_Combo_Dipendenti_responsabile()
        Dim Indice As Integer

        Dim Cnn_Dipendenti As New SqlConnection

        ComboBox1.Items.Clear()

        Cnn_Dipendenti.ConnectionString = Homepage.sap_tirelli
        Cnn_Dipendenti.Open()

        Dim Cmd_Dipendenti As New SqlCommand
        Dim Reader_Dipendenti As SqlDataReader

        Cmd_Dipendenti.Connection = Cnn_Dipendenti
        Cmd_Dipendenti.CommandText = "select * FROM [TIRELLI_40].[dbo].OHEM
WHERE active='Y' 
ORDER BY lastName, firstName"

        Reader_Dipendenti = Cmd_Dipendenti.ExecuteReader()
        Indice = 0
        ComboBox1.Items.Clear()
        ComboBox1.Items.Add("")


        Do While Reader_Dipendenti.Read()
            Elenco_responsabili(Indice) = Reader_Dipendenti("empID")
            ComboBox1.Items.Add(Reader_Dipendenti("lastName") & " " & Reader_Dipendenti("firstName"))
            Indice = Indice + 1
        Loop
        Num_Dipendenti_responsabili = Indice
        Cnn_Dipendenti.Close()
    End Sub





    Private Sub Modifica_Scheda_Combinazione_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Dim imageColumn As New DataGridViewImageColumn

        imageColumn.ImageLayout = DataGridViewImageCellLayout.Zoom
        imageColumn.HeaderText = "Immagine"

        Txt_Vel_Richiesta.Enabled = False
    End Sub

    Private Sub Cmd_Azione_Click(sender As Object, e As EventArgs) Handles Cmd_Azione.Click
        Dim Collaudato As Integer
        If Check_Collaudato.Checked = True Then
            Collaudato = 1
        Else
            Collaudato = 0

        End If
        Dim Video As Integer
        If Check_Video.Checked = True Then
            Video = 1
        Else
            Video = 0

        End If
        Dim Dipendente As Integer
        If Combo_Dipendenti.Text = "" Then
            Dipendente = 0
        Else
            Dipendente = Elenco_Dipendenti(Combo_Dipendenti.SelectedIndex)
        End If

        Dim responsabile As Integer
        If ComboBox1.Text = "" Then
            responsabile = 0
        Else
            responsabile = Elenco_responsabili(ComboBox1.SelectedIndex)
        End If

        Dim Cnn_Combinazioni As New SqlConnection
        Cnn_Combinazioni.ConnectionString = Homepage.sap_tirelli
        Cnn_Combinazioni.Open()
        Dim Cmd_Combinazioni As New SqlCommand
        Cmd_Combinazioni.Connection = Cnn_Combinazioni





        Cmd_Combinazioni.CommandText = "update [TIRELLI_40].[DBO].COLL_Combinazioni

set vel_effettiva='" & Val(Txt_Vel_Effettiva.Text).ToString & "'
,Vel_Richiesta = '" & Val(Txt_Vel_Richiesta.Text).ToString & "'
,Video = '" & Video.ToString & "'
,Firma_Collaudo = '" & Dipendente & "'
,Ricetta = '" & Val(Txt_Ricetta.Text).ToString & "'
,Collaudato = '" & Collaudato.ToString & "'
,Num_Campioni = '" & Num_Elementi.ToString & "'
,Note = '" & Txt_Note.Text & "'
,FIRMA_RESPONSABILE = '" & responsabile & "'
,collaudato_con='" & ComboBox2.Text & "'


        where id_combinazione='" & Lbl_ID.Text & "'"

        Cmd_Combinazioni.ExecuteNonQuery()
        Cnn_Combinazioni.Close()
        Form_Scheda_Collaudi.Show()

        Form_Scheda_Collaudi.inizializzazione_form(Form_Scheda_Collaudi.Lbl_Commessa.Text)
        Me.Close()
    End Sub

    Private Sub Cmd_Canc_Click(sender As Object, e As EventArgs)
        Compila_Combo_Dipendenti()
    End Sub

    Private Sub Combo_Dipendenti_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Dipendenti.SelectedIndexChanged

    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Check_Video_CheckedChanged(sender As Object, e As EventArgs) Handles Check_Video.CheckedChanged
        If Check_Video.Checked Then
            Check_Video.BackColor = Color.LightGreen
        Else
            Check_Video.BackColor = SystemColors.Control
        End If
    End Sub

    Private Sub Check_Collaudato_CheckedChanged(sender As Object, e As EventArgs) Handles Check_Collaudato.CheckedChanged
        If Check_Collaudato.Checked Then
            Check_Collaudato.BackColor = Color.LightGreen
        Else
            Check_Collaudato.BackColor = SystemColors.Control
        End If
    End Sub
End Class