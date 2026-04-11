Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Structure Tab_Combinazione
    Public Id_Campione As Integer
    Public Nome As String
    Public Immagine As String
    Public Automatico As Integer
End Structure


Public Class Form_Inserisci_Combinazioni
    Public Codice_BP As Integer
    Public Elenco_BP(10000) As Integer
    Public Elenco_Campioni(10000) As Integer
    Private Elenco_campioni_combinazione(10) As Tab_Combinazione
    Public Elenco_Dipendenti(1000) As Integer
    Public Elenco_Combinazioni(1000) As Integer
    Public Num_Dipendenti As Integer
    Public Num_Elementi As Integer
    Public Num_Combinazioni As Integer

    Public Sub Cerca_BP_Codice()
        Dim Cnn_BP As New SqlConnection
        Cnn_BP.ConnectionString = Homepage.sap_tirelli
        Cnn_BP.Open()

        Dim Cmd_BP As New SqlCommand
        Dim Cmd_BP_Reader As SqlDataReader

        Cmd_BP.Connection = Cnn_BP
        Cmd_BP.CommandText = " SELECT CardCode,CardName FROM OCRD WHERE CardCode='" & Codice_BP & "'"
        Cmd_BP_Reader = Cmd_BP.ExecuteReader
        If Cmd_BP_Reader.Read() Then
            Txt_Lista_Campioni.Text = Codice_BP & " - " & Cmd_BP_Reader("CardName")
        End If
        Cnn_BP.Close()
        Aggiorna_Lista_Campioni()

    End Sub

    Private Sub Cmd_Esci_Click(sender As Object, e As EventArgs) Handles Cmd_Esci.Click

        Me.Close()
    End Sub

    Private Sub Cmd_Cerca_BP_Click(sender As Object, e As EventArgs) Handles Cmd_Cerca_BP.Click
        Lst_BP.Items.Clear()
        Dim Indice As Integer

        Dim Cnn_BP As New SqlConnection
        Cnn_BP.ConnectionString = Homepage.sap_tirelli
        Cnn_BP.Open()

        Dim Cmd_BP As New SqlCommand
        Dim Cmd_BP_Reader As SqlDataReader

        Indice = 0
        Cmd_BP.Connection = Cnn_BP
        Cmd_BP.CommandText = " SELECT CardCode,CardName,CardType,ValidFor FROM OCRD WHERE ValidFor='Y' AND CardType<>'S' AND CardName LIKE N'%" & TXT_Bp.Text & "%' ORDER BY CardName "
        Cmd_BP_Reader = Cmd_BP.ExecuteReader
        Do While Cmd_BP_Reader.Read()
            Lst_BP.Items.Add(Cmd_BP_Reader("CardCode") & " - " & Cmd_BP_Reader("CardName"))
            Elenco_BP(Indice) = Cmd_BP_Reader("CardCode")
            Indice = Indice + 1
        Loop
        Cnn_BP.Close()
    End Sub

    Private Sub Lst_BP_DoubleClick(sender As Object, e As EventArgs) Handles Lst_BP.DoubleClick
        If Lst_BP.SelectedIndex >= 0 Then
            Txt_Lista_Campioni.Text = Lst_BP.SelectedItem.ToString
            Codice_BP = Elenco_BP(Lst_BP.SelectedIndex)
            Aggiorna_Lista_Campioni()

        End If
    End Sub

    Private Sub Form_Inserisci_Combinazioni_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Txt_Lista_Campioni.Enabled = False
        Cmd_Azione.Text = "Inserisci"


        Dim imageColumn As New DataGridViewImageColumn

        imageColumn.ImageLayout = DataGridViewImageCellLayout.Zoom
        imageColumn.HeaderText = "Immagine"
        DataGrid_Combinazione.Columns.Add(imageColumn)

        Compila_Combo_Dipendenti()

        Lbl_ID.Text = Nuovo_ID()
        Cmd_Elimina.Enabled = False

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

    Private Function Nuovo_ID() As Integer
        Dim Cnn_ID As New SqlConnection
        Dim Risultato As Integer
        Cnn_ID.ConnectionString = Homepage.sap_tirelli
        Cnn_ID.Open()

        Dim Cmd_ID As New SqlCommand
        Dim Reader_ID As SqlDataReader

        Cmd_ID.Connection = Cnn_ID
        Cmd_ID.CommandText = "select MAX(Id_Combinazione) As 'Massimo' FROM [TIRELLI_40].[DBO].COLL_Combinazioni"

        Risultato = 0
        Reader_ID = Cmd_ID.ExecuteReader()
        If Reader_ID.Read() Then
            Risultato = Reader_ID("Massimo")
        End If
        Risultato = Risultato + 1

        Cnn_ID.Close()
        Return Risultato
    End Function

    Public Sub Aggiorna_Lista_Campioni()
        Dim Indice As Integer

        Dim Cnn_Campioni As New SqlConnection
        Cnn_Campioni.ConnectionString = Homepage.sap_tirelli
        Cnn_Campioni.Open()
        Dim Cmd_Campioni As New SqlCommand
        Dim Cmd_Campioni_Reader As SqlDataReader

        Cmd_Campioni.Connection = Cnn_Campioni
        Cmd_Campioni.CommandText = "SELECT COLL_Campioni.Nome,COLL_Campioni.id_campione,COLL_Campioni.Tipo_Campione,COLL_Tipo_Campione.Id_Tipo_Campione,COLL_Tipo_Campione.Descrizione as 'Tipo',COLL_Tipo_Campione.Iniziale_Sigla 
FROM [TIRELLI_40].[DBO].coll_campioni,[TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE WHERE COLL_Campioni.Tipo_Campione=COLL_Tipo_Campione.Id_Tipo_Campione AND Codice_BP=" & Codice_BP & "ORDER BY COLL_Tipo_Campione.Descrizione,Nome"
        Cmd_Campioni_Reader = Cmd_Campioni.ExecuteReader
        Indice = 0
        Lst_Campioni.Items.Clear()

        Do While Cmd_Campioni_Reader.Read()
            Lst_Campioni.Items.Add(Cmd_Campioni_Reader("Tipo") & " - " & Cmd_Campioni_Reader("Iniziale_Sigla") & Cmd_Campioni_Reader("Nome"))
            Elenco_Campioni(Indice) = Cmd_Campioni_Reader("Id_Campione")
            Indice = Indice + 1
        Loop
        Cnn_Campioni.Close()
    End Sub

    Private Sub Lst_Campioni_DoubleClick(sender As Object, e As EventArgs) Handles Lst_Campioni.DoubleClick
        If Lst_Campioni.SelectedIndex >= 0 Then
            Elenco_campioni_combinazione(Num_Elementi) = Get_Info_Campione(Elenco_Campioni(Lst_Campioni.SelectedIndex))
            Num_Elementi = Num_Elementi + 1
        End If
        Aggiorna_Grid_Combinazione()

    End Sub

    Private Function Get_Titolo(ID As Integer) As String

        Dim Cnn_Campione As New SqlConnection
        Dim Risultato As String

        Cnn_Campione.ConnectionString = Homepage.sap_tirelli
        Cnn_Campione.Open()

        Dim Cmd_Campione As New SqlCommand
        Dim Cmd_Campione_Reader As SqlDataReader

        Cmd_Campione.Connection = Cnn_Campione
        Cmd_Campione.CommandText = " SELECT COLL_Campioni.id_campione, COLL_Campioni.Nome, COLL_Campioni.Immagine, COLL_Tipo_Campione.Iniziale_Sigla, COLL_Tipo_Campione.Descrizione as 'Descrizione_Campione' 
FROM [TIRELLI_40].[DBO].coll_campioni, [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE WHERE COLL_Campioni.Tipo_Campione=COLL_Tipo_Campione.Id_Tipo_Campione AND Id_Campione=" & ID
        Cmd_Campione_Reader = Cmd_Campione.ExecuteReader
        If Cmd_Campione_Reader.Read() Then
            Risultato = Cmd_Campione_Reader("Iniziale_Sigla") & Cmd_Campione_Reader("Nome")
        End If

        Cnn_Campione.Close()
        Return Risultato
    End Function

    Private Sub Aggiorna_Grid_Combinazione()
        DataGrid_Combinazione.Rows.Clear()

        Dim i As Integer
        For i = 0 To Num_Elementi - 1 Step 1
            Dim Auto As Boolean
            If Elenco_campioni_combinazione(i).Automatico = 0 Then
                Auto = False
            Else
                Auto = True
            End If
            DataGrid_Combinazione.Rows.Add(Elenco_campioni_combinazione(i).Nome, Auto)
            If Elenco_campioni_combinazione(i).Immagine.Length > 1 Then
                Try
                    DataGrid_Combinazione.Rows(i).Cells(3).Value = Image.FromFile(Homepage.Percorso_immagini & Elenco_campioni_combinazione(i).Immagine)
                Catch ex As Exception

                End Try
            End If

        Next

    End Sub

    Private Function Get_Info_Campione(ID As Integer) As Tab_Combinazione
        Dim Cnn_Campione As New SqlConnection
        Dim Risultato As Tab_Combinazione

        Cnn_Campione.ConnectionString = Homepage.sap_tirelli
        Cnn_Campione.Open()

        Dim Cmd_Campione As New SqlCommand
        Dim Cmd_Campione_Reader As SqlDataReader

        Cmd_Campione.Connection = Cnn_Campione
        Cmd_Campione.CommandText = " SELECT COLL_Campioni, COLL_Campioni.Nome, COLL_Campioni.Immagine, COLL_Tipo_Campione.Iniziale_Sigla, COLL_Tipo_Campione.Descrizione as 'Descrizione_Campione' 
FROM [TIRELLI_40].[DBO].coll_campioni, [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE WHERE COLL_Campioni.Tipo_Campione=COLL_Tipo_Campione.Id_Tipo_Campione AND Id_Campione=" & ID
        Cmd_Campione_Reader = Cmd_Campione.ExecuteReader
        If Cmd_Campione_Reader.Read() Then
            Risultato = Cmd_Campione_Reader("ID_Campione")
            Risultato.Nome = Cmd_Campione_Reader("Descrizione_Campione") & " - " & Cmd_Campione_Reader("Iniziale_Sigla") & Cmd_Campione_Reader("Nome")
            Risultato.Automatico = False
            Risultato.Immagine = Cmd_Campione_Reader("Immagine")
        End If

        Cnn_Campione.Close()
        Return Risultato
    End Function

    Private Sub DataGrid_Combinazione_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGrid_Combinazione.CellContentClick
        If e.RowIndex >= 0 Then


            If e.ColumnIndex.ToString = 1 Then
                If Elenco_campioni_combinazione(e.RowIndex).Automatico = 1 Then
                    Elenco_campioni_combinazione(e.RowIndex).Automatico = 0
                Else
                    Elenco_campioni_combinazione(e.RowIndex).Automatico = 1
                End If
                Aggiorna_Grid_Combinazione()
            End If


            If e.ColumnIndex.ToString = 2 Then
                Dim Indice As Integer

                For Indice = e.RowIndex To Num_Elementi
                    Elenco_campioni_combinazione(Indice) = Elenco_campioni_combinazione(Indice + 1)
                Next
                Num_Elementi = Num_Elementi - 1
                Pulisci_Elenco()
                Aggiorna_Grid_Combinazione()
            End If
        End If
    End Sub

    Private Sub Pulisci_Elenco()

    End Sub



    Private Sub Cmd_Canc_Click(sender As Object, e As EventArgs) Handles Cmd_Canc.Click
        Compila_Combo_Dipendenti()
    End Sub

    Public Sub Aggiorna_Lista_Combinazioni()
        Dim Cnn_Combinazioni As New SqlConnection


        Num_Combinazioni = 0
        Lst_Combinazioni.Items.Clear()

        Cnn_Combinazioni.ConnectionString = Homepage.sap_tirelli
        Cnn_Combinazioni.Open()

        Dim Cmd_Combinazioni As New SqlCommand
        Dim Cmd_Combinazioni_Reader As SqlDataReader

        Cmd_Combinazioni.Connection = Cnn_Combinazioni
        Cmd_Combinazioni.CommandText = "SELECT 
* FROM [TIRELLI_40].[DBO].COLL_Combinazioni WHERE Commessa='" & Lbl_Commessa.Text & "'"
        Cmd_Combinazioni_Reader = Cmd_Combinazioni.ExecuteReader


        Do While Cmd_Combinazioni_Reader.Read()
            Elenco_Combinazioni(Num_Combinazioni) = Cmd_Combinazioni_Reader("Id_Combinazione")
            Num_Combinazioni = Num_Combinazioni + 1
            Dim Titolo As String
            Titolo = Cmd_Combinazioni_Reader("Id_Combinazione") & " - "
            If (Cmd_Combinazioni_Reader("Campione_1")) > 0 Then
                Titolo = Titolo & Get_Titolo(Cmd_Combinazioni_Reader("Campione_1"))
                If Cmd_Combinazioni_Reader("Automatico_1") > 0 Then
                    Titolo = Titolo & "A"
                Else
                    Titolo = Titolo & "M"
                End If
            End If
            If (Cmd_Combinazioni_Reader("Campione_2")) > 0 Then
                Titolo = Titolo & Get_Titolo(Cmd_Combinazioni_Reader("Campione_2"))
                If Cmd_Combinazioni_Reader("Automatico_2") > 0 Then
                    Titolo = Titolo & "A"
                Else
                    Titolo = Titolo & "M"
                End If
            End If
            If (Cmd_Combinazioni_Reader("Campione_3")) > 0 Then
                Titolo = Titolo & Get_Titolo(Cmd_Combinazioni_Reader("Campione_3"))
                If Cmd_Combinazioni_Reader("Automatico_3") > 0 Then
                    Titolo = Titolo & "A"
                Else
                    Titolo = Titolo & "M"
                End If
            End If
            If (Cmd_Combinazioni_Reader("Campione_4")) > 0 Then
                Titolo = Titolo & Get_Titolo(Cmd_Combinazioni_Reader("Campione_4"))
                If Cmd_Combinazioni_Reader("Automatico_4") > 0 Then
                    Titolo = Titolo & "A"
                Else
                    Titolo = Titolo & "M"
                End If
            End If
            If (Cmd_Combinazioni_Reader("Campione_5")) > 0 Then
                Titolo = Titolo & Get_Titolo(Cmd_Combinazioni_Reader("Campione_5"))
                If Cmd_Combinazioni_Reader("Automatico_5") > 0 Then
                    Titolo = Titolo & "A"
                Else
                    Titolo = Titolo & "M"
                End If
            End If
            If (Cmd_Combinazioni_Reader("Campione_6")) > 0 Then
                Titolo = Titolo & Get_Titolo(Cmd_Combinazioni_Reader("Campione_6"))
                If Cmd_Combinazioni_Reader("Automatico_6") > 0 Then
                    Titolo = Titolo & "A"
                Else
                    Titolo = Titolo & "M"
                End If
            End If
            If (Cmd_Combinazioni_Reader("Campione_7")) > 0 Then
                Titolo = Titolo & Get_Titolo(Cmd_Combinazioni_Reader("Campione_7"))
                If Cmd_Combinazioni_Reader("Automatico_7") > 0 Then
                    Titolo = Titolo & "A"
                Else
                    Titolo = Titolo & "M"
                End If
            End If
            If (Cmd_Combinazioni_Reader("Campione_8")) > 0 Then
                Titolo = Titolo & Get_Titolo(Cmd_Combinazioni_Reader("Campione_8"))
                If Cmd_Combinazioni_Reader("Automatico_8") > 0 Then
                    Titolo = Titolo & "A"
                Else
                    Titolo = Titolo & "M"
                End If
            End If
            If (Cmd_Combinazioni_Reader("Campione_9")) > 0 Then
                Titolo = Titolo & Get_Titolo(Cmd_Combinazioni_Reader("Campione_9"))
                If Cmd_Combinazioni_Reader("Automatico_9") > 0 Then
                    Titolo = Titolo & "A"
                Else
                    Titolo = Titolo & "M"
                End If
            End If
            If (Cmd_Combinazioni_Reader("Campione_10")) > 0 Then
                Titolo = Titolo & Get_Titolo(Cmd_Combinazioni_Reader("Campione_10"))
                If Cmd_Combinazioni_Reader("Automatico_10") > 0 Then
                    Titolo = Titolo & "A"
                Else
                    Titolo = Titolo & "M"
                End If
            End If
            Lst_Combinazioni.Items.Add(Titolo)
        Loop


        Cnn_Combinazioni.Close()
    End Sub

    Private Sub Lst_Combinazioni_DoubleClick(sender As Object, e As EventArgs) Handles Lst_Combinazioni.DoubleClick
        If Lst_Combinazioni.SelectedIndex >= 0 Then
            Dim Id As Integer
            Id = Elenco_Combinazioni(Lst_Combinazioni.SelectedIndex)

            Dim Cnn_Combinazioni As New SqlConnection
            Cnn_Combinazioni.ConnectionString = Homepage.sap_tirelli
            Cnn_Combinazioni.Open()

            Dim Cmd_Combinazioni As New SqlCommand
            Dim Cmd_Combinazioni_Reader As SqlDataReader

            Cmd_Combinazioni.Connection = Cnn_Combinazioni
            Cmd_Combinazioni.CommandText = "SELECT * FROM [TIRELLI_40].[DBO].COLL_Combinazioni WHERE Id_Combinazione='" & Id.ToString & "'"
            Cmd_Combinazioni_Reader = Cmd_Combinazioni.ExecuteReader


            If Cmd_Combinazioni_Reader.Read() Then
                Lbl_ID.Text = Cmd_Combinazioni_Reader("Id_Combinazione").ToString
                Txt_Vel_Effettiva.Text = Cmd_Combinazioni_Reader("Vel_Effettiva").ToString
                Txt_Vel_Richiesta.Text = Cmd_Combinazioni_Reader("Vel_Richiesta").ToString
                Txt_Note.Text = Cmd_Combinazioni_Reader("Note").ToString
                Txt_Ricetta.Text = Cmd_Combinazioni_Reader("Ricetta").ToString
                Num_Elementi = Cmd_Combinazioni_Reader("Num_Campioni")

                If Num_Elementi >= 1 Then
                    Elenco_campioni_combinazione(0) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_1"))
                    Elenco_campioni_combinazione(0) = Cmd_Combinazioni_Reader("Campione_1")
                    Elenco_campioni_combinazione(0).Automatico = Cmd_Combinazioni_Reader("Automatico_1")
                End If
                If Num_Elementi >= 2 Then
                    Elenco_campioni_combinazione(1) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_2"))
                    Elenco_campioni_combinazione(1) = Cmd_Combinazioni_Reader("Campione_2")
                    Elenco_campioni_combinazione(1).Automatico = Cmd_Combinazioni_Reader("Automatico_2")
                End If
                If Num_Elementi >= 3 Then
                    Elenco_campioni_combinazione(2) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_3"))
                    Elenco_campioni_combinazione(2) = Cmd_Combinazioni_Reader("Campione_3")
                    Elenco_campioni_combinazione(2).Automatico = Cmd_Combinazioni_Reader("Automatico_3")
                End If
                If Num_Elementi >= 4 Then
                    Elenco_campioni_combinazione(3) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_4"))
                    Elenco_campioni_combinazione(3) = Cmd_Combinazioni_Reader("Campione_4")
                    Elenco_campioni_combinazione(3).Automatico = Cmd_Combinazioni_Reader("Automatico_4")
                End If
                If Num_Elementi >= 5 Then
                    Elenco_campioni_combinazione(4) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_5"))
                    Elenco_campioni_combinazione(4) = Cmd_Combinazioni_Reader("Campione_5")
                    Elenco_campioni_combinazione(4).Automatico = Cmd_Combinazioni_Reader("Automatico_5")
                End If
                If Num_Elementi >= 6 Then
                    Elenco_campioni_combinazione(5) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_6"))
                    Elenco_campioni_combinazione(5) = Cmd_Combinazioni_Reader("Campione_6")
                    Elenco_campioni_combinazione(5).Automatico = Cmd_Combinazioni_Reader("Automatico_6")
                End If
                If Num_Elementi >= 7 Then
                    Elenco_campioni_combinazione(6) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_7"))
                    Elenco_campioni_combinazione(6) = Cmd_Combinazioni_Reader("Campione_7")
                    Elenco_campioni_combinazione(6).Automatico = Cmd_Combinazioni_Reader("Automatico_7")
                End If
                If Num_Elementi >= 8 Then
                    Elenco_campioni_combinazione(7) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_8"))
                    Elenco_campioni_combinazione(7) = Cmd_Combinazioni_Reader("Campione_8")
                    Elenco_campioni_combinazione(7).Automatico = Cmd_Combinazioni_Reader("Automatico_8")
                End If
                If Num_Elementi >= 9 Then
                    Elenco_campioni_combinazione(8) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_9"))
                    Elenco_campioni_combinazione(8) = Cmd_Combinazioni_Reader("Campione_9")
                    Elenco_campioni_combinazione(8).Automatico = Cmd_Combinazioni_Reader("Automatico_9")
                End If
                If Num_Elementi >= 10 Then
                    Elenco_campioni_combinazione(9) = Get_Info_Campione(Cmd_Combinazioni_Reader("Campione_10"))
                    Elenco_campioni_combinazione(9) = Cmd_Combinazioni_Reader("Campione_10")
                    Elenco_campioni_combinazione(9).Automatico = Cmd_Combinazioni_Reader("Automatico_10")
                End If

                Compila_Combo_Dipendenti()
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

                Aggiorna_Grid_Combinazione()
                Cmd_Azione.Text = "Aggiorna"
                Cmd_Elimina.Enabled = True
            End If
            Cnn_Combinazioni.Close()
        End If
    End Sub

    Private Sub Cmd_Nuova_Click(sender As Object, e As EventArgs) Handles Cmd_Nuova.Click
        Num_Elementi = 0
        Pulisci_Elenco()
        Aggiorna_Grid_Combinazione()
        Txt_Note.Text = ""
        Txt_Ricetta.Text = ""
        Txt_Vel_Effettiva.Text = ""
        Txt_Vel_Richiesta.Text = ""
        Check_Collaudato.Checked = False
        Check_Video.Checked = False
        Compila_Combo_Dipendenti()
        Nuovo_ID()
        Aggiorna_Lista_Combinazioni()
        Cmd_Azione.Text = "Inserisci"
        Cmd_Elimina.Enabled = False
    End Sub

    Private Sub Cmd_Elimina_Click(sender As Object, e As EventArgs) Handles Cmd_Elimina.Click
        Dim Cnn_Combinazioni As New SqlConnection
        Cnn_Combinazioni.ConnectionString = homepage.sap_tirelli
        Cnn_Combinazioni.Open()
        Dim Cmd_Combinazioni As New SqlCommand
        Cmd_Combinazioni.Connection = Cnn_Combinazioni
        Cmd_Combinazioni.CommandText = "DELETE FROM [TIRELLI_40].[DBO].COLL_Combinazioni WHERE Id_Combinazione=" & Lbl_ID.Text
        Cmd_Combinazioni.ExecuteNonQuery()
        Cnn_Combinazioni.Close()
        Num_Elementi = 0
        Pulisci_Elenco()
        Aggiorna_Grid_Combinazione()
        Txt_Note.Text = ""
        Txt_Ricetta.Text = ""
        Txt_Vel_Effettiva.Text = ""
        Txt_Vel_Richiesta.Text = ""
        Check_Collaudato.Checked = False
        Check_Video.Checked = False
        Compila_Combo_Dipendenti()
        Nuovo_ID()
        Aggiorna_Lista_Combinazioni()
        Cmd_Azione.Text = "Inserisci"
        Cmd_Elimina.Enabled = False
    End Sub

    Private Sub Cmd_Copia_Da_Click(sender As Object, e As EventArgs) Handles Cmd_Copia_Da.Click
        Form_Copia_Combinazioni_Da.Show()
        Form_Copia_Combinazioni_Da.Txt_Destinazione.Text = Lbl_Commessa.Text
        Me.Hide()
    End Sub

    Private Sub DataGrid_Combinazione_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGrid_Combinazione.CellClick

    End Sub

    Private Sub Lst_Combinazioni_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lst_Combinazioni.SelectedIndexChanged

    End Sub

    Private Sub Cmd_Azione_Click(sender As Object, e As EventArgs) Handles Cmd_Azione.Click

    End Sub
End Class