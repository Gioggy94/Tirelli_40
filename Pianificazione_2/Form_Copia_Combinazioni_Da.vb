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

Public Class Form_Copia_Combinazioni_Da
    Private Sub Cmd_Esci_Click(sender As Object, e As EventArgs) Handles Cmd_Esci.Click
        Form_Inserisci_Combinazioni.Show()
        Form_Inserisci_Combinazioni.Owner = Me
        Me.Close()
    End Sub

    Private Function Nuovo_ID() As Integer
        Dim Cnn_ID As New SqlConnection
        Dim Risultato As Integer
        Cnn_ID.ConnectionString = homepage.sap_tirelli
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

    Public Sub Copia_Combinazioni()
        Dim Cnn_Combinazioni As New SqlConnection

        Cnn_Combinazioni.ConnectionString = homepage.sap_tirelli
        Cnn_Combinazioni.Open()

        Dim Cmd_Combinazioni As New SqlCommand
        Dim Cmd_Combinazioni_Reader As SqlDataReader

        Cmd_Combinazioni.Connection = Cnn_Combinazioni
        Cmd_Combinazioni.CommandText = "SELECT * FROM [TIRELLI_40].[DBO].COLL_Combinazioni WHERE Commessa='" & Txt_Matricola_Da.Text & "'"
        Cmd_Combinazioni_Reader = Cmd_Combinazioni.ExecuteReader


        Do While Cmd_Combinazioni_Reader.Read()
            Dim Cnn_Nuovo As New SqlConnection
            Cnn_Nuovo.ConnectionString = homepage.sap_tirelli
            Cnn_Nuovo.Open()
            Dim Cmd_Nuovo As New SqlCommand
            Cmd_Nuovo.Connection = Cnn_Nuovo

            Cmd_Nuovo.CommandText = "INSERT INTO [TIRELLI_40].[DBO].COLL_Combinazioni
                (Id_Combinazione,Commessa,Campione_1,Campione_2,Campione_3,Campione_4,Campione_5,Campione_6,Campione_7,Campione_8,
                Campione_9,Campione_10,Automatico_1,Automatico_2,Automatico_3,Automatico_4,Automatico_5,Automatico_6,
                Automatico_7,Automatico_8,Automatico_9,Automatico_10,Vel_Effettiva,Vel_Richiesta,Video,Firma_Collaudo,
                Ricetta,Num_Campioni,Collaudato,Note) VALUES (" &
                Nuovo_ID.ToString & ",'" &
                Txt_Destinazione.Text & "'," &
                Cmd_Combinazioni_Reader("Campione_1") & "," &
                Cmd_Combinazioni_Reader("Campione_2") & "," &
                Cmd_Combinazioni_Reader("Campione_3") & "," &
                Cmd_Combinazioni_Reader("Campione_4") & "," &
                Cmd_Combinazioni_Reader("Campione_5") & "," &
                Cmd_Combinazioni_Reader("Campione_6") & "," &
                Cmd_Combinazioni_Reader("Campione_7") & "," &
                Cmd_Combinazioni_Reader("Campione_8") & "," &
                Cmd_Combinazioni_Reader("Campione_9") & "," &
                Cmd_Combinazioni_Reader("Campione_10") & "," &
                Cmd_Combinazioni_Reader("Automatico_1") & "," &
                Cmd_Combinazioni_Reader("Automatico_2") & "," &
                Cmd_Combinazioni_Reader("Automatico_3") & "," &
                Cmd_Combinazioni_Reader("Automatico_4") & "," &
                Cmd_Combinazioni_Reader("Automatico_5") & "," &
                Cmd_Combinazioni_Reader("Automatico_6") & "," &
                Cmd_Combinazioni_Reader("Automatico_7") & "," &
                Cmd_Combinazioni_Reader("Automatico_8") & "," &
                Cmd_Combinazioni_Reader("Automatico_9") & "," &
                Cmd_Combinazioni_Reader("Automatico_10") & "," &
                Cmd_Combinazioni_Reader("Vel_Effettiva") & "," &
                Cmd_Combinazioni_Reader("Vel_Richiesta") & "," &
                Cmd_Combinazioni_Reader("Video") & "," &
                Cmd_Combinazioni_Reader("Firma_Collaudo") & "," &
                Cmd_Combinazioni_Reader("Ricetta") & "," &
                Cmd_Combinazioni_Reader("Num_Campioni") & "," &
                Cmd_Combinazioni_Reader("Collaudato") & ",'" &
                Cmd_Combinazioni_Reader("Note") & "')"
            Cmd_Nuovo.ExecuteNonQuery()
            Cnn_Nuovo.Close()
        Loop
        Cnn_Combinazioni.Close()
    End Sub

    Private Sub Cmd_Copia_Click(sender As Object, e As EventArgs) Handles Cmd_Copia.Click
        Copia_Combinazioni()
        Form_Inserisci_Combinazioni.Show()
        Form_Inserisci_Combinazioni.Owner = Me
        Form_Inserisci_Combinazioni.Aggiorna_Lista_Combinazioni()
        Me.Close()
    End Sub

    Private Sub Form_Copia_Combinazioni_Da_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class