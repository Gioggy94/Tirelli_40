Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class Form_Data_Firma
    Private Elenco_campioni_combinazione(10) As Tab_Combinazione
    Public Elenco_Dipendenti(1000) As Integer
    Public Elenco_Combinazioni(1000) As Integer
    Public Num_Dipendenti As Integer
    Public Num_Elementi As Integer
    Public Num_Combinazioni As Integer

    Private Sub Compila_Combo_Dipendenti()
        Dim Indice As Integer

        Dim Cnn_Dipendenti As New SqlConnection

        Combo_Dipendenti.Items.Clear()

        Cnn_Dipendenti.ConnectionString = homepage.sap_tirelli
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

    Private Sub Cmd_Annulla_Click(sender As Object, e As EventArgs) Handles Cmd_Annulla.Click
        Me.Close()
    End Sub

    Private Sub Form_Data_Firma_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Compila_Combo_Dipendenti()
    End Sub


End Class