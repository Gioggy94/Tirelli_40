Imports System.IO
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient

Public Class Form_Cambia_Reparto
    Public Elenco_Reparti(1000) As Integer
    Private codice_reparto As String = ""


    Private Sub Btn_Cancella_Click(sender As Object, e As EventArgs) Handles Btn_Cancella.Click
        Pianificazione_Tickets.Show()
        Me.Close()
    End Sub

    Private Sub Form_Cambia_Dipendente_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Conn_SAP As New SqlConnection
        Dim Com_SAP As New SqlCommand
        Dim DataRead_SAP As SqlDataReader
        Dim Indice As Integer

        Conn_SAP.ConnectionString = homepage.sap_tirelli
        Conn_SAP.Open()
        Com_SAP.Connection = Conn_SAP
        Com_SAP.CommandText = "SELECT * FROM [TIRELLI_40].[DBO].COLL_Reparti ORDER BY Descrizione"
        DataRead_SAP = Com_SAP.ExecuteReader()
        Indice = 0

        Do While DataRead_SAP.Read()
            Elenco_Reparti(Indice) = DataRead_SAP("Id_Reparto")

            Combo_Reparti.Items.Add(DataRead_SAP("Descrizione"))


            Indice = Indice + 1
        Loop
        Conn_SAP.Close()
    End Sub



    Private Sub Cmd_Seleziona_Click(sender As Object, e As EventArgs) Handles Cmd_Seleziona.Click

        If Combo_Reparti.SelectedIndex < 0 Then
            MsgBox("Selezionare un Reparto")
        Else
            codice_reparto = Elenco_Reparti(Combo_Reparti.SelectedIndex)


            Homepage.Aggiorna_INI_COMPUTER()
            '  Homepage.rileva_reparto(Homepage.codice_reparto)
            Me.Hide()
            Homepage.Enabled = True
            Homepage.operazioni_dopo_lettura_ini()
            Pianificazione_Tickets.CODICE_REPARTO = Elenco_Reparti(Combo_Reparti.SelectedIndex)
            Pianificazione_Tickets.inizializzazione_form()

        End If
    End Sub


End Class