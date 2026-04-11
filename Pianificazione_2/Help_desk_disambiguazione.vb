Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Help_desk_disambiguazione
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Help_Desk_Interventi.Inserimento_dipendenti(Help_Desk_Interventi.ComboBox3)
        Help_Desk_Interventi.Inserimento_owner(Help_Desk_Interventi.ComboBox4)
        Help_Desk_Interventi.carica_interventi_effettuati()


        Help_Desk_Interventi.Refresh()

        Help_Desk_Interventi.Show()
        Me.Hide()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form_tickets_help_desk_tabella.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form_contratti_quadro.Show()
    End Sub
End Class