Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class CQ_ModificaBP
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
        CQ_Password.TextBox1.Text = Nothing
        CQ_Password.TextBox2.Text = Nothing
        CQ_Password.RadioButton2.Checked = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If CheckBox1.Checked = True And CheckBox2.Checked = True Then

            aggiorna_mail()
            aggiorna_telefono()
            dati_anagrafici_BP()
            TextBox2.Text = Nothing
            TextBox3.Text = Nothing
            MsgBox("Telefono e mail aggiornati con successo")
            TextBox1.Enabled = True
        ElseIf CheckBox1.Checked = True And CheckBox2.Checked = False Then

            aggiorna_telefono()
            TextBox1.Enabled = True
            dati_anagrafici_BP()
            TextBox2.Text = Nothing
            MsgBox("Telefono aggiornato con successo")
        ElseIf CheckBox1.Checked = False And CheckBox2.Checked = True Then

            aggiorna_mail()
            TextBox1.Enabled = True
            dati_anagrafici_BP()
            TextBox3.Text = Nothing
            MsgBox("mail aggiornata con successo")
        Else
            MsgBox("Selezionare cosa si vuole aggiornare")
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        'inserire nei relativi campi il nome, il telefono e la mail del fornitore collegato al Cod. BP


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        dati_anagrafici_BP()
        TextBox1.Enabled = False
    End Sub

    Sub dati_anagrafici_BP()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "SELECT T0.[CardName], CASE WHEN T0.[Phone1] IS NULL THEN '' ELSE T0.[Phone1] END AS 'Phone1' , CASE WHEN t0.e_mail IS NULL THEN '' ELSE T0.E_MAIL END AS 'E_mail' FROM OCRD T0 WHERE T0.[CardCode] ='" & TextBox1.Text & " '"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() Then
            Label9.Text = cmd_SAP_reader("Cardname")
            Label3.Text = cmd_SAP_reader("Phone1")
            Label4.Text = cmd_SAP_reader("e_mail")


        End If
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Sub aggiorna_telefono()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand

        CMD_SAP_7.Connection = cnn
        CMD_SAP_7.CommandText = "update ocrd set ocrd.Phone1='" & TextBox2.Text & "' where ocrd.cardcode='" & TextBox1.Text & "'"
        CMD_SAP_7.ExecuteNonQuery()
        cnn.Close()
    End Sub

    Sub aggiorna_mail()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand

        CMD_SAP_7.Connection = cnn
        CMD_SAP_7.CommandText = "update ocrd set ocrd.e_mail='" & TextBox3.Text & "' where ocrd.cardcode='" & TextBox1.Text & "'"
        CMD_SAP_7.ExecuteNonQuery()
        cnn.Close()
    End Sub
End Class