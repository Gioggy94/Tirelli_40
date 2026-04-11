Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib

Public Class Offerta_budget
    Public id_quotazione As Integer


    Sub lISTA_MACCHINE()

        DataGridView_commesse.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli

        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select id, tipo_macchina, modello_macchina, info_macch, prodotto_trattato, velocita
from quotazione_budget"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            DataGridView_commesse.Rows.Add(cmd_SAP_reader("ID"), cmd_SAP_reader("tipo_macchina"), cmd_SAP_reader("modello_macchina"), cmd_SAP_reader("info_macch"), cmd_SAP_reader("prodotto_trattato"), cmd_SAP_reader("velocita"))

        Loop


        cmd_SAP_reader.Close()
        cnn.Close()
    End Sub

    Sub Inserimento_tipo_macchina()
        ComboBox2.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select tipo_macchina
from quotazione_budget
where Prodotto_trattato  Like '%%" & ComboBox3.Text & "%%'   and velocita Like '%%" & ComboBox1.Text & "%%' and paese_destinazione Like '%%" & ComboBox4.Text & "%%'
group by 
tipo_macchina"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer

        Do While cmd_SAP_reader.Read()

            ComboBox2.Items.Add(cmd_SAP_reader("Tipo_macchina"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()


    End Sub 'Inserisco le risorse nella combo box

    Sub Inserimento_prodotto()
        ComboBox3.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select prodotto_trattato
from quotazione_budget
where tipo_macchina  Like '%%" & ComboBox2.Text & "%%'   and velocita Like '%%" & ComboBox1.Text & "%%' and paese_destinazione Like '%%" & ComboBox4.Text & "%%'
group by 
prodotto_trattato"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer

        Do While cmd_SAP_reader.Read()

            ComboBox3.Items.Add(cmd_SAP_reader("prodotto_trattato"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Sub Inserimento_velocita()
        ComboBox1.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select velocita
from quotazione_budget
where Prodotto_trattato  Like '%%" & ComboBox3.Text & "%%'   and tipo_macchina Like '%%" & ComboBox2.Text & "%%' and paese_destinazione Like '%%" & ComboBox4.Text & "%%'
group by 
velocita
order by velocita "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer

        Do While cmd_SAP_reader.Read()

            ComboBox1.Items.Add(cmd_SAP_reader("velocita"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Sub filtra()
        Dim i = 0
        Dim parola1 As String
        Dim parola2 As String
        Dim parola4 As String
        Dim parola5 As String

        Do While i < DataGridView_commesse.RowCount


            parola1 = DataGridView_commesse.Rows(i).Cells(1).Value
            parola4 = DataGridView_commesse.Rows(i).Cells(4).Value
            parola5 = DataGridView_commesse.Rows(i).Cells(5).Value



            If parola1.Contains(UCase(ComboBox2.Text)) Then

                DataGridView_commesse.Rows(i).Visible = True


                If parola4.Contains(UCase(ComboBox3.Text)) Then

                    DataGridView_commesse.Rows(i).Visible = True


                    If parola5.Contains(ComboBox1.Text) Then

                        DataGridView_commesse.Rows(i).Visible = True


                    Else
                        DataGridView_commesse.Rows(i).Visible = False

                    End If



                Else
                    DataGridView_commesse.Rows(i).Visible = False

                End If


            Else
                DataGridView_commesse.Rows(i).Visible = False

            End If


            i = i + 1
        Loop
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        filtra()
        Inserimento_velocita()
        Inserimento_prodotto()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        filtra()
        Inserimento_tipo_macchina()
        Inserimento_velocita()


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        filtra()
        Inserimento_prodotto()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ComboBox1.Text = Nothing
        ComboBox2.Text = Nothing
        ComboBox3.Text = Nothing
        filtra()
        Inserimento_tipo_macchina()
        Inserimento_prodotto()
        Inserimento_velocita()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Cambio_BP.Show()
        Me.Hide()
        Cambio_BP.Owner = Me
    End Sub


    Private Sub DataGridView_commesse_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_commesse.CellClick
        id_quotazione = DataGridView_commesse.Rows(e.RowIndex).Cells(0).Value
        prezzo()
    End Sub

    Sub prezzo()

        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select quotazione_budget_cf
from quotazione_budget
where id =" & id_quotazione & " "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then

            TextBox1.Text = Format(cmd_SAP_reader("quotazione_budget_cf"), "Currency")
        End If
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Inserimento_tipo_macchina()
        Inserimento_prodotto()
        Inserimento_velocita()
    End Sub

    Private Sub GroupBox2_Enter(sender As Object, e As EventArgs) Handles GroupBox2.Enter

    End Sub
End Class