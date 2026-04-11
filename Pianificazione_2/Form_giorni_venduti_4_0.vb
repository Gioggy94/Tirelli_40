Imports System.Data.SqlClient

Public Class Form_giorni_venduti_4_0
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Sub inizializza_form()
        carica_interventi_effettuati()
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form_aggiungi_giorni_venduti.Show()
    End Sub

    Sub carica_interventi_effettuati()
        DataGridView1.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "select t0.id, t0.itemcode,t1.itemname, t0.giorni,t0.matricola,t3.u_final_customer_name,t3.u_progetto, t0.motivazione,t0.dipendente, concat(t2.firstName,' ', t2.lastname) as 'Nome'
from [Tirelli_40].[dbo].giorni_venduti_4_0 t0 inner join oitm t1 on t0.itemcode=t1.itemcode 
left join [TIRELLI_40].[dbo].ohem t2 on t0.dipendente=t2.empid
left join oitm t3 on t3.itemcode=t0.matricola
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            DataGridView1.Rows.Add(cmd_SAP_reader_2("id"), cmd_SAP_reader_2("dipendente"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("itemcode"), cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("matricola"), cmd_SAP_reader_2("u_final_customer_name"), cmd_SAP_reader_2("u_progetto"), cmd_SAP_reader_2("giorni"))


        Loop

        cnn1.Close()

        DataGridView1.ClearSelection()

    End Sub

    Private Sub Form_giorni_venduti_4_0_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inizializza_form()
    End Sub
End Class