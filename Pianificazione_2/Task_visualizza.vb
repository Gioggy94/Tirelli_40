Imports System.IO
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class Task_visualizza
    Public id_task As Integer
    Public inizializzazione As String = "SI"
    Public giorni As Integer

    Sub inizializzazione_form()
        Label1.Text = id_task
        rileva_task()

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()

    End Sub

    Private Sub Task_visualizza_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inizializzazione_form()
    End Sub

    Sub rileva_task()


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        CMD_SAP_2.CommandText = "select t0.id,t0.oc, t1.cardname, CASE WHEN T5.CARDNAME IS NULL THEN '' ELSE t5.cardname end as 'Cliente_finale', t0.task, t2.Nome_task, t0.reparto,t3.Descrizione,t0.riferimento, t4.riferimento as 'Nome_riferimento', t0.giorni,t0.stato,t0.linenum, t0.data_inizio, t0.data_fine, t0.id_link, t0.Data_chiusura_task, t0.Ora_chiusura_task 
from [Tirelli_40].[dbo].[Pianificazione_CDS] t0 inner join ordr t1 on t0.oc=t1.docnum
left join [Tirelli_40].[dbo].[Pianificazione_CDS_TASK] t2 on t2.id =t0.task
left join [TIRELLI_40].[DBO].COLL_Reparti t3 on t0.reparto=t3.Id_Reparto

  left join [Tirelli_40].[dbo].[Pianificazione_CDS_Riferimenti] t4 on t0.Riferimento=t4.id
left join ocrd t5 on t5.cardcode=t1.u_CODICEBP
 where t0.id= " & id_task & ""



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        If cmd_SAP_reader_2.Read() Then

            Label2.Text = cmd_SAP_reader_2("Descrizione")
            Label3.Text = cmd_SAP_reader_2("Nome_task")
            Label4.Text = cmd_SAP_reader_2("cardname")
            Label5.Text = cmd_SAP_reader_2("Data_inizio")
            Label6.Text = cmd_SAP_reader_2("Data_fine")
            giorni = cmd_SAP_reader_2("giorni")
            TextBox1.Text = cmd_SAP_reader_2("giorni")
            ComboBox1.Text = cmd_SAP_reader_2("Stato")
            DateTimePicker4.Value = cmd_SAP_reader_2("Data_inizio")
            DateTimePicker1.Value = cmd_SAP_reader_2("Data_fine")
        End If



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        inizializzazione = "NO"

    End Sub

    Sub Aggiorna_id_task()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = cnn

        If ComboBox1.Text = "P" Then

            Cmd_SAP.CommandText = "UPDATE t0 SET t0.stato='" & ComboBox1.Text & "', T0.DATA_INIZIO = CONVERT(DATETIME, '" & DateTimePicker4.Value & "', 103), T0.DATA_FINE=CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 103)
from [Tirelli_40].[dbo].[Pianificazione_CDS] t0
WHERE t0.id= " & id_task & ""
        Else
            Cmd_SAP.CommandText = "UPDATE t0 SET t0.stato='" & ComboBox1.Text & "', t0.data_chiusura_task =getdate(), t0.ora_chiusura_task=CONCAT(DATEPART(HOUR, GETDATE()),':',DATEPART(MINUTE, GETDATE())), T0.DATA_INIZIO = CONVERT(DATETIME, '" & DateTimePicker4.Value & "', 103), T0.DATA_FINE=CONVERT(DATETIME, '" & DateTimePicker1.Value & "', 103)
from [Tirelli_40].[dbo].[Pianificazione_CDS] t0
WHERE t0.id= " & id_task & ""
        End If


        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Aggiorna_id_task()
        Pianificazione_Tickets.riempi_tasks()
        MsgBox("Task aggiornata con successo")
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If inizializzazione = "NO" Then
            giorni = TextBox1.Text
            DateTimePicker1.Value = DateAdd("d", giorni, DateTimePicker4.Value)
        End If
    End Sub
End Class