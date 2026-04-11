Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Form_Visualizza_Articolo
    Private Sub Cmd_Annulla_Click(sender As Object, e As EventArgs) Handles Cmd_Annulla.Click
        Me.Close()
    End Sub

    Public Sub Aggiorna_Dati()
        Dim magazzino_tot As Integer
        Dim confermato_tot As Integer
        Dim Ordinato_tot As Integer
        Dim Disponibile As Integer




        Dim Cnn_Articolo As New SqlConnection
        Cnn_Articolo.ConnectionString = homepage.sap_tirelli
        Cnn_Articolo.Open()
        Dim Cmd_Articolo As New SqlCommand
        Dim Reader_Articolo As SqlDataReader


        DataGridView_magazzino.Rows.Clear()



        Cmd_Articolo.Connection = Cnn_Articolo
        Cmd_Articolo.CommandText = "SELECT T0.[WhsCode], CASE WHEN T0.[OnHand] is null then 0 else T0.[OnHand] END AS 'onhand' , case when T0.[IsCommited] is null then 0 else T0.[IsCommited] end as 'iscommited' , case when T0.[OnOrder] is null then 0 else T0.[OnOrder] end as 'onorder'  FROM OITW T0 WHERE (T0.[OnHand]>0 or t0.iscommited>0 or t0.onorder>0) and t0.itemcode='" & Lbl_Codice.Text & "'"


        Reader_Articolo = Cmd_Articolo.ExecuteReader

        Do While Reader_Articolo.Read()


            DataGridView_magazzino.Rows.Add(Reader_Articolo("whscode"), Reader_Articolo("onhand"), Reader_Articolo("iscommited"), Reader_Articolo("onorder"))
        Loop

        Reader_Articolo.Close()


        Cmd_Articolo.CommandText = "SELECT sum(case when T0.[OnHand] is null then 0 else T0.[OnHand] end ) as 'Magazzino_TOT', sum(case when T0.[iscoMMited] is null then 0 else T0.[iscoMMited] end) as 'Confermato_TOT', sum(case when T0.[onorder] is null then 0 else T0.[onorder] end) as 'ordinato_TOT',  sum(T0.[OnHand]-T0.[IsCommited]+T0.[OnOrder]) as 'Disponibile'
FROM OITW T0 WHERE (T0.[OnHand]>0 or t0.iscommited>0 or t0.onorder>0) and t0.itemcode='" & Lbl_Codice.Text & "'"
        Reader_Articolo = Cmd_Articolo.ExecuteReader

        If Reader_Articolo.Read() = True Then

            Try
                magazzino_tot = Reader_Articolo("Magazzino_TOT")
            Catch ex As Exception
                magazzino_tot = 0
            End Try
            Try
                confermato_tot = Reader_Articolo("Confermato_TOT")
            Catch ex As Exception
                confermato_tot = 0
            End Try
            Try
                Ordinato_tot = Reader_Articolo("ordinato_TOT")
            Catch ex As Exception
                Ordinato_tot = 0
            End Try
            Try
                Disponibile = Reader_Articolo("Disponibile")
            Catch ex As Exception
                Disponibile = 0
            End Try
        End If
        DataGridView_magazzino.Rows.Add("", "", "", "")
        DataGridView_magazzino.Rows.Add("Totale", magazzino_tot, confermato_tot, Ordinato_tot)
        DataGridView_magazzino.Rows.Add("", "", "", "")
        DataGridView_magazzino.Rows.Add("Disponibile", "", "", Disponibile)
        Cnn_Articolo.Close()

    End Sub


End Class