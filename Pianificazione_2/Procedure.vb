Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Procedure
    Public ELENCO_REPARTI(1000) As Integer
    Public filtro_id As String
    Public filtro_codice As String
    Public filtro_reparto As String
    Public filtro_sicurezza As String
    Public filtro_sap As String
    Public filtro_nome As String

    Sub riempi_datagridview_procedure()


        DataGridView1.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT top 100 [ID]
      ,t0.[Codice]
      ,t0.[Reparto]
,t1.descrizione as 'Nome_reparto'
      ,t0.[SAP/4.0]
      ,t0.[Sicurezza]
      ,t0.[Nome]
,t0.[immagine]
  FROM [TIRELLI_40].[dbo].[Procedure] t0 left join [TIRELLI_40].[DBO].COLL_Reparti t1 on t0.reparto=t1.id_reparto
where 0=0 " & filtro_id & "" & filtro_codice & "" & filtro_reparto & "" & filtro_sicurezza & "" & filtro_sap & "" & filtro_nome & ""
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read() = True
            DataGridView1.Rows.Add(cmd_SAP_reader("id"), cmd_SAP_reader("Codice"), cmd_SAP_reader("Reparto"), cmd_SAP_reader("Nome_reparto"), cmd_SAP_reader("SAP/4.0"), cmd_SAP_reader("Sicurezza"), cmd_SAP_reader("Nome"), cmd_SAP_reader("immagine"))
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

        DataGridView1.ClearSelection()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Procedure_nuova.stato_form = "Nuova"
        Procedure_nuova.Show()
        Me.Hide()
        Procedure_nuova.Owner = Me
    End Sub

    Sub carica_reparti()
        Dim Conn_SAP As New SqlConnection
        Dim Com_SAP As New SqlCommand
        Dim DataRead_SAP As SqlDataReader
        Dim Indice As Integer


        Indice = 0
        ComboBox1.Items.Add("")
        'ELENCO_REPARTI(Indice) = ""
        Indice = Indice + 1

        Conn_SAP.ConnectionString = homepage.sap_tirelli
        Conn_SAP.Open()
        Com_SAP.Connection = Conn_SAP
        Com_SAP.CommandText = "SELECT * FROM [TIRELLI_40].[DBO].COLL_Reparti ORDER BY Descrizione"
        DataRead_SAP = Com_SAP.ExecuteReader()


        Do While DataRead_SAP.Read()
            ELENCO_REPARTI(Indice) = DataRead_SAP("Id_Reparto")

            ComboBox1.Items.Add(DataRead_SAP("Descrizione"))


            Indice = Indice + 1
        Loop
        Conn_SAP.Close()

    End Sub

    Private Sub Procedure_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inizializzazione()
    End Sub

    Sub inizializzazione()
        carica_reparti()
        riempi_datagridview_procedure()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = Nothing Then
            filtro_id = ""
        Else
            filtro_id = " and t0.id Like '%" & TextBox1.Text & "%' "
        End If
        riempi_datagridview_procedure()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = Nothing Then
            filtro_nome = ""
        Else
            filtro_nome = " and t0.nome Like '%" & TextBox3.Text & "%'"
        End If
        riempi_datagridview_procedure()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = Nothing Then
            filtro_codice = ""
        Else
            filtro_codice = " and t0.codice Like '%" & TextBox2.Text & "%'"
        End If
        riempi_datagridview_procedure()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex < 1 Then
            filtro_sicurezza = ""
        Else
            filtro_sicurezza = " and t0.sicurezza= '" & ComboBox2.Text & "'"
        End If
        riempi_datagridview_procedure()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged

        If ComboBox3.SelectedIndex < 1 Then
            filtro_sap = ""
        Else
            filtro_sap = " and t0.[SAP/4.0]= '" & ComboBox3.Text & "'"

        End If
        riempi_datagridview_procedure()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex < 0 Then
            filtro_reparto = ""
        Else
            filtro_reparto = " and t0.reparto= " & ELENCO_REPARTI(ComboBox1.SelectedIndex) & ""

        End If
        riempi_datagridview_procedure()
    End Sub


    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then


            Procedure_nuova.id = DataGridView1.Rows(e.RowIndex).Cells(columnName:="id").Value
            If e.ColumnIndex = DataGridView1.Columns.IndexOf(ID) Then

                Procedure_nuova.stato_form = "Visualizza"
                Procedure_nuova.iniziazione_form()
                Procedure_nuova.Show()

            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Sub cancella_procedura()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli


        cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = CNN

        Cmd_SAP.CommandText = "delete  [Tirelli_40].dbo.[Procedure]
 where id ='" & Procedure_nuova.id & "'"
        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = "delete  [Tirelli_40].dbo.procedure_file
 where id_procedura='" & Procedure_nuova.id & "'"
        Cmd_SAP.ExecuteNonQuery()



        cnn.Close()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim password = InputBox("Confermare di eliminare la procedura " & Procedure_nuova.id & " inserendo la password")
        If UCase(password) = "-" Or UCase(password) = "." Then
            cancella_procedura()
            riempi_datagridview_procedure()
            MsgBox("procedura cancellata con successo")



        Else
            MsgBox("Password errata")
        End If

    End Sub
End Class