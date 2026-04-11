Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class CQ_Modifica_Autocontrollo
    Public id_autocontrollo As Integer
    Public Elenco_GRUPPO_risorse(1000) As String
    Public Gruppo_risorsa_id As String

    Sub carica_controlli()
        DataGridView1.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "SELECT T0.itemcode,T1.[ResGrpNam],T0.controllo,T0.tipo_controllo,T0.id
from autocontrollo_config T0 LEFT JOIN ORSB T1 ON T0.resgrpcod=T1.resgrpcod  
order by id"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            DataGridView1.Rows.Add(cmd_SAP_reader_2("itemcode"), cmd_SAP_reader_2("ResGrpNam"), cmd_SAP_reader_2("controllo"), cmd_SAP_reader_2("tipo_controllo"), cmd_SAP_reader_2("id"))


        Loop

        cnn1.Close()

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            id_autocontrollo = DataGridView1.Rows(e.RowIndex).Cells(4).Value
        End If
    End Sub

    Sub ELIMINA_RECORD()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "DELETE autocontrollo_config WHERE ID='" & id_autocontrollo & "'"
        CMD_SAP.ExecuteNonQuery()

        cnn.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ELIMINA_RECORD()
        carica_controlli()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = Nothing Then
            MsgBox("Inserire un controllo")
        Else
            If ComboBox1.SelectedIndex < 0 Then
                MsgBox("Selezionare il tipo controllo")
            Else
                If ComboBox1.SelectedIndex < 0 Then
                    MsgBox("Selezionare il gruppo macchina")
                Else
                    Trova_ID()
                    inserisci_RECORD()
                    carica_controlli()
                    pulizia()
                End If
            End If
        End If

    End Sub

    Sub pulizia()
        TextBox1.Text = Nothing
        TextBox2.Text = Nothing
        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1
    End Sub

    Sub Trova_ID()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select max(id)+1 as 'ID' from autocontrollo_config"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                id_autocontrollo = cmd_SAP_reader_2("ID")
            Else
                id_autocontrollo = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub

    Sub inserisci_RECORD()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "insert into autocontrollo_config (itemcode,resgrpcod,controllo,tipo_controllo,id) values ('" & TextBox1.Text & "','" & Gruppo_risorsa_id & "','" & TextBox2.Text & "','" & ComboBox1.Text & "', '" & id_autocontrollo & "') "
        CMD_SAP.ExecuteNonQuery()

        cnn.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Sub Inserimento_GRUPPO_RISORSE()
        ComboBox2.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()


        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = cnn
        CMD_SAP_docentry.CommandText = "SELECT T0.[ResGrpCod], T0.[ResGrpNam] FROM ORSB T0
ORDER BY T0.[ResGrpCod]"

        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_docentry_reader.Read()
            Elenco_GRUPPO_risorse(Indice) = cmd_SAP_docentry_reader("ResGrpCod")
            ComboBox2.Items.Add(cmd_SAP_docentry_reader("ResGrpNam"))
            Indice = Indice + 1
        Loop
        cmd_SAP_docentry_reader.Close()
        cnn.Close()


    End Sub

    Private Sub CQ_Modifica_Autocontrollo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Inserimento_GRUPPO_RISORSE()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Try
            Gruppo_risorsa_id = Elenco_GRUPPO_risorse(ComboBox2.SelectedIndex)
        Catch ex As Exception

        End Try

    End Sub


End Class