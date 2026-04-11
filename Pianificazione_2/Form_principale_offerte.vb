Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Pianificazione_offerte
    Private filtro_off As String
    Private filtro_cliente As String
    Private filtro_cliente_f As String
    Private filtro_owner As String
    Private filtro_stato As String
    Public n_offerta As Integer
    Public tabella_intestazione = "OQUT"
    Public tabella_righe = "QUT1"

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub
    Sub inizializza_form()
        riempi_datagridview_offerte(tabella_intestazione, tabella_righe)
    End Sub

    Sub riempi_datagridview_offerte(Par_tabella_intestazione As String, par_tabella_righe As String)

        DataGridView_offerte.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = cnn1



        CMD_SAP_1.CommandText = "select top 100 t0.docnum
,t0.docdate
,t0.docstatus
,t0.cardcode
,t0.cardname
,t2.lastname
, t1.cardcode as 'Cardcode_f'
, t1.cardname as 'Cardname_F'
,t0.doccur
,t0.doctotal,
case when t0.doccur='USD' then t0.docrate end as 'docrate'
,case when t0.doccur='USD' then t0.doctotal*t0.docrate  end as 'Total_$'
from " & Par_tabella_intestazione & " t0 left join ocrd t1 on t0.U_CodiceBP=t1.cardcode
left join [TIRELLI_40].[dbo].ohem t2 on t2.code=t0.ownercode
where 0=0 " & filtro_off & " " & filtro_cliente & "" & filtro_cliente_f & "" & filtro_owner & "" & filtro_stato & "
order by t0.docnum DESC"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()

            DataGridView_offerte.Rows.Add(cmd_SAP_reader_1("DocNum"), cmd_SAP_reader_1("docdate"), cmd_SAP_reader_1("docstatus"), cmd_SAP_reader_1("Lastname"), cmd_SAP_reader_1("cardcode"), cmd_SAP_reader_1("cardname"), cmd_SAP_reader_1("cardcode_f"), cmd_SAP_reader_1("cardname_f"), cmd_SAP_reader_1("doccur"), cmd_SAP_reader_1("docrate"), cmd_SAP_reader_1("doctotal"), cmd_SAP_reader_1("total_$"))

        Loop

        cnn1.Close()
        DataGridView_offerte.ClearSelection()
    End Sub

    Private Sub TextBox_off_TextChanged(sender As Object, e As EventArgs) Handles TextBox_OFF.TextChanged
        If TextBox_OFF.Text = Nothing Then
            filtro_off = ""
        Else
            filtro_off = " and t0.[docnum] Like '%%" & TextBox_OFF.Text & "%%'  "
        End If
        riempi_datagridview_offerte(tabella_intestazione, tabella_righe)
    End Sub

    Private Sub TextBox_cliente_TextChanged(sender As Object, e As EventArgs) Handles TextBox_cliente.TextChanged
        If TextBox_cliente.Text = Nothing Then
            filtro_cliente = ""
        Else
            filtro_cliente = " and t0.[cardname] Like '%%" & TextBox_cliente.Text & "%%'  "
        End If
        riempi_datagridview_offerte(tabella_intestazione, tabella_righe)
    End Sub

    Private Sub TextBox_Cliente_f_TextChanged(sender As Object, e As EventArgs) Handles TextBox_Cliente_f.TextChanged
        If TextBox_Cliente_f.Text = Nothing Then
            filtro_cliente_f = ""
        Else
            filtro_cliente_f = " and t1.[cardname] Like '%%" & TextBox_Cliente_f.Text & "%%'  "
        End If
        riempi_datagridview_offerte(tabella_intestazione, tabella_righe)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = Nothing Then
            filtro_owner = ""
        Else
            filtro_owner = " and t2.[lastname] Like '%%" & TextBox4.Text & "%%'  "
        End If
        riempi_datagridview_offerte(tabella_intestazione, tabella_righe)
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text = Nothing Then
            filtro_stato = ""
        Else
            filtro_stato = " and t0.docstatus Like '%%" & TextBox5.Text & "%%'  "
        End If
        riempi_datagridview_offerte(tabella_intestazione, tabella_righe)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Form_nuova_offerta.Show()
        Form_nuova_offerta.tipo_offerta = "Nuova"
    End Sub


    Private Sub DataGridView_offerte_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_offerte.CellClick
        n_offerta = DataGridView_offerte.Rows(e.RowIndex).Cells(columnName:="N_DOC").Value
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form_nuova_offerta.Show()

        Form_nuova_offerta.TextBox10.Text = n_offerta
        Form_nuova_offerta.tipo_offerta = "Visualizzazione"
        Form_nuova_offerta.inizializzazione_form(n_offerta, tabella_intestazione, tabella_righe, "Offerta")

    End Sub


End Class