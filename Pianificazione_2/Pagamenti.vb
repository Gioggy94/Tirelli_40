Imports System.Data.SqlClient

Public Class Pagamenti
    Private filtro_N_pagamento As String
    Private filtro_codice_cliente As String
    Private filtro_cliente As String
    Private filtro_N_documento As String

    Private Sub Pagamenti_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        inizializza_form()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()

    End Sub

    Sub inizializza_form()
        riempi_datagridview_pagamenti()
    End Sub

    Sub riempi_datagridview_pagamenti()

        DataGridView1.Rows.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "

select top 100 *
from
(
SELECT t0.docentry, T0.DocNum as 'N° Doc', t0.createdate, T0.DocDate as 'Data',  DATEDIFF(DAY,  t0.DocDate,t0.createdate) AS 'Ritardo_registrazione',  t0.doctime as 'Ora',t0.cardcode, T0.CardName as 'Cliente', T0.TrsfrAcct as 'Conto Co.Ge.',   T0.Comments as 'Note',
case when t1.invtype is null then 99999 else t1.invtype end as 'Invtype',
case when t2.docnum is not null then 'Anticipo cliente'
when t3.docnum is not null then 'Fattura cliente'
when t4.docnum is not null then 'Not di credito cliente'
when t1.invtype=24 then 'Registrazione_prima_nota'
when t1.invtype=30 then 'Registrazione_prima_nota_1'
when t1.invtype=19 then 'Nota credito fornitore'
when t1.invtype=18 then 'Fattura_fornitore'
when t1.invtype=46 then 'Registrazione_prima_nota_2'

end as 'Tipo_documento'
,
case when t2.docnum is not null then t2.docnum
when t3.docnum is not null then t3.docnum
when t4.docnum is not null then t4.docnum
when t1.invtype=24 then ''
when t1.invtype=30 then ''
when t1.invtype=19 then T5.DOCNUM
when t1.invtype=46 then ''

end as 'Numero_documento'
,  T0.DocTotal as 'Pagamento_pervenuto'
, t1.sumapplied as 'Pagamento_suddiviso'
FROM ORCT t0 left join rct2 t1 on t0.docentry=t1.docnum
left join odpi t2 on t2.docentry=t1.docentry and t1.invtype=203
left join oinv t3 on t3.docentry=t1.docentry and t1.invtype=13
left join ORIN t4 on t4.docentry=t1.docentry and t1.invtype=14
left join ORPC t5 on t5.docentry=t1.docentry and t1.invtype=19


)
as t10

where 0=0 " & filtro_N_pagamento & filtro_cliente & filtro_codice_cliente & filtro_N_documento & "

ORDER BY t10.DOCENTRY DESC

"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            DataGridView1.Rows.Add(cmd_SAP_reader("N° Doc"), cmd_SAP_reader("tipo_documento"), cmd_SAP_reader("Numero_documento"), cmd_SAP_reader("createdate"), cmd_SAP_reader("Data"), cmd_SAP_reader("ritardo_registrazione"), cmd_SAP_reader("Ora"), cmd_SAP_reader("cliente"), cmd_SAP_reader("Conto Co.Ge."), cmd_SAP_reader("note"), cmd_SAP_reader("invtype"), cmd_SAP_reader("Pagamento_pervenuto"), cmd_SAP_reader("Pagamento_suddiviso"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            filtro_N_pagamento = ""
        Else
            filtro_N_pagamento = " and t10.[N° Doc]= '" & TextBox1.Text & "'"
        End If
        riempi_datagridview_pagamenti()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = "" Then
            filtro_codice_cliente = ""
        Else
            filtro_codice_cliente = " and t10.[cardcode] Like '%%" & TextBox2.Text & "%%' "
        End If
        riempi_datagridview_pagamenti()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Then
            filtro_cliente = ""
        Else
            filtro_cliente = " and t10.[cliente] Like '%%" & TextBox3.Text & "%%' "
        End If
        riempi_datagridview_pagamenti()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = "" Then
            filtro_N_documento = ""
        Else
            filtro_N_documento = " and t10.[Numero documento] Like '%%" & TextBox4.Text & "%%' "
        End If
        riempi_datagridview_pagamenti()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = DataGridView1.Columns.IndexOf(N_doc_) Then

                If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Invtype").Value = 13 Then

                    Form_nuova_offerta.Show()

                    Form_nuova_offerta.TextBox10.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="N_doc_").Value
                    Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                    Form_nuova_offerta.inizializzazione_form(DataGridView1.Rows(e.RowIndex).Cells(columnName:="N_doc_").Value, "OINV", "INV1", "Fattura")

                ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Invtype").Value = 203 Then


                    Form_nuova_offerta.Show()

                    Form_nuova_offerta.TextBox10.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="N_doc_").Value
                    Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                    Form_nuova_offerta.inizializzazione_form(DataGridView1.Rows(e.RowIndex).Cells(columnName:="N_doc_").Value, "ODPI", "DPI1", "Fattura anticipo")

                Else

                    MsgBox("Visualizzazione di questo documento non ancora sviluppata")

                End If





            End If

        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting

        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Invtype").Value = 13 Then

            DataGridView1.Rows(e.RowIndex).Cells(columnName:="Tipo_documento").Style.BackColor = Color.Aqua

        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Invtype").Value = 203 Then

            DataGridView1.Rows(e.RowIndex).Cells(columnName:="Tipo_documento").Style.BackColor = Color.Coral

        End If


    End Sub
End Class