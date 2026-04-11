
Imports System.Data.SqlClient

Public Class Form201


    Private Sub Button_cerca_Click(sender As Object, e As EventArgs)
        If Consuntivo1.ComboBox_documento.Text = "ODP" Then
            elenco_ODP_aperti()
        ElseIf Consuntivo1.ComboBox_documento.Text = "OC" Then
            elenco_OC_aperti()
        End If
    End Sub

    Sub elenco_ODP_aperti()

        DataGridView_ODP.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1
        If TextBox_commessa.Text = Nothing Then
            CMD_SAP_1.CommandText = " SELECT top 100 T0.[DocNum] as 'N ODP', t0.status as 'Stato ODP', T0.[ItemCode] as 'Codice', T1.itemname as 'Descrizione', T1.[U_Disegno] as 'Disegno', T0.[PlannedQty] as 'Quantita', T0.[U_PRG_AZS_Commessa] as 'Commessa', T0.[U_UTILIZZ] as 'Cliente'
FROM OWOR T0 inner join OITM T1 on t0.itemcode=t1.itemcode

WHERE  T0.[itemcode]  Like '%%" & TextBox_codice.Text & "%' and T1.[itemname]  Like '%%" & TextBox_descrizione.Text & "%'  and t0.docnum Like '%%" & TextBox1.Text & "%'
order by t0.status DESC"
        Else
            CMD_SAP_1.CommandText = " SELECT T0.[DocNum] as 'N ODP', t0.status as 'Stato ODP', T0.[ItemCode] as 'Codice', T1.itemname as 'Descrizione', T1.[U_Disegno] as 'Disegno', T0.[PlannedQty] as 'Quantita', T0.[U_PRG_AZS_Commessa] as 'Commessa', T0.[U_UTILIZZ] as 'Cliente'
FROM OWOR T0 inner join OITM T1 on t0.itemcode=t1.itemcode

WHERE T0.[U_PRG_AZS_Commessa]  Like '%%" & TextBox_commessa.Text & "%' and T0.[itemcode]  Like '%%" & TextBox_codice.Text & "%' and T1.[itemname]  Like '%%" & TextBox_descrizione.Text & "%'  and t0.docnum Like '%%" & TextBox1.Text & "%' order by t0.status DESC "

        End If



        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()

            DataGridView_ODP.Rows.Add(cmd_SAP_reader_1("N ODP"), cmd_SAP_reader_1("Stato ODP"), cmd_SAP_reader_1("Codice"), cmd_SAP_reader_1("Descrizione"), cmd_SAP_reader_1("Disegno"), Math.Round(cmd_SAP_reader_1("Quantita")), cmd_SAP_reader_1("Commessa"), cmd_SAP_reader_1("Cliente"))

        Loop
        Cnn1.Close()
    End Sub

    Sub elenco_OC_aperti()

        DataGridView_OC.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1

        CMD_SAP_1.CommandText = " SELECT T0.[DocNum] as 'Documento', T0.[CardCode] as 'Codice BP', T0.[CardName] as 'Nome BP',t0.u_matrcds as 'CdS' FROM ORDR T0 WHERE T0.[DocStatus] ='O' 
        AND T0.[DocNum] Like '%%" & TextBox_documento.Text & "%'
        AND T0.[CARDCODE] Like '%%" & TextBox_codice_cliente.Text & "%'
        AND T0.CARDNAME Like '%%" & TextBox_nome_cliente.Text & "%'
        AND COALESCE(t0.u_matrcds,'') Like '%%" & TextBox_CDS.Text & "%'"


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()

            DataGridView_OC.Rows.Add(cmd_SAP_reader_1("Documento"), cmd_SAP_reader_1("Codice BP"), cmd_SAP_reader_1("Nome BP"), cmd_SAP_reader_1("CdS"))

        Loop
        Cnn1.Close()
    End Sub


    Private Sub DataGridView_ODP_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellClick
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = 0 Then
                Consuntivo1.TextBox_numero.Text = DataGridView_ODP.Rows(e.RowIndex).Cells(0).Value
            End If

            Me.Hide()
        End If

    End Sub


    Private Sub DataGridView_OC_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_OC.CellClick
        If e.ColumnIndex = 0 Then
            Consuntivo1.TextBox_numero.Text = DataGridView_OC.Rows(e.RowIndex).Cells(0).Value
            Me.Hide()
        End If


    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        e.Cancel = True ' Impedisce il cambio di scheda



    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox_codice_TextChanged(sender As Object, e As EventArgs) Handles TextBox_codice.TextChanged

    End Sub

    Private Sub TextBox_descrizione_TextChanged(sender As Object, e As EventArgs) Handles TextBox_descrizione.TextChanged

    End Sub

    Private Sub TextBox_commessa_TextChanged(sender As Object, e As EventArgs) Handles TextBox_commessa.TextChanged

    End Sub

    Private Sub TextBox_documento_TextChanged(sender As Object, e As EventArgs) Handles TextBox_documento.TextChanged
        elenco_OC_aperti()
    End Sub

    Private Sub TextBox_codice_cliente_TextChanged(sender As Object, e As EventArgs) Handles TextBox_codice_cliente.TextChanged
        elenco_OC_aperti()
    End Sub

    Private Sub TextBox_nome_cliente_TextChanged(sender As Object, e As EventArgs) Handles TextBox_nome_cliente.TextChanged
        elenco_OC_aperti()
    End Sub

    Private Sub TextBox_CDS_TextChanged(sender As Object, e As EventArgs) Handles TextBox_CDS.TextChanged
        elenco_OC_aperti()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        Scheda_commessa_Pianificazione.carica_commesse(DataGridView, TextBox5.Text, TextBox6.Text, TextBox4.Text, TextBox3.Text, "", "", TextBox2.Text, "", "")
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Scheda_commessa_Pianificazione.carica_commesse(DataGridView, TextBox5.Text, TextBox6.Text, TextBox4.Text, TextBox3.Text, "", "", TextBox2.Text, "", "")
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Scheda_commessa_Pianificazione.carica_commesse(DataGridView, TextBox5.Text, TextBox6.Text, TextBox4.Text, TextBox3.Text, "", "", TextBox2.Text, "", "")
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Scheda_commessa_Pianificazione.carica_commesse(DataGridView, TextBox5.Text, TextBox6.Text, TextBox4.Text, TextBox3.Text, "", "", TextBox2.Text, "", "")
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Scheda_commessa_Pianificazione.carica_commesse(DataGridView, TextBox5.Text, TextBox6.Text, TextBox4.Text, TextBox3.Text, "", "", TextBox2.Text, "", "")
    End Sub

    Private Sub DataGridView_ODP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellContentClick

    End Sub

    Private Sub DataGridView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView.CellContentClick

    End Sub

    Private Sub DataGridView_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView.CellClick
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = 0 Then
                Consuntivo1.TextBox_numero.Text = DataGridView.Rows(e.RowIndex).Cells(0).Value
            End If

            Me.Close()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        elenco_ODP_aperti()
    End Sub
End Class