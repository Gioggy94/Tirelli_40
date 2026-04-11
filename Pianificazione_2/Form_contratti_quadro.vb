Imports System.Data.SqlClient

Public Class Form_contratti_quadro

    Public filtro_cliente As String
    Public filtro_n As String
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Sub inizializza_form()
        carica_contratti_quadro(DataGridView1)
    End Sub

    Sub carica_contratti_quadro(par_datagridview As DataGridView)

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = " select
t0.number, T0.AbsID, T0.BpCode, T0.BpName, T0.CntctCode, T0.StartDate, T0.EndDate, T0.TermDate, 
T0.Descript, T0.Type, T0.Status, T0.Owner, T0.Renewal, T0.UseDiscnt, T0.RemindVal, T0.RemindUnit, T0.Remarks, T0.AtchEntry, T0.LogInstanc, T0.UserSign, T0.UserSign2, T0.UpdtDate, T0.CreateDate, T0.Cancelled, T0.DataSource, T0.Transfered, T0.RemindFlg, T0.Attachment, T0.SettleProb, T0.UpdtTime, T0.Method, T0.GroupNum, T0.ListNum, T0.SignDate, T0.AmendedTo, T0.Series, T0.Number, T0.ObjType, T0.Handwrtten, T0.PIndicator, T0.BpType, T0.PayMethod, T0.NumAtCard, T0.BPCurr, T0.FixedRate, T0.TrnspCode, T0.Project, T0.PriceMode, T0.WddStatus, T0.FromStat, T0.DPPStatus, T0.SAPPassprt, T0.EncryptIV, T0.U_PRG_AZS_NoteChiu, T0.U_PRG_AZS_NoteAper 
FROM OOAT T0 
WHERE t0.bptype='c' and t0.status='A' " & filtro_cliente & filtro_N & "


"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(cmd_SAP_reader_2("AbsID"), cmd_SAP_reader_2("number"), cmd_SAP_reader_2("bpname"), cmd_SAP_reader_2("startdate"), cmd_SAP_reader_2("enddate"), cmd_SAP_reader_2("descript"))


        Loop

        Cnn1.Close()

        par_datagridview.ClearSelection()

    End Sub

    Private Sub Form_contratti_quadro_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inizializza_form()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged


        If TextBox1.Text = "" Then

            filtro_cliente = ""
        Else
            filtro_cliente = " and T0.BpName Like '%%" & TextBox1.Text & "%%' "

        End If

        carica_contratti_quadro(DataGridView1)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

        If TextBox2.Text = "" Then

            filtro_n = ""
        Else
            filtro_n = " and T0.number = '" & TextBox2.Text & "' "

        End If

        carica_contratti_quadro(DataGridView1)
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then


            Form_contratto_quadro_anagrafica.Show()
            Form_contratto_quadro_anagrafica.inizializza_form(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Docentry").Value)



        End If
    End Sub
End Class