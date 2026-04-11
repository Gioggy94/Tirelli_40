Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Windows.Controls

Public Class CQ_AttivitaAperte
    Public N_attivita As String

    Public filtro_codice As String
    Public filtro_fornitore As String
    Public filtro_matricola As String


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If N_attivita = Nothing Then
            MsgBox("Selezionare un attività")
        Else
            CQ_Modulo_operativo.Anagrafica_attività()
            CQ_Modulo_operativo.Show()
            CQ_Modulo_operativo.Owner = Me
            CQ_Modulo_operativo.Inserimento_Esito_controllo()
            CQ_Modulo_operativo.Inserimento_dipendenti()
            CQ_Modulo_operativo.Inserimento_imputazione()

            If File.Exists(Homepage.percorso_disegni_generico & "PDF\"  & CQ_Modulo_operativo.Label9.Text & ".PDF") Then

                CQ_Modulo_operativo.Button3.BackColor = Color.Lime
                CQ_Modulo_operativo.Button3.Enabled = True

            Else

                CQ_Modulo_operativo.Button3.BackColor = Color.SteelBlue
                CQ_Modulo_operativo.Button3.Enabled = False
            End If
        End If

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Sub Attività_aperte()
        If Homepage.ERP_provenienza = "SAP" Then


            DataGridView1.Rows.Clear()
            Dim CNN As New SqlConnection
            CNN.ConnectionString = Homepage.sap_tirelli
            CNN.Open()

            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

            CMD_SAP.Connection = CNN

            CMD_SAP.CommandText = "
select *
from
(SELECT T0.[ClgCode], t0.u_prg_qlt_itemcode,t1.itemname, t1.u_disegno,case when T0.DOCTYPE='59' then 'MU TIRELLI' WHEN T0.DOCTYPE='20' THEN T2.CARDNAME END as 'cardname' , T0.DOCTYPE, T0.CntctDate
FROM OCLG T0 inner join oitm t1 on t0.u_prg_qlt_itemcode = t1.itemcode
left join ocrd t2 on t2.cardcode=t0.cardcode
WHERE T0.[U_Stato] ='O' and t0.cntcttype=10
)
as t10
where 0=0 " & filtro_codice & " " & filtro_fornitore & "
ORDER BY T10.[ClgCode]"

            cmd_SAP_reader = CMD_SAP.ExecuteReader

            Do While cmd_SAP_reader.Read()
                DataGridView1.Rows.Add(cmd_SAP_reader("ClgCode"), cmd_SAP_reader("u_prg_qlt_itemcode"), cmd_SAP_reader("itemname"), cmd_SAP_reader("U_Disegno"), cmd_SAP_reader("Cardname"), "", cmd_SAP_reader("CntctDate"))
            Loop
            cmd_SAP_reader.Close()
            CNN.Close()

            DataGridView1.ClearSelection()
        End If
    End Sub 'Inserisco le risorse nella combo box




    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = Nothing Then
            filtro_codice = ""
        Else
            filtro_codice = " and t10.u_prg_qlt_itemcode Like '%%" & TextBox2.Text & "%%'  "
        End If
        Attività_aperte()

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = Nothing Then
            filtro_fornitore = ""
        Else
            filtro_fornitore = " and t10.cardname Like '%%" & TextBox3.Text & "%%'  "
        End If
        Attività_aperte()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then


            N_attivita = Nothing
            N_attivita = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        End If
    End Sub



    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting

        If DataGridView1.Rows(e.RowIndex).Cells(4).Value = "MU TIRELLI" Then
            DataGridView1.Rows(e.RowIndex).Cells(4).Style.BackColor = Color.Aqua
        Else
            DataGridView1.Rows(e.RowIndex).Cells(4).Style.BackColor = Color.Lime
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        CQ_nuovo_controllo.Inserimento_dipendenti()
        CQ_nuovo_controllo.Inserimento_imputazione()
        CQ_nuovo_controllo.Inserimento_Esito_controllo()



        CQ_nuovo_controllo.trova_ID()
        CQ_nuovo_controllo.Label1.Text = CQ_nuovo_controllo.id

        CQ_nuovo_controllo.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub CQ_AttivitaAperte_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class