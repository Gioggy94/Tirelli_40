
Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Form_Richieste_Trasferimento
    Public Mag_partenza As String = "01"
    Public Mag_arrivo As String = "B01"

    Sub inizializione_form()

        riempi_datagridview(DataGridView1, Mag_partenza, Mag_arrivo)

    End Sub
    Sub riempi_datagridview(PAR_DATAGRIDVIEW As DataGridView, PAR_MAG_P As String, PAR_MAG_A As String)
        PAR_DATAGRIDVIEW.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn




        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "
SELECT T0.[DocDate], T0.DOCDUEDATE, T0.[DocNum], T1.[ItemCode], T1.[Dscription], T1.[Quantity] 
FROM OWTQ T0  
INNER JOIN WTQ1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T1.[LineStatus] ='o'  AND T1.[FromWhsCod]='" & PAR_MAG_P & "' and   T1.[WhsCode]='" & PAR_MAG_A & "'
ORDER BY T0.DOCDUEDATE


"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            PAR_DATAGRIDVIEW.Rows.Add(cmd_SAP_reader("DOCNUM"), cmd_SAP_reader("itemCODE"), cmd_SAP_reader("DSCRIPTION"), cmd_SAP_reader("QUANTITY"), cmd_SAP_reader("DOCDATE"), cmd_SAP_reader("DOCDUEDATE"))

        Loop


        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Form_Richieste_Trasferimento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inizializione_form()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button8.BackColor = Color.Gold
        Button2.BackColor = Color.Lime
        Button3.BackColor = Color.Gold
        Button7.BackColor = Color.Gold
        Button5.BackColor = Color.Gold
        Button6.BackColor = Color.Gold
        Mag_partenza = "ferretto"
        Mag_arrivo = "06"
        riempi_datagridview(DataGridView1, Mag_partenza, Mag_arrivo)

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Button8.BackColor = Color.Gold
        Button2.BackColor = Color.Gold
        Button3.BackColor = Color.Gold
        Button7.BackColor = Color.Gold
        Button5.BackColor = Color.Lime
        Button6.BackColor = Color.Gold
        Mag_partenza = "01"
        Mag_arrivo = "B01"
        riempi_datagridview(DataGridView1, Mag_partenza, Mag_arrivo)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Button8.BackColor = Color.Gold
        Button2.BackColor = Color.Gold
        Button3.BackColor = Color.Lime
        Button7.BackColor = Color.Gold
        Button5.BackColor = Color.Gold
        Button6.BackColor = Color.Gold
        Mag_partenza = "01"
        Mag_arrivo = "ferretto"
        riempi_datagridview(DataGridView1, Mag_partenza, Mag_arrivo)

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Button8.BackColor = Color.Gold
        Button2.BackColor = Color.Gold
        Button3.BackColor = Color.Gold
        Button7.BackColor = Color.Gold
        Button5.BackColor = Color.Gold
        Button6.BackColor = Color.Lime

        Mag_partenza = "Ferretto"
        Mag_arrivo = "B01"
        riempi_datagridview(DataGridView1, Mag_partenza, Mag_arrivo)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Button8.BackColor = Color.Lime
        Button2.BackColor = Color.Gold
        Button3.BackColor = Color.Gold
        Button7.BackColor = Color.Gold
        Button5.BackColor = Color.Gold
        Button6.BackColor = Color.Gold

        Mag_partenza = "09"
        Mag_arrivo = "01"
        riempi_datagridview(DataGridView1, Mag_partenza, Mag_arrivo)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Button8.BackColor = Color.Gold
        Button2.BackColor = Color.Gold
        Button3.BackColor = Color.Gold
        Button7.BackColor = Color.Lime
        Button5.BackColor = Color.Gold
        Button6.BackColor = Color.Gold

        Mag_partenza = "06"
        Mag_arrivo = "CAP2"
        riempi_datagridview(DataGridView1, Mag_partenza, Mag_arrivo)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then


            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Cod) Then

                Magazzino.Codice_SAP = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Cod").Value

                ' Ripristina la finestra se è minimizzata
                If Magazzino.WindowState = FormWindowState.Minimized Then
                    Magazzino.WindowState = FormWindowState.Normal
                End If

                ' Porta la finestra in primo piano
                Magazzino.BringToFront()
                Magazzino.Activate()
                Magazzino.Show()

                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)


            End If
        End If
    End Sub
End Class