Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Public Class Form_Seleziona_BP_Trattamento

	Private Sub Cmd_Exit_Click(sender As Object, e As EventArgs) Handles Cmd_Exit.Click
		Form_Entrate_Merci.Risposta_BP_Trattamento = " "
		Me.Close()

	End Sub

	Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGrid_BP.CellContentClick
		If e.ColumnIndex = 1 Then
			Form_Entrate_Merci.Risposta_BP_Trattamento = DataGrid_BP.Rows(e.RowIndex).Cells(0).Value
			Me.Close()
		End If
	End Sub

	Private Sub Form_Seleziona_BP_Trattamento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		Dim Cnn_Entrate_Merci As New SqlConnection
		Dim Cmd_Entrate_Merci As New SqlCommand
		Dim Cmd_Entrate_Merci_Reader As SqlDataReader
		'Intestazione Entrata Merci

		Cnn_Entrate_Merci.ConnectionString = homepage.sap_tirelli
		Cnn_Entrate_Merci.Open()
		Cmd_Entrate_Merci.Connection = Cnn_Entrate_Merci
		Cmd_Entrate_Merci.CommandText = "SELECT  T0.[BinCode], T0.[WhsCode] FROM OBIN T0
WHERE T0.WHSCODE='12'"
		Cmd_Entrate_Merci_Reader = Cmd_Entrate_Merci.ExecuteReader
		DataGrid_BP.Rows.Clear()
		Do While Cmd_Entrate_Merci_Reader.Read()
			DataGrid_BP.Rows.Add(Cmd_Entrate_Merci_Reader("BinCode"), "Seleziona...")
		Loop
		Cnn_Entrate_Merci.Close()
	End Sub
End Class