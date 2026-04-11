Imports System.Data.SqlClient
Public Class Form109
    Private Sub Button_inserisci_Click(sender As Object, e As EventArgs) Handles Button_inserisci.Click
        If TextBox_Commessa.Text <> "" And TextBox_descrizione_commessa.Text <> "" And TextBox_consegna.Text <> "" Then
            Dim CNN As New SqlConnection
            CNN.ConnectionString = Homepage.sap_tirelli
            cnn.Open()
            Dim CMD_SAP As New SqlCommand
            CMD_SAP.Connection = CNN
            CMD_SAP.CommandText = "INSERT INTO [Tirelli_40].[dbo].[PIANIFICAZIONE_COMMESSA] (Pianificazione_commessa.[Commessa],Pianificazione_commessa.[Descrizione],Pianificazione_commessa.[Stato],Pianificazione_commessa.[cliente],Pianificazione_commessa.[OC],Pianificazione_commessa.[consegna]) Values ('" & Trim(TextBox_Commessa.Text) & "','" & TextBox_descrizione_commessa.Text & "','O', '" & TextBox_cliente.Text & "','" & TextBox_OC.Text & "',CONVERT(DATETIME, '" & TextBox_consegna.Text & "',103)) "
            CMD_SAP.ExecuteNonQuery()
            cnn.Close()
        Else
            MsgBox("Mancano delle informazioni")
        End If

        MsgBox("Commessa inserita con successo")
        Pianificazione.Commesse_aperte()
    End Sub
End Class