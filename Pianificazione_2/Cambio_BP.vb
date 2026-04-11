Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib
Public Class Cambio_BP
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox2.Text = Nothing Or TextBox1.Text = Nothing Then
            MsgBox("Mancano delle informazioni")
        Else
            OPPORTUNITA()
        End If

    End Sub

    Sub OPPORTUNITA()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = cnn1
        Cmd_SAP.CommandText = "UPDATE OOPR
SET OOPR.CARDCODE=T10.BPCODE, OOPR.CARDNAME=T10.CLIENTE
FROM
(
SELECT T0.CARDCODE AS 'BPCODE', T0.CARDNAME AS 'CLIENTE'
FROM OCRD T0
WHERE T0.CARDCODE='" & TextBox2.Text & "'
)
AS T10
WHERE OOPR.Opprid='" & TextBox1.Text & "'"

        Cmd_SAP.ExecuteNonQuery()
        cnn1.Close()
        TextBox1.Text = Nothing
        TextBox2.Text = Nothing
        MsgBox("Opportunità aggiornata con successo")

        MsgBox("ATTENZIONE, QUESTA PROCEDURA MODIFICA SOLO LE ANAGRAFICHE ESSENZIALI. CONTROLLARE CONDIZIONI DI PAGAMENTO E ALTRE INFORMAZIONI CHE SONO RIMASTE QUELLE DEL BP PRECEDENTE")

    End Sub

    Sub OFFERTA()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = cnn1
        Cmd_SAP.CommandText = "UPDATE OQUT
SET OQUT.CARDCODE=T10.BPCODE, OQUT.CARDNAME=T10.CLIENTE, OQUT.LICTRADNUM=T10.LICTRADNUM
FROM
(
SELECT T0.CARDCODE AS 'BPCODE', T0.CARDNAME AS 'CLIENTE', T0.LICTRADNUM
FROM OCRD T0
WHERE T0.CARDCODE='" & TextBox3.Text & "'
)
AS T10
WHERE OQUT.Docnum='" & TextBox4.Text & "'"

        Cmd_SAP.ExecuteNonQuery()
        cnn1.Close()
        TextBox3.Text = Nothing
        TextBox4.Text = Nothing
        MsgBox("Offerta aggiornata con successo")
        MsgBox("ATTENZIONE, QUESTA PROCEDURA MODIFICA SOLO LE ANAGRAFICHE ESSENZIALI. CONTROLLARE CONDIZIONI DI PAGAMENTO E ALTRE INFORMAZIONI CHE SONO RIMASTE QUELLE DEL BP PRECEDENTE")
    End Sub

    Sub ORDINE()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = cnn1
        Cmd_SAP.CommandText = "UPDATE ORDR
SET ORDR.CARDCODE=T10.BPCODE, ORDR.CARDNAME=T10.CLIENTE, ORDR.LICTRADNUM=T10.LICTRADNUM
FROM
(
SELECT T0.CARDCODE AS 'BPCODE', T0.CARDNAME AS 'CLIENTE', T0.LICTRADNUM
FROM OCRD T0
WHERE T0.CARDCODE='" & TextBox5.Text & "'
)
AS T10
WHERE ORDR.Docnum='" & TextBox6.Text & "'

update owor set owor.cardcode='" & TextBox5.Text & "' from owor where owor.originnum='" & TextBox6.Text & "'


"

        Cmd_SAP.ExecuteNonQuery()
        cnn1.Close()
        TextBox5.Text = Nothing
        TextBox6.Text = Nothing
        MsgBox("Ordine aggiornato con successo")
        MsgBox("ATTENZIONE, QUESTA PROCEDURA MODIFICA SOLO LE ANAGRAFICHE ESSENZIALI. CONTROLLARE CONDIZIONI DI PAGAMENTO E ALTRE INFORMAZIONI CHE SONO RIMASTE QUELLE DEL BP PRECEDENTE")
    End Sub

    Sub CDS()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = cnn1
        Cmd_SAP.CommandText = "UPDATE OSCL
SET OSCL.customer=T10.BPCODE, OSCL.custmrName=T10.CLIENTE
FROM
(
SELECT T0.CARDCODE AS 'BPCODE', T0.CARDNAME AS 'CLIENTE'
FROM OCRD T0
WHERE T0.CARDCODE='" & TextBox7.Text & "'
)
AS T10
WHERE OSCL.callID='" & TextBox8.Text & "'


"

        Cmd_SAP.ExecuteNonQuery()
        cnn1.Close()
        TextBox5.Text = Nothing
        TextBox6.Text = Nothing
        MsgBox("Ordine aggiornato con successo")
        MsgBox("ATTENZIONE, QUESTA PROCEDURA MODIFICA SOLO LE ANAGRAFICHE ESSENZIALI. CONTROLLARE CONDIZIONI DI PAGAMENTO E ALTRE INFORMAZIONI CHE SONO RIMASTE QUELLE DEL BP PRECEDENTE")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox3.Text = Nothing Or TextBox4.Text = Nothing Then
            MsgBox("Mancano delle informazioni")
        Else
            OFFERTA()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If TextBox5.Text = Nothing Or TextBox6.Text = Nothing Then
            MsgBox("Mancano delle informazioni")
        Else
            ORDINE()
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If TextBox7.Text = Nothing Or TextBox8.Text = Nothing Then
            MsgBox("Mancano delle informazioni")
        Else
            CDS()
        End If

    End Sub
End Class