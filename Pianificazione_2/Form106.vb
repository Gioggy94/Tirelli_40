Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Form106
    Public percentuale As Integer

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click


        'Dashboard_MU.TextBox1.Text = Nothing
        'If Dashboard_MU.Button11.BackColor = Color.Red Then
        '    Autocontrollo.carica_checklist_autocontrollo()
        '    Autocontrollo.Show()
        'End If
        'Me.Close()


        Dashboard_MU_New.percentuale = 0
        ' Dashboard_MU_New.inserisci_lavorazione_a_sap()
        Dashboard_MU_New.rECORD_CARICATI(Dashboard_MU_New.DataGridView2)
        Dashboard_MU_New.pulisci_campi_manodopera()
        MsgBox("Lavorazione inserita con successo")
        Me.Owner.Show()
        Me.Close()

        'If Dashboard_MU.Button11.BackColor = Color.Red Then
        '    Autocontrollo.carica_checklist_autocontrollo()
        '    Autocontrollo.Show()
        'End If
        'Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dashboard_MU_New.percentuale = 1
        'Dashboard_MU_New.inserisci_lavorazione_a_sap()
        Dashboard_MU_New.rECORD_CARICATI(Dashboard_MU_New.DataGridView2)
        Dashboard_MU_New.pulisci_campi_manodopera()
        Me.Owner.Show()
        Me.Close()

        'percentuale = 1
        'inserisci_percentuale()
        'Lavorazioni_MES.manodopera_attrezzaggio()


        'Dashboard_MU.TextBox1.Text = Nothing
        'If Dashboard_MU.Button11.BackColor = Color.Red Then
        '    Autocontrollo.carica_checklist_autocontrollo()
        '    Autocontrollo.Show()
        'End If
        'Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        MsgBox(Dashboard_MU_New.docnum_odp)
        Me.Visible = True
        Dashboard_MU_New.percentuale = 2
        MsgBox(Dashboard_MU_New.docnum_odp)
        'Dashboard_MU_New.inserisci_lavorazione_a_sap()
        MsgBox(Dashboard_MU_New.docnum_odp)
        Dashboard_MU_New.rECORD_CARICATI(Dashboard_MU_New.DataGridView2)
        Dashboard_MU_New.pulisci_campi_manodopera()
        'Me.Owner.Show()
        Me.Hide()

        'percentuale = 2
        'inserisci_percentuale()
        'Lavorazioni_MES.manodopera_attrezzaggio()


        'Dashboard_MU.TextBox1.Text = Nothing
        'If Dashboard_MU.Button11.BackColor = Color.Red Then
        '    Autocontrollo.carica_checklist_autocontrollo()
        '    Autocontrollo.Show()
        'End If
        'Me.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dashboard_MU_New.percentuale = 3
        ' Dashboard_MU_New.inserisci_lavorazione_a_sap()
        Dashboard_MU_New.rECORD_CARICATI(Dashboard_MU_New.DataGridView2)
        Dashboard_MU_New.pulisci_campi_manodopera()
        Me.Owner.Show()
        Me.Close()

        'percentuale = 3
        'inserisci_percentuale()
        'Lavorazioni_MES.manodopera_attrezzaggio()


        'Dashboard_MU.TextBox1.Text = Nothing
        'If Dashboard_MU.Button11.BackColor = Color.Red Then
        '    Autocontrollo.carica_checklist_autocontrollo()
        '    Autocontrollo.Show()
        'End If
        'Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Dashboard_MU_New.percentuale = 4
        '  Dashboard_MU_New.inserisci_lavorazione_a_sap()
        Dashboard_MU_New.rECORD_CARICATI(Dashboard_MU_New.DataGridView2)
        Dashboard_MU_New.pulisci_campi_manodopera()
        Me.Owner.Show()
        Me.Close()

        'percentuale = 4
        'inserisci_percentuale()
        'Lavorazioni_MES.manodopera_attrezzaggio()


        'Dashboard_MU.TextBox1.Text = Nothing
        'If Dashboard_MU.Button11.BackColor = Color.Red Then
        '    Autocontrollo.carica_checklist_autocontrollo()
        '    Autocontrollo.Show()
        'End If
        'Me.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dashboard_MU_New.percentuale = 5
        '   Dashboard_MU_New.inserisci_lavorazione_a_sap()
        Dashboard_MU_New.rECORD_CARICATI(Dashboard_MU_New.DataGridView2)
        Dashboard_MU_New.pulisci_campi_manodopera()
        Me.Owner.Show()
        Me.Close()

        'percentuale = 5
        'inserisci_percentuale()
        'Lavorazioni_MES.manodopera_attrezzaggio()

        'Dashboard_MU.TextBox1.Text = Nothing
        'If Dashboard_MU.Button11.BackColor = Color.Red Then
        '    Autocontrollo.carica_checklist_autocontrollo()
        '    Autocontrollo.Show()
        'End If
        'Me.Close()
    End Sub

    Sub inserisci_percentuale()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "UPDATE MANODOPERA Set percentuale_lavorazione=" & percentuale & " WHERE ID ='" & Lavorazioni_MES.id & "'"
        CMD_SAP.ExecuteNonQuery()
        cnn.Close()


    End Sub

    Private Sub Form106_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class