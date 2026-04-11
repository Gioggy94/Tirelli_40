Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word

Public Class Form_cambia_dipendente
    Public Elenco_dipendenti(1000) As Integer
    Public Elenco_Reparti(1000) As Integer
    Private Sub Form_cambia_dipendente_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Homepage.da_aggiornare_ini = "NO"
        inserisci_dipendenti()
        inserisci_reparti()
        Homepage.da_aggiornare_ini = "SI"
    End Sub

    Sub inserisci_reparti()

        Combo_Reparti.Items.Clear()

        Dim Conn_SAP As New SqlConnection
        Dim Com_SAP As New SqlCommand
        Dim DataRead_SAP As SqlDataReader
        Dim Indice As Integer

        Conn_SAP.ConnectionString = homepage.sap_tirelli
        Conn_SAP.Open()
        Com_SAP.Connection = Conn_SAP
        Com_SAP.CommandText = "SELECT * FROM [TIRELLI_40].[DBO].COLL_Reparti ORDER BY Descrizione"
        DataRead_SAP = Com_SAP.ExecuteReader()
        Indice = 0

        Do While DataRead_SAP.Read()
            Elenco_Reparti(Indice) = DataRead_SAP("Id_Reparto")

            Combo_Reparti.Items.Add(DataRead_SAP("Descrizione"))

                Indice = Indice + 1
        Loop
        Conn_SAP.Close()
    End Sub

    Sub inserisci_dipendenti()

        Combo_dipendenti.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code where t0.active='Y'  order by T0.[lastName]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            Combo_dipendenti.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()


    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Homepage.totem = Nothing Then
            MsgBox("Selezionare se si tratta di un totem o di un PC personale prima di uscire")

        ElseIf Homepage.totem = "Y" And Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto = Nothing Then

            MsgBox("Selezionare un reparto")

        ElseIf Homepage.totem = "N" And Homepage.id_salvato = Nothing Then

            MsgBox("Selezionare un utente")

        Else
            Me.Close()
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If Homepage.da_aggiornare_ini = "SI" Then
            If RadioButton1.Checked = True Then
                Homepage.totem = "Y"
                groupbox4.Visible = False
                GroupBox2.Visible = True
                GroupBox2.Enabled = True
            Else
                Homepage.totem = "N"
                groupbox4.Visible = True
                GroupBox2.Visible = True
                GroupBox2.Enabled = False
            End If
            Homepage.Aggiorna_INI_COMPUTER()
        End If
    End Sub

    Private Sub Combo_Reparti_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_Reparti.SelectedIndexChanged
        If Homepage.da_aggiornare_ini = "SI" Then


            Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto = Elenco_Reparti(Combo_Reparti.SelectedIndex)
            'Homepage.rileva_reparto(Homepage.codice_reparto)

            Homepage.Aggiorna_INI_COMPUTER()
            Pianificazione_Tickets.Lbl_Nome_Reparto.Text = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).nome_reparto
            Pianificazione_Tickets.riempi_tickets(Pianificazione_Tickets.DataGridView1)
        End If
    End Sub

    Private Sub Combo_dipendenti_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_dipendenti.SelectedIndexChanged
        If Homepage.da_aggiornare_ini = "SI" Then
            Homepage.ID_SALVATO = Elenco_dipendenti(Combo_dipendenti.SelectedIndex)

            Homepage.Aggiorna_INI_COMPUTER()
            Combo_Reparti.Text = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).nome_reparto
        End If


    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If Homepage.da_aggiornare_ini = "SI" Then
            If RadioButton1.Checked = True Then
                Homepage.totem = "Y"
                groupbox4.Visible = False
                GroupBox2.Visible = True
            Else
                Homepage.totem = "N"
                groupbox4.Visible = True
                GroupBox2.Visible = True
                GroupBox2.Enabled = False
            End If
            Homepage.Aggiorna_INI_COMPUTER()
        End If
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged

    End Sub

    Private Sub RadioButton6_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton6.CheckedChanged
        If RadioButton6.Checked = True Then
            Homepage.Centro_di_costo = "TIR01"
        End If


    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        If RadioButton7.Checked = True Then
            Homepage.Centro_di_costo = "KTF01"
        End If
    End Sub

    Private Sub RadioButton8_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton8.CheckedChanged
        If RadioButton8.Checked = True Then
            Homepage.Centro_di_costo = "BRB01"
        End If
    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
        If RadioButton9.Checked = True Then
            Homepage.Centro_di_costo = "OH01"
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton2.Checked = True Then
            Aggiorna_Informazioni_Anagrafiche()
            MsgBox("Dati anagrafici aggiornati con succeso")
        End If
    End Sub

    Sub Aggiorna_Informazioni_Anagrafiche()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli


        cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "update [TIRELLI_40].[dbo].OHEM 
set U_CODICE_pdm='" & TextBox1.Text & "' where EMPID = '" & Homepage.ID_SALVATO & "'"
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()


    End Sub 'Aggiorno la data dell'odp

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged

    End Sub
End Class