Imports System.Data.SqlClient
Imports Tirelli.Form_gestione_utente

Public Class Form_impostazioni_ticket

    Public iniziazione As String
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Sub inizializza_form()
        compila_datagridview()
        iniziazione = "N"
    End Sub
    Sub compila_datagridview()
        inserisci_reparti_datagridview(DataGridView, TextBox1.Text)
    End Sub

    Sub inserisci_reparti_datagridview(par_datagridview As DataGridView, par_reparto As String)
        Dim filtro_reparto As String
        If par_reparto = "" Then

            filtro_reparto = ""
        Else
            filtro_reparto = " and t0.descrizione Like '%%" & par_reparto & "%%'"
        End If


        par_datagridview.Rows.Clear()

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT  t0.[Id_Reparto]
      ,t0.[Descrizione]
      ,t0.[Administrator]
      ,t0.[Fittizio]
      ,t0.[SAP_ID_Reparto]
	  ,t1.name as 'Nome_reparto_1'
      ,t0.[SAP_ID_Reparto_2]
	  ,t2.name as 'Nome_reparto_2'
	   ,t0.[Mail_1]
      ,t0.[Mail_2]
      ,t0.[Mail_3]
      ,t0.[TIR01]
      ,t0.[BRB01]
  FROM [TIRELLI_40].[DBO].COLL_Reparti t0
  left join [TIRELLI_40].[dbo].oudp t1 on t1.code=t0.[SAP_ID_Reparto]
  left join [TIRELLI_40].[dbo].oudp t2 on t2.code=t0.[SAP_ID_Reparto_2]
where 0 = 0 " & filtro_reparto & "
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader("Id_Reparto"), cmd_SAP_reader("Descrizione"), cmd_SAP_reader("Administrator"), cmd_SAP_reader("Fittizio"), cmd_SAP_reader("SAP_ID_Reparto"), cmd_SAP_reader("Nome_reparto_1"), cmd_SAP_reader("SAP_ID_Reparto_2"), cmd_SAP_reader("Nome_reparto_2"), cmd_SAP_reader("Mail_1"), cmd_SAP_reader("Mail_2"), cmd_SAP_reader("Mail_3"))


        Loop
        cmd_SAP_reader.Close()
        CNN.Close()


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        compila_datagridview()
    End Sub

    Private Sub Form_impostazioni_ticket_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Form_gestione_utente.inserisci_reparti(ComboBox1)
        Form_gestione_utente.inserisci_reparti(ComboBox3)
        inizializza_form()
    End Sub



    Private Sub DataGridView_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView.CellClick
        If e.RowIndex >= 0 Then
            compila_dettagli_anagrafici(DataGridView.Rows(e.RowIndex).Cells(columnName:="ID_reparto").Value)
        End If
    End Sub

    Sub compila_dettagli_anagrafici(par_empid As String)

        TextBox5.Text = dati_anagrafici_reparto_ticket(par_empid).descrizione
        ComboBox3.Text = dati_anagrafici_reparto_ticket(par_empid).nome_reparto_1
        ComboBox1.Text = dati_anagrafici_reparto_ticket(par_empid).nome_reparto_2
        TextBox7.Text = dati_anagrafici_reparto_ticket(par_empid).Mail_1
        TextBox8.Text = dati_anagrafici_reparto_ticket(par_empid).Mail_2
        TextBox2.Text = dati_anagrafici_reparto_ticket(par_empid).Mail_3

        TextBox3.Text = par_empid

    End Sub

    Public Function dati_anagrafici_reparto_ticket(par_id As Integer) As Dettaglirepartoticket



        Dim dettagli As New Dettaglirepartoticket()


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_sap_tirelli As New SqlCommand
        Dim cmd_sap_tirelli_reader As SqlDataReader

        CMD_sap_tirelli.Connection = Cnn
        CMD_sap_tirelli.CommandText = "SELECT  t0.[Id_Reparto]
      ,t0.[Descrizione]
      ,t0.[Administrator]
      ,t0.[Fittizio]
      ,t0.[SAP_ID_Reparto]
	  ,t1.name as 'Nome_reparto_1'
      ,coalesce(t0.[SAP_ID_Reparto_2],'') as 'SAP_ID_Reparto_2'
	  ,coalesce(t2.name,'') as 'Nome_reparto_2'
	   ,t0.[Mail_1]
      ,t0.[Mail_2]
      ,t0.[Mail_3]
      ,t0.[TIR01]
      ,t0.[BRB01]
  FROM [TIRELLI_40].[DBO].COLL_Reparti t0
  left join [TIRELLI_40].[dbo].oudp t1 on t1.code=t0.[SAP_ID_Reparto]
  left join [TIRELLI_40].[dbo].oudp t2 on t2.code=t0.[SAP_ID_Reparto_2]
where t0.[Id_Reparto]='" & par_id & "' "

        cmd_sap_tirelli_reader = CMD_sap_tirelli.ExecuteReader


        If cmd_sap_tirelli_reader.Read() Then

            dettagli.Descrizione = cmd_sap_tirelli_reader("Descrizione")
            dettagli.id_reparto_1 = cmd_sap_tirelli_reader("SAP_ID_Reparto")
            dettagli.nome_reparto_1 = cmd_sap_tirelli_reader("Nome_reparto_1")
            dettagli.id_reparto_2 = cmd_sap_tirelli_reader("SAP_ID_Reparto_2")
            dettagli.nome_reparto_2 = cmd_sap_tirelli_reader("Nome_reparto_2")
            dettagli.Mail_1 = cmd_sap_tirelli_reader("Mail_1")
            dettagli.Mail_2 = cmd_sap_tirelli_reader("Mail_2")
            dettagli.Mail_3 = cmd_sap_tirelli_reader("Mail_3")



        End If


        cmd_sap_tirelli_reader.Close()
        Cnn.Close()
        Return dettagli
    End Function

    Public Class Dettaglirepartoticket
        Public descrizione As String
        Public id_reparto_1 As String
        Public nome_reparto_1 As String
        Public id_reparto_2 As String
        Public nome_reparto_2 As String
        Public Mail_1 As String
        Public Mail_2 As String
        Public Mail_3 As String
    End Class

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox3.SelectedIndex < 0 Or ComboBox1.SelectedIndex < 0 Then
            MsgBox("Selezionare entrambi i reparti")
        Else
            Aggiorna_dati_anagrafici(TextBox3.Text, TextBox5.Text, Form_gestione_utente.Elenco_Reparti(ComboBox3.SelectedIndex), Form_gestione_utente.Elenco_Reparti(ComboBox1.SelectedIndex), TextBox7.Text, TextBox8.Text, TextBox2.Text)
            MsgBox("Reparto aggiornato con successo")
            compila_datagridview()
        End If

    End Sub

    Sub Aggiorna_dati_anagrafici(par_codice_reparto As String, par_descrizione As String, par_reparto_1 As String, par_reparto_2 As String, par_mail_1 As String, par_mail_2 As String, par_mail_3 As String)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli


        CNN.Open()

        Dim Cmd_SAP As New SqlCommand

        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "update [TIRELLI_40].[DBO].COLL_Reparti
set descrizione ='" & par_descrizione & "'
,SAP_ID_Reparto ='" & par_reparto_1 & "'
,SAP_ID_Reparto_2 ='" & par_reparto_2 & "'
,Mail_1 = '" & par_mail_1 & "'
,Mail_2 = '" & par_mail_2 & "'
,Mail_3 = '" & par_mail_3 & "'

 where Id_Reparto = '" & par_codice_reparto & "'"
        Cmd_SAP.ExecuteNonQuery()

        CNN.Close()


    End Sub

    Sub Elimina_reparto(par_codice_reparto As String)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli


        CNN.Open()

        Dim Cmd_SAP As New SqlCommand

        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "DELETE [TIRELLI_40].[DBO].COLL_Reparti


 where Id_Reparto = '" & par_codice_reparto & "'"
        Cmd_SAP.ExecuteNonQuery()

        CNN.Close()


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Chiedi se l'utente vuole aprire il video
        Dim result As DialogResult = MessageBox.Show("Sei sicuro di voler cancellare il reparto dai tickets?", "Cancella", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            Elimina_reparto(TextBox3.Text)
        End If
        compila_datagridview()

    End Sub

    Sub Inserisci_dati_anagrafici(par_codice_reparto As String, par_descrizione As String, par_reparto_1 As String, par_reparto_2 As String, par_mail_1 As String, par_mail_2 As String, par_mail_3 As String)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli

        CNN.Open()

        Dim Cmd_SAP As New SqlCommand

        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].COLL_Reparti (Id_Reparto, descrizione, SAP_ID_Reparto, SAP_ID_Reparto_2, Mail_1, Mail_2, Mail_3,administrator,fittizio) " &
                              "VALUES ('" & par_codice_reparto & "', '" & par_descrizione & "', '" & par_reparto_1 & "', '" & par_reparto_2 & "', '" & par_mail_1 & "', '" & par_mail_2 & "', '" & par_mail_3 & "',1,0)"
        Cmd_SAP.ExecuteNonQuery()

        CNN.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If ComboBox3.SelectedIndex < 0 Or ComboBox1.SelectedIndex < 0 Then
            MsgBox("Selezionare entrambi i reparti")
        Else
            Inserisci_dati_anagrafici(trova_codice_reparto_tickets(), TextBox5.Text, Form_gestione_utente.Elenco_Reparti(ComboBox3.SelectedIndex), Form_gestione_utente.Elenco_Reparti(ComboBox1.SelectedIndex), TextBox7.Text, TextBox8.Text, TextBox2.Text)
            MsgBox("Reparto inserito con successo")
            compila_datagridview()
        End If
    End Sub

    Public Function trova_codice_reparto_tickets()


        Dim max As Integer = 0


        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT  max(t0.[Id_Reparto]) as 'MAX'
      
  FROM [TIRELLI_40].[DBO].COLL_Reparti t0
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            max = cmd_SAP_reader("MAX") + 1

        End If
        cmd_SAP_reader.Close()
        CNN.Close()

        Return max
    End Function
End Class