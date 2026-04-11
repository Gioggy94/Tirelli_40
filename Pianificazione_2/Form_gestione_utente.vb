Imports System.Data.SqlClient



Public Class Form_gestione_utente
    Public parametro_attivo As String
    Public iniziazione As String = "Y"
    Public Elenco_Reparti(1000) As Integer
    Public Elenco_Reparti_Tickets(1000) As Integer
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Form_gestione_utente_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        inserisci_reparti(ComboBox1)
        inserisci_reparti_tickets(ComboBox3)
        inizializza_form()
        If Homepage.ID_SALVATO = Nothing Then
            MsgBox("Selezionare un utente e salvarlo sulla 4.0 in basso a destra")
        Else
            compila_dettagli_anagrafici(Homepage.ID_SALVATO)
        End If

    End Sub

    Sub inizializza_form()
        compila_datagridview()
        iniziazione = "N"
    End Sub

    Sub compila_datagridview()
        inserisci_dipendenti_datagridview(DataGridView, TextBox1.Text, TextBox2.Text, parametro_attivo, TextBox4.Text, TextBox3.Text, TextBox10.Text)
    End Sub

    Sub inserisci_dipendenti_datagridview(par_datagridview As DataGridView, par_nome As String, par_cognome As String, par_attivo As String, par_branch As String, par_reparto As String, par_reparto_tickets As String)
        Dim filtro_nome As String
        If par_nome = "" Then
            filtro_nome = ""
        Else
            filtro_nome = " and t0.firstname Like '%%" & par_nome & "%%'"
        End If
        Dim filtro_cognome As String

        If par_cognome = "" Then
            filtro_cognome = ""
        Else
            filtro_cognome = " and t0.lastname Like '%%" & par_cognome & "%%'"
        End If

        Dim filtro_attivo As String

        If par_attivo = "Y" Or par_attivo = "N" Then
            filtro_attivo = " and t0.active = '" & par_attivo & "'"
        ElseIf par_attivo = "ALL" Then
            filtro_attivo = ""

        End If
        Dim filtro_branch As String
        If par_branch = "" Then
            filtro_branch = ""
        Else
            filtro_branch = ""
            ' filtro_branch = " and t0.costcenter Like '%%" & par_branch & "%%'"
        End If

        Dim filtro_reparto As String
        If par_reparto = "" Then
            filtro_reparto = ""
        Else
            filtro_reparto = " and t1.name Like '%%" & par_reparto & "%%'"
        End If

        Dim filtro_reparto_tickets As String
        If par_reparto_tickets = "" Then
            filtro_reparto_tickets = ""
        Else
            filtro_reparto_tickets = " and t2.descrizione Like '%%" & par_reparto_tickets & "%%'"
        End If



        par_datagridview.Rows.Clear()

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti'
, T0.[lastName] as  'Cognome'
,T0.[firstName] AS 'Nome'
,T0.[active] 
, T1.[name] as 'Reparto'

,coalesce(t2.descrizione,'') as 'Nome_reparto_tickets'
        FROM [TIRELLI_40].[dbo].OHEM T0 
left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
left join [TIRELLI_40].[DBO].coll_reparti t2 on t2.id_reparto=t0.u_reparto_tickets
where 0=0 " & filtro_nome & filtro_cognome & filtro_attivo & filtro_branch & filtro_reparto & filtro_reparto_tickets & "
order by T0.[lastName]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader("Codice dipendenti"), cmd_SAP_reader("Nome"), cmd_SAP_reader("Cognome"), cmd_SAP_reader("Reparto"), cmd_SAP_reader("Nome_reparto_tickets"), cmd_SAP_reader("Active"))


        Loop
        cmd_SAP_reader.Close()
        CNN.Close()


    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        parametro_attivo = "Y"
        If iniziazione = "N" Then
            compila_datagridview()
        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        parametro_attivo = "N"
        compila_datagridview()
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        parametro_attivo = "ALL"
        compila_datagridview()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        compila_datagridview()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        compila_datagridview()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        compila_datagridview()
    End Sub

    Public Function dati_anagrafici_dipendente(par_id As Integer) As Dettaglidipendente


        Dim dettagli As New Dettaglidipendente()


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_sap_tirelli As New SqlCommand
        Dim cmd_sap_tirelli_reader As SqlDataReader

        CMD_sap_tirelli.Connection = Cnn
        CMD_sap_tirelli.CommandText = "SELECT T0.[lastName] as 'Cognome' 
, T0.[firstName] AS 'Nome', coalesce(T1.[name],'') as 'Name' 
, case when T2.USER_CODE is null then '' else t2.user_code end AS 'Nome_sap_tirelli'

, coalesce(t0.dept,0) as 'Dept'
, coalesce(t3.descrizione,'') as 'Descrizione'
,coalesce(t3.id_reparto,'') as 'Reparto'
, case when T0.[userid] is null then '50' else t0.userid end as 'Codice_licenza_sap_tirelli'
, coalesce(t0.u_codice_pdm,'') as 'Codice_PDM'
,t0.active
,coalesce(t0.email,'') as 'email'
,coalesce(t1.name,'') as 'Nome_reparto_SAP'
,coalesce(t0.galileo,'') as 'Galileo'

        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
        LEFT JOIN [TIRELLISRLDB].DBO.OUSR T2 ON T0.[userId] = T2.[USERID] 
left join [TIRELLI_40].[DBO].coll_reparti t3 on (t3.ID_REPARTO=t0.U_REPARTO_TICKETS)
where t0.empid='" & par_id & "' "

        cmd_sap_tirelli_reader = CMD_sap_tirelli.ExecuteReader


        If cmd_sap_tirelli_reader.Read() Then

            dettagli.Cognome = cmd_sap_tirelli_reader("Cognome")
            dettagli.Nome = cmd_sap_tirelli_reader("nome")
            dettagli.Attivo = cmd_sap_tirelli_reader("active")
            dettagli.Codice_Reparto_SAP = cmd_sap_tirelli_reader("dept")
            dettagli.Nome_reparto_SAP = cmd_sap_tirelli_reader("Nome_reparto_SAP")



            dettagli.Codice_Reparto_4_0 = cmd_sap_tirelli_reader("Reparto")
            dettagli.Nome_Reparto_4_0 = cmd_sap_tirelli_reader("Descrizione")
            dettagli.Mail = cmd_sap_tirelli_reader("email")
            '    dettagli.Costcenter = cmd_sap_tirelli_reader("costcenter")
            dettagli.Costcenter = ""
            dettagli.Codice_Galileo = cmd_sap_tirelli_reader("Galileo")

        End If


        cmd_sap_tirelli_reader.Close()
        Cnn.Close()
        Return dettagli
    End Function

    Sub compila_dettagli_anagrafici(par_empid As String)
        TextBox5.Text = dati_anagrafici_dipendente(par_empid).Nome
        TextBox6.Text = dati_anagrafici_dipendente(par_empid).Cognome
        ComboBox1.Text = dati_anagrafici_dipendente(par_empid).Nome_reparto_SAP
        TextBox11.Text = dati_anagrafici_dipendente(par_empid).Codice_Galileo
        ' ComboBox2.Text = dati_anagrafici_dipendente(par_empid).Costcenter
        ComboBox3.Text = dati_anagrafici_dipendente(par_empid).Nome_Reparto_4_0

        If dati_anagrafici_dipendente(par_empid).Attivo = "Y" Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If



        TextBox7.Text = dati_anagrafici_dipendente(par_empid).Mail
        TextBox8.Text = par_empid

    End Sub

    Public Class Dettaglidipendente
        Public Nome As String
        Public Cognome As String
        Public Attivo As String
        Public Codice_Reparto_SAP As String
        Public Nome_reparto_SAP As String
        Public Codice_Reparto_4_0 As String
        Public Nome_Reparto_4_0 As String
        Public Mail As String
        Public Costcenter As String
        Public Codice_Galileo As String


    End Class

    Sub inserisci_reparti(par_combobox As ComboBox)

        par_combobox.Items.Clear()

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT T0.Code, T0.Name, T0.Remarks, T0.Father 
FROM [TIRELLI_40].[dbo].OUDP T0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_reparti(Indice) = cmd_SAP_reader("Code")
            par_combobox.Items.Add(cmd_SAP_reader("Name"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        CNN.Close()


    End Sub

    Sub inserisci_reparti_tickets(par_combobox As ComboBox)

        par_combobox.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT  [Id_Reparto]
      ,[Descrizione]
      ,[Mail_1]
      ,[Mail_2]
      ,[Administrator]
      ,[Fittizio]
      ,[SAP_ID_Reparto]
      ,[SAP_ID_Reparto_2]
      ,[Mail_3]
      ,[Descrizione_inglese]
      ,[TIR01]
      ,[BRB01]
	  ,active
  FROM [TIRELLI_40].[DBO].[COLL_Reparti]
WHERE active ='Y'
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_Reparti_Tickets(Indice) = cmd_SAP_reader("Id_Reparto")
            par_combobox.Items.Add(cmd_SAP_reader("Descrizione"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        CNN.Close()


    End Sub

    Public Function TROVA_MAX_empid()

        Dim max As Integer = 1
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        CMD_SAP_2.CommandText = "

        Select max(empid) as 'MAx'

FROM
[TIRELLI_40].[DBO].ohem


"
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            max = cmd_SAP_reader_2("max") + 1
        Else
            max = 1
        End If

        Cnn1.Close()

        Return max
    End Function 'Inserisco le risorse nella combo box


    Sub inserisci_nuovo_utente(par_empid As String, par_nome As String, par_cognome As String, par_reparto As String, par_reparto_tickets As String, par_email As String, par_attivo As String, par_galileo As String)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli


        CNN.Open()

        Dim Cmd_SAP As New SqlCommand

        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "INSERT INTO [tirelli_40].[dbo].[OHEM]
           ([empID]
           ,[lastName]
           ,[firstName]
           ,[middleName]
           ,[sex]

           ,[dept]
           ,[image]
           ,[mobile]
           ,[email]
           
           ,[Active]
           ,[CreateDate]
           ,[U_N_Badge]
           
           ,[U_Reparto_tickets]
           ,[Galileo])
     VALUES
           (" & TROVA_MAX_empid() & "
           ,'" & par_cognome & "'
           ,'" & par_nome & "'
           ,''
           ,''
        
           ,'" & par_reparto & "'
           ,''
           ,''
           ,'" & par_email & "'

           ,'" & par_attivo & "'
           ,getdate()
           ,''
           
          
           ,'" & par_reparto_tickets & "'
           ,'" & par_galileo & "')"
        Cmd_SAP.ExecuteNonQuery()




        CNN.Close()


    End Sub

    Sub ELIMINA_utente(par_empid As String)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli


        CNN.Open()

        Dim Cmd_SAP As New SqlCommand

        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "delete  [tirelli_40].[dbo].[OHEM]
           where [empID] ='" & par_empid & "'"
        Cmd_SAP.ExecuteNonQuery()




        CNN.Close()


    End Sub

    Sub Aggiorna_dati_anagrafici(par_empid As String, par_nome As String, par_cognome As String, par_reparto As String, par_reparto_tickets As String, par_email As String, par_attivo As String, par_galileo As String)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli


        CNN.Open()

        Dim Cmd_SAP As New SqlCommand

        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "update [TIRELLI_40].[dbo].OHEM
set firstname ='" & par_nome & "'
,lastname ='" & par_cognome & "'
,dept ='" & par_reparto & "'
,email = '" & par_email & "'
,active = '" & par_attivo & "'
,u_reparto_tickets='" & par_reparto_tickets & "'
,Galileo='" & par_galileo & "'
 where EMPID = '" & par_empid & "'"
        Cmd_SAP.ExecuteNonQuery()

        CNN.Close()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim attivo As String
        If CheckBox1.Checked = True Then
            attivo = "Y"
        Else
            attivo = "N"
        End If
        Aggiorna_dati_anagrafici(TextBox8.Text, TextBox5.Text, TextBox6.Text, Elenco_Reparti(ComboBox1.SelectedIndex), Elenco_Reparti_Tickets(ComboBox3.SelectedIndex), TextBox7.Text, attivo, TextBox11.Text)
        MsgBox("Utente aggiornato con successo")
        compila_datagridview()
    End Sub

    Private Sub DataGridView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView.CellContentClick

    End Sub

    Private Sub DataGridView_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView.CellClick
        If e.RowIndex >= 0 Then
            compila_dettagli_anagrafici(DataGridView.Rows(e.RowIndex).Cells(columnName:="Codice").Value)
        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox8.Text = "" Then
            MsgBox("Selezionare un utente")
        Else
            Homepage.ID_SALVATO = TextBox8.Text

            Homepage.Aggiorna_INI_COMPUTER()
            Homepage.Enabled = True
            MsgBox("Utente salvato con successo")
        End If

    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        compila_datagridview()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        compila_datagridview()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If MessageBox.Show($"Sei sicuro di voler aggiungere e non aggiornare questo utente?", "AGGIUNGI utente", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            Dim attivo As String
            If CheckBox1.Checked = True Then
                attivo = "Y"
            Else
                attivo = "N"
            End If
            inserisci_nuovo_utente(TextBox8.Text, TextBox5.Text, TextBox6.Text, Elenco_Reparti(ComboBox1.SelectedIndex), Elenco_Reparti_Tickets(ComboBox3.SelectedIndex), TextBox7.Text, attivo, TextBox11.Text)
            MsgBox("Utente inserito con successo")
            compila_datagridview()

        End If


    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If MessageBox.Show($"Sei sicuro di voler eliminare per sempre questo utente?", "Elimina utente", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            ELIMINA_utente(TextBox8.Text)
            MsgBox("Utente eliminato con successo")
            compila_datagridview()
        End If
    End Sub
End Class