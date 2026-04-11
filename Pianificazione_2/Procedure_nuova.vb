Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop



Public Class Procedure_nuova
    Public Elenco_Reparti(1000) As Integer
    Public id As Integer
    Public id_n_riga As Integer
    Public codice as String
    Public sap As String = "N"
    Public sicurezza As String = "N"
    Public immagine As String
    Public stato_form As String
    Public note As String
    Public numero_file_da_eliminare As Integer
    Public riga_datagridview As Integer




    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            LinkLabel1.Text = OpenFileDialog1.FileName
        End If
    End Sub

    Sub iniziazione_form()
        carica_reparti()

        If stato_form = "Visualizza" Then
            DataGridView1.Enabled = True
            Panel4.Enabled = True
            Button2.Enabled = False
            informazioni_anagrafiche_procedura()
            Button2.Text = "Aggiorna procedura"
            carica_file()
        ElseIf stato_Form = "Nuova" Then
            DataGridView1.Enabled = False
            Panel4.Enabled = False
            Trova_ID()
            Button2.Enabled = True
            Button2.Text = "Inserisci nuova procedura"
        End If
        Label1.Text = id
    End Sub


    Sub carica_reparti()



        Dim Conn_SAP As New SqlConnection
        Dim Com_SAP As New SqlCommand
        Dim DataRead_SAP As SqlDataReader
        Dim Indice As Integer

        Indice = 0
        ComboBox1.Items.Add("")
        ' Elenco_Reparti(Indice) = ""
        Indice = Indice + 1


        Conn_SAP.ConnectionString = Homepage.sap_tirelli
        Conn_SAP.Open()
        Com_SAP.Connection = Conn_SAP
        Com_SAP.CommandText = "SELECT * FROM [TIRELLI_40].[DBO].COLL_Reparti ORDER BY Descrizione"
        DataRead_SAP = Com_SAP.ExecuteReader()


        Do While DataRead_SAP.Read()
            Elenco_Reparti(Indice) = DataRead_SAP("Id_Reparto")

            ComboBox1.Items.Add(DataRead_SAP("Descrizione"))


            Indice = Indice + 1
        Loop
        Conn_SAP.Close()
    End Sub

    Sub carica_file()
        DataGridView1.Rows.Clear()
        Dim Conn_SAP As New SqlConnection
        Dim Com_SAP As New SqlCommand
        Dim DataRead_SAP As SqlDataReader


        Conn_SAP.ConnectionString = Homepage.sap_tirelli
        Conn_SAP.Open()
        Com_SAP.Connection = Conn_SAP
        Com_SAP.CommandText = "Select  id_procedura,id_riga, nome_file, estensione_file
from [tirelli_40].dbo.procedure_file where id_procedura='" & Label1.Text & "'"
        DataRead_SAP = Com_SAP.ExecuteReader()


        Do While DataRead_SAP.Read()

            DataGridView1.Rows.Add(DataRead_SAP("id_procedura"), DataRead_SAP("id_riga"), DataRead_SAP("nome_file"), DataRead_SAP("estensione_file"))
        Loop
        Conn_SAP.Close()
    End Sub

    Private Sub Procedure_nuova_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        iniziazione_form()
    End Sub

    Sub inserisci_file()
        Dim Cnn3 As New SqlConnection
        My.Computer.FileSystem.CopyFile(LinkLabel1.Text,
    Homepage.Percorso_procedure & TextBox1.Text & "." & Strings.Right(LinkLabel1.Text, Len(LinkLabel1.Text) - InStrRev(LinkLabel1.Text, ".")), overwrite:=True)
        Trova_riga_file()
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "insert into [tirelli_40].[dbo].procedure_file (id_procedura, id_riga, nome_file,estensione_file)
                                                    values(" & Label1.Text & "," & id_n_riga & ",'" & TextBox1.Text & "','" & Strings.Right(LinkLabel1.Text, Len(LinkLabel1.Text) - InStrRev(LinkLabel1.Text, ".")) & "')"

        CMD_SAP_3.ExecuteNonQuery()
        cnn3.Close()
        MsgBox("File inserito con successo")

    End Sub

    Sub inserisci_procedura()
        Trova_ID()
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "insert into [tirelli_40].[dbo].[procedure] (id,codice,reparto,[sap/4.0],sicurezza,nome,immagine,note)
                                                    values(" & id & ",'" & codice & "','" & Elenco_Reparti(ComboBox1.SelectedIndex) & "','" & sap & "', '" & sicurezza & "','" & TextBox3.Text & "','" & immagine & "','" & note & "')"

        CMD_SAP_3.ExecuteNonQuery()
        cnn3.Close()


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If TextBox1.Text = "" Then
            MsgBox("Dare un nome al file")
        Else
            inserisci_file()
            carica_file()
        End If

    End Sub

    Sub Trova_ID()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select case when max(id)+1 is null then 1 else max(id)+1 end as 'ID' 
from [tirelli_40].[dbo].[procedure]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                id = cmd_SAP_reader_2("ID")
            Else
                id = 1
            End If
        Else
            id = 1
        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()
    End Sub

    Sub Trova_riga_file()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select case when max(id_riga)+1 is null then 1 else max(id_riga)+1 end as 'ID_riga' 
from [tirelli_40].[dbo].[procedure_file] where id_procedura=" & Label1.Text & ""

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID_riga") Is System.DBNull.Value Then
                id_n_riga = cmd_SAP_reader_2("ID_riga")
            Else
                id_n_riga = 1
            End If
        Else
            id_n_riga = 1
        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        codice = TextBox2.Text
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        If stato_form = "Nuova" Then
            If ComboBox2.SelectedIndex < 0 Then
                MsgBox("selezionare se è relativa a SAP o meno")
            ElseIf ComboBox3.SelectedIndex < 0 Then
                MsgBox("selezionare se è relativa a sicurezza o meno")
            Else
                inserisci_procedura()
                Procedure.inizializzazione()
                MsgBox("Procedura inserita con Successo")

            End If
        ElseIf stato_form = "Visualizza" Then


        End If



    End Sub

    Sub informazioni_anagrafiche_procedura()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = cnn
        CMD_SAP_docentry.CommandText = "SELECT  [ID]
      ,t0.[Codice]
      ,t0.[Reparto]
,CASE WHEN t1.descrizione IS NULL THEN '' ELSE T1.DESCRIZIONE END as 'Nome_reparto'
      ,t0.[SAP/4.0]
      ,t0.[Sicurezza]
      ,t0.[Nome]
,t0.[immagine]
,t0.note
  FROM [TIRELLI_40].[dbo].[Procedure] t0 left join [TIRELLI_40].[DBO].COLL_Reparti t1 on t0.reparto=t1.id_reparto
where t0.id=" & id & ""

        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader


        If cmd_SAP_docentry_reader.Read() Then
            Label1.Text = cmd_SAP_docentry_reader("id")
            ComboBox1.Text = cmd_SAP_docentry_reader("nome_reparto")
            ComboBox2.Text = cmd_SAP_docentry_reader("SAP/4.0")
            ComboBox3.Text = cmd_SAP_docentry_reader("Sicurezza")
            TextBox2.Text = cmd_SAP_docentry_reader("codice")
            TextBox3.Text = cmd_SAP_docentry_reader("nome")
            RichTextBox44.Text = cmd_SAP_docentry_reader("note")

        End If
        cmd_SAP_docentry_reader.Close()
        cnn.Close()


    End Sub 'Inserisco le risorse nella combo box

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        sicurezza = ComboBox3.Text
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        sap = ComboBox2.Text
    End Sub

    Private Sub RichTextBox44_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox44.TextChanged

        note = Replace(RichTextBox44.Text, "'", "''")
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.ColumnIndex = DataGridView1.Columns.IndexOf(Nome_file) Then
            If File.Exists(Homepage.Percorso_procedure & DataGridView1.Rows(e.RowIndex).Cells(columnName:="Nome_file").Value & "." & DataGridView1.Rows(e.RowIndex).Cells(columnName:="estensione_file").Value) Then
                Process.Start(Homepage.Percorso_procedure & DataGridView1.Rows(e.RowIndex).Cells(columnName:="Nome_file").Value & "." & DataGridView1.Rows(e.RowIndex).Cells(columnName:="estensione_file").Value)
            Else MsgBox("File non trovato")
            End If
        End If
        riga_datagridview = e.RowIndex
        numero_file_da_eliminare = DataGridView1.Rows(e.RowIndex).Cells(columnName:="id_riga").Value
    End Sub

    Sub elimina_file()
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "delete [tirelli_40].[dbo].[procedure_file] where id_procedura=" & Label1.Text & " and id_riga = " & numero_file_da_eliminare & ""

        CMD_SAP_3.ExecuteNonQuery()
        cnn3.Close()


    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim Question
        Question = MsgBox("Sei sicuro di voler eliminare il file numero " & numero_file_da_eliminare & " ?", vbYesNo)
        If Question = vbYes Then
            elimina_file()
            If File.Exists(Homepage.Percorso_procedure & DataGridView1.Rows(riga_datagridview).Cells(columnName:="Nome_file").Value & "." & DataGridView1.Rows(riga_datagridview).Cells(columnName:="estensione_file").Value) Then
                My.Computer.FileSystem.DeleteFile(Homepage.Percorso_procedure & DataGridView1.Rows(riga_datagridview).Cells(columnName:="Nome_file").Value & "." & DataGridView1.Rows(riga_datagridview).Cells(columnName:="estensione_file").Value)
            End If

            carica_file()
        End If
    End Sub



    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class