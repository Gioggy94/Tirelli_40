Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Public Class Autocontrollo
    Public tipo_autocontrollo As Integer
    Public ID As Integer
    Public ID_record_cq As Integer
    Public Elenco_dipendenti(1000) As String
    Public Codicedip As Integer
    Public risorsa As String
    Public tipo_macchina As String
    Public tipo_lav As String



    Public autocontrollo_attrezzaggio_necessario As String
    Public autocontrollo_lavorazione_necessario As String



    Sub carica_checklist_autocontrollo()

        DataGridView_AUTOCONTROLLO.Rows.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()


        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = cnn

        CMD_SAP_docentry.CommandText = " select case when t10.itemcode is null then '' else t10.itemcode end as 'Itemcode', t10.resgrpcod, t10.controllo, t10.id
from
(
select * from [TIRELLI_40].[DBO].autocontrollo_config
where tipo_controllo=1 and resgrpcod=2 and itemcode is null

union all

select * from [TIRELLI_40].[DBO].autocontrollo_config
where tipo_controllo=1 and itemcode='" & LabelCodiceSAP.Text & "'
)
as t10
order by id"


        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader
        Do While cmd_SAP_docentry_reader.Read()

            DataGridView_AUTOCONTROLLO.Rows.Add(cmd_SAP_docentry_reader("itemcode"), cmd_SAP_docentry_reader("id"), cmd_SAP_docentry_reader("Controllo"))

        Loop
        cmd_SAP_docentry_reader.Close()
        cnn.Close()


    End Sub

    Sub carica_checklist_autocontrollo_completata()

        DataGridView_AUTOCONTROLLO.Rows.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()


        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = cnn

        CMD_SAP_docentry.CommandText = "Select t1.id, t1.controllo, case when t0.ok = 'Y' then 'X' else'' end as 'ok', case when t0.np = 'Y' then 'X' else'' end as 'NP', case when t0.D = 'Y' then 'X' else'' end as 'D', case when t0.NC = 'Y' then 'X' else'' end as 'NC'
from [TIRELLI_40].[DBO].autocontrollo t0 left join [TIRELLI_40].[DBO].autocontrollo_config t1 on t0.id_config=t1.id
where t0.docnum=" & Label_ordine_SAP.Text & " and t0.tipo_autocontrollo=" & tipo_autocontrollo & " and t0.itemcode='" & risorsa & "' order by t1.id"


        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader
        Do While cmd_SAP_docentry_reader.Read()

            DataGridView_AUTOCONTROLLO.Rows.Add("", cmd_SAP_docentry_reader("id"), cmd_SAP_docentry_reader("Controllo"), cmd_SAP_docentry_reader("OK"), cmd_SAP_docentry_reader("NP"), cmd_SAP_docentry_reader("D"), cmd_SAP_docentry_reader("NC"))

        Loop
        cmd_SAP_docentry_reader.Close()
        cnn.Close()


    End Sub


    Private Sub DataGridView_AUTOCONTROLLO_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_AUTOCONTROLLO.CellClick
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = DataGridView_AUTOCONTROLLO.Columns.IndexOf(OK) Then
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="OK").Value = "X"
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="OK").Style.BackColor = Color.Lime
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NP").Value = Nothing
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NP").Style.BackColor = Color.White
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="D").Value = Nothing
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="D").Style.BackColor = Color.White
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Value = Nothing
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Style.BackColor = Color.White
            ElseIf e.ColumnIndex = DataGridView_AUTOCONTROLLO.Columns.IndexOf(NP) Then
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="OK").Value = Nothing
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="OK").Style.BackColor = Color.White
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NP").Value = "X"
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NP").Style.BackColor = Color.Yellow
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="D").Value = Nothing
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="D").Style.BackColor = Color.White
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Value = Nothing
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Style.BackColor = Color.White
            ElseIf e.ColumnIndex = DataGridView_AUTOCONTROLLO.Columns.IndexOf(D) Then
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="OK").Value = Nothing
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="OK").Style.BackColor = Color.White
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NP").Value = Nothing
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NP").Style.BackColor = Color.White
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="D").Value = "X"
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="D").Style.BackColor = Color.Orange
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Value = Nothing
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Style.BackColor = Color.White

            ElseIf e.ColumnIndex = DataGridView_AUTOCONTROLLO.Columns.IndexOf(NC) Then

                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="OK").Value = Nothing
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="OK").Style.BackColor = Color.White
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NP").Value = Nothing
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NP").Style.BackColor = Color.White
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="D").Value = Nothing
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="D").Style.BackColor = Color.White
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Value = "X"
                DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Style.BackColor = Color.Red

            End If

        End If
    End Sub

    Sub inserisci_manodopera_Cq()
        Trova_ID()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = cnn


        CMD_SAP.CommandText = "delete manodopera from manodopera where tipo_documento='ODP' and docnum=" & Label_ordine_SAP.Text & " and risorsa='R00571' and risorsa_2 ='" & risorsa & "'"

        CMD_SAP.ExecuteNonQuery()

        CMD_SAP.CommandText = "insert into manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,start,stop,consuntivo, tipologia_lavorazione,risorsa_2) 
values (" & ID & ",'ODP'," & Label_ordine_SAP.Text & ",'" & Codicedip & "','R00571',getdate(),convert(varchar, getdate(), 108),convert(varchar, getdate(), 108)," & TextBox4.Text & ",'" & tipo_lav & "','" & risorsa & "')"


        CMD_SAP.ExecuteNonQuery()
        cnn.Close()
    End Sub
    Sub CHECK_FLAG()

        Dim I = 0
        Dim OK = 0
        Dim NP = 0
        Dim D = 0
        Dim NC = 0

        Do While I < DataGridView_AUTOCONTROLLO.RowCount


            If DataGridView_AUTOCONTROLLO.Rows(I).Cells(columnName:="OK").Value = "X" Then
                OK = OK + 1
            ElseIf DataGridView_AUTOCONTROLLO.Rows(I).Cells(columnName:="np").Value = "X" Then
                NP = NP + 1
            ElseIf DataGridView_AUTOCONTROLLO.Rows(I).Cells(columnName:="D").Value = "X" Then
                D = D + 1
            ElseIf DataGridView_AUTOCONTROLLO.Rows(I).Cells(columnName:="NC").Value = "X" Then
                NC = NC + 1
            End If

            I = I + 1

        Loop

        If OK + NP + D + NC = DataGridView_AUTOCONTROLLO.RowCount Then


            Me.Hide()

            inserisci_manodopera_Cq()
            inserisci_RECORD_Cq()
            Dashboard_MU_New.docnum = Label_ordine_SAP.Text
            Dashboard_MU_New.risorsa = risorsa

            Dashboard_MU_New.CHECK_AUTOCONTROLLO_CARICATI()

            TextBox4.Text = Nothing
            Button_completato.Enabled = False
            Button1.Enabled = False


        Else
            MsgBox("Convalidare tutti i controlli oppure chiedere la deroga per proseguire ")

        End If

    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = PASSWORD.Text Then
            GroupBox4.Visible = True
            GroupBox5.Visible = True
            TextBox1.Text = Nothing
            Inserimento_dipendenti()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Panel6.Visible = True
        PASSWORD.Text = Rnd() * 10000000
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Panel6.Visible = False
        Panel9.Visible = False
        Panel14.Visible = False
        Dashboard_MU_New.CHECK_AUTOCONTROLLO_CARICATI()
        inserisci_RECORD_Cq()
        Me.Hide()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        Button_completato.Enabled = True
        Button1.Enabled = True
    End Sub

    Private Sub Button_completato_Click(sender As Object, e As EventArgs) Handles Button_completato.Click
        CHECK_FLAG()

    End Sub

    Sub inserisci_RECORD_Cq()
        Dim i As Integer
        Dim OK_STRING As String
        Dim NP_STRING As String
        Dim D_STRING As String
        Dim NC_STRING As String
        Dim derogatore As Integer
        Dim tempo_deroga As Integer
        Dim testo_deroga As String
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "delete autocontrollo from [TIRELLI_40].[DBO].autocontrollo where tipo='ODP' and docnum='" & Label_ordine_SAP.Text & "' and itemcode='" & risorsa & "' and tipo_autocontrollo= '" & tipo_autocontrollo & "'"
        CMD_SAP.ExecuteNonQuery()
        cnn.Close()


        Do While i < DataGridView_AUTOCONTROLLO.RowCount
            Trova_ID_Record_CQ()
            cnn.ConnectionString = Homepage.sap_tirelli
            cnn.Open()

            ' Dim CMD_SAP As New SqlCommand
            CMD_SAP.Connection = cnn
            If DataGridView_AUTOCONTROLLO.Rows(i).Cells(columnName:="OK").Value = "X" Then
                OK_STRING = "Y"
            Else
                OK_STRING = "N"
            End If
            If DataGridView_AUTOCONTROLLO.Rows(i).Cells(columnName:="NP").Value = "X" Then
                NP_STRING = "Y"
            Else
                NP_STRING = "N"
            End If
            If DataGridView_AUTOCONTROLLO.Rows(i).Cells(columnName:="D").Value = "X" Then
                D_STRING = "Y"
                derogatore = Codicedip
                tempo_deroga = TextBox4.Text
                testo_deroga = TextBox3.Text

            Else
                D_STRING = "N"
                derogatore = Nothing
                tempo_deroga = Nothing
                testo_deroga = Nothing
            End If
            If DataGridView_AUTOCONTROLLO.Rows(i).Cells(columnName:="NC").Value = "X" Then
                NC_STRING = "Y"
            Else
                NC_STRING = "N"
            End If



            CMD_SAP.CommandText = "insert into [TIRELLI_40].[DBO].autocontrollo (id,tipo,docnum,dipendente,itemcode,resgrpcod,id_config,tipo_autocontrollo, data,ora,ok,np,d,nc, derogatore, descrizione_deroga,tempo) 
             values (" & ID_record_cq & ",'ODP','" & Label_ordine_SAP.Text & "','" & Codicedip & "','" & risorsa & "','" & tipo_macchina & "','" & DataGridView_AUTOCONTROLLO.Rows(i).Cells(columnName:="id_config").Value & "','" & tipo_autocontrollo & "',getdate(),convert(varchar, getdate(), 108),'" & OK_STRING & "','" & NP_STRING & "','" & D_STRING & "','" & NC_STRING & "', '" & derogatore & "','" & testo_deroga & "','" & tempo_deroga & "'  )"
            CMD_SAP.ExecuteNonQuery()
            cnn.Close()
            i = i + 1
        Loop
        Codicedip = Nothing
        TextBox3.Text = Nothing
        TextBox4.Text = Nothing
    End Sub

    Sub Trova_ID()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select max(id)+1 as 'ID' from manodopera"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                ID = cmd_SAP_reader_2("ID")
            Else
                ID = 1
            End If

        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()
    End Sub

    Sub Trova_ID_Record_CQ()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select max(id)+1 as 'ID' from [TIRELLI_40].[DBO].autocontrollo"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                ID_record_cq = cmd_SAP_reader_2("ID")
            Else
                ID_record_cq = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub


    Private Sub DataGridView_AUTOCONTROLLO_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_AUTOCONTROLLO.CellFormatting

        If DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="OK").Value = "X" Then
            DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="OK").Style.BackColor = Color.Lime
        Else
            DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="OK").Style.BackColor = Color.White
        End If

        If DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NP").Value = "X" Then
            DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NP").Style.BackColor = Color.Yellow
        Else
            DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NP").Style.BackColor = Color.White
        End If

        If DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="D").Value = "X" Then
            DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="D").Style.BackColor = Color.Orange
        Else
            DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="D").Style.BackColor = Color.White
        End If

        If DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Value = "X" Then
            DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Style.BackColor = Color.Red
        Else
            DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Style.BackColor = Color.White
        End If

        If DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Value = "X" Then
            DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Style.BackColor = Color.Red
        Else
            DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="NC").Style.BackColor = Color.White
        End If

        If DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).Cells(columnName:="ITEMCODE").Value <> "" Then
            DataGridView_AUTOCONTROLLO.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Orange


        End If


    End Sub

    Sub Inserimento_dipendenti()
        Combodipendenti.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
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
            Combodipendenti.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Private Sub Combodipendenti_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combodipendenti.SelectedIndexChanged


        Codicedip = Elenco_dipendenti(Combodipendenti.SelectedIndex)


    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) 
        Me.Hide()
    End Sub

    Private Sub Button_disegno_Click(sender As Object, e As EventArgs) Handles Button_disegno.Click
        Try
            Process.Start(Homepage.percorso_disegni_generico & "PDF\"  & Button_disegno.Text & ".PDF")
        Catch ex As Exception
            MsgBox("Il disegno " & Button_disegno.Text & " non è ancora stato processato")
        End Try
    End Sub









End Class