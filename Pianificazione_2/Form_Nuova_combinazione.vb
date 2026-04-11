Imports System.Data.SqlClient
Imports System.Reflection.Emit
Imports Tirelli.Presence

Public Class Form_Nuova_combinazione
    Public codice_bp As String
    Public codice_bp_finale As String
    Public codice_commessa As String
    Public id_campione As Integer

    Public Elenco_campioni_combinazione(100) As String
    Public ID_combinazione_salvata As Integer
    Public numero_combinazioni As Integer
    Private num_collaudati As Integer
    Private riga As Integer

    Private Sub Form_Nuova_combinazione_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        inizializza_form()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Sub inizializza_form()
        Label2.Text = codice_commessa
        Label3.Text = Magazzino.OttieniDettagliAnagrafica(codice_commessa).Descrizione
        Label1.Text = Business_partner.Trova_business_partner(codice_bp)
        Scheda_tecnica.riempi_datagridview_campioni(DataGridView3, codice_bp, codice_bp_finale, "", Homepage.Percorso_immagini, Homepage.sap_tirelli)
        Scheda_tecnica.riempi_datagridview_combinazioni(DataGridView1, codice_commessa, Homepage.sap_tirelli)
    End Sub

    Sub aggiorna_numero_combinazioni(par_commessa As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn

        Cmd_SAP.CommandText = "update t11 set t11.Numero_combinazione=t10.Numero_Progressivo
from
(
SELECT 
    [Id_Combinazione],
    [Commessa],
    ROW_NUMBER() OVER (PARTITION BY [Commessa] ORDER BY [Id_Combinazione]) AS Numero_Progressivo
FROM 
    [TIRELLI_40].[DBO].COLL_Combinazioni
where [Commessa] ='" & par_commessa & "'

	)
	as t10 inner join  [TIRELLI_40].[DBO].COLL_Combinazioni t11 on t11.Id_Combinazione=t10.Id_Combinazione

"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Sub inserisci_nuova_combinazione(par_id_combinazione As Integer, par_commessa As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn

        Cmd_SAP.CommandText = "
delete [TIRELLI_40].[DBO].COLL_Combinazioni where id_combinazione = '" & par_id_combinazione & "'
INSERT INTO [TIRELLI_40].[DBO].COLL_Combinazioni
                (Id_Combinazione,Commessa,Campione_1,Campione_2,Campione_3,Campione_4,Campione_5,Campione_6,Campione_7,Campione_8,
                Campione_9,Campione_10,Vel_Richiesta,Note
,insertdate
,updatedate
,collaudato
,video
,ricetta
,ownerid
,tipo) VALUES (" & par_id_combinazione & ",'" &
                par_commessa & "',
                coalesce(" & Elenco_campioni_combinazione(0) & ",0),
                coalesce(" & Elenco_campioni_combinazione(1) & ",0),
               coalesce( " & Elenco_campioni_combinazione(2) & ",0),
               coalesce(" & Elenco_campioni_combinazione(3) & ",0),
              coalesce( " & Elenco_campioni_combinazione(4) & ",0),
              coalesce( " & Elenco_campioni_combinazione(5) & ",0),
             coalesce( " & Elenco_campioni_combinazione(6) & ",0),
             coalesce( " & Elenco_campioni_combinazione(7) & ",0),
             coalesce(  " & Elenco_campioni_combinazione(8) & ",0),
             coalesce( " & Elenco_campioni_combinazione(9) & ",0), 
             '" & TextBox1.Text & "',
                '" & Replace(RichTextBox1.Text, "'", " ") & "'
,getdate()
,getdate()
,0
,0
,0
,0
,'" & ComboBox1.Text & "')
"
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Public Function TROVA_MAX_COMBINAZIONE(par_commessa As String)

        Dim max As Integer = 1
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        CMD_SAP_2.CommandText = "

        Select coalesce(max(numero_combinazione),0) as 'MAx'

FROM
[TIRELLI_40].[DBO].COLL_Combinazioni
where [Commessa] ='" & par_commessa & "'

"
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            max = cmd_SAP_reader_2("max") + 1
        Else
            max = 1
        End If

        Cnn1.Close()

        Return max
    End Function 'Inserisco le risorse nella combo
    '
    Public Function TROVA_MAX_ID_COMBINAZIONE()

        Dim max As Integer = 1
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        CMD_SAP_2.CommandText = "

        Select max(ID_COMBINAZIONE) as 'MAx'

FROM
[TIRELLI_40].[DBO].COLL_Combinazioni


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



    Sub memorizza_campioni_combinazione(par_datagridview As DataGridView)
        ' Azzera l'array Elenco_campioni_combinazione
        For i As Integer = 0 To Elenco_campioni_combinazione.Length - 1
            Elenco_campioni_combinazione(i) = 0
        Next

        Dim contatore As Integer = 0
        For Each row As DataGridViewRow In DataGridView2.Rows
            ' Verifica che la riga non sia una nuova riga (di inserimento)
            If Not row.IsNewRow Then
                Elenco_campioni_combinazione(contatore) = row.Cells("DataGridViewTextBoxColumn1").Value.ToString()

            End If
            contatore += 1
        Next
    End Sub

    Sub info_combinazioni(par_datagridview As DataGridView, par_combinazione As Integer)
        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT TOP (1000) [Id_Combinazione]
      ,[Commessa]
      ,[Campione_1]
      ,[Campione_2]
      ,[Campione_3]
      ,[Campione_4]
      ,[Campione_5]
      ,[Campione_6]
      ,[Campione_7]
      ,[Campione_8]
      ,[Campione_9]
      ,[Campione_10]

      ,[Vel_Effettiva]
      ,[Vel_Richiesta]

      ,[Numero_combinazione]
,coalesce(note,'') as 'Note'
  FROM [TIRELLI_40].[DBO].COLL_Combinazioni
where id_combinazione ='" & par_combinazione & "'
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then

            Label5.Text = cmd_SAP_reader_2("Numero_combinazione")

            Label4.Text = cmd_SAP_reader_2("Numero_combinazione")
            TextBox1.Text = cmd_SAP_reader_2("Vel_Richiesta")
            If cmd_SAP_reader_2("Campione_1") <> 0 Then
                Scheda_tecnica.trova_info_campione(par_datagridview, cmd_SAP_reader_2("Campione_1"))
            End If
            If cmd_SAP_reader_2("Campione_2") <> 0 Then
                Scheda_tecnica.trova_info_campione(par_datagridview, cmd_SAP_reader_2("Campione_2"))
            End If
            If cmd_SAP_reader_2("Campione_3") <> 0 Then
                Scheda_tecnica.trova_info_campione(par_datagridview, cmd_SAP_reader_2("Campione_3"))
            End If
            If cmd_SAP_reader_2("Campione_4") <> 0 Then
                Scheda_tecnica.trova_info_campione(par_datagridview, cmd_SAP_reader_2("Campione_4"))
            End If
            If cmd_SAP_reader_2("Campione_5") <> 0 Then
                Scheda_tecnica.trova_info_campione(par_datagridview, cmd_SAP_reader_2("Campione_5"))
            End If
            If cmd_SAP_reader_2("Campione_6") <> 0 Then
                Scheda_tecnica.trova_info_campione(par_datagridview, cmd_SAP_reader_2("Campione_6"))
            End If
            If cmd_SAP_reader_2("Campione_7") <> 0 Then
                Scheda_tecnica.trova_info_campione(par_datagridview, cmd_SAP_reader_2("Campione_7"))
            End If
            If cmd_SAP_reader_2("Campione_8") <> 0 Then
                Scheda_tecnica.trova_info_campione(par_datagridview, cmd_SAP_reader_2("Campione_8"))
            End If
            If cmd_SAP_reader_2("Campione_9") <> 0 Then
                Scheda_tecnica.trova_info_campione(par_datagridview, cmd_SAP_reader_2("Campione_9"))
            End If

            RichTextBox1.Text = cmd_SAP_reader_2("Note")

        End If

        Cnn1.Close()
        par_datagridview.ClearSelection()
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            riga = e.RowIndex


                DataGridView2.Rows.Clear()
            ID_combinazione_salvata = DataGridView1.Rows(e.RowIndex).Cells(columnName:="id_combinazione").Value
            info_combinazioni(DataGridView2, DataGridView1.Rows(e.RowIndex).Cells(columnName:="id_combinazione").Value)
            Button3.Text = "AGGIORNA combinazione"
        End If

        If e.RowIndex = 0 Then
            Button8.Visible = False
            Button7.Visible = True
        ElseIf e.RowIndex > 0 And e.RowIndex < DataGridView1.RowCount - 2 Then
            Button8.Visible = True
            Button7.Visible = True
        ElseIf e.RowIndex = DataGridView1.RowCount - 2 Then
            Button8.Visible = True
            Button7.Visible = False
        End If


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Business_partner.Show()

        Business_partner.Provenienza = "Form_nuova_combinazione"
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        If e.RowIndex >= 0 Then
            id_campione = DataGridView3.Rows(e.RowIndex).Cells(columnName:="Campione_").Value


            If e.ColumnIndex = DataGridView3.Columns.IndexOf(Immagine_) Then
                Form_campione_visualizza.id_campione = DataGridView3.Rows(e.RowIndex).Cells(columnName:="Campione_").Value
                Form_campione_visualizza.Show()
                Form_campione_visualizza.inizializza_form()

            End If








        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Scheda_tecnica.trova_info_campione(DataGridView2, id_campione)
    End Sub





    Public Sub Elimina_combinazione(par_combinazione As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn

        Cmd_SAP.CommandText = "delete [TIRELLI_40].[DBO].COLL_Combinazioni WHERE ID_COMBINAZIONE='" & par_combinazione & "' "


        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If MessageBox.Show($"Sei sicuro di voler eliminare la combinazione?", "ESCI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Elimina_combinazione(ID_combinazione_salvata)
            aggiorna_numero_combinazioni(codice_commessa)
            inizializza_form()
            MsgBox("Combinazione " & Label5.Text & " eliminata con successo")
        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Label4.Text = TROVA_MAX_COMBINAZIONE(codice_commessa)
        DataGridView2.Rows.Clear()
        Button3.Text = "Inserisci NUOVA combinazione"
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        ' Verifica che l'indice della riga sia valido (non nell'intestazione)
        If e.RowIndex >= 0 Then
            ' Rimuove la riga corrispondente all'indice della riga cliccata
            If e.ColumnIndex = DataGridView2.Columns.IndexOf(X) Then
                DataGridView2.Rows.RemoveAt(e.RowIndex)

            End If

        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ComboBox1.SelectedIndex < 0 Then
            MsgBox("Selezionare prima se si tratta di prima fornitura o CDS (macchina base o no)")
            Return
        End If
        memorizza_campioni_combinazione(DataGridView2)


        If Button3.Text = "Inserisci NUOVA combinazione" Then

            inserisci_nuova_combinazione(TROVA_MAX_ID_COMBINAZIONE(), codice_commessa)

        Else
            inserisci_nuova_combinazione(ID_combinazione_salvata, codice_commessa)
        End If


        aggiorna_numero_combinazioni(codice_commessa)
        MsgBox("Combinazione aggiornata con successo")
        Scheda_tecnica.riempi_datagridview_combinazioni(Scheda_tecnica.DataGridView1, codice_commessa, Homepage.sap_tirelli)
        DataGridView2.Rows.Clear()
        inizializza_form()
        Label4.Text = TROVA_MAX_COMBINAZIONE(codice_commessa)

        Button3.Text = "Inserisci NUOVA combinazione"
    End Sub

    Private Sub TextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Dim textBox As TextBox = DirectCast(sender, TextBox)

        ' Permetti solo numeri interi in TextBox1
        If textBox Is TextBox1 Then
            If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
                e.Handled = True
            End If
        Else

        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim par_datagridview As DataGridView = DataGridView1

        Dim NUMERO_APPOGGIO As Integer = par_datagridview.Rows(riga + 1).Cells(columnName:="numero").Value

        par_datagridview.Rows(riga + 1).Cells(columnName:="numero").Value = par_datagridview.Rows(riga).Cells(columnName:="numero").Value
        par_datagridview.Rows(riga).Cells(columnName:="numero").Value = NUMERO_APPOGGIO

        Aggiorna_numero_combinazione(par_datagridview.Rows(riga + 1).Cells(columnName:="numero").Value, par_datagridview.Rows(riga + 1).Cells(columnName:="ID_COMBINAZIONE").Value)
        Aggiorna_numero_combinazione(par_datagridview.Rows(riga).Cells(columnName:="numero").Value, par_datagridview.Rows(riga).Cells(columnName:="ID_COMBINAZIONE").Value)
        Distinta_base_form.SpostaRigaGiù(DataGridView1, riga)
        riga += 1
        If riga = par_datagridview.RowCount - 2 Then
            Button7.Visible = False
            Button8.Visible = True
        End If
    End Sub

    Public Sub Aggiorna_numero_combinazione(Par_numero_combinazione As Integer, par_id_combinazione As Integer)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = CNN6

        CMD_SAP_5.CommandText = "
UPDATE T0  SET T0.NUMERO_COMBINAZIONE =" & Par_numero_combinazione & " 
FROM [TIRELLI_40].[DBO].COLL_Combinazioni t0
WHERE T0.ID_COMBINAZIONE =" & par_id_combinazione & ""
        CMD_SAP_5.ExecuteNonQuery()

        CNN6.Close()

    End Sub


    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click


        If riga = 0 Then
            Button8.Visible = False
        End If
        Dim par_datagridview As DataGridView = DataGridView1
        Dim NUMERO_APPOGGIO As Integer = par_datagridview.Rows(riga - 1).Cells(columnName:="numero").Value

        par_datagridview.Rows(riga - 1).Cells(columnName:="numero").Value = par_datagridview.Rows(riga).Cells(columnName:="numero").Value
        par_datagridview.Rows(riga).Cells(columnName:="numero").Value = NUMERO_APPOGGIO

        Aggiorna_numero_combinazione(par_datagridview.Rows(riga - 1).Cells(columnName:="numero").Value, par_datagridview.Rows(riga - 1).Cells(columnName:="ID_COMBINAZIONE").Value)
        Aggiorna_numero_combinazione(par_datagridview.Rows(riga).Cells(columnName:="numero").Value, par_datagridview.Rows(riga).Cells(columnName:="ID_COMBINAZIONE").Value)
        Distinta_base_form.SpostaRigaSu(DataGridView1, riga)
        riga -= 1
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Value = "M" Then
            DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Style.BackColor = Color.Orange
        ElseIf DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Value = "CDS" Then
            DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Style.BackColor = Color.YellowGreen
        End If
    End Sub
End Class