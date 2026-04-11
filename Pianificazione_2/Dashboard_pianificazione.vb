Imports System.Data.SqlClient

Public Class Dashboard_pianificazione

    Public Cn As New OleDb.OleDbConnection
    Public reader As OleDb.OleDbDataReader
    Public Elenco_risorse(10000) As String
    Public Elenco_dipendenti(1000) As String
    Public attivita As String
    Private Dipendente As String
    Public unita As String
    Public risorsa_appoggio As String
    Public macchina_standard As String
    Public tempo_preass_ODP_M As String
    Public tempo_montaggio_ODP_M As String
    Public tempo_el_ODP_M As String
    Public tempo_preass_B As String
    Public tempo_montaggio_B As String
    Public tempo_el_ODP_B As String
    Public tempo_preass_ODP_M_completati as string
    Public tempo_montaggio_ODP_M_completati as String
    Public tempo_el_ODP_M_completati As String
    Public tempo_preass_ODP_M_mancanti As String
    Public tempo_montaggio_ODP_M_mancanti As String
    Public tempo_el_ODP_M_mancanti As String
    Public commessa As String
    Public id_pianificazione As Integer




    Private Sub Button1_Click_1(sender As Object, e As EventArgs)
        Me.Hide()
        Pianificazione.Show()
    End Sub

    Sub iniziazione_form()
        Label_commessa.Text = commessa
        Pianificazione_matricola_datagridview(commessa)
        info_anagrafiche_commessa(commessa)
    End Sub



    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Inserimento_risorse()
        inserimento_dipendenti()


    End Sub

    Private Sub Button_aggiungi_riga_Click(sender As Object, e As EventArgs) Handles Button_aggiungi_riga.Click

        aggiungere_riga_pianificazione()
        ComboBox_risorse.Text = ""
        TextBox_giorni_lav.Hide()

    End Sub

    Sub Pianificazione_matricola_datagridview(par_commessa As String)

        DataGridView_Risorse.Rows.Clear()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "select case when t0.id is null then 0 else t0.id end as 'ID', t0.risorsa as 'Risorsa', T2.lastNAME+' ' +substring(t2.firstname,1,1) as 'Dipendente', t0.data_i as 'Data_i', t0.data_f as 'Data_F', t1.resname as 'Nome_risorsa', case when t0.unita is null then '' else t0.unita end as 'unita', case when t0.attivita is null then '' else t0.attivita end as 'Attivita'
from [Tirelli_40].[dbo].[PIANIFICAZIONE] t0 inner join orsc t1 on t0.risorsa=t1.visrescode
LEFT JOIN [TIRELLI_40].[dbo].OHEM T2 ON T2.EMPID=T0.DIPENDENTE
Where t0.commessa='" & par_commessa & "'
order by  t0.[risorsa],t0.[Data_I], t0.id"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            DataGridView_Risorse.Rows.Add(cmd_SAP_reader("ID"), cmd_SAP_reader("Risorsa"), cmd_SAP_reader("Nome_risorsa"), cmd_SAP_reader("Dipendente"), cmd_SAP_reader("Data_I"), cmd_SAP_reader("Data_F"), cmd_SAP_reader("unita"), cmd_SAP_reader("attivita"))

        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub

    Sub inserimento_dipendenti()
        Dim Disattiva_indiretto As String = ""
        Dim Disattiva_indiretto_2 As String = ""
        ComboBox_dipendente.Items.Clear()
        If CheckBox_dipendenti.Checked = True Then

            Disattiva_indiretto = "/*"
            Disattiva_indiretto_2 = "*/"
        End If
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "select *
from
(
select '' as 'Codice dipendenti', '' as 'Nome', '' as 'Nome 2'
union all
SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code where t0.active='Y' 
)
as t0
order by t0.nome"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            ComboBox_dipendente.Items.Add(cmd_SAP_reader("Nome"))

            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()
    End Sub

    Sub Inserimento_risorse()
        ComboBox_risorse.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select t0.visrescode as 'Risorsa', t0.resname as 'Nome_risorsa'
from orsc t0
where t0.resgrpcod=5"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_risorse(Indice) = cmd_SAP_reader("risorsa")
            ComboBox_risorse.Items.Add(cmd_SAP_reader("Nome_risorsa"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()




    End Sub 'Inserisco le risorse nella combo box

    Private Sub MonthCalendar_data_inizio_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar_data_inizio.DateChanged


        TextBox_data_inizio.Text = MonthCalendar_data_inizio.SelectionStart
        MonthCalendar_data_fine.SetDate(MonthCalendar_data_inizio.SelectionStart)
        TextBox_giorni_lav.Text = 1
        ComboBox_risorse.Enabled = True


    End Sub



    Private Sub ComboBox_risorse_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_risorse.SelectedIndexChanged

        Pianificazione.risorsa = Elenco_risorse(ComboBox_risorse.SelectedIndex)

        TextBox_data_inizio.Show()



    End Sub

    Private Sub Button_aggiorna_Click(sender As Object, e As EventArgs)

        attivita = TextBox_attivita.Text
        Dim CNN As New SqlConnection
        CNN.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "Update [Tirelli_40].[dbo].[PIANIFICAZIONE] set pianificazione.risorsa='" & Pianificazione.risorsa & "', Pianificazione.dipendente='" & Dipendente & "', Pianificazione.attivita='" & attivita & "', Pianificazione.[Data_i]=CONVERT(DATETIME, '" & TextBox_data_inizio.Text & "',103), Pianificazione.[Data_f]=CONVERT(DATETIME, '" & TextBox_data_fine.Text & "',103), pianificazione.unita='" & ComboBox_unità.Text & "' where pianificazione.id=" & id_pianificazione & ""
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()

        Pianificazione_matricola_datagridview(commessa)


    End Sub

    Private Sub MonthCalendar_data_fine_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar_data_fine.DateChanged
        TextBox_data_fine.Text = MonthCalendar_data_fine.SelectionStart

    End Sub
    Sub aggiungere_riga_pianificazione()

        trova_ID()
        attivita = TextBox_attivita.Text
        If ComboBox_dipendente.Text <> "" Then
            Dipendente = Elenco_dipendenti(ComboBox_dipendente.SelectedIndex)
        Else Dipendente = ""
        End If
        Pianificazione.risorsa = Mid(Pianificazione.risorsa, 1, 6)

        If ComboBox_risorse.Text <> "" And TextBox_data_inizio.Text <> "" And TextBox_data_fine.Text <> "" Then
            Dim CNN As New SqlConnection
            CNN.ConnectionString = Homepage.sap_tirelli
            cnn.Open()
            Dim CMD_SAP As New SqlCommand
            CMD_SAP.Connection = CNN
            CMD_SAP.CommandText = "Insert into [Tirelli_40].[dbo].[PIANIFICAZIONE] (Pianificazione.id, Pianificazione.[Commessa],Pianificazione.[Risorsa],Pianificazione.[Attivita],Pianificazione.[Dipendente],Pianificazione.[Data_I], Pianificazione.[Data_F],pianificazione.unita) 
            values (" & id_pianificazione & ",'" & commessa & "', '" & Pianificazione.risorsa & "','" & attivita & "','" & Dipendente & "',CONVERT(DATETIME, '" & TextBox_data_inizio.Text & "',103),CONVERT(DATETIME, '" & TextBox_data_fine.Text & "',103),'" & ComboBox_unità.Text & "')"
            CMD_SAP.ExecuteNonQuery()
            cnn.Close()

            risorsa_appoggio = Pianificazione.risorsa
            Pianificazione_matricola_datagridview(commessa)
            Pianificazione.risorsa = risorsa_appoggio
        Else
            MsgBox("mancano delle informazioni")
        End If
    End Sub

    Sub Elimina_riga_pianificazione(par_id As Integer)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "DELETE [Tirelli_40].[dbo].[PIANIFICAZIONE] where pianificazione.ID=" & par_id & ""
        CMD_SAP.ExecuteNonQuery()
        CNN.Close()


        Pianificazione_matricola_datagridview(commessa)


    End Sub



    Private Sub Button_aggiorna_excel_Click(sender As Object, e As EventArgs) Handles Button_aggiorna_excel.Click
        Pianificazione.commessa_appoggio = commessa
        Pianificazione.Pianificazione_output()
        commessa = Pianificazione.commessa_appoggio
        MsgBox("Excel aggiornato")
    End Sub
    Public Function AddWorkDays(ByVal startDate As Date, ByVal workDays As Integer) As Date
        Dim endDate As Date = startDate.AddDays(0)
        Dim n As Integer = 0
        If workDays > 0 Then
            n = 1
        ElseIf workDays < 0 Then
            n = -1
        End If
        If n <> 0 Then
            For i = 1 To Math.Abs(workDays)
                endDate = endDate.AddDays(n)
                While (endDate.DayOfWeek = DayOfWeek.Saturday OrElse endDate.DayOfWeek = DayOfWeek.Sunday)
                    endDate = endDate.AddDays(n)
                End While
            Next
        End If

        Return endDate
    End Function


    Private Sub TextBox_giorni_lav_TextChanged(sender As Object, e As EventArgs) Handles TextBox_giorni_lav.TextChanged
        calcolo_giorni_lavorativi()
    End Sub

    Sub calcolo_giorni_lavorativi()
        Dim Giorni_LAv = TextBox_giorni_lav.Text
        If Giorni_LAv <> "" Then
            If TextBox_data_inizio.Text <> Nothing Then
                TextBox_data_fine.Text = AddWorkDays(TextBox_data_inizio.Text, Giorni_LAv)
            End If
        End If
    End Sub

    Private Sub ComboBox_dipendente_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_dipendente.SelectedIndexChanged
        Dipendente = Elenco_dipendenti(ComboBox_dipendente.SelectedIndex)
    End Sub

    Sub trova_ID()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "SELECT 'o', max(case when t0.id is null then 0 else t0.id end )+1 as 'ID' from [Tirelli_40].[dbo].[PIANIFICAZIONE] t0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then

            If Not cmd_SAP_reader("ID") Is System.DBNull.Value Then
                id_pianificazione = cmd_SAP_reader("ID")
            Else
                id_pianificazione = 1
            End If
        Else
            id_pianificazione = 1
        End If

        cnn.Close()
        cmd_SAP_reader.Close()


    End Sub

    Private Sub CheckBox_dipendenti_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox_dipendenti.CheckedChanged
        inserimento_dipendenti()
    End Sub














    Private Sub ComboBox_stato_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_stato.SelectedIndexChanged
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        Pianificazione.stato = ComboBox_stato.Text
        CNN.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].[PIANIFICAZIONE_COMMESSA]
        SET Pianificazione_commessa.stato='" & ComboBox_stato.Text & "'
        WHERE Pianificazione_commessa.commessa='" & commessa & "'"
        Cmd_SAP.ExecuteNonQuery()

        CNN.Close()

    End Sub





    Private Sub DataGridView_Risorse_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_Risorse.CellClick
        If e.RowIndex >= 0 Then


            id_pianificazione = DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="ID").Value

            Dim CNN As New SqlConnection
            CNN.ConnectionString = Homepage.sap_tirelli
            If e.ColumnIndex = DataGridView_Risorse.Columns.IndexOf(Delete) Then

                Elimina_riga_pianificazione(id_pianificazione)

                Pianificazione_matricola_datagridview(commessa)

            ElseIf e.ColumnIndex = DataGridView_Risorse.Columns.IndexOf(Modifica) Then


                Pianificazione.data_inizio = DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="Data_i").Value
                Pianificazione.data_fine = DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="Data_f").Value

                Pianificazione.data_inizio = AddWorkDays(Pianificazione.data_inizio, DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="Inizio").Value + DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="Inizio_fine").Value)
                Pianificazione.data_fine = AddWorkDays(Pianificazione.data_fine, DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="Fine").Value + DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="Inizio_fine").Value)

                DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="Data_i").Value = Pianificazione.data_inizio
                DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="Data_f").Value = Pianificazione.data_fine

                CNN.Open()

                Dim Cmd_SAP As New SqlCommand


                Cmd_SAP.Connection = CNN
                Cmd_SAP.CommandText = "Update [Tirelli_40].[dbo].[PIANIFICAZIONE] set pianificazione.unita = '" & DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="Unità").Value & "', Pianificazione.attivita='" & DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="attivita_tab").Value & "', Pianificazione.[Risorsa]= '" & DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="risorsa").Value & "', Pianificazione.[Data_i]=CONVERT(DATETIME, '" & DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="Data_i").Value & "',103), Pianificazione.[Data_f]=CONVERT(DATETIME, '" & DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="Data_f").Value & "',103) where pianificazione.id=" & id_pianificazione & ""
                Cmd_SAP.ExecuteNonQuery()

                CNN.Close()


            Else


                trova_valori(DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="ID").Value)
                TextBox_data_inizio.Text = Pianificazione.data_inizio
                TextBox_data_fine.Text = Pianificazione.data_fine

                ComboBox_dipendente.Text = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).cognome & " " & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).nome
                TextBox_attivita.Text = DataGridView_Risorse.Rows(e.RowIndex).Cells(columnName:="Attivita_Tab").Value
                Button2.Visible = True
            End If
        End If
    End Sub

    Sub trova_valori(par_pianificazione_id As Integer)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT T0.Id as 'ID', t0.risorsa as 'Risorsa', t0.data_i as 'Data_i', t0.data_f as 'Data_F', t1.resname as 'Nome_risorsa', coalesce(t0.dipendente,0) as 'Dipendente', case when t0.attivita is null then '' else t0.attivita end as 'attivita', case when t0.unita is null then 0 else t0.unita end as 'unita'
from [Tirelli_40].[dbo].[PIANIFICAZIONE] t0 inner join orsc t1 on t0.risorsa=t1.visrescode
where t0.id=" & par_pianificazione_id & ""

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() Then
            Pianificazione.risorsa = cmd_SAP_reader("risorsa")
            Pianificazione.data_inizio = cmd_SAP_reader("Data_I")
            Pianificazione.data_fine = cmd_SAP_reader("Data_F")
            Pianificazione.nome_risorsa = cmd_SAP_reader("Nome_risorsa")
            ComboBox_risorse.Text = cmd_SAP_reader("Nome_risorsa")
            Dipendente = cmd_SAP_reader("dipendente")
            attivita = cmd_SAP_reader("attivita")
            ComboBox_unità.Text = cmd_SAP_reader("unita")
        End If
        cmd_SAP_reader.Close()
        cnn.Close()
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click



        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "delete [Tirelli_40].[dbo].[PIANIFICAZIONE_KPI]
WHERE COMMESSA='" & commessa & "'"
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()


        cnn.Open()


        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "insert into [Tirelli_40].[dbo].[PIANIFICAZIONE_KPI] (Commessa, Risorsa, data_I,data_f,unita,attivita) 

SELECT Commessa as 'Commessa', RISORSA, data_I as 'Data_I', DATA_F AS 'Data_F', unita as 'Unita', attivita
FROM [Tirelli_40].[dbo].PIANIFICAZIONE
WHERE COMMESSA='" & commessa & "'"
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()



        cnn.Open()

        Cmd_SAP.Connection = CNN
        Cmd_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].[PIANIFICAZIONE_COMMESSA] SET KPI='SI' WHERE COMMESSA='" & commessa & "' "
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()


        MsgBox("Dati per KPI importati con successo")

    End Sub

    Sub info_anagrafiche_commessa(par_commessa As String)

        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1

        CMD_SAP_1.CommandText = " select  t10.commessa  , t10.Docentry,t10.KPI,


            coalesce(t14.itemname ,t10.[Nome commessa]) As 'Nome commessa',
            coalesce(t11.docnum,t10.oc) as 'ODC',
            substring(coalesce(t11.cardname,t10.Cliente),1,25) As 'Cliente',
coalesce(t13.cardname,'') as 'Cliente Finale',
            coalesce(T12.[ShipDate],t10.consegna) as 'Consegna'
, coalesce(t12.ocrcode,'') as 'ocrcode', t10.stato

from
(
SELECT t0.commessa 'Commessa' ,t0.descrizione as 'Nome commessa', t0.consegna as 'Consegna', t0.OC as 'OC'
, max(t1.docentry) as 'Docentry'
, min(t1.linenum) as 'Linenum'

,coalesce(t0.cliente,'') as 'Cliente',
coalesce(t0.kpi,'No') as 'KPI', t0.stato

from [Tirelli_40].[dbo].[PIANIFICAZIONE_COMMESSA] t0 
left join rdr1 t1 on t1.itemcode=t0.commessa
where t0.commessa='" & par_commessa & "'
group by t0.commessa  ,t0.descrizione, t0.consegna , t0.OC, t0.stato , 
t0.cliente,
t0.kpi
)
as t10 left join ordr t11 on t11.docentry=t10.docentry
left join rdr1 t12 on t12.itemcode=t10.commessa and t10.linenum=t12.linenum
left join ocrd t13 on t13.CardCode=t11.U_CodiceBP
left join oitm t14 on t14.itemcode=t12.itemcode
where 0=0 
order by t10.commessa "


        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        If cmd_SAP_reader_1.Read() Then
            Label_descrizione.Text = cmd_SAP_reader_1("Nome commessa")
            Label_cliente.Text = cmd_SAP_reader_1("cliente")
            Label_cliente_finale.Text = cmd_SAP_reader_1("cliente finale")
            Label_consegna.Text = cmd_SAP_reader_1("consegna")
            ComboBox_stato.Text = cmd_SAP_reader_1("stato")


        End If
        cmd_SAP_reader_1.Close()
        Cnn1.Close()

    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs)

        FORM6.Show()

        FORM6.inizializza_form(commessa)
        'FORM6.Inserimento_responsabile_montaggio()
        'FORM6.Inserimento_responsabile_collaudo()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub



    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Elimina_riga_pianificazione(id_pianificazione)
        aggiungere_riga_pianificazione()
        ComboBox_risorse.Text = ""



        TextBox_giorni_lav.Hide()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        aggiungi_riga_evasione_ut(commessa, Homepage.ID_SALVATO)
        MsgBox("Quantità bloccate con successo")
    End Sub

    Sub aggiungi_riga_evasione_ut(par_commessa As String, par_utente As Integer)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "

delete [tirelli_40].[dbo].[Evasione_commesse_ut] where commessa ='" & par_commessa & "'

INSERT INTO [tirelli_40].[dbo].[Evasione_commesse_ut]
           ([Commessa]
           ,[utente]
           ,[data])
     VALUES
           ('" & par_commessa & "'
           ," & par_utente & "
           ,getdate())

           UPDATE T1 SET T1.U_Q_DATA_RILASCIO=T1.[PlannedQty]
FROM OWOR T0 INNER JOIN WOR1 T1 ON T0.DOCENTRY=T1.DOCENTRY
WHERE T0.U_PRG_AZS_COMMESSA='" & par_commessa & "' and (t0.status='P' or t0.status='R')

"
        CMD_SAP.ExecuteNonQuery()
        CNN.Close()





    End Sub
End Class
