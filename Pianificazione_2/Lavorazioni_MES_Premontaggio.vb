Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Windows.Documents

Public Class Lavorazioni_MES_Premontaggio
    Public id As Integer
    Public Elenco_dipendenti_MES(1000) As String
    Public riga As Integer
    Public numero_lavorazione As Integer
    Public ultima_lavorazione As String
    Public ID_PADRE As Integer


    Sub inserisci_start_odp()
        '        Trova_ID()
        '        Dim tempo As Integer
        '        cnn.ConnectionString = homepage.sap_tirelli
        '        cnn.Open()

        '        Dim CMD_SAP As New SqlCommand
        '        CMD_SAP.Connection = cnn

        '        If Dashboard_MU.MU = 1 Then
        '            numero_lavorazione_macchina()
        '            If Dashboard_MU.tipo_lav = "A" Then
        '                tempo = Dashboard_MU.TextBox1.Text
        '            ElseIf Dashboard_MU.tipo_lav = "L" Then
        '                tempo = Dashboard_MU.TextBox2.Text
        '            End If

        '            CMD_SAP.CommandText = "insert into manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,start,stop,consuntivo, tipologia_lavorazione,ordine_lavorazione) 
        'values (" & id & ",'ODP'," & Dashboard_MU.docnum & ",'" & Dashboard_pianificazione.Dipendente & "','" & Pianificazione.risorsa & "',getdate(),convert(varchar, getdate(), 108),convert(varchar, getdate(), 108)," & tempo & ",'" & Dashboard_MU.tipo_lav & "','" & numero_lavorazione & "')"
        '        Else

        '            CMD_SAP.CommandText = "insert into manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,start,consuntivo, tipologia_lavorazione) 
        'values (" & id & ",'ODP'," & Form_Premontaggio.DataGridView_ODP.Rows(Form_Premontaggio.riga).Cells(2).Value & ",'" & Dashboard_pianificazione.Dipendente & "','" & Pianificazione.risorsa & "',getdate(),convert(varchar, getdate(), 108),0,'" & Dashboard_MU.tipo_lav & "')"
        '        End If
        '        CMD_SAP.ExecuteNonQuery()
        '        cnn.Close()

        MsgBox("Funzione disattivata, segnalare al responsabile se si vede questo messaggio")
    End Sub

    Sub inserisci_start_4_0()
        '        Trova_ID()
        '        cnn.ConnectionString = homepage.sap_tirelli
        '        cnn.Open()

        '        Dim CMD_SAP As New SqlCommand
        '        CMD_SAP.Connection = cnn

        '        numero_lavorazione_macchina()

        '        CMD_SAP.CommandText = "insert into manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,start,consuntivo, tipologia_lavorazione,ordine_lavorazione) 
        'values (" & id & ",'ODP'," & Dashboard_MU.docnum & ",'" & Dashboard_pianificazione.Dipendente & "','" & Pianificazione.risorsa & "',getdate(),convert(varchar, getdate(), 108),0,'" & Dashboard_MU.tipo_lav & "','" & numero_lavorazione & "')"

        '        CMD_SAP.ExecuteNonQuery()
        '        cnn.Close()

        MsgBox("Funzione disattivata, segnalare al responsabile se si vede questo messaggio")
    End Sub

    Sub numero_lavorazione_macchina()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t10.id, t11.risorsa, case when t11.ordine_lavorazione is null then 0 else t11.ordine_lavorazione end as 'ordine_lavorazione'
FROM
(
SELECT MAX (T0.ID) AS 'ID'
FROM MANODOPERA T0 LEFT JOIN ORSC T1 ON T0.RISORSA= T1.VISRESCODE  WHERE DOCNUM ='" & Dashboard_MU_New.docnum & "' AND T1.[ResType]='M'
)
AS T10
LEFT JOIN MANODOPERA T11 ON T10.ID=T11.ID"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then

            If cmd_SAP_reader_2("risorsa") Is System.DBNull.Value Then
                numero_lavorazione = 0
            Else
                If cmd_SAP_reader_2("risorsa") = Dashboard_MU_New.risorsa Then


                    numero_lavorazione = cmd_SAP_reader_2("ordine_lavorazione")

                Else
                    numero_lavorazione = cmd_SAP_reader_2("ordine_lavorazione") + 1
                End If
            End If


            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub


    Sub inserisci_start_oc(par_codice_dipendente As String)
        Trova_ID()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn




        CMD_SAP.CommandText = "insert into manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,start,consuntivo, tipologia_lavorazione) 
values (" & id & ",'OC'," & Form_Premontaggio.ODP & ",'" & par_codice_dipendente & "','" & Pianificazione.risorsa & "',getdate(),convert(varchar, getdate(), 108),0,'" & Dashboard_MU_New.tipo_lav & "')"
        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()
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
                id = cmd_SAP_reader_2("ID")
            Else
                id = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub

    Private Sub ComboBox_dipendente_8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_dipendente.SelectedIndexChanged

        ' Dashboard_pianificazione.Dipendente = Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex)
        Form_Premontaggio.Check_Lavorazioni_aperte_dipendente(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
        'Form_Premontaggio.Check_Lavorazioni_aperte_dipendente()

    End Sub



    Private Sub Button_inserisci_Click(sender As Object, e As EventArgs)

        If ComboBox_risorse.Text <> "" And ComboBox_dipendente.Text <> "" Then

            Form_Premontaggio.check_dipendente = "OK"
            Form_Premontaggio.Check_Lavorazioni_aperte_dipendente(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))

            If Form_Premontaggio.check_dipendente = "OK" Then
                Pianificazione.risorsa = Form_Premontaggio.Elenco_risorse(ComboBox_risorse.SelectedIndex)
                inserisci_start_odp()
                Me.Hide()
                Form_Premontaggio.Show()
                Form_Premontaggio.stato_lavorazione = "In_esecuzione"

                Try



                    Form_Premontaggio.Cambia_stato_ODP()
                    Form_Premontaggio.DataGridView_ODP.Rows(Form_Premontaggio.riga).Cells(6).Value = "In_esecuzione"
                Catch ex As Exception

                End Try
                Form_Premontaggio.Inserimento_risorse_MES()
            End If

        Else
            MsgBox("Mancano delle informazioni fondamentali")
        End If

    End Sub

    Sub manodopera_attrezzaggio(par_codice_dipendente As String)
        Dim Cnn4 As New SqlConnection
        Cnn4.ConnectionString = Homepage.sap_tirelli
        Cnn4.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn4
        CMD_SAP_2.CommandText = "SELECT    DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop)-(DATEPART(hour, t0.START)*60+DATEPART(minute, t0.START)) + t0.consuntivo AS 'MINUTI'
FROM MANODOPERA t0

where  t0.id=" & id & ""

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        ID_PADRE = id
        If cmd_SAP_reader_2.Read() = True Then

            Trova_ID()
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()

            Dim CMD_SAP As New SqlCommand
            CMD_SAP.Connection = Cnn


            CMD_SAP.CommandText = "insert into manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,start,stop,consuntivo, tipologia_lavorazione,id_padre) 
values (" & id & ",'ODP'," & Dashboard_MU_New.docnum & ",'" & par_codice_dipendente & "','R00500',getdate(),convert(varchar, getdate(), 108),convert(varchar, getdate(), 108)," & cmd_SAP_reader_2("Minuti") & "*" & Form106.percentuale & "/5,'" & Dashboard_MU_New.tipo_lav & "','" & ID_PADRE & "')"

            CMD_SAP.ExecuteNonQuery()
            Cnn.Close()

            cmd_SAP_reader_2.Close()
        End If
        Cnn4.Close()


    End Sub

    Sub inserisci_STOP()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "UPDATE MANODOPERA Set Stop=convert(varchar, getdate(), 108) WHERE ID ='" & id & "'"
        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        If Homepage.Form_precedente = 0 Then

            Homepage.Show()

            Me.Hide()
        ElseIf Homepage.Form_precedente = 6 Then


            Form_Premontaggio.Show()
            Me.Hide()
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button_start.Click
        If ComboBox_dipendente.SelectedIndex < 0 Then
            MsgBox("Scegliere un dipendente")
        Else
            If ComboBox_risorse.Text <> "" And ComboBox_dipendente.Text <> "" Then

                Form_Premontaggio.check_dipendente = "OK"
                Form_Premontaggio.Check_Lavorazioni_aperte_dipendente(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
                If Form_Premontaggio.check_dipendente = "OK" Then

                    'da inserire quando siamo pronti

                    'Consuntivo1.check_manodopera_pregressa()
                    'If Consuntivo1.stop_ciclo = 0 Then
                    Pianificazione.risorsa = Form_Premontaggio.Elenco_risorse(ComboBox_risorse.SelectedIndex)
                    inserisci_start_odp()

                    Me.Hide()
                    Form_Premontaggio.Show()
                    Form_Premontaggio.stato_lavorazione = "In_esecuzione"

                    Try
                        Form_Premontaggio.Cambia_stato_ODP()
                        Form_Premontaggio.DataGridView_ODP.Rows(Form_Premontaggio.riga).Cells(8).Value = "In_esecuzione"
                    Catch ex As Exception

                    End Try
                    Form_Premontaggio.Inserimento_risorse_MES()

                End If

            Else
                MsgBox("Mancano delle informazioni fondamentali")
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button_stop.Click
        id = DataGridView_lavorazioni.Rows(riga).Cells(2).Value
        Form_Premontaggio.ODP = DataGridView_lavorazioni.Rows(riga).Cells(2).Value
        inserisci_STOP()

        Form_Premontaggio.CHIUDI_lavorazione()

        Form_Premontaggio.Lavorazioni_aperte()


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.ID as 'ID', t0.docnum as 'ODP', t3.itemcode as 'Itemcode',T3.PRODNAME as 'Descrizione', T3.[U_PRG_AZS_Commessa] as 'Commessa', case when t4.u_disegno is null then '' else t4.u_disegno end as 'Disegno', t3.plannedqty as 'Quantita', T1.[LastName]+' '+T1.[FirstName] as 'Dipendente', t2.resname as 'Risorsa', t0.data as 'Data', t0.start as 'Start'
FROM MANODOPERA t0 inner join [TIRELLI_40].[DBO].ohem t1 on t1.[empID]=dipendente
inner join orsc t2 on t2.visrescode=t0.risorsa
LEFT JOIN OWOR T3 ON T3.DOCNUM=t0.docnum
left join oitm t4 on t4.itemcode=t3.itemcode
where t0.id=" & id & " AND (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='')"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then


        Else



            For Each Riga As DataGridViewRow In Form_Premontaggio.DataGridView_ODP.Rows
                Dim stringa As String = Riga.Cells(2).Value.ToString
                If stringa = Form_Premontaggio.ODP Then
                    'Form_Premontaggio.DataGridView_ODP.Rows(Riga).Cells(13).Value = "ODP"

                    Form_Premontaggio.DataGridView_ODP(8, Riga.Index).Value = ""
                End If

            Next


            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()

        ComboBox_dipendente.Text = ""
        ComboBox_risorse.Text = ""

    End Sub

    Private Sub DataGridView_lavorazioni_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_lavorazioni.CellClick
        If e.RowIndex >= 0 Then
            riga = e.RowIndex
            Button_stop.Show()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ComboBox_dipendente.SelectedIndex < 0 Then
            MsgBox("Scegliere un dipendente")
        Else
            If ComboBox_risorse.Text <> "" And ComboBox_dipendente.Text <> "" Then

                Form_Premontaggio.check_dipendente = "OK"

                Form_Premontaggio.Check_Lavorazioni_aperte_dipendente(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
                If Form_Premontaggio.check_dipendente = "OK" Then
                    Consuntivo1.check_manodopera_pregressa()
                    If Consuntivo1.stop_ciclo = 0 Then

                        Pianificazione.risorsa = Form_Premontaggio.Elenco_risorse(ComboBox_risorse.SelectedIndex)
                        inserisci_start_oc(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
                        Me.Hide()


                        'Try
                        'Form_Premontaggio.Cambia_stato_ODP()
                        'Form_Premontaggio.DataGridView_ODP.Rows(Form_Premontaggio.riga).Cells(6).Value = "In_esecuzione"
                        'Catch ex As Exception

                        '  End Try
                        Form_Premontaggio.Inserimento_risorse_MES()
                    Else
                        Lavorazioni_MES.inserimento_dipendenti_MES(Consuntivo1.ComboBox_dipendente, Consuntivo1.Elenco_dipendenti)
                        Consuntivo1.Inserimento_risorse()
                        Consuntivo1.Show()



                        Consuntivo1.Lavorazioni_aperte(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), Consuntivo1.DataGridView_lavorazioni)
                        Consuntivo1.ComboBox_dipendente.Text = ComboBox_dipendente.Text
                        Consuntivo1.Show()
                        Consuntivo1.Refresh()
                        Consuntivo1.Owner = Me
                        Me.Hide()

                    End If

                End If

                Else
                MsgBox("Mancano delle informazioni fondamentali")
            End If
        End If
    End Sub



End Class