Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class Lavorazioni_MES
    Public id As Integer
    Public Elenco_dipendenti_MES(1000) As String
    Public riga As Integer
    Public numero_lavorazione As Integer
    Public ultima_lavorazione As String
    Public ID_PADRE As Integer


    Sub inserisci_start_odp(par_dipendente As String)
        Trova_ID()
        Dim tempo As Integer
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        If Dashboard_MU_New.mu = 1 Then
            numero_lavorazione_macchina()
            If Dashboard_MU_New.tipo_lav = "A" Then
                tempo = Dashboard_MU_New.TextBox1.Text
            ElseIf Dashboard_MU_New.tipo_lav = "L" Then
                tempo = Dashboard_MU_New.TextBox2.Text
            End If

            CMD_SAP.CommandText = "insert into manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,start,stop,consuntivo, tipologia_lavorazione,ordine_lavorazione) 
        values (" & id & ",'ODP'," & Label_numero_ODP_F.Text & ",'" & par_dipendente & "','" & Pianificazione.risorsa & "',getdate(),convert(varchar, getdate(), 108),convert(varchar, getdate(), 108)," & tempo & ",'" & Dashboard_MU_New.tipo_lav & "','" & numero_lavorazione & "')"
        Else

            CMD_SAP.CommandText = "insert into manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,start,consuntivo, tipologia_lavorazione) 
        values (" & id & ",'ODP'," & Label_numero_ODP_F.Text & ",'" & par_dipendente & "','" & Pianificazione.risorsa & "',getdate(),convert(varchar, getdate(), 108),0,'" & Dashboard_MU_New.tipo_lav & "')"
        End If

        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()


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


    Sub inserisci_start_oc(par_dipendente As String)
        Trova_ID()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn




        CMD_SAP.CommandText = "insert into manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,start,consuntivo, tipologia_lavorazione) 
values (" & id & ",'OC'," & FORM6.ODP & ",'" & par_dipendente & "','" & Pianificazione.risorsa & "',getdate(),convert(varchar, getdate(), 108),0,'" & Dashboard_MU_New.tipo_lav & "')"
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

        '   Dashboard_pianificazione.Dipendente = Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex)
        FORM6.Check_Lavorazioni_aperte_dipendente(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
        'Form6.Check_Lavorazioni_aperte_dipendente()

    End Sub



    Private Sub Button_inserisci_Click(sender As Object, e As EventArgs)

        If ComboBox_risorse.SelectedIndex >= 0 And ComboBox_dipendente.SelectedIndex >= 0 Then

            FORM6.check_dipendente = "OK"
            FORM6.Check_Lavorazioni_aperte_dipendente(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))

            If FORM6.check_dipendente = "OK" Then
                Pianificazione.risorsa = FORM6.Elenco_risorse(ComboBox_risorse.SelectedIndex)
                inserisci_start_odp(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
                Me.Hide()
                FORM6.Show()
                FORM6.stato_lavorazione = "In_esecuzione"

                Try



                    '   FORM6.Cambia_stato_ODP()
                    FORM6.DataGridView_ODP.Rows(FORM6.riga).Cells(6).Value = "In_esecuzione"
                Catch ex As Exception

                End Try
                Inserimento_risorse_MES(ComboBox_risorse)
            End If

        Else
            MsgBox("Mancano delle informazioni fondamentali")
        End If

    End Sub

    Sub manodopera_attrezzaggio(par_dipendente As String)
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
values (" & id & ",'ODP'," & Dashboard_MU_New.docnum & ",'" & par_dipendente & "','R00500',getdate(),convert(varchar, getdate(), 108),convert(varchar, getdate(), 108)," & cmd_SAP_reader_2("Minuti") & "*" & Form106.percentuale & "/5,'" & Dashboard_MU_New.tipo_lav & "','" & ID_PADRE & "')"

            CMD_SAP.ExecuteNonQuery()
            Cnn.Close()

            cmd_SAP_reader_2.Close()
        End If
        Cnn4.Close()


    End Sub

    Sub inserisci_STOP()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "UPDATE MANODOPERA Set Stop=convert(varchar, getdate(), 108) WHERE ID ='" & id & "'"
        CMD_SAP.ExecuteNonQuery()
        cnn.Close()


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        If Homepage.Form_precedente = 0 Then

            Homepage.Show()

            Me.Hide()
        ElseIf Homepage.Form_precedente = 6 Then


            FORM6.Show()
            Me.Hide()
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button_start.Click
        If ComboBox_dipendente.SelectedIndex < 0 Then
            MsgBox("Scegliere un dipendente")
        Else
            If ComboBox_risorse.SelectedIndex >= 0 And ComboBox_dipendente.SelectedIndex >= 0 Then

                FORM6.check_dipendente = "OK"
                FORM6.Check_Lavorazioni_aperte_dipendente(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
                If FORM6.check_dipendente = "OK" Then

                    'da inserire quando siamo pronti

                    'Consuntivo1.check_manodopera_pregressa()
                    'If Consuntivo1.stop_ciclo = 0 Then
                    Pianificazione.risorsa = FORM6.Elenco_risorse(ComboBox_risorse.SelectedIndex)
                    Dashboard_MU_New.mu = 0
                    inserisci_start_odp(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))

                    Me.Hide()
                    FORM6.Show()
                    FORM6.stato_lavorazione = "In_esecuzione"

                    Try
                        'FORM6.Cambia_stato_ODP()
                        FORM6.DataGridView_ODP.Rows(FORM6.riga).Cells(6).Value = "In_esecuzione"
                    Catch ex As Exception

                    End Try
                    Inserimento_risorse_MES(ComboBox_risorse)

                End If

            Else
                MsgBox("Mancano delle informazioni fondamentali")
            End If
        End If
    End Sub

    Sub formatta_form_8(par_n_odp As Integer)
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.[DocNum] AS 'docnum', T0.[ItemCode] as 'Itemcode', T1.[ItemName] as 'Itemname', case when T1.[U_Disegno] is null then '' else t1.u_disegno end as 'Disegno', T0.[U_PRG_AZS_Commessa] as 'Commessa', case when T0.[U_Fase] is null then '' else t0.U_fase end as 'Fase' , T2.[ItmsGrpNam] as 'Gruppo articolo'

FROM OWOR T0 INNER JOIN OITM T1 ON T0.ITEMCODE=T1.ITEMCODE
INNER JOIN OITB T2 ON T1.[ItmsGrpCod] = T2.[ItmsGrpCod] 
WHERE T0.[DocNum] ='" & par_n_odp & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            Label_numero_ODP_F.Text = par_n_odp
            Label_Codice_ODP_F.Text = cmd_SAP_reader_2("Itemcode")
            Label_descrizione.Text = cmd_SAP_reader_2("Itemname")
            Label_commessa_F.Text = cmd_SAP_reader_2("Commessa")
            Label_disegno_F.Text = cmd_SAP_reader_2("Disegno")
            Label_fase_F.Text = cmd_SAP_reader_2("Fase")
            Label_gruppo_articolo_F.Text = cmd_SAP_reader_2("Gruppo articolo")

            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub

    Sub inserimento_dipendenti_MES(par_combobox As ComboBox, par_elenco_dipendenti() As String)

        Dim filtro_regola_distribuzione As String

        If Homepage.Centro_di_costo = "BRB01" Then
            ' filtro_regola_distribuzione = " And t0.costcenter='BRB01' "
            filtro_regola_distribuzione = ""
        Else
            filtro_regola_distribuzione = ""
        End If

        par_combobox.Items.Clear()

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli

        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "select *
from
(
select '' as 'Codice dipendenti', '' as 'Nome', '' as 'Nome 2'
union all
SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 
left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
where t0.active='Y' " & filtro_regola_distribuzione & "
)
as t0
order by t0.nome"
        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti_MES(Indice) = cmd_SAP_reader("Codice dipendenti")
            par_combobox.Items.Add(cmd_SAP_reader("Nome"))

            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        CNN.Close()
        'If Homepage.totem = "N" Then
        '    par_combobox.Text = Homepage.UTENTE_NOME_SALVATO
        'End If
    End Sub

    Sub Inserimento_risorse_MES(par_combobox As ComboBox)
        par_combobox.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli

        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "Select t0.visrescode as 'Risorsa', t0.resname as 'Nome_risorsa'
from orsc t0
where t0.resgrpcod<>5 and t0.restype='L' and t0.validfor='Y' order by t0.resname"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            FORM6.Elenco_risorse(Indice) = cmd_SAP_reader("risorsa")
            par_combobox.Items.Add(cmd_SAP_reader("Nome_risorsa"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        CNN.Close()

    End Sub 'Inserisco le risorse nella combo box

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button_stop.Click
        id = DataGridView_lavorazioni.Rows(riga).Cells(0).Value
        FORM6.ODP = DataGridView_lavorazioni.Rows(riga).Cells(2).Value
        inserisci_STOP()

        FORM6.CHIUDI_lavorazione()

        Lavorazioni_aperte(DataGridView_lavorazioni, FORM6.ODP, 0)

        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "SELECT t0.ID as 'ID', t0.docnum as 'ODP', t3.itemcode as 'Itemcode',T3.PRODNAME as 'Descrizione', T3.[U_PRG_AZS_Commessa] as 'Commessa', case when t4.u_disegno is null then '' else t4.u_disegno end as 'Disegno', t3.plannedqty as 'Quantita', T1.[LastName]+' '+T1.[FirstName] as 'Dipendente', t2.resname as 'Risorsa', t0.data as 'Data', t0.start as 'Start'
FROM MANODOPERA t0 inner join [TIRELLI_40].[dbo].ohem t1 on t1.[empID]=dipendente
inner join orsc t2 on t2.visrescode=t0.risorsa
LEFT JOIN OWOR T3 ON T3.DOCNUM=t0.docnum
left join oitm t4 on t4.itemcode=t3.itemcode
where t0.id=" & id & " AND (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='')"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then


        Else



            For Each Riga As DataGridViewRow In FORM6.DataGridView_ODP.Rows
                Dim stringa As String = Riga.Cells(0).Value.ToString
                If stringa = FORM6.ODP Then
                    'Form6.DataGridView_ODP.Rows(Riga).Cells(13).Value = "ODP"

                    FORM6.DataGridView_ODP(6, Riga.Index).Value = ""
                End If

            Next


            cmd_SAP_reader_2.Close()
        End If
        cnn1.Close()

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

                FORM6.check_dipendente = "OK"

                FORM6.Check_Lavorazioni_aperte_dipendente(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
                If FORM6.check_dipendente = "OK" Then
                    Consuntivo1.check_manodopera_pregressa()
                    If Consuntivo1.stop_ciclo = 0 Then

                        Pianificazione.risorsa = FORM6.Elenco_risorse(ComboBox_risorse.SelectedIndex)
                        inserisci_start_oc(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
                        Me.Hide()


                        'Try
                        'FORM6.Cambia_stato_ODP()
                        'FORM6.DataGridView_ODP.Rows(FORM6.riga).Cells(6).Value = "In_esecuzione"
                        'Catch ex As Exception

                        '  End Try
                        Inserimento_risorse_MES(ComboBox_risorse)

                    Else
                        inserimento_dipendenti_MES(Consuntivo1.ComboBox_dipendente, Consuntivo1.Elenco_dipendenti)
                        Consuntivo1.Inserimento_risorse()
                        Consuntivo1.Show()



                        Consuntivo1.Lavorazioni_aperte(Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), DataGridView_lavorazioni)
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

    Sub Lavorazioni_aperte(par_datagridview As DataGridView, par_docnum As Integer, par_inter As Integer)
        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        If par_inter = 0 Then


            CMD_SAP_2.Connection = Cnn1
            CMD_SAP_2.CommandText = "SELECT t0.ID as 'ID', t0.docnum as 'ODP', t0.tipo_documento as 'Tipo_documento', t3.itemcode as 'Itemcode',T3.PRODNAME as 'Descrizione', T3.[U_PRG_AZS_Commessa] as 'Commessa', case when t5.itemname is null then '' else t5.itemname end as 'Nome_commessa', case when t5.u_final_customer_name is null then '' else t5.u_final_customer_name end as 'Cliente', case when t4.u_disegno is null then '' else t4.u_disegno end as 'Disegno', t3.plannedqty as 'Quantita', T1.[LastName]+' '+T1.[FirstName] as 'Dipendente', t2.resname as 'Risorsa', t0.data as 'Data', t0.start as 'Start'
FROM MANODOPERA t0 inner join [TIRELLI_40].[dbo].ohem t1 on t1.[empID]=dipendente
inner join orsc t2 on t2.visrescode=t0.risorsa
LEFT JOIN OWOR T3 ON T3.DOCNUM=t0.docnum
left join oitm t4 on t4.itemcode=t3.itemcode
LEFT JOIN OITM T5 ON T5.ITEMCODE=T3.[U_PRG_AZS_Commessa]
where t0.docnum=" & par_docnum & " AND (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='')
order by T1.[LastName]+' '+T1.[FirstName], t0.data DESC "

            cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Else
            CMD_SAP_2.Connection = Cnn1
            CMD_SAP_2.CommandText = "SELECT t0.ID as 'ID',t0.tipo_documento as 'Tipo_documento', t0.docnum as 'ODP', t3.itemcode as 'Itemcode',T3.PRODNAME as 'Descrizione', T3.[U_PRG_AZS_Commessa] as 'Commessa', case when t5.itemname is null then '' else t5.itemname end as 'Nome_commessa', case when t5.u_final_customer_name is null then '' else t5.u_final_customer_name end as 'Cliente', case when t4.u_disegno is null then '' else t4.u_disegno end as 'Disegno', t3.plannedqty as 'Quantita', T1.[LastName]+' '+T1.[FirstName] as 'Dipendente', t2.resname as 'Risorsa', t0.data as 'Data', t0.start as 'Start'
FROM MANODOPERA t0 inner join [TIRELLI_40].[dbo].ohem t1 on t1.[empID]=dipendente
inner join orsc t2 on t2.visrescode=t0.risorsa
LEFT JOIN OWOR T3 ON T3.DOCNUM=t0.docnum
left join oitm t4 on t4.itemcode=t3.itemcode
LEFT JOIN OITM T5 ON T5.ITEMCODE=T3.[U_PRG_AZS_Commessa]
where  (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='') ANd t0.start <> ''
order by T1.[LastName]+' '+T1.[FirstName], t0.data DESC"

            cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        End If

        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(cmd_SAP_reader_2("ID"), cmd_SAP_reader_2("Tipo_documento"), cmd_SAP_reader_2("ODP"), cmd_SAP_reader_2("Itemcode"), cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("Disegno"), Math.Round(cmd_SAP_reader_2("Quantita"), 2), cmd_SAP_reader_2("Commessa"), cmd_SAP_reader_2("Nome_commessa"), cmd_SAP_reader_2("Cliente"), cmd_SAP_reader_2("Dipendente"), cmd_SAP_reader_2("Risorsa"), cmd_SAP_reader_2("Data"), cmd_SAP_reader_2("Start"))
        Loop
        cmd_SAP_reader_2.Close()

        Cnn1.Close()
        par_datagridview.ClearSelection()
    End Sub

    Private Sub Lavorazioni_MES_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Dim PERCORSO As String = Homepage.PERCORSO_statistiche_lavorazioni
        Process.Start(PERCORSO)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form_statistiche_manodopera.Show()
    End Sub
End Class