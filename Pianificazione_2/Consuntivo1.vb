Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Windows.Documents
Public Class Consuntivo1
    Public Elenco_dipendenti(1000) As String
    Public Elenco_risorse(1000) As String
    Public risorsa_manodopera As String
    Public dipendente_manodopera As String
    Public data_selezione As String
    Public Documento_SAP As String
    Public RIGA As Integer
    Public settimane_controlli As Integer
    Public delta_giorni As Integer
    Public contatore As Integer
    Public DATA_min As Date
    Public DATA_max As Date
    Public minuti_giorno As Integer
    Public stop_ciclo As Integer = 0
    Public stringa_messaggio_manodopera_mancante As String = "Sistemare la manodopera pregressa prima di fare nuovi inserimenti"
    Public oggi As String

    Sub Lavorazioni_aperte(par_dipendente As Integer, par_datagridview As DataGridView)

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "Select t10.ID as 'ID', t10.documento as 'Documento',  t10.ODP as 'ODP', t10.itemcode as 'Itemcode',case when T10.Descrizione is null then '' else t10.descrizione end as 'Descrizione', T10.[Commessa] as 'Commessa', t10.Disegno as 'Disegno', t10.Quantita as 'Quantita', T10.Dipendente as 'Dipendente', t10.Risorsa as 'Risorsa', t10.Data as 'Data', t10.start as 'Start', t10.stop as 'Stop', t10.consuntivo as 'Consuntivo' , t10.minuti as 'Minuti'
from
(
SELECT t0.ID as 'ID', t0.tipo_documento as 'Documento',
t0.docnum as 'ODP',
case when t0.tipo_documento ='ODP' then t3.itemcode else '' end as 'Itemcode',
case when t0.tipo_documento ='ODP' then T3.PRODNAME WHEN T0.TIPO_DOCUMENTO = 'ALTRO' THEN T0.TIPOLOGIA_LAVORAZIONE else '' end as 'Descrizione',
COALESCE(T3.[U_PRG_AZS_Commessa],COALESCE(T0.COMMESSA,'')) as 'Commessa', case when t4.u_disegno is null and t0.tipo_documento <>'ODP' then '' else t4.u_disegno end as 'Disegno', case when t0.tipo_documento ='ODP' then t3.plannedqty else 0 end as 'Quantita', T1.[firstName]+' '+T1.[lastName] as 'Dipendente', t2.resname as 'Risorsa', t0.data as 'Data', t0.start as 'Start', t0.stop as 'Stop', t0.consuntivo as 'Consuntivo' ,
 case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end as 'Minuti'
FROM MANODOPERA t0 inner join [TIRELLI_40].[dbo].ohem t1 on t1.[empID]=t0.dipendente
inner join orsc t2 on t2.visrescode=t0.risorsa
LEFT JOIN OWOR T3 ON T3.DOCNUM=t0.docnum
left join oitm t4 on t4.itemcode=t3.itemcode
where   T0.[dipendente]  ='" & par_dipendente & "' and cast(t0.data as varchar) Like '%%" & data_selezione & "%%'
)
as t10
order by t10.id DESC"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader_2("ID"), cmd_SAP_reader_2("Documento"), cmd_SAP_reader_2("ODP"), cmd_SAP_reader_2("Itemcode"), cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("Disegno"), Math.Round(cmd_SAP_reader_2("Quantita"), 2), cmd_SAP_reader_2("Commessa"), cmd_SAP_reader_2("Dipendente"), cmd_SAP_reader_2("Risorsa"), cmd_SAP_reader_2("Data"), cmd_SAP_reader_2("Start"), cmd_SAP_reader_2("Stop"), cmd_SAP_reader_2("Consuntivo"), cmd_SAP_reader_2("Minuti"))
        Loop


        Cnn1.Close()
        par_datagridview.ClearSelection()
    End Sub






    Private Sub DataGridView_lavorazioni_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_lavorazioni.CellClick

        If e.RowIndex >= 0 Then
            Lavorazioni_MES.id = DataGridView_lavorazioni.Rows(e.RowIndex).Cells(columnName:="ID").Value

            If e.ColumnIndex = 15 Then


                If DataGridView_lavorazioni.Rows(e.RowIndex).Cells(columnName:="start").Value Is Nothing Then
                Else

                    inserisci_STOP()
                    Lavorazioni_aperte(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), DataGridView_lavorazioni)
                End If
            End If
        End If
    End Sub

    Sub inserisci_STOP()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "UPDATE MANODOPERA SET STOP=convert(varchar, getdate(), 108) WHERE ID ='" & Lavorazioni_MES.id & "' and (STOP is null or STOP='') and (consuntivo is null or consuntivo ='')"
        CMD_SAP.ExecuteNonQuery()
        CNN.Close()


    End Sub



    Sub Inserimento_risorse()
        ComboBox_risorse.Items.Clear()
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
            Elenco_risorse(Indice) = cmd_SAP_reader("risorsa")
            ComboBox_risorse.Items.Add(cmd_SAP_reader("Nome_risorsa"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        CNN.Close()


    End Sub 'Inserisco le risorse nella combo box

    Private Sub ComboBox_risorse_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_risorse.SelectedIndexChanged
        risorsa_manodopera = Elenco_risorse(ComboBox_risorse.SelectedIndex)
        Lavorazioni_aperte(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), DataGridView_lavorazioni)
    End Sub

    Private Sub ComboBox_dipendente_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_dipendente.SelectedIndexChanged

        Lavorazioni_aperte(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), DataGridView_lavorazioni)
        Minuti_progressbar(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
        Form_statistiche_manodopera.record_per_dipendente(DataGridView4, TextBox1.Text, Homepage.Centro_di_costo, Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), "")
    End Sub









    Private Sub Button_inserisci_Click(sender As Object, e As EventArgs) Handles Button_inserisci.Click



        If ComboBox_dipendente.SelectedIndex < 0 Or TextBox_numero.Text Is Nothing Or TextBox_minuti.Text Is Nothing Or risorsa_manodopera Is Nothing Or TextBox_minuti.Text = "" Then
            MsgBox("Mancano informazioni obbligatorie")

        Else


            inserisci_consuntivo(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))



            Minuti_progressbar(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
            Lavorazioni_aperte(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), DataGridView_lavorazioni)

            Form_statistiche_manodopera.record_per_dipendente(DataGridView4, TextBox1.Text, Homepage.Centro_di_costo, Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), "")
            MsgBox("Manodopera inserita con successo")
        End If


    End Sub

    Private Sub Button_numero_Click(sender As Object, e As EventArgs) Handles Button_numero.Click
        If ComboBox_documento.Text = "ODP" Then
            Form201.elenco_ODP_aperti()
            Form201.GroupBox_OC.Hide()
            Form201.GroupBox_ODP.Show()
            Form201.DataGridView_OC.Hide()
            Form201.DataGridView_ODP.Show()
            Form201.TabControl1.SelectedIndex = 0
        ElseIf ComboBox_documento.Text = "OC" Then
            Form201.elenco_OC_aperti()
            Form201.DataGridView_ODP.Hide()
            Form201.DataGridView_OC.Show()
            Form201.GroupBox_OC.Show()
            Form201.GroupBox_ODP.Hide()
            Form201.TabControl1.SelectedIndex = 1

        ElseIf ComboBox_documento.Text = "COMMESSA" Then
            Form201.TabControl1.SelectedIndex = 2
            Scheda_commessa_Pianificazione.carica_commesse(Form201.DataGridView, Form201.TextBox5.Text, Form201.TextBox6.Text, Form201.TextBox4.Text, Form201.TextBox3.Text, "", "", Form201.TextBox2.Text, "", "")
        End If
        Form201.Show()

    End Sub

    Private Sub Form200_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PB_Minuti.Maximum = 480
        'If Minuti_Totali > 480 Then
        'PB_Minuti.Value = 480
        ' Else
        'PB_Minuti.Value = Minuti_Totali
        'End If
    End Sub



    Private Sub MonthCalendar_data_inizio_DateSelected(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar_data_inizio.DateSelected
        data_selezione = MonthCalendar_data_inizio.SelectionStart.ToString("yyyy-MM-dd")
        Minuti_progressbar(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
        Lavorazioni_aperte(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), DataGridView_lavorazioni)


    End Sub

    Sub inserisci_consuntivo(par_utente As Integer)
        If par_utente = 0 Then
            MsgBox("Selezionare un utente")
        Else

            Dim CNN As New SqlConnection
            Lavorazioni_MES.Trova_ID()
            If data_selezione Is Nothing Then
                data_selezione = Today.ToString("yyyy-MM-dd")
            End If
            If ComboBox_documento.Text = "ALTRO" Then
                CNN.ConnectionString = Homepage.sap_tirelli
                CNN.Open()

                Dim CMD_SAP As New SqlCommand
                CMD_SAP.Connection = CNN


                CMD_SAP.CommandText = "insert into manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,consuntivo,tipologia_lavorazione,commessa) 
values (" & Lavorazioni_MES.id & ",'" & ComboBox_documento.Text & "',0,'" & par_utente & "','" & risorsa_manodopera & "',cast ('" & data_selezione & "' as date)," & TextBox_minuti.Text & ",'" & TextBox_numero.Text & "','" & TextBox_numero.Text & "')"
                CMD_SAP.ExecuteNonQuery()

            ElseIf ComboBox_documento.Text = "COMMESSA" Then

                CNN.ConnectionString = Homepage.sap_tirelli
                CNN.Open()
                Dim CMD_SAP As New SqlCommand
                CMD_SAP.Connection = CNN


                CMD_SAP.CommandText = "insert into manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,consuntivo,commessa) 
values (" & Lavorazioni_MES.id & ",'" & ComboBox_documento.Text & "',0,'" & par_utente & "','" & risorsa_manodopera & "',cast ('" & data_selezione & "' as date)," & TextBox_minuti.Text & ",'" & TextBox_numero.Text & "')"
                CMD_SAP.ExecuteNonQuery()

            Else

                CNN.ConnectionString = Homepage.sap_tirelli
                CNN.Open()
                Dim CMD_SAP As New SqlCommand
                CMD_SAP.Connection = CNN


                CMD_SAP.CommandText = "insert into manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,consuntivo,commessa) 
values (" & Lavorazioni_MES.id & ",'" & ComboBox_documento.Text & "','" & TextBox_numero.Text & "','" & par_utente & "','" & risorsa_manodopera & "',cast ('" & data_selezione & "' as date)," & TextBox_minuti.Text & ",'')"
                CMD_SAP.ExecuteNonQuery()
            End If


            CNN.Close()


        End If
    End Sub

    Sub check_manodopera_pregressa()
        stop_ciclo = 0
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        contatore = 0
        settimane_controlli = 3

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "Select DATENAME(weekday,getdate() ) as 'Oggi'"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then

            oggi = cmd_SAP_reader_2("oggi")

            Do While settimane_controlli >= 0


                If oggi = "Monday" Then
                    delta_giorni = 8

                ElseIf oggi = "Tuesday" Then
                    delta_giorni = 9
                ElseIf oggi = "Wednesday" Then


                    delta_giorni = 10

                ElseIf oggi = "Thursday" Then

                    delta_giorni = 11

                ElseIf oggi = "Friday" Then

                    delta_giorni = 12

                End If


                Do While contatore < 5
                    check_manodopera()
                    contatore = contatore + 1
                Loop

                contatore = 0
                settimane_controlli = settimane_controlli - 1
            Loop

            If stop_ciclo > 0 Then
                MsgBox(stringa_messaggio_manodopera_mancante)
            End If
            stringa_messaggio_manodopera_mancante = "Sistemare la manodopera delle settimane passate prima di fare nuovi inserimenti"


            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()

    End Sub

    Sub giorno_odierno()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        contatore = 0

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "Select DATENAME(weekday,getdate() )as 'Oggi'"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then

            oggi = cmd_SAP_reader_2("oggi")

        End If
        Cnn1.Close()

    End Sub

    Sub check_manodopera()
        Dim CNN2 As New SqlConnection
        CNN2.ConnectionString = Homepage.sap_tirelli
        CNN2.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = CNN2
        CMD_SAP_2.CommandText = "Select case when sum(t10.minuti) is null then 0 else sum(t10.minuti) end  as 'Minuti', T10.Data_min, T10.Data_max, t10.startdate
from
(

SELECT t1.startdate, T0.id, T0.tipo_documento, T0.docnum as 'N documento', T4.ITEMNAME AS 'Prodotto', T3.[U_PRG_AZS_Commessa] as 'commessa',  t5.itemname as 'Nome commessa', t5.u_final_customer_name as 'Cliente', T1.LASTNAME +' ' + T1.FIRSTNAME AS 'Dipendente', t6.nAME as 'Reparto',  T0.RISORSA AS 'Risorsa', t2.resname as 'Lavorazione', T0.data,T0.start,T0.stop,T0.consuntivo, case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end as 'Minuti', T0.combinazione,T0.tipologia_lavorazione, getdate()-" & (delta_giorni - contatore) & "-(" & settimane_controlli * 7 & ") AS 'data_min', getdate()-(" & (delta_giorni - contatore) - 1 & ")-(" & settimane_controlli * 7 & ") as 'Data_max'
FROM MANODOPERA t0 LEFT JOIN [TIRELLI_40].[dbo].OHEM T1 ON T1.CODE=T0.DIPENDENTE
left join orsc t2 on t2.visrescode=t0.risorsa
left join owor t3 on t3.docnum=t0.docnum and t0.tipo_documento='ODP'
LEFT JOIN OITM t4 ON T4.ITEMCODE=T3.ITEMCODE
left join oitm t5 on t5.itemcode=T3.[U_PRG_AZS_Commessa]
left join [TIRELLI_40].[dbo].oudp t6 on t1.dept=t6.code

where t0.dipendente='" & Homepage.ID_SALVATO & "' and T1.STARTDATE<=GETDATE()-30 AND T0.data >=getdate()-" & (delta_giorni - contatore) & "-(" & settimane_controlli * 7 & ") and T0.data < getdate()-(" & (delta_giorni - contatore) - 1 & ")-(" & settimane_controlli * 7 & ")
)
as t10 WHERE T10.STARTDATE<=GETDATE()-30 GROUP BY T10.Data_min, T10.Data_max, t10.startdate"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then


            DATA_min = cmd_SAP_reader_2("Data_min")
            DATA_max = cmd_SAP_reader_2("Data_max")
            minuti_giorno = cmd_SAP_reader_2("Minuti")
        Else

            DATA_min = DateAdd("d", -(delta_giorni - contatore) - (settimane_controlli * 7), Today)
            DATA_max = DateAdd("d", -(delta_giorni - contatore - 1) - (settimane_controlli * 7), Today)
            minuti_giorno = 0
        End If

        If minuti_giorno < 400 Then

            stringa_messaggio_manodopera_mancante = stringa_messaggio_manodopera_mancante & vbCrLf & "Nel giorno = " & DATA_max & vbCrLf & "risultano inseriti Minuti = " & minuti_giorno & vbCrLf


            stop_ciclo = stop_ciclo + 1
        End If



        cmd_SAP_reader_2.Close()
        CNN2.Close()

    End Sub

    Sub Minuti_progressbar(par_dipendente As Integer)
        PB_Minuti.Minimum = 0
        PB_Minuti.Maximum = 480
        If Homepage.ID_SALVATO <> Nothing And data_selezione <> Nothing Then


            DataGridView_lavorazioni.Rows.Clear()
            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()


            Dim CMD_SAP_2 As New SqlCommand
            Dim cmd_SAP_reader_2 As SqlDataReader


            CMD_SAP_2.Connection = Cnn1
            CMD_SAP_2.CommandText = "Select case when t20.minuti<0 then 0 else t20.minuti end as 'Minuti'
from
(
Select case when sum(case when t10.minuti is null then 0 else t10.minuti end) is null then 0 else sum(case when t10.minuti is null then 0 else t10.minuti end) end  as 'Minuti'
from
(
SELECT  case when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end as 'Minuti'
FROM MANODOPERA t0 

where   T0.[dipendente]  = '" & par_dipendente & "' and cast(t0.data as varchar) Like '%%" & data_selezione & "%%'
)
as t10
) 
as t20"


            cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

            If cmd_SAP_reader_2.Read() = True Then
                If cmd_SAP_reader_2("Minuti") > PB_Minuti.Maximum Then
                    PB_Minuti.Value = PB_Minuti.Maximum
                    Lbl_Minuti.Text = cmd_SAP_reader_2("Minuti")
                Else
                    Lbl_Minuti.Text = cmd_SAP_reader_2("Minuti")
                    PB_Minuti.Value = cmd_SAP_reader_2("Minuti")
                End If


                cmd_SAP_reader_2.Close()
            End If


            Cnn1.Close()
        Else
            PB_Minuti.Value = 0
        End If
    End Sub

    Private Sub Button_clear_filters_Click(sender As Object, e As EventArgs)
        Homepage.ID_SALVATO = Nothing
        ComboBox_dipendente.Text = Nothing
        data_selezione = Nothing
        risorsa_manodopera = Nothing
        ComboBox_risorse.Text = Nothing
        Minuti_progressbar(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
        Lavorazioni_aperte(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), DataGridView_lavorazioni)
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Lavorazioni_aperte(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), DataGridView_lavorazioni)
    End Sub

    Private Sub ComboBox_documento_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_documento.SelectedIndexChanged
        Documento_SAP = ComboBox_documento.Text
        If ComboBox_documento.Text = "ALTRO" Then
            ComboBox_causale.Show()
            Button_numero.Hide()

        ElseIf ComboBox_documento.Text = "COMMESSA" Then
            Button_numero.Text = "Scegli commessa"
            ComboBox_causale.Hide()
            Button_numero.Show()

        ElseIf ComboBox_documento.Text = "OC" Then
            Button_numero.Text = "N° OC"


        Else
            Button_numero.Text = "N° ODP"
            ComboBox_causale.Hide()
            Button_numero.Show()
        End If
        TextBox_numero.Text = Nothing


    End Sub

    Private Sub TextBox_minuti_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_minuti.KeyPress
        ' Accetto solo l'immissione di numeri interi e decimali

        ' Recupero il codice ascii del tasto digitato
        ' il tasto digitato è memorizzato nella proprietà "KeyChar"
        ' dell'oggetto System.Windows.Forms.KeyPressEventArgs

        Dim KeyAscii As Short = Asc(e.KeyChar)

        ' In questo caso oltre a consentire numeri, tasto Canc
        ' e tasto BackSpace, devo consentire anche l'immissione
        ' del punto e della virgola
        If KeyAscii < 48 And KeyAscii <> 24 And KeyAscii <> 8 And e.KeyChar <> "." And e.KeyChar <> "," Then
            KeyAscii = 0
        ElseIf KeyAscii > 57 Then
            KeyAscii = 0
        End If

        ' Faccio in modo che se l'utente digita la virgola
        ' mi appaia il punto
        If e.KeyChar = "," Then
            KeyAscii = 46 ' 46 è il codice ascii del punto
        End If

        ' Il punto è si consentito
        ' ma non come primo carattere
        If TextBox_minuti.TextLength = 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If

        ' ovviamente se c'è già un punto
        ' non è consentito digitarne altri
        If (KeyAscii = 46) And
            TextBox_minuti.Text.IndexOf(".") > 0 Then
            KeyAscii = 0
        End If

        ' Reimposto il keychar
        e.KeyChar = Chr(KeyAscii)
    End Sub

    Private Sub ComboBox_causale_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_causale.SelectedIndexChanged
        TextBox_numero.Text = ComboBox_causale.Text
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Lavorazioni_MES.id = Nothing Then
        Else
            Dim CNN As New SqlConnection
            CNN.ConnectionString = Homepage.sap_tirelli
            CNN.Open()

            Dim CMD_SAP As New SqlCommand
            CMD_SAP.Connection = CNN


            CMD_SAP.CommandText = "DELETE MANODOPERA WHERE MANODOPERA.ID='" & Lavorazioni_MES.id & "'"
            CMD_SAP.ExecuteNonQuery()
            CNN.Close()
        End If
        Minuti_progressbar(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
        Lavorazioni_aperte(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), DataGridView_lavorazioni)
        Form_statistiche_manodopera.record_per_dipendente(DataGridView4, TextBox1.Text, Homepage.Centro_di_costo, Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), "")
    End Sub



    Sub inserimento_manodopera_con_check()
        check_manodopera_pregressa()
        If stop_ciclo = 0 Then

            inserisci_consuntivo(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
            Minuti_progressbar(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
            Lavorazioni_aperte(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), DataGridView_lavorazioni)

        End If
    End Sub

    Sub inserimento_manodopera_senza_check()
        inserisci_consuntivo(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
        Minuti_progressbar(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex))
        Lavorazioni_aperte(Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), DataGridView_lavorazioni)
    End Sub

    Sub funzione_inserimento_manodopera()
        giorno_odierno()

        If data_selezione Is Nothing Then
            data_selezione = Today.ToString("yyyy-MM-dd")
        End If

        If oggi = "Monday" Then

            If data_selezione > DateAdd("d", -1, Today) Then
                inserimento_manodopera_con_check()
            Else
                inserimento_manodopera_senza_check()

            End If

        ElseIf oggi = "Tuesday" Then

            If data_selezione > DateAdd("d", -2, Today) Then
                inserimento_manodopera_con_check()
            Else
                inserimento_manodopera_senza_check()
            End If

        ElseIf oggi = "Wednesday" Then

            If data_selezione > DateAdd("d", -3, Today) Then
                inserimento_manodopera_con_check()
            Else
                inserimento_manodopera_senza_check()

            End If

        ElseIf oggi = "Thursday" Then

            If data_selezione > DateAdd("d", -4, Today) Then
                inserimento_manodopera_con_check()
            Else
                inserimento_manodopera_senza_check()

            End If

        ElseIf oggi = "Friday" Then


            If data_selezione > DateAdd("d", -5, Today) Then

                inserimento_manodopera_con_check()
            Else
                inserimento_manodopera_senza_check()
            End If
        End If
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Form_statistiche_manodopera.record_per_dipendente(DataGridView4, TextBox1.Text, Homepage.Centro_di_costo, Lavorazioni_MES.Elenco_dipendenti_MES(ComboBox_dipendente.SelectedIndex), "")
    End Sub

    Private Sub MonthCalendar_data_inizio_DateChanged(sender As Object, e As DateRangeEventArgs) Handles MonthCalendar_data_inizio.DateChanged

    End Sub
End Class