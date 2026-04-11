Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class CQ_Tabelle
    Public Elenco_dipendenti(1000) As String

    Public codicedip As String
    Public filtro_operatore_mu As String
    Public filtro_odp As String
    Public filtro_oa As String
    Public id_controllo As Integer
    Public filtro_esito_controllo As String
    Public filtro_imputazione As String
    Public iniziazione As Integer = 0

    'MAil

    Public objOutlook As Object
    Public objMail As Object

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Close()
    End Sub

    Sub riempi_autocontrollo()

        DataGridView.Rows.Clear()
        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1



        CMD_SAP_2.CommandText = "   SELECT 
top " & TextBox22.Text & "

T0.ID as 'ID', T0.DATA, T1.DOCNUM AS 'ODP', T1.ITEMCODE AS 'Cod', T1.PRODNAME AS 'Descrizione', T1.U_PRG_AZS_COMMESSA AS 'Commessa', T0.DIPENDENTE AS 'ID.P', (T3.LASTNAME+' '+T3.FIRSTNAME) AS 'Dipendente', T0.itemcode as 'Risorsa', t4.resname as 'Macchina', T5.[ResGrpNam] AS 'Gruppo',T0.TIPO_AUTOCONTROLLO AS 'N° Cont', T2.CONTROLLO, T0.OK, T0.NP, T0.D, T0.NC  , T0.derogatore as 'Derogatore', t0.descrizione_deroga as 'Nota deroga', coalesce(t7.docnum,0) as 'EMP_'
    FROM [TIRELLI_40].[DBO].autocontrollo T0 
INNER JOIN [TIRELLISRLDB].[dbo].OWOR T1 ON T0.Docnum = T1.Docnum
    LEFT JOIN [TIRELLI_40].[DBO].autocontrollo_CONFIG T2 ON T0.ID_CONFIG = T2.ID
    LEFT JOIN [TIRELLI_40].[dbo].OHEM T3 ON T3.empid=T0.DIPENDENTE
    left join [TIRELLISRLDB].[dbo].orsc t4 on t4.visrescode =t0.itemcode
    LEFT JOIN [TIRELLISRLDB].[dbo].ORSB T5 ON T5.ResGrpCod = T4.ResGrpCod
left join [TIRELLISRLDB].[dbo].ign1 t6 on t6.BaseRef=T1.DOCNUM and t1.itemcode=t6.itemcode
left join  [TIRELLISRLDB].[dbo].oign t7 on t7.docentry=t6.docentry

where t1.itemcode   Like '%%" & TextBox2.Text & "%%' and t1.docnum   Like '%%" & TextBox1.Text & "%%' and T1.U_PRG_AZS_COMMESSA   Like '%%" & TextBox3.Text & "%%' and T5.[ResGrpNam] Like '%%" & TextBox4.Text & "%%'

    ORDER BY T0.ID desc"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader_2.Read()

            DataGridView.Rows.Add(cmd_SAP_reader_2("id"), cmd_SAP_reader_2("data"), cmd_SAP_reader_2("odp"), cmd_SAP_reader_2("cod"), cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("commessa"), cmd_SAP_reader_2("Dipendente"), cmd_SAP_reader_2("macchina"), cmd_SAP_reader_2("gruppo"), cmd_SAP_reader_2("N° Cont"), cmd_SAP_reader_2("Controllo"), cmd_SAP_reader_2("ok"), cmd_SAP_reader_2("np"), cmd_SAP_reader_2("D"), cmd_SAP_reader_2("NC"), cmd_SAP_reader_2("Derogatore"), cmd_SAP_reader_2("EMP_"))

        Loop



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        DataGridView.ClearSelection()

    End Sub

    Sub riempi_registrazioni_controlli()

        Dim Cnn1 As New SqlConnection

        DataGridView1.Rows.Clear()

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        CMD_SAP_2.CommandText = "     declare @giorni_intervallo as integer
  set @giorni_intervallo = " & TextBox9.Text & "


select t10.id,t10.data,t10.cod,t10.itemname,t10.disegno, case when t10.odp is null then T15.DOCNUM else t10.odp end as 'ODP'
,case when t10.oa is null then t14.docnum else t10.oa end as 'OA'
,t10.emp,t10.emf,t10.ordine,t10.cod_fornitore,t10.Fornitore,t10.imputazione,t10.zona_controllo,t10.Dipendente,t10.[Q.tà Contr.],t10.[Q.tà NC],t10.[Q.tà OK],t10.attività,t10.esito,t10.campo,t10.Descrizione_NC,t10.osservazioni,t10.peso,t10.stato,t10.richiesto,t10.rilevato,t10.concedente
, t10.resname, t10.[Codice operatore], t10.[operatore MU],t10.[Data lavorazione],t10.autocontrollo
,t10.erp
from
(
 SELECT t0.ID, T0.DATA, T0.CODICE AS 'Cod',T1.ITEMNAME , T1.U_DISEGNO as 'Disegno', T0.ODP, T0.OA,
 case when t6.docnum is null then t0.emp else t6.docnum end as 'EMP',
 case when t5.docnum is null then t0.emf else t5.docnum end as 'EMF', 
 CASE WHEN ( T0.OA IS NULL OR T0.OA=0) THEN T0.ODP ELSE T0.OA END AS 'Ordine',
 CASE WHEN (T0.attività = 0 or t0.attività is null) THEN T3.CARDcode else t5.cardcode END AS 'Cod_Fornitore',
 CASE WHEN (T0.attività = 0 or t0.attività is null) THEN T3.CARDNAME else t5.cardname END AS 'Fornitore',
 
 T0.IMPUTAZIONE, T0.ZONA_CONTROLLO,  T2.[U_NAME] as 'Dipendente', T0.PZ_CONTR AS 'Q.tà Contr.', T0.PZ_NC AS 'Q.tà NC', T0.PZ_OK AS 'Q.tà OK', T0.Attività, T0.ESITO_AUTOCONTROLLO AS 'Esito', t0.campo_definizione_NC AS 'Campo', T0.DESCRIZIONE_NC as 'Descrizione_NC', T0.OSSERVAZIONI_NC as 'Osservazioni', T0.PESO_NC AS 'PESO', T0.STATO, t0.richiesto,t0.rilevato, t0.concedente, t0.autocontrollo
  ,t7.itemcode, t8.resname, t7.dipendente as 'Codice operatore', concat(t9.firstname,' ',t9.lastname) as 'operatore MU',t7.data as 'Data lavorazione'
  ,t0.erp
FROM [TIRELLI_40].[DBO].cq_nuovo_controllo T0
INNER JOIN [TIRELLISRLDB].[dbo].OITM T1 ON T1.ITEMCODE = T0.CODICE
LEFT JOIN [TIRELLISRLDB].[dbo].OUSR T2 ON T0.OPERATORE=T2.USERID 
LEFT JOIN [TIRELLISRLDB].[dbo].OPdn T3 ON T3.DOCNUM=T0.emf
left join [TIRELLISRLDB].[dbo].oclg t4 on t4.ClgCode=t0.attività and t0.erp='SAP'
left join [TIRELLISRLDB].[dbo].opdn t5 on t5.DocNum=t4.docnum and t4.doctype=20 and t0.erp='SAP'
left join [TIRELLISRLDB].[dbo].oign t6 on t6.docnum=t4.docnum and t4.doctype=59 and t0.erp='SAP'
 left join [TIRELLI_40].[DBO].autocontrollo t7 on t7.id=t0.autocontrollo
left join [TIRELLISRLDB].[dbo].orsc t8 on t8.visrescode=t7.itemcode
  left join [TIRELLI_40].[dbo].ohem t9 on t9.empid=t7.dipendente

where T0.DATA>=getdate()-@giorni_intervallo 
)
as t10
left join [TIRELLISRLDB].[dbo].opdn t11 on t11.docnum=t10.emf and t10.erp='SAP'
left join [TIRELLISRLDB].[dbo].pdn1 t12 on t12.docentry=t11.docentry AND T10.COD=T12.ITEMCODE and t10.erp='SAP'
LEFT JOIN [TIRELLISRLDB].[dbo].POR1 T13 ON T13.DOCENTRY=T12.BASEENTRY AND T13.LINENUM=T12.BASELINE AND T13.ITEMCODE=T12.ITEMCODE and t10.erp='SAP'
left join [TIRELLISRLDB].[dbo].opor t14 on t14.docentry=t13.docentry and t10.erp='SAP'
left join [TIRELLISRLDB].[dbo].oign t15 on t15.docnum=t10.EMP and t10.erp='SAP'
left join [TIRELLISRLDB].[dbo].IGN1 t16 on t16.docentry=t15.docentry AND T10.COD=T16.ITEMCODE and t10.erp='SAP'
LEFT JOIN [TIRELLISRLDB].[dbo].OWOR T17 ON T17.DOCNUM=T16.BASEREF and t10.erp='SAP'
where t10.data>=getdate()-@giorni_intervallo AND t10.COD   Like '%%" & TextBox6.Text & "%%'  and t10.disegno   Like '%%" & TextBox5.Text & "%%'" & filtro_operatore_mu & " " & filtro_odp & "  " & filtro_oa & " " & filtro_esito_controllo & " " & filtro_imputazione & " 
GROUP BY t10.id,t10.data,t10.cod,t10.itemname,t10.disegno, T10.ODP,T15.DOCNUM,

t10.emp,t10.emf,t10.ordine,t10.cod_fornitore,t10.Fornitore,t10.imputazione,t10.zona_controllo,t10.Dipendente,t10.[Q.tà Contr.],t10.[Q.tà NC],t10.[Q.tà OK],t10.attività,t10.esito,t10.campo,t10.Descrizione_NC,t10.osservazioni,t10.peso,t10.stato,t10.richiesto,t10.rilevato,t10.concedente
, t10.resname, t10.[Codice operatore], t10.[operatore MU],t10.[Data lavorazione],t10.autocontrollo,T10.OA,T14.DOCNUM,t10.erp
order by t10.id desc"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            DataGridView1.Rows.Add(cmd_SAP_reader_2("id"), cmd_SAP_reader_2("data"), cmd_SAP_reader_2("Cod"), cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("disegno"), cmd_SAP_reader_2("Q.tà Contr."), cmd_SAP_reader_2("Q.tà NC"), cmd_SAP_reader_2("imputazione"), cmd_SAP_reader_2("campo"), cmd_SAP_reader_2("descrizione_nc"), cmd_SAP_reader_2("esito"), cmd_SAP_reader_2("peso"), cmd_SAP_reader_2("osservazioni"), cmd_SAP_reader_2("odp"), cmd_SAP_reader_2("oa"), cmd_SAP_reader_2("EMF"), cmd_SAP_reader_2("ERP"), cmd_SAP_reader_2("fornitore"), cmd_SAP_reader_2("autocontrollo"), cmd_SAP_reader_2("resname"), cmd_SAP_reader_2("operatore mu"), cmd_SAP_reader_2("data lavorazione"))

        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        DataGridView1.ClearSelection()

    End Sub

    Private Sub DataGridView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView.CellContentClick

    End Sub

    Private Sub DataGridView_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView.CellClick
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = 0 Then

                CQ_nuovo_controllo.TextBox7.Text = DataGridView.Rows(e.RowIndex).Cells(columnName:="id").Value



                CQ_nuovo_controllo.TextBox1.Text = DataGridView.Rows(e.RowIndex).Cells(columnName:="EMP").Value


                Me.Close()


            End If

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        riempi_autocontrollo()

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        riempi_autocontrollo()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        riempi_autocontrollo()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        riempi_autocontrollo()
    End Sub

    Private Sub tabpage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter
        If iniziazione = 0 Then

        Else
            Inserimento_imputazione()
            Inserimento_dipendenti()
            riempi_registrazioni_controlli()
        End If


    End Sub

    Sub Inserimento_dipendenti()

        ComboBox1.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code where t0.active='Y' and t1.name='Macchine utensili' order by T0.[lastName]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice_dipendenti As Integer
        Indice_dipendenti = 0
        ComboBox1.Items.Add("")
        Indice_dipendenti = Indice_dipendenti + 1
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice_dipendenti) = cmd_SAP_reader("Codice dipendenti")
            ComboBox1.Items.Add(cmd_SAP_reader("Nome"))
            Indice_dipendenti = Indice_dipendenti + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

        codicedip = Elenco_dipendenti(ComboBox1.SelectedIndex)

        If ComboBox1.SelectedIndex = -1 Or ComboBox1.SelectedIndex = 0 Then
            filtro_operatore_mu = ""
        Else
            filtro_operatore_mu = " and t10.[codice operatore]=" & codicedip & ""
        End If

        If iniziazione = 0 Then

        Else
            riempi_registrazioni_controlli()
        End If



    End Sub



    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged

    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.Text = Nothing Then
            filtro_odp = ""
        Else
            filtro_odp = "and t0.odp like Like '%%" & TextBox7.Text & "%%'"
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        riempi_registrazioni_controlli()
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs)
        riempi_registrazioni_controlli()
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs)
        riempi_registrazioni_controlli()
    End Sub

    Private Sub TextBox5_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        riempi_registrazioni_controlli()
    End Sub

    Private Sub TextBox8_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        If TextBox8.Text = Nothing Then
            filtro_oa = ""
        Else
            filtro_oa = "and t0.odp like Like '%%" & TextBox8.Text & "%%'"
        End If
    End Sub




    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If MessageBox.Show($"Sei sicuro di voler eliminare il controllo ?", "Elimina controllo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            elimina_controllo()
            riempi_registrazioni_controlli()
            MsgBox("Controllo eliminato con successo")
        End If



    End Sub

    Sub elimina_controllo()
        Dim CNN5 As New SqlConnection
        CNN5.ConnectionString = Homepage.sap_tirelli
        cnn5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = CNN5
        CMD_SAP_5.CommandText = "delete [TIRELLI_40].[DBO].[CQ_Nuovo_controllo] where id = '" & id_controllo & "' "
        CMD_SAP_5.ExecuteNonQuery()


        cnn5.Close()

    End Sub

    Sub aggiorna_controllo()
        Dim CNN5 As New SqlConnection
        CNN5.ConnectionString = Homepage.sap_tirelli
        cnn5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = CNN5
        CMD_SAP_5.CommandText = "update [TIRELLI_40].[DBO].[CQ_Nuovo_controllo] set 
codice='" & TextBox10.Text & "'
,disegno='" & TextBox11.Text & "'
,erp='" & TextBox23.Text & "'
,pz_contr='" & TextBox16.Text & "'
,PZ_NC='" & TextBox17.Text & "'
,Imputazione='" & TextBox21.Text & "'
,Campo_definizione_NC='" & TextBox12.Text & "'
,Descrizione_NC='" & TextBox13.Text & "'
,esito_autocontrollo='" & TextBox18.Text & "'
,PEso_nc='" & TextBox19.Text & "'
,OSservazioni_NC='" & TextBox14.Text & "'
,ODP='" & TextBox15.Text & "'
,OA='" & TextBox20.Text & "'
  where id = '" & id_controllo & "' "
        CMD_SAP_5.ExecuteNonQuery()


        cnn5.Close()

    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            id_controllo = DataGridView1.Rows(e.RowIndex).Cells(columnName:="DataGridViewButtonColumn1").Value

            TextBox10.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice").Value
            TextBox11.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disegno").Value
            TextBox23.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="erp").Value
            TextBox16.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="pz_contr").Value
            TextBox17.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Pz_NC").Value
            TextBox12.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Campo_definizione_nc").Value
            TextBox13.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Descrizione_nc").Value
            TextBox14.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Osservazioni_NC").Value
            Try
                TextBox15.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ODP_").Value
            Catch ex As Exception
                TextBox15.Text = ""
            End Try

            TextBox18.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Esito_autocontrollo").Value
            TextBox19.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Peso_NC").Value
            Try
                TextBox20.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="OA").Value
            Catch ex As Exception
                TextBox20.Text = ""
            End Try
            TextBox21.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Imputazione").Value


            Label1.Text = id_controllo

            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Disegno) Then



                Magazzino.visualizza_disegno(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disegno").Value)

            End If
            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Codice) Then
                Magazzino.Codice_SAP = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice").Value


                ' Ripristina la finestra se è minimizzata
                If Magazzino.WindowState = FormWindowState.Minimized Then
                    Magazzino.WindowState = FormWindowState.Normal
                End If

                ' Porta la finestra in primo piano
                Magazzino.BringToFront()
                Magazzino.Activate()
                Magazzino.Show()

                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)

            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        aggiorna_controllo()
        riempi_registrazioni_controlli()
        MsgBox("Controllo aggiornato con successo")
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex <= 0 Then
            filtro_esito_controllo = ""
            riempi_registrazioni_controlli()
        Else
            filtro_esito_controllo = " and t10.esito = '" & ComboBox2.Text & "' "
            riempi_registrazioni_controlli()
        End If
    End Sub



    Sub Inserimento_imputazione()

        ComboBox3.Items.Clear()
        ComboBox3.Items.Add("")
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT T0.imputazione
FROM [TIRELLI_40].[DBO].cq_imputazioni t0 group by T0.imputazione"

        cmd_SAP_reader = CMD_SAP.ExecuteReader



        Do While cmd_SAP_reader.Read()

            ComboBox3.Items.Add(cmd_SAP_reader("imputazione"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub

    Private Sub CQ_Tabelle_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Inserimento_imputazione()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If iniziazione > 0 Then


            If ComboBox3.SelectedIndex <= 0 Then
                filtro_imputazione = ""
                riempi_registrazioni_controlli()
            Else
                filtro_imputazione = " and t0.imputazione = '" & ComboBox3.Text & "' "
                riempi_registrazioni_controlli()
            End If
        Else

        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Layout_documenti.Show()
        Layout_documenti.ComboBox1.SelectedIndex = 8
        Layout_documenti.TextBox1.Text = Label1.Text
        Layout_documenti.Button1.PerformClick()

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click


        Layout_documenti.nome_documento_SAP = "Non_conformità"
        Layout_documenti.Informazioni_NC(Label1.Text)

        Acquisti.nome_documento = "Ordine_acquisto"
        Acquisti.documento_sap_testata = "OPOR"
        Acquisti.documento_sap_righe = "POR1"

        Acquisti.TextBox1.Text = Layout_documenti.OA_nc
        Acquisti.GENERA_PDF_DOC(Layout_documenti.OA_nc)

        Layout_documenti.Informazioni_NC(Label1.Text)
        Layout_documenti.nome_documento_SAP = "Non_conformità"
        Layout_documenti.documento_SAP = ""
        Layout_documenti.trova_word_base(Layout_documenti.Lingua, Layout_documenti.documento_SAP, Layout_documenti.garanzia, Layout_documenti.nome_documento_SAP)
        Layout_documenti.Genera_documento_NC(Label1.Text)



        InviaEmailConAllegato()



    End Sub

    Sub InviaEmailConAllegato()


        '
        Dim strSubject As String
        Dim strBody As String
        Dim strAttachmentPath As String
        Dim strImagePath As String
        Dim intImageWidth As Integer
        Dim strFileName As String
        Dim strSenderEmailAddress As String

        If Homepage.azienda = "Tirelli" Then
            strSenderEmailAddress = "qualita.tirelli@tirelli.net"
        ElseIf Homepage.azienda = "4LIFE" Then

            strSenderEmailAddress = "qualita.tirelli@tirelli.net"
        End If


        intImageWidth = 200 ' Imposta la larghezza dell'immagine a 400 pixel


        strImagePath = Homepage.logo_azienda








        strSubject = "Non conformità NR° " & Label1.Text & ""
        strBody = "<font face='Century Gothic' size='3'>Buongiorno, <br> Il reparto qualità segnala una non conformità relativa alla fornitura allegata"

        ' strBody = strBody & "<br> Per informazioni relative al: <br><br>"
        strBody = strBody & "<br> Il pezzo lo potrete visionare e ritirare in sede Tirelli. <br>"

        strBody = strBody & "<br> Il livello d’ urgenza della ripresa è: <br>"
        strBody = strBody & "<br> Con la presente chiediamo al fornitore di comunicare, nel più breve tempo possibile, la nuova data di consegna del pezzo conforme. 
A disposizione
 <br><br>"


        strBody = strBody & "<br><font face='Century Gothic' size='3'><b>" & Layout_documenti.Compilatore & "</b></font>"
        strBody = strBody & "<br><font face='Century Gothic' size='3'>Quality department"
        strBody = strBody & "<br> "
        strBody = strBody & "<br><img src='" & strImagePath & "' width='" & intImageWidth & "'><br>"
        If Homepage.azienda = "Tirelli" Then
            strBody = strBody & "<br><font face='Century Gothic' size='3'>Tirelli SRL"
            strBody = strBody & "<br><font face='Century Gothic' size='3'>Via Vittorio Veronesi, 1 - 46045 Marmirolo (MN) - ITALY"
            strBody = strBody & "<br><font face='Century Gothic' size='3'>Tel. +39 0376 396 820 / 387 048"
            strBody = strBody & "<br><font face='Century Gothic' size='3'>www.tirelli.net"
            strBody = strBody & "<br<font face='Century Gothic' size='3'>>R.I. - C.F. e P.IVA IT01905710206"

        Else Homepage.azienda = "4LIFE"

            strBody = strBody & "<br><font face='Century Gothic' size='3'>4 LIFE</font>"
            strBody = strBody & "<br><font face='Century Gothic' size='3'>Via Progresso, 9 - 46047 Porto Mantovano (MN) - ITALY</font>"
            strBody = strBody & "<br><font face='Century Gothic' size='3'>Tel. + 39 0376 396820</font>"
            strBody = strBody & "<br><font face='Century Gothic' size='3'><a href='http://www.4lifemachinery.com' target='_blank'>www.4lifemachinery.com</a></font>"
            strBody = strBody & "<br><font face='Century Gothic' size='3'>R.I. - C.F. e P.IVA IT02694950201</font>"
            strBody = strBody & "<br><br><font face='Century Gothic' size='2'>Si rende noto che le informazioni contenute nella presente missiva sono strettamente riservate. Per maggiori informazioni, si rinvia alla pagina web www.4lifemachinery.com/privacy-policy ; se non siete i destinatari della presente, Vi preghiamo di darcene immediata notizia per telefono allo 0376 396820 o via e-mail all’indirizzo del mittente e di distruggere il messaggio nonché gli allegati. 
We hereby notify that the information in this message are confidential. For further information, please visit our Web site at www.4lifemachinery.com/en/privacy-policy; If you are not the addressees above mentioned, please contact us immediately by phone, telephone number +39 0376 396820, or by email at the sender’s address and destroy the message as well as its attachments
"
        End If












        'Crea un oggetto Outlook e una nuova email
        objOutlook = CreateObject("Outlook.Application")
        objMail = objOutlook.CreateItem(0)



        'da fare
        ' Acquisti.trova_destinatari()

        trova_destinatari()

        'Imposta i campi della nuova email
        With objMail
            '.To = strEmail
            .Subject = strSubject
            .HTMLBody = strBody
            .Display 'Apre la mail in anteprima
            .CC = "acquisti@tirelli.net; andrea.nolli@tirelli.net"



            .Attachments.Add(Layout_documenti.percorso_documento_nc_pdf)
            .Attachments.Add(Layout_documenti.percorso_documento_acquisto_per_qualità)
            .Attachments.Add(Homepage.percorso_disegni_generico & "PDF\"  & TextBox11.Text & ".PDF")



            .SentOnBehalfOfName = strSenderEmailAddress





            'strFileName = Dir(percorso_cartella & "\*.*")

            'Do While strFileName <> ""

            '    If LCase(Split(strFileName, ".")(UBound(Split(strFileName, ".")))) <> "ini" Then
            '        .Attachments.Add(percorso_cartella & "\" & strFileName)

            '    End If
            '    strFileName = Dir()
            'Loop


        End With

        'Rilascia gli oggetti creati
        objMail = Nothing
        objOutlook = Nothing


    End Sub

    Sub trova_destinatari()
        objMail.To() = ""
        Dim c As String = ""
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "Select t1.e_maill
from [TIRELLISRLDB].[dbo].OPOR t0 inner join [TIRELLISRLDB].DBO.[ocpr] t1 on t0.cardcode=t1.cardcode
where t0.docnum='" & Layout_documenti.OA_nc & "' and t1.u_riceve_mail='Y' and (t1.e_maill<>'' or  t1.e_maill <> null)"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Do While cmd_SAP_reader.Read() = True

            c = c & cmd_SAP_reader("e_maill") & " ;"


        Loop
        objMail.To() = c
        cnn.Close()
    End Sub

    Sub inizializzazione_form()
        iniziazione = 0
        riempi_autocontrollo()
        iniziazione = 1

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        riempi_registrazioni_controlli()
    End Sub
End Class