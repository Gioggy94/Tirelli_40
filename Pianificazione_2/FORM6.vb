Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Windows.Documents



Public Class FORM6
    Public ODP As String
    Public stato_lavorazione As String
    Public Elenco_dipendenti_MES(1000) As String
    Public Elenco_responsabili_Montaggio(1000) As String
    Public Elenco_responsabili_collaudo(1000) As String
    Public Elenco_risorse(1000) As String
    Public riga As Integer
    Public check_dipendente As String
    Public dt = New DataTable()
    Public id_ticket As String

    Public GRUPPI_PREASSEBLAGGIO_TOT As Integer = 0
    Public GRUPPI_MONTAGGIO_TOT As Integer = 0

    Public GRUPPI_PREASSEBLAGGIO_MANCANTI As Integer = 0
    Public GRUPPI_MONTAGGIO_MANCANTI As Integer = 0

    Public GRUPPI_PREASSEBLAGGIO_COMPLETATI As Integer = 0
    Public GRUPPI_MONTAGGIO_COMPLETATI As Integer = 0

    Public filtro_odp As String
    Public filtro_codice As String
    Public filtro_descrizione As String
    Public filtro_fase As String
    Public filtro_mag_minimo As String

    Public condizioni_where As String

    Public codice_treeview As String
    Public codice_pic As String

    Private Sub DataGridView_ODP_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellClick

        If e.RowIndex >= 0 Then
            riga = e.RowIndex
            ODP = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="N_ODP").Value
            '     id_ticket = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="TK").Value
            filtra_news()


            If e.ColumnIndex > 0 And e.RowIndex >= 0 Then
                If id_ticket > 0 Then
                    Button12.Visible = True

                Else
                    Button12.Visible = False
                End If


                If File.Exists(Homepage.percorso_DWF & DataGridView_ODP.Rows(riga).Cells(columnName:="Disegno").Value & ".iam.dwf") Then
                    Button7.BackColor = Color.Lime
                Else
                    Button7.BackColor = Color.Red
                End If



                If File.Exists(Homepage.percorso_disegni_generico & "PDF\" & DataGridView_ODP.Rows(riga).Cells(columnName:="Disegno").Value & ".PDF") Then
                    Button8.BackColor = Color.Lime
                Else
                    Button8.BackColor = Color.Red
                End If
                codice_pic = DataGridView_ODP.Rows(riga).Cells(columnName:="Disegno").Value
                Magazzino.visualizza_picture(codice_pic, PictureBox4)

            End If
            ' Compila_Albero_treeview_async(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Codice").Value, TreeView1, Homepage.sap_tirelli)

        End If
    End Sub

    Private Sub DataGridView_ODP_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_ODP.CellFormatting
        If e.RowIndex < 0 Then Return
        Dim row = DataGridView_ODP.Rows(e.RowIndex)

        ' MAG column — disponibilità materiale
        Dim magVal = row.Cells("MAG").Value
        If Not magVal Is System.DBNull.Value Then
            Dim magNum As Double
            If Double.TryParse(magVal.ToString(), magNum) Then
                If magNum >= 100 Then
                    row.Cells("MAG").Style.BackColor = Color.FromArgb(144, 238, 144)
                    row.Cells("MAG").Style.ForeColor = Color.FromArgb(10, 60, 10)
                ElseIf magNum >= 90 Then
                    row.Cells("MAG").Style.BackColor = Color.FromArgb(255, 220, 60)
                    row.Cells("MAG").Style.ForeColor = Color.FromArgb(80, 50, 0)
                Else
                    row.Cells("MAG").Style.BackColor = Color.FromArgb(210, 80, 80)
                    row.Cells("MAG").Style.ForeColor = Color.White
                End If
            End If
        End If

        ' Completato row — soft green overlay
        Dim completato = row.Cells("Completato").Value
        If Not completato Is System.DBNull.Value AndAlso completato?.ToString() = "C" Then
            row.DefaultCellStyle.BackColor = Color.FromArgb(195, 235, 195)
            row.DefaultCellStyle.ForeColor = Color.FromArgb(20, 70, 20)
        End If
    End Sub


    Sub Cambia_stato_ODP(par_docnum As String, par_utente As String, par_stato As String)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP_7 As New SqlCommand

        CMD_SAP_7.Connection = CNN

        If par_stato = "C" Then
            CMD_SAP_7.CommandText = "
delete t0 
from [Tirelli_40].dbo.[Ordini_produzione_dati] t0 where t0.docnum ='" & par_docnum & "'

INSERT INTO [dbo].[Ordini_produzione_dati]
           ([Docnum]
           ,[empid]
           ,[Docdate]
           ,[Stato])
     VALUES
           ('" & par_docnum & "'
           ,'" & par_utente & "'
           ,getdate()
           ,'" & par_stato & "')
"
        Else
            CMD_SAP_7.CommandText = "
delete t0 
from [Tirelli_40].dbo.[Ordini_produzione_dati] t0 where t0.docnum ='" & par_docnum & "'
"
        End If



        CMD_SAP_7.ExecuteNonQuery()

        CNN.Close()
    End Sub


    Sub elenco_ODP_commessa(par_commessa As String, par_datagridview As DataGridView)

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand
        Dim cmd_SAP_reader_1 As SqlDataReader
        CMD_SAP_1.Connection = Cnn1

        If Homepage.ERP_provenienza = "SAP" Then


            If par_commessa = "MAGAZZINO" Or par_commessa = "SCORTA" Then

                condizioni_where = " and (t0.status='P' or t0.status='R')  "
            Else
                condizioni_where = " and (t0.status='P' or t0.status='R' or t0.closedate>= getdate()-360 ) "
            End If


            CMD_SAP_1.CommandText = " DECLARE @Codice_matricola AS VARCHAR (20)
set @Codice_matricola='" & par_commessa & "'

SELECT t40.[N ODP] ,t40.Progressivo_commessa, t40.[Stato ODP], t40.Codice , t40.Descrizione , t40.Disegno , t40.quantita , t40.stato ,T40.[Codice fase], t40.fase , t40.N ,t40.Trasferiti ,  T40.PREM, T40.MONT ,T40.[ASS EL], T40.NEWS,t40.tipo,case when MAX(t40.Id_Ticket) is null then '' else MAX(t40.Id_Ticket) end AS 'ID_TICKET', case when t42.descrizione is null then '' else t42.descrizione end as 'Reparto_ticket'
FROM
(
SELECT  t30.[N ODP] ,t30.Progressivo_commessa, t30.[Stato ODP], t30.Codice , t30.Descrizione , t30.Disegno , t30.quantita , t30.stato ,T30.[Codice fase], t30.fase , t30.N ,t30.Trasferiti ,  T30.PREM, T30.MONT ,T30.[ASS EL], T30.NEWS,t30.tipo,MAX(B.Id_Ticket) AS 'ID_TICKET'
FROM
(

SELECT t20.[N ODP] ,t20.Progressivo_commessa, t20.[Stato ODP], t20.Codice , t20.Descrizione , t20.Disegno , t20.quantita , t20.stato ,T20.[Codice fase], t20.fase , t20.N ,t20.Trasferiti ,  case when T20.PREM is null then 0 else t20.prem end as 'PREM' , case when T20.MONT is null then 0 else t20.mont end as 'MONT' ,T20.[ASS EL], SUM(CASE WHEN a.DOCENTRY is null THEN 0 ELSE 1 END)  AS 'NEWS',t20.tipo 
FROM
(
Select T10.DOCENTRY,t10.[N ODP] as 'N ODP',t10.Progressivo_commessa, t10.[Stato ODP], t10.Codice as 'Codice', t10.Descrizione as 'Descrizione', t10.Disegno as 'Disegno', t10.quantita as 'Quantita', t10.stato as 'Stato', t10.fase as 'Fase', T10.[Codice fase], t10.N as 'N',t10.Trasferiti as 'Trasferiti',  sum(CASE WHEN T10.[Codice fase] ='P01501' THEN t10.quantita* case when T11.[Code]='R00568' OR T11.CODE='R00525' then t11.quantity else 0 end END)   as 'PREM' , sum(CASE WHEN T10.[Codice fase] ='P02001' THEN t10.quantita* case when T11.[Code]='R00568' OR T11.CODE='R00525' then t11.quantity else 0 end END)   as 'MONT' ,sum(case when t11.code='R00530' then t11.quantity else 0 end) as 'ASS EL', t10.tipo
from
(
select t5.docentry, t5.[N ODP], t5.Progressivo_commessa,t5.[Stato ODP],t5.codice,t5.descrizione,t5.disegno,t5.Quantita,t5.stato,t5.[Codice fase], t5.fase,t5.tipo
,sum(CASE WHEN T6.ITEMTYPE='4' then 1 else 0 end ) as 'N',
 sum(case when t6.U_prg_wip_qtadatrasf=0 and T6.ITEMTYPE='4'  then 1 else 0 end) as 'trasferiti'
from
(
SELECT T0.DOCENTRY,T0.[DocNum] as 'N ODP',
coalesce(t0.u_progressivo_commessa,0) as 'Progressivo_commessa',
t0.status as 'Stato ODP', T0.[ItemCode] as 'Codice', T1.itemname as 'Descrizione', coalesce(t1.u_disegno,'') as 'Disegno', T0.[PlannedQty] as 'Quantita',T0.U_stato as 'Stato', T0.[U_Fase] as 'Codice fase', T2.[Name] as 'Fase'

, SUBSTRING(T0.u_PRODUZIONE,1,1) AS 'Tipo'

FROM OWOR T0 inner join OITM T1 on t0.itemcode=t1.itemcode
 left JOIN [dbo].[@FASE]  T2 ON T0.[U_Fase] = T2.[Code] 
WHERE T0.[U_PRG_AZS_Commessa]   =@Codice_matricola " & condizioni_where & " and (T0.[U_PRODUZIONE]='ASSEMBL' or T0.[U_PRODUZIONE]='EST') 
group by
T0.DOCENTRY, T0.[DocNum] ,
 t0.U_Progressivo_commessa ,
T0.[ItemCode], T1.[U_Disegno] , T2.[Name], t1.itemname, T0.[PlannedQty], T0.[U_Fase], T0.U_stato,t0.status, t0.U_produzione
)
as t5 left join wor1 t6 on t6.docentry=t5.docentry
group by t5.docentry, t5.Progressivo_commessa,t5.[Stato ODP],t5.codice,t5.descrizione,t5.disegno,t5.Quantita,t5.stato,t5.[Codice fase], t5.fase,t5.tipo, t5.[N ODP]

)
as t10
left join itt1 t11 on t11.father=t10.codice
group by T10.DOCENTRY, t10.[N ODP],t10.Progressivo_commessa, t10.[Stato ODP], t10.Codice, t10.Descrizione, t10.Disegno, t10.quantita, t10.fase, t10.N,t10.Trasferiti, t10.[Codice Fase],t10.stato ,t10.tipo
)
AS T20

LEFT JOIN (SELECT T2.DOCENTRY
from owtr t0 left join wtr1 t1 on t0.docentry=t1.docentry
left join owor t2 on T1.U_PRG_AZS_OPDOCENTRY=T2.DOCENTRY
where t0.DocDate>=getdate()-2 and t1.WHSCODE='WIP' AND T2.U_PRG_AZS_COMMESSA=@Codice_matricola) A on A.DOCENTRY=T20.DOCENTRY 
 
GROUP BY t20.[N ODP] ,t20.Progressivo_commessa, t20.[Stato ODP], t20.Codice , t20.Descrizione , t20.Disegno , t20.quantita , t20.stato , t20.fase , t20.N ,t20.Trasferiti, T20.PREM,T20.MONT, T20.[ASS EL],T20.[Codice Fase],t20.tipo


)
AS T30
LEFT JOIN ( SELECT T2.DOCNUM,T3.Id_Ticket
 
 FROM [TIRELLI_40].[DBO].COLL_RIFERIMENTI T1 INNER JOIN OWOR T2 ON T2.DOCNUM=t1.Codice_SAP
 left join [TIRELLI_40].[DBO].coll_tickets t3 on t3.Id_Ticket=t1.Rif_Ticket and t3.aperto=1
 
 WHERE T1.Tipo_Codice='Ordine di produzione' AND T2.U_PRG_AZS_Commessa=@Codice_matricola)B ON B.DOCNUM=T30.[N ODP]
 group by t30.[N ODP] ,t30.Progressivo_commessa, t30.[Stato ODP], t30.Codice , t30.Descrizione , t30.Disegno , t30.quantita , t30.stato ,T30.[Codice fase], t30.fase , t30.N ,t30.Trasferiti ,  T30.PREM, T30.MONT ,T30.[ASS EL], T30.NEWS,t30.tipo
 )
AS T40 left join [TIRELLI_40].[DBO].coll_tickets t41 on t41.Id_Ticket=t40.Id_Ticket
left join [TIRELLI_40].[DBO].COLL_Reparti t42 on t42.Id_Reparto=t41.Destinatario

where 0=0 " & filtro_odp & " " & filtro_codice & " " & filtro_descrizione & " " & filtro_fase & " 
group by t40.[N ODP],t40.Progressivo_commessa , t40.[Stato ODP], t40.Codice , t40.Descrizione , t40.Disegno , t40.quantita , t40.stato ,T40.[Codice fase], t40.fase , t40.N ,t40.Trasferiti ,  T40.PREM, T40.MONT ,T40.[ASS EL], T40.NEWS,t40.tipo,t42.descrizione

order by t40.Progressivo_commessa, T40.TIPO, t40.[Codice Fase],  t40.[N ODP]"
        Else
            CMD_SAP_1.CommandText = "  select 
t10.numodp as 'N ODP'
,t10.posizione as 'Progressivo_commessa'
,TRIM(t10.codart) as 'Codice'
,t10.dscodart_odp as 'Descrizione'
,trim(t10.disegno) as 'Disegno'
,t10.qta_pia as 'Quantita'
,t10.qta_res as 'Quantita_res'
,t10.pianificato as 'Stato'
,t10.ciclo as 'Fase'
,t10.nrrig_imp as 'N'
,t10.NRRIG_IMPAV as 'Trasferiti'
,0 as 'Prem'
,0 as 'Mont'
,0 as 'Ass EL'
,'' as 'Tipo'
,0 as 'News'
,0 as 'ID_ticket'
,'' as 'Reparto_ticket'
,coalesce(t11.stato,'') as  'Completato'


FROM OPENQUERY(AS400, '
    SELECT *
    FROM TIR90VIS.JGALODP t0
	where matricola=''" & par_commessa & "'' " & filtro_odp & filtro_codice & filtro_descrizione & "
 and mag_ver<>''TCP''
') T10
LEFT JOIN [Tirelli_40].[dbo].[Ordini_produzione_dati] t11 
    ON t10.[NumODP] COLLATE DATABASE_DEFAULT = t11.docnum COLLATE DATABASE_DEFAULT"

        End If
        cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
        Do While cmd_SAP_reader_1.Read()
            Dim img As Image = Nothing

            Dim codiceDisegno As String = cmd_SAP_reader_1("Disegno").ToString()
            Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & codiceDisegno & ".PNG"

            If File.Exists(percorso) Then
                Using fs As New FileStream(percorso, FileMode.Open, FileAccess.Read)
                    Using tmp As Image = Image.FromStream(fs)
                        img = New Bitmap(tmp) ' evita lock sul file
                    End Using
                End Using
            End If

            Dim idx As Integer = par_datagridview.Rows.Add(
                cmd_SAP_reader_1("N ODP"),
                img,
                cmd_SAP_reader_1("Progressivo_commessa"),
                cmd_SAP_reader_1("Codice"),
                cmd_SAP_reader_1("Descrizione"),
                cmd_SAP_reader_1("Disegno"),
                "",
                Math.Round(cmd_SAP_reader_1("Quantita")),
                cmd_SAP_reader_1("Stato"),
                cmd_SAP_reader_1("Fase"),
                cmd_SAP_reader_1("trasferiti") / cmd_SAP_reader_1("N") * 100,
                cmd_SAP_reader_1("Completato")
            )

            ' Riduci altezza SOLO se non c'è immagine
            If img Is Nothing Then
                par_datagridview.Rows(idx).Height = 40
            Else
                '   par_datagridview.Rows(idx).Height = 60 ' o quello che vuoi
            End If
        Loop
        Cnn1.Close()
        DataGridView_ODP.ClearSelection()
    End Sub







    Sub Check_Lavorazioni_aperte_dipendente(par_codice_dipendente As String)

        Dim testo As String = ""
        Lavorazioni_MES.DataGridView_lavorazioni.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.id as 'ID', t0.tipo_documento as 'Documento', t0.docnum as 'ODP', t1.itemcode as 'Itemcode', t2.itemname as 'Descrizione', case when t2.u_disegno is null then '' else t2.u_disegno end as 'Disegno', t1.plannedqty as 'quantita', T1.[U_PRG_AZS_Commessa] as 'Commessa', t4.lastname +' '+t4.firstname as 'Dipendente', t3.resname as 'Risorsa', t0.data as 'Data', t0.start as 'Start', t0.stop, t0.consuntivo
from manodopera t0 left join owor t1 on t0.docnum = t1.docnum
left join oitm t2 on t2.itemcode=t1.itemcode
inner join orsc t3 on t3.visrescode=t0.risorsa
left join [TIRELLI_40].[dbo].OHEM T4 ON T4.[empID]=T0.DIPENDENTE
where (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='') and t0.dipendente='" & par_codice_dipendente & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            testo = "A"
            Lavorazioni_MES.DataGridView_lavorazioni.Rows.Add(cmd_SAP_reader_2("ID"), cmd_SAP_reader_2("Documento"), cmd_SAP_reader_2("ODP"), cmd_SAP_reader_2("Itemcode"), cmd_SAP_reader_2("Descrizione"), cmd_SAP_reader_2("Disegno"), Math.Round(cmd_SAP_reader_2("Quantita"), 2), cmd_SAP_reader_2("Commessa"), cmd_SAP_reader_2("Dipendente"), cmd_SAP_reader_2("Risorsa"), cmd_SAP_reader_2("Data"), cmd_SAP_reader_2("Start"))
        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        If testo = "A" Then
            MsgBox("risultano le seguenti lavorazioni aperte su questo dipendente, devono essere concluse prima di poterne aprire delle altre")

            'Lavorazioni_MES.Button_start.Hide()
            Lavorazioni_MES.Button_stop.Show()
            check_dipendente = "STOP"

        End If


        testo = ""

    End Sub

    Sub CHIUDI_lavorazione()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.ID as 'ID', t0.docnum as 'ODP', t3.itemcode as 'Itemcode',T3.PRODNAME as 'Descrizione', T3.[U_PRG_AZS_Commessa] as 'Commessa', case when t4.u_disegno is null then '' else t4.u_disegno end as 'Disegno', t3.plannedqty as 'Quantita', T1.[firstName]+' '+T1.[lastName] as 'Dipendente', t2.resname as 'Risorsa', t0.data as 'Data', t0.start as 'Start'
FROM MANODOPERA t0 inner join [TIRELLI_40].[dbo].ohem t1 on t1.[empID]=dipendente
inner join orsc t2 on t2.visrescode=t0.risorsa
LEFT JOIN OWOR T3 ON T3.DOCNUM=t0.docnum
left join oitm t4 on t4.itemcode=t3.itemcode
where t0.docnum=" & ODP & " and (t0.stop is null or t0.stop ='') and (t0.consuntivo is null or t0.consuntivo='')"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If Not cmd_SAP_reader_2.Read() = True Then
            stato_lavorazione = ""
            ' Cambia_stato_ODP()

        End If

        Cnn1.Close()

    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Close()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Lavorazioni_MES.Button_start.Show()
        Lavorazioni_MES.Button_stop.Hide()
        Lavorazioni_MES.Show()
        Lavorazioni_MES.DataGridView_lavorazioni.Rows.Clear()


        Lavorazioni_MES.ComboBox_risorse.Text = ""
        Lavorazioni_MES.ComboBox_dipendente.Text = ""
        Lavorazioni_MES.formatta_form_8(DataGridView_ODP.Rows(riga).Cells(columnName:="N_ODP").Value)
        Lavorazioni_MES.inserimento_dipendenti_MES(Lavorazioni_MES.ComboBox_dipendente, Lavorazioni_MES.Elenco_dipendenti_MES)
        Lavorazioni_MES.Inserimento_risorse_MES(Lavorazioni_MES.ComboBox_risorse)
        Lavorazioni_MES.Button_start.Show()
        Lavorazioni_MES.GroupBox3.Show()
        Lavorazioni_MES.GroupBox1.Show()
        Lavorazioni_MES.GroupBox2.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Homepage.Form_precedente = 6
        Lavorazioni_MES.Button_start.Hide()
        Lavorazioni_MES.Button_stop.Show()
        Lavorazioni_MES.Show()

        Lavorazioni_MES.ComboBox_risorse.Text = ""
        Lavorazioni_MES.ComboBox_dipendente.Text = ""

        Lavorazioni_MES.Lavorazioni_aperte(Lavorazioni_MES.DataGridView_lavorazioni, DataGridView_ODP.Rows(riga).Cells(columnName:="N_ODP").Value, 0)
        Lavorazioni_MES.inserimento_dipendenti_MES(Lavorazioni_MES.ComboBox_dipendente, Lavorazioni_MES.Elenco_dipendenti_MES)
        Lavorazioni_MES.Inserimento_risorse_MES(Lavorazioni_MES.ComboBox_risorse)
        Lavorazioni_MES.GroupBox3.Show()
        Lavorazioni_MES.GroupBox1.Hide()
        Lavorazioni_MES.GroupBox2.Hide()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Cambia_stato_ODP(ODP, Homepage.ID_SALVATO, "C")
        elenco_ODP_commessa(Pianificazione.commessa, DataGridView_ODP)
        'If DataGridView_ODP.Rows(riga).Cells(columnName:="STATO").Value Is System.DBNull.Value Then
        '    stato_lavorazione = "Completato"
        '    Cambia_stato_ODP(ODP, Homepage.ID_SALVATO, "C")
        '    DataGridView_ODP.Rows(riga).Cells(columnName:="STATO").Value = "Completato"
        '    ' completamento_gruppi_preassemblaggio_assemblaggio()
        'ElseIf DataGridView_ODP.Rows(riga).Cells(columnName:="STATO").Value = "In_esecuzione" Then
        '    MsgBox("Risulta una lavorazione aperta, chiudere tutte le lavorazioni per poter completare l'operazione")

        'Else
        '    stato_lavorazione = "Completato"
        '    Cambia_stato_ODP()
        '    DataGridView_ODP.Rows(riga).Cells(columnName:="STATO").Value = "Completato"
        '    ' completamento_gruppi_preassemblaggio_assemblaggio()
        'End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Cambia_stato_ODP(ODP, Homepage.ID_SALVATO, "X")
        elenco_ODP_commessa(Pianificazione.commessa, DataGridView_ODP)
        'Try
        '    If DataGridView_ODP.Rows(riga).Cells(columnName:="STATO").Value = "Completato" Then
        '        stato_lavorazione = ""
        '        Cambia_stato_ODP()
        '        DataGridView_ODP.Rows(riga).Cells(columnName:="STATO").Value = ""
        '        '  completamento_gruppi_preassemblaggio_assemblaggio()
        '    End If
        'Catch ex As Exception

        'End Try

    End Sub


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)
        filtra_data()
    End Sub


    Sub filtra_data()
        Dim i = 0
        Do While i < DataGridView_ODP.RowCount
            Dim parola As String
            parola = UCase(DataGridView_ODP.Rows(i).Cells(2).Value)

            If parola.Contains(UCase(TextBox1.Text)) Then
                DataGridView_ODP.Rows(i).Visible = True

            Else
                DataGridView_ODP.Rows(i).Visible = False

            End If
            i = i + 1
        Loop
    End Sub

    Private Sub Button_commessa_Click(sender As Object, e As EventArgs) Handles Button_commessa.Click
        Mostra.Hide()
        Homepage.mostra_dashboard()
        Mostra.Owner = Me
        Mostra.Show()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs)
        'Me.WindowState = FormWindowState.Normal

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub




    Sub filtra_news()
        Dim i = 0

        Dim parola4 As String


        Do While i < DataGridView1.RowCount
            Try


                parola4 = UCase(DataGridView1.Rows(i).Cells(4).Value)


                If parola4.Contains(ODP) Then
                    DataGridView1.Rows(i).Visible = True

                Else
                    DataGridView1.Rows(i).Visible = False

                End If


            Catch ex As Exception
                DataGridView1.Rows(i).Visible = False
            End Try
            i = i + 1
        Loop
    End Sub




    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click


        If ODP = Nothing Then
            MsgBox("Scegliere un ordine di produzione")
        Else
            Dim new_form_odp_form = New ODP_Form
            new_form_odp_form.docnum_odp = ODP
            new_form_odp_form.Show()
            new_form_odp_form.inizializza_form(ODP)

        End If

    End Sub

    Private Sub Button8_Click_1(sender As Object, e As EventArgs) Handles Button8.Click


        Magazzino.visualizza_disegno(DataGridView_ODP.Rows(riga).Cells(columnName:="disegno").Value)


    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        If File.Exists(Homepage.percorso_DWF & DataGridView_ODP.Rows(riga).Cells(columnName:="Disegno").Value & ".iam.dwf") Then
            Console.WriteLine(Homepage.percorso_DWF & DataGridView_ODP.Rows(riga).Cells(columnName:="Disegno").Value & ".iam.dwf")
            Process.Start(Homepage.percorso_DWF & DataGridView_ODP.Rows(riga).Cells(columnName:="Disegno").Value & ".iam.dwf")
        Else
            MsgBox("3D non presente")
        End If
    End Sub



    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        elenco_ODP_commessa(Pianificazione.commessa, DataGridView_ODP)
        news_materiale()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Form_nuovo_ticket.Show()
        Form_nuovo_ticket.Inserimento_dipendenti()

        Form_nuovo_ticket.Reparto = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto
        Form_nuovo_ticket.Administrator = 0
        Form_nuovo_ticket.Startup()
        Form_nuovo_ticket.Txt_Commessa.Text = Pianificazione.commessa
        Form_nuovo_ticket.ComboBox2.Text = Homepage.business
    End Sub



    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Form_Visualizza_Ticket.Txt_Id.Text = id_ticket
        Form_Visualizza_Ticket.Startup()
        Form_Visualizza_Ticket.Show()


    End Sub



    '    Sub completamento_gruppi_preassemblaggio_assemblaggio()


    '        GRUPPI_PREASSEBLAGGIO_TOT = 0
    '        GRUPPI_MONTAGGIO_TOT = 0

    '        GRUPPI_PREASSEBLAGGIO_MANCANTI = 0
    '        GRUPPI_MONTAGGIO_MANCANTI = 0

    '        GRUPPI_PREASSEBLAGGIO_COMPLETATI = 0
    '        GRUPPI_MONTAGGIO_COMPLETATI = 0

    '        Label10.Text = 0
    '        Label3.Text = 0

    '        Lavorazioni_MES.DataGridView_lavorazioni.Rows.Clear()

    '        Dim Cnn1 As New SqlConnection
    '        Cnn1.ConnectionString = Homepage.sap_tirelli
    '        Cnn1.Open()

    '        Dim CMD_SAP_2 As New SqlCommand
    '        Dim cmd_SAP_reader_2 As SqlDataReader


    '        CMD_SAP_2.Connection = cnn1
    '        CMD_SAP_2.CommandText = "SELECT case when t0.u_fase is null then '' else t0.u_fase end as 'u_fase', case when t0.u_stato is null then '' else t0.u_stato end as 'u_stato', count(T0.[DocNum]) as 'Numero_gruppi'
    'FROM OWOR T0 
    'WHERE T0.STATUS <>'C' AND t0.u_prg_azs_commessa='" & pianificazione.commessa & "' and (t0.u_produzione='ASSEMBL' or t0.u_produzione='EST')
    'group by t0.u_fase, t0.u_stato"

    '        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

    '        Do While cmd_SAP_reader_2.Read()


    '            If cmd_SAP_reader_2("u_fase") = "P01501" Then

    '                If cmd_SAP_reader_2("u_stato") = "Completato" Then
    '                    ProgressBar_PREASS.Maximum = cmd_SAP_reader_2("Numero_gruppi") + 1
    '                    GRUPPI_PREASSEBLAGGIO_COMPLETATI = cmd_SAP_reader_2("Numero_gruppi")
    '                    ProgressBar_PREASS.Value = cmd_SAP_reader_2("Numero_gruppi")
    '                    Label10.Text = cmd_SAP_reader_2("Numero_gruppi")

    '                End If

    '                GRUPPI_PREASSEBLAGGIO_TOT = GRUPPI_PREASSEBLAGGIO_TOT + cmd_SAP_reader_2("Numero_gruppi")
    '            End If

    '            If cmd_SAP_reader_2("u_fase") = "P02001" Then

    '                If cmd_SAP_reader_2("u_stato") = "Completato" Then
    '                    ProgressBar_montaggio.Maximum = cmd_SAP_reader_2("Numero_gruppi")
    '                    GRUPPI_MONTAGGIO_COMPLETATI = cmd_SAP_reader_2("Numero_gruppi")
    '                    ProgressBar_montaggio.Value = cmd_SAP_reader_2("Numero_gruppi")
    '                    Label3.Text = cmd_SAP_reader_2("Numero_gruppi")
    '                End If
    '                GRUPPI_MONTAGGIO_TOT = GRUPPI_MONTAGGIO_TOT + cmd_SAP_reader_2("Numero_gruppi")
    '            End If

    '        Loop
    '        Label7.Text = GRUPPI_PREASSEBLAGGIO_TOT
    '        ProgressBar_PREASS.Maximum = GRUPPI_PREASSEBLAGGIO_TOT

    '        Try
    '            GRUPPI_PREASSEBLAGGIO_MANCANTI = GRUPPI_PREASSEBLAGGIO_TOT - Label10.Text
    '            Label8.Text = GRUPPI_PREASSEBLAGGIO_TOT - Label10.Text
    '            Label2.Text = Math.Round(Label10.Text / GRUPPI_PREASSEBLAGGIO_TOT * 100) & "%"
    '        Catch ex As Exception

    '        End Try


    '        Label5.Text = GRUPPI_MONTAGGIO_TOT
    '        Try
    '            GRUPPI_MONTAGGIO_MANCANTI = GRUPPI_MONTAGGIO_TOT - Label3.Text
    '            Label4.Text = GRUPPI_MONTAGGIO_TOT - Label3.Text
    '            Label_tempo_mont_XC.Text = Math.Round(Label3.Text / GRUPPI_MONTAGGIO_TOT * 100) & "%"
    '        Catch ex As Exception

    '        End Try
    '        ProgressBar_montaggio.Maximum = GRUPPI_MONTAGGIO_TOT

    '        cmd_SAP_reader_2.Close()
    '        cnn1.Close()

    '        Commesse_MES.date_inizio_fine_commesse()


    '    End Sub





    Private Sub TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = Nothing Then
            filtro_odp = ""
        Else

            If Homepage.ERP_provenienza = "SAP" Then
                filtro_odp = " and t40.[N ODP] Like '%%" & TextBox1.Text & "%%'  "
            Else
                filtro_odp = " and t0.numodp   LIKE ''%" & TextBox1.Text & "%'' "
            End If

        End If
        elenco_ODP_commessa(Pianificazione.commessa, DataGridView_ODP)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = Nothing Then
            filtro_codice = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_codice = " and t40.Codice Like '%%" & TextBox2.Text & "%%'  "
            Else
                filtro_codice = "  and upper(t0.codart) LIKE ''%%" & TextBox2.Text.ToUpper & "%%'' "
            End If

        End If
        elenco_ODP_commessa(Pianificazione.commessa, DataGridView_ODP)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = Nothing Then
            filtro_fase = ""
        Else
            filtro_fase = " and t40.[fase] Like '%%" & TextBox4.Text & "%%'  "
        End If
        elenco_ODP_commessa(Pianificazione.commessa, DataGridView_ODP)
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = Nothing Then
            filtro_descrizione = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_descrizione = " and t40.descrizione Like '%%" & TextBox3.Text & "%%'  "
            Else
                filtro_descrizione = " AND UPPER(t0.dscodart_odp) LIKE ''%" & TextBox3.Text.ToUpper() & "%'' "
            End If

        End If
        elenco_ODP_commessa(Pianificazione.commessa, DataGridView_ODP)
    End Sub

    Private Sub DataGridView_ODP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellContentClick

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        DataGridView_ODP.Columns("PDF").Visible = True
        For Each row As DataGridViewRow In DataGridView_ODP.Rows
            If row.Cells("Codice").Value.ToString().StartsWith("C", StringComparison.OrdinalIgnoreCase) Then
                ' Fai qualcosa se la prima lettera è C
            Else
                Dim percorso As String = Homepage.percorso_disegni_generico & "PDF\"  & row.Cells("disegno").Value & ".PDF"
                If File.Exists(percorso) Then

                    row.Cells("PDF").Value = "SI"

                Else

                    row.Cells("PDF").Value = "NO"
                End If
            End If

        Next
        MsgBox("FINE")
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Form_lotto_di_prelievo.estrai_datagridview_in_excel(DataGridView_ODP)
    End Sub



    Public Sub Compila_Albero_treeview_async(par_codice As String, par_treeview As TreeView, connectionString As String)
        ' Controlla se la stringa di connessione è valida
        If String.IsNullOrWhiteSpace(connectionString) Then
            MessageBox.Show("⚠️ Errore: la stringa di connessione non è impostata.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' Mostra un nodo di caricamento
        par_treeview.Invoke(Sub()
                                par_treeview.Nodes.Clear()
                                par_treeview.Nodes.Add(New TreeNode("🔄 Caricamento..."))
                            End Sub)

        Task.Run(Sub()
                     Dim rootNode As TreeNode = Nothing

                     Try
                         Using Cnn_Tree As New SqlConnection(connectionString)
                             Cnn_Tree.Open()

                             Using Cmd_Tree As New SqlCommand("SELECT T0.[father], T0.[code], T1.[itemname] as 'Nome_padre'
                                                           FROM itt1 T0 
                                                           INNER JOIN oitm T1 ON T1.itemcode = T0.father
                                                           WHERE T0.[father] = @par_codice
                                                           ORDER BY T0.VisOrder", Cnn_Tree)

                                 Cmd_Tree.Parameters.AddWithValue("@par_codice", par_codice)

                                 Using Reader_Tree As SqlDataReader = Cmd_Tree.ExecuteReader()
                                     If Reader_Tree.Read() Then
                                         rootNode = New TreeNode(Reader_Tree("father") & "-" & Reader_Tree("nome_padre"))
                                         Trova_Figli_Background(rootNode, Reader_Tree("father"), par_treeview, connectionString)
                                     End If
                                 End Using
                             End Using
                         End Using
                     Catch ex As Exception
                         Debug.WriteLine("❌ ERRORE: " & ex.Message)
                     End Try

                     ' Aggiorna la TreeView nel thread principale
                     par_treeview.Invoke(Sub()
                                             par_treeview.Nodes.Clear()
                                             If rootNode IsNot Nothing Then
                                                 par_treeview.Nodes.Add(rootNode)
                                                 par_treeview.ExpandAll()
                                                 par_treeview.TopNode = par_treeview.Nodes(0)
                                             Else
                                                 par_treeview.Nodes.Add(New TreeNode("❌ Nessun dato trovato"))
                                             End If
                                         End Sub)
                 End Sub)
    End Sub


    Private Sub Trova_Figli_Background(parentNode As TreeNode, par_codice As String, par_treeview As TreeView, connectionString As String)
        Dim childNodes As New List(Of TreeNode)

        Using Cnn_Tree As New SqlConnection(connectionString)
            Cnn_Tree.Open()

            Using Cmd_Tree As New SqlCommand("SELECT T0.[father], T0.[code], SUBSTRING(T0.code,1,1) as 'Prima_lettera', 
                                                 T1.[itemname] as 'Nome_padre', COALESCE(T1.u_disegno, '') as 'Disegno', 
                                                 T2.[itemname] as 'Nome_figlio', T0.[Quantity]
                                          FROM itt1 T0 
                                          INNER JOIN oitm T1 ON T1.itemcode = T0.father
                                          INNER JOIN oitm T2 ON T2.itemcode = T0.code
                                          INNER JOIN oitt T3 ON T3.code = T0.code
                                          WHERE T0.[father] = @par_codice 
                                          AND SUBSTRING(T0.code,1,1) = '0'
                                          ORDER BY T0.VisOrder", Cnn_Tree)

                Cmd_Tree.Parameters.AddWithValue("@par_codice", par_codice)

                Using Reader_Tree As SqlDataReader = Cmd_Tree.ExecuteReader()
                    While Reader_Tree.Read()
                        Dim newNode As New TreeNode(Reader_Tree("code") & " - " & Reader_Tree("nome_figlio") & " Q : = " & Reader_Tree("Quantity"))

                        ' Imposta l'immagine in base alla prima lettera
                        Select Case Reader_Tree("prima_lettera").ToString()
                            Case "C"
                                newNode.ImageIndex = 2
                            Case "D"
                                newNode.ImageIndex = 1
                            Case "0"
                                newNode.ImageIndex = 0
                            Case "R"
                                newNode.ImageIndex = 3
                        End Select

                        ' **Aggiungo il nodo subito, così so che esiste**
                        childNodes.Add(newNode)
                    End While
                End Using
            End Using
        End Using

        ' **Aggiorniamo i nodi nel thread principale**
        par_treeview.Invoke(Sub()
                                parentNode.Nodes.Clear()
                                If childNodes.Count > 0 Then
                                    parentNode.Nodes.AddRange(childNodes.ToArray())
                                End If

                                ' **Adesso richiamo la funzione per i figli di ogni nodo**
                                For Each nodo As TreeNode In childNodes
                                    Trova_Figli_Background(nodo, nodo.Text.Split("-")(0).Trim(), par_treeview, connectionString)
                                Next
                            End Sub)
    End Sub





    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs)
        codice_treeview = e.Node.Text.Split("-"c)(0).Trim()

        If File.Exists(Homepage.percorso_DWF & Magazzino.OttieniDettagliAnagrafica(codice_treeview).Disegno & ".iam.dwf") Then
            Apri_DXF.BackColor = Color.Lime
        Else
            Apri_DXF.BackColor = Color.Red
        End If

        If File.Exists(Homepage.percorso_disegni_generico & "PDF\" & Magazzino.OttieniDettagliAnagrafica(codice_treeview).Disegno & ".PDF") Then
            Apri_PDF.BackColor = Color.Lime
        Else
            Apri_DXF.BackColor = Color.Red
        End If


    End Sub

    Private Sub Apri_PDF_Click(sender As Object, e As EventArgs) Handles Apri_PDF.Click
        Magazzino.visualizza_disegno(Magazzino.OttieniDettagliAnagrafica(codice_treeview).Disegno)
    End Sub

    Private Sub Apri_DXF_Click(sender As Object, e As EventArgs) Handles Apri_DXF.Click
        If File.Exists(Homepage.percorso_DWF & Magazzino.OttieniDettagliAnagrafica(codice_treeview).Disegno & ".iam.dwf") Then
            Console.WriteLine(Homepage.percorso_DWF & Magazzino.OttieniDettagliAnagrafica(codice_treeview).Disegno & ".iam.dwf")
            Process.Start(Homepage.percorso_DWF & Magazzino.OttieniDettagliAnagrafica(codice_treeview).Disegno & ".iam.dwf")
        Else
            MsgBox("3D non presente")
        End If
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs)

    End Sub

    Private Sub Apri_Distinta_Click(sender As Object, e As EventArgs) Handles Apri_Distinta.Click
        Dim new_form_distinta_form As New Distinta_base_form

        ' Imposta il valore della TextBox
        new_form_distinta_form.TextBox1.Text = codice_treeview

        ' Calcola 1 cm in pixel (conversione da cm a pixel, 96 dpi)
        Dim oneCmInPixels As Integer = CInt(1 / 2.54 * 96) ' 1 cm ≈ 37,8 pixel a 96 dpi

        ' Imposta la posizione della nuova form rispetto alla posizione dell'attuale form
        new_form_distinta_form.StartPosition = FormStartPosition.Manual
        new_form_distinta_form.Location = New Point(Me.Location.X + oneCmInPixels, Me.Location.Y + oneCmInPixels)

        ' Mostra la form
        new_form_distinta_form.Show()
    End Sub

    Private Sub FORM6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ApplicaStile()
    End Sub

    Sub ApplicaStile()
        Dim navy As Color = Color.FromArgb(22, 45, 84)
        Dim navyDark As Color = Color.FromArgb(10, 26, 55)
        Dim navyMid As Color = Color.FromArgb(16, 34, 70)
        Dim accent As Color = Color.FromArgb(70, 130, 210)
        Dim bgApp As Color = Color.FromArgb(238, 242, 247)
        Dim fontBase As New Font("Segoe UI", 9, FontStyle.Regular)
        Dim fontBold As New Font("Segoe UI", 9, FontStyle.Bold)
        Dim fontSmall As New Font("Segoe UI", 8, FontStyle.Regular)

        Me.BackColor = bgApp

        ' ── Left sidebar (Panel1) ──
        Panel1.BackColor = navyDark
        Button_commessa.BackColor = navy
        Button_commessa.ForeColor = Color.White
        Button_commessa.FlatStyle = FlatStyle.Flat
        Button_commessa.FlatAppearance.BorderSize = 0
        Button_commessa.Font = New Font("Segoe UI", 15, FontStyle.Bold)
        Panel_sep1.BackColor = accent
        Panel_sep2.BackColor = accent
        Panel16.BackColor = navyDark
        TLP_InfoCommessa.BackColor = Color.Transparent
        For Each lbl As Label In TLP_InfoCommessa.Controls.OfType(Of Label)()
            If lbl.Name.StartsWith("Label_t_") Then
                lbl.ForeColor = Color.FromArgb(140, 175, 220)
                lbl.Font = fontSmall
            Else
                lbl.ForeColor = Color.White
                lbl.Font = fontBold
            End If
        Next

        ' Premontaggio/Montaggio area
        Panel17.BackColor = navyDark
        Panel2.BackColor = navyMid
        Panel18.BackColor = navyMid
        Label1.ForeColor = Color.White
        Label1.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        Label21.ForeColor = Color.White
        Label21.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        Label2.ForeColor = Color.FromArgb(144, 238, 144)
        Label_tempo_mont_XC.ForeColor = Color.FromArgb(144, 238, 144)
        For Each gb As GroupBox In {GroupBox18, GroupBox17, GroupBox16, GroupBox20, GroupBox19,
                                     GroupBox11, GroupBox12, GroupBox13, GroupBox14, GroupBox15}
            gb.BackColor = navyDark
            gb.ForeColor = Color.FromArgb(140, 175, 220)
            gb.Font = fontSmall
        Next
        For Each lbl As Label In {Label10, Label8, Label7, Label12, Label11,
                                   Label3, Label4, Label5, Label6, Label9}
            lbl.ForeColor = Color.White
            lbl.Font = fontBold
        Next

        ' News materiali
        Panel24.BackColor = navyDark
        GroupBox10.BackColor = navyDark
        GroupBox10.ForeColor = Color.White
        GroupBox10.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        DataGridView1.BackgroundColor = navyDark
        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = navy
        DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        DataGridView1.ColumnHeadersDefaultCellStyle.Font = fontBold
        DataGridView1.DefaultCellStyle.BackColor = bgApp
        DataGridView1.DefaultCellStyle.ForeColor = Color.FromArgb(30, 40, 60)
        DataGridView1.DefaultCellStyle.Font = New Font("Segoe UI", 8)
        DataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.White
        DataGridView1.GridColor = Color.FromArgb(200, 210, 230)
        DataGridView1.EnableHeadersVisualStyles = False

        ' ── Right action panel (Panel3) ──
        Panel3.BackColor = navyDark
        Button1.BackColor = navy
        Button1.ForeColor = Color.White
        Button1.FlatStyle = FlatStyle.Flat
        Button1.FlatAppearance.BorderColor = Color.FromArgb(40, 70, 130)
        Button3.BackColor = Color.FromArgb(150, 35, 35)
        Button3.ForeColor = Color.White
        Button3.FlatStyle = FlatStyle.Flat
        Button3.FlatAppearance.BorderSize = 0
        Button14.BackColor = Color.FromArgb(30, 60, 110)
        Button14.ForeColor = Color.White
        Button14.FlatStyle = FlatStyle.Flat
        Button14.Font = fontSmall
        Button10.BackColor = Color.FromArgb(30, 60, 110)
        Button10.ForeColor = Color.White
        Button10.FlatStyle = FlatStyle.Flat
        Button10.Font = fontSmall
        GroupBox113.BackColor = navyMid
        GroupBox113.ForeColor = Color.White
        GroupBox113.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        Button13.FlatStyle = FlatStyle.Flat
        Button13.FlatAppearance.BorderColor = accent
        Button9.FlatStyle = FlatStyle.Flat
        Button9.FlatAppearance.BorderColor = accent
        Button12.BackColor = Color.FromArgb(30, 100, 60)
        Button12.ForeColor = Color.White
        Button12.Font = fontBold
        Button11.BackColor = Color.FromArgb(22, 75, 155)
        Button11.ForeColor = Color.White
        Button11.FlatStyle = FlatStyle.Flat
        Button11.FlatAppearance.BorderSize = 0
        Button11.Font = New Font("Segoe UI", 12, FontStyle.Bold)
        Button8.BackColor = Color.FromArgb(14, 50, 100)
        Button8.ForeColor = Color.White
        Button8.FlatStyle = FlatStyle.Flat
        Button7.BackColor = Color.FromArgb(14, 50, 100)
        Button7.ForeColor = Color.White
        Button7.FlatStyle = FlatStyle.Flat
        Button5.BackColor = Color.FromArgb(30, 110, 60)
        Button5.ForeColor = Color.White
        Button5.FlatStyle = FlatStyle.Flat
        Button5.FlatAppearance.BorderSize = 0
        Button6.FlatStyle = FlatStyle.Flat
        Button6.FlatAppearance.BorderSize = 0

        ' ── Filter bar (Panel4) ──
        Panel4.BackColor = navyDark
        For Each gb As GroupBox In TableLayoutPanel14.Controls.OfType(Of GroupBox)()
            gb.BackColor = Color.Transparent
            gb.ForeColor = Color.FromArgb(140, 175, 220)
            gb.Font = fontSmall
            For Each tb As TextBox In gb.Controls.OfType(Of TextBox)()
                tb.BackColor = navyMid
                tb.ForeColor = Color.White
                tb.BorderStyle = BorderStyle.FixedSingle
            Next
        Next

        ' ── DataGridView_ODP ──
        DataGridView_ODP.BackgroundColor = bgApp
        DataGridView_ODP.ColumnHeadersDefaultCellStyle.BackColor = navy
        DataGridView_ODP.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        DataGridView_ODP.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        DataGridView_ODP.DefaultCellStyle.BackColor = Color.White
        DataGridView_ODP.DefaultCellStyle.ForeColor = Color.FromArgb(30, 40, 60)
        DataGridView_ODP.DefaultCellStyle.Font = New Font("Segoe UI", 9)
        DataGridView_ODP.AlternatingRowsDefaultCellStyle.BackColor = bgApp
        DataGridView_ODP.GridColor = Color.FromArgb(200, 210, 230)
        DataGridView_ODP.EnableHeadersVisualStyles = False
    End Sub

    Sub inizializza_form(par_commessa As String)
        Button_commessa.Text = par_commessa
        Label_descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(par_commessa).Descrizione_commessa
        Label_ordine_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(par_commessa).ordine_cliente_commessa
        Label_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(par_commessa).Cliente_commessa
        Label_cliente_finale.Text = Commesse_MES.SCHEDA_COMMESSA(par_commessa).Cliente_finale_commessa
        Label_consegna.Text = Commesse_MES.SCHEDA_COMMESSA(par_commessa).Consegna_commessa
        'completamento_gruppi_preassemblaggio_assemblaggio()
        elenco_ODP_commessa(Pianificazione.commessa, DataGridView_ODP)
        news_materiale()
        Dim senso_rotazione As String = Scheda_tecnica.trova_SENSO_ORIENTAMENTO_commessa(par_commessa)
        GroupBox113.Text = "Senso : " & senso_rotazione


        If senso_rotazione = "SX-DX(CW)" Then
            PictureBox3.Image = Image.FromFile(Homepage.percorso_server & "00-Tirelli 4.0\Immagini\Img Scheda tecnica\Oraria.png")
            ' Imposta la modalità di ridimensionamento

        ElseIf senso_rotazione = "DX-SX(CCW)" Then
            PictureBox3.Image = Image.FromFile(Homepage.percorso_server & "00-Tirelli 4.0\Immagini\Img Scheda tecnica\Antioraria.png")

        End If
        PictureBox3.SizeMode = PictureBoxSizeMode.Zoom



    End Sub

    Sub news_materiale()
        If Homepage.ERP_provenienza = "SAP" Then


            DataGridView1.Rows.Clear()
            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_1 As New SqlCommand
            Dim cmd_SAP_reader_1 As SqlDataReader
            CMD_SAP_1.Connection = Cnn1


            CMD_SAP_1.CommandText = " SELECT T2.[ItemCode], T2.[ItemName], T1.[Quantity],t3.docnum, T3.PRODNAME, T3.[U_PRG_AZS_Commessa], T0.DOCDATE,T0.DOCTIME FROM OWTR T0  INNER JOIN WTR1 T1 ON T0.[DocEntry] = T1.[DocEntry] INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode]
LEFT JOIN OWOR T3 ON T3.DOCENTRY= T1.[U_PRG_AZS_OpDocEntry] WHERE T0.[DocDate] >GETDATE()-2 AND T1.WHSCODE='WIP' and T3.[U_PRG_AZS_Commessa] ='" & Pianificazione.commessa & "' 
order by t0.docdate DESC, t0.doctime DESC"

            cmd_SAP_reader_1 = CMD_SAP_1.ExecuteReader
            Do While cmd_SAP_reader_1.Read()

                DataGridView1.Rows.Add(cmd_SAP_reader_1("ItemCode"), cmd_SAP_reader_1("ItemName"), cmd_SAP_reader_1("Quantity"), cmd_SAP_reader_1("PRODNAME"), cmd_SAP_reader_1("docnum"), cmd_SAP_reader_1("docdate"), cmd_SAP_reader_1("doctime"))
            Loop
            Cnn1.Close()
            DataGridView1.ClearSelection()
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Process.Start("https://jpm.tirelli.net/jpm-share/?path=dclOlnCfg")
    End Sub

    Private Sub PictureBox4_Click(sender As Object, e As EventArgs) Handles PictureBox4.Click

        Magazzino.apri_picture(codice_pic)
    End Sub
End Class