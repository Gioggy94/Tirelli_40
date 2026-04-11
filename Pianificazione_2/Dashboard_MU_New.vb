Imports System.Collections
'Imports Microsoft.Office.Interop.Excel
Imports System.ComponentModel.Design
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Net.Mail
Imports System.Windows.Forms.DataVisualization.Charting
Imports AxFOXITREADERLib
Imports Microsoft.Office.Interop
Imports Tirelli.ODP_Form
Imports Tirelli.Presence

Public Class Dashboard_MU_New

    Public Elenco_macchinari(1000) As String
    Public Elenco_tipo_macchina(1000) As String
    Public Elenco_dipendenti(1000) As String
    Public Elenco_imprevisti(100) As String
    Public Codicedip As Integer
    Public risorsa As String
    Public tipo_macchina As String
    Private filtro_odp As String
    Private filtro_codice As String
    Private filtro_descrizione As String
    Private filtro_materiale As String
    Private filtro_spessore As String
    Public docnum As String
    Public docnum_odp As Integer
    Public docentry As Integer
    Public codice As String
    Public Descart As String
    Public disegno As String
    Public quantodp As Integer
    Public N_lavorazione As String
    Public tipo_lav As String
    Public num_bolla As String
    Public Property Nesting_riga As Integer
    Public time As Integer
    Public id_manodopera As Integer
    Public percentuale As Integer
    Private lavorazione As String
    Private id_padre As Integer
    Private risorsa_old As String
    Private numero_lavorazione As Integer
    Private id_imprevisto As Integer
    Private codice_imprevisto As Integer
    Public tipo_autocontrollo As Integer

    Public autocontrollo_attrezzaggio_necessario As String
    Public autocontrollo_lavorazione_necessario As String
    Private ordine_completabile As String
    Public tempo_cambio_pezzo As Integer
    Public tempo_CICLO_pezzo As Integer

    Public mu As Integer = 1
    Public Num_vis_riga As Integer
    Private blocco_tab As Integer = 0
    Public code_fase As String


    Sub inizializzazione_dashboard_mu()


        'If Homepage.Centro_di_costo = "BRB01" Then
        '    blocco_tab = 0
        '    Inserimento_dipendenti()
        '    Inserimento_risorse()
        '    Inserimento_causali_imprevisti()
        '    compila_datagridvied_lista_odp_bint()
        '    TabControl1.SelectedIndex = 3
        '    blocco_tab = 1
        'Else
        Inserimento_dipendenti()
            Inserimento_risorse()
            Inserimento_causali_imprevisti()
        '  End If


        'TabPage6.Hide()



    End Sub

    Sub compila_datagridvied_lista_odp_bint()
        lISTA_odp_BINT(DataGridView5, TextBox15.Text, TextBox16.Text, TextBox14.Text)
    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting
        If blocco_tab = 1 Then
            e.Cancel = True ' Impedisce il cambio di scheda
        End If


    End Sub

    Sub Inserimento_dipendenti()
        Combodipendenti.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code where t0.active='Y' and t1.name='Macchine utensili' order by T0.[lastName]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            Combodipendenti.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Sub Inserimento_risorse()
        Combomacchinari.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = Cnn
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_docentry.CommandText = "SELECT T0.[VisResCode] AS 'Risorsa'
,T0.[ResGrpCod] as 'Tipo macchina', T0.[ResName] as 'Nome' 
FROM ORSC T0 WHERE T0.[ResType] ='M' and t0.validfor='Y' 
ORDER BY t0.[resname]"
        Else
            CMD_SAP_docentry.CommandText = "select risorsa
,cent_lav as 'Tipo macchina'
,concat(trim(centrolavoro),' ',trim(macchina) ) as 'Nome' 
from
[AS400].[S786FAD1].[TIR90VIS].[JGALRIS]
order by cent_lav, concat(trim(centrolavoro),' ',trim(macchina) )"
        End If


        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_docentry_reader.Read()
            Elenco_macchinari(Indice) = cmd_SAP_docentry_reader("Risorsa")
            Elenco_tipo_macchina(Indice) = cmd_SAP_docentry_reader("Tipo macchina")
            Combomacchinari.Items.Add(cmd_SAP_docentry_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_docentry_reader.Close()
        Cnn.Close()


    End Sub 'Inserisco le risorse nella combo box

    Sub Inserimento_causali_imprevisti()
        ComboBox1.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = Cnn
        CMD_SAP_docentry.CommandText = "Select id,descrizione
from [TIRELLI_40].[DBO].causali_imprevisti_mu"

        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_docentry_reader.Read()
            Elenco_imprevisti(Indice) = cmd_SAP_docentry_reader("id")

            ComboBox1.Items.Add(cmd_SAP_docentry_reader("Descrizione"))
            Indice = Indice + 1
        Loop
        cmd_SAP_docentry_reader.Close()
        Cnn.Close()


    End Sub 'Inserisco le risorse nella combo box

    Private Sub Combodipendenti_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combodipendenti.SelectedIndexChanged
        If Combodipendenti.Text <> "" Then
            Codicedip = Elenco_dipendenti(Combodipendenti.SelectedIndex)
            ' Dashboard_pianificazione.Dipendente = Elenco_dipendenti(Combodipendenti.SelectedIndex)

            'dipendente = Combodipendenti.Text

            kpi_mensile_NC()
            kpi_mensile_FASI()
        Else
            MsgBox("Inserire codice dipendente valido")
        End If


        visibilità()

    End Sub

    Private Sub Combomacchinari_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combomacchinari.SelectedIndexChanged
        If Combomacchinari.Text <> "" Then
            risorsa = Elenco_macchinari(Combomacchinari.SelectedIndex)
            tipo_macchina = Elenco_tipo_macchina(Combomacchinari.SelectedIndex)



            Button3.Text = ""
            riempi_datagridview(TextBox8.Text, TextBox7.Text, TextBox5.Text, TextBox4.Text, TextBox17.Text)

            'lISTA_odp_per_tipo_macchina(DataGridView_ODP_TIPO_MACCHINA)



        Else
            MsgBox("Inserire macchinario")
        End If

        visibilità()
    End Sub

    Sub riempi_datagridview(par_odp As String, par_codice As String, par_descrizione As String, par_materiale As String, par_magazzino_destinazione As String)
        'If tipo_macchina = 7 Then
        '    Dim par_tipo_appoggio As String = "ODP_TAGLIO"
        '    DataGridView_ODP_TIPO_MACCHINA.Rows.Clear()
        '    ODP_Tree.PULISCI_APPOGGIO(Homepage.ID_SALVATO, par_tipo_appoggio)
        '    lISTA_odp_taglio(DataGridView_ODP_TIPO_MACCHINA, par_tipo_appoggio)
        'Else
        lISTA_odp_per_tipo_macchina(DataGridView_ODP_TIPO_MACCHINA, par_odp, par_codice, par_descrizione, par_materiale, par_magazzino_destinazione)
        ' End If

        Presenza_attrezzaggio_lavorazione()
    End Sub

    Sub visibilità()
        If Combodipendenti.SelectedIndex >= 0 And Combomacchinari.SelectedIndex >= 0 Then
            ' TableLayoutPanel5.Visible = True
            'TableLayoutPanel8.Visible = True
            TabControl1.Visible = True

        Else
            '   TableLayoutPanel5.Visible = False
            ' TableLayoutPanel8.Visible = False
            TabControl1.Visible = False

        End If
    End Sub


    Sub kpi_mensile_NC()
        If Homepage.ERP_provenienza = "SAP" Then


            Button16.Text = 0
            Button16.BackColor = Color.Red
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()

            Dim CMD_SAP_docentry As New SqlCommand
            Dim cmd_SAP_docentry_reader As SqlDataReader

            CMD_SAP_docentry.Connection = Cnn
            CMD_SAP_docentry.CommandText = "SELECT count(t0.id) as 'N°_NC'
from [TIRELLI_40].[DBO].CQ_Nuovo_controllo t0
left join ousr t1 on t0.operatore=t1.userid
left join opor t2 on t2.docnum=t0.oa
left join owor t3 on t3.docnum=t0.odp
left join oitm t4 on t4.itemcode=t3.u_prg_azs_commessa
inner join [TIRELLI_40].[DBO].autocontrollo t5 on t5.id=t0.autocontrollo
left join orsc t6 on t6.visrescode=t5.itemcode
left join [TIRELLI_40].[dbo].ohem t7 on t7.empid=t5.dipendente
where t0.data>=getdate()-30 and t5.dipendente='" & Codicedip & "' and t0.esito_autocontrollo<>'Conforme'"

            cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader



            If cmd_SAP_docentry_reader.Read() Then
                If Not cmd_SAP_docentry_reader("N°_NC") Is System.DBNull.Value Then
                    Button16.BackColor = Color.Red
                    Button16.Text = cmd_SAP_docentry_reader("N°_NC")
                Else

                    Button16.Text = 0
                End If
            End If
            If Button16.Text = 0 Then
                Button16.BackColor = Color.Lime


            End If
            cmd_SAP_docentry_reader.Close()
            Cnn.Close()

        End If
    End Sub 'Inserisco le risorse nella combo box

    Sub kpi_mensile_FASI()
        If Homepage.ERP_provenienza = "SAP" Then


            Button17.Text = 0
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()

            Dim CMD_SAP_docentry As New SqlCommand
            Dim cmd_SAP_docentry_reader As SqlDataReader

            CMD_SAP_docentry.Connection = Cnn
            CMD_SAP_docentry.CommandText = "SELECT T0.DIPENDENTE,T1.LASTNAME +' ' + T1.FIRSTNAME AS 'Dipendente', COUNT(CONCAT(t0.docnum,T0.RISORSA)) as 'ODP'
FROM MANODOPERA t0 LEFT JOIN [TIRELLI_40].[dbo].OHEM T1 ON T1.empid=T0.DIPENDENTE
left join orsc t2 on t2.visrescode=t0.risorsa
left join owor t3 on t3.docnum=t0.docnum and t0.tipo_documento='ODP'
LEFT JOIN OITM t4 ON T4.ITEMCODE=T3.ITEMCODE
left join oitm t5 on t5.itemcode=T3.[U_PRG_AZS_Commessa]
left join [TIRELLI_40].[dbo].oudp t6 on t1.dept=t6.code

where t2.restype='M' AND T0.DIPENDENTE='" & Codicedip & "' AND T0.DATA>=GETDATE()-30
group by T0.DIPENDENTE, T1.LASTNAME +' ' + T1.FIRSTNAME "

            cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader



            If cmd_SAP_docentry_reader.Read() Then
                If Not cmd_SAP_docentry_reader("ODP") Is System.DBNull.Value Then

                    Button17.Text = cmd_SAP_docentry_reader("odp")
                Else

                    Button17.Text = 0
                End If
            Else
                Button17.Text = 0
            End If

            cmd_SAP_docentry_reader.Close()
            Cnn.Close()

        End If

    End Sub

    Sub lISTA_odp_per_tipo_macchina(par_datagridview As DataGridView, par_odp As String, par_codice As String, par_descrizione As String, par_materiale As String, par_magazzino_destinazione As String)

        Dim filtro_odp As String
        If par_odp = "" Then
            filtro_odp = ""
        Else
            filtro_odp = " and odp Like ''%%" & par_odp & "%%''"
        End If
        Dim filtro_codice As String
        If par_codice = "" Then
            filtro_codice = ""
        Else
            filtro_codice = " and codart_odp   Like ''%%" & par_codice & "%%''"
        End If
        Dim filtro_descrizione As String
        If par_descrizione = "" Then

        Else
            filtro_descrizione = " and dscodart_odp   Like ''%%" & par_descrizione & "%%''"
        End If
        Dim filtro_materiale As String
        If par_materiale = "" Then
            filtro_materiale = ""
        Else
            filtro_materiale = " and tipomate   Like ''%%" & par_materiale & "%%''"
        End If

        Dim filtro_magazzino_dest As String
        If par_magazzino_destinazione = "" Then
            filtro_magazzino_dest = ""
        Else
            filtro_magazzino_dest = " and odp Like ''%%" & par_odp & "%%''"
        End If

        par_datagridview.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = Cnn
        If Homepage.ERP_provenienza = "SAP" Then



            If CheckBox1.Checked = True Then


                CMD_SAP_docentry.CommandText = "Select t30.ODP,T30.DOCENTRY, T30.DOC, T30.[N PEZ] , T30.COD, T30.Nome, t30.DIS, t30.RIS, t30.MAC, T30.ATT, T30.LAV, T30.Priorità, t30.u_lavorazione, case when t30.Righe=t30.Trasferiti then 'OK' when t30.[Trasferiti parziali]>0 then 'PARZ' else 'NO' end as 'materiale'
from
(
SELECT t20.ODP,T20.DOCENTRY, T20.DOC, T20.[N PEZ] , T20.COD, T20.Nome, t20.DIS, t20.RIS, t20.MAC, T20.ATT, T20.LAV, T20.Priorità, t20.u_lavorazione, sum(case when t21.itemtype=4 then 1 else 0 end) as 'Righe', sum(case when t21.itemtype=4 and T21.[U_PRG_WIP_QtaDaTrasf]=0  then 1 else 0 end) as 'Trasferiti' ,sum(case when t21.itemtype=4 and T21.[U_PRG_WIP_QtaSpedita]>0 and T21.[U_PRG_WIP_QtaDaTrasf]>0  then 1 else 0 end) as 'Trasferiti parziali' 
FROM
(
Select t13.docnum as 'ODP', t13.docentry, 'ODP' AS 'DOC', T13.[PlannedQty] as 'N PEZ' , t13.itemcode as 'COD', t14.itemname as 'Nome', t14.u_disegno as 'DIS', t11.itemcode as 'RIS', t12.itemname as 'MAC', T11.[AdditQty] as 'ATT', T11.[PlannedQty] as 'LAV', T13.[U_Priorita_MES] AS 'Priorità', t13.u_lavorazione
from
(
SELECT t0.docentry, min(T0.[VisOrder]) as 'Visorder' , t0.u_stato_lavorazione
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry]
inner join ORSC T2 on t2.visrescode=t0.itemcode 
WHERE T1.[Status] ='R' and  T2.[ResType] ='M' and  T0.[U_Stato_lavorazione] = 'O' and substring(t1.u_produzione,1,3)='INT'
group by 
t0.docentry, t0.u_stato_lavorazione
)
as t10
inner join wor1 t11 on t11.docentry=t10.docentry and t10.visorder=t11.visorder
inner join oitm t12 on t12.itemcode=t11.itemcode
inner join owor t13 on t13.docentry=t11.docentry
inner join oitm t14 on t14.itemcode=t13.itemcode
inner join orsc t15 on t15.visrescode=t11.itemcode
where t15.resgrpcod=" & tipo_macchina & "
)
AS T20 left join wor1 t21 on t21.docentry=t20.docentry

group by
t20.ODP,T20.DOCENTRY, T20.DOC, T20.[N PEZ] , T20.COD, T20.Nome, t20.DIS, t20.RIS, t20.MAC, T20.ATT, T20.LAV, T20.Priorità, t20.u_lavorazione
)
as t30
where 0=0 " & filtro_materiale & filtro_spessore & filtro_descrizione & filtro_codice & filtro_odp & "
order by T30.[Priorità]"



            Else
                CMD_SAP_docentry.CommandText = "Select t30.ODP,T30.DOCENTRY, T30.DOC, T30.[N PEZ] , T30.COD, T30.Nome, t30.DIS, t30.RIS, t30.MAC, T30.ATT, T30.LAV, T30.Priorità, t30.u_lavorazione, case when t30.Righe=t30.Trasferiti then 'OK' when t30.[Trasferiti parziali]>0 then 'PARZ' else 'NO' end as 'materiale', T30.U_PRG_TIR_MATERIALE, T30.BHEIGHT1,T30.BWIDTH1, T30.BLENGTH1
from
(
SELECT t20.ODP,T20.DOCENTRY, T20.DOC, T20.[N PEZ] , T20.COD, T20.Nome, t20.DIS, t20.RIS, t20.MAC, T20.ATT, T20.LAV, T20.Priorità, t20.u_lavorazione, sum(case when t21.itemtype=4 then 1 else 0 end) as 'Righe', sum(case when t21.itemtype=4 and T21.[U_PRG_WIP_QtaDaTrasf]=0  then 1 else 0 end) as 'Trasferiti' ,sum(case when t21.itemtype=4 and T21.[U_PRG_WIP_QtaSpedita]>0 and T21.[U_PRG_WIP_QtaDaTrasf]>0  then 1 else 0 end) as 'Trasferiti parziali' 
, T20.U_PRG_TIR_MATERIALE, T20.BHEIGHT1,T20.BWIDTH1, T20.BLENGTH1
FROM
(
Select t13.docnum as 'ODP',T13.DOCENTRY, 'ODP' AS 'DOC', T13.[PlannedQty] as 'N PEZ' , t13.itemcode as 'COD', t14.itemname as 'Nome', t14.u_disegno as 'DIS', t11.itemcode as 'RIS', t12.itemname as 'MAC', T11.[AdditQty] as 'ATT', T11.[PlannedQty] as 'LAV', T13.[U_Priorita_MES] AS 'Priorità', t13.u_lavorazione, CASE WHEN T14.U_PRG_TIR_MATERIALE IS NULL THEN '' ELSE T14.U_PRG_TIR_MATERIALE END AS 'U_PRG_TIR_MATERIALE' , CASE WHEN T14.BHEIGHT1 IS NULL THEN 0 ELSE T14.BHEIGHT1 END AS 'BHEIGHT1',T14.BWIDTH1, T14.BLENGTH1
from
(
SELECT t0.docentry, min(T0.[VisOrder]) as 'Visorder' , t0.u_stato_lavorazione
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry]
inner join ORSC T2 on t2.visrescode=t0.itemcode WHERE T1.[Status] ='R' and  T2.[ResType] ='M' and  T0.[U_Stato_lavorazione] = 'O' and substring(t1.u_produzione,1,3)='INT'
group by 
t0.docentry, t0.u_stato_lavorazione
)
as t10
inner join wor1 t11 on t11.docentry=t10.docentry and t10.visorder=t11.visorder
inner join oitm t12 on t12.itemcode=t11.itemcode
inner join owor t13 on t13.docentry=t11.docentry
inner join oitm t14 on t14.itemcode=t13.itemcode
inner join orsc t15 on t15.visrescode=t11.itemcode
where t11.itemcode='" & risorsa & "'
)
AS T20 left join wor1 t21 on t21.docentry=t20.docentry

group by
t20.ODP,T20.DOCENTRY, T20.DOC, T20.[N PEZ] , T20.COD, T20.Nome, t20.DIS, t20.RIS, t20.MAC, T20.ATT, T20.LAV, T20.Priorità, t20.u_lavorazione, T20.U_PRG_TIR_MATERIALE, T20.BHEIGHT1,T20.BWIDTH1, T20.BLENGTH1
)
as t30
where 0=0 " & filtro_materiale & filtro_spessore & filtro_descrizione & filtro_codice & filtro_odp & "
order by T30.Priorità"

            End If
        Else
            CMD_SAP_docentry.CommandText = "
select 
t10.odp
,t10.odp as 'docentry'
,'ODP' as 'DOC'
,qta_odp as 'N PEZ'
,trim(codart_odp) as 'COD'
,dscodart_odp as 'Nome'
,trim(disegno) as 'DIS'
,cod_risorsa as 'Ris'
,risorsa as 'MAC'
,attman as 'ATT'
,lavman as 'LAV'
,attmac as 'Attrezzaggio_macchina'
,lavmac as 'Lavorazione_macchina'
,priorita as 'Priorità'
,999 as 'U_lavorazione'
,'MANCA' as 'Materiale'
,tipomate as 'U_prg_tir_materiale'
,altezza as 'Bheight1'
,larghezza as 'Bwidth1'
,lunghezza as 'Blenght1'
,fase_av

FROM OPENQUERY(AS400, '
select *
from
S786FAD1.TIR90VIS.JGALodpmu
where 
dest=''" & par_magazzino_destinazione & "'' and
cod_risorsa=''" & risorsa & "''
" & filtro_odp & "
" & filtro_codice & "
" & filtro_descrizione & "
" & filtro_materiale & "
and fase_av=''*'' and stato=''R''

order by data_iniz
'
) as t10
order by t10.priorita"

        End If
        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader
        Do While cmd_SAP_docentry_reader.Read()
            Dim img As Image = Nothing
            Dim codiceDisegno As String = If(IsDBNull(cmd_SAP_docentry_reader("DIS")), "", cmd_SAP_docentry_reader("DIS").ToString())
            Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & codiceDisegno & ".PNG"

            If codiceDisegno <> "" AndAlso File.Exists(percorso) Then
                Using fs As New FileStream(percorso, FileMode.Open, FileAccess.Read)
                    Using tmp As Image = Image.FromStream(fs)
                        img = New Bitmap(tmp) ' evita lock sul file
                    End Using
                End Using
            End If
            par_datagridview.Rows.Add(cmd_SAP_docentry_reader("ODP"), Math.Round(cmd_SAP_docentry_reader("N PEZ")), cmd_SAP_docentry_reader("COD"), cmd_SAP_docentry_reader("Nome"), cmd_SAP_docentry_reader("DIS"), img, cmd_SAP_docentry_reader("RIS"), cmd_SAP_docentry_reader("MAC"), Math.Round(cmd_SAP_docentry_reader("ATT")), Math.Round(cmd_SAP_docentry_reader("LAV")), cmd_SAP_docentry_reader("u_lavorazione"), cmd_SAP_docentry_reader("Priorità"), cmd_SAP_docentry_reader("Materiale"), cmd_SAP_docentry_reader("U_PRG_TIR_MATERIALE"), cmd_SAP_docentry_reader("BHEIGHT1"))

        Loop
        cmd_SAP_docentry_reader.Close()
        Cnn.Close()


    End Sub

    Sub lISTA_odp_taglio(par_datagridview As DataGridView, par_tipo_appoggio As String)



        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = Cnn




        CMD_SAP_docentry.CommandText = "
Select t30.ODP,T30.DOCENTRY, T30.DOC, T30.[N PEZ] , T30.COD, T30.Nome, t30.DIS, t30.RIS, t30.MAC, T30.ATT, T30.LAV, T30.Priorità, t30.u_lavorazione, case when t30.Righe=t30.Trasferiti then 'OK' when t30.[Trasferiti parziali]>0 then 'PARZ' else 'NO' end as 'materiale', T30.U_PRG_TIR_MATERIALE, T30.BHEIGHT1,T30.BWIDTH1, T30.BLENGTH1
from
(
SELECT t20.ODP,T20.DOCENTRY, T20.DOC, T20.[N PEZ] , T20.COD, T20.Nome, t20.DIS, t20.RIS, t20.MAC, T20.ATT, T20.LAV, T20.Priorità, t20.u_lavorazione, sum(case when t21.itemtype=4 then 1 else 0 end) as 'Righe', sum(case when t21.itemtype=4 and T21.[U_PRG_WIP_QtaDaTrasf]=0  then 1 else 0 end) as 'Trasferiti' ,sum(case when t21.itemtype=4 and T21.[U_PRG_WIP_QtaSpedita]>0 and T21.[U_PRG_WIP_QtaDaTrasf]>0  then 1 else 0 end) as 'Trasferiti parziali' 
, T20.U_PRG_TIR_MATERIALE, T20.BHEIGHT1,T20.BWIDTH1, T20.BLENGTH1
FROM
(
Select t13.docnum as 'ODP', t13.docentry, 'ODP' AS 'DOC', T13.[PlannedQty] as 'N PEZ' , t13.itemcode as 'COD', t14.itemname as 'Nome', t14.u_disegno as 'DIS', t11.itemcode as 'RIS', t12.itemname as 'MAC', T11.[AdditQty] as 'ATT', T11.[PlannedQty] as 'LAV', T13.[U_Priorita_MES] AS 'Priorità', t13.u_lavorazione, CASE WHEN T14.U_PRG_TIR_MATERIALE IS NULL THEN '' ELSE T14.U_PRG_TIR_MATERIALE END AS 'U_PRG_TIR_MATERIALE' , CASE WHEN T14.BHEIGHT1 IS NULL THEN 0 ELSE T14.BHEIGHT1 END AS 'BHEIGHT1',T14.BWIDTH1, T14.BLENGTH1
from
(
SELECT t0.docentry, min(T0.[VisOrder]) as 'Visorder' , t0.u_stato_lavorazione
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry]
inner join ORSC T2 on t2.visrescode=t0.itemcode 
LEFT JOIN [TIRELLI_40].[DBO].APPOGGIO T3 ON T1.DOCNUM=T3.VALORE AND T3.TIPO='" & par_tipo_appoggio & "' AND T3.UTENTE=" & Homepage.ID_SALVATO & "
WHERE T1.[Status] ='R' and  T2.[ResType] ='M' and  T0.[U_Stato_lavorazione] = 'O' AND T3.VALORE IS NULL and substring(t1.u_produzione,1,3)='INT'
group by 
t0.docentry, t0.u_stato_lavorazione
)
as t10
inner join wor1 t11 on t11.docentry=t10.docentry and t10.visorder=t11.visorder
inner join oitm t12 on t12.itemcode=t11.itemcode
inner join owor t13 on t13.docentry=t11.docentry
inner join oitm t14 on t14.itemcode=t13.itemcode
inner join orsc t15 on t15.visrescode=t11.itemcode
where t15.resgrpcod=" & tipo_macchina & "
)
AS T20 left join wor1 t21 on t21.docentry=t20.docentry

group by
t20.ODP,T20.DOCENTRY, T20.DOC, T20.[N PEZ] , T20.COD, T20.Nome, t20.DIS, t20.RIS, t20.MAC, T20.ATT, T20.LAV, T20.Priorità, t20.u_lavorazione, T20.U_PRG_TIR_MATERIALE, T20.BHEIGHT1,T20.BWIDTH1, T20.BLENGTH1
)
as t30
where 0=0 " & filtro_materiale & filtro_spessore & filtro_descrizione & filtro_codice & filtro_odp & "
order by T30.[Priorità]
"





        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader
        If cmd_SAP_docentry_reader.Read() Then

            par_datagridview.Rows.Add(cmd_SAP_docentry_reader("ODP"), cmd_SAP_docentry_reader("DOC"), Math.Round(cmd_SAP_docentry_reader("N PEZ")), cmd_SAP_docentry_reader("COD"), cmd_SAP_docentry_reader("Nome"), cmd_SAP_docentry_reader("DIS"), cmd_SAP_docentry_reader("RIS"), cmd_SAP_docentry_reader("MAC"), Math.Round(cmd_SAP_docentry_reader("ATT")), Math.Round(cmd_SAP_docentry_reader("LAV")), cmd_SAP_docentry_reader("u_lavorazione"), cmd_SAP_docentry_reader("Priorità"), cmd_SAP_docentry_reader("Materiale"), cmd_SAP_docentry_reader("U_PRG_TIR_MATERIALE"), cmd_SAP_docentry_reader("BHEIGHT1"), "NO")
            ODP_Tree.AGGIUNGI_RECORD_APPOGGIO(Homepage.ID_SALVATO, par_tipo_appoggio, cmd_SAP_docentry_reader("ODP"))
            trova_taglio_con_stessa_materia_prima(par_datagridview, cmd_SAP_docentry_reader("ODP"), par_tipo_appoggio, tipo_macchina)
        End If
        cmd_SAP_docentry_reader.Close()
        Cnn.Close()


    End Sub

    Sub trova_taglio_con_stessa_materia_prima(par_datagridview As DataGridView, par_odp As Integer, PAR_TIPO_APPOGGIO As String, par_tipo_macchina As Integer)


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = Cnn




        CMD_SAP_docentry.CommandText = " 
Select t30.ODP,T30.DOCENTRY, T30.DOC, T30.[N PEZ] , T30.COD, T30.Nome, t30.DIS, t30.RIS, t30.MAC, T30.ATT, T30.LAV, coalesce(T31.u_Priorita_mes,0) as 'Priorità', coalesce(T31.u_lavorazione,0) as 'U_lavorazione', case when t30.Righe=t30.Trasferiti then 'OK' when t30.[Trasferiti parziali]>0 then 'PARZ' else 'NO' end as 'materiale', T30.U_PRG_TIR_MATERIALE, T30.BHEIGHT1,T30.BWIDTH1, T30.BLENGTH1
from
(
SELECT t20.ODP,T20.DOCENTRY, T20.DOC, T20.[N PEZ] , T20.COD, T20.Nome, t20.DIS, t20.RIS, t20.MAC, T20.ATT, T20.LAV, T20.Priorità, t20.u_lavorazione, sum(case when t21.itemtype=4 then 1 else 0 end) as 'Righe', sum(case when t21.itemtype=4 and T21.[U_PRG_WIP_QtaDaTrasf]=0  then 1 else 0 end) as 'Trasferiti' ,sum(case when t21.itemtype=4 and T21.[U_PRG_WIP_QtaSpedita]>0 and T21.[U_PRG_WIP_QtaDaTrasf]>0  then 1 else 0 end) as 'Trasferiti parziali' 
, T20.U_PRG_TIR_MATERIALE, T20.BHEIGHT1,T20.BWIDTH1, T20.BLENGTH1
FROM
(
select t3.docnum AS 'ODP','ODP' AS 'DOC', t3.docentry, T3.[PlannedQty] as 'N PEZ' , t3.itemcode as 'COD', t4.itemname as 'Nome', t4.u_disegno as 'DIS', t1.itemcode as 'RIS', t5.itemname as 'MAC', T2.[AdditQty] as 'ATT', T2.[PlannedQty] as 'LAV', T0.[U_Priorita_MES] AS 'Priorità', t0.u_lavorazione
, COALESCE(T4.U_PRG_TIR_MATERIALE,'')  AS 'U_PRG_TIR_MATERIALE' 
, COALESCE(T4.BHEIGHT1 , 0 ) AS 'BHEIGHT1',T4.BWIDTH1, T4.BLENGTH1
from
owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
INNER JOIN WOR1 T2 ON T2.ITEMCODE=T1.ITEMCODE 
INNER JOIN OWOR T3 ON T3.DOCENTRY=T2.DOCENTRY

inner join
(select t1.docentry
FROM WOR1 T0 inner join owor t1 on t0.docentry=T1.docentry
inner join ORSC T2 on t2.visrescode=t0.itemcode 
where T1.[Status] ='R' and  T2.[ResType] ='M' and  T0.[U_Stato_lavorazione] = 'O' and substring(t1.u_produzione,1,3)='INT' and t2.resgrpcod=" & par_tipo_macchina & " 
group by t1.docentry) A on A.docentry=t2.docentry
INNER JOIN OITM T4 ON T4.ITEMCODE=T3.ITEMCODE
INNER JOIN OITM T5 ON T5.ITEMCODE=T2.ITEMCODE
LEFT JOIN [TIRELLI_40].[DBO].APPOGGIO T6 ON T3.DOCNUM=T6.VALORE AND T6.TIPO='" & PAR_TIPO_APPOGGIO & "' AND T6.UTENTE='" & Homepage.ID_SALVATO & "'

inner join 
(
select t10.docentry
from
(
select t0.docentry, min(T0.[VisOrder]) as 'Visorder' ,t2.resgrpcod
from wor1 t0 inner join owor t1 on t0.docentry=t1.docentry
inner join ORSC T2 on t2.visrescode=t0.itemcode
where t1.status='R' and t0.u_stato_lavorazione ='O' and   T2.[ResType] ='M'
group by t0.docentry ,t2.resgrpcod
)
as t10
where t10.resgrpcod=7
)
B on B.docentry=t1.docentry

where t3.Status='R' and substring(t3.u_produzione,1,3)='INT' AND T6.VALORE IS NULL AND t1.ITEMTYPE =4 and t2.ITEMTYPE =4 AND T0.DOCNUM =" & par_odp & " and t1.visorder =" & trova_prima_riga_macchina(par_odp) & "
)
AS T20 left join wor1 t21 on t21.docentry=t20.docentry

group by
t20.ODP,T20.DOCENTRY, T20.DOC, T20.[N PEZ] , T20.COD, T20.Nome, t20.DIS, t20.RIS, t20.MAC, T20.ATT, T20.LAV, T20.Priorità, t20.u_lavorazione, T20.U_PRG_TIR_MATERIALE, T20.BHEIGHT1,T20.BWIDTH1, T20.BLENGTH1
)
as t30 inner join owor t31 on t31.docnum=t30.odp
where 0=0 " & filtro_materiale & filtro_spessore & filtro_descrizione & filtro_codice & filtro_odp & "
order by coalesce(T31.u_Priorita_mes,0)


"


        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader
        Do While cmd_SAP_docentry_reader.Read()

            par_datagridview.Rows.Add(cmd_SAP_docentry_reader("ODP"), cmd_SAP_docentry_reader("DOC"), Math.Round(cmd_SAP_docentry_reader("N PEZ")), cmd_SAP_docentry_reader("COD"), cmd_SAP_docentry_reader("Nome"), cmd_SAP_docentry_reader("DIS"), cmd_SAP_docentry_reader("RIS"), "", Math.Round(cmd_SAP_docentry_reader("ATT")), Math.Round(cmd_SAP_docentry_reader("LAV")), cmd_SAP_docentry_reader("u_lavorazione"), cmd_SAP_docentry_reader("Priorità"), cmd_SAP_docentry_reader("Materiale"), cmd_SAP_docentry_reader("U_PRG_TIR_MATERIALE"), cmd_SAP_docentry_reader("BHEIGHT1"), "SI")

            ODP_Tree.AGGIUNGI_RECORD_APPOGGIO(Homepage.ID_SALVATO, PAR_TIPO_APPOGGIO, cmd_SAP_docentry_reader("ODP"))

        Loop
        cmd_SAP_docentry_reader.Close()
        Cnn.Close()

        lISTA_odp_taglio(par_datagridview, PAR_TIPO_APPOGGIO)
    End Sub



    Sub lISTA_odp_BINT(par_datagridview As DataGridView, par_odp As String, par_itemcode As String, par_itemname As String)

        Dim filtro_odp As String
        Dim filtro_itemcode As String
        Dim filtro_itemname As String
        Dim filtro_commessa As String
        Dim filtro_cliente As String

        If par_odp = "" Then

            filtro_odp = ""
        Else
            filtro_odp = " and t0.docnum Like '%%" & par_odp & "%%' "
        End If
        If par_itemcode = "" Then

            filtro_itemcode = ""
        Else
            filtro_itemcode = " and t0.itemcode Like '%%" & par_itemcode & "%%' "
        End If

        If par_itemname = "" Then

            filtro_itemname = ""
        Else
            filtro_itemname = " and t0.prodname Like '%%" & par_itemname & "%%' "
        End If




        par_datagridview.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = Cnn


        CMD_SAP_docentry.CommandText = "select t0.docnum,T0.ITEMCODE, COALESCE(T1.U_CODICE_BRB,'') AS 'Codice_BRB', T1.U_Disegno, T0.PRODNAME, T0.PlannedQty-T0.CmpltQty AS 'Q', T0.U_PRG_AZS_Commessa
,coalesce(t2.U_Final_customer_name,t0.u_utilizz) as 'Cliente'
, coalesce(t4.ocrcode,case when coalesce(t3.location,'')='13' then 'BRB01' ELSE 'TIR01' END) AS 'DIV'
,T0.PostDate AS 'DATA_CREAZIONE'
,T0.DUEDATE AS 'DATA_CONSEGNA'
,coalesce(t0.u_stato,'') as 'Status'
from owor T0 INNER JOIN OITM T1 ON T0.ITEMCODE=T1.ITEMCODE
left join oitm t2 on t2.itemcode=t0.u_prg_azs_commessa
left join owhs t3 on t3.whscode=t0.Warehouse
left join rdr1 t4 on t4.itemcode=t0.u_prg_azs_commessa and t4.openqty>0
where T0.u_produzione='B_INT' AND (T0.STATUS='P' OR T0.STATUS='R') " & filtro_odp & filtro_itemcode & filtro_itemname & filtro_commessa & filtro_cliente & "
ORDER BY T0.DUEDATE"


        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader
        Do While cmd_SAP_docentry_reader.Read()

            par_datagridview.Rows.Add(cmd_SAP_docentry_reader("docnum"), cmd_SAP_docentry_reader("ITEMCODE"), cmd_SAP_docentry_reader("Codice_BRB"), cmd_SAP_docentry_reader("U_DISEGNO"), cmd_SAP_docentry_reader("PRODNAME"), cmd_SAP_docentry_reader("q"), cmd_SAP_docentry_reader("U_PRG_AZS_Commessa"), cmd_SAP_docentry_reader("Cliente"), cmd_SAP_docentry_reader("DIV"), cmd_SAP_docentry_reader("Status"), cmd_SAP_docentry_reader("DATA_CREAZIONE"), cmd_SAP_docentry_reader("DATA_CONSEGNA"))

        Loop
        cmd_SAP_docentry_reader.Close()
        Cnn.Close()


    End Sub

    Public Function trova_prima_riga_macchina(par_docnum As Integer)
        Dim visorder As Integer




        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1



        CMD_SAP_2.CommandText = "
select t1.docentry, min( t0.VisOrder ) as 'Visorder'
FROM WOR1 T0 inner join owor t1 on t0.docentry=T1.docentry
inner join ORSC T2 on t2.visrescode=t0.itemcode 
where T1.[Status] ='R' and  T2.[ResType] ='M' and  T0.[U_Stato_lavorazione] = 'O' and substring(t1.u_produzione,1,3)='INT' and t1.docentry=" & par_docnum & "
group by t1.docentry

"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            visorder = cmd_SAP_reader_2("visorder")

        End If



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        Return visorder

    End Function


    Sub lISTA_tutti_gli_odp()

        DataGridView_ODP_TIPO_MACCHINA.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = Cnn

        CMD_SAP_docentry.CommandText = "Select t30.ODP, T30.DOCENTRY, T30.DOC, T30.[N PEZ] , T30.COD, T30.Nome, t30.DIS, t30.RIS, t30.MAC, T30.ATT, T30.LAV, T30.Priorità, t30.u_lavorazione, case when t30.Righe=t30.Trasferiti then 'OK' when t30.[Trasferiti parziali]>0 then 'PARZ' else 'NO' end as 'materiale', T30.U_PRG_TIR_MATERIALE, T30.BHEIGHT1,T30.BWIDTH1, T30.BLENGTH1
from
(
SELECT t20.ODP,T20.DOCENTRY, T20.DOC, T20.[N PEZ] , T20.COD, T20.Nome, t20.DIS, t20.RIS, t20.MAC, T20.ATT, T20.LAV, T20.Priorità, t20.u_lavorazione, sum(case when t21.itemtype=4 then 1 else 0 end) as 'Righe', sum(case when t21.itemtype=4 and T21.[U_PRG_WIP_QtaDaTrasf]=0  then 1 else 0 end) as 'Trasferiti' ,sum(case when t21.itemtype=4 and T21.[U_PRG_WIP_QtaSpedita]>0 and T21.[U_PRG_WIP_QtaDaTrasf]>0  then 1 else 0 end) as 'Trasferiti parziali' 
, T20.U_PRG_TIR_MATERIALE, T20.BHEIGHT1,T20.BWIDTH1, T20.BLENGTH1
FROM
(
Select t13.docnum as 'ODP',T13.DOCENTRY, 'ODP' AS 'DOC', T13.[PlannedQty] as 'N PEZ' , t13.itemcode as 'COD', t14.itemname as 'Nome', t14.u_disegno as 'DIS', t11.itemcode as 'RIS', t12.itemname as 'MAC', T11.[AdditQty] as 'ATT', T11.[PlannedQty] as 'LAV', T13.[U_Priorita_MES] AS 'Priorità', t13.u_lavorazione, CASE WHEN T14.U_PRG_TIR_MATERIALE IS NULL THEN '' ELSE T14.U_PRG_TIR_MATERIALE END AS 'U_PRG_TIR_MATERIALE' , CASE WHEN T14.BHEIGHT1 IS NULL THEN 0 ELSE T14.BHEIGHT1 END AS 'BHEIGHT1',T14.BWIDTH1, T14.BLENGTH1
from
(
SELECT t0.docentry, min(T0.[VisOrder]) as 'Visorder' , t0.u_stato_lavorazione
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry]
inner join ORSC T2 on t2.visrescode=t0.itemcode WHERE T1.[Status] ='R' and  T2.[ResType] ='M' and  T0.[U_Stato_lavorazione] = 'O'
group by 
t0.docentry, t0.u_stato_lavorazione
)
as t10
inner join wor1 t11 on t11.docentry=t10.docentry and t10.visorder=t11.visorder
inner join oitm t12 on t12.itemcode=t11.itemcode
inner join owor t13 on t13.docentry=t11.docentry
inner join oitm t14 on t14.itemcode=t13.itemcode
inner join orsc t15 on t15.visrescode=t11.itemcode
)
AS T20 left join wor1 t21 on t21.docentry=t20.docentry

group by
t20.ODP,T20.DOCENTRY, T20.DOC, T20.[N PEZ] , T20.COD, T20.Nome, t20.DIS, t20.RIS, t20.MAC, T20.ATT, T20.LAV, T20.Priorità, t20.u_lavorazione, T20.U_PRG_TIR_MATERIALE, T20.BHEIGHT1,T20.BWIDTH1, T20.BLENGTH1
)
as t30
where 0=0 " & filtro_materiale & filtro_spessore & filtro_descrizione & filtro_codice & filtro_odp & "
order by T30.Priorità"

        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader
        Do While cmd_SAP_docentry_reader.Read()

            DataGridView_ODP_TIPO_MACCHINA.Rows.Add(cmd_SAP_docentry_reader("ODP"), cmd_SAP_docentry_reader("DOC"), Math.Round(cmd_SAP_docentry_reader("N PEZ")), cmd_SAP_docentry_reader("COD"), cmd_SAP_docentry_reader("Nome"), cmd_SAP_docentry_reader("DIS"), cmd_SAP_docentry_reader("RIS"), cmd_SAP_docentry_reader("MAC"), Math.Round(cmd_SAP_docentry_reader("ATT")), Math.Round(cmd_SAP_docentry_reader("LAV")), cmd_SAP_docentry_reader("u_lavorazione"), cmd_SAP_docentry_reader("Priorità"), cmd_SAP_docentry_reader("Materiale"), cmd_SAP_docentry_reader("U_PRG_TIR_MATERIALE"), cmd_SAP_docentry_reader("BHEIGHT1"))

        Loop
        cmd_SAP_docentry_reader.Close()
        Cnn.Close()


    End Sub


    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        If TextBox8.Text = Nothing Then
            filtro_odp = ""
        Else
            filtro_odp = " and T30.odp Like '%%" & TextBox8.Text & "%%'  "
        End If
        riempi_datagridview(TextBox8.Text, TextBox7.Text, TextBox5.Text, TextBox4.Text, TextBox17.Text)
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.Text = Nothing Then
            filtro_codice = ""
        Else
            filtro_codice = " and T30.cod Like '%%" & TextBox7.Text & "%%'  "
        End If
        riempi_datagridview(TextBox8.Text, TextBox7.Text, TextBox5.Text, TextBox4.Text, TextBox17.Text)
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text = Nothing Then
            filtro_descrizione = ""
        Else
            filtro_descrizione = " and T30.Nome Like '%%" & TextBox5.Text & "%%'  "
        End If
        riempi_datagridview(TextBox8.Text, TextBox7.Text, TextBox5.Text, TextBox4.Text, TextBox17.Text)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = Nothing Then
            filtro_materiale = ""
        Else
            filtro_materiale = " and T30.U_PRG_TIR_MATERIALE Like '%%" & TextBox4.Text & "%%'  "
        End If
        riempi_datagridview(TextBox8.Text, TextBox7.Text, TextBox5.Text, TextBox4.Text, TextBox17.Text)
    End Sub





    Private Sub DataGridView_ODP_TIPO_MACCHINA_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP_TIPO_MACCHINA.CellClick

        If e.RowIndex >= 0 Then

            If e.ColumnIndex = DataGridView_ODP_TIPO_MACCHINA.Columns.IndexOf(ODP) Then

                docnum = DataGridView_ODP_TIPO_MACCHINA.Rows(e.RowIndex).Cells(columnName:="ODP").Value
                Button3.Text = docnum
                invia_Odp(docnum)
                righe_ODP_macchine(DataGridView1, docnum)
                materia_prima_macchina(DataGridView6, docnum)
                rECORD_CARICATI(DataGridView2)
                Presenza_attrezzaggio_lavorazione()
                tipo_autocontrollo = 1
                check_autocontrollo_attrezzaggio()
                tipo_autocontrollo = 2
                check_autocontrollo_lavorazione()
                Panel19.Visible = True
                CHECK_AUTOCONTROLLO_CARICATI()


            End If

            If e.ColumnIndex = DataGridView_ODP_TIPO_MACCHINA.Columns.IndexOf(Nesting) Then
                Nesting_riga = DataGridView_ODP_TIPO_MACCHINA.Rows(e.RowIndex).Cells(columnName:="Nesting").Value
                filtro_nesting()
            End If


            If e.ColumnIndex = DataGridView_ODP_TIPO_MACCHINA.Columns.IndexOf(COD) Then
                Magazzino.Codice_SAP = DataGridView_ODP_TIPO_MACCHINA.Rows(e.RowIndex).Cells(columnName:="COD").Value
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

    Sub invia_Odp(par_docnum_odp)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = Cnn
        If Homepage.ERP_provenienza = "SAP" Then


            CMD_SAP_docentry.CommandText = "SELECT t0.docentry AS 'Doc'
, 0 as 'numbolla'
, as 'Code_fase'
, case when t0.u_stato is null then '' else t0.u_stato end as 'U_stato', T0.[ItemCode] as 'Cod',CAST(case when t1.u_disegno is null then 'Senza disegno' else T1.[u_disegno] end as Varchar(20)) AS 'Disegno', T0.PRODNAME as 'Nome',T0.[PlannedQty] as 'Quant', case when t0.U_lavorazione is null then 0 else t0.U_lavorazione end as 'N lavorazione' , case when T2.CODICE_CAM is null then '' else t2.codice_CAM end as 'Codice_CAM'
,T0.warehouse as 'Mag_destinazione'
FROM OWOR T0 INNER JOIN OITM T1 ON T0.ITEMCODE=T1.ITEMCODE LEFT JOIN CAM T2 ON T2.CODICE_DISEGNO=T1.U_DISEGNO
WHERE T0.DOCNUM= '" & par_docnum_odp & "' AND (T0.STATUS='R' OR T0.STATUS='P')"
        Else
            CMD_SAP_docentry.CommandText = "
select t10.odp as 'Doc'
,t10.stato as 'U_stato'
,trim(t10.codart_odp) as 'Cod'
,trim(t10.disegno) as 'Disegno'
,dscodart_odp as 'Nome'
,qta_res as 'Quant'
,999 as 'N lavorazione'
,'' as 'Codice CAM'
,'' as 'mag_destinazione'
,numbolla
,code_fase

FROM OPENQUERY(AS400, '
select *
from
S786FAD1.TIR90VIS.JGALodpmu
where odp=''" & par_docnum_odp & "'' and cod_risorsa=''" & risorsa & "'''

) as t10"
        End If
        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader
        If cmd_SAP_docentry_reader.Read() Then

            docentry = Val(cmd_SAP_docentry_reader("Doc"))
            codice = cmd_SAP_docentry_reader("Cod")
            Descart = cmd_SAP_docentry_reader("Nome")
            disegno = cmd_SAP_docentry_reader("Disegno")
            quantodp = Val(cmd_SAP_docentry_reader("Quant"))
            ' codice_CAM = Val(cmd_SAP_docentry_reader("Codice_CAM"))
            N_lavorazione = Val(cmd_SAP_docentry_reader("N lavorazione"))
            'Label29.Text = cmd_SAP_docentry_reader("Mag_destinazione")
            num_bolla = cmd_SAP_docentry_reader("numbolla")
            Label2.Text = num_bolla
            'If Label29.Text = "CAP2" Then
            '    Label29.BackColor = Color.Blue

            'Else
            '    Label29.BackColor = Color.Green
            'End If

            Button2.Text = codice
            Label24.Text = Descart
            Button4.Text = disegno
            Label1.Text = (cmd_SAP_docentry_reader("u_stato"))
            LabelQuantitàSAP.Text = quantodp
            TextBox11.Text = quantodp
            code_fase = cmd_SAP_docentry_reader("Code_fase")
            Magazzino.visualizza_picture(disegno, PictureBox2)

        Else
            MsgBox("Ordine non trovato")
        End If
        cmd_SAP_docentry_reader.Close()
        Cnn.Close()

        Button28.Enabled = True
        rECORD_CARICATI(DataGridView2)
        RichTextBox2.Text = 0
        RichTextBox3.Text = 0
    End Sub

    Sub righe_ODP_macchine(par_datagridview As DataGridView, par_docnum As String)

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_2.CommandText = "SELECT t0.itemname, T0.[AdditQty], T0.[PlannedQty], T0.[U_Stato_lavorazione]  FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] inner join ORSC T2 on T2.[VisResCode]=t0.itemcode
        WHERE T1.[DocNum] ='" & par_docnum & "' and T2.[ResType]='M' order by T0.[VisOrder]"

        Else

            CMD_SAP_2.CommandText = "

		select odp as 'N_ODP'
		,coalesce(risorsa,desc_fase) as 'itemname'
		,trim(code_fase) as 'Code_fase'
		,999 as 'additqty'
		,999 as 'Plannedqty'
		,stato_fase as 'u_stato_lavorazione'
		FROM OPENQUERY(AS400, '
select *
from
S786FAD1.TIR90VIS.JGALodpmu
where odp=''" & par_docnum & "'' and tipo_macchina=''M''
order by code_fase
') as t10"


        End If


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()


            par_datagridview.Rows.Add(cmd_SAP_reader_2("code_fase"), cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("AdditQty"), cmd_SAP_reader_2("PlannedQty") - cmd_SAP_reader_2("AdditQty"), cmd_SAP_reader_2("u_stato_lavorazione"))

        Loop


        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        par_datagridview.ClearSelection()
    End Sub

    Sub materia_prima_macchina(par_datagridview As DataGridView, par_docnum As Integer)

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_2.CommandText = "SELECT t0.itemcode, t0.itemname, T0.[AdditQty], T0.[PlannedQty], T0.[U_Stato_lavorazione]  
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] 
        WHERE T1.[DocNum] ='" & docnum & "' and t0.itemtype='4'
order by T0.[VisOrder]"
        Else
            CMD_SAP_2.CommandText = "select odp 
,trim(codart) as 'itemcode'
,des_code as 'itemname' 
,0 as 'additqty'
,qtapia as 'Plannedqty'
,qtadatra as 'Da prelevare'
,qtatra as 'Prelevato'
,saldo_imp as 'u_stato_lavorazione'

FROM OPENQUERY(AS400, '
select t0.odp
,t0.codart
, t1.des_code,t0.qtapia,t0.qtadatra,t0.qtatra,t0.saldo_imp
from
S786FAD1.TIR90VIS.JGALimp t0
LEFT JOIN S786FAD1.TIR90VIS.JGALart t1 
        ON t0.codart = t1.code
where odp=''" & docnum & "''


'
) as t10"
        End If


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Try

            Do While cmd_SAP_reader_2.Read()


                par_datagridview.Rows.Add(cmd_SAP_reader_2("itemcode"), cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("PlannedQty"))

            Loop

        Catch ex As Exception
            MsgBox("L'ordine" & cmd_SAP_reader_2("N° Ordine") & "presenta un errore")
        End Try
        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    Sub rECORD_CARICATI(par_datagridview As DataGridView)


        par_datagridview.Rows.Clear()

        Dim Cnn1 As New SqlConnection

            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_2 As New SqlCommand
            Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then

            CMD_SAP_2.CommandText = "SELECT T0.ID,T1.LASTNAME,T0.START, T0.CONSUNTIVO, T0.TIPOLOGIA_LAVORAZIONE
FROM MANODOPERA T0 LEFT JOIN [TIRELLI_40].[dbo].OHEM T1 ON T1.EMPID=T0.DIPENDENTE
LEFT JOIN ORSC T2 ON T2.VISRESCODE=T0.RISORSA
WHERE DOCNUM='" & docnum & "' AND T2.[ResType]='m'
ORDER BY T0.ID"
        Else
            CMD_SAP_2.CommandText = "select 
t10.rigayp as 'ID'
,CDDIYP 
,t10.desc_dip as 'Lastname'
,DTMOYP as 'Start'
,FATMYP
,FATUYP
,FLAMYP
,FLAUYP
,SAACYP

FROM OPENQUERY(AS400, '
select *
from
S786FAD1.TIR90VIS.YPCMOV0F t0
left join S786FAD1.TIR90VIS.JGALDIP t1 on t0.CDDIYP=t1.cod_dip
where orpryp=''" & docnum & "''
'
)
as t10
order by t10.rigayp"
        End If
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


            Do While cmd_SAP_reader_2.Read()

            If Homepage.ERP_provenienza = "SAP" Then


                par_datagridview.Rows.Add(cmd_SAP_reader_2("ID"), cmd_SAP_reader_2("Lastname"), cmd_SAP_reader_2("Start"), cmd_SAP_reader_2("Consuntivo"), cmd_SAP_reader_2("tipologia_lavorazione"), cmd_SAP_reader_2("SAACYP"))
            Else
                par_datagridview.Rows.Add(
        cmd_SAP_reader_2("ID"),
        cmd_SAP_reader_2("Lastname"),
        cmd_SAP_reader_2("Start"),
        cmd_SAP_reader_2("FATMYP"),
        cmd_SAP_reader_2("FATUYP"),
        cmd_SAP_reader_2("FLAMYP"),
        cmd_SAP_reader_2("FLAUYP")
    )
            End If
        Loop


            cmd_SAP_reader_2.Close()
            Cnn1.Close()


    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter
        Timer4_0.Start()
    End Sub
    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Enter
        Timer4_0.Stop()
    End Sub

    Sub Macchine_In_Movimento()
        Dim CNN2 As New SqlConnection
        CNN2.ConnectionString = Pianificazione.DATABASE_MU_4_0
        CNN2.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        CMD_SAP_2.Connection = CNN2
        CMD_SAP_2.CommandText = "SELECT top 1000 t0.machineid,
case when t0.machineid=1 then 'Doosan 4' when t0.machineid=2 then 'Doosan 6' end as 'Macchina'
,t0.inizio,t0.fine
      ,t0.[DurataTotale], t0.activityid, t1.name , t0.[PartProgramNumber]
                            From [dbo].[Transazioni] t0 
                            LEFT JOIN [dbo].[Activities] t1 ON t0.activityid = t1.ActivityId 
                            where t0.fine is null
                            Order By t0.IDTransazione DESC"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            Dim machineId As Integer = CInt(cmd_SAP_reader_2("machineid"))
            Dim activityId As Integer = CInt(cmd_SAP_reader_2("activityid"))

            If machineId = 1 Then
                If activityId = 11 Then
                    AggiornaPannello(Label11, Label6, Label13, Label8, Panel45, Color.Lime)
                ElseIf activityId = 10 Then
                    AggiornaPannello(Label13, Label8, Label11, Label6, Panel45, Color.Red)


                End If
                Label4.Text = AggiornaEtichetta(cmd_SAP_reader_2("name"), cmd_SAP_reader_2("inizio"))
            ElseIf machineId = 2 Then
                If activityId = 11 Then
                    AggiornaPannello(Label19, Label21, Label17, Label23, Panel61, Color.Lime)
                ElseIf activityId = 10 Then
                    AggiornaPannello(Label17, Label23, Label19, Label21, Panel61, Color.Red)
                End If
                Label5.Text = AggiornaEtichetta(cmd_SAP_reader_2("name"), cmd_SAP_reader_2("inizio"))
            End If
        Loop

        cmd_SAP_reader_2.Close()
        CNN2.Close()

    End Sub

    Sub log_4_0()
        DataGridView3.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Pianificazione.DATABASE_MU_4_0

        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT top 50
IDTransazione
,t0.machineid
,case when t0.machineid=1 then 'Doosan 4' when t0.machineid=2 then 'Doosan 6' end as 'Macchina'
,t0.inizio
,t0.fine
 ,t0.[DurataTotale]/60000 as 'Minuti'
, t0.activityid
, t1.name
, t0.[PartProgramNumber]

                            From [dbo].[Transazioni] t0 
                            LEFT JOIN [dbo].[Activities] t1 ON t0.activityid = t1.ActivityId 
                            where T0.[MACHINEID]   Like '%%" & ComboBox3.Text & "%%' 
                            Order By t0.IDTransazione DESC"



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()

            DataGridView3.Rows.Add(cmd_SAP_reader_2("IDTransazione"), cmd_SAP_reader_2("machineid"), cmd_SAP_reader_2("Macchina"), cmd_SAP_reader_2("Inizio"), cmd_SAP_reader_2("Fine"), cmd_SAP_reader_2("Minuti"), cmd_SAP_reader_2("name"), cmd_SAP_reader_2("partprogramnumber"))
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    Sub AggiornaPannello(label1 As Label, label2 As Label, label3 As Label, label4 As Label, panel As Panel, colore As Color)

        panel.BackColor = colore
    End Sub

    Function AggiornaEtichetta(name As Object, inizio As Object) As String
        Return $"{name} dalle: {Hour(inizio)}:{Minute(inizio)}:{Second(inizio)}"
    End Function

    Sub Presenza_attrezzaggio_lavorazione()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = Cnn
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_docentry.CommandText = "SELECT case when t0.U_Presenza_attrezzaggio is null 
then 'Y' else t0.u_presenza_attrezzaggio end as 'Presenza_attrezzaggio',
case when t0.U_Presenza_lavorazione is null then 'Y' else t0.u_presenza_lavorazione end as 'Presenza_lavorazione'
FROM ORSC t0
where t0.VisResCode='" & risorsa & "'"
        Else
            CMD_SAP_docentry.CommandText = "SELECT 'Y' as 'Presenza_attrezzaggio',
'Y'  as 'Presenza_lavorazione'"
        End If


        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader
        If cmd_SAP_docentry_reader.Read() Then

            If cmd_SAP_docentry_reader("presenza_attrezzaggio") = "Y" Then
                GroupBox12.Visible = True
            Else
                GroupBox12.Visible = False
            End If

            If cmd_SAP_docentry_reader("presenza_lavorazione") = "Y" Then
                GroupBox13.Visible = True
            Else
                GroupBox13.Visible = False
            End If


        Else
            MsgBox("Ordine non trovato")
        End If
        cmd_SAP_docentry_reader.Close()
        Cnn.Close()


    End Sub

    Sub filtro_nesting()

        DataGridView_ODP_TIPO_MACCHINA.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = Cnn



        CMD_SAP_docentry.CommandText = "Select t30.ODP,T30.DOCENTRY, T30.DOC, T30.[N PEZ] , T30.COD, T30.Nome, t30.DIS, t30.RIS, t30.MAC, T30.ATT, T30.LAV, T30.Priorità, t30.u_lavorazione, case when t30.Righe=t30.Trasferiti then 'OK' when t30.[Trasferiti parziali]>0 then 'PARZ' else 'NO' end as 'materiale'
from
(
SELECT t20.ODP,T20.DOCENTRY, T20.DOC, T20.[N PEZ] , T20.COD, T20.Nome, t20.DIS, t20.RIS, t20.MAC, T20.ATT, T20.LAV, T20.Priorità, t20.u_lavorazione, sum(case when t21.itemtype=4 then 1 else 0 end) as 'Righe', sum(case when t21.itemtype=4 and T21.[U_PRG_WIP_QtaDaTrasf]=0  then 1 else 0 end) as 'Trasferiti' ,sum(case when t21.itemtype=4 and T21.[U_PRG_WIP_QtaSpedita]>0 and T21.[U_PRG_WIP_QtaDaTrasf]>0  then 1 else 0 end) as 'Trasferiti parziali' 
FROM
(
Select t13.docnum as 'ODP', t13.docentry, 'ODP' AS 'DOC', T13.[PlannedQty] as 'N PEZ' , t13.itemcode as 'COD', t14.itemname as 'Nome', t14.u_disegno as 'DIS', t11.itemcode as 'RIS', t12.itemname as 'MAC', T11.[AdditQty] as 'ATT', T11.[PlannedQty] as 'LAV', T13.[U_Priorita_MES] AS 'Priorità', t13.u_lavorazione
from
(
SELECT t0.docentry, min(T0.[VisOrder]) as 'Visorder' , t0.u_stato_lavorazione
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry]
inner join ORSC T2 on t2.visrescode=t0.itemcode 
WHERE T1.[Status] ='R' and  T2.[ResType] ='M'  and  T0.[U_Stato_lavorazione] = 'O'
group by 
t0.docentry, t0.u_stato_lavorazione
)
as t10
inner join wor1 t11 on t11.docentry=t10.docentry and t10.visorder=t11.visorder
inner join oitm t12 on t12.itemcode=t11.itemcode
inner join owor t13 on t13.docentry=t11.docentry
inner join oitm t14 on t14.itemcode=t13.itemcode
inner join orsc t15 on t15.visrescode=t11.itemcode
where t13.u_lavorazione = '" & Nesting_riga & "'
)
AS T20 left join wor1 t21 on t21.docentry=t20.docentry

group by
t20.ODP,T20.DOCENTRY, T20.DOC, T20.[N PEZ] , T20.COD, T20.Nome, t20.DIS, t20.RIS, t20.MAC, T20.ATT, T20.LAV, T20.Priorità, t20.u_lavorazione
)
as t30 order by T30.[Priorità]"


        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader
        Do While cmd_SAP_docentry_reader.Read()
            Dim img As Image = Nothing
            Dim codiceDisegno As String = If(IsDBNull(cmd_SAP_docentry_reader("DIS")), "", cmd_SAP_docentry_reader("DIS").ToString())
            Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & codiceDisegno & ".PNG"

            If codiceDisegno <> "" AndAlso File.Exists(percorso) Then
                Using fs As New FileStream(percorso, FileMode.Open, FileAccess.Read)
                    Using tmp As Image = Image.FromStream(fs)
                        Img = New Bitmap(tmp) ' evita lock sul file
                    End Using
                End Using
            End If


            DataGridView_ODP_TIPO_MACCHINA.Rows.Add(cmd_SAP_docentry_reader("ODP"), cmd_SAP_docentry_reader("DOC"), Math.Round(cmd_SAP_docentry_reader("N PEZ")), cmd_SAP_docentry_reader("COD"), cmd_SAP_docentry_reader("Nome"), cmd_SAP_docentry_reader("DIS"), img, cmd_SAP_docentry_reader("RIS"), cmd_SAP_docentry_reader("MAC"), Math.Round(cmd_SAP_docentry_reader("ATT")), Math.Round(cmd_SAP_docentry_reader("LAV")), cmd_SAP_docentry_reader("u_Lavorazione"), cmd_SAP_docentry_reader("Priorità"), cmd_SAP_docentry_reader("Materiale"))


        Loop
        cmd_SAP_docentry_reader.Close()
        Cnn.Close()


    End Sub





    Private Sub DataGridView_ODP_TIPO_MACCHINA_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_ODP_TIPO_MACCHINA.CellFormatting
        If DataGridView_ODP_TIPO_MACCHINA.Rows(e.RowIndex).Cells(columnName:="catena").Value = "SI" Then
            DataGridView_ODP_TIPO_MACCHINA.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Yellow

        ElseIf DataGridView_ODP_TIPO_MACCHINA.Rows(e.RowIndex).Cells(columnName:="Nesting").Value > 0 Then
            DataGridView_ODP_TIPO_MACCHINA.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Aqua

        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Carico_macchine.start_carico_macchine()

        Carico_macchine.Show()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        TableLayoutPanel17.Visible = True
        Button11.BackColor = Color.Lime
        Button9.BackColor = Color.Transparent
        Button14.Text = "INSERISCI LAVORAZIONE"
        Button14.Visible = True
        GroupBox9.Text = "Tempo ciclo"
        GroupBox10.Text = "Tempo cambio pezzo"
        GroupBox10.Visible = True
        GroupBox33.Text = "Imprevisti"
        GroupBox34.Visible = True
        GroupBox35.Visible = True

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TableLayoutPanel17.Visible = True
        Button9.BackColor = Color.Lime
        Button11.BackColor = Color.Transparent
        Button14.Text = "INSERISCI ATTREZZAGGIO"
        Button14.Visible = True
        GroupBox9.Text = "Tempo attrezzaggio"
        GroupBox10.Visible = False
        GroupBox34.Visible = False
        GroupBox35.Visible = False
    End Sub



    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
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
        If TextBox1.TextLength = 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If

        ' ovviamente se c'è già un punto
        ' non è consentito digitarne altri
        If (KeyAscii = 46) And
            TextBox1.Text.IndexOf(".") > 0 Then
            KeyAscii = 0
        End If

        ' Reimposto il keychar
        e.KeyChar = Chr(KeyAscii)

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If GroupBox10.Visible = True Then
            If TextBox2.Text = Nothing Then
                tempo_cambio_pezzo = 0
            Else
                tempo_cambio_pezzo = TextBox2.Text
            End If

            If TextBox1.Text = Nothing Then
                tempo_CICLO_pezzo = 0
            Else
                tempo_CICLO_pezzo = TextBox1.Text
            End If

            time = tempo_CICLO_pezzo * TextBox11.Text + tempo_cambio_pezzo * TextBox11.Text
            Label27.Text = time
        End If
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
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
        If TextBox2.TextLength = 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If

        ' ovviamente se c'è già un punto
        ' non è consentito digitarne altri
        If (KeyAscii = 46) And
            TextBox2.Text.IndexOf(".") > 0 Then
            KeyAscii = 0
        End If

        ' Reimposto il keychar
        e.KeyChar = Chr(KeyAscii)



    End Sub



    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
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
        If TextBox9.TextLength = 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If

        ' ovviamente se c'è già un punto
        ' non è consentito digitarne altri
        If (KeyAscii = 46) And
            TextBox9.Text.IndexOf(".") > 0 Then
            KeyAscii = 0
        End If

        ' Reimposto il keychar
        e.KeyChar = Chr(KeyAscii)
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        docnum_odp = Button3.Text

        If TextBox9.Text <> Nothing Then
            If ComboBox1.SelectedIndex = -1 Then
                MsgBox("Scegliere una causale imprevisto")

            Else
                tipo_lav = "I"
                time = TextBox9.Text
                percentuale = 5
                lavorazione = risorsa

                inserisci_lavorazione_a_sap_macchina(lavorazione)
                inserisci_lavorazione_a_sap_macchina("R00500")
                inserisci_record_imprevisto()
                ciclo_inserimento_lavorazione()


            End If

        Else
            ciclo_inserimento_lavorazione()

        End If
        kpi_mensile_FASI()

    End Sub

    Sub ciclo_inserimento_lavorazione()
        risorsa_old = ""
        If Button9.BackColor = Color.Lime Then

            If TextBox1.Text = Nothing And TextBox9.Text = Nothing Then
                MsgBox("Inserire un tempo")
            Else

                tipo_autocontrollo = 1
                tipo_lav = "A"
                time = TextBox1.Text
                percentuale = 5
                lavorazione = risorsa

                inserisci_lavorazione_a_sap_macchina(lavorazione)

                risorsa_old = lavorazione
                lavorazione = "R00500"
                id_padre = id_manodopera
                inserisci_lavorazione_a_sap_manodopera()

                rECORD_CARICATI(DataGridView2)
                pulisci_campi_manodopera()



                If Button10.BackColor = Color.Red And Button10.Visible = True Then

                    inizializzazione_autocontrollo()

                End If

            End If

        ElseIf Button11.BackColor = Color.Lime Then

            If TextBox1.Text = Nothing And TextBox9.Text = Nothing Then
                MsgBox("Inserire un tempo")
            Else
                tipo_autocontrollo = 2
                tipo_lav = "L"

                If TextBox2.Text = Nothing Then
                    tempo_cambio_pezzo = 0
                Else
                    tempo_cambio_pezzo = TextBox2.Text
                End If

                If TextBox1.Text = Nothing Then
                    tempo_CICLO_pezzo = 0
                Else
                    tempo_CICLO_pezzo = TextBox1.Text
                End If
                time = tempo_CICLO_pezzo * TextBox11.Text + tempo_cambio_pezzo * TextBox11.Text
                percentuale = 5
                lavorazione = risorsa

                inserisci_lavorazione_a_sap_macchina(lavorazione)
                risorsa_old = lavorazione
                lavorazione = "R00500"

                'Form106.Owner = Me
                'Form106.Show()
                'Me.Visible = False
                id_padre = id_manodopera

                Panel7.Visible = False








            End If
        End If

    End Sub

    Sub pulisci_campi_manodopera()
        TextBox1.Text = Nothing
        TextBox2.Text = Nothing
        TextBox9.Text = Nothing
        ComboBox1.SelectedIndex = -1
        TextBox10.Text = Nothing
    End Sub
    Sub inserisci_lavorazione_a_sap_macchina(PAR_RISORSA As String)
        Trova_ID()
        numero_lavorazione_macchina()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn



        CMD_SAP.CommandText = "insert into manodopera (id, tipo_documento, docnum, Dipendente, risorsa, Data, start,Stop, consuntivo, tipologia_lavorazione, ordine_lavorazione)
                        values (" & id_manodopera & ",'ODP'," & docnum_odp & ",'" & Codicedip & "','" & PAR_RISORSA & "',getdate(),convert(varchar, getdate(), 108),convert(varchar, getdate(), 108)," & Replace(time / 5 * percentuale, ",", ".") & ",'" & tipo_lav & "','" & numero_lavorazione & "')"

        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()
    End Sub

    Sub inserisci_lavorazione_a_Galileo_macchina(PAR_fase As String, par_ordine_produzione As Integer, par_tempo As String)
        'Trova_ID()
        'numero_lavorazione_macchina()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "INSERT INTO [AS400].[S786FAD1].[TIR90VIS].[YPCMOV0F] 
(PROFYP, DT01YP, CDDTYP, DTOPYP,      
NROPYP, RIGAYP, DTMOYP, ORPRYP, CDFAYP,                 CAMOYP,     
TEMOYP) 
VALUES ('TIR40'
, CONVERT(int, CONVERT(varchar(8), GETDATE(), 112))
, '01' 
, CONVERT(int, CONVERT(varchar(8), GETDATE(), 112))
, 400
, 1 
, CONVERT(int, CONVERT(varchar(8), GETDATE(), 112))
,'" & par_ordine_produzione & "'
, '" & PAR_fase & "' 
, '01' -- causale tempo
, " & par_tempo & ")
"

        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()
    End Sub

    Sub inserisci_lavorazione_a_sap_manodopera()
        Trova_ID()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn



        CMD_SAP.CommandText = "insert into manodopera (id, tipo_documento, docnum, Dipendente, risorsa, Data, start,Stop, consuntivo, tipologia_lavorazione,risorsa_2,id_padre)
                        values (" & id_manodopera & ",'ODP'," & docnum_odp & ",'" & Codicedip & "','" & lavorazione & "',getdate(),convert(varchar, getdate(), 108),convert(varchar, getdate(), 108),round(case when " & Replace(time / 5 * percentuale, ",", ".") & "<1 then 1 else " & Replace(time / 5 * percentuale, ",", ".") & " end ,0),'" & tipo_lav & "','" & risorsa_old & "','" & id_padre & "')"

        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()


    End Sub

    Sub inserisci_record_imprevisto()
        Trova_ID_imprevisto()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn



        CMD_SAP.CommandText = "insert into [Tirelli_40].[dbo].[record_imprevisti_mu] (id,id_manodopera,codice_imprevisto,altro) 
values (" & id_imprevisto & "," & id_manodopera & "," & codice_imprevisto & ",'" & TextBox10.Text & "')"

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
        CMD_SAP_2.CommandText = "select max(id)+1 as 'ID'
from manodopera"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                id_manodopera = cmd_SAP_reader_2("ID")
            Else
                id_manodopera = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub

    Sub Trova_ID_imprevisto()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select max(id)+1 as 'ID' from [Tirelli_40].[dbo].[record_imprevisti_mu]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                id_imprevisto = cmd_SAP_reader_2("ID")
            Else
                id_imprevisto = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub

    Sub numero_lavorazione_macchina()
        If Homepage.ERP_provenienza = "SAP" Then


            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_2 As New SqlCommand
            Dim cmd_SAP_reader_2 As SqlDataReader


            CMD_SAP_2.Connection = Cnn1
            CMD_SAP_2.CommandText = "SELECT t10.id, t11.risorsa
, case when t11.ordine_lavorazione is null then 0 else t11.ordine_lavorazione end as 'ordine_lavorazione'
FROM
(
SELECT MAX (T0.ID) AS 'ID'
FROM MANODOPERA T0 LEFT JOIN ORSC T1 ON T0.RISORSA= T1.VISRESCODE  WHERE DOCNUM ='" & Button3.Text & "' AND T1.[ResType]='M'
)
AS T10
LEFT JOIN MANODOPERA T11 ON T10.ID=T11.ID"

            cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

            If cmd_SAP_reader_2.Read() = True Then

                If cmd_SAP_reader_2("risorsa") Is System.DBNull.Value Then
                    numero_lavorazione = 0
                Else
                    If cmd_SAP_reader_2("risorsa") = risorsa Then


                        numero_lavorazione = cmd_SAP_reader_2("ordine_lavorazione")

                    Else
                        numero_lavorazione = cmd_SAP_reader_2("ordine_lavorazione") + 1
                    End If
                End If


                cmd_SAP_reader_2.Close()
            End If
            Cnn1.Close()
        End If
    End Sub



    Private Sub Button20_Click(sender As Object, e As EventArgs)
        percentuale = 0

        inserisci_lavorazione_a_sap_manodopera()
        inserisci_percentuale()
        rECORD_CARICATI(DataGridView2)
        pulisci_campi_manodopera()



        Panel7.Visible = True

        If Button12.BackColor = Color.Red And Button12.Visible = True Then

            inizializzazione_autocontrollo()

        End If
        kpi_mensile_FASI()
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs)
        percentuale = 1
        inserisci_lavorazione_a_sap_manodopera()
        inserisci_percentuale()
        rECORD_CARICATI(DataGridView2)
        pulisci_campi_manodopera()


        Panel7.Visible = True
        If Button12.BackColor = Color.Red And Button12.Visible = True Then

            inizializzazione_autocontrollo()

        End If
        kpi_mensile_FASI()
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs)
        percentuale = 2
        inserisci_lavorazione_a_sap_manodopera()
        inserisci_percentuale()
        rECORD_CARICATI(DataGridView2)
        pulisci_campi_manodopera()


        Panel7.Visible = True
        If Button12.BackColor = Color.Red And Button12.Visible = True Then

            inizializzazione_autocontrollo()

        End If
        kpi_mensile_FASI()

    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs)
        percentuale = 3
        inserisci_lavorazione_a_sap_manodopera()
        inserisci_percentuale()
        rECORD_CARICATI(DataGridView2)
        pulisci_campi_manodopera()


        Panel7.Visible = True
        If Button12.BackColor = Color.Red And Button12.Visible = True Then

            inizializzazione_autocontrollo()

        End If

        kpi_mensile_FASI()
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs)
        percentuale = 4
        inserisci_lavorazione_a_sap_manodopera()
        inserisci_percentuale()
        rECORD_CARICATI(DataGridView2)
        pulisci_campi_manodopera()


        Panel7.Visible = True
        If Button12.BackColor = Color.Red And Button12.Visible = True Then

            inizializzazione_autocontrollo()

        End If

        kpi_mensile_FASI()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs)
        percentuale = 5
        inserisci_lavorazione_a_sap_manodopera()
        inserisci_percentuale()
        rECORD_CARICATI(DataGridView2)
        pulisci_campi_manodopera()


        Panel7.Visible = True
        If Button12.BackColor = Color.Red And Button12.Visible = True Then

            inizializzazione_autocontrollo()

        End If
        kpi_mensile_FASI()

    End Sub

    Sub inserisci_percentuale()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "UPDATE MANODOPERA Set percentuale_lavorazione=" & percentuale & " WHERE ID ='" & id_manodopera & "'"
        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()


    End Sub

    Sub inserisci_commento_odp(PAR_N_ODP As Integer, par_testo As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "UPDATE OWOR Set COMMENTS='" & par_testo & "'
WHERE DOCNUM ='" & PAR_N_ODP & "'"
        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()


    End Sub

    Sub check_autocontrollo_attrezzaggio()
        If Homepage.ERP_provenienza = "SAP" Then


            autocontrollo_attrezzaggio_necessario = "N"
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()


            Dim CMD_SAP_docentry As New SqlCommand
            Dim cmd_SAP_docentry_reader As SqlDataReader

            CMD_SAP_docentry.Connection = Cnn

            CMD_SAP_docentry.CommandText = "select *
from
(
select * from [TIRELLI_40].[DBO].autocontrollo_config
where tipo_controllo=" & tipo_autocontrollo & " and resgrpcod=" & tipo_macchina & " and itemcode is null

union all

select * from [TIRELLI_40].[DBO].autocontrollo_config
where tipo_controllo=" & tipo_autocontrollo & " and itemcode='" & Button2.Text & "'
)
as t10
order by id"


            cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader
            If cmd_SAP_docentry_reader.Read() Then

                autocontrollo_attrezzaggio_necessario = "Y"

            End If
            cmd_SAP_docentry_reader.Close()
            Cnn.Close()

            If autocontrollo_attrezzaggio_necessario = "Y" Then
                Button10.BackColor = Color.Red
                Button10.Visible = True
            Else
                Button10.Visible = False
                Button10.BackColor = Color.Lime
            End If
            CHECK_AUTOCONTROLLO_CARICATI()
        End If
    End Sub

    Sub check_autocontrollo_lavorazione()
        If Homepage.ERP_provenienza = "SAP" Then


            autocontrollo_lavorazione_necessario = "N"
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()


            Dim CMD_SAP_docentry As New SqlCommand
            Dim cmd_SAP_docentry_reader As SqlDataReader

            CMD_SAP_docentry.Connection = Cnn

            CMD_SAP_docentry.CommandText = "select *
from
(
select * from [TIRELLI_40].[DBO].autocontrollo_config
where tipo_controllo=" & tipo_autocontrollo & " and resgrpcod=" & tipo_macchina & " and itemcode is null

union all

select * from [TIRELLI_40].[DBO].autocontrollo_config
where tipo_controllo=" & tipo_autocontrollo & " and itemcode='" & Button2.Text & "'
)
as t10
order by id"


            cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader
            If cmd_SAP_docentry_reader.Read() Then

                autocontrollo_lavorazione_necessario = "Y"

            End If
            cmd_SAP_docentry_reader.Close()
            Cnn.Close()

            If autocontrollo_attrezzaggio_necessario = "Y" Then
                Button12.BackColor = Color.Red
                Button12.Visible = True
            Else
                Button12.Visible = False
                Button12.BackColor = Color.Lime
            End If
        Else

        End If
    End Sub

    Sub inizializzazione_autocontrollo()

        Autocontrollo.Label1.Text = Combodipendenti.Text
        Autocontrollo.Label2.Text = Combomacchinari.Text
        Autocontrollo.Codicedip = Codicedip
        Autocontrollo.risorsa = risorsa
        Autocontrollo.Label_ordine_SAP.Text = Button3.Text
        Autocontrollo.LabelDescrizioneSAP.Text = Label24.Text
        Autocontrollo.Button_disegno.Text = Button4.Text
        Autocontrollo.LabelQuantitàSAP.Text = LabelQuantitàSAP.Text

        Autocontrollo.tipo_macchina = tipo_macchina
        Autocontrollo.LabelCodiceSAP.Text = Button2.Text
        Autocontrollo.tipo_autocontrollo = tipo_autocontrollo
        Autocontrollo.tipo_lav = tipo_lav
        Autocontrollo.carica_checklist_autocontrollo()
        Autocontrollo.Show()

    End Sub


    Sub CHECK_AUTOCONTROLLO_CARICATI()



        Dim AUTOCONTROLLO_1 As String = "NO"
        Dim AUTOCONTROLLO_2 As String = "NO"

        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT TIPO_AUTOCONTROLLO
FROM [TIRELLI_40].[DBO].autocontrollo WHERE DOCNUM='" & docnum & "' and itemcode= '" & risorsa & "'
        GROUP BY TIPO_AUTOCONTROLLO"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader




        Do While cmd_SAP_reader_2.Read()


            If cmd_SAP_reader_2("TIPO_AUTOCONTROLLO") = 1 Then
                AUTOCONTROLLO_1 = "Y"
            End If
            If cmd_SAP_reader_2("TIPO_AUTOCONTROLLO") = 2 Then
                AUTOCONTROLLO_2 = "Y"
            End If

        Loop

        If AUTOCONTROLLO_1 = "Y" Then

            Button10.BackColor = Color.Lime

        Else

            Button10.BackColor = Color.Red
        End If



        If AUTOCONTROLLO_2 = "Y" Then
            Button12.BackColor = Color.Lime
        Else
            Button12.BackColor = Color.Red
        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If Button10.BackColor = Color.Lime Then

            Autocontrollo.tipo_autocontrollo = 1
            Autocontrollo.risorsa = risorsa
            inizializzazione_autocontrollo()

            Autocontrollo.carica_checklist_autocontrollo_completata()

            Autocontrollo.Show()


        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click



        ODP_Form.docnum_odp = Button3.Text
        ODP_Form.Show()
        ODP_Form.inizializza_form(Button3.Text)


    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Process.Start(Homepage.percorso_disegni_generico & "PDF\" & Button4.Text & ".PDF")
        Catch ex As Exception
            MsgBox("Il disegno " & Button4.Text & " non è ancora stato processato")
        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = " update owor set PLANNEDQTY= " & TextBox3.Text & " where docnum=" & Button3.Text & ""
        Cmd_SAP.ExecuteNonQuery()


        Cmd_SAP.CommandText = " update t41 set t41.onorder=t40.ORDINATI
from
(
SELECT t10.itemcode,sum(t10.Ordinati) AS 'ORDINATI', t10.MAG
FROM
(
SELECT T0.ITEMCODE, 0 AS 'ORDINATI', T0.WHSCODE AS 'MAG'
FROM OITW T0
WHERE T0.ONORDER>0 OR T0.ONORDER<0
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[PlannedQty]) AS 'CONFERMATI', T0.[wareHouse] AS 'mag'
FROM OWOR T0   WHERE (T0.[STATUS] ='P' OR  T0.[STATUS] ='R')
GROUP BY T0.[ItemCode],T0.[wareHouse]

UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.[WhsCode] 
FROM POR1 T0  INNER JOIN OPOR T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T1.[DocStatus] ='O' AND T0.[OpenQty]>0
GROUP BY 
T0.[ItemCode],  T0.[WhsCode]
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.WHSCODE
FROM WTQ1 T0  INNER JOIN OWTQ T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T0.[OpenQty] >0 AND  T1.[DocStatus] ='O'
GROUP BY 
T0.[ItemCode], T0.WHSCODE
)
AS T10
group by t10.itemcode, t10.MAG
)
as t40 inner join oitw t41 on t41.itemcode=t40.itemcode and t41.whscode=t40.MAG"


        Cmd_SAP.ExecuteNonQuery()

        Cmd_SAP.CommandText = " update t41 set t41.iscommited=t40.confermati
from
(

SELECT t10.itemcode,sum(t10.confermati) AS 'CONFERMATI', t10.MAG
FROM
(
SELECT T0.ITEMCODE, 0 AS 'CONFERMATI', T0.WHSCODE AS 'MAG'
FROM OITW T0
WHERE T0.ISCOMMITED>0 OR T0.ISCOMMITED<0
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[PlannedQty]) AS 'CONFERMATI', T0.[wareHouse] AS 'mag'
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE (T1.[STATUS] ='P' OR  T1.[STATUS] ='R') AND T1.[CmpltQty]< T1.[PlannedQty] 
GROUP BY T0.[ItemCode],T0.[wareHouse]

UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.[WhsCode] 
FROM RDR1 T0  INNER JOIN ORDR T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T1.[DocStatus] ='O' AND T0.[OpenQty]>0
GROUP BY 
T0.[ItemCode],  T0.[WhsCode]
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.FROMWHSCOD 
FROM WTQ1 T0  INNER JOIN OWTQ T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T0.[OpenQty] >0 AND  T1.[DocStatus] ='O'
GROUP BY 
T0.[ItemCode], T0.FROMWHSCOD
)
AS T10
group by t10.itemcode, t10.MAG

)
as t40 inner join oitw t41 on t41.itemcode=t40.itemcode and t41.whscode=t40.mag"


        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

        invia_Odp(docnum)

        MsgBox("Quantità cambiata con successo")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Magazzino.Codice_SAP = Button2.Text
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
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        CQ_Tabelle.Show()

        CQ_Tabelle.Inserimento_dipendenti()


        CQ_Tabelle.TabControl1.SelectTab(1)
        CQ_Tabelle.ComboBox3.SelectedIndex = 5
        CQ_Tabelle.codicedip = Codicedip
        CQ_Tabelle.ComboBox1.Text = Combodipendenti.Text
        CQ_Tabelle.ComboBox2.SelectedIndex = 2
        CQ_Tabelle.riempi_registrazioni_controlli()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs)
        chiudi_lavorazione()
        righe_ODP_macchine(DataGridView1, docnum)
        materia_prima_macchina(DataGridView6, docnum)


        'If Homepage.Centro_di_costo = "BRB01" Then
        '    MsgBox("Fase conclusa con successo")
        '    blocco_tab = 0
        '    Inserimento_dipendenti()
        '    Inserimento_risorse()
        '    Inserimento_causali_imprevisti()
        '    compila_datagridvied_lista_odp_bint()
        '    TabControl1.SelectedIndex = 3
        '    blocco_tab = 1
        'Else

        'If tipo_macchina = 7 Then
        '        Dim par_tipo_appoggio As String = "ODP_TAGLIO"
        '        DataGridView_ODP_TIPO_MACCHINA.Rows.Clear()
        '        ODP_Tree.PULISCI_APPOGGIO(Homepage.ID_SALVATO, par_tipo_appoggio)
        '        lISTA_odp_taglio(DataGridView_ODP_TIPO_MACCHINA, par_tipo_appoggio)
        '    Else
        '        lISTA_odp_per_tipo_macchina(DataGridView_ODP_TIPO_MACCHINA, TextBox8.Text, TextBox7.Text, TextBox5.Text, TextBox4.Text)
        '    End If


        l_ordine_completabile()
        Panel19.Visible = False
        kpi_mensile_FASI()

        If ordine_completabile = "SI" Then
            completa_ordine()
            ELIMINA_risorse_IN_ODP_COMPLETATI(" AND T0.DOCNUM ='" & docnum & "' ")
            carica_lavorazioni_in_ODP(" AND T0.DOCNUM ='" & docnum & "' ")

            MsgBox("Manodopera caricata con successo")
            '  MsgBox("ORDINE " & docnum & " Completato con successo, portare l'ordine finito nello scomparto " & vbCrLf & vbCrLf & Label29.Text)
        Else
            MsgBox("FASE conclusa con successo. portare la merce alla fase successiva e selezionare il prossimo ordine di produzione ")
        End If
        '   End If


        'TabPage6.Hide()





    End Sub

    Sub chiudi_lavorazione()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = " UPDATE T21 SET T21.[U_Stato_lavorazione] = 'C'
FROM
(
Select T13.DOCENTRY AS 'DOCENTRY', T11.VISORDER AS 'VISORDER', t13.docnum as 'ODP', T13.[PlannedQty] as 'N PEZ' , t13.itemcode as 'COD', t14.itemname as 'Nome', t14.u_disegno as 'DIS', t11.itemcode as 'RIS', t12.itemname as 'MAC', T11.[AdditQty] as 'ATT', T11.[PlannedQty] as 'LAV', T13.[U_Priorita_MES] AS 'Priorità'
from
(
SELECT t0.docentry, min(T0.[VisOrder]) as 'Visorder' , t0.u_stato_lavorazione
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry]
inner join ORSC T2 on t2.visrescode=t0.itemcode WHERE T1.[Status] ='R' and  T2.[ResType] ='M' and  T0.[U_Stato_lavorazione] = 'O'
group by 
t0.docentry, t0.u_stato_lavorazione
)
as t10
inner join wor1 t11 on t11.docentry=t10.docentry and t10.visorder=t11.visorder
inner join oitm t12 on t12.itemcode=t11.itemcode
inner join owor t13 on t13.docentry=t11.docentry
inner join oitm t14 on t14.itemcode=t13.itemcode
inner join orsc t15 on t15.visrescode=t11.itemcode

WHERE t15.resgrpcod=" & tipo_macchina & " AND t13.docnum= " & Button3.Text & "

) 
AS T20
INNER JOIN WOR1 T21 ON T21.DOCENTRY=T20.DOCENTRY AND T21.VISORDER=T20.VISORDER "
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()
    End Sub

    Sub l_ordine_completabile()
        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP_docentry As New SqlCommand
        Dim cmd_SAP_docentry_reader As SqlDataReader

        CMD_SAP_docentry.Connection = Cnn
        CMD_SAP_docentry.CommandText = "SELECT t0.u_stato_lavorazione FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry]
 inner join ORSC T2 on t0.itemcode=t2.visrescode WHERE T1.[DocNum] ='" & Button3.Text & "' and  T2.[ResType] ='M' and t0.u_stato_lavorazione='O'"
        cmd_SAP_docentry_reader = CMD_SAP_docentry.ExecuteReader

        If cmd_SAP_docentry_reader.Read = True Then

            ordine_completabile = "NO"

        Else
            ordine_completabile = "SI"

        End If

        cmd_SAP_docentry_reader.Close()
        Cnn.Close()
    End Sub

    Sub completa_ordine()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand


        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE OWOR SET OWOR.U_stato='Completato' where owor.docnum = " & Button3.Text & ""
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()
    End Sub



    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        'Dim ID_LAVORAZIONE As Integer
        'If e.RowIndex >= 0 Then

        '    If e.ColumnIndex = DataGridView2.Columns.IndexOf(X) Then

        '        ID_LAVORAZIONE = DataGridView2.Rows(e.RowIndex).Cells(columnName:="ID").Value
        '        Dim Cnn As New SqlConnection
        '        Cnn.ConnectionString = Homepage.sap_tirelli

        '        Cnn.Open()

        '        Dim Cmd_SAP As New SqlCommand


        '        Cmd_SAP.Connection = Cnn
        '        Cmd_SAP.CommandText = " DELETE MANODOPERA WHERE ID=" & ID_LAVORAZIONE & ""
        '        Cmd_SAP.ExecuteNonQuery()


        '        Cmd_SAP.Connection = Cnn
        '        Cmd_SAP.CommandText = " DELETE MANODOPERA WHERE ID_PADRE=" & ID_LAVORAZIONE & ""
        '        Cmd_SAP.ExecuteNonQuery()

        '        Cmd_SAP.Connection = Cnn
        '        Cmd_SAP.CommandText = " DELETE [Tirelli_40].[dbo].[record_imprevisti_mu] WHERE id_manodopera=" & ID_LAVORAZIONE & ""
        '        Cmd_SAP.ExecuteNonQuery()

        '        Cnn.Close()
        '        rECORD_CARICATI()
        '    End If
        'End If
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        If Combomacchinari.Text <> "" Then
            risorsa = Elenco_macchinari(Combomacchinari.SelectedIndex)
            tipo_macchina = Elenco_tipo_macchina(Combomacchinari.SelectedIndex)



            lISTA_odp_per_tipo_macchina(DataGridView_ODP_TIPO_MACCHINA, TextBox8.Text, TextBox7.Text, TextBox5.Text, TextBox4.Text, TextBox17.Text)
            Presenza_attrezzaggio_lavorazione()


        Else
            MsgBox("Inserire macchinario")
        End If

        visibilità()
    End Sub



    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox11.KeyPress
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
        If TextBox11.TextLength = 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If

        ' ovviamente se c'è già un punto
        ' non è consentito digitarne altri
        If (KeyAscii = 46) And
            TextBox11.Text.IndexOf(".") > 0 Then
            KeyAscii = 0
        End If

        ' Reimposto il keychar
        e.KeyChar = Chr(KeyAscii)





    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If GroupBox10.Visible = True Then




            If TextBox2.Text = Nothing Then
                tempo_cambio_pezzo = 0
            Else
                tempo_cambio_pezzo = TextBox2.Text
            End If

            If TextBox1.Text = Nothing Then
                tempo_CICLO_pezzo = 0
            Else
                tempo_CICLO_pezzo = TextBox1.Text
            End If


            time = tempo_CICLO_pezzo * TextBox11.Text + tempo_cambio_pezzo * TextBox11.Text
            Label27.Text = time
        End If

    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged

        If TextBox1.Text > LabelQuantitàSAP.Text Then

            MsgBox("Non è possibile inserire una quantità maggiore di quella indicata dall'ordine di produzione")
        Else
            If GroupBox10.Visible = True Then
                If TextBox2.Text = Nothing Then
                    tempo_cambio_pezzo = 0
                Else
                    tempo_cambio_pezzo = TextBox2.Text
                End If

                If TextBox1.Text = Nothing Then
                    tempo_CICLO_pezzo = 0
                Else
                    tempo_CICLO_pezzo = TextBox1.Text
                End If

                time = tempo_CICLO_pezzo * TextBox11.Text + tempo_cambio_pezzo * TextBox11.Text
                Label27.Text = time
            End If
        End If
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        lISTA_tutti_gli_odp()
    End Sub

    Sub ELIMINA_risorse_IN_ODP_COMPLETATI(PAR_FILTRO_DOCNUM As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        CMD_SAP_2.Connection = Cnn


        CMD_SAP_2.CommandText = "Select t10.docentry 
from
(
SELECT   T3.DOCENTRY AS 'Docentry', t0.docnum as 'ODP', t3.plannedqty as 'Quantità', t2.u_ordine as 'Ordine', t0.risorsa as 'Risorsa', t2.resname as 'Nome risorsa', t0.tipologia_lavorazione as 'Lavorazione', sum(CASE when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end) as 'Minuti'
FROM MANODOPERA t0 inner join [TIRELLI_40].[dbo].ohem t1 on t1.[empID]=t0.dipendente
inner join orsc t2 on t2.visrescode=t0.risorsa
LEFT JOIN OWOR T3 ON T3.DOCNUM=t0.docnum
left join oitm t4 on t4.itemcode=t3.itemcode
where  t0.tipo_documento='ODP' AND T3.CMPLTQTY<T3.PLANNEDQTY AND ((substring(t3.U_produzione,1,3)='INT' AND  t3.u_stato='Completato') or substring(t3.U_produzione,1,5)='B_INT') and (T3.STAtUS ='R' OR substring(t3.U_produzione,1,5)='B_INT') 
" & PAR_FILTRO_DOCNUM & "
group by 
t0.docnum,T3.DOCENTRY,t0.risorsa,t2.resname, t0.tipologia_lavorazione, t3.plannedqty, t2.u_ordine
)
as t10
group by 
t10.docentry 
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()


            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli

            Cnn1.Open()


            Dim Cmd_SAP_1 As New SqlCommand


            'Inserisco i valori nell'odp
            Cmd_SAP_1.Connection = Cnn1
            Cmd_SAP_1.CommandText = "DELETE WOR1 WHERE DOCENTRY= " & cmd_SAP_reader_2("Docentry") & " AND ITEMTYPE=290 "
            Cmd_SAP_1.ExecuteNonQuery()


            Cnn1.Close()


        Loop

        cmd_SAP_reader_2.Close()
        Cnn.Close()

    End Sub

    Sub carica_lavorazioni_in_ODP(PAR_FILTRO_DOCNUM As String)
        Dim prezzo As String
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        CMD_SAP_2.Connection = Cnn


        CMD_SAP_2.CommandText = "select t20.duedate,t20.docentry,t20.odp,t20.ordine,t20.risorsa,t20.[nome risorsa], t20.quantità ,sum(t20.attrezzaggio) as 'Attrezzaggio' ,sum(t20.lavorazione) as 'Lavorazione', min(t20.ordine_lavorazione),t20.price
from
(
Select t10.duedate,  t10.docentry as 'Docentry', t10.ODP as 'ODP',t10.ordine, t10.Risorsa as 'Risorsa',  t10.[Nome risorsa] as 'Nome risorsa', t10.quantità as 'Quantità', sum(case when t10.lavorazione = 'A' then t10.Minuti else 0 end) as 'Attrezzaggio', case when sum(case when t10.lavorazione = 'L' then t10.Minuti else 0 end)=0 then 1 else sum(case when t10.lavorazione = 'L' then t10.Minuti else 0 end) end as 'Lavorazione', t10.ordine_lavorazione, t11.price
from
(
SELECT  T3.DUEDATE, T3.DOCENTRY AS 'Docentry',  t0.docnum as 'ODP', t3.plannedqty as 'Quantità', t2.u_ordine as 'Ordine', t0.risorsa as 'Risorsa', t2.resname as 'Nome risorsa', t0.tipologia_lavorazione as 'Lavorazione', sum(CASE when t0.start is null or t0.stop is null then t0.consuntivo else case when DATEPART(hour, t0.start)<12 and ((DATEPART(hour, t0.stop)>=13 and DATEPART(minute, t0.stop)>30) or datepart(hour,t0.stop)>=14) then (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo -90 else  (DATEPART(hour, t0.stop)*60+DATEPART(minute, t0.stop))-(DATEPART(hour, t0.start)*60+DATEPART(minute, t0.start))+t0.consuntivo end end) as 'Minuti', t0.ordine_lavorazione
FROM MANODOPERA t0 inner join [TIRELLI_40].[dbo].ohem t1 on t1.[empID]=t0.dipendente
inner join orsc t2 on t2.visrescode=t0.risorsa
LEFT JOIN OWOR T3 ON T3.DOCNUM=t0.docnum
left join oitm t4 on t4.itemcode=t3.itemcode
where  t0.tipo_documento='ODP' and T3.CMPLTQTY<T3.PLANNEDQTY   AND ((substring(t3.U_produzione,1,3)='INT' AND  t3.u_stato='Completato') or substring(t3.U_produzione,1,5)='B_INT')  and (t3.status='P' or t3.status='R')  " & PAR_FILTRO_DOCNUM & "
group by 
t3.duedate, t0.docnum,T3.DOCENTRY,t0.risorsa,t2.resname, t0.tipologia_lavorazione, t3.plannedqty, t2.u_ordine, t0.ordine_lavorazione
)
as t10 INNER JOIN itm1 t11 on t11.itemcode=t10.risorsa

where t11.pricelist='2'

group by 
t10.duedate, t10.docentry , t10.ODP , t10.Risorsa , t10.[Nome risorsa],t10.quantità, T10.ORDINE,t10.ordine_lavorazione,t11.price

) 
as t20
group by 
t20.duedate,t20.docentry,t20.odp,t20.ordine,t20.risorsa,t20.[nome risorsa], t20.quantità, t20.price
order by t20.docentry, min(t20.ordine_lavorazione)"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            Dim CNN3 As New SqlConnection
            CNN3.ConnectionString = Homepage.sap_tirelli
            CNN3.Open()

            Dim Cmd_SAP As New SqlCommand

            Dim Num_Riga As Integer
            ' Individuo la riga in cui scrivere
            Dim CMD_SAP_max_LINE As New SqlCommand
            Dim cmd_SAP_max_LINE_reader As SqlDataReader

            CMD_SAP_max_LINE.Connection = CNN3
            CMD_SAP_max_LINE.CommandText = "SELECT 
case when max(T0.[LineNum]) is null then 0 else max(T0.[LineNum]) end as 'max'
, case when max(T0.[VISORDER]) is null then 0 else max(T0.[VISORDER]) end  as 'max_VIS'
FROM WOR1 T0 inner join owor t1 on t0.docentry=t1.docentry WHERE T1.docentry=" & cmd_SAP_reader_2("Docentry") & ""
            cmd_SAP_max_LINE_reader = CMD_SAP_max_LINE.ExecuteReader

            If cmd_SAP_max_LINE_reader.Read = True Then
                Num_Riga = Val((cmd_SAP_max_LINE_reader("max"))) + 1
                Num_vis_riga = Val((cmd_SAP_max_LINE_reader("max_VIS"))) + 1
            Else
                Num_Riga = 0
            End If
            cmd_SAP_max_LINE_reader.Close()

            CNN3.Close()

            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli

            Cnn1.Open()


            Dim Cmd_SAP_1 As New SqlCommand

            prezzo = Replace(cmd_SAP_reader_2("price"), ",", ".")

            Cmd_SAP_1.Connection = Cnn1
            Cmd_SAP_1.CommandText = "insert into WOR1(WOR1.DOCENTRY, WOR1.LINENUM,WOR1.ITEMCODE,WOR1.ITEMNAME,WOR1.VISORDER
,WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType, wor1.[ResAlloc], wor1.[StartDate], wor1.[EndDate],wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden, WOR1.U_STATO_lavorazione,wor1.u_prezzolis,wor1.u_DIPENDENTE) 
        VALUES (" & cmd_SAP_reader_2("Docentry") & ", '" & Num_Riga & "','" & cmd_SAP_reader_2("Risorsa") & "', '" & cmd_SAP_reader_2("Nome risorsa") & "'," & Num_vis_riga & "
        ,CASE WHEN round(" & cmd_SAP_reader_2("Lavorazione") & "/" & cmd_SAP_reader_2("Quantità") & ",0) < 1 THEN 1 ELSE round(" & cmd_SAP_reader_2("Lavorazione") & "/" & cmd_SAP_reader_2("Quantità") & ",0) END ,'" & cmd_SAP_reader_2("Attrezzaggio") & "'," & cmd_SAP_reader_2("Attrezzaggio") & "+" & cmd_SAP_reader_2("Lavorazione") & ",'RIS',290,'m','F', CONVERT(DATETIME, '" & cmd_SAP_reader_2("duedate") & "',103), CONVERT(DATETIME, '" & cmd_SAP_reader_2("duedate") & "',103),0,0,0,0,0,'C'," & prezzo & ",1)"
            Cmd_SAP_1.ExecuteNonQuery()


            Cnn1.Close()


        Loop

        cmd_SAP_reader_2.Close()
        Cnn.Close()



    End Sub




    Private Sub Timer4_0_Tick(sender As Object, e As EventArgs) Handles Timer4_0.Tick

        DateTimePicker2.Value = Now

        'test(DateTimePicker1, DateTimePicker2)
        Percentuale_Lavorazione(DateTimePicker1, DateTimePicker2)
        Macchine_In_Movimento()

    End Sub

    Sub test(par_datetimepicker1 As DateTimePicker, par_datetimepicker2 As DateTimePicker)
        Dim test As String
        test = "DECLARE @DataInizio AS DATETIME;
DECLARE @DataInizio_00 AS DATETIME;
DECLARE @DataFine AS DATETIME;
DECLARE @DataFine_59 AS DATETIME;

-- Imposta @DataInizio e @DataFine con la data specificata
SET @DataInizio = CONVERT(DATETIME, '" & par_datetimepicker1.Value & " ', 120);


SET @DataFine = CONVERT(DATETIME, '" & par_datetimepicker2.Value & " ', 120);


-- Imposta @DataInizio_00 alla mezzanotte del giorno specificato
SET @DataInizio_00 = CAST(CONVERT(DATE,'" & par_datetimepicker1.Value & " ', 103) AS DATETIME);

-- Imposta @DataFine_59 all'ultimo secondo del giorno specificato
SET @DataFine_59 = DATEADD(SECOND, -1, DATEADD(DAY, DATEDIFF(DAY, 0, '" & par_datetimepicker2.Value & " ') + 1, 0));


select @datainizio,@DataFine
select t20.MachineId,t20.Lavorazione,t20.fermo, DATEDIFF(second,@Datafine,@DataInizio_00)-t20.Lavorazione-t20.fermo as 'Spento'
from
(
select t10.MachineId,sum(t10.Lavorazione) as 'Lavorazione',sum(t10.fermo) as 'Fermo'
from
(

SELECT t0.machineid,t0.inizio, t0.fine,
CASE 
    WHEN t0.ActivityId = 11 AND t0.fine IS NULL THEN DATEDIFF(second, t0.inizio, GETDATE())
WHEN t0.ActivityId = 11 AND fine IS NOT NULL THEN durata/1000
else 0 end as 'Lavorazione',
 case 
 WHEN t0.ActivityId = 10 AND t0.fine IS NULL THEN DATEDIFF(second, t0.inizio, GETDATE())
when t0.ActivityId = 10 AND fine IS NOT NULL THEN durata/1000
else 0 end as 'Fermo'
      
FROM [dbo].[Transazioni] t0
LEFT JOIN [dbo].[Activities] t1 ON t0.activityid = t1.ActivityId
WHERE CONVERT(DATE,t0.inizio,103) >= @DataInizio_00 AND (CONVERT(DATE,t0.fine,103) <= @DataFine_59 OR t0.fine IS NULL) 
)
as t10
group by t10.MachineId
)
as t20

"
        Console.WriteLine(test)
    End Sub


    Sub Percentuale_Lavorazione(par_datetimepicker1 As DateTimePicker, par_datetimepicker2 As DateTimePicker)
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Pianificazione.DATABASE_MU_4_0
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        ' Imposta la connessione per il comando
        CMD_SAP_2.Connection = Cnn1

        ' Modifica della query per utilizzare i valori di par_datetimepicker1 e par_datetimepicker2
        CMD_SAP_2.CommandText = "DECLARE @DataInizio AS DATETIME;
DECLARE @DataInizio_00 AS DATETIME;
DECLARE @DataFine AS DATETIME;
DECLARE @DataFine_59 AS DATETIME;

-- Imposta @DataInizio e @DataFine con la data specificata
SET @DataInizio = CONVERT(DATETIME, '" & par_datetimepicker1.Value & " ', 103);


SET @DataFine = CONVERT(DATETIME, '" & par_datetimepicker2.Value & " ', 103);


-- Imposta @DataInizio_00 alla mezzanotte del giorno specificato
SET @DataInizio_00 = CAST(CONVERT(DATE,'" & par_datetimepicker1.Value & " ', 103) AS DATETIME);

-- Imposta @DataFine_59 all'ultimo secondo del giorno specificato
SET @DataFine_59 = DATEADD(SECOND, -1, DATEADD(DAY, DATEDIFF(DAY, 0, '" & par_datetimepicker2.Value & " ') + 1, 0));



select t20.MachineId,t20.Lavorazione,t20.fermo, -DATEDIFF(second,@Datafine,@DataInizio_00)-t20.Lavorazione-t20.fermo as 'Spento'
from
(
select t10.MachineId,sum(t10.Lavorazione) as 'Lavorazione',sum(t10.fermo) as 'Fermo'
from
(

SELECT t0.machineid,t0.inizio, t0.fine,
CASE 
    WHEN t0.ActivityId = 11 AND t0.fine IS NULL THEN DATEDIFF(second, t0.inizio, GETDATE())
WHEN t0.ActivityId = 11 AND fine IS NOT NULL THEN durata/1000
else 0 end as 'Lavorazione',
 case 
 WHEN t0.ActivityId = 10 AND t0.fine IS NULL THEN DATEDIFF(second, t0.inizio, GETDATE())
when t0.ActivityId = 10 AND fine IS NOT NULL THEN durata/1000
else 0 end as 'Fermo'
      
FROM [dbo].[Transazioni] t0
LEFT JOIN [dbo].[Activities] t1 ON t0.activityid = t1.ActivityId
WHERE CONVERT(DATE,t0.inizio,103) >= @DataInizio_00 AND (CONVERT(DATE,t0.fine,103) <= @DataFine_59 OR t0.fine IS NULL) 
)
as t10
group by t10.MachineId
)
as t20"



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            'MsgBox(cmd_SAP_reader_2("machineid"))
            Dim machineId As Integer = CInt(cmd_SAP_reader_2("machineid"))

            If machineId = 1 Then

                AggiornaLabelTempo(Label11, Label13, Label15, cmd_SAP_reader_2)
                AggiornaLabelPercentuali(Label6, Label8, Label9, cmd_SAP_reader_2)
            ElseIf machineId = 2 Then

                AggiornaLabelTempo(Label19, Label17, Label7, cmd_SAP_reader_2)
                AggiornaLabelPercentuali(Label21, Label23, Label22, cmd_SAP_reader_2)
            End If
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
    End Sub



    Sub AggiornaLabelTempo(labelLavorazione As Label, labelFermo As Label, labelSpento As Label, reader As SqlDataReader)
        AggiornaLabelTempoSingolo(labelLavorazione, reader("Lavorazione"))
        AggiornaLabelTempoSingolo(labelFermo, reader("Fermo"))
        AggiornaLabelTempoSingolo(labelSpento, reader("Spento"))
    End Sub

    Sub AggiornaLabelTempoSingolo(label As Label, millisecondi As Object)
        Dim iSecond As Double = CDbl(millisecondi)
        Dim iSpan As TimeSpan = TimeSpan.FromSeconds(iSecond)
        label.Text = iSecond
        label.Text = $"{iSpan.Days}d {iSpan.Hours.ToString().PadLeft(2, "0"c)}:{iSpan.Minutes.ToString().PadLeft(2, "0"c)}:{iSpan.Seconds.ToString().PadLeft(2, "0"c)}"

    End Sub

    Sub AggiornaLabelPercentuali(labelLavorazione As Label, labelFermo As Label, labelSpento As Label, reader As SqlDataReader)
        Dim totale As Double = reader("Lavorazione") + reader("Fermo") + reader("Spento")
        labelLavorazione.Text = FormatPercent(reader("Lavorazione") / totale)
        labelFermo.Text = FormatPercent(reader("Fermo") / totale)
        labelSpento.Text = FormatPercent(reader("Spento") / totale)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            log_4_0()
        Catch ex As Exception

        End Try
        codice_imprevisto = ComboBox1.SelectedIndex + 1

    End Sub

    Private Sub TabPage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Enter
        Try
            log_4_0()
        Catch ex As Exception

        End Try

    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.SelectedIndex >= 0 Then
            log_4_0()

        End If
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub DataGridView3_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView3.CellFormatting
        If DataGridView3.Rows(e.RowIndex).Cells(columnName:="id_macchina").Value = 1 Then



            DataGridView3.Rows(e.RowIndex).Cells(columnName:="machinename").Style.BackColor = Color.Aqua
        Else
            DataGridView3.Rows(e.RowIndex).Cells(columnName:="machinename").Style.BackColor = Color.Beige
        End If


    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Dim partprogram As String = DataGridView3.Rows(e.RowIndex).Cells(columnName:="partprogram").Value
        Label28.Text = DataGridView3.Rows(e.RowIndex).Cells(columnName:="partprogram").Value

        trova_tempi_4_0_lavorazione(partprogram)
        trova_tempi_4_0_attrezzaggio(partprogram)


    End Sub

    Sub trova_tempi_4_0_lavorazione(par_partprogram As Integer)


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Pianificazione.DATABASE_MU_4_0
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT 
t0.[PartProgramNumber]
, t0.activityid
, t1.name
 ,sum(t0.[DurataTotale])/60000 as 'Minuti'


                            From [dbo].[Transazioni] t0 
                            LEFT JOIN [dbo].[Activities] t1 ON t0.activityid = t1.ActivityId 
                           where  t0.[PartProgramNumber]='" & par_partprogram & "' and  t0.activityid ='11'
                         

							group by t0.[PartProgramNumber]
, t0.activityid
, t1.name"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then


            TextBox13.Text = cmd_SAP_reader_2("minuti")
        Else
            TextBox13.Text = 0

        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    Sub trova_tempi_4_0_attrezzaggio(par_partprogram As Integer)


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Pianificazione.DATABASE_MU_4_0
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT 
t0.[PartProgramNumber]
, t0.activityid
, t1.name
 ,sum(t0.[DurataTotale])/60000 as 'Minuti'


                            From [dbo].[Transazioni] t0 
                            LEFT JOIN [dbo].[Activities] t1 ON t0.activityid = t1.ActivityId 
                           where  t0.[PartProgramNumber]='" & par_partprogram & "' and  t0.activityid ='10'
                         

							group by t0.[PartProgramNumber]
, t0.activityid
, t1.name"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then


            TextBox12.Text = cmd_SAP_reader_2("minuti")
        Else
            TextBox12.Text = 0

        End If

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    Sub inserisci_lavorazione_a_sap_4_0()
        Trova_ID()

        numero_lavorazione_macchina()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn


        tipo_lav = "A"
        CMD_SAP.CommandText = "insert into manodopera (id, tipo_documento, docnum, Dipendente, risorsa, Data, start,Stop, consuntivo, tipologia_lavorazione, ordine_lavorazione)
                        values (" & id_manodopera & ",'ODP'," & docnum & ",'" & Codicedip & "','" & risorsa & "',getdate(),convert(varchar, getdate(), 108),convert(varchar, getdate(), 108),'" & TextBox12.Text & "','" & tipo_lav & "','" & numero_lavorazione & "')"
        CMD_SAP.ExecuteNonQuery()
        tipo_lav = "L"
        CMD_SAP.CommandText = "insert into manodopera (id, tipo_documento, docnum, Dipendente, risorsa, Data, start,Stop, consuntivo, tipologia_lavorazione, ordine_lavorazione)
                        values (" & id_manodopera + 1 & ",'ODP'," & docnum & ",'" & Codicedip & "','" & risorsa & "',getdate(),convert(varchar, getdate(), 108),convert(varchar, getdate(), 108),'" & TextBox13.Text & "','" & tipo_lav & "','" & numero_lavorazione & "')"

        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        If docnum = "" Or docnum Is Nothing Then
            MsgBox("Scegliere ordine di produzione")

        Else


            If Codicedip = 0 Then
                MsgBox("Selezionare un operatore")

            Else
                inserisci_lavorazione_a_sap_4_0()
                MsgBox("Lavorazione caricata con successo" & vbCrLf & "Attrezzaggio : " & TextBox12.Text & vbCrLf & "Lavorazione:" & TextBox13.Text & vbCrLf & "ODP: " & docnum)
            End If

        End If

    End Sub

    Private Sub Dashboard_MU_New_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs)
        compila_datagridvied_lista_odp_bint()
    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs)
        compila_datagridvied_lista_odp_bint()
    End Sub



    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs)
        compila_datagridvied_lista_odp_bint()
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs)
        compila_datagridvied_lista_odp_bint()
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs)
        compila_datagridvied_lista_odp_bint()
    End Sub

    Private Sub DataGridView5_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridView5_CellClick(sender As Object, e As DataGridViewCellEventArgs)




    End Sub

    Private Sub DataGridView5_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        Dim divValue As String = DataGridView5.Rows(e.RowIndex).Cells("Status").Value.ToString()
        Select Case divValue
            Case "Pianificato"
                DataGridView5.Rows(e.RowIndex).Cells("Status").Style.BackColor = Color.Yellow
            Case "Completato"
                DataGridView5.Rows(e.RowIndex).Cells("DIV").Style.BackColor = Color.Lime

        End Select
    End Sub

    Private Sub RichTextBox1_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox1.TextChanged

    End Sub

    Private Sub RichTextBox1_Leave(sender As Object, e As EventArgs) Handles RichTextBox1.Leave
        inserisci_commento_odp(Button3.Text, RichTextBox1.Text)
        MsgBox("Commento aggiornato")
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)

    End Sub

    Private Sub DataGridView6_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellContentClick

    End Sub

    Private Sub DataGridView6_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView6.CellClick
        Dim par_datagridview As DataGridView = DataGridView6

        If e.RowIndex >= 0 Then

            If e.ColumnIndex = par_datagridview.Columns.IndexOf(DataGridViewTextBoxColumn1) Then

                Magazzino.Codice_SAP = par_datagridview.Rows(e.RowIndex).Cells(columnName:="DataGridViewTextBoxColumn1").Value

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



    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        Form_Cicli_di_lavoro.Show()
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs)
        Macchine_In_Movimento()
        Percentuale_Lavorazione(DateTimePicker1, DateTimePicker2)
    End Sub

    Private Sub DataGridView_ODP_TIPO_MACCHINA_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP_TIPO_MACCHINA.CellContentClick

    End Sub

    Private Sub Button28_Click_1(sender As Object, e As EventArgs) Handles Button28.Click
        If RichTextBox2.Text = "" Then
            RichTextBox2.Text = 0
        End If
        If RichTextBox3.Text = "" Then
            RichTextBox3.Text = 0
        End If
        If Button3.Text = "" Then
            MsgBox("selezionare un ordine di produzione")
        End If
        If Homepage.trova_Dettagli_dipendente(Codicedip).utente_galileo = "" Then
            MsgBox("Non risulta nessun dipendente galileo assegnato")
            Return
        End If

        If ODP_Form.trova_esistenza_odp(Button3.Text) = False Then
            MsgBox("Ordine di produzione non esistente, segnalare al responsabile")
            Return
        End If

        If MessageBox.Show($"Vuoi chiudere la fase?", "Chiudi fase", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            inserisci_tempo_a_Galileo(Button3.Text, Homepage.trova_Dettagli_dipendente(Codicedip).utente_galileo, risorsa, code_fase, num_bolla, RichTextBox2.Text, RichTextBox3.Text, TrackBar1.Value, "S")
            Button28.Enabled = False
            riempi_datagridview(TextBox8.Text, TextBox7.Text, TextBox5.Text, TextBox4.Text, TextBox17.Text)
        Else
            inserisci_tempo_a_Galileo(Button3.Text, Homepage.trova_Dettagli_dipendente(Codicedip).utente_galileo, risorsa, code_fase, num_bolla, RichTextBox2.Text, RichTextBox3.Text, TrackBar1.Value, "")
        End If

        rECORD_CARICATI(DataGridView2)
        MsgBox("Tempo inserito con successo")

    End Sub

    Sub inserisci_tempo_a_Galileo(par_n_odp As String, par_dipendente As String, PAR_RISORSA As String, par_code_fase As String, num_bolla As String, par_tempo_attrezzaggio As Integer, par_tempo_lavorazione As Integer, par_percentuale As Integer, par_stato_fase As String)


        Dim centiore_attrezzaggio As Decimal = Math.Round(par_tempo_attrezzaggio / 60D, 2)
        Dim centiore_lavorazione_macchina As Decimal = Math.Round(par_tempo_lavorazione / 60D, 2)
        Dim centiore_lavorazione_uomo = centiore_lavorazione_macchina * par_percentuale / 10
        Dim dataOggiInt As Integer = CInt(Date.Today.ToString("yyyyMMdd"))
        Dim ID_giornata As Integer = Trova_ID_giornata(par_n_odp, dataOggiInt)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn



        CMD_SAP.CommandText = "INSERT INTO [AS400].[S786FAD1].[TIR90VIS].[YPCMOV0F] 

(PROFYP
,CDDIYP
, DT01YP
, CDDTYP
, DTOPYP
,NROPYP
, RIGAYP
, DTMOYP
, ORPRYP
, CDFAYP
,CAMOYP

,NRBLYP
,CDCLYP
,CDMUYP
,FATMYP
,FATUYP
,FLAMYP
,FLAUYP
,SAACYP
) 
VALUES 
('TIR40'
,'" & par_dipendente & "' 
, " & dataOggiInt & "
, '01' 
, " & dataOggiInt & "
, 400
, " & ID_giornata & "
, " & dataOggiInt & "
,'" & par_n_odp & "' 
, '" & par_code_fase & "'
, '01'

,'" & num_bolla & "' 
,substring('" & PAR_RISORSA & "', 1, 3)
,substring('" & PAR_RISORSA & "', 4, 3)
, " & Replace(centiore_attrezzaggio, ",", ".") & "
, " & Replace(centiore_attrezzaggio, ",", ".") & "
, " & Replace(centiore_lavorazione_macchina, ",", ".") & "
, " & Replace(centiore_lavorazione_uomo, ",", ".") & "
,'" & par_stato_fase & "'
)"

        CMD_SAP.ExecuteNonQuery()
        Cnn.Close()
    End Sub







    Public Function Trova_ID_giornata(par_n_odp As String, par_data As String)
        Dim id As Integer = 1
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select *
FROM OPENQUERY(AS400, '
select max(RIGAYP) as RIGAYP
from
S786FAD1.TIR90VIS.YPCMOV0F
where DT01YP =''" & par_data & "'' 

'
) as t10"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("RIGAYP") Is System.DBNull.Value Then
                id = cmd_SAP_reader_2("RIGAYP") + 1
            Else
                id = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
        Return id
    End Function
    Private Sub RichTextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles RichTextBox2.KeyPress
        ' Consente solo numeri e tasti di controllo (backspace, delete, ecc.)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub
    Private Sub RichTextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles RichTextBox3.KeyPress
        ' Consente solo numeri e tasti di controllo (backspace, delete, ecc.)
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub



    Private Sub Button29_Click(sender As Object, e As EventArgs)
        MsgBox(TrackBar1.Value)
    End Sub

    Private Sub RichTextBox2_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox2.TextChanged
        RichTextBox2.SelectionStart = RichTextBox2.TextLength
        RichTextBox2.SelectionLength = 0
        RichTextBox2.SelectionAlignment = HorizontalAlignment.Center
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        Label3.Text = TrackBar1.Value * 10 & " % "
        Dim tempo As Integer = 0
        If RichTextBox3.Text = "" Then
            tempo = 0
        Else
            tempo = RichTextBox3.Text
        End If
        Label30.Text = TrackBar1.Value * tempo / 10 & " Min"
    End Sub

    Private Sub RichTextBox3_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox3.TextChanged
        RichTextBox3.SelectionStart = RichTextBox3.TextLength
        RichTextBox3.SelectionLength = 0
        RichTextBox3.SelectionAlignment = HorizontalAlignment.Center
        Label3.Text = TrackBar1.Value * 10 & " % "

        Dim valore As Decimal

        If Not Decimal.TryParse(Label30.Text, valore) Then
            valore = 0
        End If




        Label30.Text = TrackBar1.Value * valore / 10 & " Min"
    End Sub

    Private Sub GroupBox47_Enter(sender As Object, e As EventArgs) Handles GroupBox47.Enter

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Try


            If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "S" Then
                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
            ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "O" Then
                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Yellow
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DataGridView1_CellLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellLeave

    End Sub

    Private Sub DataGridView5_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellContentClick

    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        riempi_datagridview(TextBox8.Text, TextBox7.Text, TextBox5.Text, TextBox4.Text, TextBox17.Text)
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Magazzino.apri_picture(Button4.Text)
    End Sub
End Class