Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib

Public Class Business_partner

    Public Codice_BP_selezionato As String
    Public Nome_BP_selezionato As String
    Public Codice_BP_JG_selezionato As String

    Public Provenienza As String


    Sub carica_controlli(par_datagridview As DataGridView, par_codice_bp_sap As String, par_codice_bp_galileo As String, par_nome_bp As String)
        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        If Provenienza = "Magazzino" Or Provenienza = "Form_entrata_merce" Then
            CMD_SAP_2.Connection = Cnn1

            If Homepage.ERP_provenienza = "SAP" Then
                CMD_SAP_2.CommandText = "SELECT top 100 t0.cardcode, t0.cardname ,'' as 'cardcode_jG','' as 'clifor'
from TIRELLISRLDB.DBO.ocrd t0 where (t0.cardtype='S') and t0.frozenFor='N' 
and t0.cardcode Like '%%" & par_codice_bp_sap & "%%' and t0.cardname Like '%%" & par_nome_bp & "%%'"

            Else

                CMD_SAP_2.CommandText = "select
coalesce(trim(t0.codesap),'') as 'Cardcode'
,coalesce(t0.conto,'') as 'Cardcode_JG'
,coalesce(t0.ds_conto,'') as 'Cardname'
,*
FROM OPENQUERY(AS400, '
    SELECT *
    FROM S786FAD1.TIR90VIS.JGALACF
where t0.clifor<>''D'' and t0.codesap Like ''%%" & par_codice_bp_sap & "%%'' and t0.conto Like ''%%" & par_codice_bp_galileo & "%%'' and upper(t0.cardname) Like ''%%" & par_nome_bp & "%%''

'
)
as t0"

            End If

        Else

            CMD_SAP_2.Connection = Cnn1
            If Homepage.ERP_provenienza = "SAP" Then
                CMD_SAP_2.CommandText = "SELECT top 100 t0.cardcode, t0.cardname ,'' as 'cardcode_jG','' as 'clifor'
from TIRELLISRLDB.DBO.ocrd t0 
where (t0.cardtype='C' or t0.cardtype='L') and t0.frozenFor='N' 
and t0.cardcode Like '%%" & par_codice_bp_sap & "%%' and t0.cardname Like '%%" & par_nome_bp & "%%'"
            Else

                CMD_SAP_2.CommandText = "SELECT
    COALESCE(TRIM(t0.codesap), '') AS Cardcode,
    COALESCE(t0.conto, '') AS Cardcode_JG,
    COALESCE(trim(t0.ds_conto), '') AS Cardname
,t0.clifor
FROM OPENQUERY(AS400, '
    SELECT *
    FROM S786FAD1.TIR90VIS.JGALACF
where clifor<>''D'' and codesap LIKE ''%" & par_codice_bp_sap & "%'' and conto LIKE ''%" & par_codice_bp_galileo & "%'' and ds_conto LIKE ''%" & par_nome_bp & "%''

'
)
as t0"
            End If


        End If

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(cmd_SAP_reader_2("cardcode"), cmd_SAP_reader_2("cardcode_JG"), cmd_SAP_reader_2("cardname"), cmd_SAP_reader_2("clifor"))


        Loop

        Cnn1.Close()

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        filtra_bp()
    End Sub
    Sub filtra_bp()
        carica_controlli(DataGridView1, TextBox1.Text, TextBox3.Text.ToUpper, TextBox2.Text.ToUpper)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        filtra_bp()
    End Sub



    Private Sub Business_partner_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        filtra_bp()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            Codice_BP_selezionato = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice_BP").Value
            Codice_BP_JG_selezionato = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice_JGalileo").Value
            Nome_BP_selezionato = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Nome_BP").Value

            If Provenienza = "Help_desk_interventi_BP" Then

                Form_scheda_intervento.Show()

                Form_scheda_intervento.Codice_BP = Codice_BP_selezionato
                Form_scheda_intervento.Codice_BP_jG = Codice_BP_JG_selezionato
                Form_scheda_intervento.Label3.Text = Nome_BP_selezionato



            ElseIf Provenienza = "Help_desk_tickets_BP" Then

                Form_TICKETS_HELP_DESK.Show()

                Form_TICKETS_HELP_DESK.codice_bp = Codice_BP_selezionato
                Form_TICKETS_HELP_DESK.Codice_BP_JG_selezionato = Codice_BP_JG_selezionato
                Form_TICKETS_HELP_DESK.Label1.Text = Nome_BP_selezionato


            ElseIf Provenienza = "Help_desk_interventi_BP_finale" Then

                Form_scheda_intervento.Show()

                Form_scheda_intervento.Codice_BP_finale = Codice_BP_selezionato
                Form_scheda_intervento.Label4.Text = Nome_BP_selezionato
                Form_scheda_intervento.Codice_BP_finale_jG = Codice_BP_JG_selezionato


            ElseIf Provenienza = "UT" Then
                UT.TextBox10.Text = Codice_BP_selezionato
                UT.TextBox5.Text = Nome_BP_selezionato

            ElseIf Provenienza = "Form_nuovo_campione" Then
                Form_nuovo_campione.Codice_BP_selezionato = Codice_BP_selezionato
                Form_nuovo_campione.Label1.Text = Nome_BP_selezionato
                Form_nuovo_campione.Codice_BP_JG_selezionato = Codice_BP_JG_selezionato
            ElseIf Provenienza = "Form_campione_visualizza" Then
                Form_campione_visualizza.cliente_cambiato = True
                Form_campione_visualizza.codice_bp_SELEZIONATO = Codice_BP_selezionato
                Form_campione_visualizza.codice_bp_jgal_SELEZIONATO = Codice_BP_JG_selezionato
                Form_campione_visualizza.Label1.Text = Nome_BP_selezionato

            ElseIf Provenienza = "Magazzino" Then
                Magazzino.codice_fornitore = Codice_BP_selezionato
                Magazzino.Label3.Text = Nome_BP_selezionato

            ElseIf Provenienza = "Form_nuova_combinazione" Then
                Form_Nuova_combinazione.codice_bp = Codice_BP_selezionato
                Form_Nuova_combinazione.Label1.Text = Nome_BP_selezionato
                Form_Nuova_combinazione.inizializza_form()
            ElseIf Provenienza = "Ciclo_di_lavoro" Then
                Form_Cicli_di_lavoro.codice_bp = Codice_BP_selezionato
                Form_Nuova_combinazione.Label3.Text = Nome_BP_selezionato

            ElseIf Provenienza = "Form_entrata_merce" Then

                Form_Entrate_Merci.Show()

                Form_Entrate_Merci.Codice_BP_finale = Codice_BP_selezionato
                Form_Entrate_Merci.Label7.Text = Nome_BP_selezionato

            End If
            Me.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Public Function Trova_business_partner(par_codice_business_partner As String)

        Dim cardname As String = ""


        Dim Cnn1 As New SqlConnection


        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_2.CommandText = "select coalesce(t0.cardname,'') as 'Cardname'
from ocrd t0
where t0.cardcode='" & par_codice_business_partner & "' "
        Else
            CMD_SAP_2.CommandText =
                "select
coalesce(t0.codesap,'') as 'Cardcode'
,coalesce(t0.conto,'') as 'Cardcode_JG'
,coalesce(trim(t0.ds_conto),'') as 'Cardname'

FROM OPENQUERY(AS400, '
    SELECT *
    FROM S786FAD1.TIR90VIS.JGALACF t0
where t0.codesap = ''" & par_codice_business_partner & "'' 

'
)
as t0"
        End If



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then

            cardname = cmd_SAP_reader_2("cardname")

        End If

        cmd_SAP_reader_2.Close()

        Cnn1.Close()

        Return cardname
    End Function

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class