Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib

Public Class Help_Desk_Interventi
    Public Elenco_dipendenti(1000) As String
    Public Elenco_owner(1000) As String

    Public codicedip As Integer
    Public esito_controllo As String
    Public inserimento As String

    Public Codice_BP As String
    Public Codice_BP_finale As String

    Public data_selezione_inizio As String
    Public data_selezione_fine As String

    Public id_intervento
    Public id_intervento_da_modificare
    Private FILTRO_DIP As String
    Private filtro_commessa As String
    Private filtro_cliente As String
    Private filtro_stato As String
    Private FILTRO_owner As String
    Private filtro_divisione As String

    Sub Inserimento_dipendenti(par_combobox As ComboBox)
        par_combobox.Items.Clear()

        par_combobox.Items.Add("")
        Dim Indice As Integer
        Indice = 0
        Indice = Indice + 1
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T1.[empID] as 'Codice dipendenti', T1.[lastName] + ' ' + T1.[firstName] AS 'Nome'
        FROM [Tirelli_40].[dbo].[Help_desk_interventi_effettuati_jc] T0 
inner join [TIRELLI_40].[DBO].OHEM t1 on T0.dipendente=T1.EMPID 
group by T1.[empID],T1.[lastName] + ' ' + T1.[firstName]
order by T1.[lastName] + ' ' + T1.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")


            par_combobox.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Private Sub TextBox1KeyPress(sender As Object, e As KeyPressEventArgs)
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
    End Sub 'permetto solo numeri come input

    Private Sub Help_Desk_Interventi_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Inserimento_dipendenti(ComboBox3)
        Inserimento_owner(ComboBox4)
        carica_interventi_effettuati()
    End Sub

    Sub inizializza_form()
        carica_interventi_effettuati()
    End Sub


    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
            codicedip = Elenco_dipendenti(ComboBox3.SelectedIndex)

        Catch ex As Exception

        End Try

        If ComboBox3.SelectedIndex <= 0 Then
            FILTRO_DIP = ""
        Else

            FILTRO_DIP = " and t10.empid= " & codicedip & " "


        End If
        ' carica_interventi_effettuati()'


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Business_partner.Show()
        Me.Hide()
        Business_partner.Owner = Me
        Business_partner.Provenienza = "Help_desk_interventi_BP"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        Business_partner.Show()
        Me.Hide()
        Business_partner.Owner = Me
        Business_partner.Provenienza = "Help_desk_interventi_BP_finale"
    End Sub





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form_scheda_intervento.Show()
        Form_scheda_intervento.Button4.Visible = False
        Form_scheda_intervento.Button7.Visible = False
        Form_scheda_intervento.Button1.Visible = True
    End Sub

    Sub Trova_ID()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select max(id)+1 as 'ID' from 
[Tirelli_40].[dbo].[Help_desk_interventi_effettuati_jc]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                id_intervento = cmd_SAP_reader_2("ID")
            Else
                id_intervento = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub



    Sub carica_interventi_effettuati()

        DataGridView1.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        CMD_SAP_2.CommandTimeout = 180

        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then


            CMD_SAP_2.CommandText = "select top 99999 t0.id,t1.lastname as 'owner', t2.lastname as 'dipendente' ,t0.commessa,coalesce(t0.ocrcode,'') as 'ocrcode',case when t5.itemname is null then '' else t5.itemname end as 'itemname',t3.cardname,case when t4.cardname is null then '' else t4.cardname end as 'cardname_final',t0.data_inizio,t0.data_fine,t0.ore,t0.causale,t0.stato,case when t0.comments is null then '' else t0.comments end as 'comments'
,t0.fattura,coalesce(t0.nota_spese,0),t0.costo_vitto_alloggio,t0.costo_trasporto

from [Tirelli_40].[dbo].[Help_desk_interventi_effettuati] t0 
inner join [TIRELLI_40].[DBO].ohem t1 on t1.empid=t0.owner 

inner join [TIRELLI_40].[DBO].ohem t2 on t2.empid=t0.dipendente
inner join ocrd t3 on t0.cardcode=t3.cardcode
left join ocrd t4 on t0.cardcode_final=t4.cardcode
left join oitm t5 on t5.itemcode=t0.commessa

WHERE 0=0 " & FILTRO_DIP & " " & filtro_commessa & " " & filtro_cliente & "" & filtro_stato & FILTRO_owner & filtro_divisione & "

order by t0.data_inizio DESC, t0.id desc


"
        Else
            CMD_SAP_2.CommandText = "select *
from
(
SELECT 
    t0.id,
    t0.dipendente AS 'empid',
    t1.lastname   AS 'owner',
    t2.lastname   AS 'dipendente',
    t0.commessa,
    COALESCE(t0.ocrcode, '')    AS 'ocrcode',
    ''     AS 'itemname',
    trim(t3.DS_CONTO)                 AS 'Cardname',
    ISNULL(trim(t4.DS_CONTO), '')     AS 'cardname_final',
    t0.data_inizio,
    t0.data_fine,
    t0.ore,
    t0.causale,
    t0.stato,
    ISNULL(t0.comments, '')     AS 'comments',
    t0.fattura,
    coalesce(t0.nota_spese,0) as 'Nota_spese',
    t0.costo_vitto_alloggio,
    t0.costo_trasporto
FROM [Tirelli_40].[dbo].[Help_desk_interventi_effettuati_jC] t0
LEFT JOIN [TIRELLI_40].[DBO].ohem t1
    ON t1.empid = t0.owner
LEFT JOIN [TIRELLI_40].[DBO].ohem t2
    ON t2.empid = t0.dipendente
LEFT JOIN [AS400].[S786FAD1].[TIR90VIS].[JGALACF] t3
    ON t3.conto = t0.cardcode_JG
LEFT JOIN [AS400].[S786FAD1].[TIR90VIS].[JGALACF] t4
    ON t4.conto = t0.cardcode_Final_JG
   AND COALESCE(t0.cardcode_final_jg, '') <> ''

	)
	as t10
	WHERE 0=0   " & FILTRO_DIP & " " & filtro_commessa & " " & filtro_cliente & "" & filtro_stato & FILTRO_owner & filtro_divisione & "
ORDER BY
    t10.data_inizio DESC,
    t10.id DESC"



        End If
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            DataGridView1.Rows.Add(cmd_SAP_reader_2("id"), cmd_SAP_reader_2("owner"), cmd_SAP_reader_2("dipendente"), cmd_SAP_reader_2("commessa"), cmd_SAP_reader_2("ocrcode"), cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("cardname"), cmd_SAP_reader_2("cardname_final"), cmd_SAP_reader_2("data_inizio"), cmd_SAP_reader_2("data_fine"), cmd_SAP_reader_2("ore"), cmd_SAP_reader_2("causale"), cmd_SAP_reader_2("stato"), cmd_SAP_reader_2("comments"), cmd_SAP_reader_2("fattura"), cmd_SAP_reader_2("nota_spese"), cmd_SAP_reader_2("costo_vitto_alloggio"), cmd_SAP_reader_2("costo_trasporto"))


        Loop

        Cnn1.Close()

        DataGridView1.ClearSelection()

    End Sub




    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            Form_scheda_intervento.id_intervento_da_modificare = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ID").Value

            Form_scheda_intervento.Show()
            Form_scheda_intervento.inizializza_form(Form_scheda_intervento.id_intervento_da_modificare)


            Form_scheda_intervento.Button4.Visible = True
            Form_scheda_intervento.Button7.Visible = True
            Form_scheda_intervento.Button1.Visible = False


        End If
    End Sub

    Sub ELIMINA_intervento()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "DELETE [Tirelli_40].[dbo].[Help_desk_interventi_effettuati_jc] 
WHERE ID='" & id_intervento_da_modificare & "'"
        CMD_SAP.ExecuteNonQuery()

        cnn.Close()
    End Sub

    Sub conferma_intervento()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].[Help_desk_interventi_effettuati_jc] SET STATO='R' WHERE ID='" & id_intervento_da_modificare & "'"
        CMD_SAP.ExecuteNonQuery()

        cnn.Close()
    End Sub



    Sub PULIZIA_caselle()
        ComboBox3.SelectedIndex = -1
        TextBox2.Text = Nothing

        Codice_BP = Nothing
        Codice_BP_finale = Nothing


        TextBox1.Text = Nothing

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        ELIMINA_intervento()
        carica_interventi_effettuati()


    End Sub



    Sub Inserimento_owner(PAR_COMBOBOX As ComboBox)
        PAR_COMBOBOX.Items.Clear()
        PAR_COMBOBOX.Items.Add("")
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T1.[empID] as 'Codice dipendenti', T1.[lastName] + ' ' + T1.[firstName] AS 'Nome'
        FROM [Tirelli_40].[dbo].[Help_desk_interventi_effettuati_jc] T0 
inner join [TIRELLI_40].[DBO].OHEM t1 on T0.OWNER=T1.EMPID 
group by T1.[empID],T1.[lastName] + ' ' + T1.[firstName]
order by T1.[lastName] + ' ' + T1.[firstName]  "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 1
        Do While cmd_SAP_reader.Read()
            Elenco_owner(Indice) = cmd_SAP_reader("Codice dipendenti")

            PAR_COMBOBOX.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
        Homepage.Show()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Process.Start(Homepage.percorso_server & "09-Service\Gantt interventi.xlsx")
    End Sub


    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = "" Then

            filtro_commessa = ""
        Else
            filtro_commessa = " and t10.commessa   Like '%%" & TextBox2.Text.ToUpper & "%%' "

        End If


    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then

            filtro_cliente = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then



                filtro_cliente = " and (t3.cardname   Like '%%" & TextBox1.Text.ToUpper & "%%'   or t4.cardname  Like '%%" & TextBox1.Text & "%%') "
            Else



                filtro_cliente = " AND  (UPPER(t10.cardname_final) LIKE '%" & TextBox1.Text.ToUpper & "%' " &
    "OR UPPER(t10.Cardname) LIKE '%" & TextBox1.Text.ToUpper & "%')"


            End If
        End If

        ' carica_interventi_effettuati()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContextMenuStripChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContextMenuStripChanged

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then

            filtro_stato = ""


        End If

        '  carica_interventi_effettuati()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

        If RadioButton2.Checked = True Then



            filtro_stato = " and t0.stato= 'P'"

        End If

        ' carica_interventi_effettuati()
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = True Then


            filtro_stato = " and t0.stato= 'R'"

        End If

        ' carica_interventi_effettuati()
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting

        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="causale").Value = "Vendita" Then



            DataGridView1.Rows(e.RowIndex).Cells(columnName:="causale").Style.BackColor = Color.Lime

        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="causale").Value.ToString().Contains("Completamento") Then

            DataGridView1.Rows(e.RowIndex).Cells(columnName:="causale").Style.BackColor = Color.Gold

        Else
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="causale").Style.BackColor = Color.Orange


        End If


        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "P" Or DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "p" Then



            DataGridView1.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Color.Aqua





        Else
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="stato").Style.BackColor = Color.Gray


        End If

        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="divisione").Value = "TIR01" Then



            DataGridView1.Rows(e.RowIndex).Cells(columnName:="divisione").Style.BackColor = Color.LightBlue

        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="divisione").Value = "BRB01" Then


            DataGridView1.Rows(e.RowIndex).Cells(columnName:="divisione").Style.BackColor = Color.Yellow

        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="divisione").Value = "KTF01" Then


            DataGridView1.Rows(e.RowIndex).Cells(columnName:="divisione").Style.BackColor = Color.Green


        End If


    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged


        If ComboBox4.SelectedIndex <= 0 Then
            FILTRO_owner = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then


                FILTRO_owner = " and t0.owner= '" & Elenco_owner(ComboBox4.SelectedIndex) & "' "
            Else
                FILTRO_owner = " and t10.owner= '" & Elenco_owner(ComboBox4.SelectedIndex) & "' "
            End If
        End If
        'carica_interventi_effettuati()
    End Sub



    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Then

            filtro_divisione = ""
        Else

            filtro_divisione = " and t10.ocrcode   Like '%%" & TextBox3.Text & "%%'  "



        End If

        ' carica_interventi_effettuati()
    End Sub

    Private Sub GroupBox3_Enter(sender As Object, e As EventArgs) Handles GroupBox3.Enter

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        carica_interventi_effettuati()
    End Sub

    Private Sub GroupBox4_Enter(sender As Object, e As EventArgs) Handles GroupBox4.Enter

    End Sub
End Class