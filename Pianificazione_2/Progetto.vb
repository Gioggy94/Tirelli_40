Imports System.Data.SqlClient
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop
Imports Microsoft.VisualBasic.Logging

Public Class Progetto
    Public Elenco_dipendenti(1000) As String
    Public Elenco_acquisti(1000) As String
    Public Elenco_stati_progetto(100) As String
    Public absentry As Integer
    Public cartella_progetto As String
    Private id As Integer
    Private codicedip As Integer
    Public commessa As String
    Private NUMERO_GROUPBOX As Integer = 15
    ' Dichiarazione di una variabile per memorizzare i contenuti del file copiato
    Private fileContents As Byte()
    Private n_documento As Integer
    Private sottocartella As String
    Private cartella_opportunità As String

    Public Sel_Stampante As New PrintDialog
    Public Stampante_Selezionata As Boolean

    Public altezza_Scontrino As Integer
    Public larghezza_scontrino As Integer
    Public numero_combinazioni As Integer = 0
    Private num_collaudati As Integer
    Private codice_bp As String
    Public resp_acquisti As Integer = 0
    Public numero_ultima_revisione As Integer = 0
    Public N_rev_visualizza As Integer
    Public codice_progetto As String

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Sub inizializza_progetto()
        Inserimento_dipendenti()
        Inserimento_dipendenti_acquisti(ComboBox2)
        anagrafica_progetto()
        crea_campi()
        compila_campi()
        Gantt_progetto()
        commesse_progetto()
        acconti_progetto()
        Scheda_tecnica.riempi_datagridview_campioni(DataGridView3, codice_bp, "", "", Homepage.Percorso_immagini, Homepage.sap_tirelli)

        trova_ultima_revisione_progetto(Label4.Text, codice_progetto)
        mostra_file_async(LinkLabel1.Text, TreeView1)
        ' esplodi_cartelle_macchina_del_progetto(TreeView2, Label4.Text)
        esplodi_cartelle_macchina_del_progetto_async(TreeView2, Label4.Text, Homepage.sap_tirelli, Homepage.percorso_cartelle_macchine)
        inizializza_scheda_tecnica_progetto(Label4.Text, codice_progetto)

    End Sub

    Sub inizializza_scheda_tecnica_progetto(par_numero_progetto As Integer, par_codice_progetto As String)


        elenca_revisioni_progetto(par_numero_progetto)
        trova_ultima_revisione_progetto(par_numero_progetto, par_codice_progetto)
        Label7.Text = numero_ultima_revisione

        riempi_scheda_tecnica_progetto(par_numero_progetto, numero_ultima_revisione, "Progetto")

        carica_appunti(Replace(Label4.Text, "PJ", ""), "PROGETTO", DataGridView5, "", Homepage.ID_SALVATO)
    End Sub

    Sub riempi_scheda_tecnica_progetto(par_numero_progetto As Integer, par_n_rev As Integer, par_fonte As String)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "
SELECT 
    t0.[ID],
    t0.[n_progetto],
    t0.[Rev],
    t0.[Note],
    COALESCE(t2.[Nome],'') AS [Stato_scheda],
    t0.[Imballo],
    t0.[Prezzo_Imballo],
    t0.[Trasporto],
    t0.[Prezzo_Trasporto],
    t0.[Installazione],
    t0.[Prezzo_Installazione]
,[id_rischio_cliente]
      ,[id_rischio_geografico]
      ,[indice_rischio_riempimento]
      ,[indice_rischio_handling_tappo]
      ,[indice_rischio_handling_bottiglia]
      ,[indice_rischio_handling_etichettatura]
      ,[indice_rischio_cf_tipologia]
      ,[indice_rischio_complessita_fornitura]
      ,[indice_rischio_ambiente_lavoro]
      ,[indice_rischio_vendita]
      ,[indice_rischio_tecnico]
      ,[indice_rischio_progetto]
      ,[livello_rischio_progetto]
      ,[livello_rischio_totale]
      ,[note_progetto]

FROM [Tirelli_40].[dbo].[Scheda_Tecnica_valori_progetto] t0
LEFT JOIN [Tirelli_40].[dbo].[Scheda_tecnica_revisioni_progetto] t1 
    ON t1.n_progetto = '" & par_numero_progetto & "' 
    AND t1.numero = '" & par_n_rev & "'
LEFT JOIN [Tirelli_40].[dbo].[Scheda_Tecnica_stato_progetto] t2 
    ON t1.stato = t2.id
WHERE t0.n_progetto = '" & par_numero_progetto & "' 
  AND t0.rev = '" & par_n_rev & "'
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            'If par_fonte = "Progetto" Then
            ComboBox70.Text = cmd_SAP_reader_2("Stato_scheda")

            Label11.Text = cmd_SAP_reader_2("livello_rischio_totale")
            RichTextBox1.Text = cmd_SAP_reader_2("Note")
            Scheda_tecnica.RichTextBox7.Text = cmd_SAP_reader_2("Note")

            ' ✅ nuovi campi letti dal database
            CheckBox1.Checked = (cmd_SAP_reader_2("Imballo") = True)
                TextBox1.Text = If(IsDBNull(cmd_SAP_reader_2("Prezzo_Imballo")), "", cmd_SAP_reader_2("Prezzo_Imballo").ToString())

                CheckBox2.Checked = (cmd_SAP_reader_2("Trasporto") = True)
                TextBox2.Text = If(IsDBNull(cmd_SAP_reader_2("Prezzo_Trasporto")), "", cmd_SAP_reader_2("Prezzo_Trasporto").ToString())

                CheckBox3.Checked = (cmd_SAP_reader_2("Installazione") = True)
                TextBox3.Text = If(IsDBNull(cmd_SAP_reader_2("Prezzo_Installazione")), "", cmd_SAP_reader_2("Prezzo_Installazione").ToString())

            If cmd_SAP_reader_2("id_rischio_cliente") > 0 Then
                ComboBox3.Text = cmd_SAP_reader_2("id_rischio_cliente").ToString()
            Else
                ComboBox3.SelectedIndex = -1
            End If

            If cmd_SAP_reader_2("id_rischio_geografico") > 0 Then
                ComboBox4.Text = cmd_SAP_reader_2("id_rischio_geografico").ToString()
            Else
                ComboBox4.SelectedIndex = -1
            End If

            If cmd_SAP_reader_2("indice_rischio_riempimento") > 0 Then
                ComboBox6.Text = cmd_SAP_reader_2("indice_rischio_riempimento").ToString()
            Else
                ComboBox6.SelectedIndex = -1
            End If

            If cmd_SAP_reader_2("indice_rischio_handling_tappo") > 0 Then
                ComboBox5.Text = cmd_SAP_reader_2("indice_rischio_handling_tappo").ToString()
            Else
                ComboBox5.SelectedIndex = -1
            End If

            If cmd_SAP_reader_2("indice_rischio_handling_bottiglia") > 0 Then
                ComboBox7.Text = cmd_SAP_reader_2("indice_rischio_handling_bottiglia").ToString()
            Else
                ComboBox7.SelectedIndex = -1
            End If

            If cmd_SAP_reader_2("indice_rischio_handling_etichettatura") > 0 Then
                ComboBox8.Text = cmd_SAP_reader_2("indice_rischio_handling_etichettatura").ToString()
            Else
                ComboBox8.SelectedIndex = -1
            End If

            If cmd_SAP_reader_2("indice_rischio_cf_tipologia") > 0 Then
                ComboBox9.Text = cmd_SAP_reader_2("indice_rischio_cf_tipologia").ToString()
            Else
                ComboBox9.SelectedIndex = -1
            End If

            If cmd_SAP_reader_2("indice_rischio_complessita_fornitura") > 0 Then
                ComboBox10.Text = cmd_SAP_reader_2("indice_rischio_complessita_fornitura").ToString()
            Else
                ComboBox10.SelectedIndex = -1
            End If

            If cmd_SAP_reader_2("indice_rischio_ambiente_lavoro") > 0 Then
                ComboBox11.Text = cmd_SAP_reader_2("indice_rischio_ambiente_lavoro").ToString()
            Else
                ComboBox11.SelectedIndex = -1
            End If

            ' ElseIf par_fonte = "Scheda_tecnica" Then
            '  Scheda_tecnica.RichTextBox7.Text = cmd_SAP_reader_2("Note")

            '    ' ✅ anche in questa modalità puoi leggere i nuovi campi se ti servono:
            '    Scheda_tecnica.CheckBox1.Checked = (cmd_SAP_reader_2("Imballo") = True)
            '    Scheda_tecnica.TextBox1.Text = If(IsDBNull(cmd_SAP_reader_2("Prezzo_Imballo")), "", cmd_SAP_reader_2("Prezzo_Imballo").ToString())

            '    Scheda_tecnica.CheckBox2.Checked = (cmd_SAP_reader_2("Trasporto") = True)
            '    Scheda_tecnica.TextBox2.Text = If(IsDBNull(cmd_SAP_reader_2("Prezzo_Trasporto")), "", cmd_SAP_reader_2("Prezzo_Trasporto").ToString())

            '    Scheda_tecnica.CheckBox3.Checked = (cmd_SAP_reader_2("Installazione") = True)
            '    Scheda_tecnica.TextBox3.Text = If(IsDBNull(cmd_SAP_reader_2("Prezzo_Installazione")), "", cmd_SAP_reader_2("Prezzo_Installazione").ToString())
            'End If
        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub







    Private Sub Form1_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged

        FlowLayoutPanel1.AutoScroll = True
        For Each groupBox As GroupBox In FlowLayoutPanel1.Controls.OfType(Of GroupBox)()
            ' Modifica le dimensioni delle GroupBox
            groupBox.Height = FlowLayoutPanel1.Height \ NUMERO_GROUPBOX * 0.9
            groupBox.Width = FlowLayoutPanel1.Width - 40 ' Considera lo spazio per la scrollbar

            ' Ridimensiona i controlli interni
            For Each control As Control In groupBox.Controls
                If TypeOf control Is Button Or TypeOf control Is Label Then
                    control.Width = groupBox.Width * 0.15
                ElseIf TypeOf control Is TextBox Then
                    control.Width = groupBox.Width * 0.55
                End If
            Next
        Next
    End Sub

    Sub anagrafica_progetto()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP.CommandText = "Select T0.ABSENTRY,T0.DOCNUM, T0.CARDCODE, T0.CARDNAME
, T0.U_CODICE_CLIENTE_FINALE
, COALESCE(T4.CARDNAME,'') as 'u_cliente_finale',t0.name, CONCAT(T1.LASTNAME,' ',T1.FIRSTNAME) AS 'PM', T2.SLPNAME

, coalesce(t3.lastname,' ',t3.firstname) as 'Resp_acquisti'
,coalesce(t0.U_Resp_acquisti,0) as 'U_Resp_acquisti'
From opmg T0 LEFT Join [TIRELLI_40].[dbo].OHEM T1 ON T0.OWNER=T1.EMPID
Left Join OSLP T2 ON T2.SLPCODE=T0.EMPLOYEE
left join [TIRELLI_40].[dbo].ohem t3 on t3.empid=t0.U_Resp_acquisti
LEFT JOIN OCRD T4 ON T4.CARDCODE=T0.U_CODICE_CLIENTE_FINALE

where T0.series >= 2075 And T0.STATUS ='S' and t0.absentry=" & absentry & ""
        Else
            CMD_SAP.CommandText = "SELECT top 1 
'' as 'Name'
,'' as 'docnum'
,'' as 'slpname'
,100 as 'u_resp_acquisti'
,'' as 'resp_acquisti'
,trim(t10.matricola) as 'Itemcode'
, t10.itemname
, t10.desc_supp
, T10.DSCLI_FATT as 'Cardname'
, T10.CLI_FATT as 'Cardcode',
        t10.codice_finale as 'U_CLIENTE_FINALE'
		, t10.numero_progetto as 'absentry',
        trim(numero_progetto) as 'Numero_progetto',
		T10.NAME_progetto AS 'DESC_PROGETTO',
		'' as 'Nome_stato',
        '' as 'Livello_rischio_totale', '' as 'Name',
        t10.pm as 'CODICE_PM'
		,t10.DESC_pm as 'PM'
		, T10.DSNAZ_FINALE as u_country_of_delivery,
        t10.brand AS 'CODICE_BRAND',
		T10.DESC_BRAND AS 'BRAND',
		'' as 'Baia'
		, '' as 'Zona'
		,DATA_CONSEGNA
		,T10.NOME_STATO AS 'STATO_COMMESSA'
FROM OPENQUERY(AS400, '
    SELECT *
    FROM TIR90VIS.JGALCOM t0
 
WHERE 
      UPPER(t0.matricola) LIKE ''%%%%''
      AND UPPER(t0.itemname) LIKE ''%%%%''
      AND (UPPER(t0.codice_finale) LIKE ''%%%%'' 
           OR UPPER(t0.dscli_fatt) LIKE ''%%%%'') 
      AND (
          t0.itemcode = ''" & absentry & "''  
          OR (LEFT(RTRIM(t0.itemcode),2) = ''T0'' 
              AND RIGHT(RTRIM(t0.itemcode), LENGTH(''" & absentry & "'')) = ''" & absentry & "'')
      )
	  and t0.matricola<>''''
  
ORDER BY T0.ITEMCODE 

limit 100  
') T10

"
        End If


        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then
            Label10.Text = codice_progetto
            Label1.Text = Replace(cmd_SAP_reader("cardname"), ",", ".")
            Label2.Text = Replace(cmd_SAP_reader("U_CLIENTE_FINALE"), ",", ".")
            Label3.Text = Replace(cmd_SAP_reader("name"), """", " ")
            Label4.Text = absentry
            Label5.Text = cmd_SAP_reader("PM")
            Label6.Text = cmd_SAP_reader("slpname")
            LinkLabel1.Text = Scheda_tecnica.trova_percorso_documenti(absentry, "PROGETTO", codice_progetto)
            'LinkLabel1.Text = cmd_SAP_reader("cartella")
            codice_bp = cmd_SAP_reader("CARDCODE")
            resp_acquisti = cmd_SAP_reader("U_Resp_acquisti")
            ComboBox2.Text = cmd_SAP_reader("Resp_acquisti")

        End If
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub


    Sub crea_campi()
        Dim groupBoxes(66) As GroupBox
        Dim i As Integer = 0
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "select *
                            from [TIRELLI_40].DBO.Requisiti_progetto
                            where active='Y' and documento='Progetto'
                            order by cast(ID as integer) "
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            'Crea un nuovo groupbox per ogni requisito
            groupBoxes(i) = New GroupBox()
            FlowLayoutPanel1.Controls.Add(groupBoxes(i))

            'Imposta le proprietà del groupbox
            Dim id As String = cmd_SAP_reader("ID").ToString()
            If id.Length = 1 Then
                id = "0" & id
            End If

            With groupBoxes(i)
                .Name = cmd_SAP_reader("codice_requisito")
                .Text = id & " " & cmd_SAP_reader("nome requisito")
                .Dock = DockStyle.Bottom
                .Height = FlowLayoutPanel1.Height \ NUMERO_GROUPBOX * 0.8
                .Width = FlowLayoutPanel1.Width - 20
            End With

            'Aggiungi lo spazio di caricamento dei documenti al groupbox
            Dim checkbox_complete As New CheckBox()
            With checkbox_complete
                .Name = "checkbox_" & cmd_SAP_reader("codice_requisito").ToString()
                .Text = "Completo"
                .Width = groupBoxes(i).Width * 0.2
                .Height = groupBoxes(i).Height * 0.5
                .Dock = DockStyle.Left
            End With

            AddHandler checkbox_complete.CheckedChanged, AddressOf CheckBoxValueChanged

            groupBoxes(i).Controls.Add(checkbox_complete)
            ' ...



            'Aggiungi il pulsante "Completo" al groupbox
            Dim button1 As New Button()
            button1.Name = "Button_completo_" & cmd_SAP_reader("codice_requisito").ToString()
            button1.Text = "Crea cartella"
            button1.Width = groupBoxes(i).Width * 0.2
            button1.Height = groupBoxes(i).Height * 0.5
            button1.Dock = DockStyle.Left
            AddHandler button1.Click, AddressOf Button_Click
            groupBoxes(i).Controls.Add(button1)

            'Aggiungi il pulsante "Aggiorna" al groupbox
            Dim button2 As New Button()
            button2.Name = "Button_aggiorna_" & cmd_SAP_reader("codice_requisito").ToString()
            button2.Text = "Aggiorna"
            button2.Width = groupBoxes(i).Width * 0.2
            button2.Height = groupBoxes(i).Height * 0.5
            button2.Dock = DockStyle.Left
            AddHandler button2.Click, AddressOf Button_Click
            groupBoxes(i).Controls.Add(button2)



            'Aggiungi la casella di testo al groupbox
            Dim richtextBox As New RichTextBox()
            richtextBox.Name = "Richtextbox_" & cmd_SAP_reader("codice_requisito")
            richtextBox.Dock = DockStyle.Left
            richtextBox.Width = groupBoxes(i).Width * 0.4
            richtextBox.Height = groupBoxes(i).Height * 0.5
            ' richtextBox.Height = groupBoxes(i).Height
            groupBoxes(i).Controls.Add(richtextBox)

            i = i + 1
        Loop

        cmd_SAP_reader.Close()
        cnn.Close()
    End Sub

    Private Sub CheckBoxValueChanged(sender As Object, e As EventArgs)
        ' Codice da eseguire quando il valore della checkbox cambia
    End Sub

    Sub compila_campi()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "  declare @absentry as integer
  set @absentry=" & absentry & "
  set @absentry=" & absentry & "

  select t0.campo_progetto, t0.data, t0.ora, t0.testo, t0.completo, t0.dipendente
  from
  [TIRELLI_40].DBO.progetto_logs t0 inner join 
  (
    select campo_progetto ,max(id) as 'Max_id' from [TIRELLI_40].DBO.progetto_logs
  where absentry=@absentry
  group by campo_progetto
  ) A on a.max_id= t0.id
  where t0.absentry=@absentry"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read() = True

            Dim campoProgetto As String = "Richtextbox_" & cmd_SAP_reader("campo_progetto").ToString()
            Dim rtb As RichTextBox = CType(Me.Controls.Find(campoProgetto, True)(0), RichTextBox)
            rtb.Text = cmd_SAP_reader("testo")


        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub





    Private Sub Button_Click(sender As Object, e As EventArgs)
        Dim button As Button = DirectCast(sender, Button)
        Dim actionName As String = ""
        If button.Name.StartsWith("Button_aggiorna") Then

            actionName = button.Name.Substring("Button_aggiorna_".Length)
            Dim parentGroupBox As GroupBox = DirectCast(button.Parent, GroupBox)
            Dim richTextBox As RichTextBox = parentGroupBox.Controls.OfType(Of RichTextBox)().FirstOrDefault()

            If richTextBox IsNot Nothing Then

                aggiorna_log_progetto(richTextBox.Text, actionName)
            End If

            MsgBox("Campo aggiornato con successo")
        ElseIf button.Name.StartsWith("Button_completo") Then

            Dim requisito As String = DirectCast(button.Parent, GroupBox).Text ' ottenere il testo del GroupBox genitore
            requisito = requisito.Trim()
            Dim cartellaRequisito As String = Path.Combine(Homepage.percorso_progetti & LinkLabel1.Text, requisito)

            If Not Directory.Exists(cartellaRequisito) Then
                Try
                    Directory.CreateDirectory(cartellaRequisito)
                Catch ex As Exception
                    MsgBox("Non è possibile creare la cartella per un errore interno. creare la cartella nel server accedendo al progetto")
                End Try

            End If
            mostra_file_async(LinkLabel1.Text, TreeView1)
        End If


        If TypeOf sender Is Button Then
            Dim button_1 As Button = DirectCast(sender, Button)

        End If


    End Sub

    Sub aggiorna_log_progetto(contenuto As String, campo_progetto As String)
        Trova_ID()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO [Tirelli_40].dbo.PROGETTO_LOGS (ID,ABSENTRY,CAMPO_PROGETTO,DATA,ORA,TESTO,COMPLETO,DIPENDENTE)
VALUES (" & id & ", " & absentry & ",'" & campo_progetto & "',getdate(),convert(varchar, getdate(), 108),'" & contenuto & "',''," & codicedip & ")"
        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()

    End Sub

    Sub Trova_ID()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "Select max(id)+1 As 'ID' from [TIRELLI_40].DBO.PROGETTO_LOGS"

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


    Sub Inserimento_dipendenti()
        Combodipendenti.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
        LEFT JOIN [TIRELLI_40].DBO.COLL_REPARTI T2 ON T2.SAP_ID_REPARTO=T0.DEPT
where t0.active='Y' and cast(t2.ID_REPARTO as varchar)='" & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto & "' order by T0.[lastName]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            Combodipendenti.Items.Add(cmd_SAP_reader("Nome"))
            If cmd_SAP_reader("Codice dipendenti") = Homepage.ID_SALVATO Then
                Combodipendenti.SelectedIndex = Indice
            End If
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()


    End Sub

    Sub commesse_progetto()
        DataGridView_commesse.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn

        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP.CommandText = "select *
from
(
SELECT T0.[ItemCode], T0.[ItemName]
, coalesce(t0.u_progetto,0) as 'u_progetto' 
FROM OITM T0 

WHERE  substring(t0.itemcode,1,1)='M'
--coalesce(t0.u_progetto,0)=4 
)
as t10
where  t10.u_progetto=" & absentry & "
order by t10.itemcode "
        Else
            CMD_SAP.CommandText = "SELECT top 100 trim(t10.matricola) as 'Itemcode', t10.itemname, t10.desc_supp
, T10.DSCLI_FATT as 'Cliente'
, T10.CLI_FATT as 'Codice_cliente',
        t10.codice_finale as 'Cliente_finale'
		, t10.numero_progetto as 'absentry',
        trim(numero_progetto) as 'Numero_progetto',
		T10.NAME_progetto AS 'DESC_PROGETTO',
		'' as 'Nome_stato',
        '' as 'Livello_rischio_totale', '' as 'Name',
        t10.pm as 'CODICE_PM'
		,t10.DESC_pm as 'PM'
		, T10.DSNAZ_FINALE as u_country_of_delivery,
        t10.brand AS 'CODICE_BRAND',
		T10.DESC_BRAND AS 'BRAND',
		'' as 'Baia'
		, '' as 'Zona'
		,DATA_CONSEGNA
		,T10.NOME_STATO AS 'STATO_COMMESSA'
FROM OPENQUERY(AS400, '
    SELECT *
    FROM TIR90VIS.JGALCOM t0
 
WHERE 
      UPPER(t0.matricola) LIKE ''%%%%''
      AND UPPER(t0.itemname) LIKE ''%%%%''
      AND (UPPER(t0.codice_finale) LIKE ''%%%%'' 
           OR UPPER(t0.dscli_fatt) LIKE ''%%%%'') 
      AND (
          t0.itemcode = ''" & absentry & "''  
          OR (LEFT(RTRIM(t0.itemcode),2) = ''T0'' 
              AND RIGHT(RTRIM(t0.itemcode), LENGTH(''" & absentry & "'')) = ''" & absentry & "'')
      )
	  and t0.matricola<>''''
  
ORDER BY T0.ITEMCODE 

limit 100  
') T10"

        End If


        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            DataGridView_commesse.Rows.Add(cmd_SAP_reader("ItemCode"), cmd_SAP_reader("Itemname"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Sub cardini_commesse()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "select *
from
(
SELECT T0.[ItemCode], T0.[ItemName]
, coalesce(t0.u_progetto,0) as 'u_progetto' 
FROM OITM T0 

WHERE  substring(t0.itemcode,1,1)='M'
--coalesce(t0.u_progetto,0)=4 
)
as t10
where  t10.u_progetto=" & absentry & "
order by t10.itemcode DESc"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()

            Scheda_tecnica.crea_modulo_cardini(FlowLayoutPanel13, cmd_SAP_reader("ItemCode"), cmd_SAP_reader("Itemname"), Homepage.JPM_TIRELLI)

        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub

    Sub acconti_progetto()
        If Homepage.ERP_provenienza = "SAP" Then

            DataGridView2.Rows.Clear()
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()

            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

            CMD_SAP.Connection = Cnn

            CMD_SAP.CommandText = "
declare @absentry as integer
set @absentry =" & absentry & "

select t10.tipo_doc, t10.tabella_testata, t10.tabella_riga, t10.doc, t11.docnum, t11.doctotal
from
(
SELECT 'Order' as 'Tipo_doc', 'ORDR' as 'Tabella_testata', 'RDR1' as 'Tabella_riga', 'OC' as 'Doc', max(t2.DOCENTRY) as 'Docentry'

FROM OITM T0  left join RDR1 t1 on t0.itemcode=t1.itemcode
left join ORDR t2 on t2.docentry=t1.docentry

WHERE t0.u_progetto=@absentry
)
as t10 left join ordr t11 on t11.docentry=t10.docentry


union all

select  t10.tipo_doc, t10.tabella_testata, t10.tabella_riga, t10.doc, t11.docnum, t11.doctotal
from
(
SELECT 'Consegna' as 'Tipo_doc', 'ODLN' as 'Tabella_testata', 'DLN1' as 'Tabella_riga','DDT' as 'Doc', max(t2.DOCENTRY) as 'Docentry'

FROM OITM T0  left join DLN1 t1 on t0.itemcode=t1.itemcode
left join ODLN t2 on t2.docentry=t1.docentry

WHERE t0.u_progetto=@absentry
)
as t10 left join oDLN t11 on t11.docentry=t10.docentry

union all

select  t10.tipo_doc, t10.tabella_testata, t10.tabella_riga, t10.doc, t11.docnum, t11.doctotal
from
(
SELECT 'Fattura' as 'Tipo_doc', 'OINV' as 'Tabella_testata', 'INV1' as 'Tabella_riga','FATTURA' as 'Doc', max(t2.DOCENTRY) as 'Docentry'

FROM OITM T0  left join INV1 t1 on t0.itemcode=t1.itemcode
left join OINV t2 on t2.docentry=t1.docentry

WHERE t0.u_progetto=@absentry
)
as t10 left join OINV t11 on t11.docentry=t10.docentry

"



            cmd_SAP_reader = CMD_SAP.ExecuteReader

            Do While cmd_SAP_reader.Read()
                If Homepage.ERP_provenienza = "SAP" Then
                    DataGridView2.Rows.Add(cmd_SAP_reader("tipo_doc"), cmd_SAP_reader("tabella_testata"), cmd_SAP_reader("tabella_riga"), cmd_SAP_reader("DOC"), cmd_SAP_reader("docnum"), cmd_SAP_reader("doctotal"))

                End If
            Loop


            cmd_SAP_reader.Close()



            cmd_SAP_reader.Close()
            Cnn.Close()
        Else

        End If
    End Sub

    Sub Gantt_progetto()
        '        Dim contatore As Integer = 0
        '        ' Pulisci il Chart4 prima di popolarlo con i nuovi dati
        '        Chart4.Series.Clear()
        '        Chart4.Series.Add("Date")
        '        Chart4.Series("Date").ChartType = DataVisualization.Charting.SeriesChartType.RangeBar
        '        Chart4.ChartAreas(0).AxisX.LabelStyle.Font = New Font("Arial", 6)
        '        ' Imposta la larghezza desiderata per le barre del diagramma di Gantt
        '        Chart4.Series("Date")("PixelPointWidth") = "8"
        '        Dim Cnn As New SqlConnection
        '        Cnn.ConnectionString = Homepage.sap_tirelli
        '        cnn.Open()

        '        Dim CMD_SAP As New SqlCommand
        '        Dim cmd_SAP_reader As SqlDataReader

        '        CMD_SAP.Connection = cnn
        '        CMD_SAP.CommandText = "select *
        'from
        '(
        'SELECT T0.[ItemCode], T0.[ItemName],t1.ATTIVITA,t1.RISORSA, case when t2.resname is null then '' else t2.resname end as 'Nome ris', min( t1.DATA_I) as 'Data_I',max(t1.DATA_F) as 'Data_f'
        'FROM OITM T0 inner join [TIRELLI_40].DBO.PIANIFICAZIONE t1 on t1.commessa=t0.itemcode
        'left join orsc t2 on t2.visrescode=t1.risorsa
        'WHERE t0.u_progetto='" & absentry & "' 
        'group by T0.[ItemCode], T0.[ItemName],t1.ATTIVITA,t1.RISORSA,t2.resname
        ')
        'as t10
        'order by t10.itemcode DESC, t10.Data_I DESC
        '"

        '        cmd_SAP_reader = CMD_SAP.ExecuteReader



        '        ' Crea una nuova serie nel chart per il diagramma di Gantt
        '        Dim series As New DataVisualization.Charting.Series("GanttSeries")


        '        Do While cmd_SAP_reader.Read()
        '            Dim dataInizio As DateTime = Convert.ToDateTime(cmd_SAP_reader("Data_I"))
        '            Dim dataFine As DateTime = Convert.ToDateTime(cmd_SAP_reader("Data_F"))
        '            Dim label As String = dataInizio.ToString("dd/MM") & " - " & dataFine.ToString("dd/MM")
        '            Chart4.Series("Date").Points.AddXY(cmd_SAP_reader("itemcode") & " " & cmd_SAP_reader("Nome RIS") & " " & cmd_SAP_reader("attivita"), cmd_SAP_reader("Data_F"), cmd_SAP_reader("Data_I"))
        '            Chart4.Series("Date").Points.Last().Label = label


        '            ' Impostare l'etichetta sul punto e posizionarla a destra


        '            Chart4.Series("Date").Points.Last().LabelForeColor = Color.Black
        '            Chart4.Series("Date").Points.Last().LabelToolTip = label

        '        Loop

        '        Chart4.ChartAreas(0).AlignmentOrientation = DataVisualization.Charting.AreaAlignmentOrientations.Horizontal
        '        Chart4.ChartAreas(0).AlignmentStyle = DataVisualization.Charting.AreaAlignmentStyles.PlotPosition

        '        Chart4.ChartAreas(0).AxisX.LabelStyle.Interval = 1
        '        Chart4.ChartAreas(0).AxisX.LabelStyle.TruncatedLabels = False


        '        cmd_SAP_reader.Close()
        '        cnn.Close()
    End Sub

    Sub riempi_datagridview_campioni()
        Dim Cnn1 As New SqlConnection
        DataGridView3.Rows.Clear()

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        '        CMD_SAP_2.CommandText = "SELECT  t1,  t2.INIZIALE_SIGLA + T1.NOME as 'Nome', case when (t1.immagine is null or t1.immagine ='') then '" & Homepage.Percorso_immagini & "N_A.JPG'  else t1.immagine end as 'immagine', t2.descrizione as 'Tipo', t5.onhand
        'from opmg t0 inner join COLL_Campioni t1 on cast(t0.cardcode as integer) =t1.codice_bp or cast(t0.u_codice_cliente_finale as integer) =t1.codice_bp
        'inner  join  COLL_TIPO_CAMPIONE t2 on t1.TIPO_campione= T2.ID_TIPO_CAMPIONE
        'inner join oitm t3 on t3.u_progetto=t0.absentry
        'INNER join coll_combinazioni t4 on t4.Commessa=t3.ItemCode and(t4.Campione_1=t1 or t4.Campione_2=t1 or t4.Campione_3=t1 or t4.Campione_4=t1 or t4.Campione_5=t1 or t4.Campione_6=t1 or t4.Campione_7=t1 or t4.Campione_8=t1 or t4.Campione_9=t1 or t4.Campione_10=t1 )
        'left join oitm t5 on t5.itemcode=t1.codice_sap
        'where t0.absentry=' " & absentry & "'
        'group by  t1,  t2.INIZIALE_SIGLA + T1.NOME,case when (t1.immagine is null or t1.immagine ='') then '" & Homepage.Percorso_immagini & "N_A.JPG'  else t1.immagine end,t2.descrizione, t5.onhand
        'order by t2.INIZIALE_SIGLA + T1.NOME"

        CMD_SAP_2.CommandText = "SELECT   t1.id_Campione,  t2.INIZIALE_SIGLA + T1.NOME as 'Nome', case when (t1.immagine is null or t1.immagine ='') then '" & Homepage.Percorso_immagini & "N_A.JPG'  else t1.immagine end as 'immagine', t2.descrizione as 'Tipo', t5.onhand, t1.Dato_6, t1.descrizione
from opmg t0 inner join oitm t3 on t3.u_progetto=t0.absentry
INNER join [TIRELLI_40].DBO.coll_combinazioni t4 on t4.Commessa=t3.ItemCode 
INNER JOIN [TIRELLI_40].DBO.COLL_CAMPIONI T1 ON t4.Campione_1=t1.id_Campione or t4.Campione_2=t1.id_Campione or t4.Campione_3=t1.id_Campione or t4.Campione_4=t1.id_Campione or t4.Campione_5=t1.id_Campione or t4.Campione_6=t1.id_Campione or t4.Campione_7=t1.id_Campione or t4.Campione_8=t1.id_Campione or t4.Campione_9=t1.id_Campione or t4.Campione_10=t1.id_Campione 

inner  join  [TIRELLI_40].DBO.COLL_TIPO_CAMPIONE t2 on t1.TIPO_campione= T2.ID_TIPO_CAMPIONE


left join oitm t5 on t5.itemcode=t1.codice_sap
where t0.absentry=' " & absentry & "'
group by  t1.id_Campione,  t2.INIZIALE_SIGLA ,t1.nome,case when (t1.immagine is null or t1.immagine ='') then '" & Homepage.Percorso_immagini & "N_A.JPG'  else t1.immagine end,t2.descrizione, t5.onhand, t1.Dato_6, t1.descrizione
order by

t2.INIZIALE_SIGLA ,  cast(substring(T1.NOME,1,99) as integer)"



        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        ' Dim contatore As Integer = 0
        Dim i As Integer = 0
        Do While cmd_SAP_reader_2.Read()


            Dim MyImage As Bitmap

            Try
                MyImage = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine"))
            Catch ex As Exception
                MyImage = Image.FromFile(Homepage.Percorso_immagini & "\N_A.JPG")

            End Try

            DataGridView3.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("Tipo"), MyImage, cmd_SAP_reader_2("onhand"), cmd_SAP_reader_2("dato_6"), cmd_SAP_reader_2("descrizione")) 'Image.FromFile(cmd_SAP_reader_2("immagine"))

            i = i + 1
        Loop


        cmd_SAP_reader_2.Close()
        Cnn1.Close()


        DataGridView3.ClearSelection()

    End Sub



    Private Sub DataGridView_commesse_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_commesse.CellClick
        If e.RowIndex >= 0 Then
            commessa = DataGridView_commesse.Rows(e.RowIndex).Cells(columnName:="codice").Value

            Scheda_tecnica.riempi_datagridview_combinazioni(DataGridView1, commessa, Homepage.sap_tirelli)
            If e.ColumnIndex = DataGridView_commesse.Columns.IndexOf(Codice) Then

                If DataGridView_commesse.Rows(e.RowIndex).Cells(columnName:="codice").Value >= "M04000" Then
                    Scheda_tecnica.Show()
                    Scheda_tecnica.BringToFront()
                    Scheda_tecnica.inizializza_scheda_tecnica(DataGridView_commesse.Rows(e.RowIndex).Cells(columnName:="codice").Value)
                    Try
                        Scheda_tecnica.codice_bp_campione = DataGridView_commesse.Rows(e.RowIndex).Cells(columnName:="Codice_cliente").Value
                    Catch ex As Exception

                    End Try
                Else

                    Scheda_commessa_documentazione.inizializzazione = 0
                    Scheda_commessa_documentazione.carico_iniziale = 0

                    Scheda_commessa_documentazione.Azzera_campi()

                    Scheda_commessa_documentazione.commessa = DataGridView_commesse.Rows(e.RowIndex).Cells(columnName:="Codice").Value


                    Scheda_commessa_Pianificazione.layout_scheda_tecnica()


                    Scheda_commessa_documentazione.compila_anagrafica(Scheda_commessa_documentazione.commessa)


                    Scheda_commessa_documentazione.Inserimento_dipendenti()

                    Scheda_commessa_documentazione.COMPILA_RECORD_INIZIALI()

                    '  Scheda_commessa_documentazione.Rischio_effettivo()
                    Scheda_commessa_documentazione.Ultimo_aggiornamento()
                    Scheda_commessa_documentazione.riempi_datagridview_combinazioni()
                    Scheda_tecnica.riempi_datagridview_campioni(DataGridView3, codice_bp, "", "", Homepage.Percorso_immagini, Homepage.sap_tirelli)
                    Scheda_commessa_documentazione.cerca_file()
                    Scheda_commessa_documentazione.Show()



                    Scheda_commessa_documentazione.carico_iniziale = 1
                    Scheda_commessa_documentazione.inizializzazione = 1

                End If


            End If
        End If
    End Sub



    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        If e.RowIndex >= 0 Then




            Form_campione_visualizza.id_campione = DataGridView3.Rows(e.RowIndex).Cells(columnName:="Campione_").Value
            Form_campione_visualizza.Show()

            Form_campione_visualizza.inizializza_form()

        End If
    End Sub

    Sub riempi_datagridview_combinazioni()

        Dim Larghezza_Colonna_Immagine As Integer
        Dim Larghezza_Colonna_Testo As Integer
        Dim Larghezza_Colonna_Bottone As Integer

        Larghezza_Colonna_Immagine = DataGridView1.Width * 15 / 100
        Larghezza_Colonna_Testo = DataGridView1.Width * 7 / 100
        Larghezza_Colonna_Bottone = DataGridView1.Width * 5 / 100


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        trova_percorso()
    End Sub

    Sub trova_percorso()
        Dim rootDirectory As String = Homepage.percorso_progetti

        ' Estrai la parte numerica da codice_progetto (es. "T00260" → 260)
        Dim codiceClean As String = codice_progetto.Trim()
        Dim parteNumerica As Integer = Integer.Parse(codiceClean.Substring(1)) ' Rimuove la "T"

        Dim blocco As Integer = (parteNumerica \ 1000) + 1
        Dim nomeCartellaBlocco As String = $"Progetti {blocco}-1000"
        Dim targetDirectory As String = Path.Combine(rootDirectory, nomeCartellaBlocco)

        If Not Directory.Exists(targetDirectory) Then
            MsgBox("Cartella del blocco progetti non trovata: " & targetDirectory, MsgBoxStyle.Critical)
            Exit Sub
        End If

        ' Cerca tutte le cartelle che iniziano con "Progetto T00260"
        Dim pattern As String = $"Progetto {codiceClean}"
        Dim cartelleTutte As String() = Directory.GetDirectories(targetDirectory)
        Dim matchingFolders As New List(Of String)

        For Each cartella In cartelleTutte
            Dim nomeCartella As String = Path.GetFileName(cartella)
            If nomeCartella.StartsWith(pattern, StringComparison.OrdinalIgnoreCase) Then
                matchingFolders.Add(cartella)
            End If
        Next

        Dim cartellaProgettoCompleta As String

        If matchingFolders.Count = 1 Then
            cartellaProgettoCompleta = matchingFolders(0)
        ElseIf matchingFolders.Count > 1 Then
            cartellaProgettoCompleta = matchingFolders(0) ' oppure mostra una scelta
        Else
            ' Nessuna trovata: crea una nuova cartella
            Dim cliente As String = If(Label2.Text = "", Label1.Text, Label2.Text)
            Dim nomeProgetto As String = $"Progetto {codiceClean} {cliente} {Label3.Text}"
            Dim nomeProgettoPulito As String = PulisciTesto(nomeProgetto)
            cartellaProgettoCompleta = Path.Combine(targetDirectory, nomeProgettoPulito)
            Directory.CreateDirectory(cartellaProgettoCompleta)
        End If

        ' Imposta il percorso relativo
        Dim sottocartella As String = Replace(cartellaProgettoCompleta, Homepage.percorso_progetti, "")
        LinkLabel1.Text = sottocartella
        Scheda_tecnica.Aggiorna_percorso_macchina(Replace(LinkLabel1.Text, "'", " "), codice_progetto, "PROGETTO")
    End Sub

    ' 🔧 Funzione per pulire il testo da caratteri speciali nei nomi da creare
    Function PulisciTesto(testo As String) As String
        Return testo.Replace("–", "-") _
                .Replace("—", "-") _
                .Replace("’", "'") _
                .Replace("‘", "'") _
                .Replace("“", """") _
                .Replace("”", """") _
                .Replace("`", "'") _
                .Replace("´", "'") _
                .Trim()
    End Function

    Sub crea_sottocartelle()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "select *
from [TIRELLI_40].DBO.Requisiti_progetto where active='Y' and cartella='Y'
where documento='Progetto' order by ID"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            Directory.CreateDirectory(cartella_progetto & "\" & cmd_SAP_reader("id") & "_" & cmd_SAP_reader("nome requisito"))

        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub Aggiorna_percorso_progetto_old()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Dim destinazione = Replace(LinkLabel1.Text, "'", " ")
        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE t0 SET t0.U_percorso_cartella='" & destinazione & "' 
from opmg t0
where t0.docnum=" & Replace(Label4.Text, "PJ", "") & ""
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub
    'commento
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Dim percorso As String = Homepage.percorso_progetti & LinkLabel1.Text

        Process.Start(percorso)
    End Sub
    Public Sub esplodi_cartelle_macchina_del_progetto_async(par_treeview As TreeView, par_docnum As String, connectionString As String, par_percorso_macchina As String)

    End Sub
    Public Sub mostra_file_async(par_percorso As String, par_treeview As TreeView)
        Dim codiceClean As String = codice_progetto.Trim()
        Dim rootDirectoryPath As String = Homepage.percorso_progetti & par_percorso

        If Not Directory.Exists(rootDirectoryPath) Or LinkLabel1.Text = "" Then
            MsgBox("La cartella " & rootDirectoryPath & " non esiste, né è stata trovata o creata una per questo progetto.")
            trova_percorso()

            ' Ricalcola il percorso dopo trova_percorso()
            rootDirectoryPath = Homepage.percorso_progetti & LinkLabel1.Text

            ' Se ancora non esiste, esci
            If Not Directory.Exists(rootDirectoryPath) Then
                MsgBox("Impossibile trovare o creare la cartella per il progetto " & codiceClean)
                Return
            End If
        End If

        ' Esegui l'operazione in background
        Task.Run(Sub()
                     Try
                         ' Pulisce la TreeView e aggiunge il nodo "Caricamento..."
                         par_treeview.Invoke(Sub()
                                                 par_treeview.Nodes.Clear()
                                                 Dim loadingNode As New TreeNode("🔄 Caricamento...")
                                                 par_treeview.Nodes.Add(loadingNode)
                                             End Sub)

                         Dim rootDirectory As New DirectoryInfo(rootDirectoryPath)
                         Dim rootNode As New TreeNode(rootDirectory.Name) With {.Tag = rootDirectory}

                         ' Popola il TreeView in background
                         AddDirectories(rootNode, par_treeview)
                         Addfiles(rootNode, par_treeview)

                         ' Rimuove il nodo "Caricamento..." e aggiunge i dati finali
                         par_treeview.Invoke(Sub()
                                                 par_treeview.Nodes.Clear()
                                                 par_treeview.Nodes.Add(rootNode)
                                                 par_treeview.ExpandAll()
                                                 par_treeview.AllowDrop = True
                                             End Sub)
                     Catch ex As Exception
                         par_treeview.Invoke(Sub()
                                                 MsgBox("Errore nel caricamento file: " & ex.Message, MsgBoxStyle.Critical)
                                             End Sub)
                     End Try
                 End Sub)
    End Sub


    Public Sub mostra_file_async_analisi_rischi(par_percorso As String, par_treeview As TreeView)

        Dim rootDirectoryPath As String = Homepage.percorso_progetti & par_percorso

        ' Esegui l'operazione in background
        Task.Run(Sub()
                     Try


                         ' Pulisce la TreeView e aggiunge il nodo "Caricamento..."
                         par_treeview.Invoke(Sub()
                                                 par_treeview.Nodes.Clear()
                                                 Dim loadingNode As New TreeNode("🔄 Caricamento...")
                                                 par_treeview.Nodes.Add(loadingNode)
                                             End Sub)

                         Dim rootDirectory As New DirectoryInfo(rootDirectoryPath)
                         Dim rootNode As New TreeNode(rootDirectory.Name) With {.Tag = rootDirectory}

                         ' Popola il TreeView in background
                         AddDirectories(rootNode, par_treeview)
                         Addfiles(rootNode, par_treeview)

                         ' Rimuove il nodo "Caricamento..." e aggiunge i dati finali
                         par_treeview.Invoke(Sub()
                                                 par_treeview.Nodes.Clear()
                                                 par_treeview.Nodes.Add(rootNode)
                                                 par_treeview.ExpandAll()
                                                 par_treeview.AllowDrop = True
                                             End Sub)
                     Catch ex As Exception

                     End Try
                 End Sub)
    End Sub


    Public Sub AddDirectories(parentNode As TreeNode, par_treeview As TreeView)
        Dim parentDirectory As DirectoryInfo = TryCast(parentNode.Tag, DirectoryInfo)
        If parentDirectory Is Nothing OrElse Not parentDirectory.Exists Then Exit Sub

        Debug.WriteLine("📂 Scansiono cartella: " & parentDirectory.FullName)
        Try


            ' Aggiungi icona cartella in modo sicuro
            par_treeview.Invoke(Sub() AggiungiIconaCartella())

            ' Aggiunge tutte le cartelle come nodi figli
            For Each directory As DirectoryInfo In parentDirectory.GetDirectories()
                Dim directoryNode As New TreeNode(directory.Name) With {
        .Tag = directory,
        .ImageKey = "folder",           ' Imposta l'icona della cartella
        .SelectedImageKey = "folder"     ' Imposta l'icona selezionata
    }

                ' Aggiunge il nodo in modo thread-safe
                par_treeview.Invoke(Sub() parentNode.Nodes.Add(directoryNode))

                ' Aggiunge anche i file nella cartella
                Addfiles(directoryNode, par_treeview)

                ' Chiamata ricorsiva per le sottocartelle
                AddDirectories(directoryNode, par_treeview)
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub AggiungiIconaCartella()
        If Not ImageList1.Images.ContainsKey("folder") Then
            ImageList1.Images.Add("folder", SystemIcons.WinLogo)
        End If
    End Sub

    Public Sub Addfiles(parentNode As TreeNode, par_treeview As TreeView)
        Dim parentDirectory As DirectoryInfo = TryCast(parentNode.Tag, DirectoryInfo)
        If parentDirectory Is Nothing OrElse Not parentDirectory.Exists Then Exit Sub

        Try
            For Each file As FileInfo In parentDirectory.GetFiles()
                Dim fileNode As New TreeNode(file.Name & " (" & file.LastWriteTime.ToString() & ")") With {.Tag = file}

                ' Sostituzione del percorso di rete
                Dim filepath As String = file.FullName.Replace("\\tirfs01\Tirelli", "T:")
                '  Dim filepath As String = file.FullName

                ' Ottiene l'icona del file
                Dim fileIcon As Icon = SystemIcons.WinLogo
                Try
                    fileIcon = Icon.ExtractAssociatedIcon(filepath)
                Catch ex As Exception
                    Debug.WriteLine("⚠️ ERRORE: Impossibile estrarre icona per " & filepath)
                End Try

                ' Aggiunge l'icona in modo sicuro
                par_treeview.Invoke(Sub() AggiungiIconaFile(file.Extension, fileIcon))

                fileNode.ImageKey = file.Extension

                ' Aggiunge il nodo file in modo thread-safe
                par_treeview.Invoke(Sub() parentNode.Nodes.Add(fileNode))
            Next
        Catch ex As Exception
            Debug.WriteLine("❌ ERRORE in Addfiles: " & ex.Message)
        End Try
    End Sub


    ' Funzione per aggiungere l'icona di un file in modo sicuro
    Private Sub AggiungiIconaFile(extension As String, icon As Icon)
        If Not ImageList1.Images.ContainsKey(extension) Then
            ImageList1.Images.Add(extension, icon)
        End If
    End Sub



    Private Sub TreeView2_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView2.NodeMouseDoubleClick
        ' Verifica se il nodo selezionato è un file
        If TypeOf e.Node.Tag Is FileInfo Then
            ' Apri il file con l'applicazione predefinita
            Dim file As FileInfo = DirectCast(e.Node.Tag, FileInfo)
            Process.Start(file.FullName)
        ElseIf TypeOf e.Node.Tag Is DirectoryInfo Then
            Dim directory As DirectoryInfo = DirectCast(e.Node.Tag, DirectoryInfo)
            Process.Start(directory.FullName)
        End If
    End Sub

    Private Sub TreeView1_DragEnter(sender As Object, e As DragEventArgs) Handles TreeView1.DragEnter
        ' Verifica che il file trascinato sia un file
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' Imposta il cursore come "Copia" durante il trascinamento
            e.Effect = DragDropEffects.Copy
        Else
            ' Imposta il cursore come "No" durante il trascinamento
            e.Effect = DragDropEffects.None
        End If
    End Sub
    Private Sub TreeView2_DragEnter(sender As Object, e As DragEventArgs) Handles TreeView2.DragEnter
        ' Verifica che il file trascinato sia un file
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' Imposta il cursore come "Copia" durante il trascinamento
            e.Effect = DragDropEffects.Copy
        Else
            ' Imposta il cursore come "No" durante il trascinamento
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub TreeView1_DragDrop(sender As Object, e As DragEventArgs) Handles TreeView1.DragDrop

        Dim par_treeview As TreeView = TreeView1
        ' Verifica che il file trascinato sia un file
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' Ottiene il percorso completo del file trascinato
            Dim filePath As String = CType(e.Data.GetData(DataFormats.FileDrop), String())(0)

            ' Ottiene il nodo selezionato
            Dim selectedNode As TreeNode = TreeView1.SelectedNode

            ' Verifica che il nodo selezionato sia una cartella
            If selectedNode Is Nothing OrElse Not TypeOf selectedNode.Tag Is DirectoryInfo Then
                MsgBox("Selezionare una cartella per il salvataggio del file.")
                Return
            End If

            ' Salva il file nella cartella del nodo selezionato
            Dim targetDirectory As DirectoryInfo = CType(selectedNode.Tag, DirectoryInfo)
            File.Copy(filePath, Path.Combine(targetDirectory.FullName, Path.GetFileName(filePath)), True)

            ' Aggiorna la TreeView
            selectedNode.Nodes.Clear()
            AddDirectories(selectedNode, par_treeview)
            Addfiles(selectedNode, par_treeview)

            ' Espande il nodo selezionato
            selectedNode.Expand()
        End If
        mostra_file_async(LinkLabel1.Text, par_treeview)
    End Sub

    Private Sub TreeView2_DragDrop(sender As Object, e As DragEventArgs) Handles TreeView2.DragDrop
        Dim par_treeview As TreeView = TreeView2
        ' Verifica che il file trascinato sia un file
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' Ottiene il percorso completo del file trascinato
            Dim filePath As String = CType(e.Data.GetData(DataFormats.FileDrop), String())(0)

            ' Ottiene il nodo selezionato
            Dim selectedNode As TreeNode = TreeView1.SelectedNode

            ' Verifica che il nodo selezionato sia una cartella
            If selectedNode Is Nothing OrElse Not TypeOf selectedNode.Tag Is DirectoryInfo Then
                MsgBox("Selezionare una cartella per il salvataggio del file.")
                Return
            End If

            ' Salva il file nella cartella del nodo selezionato
            Dim targetDirectory As DirectoryInfo = CType(selectedNode.Tag, DirectoryInfo)
            File.Copy(filePath, Path.Combine(targetDirectory.FullName, Path.GetFileName(filePath)), True)

            ' Aggiorna la TreeView
            selectedNode.Nodes.Clear()
            AddDirectories(selectedNode, par_treeview)
            Addfiles(selectedNode, par_treeview)
            ' Espande il nodo selezionato
            selectedNode.Expand()
        End If
        '  mostra_file_async(LinkLabel1.Text, TreeView1)
    End Sub





    Private Sub Apri_file_Click(sender As Object, e As EventArgs) Handles Apri_file.Click
        ' Verifica se il nodo selezionato è un file
        If TypeOf TreeView1.SelectedNode.Tag Is FileInfo Then
            ' Se il nodo selezionato è un file, apri il file
            Dim file As FileInfo = DirectCast(TreeView1.SelectedNode.Tag, FileInfo)
            Process.Start(file.FullName)
        ElseIf TypeOf TreeView1.SelectedNode.Tag Is DirectoryInfo Then
            ' Se il nodo selezionato è una directory, apri la cartella
            Dim directory As DirectoryInfo = DirectCast(TreeView1.SelectedNode.Tag, DirectoryInfo)
            Process.Start("explorer.exe", directory.FullName)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Crea una nuova istanza di OpenFileDialog
        Dim openFileDialog As New OpenFileDialog()

        ' Imposta le proprietà del dialogo
        openFileDialog.Filter = "Tutti i file (*.*)|*.*"
        openFileDialog.Title = "Seleziona una directory"
        openFileDialog.CheckFileExists = False
        openFileDialog.CheckPathExists = True
        openFileDialog.FileName = "Seleziona una directory"

        ' Mostra la finestra di dialogo per la selezione della directory
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' Assegna la directory selezionata al controllo LinkLabel
            LinkLabel1.Text = Path.GetDirectoryName(openFileDialog.FileName)
        End If
        Scheda_tecnica.Aggiorna_percorso_macchina(Replace(LinkLabel1.Text, "'", " "), codice_progetto, "PROGETTO")
    End Sub

    Private Sub TreeView1_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseClick
        If e.Button = MouseButtons.Right Then
            TreeView1.SelectedNode = TreeView1.GetNodeAt(e.X, e.Y)
        End If
    End Sub

    Private Sub TreeView2_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView2.NodeMouseClick
        If e.Button = MouseButtons.Right Then
            TreeView2.SelectedNode = TreeView2.GetNodeAt(e.X, e.Y)
        End If
    End Sub

    Private Sub RinominaFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RinominaFileToolStripMenuItem.Click
        Dim par_treeview As TreeView = TreeView1
        ' Controlla se un nodo è stato selezionato
        If par_treeview.SelectedNode IsNot Nothing Then
            Dim file As FileInfo = TryCast(par_treeview.SelectedNode.Tag, FileInfo)

            ' Apri una finestra di dialogo per consentire all'utente di inserire il nuovo nome del file
            Dim newFileName As String = InputBox("Inserisci il nuovo nome del file", "Rinomina file", file.Name)

            If Not String.IsNullOrEmpty(newFileName) Then
                ' Rinomina il file
                Dim newFilePath As String = Path.Combine(file.DirectoryName, newFileName)
                FileSystem.Rename(file.FullName, newFilePath)
                mostra_file_async(LinkLabel1.Text, par_treeview)


            End If
        End If

    End Sub

    Private Sub Combodipendenti_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combodipendenti.SelectedIndexChanged
        codicedip = Elenco_dipendenti(Combodipendenti.SelectedIndex)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text = "[]" Then

            Me.WindowState = FormWindowState.Maximized
            Button2.Text = "Riduci"
        ElseIf Button2.Text = "Riduci" Then
            Me.WindowState = FormWindowState.Normal
            Button2.Text = "[]"
        End If

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        mostra_file_async(LinkLabel1.Text, TreeView1)
    End Sub

    Private Sub DataGridView_commesse_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_commesse.CellContentClick

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        If e.RowIndex >= 0 Then


            If DataGridView1.Columns.IndexOf(Immagine_1) Or DataGridView1.Columns.IndexOf(Immagine_2) Or DataGridView1.Columns.IndexOf(Immagine_3) Or DataGridView1.Columns.IndexOf(Immagine_4) Or DataGridView1.Columns.IndexOf(Immagine_5) Or DataGridView1.Columns.IndexOf(Immagine_6) Or DataGridView1.Columns.IndexOf(Immagine_7) Or DataGridView1.Columns.IndexOf(Immagine_8) Or DataGridView1.Columns.IndexOf(Immagine_9) Or DataGridView1.Columns.IndexOf(Immagine_10) Then


                Form_campione_visualizza.id_campione = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex - 2).Value
                Form_campione_visualizza.Show()
                Form_campione_visualizza.inizializza_form()



            End If
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        esplodi_cartelle_macchina_del_progetto_async(TreeView2, Replace(Label4.Text, "PJ", ""), Homepage.sap_tirelli, Homepage.percorso_cartelle_macchine)
    End Sub


    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        ' Verifica se il nodo selezionato è un file
        If TypeOf TreeView2.SelectedNode.Tag Is FileInfo Then
            ' Se il nodo selezionato è un file, apri il file
            Dim file As FileInfo = DirectCast(TreeView2.SelectedNode.Tag, FileInfo)
            Process.Start(file.FullName)
        ElseIf TypeOf TreeView2.SelectedNode.Tag Is DirectoryInfo Then
            ' Se il nodo selezionato è una directory, apri la cartella
            Dim directory As DirectoryInfo = DirectCast(TreeView2.SelectedNode.Tag, DirectoryInfo)

            Process.Start(directory.FullName)
        End If
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        Dim par_treeview As TreeView = TreeView2
        ' Controlla se un nodo è stato selezionato
        If par_treeview.SelectedNode IsNot Nothing Then
            Dim file As FileInfo = TryCast(par_treeview.SelectedNode.Tag, FileInfo)

            ' Apri una finestra di dialogo per consentire all'utente di inserire il nuovo nome del file
            Dim newFileName As String = InputBox("Inserisci il nuovo nome del file", "Rinomina file", file.Name)

            If Not String.IsNullOrEmpty(newFileName) Then
                ' Rinomina il file
                Dim newFilePath As String = Path.Combine(file.DirectoryName, newFileName)
                FileSystem.Rename(file.FullName, newFilePath)
                mostra_file_async(LinkLabel1.Text, par_treeview)


            End If
        End If
    End Sub



    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub





    Private Sub Progetto_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Inserimento_STATO_PROGETTI(ComboBox70)
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.RowIndex >= 0 Then

            ComboBox1.Text = DataGridView2.Rows(e.RowIndex).Cells(columnName:="tIPO_DOC").Value
            n_documento = DataGridView2.Rows(e.RowIndex).Cells(columnName:="num").Value

            If e.ColumnIndex = DataGridView2.Columns.IndexOf(Num) Then

                Form_nuova_offerta.Show()

                Form_nuova_offerta.TextBox10.Text = DataGridView2.Rows(e.RowIndex).Cells(columnName:="num").Value
                Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                Form_nuova_offerta.inizializzazione_form(DataGridView2.Rows(e.RowIndex).Cells(columnName:="num").Value, DataGridView2.Rows(e.RowIndex).Cells(columnName:="Tabella_testata").Value, DataGridView2.Rows(e.RowIndex).Cells(columnName:="Tabella_righe").Value, DataGridView2.Rows(e.RowIndex).Cells(columnName:="tIPO_DOC").Value)

            End If
        End If

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Layout_documenti.Show()
        Layout_documenti.ComboBox1.SelectedIndex = ComboBox1.SelectedIndex
        Layout_documenti.TextBox1.Text = n_documento
        Layout_documenti.percorso_specifico = Homepage.percorso_progetti & LinkLabel1.Text & "\" & "11 Documenti spedizione\" & ComboBox1.Text & "_" & n_documento & ".doc"
        Layout_documenti.Button1.PerformClick()

        mostra_file_async(LinkLabel1.Text, TreeView1)
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect

    End Sub

    Private Sub TreeView1_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick
        ' Verifica se il nodo selezionato è una directory
        If TypeOf e.Node.Tag Is FileInfo Then
            ' Apri il file con l'applicazione predefinita
            Dim file As FileInfo = DirectCast(e.Node.Tag, FileInfo)
            Process.Start(file.FullName)
        ElseIf TypeOf e.Node.Tag Is DirectoryInfo Then
            ' Se il nodo selezionato è una directory, apri la cartella
            Dim directory As DirectoryInfo = DirectCast(e.Node.Tag, DirectoryInfo)
            Process.Start("explorer.exe", directory.FullName)
        End If


    End Sub

    Private Sub Elimina_file_Click_1(sender As Object, e As EventArgs) Handles Elimina_file.Click
        ' Verifica se il nodo selezionato è un file
        If TypeOf TreeView1.SelectedNode.Tag Is FileInfo Then
            ' Chiedi all'utente conferma prima di eliminare il file
            Dim file As FileInfo = DirectCast(TreeView1.SelectedNode.Tag, FileInfo)
            If MessageBox.Show($"Sei sicuro di voler eliminare il file '{file.Name}'?", "Elimina file", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                file.Delete()
                TreeView1.SelectedNode.Remove()
            End If
        ElseIf TypeOf TreeView1.SelectedNode.Tag Is DirectoryInfo Then
            ' Chiedi all'utente conferma prima di eliminare la directory
            Dim directory As DirectoryInfo = DirectCast(TreeView1.SelectedNode.Tag, DirectoryInfo)
            If MessageBox.Show($"Sei sicuro di voler eliminare la directory '{directory.Name}' e tutti i file contenuti al suo interno?", "Elimina directory", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                directory.Delete(True)
                TreeView1.SelectedNode.Remove()
            End If
        End If
    End Sub

    Private Sub EliminaFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EliminaFileToolStripMenuItem.Click
        ' Verifica se il nodo selezionato è un file
        If TypeOf TreeView2.SelectedNode.Tag Is FileInfo Then
            ' Chiedi all'utente conferma prima di eliminare il file
            Dim file As FileInfo = DirectCast(TreeView2.SelectedNode.Tag, FileInfo)
            If MessageBox.Show($"Sei sicuro di voler eliminare il file '{file.Name}'?", "Elimina file", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                file.Delete()
                TreeView2.SelectedNode.Remove()
            End If
        ElseIf TypeOf TreeView2.SelectedNode.Tag Is DirectoryInfo Then
            ' Chiedi all'utente conferma prima di eliminare la directory
            Dim directory As DirectoryInfo = DirectCast(TreeView2.SelectedNode.Tag, DirectoryInfo)
            If MessageBox.Show($"Sei sicuro di voler eliminare la directory '{directory.Name}' e tutti i file contenuti al suo interno?", "Elimina directory", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                directory.Delete(True)
                TreeView2.SelectedNode.Remove()
            End If
        End If
    End Sub

    Private Sub TreeView2_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView2.AfterSelect

    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Process.Start(LinkLabel2.Text)
    End Sub

    Private Sub TreeView1_DoubleClick(sender As Object, e As EventArgs) Handles TreeView1.DoubleClick

    End Sub

    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click

        'Etichetta cassetta
        Fun_Stampa(700, 185)

        'Etichetta EM
        'Fun_Stampa(200, 185)
    End Sub
    Sub Fun_Stampa(par_altezza_Scontrino As Integer, par_larghezza_scontrino As Integer)

        altezza_Scontrino = par_altezza_Scontrino
        larghezza_scontrino = par_larghezza_scontrino


        Dim preview_scontrino As Boolean
        preview_scontrino = False

        If preview_scontrino = True Then
            If Stampante_Selezionata = False Then
                Sel_Stampante.AllowSomePages = False
                Sel_Stampante.ShowHelp = False
                Sel_Stampante.Document = Scontrino

                ' Utilizza un PrintPreviewDialog per mostrare l'anteprima di stampa
                Dim previewDialog As New PrintPreviewDialog()
                previewDialog.Document = Scontrino

                Dim result As DialogResult = previewDialog.ShowDialog()

                If (result = DialogResult.OK) Then
                    Stampante_Selezionata = True

                    ' Ora la stampante è selezionata, puoi chiamare Scontrino.Print()
                    Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", par_larghezza_scontrino, par_altezza_Scontrino)
                    Scontrino.Print()
                End If
            Else
                ' Se la stampante è già stata selezionata in precedenza, stampa direttamente
                Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", par_larghezza_scontrino, par_altezza_Scontrino)
                Scontrino.Print()
            End If

        Else
            If Stampante_Selezionata = False Then
                Sel_Stampante.AllowSomePages = False
                Sel_Stampante.ShowHelp = False
                Sel_Stampante.Document = Scontrino
                Dim result As DialogResult = Sel_Stampante.ShowDialog()
                If (result = DialogResult.OK) Then
                    Stampante_Selezionata = True
                    ' Utilizza un PrintPreviewDialog per mostrare l'anteprima di stampa
                    Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", par_larghezza_scontrino, par_altezza_Scontrino)
                    Dim previewDialog As New PrintPreviewDialog()
                    previewDialog.Document = Scontrino
                    Scontrino.Print()
                End If
            Else
                Scontrino.Print()
            End If
        End If



    End Sub




    Private Sub Scontrino_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles Scontrino.PrintPage

        'Dichiaro caratteri di scrittura
        Dim Penna As New Pen(Color.Black)
        Dim Carattere_Titolo As New Font("Calibri", 18, FontStyle.Bold)
        Dim Carattere_Dati As New Font("Calibri", 25, FontStyle.Bold)

        'Dichiaro elementi query
        Dim Data As String = ""
        Dim cliente As String = ""
        Dim progetto As String = ""
        Dim PM As String = ""
        Dim matricole As String = ""

        'Collegamento a database SQL
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "
select t0.docnum, coalesce(t1.cardname,t0.cardname) as 'Cliente', t0.name as 'Nome_progetto', concat(t2.lastname, ' ', t2.firstname) as 'PM', CAST(GETDATE() AS DATE) as 'oggi',cast(GETDATE() as date) as 'Data', *

from OPMG t0

left join ocrd t1 on t0.u_Codice_cliente_finale=t1.cardcode

left join [TIRELLI_40].[dbo].ohem t2 on t2.empid=t0.owner
where t0.absentry='" & absentry & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() Then
            progetto = "N°" & cmd_SAP_reader("docnum") & " " & cmd_SAP_reader("Nome_progetto")
            cliente = cmd_SAP_reader("Cliente")
            Data = cmd_SAP_reader("data")
            PM = cmd_SAP_reader("PM")

        End If

        cmd_SAP_reader.Close()
        CNN.Close()


        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()


        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "
select t0.itemcode
from oitm t0 
where t0.u_progetto='" & absentry & "'
order by t0.itemcode DESC"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()
            matricole = cmd_SAP_reader("itemcode") & " " & matricole

        Loop

        cmd_SAP_reader.Close()
        CNN.Close()

        'Programmazione scontrino
        With e.Graphics
            .SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

            'Dichiaro libreria che permettere scritte verticali'
            Dim state As Drawing2D.GraphicsState = .Save()

            'Contornatura scontrino
            .DrawRectangle(Penna, 1, 1, larghezza_scontrino - 1, altezza_Scontrino - 1)

            'Dichiaro dimensionamento campi
            Dim altezza_ret_titoli As Integer = altezza_Scontrino / 4
            Dim altezza_ret_dati As Integer = altezza_Scontrino / 4 * 3
            Dim larghezza_ret As Integer = larghezza_scontrino / 5

            'Tabella base TITOLI
            .DrawRectangle(Penna, 1, 1, larghezza_ret, altezza_ret_titoli)
            .DrawRectangle(Penna, larghezza_ret * 1, 1, altezza_ret_titoli * 1, altezza_ret_titoli)
            .DrawRectangle(Penna, larghezza_ret * 2, 1, altezza_ret_titoli * 2, altezza_ret_titoli)
            .DrawRectangle(Penna, larghezza_ret * 3, 1, altezza_ret_titoli * 3, altezza_ret_titoli)
            .DrawRectangle(Penna, larghezza_ret * 4, 1, altezza_ret_titoli * 4, altezza_ret_titoli)

            'Scrittura base TITOLI VERTICALI

            'Rotazione testo in gradi
            .RotateTransform(+90)

            ' Scritta in verticale (ATTENZIONE Il rifermiento delle scritte è lo spigolo in alto a sinistra)
            .DrawString("DATA", Carattere_Titolo, Brushes.Black, 5, -larghezza_ret)
            .DrawString("PM", Carattere_Titolo, Brushes.Black, 5, -larghezza_ret * 2)
            .DrawString("MATRICOLA/E", Carattere_Titolo, Brushes.Black, 5, -larghezza_ret * 3)
            .DrawString("CLIENTE", Carattere_Titolo, Brushes.Black, 5, -larghezza_ret * 4)
            .DrawString("PROGETTO", Carattere_Titolo, Brushes.Black, 5, -larghezza_ret * 5)

            'Ripristino condizioni iniziali di scrittura
            .Restore(state)

            'Tabella base DATI
            .DrawRectangle(Penna, 1, altezza_ret_titoli, larghezza_ret, altezza_ret_dati)
            .DrawRectangle(Penna, larghezza_ret * 1, altezza_ret_titoli, larghezza_ret * 1, altezza_ret_dati)
            .DrawRectangle(Penna, larghezza_ret * 2, altezza_ret_titoli, larghezza_ret * 2, altezza_ret_dati)
            .DrawRectangle(Penna, larghezza_ret * 3, altezza_ret_titoli, larghezza_ret * 3, altezza_ret_dati)
            .DrawRectangle(Penna, larghezza_ret * 4, altezza_ret_titoli, larghezza_ret * 4, altezza_ret_dati)

            'Scrittura base DATI VERTICALI

            'Rotazione testo in gradi
            .RotateTransform(+90)

            ' Scritta in verticale (ATTENZIONE Il rifermiento delle scritte è lo spigolo in alto a sinistra)
            .DrawString(Data, Carattere_Dati, Brushes.Black, altezza_ret_titoli + 5, -larghezza_ret)
            .DrawString(PM, Carattere_Dati, Brushes.Black, altezza_ret_titoli + 5, -larghezza_ret * 2)
            .DrawString(matricole, Carattere_Dati, Brushes.Black, altezza_ret_titoli + 5, -larghezza_ret * 3)
            .DrawString(cliente, Carattere_Dati, Brushes.Black, altezza_ret_titoli + 5, -larghezza_ret * 4)
            .DrawString(progetto, Carattere_Dati, Brushes.Black, altezza_ret_titoli + 5, -larghezza_ret * 5)

            'Ripristino condizioni iniziali di scrittura
            .Restore(state)

        End With
    End Sub



    Sub Inserimento_dipendenti_acquisti(par_combobox As ComboBox)
        par_combobox.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "select t0.empid, concat(t0.lastname,' ' ,t0.firstname) as 'Acq'
from [TIRELLI_40].[dbo].ohem t0
where t0.active='Y'
order by concat(t0.lastname,' ' ,t0.firstname)"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_acquisti(Indice) = cmd_SAP_reader("empid")
            par_combobox.Items.Add(cmd_SAP_reader("Acq"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If ComboBox2.SelectedIndex < 0 Then
            MsgBox("Assegnare un responsabile acquisti")
            Return
        End If
        Aggiorna_RESP_acquisti()
        MsgBox("Responsabile aggiornato con successo")
    End Sub

    Sub Aggiorna_RESP_acquisti()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Dim destinazione = Replace(LinkLabel1.Text, "'", " ")
        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE t0 SET t0.U_resp_acquisti='" & Elenco_acquisti(ComboBox2.SelectedIndex) & "' 
from opmg t0
where t0.docnum=" & Replace(Label4.Text, "PJ", "") & ""
        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click



        ' Controlla se la TabPage 5 (indice 4) è selezionata
        'MsgBox(TabControl1.SelectedIndex)
        'MsgBox(TabControl1.SelectedIndex)
        'End
        Dim STATO_SCHEDA As String = ""
        If ComboBox70.SelectedIndex >= 0 Then
            STATO_SCHEDA = Elenco_stati_progetto(ComboBox70.SelectedIndex)
        End If

        If TabControl1.SelectedIndex = 1 Then

        End If
        inserisci_valori_progetto_scheda_tecnica(Replace(Label4.Text, "PJ", ""), codice_progetto)
        inserisci_numero_nuova_revisione_progetto(Replace(Label4.Text, "PJ", ""), codice_progetto, STATO_SCHEDA)
        elenca_revisioni_progetto(Replace(Label4.Text, "PJ", ""))
        trova_ultima_revisione_progetto(Replace(Label4.Text, "PJ", ""), codice_progetto)
        Label7.Text = numero_ultima_revisione
            ' Label7.Text = numero_ultima_revisione + 1
            MsgBox("Revisione N° " & numero_ultima_revisione & " inserita con successo")
        'End If





    End Sub

    Sub inserisci_valori_analisi_dei_rischi(par_numero_progetto As Integer, par_codice_progetto As String)


        trova_ultima_revisione_progetto(par_numero_progetto, par_codice_progetto)


        Dim Imballo As Integer = If(CheckBox1.Checked, 1, 0)
        Dim Prezzo_Imballo As String = If(String.IsNullOrWhiteSpace(TextBox1.Text), "NULL", Replace(TextBox1.Text, ",", "."))
        Dim Trasporto As Integer = If(CheckBox2.Checked, 1, 0)
        Dim Prezzo_Trasporto As String = If(String.IsNullOrWhiteSpace(TextBox2.Text), "NULL", Replace(TextBox2.Text, ",", "."))
        Dim Installazione As Integer = If(CheckBox3.Checked, 1, 0)
        Dim Prezzo_Installazione As String = If(String.IsNullOrWhiteSpace(TextBox3.Text), "NULL", Replace(TextBox3.Text, ",", "."))
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "INSERT INTO [Tirelli_40].[dbo].[Scheda_Tecnica_valori_progetto] " &
    "([n_progetto], [Rev], [Note], [Imballo], [Prezzo_Imballo], [Trasporto], [Prezzo_Trasporto], [Installazione], [Prezzo_Installazione]) " &
    "VALUES (" & par_numero_progetto & ", " &
    numero_ultima_revisione & " + 1, '" &
    Replace(RichTextBox1.Text, "'", "''") & "', " &
    Imballo & ", " &
    Prezzo_Imballo & ", " &
    Trasporto & ", " &
    Prezzo_Trasporto & ", " &
    Installazione & ", " &
    Prezzo_Installazione & ")"



        CMD_SAP_3.ExecuteNonQuery()

        Cnn3.Close()


    End Sub

    Sub inserisci_valori_progetto_scheda_tecnica(par_numero_progetto As Integer, par_codice_progetto As String)


        trova_ultima_revisione_progetto(par_numero_progetto, par_codice_progetto)


        Dim Imballo As Integer = If(CheckBox1.Checked, 1, 0)
        Dim Prezzo_Imballo As String = If(String.IsNullOrWhiteSpace(TextBox1.Text), "NULL", Replace(TextBox1.Text, ",", "."))
        Dim Trasporto As Integer = If(CheckBox2.Checked, 1, 0)
        Dim Prezzo_Trasporto As String = If(String.IsNullOrWhiteSpace(TextBox2.Text), "NULL", Replace(TextBox2.Text, ",", "."))
        Dim Installazione As Integer = If(CheckBox3.Checked, 1, 0)
        Dim Prezzo_Installazione As String = If(String.IsNullOrWhiteSpace(TextBox3.Text), "NULL", Replace(TextBox3.Text, ",", "."))


        Dim id_rischio_cliente As Integer = If(ComboBox3.SelectedIndex < 0, 0, ComboBox3.Text)
        Dim id_rischio_geografico As Integer = If(ComboBox4.SelectedIndex < 0, 0, ComboBox4.Text)
        Dim indice_rischio_riempimento As Integer = If(ComboBox6.SelectedIndex < 0, 0, ComboBox6.Text)
        Dim indice_rischio_handling_tappo As Integer = If(ComboBox5.SelectedIndex < 0, 0, ComboBox5.Text)
        Dim indice_rischio_handling_bottiglia As Integer = If(ComboBox7.SelectedIndex < 0, 0, ComboBox7.Text)
        Dim indice_rischio_handling_etichettatura As Integer = If(ComboBox8.SelectedIndex < 0, 0, ComboBox8.Text)
        Dim indice_rischio_cf_tipologia As Integer = If(ComboBox9.SelectedIndex < 0, 0, ComboBox9.Text)
        Dim indice_rischio_complessita_fornitura As Integer = If(ComboBox10.SelectedIndex < 0, 0, ComboBox10.Text)
        Dim indice_rischio_ambiente_lavoro As Integer = If(ComboBox11.SelectedIndex < 0, 0, ComboBox11.Text)

        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "INSERT INTO [Tirelli_40].[dbo].[Scheda_Tecnica_valori_progetto] " &
    "([n_progetto], [Rev], [Note], [Imballo], [Prezzo_Imballo], [Trasporto], [Prezzo_Trasporto], [Installazione], [Prezzo_Installazione]
    ,[id_rischio_cliente]

           ,[id_rischio_geografico]

           ,[indice_rischio_riempimento]

           ,[indice_rischio_handling_tappo]

           ,[indice_rischio_handling_bottiglia]

           ,[indice_rischio_handling_etichettatura]

           ,[indice_rischio_cf_tipologia]

           ,[indice_rischio_complessita_fornitura]

           ,[indice_rischio_ambiente_lavoro]

           ,[indice_rischio_vendita]

           ,[indice_rischio_tecnico]
           ,[indice_rischio_progetto]
           ,[livello_rischio_progetto]
           ,[livello_rischio_totale]
           ,[note_progetto]
    ,codice_progetto) " &
    "VALUES (" & par_numero_progetto & ", " &
    numero_ultima_revisione & " + 1, '" &
    Replace(RichTextBox1.Text, "'", "''") & "', " &
    Imballo & ", " &
    Prezzo_Imballo & ", " &
    Trasporto & ", " &
    Prezzo_Trasporto & ", " &
    Installazione & ", " &
    Prezzo_Installazione & "
, " & id_rischio_cliente & "
, " & id_rischio_geografico & "
, " & indice_rischio_riempimento & "
, " & indice_rischio_handling_tappo & "
, " & indice_rischio_handling_bottiglia & "
, " & indice_rischio_handling_etichettatura & "
, " & indice_rischio_cf_tipologia & "
, " & indice_rischio_complessita_fornitura & "
, " & indice_rischio_ambiente_lavoro & "
, '" & RichTextBox3.Text & "'
, '" & RichTextBox4.Text & "'
, '" & RichTextBox5.Text & "'
, '" & RichTextBox7.Text & "'
, '" & RichTextBox8.Text & "'
, '" & RichTextBox6.Text & "'
,'" & par_codice_progetto & "'
)"

        CMD_SAP_3.ExecuteNonQuery()

        CMD_SAP_3.CommandText = "delete [Tirelli_40].[dbo].[Scheda_Tecnica_valori_progetto_last] where [n_progetto]= " & par_numero_progetto & "

INSERT INTO [dbo].[Scheda_Tecnica_valori_progetto_last]
             ([n_progetto]
,codice_progetto
           ,[Rev]
           ,[livello_rischio_totale]
           ,[data])
     VALUES
          (" & par_numero_progetto & "
          ,'" & par_codice_progetto & "'
           ," & numero_ultima_revisione & " + 1
           ,'" & RichTextBox8.Text & "'
           ,getdate())
"



        CMD_SAP_3.ExecuteNonQuery()

        Cnn3.Close()


    End Sub



    Sub Inserimento_STATO_PROGETTI(par_combobox As ComboBox)
        par_combobox.Items.Clear()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[id] , T0.[nome] 
        FROM [Tirelli_40].dbo.scheda_tecnica_stato_progetto T0
 order by T0.[ordine]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_stati_progetto(Indice) = cmd_SAP_reader("id")
            par_combobox.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub 'Inserisco le risorse nella combo box
    Private Sub DataGridView_revisione_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_revisione.CellClick
        N_rev_visualizza = DataGridView_revisione.Rows(e.RowIndex).Cells(columnName:="N_Rev").Value
    End Sub



    Sub trova_ultima_revisione_progetto(par_numero_progetto As String, par_codice_progetto As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "
               Select  coalesce(t11.numero,0) as 'Ultima_rev',t11.utente, coalesce(CONCAT(T12.LASTNAME,' ',T12.FIRSTNAME),'-') as 'Nome_utente', t11.data,t11.ora
,coalesce(t13.nome,'') as 'Stato_scheda'
from
(
SELECT MAX(t0.id) as 'Ultimo_id'
     
  FROM [Tirelli_40].[dbo].[Scheda_tecnica_revisioni_progetto] t0
where t0.n_progetto ='" & par_numero_progetto & "' or t0.codice_progetto='" & par_codice_progetto & "'

)
as t10 left join [Tirelli_40].[dbo].[Scheda_tecnica_revisioni_progetto] t11 on t10.ultimo_id=t11.id
left join [TIRELLI_40].[dbo].ohem t12 on t12.empid=t11.utente
left join [TIRELLI_40].[dbo].Scheda_Tecnica_stato_progetto t13 on t11.stato=t13.id

"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            numero_ultima_revisione = cmd_SAP_reader("ultima_Rev")
            ComboBox70.Text = cmd_SAP_reader("Stato_scheda")

            If Not cmd_SAP_reader("Data") Is System.DBNull.Value Then
                Label8.Text = cmd_SAP_reader("Data") & " | " & cmd_SAP_reader("ORA")
            Else
                Label8.Text = "-"
            End If


            Label9.Text = cmd_SAP_reader("Nome_utente")

        Else
            numero_ultima_revisione = 0
            Label8.Text = "-"
            Label9.Text = "-"
        End If

        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub 'Inserisco le risorse nella combo box


    Sub inserisci_numero_nuova_revisione_progetto(par_numero_progetto As Integer, par_codice_progetto As String, par_stato As String)


        trova_ultima_revisione_progetto(par_numero_progetto, par_codice_progetto)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "

INSERT INTO [Tirelli_40].[dbo].[Scheda_tecnica_revisioni_progetto]
           ([n_progetto]
,codice_progetto
           ,[Numero]
           ,[utente]
           ,[Data]
           ,[ora]
,[stato])
     VALUES
           ('" & par_numero_progetto & "'
           ,'" & codice_progetto & "'
           ," & numero_ultima_revisione & "+1
           ," & Homepage.ID_SALVATO & "
,getdate()
           ,convert(varchar, getdate(), 108)
,'" & par_stato & "')"


        CMD_SAP_3.ExecuteNonQuery()

        Cnn3.Close()


    End Sub

    Sub elenca_revisioni_progetto(par_numero_progetto As Integer)
        Dim Cnn1 As New SqlConnection
        DataGridView_revisione.Rows.Clear()
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.[ID]
      ,t0.[n_progetto]
	  ,t0.numero
      ,t0.[utente]
      ,t0.[Data]
      ,t0.[ora]
,coalesce(t2.nome,'') as 'Stato'
,CONCAT(T1.LASTNAME,' ',T1.FIRSTNAME) as 'Nome_utente'
  FROM [Tirelli_40].[dbo].[Scheda_tecnica_revisioni_progetto] t0
left join [TIRELLI_40].[dbo].ohem t1 on t1.empid=t0.utente
left join [Tirelli_40].[dbo].[Scheda_Tecnica_stato_progetto] t2 on t2.id=t0.stato
where t0.[n_progetto]='" & par_numero_progetto & "'
order by t0.ID DESC
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            DataGridView_revisione.Rows.Add(cmd_SAP_reader_2("numero"), cmd_SAP_reader_2("utente"), cmd_SAP_reader_2("nome_utente"), cmd_SAP_reader_2("data"), cmd_SAP_reader_2("ora"), cmd_SAP_reader_2("Stato"))


        Loop


        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Try
            riempi_scheda_tecnica_progetto(Replace(Label4.Text, "PJ", ""), N_rev_visualizza, "Progetto")
            Label7.Text = N_rev_visualizza
            MsgBox("Stai ora visualizzando la revisione " & N_rev_visualizza)
        Catch ex As Exception
            MsgBox("Selezionare un numero di revisione")
        End Try
    End Sub


    Private Async Sub tabpage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter


        inizializza_scheda_tecnica_progetto(Replace(Label4.Text, "PJ", ""), codice_progetto)


    End Sub

    '   Private Async Sub tabpage4_Click(sender As Object, e As EventArgs) Handles TabPage4.Enter


    '    cardini_commesse()



    'End Sub


    Private Async Sub tabpage3_Click(sender As Object, e As EventArgs) Handles TabPage3.Enter


        carica_appunti(Replace(Label4.Text, "PJ", ""), "PROGETTO", DataGridView5, "", Homepage.ID_SALVATO)


    End Sub

    Private Sub DataGridView1_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Rows(e.RowIndex).Cells("Collaudo").Value = 1 Then
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
        End If
        If DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Value = "M" Then
            DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Style.BackColor = Color.Orange
        ElseIf DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Value = "CDS" Then
            DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Style.BackColor = Color.YellowGreen
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click


        inserisci_nuovo_appunto(Homepage.ID_SALVATO, "PROGETTO", "", Replace(Label4.Text, "PJ", ""), Replace(RichTextBox2.Text, "'", " "), False, False, False)


        carica_appunti(Replace(Label4.Text, "PJ", ""), "PROGETTO", DataGridView5, "", Homepage.ID_SALVATO)
        RichTextBox2.Text = ""

    End Sub

    Sub inserisci_nuovo_appunto(par_dipendente As Integer, par_tipo As String, par_commessa As String, par_n_progetto As Integer, par_contenuto As String, par_grassetto As Boolean, par_corsivo As Boolean, par_risolto As Boolean)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand
        CMD_SAP_5.Connection = CNN6
        CMD_SAP_5.CommandText = "INSERT INTO [TIRELLI_40].[DBO].dati_mancanti_progetto
       ([Dipendente]
       ,[Data]
       ,[Ora]
,[tipo]
       ,[Commessa]
,[progetto]
       ,[Contenuto]
       ,[Grassetto]
       ,[Corsivo]
       ,[risolto])
 VALUES
       (@Dipendente
       ,getdate()
       ,@Ora
,@tipo
       ,@Commessa
,@progetto
       ,@Contenuto
       ,@Grassetto
       ,@Corsivo
       ,@risolto)"

        ' Aggiunta dei parametri
        CMD_SAP_5.Parameters.AddWithValue("@Dipendente", par_dipendente)

        CMD_SAP_5.Parameters.AddWithValue("@Ora", DateTime.Now.TimeOfDay) ' Ora attuale del sistema
        CMD_SAP_5.Parameters.AddWithValue("@Commessa", par_commessa)
        CMD_SAP_5.Parameters.AddWithValue("@Contenuto", par_contenuto)
        '  CMD_SAP_5.Parameters.AddWithValue("@Font", par_font)
        CMD_SAP_5.Parameters.AddWithValue("@Grassetto", par_grassetto)
        CMD_SAP_5.Parameters.AddWithValue("@Corsivo", par_corsivo)
        CMD_SAP_5.Parameters.AddWithValue("@risolto", par_risolto)
        CMD_SAP_5.Parameters.AddWithValue("@tipo", par_tipo)
        CMD_SAP_5.Parameters.AddWithValue("@progetto", par_n_progetto)

        CMD_SAP_5.ExecuteNonQuery()

        CNN6.Close()
    End Sub

    Sub carica_appunti(par_commessa As String, par_tipo As String, par_Datagridview As DataGridView, par_annulla_filtro As String, par_dipendente As Integer)

        Dim filtro_dipendente As String
        If par_dipendente = 0 Then
            filtro_dipendente = ""
        Else
            filtro_dipendente = " and t0.dipendente = " & par_dipendente & ""
        End If

        ' par_richtextbox.Text = ""
        par_Datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT  t0.[ID]
      ,t0.[Dipendente]
	  ,concat(t1.lastname, ' ', substring(t1.firstname,1,1)) as 'Nome'
      ,t0.[Data]
      ,t0.[Ora]
      ,t0.[progetto]
,COALESCE(t2.[U_FINAL_cUSTOMER_NAME],'') AS 'U_FINAL_CUSTOMER_NAME'
      ,t0.[Contenuto]
      ,t0.[Font]
      ,t0.[Grassetto]
      ,t0.[Corsivo]
      ,t0.[risolto]
  FROM [TIRELLI_40].[DBO].dati_mancanti_progetto t0 
  left join [TIRELLI_40].[dbo].ohem t1 on t0.dipendente=t1.empid
left join [TIRELLISRLDB].[dbo].oitm t2 on t2.itemcode=t0.commessa
where  cast(t0.[Progetto] as varchar) ='" & par_commessa & "' and t0.[Tipo]='" & par_tipo & "' " & par_annulla_filtro & filtro_dipendente & " 
  order by t0.[Commessa],t0.[ID]
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            par_Datagridview.Rows.Add(cmd_SAP_reader_2("id"), cmd_SAP_reader_2("progetto"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("data"), cmd_SAP_reader_2("contenuto"), cmd_SAP_reader_2("Risolto"), cmd_SAP_reader_2("U_FINAL_CUSTOMER_NAME"))
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        par_Datagridview.ClearSelection()

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        carica_appunti(Replace(Label4.Text, "PJ", ""), "PROGETTO", DataGridView5, "", Homepage.ID_SALVATO)
    End Sub

    Private Sub CancellaRigaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CancellaRigaToolStripMenuItem.Click
        'If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Trasferito").Value > 0 Then
        '    MsgBox("Impossibile cancellare riga di un codice che risulta TRASFERITO")
        'Else
        Dim PAR_DATAGRIDVIEW As DataGridView
        PAR_DATAGRIDVIEW = DataGridView5

        ' Supponendo che COLONNAID sia il nome della colonna che vuoi usare per cancellare il commento
        Dim COLONNAID As String = "ID_"
        Dim selectedRow As DataGridViewRow = PAR_DATAGRIDVIEW.CurrentRow

        ' Verifica che ci sia una riga selezionata prima di procedere
        If selectedRow IsNot Nothing Then
            ' Chiede conferma all'utente se vuole cancellare il commento
            Dim result As DialogResult = MessageBox.Show("Sei sicuro di voler cancellare il commento?" & vbCrLf & selectedRow.Cells("Commento_").Value, "Conferma Cancellazione", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            ' Se l'utente conferma, procedi con la cancellazione
            If result = DialogResult.Yes Then
                ' Passa l'ID della riga alla funzione cancella_commento
                cancella_commento(selectedRow.Cells(COLONNAID).Value)

                ' Rimuovi la riga selezionata
                PAR_DATAGRIDVIEW.Rows.RemoveAt(selectedRow.Index)
            End If
        Else
            MessageBox.Show("Seleziona una riga prima di cancellarla.")
        End If
    End Sub

    Public Sub mail_dati_mancanti(par_lista_distribuzione As String)

        Dim emailBody As New StringBuilder()
        Dim destinatariUnici As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        ' Variabili a livello di mail_progetti
        Dim colori As String() = {"blue", "green", "red", "purple", "orange"}
        Dim colorIndex As Integer = 0
        Dim coloriTag As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)


        Using Cnn1 As New SqlConnection(Homepage.sap_tirelli)
            Cnn1.Open()

            ' --- Query progetti ---
            Using CMD_SAP_2 As New SqlCommand("
         select t30.progetto_commessa, t30.cardname,
          t30.Cliente_finale,
          t30.name, t30.data_min
		  ,t30.Rev_Max , t31.Stato, t32.Nome as 'Nome_stato', t32.Ordine
  from
  (
  SELECT t20.progetto_commessa, t21.cardname,
          COALESCE(t22.cardname,'') AS Cliente_finale,
          t21.name, t20.data_min
		  ,max(t23.numero) as 'Rev_Max'
   FROM (
       SELECT t10.progetto_commessa, MIN(t10.DOCDUEDATE) AS data_min
       FROM (
           SELECT T0.ID, T0.tipo,
                  CAST(T0.Commessa AS VARCHAR) AS Commessa,
                  CAST(T0.progetto AS VARCHAR) AS Progetto,
                  CASE WHEN T0.tipo='Progetto'
                       THEN CAST(T0.progetto AS VARCHAR)
                       ELSE CAST(T2.u_progetto AS VARCHAR)
                  END AS Progetto_commessa,
                  CONCAT(T1.lastname, ' ', SUBSTRING(T1.firstname,1,1)) AS Nome,
                  T0.[Data], COALESCE(T2.itemname,'') AS Nome_macchina,
                  COALESCE(T2.[U_Final_customer_name],'') AS Cliente_finale,
                  T0.[Contenuto], T0.[risolto], A.DOCDUEDATE
           FROM [TIRELLI_40].[DBO].dati_mancanti_progetto T0
           LEFT JOIN [TIRELLI_40].[dbo].OHEM T1 ON T0.dipendente=T1.empid
           LEFT JOIN OITM T2 ON T2.itemcode=T0.commessa
           LEFT JOIN (
               SELECT T10.ITEMCODE, T12.DOCNUM, T12.DOCDUEDATE
               FROM (
                   SELECT MIN(T1.DocEntry) AS DOCENTRY, T0.ITEMCODE
                   FROM RDR1 T0
                   INNER JOIN ORDR T1 ON T0.DocEntry=T1.DocEntry
                   WHERE T1.CANCELED<>'Y' AND SUBSTRING(T0.ITEMCODE,1,1)='M'
                   GROUP BY T0.ITEMCODE
               ) T10
               LEFT JOIN RDR1 T11 ON T11.DocEntry=T10.DocEntry AND T10.ITEMCODE=T11.ITEMCODE
               LEFT JOIN ORDR T12 ON T12.DocEntry=T11.DocEntry
           ) A ON A.ItemCode=T0.Commessa
           WHERE 1=1 
       ) t10
       GROUP BY t10.progetto_commessa
   ) t20
   LEFT JOIN OPMG t21 ON CAST(t20.Progetto_commessa AS VARCHAR)=CAST(T21.DocNum AS VARCHAR)
   LEFT JOIN OCRD t22 ON t22.CardCode=t21.U_Codice_cliente_finale
   left join [Tirelli_40].[dbo].[Scheda_tecnica_revisioni_progetto] t23 on t23.n_progetto=t20.Progetto_commessa
  group by 
  t20.progetto_commessa, t21.cardname,
          COALESCE(t22.cardname,'') ,
          t21.name, t20.data_min
		  )
		  as t30
		  left join [Tirelli_40].[dbo].[Scheda_tecnica_revisioni_progetto] t31 on t31.n_progetto=t30.Progetto_commessa and t30.rev_max=t31.Numero
		  left join [Tirelli_40].[dbo].[Scheda_Tecnica_stato_progetto] t32 on t32.ID=t31.Stato
  ORDER BY t32.ordine, t30.data_min", Cnn1)

                Using rdr As SqlDataReader = CMD_SAP_2.ExecuteReader()
                    Dim currentProgetto As String = String.Empty
                    Dim currentOrdine As Integer = -1 ' <- Per tracciare il cambio di stato

                    While rdr.Read()
                        Dim ordine As Integer = If(IsDBNull(rdr("Ordine")), -1, CInt(rdr("Ordine")))
                        Dim nomeStato As String = rdr("Nome_stato").ToString()

                        ' --- Se cambia lo stato (ordine) aggiungi titolo di sezione ---
                        If ordine <> currentOrdine Then
                            currentOrdine = ordine
                            emailBody.AppendLine("<br /><h3 style='color:#003366;'>" & nomeStato & "</h3><br />")
                        End If

                        ' --- Poi gestisci il progetto ---
                        Dim progetto As String = "Progetto " & rdr("progetto_commessa").ToString() & " " &
                                                 rdr("cardname").ToString() & " " &
                                                 rdr("Cliente_finale").ToString() & " " &
                                                 rdr("name").ToString()

                        If currentProgetto <> progetto Then
                            If currentProgetto <> String.Empty Then emailBody.AppendLine("<br />")
                            currentProgetto = progetto
                            emailBody.AppendLine("<b>" & progetto & "</b><br /><br />")
                        End If

                        ' --- Chiamata della sub ---
                        trova_appunti(rdr("progetto_commessa").ToString(), emailBody, destinatariUnici, colori, colorIndex, coloriTag)
                    End While
                End Using
            End Using

            ' --- Aggiungi destinatari dalla lista di distribuzione ---
            Using CMD_Email As New SqlCommand("
            SELECT Mail
            FROM [Tirelli_40].[dbo].[Lista_distribuzione_riunioni]
            WHERE Nome_lista=@lista", Cnn1)

                CMD_Email.Parameters.AddWithValue("@lista", par_lista_distribuzione)

                Using rdrMail As SqlDataReader = CMD_Email.ExecuteReader()
                    While rdrMail.Read()
                        destinatariUnici.Add(rdrMail("Mail").ToString())
                    End While
                End Using
            End Using
        End Using  ' <-- qui chiudiamo l'uso della connessione

        ' --- Prepara lista finale destinatari ---
        Dim emailList As String = String.Join(";", destinatariUnici)

        ' --- Invia tramite Outlook ---
        Dim outlookApp As New Outlook.Application
        Dim mailItem As Outlook.MailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem)

        With mailItem
            .Subject = "Punti aperti"
            .HTMLBody = emailBody.ToString()
            .To = emailList
            .Display() ' oppure .Send()
        End With
    End Sub

    Sub trova_appunti(par_progetto As String,
                  ByRef emailBody As StringBuilder,
                  ByRef destinatariUnici As HashSet(Of String),
                  colori() As String, ByRef colorIndex As Integer,
                  coloriTag As Dictionary(Of String, String))

        '  Dim filtro_Dipendente As String = If(par_id_dipendente = 0, "", " AND t0.dipendente = " & par_id_dipendente)

        Using Cnn1 As New SqlConnection(Homepage.sap_tirelli)
            Cnn1.Open()

            Using CMD_SAP_2 As New SqlCommand("
            SELECT T0.Commessa, COALESCE(T2.itemname,'') AS Nome_macchina, 
                   COALESCE(T2.[U_Final_customer_name],'') AS Cliente_finale,
                   T0.Contenuto, T0.risolto
            FROM [TIRELLI_40].[DBO].dati_mancanti_progetto T0
            LEFT JOIN OITM T2 ON T2.itemcode=T0.commessa
            WHERE (CASE WHEN T0.tipo='Progetto' THEN CAST(T0.progetto AS VARCHAR) ELSE CAST(T2.u_progetto AS VARCHAR) END) = @progetto
            
            ORDER BY T0.ID", Cnn1)

                CMD_SAP_2.Parameters.AddWithValue("@progetto", par_progetto)

                Using rdr As SqlDataReader = CMD_SAP_2.ExecuteReader()
                    Dim currentCommessa As String = String.Empty

                    While rdr.Read()
                        Dim commessa As String = rdr("Commessa").ToString() & " " &
                                             rdr("Nome_macchina").ToString() & " " &
                                             rdr("Cliente_finale").ToString()

                        If currentCommessa <> commessa Then
                            If currentCommessa <> String.Empty Then emailBody.AppendLine("<br />")
                            currentCommessa = commessa
                            '   emailBody.AppendLine("<b>" & commessa & "</b><br />")

                            emailBody.AppendLine("<b><div style='margin-left:30px;'>" & commessa & "<br /><br /></div> </b>")
                        End If

                        Dim contenuto As String = rdr("Contenuto").ToString()
                        If Convert.ToBoolean(rdr("risolto")) Then contenuto = "<s>" & contenuto & "</s>"

                        ' --- Gestione tag @ senza ByRef nella lambda ---
                        Dim regex As New Regex("@\S+")
                        Dim matches = regex.Matches(contenuto)
                        For Each match As Match In matches
                            Dim nome As String = match.Value.Substring(1).ToLower()
                            Dim indirizzo As String = nome & "@tirelli.net"
                            destinatariUnici.Add(indirizzo)

                            Dim coloreCorrente As String
                            If Not coloriTag.TryGetValue(nome, coloreCorrente) Then
                                coloreCorrente = colori(colorIndex)
                                coloriTag(nome) = coloreCorrente
                                colorIndex = (colorIndex + 1) Mod colori.Length
                            End If

                            contenuto = contenuto.Replace(match.Value,
                                    "<span style='color:" & coloreCorrente & "; font-weight:bold;'>" & match.Value & "</span>")
                        Next

                        emailBody.AppendLine("<div style='margin-left:60px;'>" & contenuto & "<br /><br /></div>")
                    End While
                End Using
            End Using
        End Using
    End Sub

    Sub cancella_commento(par_ID As Integer)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand
        CMD_SAP_5.Connection = CNN6
        CMD_SAP_5.CommandText = "DELETE [TIRELLI_40].[DBO].dati_mancanti_progetto

      
WHERE ID=@ID"

        ' Aggiunta dei parametri

        CMD_SAP_5.Parameters.AddWithValue("@ID", par_ID)



        CMD_SAP_5.ExecuteNonQuery()

        CNN6.Close()
    End Sub

    Private Sub CompletatoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CompletatoToolStripMenuItem.Click
        'If DataGridView_ODP.Rows(DATAGRIDVIEW_odp_RIGA).Cells(columnName:="Trasferito").Value > 0 Then
        '    MsgBox("Impossibile cancellare riga di un codice che risulta TRASFERITO")
        'Else
        Dim PAR_DATAGRIDVIEW As DataGridView
        PAR_DATAGRIDVIEW = DataGridView5

        ' Supponendo che COLONNAID sia il nome della colonna che vuoi usare per cancellare il commento
        Dim COLONNAID As String = "ID_"
        Dim selectedRow As DataGridViewRow = PAR_DATAGRIDVIEW.CurrentRow

        ' Verifica che ci sia una riga selezionata prima di procedere
        If selectedRow IsNot Nothing Then
            ' Chiede conferma all'utente se vuole cancellare il commento

            cambia_stato(selectedRow.Cells(COLONNAID).Value)
            carica_appunti(Replace(Label4.Text, "PJ", ""), "PROGETTO", DataGridView5, "", Homepage.ID_SALVATO)


        Else
            MessageBox.Show("Seleziona una riga prima di CAMBIARNE STATO")
        End If
    End Sub

    Sub cambia_stato(par_ID As Integer)
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP_5 As New SqlCommand
        CMD_SAP_5.Connection = CNN6
        CMD_SAP_5.CommandText = "UPDATE [TIRELLI_40].[DBO].dati_mancanti_progetto
SET  RISOLTO=CASE WHEN RISOLTO='FALSE' THEN 'TRUE' ELSE 'FALSE' END

      
WHERE ID=@ID"

        ' Aggiunta dei parametri

        CMD_SAP_5.Parameters.AddWithValue("@ID", par_ID)



        CMD_SAP_5.ExecuteNonQuery()

        CNN6.Close()
    End Sub

    Private Sub DataGridView5_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView5.CellContentClick

    End Sub

    Private Sub DataGridView5_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView5.CellFormatting
        Dim par_datagridview As DataGridView = DataGridView5
        ' Verifica se la colonna "stato" è presente (sostituisci "stato" con il nome corretto della colonna)
        Dim statoIndex As Integer = par_datagridview.Columns("stato___").Index

        ' Verifica se siamo in una riga valida e non è una riga nuova
        If e.RowIndex >= 0 AndAlso Not par_datagridview.Rows(e.RowIndex).IsNewRow Then
            ' Controlla il valore della colonna "stato"
            Dim statoValue As Boolean = Convert.ToBoolean(par_datagridview.Rows(e.RowIndex).Cells(statoIndex).Value)

            ' Se il valore è "True", applica il font barrato a tutta la riga
            If statoValue Then
                For Each cell As DataGridViewCell In par_datagridview.Rows(e.RowIndex).Cells
                    cell.Style.Font = New Font(par_datagridview.Font, FontStyle.Strikeout)
                Next
            Else
                ' Rimuove il font barrato se non è "True"
                For Each cell As DataGridViewCell In par_datagridview.Rows(e.RowIndex).Cells
                    cell.Style.Font = New Font(par_datagridview.Font, FontStyle.Regular)
                Next
            End If
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        mail_dati_mancanti("")
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click

    End Sub

    Private Sub ComboBox70_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox70.SelectedIndexChanged
        Dim coloreBordo As Color = Color.Transparent
        Dim colorefont As Color = Color.Black

        Select Case ComboBox70.Text
            Case "REV P"
                coloreBordo = Color.Purple
                colorefont = Color.White
            Case "REV 0"
                coloreBordo = Color.Yellow
            Case "REV A"
                coloreBordo = Color.Lime
            Case "SOSPESO"
                coloreBordo = Color.Aqua
            Case "CHIUSO"
                coloreBordo = Color.Gray
            Case Else
                coloreBordo = SystemColors.Control
        End Select

        TableLayoutPanel11.BackColor = coloreBordo
        TableLayoutPanel11.ForeColor = colorefont
    End Sub

    Private Sub TableLayoutPanel17_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel17.Paint

    End Sub

    Private Sub FlowLayoutPanel3_Paint(sender As Object, e As PaintEventArgs) Handles FlowLayoutPanel3.Paint

    End Sub

    Private Sub DataGridView_revisione_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_revisione.CellContentClick

    End Sub

    Private Sub DataGridView_revisione_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_revisione.CellFormatting
        Dim PAR_DATAGRIDVIEW As DataGridView = DataGridView_revisione

        If PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "REV P" Then

            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Color.Purple
            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:="Stato").Style.ForeColor = Color.White

        ElseIf PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "REV 0" Then

            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Color.Yellow

        ElseIf PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "REV A" Then

            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Color.Lime

        ElseIf PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "SOSPESO" Then

            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Color.Aqua

        ElseIf PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "CHIUSO" Then

            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Color.Gray
        Else

            PAR_DATAGRIDVIEW.Rows(e.RowIndex).Cells(columnName:="Stato").Style.BackColor = Nothing
        End If
    End Sub



    Private Sub TableLayoutPanel28_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel28.Paint

    End Sub



    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        AggiornaMaxDaCombo(ComboBox3, ComboBox4, RichTextBox3)
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        AggiornaMaxDaCombo(ComboBox3, ComboBox4, RichTextBox3)
    End Sub

    Private Sub AggiornaMaxDaCombo(comboA As ComboBox, comboB As ComboBox, output As RichTextBox)
        ' Controlla che entrambe abbiano una selezione valida
        If comboA.SelectedIndex < 0 OrElse comboB.SelectedIndex < 0 Then
            output.Text = "N/A"
            Return
        End If

        ' Entrambe selezionate: confronta i valori numerici
        Dim valA As Double = Double.Parse(comboA.Text)
        Dim valB As Double = Double.Parse(comboB.Text)

        output.Text = Math.Max(valA, valB).ToString()
        AggiornaMedia()
    End Sub

    Private Sub AggiornaMaxDaCombo_pluri(comboA As ComboBox, comboB As ComboBox, comboC As ComboBox, comboD As ComboBox, comboE As ComboBox, comboF As ComboBox, comboG As ComboBox, output As RichTextBox)
        ' Controlla che tutte abbiano una selezione valida
        If comboA.SelectedIndex < 0 Or comboB.SelectedIndex < 0 Or comboC.SelectedIndex < 0 Or comboD.SelectedIndex < 0 Or comboE.SelectedIndex < 0 Or comboF.SelectedIndex < 0 Or comboG.SelectedIndex < 0 Then
            output.Text = "N/A"
            Return
        End If

        ' Converte i valori in numeri
        Dim valori As New List(Of Double)
        valori.Add(Double.Parse(comboA.Text))
        valori.Add(Double.Parse(comboB.Text))
        valori.Add(Double.Parse(comboC.Text))
        valori.Add(Double.Parse(comboD.Text))
        valori.Add(Double.Parse(comboE.Text))
        valori.Add(Double.Parse(comboF.Text))
        valori.Add(Double.Parse(comboG.Text))

        ' Prende il massimo
        output.Text = valori.Max().ToString()
        AggiornaMedia()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        AggiornaMaxDaCombo_pluri(ComboBox6, ComboBox5, ComboBox7, ComboBox8, ComboBox9, ComboBox10, ComboBox11, RichTextBox4)
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        AggiornaMaxDaCombo_pluri(ComboBox6, ComboBox5, ComboBox7, ComboBox8, ComboBox9, ComboBox10, ComboBox11, RichTextBox4)

    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        AggiornaMaxDaCombo_pluri(ComboBox6, ComboBox5, ComboBox7, ComboBox8, ComboBox9, ComboBox10, ComboBox11, RichTextBox4)

    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        AggiornaMaxDaCombo_pluri(ComboBox6, ComboBox5, ComboBox7, ComboBox8, ComboBox9, ComboBox10, ComboBox11, RichTextBox4)

    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        AggiornaMaxDaCombo_pluri(ComboBox6, ComboBox5, ComboBox7, ComboBox8, ComboBox9, ComboBox10, ComboBox11, RichTextBox4)

    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        AggiornaMaxDaCombo_pluri(ComboBox6, ComboBox5, ComboBox7, ComboBox8, ComboBox9, ComboBox10, ComboBox11, RichTextBox4)

    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedIndexChanged
        AggiornaMaxDaCombo_pluri(ComboBox6, ComboBox5, ComboBox7, ComboBox8, ComboBox9, ComboBox10, ComboBox11, RichTextBox4)

    End Sub
    Private Sub RichTextBox3_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox3.TextChanged
        Dim par_richtextbox As RichTextBox = RichTextBox3
        Dim pos As Integer = par_richtextbox.SelectionStart
        Dim length As Integer = par_richtextbox.SelectionLength

        ' Seleziona tutto
        par_richtextbox.SelectAll()
        par_richtextbox.SelectionAlignment = HorizontalAlignment.Center

        ' Ripristina la selezione del cursore
        par_richtextbox.SelectionStart = pos
        par_richtextbox.SelectionLength = length

        AggiornaMedia()



    End Sub

    Private Sub RichTextBox4_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox4.TextChanged
        Dim par_richtextbox As RichTextBox = RichTextBox4
        Dim pos As Integer = par_richtextbox.SelectionStart
        Dim length As Integer = par_richtextbox.SelectionLength

        ' Seleziona tutto
        par_richtextbox.SelectAll()
        par_richtextbox.SelectionAlignment = HorizontalAlignment.Center

        ' Ripristina la selezione del cursore
        par_richtextbox.SelectionStart = pos
        par_richtextbox.SelectionLength = length

        AggiornaMedia()

    End Sub

    Private Sub AggiornaMedia()
        If RichTextBox3.Text = "N/A" Or RichTextBox4.Text = "N/A" Then
            RichTextBox5.Text = "N/A"
        Else
            Dim val3 As Double
            Dim val4 As Double

            ' Prova a convertire i valori in double
            If Double.TryParse(RichTextBox3.Text, val3) AndAlso Double.TryParse(RichTextBox4.Text, val4) Then
                RichTextBox5.Text = ((val3 + val4) / 2).ToString()
            Else
                RichTextBox5.Text = "Errore"
            End If
        End If

        ' Se il testo non è un numero valido
        Dim val As Double
        Dim par_richtextbox As RichTextBox = RichTextBox5

        If par_richtextbox.Text = "N/A" OrElse Not Double.TryParse(par_richtextbox.Text, val) Then
            RichTextBox7.Text = "N/A"
            RichTextBox8.Text = "N/A"
            Return
        End If

        ' Determina il livello L1/L2/L3
        If val < 3 Then
            RichTextBox7.Text = "L1"
        ElseIf val > 3.99 Then
            RichTextBox7.Text = "L3"
        Else
            RichTextBox7.Text = "L2"
        End If

        ' Imposta RichTextBox8 uguale a RichTextBox7
        RichTextBox8.Text = RichTextBox7.Text

        ' Aggiungi "T" se RichTextBox4 > 3
        Dim val5 As Double
        If Double.TryParse(RichTextBox4.Text, val5) AndAlso val5 > 3 Then
            RichTextBox8.Text &= " T"
        End If
    End Sub

    Private Sub RichTextBox5_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox5.TextChanged
        Dim par_richtextbox As RichTextBox = RichTextBox5
        Dim pos As Integer = par_richtextbox.SelectionStart
        Dim length As Integer = par_richtextbox.SelectionLength

        ' Seleziona tutto e centra il testo
        par_richtextbox.SelectAll()
        par_richtextbox.SelectionAlignment = HorizontalAlignment.Center

        ' Ripristina la selezione del cursore
        par_richtextbox.SelectionStart = pos
        par_richtextbox.SelectionLength = length


    End Sub

    Private Sub RichTextBox7_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox7.TextChanged
        Dim par_richtextbox As RichTextBox = RichTextBox7
        Dim pos As Integer = par_richtextbox.SelectionStart
        Dim length As Integer = par_richtextbox.SelectionLength

        ' Seleziona tutto
        par_richtextbox.SelectAll()
        par_richtextbox.SelectionAlignment = HorizontalAlignment.Center

        ' Ripristina la selezione del cursore
        par_richtextbox.SelectionStart = pos
        par_richtextbox.SelectionLength = length
    End Sub

    Private Sub TableLayoutPanel37_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel37.Paint

    End Sub

    Private Sub RichTextBox8_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox8.TextChanged
        Dim par_richtextbox As RichTextBox = RichTextBox8
        Dim pos As Integer = par_richtextbox.SelectionStart
        Dim length As Integer = par_richtextbox.SelectionLength

        ' Seleziona tutto
        par_richtextbox.SelectAll()
        par_richtextbox.SelectionAlignment = HorizontalAlignment.Center

        ' Ripristina la selezione del cursore
        par_richtextbox.SelectionStart = pos
        par_richtextbox.SelectionLength = length



    End Sub

    Private Sub FlowLayoutPanel4_Paint(sender As Object, e As PaintEventArgs) Handles FlowLayoutPanel4.Paint

    End Sub


    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
    Try
        ' 1. Controllo validità dati iniziali
        Dim nomeProgetto As String = LinkLabel1.Text
            Dim suffissoFile As String = "Progetto " & Label4.Text ' Il valore da appendere al nome del file

            If String.IsNullOrEmpty(nomeProgetto) Then 
            MsgBox("Selezionare prima un progetto.", MsgBoxStyle.Exclamation)
            Exit Sub 
        End If

        ' 2. Definizione percorsi cartelle
        Dim directoryProgetto As String = IO.Path.Combine(Homepage.percorso_progetti, nomeProgetto)
        Dim nomeSottoCartella As String = "Azioni contenitive"
        Dim percorsoCompletoSottoCartella As String = IO.Path.Combine(directoryProgetto, nomeSottoCartella)

        ' 3. Crea la cartella se non esiste
        If Not IO.Directory.Exists(percorsoCompletoSottoCartella) Then
            IO.Directory.CreateDirectory(percorsoCompletoSottoCartella)
        End If

            ' 4. Definizione file (Sorgente e Destinazione con nuovo nome)
            Dim fileSorgente As String = "\\tirfs01\TIRELLI\04-Leads\SENG\Azioni contenitive.xlsx"

            ' Nuovo nome file: "Azioni contenitive [TestoLabel4].xlsx"
            Dim nomeFileDestinazione As String = "Azioni contenitive " & suffissoFile & ".xlsx"
        Dim fileDestinazione As String = IO.Path.Combine(percorsoCompletoSottoCartella, nomeFileDestinazione)

        ' 5. Esecuzione copia e aggiornamento TreeView
        If IO.File.Exists(fileSorgente) Then
            IO.File.Copy(fileSorgente, fileDestinazione, True)
            
            ' Aggiorna la TreeView 3 passando il percorso relativo
            ' La tua funzione farà il resto: Homepage.percorso_progetti & "Progetto\Azioni contenitive"
            Dim percorsoRelativoPerTreeView As String = nomeProgetto & "\" & nomeSottoCartella
                mostra_file_async_analisi_rischi(percorsoRelativoPerTreeView, TreeView3)

                ' Messaggio opzionale o apertura cartella
                ' Process.Start("explorer.exe", percorsoCompletoSottoCartella)
            Else
            MessageBox.Show("Sorgente non trovata: " & fileSorgente, "Errore File", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    Catch ex As Exception
        MessageBox.Show("Errore durante l'operazione: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try
End Sub

    Private Async Sub tabpage5_Click(sender As Object, e As EventArgs) Handles TabPage5.Enter
        Dim nomeProgetto As String = LinkLabel1.Text
        Dim nomeSottoCartella As String = "Azioni contenitive"
        Dim percorsoRelativo As String = IO.Path.Combine(nomeProgetto, nomeSottoCartella)
        mostra_file_async_analisi_rischi(percorsoRelativo, TreeView3)
    End Sub

    Private Sub TreeView3_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView3.NodeMouseDoubleClick
        Try
            ' 1. Verifica se il nodo cliccato ha un Tag valido
            If e.Node.Tag Is Nothing Then Exit Sub

            ' 2. Gestione se il nodo è un FILE
            If TypeOf e.Node.Tag Is IO.FileInfo Then
                Dim file As IO.FileInfo = DirectCast(e.Node.Tag, IO.FileInfo)

                ' Verifica che il file esista ancora prima di provare ad aprirlo
                If file.Exists Then
                    Process.Start(New ProcessStartInfo(file.FullName) With {.UseShellExecute = True})
                Else
                    MsgBox("Il file non è più disponibile nel percorso: " & file.FullName, MsgBoxStyle.Critical)
                End If

                ' 3. Gestione se il nodo è una CARTELLA (Directory)
            ElseIf TypeOf e.Node.Tag Is IO.DirectoryInfo Then
                Dim directory As IO.DirectoryInfo = DirectCast(e.Node.Tag, IO.DirectoryInfo)

                If directory.Exists Then
                    Process.Start("explorer.exe", directory.FullName)
                Else
                    MsgBox("La cartella non è raggiungibile.", MsgBoxStyle.Exclamation)
                End If
            End If

        Catch ex As Exception
            MsgBox("Impossibile aprire l'elemento: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        Dim nomeProgetto As String = LinkLabel1.Text
        Dim nomeSottoCartella As String = "Azioni contenitive"
        Dim percorsoRelativoPerTreeView As String = nomeProgetto & "\" & nomeSottoCartella
        mostra_file_async_analisi_rischi(percorsoRelativoPerTreeView, TreeView3)
    End Sub

    Private Sub TableLayoutPanel11_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel11.Paint

    End Sub
End Class