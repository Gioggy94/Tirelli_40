Imports System.Data.SqlClient

Public Class Opportunità
    Public filtro_status As String = "and t1.status='O' "
    Private filtro_FASE As String
    Private filtro_num_opportunità As String
    Private filtro_cliente_opportunità As String
    Private filtro_paese As String
    Private filtro_type As String
    Private filtro_salesman As String
    Private filtro_phase_owner As String

    Sub inizializza_opportunità()
        carica_checkedlistbox_status()
        carica_checkedlistbox_fase()

        riempi_datagridview_opportunità(filtro_status, filtro_FASE)
        riempi_datagridview_sommario(filtro_status, filtro_FASE)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
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


    Sub riempi_datagridview_opportunità(filtro_status As String, par_filtro_fase As String)
        DataGridView1.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "
select top 100 *
from
(
SELECT T1.[OpprId] as 'OPP NR°',   T1.[CardCode] as 'BP CODE', T1.[CardName] as 'BP Name', T1.u_CLIENTEFINALE AS 'End user',t5.name as 'Country', t4.Descript as 'Fase',
    CONCAT (t1.U_clientefinale,case when t1.u_clientefinale  is null then '' else ': ' end,T1.[U_Descrizioneprogetto]) AS 'Description', t0.opendate as 'Insert date',
 DATEDIFF(dd, T0.[opendate], GETDATE()) + 1 - (DATEDIFF(wk, T0.[opendate], GETDATE()) * 2)
        - (CASE WHEN DATEPART(dw, T0.[opendate]) = 1 THEN 1 ELSE 0 END)
        - (CASE WHEN DATEPART(dw, T0.[opendate]) = 7 THEN 1 ELSE 0 END)-1 AS 'permanenza'
,
T0.[U_PRIORITA] as 'Priorità',t0.u_data_richiesta as 'Data richiesta', t0.U_layout as 'Layout',   T0.[U_Informazioni] as 'Informations', T0.[Line], t6.firstname+' '+t6.lastname  as 'Tirelli owner', t2.firstname+' '+t2.lastname as 'Tirelli phase owner', T3.[Slpname] as 'Tirelli Salesman', t0.U_prg_azs_notelivopp AS 'Note livello',T0.[step_id],  t0.status as 'Status', T0.[DocNumber] as 'Document N°' 

    FROM OPR1 T0 inner JOIN OOPR T1 ON T0.[OpprId] = T1.[OpprId]
    left join [TIRELLI_40].[DBO].OHEM T6 ON t6.empid=t1.owner
    left join [TIRELLI_40].[DBO].OHEM T2 ON t2.empid=t0.owner
    left join  OSLP T3 ON T3.slpcode =t0.slpcode
    left join oost t4 on t0.step_id=t4.num
    left join [dbo].[@BNCCRY] t5 on t5.code= t1.u_destinazione

INNER JOIN 
	(SELECT T0.OPPRID, MAX(T0.LINE) AS 'MAXLINE'
	FROM OPR1 T0
	GROUP BY T0.OpprId
	) A ON A.OPPRID=T0.OPPRID AND A.MAXLINE=T0.LINE


    WHERE 0=0  " & filtro_status & "" & par_filtro_fase & "
    and T0.ObjType<>'22' and t0.step_id<>'11'and t0.step_id<>'10'and t0.step_id<>'9'and t0.step_id<>'6'and t0.step_id<>'1' AND (T1.U_UFFICIO is null or T1.U_UFFICIO='COMMERCIALE')
)
as t10
where 0 =0 " & filtro_num_opportunità & " " & filtro_cliente_opportunità & " " & filtro_paese & " " & filtro_type & " " & filtro_salesman & " " & filtro_phase_owner & "
order by
DATEDIFF(dd, T10.[Insert date], GETDATE()) + 1 - (DATEDIFF(wk, T10.[Insert date], GETDATE()) * 2)
        - (CASE WHEN DATEPART(dw, T10.[Insert date]) = 1 THEN 1 ELSE 0 END)
        - (CASE WHEN DATEPART(dw, T10.[Insert date]) = 7 THEN 1 ELSE 0 END) DESC
"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()



            DataGridView1.Rows.Add(cmd_SAP_reader("OPP NR°"), cmd_SAP_reader("BP CODE"), cmd_SAP_reader("BP Name"), cmd_SAP_reader("End user"), cmd_SAP_reader("Country"), 0, 0, cmd_SAP_reader("fase"), cmd_SAP_reader("Description"), cmd_SAP_reader("Insert date"), cmd_SAP_reader("Priorità"), cmd_SAP_reader("permanenza"), cmd_SAP_reader("Layout"), cmd_SAP_reader("Informations"), cmd_SAP_reader("Line"), cmd_SAP_reader("Tirelli owner"), cmd_SAP_reader("Tirelli phase owner"), cmd_SAP_reader("Tirelli Salesman"), cmd_SAP_reader("Note livello"), cmd_SAP_reader("step_id"), cmd_SAP_reader("Status"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

        DataGridView1.ClearSelection()

    End Sub

    Private Sub tabpage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter
        riempi_datagridview_costi_macchina(DataGridView3)
    End Sub

    Sub riempi_datagridview_costi_macchina(par_datagridview As DataGridView)

        par_datagridview.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "select 
t11.id
      ,t11.[Tipo],

      t11.[Opportunità]
	  ,COALESCE(T14.CARDNAME,'') AS 'Cardname'
      ,t11.Rev
      ,t11.[utente]
,t11.valuta
,t11.cambio
	 , concat(t12.lastname,' ', t12.firstname) as 'Nome'
      ,t11.[Data]
      ,t11.[ora]
      ,t11.[Note]
from
(
SELECT 
--t0.id
      --t0.[Tipo],

      t0.[Opportunità]
	 -- ,COALESCE(T3.CARDNAME,'') AS 'Cardname'
      ,max(t0.[REV]) as 'Rev'
      --,t0.[utente]
--,t0.valuta
--,t0.cambio
--	 , concat(t1.lastname,' ', t1.firstname) as 'Nome'
      --,max(t0.[Data]
      --,t0.[ora]
      --,t0.[Note]
  FROM [Tirelli_40].[dbo].[Superlistino_log_costificazioni_testata] t0

  
  where t0.[Data]>=getdate()-60
    group by t0.[Opportunità]
  --order by id desc
  )
  as t10 inner join [Tirelli_40].[dbo].[Superlistino_log_costificazioni_testata] t11 on t10.Opportunità=t11.Opportunità and t11.REV=t10.rev
    left join [TIRELLI_40].[DBO].ohem t12 on t12.empid=t11.utente
  left join OOPR t13 on t13.OpprId=t11.[Opportunità]
  left join ocrd t14 on t14.cardcode=t13.CardCode
  order by id desc
"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()


            par_datagridview.Rows.Add(cmd_SAP_reader("ID"), cmd_SAP_reader("Tipo"), cmd_SAP_reader("Opportunità"), cmd_SAP_reader("Cardname"), cmd_SAP_reader("rev"), cmd_SAP_reader("Nome"), cmd_SAP_reader("Valuta"), cmd_SAP_reader("Cambio"), cmd_SAP_reader("Data"), cmd_SAP_reader("Note"))

        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

        par_datagridview.ClearSelection()

    End Sub



    Sub riempi_datagridview_sommario(filtro_status As String, par_filtro_fase As String)
        Dim contatore As Integer = 0
        DataGridView2.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT count(T1.[OpprId]) as 'N°', t4.Descript as 'Fase',T0.[step_id] 

FROM OPR1 T0 inner JOIN OOPR T1 ON T0.[OpprId] = T1.[OpprId]
left join [TIRELLI_40].[DBO].OHEM T6 ON t6.empid=t1.owner
left join [TIRELLI_40].[DBO].OHEM T2 ON t2.empid=t0.owner
left join  OSLP T3 ON T3.slpcode =t0.slpcode
left join oost t4 on t0.step_id=t4.num
left join [dbo].[@BNCCRY] t5 on t5.code= t1.u_destinazione


WHERE  T0.[opendate] >= getdate()-365 " & filtro_status & "" & par_filtro_fase & "
and T0.ObjType<>'22'  AND (T1.U_UFFICIO is null or T1.U_UFFICIO='COMMERCIALE')
group by t4.Descript ,T0.[step_id]
"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()



            DataGridView2.Rows.Add(cmd_SAP_reader("step_id"), cmd_SAP_reader("Fase"), cmd_SAP_reader("N°"))
            contatore = contatore + cmd_SAP_reader("N°")
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()
        Label1.Text = contatore

        DataGridView2.ClearSelection()
    End Sub

    Sub carica_checkedlistbox_fase()
        filtro_FASE = "and "
        CheckedListBox1.Items.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "SELECT T0.stepid, t0.num, T0.Descript, T0.Canceled, T0.SalesStage, T0.PurStage FROM OOST T0 "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            If cmd_SAP_reader_2("num") = 2 Or cmd_SAP_reader_2("num") = 3 Or cmd_SAP_reader_2("num") = 5 Or cmd_SAP_reader_2("num") = 8 Or cmd_SAP_reader_2("num") = 14 Then
                CheckedListBox1.Items.Add(cmd_SAP_reader_2("num") & " " & cmd_SAP_reader_2("Descript"), True)
            ElseIf cmd_SAP_reader_2("num") = 7 Or cmd_SAP_reader_2("num") = 9 Or cmd_SAP_reader_2("num") = 12 Or cmd_SAP_reader_2("num") = 15 Then

                CheckedListBox1.Items.Add(cmd_SAP_reader_2("num") & " " & cmd_SAP_reader_2("Descript"), False)

            End If


        Loop
        cmd_SAP_reader_2.Close()
        cnn1.Close()
        componi_filtro_fase()

    End Sub

    Sub carica_checkedlistbox_status()
        filtro_status = "and "
        CheckedListBox2.Items.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "SELECT T0.[Status] FROM OOPR T0 group BY T0.[Status]"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            If cmd_SAP_reader_2("status") = "O" Then
                CheckedListBox2.Items.Add(cmd_SAP_reader_2("status"), True)
            Else
                CheckedListBox2.Items.Add(cmd_SAP_reader_2("status"), False)
            End If

        Loop
        cmd_SAP_reader_2.Close()
        cnn1.Close()
        componi_filtro_status()

    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        componi_filtro_fase()
        'riempi_datagridview_opportunità(filtro_status, filtro_FASE)
        'riempi_datagridview_sommario(filtro_status, filtro_FASE)
    End Sub

    Sub componi_filtro_fase()

        filtro_FASE = "and "
        For Each selectedItem As Object In CheckedListBox1.CheckedItems
            Dim selectedText As String = selectedItem.ToString()
            Dim firstPart As String = Scheda_commessa_Pianificazione.GetFirstPart(selectedText)

            ' Utilizza il valore di firstPart come desideri

            If filtro_FASE = "and " Then
                filtro_FASE = filtro_FASE & "(t0.step_id= " & firstPart & ""
            Else
                filtro_FASE = filtro_FASE & " or t0.step_id= " & firstPart
            End If
        Next
        filtro_FASE = filtro_FASE & ")"


    End Sub

    Sub componi_filtro_status()

        filtro_status = "and "
        For Each selectedItem As Object In CheckedListBox2.CheckedItems
            Dim selectedText As String = selectedItem.ToString()

            ' Utilizza il valore di firstPart come desideri

            If filtro_status = "and " Then
                filtro_status = filtro_status & "(t1.status= '" & selectedText & "'"
            Else
                filtro_status = filtro_status & " or t1.status= '" & selectedText & "'"
            End If
        Next
        filtro_status = filtro_status & ")"


    End Sub



    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Step_id").Value = "3" Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="type").Style.BackColor = Color.Violet

        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Step_id").Value = "2" Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="type").Style.BackColor = Color.Gray


        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Step_id").Value = "5" Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="type").Style.BackColor = Color.Blue

        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Step_id").Value = "8" Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="type").Style.BackColor = Color.Orange
        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Step_id").Value = "14" Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="type").Style.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        If DataGridView2.Rows(e.RowIndex).Cells(columnName:="codice_categoria").Value = "3" Then
            DataGridView2.Rows(e.RowIndex).Cells(columnName:="categoria").Style.BackColor = Color.Violet

        ElseIf DataGridView2.Rows(e.RowIndex).Cells(columnName:="codice_categoria").Value = "2" Then
            DataGridView2.Rows(e.RowIndex).Cells(columnName:="categoria").Style.BackColor = Color.Gray


        ElseIf DataGridView2.Rows(e.RowIndex).Cells(columnName:="codice_categoria").Value = "5" Then
            DataGridView2.Rows(e.RowIndex).Cells(columnName:="categoria").Style.BackColor = Color.Blue

        ElseIf DataGridView2.Rows(e.RowIndex).Cells(columnName:="codice_categoria").Value = "8" Then
            DataGridView2.Rows(e.RowIndex).Cells(columnName:="categoria").Style.BackColor = Color.Orange
        ElseIf DataGridView2.Rows(e.RowIndex).Cells(columnName:="codice_categoria").Value = "14" Then
            DataGridView2.Rows(e.RowIndex).Cells(columnName:="categoria").Style.BackColor = Color.Yellow
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        If e.RowIndex >= 0 Then

            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Opp_nr) Then


                Opportunità_N.n_opportunità = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Opp_nr").Value
                Opportunità_N.Show()
                Opportunità_N.inizializza_opportunità(Opportunità_N.n_opportunità)

            End If
        End If

    End Sub

    Private Sub CheckedListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox2.SelectedIndexChanged
        componi_filtro_status()

        'riempi_datagridview_opportunità(filtro_status, filtro_FASE)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = "" Then
            filtro_num_opportunità = ""
        Else
            filtro_num_opportunità = " and t10.[OPP NR°] Like '%%" & TextBox2.Text & "%%'"
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            filtro_cliente_opportunità = ""
        Else
            filtro_cliente_opportunità = " and (t10.[bp name] Like '%%" & TextBox1.Text & "%%' or t10.[end user] Like '%%" & TextBox1.Text & "%%') "
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Then
            filtro_paese = ""
        Else
            filtro_paese = " and t10.[country] Like '%%" & TextBox3.Text & "%%'"
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = "" Then
            filtro_type = ""
        Else
            filtro_type = " and t10.[fase] Like '%%" & TextBox4.Text & "%%'"
        End If
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text = "" Then
            filtro_salesman = ""
        Else
            filtro_salesman = " and t10.[Tirelli Salesman] Like '%%" & TextBox5.Text & "%%'"
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If TextBox6.Text = "" Then
            filtro_phase_owner = ""
        Else
            filtro_phase_owner = " and t10.[Tirelli phase owner] Like '%%" & TextBox6.Text & "%%'"
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        riempi_datagridview_opportunità(filtro_status, filtro_FASE)
        riempi_datagridview_sommario(filtro_status, filtro_FASE)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Cmd_Cerca_Click(sender As Object, e As EventArgs) Handles Cmd_Cerca.Click
        Opportunità_N.n_opportunità = TextBox7.Text
        Opportunità_N.Show()
        Opportunità_N.inizializza_opportunità(Opportunità_N.n_opportunità)
    End Sub



    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Dim par_datagridview As DataGridView = DataGridView3
        If e.RowIndex < 0 Then Exit Sub
        If e.ColumnIndex = par_datagridview.Columns.IndexOf(Rev) Then
            Form_configuratore_vendita.Show()
            Form_configuratore_vendita.ricrea_form_da_database_new(par_datagridview.Rows(e.RowIndex).Cells(columnName:="n_opp").Value, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Rev").Value)
        End If
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub
End Class