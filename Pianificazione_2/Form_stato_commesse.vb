Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop

Public Class Form_stato_commesse
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Function GetOrderBy() As String
        Select Case ComboBox_ordinamento.SelectedItem?.ToString()
            Case "ODP"
                Return "t0.odp, t0.ordine_stato DESC, t0.itemcode"
            Case "Commessa"
                Return "t0.matricola, t0.sottocommessa, t0.ordine_stato DESC, t0.itemcode"
            Case "Codice art."
                Return "t0.itemcode, t0.ordine_stato DESC"
            Case "Data consegna"
                Return "t0.data_c, t0.ordine_stato DESC, t0.progetto, t0.sottocommessa"
            Case Else ' "Urgenza" o default
                Return "t0.ordine_stato DESC, t0.progetto, t0.sottocommessa, t0.oc, t0.odp, t0.itemcode"
        End Select
    End Function

    Sub carica_datagridview_stato_commesse(par_codice_art As String, par_datagridview As DataGridView, par_label As Label, par_progetto As String, par_sottocommessa As String, par_matricola As String, par_odp_padre As String, par_stato As String, par_mag_dest As String, par_ubicazione As String, par_mag_imp As String, par_pianificato As String)

        Dim filtro_codice As String = ""
        If par_codice_art <> "" Then
            filtro_codice = " And t0.itemcode='" & par_codice_art & "' "
        End If

        Dim filtro_progetto As String = ""
        If par_progetto <> "" Then
            filtro_progetto = " And t0.progetto='" & par_progetto & "' "
        End If

        Dim filtro_sottocommessa As String = ""
        If par_sottocommessa <> "" Then
            filtro_sottocommessa = " And t0.sottocommessa='" & par_sottocommessa & "' "
        End If

        Dim filtro_matricola As String = ""
        If par_matricola <> "" Then
            filtro_matricola = " And t0.matricola='" & par_matricola & "' "
        End If

        Dim filtro_odp_padre As String = ""
        If par_odp_padre <> "" Then
            filtro_odp_padre = " And t0.odp='" & par_odp_padre & "' "
        End If

        Dim filtro_stato As String = ""
        If par_stato <> "" Then
            filtro_stato = " And t0.stato Like '%%" & par_stato & "%%' "
        End If

        Dim filtro_MAG_dEST As String = ""
        If par_mag_dest <> "" Then
            filtro_MAG_dEST = " And t0.mag_odp Like '%%" & par_mag_dest & "%%' "
        End If

        Dim filtro_ubicazione As String = ""
        If par_ubicazione <> "" Then
            filtro_ubicazione = " And t0.ubicazione Like '%%" & par_ubicazione & "%%' "
        End If

        Dim filtro_mag_impegno As String = ""
        If par_mag_imp <> "" Then
            filtro_mag_impegno = " And t0.[mag_ver] Like '%%" & par_mag_imp & "%%' "
        End If

        Dim filtro_pianificato As String = ""
        If par_pianificato <> "" Then
            filtro_pianificato = " And coalesce(t0.[pianificato],'') Like '%%" & par_pianificato & "%%' "
        End If

        Dim contatore As Integer = 0
        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        CMD_SAP_2.Connection = Cnn1

        CMD_SAP_2.CommandText = "SELECT top 10000 t0.[codice_univoco]
      ,t0.[itemcode]
 ,des_code
	,disegno

      ,t0.[documento]
      ,t0.[odp]
      ,t0.[oc]
      ,t0.[mag_ver]
      ,t0.[itemname]
      ,t0.[progetto]
      ,t0.[sottocommessa]
      ,t0.[matricola]
      ,t0.[desc_commessa]
,coalesce(t2.Nome_Baia,'') as 'Nome_baia'
      ,t0.[dtasca]
      ,t0.[qtapia]
      ,t0.[qtatra]
      ,t0.[qtadatra]
,t0.saldo_imp
      ,t0.[stato]
,t0.ubicazione
      ,t0.[ordine_stato]
,coalesce(t0.[documento_ordinato],'') as 'documento_ordinato'
      ,coalesce(t0.[n_documento_ord],'') as 'n_documento_ord'
      ,t0.[desc_for]
      ,t0.[q]
      ,CONVERT(date, t0.data_r, 112) AS data_r
       ,CONVERT(date, t0.data_c, 112) AS data_c
      ,t0.[mag_odp]
      ,CONVERT(date, coalesce(t0.data_immissione, t0.data_r), 112) AS data_immissione
      ,CONVERT(date, t0.data_i, 112) AS data_i
      ,coalesce(t0.[pianificato],'') as 'pianificato'
  FROM [Tirelli_40].[dbo].[stato_commesse_output] t0
  left join [Tirelli_40].[dbo].[Layout_CAP1] t1 on t1.Commessa=t0.matricola and t1.Stato='O'
  left join [Tirelli_40].[dbo].[Layout_CAP1_nomi] t2 on t2.numero_baia=t1.Baia
where 0=0 and visibile='Y'  " & filtro_codice & filtro_progetto & filtro_sottocommessa & filtro_matricola & filtro_odp_padre & filtro_stato & filtro_MAG_dEST & filtro_ubicazione & filtro_mag_impegno & filtro_pianificato & "
  ORDER BY " & GetOrderBy()

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            Dim img As Image = Nothing
            Dim codiceDisegno As String = cmd_SAP_reader_2("Disegno")
            Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & codiceDisegno & ".PNG"

            If File.Exists(percorso) Then
                Using fs As New FileStream(percorso, FileMode.Open, FileAccess.Read)
                    Using tmp As Image = Image.FromStream(fs)
                        img = New Bitmap(tmp)
                    End Using
                End Using
            End If

            par_datagridview.Rows.Add(
        cmd_SAP_reader_2("codice_univoco"),
        cmd_SAP_reader_2("itemcode"),
        cmd_SAP_reader_2("des_code"),
        img,
        cmd_SAP_reader_2("Disegno"),
        cmd_SAP_reader_2("documento"),
        cmd_SAP_reader_2("odp"),
        cmd_SAP_reader_2("pianificato"),
            cmd_SAP_reader_2("data_i"),
        cmd_SAP_reader_2("itemname"),
        cmd_SAP_reader_2("oc"),
            cmd_SAP_reader_2("mag_odp"),
        cmd_SAP_reader_2("mag_ver"),
        cmd_SAP_reader_2("progetto"),
        cmd_SAP_reader_2("sottocommessa"),
        cmd_SAP_reader_2("matricola"),
        cmd_SAP_reader_2("desc_commessa"),
        cmd_SAP_reader_2("nome_baia"),
        cmd_SAP_reader_2("qtapia"),
        cmd_SAP_reader_2("qtatra"),
        cmd_SAP_reader_2("qtadatra"),
        cmd_SAP_reader_2("saldo_imp"),
            cmd_SAP_reader_2("stato"),
             cmd_SAP_reader_2("ubicazione"),
        cmd_SAP_reader_2("ordine_stato"),
             cmd_SAP_reader_2("documento_ordinato"),
         cmd_SAP_reader_2("n_documento_ord"),
          cmd_SAP_reader_2("desc_for"),
           cmd_SAP_reader_2("q"),
                 cmd_SAP_reader_2("data_immissione"),
            cmd_SAP_reader_2("Data_c"))

            contatore += 1
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        par_datagridview.ClearSelection()
        par_label.Text = contatore
    End Sub

    Sub carica_datagridview_per_ODP(par_progetto As String, par_sottocommessa As String, par_matricola As String, par_odp As String)

        Dim filtro_progetto As String = ""
        If par_progetto <> "" Then
            filtro_progetto = " And t0.progetto='" & par_progetto & "' "
        End If

        Dim filtro_sottocommessa As String = ""
        If par_sottocommessa <> "" Then
            filtro_sottocommessa = " And t0.sottocommessa='" & par_sottocommessa & "' "
        End If

        Dim filtro_matricola As String = ""
        If par_matricola <> "" Then
            filtro_matricola = " And t0.matricola='" & par_matricola & "' "
        End If

        Dim filtro_odp As String = ""
        If par_odp <> "" Then
            filtro_odp = " And t0.odp='" & par_odp & "' "
        End If

        DataGridView_per_ODP.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD As New SqlCommand
        CMD.Connection = Cnn1
        CMD.CommandText = "
SELECT
    t0.odp,
    t0.matricola,
    t0.desc_commessa,
    coalesce(t2.Nome_Baia,'') as Nome_baia,
    MAX(t0.ordine_stato) as ord_stato_max,
    SUM(CASE WHEN t0.ordine_stato=6 THEN 1 ELSE 0 END) as n_da_ord,
    SUM(CASE WHEN t0.ordine_stato=5 THEN 1 ELSE 0 END) as n_in_fut,
    SUM(CASE WHEN t0.ordine_stato=4 THEN 1 ELSE 0 END) as n_in_pass,
    SUM(CASE WHEN t0.ordine_stato=3 THEN 1 ELSE 0 END) as n_tcp,
    SUM(CASE WHEN t0.ordine_stato=2 THEN 1 ELSE 0 END) as n_cq,
    SUM(CASE WHEN t0.ordine_stato=1 THEN 1 ELSE 0 END) as n_ass,
    SUM(CASE WHEN t0.ordine_stato=0 THEN 1 ELSE 0 END) as n_trasf,
    COUNT(*) as n_tot
FROM [Tirelli_40].[dbo].[stato_commesse_output] t0
left join [Tirelli_40].[dbo].[Layout_CAP1] t1 on t1.Commessa=t0.matricola and t1.Stato='O'
left join [Tirelli_40].[dbo].[Layout_CAP1_nomi] t2 on t2.numero_baia=t1.Baia
WHERE visibile='Y'
" & filtro_progetto & filtro_sottocommessa & filtro_matricola & filtro_odp & "
GROUP BY t0.odp, t0.matricola, t0.desc_commessa, t2.Nome_Baia
ORDER BY MAX(t0.ordine_stato) DESC, t0.matricola, t0.odp"

        Dim rd As SqlDataReader = CMD.ExecuteReader

        Do While rd.Read()
            Dim ordStato As Integer = CInt(rd("ord_stato_max"))
            Dim statoTesto As String = ""
            Select Case ordStato
                Case 6 : statoTesto = "Da ordinare"
                Case 5 : statoTesto = "In ordine futuro"
                Case 4 : statoTesto = "In ordine passato"
                Case 3 : statoTesto = "Trasferibile TCP"
                Case 2 : statoTesto = "CQ"
                Case 1 : statoTesto = "DA_ASS"
                Case 0 : statoTesto = "Trasferibile"
            End Select

            DataGridView_per_ODP.Rows.Add(
                rd("odp"),
                rd("matricola"),
                rd("desc_commessa"),
                rd("Nome_baia"),
                statoTesto,
                ordStato,
                IIf(CInt(rd("n_da_ord")) > 0, rd("n_da_ord").ToString(), ""),
                IIf(CInt(rd("n_in_fut")) > 0, rd("n_in_fut").ToString(), ""),
                IIf(CInt(rd("n_in_pass")) > 0, rd("n_in_pass").ToString(), ""),
                IIf(CInt(rd("n_tcp")) > 0, rd("n_tcp").ToString(), ""),
                IIf(CInt(rd("n_cq")) > 0, rd("n_cq").ToString(), ""),
                IIf(CInt(rd("n_ass")) > 0, rd("n_ass").ToString(), ""),
                IIf(CInt(rd("n_trasf")) > 0, rd("n_trasf").ToString(), ""),
                rd("n_tot"))
        Loop

        rd.Close()
        Cnn1.Close()
        DataGridView_per_ODP.ClearSelection()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TabControl_viste.SelectedTab Is TabPage_lista Then
            carica_datagridview_stato_commesse(TextBox7.Text.ToUpper, DataGridView_stato_commesse, Label1, TextBox1.Text.ToUpper, TextBox15.Text.ToUpper, TextBox2.Text.ToUpper, TextBox3.Text.ToUpper, TextBox9.Text.ToUpper, TextBox5.Text.ToUpper, TextBox4.Text.ToUpper, TextBox6.Text.ToUpper, TextBox_StatoODP.Text.ToUpper)
        Else
            carica_datagridview_per_ODP(TextBox1.Text.ToUpper, TextBox15.Text.ToUpper, TextBox2.Text.ToUpper, TextBox3.Text.ToUpper)
        End If
    End Sub

    Private Sub TabControl_viste_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl_viste.SelectedIndexChanged
        If TabControl_viste.SelectedTab Is TabPage_odp Then
            carica_datagridview_per_ODP(TextBox1.Text.ToUpper, TextBox15.Text.ToUpper, TextBox2.Text.ToUpper, TextBox3.Text.ToUpper)
        ElseIf TabControl_viste.SelectedTab Is TabPage_albero Then
            If TreeView_mancanti.Nodes.Count = 0 Then
                carica_albero_mancanti(TextBox_filtro_albero.Text.ToUpper)
            End If
        End If
    End Sub

    Private Sub Button_aggiorna_albero_Click(sender As Object, e As EventArgs) Handles Button_aggiorna_albero.Click
        carica_albero_mancanti(TextBox_filtro_albero.Text.ToUpper)
    End Sub

    Private Sub Button_espandi_albero_Click(sender As Object, e As EventArgs) Handles Button_espandi_albero.Click
        If TreeView_mancanti.Nodes.Count = 0 Then Return
        ' Toggle: se almeno un nodo è collassato, espande tutto; altrimenti collassa tutto
        Dim almenoUnoCollassato As Boolean = False
        For Each n As TreeNode In TreeView_mancanti.Nodes
            If Not n.IsExpanded Then almenoUnoCollassato = True : Exit For
        Next
        TreeView_mancanti.BeginUpdate()
        If almenoUnoCollassato Then
            TreeView_mancanti.ExpandAll()
            Button_espandi_albero.Text = "Comprimi tutto"
        Else
            TreeView_mancanti.CollapseAll()
            Button_espandi_albero.Text = "Espandi tutto"
        End If
        TreeView_mancanti.EndUpdate()
    End Sub

    Private Sub TextBox_filtro_albero_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_filtro_albero.KeyDown
        If e.KeyCode = Keys.Enter Then
            carica_albero_mancanti(TextBox_filtro_albero.Text.ToUpper)
        End If
    End Sub

    Private Sub TreeView_mancanti_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView_mancanti.AfterSelect
        Dim tag As String = If(e.Node.Tag?.ToString(), "")
        If tag = "" Then Return

        Dim parts() As String = tag.Split("|"c)
        If parts.Length < 2 Then Return

        Dim tipoNodo As String = parts(0)
        Dim valore As String = parts(1)

        Dim filtroMatricola As String = ""
        Dim filtroOdp As String = ""
        Dim filtroItem As String = ""

        Select Case tipoNodo
            Case "C" ' commessa
                filtroMatricola = valore
            Case "O" ' odp
                Dim pC() As String = parts(1).Split("~"c)
                filtroMatricola = pC(0)
                filtroOdp = pC(1)
            Case "A" ' articolo
                Dim pA() As String = parts(1).Split("~"c)
                filtroMatricola = pA(0)
                filtroOdp = pA(1)
                filtroItem = pA(2)
        End Select

        carica_dettaglio_albero(filtroMatricola, filtroOdp, filtroItem)
    End Sub

    Sub inizializza_colonne_dettaglio_albero()
        If DataGridView_dettaglio_albero.Columns.Count > 0 Then Return
        DataGridView_dettaglio_albero.Columns.Add("alb_stato", "Stato")
        DataGridView_dettaglio_albero.Columns.Add("alb_itemcode", "Cod. Art.")
        DataGridView_dettaglio_albero.Columns.Add("alb_des_code", "Descrizione")
        DataGridView_dettaglio_albero.Columns.Add("alb_odp", "ODP")
        DataGridView_dettaglio_albero.Columns.Add("alb_itemname", "Desc ODP")
        DataGridView_dettaglio_albero.Columns.Add("alb_qtapia", "Q Prev.")
        DataGridView_dettaglio_albero.Columns.Add("alb_qtatra", "Q Trasf.")
        DataGridView_dettaglio_albero.Columns.Add("alb_qtadatra", "Q Da Trasf.")
        DataGridView_dettaglio_albero.Columns.Add("alb_ubicazione", "Ubicazione")
        DataGridView_dettaglio_albero.Columns.Add("alb_doc_ord", "Doc Ord")
        DataGridView_dettaglio_albero.Columns.Add("alb_data_c", "Data Cons.")
        DataGridView_dettaglio_albero.Columns.Add("alb_data_imm", "Data Imm.")
        DataGridView_dettaglio_albero.Columns.Add("alb_data_i", "Data ODP")
        DataGridView_dettaglio_albero.Columns.Add("alb_desc_for", "Fornitore")
        DataGridView_dettaglio_albero.Columns.Add("alb_ordine_stato", "OrdStato")
        DataGridView_dettaglio_albero.Columns("alb_ordine_stato").Visible = False
        DataGridView_dettaglio_albero.Columns("alb_stato").FillWeight = 80
        DataGridView_dettaglio_albero.Columns("alb_itemcode").FillWeight = 80
        DataGridView_dettaglio_albero.Columns("alb_des_code").FillWeight = 160
        DataGridView_dettaglio_albero.Columns("alb_odp").FillWeight = 60
        DataGridView_dettaglio_albero.Columns("alb_itemname").FillWeight = 160
        DataGridView_dettaglio_albero.Columns("alb_qtapia").FillWeight = 50
        DataGridView_dettaglio_albero.Columns("alb_qtatra").FillWeight = 50
        DataGridView_dettaglio_albero.Columns("alb_qtadatra").FillWeight = 60
        DataGridView_dettaglio_albero.Columns("alb_ubicazione").FillWeight = 70
        DataGridView_dettaglio_albero.Columns("alb_doc_ord").FillWeight = 60
        DataGridView_dettaglio_albero.Columns("alb_data_c").FillWeight = 70
        DataGridView_dettaglio_albero.Columns("alb_data_imm").FillWeight = 70
        DataGridView_dettaglio_albero.Columns("alb_data_i").FillWeight = 70
        DataGridView_dettaglio_albero.Columns("alb_desc_for").FillWeight = 120
    End Sub

    Sub carica_dettaglio_albero(matricola As String, odp As String, itemcode As String)
        inizializza_colonne_dettaglio_albero()
        DataGridView_dettaglio_albero.Rows.Clear()

        Dim filtroM As String = If(matricola <> "", " And t0.matricola='" & matricola & "' ", "")
        Dim filtroO As String = If(odp <> "", " And t0.odp='" & odp & "' ", "")
        Dim filtroI As String = If(itemcode <> "", " And t0.itemcode='" & itemcode & "' ", "")

        Dim Cnn As New SqlConnection(Homepage.sap_tirelli)
        Cnn.Open()
        Dim CMD As New SqlCommand
        CMD.Connection = Cnn
        CMD.CommandText = "SELECT t0.stato, t0.itemcode, t0.des_code, t0.odp, t0.itemname,
            t0.qtapia, t0.qtatra, t0.qtadatra, t0.ubicazione,
            coalesce(t0.documento_ordinato,'') as documento_ordinato,
            CONVERT(date, t0.data_c, 112) AS data_c,
            CONVERT(date, coalesce(t0.data_immissione, t0.data_r), 112) AS data_immissione,
            CONVERT(date, t0.data_i, 112) AS data_i,
            coalesce(t0.desc_for,'') as desc_for,
            t0.ordine_stato
          FROM [Tirelli_40].[dbo].[stato_commesse_output] t0
          WHERE visibile='Y' " & filtroM & filtroO & filtroI & "
          ORDER BY t0.ordine_stato DESC, t0.itemcode"
        Dim rd As SqlDataReader = CMD.ExecuteReader
        Do While rd.Read()
            Dim rowIdx As Integer = DataGridView_dettaglio_albero.Rows.Add(
                rd("stato"), rd("itemcode"), rd("des_code"), rd("odp"), rd("itemname"),
                rd("qtapia"), rd("qtatra"), rd("qtadatra"), rd("ubicazione"),
                rd("documento_ordinato"), rd("data_c"), rd("data_immissione"), rd("data_i"), rd("desc_for"), rd("ordine_stato"))
            Dim ordS As Integer = CInt(rd("ordine_stato"))
            DataGridView_dettaglio_albero.Rows(rowIdx).DefaultCellStyle.BackColor = colore_da_ordine_stato(ordS)
        Loop
        rd.Close()
        Cnn.Close()
        DataGridView_dettaglio_albero.ClearSelection()
    End Sub

    Private Function colore_da_ordine_stato(ordS As Integer) As Color
        Select Case ordS
            Case 6 : Return Color.Crimson
            Case 5 : Return Color.Tomato
            Case 4 : Return Color.OrangeRed
            Case 3 : Return Color.Orange
            Case 2 : Return Color.Gold
            Case 1 : Return Color.YellowGreen
            Case 0 : Return Color.Lime
            Case Else : Return Color.White
        End Select
    End Function

    Sub carica_albero_mancanti(filtro_commessa As String)
        TreeView_mancanti.BeginUpdate()
        TreeView_mancanti.Nodes.Clear()

        Dim filtroC As String = If(filtro_commessa <> "", " And t0.matricola Like '%%" & filtro_commessa & "%%' ", "")

        Dim Cnn As New SqlConnection(Homepage.sap_tirelli)
        Cnn.Open()
        Dim CMD As New SqlCommand
        CMD.Connection = Cnn
        ' Ordine: matricola -> odp -> ordine_stato DESC (peggiori prima dentro ogni ODP)
        CMD.CommandText = "SELECT t0.matricola, t0.desc_commessa, t0.odp, t0.itemname,
            t0.itemcode, t0.des_code, t0.qtadatra, t0.stato, t0.ordine_stato, t0.ubicazione,
            CONVERT(date, t0.data_c, 112) AS data_c
          FROM [Tirelli_40].[dbo].[stato_commesse_output] t0
          WHERE visibile='Y' " & filtroC & "
          ORDER BY t0.matricola, t0.odp, t0.ordine_stato DESC, t0.itemcode"

        Dim rd As SqlDataReader = CMD.ExecuteReader

        Dim nodoCommessa As TreeNode = Nothing
        Dim nodoOdp As TreeNode = Nothing
        Dim ultimaCommessa As String = ""
        Dim ultimoOdp As String = ""
        Dim peggiorStato_comm As Integer = 0
        Dim peggiorStato_odp As Integer = 0

        Do While rd.Read()
            Dim matricola As String = rd("matricola").ToString()
            Dim descComm As String = rd("desc_commessa").ToString()
            Dim odp As String = rd("odp").ToString()
            Dim itemname As String = rd("itemname").ToString()
            Dim itemcode As String = rd("itemcode").ToString()
            Dim des_code As String = rd("des_code").ToString()
            Dim qtadatra As String = rd("qtadatra").ToString()
            Dim stato As String = rd("stato").ToString()
            Dim ordStato As Integer = CInt(rd("ordine_stato"))
            Dim ubicazione As String = rd("ubicazione").ToString()
            Dim dataC As String = If(IsDBNull(rd("data_c")), "", CDate(rd("data_c")).ToString("dd/MM/yyyy"))

            ' Nuovo nodo commessa
            If matricola <> ultimaCommessa Then
                ' Chiudi e colora nodi precedenti
                If nodoOdp IsNot Nothing Then nodoOdp.ForeColor = colore_da_ordine_stato(peggiorStato_odp)
                If nodoCommessa IsNot Nothing Then nodoCommessa.ForeColor = colore_da_ordine_stato(peggiorStato_comm)

                nodoCommessa = New TreeNode(matricola & " - " & descComm)
                nodoCommessa.Tag = "C|" & matricola
                nodoCommessa.NodeFont = New System.Drawing.Font(TreeView_mancanti.Font, System.Drawing.FontStyle.Bold)
                TreeView_mancanti.Nodes.Add(nodoCommessa)
                ultimaCommessa = matricola
                ultimoOdp = ""
                peggiorStato_comm = ordStato
                nodoOdp = Nothing
            Else
                If ordStato > peggiorStato_comm Then peggiorStato_comm = ordStato
            End If

            ' Nuovo nodo ODP
            If odp <> ultimoOdp Then
                If nodoOdp IsNot Nothing Then nodoOdp.ForeColor = colore_da_ordine_stato(peggiorStato_odp)

                nodoOdp = New TreeNode("ODP " & odp & " - " & itemname)
                nodoOdp.Tag = "O|" & matricola & "~" & odp
                nodoCommessa.Nodes.Add(nodoOdp)
                ultimoOdp = odp
                peggiorStato_odp = ordStato
            Else
                If ordStato > peggiorStato_odp Then peggiorStato_odp = ordStato
            End If

            ' Nodo articolo
            Dim labelArt As String = "[" & stato & "] " & itemcode & " - " & des_code & "  (Q: " & qtadatra & ")"
            If dataC <> "" Then labelArt &= "  Cons: " & dataC
            If ubicazione <> "" Then labelArt &= "  Ubi: " & ubicazione
            Dim nodoArt As New TreeNode(labelArt)
            nodoArt.Tag = "A|" & matricola & "~" & odp & "~" & itemcode
            nodoArt.ForeColor = colore_da_ordine_stato(ordStato)
            nodoOdp.Nodes.Add(nodoArt)
        Loop
        rd.Close()
        Cnn.Close()

        ' Colora gli ultimi nodi rimasti aperti
        If nodoOdp IsNot Nothing Then nodoOdp.ForeColor = colore_da_ordine_stato(peggiorStato_odp)
        If nodoCommessa IsNot Nothing Then nodoCommessa.ForeColor = colore_da_ordine_stato(peggiorStato_comm)

        TreeView_mancanti.EndUpdate()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub DataGridView_stato_commesse_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_stato_commesse.CellContentClick

    End Sub

    Private Sub DataGridView_stato_commesse_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_stato_commesse.CellClick
        Dim par_datagridview As DataGridView
        par_datagridview = DataGridView_stato_commesse
        If e.RowIndex >= 0 Then

            Dim codiceDisegno As String = par_datagridview.Rows(e.RowIndex).Cells(columnName:="Disegno").Value
            Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & codiceDisegno & ".PNG"

            If e.ColumnIndex = par_datagridview.Columns.IndexOf(Itemcode) Then
                Magazzino.Codice_SAP = par_datagridview.Rows(e.RowIndex).Cells(columnName:="itemcode").Value

                If Magazzino.WindowState = FormWindowState.Minimized Then
                    Magazzino.WindowState = FormWindowState.Normal
                End If

                Magazzino.BringToFront()
                Magazzino.Activate()
                Magazzino.Show()

                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)

            ElseIf e.ColumnIndex = par_datagridview.Columns.IndexOf(ODP) Then
                Dim new_form_odp_form = New ODP_Form
                new_form_odp_form.docnum_odp = par_datagridview.Rows(e.RowIndex).Cells(columnName:="ODP").Value
                new_form_odp_form.Show()
                new_form_odp_form.inizializza_form(new_form_odp_form.docnum_odp)

            ElseIf e.ColumnIndex = par_datagridview.Columns.IndexOf(N_doc) And par_datagridview.Rows(e.RowIndex).Cells(columnName:="Doc_ord").Value = "OP" Then
                Dim new_form_odp_form = New ODP_Form
                new_form_odp_form.docnum_odp = par_datagridview.Rows(e.RowIndex).Cells(columnName:="N_doc").Value
                new_form_odp_form.Show()
                new_form_odp_form.inizializza_form(new_form_odp_form.docnum_odp)

            End If
        End If
    End Sub

    Private Sub DataGridView_stato_commesse_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_stato_commesse.CellFormatting
        Dim nome_colonna As String = "Ordine_stato"
        If DataGridView_stato_commesse.Rows(e.RowIndex).Cells(columnName:=nome_colonna).Value = "6" Then
            DataGridView_stato_commesse.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Crimson
        ElseIf DataGridView_stato_commesse.Rows(e.RowIndex).Cells(columnName:=nome_colonna).Value = "5" Then
            DataGridView_stato_commesse.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Tomato
        ElseIf DataGridView_stato_commesse.Rows(e.RowIndex).Cells(columnName:=nome_colonna).Value = "4" Then
            DataGridView_stato_commesse.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.OrangeRed
        ElseIf DataGridView_stato_commesse.Rows(e.RowIndex).Cells(columnName:=nome_colonna).Value = "3" Then
            DataGridView_stato_commesse.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Orange
        ElseIf DataGridView_stato_commesse.Rows(e.RowIndex).Cells(columnName:=nome_colonna).Value = "2" Then
            DataGridView_stato_commesse.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Gold
        ElseIf DataGridView_stato_commesse.Rows(e.RowIndex).Cells(columnName:=nome_colonna).Value = "1" Then
            DataGridView_stato_commesse.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.YellowGreen
        ElseIf DataGridView_stato_commesse.Rows(e.RowIndex).Cells(columnName:=nome_colonna).Value = "0" Then
            DataGridView_stato_commesse.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
        End If

        ' Formattazione colonna Stato ODP (R / P)
        Dim colPian As Integer = DataGridView_stato_commesse.Columns.IndexOf(Pianificato)
        If e.ColumnIndex = colPian Then
            Dim valPian As Object = DataGridView_stato_commesse.Rows(e.RowIndex).Cells(colPian).Value
            If valPian IsNot Nothing AndAlso Not IsDBNull(valPian) Then
                Select Case valPian.ToString().Trim().ToUpper()
                    Case "R"
                        e.CellStyle.BackColor = Color.LightCoral
                        e.CellStyle.Font = New Font(DataGridView_stato_commesse.Font, FontStyle.Bold)
                    Case "P"
                        e.CellStyle.BackColor = Color.LightGreen
                        e.CellStyle.Font = New Font(DataGridView_stato_commesse.Font, FontStyle.Bold)
                End Select
            End If
        End If

        ' Evidenzia Data_Imm e Data_I se lanciate nell'ultima settimana
        Dim colDataImm As Integer = DataGridView_stato_commesse.Columns.IndexOf(Data_Imm)
        Dim colDataI As Integer = DataGridView_stato_commesse.Columns.IndexOf(Data_I)
        If e.ColumnIndex = colDataImm OrElse e.ColumnIndex = colDataI Then
            Dim valData As Object = DataGridView_stato_commesse.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            If valData IsNot Nothing AndAlso Not IsDBNull(valData) Then
                If CDate(valData) >= Date.Today.AddDays(-7) Then
                    e.CellStyle.BackColor = Color.Cyan
                    e.CellStyle.Font = New Font(DataGridView_stato_commesse.Font, FontStyle.Bold)
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView_dettaglio_albero_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_dettaglio_albero.CellFormatting
        ' Colore riga per ordine_stato (colonna nascosta)
        Dim idxOrd As Integer = DataGridView_dettaglio_albero.Columns.IndexOf(DataGridView_dettaglio_albero.Columns("alb_ordine_stato"))
        If idxOrd >= 0 Then
            Dim vOrd As Object = DataGridView_dettaglio_albero.Rows(e.RowIndex).Cells(idxOrd).Value
            If vOrd IsNot Nothing AndAlso Not IsDBNull(vOrd) Then
                DataGridView_dettaglio_albero.Rows(e.RowIndex).DefaultCellStyle.BackColor = colore_da_ordine_stato(CInt(vOrd))
            End If
        End If
        ' Evidenzia data_immissione e data_i se lanciate nell'ultima settimana
        Dim idxImm As Integer = DataGridView_dettaglio_albero.Columns.IndexOf(DataGridView_dettaglio_albero.Columns("alb_data_imm"))
        Dim idxDataI As Integer = DataGridView_dettaglio_albero.Columns.IndexOf(DataGridView_dettaglio_albero.Columns("alb_data_i"))
        If (idxImm >= 0 AndAlso e.ColumnIndex = idxImm) OrElse (idxDataI >= 0 AndAlso e.ColumnIndex = idxDataI) Then
            Dim valData As Object = DataGridView_dettaglio_albero.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
            If valData IsNot Nothing AndAlso Not IsDBNull(valData) Then
                If CDate(valData) >= Date.Today.AddDays(-7) Then
                    e.CellStyle.BackColor = Color.Cyan
                    e.CellStyle.Font = New Font(DataGridView_dettaglio_albero.Font, FontStyle.Bold)
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView_per_ODP_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_per_ODP.CellFormatting
        Dim ordStato As String = DataGridView_per_ODP.Rows(e.RowIndex).Cells(columnName:="ODP_Col_OrdStato").Value?.ToString()
        Select Case ordStato
            Case "6" : DataGridView_per_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Crimson
            Case "5" : DataGridView_per_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Tomato
            Case "4" : DataGridView_per_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.OrangeRed
            Case "3" : DataGridView_per_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Orange
            Case "2" : DataGridView_per_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Gold
            Case "1" : DataGridView_per_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.YellowGreen
            Case "0" : DataGridView_per_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
        End Select
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim par_datagridview As DataGridView = DataGridView_stato_commesse
        Dim excelApp As New Excel.Application
        excelApp.Visible = True

        Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add
        Dim excelWorksheet As Excel.Worksheet = CType(excelWorkbook.Worksheets(1), Excel.Worksheet)

        For col As Integer = 1 To par_datagridview.Columns.Count
            excelWorksheet.Cells(1, col) = par_datagridview.Columns(col - 1).HeaderText
        Next

        For row As Integer = 0 To par_datagridview.Rows.Count - 1
            For col As Integer = 0 To par_datagridview.Columns.Count - 1
                excelWorksheet.Cells(row + 2, col + 1) = par_datagridview.Rows(row).Cells(col).Value
            Next
        Next

        Dim saveFileDialog As New SaveFileDialog
        saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            excelWorkbook.SaveAs(saveFileDialog.FileName)
            MessageBox.Show("Esportazione completata con successo!", "Successo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        excelApp.Quit()
        ReleaseComObject(excelApp)

    End Sub

    Public Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Class
