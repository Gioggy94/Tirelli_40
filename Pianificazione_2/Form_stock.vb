Imports System.Data.SqlClient

Public Class Form_stock
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Sub inizializza_form()
        Inserimento_stock(DataGridView, TextBox1.Text, TextBox2.Text, True, "")
    End Sub

    Sub Inserimento_stock(par_datagridview As DataGridView, par_codice As String, par_descrizione As String, par_cancella_dati As Boolean, par_parente As String)
        Dim filtro_codice As String
        Dim filtro_descrizione As String

        If par_codice = "" Then
            filtro_codice = ""
        Else
            filtro_codice = " and t0.itemcode Like '%%" & par_codice & "%%'  "

        End If

        If par_descrizione = "" Then
            filtro_descrizione = ""

        Else
            filtro_descrizione = " and t0.itemname  Like '%%" & par_descrizione & "%%'  "
        End If
        If par_cancella_dati = True Then
            par_datagridview.Rows.Clear()
        End If

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "select t0.itemcode, t0.itemname,coalesce(t0.U_Disegno,'') as 'Disegno',
coalesce(t0.u_ubimag,'') as 'Motivazione_stock',
t0.u_data_valutazione_stock,
coalesce(t0.U_Codice_BRB,'') as 'Codice_BRB',t2.ItmsGrpNam,t0.u_gestione_magazzino,coalesce(t0.MinLevel,0) as 'Minlevel'
,  coalesce(t0.MinOrdrQty,0) as 'Q_min_ord', t1.price, coalesce(t0.MinLevel,0) * t1.price as 'Valore a mag', coalesce(t0.MinOrdrQty,0)*t1.price as 'Valore_ord'
from oitm t0 inner join itm1 t1 on t0.itemcode=t1.itemcode and t1.pricelist=2
left join oitb t2 on t2.ItmsGrpCod=t0.ItmsGrpCod
where (t0.u_gestione_magazzino ='SCORTA' OR t0.u_gestione_magazzino ='STOCK'
or t0.MinLevel>0 or t0.MinOrdrQty>0 ) " & filtro_codice & filtro_descrizione & "

order by coalesce(t0.MinLevel,0) * t1.price DESC"

        cmd_SAP_reader = CMD_SAP.ExecuteReader



        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add(par_parente,
                cmd_SAP_reader("itemcode"),
                              cmd_SAP_reader("itemname"),
                              cmd_SAP_reader("Disegno"),
                              cmd_SAP_reader("Codice_BRB"),
                              cmd_SAP_reader("ItmsGrpNam"),
                              cmd_SAP_reader("u_gestione_magazzino"),
                              cmd_SAP_reader("Motivazione_stock"),
                              cmd_SAP_reader("u_data_valutazione_Stock"),
                              cmd_SAP_reader("Minlevel"),
                              cmd_SAP_reader("Q_min_ord"),
                              cmd_SAP_reader("price"),
                              cmd_SAP_reader("Valore a mag"),
                              cmd_SAP_reader("Valore_ord"))
        Loop
        cmd_SAP_reader.Close()
        CNN.Close()
        par_datagridview.ClearSelection()

    End Sub 'Inserisco le risorse nella combo box




    Private Sub Form_stock_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inizializza_form()
    End Sub

    Private Sub DataGridView_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView.CellFormatting
        Dim par_datagridview As DataGridView = DataGridView
        If par_datagridview.Rows(e.RowIndex).Cells(columnName:="Min").Value > 0 Then

            par_datagridview.Rows(e.RowIndex).Cells(columnName:="Min").Style.BackColor = Color.Aqua

        End If
     
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        controllo_padri_figli(DataGridView)
    End Sub

    Sub controllo_padri_figli(par_datagridview As DataGridView)
        For Each row As DataGridViewRow In par_datagridview.Rows
            ' Controlla che la riga non sia una riga nuova
            If Not row.IsNewRow Then
                ' Ottieni il valore della colonna "Codice" nella riga corrente
                Dim codice As String = row.Cells("Codice").Value.ToString()

                ' Esegui la funzione trova_figli passando i parametri richiesti
                trova_figli(codice, codice & "-" & row.Cells("Descrizione").Value.ToString()) ' Cambia "codice" con il valore desiderato per PAR_STRINGA se necessario
            End If
        Next
    End Sub

    Sub trova_figli(par_Codice As String, ByRef PAR_STRINGA As String)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli

        Try
            CNN.Open()

            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

            CMD_SAP.Connection = CNN
            CMD_SAP.CommandText = "SELECT t0.code, t1.itemname, t1.ItmsGrpCod, coalesce(t1.u_gestione_magazzino,'') as 'u_gestione_magazzino', coalesce(t1.MinLevel,0) as 'minlevel' 
                                  From itt1 t0 INNER Join oitm t1 ON t1.itemcode = t0.code 
        Where t0.father = '" & par_Codice & "'"


            cmd_SAP_reader = CMD_SAP.ExecuteReader()

            ' Variabile per tenere traccia se almeno un nodo soddisfa i criteri
            Dim nodoCreato As Boolean = False
            Dim rootNode As TreeNode = Nothing

            ' Itera sui risultati e aggiungi i nodi solo se rispettano i criteri dell'IF
            Do While cmd_SAP_reader.Read()
                If cmd_SAP_reader("u_gestione_magazzino").ToString() = "SCORTA" OrElse
               cmd_SAP_reader("u_gestione_magazzino").ToString() = "STOCK" OrElse
               Convert.ToInt32(cmd_SAP_reader("minlevel")) > 0 And cmd_SAP_reader("itmsgrpcod") <> "121" Then

                    ' Crea il nodo principale una sola volta, solo se necessario
                    If Not nodoCreato Then
                        rootNode = TreeView1.Nodes.Add(PAR_STRINGA)
                        nodoCreato = True
                    End If

                    ' Crea il testo del figlio come "codice - descrizione"
                    Dim childText As String = cmd_SAP_reader("code").ToString() & " - " & cmd_SAP_reader("itemname").ToString()

                    ' Aggiungi il figlio come nodo sotto rootNode
                    Dim childNode As TreeNode = rootNode.Nodes.Add(childText)

                    ' Aggiorna PAR_STRINGA per la ricorsione
                    Dim nuovoParStringa As String = PAR_STRINGA & "_" & childText

                    ' Chiama ricorsivamente trova_figli per aggiungere eventuali figli del nodo corrente
                    trova_figli(cmd_SAP_reader("code").ToString(), nuovoParStringa)
                End If
            Loop

            cmd_SAP_reader.Close()
        Catch ex As Exception
            MessageBox.Show("Errore: " & ex.Message)
        Finally
            CNN.Close()
        End Try
        TreeView1.ExpandAll()
    End Sub

    Private Sub DataGridView_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView.CellContentClick

    End Sub

    Private Sub DataGridView_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView.CellClick
        Dim par_datagridview As DataGridView = DataGridView
        If e.ColumnIndex = par_datagridview.Columns.IndexOf(Codice) Then

            Magazzino.Codice_SAP = par_datagridview.Rows(e.RowIndex).Cells(columnName:="Codice").Value

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

        ElseIf e.ColumnIndex = par_datagridview.Columns.IndexOf(Disegno) Then


            Magazzino.visualizza_disegno(par_datagridview.Rows(e.RowIndex).Cells(columnName:="Disegno").Value)
        ElseIf e.ColumnIndex = par_datagridview.Columns.IndexOf(Descrizione) Then

            Form_Cicli_di_lavoro.ControllaEVisualizzaPDF(CheckBox1, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Disegno").Value, WebBrowser1)

        End If
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect

    End Sub

    Private Sub TreeView1_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick
        ' Ottieni il testo del nodo selezionato
        Dim nodoTesto As String = e.Node.Text
        ' Separa il testo fino al primo "-"
        Dim testoFinoAlTrattino As String = nodoTesto.Split("-"c)(0).Trim()

        If e.Node.Nodes.Count > 0 Then
            ' Il nodo è un nodo padre: apri la distinta base
            Distinta_base_form.Show()
            Distinta_base_form.TextBox1.Text = testoFinoAlTrattino
        Else
            ' Il nodo è un nodo figlio: apri la parte di magazzino
            Magazzino.Codice_SAP = testoFinoAlTrattino

            ' Ripristina la finestra se è minimizzata e portala in primo piano
            If Magazzino.WindowState = FormWindowState.Minimized Then
                Magazzino.WindowState = FormWindowState.Normal
            End If
            Magazzino.BringToFront()
            Magazzino.Activate()
            Magazzino.Show()

            ' Imposta il codice SAP e aggiorna i dettagli
            Magazzino.TextBox2.Text = Magazzino.Codice_SAP
            Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)
        End If
    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        ' Verifica se il tasto premuto è "Invio"
        If e.KeyCode = Keys.Enter Then
            Inserimento_stock(DataGridView, TextBox1.Text, TextBox2.Text, True, "")
            ' Previene il beep standard del tasto Invio
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        ' Verifica se il tasto premuto è "Invio"
        If e.KeyCode = Keys.Enter Then
            Inserimento_stock(DataGridView, TextBox1.Text, TextBox2.Text, True, "")
            ' Previene il beep standard del tasto Invio
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        ' Verifica se esiste un nodo selezionato nella TreeView
        If TreeView1.SelectedNode Is Nothing Then
            MessageBox.Show("Seleziona un nodo prima di procedere.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        ' Rendi visibile la colonna "PARENTE" nella DataGridView
        If DataGridView.Columns.Contains("PARENTE") Then
            DataGridView.Columns("PARENTE").Visible = True
        End If
        DataGridView.Rows.Clear()
        ' Ottieni il testo del nodo selezionato
        Dim nodoTesto As String = TreeView1.SelectedNode.Text
        ' Separa il testo fino al primo "-"
        Dim testoFinoAlTrattino As String = nodoTesto.Split("-"c)(0).Trim()

        ' Controlla se il nodo selezionato è un nodo padre o figlio
        If TreeView1.SelectedNode.Nodes.Count > 0 Then


            Inserimento_stock(DataGridView, testoFinoAlTrattino, "", False, "PADRE")
            ' Il nodo è un nodo padre: apri la distinta base
            ' Esegui la funzione Inserimento_stock per ciascun nodo figlio
            For Each childNode As TreeNode In TreeView1.SelectedNode.Nodes
                Dim childTesto As String = childNode.Text.Split("-"c)(0).Trim()
                Inserimento_stock(DataGridView, childTesto, "", False, "FIGLIO")
            Next

        Else

            Inserimento_stock(DataGridView, testoFinoAlTrattino, "", False, "FIGLIO")
            ' Il nodo è un nodo padre: apri la distinta base
        End If
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Inserimento_stock(DataGridView, TextBox1.Text, TextBox2.Text, False, "FIGLIO")
    End Sub
End Class