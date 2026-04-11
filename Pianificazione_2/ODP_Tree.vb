Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib




Public Structure Risultato_ODP
    Public Num_ODP As String
    Public Tipo As String
    Public Num_Cassetta As String
    Public lotto_prelievo As Integer
    Public disegno As String
    Public qta As String
    Public commessa As String
    Public stato As String
    Public Num_Articoli As Integer

End Structure

Public Class ODP_Tree

    Private isShiftKeyDown As Boolean = False
    Private startIndex As Integer = -1

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        ' Controlla se il tasto Shift è stato premuto
        isShiftKeyDown = (e.KeyCode = Keys.ShiftKey)
    End Sub

    Private Sub DataGridView1_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyUp
        ' Controlla se il tasto Shift è stato rilasciato
        isShiftKeyDown = (e.KeyCode <> Keys.ShiftKey)
    End Sub

    Private Sub DataGridView1_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDown
        If e.Button = MouseButtons.Left AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If ModifierKeys = Keys.Shift AndAlso startIndex >= 0 Then
                'Se è premuto Shift, cambia il flag per le righe comprese tra startIndex ed e.RowIndex
                Dim endIndex As Integer = e.RowIndex
                Dim minIndex As Integer = Math.Min(startIndex, endIndex) + 1
                Dim maxIndex As Integer = Math.Max(startIndex, endIndex) - 1

                For i As Integer = minIndex To maxIndex
                    DataGridView1.Rows(i).SetValues(True)
                Next i
            Else
                '  Altrimenti, imposta startIndex alla riga corrente
                startIndex = e.RowIndex
            End If
        End If
    End Sub
    Private Sub ODP_Tree_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Txt_DocNum.Text = Ordine_di_produzione_lista.Trova_nuovo_lotto_di_prelievo()
    End Sub

    Public Sub Compila_Albero(par_codice As String, PAR_TIPO_APPOGGIO As String, par_solo_gruppi As Boolean)
        PULISCI_APPOGGIO(Homepage.ID_SALVATO, PAR_TIPO_APPOGGIO)
        TV_Progetto.Nodes.Clear()
        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT T0.[DocNum], T0.[ProdName], T0.[DocNum], T0.[U_PRG_AZS_Commessa],T0.[Status] 
FROM OWOR T0 
LEFT JOIN [TIRELLI_40].[DBO].APPOGGIO T1 ON T0.DOCNUM=T1.VALORE AND T1.TIPO='" & PAR_TIPO_APPOGGIO & "' AND T1.UTENTE=" & Homepage.ID_SALVATO & "

WHERE T0.[ItemCode] = '" & par_codice & "' 
AND (T0.[Status]='R' OR T0.[Status]='P') 
and t0.u_produzione='ASSEMBL' AND T1.VALORE IS NULL"

        Reader_Tree = Cmd_Tree.ExecuteReader()

        If Reader_Tree.Read() Then
            AGGIUNGI_RECORD_APPOGGIO(Homepage.ID_SALVATO, PAR_TIPO_APPOGGIO, Reader_Tree("DocNum"))
            TV_Progetto.Nodes.Add(Reader_Tree("DocNum") & "-" & Reader_Tree("ProdName"))
            Trova_Figli(TV_Progetto.Nodes(0).Nodes, Reader_Tree("DocNum"), par_codice, PAR_TIPO_APPOGGIO, par_solo_gruppi)
        End If
        TV_Progetto.ExpandAll()
        Cnn_Tree.Close()
    End Sub



    Private Function Trova_Figli(Nodi As TreeNodeCollection, ODP As String, par_commessa As String, PAR_TIPO_APPOGGIO As String, par_solo_gruppi As Boolean) As Integer

        Dim filtro_solo_gruppi As String
        If par_solo_gruppi = True Then
            filtro_solo_gruppi = " And substring(t1.itemcode,1,1)='0' "
        Else
            filtro_solo_gruppi = " "
        End If

        Dim Nodi_Count As Integer
        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT 	T0.ItemCode AS 'Cod. articolo', T1.ItemName AS 'Descrizione articolo', T2.DocNum AS 'Numero documento', T2.ItemCode AS 'Cod. articolo', 
	T0.PlannedQty AS 'Base', T0.U_PRG_WIP_QtaDaTrasf as 'Da Trasferire', T0.U_PRG_WIP_QtaSpedita as 'Trasferito'

FROM  [dbo].[OITM] T1 INNER JOIN [dbo].[WOR1] T0 ON T1.ItemCode = T0.ItemCode
INNER JOIN [dbo].[OWOR] T2 ON T2.DocEntry = T0.DocEntry 


WHERE T2.DocNum='" & ODP & "' AND T2.U_PRODUZIONE='ASSEMBL' " & filtro_solo_gruppi & " 

ORDER BY T0.ItemCode,T1.ItemName,T2.DocNum,T2.Status"

        Reader_Tree = Cmd_Tree.ExecuteReader()
        Nodi_Count = 0
        Do While Reader_Tree.Read()
            Dim Risultato As Risultato_ODP
            Risultato = Trova_ODP_Appropriato(Reader_Tree("Cod. articolo"), par_commessa, PAR_TIPO_APPOGGIO)

            Nodi.Add(Risultato.Num_ODP & " - " & Reader_Tree("Cod. articolo") & " - " & Reader_Tree("Descrizione articolo") & " - " & Risultato.Tipo & " " & Risultato.Num_Cassetta)

            If Risultato.Num_ODP <> "*" Then
                AGGIUNGI_RECORD_APPOGGIO(Homepage.ID_SALVATO, PAR_TIPO_APPOGGIO, Risultato.Num_ODP)
                Trova_Figli(Nodi(Nodi_Count).Nodes, Risultato.Num_ODP, par_commessa, PAR_TIPO_APPOGGIO, par_solo_gruppi)
            End If
            Nodi_Count = Nodi_Count + 1
        Loop
        TV_Progetto.ExpandAll()
        Cnn_Tree.Close()

        Return Nodi_Count
    End Function

    Private Function Trova_ODP_Appropriato(Codice As String, par_commessa As String, par_tipo_appoggio As String) As Risultato_ODP

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Dim Risultato As Risultato_ODP

        Cmd_Tree.CommandText = "SELECT 	T1.DocNum AS 'Numero documento', T0.ItemCode AS 'Cod. articolo',
T1.PlannedQty As 'Quantità', T0.ItemName AS 'Descrizione articolo', T1.U_PRODUZIONE AS 'Reparto', 
	T1.U_PRG_AZS_Commessa,T1.OriginNum, T1.U_UTILIZZ AS 'Rif a Cliente.'
, T1.Status AS 'Stato'
,coalescE(T1.[U_Progressivo_commessa],0) As 'Cassetta'
,coalescE(A.[ID],0) As 'lotto'
,coalesce(t1.u_disegno,'') as 'Disegno'

FROM  [dbo].[OITM] T0 inner join [dbo].[OWOR] T1 on t0.itemcode=t1.itemcode 
left join (select  max(t0.id) as 'ID', t0.docnum 
from [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 inner join owor t1 on t0.docnum=t1.docnum
where t1.itemcode='" & Codice & "'
group by t0.docnum ) A on A.docnum=t1.docnum
LEFT JOIN [TIRELLI_40].[DBO].APPOGGIO T2 ON T1.DOCNUM=T2.VALORE AND T2.TIPO='" & par_tipo_appoggio & "' AND T2.UTENTE=" & Homepage.ID_SALVATO & "

WHERE   (T1.Status <> N'L' )  AND  (T1.Status <> N'C' ) AND T0.ItemCode='" & Codice & "' AND T1.U_PRG_AZS_Commessa='" & par_commessa & "' and t1.u_produzione='ASSEMBL' AND T2.VALORE IS NULL"

        Reader_Tree = Cmd_Tree.ExecuteReader()
        If Reader_Tree.Read() Then
            Risultato.Num_ODP = Reader_Tree("Numero documento")
            Risultato.Tipo = "Sul Carrello : "
            Risultato.Num_Cassetta = Reader_Tree("Cassetta")
            Risultato.qta = Reader_Tree("Quantità")
            Risultato.disegno = Reader_Tree("Disegno")
            Risultato.lotto_prelievo = Reader_Tree("lotto")
            Risultato.commessa = Reader_Tree("U_PRG_AZS_Commessa")
            Risultato.stato = Reader_Tree("stato")

            Cnn_Tree.Close()
            Return Risultato
        End If

        Cnn_Tree.Close()
        Cnn_Tree.Open()
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT 	T1.DocNum AS 'Numero documento', T0.ItemCode AS 'Cod. articolo',T1.PlannedQty As 'Quantità', T0.ItemName AS 'Descrizione articolo', T0.CodeBars As 'Disegno', T1.U_PRODUZIONE AS 'Reparto', 
	T1.U_PRG_AZS_Commessa,T1.OriginNum, T1.U_UTILIZZ AS 'Rif a Cliente.', T1.Status AS 'Stato',case when T1.[U_Progressivo_commessa] is null then 0 else  T1.[U_Progressivo_commessa] end As 'Cassetta'

FROM  [dbo].[OITM] T0 inner join  [dbo].[OWOR] T1 on t0.itemcode=t1.itemcode
LEFT JOIN [TIRELLI_40].[DBO].APPOGGIO T2 ON T1.DOCNUM=T2.VALORE AND T2.TIPO='" & par_tipo_appoggio & "' AND T2.UTENTE=" & Homepage.ID_SALVATO & "
WHERE (T1.Status <> N'L' )  AND  (T1.Status <> N'C' ) AND T0.ItemCode='" & Codice & "' AND T1.U_PRG_AZS_Commessa='SCORTA' AND t1.u_produzione='ASSEMBL' AND T2.VALORE IS NULL"

        Reader_Tree = Cmd_Tree.ExecuteReader()
        If Reader_Tree.Read() Then
            Risultato.Num_ODP = Reader_Tree("Numero documento")
            Risultato.Tipo = "Premontaggio-Scorta : "
            Risultato.Num_Cassetta = Reader_Tree("Cassetta")
            Cnn_Tree.Close()
            Return Risultato
        End If
        Cnn_Tree.Close()

        Risultato.Num_ODP = "*"
        Risultato.Tipo = "Premontato"
        Risultato.Num_Cassetta = ""
        Cnn_Tree.Close()
        Return Risultato
    End Function



    Private Sub Cmd_Exit_Click(sender As Object, e As EventArgs) Handles Cmd_Exit.Click
        Me.Close()
    End Sub

    Private Sub TV_Progetto_NodeMouseClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TV_Progetto.NodeMouseClick
        If Microsoft.VisualBasic.Left(e.Node.Text, 1) <> "*" Then
            TXT_ODP.Text = Microsoft.VisualBasic.Left(e.Node.Text, 6)

            'Compila_Albero_datagrid(Lbl_Matricola.Text, TXT_ODP.Text)
            DataGridView1.Rows.Clear()
            PULISCI_APPOGGIO(Homepage.ID_SALVATO, "ODP_TREE")
            Trova_Figli_datagrid(DataGridView1, TXT_ODP.Text, 0, "", Lbl_Matricola.Text, "ODP_TREE", CheckBox3.Checked)

        End If

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        If TXT_ODP.Text <> "" Then


            ODP_Form.docnum_odp = TXT_ODP.Text
            ODP_Form.Show()
            ODP_Form.inizializza_form(TXT_ODP.Text)

        End If
    End Sub

    Private Function Trova_Figli_datagrid(DGV As DataGridView, ODP As String, par_livello As Integer, par_stringa As String, par_commessa As String, par_tipo_appoggio As String, par_filtro_solo_gruppi As Boolean) As Integer

        Dim filtro_solo_gruppi As String

        If par_filtro_solo_gruppi = True Then
            filtro_solo_gruppi = " and substring(t0.itemcode,1,1)='0' "
        Else
            filtro_solo_gruppi = ""
        End If

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT T0.ItemCode AS 'Cod. articolo', T1.ItemName AS 'Descrizione articolo', T2.DocNum AS 'Numero documento', T2.ItemCode AS 'Cod. articolo', 
                            T0.PlannedQty AS 'Base', T0.U_PRG_WIP_QtaDaTrasf as 'Da Trasferire', T0.U_PRG_WIP_QtaSpedita as 'Trasferito'
,coalesce(t2.u_prg_azs_commessa,'') as 'Commessa'
,t2.status
,COALESCE(T1.U_PRG_TIR_EXPLOSION,'N') AS 'Fantasma'
FROM  [dbo].[OITM] T1 INNER JOIN [dbo].[WOR1] T0 ON T1.ItemCode = T0.ItemCode
INNER JOIN [dbo].[OWOR] T2 ON T2.DocEntry = T0.DocEntry 

                            
WHERE T2.DocNum='" & ODP & "' AND T2.U_PRODUZIONE='ASSEMBL' " & filtro_solo_gruppi & "
                            ORDER BY T0.ItemCode,T1.ItemName,T2.DocNum,T2.Status"

        Reader_Tree = Cmd_Tree.ExecuteReader()
        Dim rowCounter As Integer = 0
        Do While Reader_Tree.Read()
            Dim Risultato As Risultato_ODP
            Risultato = Trova_ODP_Appropriato(Reader_Tree("Cod. articolo"), par_commessa, par_tipo_appoggio)


            If Risultato.Num_ODP <> "*" Then
                AGGIUNGI_RECORD_APPOGGIO(Homepage.ID_SALVATO, "ODP_TREE", Risultato.Num_ODP)
                DGV.Rows.Add(False, par_stringa & par_livello, Risultato.Num_ODP, Risultato.commessa, Reader_Tree("Cod. articolo"), Reader_Tree("Descrizione articolo"), Risultato.disegno, Risultato.stato, Risultato.qta, Risultato.Tipo & " " & Risultato.Num_Cassetta, Risultato.lotto_prelievo)

                Trova_Figli_datagrid(DGV, Risultato.Num_ODP, par_livello + 1, par_stringa & "+", par_commessa, par_tipo_appoggio, par_filtro_solo_gruppi)
            End If
            rowCounter += 1
        Loop

        Cnn_Tree.Close()
        Return rowCounter
    End Function

    ' Call this function to set up the DataGridView with columns
    Private Sub SetupDataGridViewColumns()
        DataGridView1.Columns.Add("NumeroDocumento", "Numero documento")
        DataGridView1.Columns.Add("CodiceArticolo", "Cod. articolo")
        DataGridView1.Columns.Add("DescrizioneArticolo", "Descrizione articolo")
        DataGridView1.Columns.Add("TipoCassetta", "Tipo Cassetta")
    End Sub

    Sub PULISCI_APPOGGIO(PAR_UTENTE As Integer, PAR_TIPO As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = homepage.sap_tirelli


        cnn.Open()

        Dim Cmd_SAP As New SqlCommand



        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "DELETE [TIRELLI_40].[DBO].APPOGGIO WHERE UTENTE='" & PAR_UTENTE & "' AND TIPO='" & PAR_TIPO & "'"

        Cmd_SAP.ExecuteNonQuery()

        cnn.Close()


    End Sub

    Sub AGGIUNGI_RECORD_APPOGGIO(PAR_UTENTE As Integer, PAR_TIPO As String, PAR_VALORE As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli


        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand



        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO [Tirelli_40].[dbo].[Appoggio]
           ([Utente]
           ,[Tipo]
           ,[valore])
     VALUES
           (" & PAR_UTENTE & "
           ,'" & PAR_TIPO & "'
           ,'" & PAR_VALORE & "') "

        Cmd_SAP.ExecuteNonQuery()

        Cnn.Close()


    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Itera attraverso le righe della DataGridView "datagridview_odp"
        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Verifica se la cella della colonna "seleziona" è flaggata
            If CBool(row.Cells("seleziona").Value) = True Then
                ' Crea una nuova riga nella DataGridView "datagridview1"
                Dim index As Integer = DataGridView2.Rows.Add()

                ' Copia i valori dalle colonne necessarie
                DataGridView2.Rows(index).Cells("Numero_ODP").Value = row.Cells("N_ODP").Value
                DataGridView2.Rows(index).Cells("Commessa_odp").Value = row.Cells("Commessa").Value
                DataGridView2.Rows(index).Cells("disegno_odp").Value = row.Cells("Disegno").Value
                DataGridView2.Rows(index).Cells("Stato_odp").Value = row.Cells("stato").Value
                row.Cells("seleziona").Value = False
            End If
        Next
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Itera all'indietro attraverso le righe selezionate nella DataGridView "datagridview1"
        For i As Integer = DataGridView1.SelectedRows.Count - 1 To 0 Step -1
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(i)
            DataGridView1.Rows.Remove(row)
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click




        ODP_Form.stampa_etichetta = "NO"


        For Each Row As DataGridViewRow In DataGridView1.Rows





            If CheckBox1.Checked = True Then

                ODP_Form.stampa_etichetta = "YES"
                FORM6.ODP = Row.Cells("Numero_ODP").Value
                ODP_Form.docnum_odp = Row.Cells("Numero_ODP").Value
                ODP_Form.percorso_documento = Homepage.PERCORSO_DOCUMENTO_ODP_ETICHETTA
                ODP_Form.Genera_ordine()

            ElseIf CheckBox2.Checked = True Then
                ODP_Form.testata_odp(Row.Cells("Numero_ODP").Value)
                ODP_Form.Fun_Stampa()


            End If


        Next
        MsgBox("FINE STAMPE")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Ordine_di_produzione_lista.check_esistenza_lotto_di_prelievo(CInt(Txt_DocNum.Text), DataGridView2)
    End Sub

    Private Sub Cmd_Indietro_Click(sender As Object, e As EventArgs) Handles Cmd_Indietro.Click
        Txt_DocNum.Text = Int(Txt_DocNum.Text) - 1
        Ordine_di_produzione_lista.RIEMPI_datagridview_documenti_lotto(Txt_DocNum.Text, DataGridView2, DataGridView2)
    End Sub

    Private Sub Cmd_Avanti_Click(sender As Object, e As EventArgs) Handles Cmd_Avanti.Click
        Txt_DocNum.Text = Int(Txt_DocNum.Text) + 1
        Ordine_di_produzione_lista.RIEMPI_datagridview_documenti_lotto(Txt_DocNum.Text, DataGridView2, DataGridView2)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Ordine_di_produzione_lista.RIEMPI_datagridview_documenti_lotto(Txt_DocNum.Text, DataGridView2, DataGridView2)
    End Sub

    Private Sub TV_Progetto_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TV_Progetto.AfterSelect

    End Sub
End Class