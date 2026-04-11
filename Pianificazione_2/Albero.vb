Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports BrightIdeasSoftware

Imports System.Windows.Forms
Imports TenTec.Windows.iGridLib
Imports System.ComponentModel
Imports ADGV
Imports System.Windows.Documents
Imports System.Reflection.Emit




Public Class Albero

    Dim dataTable As New DataTable()

    Public commessa As String
    Public livello As Integer = 1
    Public tipo_esplosione As String = "Parziale"
    Private filtro_query As String = ""
    Private filtro_odp As String = ""
    ' Private operazione As String
    Private isShiftKeyDown As Boolean = False
    Private startIndex As Integer = -1

    ' Scrivi i dati su Excel
    Public excelApp As New Microsoft.Office.Interop.Excel.Application()
    Public excelWorkbook As Microsoft.Office.Interop.Excel.Workbook
    Public excelWorksheet As Microsoft.Office.Interop.Excel.Worksheet
    Private Codice_sap As String
    Private quantità_Selezionata As Decimal

    Sub inizializza_albero(par_commessa As String)
        Compila_Albero_treeview(commessa, CheckBox2.Checked, Tree_Boom)
    End Sub

    Public Sub Compila_Albero_treeview(par_commessa As String, par_solo_gruppi As Boolean, par_treeview As TreeView)

        Dim filtro_solo_gruppi As String
        If par_solo_gruppi = True Then
            filtro_solo_gruppi = " and substring(t0.code,1,1)='0' "
        Else
            filtro_solo_gruppi = ""
        End If

        If tipo_esplosione = "Parziale" Then
            filtro_query = "and (substring(t0.father,1,1)='0' or substring(t0.father,1,1)='M' )"
        Else
            filtro_query = ""
        End If

        par_treeview.Nodes.Clear()
        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT T0.[father], T0.[code], T1.[itemname] as 'Nome_padre'
FROM itt1 T0 inner join oitm t1 on t1.itemcode=t0.father

WHERE T0.[father] = '" & par_commessa & "' " & filtro_solo_gruppi & "
ORDER BY T0.VisOrder"

        Reader_Tree = Cmd_Tree.ExecuteReader()

        If Reader_Tree.Read() Then
            par_treeview.Nodes.Add(Reader_Tree("father") & "-" & Reader_Tree("nome_padre"))
            Trova_Figli(par_treeview.Nodes(0).Nodes, Reader_Tree("father"), par_solo_gruppi)
        End If
        par_treeview.ExpandAll()
        Cnn_Tree.Close()
    End Sub

    Private Function Trova_Figli(Nodi As TreeNodeCollection, par_codice As String, par_solo_gruppi As Boolean) As Integer

        Dim filtro_solo_gruppi As String
        If par_solo_gruppi = True Then
            filtro_solo_gruppi = " and substring(t0.code,1,1)='0' "
        Else
            filtro_solo_gruppi = ""
        End If

        Dim Nodi_Count As Integer
        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT T0.[father], T0.[code],substring(t0.code,1,1) as 'Prima_lettera', T1.[itemname] as 'Nome_padre',coalesce(t1.u_disegno,'') as 'Disegno', T2.[itemname] as 'Nome_figlio' 
,T0.[Quantity]
FROM itt1 T0 inner join oitm t1 on t1.itemcode=t0.father
inner join oitm t2 on t2.itemcode=t0.code

WHERE T0.[father] = '" & par_codice & "' " & filtro_query & " " & filtro_solo_gruppi & "
ORDER BY T0.VisOrder"

        Reader_Tree = Cmd_Tree.ExecuteReader()
        Nodi_Count = 0
        Do While Reader_Tree.Read()



            ' Aggiungi il nodo al TreeView
            Dim newNode As New TreeNode(Reader_Tree("code") & " - " & Reader_Tree("nome_figlio") & " - " & Reader_Tree("Disegno") & " Q : = " & Reader_Tree("Quantity"))

            ' Assegna l'immagine in base alla prima lettera del codice
            If Reader_Tree("prima_lettera") = "C" Then
                newNode.ImageIndex = 2 ' Imposta l'indice dell'immagine per la prima lettera '1'
            ElseIf Reader_Tree("prima_lettera") = "D" Then
                newNode.ImageIndex = 1 ' Imposta l'indice dell'immagine per la prima lettera '2'

            ElseIf Reader_Tree("prima_lettera") = "0" Then
                newNode.ImageIndex = 0
            ElseIf Reader_Tree("prima_lettera") = "R" Then
                newNode.ImageIndex = 3

            End If
            Nodi.Add(newNode)


            If Trova_esistenza_distinta_base(Reader_Tree("code")) = "Y" Then
                Trova_Figli(Nodi(Nodi_Count).Nodes, Reader_Tree("code"), par_solo_gruppi)
            End If


            Nodi_Count = Nodi_Count + 1
        Loop
        Tree_Boom.ExpandAll()
        Cnn_Tree.Close()

        Return Nodi_Count
    End Function

    Public Sub Compila_DataGridView(dataGridView As DataGridView, par_commessa As String, par_selected_quantità As Decimal)
        If tipo_esplosione = "Parziale" Then
            filtro_query = "and (substring(t0.father,1,1)='0' or substring(t0.father,1,1)='M' )"
        Else
            filtro_query = ""
        End If

        dataGridView.Rows.Clear()

        Dim dt As New DataTable

        Using Cnn_Tree As New SqlConnection(Homepage.sap_tirelli)
            Cnn_Tree.Open()
            Dim command As String = "SELECT T0.[father], T0.[code], T1.[itemname] as 'Nome_padre', sum(coalesce(t2.onhand,0)+coalesce(t2.onorder,0)-coalesce(t2.iscommited,0)) as 'Disp'
,coalesce(t1.u_disegno,'') as 'u_disegno'
,t4.price
,coalesce(t1.u_codice_brb,'') as 'Codice_BRB'
,coalesce(t1.u_prg_tir_explosion,'N') as 'Fantasma'
,coalesce(t1.u_tipo_montaggio,'') as 'Tipo_montaggio'
,coalesce(t1.U_PRG_TIR_Materiale,'') as 'Materiale'
,coalesce(t5.ItmsGrpNam,'') as 'Gruppo_articoli'
,coalesce(t1.U_PRG_TIR_Trattamento,'') as 'Trattamento'
,case when t6.code is null then 'N' else 'Y' end as 'Padre'
, COALESCE(t1.minlevel,0) as 'Minimo'
,COALESCE(t1.minordrqty,0) AS 'minordrqty'
,coalesce(t1.PrcrmntMtd,'M') as 'MAke'

                                          FROM itt1 T0 
                                          INNER JOIN oitm t1 ON t1.itemcode=t0.father
										  inner join oitw t2 on t0.father=t2.itemcode
										  inner join owhs t3 on t3.whscode=t2.whscode 
                                          inner join itm1 t4 on t4.itemcode=t0.father and t4.pricelist=2
                                          INNER JOIN OITB T5 ON T5.[ItmsGrpCod] = T1.[ItmsGrpCod]
                                          left join oitt t6 on t6.code=t0.father

                                          WHERE T0.[father] = '" & par_commessa & "' 
                                           group by t0.visorder,t1.u_tipo_montaggio,t1.PrcrmntMtd, T0.[father], T0.[code], T1.[itemname],t1.u_disegno, t4.price,t1.u_codice_brb,t1.u_prg_tir_explosion,t1.U_PRG_TIR_Materiale,t5.ItmsGrpNam,t1.U_PRG_TIR_Trattamento,t1.minlevel,t1.minordrqty, t6.code
                                          ORDER BY T0.VisOrder"

            Using Cmd_Tree As New SqlCommand(command, Cnn_Tree)

                Using Reader_Tree As SqlDataReader = Cmd_Tree.ExecuteReader()
                    dt.Load(Reader_Tree)
                End Using
            End Using
        End Using

        If dt.Rows.Count > 0 Then
            Dim row As DataRow = dt.Rows(0)
            dataGridView.Rows.Add(False, 0, 0, row("father"), row("nome_padre"), row("u_disegno"), row("Codice_BRB"), row("Fantasma"), row("Materiale"), row("Gruppo_articoli"), row("Trattamento"), par_selected_quantità, row("Disp"), par_selected_quantità, row("price"), row("padre"), row("minimo"), row("minordrqty"), row("tipo_montaggio"), row("Make"))
            Trova_Figli_DataGridView(dataGridView, row("father"), 0, 1, par_selected_quantità)
        End If
    End Sub

    Private Sub Trova_Figli_DataGridView(dataGridView As DataGridView, par_codice As String, ByRef nodoCount As Integer, ByVal livello As Integer, par_selected_quantità As Decimal)
        Dim dt As New DataTable

        Using Cnn_Tree As New SqlConnection(Homepage.sap_tirelli)
            Dim COMMAND As String = "SELECT T0.[father], T0.[code]
, substring(t0.code,1,1) as 'Prima_lettera'
, T1.[itemname] as 'Nome_padre'
, coalesce(t2.u_disegno,'') as 'Disegno'
, T2.[itemname] as 'Nome_figlio'
, T0.[Quantity]
, sum(coalesce(t3.onhand,0)+coalesce(t3.onorder,0)-coalesce(t3.iscommited,0)) as 'Disp'
,t5.price
,coalesce(t2.u_codice_brb,'') as 'Codice_BRB'
,coalesce(t2.u_prg_tir_explosion,'') as 'Fantasma'
,coalesce(t2.U_PRG_TIR_Materiale,'') as 'Materiale'
,coalesce(t6.ItmsGrpNam,'') as 'Gruppo_articoli'
,coalesce(t2.U_PRG_TIR_Trattamento,'') as 'Trattamento'
,  COALESCE(t2.minlevel,0)  as 'Minimo'
,COALESCE(t2.minordrqty,0) AS 'minordrqty'
, case when t7.code is null then 'N' else 'Y' end as 'Padre'
, coalesce(t1.u_tipo_montaggio,'') as 'Tipo_montaggio'
,coalesce(t2.PrcrmntMtd,'M') as 'MAke'

                                      FROM itt1 T0 
                                      INNER JOIN oitm t1 ON t1.itemcode=t0.father
                                      INNER JOIN oitm t2 ON t2.itemcode=t0.code
									   inner join oitw t3 on t0.code=t3.itemcode
										  inner join owhs t4 on t4.whscode=t3.whscode 
                                          inner join itm1 t5 on t5.itemcode=t0.code and t5.pricelist=2
                                          INNER JOIN OITB T6 ON T6.[ItmsGrpCod] = T2.[ItmsGrpCod]
                                          left join oitt t7 on t7.code=t0.code

                                      WHERE T0.[father] = '" & par_codice & "' " & filtro_query & " " & filtro_odp & "
                                      group by T0.[father], T0.[code]

, T1.[itemname] 
,t1.u_tipo_montaggio
, t2.u_disegno
, T2.[itemname] 
, T0.[Quantity]
,T0.VisOrder
,t5.price
,t2.u_codice_brb
,t2.u_prg_tir_Explosion
,t2.U_PRG_TIR_Materiale
,t6.ItmsGrpNam
,t2.U_PRG_TIR_Trattamento
,t2.minlevel
,t2.minordrqty
,t7.code
,t2.PrcrmntMtd
									  ORDER BY T0.VisOrder"
            Cnn_Tree.Open()
            Using Cmd_Tree As New SqlCommand(COMMAND, Cnn_Tree)
                Using Reader_Tree As SqlDataReader = Cmd_Tree.ExecuteReader()
                    dt.Load(Reader_Tree)
                End Using
            End Using
        End Using

        For Each row As DataRow In dt.Rows
            Dim newRow As DataGridViewRow = New DataGridViewRow()
            newRow.CreateCells(dataGridView)

            nodoCount += 1
            Dim colonna As Integer = 0
            newRow.Cells(colonna).Value = False
            colonna += 1
            newRow.Cells(colonna).Value = nodoCount.ToString() ' Numero del nodo
            colonna += 1
            newRow.Cells(colonna).Value = New String("+"c, livello) & " " & livello.ToString() ' Livello del nodo
            'newRow.Cells(colonna).Value = livello.ToString() ' Livello del nodo
            colonna += 1

            ' Assegna l'immagine in base alla prima lettera del codice
            If row("prima_lettera") = "C" Then
                newRow.Cells(2).Style.BackColor = Color.Yellow ' Imposta lo sfondo in giallo per la prima lettera 'C'
            ElseIf row("prima_lettera") = "D" Then
                newRow.Cells(2).Style.BackColor = Color.LightBlue ' Imposta lo sfondo in azzurro per la prima lettera 'D'
            ElseIf row("prima_lettera") = "0" Then
                newRow.Cells(2).Style.BackColor = Color.White ' Imposta lo sfondo bianco per la prima lettera '0'
            ElseIf row("prima_lettera") = "R" Then
                newRow.Cells(2).Style.BackColor = Color.LightGreen ' Imposta lo sfondo verde chiaro per la prima lettera 'R'
            End If

            newRow.Cells(colonna).Value = row("code")
            colonna += 1
            newRow.Cells(colonna).Value = row("nome_figlio")
            colonna += 1
            newRow.Cells(colonna).Value = row("Disegno")
            colonna += 1
            newRow.Cells(colonna).Value = row("Codice_BRB")
            colonna += 1
            newRow.Cells(colonna).Value = row("Fantasma")
            colonna += 1
            newRow.Cells(colonna).Value = row("Materiale")
            colonna += 1
            newRow.Cells(colonna).Value = row("Gruppo_articoli")
            colonna += 1
            newRow.Cells(colonna).Value = row("Trattamento")
            colonna += 1


            ' Assicurati che ci siano almeno 6 colonne nella DataGridView

            newRow.Cells(colonna).Value = row("Quantity")
            colonna += 1
            newRow.Cells(colonna).Value = row("Disp")
            colonna += 1

            newRow.Cells(colonna).Value = row("Quantity") * par_selected_quantità
            colonna += 1
            newRow.Cells(colonna).Value = row("price")
            colonna += 1
            newRow.Cells(colonna).Value = row("padre")
            colonna += 1
            newRow.Cells(colonna).Value = row("minimo")
            colonna += 1
            newRow.Cells(colonna).Value = row("minordrqty")
            colonna += 3

            newRow.Cells(colonna).Value = row("tipo_montaggio")
            colonna += 1
            newRow.Cells(colonna).Value = row("Make")

            dataGridView.Rows.Add(newRow)

            If Trova_esistenza_distinta_base(row("code")) = "Y" Then
                Trova_Figli_DataGridView(dataGridView, row("code"), nodoCount, livello + 1, row("Quantity") * par_selected_quantità)
            End If
        Next
    End Sub



    Private Sub UpdateDataGridView(selectedCode As String, selectedQuantita As Decimal)
        ' Clear the existing data in the DataGridView
        DataGridView1.Rows.Clear()

        ' Use your existing logic or modify it as needed to populate the DataGridView based on the selected code
        ' For example, you can execute a SQL query to get the data related to the selected code

        ' ...

        ' Add rows to the DataGridView
        Compila_DataGridView(DataGridView1, selectedCode, selectedQuantita)




        ' ...
    End Sub

    Public Function Trova_esistenza_distinta_base(Codice As String)

        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Dim Risultato As String = "N"

        Cmd_Tree.CommandText = "select t0.code
from oitt t0 where t0.code='" & Codice & "'"

        Reader_Tree = Cmd_Tree.ExecuteReader()
        If Reader_Tree.Read() Then
            Risultato = "Y"

        End If
        'Cmd_Tree.ExecuteReader.Close()
        Cnn_Tree.Close()

        Return Risultato
    End Function



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub



    Private Sub Albero_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Inizializza il form
        '  InitializeComponent()
        Me.BackColor = Homepage.colore_sfondo


        Acquisti.Inserimento_fasi(ComboBox4)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Tree_Boom.CollapseAll()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Tree_Boom.ExpandAll()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Compila_Albero_treeview(commessa, CheckBox2.Checked, Tree_Boom)
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            tipo_esplosione = "Totale"
        Else
            tipo_esplosione = "Parziale"
        End If
    End Sub

    Public Sub Compila_Albero_excel(par_commessa As String)



        ' Apri un nuovo foglio di lavoro o carica un esistente
        excelWorkbook = excelApp.Workbooks.Add()
        excelWorksheet = excelWorkbook.Sheets(1)

        If tipo_esplosione = "Parziale" Then
            filtro_query = "and (substring(t0.father,1,1)='0' or substring(t0.father,1,1)='M' )"
        Else
            filtro_query = ""
        End If


        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT T0.[father], T0.[code], T1.[itemname] as 'Nome_padre'
FROM itt1 T0 inner join oitm t1 on t1.itemcode=t0.father

WHERE T0.[father] = '" & par_commessa & "'
ORDER BY T0.VisOrder"

        Reader_Tree = Cmd_Tree.ExecuteReader()
        Dim riga As Integer = 1

        If Reader_Tree.Read() Then

            ' Aggiungi i dati del nuovo nodo alla prima riga del foglio di lavoro
            excelWorksheet.Cells(riga, 1).Value = riga
            excelWorksheet.Cells(riga, 2).Value = "Code"
            excelWorksheet.Cells(riga, 3).Value = "Nome Figlio"
            excelWorksheet.Cells(riga, 4).Value = "Prima Lettera"
            excelWorksheet.Cells(riga, 5).Value = "Image Index"

            Trova_Figli_excel(riga, Reader_Tree("father"))
        End If
        Tree_Boom.ExpandAll()
        Cnn_Tree.Close()
    End Sub

    Public Function controllo_disponibile(par_codice As String, par_quantità_lancio As String)


        Dim ritorno As String = ""

        Dim Cnn_Ticket As New SqlConnection

        Cnn_Ticket.ConnectionString = Homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader
        Cmd_Ticket.Connection = Cnn_Ticket

        Cmd_Ticket.CommandText = "
select t0.itemcode,t0.u_gestione_magazzino
, sum(coalesce(t1.onhand,0)+coalesce(t1.onorder,0)-coalesce(t1.iscommited,0)) as 'Disp'
, sum(coalesce(t1.onhand,0)+coalesce(t1.onorder,0)-coalesce(t1.iscommited,0)) + " & par_quantità_lancio & " AS 'Disp_con_lancio'
from oitm t0
inner join oitw t1 on t1.itemcode =t0.itemcode
inner join owhs t2 on t1.whscode=t2.whscode 
where t0.itemcode='" & par_codice & "'
group by t0.itemcode,t0.u_gestione_magazzino "

        Reader_Ticket = Cmd_Ticket.ExecuteReader()

        If Reader_Ticket.Read() Then
            If Reader_Ticket("u_gestione_magazzino") <> "COMMESSA" Then

                ritorno = ritorno & "Codice " & Reader_Ticket("itemcode") & " gestito a stock/scorta. lanciare a parte." & vbCrLf

            End If

            If Reader_Ticket("Disp_con_lancio") > 0 Then

                ritorno = ritorno & "Codice " & Reader_Ticket("itemcode") & " Lanciando ordine di " & par_quantità_lancio & " Disponibile >0." & vbCrLf

            End If

        End If
        Reader_Ticket.Close()
        Cnn_Ticket.Close()
        Return ritorno

    End Function

    Public Function disponibile_codice(par_codice As String)


        Dim ritorno As String = ""

        Dim Cnn_Ticket As New SqlConnection

        Cnn_Ticket.ConnectionString = Homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader
        Cmd_Ticket.Connection = Cnn_Ticket

        Cmd_Ticket.CommandText = "
select t0.itemcode, sum(coalesce(t1.onhand,0)+coalesce(t1.onorder,0)-coalesce(t1.iscommited,0)) as 'Disp'
from oitm t0
inner join oitw t1 on t1.itemcode =t0.itemcode
inner join owhs t2 on t1.whscode=t2.whscode 
where t0.itemcode='" & par_codice & "'
group by t0.itemcode"

        Reader_Ticket = Cmd_Ticket.ExecuteReader()

        If Reader_Ticket.Read() Then


            ritorno = Reader_Ticket("disp")

        End If


        Reader_Ticket.Close()
        Cnn_Ticket.Close()
        Return ritorno

    End Function





    Sub Trova_Figli_excel(par_riga As Integer, par_codice As String)
        Dim Nodi_Count As Integer
        Dim Cnn_Tree As New SqlConnection
        Cnn_Tree.ConnectionString = Homepage.sap_tirelli
        Cnn_Tree.Open()
        Dim Cmd_Tree As New SqlCommand
        Dim Reader_Tree As SqlDataReader
        Cmd_Tree.Connection = Cnn_Tree
        Cmd_Tree.CommandText = "SELECT T0.[father], T0.[code],substring(t0.code,1,1) as 'Prima_lettera', T1.[itemname] as 'Nome_padre', T2.[itemname] as 'Nome_figlio' 
FROM itt1 T0 inner join oitm t1 on t1.itemcode=t0.father
inner join oitm t2 on t2.itemcode=t0.code

WHERE T0.[father] = '" & par_codice & "' " & filtro_query & "
ORDER BY T0.VisOrder"

        Reader_Tree = Cmd_Tree.ExecuteReader()
        Nodi_Count = 0
        Do While Reader_Tree.Read()

            excelWorksheet.Cells(par_riga, 2).Value = Reader_Tree("code")
            excelWorksheet.Cells(par_riga, 3).Value = Reader_Tree("nome_figlio")
            excelWorksheet.Cells(par_riga, 4).Value = Reader_Tree("prima_lettera")
            excelWorksheet.Cells(par_riga, 5).Value = ""



            If Trova_esistenza_distinta_base(Reader_Tree("code")) = "Y" Then
                Trova_Figli_excel(par_riga, Reader_Tree("code"))
            End If


            par_riga = par_riga + 1


        Loop

        Cnn_Tree.Close()


    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form_lotto_di_prelievo.estrai_datagridview_in_excel(DataGridView1)
    End Sub

    ' Funzione di rilascio delle risorse COM
    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            filtro_odp = " AND substring(t0.code,1,1)='0' "
        Else
            filtro_odp = ""
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click



        If TextBox8.Text = "" Then
            MsgBox("Selezionare una commessa")
        ElseIf ComboBox4.SelectedIndex < 0 Or ComboBox4.Text = "" Or Acquisti.elenco_fasi(ComboBox4.SelectedIndex) = "" Then
            MsgBox("Selezionare una fase")
        ElseIf ComboBox5.SelectedIndex < 0 Then
            MsgBox("Selezionare una produzione")
        ElseIf ComboBox2.SelectedIndex < 0 Then
            MsgBox("Selezionare un magazzino destinazione")
        Else

            For Each row As DataGridViewRow In DataGridView1.Rows
                ' Verifica se la cella nella colonna "Seleziona" è True e "Fant" è "Y"
                If Convert.ToBoolean(row.Cells("Seleziona").Value) And row.Cells("Fant").Value = "Y" Then

                    Dim codice As String = row.Cells("Codice").Value
                    Dim descrizione As String = row.Cells("Column2").Value
                    Dim qTot As Double = Convert.ToDouble(row.Cells("Q_tot").Value, System.Globalization.CultureInfo.InvariantCulture)
                    Dim qMinOrd As Double = Convert.ToDouble(row.Cells("QMinORD").Value, System.Globalization.CultureInfo.InvariantCulture)

                    Dim disponibilita As Double = disponibile_codice(codice) * -1 + Magazzino.OttieniDettagliAnagrafica(row.Cells("Codice").Value).Minimo
                    Dim contaDefezioni As Integer = 0

                    ' Controllo se il codice è da acquistare
                    If Magazzino.OttieniDettagliAnagrafica(codice).Approvvigionamento = "B" Then
                        If MessageBox.Show($"Il codice {codice} ({descrizione}) è segnalato come ACQUISTARE. Proseguire con ODP?", "Prosegui", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                            DataGridView2.Rows.Add(codice, descrizione, "ACQUISTARE")
                            contaDefezioni += 1
                        End If
                    End If

                    ' Controllo disponibilità
                    If controllo_disponibile(codice, qTot.ToString()) <> "" Then
                        If MessageBox.Show($"Il codice {codice} ({descrizione}) andrebbe a disponibile >0 con il lancio di questo ODP. Proseguire?", "Prosegui", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                            DataGridView2.Rows.Add(codice, descrizione, "Disp>0")
                            contaDefezioni += 1
                        End If
                    End If

                    ' Se non ci sono defezioni, procedo con il lancio ODP
                    If contaDefezioni = 0 Then
                        If qMinOrd >= disponibilita And qMinOrd >= qTot Then
                            ' Lancio ODP con quantità minima ordinabile
                            If qMinOrd > 0 Then
                                Acquisti.procedura_lancio_odp(row.Cells("Codice").Value, ComboBox5.Text, Acquisti.elenco_fasi(ComboBox4.SelectedIndex), DateTimePicker4.Value, DateTimePicker4.Value, TextBox8.Text, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).Cardname, qMinOrd, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).DocNum, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).Cardcode, ComboBox2.Text)
                            Else
                                MessageBox.Show($"Del codice {codice} {descrizione} la quantità è minore di 1")
                            End If
                        Else
                            ' Gestione lancio ODP con disponibilità
                            If disponibilita > qTot Then
                                If MessageBox.Show($"Del codice {codice} ({descrizione}) andrebbe lanciata una quantità maggiore pari a {disponibilita}. Lanciare?", "Lanciare ODP", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

                                    Acquisti.procedura_lancio_odp(row.Cells("Codice").Value, ComboBox5.Text, Acquisti.elenco_fasi(ComboBox4.SelectedIndex), DateTimePicker4.Value, DateTimePicker4.Value, TextBox8.Text, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).Cardname, disponibilita, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).DocNum, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).Cardcode, ComboBox2.Text)
                                ElseIf qTot > 0 Then
                                    Acquisti.procedura_lancio_odp(row.Cells("Codice").Value, ComboBox5.Text, Acquisti.elenco_fasi(ComboBox4.SelectedIndex), DateTimePicker4.Value, DateTimePicker4.Value, TextBox8.Text, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).Cardname, qTot, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).DocNum, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).Cardcode, ComboBox2.Text)
                                Else
                                    MessageBox.Show($"Del codice {codice} {descrizione} la quantità è minore di 1")
                                End If
                            ElseIf disponibilita < qTot And disponibilita > 0 Then
                                Acquisti.procedura_lancio_odp(row.Cells("Codice").Value, ComboBox5.Text, Acquisti.elenco_fasi(ComboBox4.SelectedIndex), DateTimePicker4.Value, DateTimePicker4.Value, TextBox8.Text, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).Cardname, disponibilita, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).DocNum, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).Cardcode, ComboBox2.Text)
                            ElseIf qTot > 0 Then
                                Acquisti.procedura_lancio_odp(row.Cells("Codice").Value, ComboBox5.Text, Acquisti.elenco_fasi(ComboBox4.SelectedIndex), DateTimePicker4.Value, DateTimePicker4.Value, TextBox8.Text, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).Cardname, qTot, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).DocNum, Acquisti.Cliente_relativo_alla_commessa(TextBox8.Text).Cardcode, ComboBox2.Text)
                            Else
                                MessageBox.Show($"Del codice {codice} {descrizione} la quantità è minore di 1")
                            End If
                        End If
                    End If
                End If
            Next

            MsgBox("ODP lanciato/I con successo")
            UpdateDataGridView(Codice_sap, quantità_Selezionata)

        End If
    End Sub

    Sub cambia_stato_fantasma(par_codice_sap As String)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand

        CMD_SAP_1.Connection = Cnn1
        CMD_SAP_1.CommandText = "
UPDATE T11 SET T11.U_PRG_TIR_EXPLOSION=T10.NUOVO_STATO
FROM
(
SELECT T0.ITEMCODE, CASE
WHEN COALESCE(T0.U_PRG_TIR_EXPLOSION,'N')='Y' THEN 'N' ELSE 'Y' END AS NUOVO_STATO
FROM OITM T0 
WHERE T0.ITEMCODE='" & par_codice_sap & "'
)
AS T10 INNER JOIN OITM T11 ON T10.ITEMCODE=T11.ITEMCODE"

        CMD_SAP_1.ExecuteNonQuery()

        Cnn1.Close()
    End Sub



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        commessa = TextBox1.Text
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Tree_Boom_Click(sender As Object, e As EventArgs) Handles Tree_Boom.Click
        If Tree_Boom.SelectedNode IsNot Nothing Then
            ' Get the selected node text (e.g., "code-nome_figlio-quantità")
            Dim selectedNodeText As String = Tree_Boom.SelectedNode.Text

            ' Extract the code from the selected node text
            Dim selectedCode As String = selectedNodeText.Split("-"c)(0)


            ' Extract the quantità from the selected node text
            Dim selectedQuantita As Decimal
            Try
                ' Split the string by "Q : =" and take the part after it
                Dim parts As String() = selectedNodeText.Split(New String() {"Q : ="}, StringSplitOptions.None)

                ' Check if there is a part after "Q : ="
                If parts.Length > 1 Then
                    ' Trim any leading/trailing whitespace from the extracted part
                    Dim quantitaString As String = parts(1).Trim()

                    ' Try to convert the extracted part to a decimal
                    If Decimal.TryParse(quantitaString, selectedQuantita) = False Then
                        ' If conversion fails, set default value
                        selectedQuantita = 1
                    End If
                Else
                    ' If "Q : =" not found or no part after it, set default value
                    selectedQuantita = 1
                End If
            Catch ex As Exception
                ' Handle any unexpected exceptions by setting default value
                selectedQuantita = 1
            End Try



            ' Update the DataGridView based on the selected code and quantità
            Codice_sap = selectedCode
            quantità_Selezionata = selectedQuantita
            UpdateDataGridView(selectedCode, selectedQuantita)


        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disp").Value < 0 Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disp").Style.ForeColor = Color.Red
        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disp").Value > 0 Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disp").Style.ForeColor = Color.Green
        Else

            DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disp").Value = Nothing
        End If

        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="PDF").Value = "NO" Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="PDF").Style.ForeColor = Color.Red
        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="PDF").Value = "SI" Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="PDF").Style.ForeColor = Color.Green
        Else

            DataGridView1.Rows(e.RowIndex).Cells(columnName:="PDF").Style.ForeColor = Color.Yellow
        End If

        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Fant").Value = "N" Then
            DataGridView1.Rows(e.RowIndex).Cells(columnName:="Fant").Style.BackColor = Color.Orange
        End If
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

    Private Sub DataGridView_ODP_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        ' Controlla se il tasto Shift è stato premuto
        isShiftKeyDown = (e.KeyCode = Keys.ShiftKey)
    End Sub

    Private Sub DataGridView_ODP_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyUp
        ' Controlla se il tasto Shift è stato rilasciato
        isShiftKeyDown = (e.KeyCode <> Keys.ShiftKey)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click


        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells("Codice").Value.ToString().StartsWith("C", StringComparison.OrdinalIgnoreCase) Then
                ' Fai qualcosa se la prima lettera è C
            Else
                Dim percorso As String = Homepage.percorso_disegni_generico & "PDF\" & row.Cells("disegno").Value & ".PDF"

                If File.Exists(percorso) Then

                    row.Cells("PDF").Value = "SI"

                Else

                    row.Cells("PDF").Value = "NO"
                End If
            End If

        Next
        MsgBox("FINE")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        For Each row As DataGridViewRow In DataGridView1.Rows
            If row.Cells("Codice").Value.ToString().StartsWith("C", StringComparison.OrdinalIgnoreCase) Then
                ' Fai qualcosa se la prima lettera è C
            Else
                Dim percorso As String = Homepage.percorso_DXF & row.Cells("disegno").Value & ".DXF"
                If File.Exists(percorso) Then

                    row.Cells("DXF").Value = "SI"

                Else

                    row.Cells("DXF").Value = "NO"
                End If
            End If


        Next
        MsgBox("FINE")
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        ' Crea una cartella nel desktop
        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim targetFolderPath As String = Path.Combine(desktopPath, "Files_Simili")

        ' Crea la cartella se non esiste
        If Not Directory.Exists(targetFolderPath) Then
            Directory.CreateDirectory(targetFolderPath)

        End If

        If Not Directory.Exists(targetFolderPath & "\Nuovi") Then
            Directory.CreateDirectory(targetFolderPath & "\Nuovi")
        End If

        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Controlla se la colonna "disegno" è valida
            If row.Cells("PDF").Value = "NO" And row.Cells("Disegno").Value <> "" Then
                ' Costruisci il percorso del file con le prime 10 lettere del nome del disegno
                Dim partialFileName As String = row.Cells("disegno").Value.ToString().Substring(0, Math.Min(10, row.Cells("disegno").Value.ToString().Length))

                Dim contatore As Integer = 0
                Do While contatore <= 2
                    Dim filePath As String
                    Dim sourceFilePath As String
                    If contatore = 0 Then
                        filePath = Homepage.percorso_disegni_generico & "PDF\" & partialFileName & ".PDF"
                        sourceFilePath = Homepage.percorso_disegni_generico & "PDF\" & partialFileName & ".PDF"
                    ElseIf contatore = 1 Then
                        filePath = Homepage.percorso_disegni_generico & "PDF\" & partialFileName & "-Sheet1.PDF"
                        sourceFilePath = Homepage.percorso_disegni_generico & "PDF\" & partialFileName & "-Sheet1.PDF"
                    ElseIf contatore = 2 Then
                        filePath = Homepage.percorso_disegni_generico & "PDF\" & partialFileName & "-Sheet2.PDF"
                        sourceFilePath = Homepage.percorso_disegni_generico & "PDF\" & partialFileName & "-Sheet2.PDF"
                    End If


                    ' Verifica se il file con le prime 10 lettere del nome del disegno esiste
                    If File.Exists(filePath) Then
                        ' Il file esiste, imposta il valore "SI" nella colonna "DXF"
                        row.Cells("PDF").Value = "SIMILE"



                        ' Costruisci il percorso di destinazione nella cartella creata
                        Dim targetFilePath As String = Path.Combine(targetFolderPath, Path.GetFileName(sourceFilePath))

                        ' Sposta il file nella cartella di destinazione
                        If File.Exists(targetFilePath) Then
                        Else
                            File.Copy(sourceFilePath, targetFilePath)
                        End If



                        If File.Exists(targetFolderPath & "\Nuovi\" & row.Cells("disegno").Value & ".PDF") Then
                        Else
                            File.Copy(sourceFilePath, targetFolderPath & "\Nuovi\" & row.Cells("disegno").Value & ".PDF")
                        End If



                    End If
                    contatore = contatore + 1
                Loop

            End If
        Next

        MsgBox("FINE")
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
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

            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(Disegno) Then
                Magazzino.visualizza_disegno(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Disegno").Value)
            End If
        End If



    End Sub



    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        Dim percorso As String
        Dim ultima_revisione As String
        Dim sottostringaSinistra As String


        For Each row As DataGridViewRow In DataGridView1.Rows
            sottostringaSinistra = ""
            ultima_revisione = ""
            If row.Cells("Codice").Value.ToString().StartsWith("C", StringComparison.OrdinalIgnoreCase) Then
                ' Fai qualcosa se la prima lettera è C
            Else
                Dim contatore As Integer = 0

                If row.Cells("codice_brb").Value <> "" Then
                    Do While contatore <= 9



                        percorso = Homepage.percorso_disegni_generico & "PDF\" & row.Cells("codice_brb").Value & "#" & contatore & ".PDF"

                        If File.Exists(percorso) Then
                            ' Trova la posizione della parola "orso" nel percorso
                            Dim posizionepdf As Integer = InStr(percorso, ".PDF")



                            sottostringaSinistra = percorso.Substring(0, posizionepdf - 1)



                            ultima_revisione = row.Cells("codice_brb").Value & "#" & contatore
                        End If


                        contatore = contatore + 1
                    Loop
                    If Magazzino.OttieniDettagliAnagrafica(row.Cells("Codice").Value).Disegno <> ultima_revisione And ultima_revisione <> "" Then



                        If MessageBox.Show($"Disegno = " & Magazzino.OttieniDettagliAnagrafica(row.Cells("Codice").Value).Disegno & vbCrLf & "Ultima rev = " & ultima_revisione, "Sostituisci", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

                            Magazzino.cambiare_disegno(row.Cells("Codice").Value, ultima_revisione, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)

                            Magazzino.update_AITM(row.Cells("Codice").Value, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)

                        End If




                    End If
                End If
            End If

        Next
        MsgBox("FINE")

    End Sub



    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        For Each row As DataGridViewRow In DataGridView1.Rows ' Sostituisci DataGridView1 con il nome effettivo del tuo controllo DataGridView
            ' Verifica se la cella nella colonna "Seleziona" è True
            If Convert.ToBoolean(row.Cells("Seleziona").Value) = True Then
                cambia_stato_fantasma(row.Cells("Codice").Value)
            End If
        Next
        MsgBox("valori Fantasma cambiati con successo")
        UpdateDataGridView(Codice_sap, quantità_Selezionata)
    End Sub

    Private Sub Cmd_Cerca_Click(sender As Object, e As EventArgs) Handles Cmd_Cerca.Click
        inizializza_albero(TextBox1.Text)
    End Sub

    Private Sub Tree_Boom_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles Tree_Boom.AfterSelect

    End Sub
End Class
