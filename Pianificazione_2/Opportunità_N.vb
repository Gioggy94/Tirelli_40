Imports System.IO

Imports System.Data.SqlClient

Imports System.Windows.Controls
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Imports System.Drawing

Imports System.Threading.Tasks
Imports System.Reflection.Emit




Public Class Opportunità_N
    Public n_opportunità As Integer
    Private codice_OWNER(1000) As String
    Private codice_type_ARRAY(100) As String
    Private codice_addetto_ARRAY(1000) As String
    Private ultima_riga_opportunità As Integer
    Private codice_documento_array(100) As String
    Private sottocartella As String

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Sub inizializza_opportunità(par_n_opportunità As Integer)
        n_opportunità = par_n_opportunità
        Testata_opportunità(par_n_opportunità)
        livello_opportunità(par_n_opportunità)
        trova_percorso(par_n_opportunità)
        mostra_file_async(Replace(LinkLabel1.Text, "\\tirfs01\TIRELLI", "T:"), TreeView1)

        cartelle_opportunità()
    End Sub

    Public Sub mostra_file_async(par_percorso As String, par_treeview As System.Windows.Forms.TreeView)
        Dim rootDirectoryPath As String = par_percorso

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

    Public Sub AddDirectories(parentNode As TreeNode, par_treeview As System.Windows.Forms.TreeView)
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




    ' Funzione per aggiungere l'icona della cartella in modo sicuro
    Private Sub AggiungiIconaCartella()
        If Not ImageList1.Images.ContainsKey("folder") Then
            ImageList1.Images.Add("folder", SystemIcons.WinLogo)
        End If
    End Sub

    Public Sub Addfiles(parentNode As TreeNode, par_treeview As System.Windows.Forms.TreeView)
        Dim parentDirectory As DirectoryInfo = TryCast(parentNode.Tag, DirectoryInfo)
        If parentDirectory Is Nothing OrElse Not parentDirectory.Exists Then Exit Sub

        Try
            For Each file As FileInfo In parentDirectory.GetFiles()
                Dim fileNode As New TreeNode(file.Name & " (" & file.LastWriteTime.ToString() & ")") With {.Tag = file}

                ' Sostituzione del percorso di rete
                Dim filepath As String = file.FullName.Replace("\\tirfs01\Tirelli", "T:")

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





    Sub Testata_opportunità(par_N_opportunità As Integer)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "
SELECT T0.[OpprId], T0.[CardCode],t0.cardname, case when T1.NAME is null then '' else t1.name end as 'name' , T2.[lastName]+' ' +T2.[firstName] as 'Compilatore', T3.[slpname] as 'Venditore', t0.status, t0.opendate
, case when t0.u_descrizioneprogetto is null then '' else t0.u_descrizioneprogetto end as 'Descrizione_progetto'
,cast (t0.maxsumloc as decimal) as 'maxsumloc', t0.preddate
, CASE WHEN T0.U_CLIENTEFINALE IS NULL THEN '' ELSE T0.U_CLIENTEFINALE END AS 'Cliente_finale'
FROM OOPR T0 
left join ocpr t1 on t1.cardcode = t0.cardcode AND T0.CPRCODE=T1.CNTCTCODE 
left join [TIRELLI_40].[DBO].OHEM T2 on t2.empid=t0.owneR
left join oslp T3 on t3.slpcode=t0.SLPCODE


WHERE T0.[OpprId] =" & par_N_opportunità & ""

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() Then

            TextBox6.Text = cmd_SAP_reader("OpprId")
            TextBox1.Text = cmd_SAP_reader("CARDCODE")
            TextBox2.Text = cmd_SAP_reader("CARDNAME")
            TextBox3.Text = cmd_SAP_reader("name")
            ComboBox1.Text = cmd_SAP_reader("venditore")
            DateTimePicker1.Value = cmd_SAP_reader("opendate")
            TextBox8.Text = cmd_SAP_reader("status")
            RichTextBox1.Text = cmd_SAP_reader("Descrizione_progetto")
            TextBox4.Text = String.Format("{0:C}", Convert.ToDecimal(cmd_SAP_reader("maxsumloc")))
            DateTimePicker2.Value = cmd_SAP_reader("preddate")
            TextBox7.Text = cmd_SAP_reader("Cliente_finale")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub

    Private Sub Cmd_Indietro_Click(sender As Object, e As EventArgs) Handles Cmd_Indietro.Click
        n_opportunità = Int(n_opportunità) - 1
        inizializza_opportunità(n_opportunità)
    End Sub

    Private Sub Cmd_Avanti_Click(sender As Object, e As EventArgs) Handles Cmd_Avanti.Click
        n_opportunità = Int(n_opportunità) + 1
        inizializza_opportunità(n_opportunità)
    End Sub

    Private Sub Cmd_Cerca_Click(sender As Object, e As EventArgs) Handles Cmd_Cerca.Click
        n_opportunità = TextBox6.Text
        inizializza_opportunità(n_opportunità)
    End Sub

    Sub livello_opportunità(par_N_opportunità As Integer)

        DataGridView1.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "
SELECT T0.[Line], T0.[OpenDate], T0.[CloseDate], T0.[SlpCode],t1.slpname, T0.[Step_Id],t4.Descript as 'Fase',T0.[DocChkbox], T0.[ObjType], T0.[DocNumber], T0.[Owner],T2.[lastName]+' ' +T2.[firstName] as 'Compilatore', T0.[U_PRG_AZS_NoteLivOPP], T0.[U_Informazioni], T0.[U_Priorita], T0.[U_Layout] ,t0.status

FROM  OPR1 T0 
left join oslp T1 on t1.slpcode=t0.SLPCODE
left join oost t4 on t0.step_id=t4.num
left join [TIRELLI_40].[DBO].OHEM T2 on t2.empid=t0.owneR

WHERE T0.[OpprId] =" & par_N_opportunità & "
order by t0.line
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim nome_documento As String = ""

        Do While cmd_SAP_reader.Read()

            If cmd_SAP_reader("ObjType") = "-1" Then
                nome_documento = ""

            ElseIf cmd_SAP_reader("ObjType") = "-23" Then
                nome_documento = "Offerta"
            ElseIf cmd_SAP_reader("ObjType") = "22" Then
                nome_documento = "Ordine di acquisto"
            ElseIf cmd_SAP_reader("ObjType") = "13" Then
                nome_documento = ""
            ElseIf cmd_SAP_reader("ObjType") = "Fattura di vendita" Then
                nome_documento = ""
            ElseIf cmd_SAP_reader("ObjType") = "17" Then
                nome_documento = "Ordine cliente"

            Else nome_documento = ""

            End If





            DataGridView1.Rows.Add(
            cmd_SAP_reader("Line"),
            cmd_SAP_reader("OpenDate"),
            cmd_SAP_reader("CloseDate"),
            cmd_SAP_reader("SlpCode"),
            cmd_SAP_reader("slpname"),
            cmd_SAP_reader("Step_Id"),
            cmd_SAP_reader("Fase"),
            cmd_SAP_reader("DocChkbox"),
            cmd_SAP_reader("ObjType"),
            nome_documento,
            cmd_SAP_reader("DocNumber"),
            cmd_SAP_reader("Owner"),
            cmd_SAP_reader("Compilatore"),
            cmd_SAP_reader("U_PRG_AZS_NoteLivOPP"),
            cmd_SAP_reader("U_Informazioni"),
            cmd_SAP_reader("U_Priorita"),
            cmd_SAP_reader("U_Layout"),
            cmd_SAP_reader("status")
        )

            '      DataGridView1.Rows.Add(
            '    cmd_SAP_reader("Line"),
            '    cmd_SAP_reader("OpenDate"),
            '    cmd_SAP_reader("CloseDate"),
            '    cmd_SAP_reader("SlpCode"),
            '    cmd_SAP_reader("slpname"),
            '    cmd_SAP_reader("Step_Id"),
            '    "",
            '    cmd_SAP_reader("DocChkbox"),
            '    cmd_SAP_reader("ObjType"),
            '    nome_documento,
            '    cmd_SAP_reader("DocNumber"),
            '    cmd_SAP_reader("Owner"),
            '    cmd_SAP_reader("Compilatore"),
            '    cmd_SAP_reader("U_PRG_AZS_NoteLivOPP"),
            '    cmd_SAP_reader("U_Informazioni"),
            '    cmd_SAP_reader("U_Priorita"),
            '    cmd_SAP_reader("U_Layout"),
            '    cmd_SAP_reader("status")
            ')

        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub

    Sub inserisci_combobox_addetto_vendite()


        Addetto_vendite.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "select t0.slpcode,t0.SlpName
from oslp t0
where t0.active='Y'
order by t0.SlpName
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            codice_addetto_ARRAY(Indice) = cmd_SAP_reader("slpcode")
            Addetto_vendite.Items.Add(cmd_SAP_reader("SlpName"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()



    End Sub

    Sub inserisci_combobox_type()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.stepid, t0.num, T0.Descript, T0.Canceled, T0.SalesStage, T0.PurStage 
FROM OOST T0 "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Dim Indice As Integer
        Indice = 0


        Do While cmd_SAP_reader_2.Read()


            ' If cmd_SAP_reader_2("num") = 2 Or cmd_SAP_reader_2("num") = 3 Or cmd_SAP_reader_2("num") = 5 Or cmd_SAP_reader_2("num") = 8 Or cmd_SAP_reader_2("num") = 14 Then
            Type.Items.Add(cmd_SAP_reader_2("Descript"))
            codice_type_ARRAY(Indice) = cmd_SAP_reader_2("num")

            'ElseIf cmd_SAP_reader_2("num") = 7 Or cmd_SAP_reader_2("num") = 9 Or cmd_SAP_reader_2("num") = 12 Or cmd_SAP_reader_2("num") = 15 Then
            '    Type.Items.Add(cmd_SAP_reader_2("Descript"))
            '    codice_type_ARRAY(Indice) = cmd_SAP_reader_2("num")

            'End If

            Indice = Indice + 1
        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Sub inserisci_combobox_owner()
        Titolare.Items.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.empid, T0.[lastName]+' ' +T0.[firstName] as 'Compilatore' 
FROM [TIRELLI_40].[DBO].OHEM T0 "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Dim Indice As Integer
        Indice = 0


        Do While cmd_SAP_reader_2.Read()

            Titolare.Items.Add(cmd_SAP_reader_2("compilatore"))
            codice_OWNER(Indice) = cmd_SAP_reader_2("empid")
            Indice = Indice + 1
        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Sub inserisci_informazioni()
        Info.Items.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT code, name
from [@informazioni] "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader_2.Read()

            Info.Items.Add(cmd_SAP_reader_2("code"))

        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Sub inserisci_priorità()
        Priorità.Items.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT *
from [@priorita] "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Priorità.Items.Add("")


        Do While cmd_SAP_reader_2.Read()

            Priorità.Items.Add(cmd_SAP_reader_2("code"))

        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Sub inserisci_layout()
        Layout.Items.Clear()
        Layout.Items.Add("")
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT *
from [@layout] "


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader_2.Read()

            Layout.Items.Add(cmd_SAP_reader_2("code"))

        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Sub inserisci_combobox_tipo_documento()
        Dim Indice As Integer
        Indice = 0
        Dim codice_documento(100) As String

        Tipo_di_documento.Items.Add("")
        codice_documento_array(Indice) = "-1"
        Indice = Indice + 1

        Tipo_di_documento.Items.Add("Offerta")
        codice_documento_array(Indice) = "23"

        Indice = Indice + 1

        Tipo_di_documento.Items.Add("Ordine di acquisto")
        codice_documento_array(Indice) = "22"

        Indice = Indice + 1

        Tipo_di_documento.Items.Add("Fattura di vendita")
        codice_documento_array(Indice) = "13"

        Indice = Indice + 1

        Tipo_di_documento.Items.Add("Ordine cliente")
        codice_documento_array(Indice) = "17"

        Indice = Indice + 1

    End Sub

    Private Sub Opportunità_N_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inserisci_combobox_addetto_vendite()
        inserisci_combobox_type()
        inserisci_combobox_tipo_documento()
        inserisci_combobox_owner()
        inserisci_informazioni()
        inserisci_priorità()
        inserisci_layout()
    End Sub

    Sub trova_ultima_riga(par_numero_opportunità As Integer)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT max(t0.line) as 'MAX' FROM OPR1 T0 WHERE T0.[OpprId] = " & par_numero_opportunità & ""


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        If cmd_SAP_reader_2.Read() Then

            ultima_riga_opportunità = cmd_SAP_reader_2("max")



        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Sub cancella_riga_ultima_riga(par_numero_opportunità As Integer, par_ultima_riga_opportunità As Integer)
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = cnn3


        CMD_SAP_3.CommandText = "delete opr1 where line =" & par_ultima_riga_opportunità & " and opr1.opprid=" & par_numero_opportunità & "
update opr1 set status='O' where line =" & par_ultima_riga_opportunità & "-1 and opr1.opprid=" & par_numero_opportunità & "
"

        CMD_SAP_3.ExecuteNonQuery()
        cnn3.Close()


    End Sub

    Private Sub CancellaUltimaRigaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        If MessageBox.Show($"Sei sicuro di voler eliminare l'ultima riga?", "Elimina", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            trova_ultima_riga(n_opportunità)
            cancella_riga_ultima_riga(n_opportunità, ultima_riga_opportunità)
            livello_opportunità(n_opportunità)
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting

    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If MessageBox.Show($"Sei sicuro di voler eliminare l'ultima riga?", "Elimina", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            trova_ultima_riga(n_opportunità)
            If ultima_riga_opportunità = 0 Then
                MsgBox("Non è possibile cancellare l'unica riga")
            Else
                cancella_riga_ultima_riga(n_opportunità, ultima_riga_opportunità)
                Opportunità_aggiungi_riga.allinea_importo_potenziale(n_opportunità)
                livello_opportunità(n_opportunità)
            End If

        End If


    End Sub



    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        ' Controlla se la colonna corrente è la colonna "Titolare"
        If DataGridView1.Columns(e.ColumnIndex).Name = "Titolare" Then

            ' Ottieni il valore modificato nella cella corrente
            Dim comboBoxCell As DataGridViewComboBoxCell = CType(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex), DataGridViewComboBoxCell)
            Dim selectedValue As String = comboBoxCell.EditedFormattedValue.ToString()

            ' Ottieni l'indice della stringa selezionata nella ComboBox
            Dim selectedIndex As Integer = comboBoxCell.Items.IndexOf(selectedValue)

            ' Verifica se l'indice è valido e ottieni il codice_OWNER corrispondente
            If selectedIndex >= 0 Then
                DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice_titolare").Value = codice_OWNER(selectedIndex)
            Else
                ' L'indice non è valido, esegui qui la logica desiderata
            End If

        ElseIf DataGridView1.Columns(e.ColumnIndex).Name = "Type" Then

            ' Ottieni il valore modificato nella cella corrente
            Dim comboBoxCell As DataGridViewComboBoxCell = CType(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex), DataGridViewComboBoxCell)
            Dim selectedValue As String = comboBoxCell.EditedFormattedValue.ToString()

            ' Ottieni l'indice della stringa selezionata nella ComboBox
            Dim selectedIndex As Integer = comboBoxCell.Items.IndexOf(selectedValue)

            ' Verifica se l'indice è valido e ottieni il codice_OWNER corrispondente
            If selectedIndex >= 0 Then
                DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice_type").Value = codice_type_ARRAY(selectedIndex)
            Else
                ' L'indice non è valido, esegui qui la logica desiderata
            End If

        ElseIf DataGridView1.Columns(e.ColumnIndex).Name = "Addetto_vendite" Then

            ' Ottieni il valore modificato nella cella corrente
            Dim comboBoxCell As DataGridViewComboBoxCell = CType(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex), DataGridViewComboBoxCell)
            Dim selectedValue As String = comboBoxCell.EditedFormattedValue.ToString()

            ' Ottieni l'indice della stringa selezionata nella ComboBox
            Dim selectedIndex As Integer = comboBoxCell.Items.IndexOf(selectedValue)

            ' Verifica se l'indice è valido e ottieni il codice_OWNER corrispondente
            If selectedIndex >= 0 Then
                DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice_addetto").Value = codice_addetto_ARRAY(selectedIndex)
            Else
                ' L'indice non è valido, esegui qui la logica desiderata
            End If

        ElseIf DataGridView1.Columns(e.ColumnIndex).Name = "Tipo_di_documento" Then

            ' Ottieni il valore modificato nella cella corrente
            Dim comboBoxCell As DataGridViewComboBoxCell = CType(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex), DataGridViewComboBoxCell)
            Dim selectedValue As String = comboBoxCell.EditedFormattedValue.ToString()

            ' Ottieni l'indice della stringa selezionata nella ComboBox
            Dim selectedIndex As Integer = comboBoxCell.Items.IndexOf(selectedValue)

            ' Verifica se l'indice è valido e ottieni il codice_OWNER corrispondente
            If selectedIndex >= 0 Then
                DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice_documento").Value = codice_documento_array(selectedIndex)
            Else
                ' L'indice non è valido, esegui qui la logica desiderata
            End If
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Opportunità_aggiungi_riga.Show()
        Opportunità_aggiungi_riga.n_opportunità = n_opportunità
    End Sub

    Sub trova_percorso(par_n_opportunità As Integer)
        Dim rootDirectory As String = Homepage.PERCORSO_CARTELLE_OPPORTUNITà
        Dim labelNumber As Integer = par_n_opportunità

        ' Divide il valore del numero della label in X e Y
        Dim x As Integer = labelNumber \ 1000 ' divisione intera per ottenere X
        Dim y As Integer = labelNumber Mod 1000 ' modulo per ottenere Y

        ' Crea il pattern di ricerca per la directory
        ' Dim searchPattern As String = String.Format("Opp. {0}-*", x)

        Dim searchPattern As String = String.Format("Opp. {0}*", x)

        ' Ottiene un elenco di tutte le directory che corrispondono al pattern di ricerca
        Dim directories As String() = System.IO.Directory.GetDirectories(rootDirectory, searchPattern)

        ' Cerca la directory corretta che soddisfa il criterio 
        Dim targetDirectory As String = Nothing
        For Each directory In directories
            ' Estrae il numero Y dal nome della directory
            Dim directoryName As String = New System.IO.DirectoryInfo(directory).Name
            Dim directoryNumber As Integer
            If Integer.TryParse(directoryName.Split("-"c)(1), directoryNumber) AndAlso directoryNumber > y Then
                targetDirectory = directory
                Exit For
            End If
        Next

        ' Se la directory corretta è stata trovata, fa qualcosa con essa
        If Not String.IsNullOrEmpty(targetDirectory) Then
            ' Esempio: visualizza il percorso della directory corretta
            sottocartella = targetDirectory
            Console.WriteLine("Directory trovata: " & targetDirectory)
        Else
            ' La directory corretta non è stata trovata
            Console.WriteLine("Nessuna directory corrispondente trovata")
        End If


        searchPattern = String.Format("Opp. {0}*", par_n_opportunità)
        directories = System.IO.Directory.GetDirectories(targetDirectory, searchPattern)

        For Each directory In directories
            ' Estrae il numero Y dal nome della directory
            Dim directoryName As String = Nothing
            directoryName = New System.IO.DirectoryInfo(directory).Name

            If directoryName <> Nothing Then
                targetDirectory = directory
                Exit For
            End If
        Next

        If Not String.IsNullOrEmpty(targetDirectory) Then

            LinkLabel1.Text = targetDirectory
        Else
            ' La directory corretta non è stata trovata
            Console.WriteLine("Nessuna directory corrispondente trovata")
        End If


    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start(LinkLabel1.Text)
    End Sub





    ' Funzione per aggiungere l'icona di un file in modo sicuro
    Private Sub AggiungiIconaFile(extension As String, icon As Icon)
        If Not ImageList1.Images.ContainsKey(extension) Then
            ImageList1.Images.Add(extension, icon)
        End If
    End Sub



    Public Sub AddDirectories(parentNode As TreeNode)
        Dim parentDirectory As DirectoryInfo = TryCast(parentNode.Tag, DirectoryInfo)
        If parentDirectory Is Nothing OrElse Not parentDirectory.Exists Then
            Debug.WriteLine("❌ ERRORE: Directory non valida -> " & If(parentDirectory?.FullName, "NULL"))
            Exit Sub
        End If

        Debug.WriteLine("📂 Scansiono cartella: " & parentDirectory.FullName)

        ' Aggiunge tutte le cartelle come nodi figli
        For Each directory As DirectoryInfo In parentDirectory.GetDirectories()
            Dim directoryNode As New TreeNode(directory.Name)
            directoryNode.Tag = directory

            ' Aggiunge l'icona della cartella se non esiste già
            If Not ImageList1.Images.ContainsKey("folder") Then
                ImageList1.Images.Add("folder", SystemIcons.WinLogo)
            End If
            directoryNode.ImageKey = "folder"

            parentNode.Nodes.Add(directoryNode)

            ' Aggiunge i file della cartella
            Addfiles(directoryNode)

            ' Chiamata ricorsiva per le sottocartelle
            AddDirectories(directoryNode)
        Next
    End Sub

    Public Sub Addfiles(parentNode As TreeNode)
        Dim parentDirectory As DirectoryInfo = TryCast(parentNode.Tag, DirectoryInfo)
        If parentDirectory Is Nothing OrElse Not parentDirectory.Exists Then
            Debug.WriteLine("❌ ERRORE: Directory non valida in Addfiles -> " & If(parentDirectory?.FullName, "NULL"))
            Exit Sub
        End If

        Try
            For Each file As FileInfo In parentDirectory.GetFiles()
                Dim fileNode As New TreeNode(file.Name & " (" & file.LastWriteTime.ToString() & ")")
                fileNode.Tag = file

                ' Sostituzione del percorso di rete
                Dim filepath As String = file.FullName
                filepath = Replace(filepath, "\\tirfs01\Tirelli", "T:")

                ' Ottiene l'icona del file
                Dim fileIcon As Icon = SystemIcons.WinLogo
                Try
                    fileIcon = Icon.ExtractAssociatedIcon(filepath)
                Catch ex As Exception
                    Debug.WriteLine("⚠️ ERRORE: Impossibile estrarre icona per " & filepath)
                End Try

                ' Aggiunge l'icona alla ImageList
                If Not ImageList1.Images.ContainsKey(file.Extension) Then
                    ImageList1.Images.Add(file.Extension, fileIcon)
                End If

                fileNode.ImageKey = file.Extension
                parentNode.Nodes.Add(fileNode)
            Next
        Catch ex As Exception
            Debug.WriteLine("❌ ERRORE in Addfiles: " & ex.Message)
        End Try
    End Sub

    Private Sub Elimina_file_Click(sender As Object, e As EventArgs) Handles Elimina_file.Click
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

    Private Sub RinominaFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RinominaFileToolStripMenuItem.Click

        ' Controlla se un nodo è stato selezionato
        If TreeView1.SelectedNode IsNot Nothing Then
            Dim file As FileInfo = TryCast(TreeView1.SelectedNode.Tag, FileInfo)

            ' Apri una finestra di dialogo per consentire all'utente di inserire il nuovo nome del file
            Dim newFileName As String = InputBox("Inserisci il nuovo nome del file", "Rinomina file", file.Name)

            If Not String.IsNullOrEmpty(newFileName) Then
                ' Rinomina il file
                Dim newFilePath As String = IO.Path.Combine(file.DirectoryName, newFileName)
                FileSystem.Rename(file.FullName, newFilePath)
                'mostra_file()
                mostra_file_async(Replace(LinkLabel1.Text, "\\tirfs01\TIRELLI", "T:"), TreeView1)


            End If
        End If
    End Sub

    Private Sub TreeView1_DragDrop(sender As Object, e As DragEventArgs) Handles TreeView1.DragDrop
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
            File.Copy(filePath, IO.Path.Combine(targetDirectory.FullName, IO.Path.GetFileName(filePath)), True)

            ' Aggiorna la TreeView
            selectedNode.Nodes.Clear()
            AddDirectories(selectedNode)
            Addfiles(selectedNode)

            ' Espande il nodo selezionato
            selectedNode.Expand()
        End If
        'mostra_file()
        mostra_file_async(Replace(LinkLabel1.Text, "\\tirfs01\TIRELLI", "T:"), TreeView1)
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

    Private Sub TreeView1_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeView1.NodeMouseDoubleClick
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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'mostra_file()
        mostra_file_async(Replace(LinkLabel1.Text, "\\tirfs01\TIRELLI", "T:"), TreeView1)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click

        Dim cliente As String
        If TextBox7.Text = "" Then
            cliente = TextBox2.Text
        Else
            cliente = TextBox7.Text
        End If


        If LinkLabel1.Text = "" Or LinkLabel1.Text = Nothing Then

            Directory.CreateDirectory(sottocartella & "\" & "Opp. " & n_opportunità & " " & cliente)
            trova_percorso(n_opportunità)
            'mostra_file()
            mostra_file_async(Replace(LinkLabel1.Text, "\\tirfs01\TIRELLI", "T:"), TreeView1)

        End If
    End Sub




    Sub cartelle_opportunità()

        DataGridView2.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "select *
                            from [Tirelli_40].[dbo].[Requisiti_progetto]
                            where active='Y' and documento='OPP' and id   NOT Like '%%-%%' 
                            order by ID"

        cmd_SAP_reader = CMD_SAP.ExecuteReader



        Do While cmd_SAP_reader.Read()
            DataGridView2.Rows.Add(cmd_SAP_reader("ID") & " " & cmd_SAP_reader("Nome requisito"))


        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick




        If LinkLabel1.Text = "" Or LinkLabel1.Text = Nothing Then

        Else


            If e.RowIndex >= 0 Then



                If e.ColumnIndex = DataGridView2.Columns.IndexOf(Cartella) Then

                    If MessageBox.Show($"Sei sicuro di voler creare la cartella " & DataGridView2.Rows(e.RowIndex).Cells(columnName:="Cartella").Value & " ?", "Crea cartella", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

                        Directory.CreateDirectory(LinkLabel1.Text & "\" & DataGridView2.Rows(e.RowIndex).Cells(columnName:="Cartella").Value)

                        Dim cartella_padre As String = LinkLabel1.Text & "\" & DataGridView2.Rows(e.RowIndex).Cells(columnName:="Cartella").Value
                        Dim cellValue As String = DataGridView2.Rows(e.RowIndex).Cells(columnName:="Cartella").Value.ToString()
                        Dim hyphenIndex As Integer = cellValue.IndexOf(" "c)

                        If hyphenIndex >= 0 Then
                            Dim firstDigits As String = cellValue.Substring(0, hyphenIndex)
                            crea_sottocartelle_opportunità(firstDigits, cartella_padre)
                        Else
                            Console.WriteLine("Nessun trattino trovato nella cella.")
                        End If

                        'mostra_file()
                        mostra_file_async(Replace(LinkLabel1.Text, "\\tirfs01\TIRELLI", "T:"), TreeView1)
                    End If

                End If


            End If



        End If
    End Sub

    Sub crea_sottocartelle_opportunità(par_prime_cifre As String, par_cartella_padre As String)

        Dim Cnn As New SqlConnection

        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "select *
                            from [Tirelli_40].[dbo].[Requisiti_progetto]
                            where active='Y' and documento='OPP' and id Like '%%-%%' and id Like '%%" & par_prime_cifre & "%%' 
                            order by ID"

        cmd_SAP_reader = CMD_SAP.ExecuteReader



        Do While cmd_SAP_reader.Read()

            Directory.CreateDirectory(par_cartella_padre & "\" & cmd_SAP_reader("ID") & " " & cmd_SAP_reader("Nome requisito"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim columnName As String = DataGridView1.Columns(e.ColumnIndex).Name

            If columnName = "Status" AndAlso DataGridView1.Rows(e.RowIndex).Cells(columnName).Value.ToString() <> "O" Then
                e.CellStyle.BackColor = Color.Gray
            ElseIf columnName = "codice_type" AndAlso DataGridView1.Rows(e.RowIndex).Cells(columnName).Value IsNot Nothing Then
                Dim codiceTypeValue As String = DataGridView1.Rows(e.RowIndex).Cells(columnName).Value.ToString()

                Select Case codiceTypeValue
                    Case "3"
                        DataGridView1.Rows(e.RowIndex).Cells(columnName:="type").Style.BackColor = Color.Violet

                        Type.DefaultCellStyle.BackColor = Color.Violet
                    Case "2"
                        DataGridView1.Rows(e.RowIndex).Cells(columnName:="type").Style.BackColor = Color.Gray
                        Type.DefaultCellStyle.BackColor = Color.Gray
                    Case "5"
                        DataGridView1.Rows(e.RowIndex).Cells(columnName:="type").Style.BackColor = Color.Blue
                        Type.DefaultCellStyle.BackColor = Color.Blue
                    Case "8"
                        DataGridView1.Rows(e.RowIndex).Cells(columnName:="type").Style.BackColor = Color.Orange
                        Type.DefaultCellStyle.BackColor = Color.Orange
                    Case "14"
                        DataGridView1.Rows(e.RowIndex).Cells(columnName:="type").Style.BackColor = Color.Yellow
                        Type.DefaultCellStyle.BackColor = Color.Yellow

                End Select
            End If
        End If
    End Sub
End Class