Imports System.Data
Imports System.Data.SqlClient

Public Class Form103

    Public itemcode As String
    Public quantity As Integer
    Public commessa As String
    Public disegno As String
    Public Itemname As String
    Public Docduedate As String
    Public Imputabilita_modifica As String
    Public DOCNUM As Integer
    Public elenco_fasi(100) As String
    Public codice_fase_inserimento As String
    Public itemcode_riga As String
    Public itemname_riga As String
    Public riga As Integer


    Private Sub Button_genera_Click(sender As Object, e As EventArgs) Handles Button_genera.Click

        If TextBox1.Text = Nothing Then
            MsgBox("Selezionare un codice valido")
        Else
            If TextBox_commessa.Text = Nothing Then
                MsgBox("Selezionare una commessa")
            Else

                If TextBox_quantity.Text = Nothing Then
                    MsgBox("Selezionare una quantità")
                Else

                    If ComboBox1.SelectedIndex < 0 Then
                        MsgBox("Selezionare una fase")
                    Else
                        If ComboBox_imputabilita_modifica.SelectedIndex < 0 Then
                            MsgBox("Selezionare un'imputabilità modifica")
                        Else

                            If Docduedate = Nothing Then
                                MsgBox("Selezionare una data di consegna")

                            Else



                                itemcode = TextBox_itemcode.Text
                                    quantity = TextBox_quantity.Text
                                    commessa = TextBox_commessa.Text
                                    Imputabilita_modifica = ComboBox_imputabilita_modifica.Text
                                    creazione_ordine_di_modifica()
                                    TextBox_itemcode.Text = Nothing
                                    TextBox_commessa.Text = Nothing
                                    TextBox_quantity.Text = Nothing
                                    ComboBox_imputabilita_modifica.Text = Nothing

                                Dashboard_MU_New.Show()
                                Button_genera.Enabled = False



                            End If
                        End If

                    End If
                End If

            End If
            End If

    End Sub

    Sub creazione_ordine_di_modifica()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand
        'Dim Cmd_SAP_Reader As SqlDataReader

        Dim CMD_SAP_max_LINE As New SqlCommand
        Dim cmd_SAP_max_LINE_reader As SqlDataReader

        CMD_SAP_max_LINE.Connection = Cnn
        CMD_SAP_max_LINE.CommandText = "SELECT CASE WHEN max(T0.docnum) IS NULL THEN 0 Else max(T0.docnum) end as 'max' FROM [Tirelli_40].[dbo].[OMODIFICA] T0"
        cmd_SAP_max_LINE_reader = CMD_SAP_max_LINE.ExecuteReader

        If cmd_SAP_max_LINE_reader.Read = True Then
            DOCNUM = Val((cmd_SAP_max_LINE_reader("max"))) + 1
        Else
            DOCNUM = 0
        End If
        cmd_SAP_max_LINE_reader.Close()

        'Inserisco i valori nell'odp
        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "insert into [Tirelli_40].[dbo].[OMODIFICA](OMODIFICA.Docnum, OMODIFICA.STATUS,omodifica.itemcode, omodifica.docdate, omodifica.duedate, omodifica.quantity, omodifica.commessa, omodifica.Imputabilita_modifica, OMODIFICA.u_lavorazione,omodifica.u_fase)
            VALUES (" & DOCNUM & ",'P','" & itemcode & "',getdate(),CONVERT(DATETIME, '" & Docduedate & "',103)," & quantity & " ,'" & commessa & "','" & Imputabilita_modifica & "',0,'" & codice_fase_inserimento & "')"
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

        For Each Riga As DataGridViewRow In DataGridView1.Rows

            If DataGridView1(0, Riga.Index).Value <> Nothing Then

                Dim CNN1 As New SqlConnection
                CNN1.ConnectionString = Homepage.sap_tirelli
                CNN1.Open()
                Dim CMD_SAP_1 As New SqlCommand
                CMD_SAP_1.Connection = CNN1

                CMD_SAP_1.CommandText = "insert into [Tirelli_40].[dbo].[MODIFICA1] (docnum,U_STATO, itemcode, LINENUM,visorder, additqty, plannedqty) VALUES (" & DOCNUM & ",'O','" & DataGridView1(0, Riga.Index).Value & "'," & Riga.Index & "," & Riga.Index & ",'" & DataGridView1(2, Riga.Index).Value & "','" & DataGridView1(4, Riga.Index).Value & "') "
                CMD_SAP_1.ExecuteNonQuery()
                CNN1.Close()
            End If
        Next



        MsgBox("L'Ordine di modifica " & vbCrLf & vbCrLf & DOCNUM & vbCrLf & vbCrLf & "è stato creato correttamente")
    End Sub



    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Inserimento_fasi()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()

    End Sub

    Sub Inserimento_fasi()
        Dim CNN As New SqlConnection
        ComboBox1.Items.Clear()
        cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.[Code], T0.[Name] FROM [dbo].[@FASE]  T0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            elenco_fasi(Indice) = cmd_SAP_reader("Code")
            ComboBox1.Items.Add(cmd_SAP_reader("Name"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Docduedate = DateTimePicker1.Value
    End Sub

    Private Sub TextBox_itemcode_TextChanged(sender As Object, e As EventArgs) Handles TextBox_itemcode.TextChanged

        itemcode = TextBox_itemcode.Text
        anagrafica_codice()

    End Sub

    Sub anagrafica_codice()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "select t0.itemcode, t0.itemname, case when t1.code is null then '' else t1.code end as 'code'
        from oitm t0 left join oitt t1 on t0.itemcode=t1.code
        where t0.itemcode= '" & itemcode & "'"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then
            TextBox1.Text = cmd_SAP_reader("Itemname")
        Else
            TextBox1.Text = Nothing
        End If
        cmd_SAP_reader.Close()
        cnn.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        codice_fase_inserimento = elenco_fasi(ComboBox1.SelectedIndex)
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If e.RowIndex >= 0 Then


            If e.ColumnIndex = 0 Then



                itemcode_riga = UCase(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Risorsa").Value)
                informazioni_articolo_riga()


            End If
        End If

    End Sub

    Sub informazioni_articolo_riga()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "Select Case When T2.[VisResCode] Is null Then T0.objTYPE Else '290' end as 'objtype',T0.[ItemName], case when T0.[DfltWH] is null then '01' else T0.[DfltWH] end as 'DfltWH', T1.[Price], T0.VALIDFOR FROM OITM T0  INNER JOIN ITM1 T1 ON T0.[ItemCode] = T1.[ItemCode] left join orsc t2 on T2.[VisResCode]=t0.itemcode WHERE T0.[ItemCode] ='" & itemcode_riga & "' AND  T1.[PriceList] =2"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then


            itemname_riga = cmd_SAP_reader("ItemName")


            DataGridView1.Rows(riga).Cells(columnName:="Descrizione").Value = cmd_SAP_reader("ItemName")

            DataGridView1.Rows(riga).Cells(columnName:="Attrezzaggio").Value = 0
            DataGridView1.Rows(riga).Cells(columnName:="Lavorazione").Value = 1
            DataGridView1.Rows(riga).Cells(columnName:="Totale").Value = DataGridView1.Rows(riga).Cells(columnName:="Attrezzaggio").Value + DataGridView1.Rows(riga).Cells(columnName:="Lavorazione").Value
            Button_genera.Enabled = True
        Else
            Button_genera.Enabled = False

        End If
        cmd_SAP_reader.Close()
        cnn.Close()


    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            riga = e.RowIndex
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        MsgBox(DataGridView1.RowCount)
    End Sub
End Class