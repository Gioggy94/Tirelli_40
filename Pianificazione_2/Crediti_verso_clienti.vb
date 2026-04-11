Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports Word = Microsoft.Office.Interop.Word

Public Class Crediti_verso_clienti
    Private filtro_N_fattura As String
    Private filtro_codice_cliente As String
    Private filtro_cliente As String
    Private filtro_department As String
    Private startIndex As Integer = -1

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Sub inizializza_form()
        riempi_datagridview_crediti()
    End Sub

    Sub riempi_datagridview_crediti()

        DataGridView1.Rows.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "
select *
from 
(
SELECT t0.docentry, t0.docnum as 'Fattura',  T0.[DocDate] as 'Data Fatt', t0.docduedate as 'Data Scad',DATEDIFF(DAY, t0.docduedate, GETDATE()) as 'Overdue', T0.[CardCode] as 'BP code', T0.[CardName] as 'BP Name',t11.cardcode as 'Final_BP_code', t11.cardname as 'Final BP', T2.[Indicator] as 'Year',  t0.U_uffcompetenza as 'Department' ,t9.balance,   t5.slpname as 'Salesman', t10.name ,  T0.[DocTotal] as 'Total', T0.[PaidToDate] as 'Paid',t0.U_aggiustamentofattura as'Adjustment', T0.[DocTotal]-T0.[PaidToDate]-case when t0.U_aggiustamentofattura is null then '0' else t0.U_aggiustamentofattura end as 'Credit',   T0.[GroupNum], T3.[PymntGroup], t0.u_settore as 'Settore'
FROM OINV T0
INNER JOIN NNM1 T2 ON T0.[Series] = T2.[Series] 
inner join octg t3 on T0.[GroupNum]= t3.[GroupNum]

inner join OSLP T5 ON T5.slpcode =t0.slpcode

LEFT JOIN OCRD T9 ON T9.[CardCode] = T0.[CardCode]
INNER join OCRY t10 on t10.code = t9.country
left join ocrd t11 on t11.cardcode=t0.u_codicebp


WHERE T0.[docDate] >= (CONVERT(DATETIME, '20141001', 112) ) and T0.[DocTotal]-T0.[PaidToDate] >0 and t0.docentry<>'5848'and t0.docentry<>'6447' and t0.docentry<>'5199' and t0.docentry <>'7925' and t0.docentry<>'6882'and t0.docentry<>'7785'and t0.docentry<>'7528'and t0.docentry<>'7168' and t0.docentry<>'7426' and t0.docentry<>'7573' and t0.docentry<>'7932' and t0.docentry<>'8436'
)
as t10

where 0=0 " & filtro_codice_cliente & filtro_cliente & filtro_department & filtro_N_fattura & "

order by t10.overdue DESC 
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()
            DataGridView1.Rows.Add(False, cmd_SAP_reader("docentry"), cmd_SAP_reader("Fattura"), cmd_SAP_reader("Data Fatt"), cmd_SAP_reader("Data Scad"), cmd_SAP_reader("Overdue"), cmd_SAP_reader("BP code"), cmd_SAP_reader("BP Name"), cmd_SAP_reader("Final_BP_code"), cmd_SAP_reader("Final BP"), cmd_SAP_reader("Year"), cmd_SAP_reader("Department"), cmd_SAP_reader("Balance"), cmd_SAP_reader("Salesman"), cmd_SAP_reader("name"), cmd_SAP_reader("Total"), cmd_SAP_reader("Paid"), cmd_SAP_reader("Adjustment"), cmd_SAP_reader("Credit"), cmd_SAP_reader("PymntGroup"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()




    End Sub

    Private Sub Crediti_verso_clienti_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        inizializza_form()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Rows(e.RowIndex).Cells(columnName:="Overdue").Value >= 90 Then



            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightCoral

        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Overdue").Value >= 60 Then

            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Orange

        ElseIf DataGridView1.Rows(e.RowIndex).Cells(columnName:="Overdue").Value >= 30 Then
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LemonChiffon
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = "" Then
            filtro_N_fattura = ""
        Else
            filtro_N_fattura = " and t10.fattura= " & TextBox1.Text & ""
        End If
        riempi_datagridview_crediti()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = "" Then
            filtro_codice_cliente = ""
        Else
            filtro_codice_cliente = " and t10.[BP Code]   Like '%%" & TextBox2.Text & "%%' "
        End If
        riempi_datagridview_crediti()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = "" Then
            filtro_cliente = ""
        Else
            filtro_cliente = " and (t10.[BP name] Like '%%" & TextBox3.Text & "%%'  or t10.[Final BP]Like '%%" & TextBox3.Text & "%%') "
        End If
        riempi_datagridview_crediti()
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = "" Then
            filtro_department = ""
        Else
            filtro_department = " and t10.Department Like '%%" & TextBox4.Text & "%%' "
        End If
        riempi_datagridview_crediti()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Verifica se la cella della colonna "seleziona" è flaggata
            If CBool(row.Cells("seleziona").Value) = True Then
                ' Crea una nuova riga nella DataGridView "datagridview1"
                Dim index As Integer = DataGridView1.Rows.Add()

                ' Copia i valori dalle colonne necessarie


                Layout_documenti.Show()
                Layout_documenti.nome_documento_SAP = "Invoice"
                Layout_documenti.documento_SAP = "OINV"
                Layout_documenti.righe_SAP = "INV1"
                Layout_documenti.TextBox1.Text = row.Cells("Fattura").Value
                Layout_documenti.docnum = row.Cells("Fattura").Value



                Layout_documenti.Informazioni_documento(row.Cells("Fattura").Value)
                Layout_documenti.trova_word_base(Layout_documenti.Lingua, Layout_documenti.documento_SAP, Layout_documenti.garanzia, Layout_documenti.nome_documento_SAP)
                Layout_documenti.Genera_documento("")

                row.Cells("seleziona").Value = False


            End If
        Next
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

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Fattura) Then
                Form_nuova_offerta.Show()


                Form_nuova_offerta.TextBox10.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Fattura").Value
                Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                Form_nuova_offerta.inizializzazione_form(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Fattura").Value, "OINV", "INV1", "Fattura")
            End If

        End If

    End Sub
End Class