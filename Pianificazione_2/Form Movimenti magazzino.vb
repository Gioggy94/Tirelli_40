Imports System.Data.SqlClient
Imports Microsoft.Office.Interop

Public Class Form_Movimenti_magazzino
    Private Sub Button4_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Sub movimenti(par_datagridview As DataGridView, par_utente_galileo As String, par_codice_sap As String, par_mag As String, par_commessa As String, par_sottocommessa As String, par_datetimepicker1 As DateTimePicker, par_datetimepicker2 As DateTimePicker)


        Dim dataInizio As String = par_datetimepicker1.Value.ToString("yyyyMMdd")
        Dim dataFine As String = par_datetimepicker2.Value.ToString("yyyyMMdd")

        par_datagridview.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "SELECT * FROM OPENQUERY(AS400, '
     SELECT trim(t0.CODART) as codart
	, t0.DOCDATE
	, t0.time, t0.CODCAU, t0.TRANSNAME
	, trim(t0.REF1) as ref1, t0.PROMAG
	, trim(t0.LOCCODE) as Loccode
	,  case when segno =''-'' then -t0.movimento else t0.movimento end as movimento
	, case when trim(t0.codcomm)='''' then coalesce(trim(t2.cod_commessa),'''') else trim(t0.codcomm) end CODCOMM
	, t0.DESCOM, t0.CLICOM, trim(t0.DS_CLICOM) as DS_CLICOM, t0.RIFMAG
, trim(t0.CARDNAME) as cardname, t0.PRICE
, case when T0.MATRICOLA ='''' then coalesce(trim(t2.matricola),'''') else trim(t0.matricola) end as matricola
, case when T0.SOTTOCOMMESSA='''' then coalesce(trim(t2.cod_sottocommessa),'''') else trim(t0.sottocommessa) end as Sottocommessa
, t0.LASTNAME, t0.CTRL_QUA, t0.SEGNO, t0.DATA_REG, t0.NUM_OPE, t0.RIG_OPE
   
FROM S786FAD1.TIR90VIS.JGALMOV t0
left join TIR90VIS.JGALodp t2 on trim(t0.commento)=t2.numodp and trim(t0.commento)<>''''
    WHERE 0=0
    AND UPPER(t0.LASTNAME) LIKE ''%" & par_utente_galileo & "%''
    AND UPPER(t0.CODART)   LIKE ''%" & par_codice_sap & "%''
    AND t0.DOCDATE >= " & dataInizio & "
    AND t0.DOCDATE <= " & dataFine & "
AND UPPER(t0.loccode)  LIKE ''%" & par_mag & "%''
AND UPPER(case when trim(t0.codcomm)='''' then coalesce(trim(t2.cod_commessa),'''') else trim(t0.codcomm) end)  LIKE ''%" & par_commessa & "%''
AND UPPER(case when T0.SOTTOCOMMESSA='''' then coalesce(trim(t2.cod_sottocommessa),'''') else trim(t0.sottocommessa) end)  LIKE ''%" & par_sottocommessa & "%''
    ORDER BY t0.TIME
    LIMIT 30000
') AS t10"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Dim contatore As Integer = 0
        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add(
                    cmd_SAP_reader("CODART"),
                    cmd_SAP_reader("DOCDATE"),
                     cmd_SAP_reader("time"),
                    cmd_SAP_reader("CODCAU"),
                    cmd_SAP_reader("TRANSNAME"),
                    cmd_SAP_reader("REF1"),
                    cmd_SAP_reader("PROMAG"),
                    cmd_SAP_reader("LOCCODE"),
                    cmd_SAP_reader("MOVIMENTO"),
                    cmd_SAP_reader("CODCOMM"),
                    cmd_SAP_reader("SOTTOCOMMESSA"),
                    cmd_SAP_reader("DESCOM"),
                    cmd_SAP_reader("CLICOM"),
                    cmd_SAP_reader("DS_CLICOM"),
                    cmd_SAP_reader("RIFMAG"),
                    cmd_SAP_reader("CARDNAME"),
                    cmd_SAP_reader("PRICE"),
                    cmd_SAP_reader("MATRICOLA"),
                    cmd_SAP_reader("LASTNAME"),
                    cmd_SAP_reader("CTRL_QUA"),
                    cmd_SAP_reader("SEGNO"),
                    cmd_SAP_reader("DATA_REG"),
                    cmd_SAP_reader("NUM_OPE"),
                    cmd_SAP_reader("RIG_OPE")
                )
            contatore += 1
        Loop
        Label1.Text = contatore
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub


    Private Sub Form_stampe_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        filtra()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        filtra()
    End Sub

    Private Sub Cmd_Indietro_Click(sender As Object, e As EventArgs) Handles Cmd_Indietro.Click
        DateTimePicker1.Value = DateTimePicker1.Value.AddDays(-1)
        DateTimePicker2.Value = DateTimePicker2.Value.AddDays(-1)
        filtra()
    End Sub

    Private Sub Cmd_Avanti_Click(sender As Object, e As EventArgs) Handles Cmd_Avanti.Click
        DateTimePicker1.Value = DateTimePicker1.Value.AddDays(1)
        DateTimePicker2.Value = DateTimePicker2.Value.AddDays(1)
        filtra()
    End Sub

    Sub filtra()
        movimenti(DataGridView1, TextBox1.Text.ToUpper, TextBox2.Text.ToUpper, TextBox3.Text.ToUpper, TextBox5.Text.ToUpper, TextBox6.Text.ToUpper, DateTimePicker1, DateTimePicker2)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim par_datagridview As DataGridView = DataGridView1
        ' Creare un'applicazione Excel
        Dim excelApp As New Excel.Application
        excelApp.Visible = True ' Mostrare Excel all'utente

        ' Creare un nuovo foglio di lavoro
        Dim excelWorkbook As Excel.Workbook = excelApp.Workbooks.Add
        Dim excelWorksheet As Excel.Worksheet = CType(excelWorkbook.Worksheets(1), Excel.Worksheet)

        ' Aggiungere intestazioni alla prima riga del foglio di lavoro (facoltativo)
        For col As Integer = 1 To par_datagridview.Columns.Count
            excelWorksheet.Cells(1, col) = par_datagridview.Columns(col - 1).HeaderText
        Next

        ' Aggiungere dati alla DataGridView al foglio di lavoro
        For row As Integer = 0 To par_datagridview.Rows.Count - 1
            For col As Integer = 0 To par_datagridview.Columns.Count - 1
                excelWorksheet.Cells(row + 2, col + 1) = par_datagridview.Rows(row).Cells(col).Value
            Next
        Next

        ' Salvare il file Excel
        Dim saveFileDialog As New SaveFileDialog
        saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx"
        If saveFileDialog.ShowDialog() = DialogResult.OK Then
            excelWorkbook.SaveAs(saveFileDialog.FileName)
            MessageBox.Show("Esportazione completata con successo!", "Successo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        ' Chiudere Excel
        excelApp.Quit()
        Form_stato_commesse.ReleaseComObject(excelApp)
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub Button4_Click_2(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
End Class