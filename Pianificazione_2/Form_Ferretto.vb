Imports System.Data.SqlClient
Imports System.Drawing.Printing

Public Class Form_Ferretto
    Public id_da_stampare As Integer = 0
    Public Stampante_Selezionata As Boolean
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        movimenti(DataGridView1)
    End Sub


    Sub movimenti(par_datagridview As DataGridView)


        par_datagridview.Rows.Clear()

        ' --- Connessione e query ---
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT TOP (1000)
    [id],
    [recordStatus],
    [recordWritingDate],
    [recordImportationDate],
    [plantId],
    [response],
    [listType],
    [listNumber],
    CASE 
        WHEN CHARINDEX('OD:', listNumber) > 0 AND CHARINDEX('-COM:', listNumber) > 0
             AND CHARINDEX('-COM:', listNumber) > CHARINDEX('OD:', listNumber)
        THEN SUBSTRING(listNumber, CHARINDEX('OD:', listNumber) + 3,
            CHARINDEX('-COM:', listNumber) - CHARINDEX('OD:', listNumber) - 3)
        ELSE NULL
    END AS OD_Number,
    CASE 
        WHEN listNumber LIKE 'OC-%'
        THEN SUBSTRING(listNumber, 4, LEN(listNumber))
        ELSE NULL
    END AS OC_Number,
    CASE 
        WHEN CHARINDEX('COM:', listNumber) > 0 AND CHARINDEX('-ART:', listNumber) > 0
             AND CHARINDEX('-ART:', listNumber) > CHARINDEX('COM:', listNumber)
        THEN SUBSTRING(listNumber, CHARINDEX('COM:', listNumber) + 4,
            CHARINDEX('-ART:', listNumber) - CHARINDEX('COM:', listNumber) - 4)
        ELSE NULL
    END AS COM,
    CASE 
        WHEN CHARINDEX('-ID:', listNumber) > 0 AND CHARINDEX('-MG:', listNumber) > 0
             AND CHARINDEX('-MG:', listNumber) > CHARINDEX('-ID:', listNumber)
        THEN REPLACE(SUBSTRING(listNumber, CHARINDEX('-ID:', listNumber) + 3,
            CHARINDEX('-MG:', listNumber) - CHARINDEX('-ID:', listNumber) - 3), ':', '')
        ELSE NULL
    END AS ID_Value,
    CASE 
        WHEN CHARINDEX('-MG:', listNumber) > 0
        THEN SUBSTRING(listNumber, CHARINDEX('-MG:', listNumber) + 4, LEN(listNumber))
        ELSE NULL
    END AS Magazzino,
    [lineNumber],
    [item],
    [batch],
    [serialNumber],
    [requestedQty],
    [processedQty],
    [errorCause],
    [wmsGenerated]
FROM [FGWmsErp].[dbo].[LISTS_RESULT]
WHERE response <> 11 and listtype=0 and [processedQty]>0
ORDER BY id DESC;"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Dim contatore As Integer = 0
        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add(
            cmd_SAP_reader("id"),
            cmd_SAP_reader("recordStatus"),
            cmd_SAP_reader("recordWritingDate"),
            cmd_SAP_reader("recordImportationDate"),
            cmd_SAP_reader("plantId"),
            cmd_SAP_reader("response"),
            cmd_SAP_reader("listType"),
            cmd_SAP_reader("listNumber"),
            cmd_SAP_reader("OD_Number"),
            cmd_SAP_reader("OC_Number"),
            cmd_SAP_reader("COM"),
            cmd_SAP_reader("ID_Value"),
            cmd_SAP_reader("Magazzino"),
            cmd_SAP_reader("lineNumber"),
            cmd_SAP_reader("item"),
            cmd_SAP_reader("batch"),
            cmd_SAP_reader("serialNumber"),
            cmd_SAP_reader("requestedQty"),
            cmd_SAP_reader("processedQty"),
            cmd_SAP_reader("errorCause"),
            cmd_SAP_reader("wmsGenerated")
        )
            contatore += 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.AutoGenerateColumns = False
        genera_colonne(DataGridView1)
        movimenti(DataGridView1)
    End Sub

    Sub genera_colonne(par_datagridview As DataGridView)
        par_datagridview.Columns.Clear()
        par_datagridview.AutoGenerateColumns = False
        par_datagridview.Columns.Add("id", "ID")
        par_datagridview.Columns.Add("recordStatus", "Stato Record")
        par_datagridview.Columns.Add("recordWritingDate", "Data Scrittura")
        par_datagridview.Columns.Add("recordImportationDate", "Data Importazione")
        par_datagridview.Columns.Add("plantId", "Plant")
        par_datagridview.Columns.Add("response", "Response")
        par_datagridview.Columns.Add("listType", "Tipo Lista")
        par_datagridview.Columns.Add("listNumber", "Numero Lista")
        par_datagridview.Columns.Add("OD_Number", "N° OD")
        par_datagridview.Columns.Add("OC_Number", "N° OC")
        par_datagridview.Columns.Add("COM", "COM")
        par_datagridview.Columns.Add("ID_Value", "ID Valore")
        par_datagridview.Columns.Add("Magazzino", "Magazzino")
        par_datagridview.Columns.Add("lineNumber", "N° Riga")
        par_datagridview.Columns.Add("item", "Articolo")
        par_datagridview.Columns.Add("batch", "Batch")
        par_datagridview.Columns.Add("serialNumber", "Matricola")
        par_datagridview.Columns.Add("requestedQty", "Qtà Richiesta")
        par_datagridview.Columns.Add("processedQty", "Qtà Processata")
        par_datagridview.Columns.Add("errorCause", "Causa Errore")
        par_datagridview.Columns.Add("wmsGenerated", "WMS Generato")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellBorderStyleChanged(sender As Object, e As EventArgs) Handles DataGridView1.CellBorderStyleChanged

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            id_da_stampare = DataGridView1.Rows(e.RowIndex).Cells(columnName:="id").Value
            RichTextBox1.Text = id_da_stampare
        End If
    End Sub

    ' =============================================
    ' VARIABILI GLOBALI DI STAMPA
    ' =============================================
    Dim stampa_OD As String = ""
    Dim stampa_OC As String = ""
    Dim stampa_COM As String = ""
    Dim stampa_ID As String = ""
    Dim stampa_articolo As String = ""
    Dim stampa_qta_richiesta As String = ""
    Dim stampa_qta_processata As String = ""
    Dim stampa_numero_scontrino As String = ""
    Dim stampa_baia As String = ""
    Dim stampa_stato_odp As String = ""
    Dim stampa_magazzino_destinazione As String = ""

    ' =============================================
    ' CLICK PULSANTE STAMPA
    ' =============================================
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Prende l'id dalla riga selezionata nella datagridview
        If id_da_stampare = 0 Then
            MessageBox.Show("Seleziona una riga da stampare.")
            Exit Sub
        End If



        Stampa_stampa(id_da_stampare)
    End Sub

    ' =============================================
    ' CARICA DATI E AVVIA STAMPA
    ' =============================================
    Sub Stampa_stampa(par_id_stampa As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD As New SqlCommand
        CMD.Connection = Cnn
        CMD.CommandText = "SELECT
    [id],
    [listNumber],
    [item],
    [requestedQty],
    [processedQty],
    CASE 
        WHEN CHARINDEX('OD:', listNumber) > 0 AND CHARINDEX('-COM:', listNumber) > 0
             AND CHARINDEX('-COM:', listNumber) > CHARINDEX('OD:', listNumber)
        THEN SUBSTRING(listNumber, CHARINDEX('OD:', listNumber) + 3,
            CHARINDEX('-COM:', listNumber) - CHARINDEX('OD:', listNumber) - 3)
        ELSE NULL
    END AS OD_Number,
    CASE 
        WHEN listNumber LIKE 'OC-%'
        THEN SUBSTRING(listNumber, 4, LEN(listNumber))
        ELSE NULL
    END AS OC_Number,
    CASE 
        WHEN CHARINDEX('COM:', listNumber) > 0 AND CHARINDEX('-ART:', listNumber) > 0
             AND CHARINDEX('-ART:', listNumber) > CHARINDEX('COM:', listNumber)
        THEN SUBSTRING(listNumber, CHARINDEX('COM:', listNumber) + 4,
            CHARINDEX('-ART:', listNumber) - CHARINDEX('COM:', listNumber) - 4)
        ELSE NULL
    END AS COM,
    CASE 
        WHEN CHARINDEX('-ID:', listNumber) > 0 AND CHARINDEX('-MG:', listNumber) > 0
             AND CHARINDEX('-MG:', listNumber) > CHARINDEX('-ID:', listNumber)
        THEN REPLACE(SUBSTRING(listNumber, CHARINDEX('-ID:', listNumber) + 3,
            CHARINDEX('-MG:', listNumber) - CHARINDEX('-ID:', listNumber) - 3), ':', '')
        ELSE NULL
    END AS ID_Value
FROM [FGWmsErp].[dbo].[LISTS_RESULT]
WHERE id = " & par_id_stampa

        Dim reader As SqlDataReader = CMD.ExecuteReader

        If reader.Read() Then
            stampa_numero_scontrino = reader("id").ToString()
            stampa_OD = If(IsDBNull(reader("OD_Number")), "", reader("OD_Number").ToString())
            stampa_OC = If(IsDBNull(reader("OC_Number")), "", reader("OC_Number").ToString())
            stampa_COM = If(IsDBNull(reader("COM")), "", reader("COM").ToString())
            stampa_ID = If(IsDBNull(reader("ID_Value")), "", reader("ID_Value").ToString())
            stampa_articolo = If(IsDBNull(reader("item")), "", reader("item").ToString())
            stampa_qta_richiesta = If(IsDBNull(reader("requestedQty")), "", reader("requestedQty").ToString())
            stampa_qta_processata = If(IsDBNull(reader("processedQty")), "", reader("processedQty").ToString())

            ' --- BLOCCA SE QTA PROCESSATA È VUOTA O ZERO ---
            If stampa_qta_processata = "" OrElse stampa_qta_processata = "0" OrElse stampa_qta_processata = "0.000" Then
                MessageBox.Show("Nessuna quantità processata, stampa annullata.")
                reader.Close()
                Cnn.Close()
                Exit Sub
            End If

            ' --- BAIA e STATO dall'OD ---
            If stampa_OD <> "" Then
                Dim info_odp = ODP_Form.ottieni_informazioni_odp("Numero", 0, stampa_OD)
                stampa_baia = info_odp.nome_baia
                stampa_stato_odp = info_odp.stato
                stampa_magazzino_destinazione = info_odp.warehouse
            Else
                stampa_baia = ""
                stampa_stato_odp = ""
                stampa_magazzino_destinazione = ""
            End If

        Else
            MessageBox.Show("ID non trovato: " & par_id_stampa)
            reader.Close()
            Cnn.Close()
            Exit Sub
        End If

        reader.Close()
        Cnn.Close()

        ' --- GESTIONE STAMPA / PREVIEW ---



        ' --- STAMPA DIRETTA SENZA PREVIEW ---
        If Stampante_Selezionata = False Then
                Dim sel As New PrintDialog
                sel.Document = Scontrino

                If sel.ShowDialog() = DialogResult.OK Then
                Stampante_Selezionata = True
                Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Scontrino", 185, 215)
                Scontrino.Print()
            End If
            Else
            ' Stampante già selezionata → stampa diretta
            Scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Scontrino", 185, 215)
            Scontrino.Print()
        End If




    End Sub

    Private Sub Scontrino_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles Scontrino.PrintPage

        Dim Penna As New Pen(Color.Black)
        Dim fTitolo As New Font("Calibri", 6, FontStyle.Italic)
        Dim fSmall As New Font("Calibri", 7, FontStyle.Regular)
        Dim fGrande As New Font("Calibri", 14, FontStyle.Bold)
        Dim fArticolo As New Font("Calibri", 14, FontStyle.Italic)
        Dim fQta As New Font("Calibri", 11, FontStyle.Bold)

        Dim g As System.Drawing.Graphics = e.Graphics
        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

        ' --- BORDO ESTERNO ---
        g.DrawRectangle(Penna, 1, 1, 183, 210)

        ' --- INTESTAZIONE: numero scontrino ---
        g.DrawString("N° " & stampa_numero_scontrino, fTitolo, Brushes.Gray, 3, 2)

        ' -------------------------------------------------------
        ' RIGA 1: [N° OD + Stato (sx)] [N° OC (dx)]
        ' -------------------------------------------------------
        g.DrawRectangle(Penna, 3, 10, 85, 45)
        g.DrawString("N° OD", fTitolo, Brushes.Black, 5, 12)
        g.DrawString(stampa_OD, fGrande, Brushes.Black, 5, 20)
        g.DrawString("Stato: " & stampa_stato_odp, fSmall, Brushes.Gray, 5, 38)

        g.DrawRectangle(Penna, 95, 10, 87, 45)
        g.DrawString("N° OC", fTitolo, Brushes.Black, 97, 12)
        g.DrawString(stampa_OC, fGrande, Brushes.Black, 97, 20)

        ' -------------------------------------------------------
        ' RIGA 2: [COM (sx 70px)] [ID (40px)] [Mag (dx 52px)]
        ' -------------------------------------------------------
        g.DrawRectangle(Penna, 3, 60, 70, 35)
        g.DrawString("COM", fTitolo, Brushes.Black, 5, 62)
        g.DrawString(stampa_COM, fGrande, Brushes.Black, 5, 72)

        g.DrawRectangle(Penna, 78, 60, 40, 35)
        g.DrawString("ID", fTitolo, Brushes.Black, 80, 62)
        g.DrawString(stampa_ID, fTitolo, Brushes.Black, 80, 72)

        g.DrawRectangle(Penna, 123, 60, 59, 35)
        g.DrawString("Mag", fTitolo, Brushes.Black, 125, 62)
        g.DrawString(stampa_magazzino_destinazione, fSmall, Brushes.Black, 125, 72)

        ' -------------------------------------------------------
        ' RIGA 3: [Articolo (sx 120px)] [Baia (dx 62px)]
        ' -------------------------------------------------------
        Dim articolo_troncato As String = stampa_articolo.Substring(0, Math.Min(8, stampa_articolo.Length))
        g.DrawRectangle(Penna, 3, 100, 117, 35)
        g.DrawString("Articolo", fTitolo, Brushes.Black, 5, 102)
        g.DrawString(articolo_troncato, fArticolo, Brushes.Black, 5, 110)

        g.DrawRectangle(Penna, 125, 100, 57, 35)
        g.DrawString("Baia", fTitolo, Brushes.Black, 127, 102)
        g.DrawString(stampa_baia, fGrande, Brushes.Black, 127, 110)

        ' -------------------------------------------------------
        ' RIGA 4: Qtà Processata (larghezza piena)
        ' -------------------------------------------------------
        g.DrawRectangle(Penna, 3, 140, 179, 30)
        g.DrawString("Qtà Processata", fTitolo, Brushes.Black, 5, 142)
        g.DrawString(stampa_qta_processata, fQta, Brushes.Black, 5, 151)

        ' -------------------------------------------------------
        ' FOOTER: data/ora + utente
        ' -------------------------------------------------------
        Dim DataOggi As String = DateTime.Now.ToString("dd/MM/yy HH:mm")
        g.DrawString(DataOggi, fSmall, Brushes.Black, 120, 197)

        Dim NomeDipendente As String = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).COGNOME &
                               " " & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).NOME
        Dim NomeTroncato As String = NomeDipendente.Substring(0, Math.Min(12, NomeDipendente.Length))
        g.DrawString(NomeTroncato, fSmall, Brushes.Black, 5, 197)

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim id_da_stampare_auto As Integer = max_id_ferretto()
        If id_da_stampare_auto <> id_da_stampare Then
            id_da_stampare = id_da_stampare_auto
            Stampa_stampa(id_da_stampare)
        End If

    End Sub


    Public Function max_id_ferretto()

        Dim id As Integer = 0

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD As New SqlCommand
        CMD.Connection = Cnn
        CMD.CommandText = "SELECT top 1
    [id]
    
FROM [FGWmsErp].[dbo].[LISTS_RESULT]
where response<>11 and listtype=0 and [processedQty]>0
order by id desc"

        Dim reader As SqlDataReader = CMD.ExecuteReader

        If reader.Read() Then
            id = reader("id").ToString()

        End If

        reader.Close()
        Cnn.Close()
        Return id
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Timer1.Start()
        Button3.Visible = True
        Panel1.BackColor = Color.Lime
        Button2.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Timer1.Stop()
        Button2.Visible = True
        Button3.Visible = False
        Panel1.BackColor = Color.IndianRed
    End Sub
End Class