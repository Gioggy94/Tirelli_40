Imports System.Data.SqlClient

Public Class Form_statistiche_manodopera



    Sub record_per_dipendente(par_datagridview As DataGridView, par_numero_giorni As Integer, par_centro_di_costo As String, par_dipendente As Integer, par_reparto As String)
        Dim par_filtro_centro_di_costo As String

        If par_centro_di_costo = "" Then
            par_filtro_centro_di_costo = ""
        Else
            par_filtro_centro_di_costo = ""
            ' par_filtro_centro_di_costo = " and T1.COSTCENTER  Like ''%%" & par_centro_di_costo & "%%'' "
        End If

        Dim par_filtro_reparto As String

        If par_reparto = "" Then
            par_filtro_reparto = ""
        Else
            par_filtro_reparto = " and T2.name  Like ''%%" & par_reparto & "%%'' "
        End If



        Dim par_filtro_dipendente As String
        If par_dipendente = 0 Then
            par_filtro_dipendente = ""
        Else
            par_filtro_dipendente = " And t1.empid = ''" & par_dipendente & "'' "
            par_filtro_centro_di_costo = " "
        End If

        ' Inizializzo la connessione
        par_datagridview.Rows.Clear()
        par_datagridview.Columns.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "
DECLARE @cols  NVARCHAR(MAX),
        @query NVARCHAR(MAX),
        @num_days INT = " & par_numero_giorni & ";

------------------------------------------------------------
-- 1. Generazione colonne (ultimi @num_days giorni, no domenica)
------------------------------------------------------------
SET @cols = STUFF((
    SELECT ',' + QUOTENAME(CONVERT(VARCHAR(10), DATEADD(DAY, -n, GETDATE()), 120))
    FROM (
        SELECT TOP (@num_days)
               ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) - 1 AS n
        FROM master..spt_values
    ) AS Numbers
    WHERE DATENAME(WEEKDAY, DATEADD(DAY, -n, GETDATE())) <> 'Sunday'
    ORDER BY n
    FOR XML PATH(''), TYPE
).value('.', 'NVARCHAR(MAX)'), 1, 1, '');

------------------------------------------------------------
-- 2. Query dinamica
------------------------------------------------------------
SET @query = '
WITH DatiManodopera AS (
    SELECT
        T1.LASTNAME + '' '' + T1.FIRSTNAME AS Dipendente,
        T2.[Name] AS Reparto,
        '''' AS CentroCosto,
        CAST(T0.DATA AS DATE) AS Data,
        CASE 
            WHEN T0.start IS NULL OR T0.stop IS NULL THEN T0.consuntivo
            ELSE 
                CASE 
                    WHEN DATEPART(hour, T0.start) < 12
                     AND (
                            (DATEPART(hour, T0.stop) >= 13 AND DATEPART(minute, T0.stop) > 30)
                          OR DATEPART(hour, T0.stop) >= 14
                         )
                    THEN (DATEPART(hour, T0.stop) * 60 + DATEPART(minute, T0.stop))
                       - (DATEPART(hour, T0.start) * 60 + DATEPART(minute, T0.start))
                       + T0.consuntivo - 90
                    ELSE (DATEPART(hour, T0.stop) * 60 + DATEPART(minute, T0.stop))
                       - (DATEPART(hour, T0.start) * 60 + DATEPART(minute, T0.start))
                       + T0.consuntivo
                END
        END AS Minuti
    FROM [TIRELLI_40].[dbo].OHEM T1
    LEFT JOIN MANODOPERA T0
        ON T1.empid = T0.DIPENDENTE
       AND T0.DATA >= CAST(GETDATE() - @num_days AS DATE)
       AND DATENAME(WEEKDAY, T0.DATA) <> ''Sunday''
    LEFT JOIN [TIRELLI_40].[dbo].OUDP T2
        ON T1.[dept] = T2.[Code]
    WHERE 0 = 0
      AND T1.active = ''Y''
   " & par_filtro_centro_di_costo & par_filtro_dipendente & par_filtro_reparto & "
     
)
SELECT Dipendente, Reparto, CentroCosto, ' + @cols + '
FROM (
    SELECT
        Dipendente,
        Reparto,
        CentroCosto,
        Minuti,
        CONVERT(VARCHAR(10), Data, 120) AS DataPivot
    FROM DatiManodopera
) AS SourceData
PIVOT (
    SUM(Minuti)
    FOR DataPivot IN (' + @cols + ')
) AS PivotTable
ORDER BY Reparto, CentroCosto, Dipendente;
';

------------------------------------------------------------
-- 3. Esecuzione
------------------------------------------------------------
EXEC sp_executesql
    @query,
    N'@num_days INT',
    @num_days;
"

        ' Esegui la query e ottieni il reader
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        ' Aggiungi le colonne dinamiche nella DataGridView
        par_datagridview.Columns.Add("Dipendente", "Dipendente")
        par_datagridview.Columns.Add("Reparto", "Reparto")
        par_datagridview.Columns.Add("CentroCosto", "Centro Costo") ' Aggiungi la colonna CentroCosto

        ' Aggiungo le colonne per le date (gli ultimi n giorni)
        Dim giorni As New List(Of String)
        For i = 0 To par_numero_giorni - 1
            Dim data As Date = DateTime.Now.AddDays(-i)
            If data.DayOfWeek <> DayOfWeek.Sunday Then ' Escludi le domeniche
                ' Usa il formato yyyy-MM-dd per farlo corrispondere a quello della query
                Dim colName As String = data.ToString("yyyy-MM-dd")
                par_datagridview.Columns.Add(colName, colName)
                giorni.Add(colName) ' Memorizza i nomi delle colonne data
            End If
        Next

        ' Imposta le prime colonne (Dipendente e CentroCosto) come fisse
        par_datagridview.Columns(0).Frozen = True
        par_datagridview.Columns(1).Frozen = True
        par_datagridview.Columns(2).Frozen = True

        ' Imposta la DataGridView per adattarsi al contenuto
        For Each column As DataGridViewColumn In par_datagridview.Columns
            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells
        Next

        ' Inserisci i dati nelle righe della DataGridView
        Do While cmd_SAP_reader.Read()
            Dim row As New DataGridViewRow
            row.CreateCells(par_datagridview)
            row.Cells(0).Value = cmd_SAP_reader("Dipendente")
            row.Cells(1).Value = cmd_SAP_reader("Reparto")
            row.Cells(2).Value = cmd_SAP_reader("CentroCosto") ' Imposta il valore della cella CentroCosto
            row.Cells(2).Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' Compila le celle per ogni giorno
            For Each giorno In giorni
                Dim cellValue As Object = If(IsDBNull(cmd_SAP_reader(giorno)), 0, cmd_SAP_reader(giorno))
                row.Cells(par_datagridview.Columns(giorno).Index).Value = cellValue
                row.Cells(par_datagridview.Columns(giorno).Index).Style.Alignment = DataGridViewContentAlignment.MiddleCenter

                ' Colora le celle con valori inferiori a 400 in rosso
                If Convert.ToInt32(cellValue) < 400 Then
                    row.Cells(par_datagridview.Columns(giorno).Index).Style.BackColor = Color.Red
                    row.Cells(par_datagridview.Columns(giorno).Index).Style.ForeColor = Color.White
                End If
            Next

            par_datagridview.Rows.Add(row)
        Loop

        cmd_SAP_reader.Close()
        Cnn.Close()

        ' Scorre alla fine della DataGridView
        Try
            par_datagridview.FirstDisplayedScrollingRowIndex = par_datagridview.RowCount - 1
        Catch ex As Exception
            ' Gestione dell'errore di scorrimento
        End Try
        par_datagridview.ClearSelection()
    End Sub


    Private Sub Form_statistiche_manodopera_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inizializza_form()
    End Sub

    Sub inizializza_form()
        record_per_dipendente(DataGridView4, TextBox1.Text, TextBox2.Text, 0, TextBox3.Text)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        inizializza_form()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        inizializza_form()
    End Sub

    Private Sub DataGridView4_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView4.CellFormatting

        Dim par_datagridview As DataGridView = DataGridView4
        Dim divValue As String = par_datagridview.Rows(e.RowIndex).Cells("CentroCosto").Value.ToString()
        Select Case divValue
            Case "BRB01"
                par_datagridview.Rows(e.RowIndex).Cells("CentroCosto").Style.BackColor = Color.Yellow
            Case "TIR01"
                par_datagridview.Rows(e.RowIndex).Cells("CentroCosto").Style.BackColor = Color.LightBlue
            Case "KTF01"
                par_datagridview.Rows(e.RowIndex).Cells("CentroCosto").Style.BackColor = Color.Green
        End Select
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        inizializza_form()
    End Sub
End Class