Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text
Imports System.ComponentModel

' ============================================================
' Entrate_merci.vb
' Form per il riconoscimento automatico di bolle DDT (PDF)
' e inserimento in SQL nella tabella Entrate_Merci.
'
' Prerequisiti Python:
'   pip install pymupdf anthropic
'
' SQL - eseguire Entrate_merci_ALTER.sql una volta su SQL Server
' per aggiungere i campi necessari a [Tirelli_40].[dbo].[Entrate_merci]
' ============================================================

Public Class Entrate_merci

    Private Const API_KEY_FILE As String = ".\anthropic_key.txt"
    Private _pdfPath As String = ""
    Private _bgWorker As BackgroundWorker

    ' ----------------------------------------------------------------
    ' REGISTRO PECULIARITÀ FORNITORE
    ' Chiave = COD_forn AS400 (trim, uppercase)
    ' Valore = array di coppie (pattern regex, sostituzione) da applicare
    '          al codice/disegno del DDT prima del confronto con AS400
    ' Per aggiungere un fornitore: d("COD") = New String() {"pattern", "repl"}
    ' ----------------------------------------------------------------
    Private Shared ReadOnly RegoleFornitore As Dictionary(Of String, String()) = CreaRegoleFornitore()

    Private Shared Function CreaRegoleFornitore() As Dictionary(Of String, String())
        Dim d As New Dictionary(Of String, String())(StringComparer.OrdinalIgnoreCase)
        ' Lasergi (1410002492): scrive "#N" invece di "-RN" e "-" invece di "_"
        d("1410002492") = New String() {"#(\d+)", "-R$1", "_", "-"}
        Return d
    End Function

    ' ----------------------------------------------------------------
    ' INIZIALIZZAZIONE
    ' ----------------------------------------------------------------
    Private Sub Entrate_merci_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CaricaChiaveApi()
        ImpostaGriglia()
        IniziaBgWorker()
    End Sub

    Private Sub CaricaChiaveApi()
        Dim keyPath As String = Path.Combine(Application.StartupPath, "anthropic_key.txt")
        If File.Exists(keyPath) Then
            txtApiKey.Text = File.ReadAllText(keyPath).Trim()
        End If
    End Sub

    Private Sub SalvaChiaveApi()
        Dim keyPath As String = Path.Combine(Application.StartupPath, "anthropic_key.txt")
        Try
            File.WriteAllText(keyPath, txtApiKey.Text.Trim())
        Catch
        End Try
    End Sub

    Private Sub ImpostaGriglia()
        dgvRighe.RowTemplate.Height = 26
        dgvRighe.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 245, 255)
        dgvRighe.DefaultCellStyle.Font = New Font("Segoe UI", 9)
        dgvRighe.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        dgvRighe.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(45, 85, 140)
        dgvRighe.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvRighe.EnableHeadersVisualStyles = False

        ' Colonna disegno (aggiunta qui per non modificare il Designer)
        Dim colDisegno As New DataGridViewTextBoxColumn()
        colDisegno.Name = "colDisegno"
        colDisegno.HeaderText = "Disegno"
        colDisegno.Width = 120
        colDisegno.ReadOnly = True
        dgvRighe.Columns.Add(colDisegno)

        ' Colonna stato confronto ordini (aggiunta qui per non modificare il Designer)
        Dim colStato As New DataGridViewTextBoxColumn()
        colStato.Name = "colStato"
        colStato.HeaderText = "Stato / Scostamento"
        colStato.Width = 220
        colStato.ReadOnly = True
        dgvRighe.Columns.Add(colStato)
    End Sub

    Private Sub IniziaBgWorker()
        _bgWorker = New BackgroundWorker()
        _bgWorker.WorkerReportsProgress = False
        _bgWorker.WorkerSupportsCancellation = False
        AddHandler _bgWorker.DoWork, AddressOf BgWorker_DoWork
        AddHandler _bgWorker.RunWorkerCompleted, AddressOf BgWorker_Completed
    End Sub

    ' ----------------------------------------------------------------
    ' SELEZIONE FILE
    ' ----------------------------------------------------------------
    Private Sub btnSfoglia_Click(sender As Object, e As EventArgs) Handles btnSfoglia.Click
        Using dlg As New OpenFileDialog()
            dlg.Title = "Seleziona bolla DDT"
            dlg.Filter = "File PDF (*.pdf)|*.pdf|Tutti i file (*.*)|*.*"
            dlg.FilterIndex = 1
            If dlg.ShowDialog() = DialogResult.OK Then
                _pdfPath = dlg.FileName
                txtFilePath.Text = _pdfPath
                dgvRighe.Rows.Clear()
                ImpostaMsg("", Color.Black)
                lblStatus.Text = "File selezionato. Premi 'Analizza PDF' per avviare il riconoscimento."
            End If
        End Using
    End Sub

    ' ----------------------------------------------------------------
    ' ANALISI PDF (Background Worker)
    ' ----------------------------------------------------------------
    Private Sub btnAnalizza_Click(sender As Object, e As EventArgs) Handles btnAnalizza.Click
        If String.IsNullOrWhiteSpace(_pdfPath) Then
            MsgBox("Seleziona prima un file PDF.", MsgBoxStyle.Exclamation, "File mancante")
            Return
        End If
        If String.IsNullOrWhiteSpace(txtApiKey.Text) Then
            MsgBox("Inserisci la chiave API Claude.", MsgBoxStyle.Exclamation, "Chiave API mancante")
            Return
        End If

        SalvaChiaveApi()
        dgvRighe.Rows.Clear()
        ImpostaMsg("", Color.Black)
        ImpostaControlliAbilitati(False)
        lblStatus.Text = "Analisi in corso... (puo' richiedere alcuni minuti)"

        _bgWorker.RunWorkerAsync(New String() {_pdfPath, txtApiKey.Text.Trim()})
    End Sub

    Private Sub BgWorker_DoWork(sender As Object, e As DoWorkEventArgs)
        Dim args() As String = CType(e.Argument, String())
        Dim pdfPath As String = args(0)
        Dim apiKey As String = args(1)

        ' Trova lo script Python
        Dim scriptPath As String = Path.Combine(Application.StartupPath, "entrate_merci_ocr.py")
        If Not File.Exists(scriptPath) Then
            ' Prova nella stessa cartella del sorgente (sviluppo)
            scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "entrate_merci_ocr.py")
        End If
        If Not File.Exists(scriptPath) Then
            e.Result = "{""errore"":""Script entrate_merci_ocr.py non trovato in: " & Application.StartupPath & """}"
            Return
        End If

        ' Prepara argomenti (gestione percorsi con spazi)
        Dim pdfArg As String = """" & pdfPath & """"
        Dim keyArg As String = """" & apiKey & """"
        Dim scriptArg As String = """" & scriptPath & """"

        Dim psi As New Diagnostics.ProcessStartInfo()
        psi.FileName = "python"
        psi.Arguments = scriptArg & " " & pdfArg & " " & keyArg
        psi.UseShellExecute = False
        psi.RedirectStandardOutput = True
        psi.RedirectStandardError = True
        psi.CreateNoWindow = True
        psi.StandardOutputEncoding = Encoding.UTF8

        Dim output As String = ""
        Dim errout As String = ""

        Try
            Using proc As Diagnostics.Process = Diagnostics.Process.Start(psi)
                output = proc.StandardOutput.ReadToEnd()
                errout = proc.StandardError.ReadToEnd()
                proc.WaitForExit(300000) ' 5 minuti timeout
            End Using
        Catch ex As Exception
            e.Result = "{""errore"":""Impossibile avviare Python: " & ex.Message.Replace("""", "'") & """}"
            Return
        End Try

        If String.IsNullOrWhiteSpace(output) Then
            Dim msg As String = If(String.IsNullOrWhiteSpace(errout), "Nessun output dallo script Python.", errout)
            e.Result = "{""errore"":""" & msg.Replace("""", "'").Replace(Chr(13), " ").Replace(Chr(10), " ") & """}"
            Return
        End If

        e.Result = output.Trim()
    End Sub

    Private Sub BgWorker_Completed(sender As Object, e As RunWorkerCompletedEventArgs)
        ImpostaControlliAbilitati(True)

        If e.Error IsNot Nothing Then
            lblStatus.Text = "Errore: " & e.Error.Message
            Return
        End If

        Dim json As String = CStr(e.Result)

        ' Controlla se e' un errore
        If json.TrimStart().StartsWith("{") Then
            Dim errMsg As String = EstraiCampoJson(json, "errore")
            lblStatus.Text = "Errore analisi: " & errMsg
            ImpostaMsg("Errore: " & errMsg, Color.Red)
            Return
        End If

        ' Popola la griglia
        Dim righe As List(Of Dictionary(Of String, String)) = ParseJsonArray(json)

        If righe.Count = 0 Then
            lblStatus.Text = "Nessun articolo trovato nel documento."
            ImpostaMsg("Nessun articolo riconosciuto.", Color.OrangeRed)
            Return
        End If

        For Each riga In righe
            Dim idx As Integer = dgvRighe.Rows.Add()
            dgvRighe.Rows(idx).Cells("colSel").Value = True
            dgvRighe.Rows(idx).Cells("colDDT").Value = GetVal(riga, "ddt_numero")
            dgvRighe.Rows(idx).Cells("colData").Value = GetVal(riga, "ddt_data")
            dgvRighe.Rows(idx).Cells("colFornitore").Value = GetVal(riga, "fornitore")
            dgvRighe.Rows(idx).Cells("colOrdine").Value = GetVal(riga, "ordine")
            dgvRighe.Rows(idx).Cells("colCodice").Value = GetVal(riga, "codice")
            dgvRighe.Rows(idx).Cells("colDescrizione").Value = GetVal(riga, "descrizione")
            dgvRighe.Rows(idx).Cells("colUM").Value = GetVal(riga, "um")
            dgvRighe.Rows(idx).Cells("colQuantita").Value = GetVal(riga, "quantita")
        Next

        ImpostaMsg($"{righe.Count} righe riconosciute. Confronto ordini in corso...", Color.DarkBlue)
        lblStatus.Text = $"Analisi completata: {righe.Count} righe. Confronto con ordini aperti..."
        VerificaOrdiniAS400()
    End Sub

    ' ----------------------------------------------------------------
    ' SELEZIONA / DESELEZIONA
    ' ----------------------------------------------------------------
    Private Sub btnSelTutto_Click(sender As Object, e As EventArgs) Handles btnSelTutto.Click
        For Each row As DataGridViewRow In dgvRighe.Rows
            row.Cells("colSel").Value = True
        Next
    End Sub

    Private Sub btnDeselTutto_Click(sender As Object, e As EventArgs) Handles btnDeselTutto.Click
        For Each row As DataGridViewRow In dgvRighe.Rows
            row.Cells("colSel").Value = False
        Next
    End Sub

    ' ----------------------------------------------------------------
    ' SALVATAGGIO SQL
    ' ----------------------------------------------------------------
    Private Sub btnSalva_Click(sender As Object, e As EventArgs) Handles btnSalva.Click
        Dim righeSelezionate As New List(Of DataGridViewRow)
        For Each row As DataGridViewRow In dgvRighe.Rows
            If row.Cells("colSel").Value IsNot Nothing AndAlso CBool(row.Cells("colSel").Value) = True Then
                righeSelezionate.Add(row)
            End If
        Next

        If righeSelezionate.Count = 0 Then
            MsgBox("Nessuna riga selezionata.", MsgBoxStyle.Information)
            Return
        End If

        Dim connStr As String = Homepage.sap_tirelli
        Dim salvate As Integer = 0
        Dim errori As Integer = 0
        Dim nomeFile As String = If(String.IsNullOrEmpty(_pdfPath), "", Path.GetFileName(_pdfPath))
        Dim utente As String = Homepage.ID_SALVATO
        Dim utenteGalileo As String = Homepage.Trova_Dettagli_dipendente(Homepage.ID_SALVATO).Utente_Galileo

        Using conn As New SqlConnection(connStr)
            Try
                conn.Open()
            Catch ex As Exception
                MsgBox("Impossibile connettersi al database:" & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Errore connessione")
                Return
            End Try

            For Each row As DataGridViewRow In righeSelezionate
                Try
                    Dim ddt As String = CStr(row.Cells("colDDT").Value)
                    Dim dataStr As String = CStr(row.Cells("colData").Value)
                    Dim fornitore As String = CStr(row.Cells("colFornitore").Value)
                    Dim ordine As String = CStr(row.Cells("colOrdine").Value)
                    Dim codice As String = CStr(row.Cells("colCodice").Value)
                    Dim disegno As String = If(row.Cells("colDisegno").Value IsNot Nothing, CStr(row.Cells("colDisegno").Value), "")
                    Dim um As String = CStr(row.Cells("colUM").Value)
                    Dim qtaStr As String = CStr(row.Cells("colQuantita").Value)
                    Dim stato As String = If(row.Cells("colStato").Value IsNot Nothing, CStr(row.Cells("colStato").Value), "")

                    Dim ddtData As Object = DBNull.Value
                    Dim dtParsed As Date
                    If Date.TryParse(dataStr, dtParsed) Then ddtData = dtParsed

                    Dim qta As Object = DBNull.Value
                    Dim qtaDec As Decimal
                    If Decimal.TryParse(qtaStr, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, qtaDec) Then
                        qta = qtaDec
                    End If

                    Dim cmd As New SqlCommand(
                        "INSERT INTO [TIRELLI_40].[dbo].[Entrate_merci] (DDT_Numero, DDT_Data, Fornitore, Ordine_Acquisto, " &
                        "Codice_Articolo, Disegno, UM, Quantita, Stato, PDF_File, Data_Inserimento, Utente, Utente_Galileo) " &
                        "VALUES (@ddt, @data, @forn, @ord, @cod, @dis, @um, @qta, @stato, @file, GETDATE(), @utente, @utenteGalileo)",
                        conn)

                    cmd.Parameters.AddWithValue("@ddt", If(String.IsNullOrEmpty(ddt), DBNull.Value, CObj(ddt)))
                    cmd.Parameters.AddWithValue("@data", ddtData)
                    cmd.Parameters.AddWithValue("@forn", If(String.IsNullOrEmpty(fornitore), DBNull.Value, CObj(fornitore)))
                    cmd.Parameters.AddWithValue("@ord", If(String.IsNullOrEmpty(ordine), DBNull.Value, CObj(ordine)))
                    cmd.Parameters.AddWithValue("@cod", If(String.IsNullOrEmpty(codice), DBNull.Value, CObj(codice)))
                    cmd.Parameters.AddWithValue("@dis", If(String.IsNullOrEmpty(disegno), DBNull.Value, CObj(disegno)))
                    cmd.Parameters.AddWithValue("@um", If(String.IsNullOrEmpty(um), DBNull.Value, CObj(um)))
                    cmd.Parameters.AddWithValue("@qta", qta)
                    cmd.Parameters.AddWithValue("@stato", If(String.IsNullOrEmpty(stato), DBNull.Value, CObj(stato)))
                    cmd.Parameters.AddWithValue("@file", If(String.IsNullOrEmpty(nomeFile), DBNull.Value, CObj(nomeFile)))
                    cmd.Parameters.AddWithValue("@utente", If(String.IsNullOrEmpty(utente), DBNull.Value, CObj(utente)))
                    cmd.Parameters.AddWithValue("@utenteGalileo", If(String.IsNullOrEmpty(utenteGalileo), DBNull.Value, CObj(utenteGalileo)))

                    cmd.ExecuteNonQuery()
                    salvate += 1
                Catch ex As Exception
                    errori += 1
                End Try
            Next
        End Using

        Dim msg As String = $"{salvate} righe salvate in SQL."
        If errori > 0 Then msg &= $" {errori} errori."
        ImpostaMsg(msg, If(errori = 0, Color.DarkGreen, Color.OrangeRed))
        lblStatus.Text = msg
        MsgBox(msg, MsgBoxStyle.Information, "Salvataggio completato")
    End Sub

    ' ----------------------------------------------------------------
    ' CONFRONTO ORDINI APERTI AS400
    ' ----------------------------------------------------------------
    Private Sub VerificaOrdiniAS400()
        Dim dtOrdini As New DataTable()

        Try
            Using conn As New SqlConnection(Homepage.sap_tirelli)
                conn.Open()
                Dim query As String =
                    "SELECT * FROM OPENQUERY(AS400, 
'SELECT
    trim(codart) as codart,
    trim(disegno) as disegno,
    t0.numdoc as n_documento,
    t0.qta_ord as Q,
    data_richiesta ,
    evaso,
    trim(cod_forn) as cod_forn
 FROM TIR90VIS.JGALord t0
 WHERE  
   DOC = ''OA''
   and (evaso <> ''S''
   or data_richiesta >= 
       INTEGER(TO_CHAR(CURRENT DATE - 100 DAYS, ''YYYYMMDD'')))
')"
                Dim da As New SqlDataAdapter(query, conn)
                da.Fill(dtOrdini)
            End Using

        Catch ex As Exception
            lblStatus.Text = $"Riconoscimento completato. Confronto ordini non disponibile: {ex.Message}"
            ImpostaMsg($"{dgvRighe.Rows.Count} righe riconosciute (confronto ordini non disponibile).", Color.DarkGoldenrod)
            Return
        End Try

        Dim ok As Integer = 0
        Dim warn As Integer = 0
        Dim err As Integer = 0

        For Each row As DataGridViewRow In dgvRighe.Rows
            Dim ordineStr As String = If(row.Cells("colOrdine").Value IsNot Nothing, CStr(row.Cells("colOrdine").Value), "")
            Dim codice As String = If(row.Cells("colCodice").Value IsNot Nothing, CStr(row.Cells("colCodice").Value).Trim().ToUpper(), "")
            Dim qtaDDT As Decimal = 0
            Dim qtaStr As String = If(row.Cells("colQuantita").Value IsNot Nothing, CStr(row.Cells("colQuantita").Value), "")
            Decimal.TryParse(qtaStr, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, qtaDDT)

            Dim ordineNorm As String = NormalizzaNumero(ordineStr)

            Dim ordineFound As Boolean = False
            Dim codiceFound As Boolean = False
            Dim qtaOC As Decimal = 0
            Dim bestCodArtCanonical As String = ""
            Dim bestDisArtCanonical As String = ""
            Dim bestDelta As Decimal = Decimal.MaxValue

            For Each dr As DataRow In dtOrdini.Rows
                Dim nDoc As String = NormalizzaNumero(dr("n_documento").ToString())
                If ordineNorm <> "" AndAlso OrdiniCorrispondono(nDoc, ordineNorm) Then
                    ordineFound = True
                    Dim codArt As String = dr("codart").ToString().Trim().ToUpper()
                    Dim disArt As String = dr("disegno").ToString().Trim().ToUpper()
                    Dim codForn As String = dr("cod_forn").ToString().Trim().ToUpper()
                    If CodiciCorrispondono(codArt, codice, codForn) OrElse (disArt <> "" AndAlso CodiciCorrispondono(disArt, codice, codForn)) Then
                        ' Calcola la quantità AS400 per questa riga
                        Dim qRaw As String = dr("Q").ToString()
                        Dim qtaOC_A As Decimal = 0
                        Dim qtaOC_B As Decimal = 0
                        Dim qtaOC_C As Decimal = 0
                        Decimal.TryParse(qRaw.Replace(",", "."), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, qtaOC_A)
                        Decimal.TryParse(qRaw, Globalization.NumberStyles.Any, Globalization.CultureInfo.GetCultureInfo("it-IT"), qtaOC_B)
                        Decimal.TryParse(qRaw.Replace(".", "").Replace(",", ""), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, qtaOC_C)
                        Dim qtaCandidato As Decimal
                        If Math.Abs(qtaDDT - qtaOC_A) < 0.001D Then
                            qtaCandidato = qtaOC_A
                        ElseIf Math.Abs(qtaDDT - qtaOC_B) < 0.001D Then
                            qtaCandidato = qtaOC_B
                        ElseIf Math.Abs(qtaDDT - qtaOC_C) < 0.001D Then
                            qtaCandidato = qtaOC_C
                        Else
                            qtaCandidato = qtaOC_A
                        End If
                        ' Scegli la riga AS400 con quantità più vicina al DDT
                        Dim delta As Decimal = Math.Abs(qtaDDT - qtaCandidato)
                        If Not codiceFound OrElse delta < bestDelta Then
                            codiceFound = True
                            bestDelta = delta
                            qtaOC = qtaCandidato
                            bestCodArtCanonical = dr("codart").ToString().Trim()
                            bestDisArtCanonical = dr("disegno").ToString().Trim()
                        End If
                    End If
                End If
            Next

            ' Applica i valori canonici della riga AS400 selezionata
            If codiceFound Then
                row.Cells("colCodice").Value = bestCodArtCanonical
                row.Cells("colDisegno").Value = bestDisArtCanonical
            End If

            If Not ordineFound Then
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 180, 180)
                row.Cells("colStato").Value = "Ordine non trovato"
                err += 1
            ElseIf Not codiceFound Then
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 210, 140)
                row.Cells("colStato").Value = "Codice non in questo ordine"
                warn += 1
            ElseIf Math.Abs(qtaDDT - qtaOC) < 0.001D Then
                row.DefaultCellStyle.BackColor = Color.FromArgb(190, 235, 190)
                row.Cells("colStato").Value = "OK"
                ok += 1
            Else
                row.DefaultCellStyle.BackColor = Color.FromArgb(255, 255, 170)
                row.Cells("colStato").Value = $"Q DDT={qtaDDT:0.##}  OA={qtaOC:0.##}"
                warn += 1
            End If
        Next

        lblStatus.Text = $"Confronto completato: {ok} OK, {warn} scostamenti, {err} non trovati."
        ImpostaMsg($"{ok} OK  |  {warn} scostamenti  |  {err} non trovati", If(err > 0, Color.Red, If(warn > 0, Color.DarkGoldenrod, Color.DarkGreen)))
    End Sub

    Private Function NormalizzaNumero(s As String) As String
        s = s.Trim()
        Dim n As Long
        If Long.TryParse(s, n) Then Return n.ToString()
        Return s.ToUpper()
    End Function

    ''' <summary>
    ''' Confronta due numeri ordine: uguaglianza esatta oppure,
    ''' se hanno lunghezze diverse, il più lungo termina con il più corto
    ''' (es. "A0202211" corrisponde a "202211").
    ''' </summary>
    Private Function OrdiniCorrispondono(a As String, b As String) As Boolean
        If a = b Then Return True
        If a.Length > b.Length Then Return a.EndsWith(b)
        If b.Length > a.Length Then Return b.EndsWith(a)
        Return False
    End Function

    ''' <summary>
    ''' Confronta due codici articolo normalizzando "#N" come "-RN"
    ''' (es. "D115918#02" corrisponde a "D115918-R02").
    ''' </summary>
    Private Function CodiciCorrispondono(a As String, b As String, Optional codForn As String = "") As Boolean
        If a = b Then Return True
        Return NormalizzaCodice(a, codForn) = NormalizzaCodice(b, codForn)
    End Function

    Private Function NormalizzaCodice(s As String, Optional codForn As String = "") As String
        Dim result As String = s
        ' Regole specifiche del fornitore dal registro
        If codForn <> "" AndAlso RegoleFornitore.ContainsKey(codForn) Then
            Dim regole As String() = RegoleFornitore(codForn)
            Dim i As Integer = 0
            While i < regole.Length - 1
                result = System.Text.RegularExpressions.Regex.Replace(result, regole(i), regole(i + 1))
                i += 2
            End While
        End If
        Return result
    End Function

    ' ----------------------------------------------------------------
    ' CHIUDI
    ' ----------------------------------------------------------------
    Private Sub btnStorico_Click(sender As Object, e As EventArgs) Handles btnStorico.Click
        Dim f As New Entrate_merci_storico()
        f.ShowDialog(Me)
    End Sub

    Private Sub btnChiudi_Click(sender As Object, e As EventArgs) Handles btnChiudi.Click
        Me.Close()
    End Sub

    ' ----------------------------------------------------------------
    ' UTILITY
    ' ----------------------------------------------------------------
    Private Sub ImpostaControlliAbilitati(abilitati As Boolean)
        btnAnalizza.Enabled = abilitati
        btnSfoglia.Enabled = abilitati
        btnSalva.Enabled = abilitati
        btnSelTutto.Enabled = abilitati
        btnDeselTutto.Enabled = abilitati
        txtApiKey.Enabled = abilitati
    End Sub

    Private Sub ImpostaMsg(testo As String, colore As Color)
        lblMsg.Text = testo
        lblMsg.ForeColor = colore
    End Sub

    Private Function GetVal(dict As Dictionary(Of String, String), chiave As String) As String
        Dim v As String = ""
        dict.TryGetValue(chiave, v)
        Return If(v, "")
    End Function

    ' ----------------------------------------------------------------
    ' PARSER JSON MINIMALE (senza dipendenze esterne)
    ' Gestisce array di oggetti piatti {"key":"value",...}
    ' ----------------------------------------------------------------
    Private Function ParseJsonArray(json As String) As List(Of Dictionary(Of String, String))
        Dim risultato As New List(Of Dictionary(Of String, String))
        json = json.Trim()
        If Not json.StartsWith("[") Then Return risultato

        ' Separa gli oggetti
        Dim depth As Integer = 0
        Dim inString As Boolean = False
        Dim objStart As Integer = -1

        For i As Integer = 0 To json.Length - 1
            Dim c As Char = json(i)

            If c = "\"c AndAlso inString Then
                i += 1 ' skip escaped char
                Continue For
            End If

            If c = """"c Then
                inString = Not inString
                Continue For
            End If

            If inString Then Continue For

            If c = "{"c Then
                depth += 1
                If depth = 1 Then objStart = i
            ElseIf c = "}"c Then
                depth -= 1
                If depth = 0 AndAlso objStart >= 0 Then
                    Dim objStr As String = json.Substring(objStart, i - objStart + 1)
                    risultato.Add(ParseJsonObject(objStr))
                    objStart = -1
                End If
            End If
        Next

        Return risultato
    End Function

    Private Function ParseJsonObject(obj As String) As Dictionary(Of String, String)
        Dim dict As New Dictionary(Of String, String)
        obj = obj.Trim()
        If Not obj.StartsWith("{") Then Return dict

        ' Regex-free: cerca coppie "key":"value"
        Dim i As Integer = 1
        While i < obj.Length - 1
            ' Salta spazi/virgole
            If obj(i) = " "c OrElse obj(i) = ","c OrElse obj(i) = Chr(10) OrElse obj(i) = Chr(13) Then
                i += 1
                Continue While
            End If

            ' Legge chiave
            If obj(i) = """"c Then
                Dim key As String = LeggiStringaJson(obj, i)
                i += key.Length + 2 ' +2 per le virgolette
                ' Salta ':'
                While i < obj.Length AndAlso obj(i) <> ":"c
                    i += 1
                End While
                i += 1 ' salta ':'
                ' Salta spazi
                While i < obj.Length AndAlso obj(i) = " "c
                    i += 1
                End While
                ' Legge valore
                Dim value As String = ""
                If i < obj.Length AndAlso obj(i) = """"c Then
                    value = LeggiStringaJson(obj, i)
                    i += value.Length + 2
                ElseIf i < obj.Length Then
                    ' numero o null
                    Dim fine As Integer = i
                    While fine < obj.Length AndAlso obj(fine) <> ","c AndAlso obj(fine) <> "}"c
                        fine += 1
                    End While
                    value = obj.Substring(i, fine - i).Trim()
                    i = fine
                End If
                dict(key) = value
            Else
                i += 1
            End If
        End While

        Return dict
    End Function

    Private Function LeggiStringaJson(testo As String, inizio As Integer) As String
        ' inizio punta al " di apertura
        Dim sb As New StringBuilder()
        Dim i As Integer = inizio + 1
        While i < testo.Length
            Dim c As Char = testo(i)
            If c = "\"c AndAlso i + 1 < testo.Length Then
                Dim nextChar As Char = testo(i + 1)
                Select Case nextChar
                    Case """"c : sb.Append(""""c)
                    Case "\"c : sb.Append("\"c)
                    Case "n"c : sb.Append(Chr(10))
                    Case "r"c : sb.Append(Chr(13))
                    Case "t"c : sb.Append(Chr(9))
                    Case Else : sb.Append(nextChar)
                End Select
                i += 2
            ElseIf c = """"c Then
                Exit While
            Else
                sb.Append(c)
                i += 1
            End If
        End While
        Return sb.ToString()
    End Function

    Private Function EstraiCampoJson(json As String, campo As String) As String
        Dim chiave As String = """" & campo & """"
        Dim idx As Integer = json.IndexOf(chiave)
        If idx < 0 Then Return json
        Dim colon As Integer = json.IndexOf(":", idx)
        If colon < 0 Then Return json
        Dim valStart As Integer = json.IndexOf("""", colon)
        If valStart < 0 Then Return json
        Dim valEnd As Integer = json.IndexOf("""", valStart + 1)
        If valEnd < 0 Then Return json
        Return json.Substring(valStart + 1, valEnd - valStart - 1)
    End Function

End Class
