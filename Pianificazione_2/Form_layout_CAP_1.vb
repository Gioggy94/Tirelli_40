Imports System.Data.SqlClient
Imports Npgsql
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports Tirelli.ODP_Form

Public Class Form_layout_CAP_1

    ' ── Proprietà pubbliche ────────────────────────────────────────────────────
    Public Codice_commessa As String = ""
    Public Codice_risorsa_ As String = ""
    Public tipo_spostamento As String = "IN"
    Public parametro_trascinato As String = ""
    Public zona As String = "Officina"
    Public stato_kpi As Boolean = False

    ' ── Stato UI ───────────────────────────────────────────────────────────────
    Private coloreOriginaleGroupBoxes As New Dictionary(Of GroupBox, Color)
    Private dragRowIndex As Integer = -1
    Private sourceDataGridView As DataGridView = Nothing

    ' ── Costanti range GroupBox ────────────────────────────────────────────────
    Private Shared ReadOnly RangeOfficina As (Basso As Integer, Alto As Integer) = (1, 27)
    Private Shared ReadOnly RangeMagazzino As (Basso As Integer, Alto As Integer) = (52, 82)
    Private Shared ReadOnly RangeEsterno As (Basso As Integer, Alto As Integer) = (201, 230)
    Private Shared ReadOnly DeltaMagazzino As Integer = 24
    Private Shared ReadOnly DeltaEsterno As Integer = 138

#Region "Inizializzazione e Load"

    Private Sub Form_layout_CAP_1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Color.White
        PictureBox1.Image = Image.FromFile(Homepage.logo_azienda)
        compila_datagridview_macchine(DataGridView1)
        AbilitaDragDropZona("Officina")
    End Sub

#End Region

#Region "Helpers zona"

    ''' <summary>Restituisce il range (basso, alto) dei GroupBox per la zona.</summary>
    Private Function RangePerZona(par_zona As String) As (Basso As Integer, Alto As Integer)
        Select Case par_zona
            Case "Officina" : Return RangeOfficina
            Case "Magazzino" : Return RangeMagazzino
            Case "Esterno" : Return RangeEsterno
            Case Else : Return (0, 0)
        End Select
    End Function

    ''' <summary>Restituisce il delta da aggiungere al numero baia reale per trovare il GroupBox.</summary>
    Private Function DeltaPerZona(par_zona As String) As Integer
        Select Case par_zona
            Case "Magazzino" : Return DeltaMagazzino
            Case "Esterno" : Return DeltaEsterno
            Case Else : Return 0
        End Select
    End Function

    ''' <summary>Trova un GroupBox per nome nel form (ricerca ricorsiva).</summary>
    Private Function TrovaGroupBox(nome As String) As GroupBox
        Dim ctrl = Me.Controls.Find(nome, True).FirstOrDefault()
        Return If(TypeOf ctrl Is GroupBox, CType(ctrl, GroupBox), Nothing)
    End Function

#End Region

#Region "Drag & Drop – Registrazione handler"

    ''' <summary>
    ''' Registra gli handler DragEnter/DragDrop/DragLeave su tutti i GroupBox
    ''' validi per la zona indicata. Sostituisce EnableDragDropOnGroupBoxes +
    ''' EnableDragDropRecursively (che facevano la stessa cosa in modo ridondante).
    ''' </summary>
    Private Sub AbilitaDragDropZona(par_zona As String)
        Dim range = RangePerZona(par_zona)

        For i As Integer = range.Basso To range.Alto
            Dim grp As GroupBox = TrovaGroupBox("GroupBox" & i)
            If grp Is Nothing Then Continue For

            grp.AllowDrop = True
            RemoveHandler grp.DragEnter, AddressOf GroupBox_DragEnter
            RemoveHandler grp.DragDrop, AddressOf GroupBox_DragDrop
            RemoveHandler grp.DragLeave, AddressOf GroupBox_DragLeave
            AddHandler grp.DragEnter, AddressOf GroupBox_DragEnter
            AddHandler grp.DragDrop, AddressOf GroupBox_DragDrop
            AddHandler grp.DragLeave, AddressOf GroupBox_DragLeave
        Next
    End Sub

#End Region

#Region "Drag & Drop – Handler"

    Private Sub GroupBox_DragEnter(sender As Object, e As DragEventArgs)
        If e.Data.GetDataPresent(GetType(String)) AndAlso IsValidTarget(CType(sender, Control)) Then
            e.Effect = DragDropEffects.Copy
            Dim gb As GroupBox = CType(sender, GroupBox)
            If Not coloreOriginaleGroupBoxes.ContainsKey(gb) Then
                coloreOriginaleGroupBoxes(gb) = gb.BackColor
            End If
            gb.BackColor = Color.LightGreen
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub GroupBox_DragLeave(sender As Object, e As EventArgs)
        Dim gb As GroupBox = CType(sender, GroupBox)
        If coloreOriginaleGroupBoxes.ContainsKey(gb) Then
            gb.BackColor = coloreOriginaleGroupBoxes(gb)
            coloreOriginaleGroupBoxes.Remove(gb)
        End If
    End Sub

    Private Function IsValidTarget(ctrl As Control) As Boolean
        If Not TypeOf ctrl Is GroupBox OrElse Not ctrl.Name.StartsWith("GroupBox") Then
            Return False
        End If
        Dim num As Integer
        If Not Integer.TryParse(ctrl.Name.Replace("GroupBox", ""), num) Then Return False
        Return (num >= RangeOfficina.Basso AndAlso num <= RangeOfficina.Alto) OrElse
               (num >= RangeMagazzino.Basso AndAlso num <= RangeMagazzino.Alto) OrElse
               (num >= RangeEsterno.Basso AndAlso num <= RangeEsterno.Alto)
    End Function

    Private Sub GroupBox_DragDrop(sender As Object, e As DragEventArgs)
        Dim groupBox As GroupBox = CType(sender, GroupBox)

        ' Ripristina colore
        If coloreOriginaleGroupBoxes.ContainsKey(groupBox) Then
            groupBox.BackColor = coloreOriginaleGroupBoxes(groupBox)
            coloreOriginaleGroupBoxes.Remove(groupBox)
        End If

        ' ── 1. Dati di base ────────────────────────────────────────────────────
        Dim testo As String = CType(e.Data.GetData(GetType(String)), String)
        If String.IsNullOrEmpty(testo) Then Return

        Dim num As Integer
        If Not Integer.TryParse(groupBox.Name.Replace("GroupBox", ""), num) Then Return
        Dim numero_del_groupbox As Integer = num

        ' ── 2. Conversione GroupBox virtuale → numero baia reale ───────────────
        Dim numero_del_groupbox_mag As Integer = 0
        If numero_del_groupbox >= RangeMagazzino.Basso AndAlso numero_del_groupbox <= RangeMagazzino.Alto Then
            numero_del_groupbox_mag = numero_del_groupbox - DeltaMagazzino
        ElseIf numero_del_groupbox >= RangeEsterno.Basso AndAlso numero_del_groupbox <= RangeEsterno.Alto Then
            numero_del_groupbox_mag = numero_del_groupbox - DeltaEsterno
        End If

        ' ── 3. Verifica destinazione valida ────────────────────────────────────
        If Not ((numero_del_groupbox >= RangeOfficina.Basso AndAlso numero_del_groupbox <= RangeOfficina.Alto) OrElse
                numero_del_groupbox_mag <> 0) Then Return

        ' ── 4. Gestione per tipo ───────────────────────────────────────────────
        Select Case parametro_trascinato
            Case "RISORSA"
                ' (da implementare)
                Return

            Case "COMMESSA"
                GestisciDropCommessa(testo, groupBox, numero_del_groupbox, numero_del_groupbox_mag)
        End Select
    End Sub

    ''' <summary>Tutta la logica di drop per una COMMESSA, estratta per leggibilità.</summary>
    Private Sub GestisciDropCommessa(testo As String, groupBox As GroupBox,
                                     numero_del_groupbox As Integer,
                                     numero_del_groupbox_mag As Integer)

        ' ── 4a. Dati commessa ──────────────────────────────────────────────────
        Dim info = check_baia_layout(testo)
        Dim numero_baia_partenza As Integer = info.numero_baia
        Dim zonaPartenza As String = info.zona_layout
        Dim statoCommessa As String = info.Stato

        ' ── 4b. Numero baia reale destinazione ────────────────────────────────
        Dim numero_def_baia As Integer =
            If(numero_del_groupbox_mag <> 0, numero_del_groupbox_mag, numero_del_groupbox)
        Dim zonaArrivo As String = zona_della_baia(numero_def_baia)

        ' ── 4c. Blocco: commessa "O" senza baia (inserimento ex-novo vietato) ──
        If statoCommessa = "O" AndAlso numero_baia_partenza = 0 Then
            MsgBox("La commessa " & testo & " è già assegnata con stato attivo (O)." &
                   vbCrLf & "Rimuoverla prima di inserirla nuovamente.",
                   MsgBoxStyle.Exclamation, "Inserimento non consentito")
            Return
        End If

        ' ── 4d. Blocco: baia di destinazione già occupata ─────────────────────
        If numero_baia_partenza > 0 AndAlso
           tipo_spostamento = "IN" AndAlso
           (zonaPartenza = "Officina" OrElse zonaPartenza = zonaArrivo) AndAlso
           statoCommessa <> "P" Then

            Dim nomeGbPartenza As String = "GroupBox" & (numero_baia_partenza + DeltaPerZona(zonaPartenza))
            Dim grpPartenza As GroupBox = TrovaGroupBox(nomeGbPartenza)
            Dim nomeBaia As String = If(grpPartenza IsNot Nothing, grpPartenza.Text, numero_baia_partenza.ToString())
            MsgBox("Macchina già presente nella baia " & nomeBaia, MsgBoxStyle.Information, "Attenzione")
            Return
        End If

        ' ── 4e. Blocco: commessa già nella baia di destinazione ───────────────
        If numero_baia_partenza = numero_def_baia AndAlso
           zonaPartenza = "Officina" AndAlso
           statoCommessa <> "P" Then
            MsgBox("Macchina già presente nella baia " & info.Nome_baia, MsgBoxStyle.Information, "Attenzione")
            Return
        End If

        ' ── 4f. Conferma spostamento ──────────────────────────────────────────
        If MessageBox.Show("Vuoi spostare " & testo & " nella baia " & groupBox.Text,
                           "Conferma spostamento",
                           MessageBoxButtons.YesNo,
                           MessageBoxIcon.Question) = DialogResult.No Then Return

        ' ── 4g. Check macchina già spedita ────────────────────────────────────
        Dim avviso As String = check_macchina_già_spedita(testo)
        If avviso <> "" Then
            If MsgBox(avviso & vbCrLf & "Vuoi proseguire comunque?",
                      vbYesNo + vbQuestion, "Conferma") = vbNo Then Return
        End If

        ' ── 4h. Esecuzione spostamento ────────────────────────────────────────
        Dim statoNuovo As String = If(zonaArrivo = "Magazzino", "O", "P")
        inserisci_record_baia(testo, numero_def_baia, statoNuovo)
        inserisci_record_baia_log(testo, numero_def_baia, tipo_spostamento)
        check_presenza_commessa_baia_layout(numero_def_baia, zona)
    End Sub

#End Region

#Region "DataGridView – Drag & Drop"

    Private Sub DataGridView_MouseDown(sender As Object, e As MouseEventArgs) _
        Handles DataGridView1.MouseDown

        tipo_spostamento = "IN"
        Dim dgv = DirectCast(sender, DataGridView)
        Dim hit = dgv.HitTest(e.X, e.Y)

        If hit.Type <> DataGridViewHitTestType.Cell OrElse hit.RowIndex < 0 Then Return

        dragRowIndex = hit.RowIndex
        sourceDataGridView = dgv

        If dgv Is DataGridView1 Then
            Dim valore As String = dgv.Rows(dragRowIndex).Cells("commessa").Value?.ToString()
            If Not String.IsNullOrEmpty(valore) Then
                parametro_trascinato = "COMMESSA"
                dgv.DoDragDrop(valore, DragDropEffects.Copy)
            End If
        End If
    End Sub

    Private Sub DataGridView_DragOver(sender As Object, e As DragEventArgs) _
        Handles DataGridView1.DragOver
        e.Effect = DragDropEffects.Copy
    End Sub

    Private Sub DataGridView_DragDrop(sender As Object, e As DragEventArgs) _
        Handles DataGridView1.DragDrop

        Dim targetDGV = DirectCast(sender, DataGridView)
        Dim hit = targetDGV.HitTest(
            targetDGV.PointToClient(New Point(e.X, e.Y)).X,
            targetDGV.PointToClient(New Point(e.X, e.Y)).Y)

        If hit.Type = DataGridViewHitTestType.Cell AndAlso hit.RowIndex >= 0 Then
            If targetDGV IsNot sourceDataGridView OrElse hit.RowIndex <> dragRowIndex Then
                MsgBox("Errore, il comando è stato annullato")
            End If
        End If
    End Sub

#End Region

#Region "Layout – check presenza e pulizia GroupBox"

    Public Sub check_presenza_commessa_baia_layout(par_baia As Integer, par_zona As String)
        Dim range = RangePerZona(par_zona)
        Dim delta = DeltaPerZona(par_zona)

        ' Pulizia GroupBox
        If par_baia = 0 Then
            For i As Integer = range.Basso To range.Alto
                PulisciGroupBox_PerNome("GroupBox" & i)
            Next
        Else
            PulisciGroupBox(par_baia, par_zona)
        End If

        Dim filtro_baia As String = If(par_baia > 0, " AND t0.baia = " & par_baia, "")

        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "SELECT t0.baia 
                 FROM [tirelli_40].[dbo].[Layout_CAP1] t0
                 LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1_nomi] t1 ON t1.numero_baia = t0.baia
                 WHERE (t0.stato='O' OR t0.stato='P')" & filtro_baia &
                " AND t1.zona = @zona
                 GROUP BY t0.baia ORDER BY t0.baia", CNN)

                CMD.Parameters.AddWithValue("@zona", par_zona)
                Using reader As SqlDataReader = CMD.ExecuteReader()
                    Dim baie As New List(Of Integer)
                    Do While reader.Read()
                        baie.Add(CInt(reader("baia")))
                    Loop
                    reader.Close()

                    For Each baiaCorrente As Integer In baie
                        Dim grp As GroupBox = TrovaGroupBox("GroupBox" & (baiaCorrente + delta))
                        If grp Is Nothing Then Continue For

                        ' Rimuovi panel esistente
                        For Each p In grp.Controls.OfType(Of Panel)().ToList()
                            grp.Controls.Remove(p)
                            p.Dispose()
                        Next

                        Dim pnl As New FlowLayoutPanel() With {
                            .FlowDirection = FlowDirection.TopDown,
                            .WrapContents = False,
                            .AutoScroll = True,
                            .AutoScrollMargin = New Size(0, 0)
                        }
                        pnl.HorizontalScroll.Maximum = 0
                        pnl.HorizontalScroll.Visible = False

                        crea_panel_nel_groupbox(pnl, grp, baiaCorrente)
                        crea_moduli_nel_panel(pnl, baiaCorrente, par_zona)
                    Next
                End Using
            End Using
        End Using

        compila_datagridview_macchine(DataGridView1)
        compila_risorse()
    End Sub

    Private Sub PulisciGroupBox(numero As Integer, par_zona As String)
        PulisciGroupBox_PerNome("GroupBox" & (numero + DeltaPerZona(par_zona)))
    End Sub

    Private Sub PulisciGroupBox_PerNome(nomeGroupBox As String)
        Dim grp As GroupBox = TrovaGroupBox(nomeGroupBox)
        If grp Is Nothing Then Return
        For Each p In grp.Controls.OfType(Of Panel)().ToList()
            grp.Controls.Remove(p)
            p.Dispose()
        Next
    End Sub

#End Region

#Region "Layout – creazione panel e moduli"

    Sub crea_panel_nel_groupbox(pnl As Panel, par_groupbox As GroupBox, baia_corrente As Integer)
        pnl.Name = "Panel" & baia_corrente
        pnl.Dock = DockStyle.Fill
        pnl.AutoScroll = True
        par_groupbox.Controls.Add(pnl)
    End Sub

    Sub crea_moduli_nel_panel(pnl As Panel, baiacorrente As Integer, par_zona As String)
        ' Recupera tutte le commesse della baia in un'unica query con JOIN
        ' (elimina le N connessioni annidate del codice originale)
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "SELECT t0.commessa,
                        COALESCE(T10.ITEMNAME,'') AS itemname,
                        SUBSTRING(COALESCE(T3.CARDNAME, COALESCE(T2.CARDNAME, T10.U_FINAL_CUSTOMER_NAME)),1,25) AS CLIENTE,
                        COALESCE(t2.u_destinazione, t10.u_country_of_delivery) AS u_destinazione,
                        COALESCE(t10.u_brand,'') AS Brand
                 FROM [tirelli_40].[dbo].[Layout_CAP1] t0
                 LEFT JOIN TIRELLISRLDB.DBO.OITM T10 ON t0.commessa = T10.itemcode
                 LEFT JOIN (
                     SELECT T1.ITEMCODE, MAX(T1.DOCENTRY) AS DOCENTRY
                     FROM TIRELLISRLDB.DBO.RDR1 T1
                     GROUP BY T1.ITEMCODE
                 ) A ON A.ITEMCODE = t0.commessa
                 LEFT JOIN TIRELLISRLDB.DBO.ORDR T2 ON T2.DOCENTRY = A.DOCENTRY
                 LEFT JOIN TIRELLISRLDB.DBO.OCRD T3 ON T3.CARDCODE = T2.U_CodiceBP
                 LEFT JOIN (
                     SELECT commessa, MAX(numero) AS numero
                     FROM [TIRELLI_40].[dbo].[Scheda_tecnica_revisioni]
                     GROUP BY commessa
                 ) B ON B.commessa = t0.commessa
                 LEFT JOIN [TIRELLI_40].[dbo].[Scheda_Tecnica_valori] t4
                     ON t4.rev = B.numero AND t4.commessa = t0.commessa
                 WHERE (t0.stato='O' OR t0.stato='P') AND t0.baia = @baia
                 ORDER BY t0.commessa DESC", CNN)

                CMD.Parameters.AddWithValue("@baia", baiacorrente)

                Using reader As SqlDataReader = CMD.ExecuteReader()
                    Do While reader.Read()
                        Dim commessa As String = reader("commessa").ToString()
                        Dim destinazione As String = reader("u_destinazione").ToString()
                        Dim brand As String = reader("Brand").ToString()

                        Dim flagImage As Bitmap = Nothing
                        Try
                            flagImage = New Bitmap("\\tirfs01\00-Tirelli 4.0\Immagini\Flags\" & destinazione & ".png")
                        Catch
                        End Try

                        If par_zona = "Officina" Then
                            Dim modulo As New Modulo_baia()
                            modulo.Titolo = commessa
                            modulo.numero_baia = baiacorrente
                            modulo.PictureBox1.Image = flagImage
                            modulo.Label2.Text = GetFirstTwoWords(reader("itemname").ToString())
                            modulo.Label3.Text = GetFirstTwoWords(reader("CLIENTE").ToString())
                            modulo.Label1.ForeColor = ColorePerBrand(brand)
                            pnl.Controls.Add(modulo)
                            pnl.Controls.SetChildIndex(modulo, 0)
                            modulo.inizializza_modulo(commessa, pnl)

                        ElseIf par_zona = "Magazzino" OrElse par_zona = "Esterno" Then
                            Dim modulo As New Modulo_mag()
                            modulo.Label1.Text = commessa
                            modulo.numero_baia = baiacorrente
                            modulo.PictureBox1.Image = flagImage
                            modulo.Label2.Text = GetFirstTwoWords(reader("itemname").ToString())
                            modulo.Label3.Text = GetFirstTwoWords(reader("CLIENTE").ToString())
                            modulo.Label1.ForeColor = ColorePerBrand(brand, True)
                            pnl.Controls.Add(modulo)
                            pnl.Controls.SetChildIndex(modulo, 0)
                            modulo.inizializza_modulo(commessa, pnl)
                        End If
                    Loop
                End Using
            End Using
        End Using
    End Sub

    ''' <summary>Restituisce il colore del brand. isMag=True usa Gold invece di Blue per BRB.</summary>
    Private Function ColorePerBrand(brand As String, Optional isMag As Boolean = False) As Color
        Select Case brand.ToUpper()
            Case "TIRELLI" : Return Color.Blue
            Case "KTF" : Return Color.DarkGreen
            Case "BRB" : Return If(isMag, Color.Gold, Color.Gold)
            Case Else : Return SystemColors.ControlText
        End Select
    End Function

#End Region

#Region "Database – Baia"

    Public Sub inserisci_record_baia(par_commessa As String, par_baia As Integer, par_stato As String)
        Dim baia_precedente As Integer = par_baia
        Dim zona_baia As String = zona_della_baia(par_baia)

        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()

            ' Leggi baia precedente
            Using CMD As New SqlCommand(
                "SELECT TOP 1 t0.Baia, t1.zona
                 FROM [tirelli_40].[dbo].[Layout_CAP1] t0
                 LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1_nomi] t1 ON t0.baia = t1.numero_baia
                 WHERE t0.Commessa = @Commessa", CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                Using reader = CMD.ExecuteReader()
                    If reader.Read() Then
                        baia_precedente = CInt(reader("Baia"))
                        zona_baia = If(reader("zona") IsNot DBNull.Value, reader("zona").ToString(), zona_baia)
                    End If
                End Using
            End Using

            ' Delete + Insert nella stessa zona
            Using CMD As New SqlCommand(
                "DELETE t0
                 FROM [tirelli_40].[dbo].[Layout_CAP1] t0
                 LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1_nomi] t1 ON t1.numero_baia = t0.baia
                 WHERE t0.Commessa = @Commessa AND t1.zona = @zona;

                 INSERT INTO [tirelli_40].[dbo].[Layout_CAP1] ([Commessa],[Baia],[Stato])
                 VALUES (@Commessa, @Baia, @Stato);", CNN)

                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                CMD.Parameters.AddWithValue("@Baia", par_baia)
                CMD.Parameters.AddWithValue("@Stato", par_stato)
                CMD.Parameters.AddWithValue("@zona", zona_della_baia(par_baia))
                CMD.ExecuteNonQuery()
            End Using
        End Using

        check_presenza_commessa_baia_layout(baia_precedente, zona)
    End Sub

    Public Sub cancella_record_baia(par_commessa As String, par_baia As Integer, par_zona As String)
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
            "DELETE t0 FROM [tirelli_40].[dbo].[Layout_CAP1] t0
             LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1_nomi] t1 ON t0.baia = t1.numero_baia
             WHERE t0.Commessa = @Commessa AND t0.baia = @Baia AND t1.zona = @Zona", CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                CMD.Parameters.AddWithValue("@Baia", par_baia)
                CMD.Parameters.AddWithValue("@Zona", par_zona)
                CMD.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub inserisci_record_baia_log(par_commessa As String, par_baia As Integer, par_tipo_mov As String)
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "INSERT INTO [tirelli_40].[dbo].[Layout_CAP1_LOG]
                 ([Commessa],[Baia],[Tipo_mov],[data],[utente])
                 VALUES (@Commessa, @Baia, @TipoMov, GETDATE(), @Utente)", CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                CMD.Parameters.AddWithValue("@Baia", par_baia)
                CMD.Parameters.AddWithValue("@TipoMov", par_tipo_mov)
                CMD.Parameters.AddWithValue("@Utente", Homepage.ID_SALVATO)
                CMD.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub cancella_record_baia_log(par_commessa As String, par_zona As String)
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "DELETE t0 FROM [tirelli_40].[dbo].[Layout_CAP1_log] t0
                 LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1_nomi] t1 ON t0.baia = t1.numero_baia
                 WHERE t0.Commessa = @Commessa AND t1.zona = @zona", CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                CMD.Parameters.AddWithValue("@zona", par_zona)
                CMD.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Public Sub cancella_record_baia_log_per_baia(par_commessa As String, par_baia As Integer)
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
            "DELETE FROM [tirelli_40].[dbo].[Layout_CAP1_log]
             WHERE Commessa = @Commessa AND Baia = @Baia", CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                CMD.Parameters.AddWithValue("@Baia", par_baia)
                CMD.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub inserisci_record_baia_spedizione(par_commessa As String, par_baia As Integer, par_tipo_mov As String)
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "INSERT INTO [tirelli_40].[dbo].[Layout_CAP1_spedite]
                 ([Commessa],[Baia],[data],[utente])
                 VALUES (@Commessa, @Baia, GETDATE(), @Utente)", CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                CMD.Parameters.AddWithValue("@Baia", par_baia)
                CMD.Parameters.AddWithValue("@Utente", Homepage.ID_SALVATO)
                CMD.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub aggiorna_stato_baia(par_commessa As String, par_stato As String,
                                   par_zona As String, par_numero_baia As Integer)
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "UPDATE [tirelli_40].[dbo].[Layout_CAP1]
                 SET stato = @Stato
                 WHERE commessa = @Commessa AND baia = @Baia;

                 UPDATE t0
                 SET t0.data = GETDATE()
                 FROM [tirelli_40].[dbo].[Layout_CAP1_LOG] t0
                 LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1_nomi] t1 ON t1.numero_baia = t0.baia
                 WHERE t0.commessa = @Commessa AND t0.tipo_mov = 'IN' AND t1.zona = @zona", CNN)
                CMD.Parameters.AddWithValue("@Stato", par_stato)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                CMD.Parameters.AddWithValue("@Baia", par_numero_baia)
                CMD.Parameters.AddWithValue("@zona", par_zona)
                CMD.ExecuteNonQuery()
            End Using
        End Using
    End Sub

#End Region

#Region "Database – Check e query"

    Public Function check_baia_layout(par_commessa As String) As Dettagli_commessa
        Dim d As New Dettagli_commessa()
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "SELECT TOP 1 t0.baia, t1.nome_baia, t1.zona, t0.STATO
                 FROM [tirelli_40].[dbo].[Layout_CAP1] t0
                 LEFT JOIN [tirelli_40].[dbo].[Layout_CAP1_nomi] t1 ON t1.numero_baia = t0.baia
                 WHERE t0.commessa = @Commessa
                 ORDER BY t0.stato", CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                Using reader = CMD.ExecuteReader()
                    If reader.Read() Then
                        d.numero_baia = CInt(reader("baia"))
                        d.zona_layout = reader("zona").ToString()
                        d.Nome_baia = reader("nome_baia").ToString()
                        d.Stato = reader("Stato").ToString()
                    End If
                End Using
            End Using
        End Using
        Return d
    End Function
    Public Function check_baia_layout_aperte(par_commessa As String) As Dettagli_commesse_aperte
        Dim d As New Dettagli_commesse_aperte()
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
            "SELECT TOP 1 t0.baia, t1.nome_baia, t1.zona, t0.STATO
             FROM [tirelli_40].[dbo].[Layout_CAP1] t0
             LEFT JOIN [tirelli_40].[dbo].[Layout_CAP1_nomi] t1 ON t1.numero_baia = t0.baia
             WHERE t0.commessa = @Commessa AND t0.stato = 'O'
             ORDER BY t0.stato", CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                Using reader = CMD.ExecuteReader()
                    If reader.Read() Then
                        d.numero_baia = CInt(reader("baia"))
                        d.zona_layout = reader("zona").ToString()
                        d.Nome_baia = reader("nome_baia").ToString()
                        d.Stato = reader("Stato").ToString()
                    End If
                End Using
            End Using
        End Using
        Return d
    End Function

    Public Function check_baia_layout_A_numero_baia(par_commessa As String, par_numero_baia As String) As Dettagli_commesse_A_numero_baia
        Dim d As New Dettagli_commesse_A_numero_baia()
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
            "SELECT TOP 1 t0.baia, t1.nome_baia, t1.zona, t0.STATO
             FROM [tirelli_40].[dbo].[Layout_CAP1] t0
             LEFT JOIN [tirelli_40].[dbo].[Layout_CAP1_nomi] t1 ON t1.numero_baia = t0.baia
             WHERE t0.commessa = @Commessa AND t0.baia = @Baia
             ORDER BY t0.stato", CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                CMD.Parameters.AddWithValue("@Baia", par_numero_baia)
                Using reader = CMD.ExecuteReader()
                    If reader.Read() Then
                        d.numero_baia = CInt(reader("baia"))
                        d.zona_layout = reader("zona").ToString()
                        d.Nome_baia = reader("nome_baia").ToString()
                        d.Stato = reader("Stato").ToString()
                    End If
                End Using
            End Using
        End Using
        Return d
    End Function

    Public Function check_macchina_già_spedita(par_commessa As String) As String
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "SELECT TOP 1 [data]
                 FROM [Tirelli_40].[dbo].[Layout_CAP1_SPEDITE]
                 WHERE commessa = @Commessa", CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                Using reader = CMD.ExecuteReader()
                    If reader.Read() Then
                        Return "La macchina risulta già spedita il " & reader("data").ToString()
                    End If
                End Using
            End Using
        End Using
        Return ""
    End Function

    Public Function zona_della_baia(par_numero_baia As Integer) As String
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "SELECT zona FROM [tirelli_40].[dbo].[Layout_CAP1_nomi]
                 WHERE numero_baia = @Baia", CNN)
                CMD.Parameters.AddWithValue("@Baia", par_numero_baia)
                Dim result = CMD.ExecuteScalar()
                Return If(result IsNot Nothing, result.ToString(), "")
            End Using
        End Using
    End Function

    Public Function ingresso_commessa(par_commessa As String, par_zona As String) As Date
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "SELECT TOP 1 [data]
                 FROM [Tirelli_40].[dbo].[Layout_CAP1_LOG] t0
                 LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1_nomi] t1 ON t0.Baia = t1.numero_baia
                 WHERE t0.Commessa = @Commessa AND t0.tipo_mov = 'IN'
                 ORDER BY t0.id DESC", CNN)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                Using reader = CMD.ExecuteReader()
                    If reader.Read() Then Return CDate(reader("data"))
                End Using
            End Using
        End Using
        Return Nothing
    End Function

#End Region

#Region "Database – Risorse"

    Public Sub inserisci_record_risorsa(par_risorsa As String, par_commessa As String)
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "DELETE FROM [tirelli_40].[dbo].[Layout_CAP1_risorse] WHERE Risorsa = @Risorsa;
                 INSERT INTO [tirelli_40].[dbo].[Layout_CAP1_risorse] ([Commessa],[Risorsa])
                 VALUES (@Commessa, @Risorsa)", CNN)
                CMD.Parameters.AddWithValue("@Risorsa", par_risorsa)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                CMD.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Sub Cancella_record(par_risorsa As String)
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "DELETE FROM [tirelli_40].[dbo].[Layout_CAP1_risorse] WHERE Risorsa = @Risorsa", CNN)
                CMD.Parameters.AddWithValue("@Risorsa", par_risorsa)
                CMD.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Function carica_impieghi_risorse() As Dictionary(Of String, (Commessa As String, NomeBaia As String, Cliente As String))
        Dim risultati As New Dictionary(Of String, (String, String, String))(StringComparer.OrdinalIgnoreCase)
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "SELECT t0.Risorsa, t0.Commessa,
                        COALESCE(t2.nome_baia,'') AS Nome_baia,
                        COALESCE(T3.U_FINAL_CUSTOMER_NAME,'') AS Cliente
                 FROM [Tirelli_40].[dbo].[Layout_CAP1_risorse] t0
                 LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1] t1 ON t0.commessa = t1.commessa
                 LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1_nomi] t2 ON t2.numero_baia = t1.baia
                 LEFT JOIN TIRELLISRLDB.DBO.oitm t3 ON t3.itemcode = t0.commessa", CNN)
                Using reader = CMD.ExecuteReader()
                    Do While reader.Read()
                        Dim risorsa As String = reader("Risorsa").ToString()
                        If Not risultati.ContainsKey(risorsa) Then
                            risultati.Add(risorsa, (reader("Commessa").ToString(),
                                                    reader("Nome_baia").ToString(),
                                                    reader("Cliente").ToString()))
                        End If
                    Loop
                End Using
            End Using
        End Using
        Return risultati
    End Function

    Public Sub lista_risorse(par_datagridview As DataGridView, par_nome As String, par_reparto As String)
        par_datagridview.Rows.Clear()
        Dim impieghi = carica_impieghi_risorse()

        Dim filtroNome As String = If(par_nome <> "", " AND LOWER(res.resdsc) LIKE '%" & par_nome.ToLower() & "%'", "")
        Dim filtroReparto As String = If(par_reparto <> "", " AND LOWER(gl.grpdsc) LIKE '%" & par_reparto.ToLower() & "%'", "")

        Dim contatori As New Dictionary(Of String, Integer) From {
            {"ferie", 0}, {"malattia", 0}, {"altro", 0},
            {"mecc_TIR", 0}, {"mecc_KTF", 0}, {"mecc_BRB", 0},
            {"elettrico", 0}, {"collaudatori", 0}
        }

        Using conn As New NpgsqlConnection(Homepage.JPM_TIRELLI)
            conn.Open()
            Using CMD As New NpgsqlCommand(
                "SELECT res.uid AS resuid, res.rescod AS codice_risorsa,
                        res.resdsc AS descrizione_risorsa, res.dteval,
                        COALESCE(grp.grpcod,'') AS codice_gruppo,
                        gl.grpdsc AS descr_gruppo,
                        lvl.lvlcod AS codice_org, lvl.lvldsc AS descr_org
                 FROM angres res
                 LEFT JOIN angresgrp rg ON res.uid = rg.resuid AND rg.prjgrppri = -1
                 LEFT JOIN anggrp grp ON grp.uid = rg.grpuid
                 LEFT JOIN anggrplng gl ON grp.uid = gl.recuid AND gl.lnguid = 1
                 LEFT JOIN orglvlres ol ON ol.resuid = res.uid
                 LEFT JOIN orglvl lvl ON lvl.uid = ol.lvluid
                 WHERE res.logdel = 0 AND COALESCE(grp.grpcod,'') <> ''
                 AND (res.dteval IS NULL OR res.dteval > NOW())" &
                filtroNome & filtroReparto & "
                 ORDER BY gl.grpdsc, res.resdsc", conn)

                Using reader = CMD.ExecuteReader()
                    Do While reader.Read()
                        Dim uid As String = reader("resuid").ToString()
                        Dim commessa As String = ""
                        Dim nomeBaia As String = ""
                        Dim cliente As String = ""

                        If impieghi.ContainsKey(uid) Then
                            Dim d = impieghi(uid)
                            commessa = d.Commessa
                            nomeBaia = d.NomeBaia
                            cliente = d.Cliente
                        End If

                        par_datagridview.Rows.Add(uid, reader("codice_risorsa"), reader("descrizione_risorsa"),
                                                  reader("descr_gruppo"), commessa, cliente, nomeBaia)

                        ' Contatori speciali
                        If commessa = GroupBox39.Text Then contatori("ferie") += 1
                        If commessa = GroupBox40.Text Then contatori("malattia") += 1
                        If commessa = GroupBox41.Text Then contatori("altro") += 1

                        If commessa = "" Then
                            Select Case reader("codice_gruppo").ToString().ToUpper()
                                Case "MONT MECC TIR" : contatori("mecc_TIR") += 1
                                Case "MONT MECC KTF" : contatori("mecc_KTF") += 1
                                Case "MONT MECC BRB" : contatori("mecc_BRB") += 1
                                Case "ELETTRICO" : contatori("elettrico") += 1
                                Case "COLL TIR" : contatori("collaudatori") += 1
                            End Select
                        End If
                    Loop
                End Using
            End Using
        End Using

        Label3.Text = contatori("ferie")
        Label4.Text = contatori("malattia")
        Label5.Text = contatori("altro")
        Label6.Text = contatori("mecc_TIR")
        Label9.Text = contatori("mecc_KTF")
        Label10.Text = contatori("mecc_BRB")
        Label11.Text = contatori("mecc_TIR") + contatori("mecc_KTF") + contatori("mecc_BRB")
        Label7.Text = contatori("elettrico")
        Label8.Text = contatori("collaudatori")

        par_datagridview.ClearSelection()
    End Sub

#End Region

#Region "Database – Galileo / Spedizioni"

    Public Sub copia_dati_T40_in_Galileo()
        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "DELETE [AS400].[S786FAD1].[TIR90VIS].[YANCOM0F];
                 INSERT INTO [AS400].[S786FAD1].[TIR90VIS].[YANCOM0F]
                 (YIDRCCA,YTPELCA,YDTINCA,YCDCOCA,YCDSCCA,YCDARCA,YCLA3CA,YBAIACA)
                 SELECT t0.ID, 1, GETDATE(), '', '', t0.Commessa,
                        RIGHT('00' + CAST(t0.Baia AS VARCHAR(2)), 2),
                        t1.Nome_Baia
                 FROM [Tirelli_40].[dbo].[Layout_CAP1] t0
                 LEFT JOIN [Tirelli_40].[dbo].[Layout_CAP1_nomi] t1 ON t0.Baia = t1.numero_baia
                 WHERE stato <> 'P'", CNN)
                CMD.ExecuteNonQuery()
            End Using
        End Using
    End Sub

#End Region

#Region "JPM – Date e query"

    Public Sub date_jpm_commessa(par_commessa As String, par_datagridview As DataGridView)
        par_datagridview.Rows.Clear()
        Using conn As New NpgsqlConnection(Homepage.JPM_TIRELLI)
            conn.Open()
            Using CMD As New NpgsqlCommand(
                "SELECT W.WBSLVLCOD AS ATT_WBS, L.TSKDSC AS ATT_DES,
                        T.TSKRSDTSSSTR AS ATT_DTPIAINI, T.TSKRSDTSSEND AS ATT_DTPIAFIN,
                        T.TSKORGRSDEFR AS DURATAORI
                 FROM PRJTSK T
                 LEFT JOIN PRJTSKDET TD ON TD.TSKUID = T.UID
                 LEFT JOIN PRJTSKLNG L ON L.RECUID = T.UID AND L.LNGUID = 1
                 LEFT JOIN ANGWBSLVL W ON T.WBSLVLUID = W.UID
                 LEFT JOIN PRJ P ON T.PRJUID = P.UID
                 WHERE T.LOGDEL = 0 AND P.PRJcod = @Commessa AND T.TSKKND = '1'", conn)
                CMD.Parameters.AddWithValue("@Commessa", par_commessa)
                Using reader = CMD.ExecuteReader()
                    Do While reader.Read()
                        par_datagridview.Rows.Add(reader("ATT_WBS"), reader("ATT_DES"),
                                                  reader("ATT_DTPIAINI"), reader("ATT_DTPIAFIN"),
                                                  reader("DURATAORI"))
                    Loop
                End Using
            End Using
        End Using
    End Sub

#End Region

#Region "Utilità"

    Private Function GetFirstTwoWords(input As String) As String
        Dim words = input.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
        Return If(words.Length >= 2, words(0) & " " & words(1), If(words.Length = 1, words(0), ""))
    End Function

    Public Function giorni_lavorativi_tra(dataInizio As Date, dataFine As Date) As Integer
        Dim giorni As Integer = 0
        Dim giorno As Date = dataInizio
        Do While giorno < dataFine
            If giorno.DayOfWeek <> DayOfWeek.Saturday AndAlso giorno.DayOfWeek <> DayOfWeek.Sunday Then
                giorni += 1
            End If
            giorno = giorno.AddDays(1)
        Loop
        Return giorni
    End Function

    Sub nomina_baie(par_zona As String)
        Dim delta = DeltaPerZona(par_zona)
        Dim range = RangePerZona(par_zona)
        Dim i As Integer = range.Basso

        Using CNN As New SqlConnection(Homepage.sap_tirelli)
            CNN.Open()
            Using CMD As New SqlCommand(
                "SELECT numero_baia, Nome_Baia
                 FROM [Tirelli_40].[dbo].[Layout_CAP1_nomi]
                 WHERE Zona = @zona ORDER BY numero_baia", CNN)
                CMD.Parameters.AddWithValue("@zona", par_zona)
                Using reader = CMD.ExecuteReader()
                    Do While reader.Read()
                        Dim grp As GroupBox = TrovaGroupBox("GroupBox" & i)
                        If grp IsNot Nothing Then grp.Text = reader("Nome_Baia").ToString()
                        i += 1
                    Loop
                End Using
            End Using
        End Using
    End Sub

    Sub compila_datagridview_macchine(par_datagridview As DataGridView)
        Scheda_commessa_Pianificazione.carica_commesse(DataGridView1, TextBox1.Text, TextBox2.Text,
                                                       TextBox3.Text, TextBox13.Text, "", "", "", "",
                                                       TextBox15.Text)
    End Sub

    Sub compila_risorse()
        lista_risorse(DataGridView3, TextBox5.Text, TextBox7.Text)
    End Sub

#End Region

#Region "Classi dati"

    Public Class Dettagli_commessa
        Public numero_baia As Integer
        Public zona_layout As String
        Public Nome_baia As String
        Public Stato As String
    End Class

    Public Class Dettagli_commesse_aperte
        Public numero_baia As Integer
        Public zona_layout As String
        Public Nome_baia As String
        Public Stato As String
    End Class

    Public Class Dettagli_commesse_A_numero_baia
        Public numero_baia As Integer
        Public zona_layout As String
        Public Nome_baia As String
        Public Stato As String
    End Class

    Public Class Dettagli_risorsa
        Public commessa As String
        Public cliente As String
    End Class

#End Region

#Region "Event handlers UI"

    Private Sub Label2_MouseDown(sender As Object, e As MouseEventArgs) Handles Label2.MouseDown
        If e.Button = MouseButtons.Left Then Label2.DoDragDrop(Label2.Text, DragDropEffects.Copy)
    End Sub

    Private Sub officina_Enter(sender As Object, e As EventArgs) Handles Officina.Enter
        zona = "Officina"
        AbilitaDragDropZona(zona)
    End Sub

    Private Sub Magazzino_Enter(sender As Object, e As EventArgs) Handles Magazzino.Enter
        zona = "Magazzino"
        nomina_baie(zona)
        AbilitaDragDropZona(zona)
    End Sub

    Private Sub esterno_Enter(sender As Object, e As EventArgs) Handles Esterno.Enter
        zona = "Esterno"
        AbilitaDragDropZona(zona)
    End Sub

    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        compila_risorse()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        compila_datagridview_macchine(DataGridView1)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        compila_datagridview_macchine(DataGridView1)
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        compila_datagridview_macchine(DataGridView1)
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        compila_risorse()
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        compila_risorse()
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        compila_datagridview_macchine(DataGridView1)
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        compila_datagridview_macchine(DataGridView1)
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        FiltraDataGridView(DataGridView3, "comm", TextBox10.Text)
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        FiltraDataGridView(DataGridView3, "cliente_", TextBox11.Text)
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        FiltraDataGridView(DataGridView3, "baia_", TextBox12.Text)
    End Sub

    ''' <summary>Filtra le righe di una DGV per colonna e testo.</summary>
    Private Sub FiltraDataGridView(dgv As DataGridView, colonna As String, filtro As String)
        Dim f = filtro.Trim().ToLower()
        For Each riga As DataGridViewRow In dgv.Rows
            If Not riga.IsNewRow Then
                riga.Visible = If(riga.Cells(colonna).Value, "").ToString().ToLower().Contains(f)
            End If
        Next
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex < 0 Then Return
        Codice_commessa = DataGridView1.Rows(e.RowIndex).Cells("Commessa").Value?.ToString()
        TextBox4.Text = Codice_commessa

        If Codice_commessa >= "M04000" Then
            Scheda_tecnica.Close()
            Scheda_tecnica.Show()
            Scheda_tecnica.BringToFront()
            Scheda_tecnica.inizializza_scheda_tecnica(Codice_commessa)
            Try
                Scheda_tecnica.codice_bp_campione = DataGridView1.Rows(e.RowIndex).Cells("Codice_cliente").Value
            Catch
            End Try
        ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(codice_Progetto) Then
            Progetto.Show()
            Progetto.BringToFront()
            Progetto.absentry = DataGridView1.Rows(e.RowIndex).Cells("absentry_progetto").Value
            Progetto.inizializza_progetto()
        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Rows(e.RowIndex).Cells("Baia").Value?.ToString() <> "" Then
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.Font =
                New Font(DataGridView1.DefaultCellStyle.Font, FontStyle.Bold)
            Select Case DataGridView1.Rows(e.RowIndex).Cells("Zona_baia_data").Value?.ToString()
                Case "Magazzino" : DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.Blue
                Case "Esterno" : DataGridView1.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.Orange
            End Select
        End If
    End Sub

    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting
        If e.RowIndex < 0 OrElse e.ColumnIndex <> 0 Then Return
        Dim dgv = CType(sender, DataGridView)
        If Not dgv.Columns.Contains("Inizio_pian") OrElse Not dgv.Columns.Contains("Fine_pian") Then Return

        Dim riga = dgv.Rows(e.RowIndex)
        If Not IsDate(riga.Cells("Inizio_pian").Value) OrElse Not IsDate(riga.Cells("Fine_pian").Value) Then Return

        Dim inizio = CDate(riga.Cells("Inizio_pian").Value).Date
        Dim fine = CDate(riga.Cells("Fine_pian").Value).Date
        Dim oggi = Date.Today

        riga.DefaultCellStyle.BackColor =
            If(fine < oggi, Color.LightGreen,
            If(inizio <= oggi AndAlso oggi <= fine, Color.Yellow,
            Color.Orange))
    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        If e.RowIndex < 0 Then Return
        Codice_risorsa_ = DataGridView3.Rows(e.RowIndex).Cells("Resuid").Value?.ToString()
        TextBox8.Text = DataGridView3.Rows(e.RowIndex).Cells("Risorsa").Value?.ToString()
    End Sub

    Private Sub DataGridView3_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView3.CellFormatting
        If DataGridView3.Rows(e.RowIndex).Cells("Comm").Value?.ToString() <> "" Then
            DataGridView3.Rows(e.RowIndex).DefaultCellStyle.Font =
                New Font(DataGridView3.DefaultCellStyle.Font, FontStyle.Bold)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "[]" Then
            Panel1.Dock = DockStyle.Fill
            Button1.Text = "-"
            SplitContainer1.Panel1Collapsed = True
            Timer1.Start()
        Else
            Button1.Text = "[]"
            SplitContainer1.Panel1Collapsed = False
            Panel1.Dock = DockStyle.Fill
            Timer1.Stop()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Panel1.Dock = DockStyle.Left
        Panel1.AutoScroll = True
        Panel1.Width += 500
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If Codice_commessa = "" Then MsgBox("Indicare una commessa") : Return
        If ComboBox1.SelectedIndex < 0 Then MsgBox("Selezionare una baia") : Return

        Dim info = check_baia_layout(Codice_commessa)
        If info.numero_baia > 0 Then
            ComboBox1.SelectedIndex = info.numero_baia
            MsgBox("Macchina già presente nella baia " & ComboBox1.Text)
            Return
        End If

        inserisci_record_baia(Codice_commessa, ComboBox1.SelectedIndex, "P")
        inserisci_record_baia_log(Codice_commessa, ComboBox1.SelectedIndex, "IN")
        check_presenza_commessa_baia_layout(ComboBox1.SelectedIndex, zona)
        MsgBox(Codice_commessa & " inserita con successo nella baia " & ComboBox1.Text)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        check_presenza_commessa_baia_layout(0, zona)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Panel1.Dock = DockStyle.Fill
        SplitContainer1.Panel1Collapsed = True
        Button1.Text = "-"

        Dim bmp As New Bitmap(Panel1.Width, Panel1.Height)
        Panel1.DrawToBitmap(bmp, New Rectangle(0, 0, bmp.Width, bmp.Height))

        Using sfd As New SaveFileDialog() With {
            .Title = "Salva Panel1 come PDF",
            .Filter = "File PDF (*.pdf)|*.pdf",
            .FileName = "Panel1.pdf"}

            If sfd.ShowDialog() <> DialogResult.OK Then
                bmp.Dispose()
                Return
            End If

            Using doc As New PdfDocument()
                Dim page As PdfPage = doc.AddPage()
                page.Size = PdfSharp.PageSize.A4
                page.Orientation = PdfSharp.PageOrientation.Landscape

                Using gfx = XGraphics.FromPdfPage(page)
                    Using ximg = XImage.FromGdiPlusImage(bmp)
                        Dim scale = Math.Min(page.Width / ximg.PixelWidth, page.Height / ximg.PixelHeight)
                        Dim fw = ximg.PixelWidth * scale
                        Dim fh = ximg.PixelHeight * scale
                        gfx.DrawImage(ximg, (page.Width - fw) / 2, (page.Height - fh) / 2, fw, fh)
                    End Using
                End Using
                doc.Save(sfd.FileName)
                Process.Start(sfd.FileName)
            End Using
        End Using

        bmp.Dispose()
        MessageBox.Show("PDF creato in orizzontale!", "Successo", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If Codice_risorsa_ = "" Then MsgBox("Indicare una risorsa") : Return
        If TextBox9.Text = "" Then MsgBox("Indicare una commessa") : Return

        inserisci_record_risorsa(Codice_risorsa_, TextBox9.Text)

        For Each row As DataGridViewRow In DataGridView3.Rows
            If Not row.IsNewRow AndAlso row.Cells("Resuid").Value?.ToString() = Codice_risorsa_ Then
                row.Cells("comm").Value = TextBox9.Text
                Exit For
            End If
        Next

        MsgBox(TextBox8.Text & " assegnato con successo alla commessa " & TextBox9.Text)
        compila_risorse()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        If Codice_risorsa_ = "" Then MsgBox("Indicare una risorsa") : Return

        Cancella_record(Codice_risorsa_)

        For Each row As DataGridViewRow In DataGridView3.Rows
            If Not row.IsNewRow AndAlso row.Cells("Resuid").Value?.ToString() = Codice_risorsa_ Then
                row.Cells("comm").Value = ""
                Exit For
            End If
        Next

        MsgBox(TextBox8.Text & " cancellato dalla commessa")
        compila_risorse()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Process.Start("\\tirfs01\00-Report aziendali\Pianificazione\PRODUZIONE\OPERATIONS\layout produzione\Layout officina.dwg")
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        copia_dati_T40_in_Galileo()
        Beep()
        MsgBox("Esportazione eseguita")
    End Sub

    Private Sub GroupBox39_Click(sender As Object, e As EventArgs) Handles GroupBox39.Click
        TextBox9.Text = GroupBox39.Text
    End Sub

    Private Sub GroupBox40_Click(sender As Object, e As EventArgs) Handles GroupBox40.Click
        TextBox9.Text = GroupBox40.Text
    End Sub

    Private Sub GroupBox41_Click(sender As Object, e As EventArgs) Handles GroupBox41.Click
        TextBox9.Text = GroupBox41.Text
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        stato_kpi = CheckBox1.Checked
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        check_presenza_commessa_baia_layout(0, zona)
    End Sub

    Private Sub modulo_baia_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        If e.Button = MouseButtons.Left Then DoDragDrop(Me.Label1.Text, DragDropEffects.Copy)
    End Sub

#End Region

End Class