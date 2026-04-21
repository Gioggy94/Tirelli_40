Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop
Imports System.Windows.Documents
Imports System.Diagnostics
Imports System.Net.Http
Imports System.Text
Imports Newtonsoft.Json.Linq


Public Class UT
    Public Elenco_gruppi(10000) As String

    Private filtro_fam_disegno As String
    Private _picRicercaImg As New PictureBox()
    Private _btnCercaImmagine As New Button()
    Private _lblPastaHint As New Label()
    Private visualizzazione As String = "TIRELLI"
    Private filtro_descrizione As String
    Public filtro_descrizione_supp As String
    Private filtro_disegno As String
    Private filtro_catalogo As String

    ' Cronologia conversazione AI
    Private _aiHistory As New JArray()
    Private _lastSqlImpegni As String = ""
    Private _btnCopiaSql As New Button()

    ' Tab impegni ODP
    Private dgvImpegni As DataGridView
    Private tabImpegni As TabPage

    ' Tab ordini di produzione
    Private dgvODP As DataGridView
    Private tabODP As TabPage
    Private _txtOdpCommessa As TextBox
    Private _txtOdpSottocommessa As TextBox
    Private _txtOdpMatricola As TextBox
    Private _txtOdpCodice As TextBox
    Private _txtOdpDescrizione As TextBox

    ' Filter mode switcher
    Private _pnlFiltriSwitcher As Panel
    Private _pnlFiltriImpegni As Panel
    Private _pnlFiltriODP As Panel
    Private _btnSwitchAna As Button
    Private _btnSwitchImp As Button
    Private _btnSwitchOdp As Button

    ' Impegni filter textboxes
    Private _txtImpCommessa As TextBox
    Private _txtImpSottocommessa As TextBox
    Private _txtImpCodice As TextBox
    Private _txtImpDescrizione As TextBox
    Private _txtImpUltimi As TextBox

    Private Const AI_SYSTEM_PROMPT As String =
        "Sei un assistente per la ricerca articoli nel gestionale Tirelli (AS400/SAP)." & vbLf & vbLf &
        "SCHEMA AS400 (S786FAD1.TIR90VIS):" & vbLf &
        "JGALART: CODE=codice, DES_CODE=descrizione, DISEGNO=disegno," & vbLf &
        "  GRUP_ART=codice-gruppo, DESC_GRP=descrizione-gruppo, PROD_FOR=produttore, CODAR_FOR=codice-catalogo-fornitore," & vbLf &
        "  DESC_FOR=fornitore-preferito, UBI_CODE=ubicazione, COSTO_STD=costo," & vbLf &
        "  STAT_CODE=stato(A=attivo/I=inattivo), UMIS=unità-misura, TIPO_PARTE=approvvigionamento(A=acquisto/P=produzione)" & vbLf &
        "JGALMAG: CODART, MAG, QTA_MAG=giacenza, QTA_DISP=disponibile" & vbLf &
        "JGALord: CODART, DISEGNO, NUMDOC=n.ordine, QTA_ORD, DATA_RICHIESTA, EVASO(S/N), COD_FORN" & vbLf &
        "JGALODP (ordini di produzione): NUMODP=n.ordine-prod, CODART, DSCODART_ODP=descrizione, COD_COMMESSA=commessa(es.T000001), COD_SOTTOCOMMESSA=sottocommessa(es.05412), QTA_PIA=q.pianificata, QTA_RES=q.residua, MAG_VER=mag-destinazione, MATRICOLA=matricola-vecchio-formato(es.M05412), DATA_SCAD=data-scadenza(YYYYMMDD)" & vbLf &
        "JGALIMP (impegni di produzione): CODART, ITEMNAME=descrizione, ODP=n.odp, QTAPIA=q.pianificata, QTATRA=q.trasferita, QTADATRA=q.da-trasferire, EVASO_ODP(S/N), MATRICOLA=commessa(old), COD_COMMESSA=commessa(es.T000001), COD_SOTTOCOMMESSA=sottocommessa(es.05000), DTASCA=data-avvio-odp(YYYYMMDD)" & vbLf & vbLf &
        "REGOLA OBBLIGATORIA: rispondi SEMPRE con JSON puro (niente testo prima/dopo, niente ```json)." & vbLf & vbLf &
        "Se la richiesta è una RICERCA di articoli, usa:" & vbLf &
        "{""azione"":""cerca"",""messaggio"":""spiegazione breve"",""filtri"":{" & vbLf &
        "  ""codice"":"""",""descrizione"":"""",""desc_supp"":"""",""disegno"":""""," & vbLf &
        "  ""catalogo"":"""",""produttore"":"""",""fornitore"":"""",""ubicazione"":"""",""gruppo"":""""," & vbLf &
        "  ""prezzo_min"":"""",""prezzo_max"":"""",""stato"":"""",""tipo_parte"":""""}}" & vbLf & vbLf &
        "Il campo 'gruppo' deve contenere la DESCRIZIONE del gruppo articoli (DESC_GRP), non il codice numerico." & vbLf &
        "Esempi di descrizioni gruppo: 'TAGLIO LASER', 'CUSCINETTI', 'VITI E BULLONI', ecc." & vbLf & vbLf &
        "FILTRI NUMERICI E STATO:" & vbLf &
        "  prezzo_min: numero decimale (es. 100), lascia vuoto se nessun limite minimo" & vbLf &
        "  prezzo_max: numero decimale (es. 500), lascia vuoto se nessun limite massimo" & vbLf &
        "  stato: 'A'=solo attivi, 'I'=solo inattivi, ''=tutti" & vbLf &
        "  tipo_parte: 'A'=solo acquisto esterno, 'P'=solo produzione interna, ''=tutti" & vbLf & vbLf &
        "Se vuole gli IMPEGNI di produzione (componenti/materiali richiesti da un ODP, in JGALIMP), usa:" & vbLf &
        "{""azione"":""impegni"",""messaggio"":""spiegazione"",""filtri"":{""commessa"":"""",""sottocommessa"":"""",""codice_articolo"":"""",""descrizione_articolo"":"""",""ultimi_odp"":""""}}" & vbLf &
        "commessa: codice commessa (es. T000001) — lascia vuoto se vuole ultimi ODP." & vbLf &
        "sottocommessa: es. 05700 (lascia vuoto se non specificata)." & vbLf &
        "codice_articolo: filtro su CODART — usare quando l'utente menziona un CODICE articolo (es. C00001, D115918). Lascia vuoto se cerca per descrizione." & vbLf &
        "descrizione_articolo: filtro su ITEMNAME — usare per parole nella DESCRIZIONE (es. sensore, fotocell). NON usare per codici articolo." & vbLf &
        "ultimi_odp: numero di righe da mostrare ordinate per data ODP piu recente (es. 200). Lascia vuoto se cerchi per commessa o codice." & vbLf & vbLf &
        "Se vuole cercare ORDINI DI PRODUZIONE (articoli da produrre, in JGALODP), usa:" & vbLf &
        "{""azione"":""odp"",""messaggio"":""spiegazione"",""filtri"":{""commessa"":"""",""sottocommessa"":"""",""matricola"":"""",""codice_articolo"":"""",""descrizione_articolo"":""""}}" & vbLf &
        "commessa: COD_COMMESSA (es. T000001) — ricerca parziale, lascia vuoto se non specificata." & vbLf &
        "sottocommessa: COD_SOTTOCOMMESSA (es. 05412) — lascia vuoto se non specificata." & vbLf &
        "matricola: MATRICOLA vecchio formato (es. M05412) — alternativa a commessa." & vbLf &
        "codice_articolo: filtro su CODART — usare quando l'utente menziona un CODICE articolo (es. C00001). Lascia vuoto se cerca per descrizione." & vbLf &
        "descrizione_articolo: filtro su DSCODART_ODP (es. fotocell, sensore) — NON usare per codici articolo." & vbLf & vbLf &
        "DIFFERENZA impegni vs odp:" & vbLf &
        "  impegni = componenti/materiali USATI per produrre (JGALIMP)" & vbLf &
        "  odp     = articoli DA PRODURRE, gli ordini di produzione stessi (JGALODP)" & vbLf & vbLf &
        "Se vuole ORDINARE i risultati già mostrati, usa:" & vbLf &
        "{""azione"":""ordina"",""messaggio"":""spiegazione"",""colonna"":""NomeColonna"",""direzione"":""ASC|DESC""}" & vbLf &
        "Colonne disponibili per ordinamento: Codice, Descrizione (=Nome), Costo, Giacenza (=Giagenza), Disp, Fornitore (=Fornitore_preferito), Produttore, Gruppo_articoli" & vbLf & vbLf &
        "Se la richiesta è INFORMATIVA (definizioni, spiegazioni, dati da DB non cercabili tramite filtri), usa:" & vbLf &
        "{""azione"":""info"",""messaggio"":""risposta completa""}" & vbLf & vbLf &
        "SINTASSI VALORI FILTRO (campi testo):" & vbLf &
        "  parola     = contiene 'parola'" & vbLf &
        "  *parola    = inizia con 'parola'" & vbLf &
        "  parola*    = finisce con 'parola'" & vbLf &
        "  p1*p2      = contiene p1 E p2" & vbLf &
        "  """"        = nessun filtro su quel campo" & vbLf & vbLf &
        "ESEMPI:" & vbLf &
        "Utente: 'cerca cuscinetti SKF 6205' → {""azione"":""cerca"",""messaggio"":""Cerco cuscinetti con catalogo 6205 e produttore SKF"",""filtri"":{""codice"":"""",""descrizione"":""cuscinett*"",""desc_supp"":"""",""disegno"":"""",""catalogo"":""6205"",""produttore"":""SKF"",""fornitore"":"""",""ubicazione"":"""",""gruppo"":"""",""prezzo_min"":"""",""prezzo_max"":"""",""stato"":"""",""tipo_parte"":""""}}" & vbLf &
        "Utente: 'articoli che costano più di 100 euro' → {""azione"":""cerca"",""messaggio"":""Cerco articoli con prezzo > 100"",""filtri"":{""codice"":"""",""descrizione"":"""",""desc_supp"":"""",""disegno"":"""",""catalogo"":"""",""produttore"":"""",""fornitore"":"""",""ubicazione"":"""",""gruppo"":"""",""prezzo_min"":""100"",""prezzo_max"":"""",""stato"":"""",""tipo_parte"":""""}}" & vbLf &
        "Utente: 'codici taglio laser di Galati' → {""azione"":""cerca"",""messaggio"":""Cerco articoli taglio laser con fornitore Galati"",""filtri"":{""codice"":"""",""descrizione"":"""",""desc_supp"":"""",""disegno"":"""",""catalogo"":"""",""produttore"":"""",""fornitore"":""galati"",""ubicazione"":"""",""gruppo"":""taglio laser"",""prezzo_min"":"""",""prezzo_max"":"""",""stato"":"""",""tipo_parte"":""""}}" & vbLf &
        "Utente: 'fotocellule usate nella commessa T000001 sottocommessa 05000' → {""azione"":""impegni"",""messaggio"":""Cerco fotocellule negli impegni commessa T000001/05000"",""filtri"":{""commessa"":""T000001"",""sottocommessa"":""05000"",""codice_articolo"":"""",""descrizione_articolo"":""fotocell"",""ultimi_odp"":""""}}" & vbLf &
        "Utente: 'tutti i componenti della commessa T000001' → {""azione"":""impegni"",""messaggio"":""Cerco impegni commessa T000001"",""filtri"":{""commessa"":""T000001"",""sottocommessa"":"""",""codice_articolo"":"""",""descrizione_articolo"":"""",""ultimi_odp"":""""}}" & vbLf &
        "Utente: 'sensori negli ultimi ODP' → {""azione"":""impegni"",""messaggio"":""Cerco sensori negli impegni ODP piu recenti"",""filtri"":{""commessa"":"""",""sottocommessa"":"""",""codice_articolo"":"""",""descrizione_articolo"":""sensor"",""ultimi_odp"":""200""}}" & vbLf &
        "Utente: 'componenti degli ultimi ordini di produzione' → {""azione"":""impegni"",""messaggio"":""Mostro componenti ODP piu recenti"",""filtri"":{""commessa"":"""",""sottocommessa"":"""",""codice_articolo"":"""",""descrizione_articolo"":"""",""ultimi_odp"":""200""}}" & vbLf &
        "Utente: 'impieghi del codice C00001' → {""azione"":""impegni"",""messaggio"":""Cerco impieghi del codice articolo C00001"",""filtri"":{""commessa"":"""",""sottocommessa"":"""",""codice_articolo"":""C00001"",""descrizione_articolo"":"""",""ultimi_odp"":""200""}}" & vbLf &
        "Utente: 'dove viene usato il codice D115918' → {""azione"":""impegni"",""messaggio"":""Cerco impieghi del codice D115918"",""filtri"":{""commessa"":"""",""sottocommessa"":"""",""codice_articolo"":""D115918"",""descrizione_articolo"":"""",""ultimi_odp"":""200""}}" & vbLf &
        "Utente: 'cosa significa TIPO_PARTE=A?' → {""azione"":""info"",""messaggio"":""TIPO_PARTE=A significa acquisto esterno (il componente viene acquistato, non prodotto internamente).""}" & vbLf &
        "Utente: 'ordinali per prezzo' → {""azione"":""ordina"",""messaggio"":""Ordino per prezzo crescente"",""colonna"":""Costo"",""direzione"":""ASC""}" & vbLf &
        "Utente: 'dal più costoso al meno costoso' → {""azione"":""ordina"",""messaggio"":""Ordino per prezzo decrescente"",""colonna"":""Costo"",""direzione"":""DESC""}" & vbLf &
        "Utente: 'fotocellule nella commessa T000001 sottocommessa 05412' → {""azione"":""odp"",""messaggio"":""Cerco fotocellule in JGALODP commessa T000001/05412"",""filtri"":{""commessa"":""T000001"",""sottocommessa"":""05412"",""matricola"":"""",""codice_articolo"":"""",""descrizione_articolo"":""fotocell""}}" & vbLf &
        "Utente: 'articoli da produrre con matricola M05412' → {""azione"":""odp"",""messaggio"":""Cerco ODP con matricola M05412"",""filtri"":{""commessa"":"""",""sottocommessa"":"""",""matricola"":""M05412"",""codice_articolo"":"""",""descrizione_articolo"":""""}}" & vbLf &
        "Utente: 'sensori presenti nella sottocommessa 05412' → {""azione"":""odp"",""messaggio"":""Cerco sensori in JGALODP sottocommessa 05412"",""filtri"":{""commessa"":"""",""sottocommessa"":""05412"",""matricola"":"""",""codice_articolo"":"""",""descrizione_articolo"":""sensor""}}" & vbLf &
        "Utente: 'tutti gli ODP della commessa T000001' → {""azione"":""odp"",""messaggio"":""Cerco tutti gli ODP commessa T000001"",""filtri"":{""commessa"":""T000001"",""sottocommessa"":"""",""matricola"":"""",""codice_articolo"":"""",""descrizione_articolo"":""""}}" & vbLf &
        "Utente: 'ODP del codice C00001' → {""azione"":""odp"",""messaggio"":""Cerco ordini produzione per codice C00001"",""filtri"":{""commessa"":"""",""sottocommessa"":"""",""matricola"":"""",""codice_articolo"":""C00001"",""descrizione_articolo"":""""}}"

    Private ReadOnly _colFiltroAttivo As Color = Color.FromArgb(255, 250, 180)
    Private _filtroPrezzoMin As Decimal = Decimal.MinValue
    Private _filtroPrezzoMax As Decimal = Decimal.MaxValue
    Private _filtroStato As String = ""
    Private _filtroTipoParte As String = ""
    Private _modalitaImpegni As Boolean = False
    Private _modalitaODP As Boolean = False
    Private _headersOriginali As New Dictionary(Of String, String)()

    Private Sub SetStatus(msg As String, Optional isError As Boolean = False)
        txbStatus.Text = msg
        txbStatus.ForeColor = If(isError, Color.DarkRed, Color.DarkGreen)
        txbStatus.Refresh()
    End Sub

    Private Sub AggiornaEvidenziazioneFiltri()
        Dim evidenzia = Sub(tb As Control, attivo As Boolean)
                            tb.BackColor = If(attivo, _colFiltroAttivo, SystemColors.Window)
                        End Sub
        evidenzia(TextBox_codice_SAP_RICERCA, TextBox_codice_SAP_RICERCA.Text <> "")
        evidenzia(TextBox_Descrizione_ricerca, TextBox_Descrizione_ricerca.Text <> "")
        evidenzia(TextBox3, TextBox3.Text <> "")
        evidenzia(TextBox_disegno_SAP, TextBox_disegno_SAP.Text <> "")
        evidenzia(TextBox6, TextBox6.Text <> "")
        evidenzia(TextBox2, TextBox2.Text <> "")
        evidenzia(TextBox4, TextBox4.Text <> "")
        evidenzia(ComboBox4, ComboBox4.Text <> "")
        evidenzia(ComboBox6, ComboBox6.SelectedIndex > 0)
        ComboBox1.BackColor = If(ComboBox1.SelectedIndex > 0, _colFiltroAttivo, SystemColors.Window)
    End Sub

    Sub inizializzazione_ut()
        AggiungiColonnaImmagine()
        InitComboBoxesAsync()
    End Sub

    Private Async Sub InitComboBoxesAsync()
        ' Leggi le proprietà di Homepage sul thread UI prima di entrare nei Task
        Dim connStr As String = Homepage.sap_tirelli
        Dim erp As String = Homepage.ERP_provenienza

        Dim tGruppi = Task.Run(Function() CaricaGruppiData(connStr, erp))
        Dim tProduttori = Task.Run(Function() CaricaProduttoriData(connStr, erp))
        Dim tFornitori = Task.Run(Function() CaricaFornitoriData(connStr, erp))
        Await Task.WhenAll(tGruppi, tProduttori, tFornitori)

        ' Gruppi — una sola query, popola entrambi i combo
        ComboBox1.Items.Clear() : ComboBox1.Items.Add("")
        ComboBox9.Items.Clear() : ComboBox9.Items.Add("")
        Dim idx As Integer = 1
        For Each g In tGruppi.Result
            Elenco_gruppi(idx) = g.Item1
            ComboBox1.Items.Add(g.Item2)
            ComboBox9.Items.Add(g.Item2)
            idx += 1
        Next
        ComboBox1.SelectedIndex = 0
        ComboBox9.SelectedIndex = 0

        ' Produttori
        ComboBox4.Items.Clear() : ComboBox4.Items.Add("") : ComboBox4.SelectedIndex = 0
        For Each p In tProduttori.Result
            ComboBox4.Items.Add(p)
        Next

        ' Fornitori
        ComboBox6.Items.Clear()
        For Each f In tFornitori.Result
            ComboBox6.Items.Add(f)
        Next
    End Sub

    Private Function CaricaGruppiData(connStr As String, erp As String) As List(Of Tuple(Of String, String))
        Dim result As New List(Of Tuple(Of String, String))()
        Using Cnn As New SqlConnection(connStr)
            Cnn.Open()
            Dim sql As String
            If erp = "SAP" Then
                sql = "SELECT T0.[ItmsGrpCod] AS Gruppo, T0.[ItmsGrpNam] AS [Nome gruppo] FROM OITB T0"
            Else
                sql = "SELECT trim(t0.grup_art) as Gruppo, trim(t0.desc_grp) as [Nome gruppo]
FROM OPENQUERY(AS400, '
    SELECT DISTINCT grup_art, desc_grp
    FROM S786FAD1.TIR90VIS.JGALART
    WHERE stat_code <> ''I'' AND grup_art <> '' ''
') as t0
ORDER BY t0.grup_art"
            End If
            Using cmd As New SqlCommand(sql, Cnn)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        result.Add(Tuple.Create(rdr("Gruppo").ToString(), rdr("Nome gruppo").ToString()))
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Private Function CaricaProduttoriData(connStr As String, erp As String) As List(Of String)
        Dim result As New List(Of String)()
        Using Cnn As New SqlConnection(connStr)
            Cnn.Open()
            Dim sql As String
            If erp = "SAP" Then
                sql = "Select t0.firmname as nome from tirellisrldb.dbo.omrc t0 ORDER BY T0.FIRMNAME"
            Else
                sql = "SELECT trim(t0.prod_for) as nome
FROM OPENQUERY(AS400, '
    SELECT DISTINCT prod_for
    FROM S786FAD1.TIR90VIS.JGALART
    WHERE stat_code <> ''I'' AND prod_for <> '' ''
') as t0
ORDER BY t0.prod_for"
            End If
            Using cmd As New SqlCommand(sql, Cnn)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim v = rdr("nome").ToString().Trim()
                        If v <> "" Then result.Add(v)
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Private Function CaricaFornitoriData(connStr As String, erp As String) As List(Of String)
        Dim result As New List(Of String)()
        Using Cnn As New SqlConnection(connStr)
            Cnn.Open()
            Dim sql As String
            If erp = "SAP" Then
                sql = "Select t0.CARDname from OCRD t0 WHERE T0.CARDTYPE='S' and t0.validfor='Y' ORDER BY T0.CARDNAME"
            Else
                sql = "SELECT trim(t0.ds_Conto) as ds_Conto
FROM OPENQUERY(AS400, '
    SELECT conto, ds_Conto
    FROM S786FAD1.TIR90VIS.JGALACF
    WHERE clifor=''F'' and stato<>''S''
') as t0
ORDER BY t0.ds_Conto"
            End If
            Using cmd As New SqlCommand(sql, Cnn)
                Using rdr = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim campo = If(erp = "SAP", "CARDname", "ds_Conto")
                        Dim v = rdr(campo).ToString().Trim()
                        If v <> "" Then result.Add(v)
                    End While
                End Using
            End Using
        End Using
        Return result
    End Function

    Private Sub AggiungiColonnaImmagine()
        If DataGridView_SAP.Columns.Contains("Immagine") Then Return
        Dim imgCol As New DataGridViewImageColumn()
        imgCol.Name = "Immagine"
        imgCol.HeaderText = ""
        imgCol.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
        imgCol.Width = 130
        imgCol.MinimumWidth = 130
        imgCol.FillWeight = 1
        imgCol.Resizable = DataGridViewTriState.False
        imgCol.ImageLayout = DataGridViewImageCellLayout.Zoom
        imgCol.DefaultCellStyle.NullValue = Nothing
        DataGridView_SAP.Columns.Add(imgCol)
        imgCol.DisplayIndex = 1
        DataGridView_SAP.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None
        DataGridView_SAP.RowTemplate.Height = 70
    End Sub

    Sub cerca(Par_codice As String, par_descrizione As String, par_descrizone_supp As String, par_disegno As String, par_fam_disegno As String, par_gruppo_art As String, par_produttore As String, par_cat_forn As String, par_fornitore As String, par_ubicazione As String)
        Dim produttore As String
        Dim fornitore_preferito As String
        Dim ubic As String

        If ComboBox4.Text = "" Then
            produttore = ""
        Else
            produttore = " and t2.firmname like '%%" & ComboBox4.Text & "%%'"
        End If

        If ComboBox6.SelectedIndex < 0 Then
            fornitore_preferito = ""
        Else
            fornitore_preferito = " and t3.cardname= '" & ComboBox6.Text & "'"
        End If

        If TextBox4.Text = "" Then
            ubic = ""
        Else
            If visualizzazione = "TIRELLI" Then
                ubic = " And T0.[u_ubicazione] Like '%%" & TextBox4.Text & "%%' "
            Else
                ubic = " and T0.[ubicazione] like '%%" & TextBox4.Text & "%%' "
            End If
        End If

        If TextBox_disegno_SAP.Text = Nothing Then
            TextBox_disegno_SAP.Text = ""
        End If

        ImpostaModeFiltri("anagrafica")
        ImpostaModalitaImpegni(False)
        ImpostaModalitaODP(False)
        DataGridView_SAP.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim Cmd_SAP As New SqlCommand
        Dim Cmd_SAP_Reader As SqlDataReader

        Cmd_SAP.Connection = Cnn
        If Homepage.ERP_provenienza = "sap" Then

            Cmd_SAP.CommandText = "select top " & TextBox9.Text & " t10.codice,t10.nome,t10.[Nome supp],t10.disegno,t10.[Gruppo art],t10.FirmName, t10.SuppCatNum,t10.CardName,t10.u_ubicazione, t10.u_famiglia_disegno,
t10.onhand,t10.WIP,sum(coalesce(t11.onhand,0)+coalesce(t11.onorder,0)-coalesce(t11.iscommited,0)) as 'Disp',t10.price,t10.frozenfor
from
(
SELECT top 100 T0.[ItemCode] as 'Codice', coalesce(T0.[ItemName],'') as 'Nome', coalesce(T0.[FrgnName],'') as 'Nome supp',t0.u_disegno as 'Disegno', T1.[ItmsGrpNam] as 'Gruppo art' , t2.firmname, t0.suppcatnum,coalesce(t3.cardname,'') as 'Cardname', t0.u_ubicazione, t0.u_famiglia_disegno,
t0.onhand-t5.onhand-t6.onhand as 'onhand',t5.onhand+t6.onhand as 'WIP',t4.price,t0.frozenfor
        FROM OITM T0  INNER JOIN OITB T1 ON T0.[ItmsGrpCod] = T1.[ItmsGrpCod]
left join omrc t2 on t2.firmcode=t0.firmcode
left join ocrd t3 on t3.cardcode=t0.cardcode
inner join itm1 t4 on t4.itemcode=t0.itemcode and t4.pricelist=2
left join oitw t5 on t5.itemcode =t0.itemcode and t5.whscode='WIP'
left join oitw t6 on t6.itemcode =t0.itemcode and t6.whscode='BWIP'
        WHERE t0.itemcode Like '%%" & TextBox_codice_SAP_RICERCA.Text & "%%' " & filtro_descrizione & fornitore_preferito & filtro_catalogo & filtro_disegno & filtro_descrizione_supp & "  and t0.itmsgrpcod like '%%" & Elenco_gruppi(ComboBox1.SelectedIndex) & "%%'" & produttore & ubic & filtro_fam_disegno &
             ")
		as t10 left join oitw t11 on t10.codice=t11.itemcode

		group by t10.codice,t10.nome,t10.[Nome supp],t10.disegno,t10.[Gruppo art],t10.FirmName, t10.SuppCatNum,t10.CardName,t10.u_ubicazione, t10.u_famiglia_disegno,
t10.onhand,t10.WIP,t10.price,t10.frozenfor"
        Else
            ' Costruisci WHERE AS400 con supporto asterischi per ogni campo
            Dim where400 As String = "where 1=1"
            where400 &= CostruisciCondizioneAS400("code", Par_codice)
            where400 &= CostruisciCondizioneAS400("t0.des_code", par_descrizione)
            where400 &= CostruisciCondizioneAS400("t0.disegno", par_disegno)
            where400 &= CostruisciCondizioneAS400("t0.grup_art", par_gruppo_art)
            where400 &= CostruisciCondizioneAS400("t0.prod_for", par_produttore)
            where400 &= CostruisciCondizioneAS400("t0.codar_for", par_cat_forn)
            where400 &= CostruisciCondizioneAS400("t0.desc_for", par_fornitore)
            where400 &= CostruisciCondizioneAS400("t0.ubi_code", par_ubicazione)

            Cmd_SAP.CommandText = "select
    trim(t10.code) as 'Codice'
    ,t10.des_code as 'Nome'
    ,'MANCA' AS 'Nome supp'
    ,trim(t10.disegno) as 'Disegno'
    ,trim(t10.desc_grp) as 'Gruppo art'
    ,t10.prod_for as 'firmname'
    ,t10.codar_for as 'Suppcatnum'
    ,t10.desc_for as 'Cardname'
    ,t10.ubi_cOde as 'u_ubicazione'
    ,'manca' as 'u_famiglia_disegno'
    ,T10.onhand as 'onhand'
    ,T10.wip_tot as 'WIP'
    ,T10.disp_tot as 'DISP'
    ,T10.costo_std as 'PRICE'
    ,trim(t10.stat_code) as 'frozenfor'
from openquery(AS400,'
    select
        T0.CODE,
        T0.DES_CODE,
        T0.DISEGNO,
        T0.GRUP_ART,
        T0.DESC_GRP,
        T0.STAT_CODE,
        T0.PROD_FOR,
        T0.CODAR_FOR,
        T0.DESC_FOR,
        T0.UBI_CODE,
        T0.COSTO_STD,
        coalesce(T1.TOT_QTA, 0) as onhand,
        coalesce(T1.TOT_DISP, 0) as disp_tot,
        coalesce(T2.TOT_WIP, 0) as wip_tot
    from S786FAD1.TIR90VIS.JGALART T0
    left join (
        select
            CODART,
            sum(QTA_MAG) as TOT_QTA,
            sum(QTA_DISP) as TOT_DISP
        from TIR90VIS.JGALMAG
        group by CODART
    ) T1 ON T0.CODE = T1.CODART
    left join (
        select
            CODART,
            sum(QTATRA) as TOT_WIP
        from TIR90VIS.JGALIMP
        where EVASO_ODP <> ''S''
        group by CODART
    ) T2 ON T0.CODE = T2.CODART
" & where400 & "
'
)
as t10"

        End If

        Try
            Cmd_SAP_Reader = Cmd_SAP.ExecuteReader
            Do While Cmd_SAP_Reader.Read()
                Dim disVal = If(IsDBNull(Cmd_SAP_Reader("Disegno")), "", Cmd_SAP_Reader("Disegno").ToString().Trim())
                Dim rowIdx = DataGridView_SAP.Rows.Add(Cmd_SAP_Reader("Codice"), Cmd_SAP_Reader("Nome"), Cmd_SAP_Reader("Nome supp"), Cmd_SAP_Reader("u_famiglia_disegno"), Cmd_SAP_Reader("Disegno"), Cmd_SAP_Reader("Gruppo art"), Cmd_SAP_Reader("firmname"), Cmd_SAP_Reader("suppcatnum"), Cmd_SAP_Reader("cardname"), Cmd_SAP_Reader("u_ubicazione"), Cmd_SAP_Reader("onhand"), Cmd_SAP_Reader("wip"), Cmd_SAP_Reader("disp"), Cmd_SAP_Reader("price"), Cmd_SAP_Reader("frozenfor"))
                CaricaImmagineRiga(rowIdx, disVal)
            Loop
            Cmd_SAP_Reader.Close()
            ApplicaPostFiltri()
            SetStatus("Ricerca completata — " & DataGridView_SAP.Rows.Count & " risultati" &
                      If(_filtroPrezzoMin > Decimal.MinValue, "  prezzo≥" & _filtroPrezzoMin, "") &
                      If(_filtroPrezzoMax < Decimal.MaxValue, "  prezzo≤" & _filtroPrezzoMax, "") &
                      If(_filtroStato <> "", "  stato=" & _filtroStato, "") &
                      If(_filtroTipoParte <> "", "  tipo=" & _filtroTipoParte, ""))
            TabControl1.SelectedIndex = 0
        Catch ex As Exception
            SetStatus("Errore: " & ex.Message, True)
        Finally
            Cnn.Close()
        End Try
    End Sub

    Private Sub ApplicaPostFiltri()
        Dim hasPrezzo = _filtroPrezzoMin > Decimal.MinValue OrElse _filtroPrezzoMax < Decimal.MaxValue
        If Not hasPrezzo AndAlso _filtroStato = "" AndAlso _filtroTipoParte = "" Then Return
        For Each row As DataGridViewRow In DataGridView_SAP.Rows
            Dim nascondi = False
            If hasPrezzo Then
                Dim priceCell = row.Cells("Costo").Value
                Dim price As Decimal = 0
                If priceCell IsNot Nothing AndAlso priceCell.ToString() <> "" Then
                    Decimal.TryParse(priceCell.ToString(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, price)
                End If
                If price < _filtroPrezzoMin OrElse price > _filtroPrezzoMax Then nascondi = True
            End If
            If Not nascondi AndAlso _filtroStato <> "" Then
                Dim attivoCell = row.Cells("Attivo").Value?.ToString()
                Dim isInattivo = (attivoCell = "Y" OrElse attivoCell = "I")
                If _filtroStato = "A" AndAlso isInattivo Then nascondi = True
                If _filtroStato = "I" AndAlso Not isInattivo Then nascondi = True
            End If
            row.Visible = Not nascondi
        Next
    End Sub

    Sub cerca_BRB()
        Dim produttore As String
        Dim catalogo_fornitore As String
        Dim ubic As String

        If ComboBox4.Text = "" Then
            produttore = ""
        Else
            produttore = " and (t10.firmname like '%%" & ComboBox4.Text & "%%' or t10.firmname like '%%" & ComboBox4.Text & "%%') "
        End If

        If TextBox2.Text = "" Then
            catalogo_fornitore = ""
        Else
            catalogo_fornitore = " And (t10.SUPPCATNUM Like '%%" & TextBox2.Text & "%%' or t10.catalogo_fornitore like '%%" & TextBox2.Text & "%%') "
        End If

        If TextBox4.Text = "" Then
            ubic = ""
        Else
            ubic = " and (T10.[u_ubicazione_labelling] like '%%" & TextBox4.Text & "%%' or T10.[ubicazione] like '%%" & TextBox4.Text & "%%' )"
        End If

        DataGridView1.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim Cmd_SAP As New SqlCommand
        Dim Cmd_SAP_Reader As SqlDataReader

        Cmd_SAP.Connection = Cnn
        If TextBox_disegno_SAP.Text = Nothing Then
            Cmd_SAP.CommandText = "		select *
from
(

select
coalesce(t1.[ID],0) as 'ID'
      ,coalesce(t1.[Codice_BRB],t0.u_codice_brb) as 'Codice_BRB'
      ,coalesce(t1.[Descrizione_BRB],t0.itemname) as 'Descrizione_BRB'
      ,coalesce(t1.[Descrizione_supp_BRB],t0.frgnname) as 'Descrizione_supp_BRB'
      ,coalesce(t1.[Catalogo_fornitore],t0.suppcatnum) as 'Catalogo_fornitore'
      ,coalesce(t1.[Fornitore],t4.cardname) as 'Fornitore'
      ,cast(coalesce(t1.[Costo],coalesce(t5.price,0)) as integer) as 'Costo'
,coalesce(t1.ubicazione,COALESCE(t0.u_ubicazione_labelling,'')) as 'Ubicazione'
,coalesce(t0.itemcode,'') as 'Itemcode'
,coalesce(t2.tirelli,'') as 'Tirelli'
, coalesce(t2.Gruppo_articoli,t0.[ItmsGrpCod]) as 'Gruppo_articoli'
,coalesce(t3.ItmsGrpNam,t6.ItmsGrpNam) as 'ItmsGrpNam'
, coalesce(t0.itemname,'') as 'Desc_tirelli'
,coalesce(t0.frgnname,'') as 'Desc_supp_tirelli'
,t0.u_codice_brb

  FROM [TIRELLISRLDB].[dbo].[OITM] t0
   left join [TIRELLI_40].[DBO].BRB_Codici t1 on t0.u_codice_BRB=t1.Codice_BRB
  left join [TIRELLI_40].[DBO].[Frontiera_PDM_BRB_SAP_Prima_Lettera] t2 on t2.brb=substring(t1.codice_BRB,1,1)
  left join [TIRELLISRLDB].[dbo].[oitb] t3 on t3.ItmsGrpCod=t2.Gruppo_articoli
  left join ocrd t4 on t4.cardcode=t0.cardcode
  left join itm1 t5 on t5.itemcode=t0.itemcode and t5.pricelist=2
left join [TIRELLISRLDB].[dbo].[oitb] t6 on t6.ItmsGrpCod=t0.[ItmsGrpCod]
left join omrc t7 on t7.firmcode=t0.firmcode

union all

sELECT  coalesce(t1.[ID],0) as 'ID'
      ,coalesce(t1.[Codice_BRB],t0.u_codice_brb) as 'Codice_BRB'
      ,coalesce(t1.[Descrizione_BRB],t0.itemname) as 'Descrizione_BRB'
      ,coalesce(t1.[Descrizione_supp_BRB],t0.frgnname) as 'Descrizione_supp_BRB'
      ,coalesce(t1.[Catalogo_fornitore],t0.suppcatnum) as 'Catalogo_fornitore'
      ,coalesce(t1.[Fornitore],t4.cardname) as 'Fornitore'
      ,cast(coalesce(t1.[Costo],coalesce(t5.price,0)) as integer) as 'Costo'
,coalesce(t1.ubicazione,t0.u_ubicazione_labelling) as 'Ubicazione'
,coalesce(t0.itemcode,'') as 'Itemcode'
,coalesce(t2.tirelli,'') as 'Tirelli'
, coalesce(t2.Gruppo_articoli,t0.[ItmsGrpCod]) as 'Gruppo_articoli'
,coalesce(t3.ItmsGrpNam,t6.ItmsGrpNam) as 'ItmsGrpNam'
, coalesce(t0.itemname,'') as 'Desc_tirelli'
,coalesce(t0.frgnname,'') as 'Desc_supp_tirelli'
,t0.u_codice_brb


 from [TIRELLI_40].[DBO].BRB_Codici t1
left join [TIRELLISRLDB].[dbo].[OITM] t0 on t0.u_codice_BRB=t1.Codice_BRB
  left join [TIRELLI_40].[DBO].[Frontiera_PDM_BRB_SAP_Prima_Lettera] t2 on t2.brb=substring(t1.codice_BRB,1,1)
  left join [TIRELLISRLDB].[dbo].[oitb] t3 on t3.ItmsGrpCod=t2.Gruppo_articoli
  left join ocrd t4 on t4.cardcode=t0.cardcode
  left join itm1 t5 on t5.itemcode=t0.itemcode and t5.pricelist=2
left join [TIRELLISRLDB].[dbo].[oitb] t6 on t6.ItmsGrpCod=t0.[ItmsGrpCod]
left join omrc t7 on t7.firmcode=t0.firmcode

        WHERE t0.itemcode  is null

)
as t10



        WHERE (t10.Codice_BRB Like '%%" & TextBox_codice_SAP_RICERCA.Text & "%%' or t10.u_codice_brb Like '%%" & TextBox_codice_SAP_RICERCA.Text & "%%')   " & filtro_descrizione_supp & filtro_descrizione & catalogo_fornitore & ubic

        End If

        Cmd_SAP_Reader = Cmd_SAP.ExecuteReader

        Try
            Cmd_SAP_Reader = Cmd_SAP.ExecuteReader
            Do While Cmd_SAP_Reader.Read()
                DataGridView1.Rows.Add(Cmd_SAP_Reader("Codice_BRB"), Cmd_SAP_Reader("itemcode"), Cmd_SAP_Reader("Descrizione_BRB"), Cmd_SAP_Reader("descrizione_supp_Brb"), Cmd_SAP_Reader("Fornitore"), Cmd_SAP_Reader("Catalogo_fornitore"), Cmd_SAP_Reader("Ubicazione"), Cmd_SAP_Reader("Costo"), Cmd_SAP_Reader("Tirelli"), Cmd_SAP_Reader("Gruppo_articoli"), Cmd_SAP_Reader("ItmsGrpNam"))
            Loop
            Cmd_SAP_Reader.Close()
            SetStatus("Ricerca BRB completata — " & DataGridView1.Rows.Count & " risultati")
        Catch ex As Exception
            SetStatus("Errore BRB: " & ex.Message, True)
        Finally
            Cnn.Close()
        End Try
    End Sub

    Sub inserimento_gruppi(par_combobox As ComboBox)

        par_combobox.Items.Clear()
        par_combobox.Items.Add("")
        par_combobox.SelectedIndex = 0

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP.CommandText = "SELECT T0.[ItmsGrpCod] AS 'Gruppo', T0.[ItmsGrpNam] as 'Nome gruppo' FROM OITB T0"
        Else
            CMD_SAP.CommandText = "SELECT trim(t0.grup_art) as Gruppo, trim(t0.desc_grp) as 'Nome gruppo'
FROM OPENQUERY(AS400, '
    SELECT DISTINCT grup_art, desc_grp
    FROM S786FAD1.TIR90VIS.JGALART
    WHERE stat_code <> ''I'' AND grup_art <> '' ''
') as t0
ORDER BY t0.grup_art"
        End If

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer = 1
        Do While cmd_SAP_reader.Read()
            Elenco_gruppi(Indice) = cmd_SAP_reader("Gruppo")
            par_combobox.Items.Add(cmd_SAP_reader("Nome gruppo"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub inserimento_produttore()
        ComboBox4.Items.Clear()
        ComboBox4.Items.Add("")
        ComboBox4.SelectedIndex = 0

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn

        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP.CommandText = "Select t0.firmname as nome from tirellisrldb.dbo.omrc t0 ORDER BY T0.FIRMNAME"
        Else
            CMD_SAP.CommandText = "SELECT trim(t0.prod_for) as nome
FROM OPENQUERY(AS400, '
    SELECT DISTINCT prod_for
    FROM S786FAD1.TIR90VIS.JGALART
    WHERE stat_code <> ''I'' AND prod_for <> '' ''
') as t0
ORDER BY t0.prod_for"
        End If

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        Do While cmd_SAP_reader.Read()
            Dim val = cmd_SAP_reader("nome").ToString().Trim()
            If val <> "" Then ComboBox4.Items.Add(val)
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub

    Sub inserimento_FORNITORI()
        ComboBox6.Items.Clear()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD As New SqlCommand
        Dim rdr As SqlDataReader

        CMD.Connection = Cnn

        If Homepage.ERP_provenienza = "SAP" Then
            CMD.CommandText = "Select t0.CARDCODE, t0.CARDname from OCRD t0 WHERE T0.CARDTYPE='S' and t0.validfor='Y' ORDER BY T0.CARDNAME"
            rdr = CMD.ExecuteReader
            Do While rdr.Read()
                ComboBox6.Items.Add(rdr("CARDname"))
            Loop
        Else
            CMD.CommandText = "SELECT trim(t0.conto) as conto, trim(t0.ds_Conto) as ds_Conto
FROM OPENQUERY(AS400, '
    SELECT conto, ds_Conto
    FROM S786FAD1.TIR90VIS.JGALACF
    WHERE clifor=''F'' and stato<>''S''
') as t0
ORDER BY t0.ds_Conto"
            rdr = CMD.ExecuteReader
            Do While rdr.Read()
                Dim nome = rdr("ds_Conto").ToString().Trim()
                If nome <> "" Then ComboBox6.Items.Add(nome)
            Loop
        End If

        rdr.Close()
        Cnn.Close()
    End Sub

    Private Sub UT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Me.BackColor = Homepage.colore_sfondo
        inizializzazione_ut()
        ApplicaStileNavy()
    End Sub

    Private Sub ApplicaStileNavy()
        Dim navy As Color = Color.FromArgb(22, 45, 84)
        Dim navyDark As Color = Color.FromArgb(10, 26, 55)
        Dim navyHover As Color = Color.FromArgb(30, 63, 122)
        Dim bgApp As Color = Color.FromArgb(238, 242, 247)

        ' Barre filtri — sfondo medio, label navy leggibili su bianco
        Dim bgBar As Color = Color.FromArgb(224, 232, 245)    ' azzurro-grigio chiaro
        Panel3.BackColor = bgBar
        Panel26.BackColor = bgBar

        ' Pulsanti X e − — navy scuro su sfondo barra
        For Each b As Button In New Button() {Button3, Button20}
            b.ForeColor = navy
            b.BackColor = bgBar
            b.FlatStyle = FlatStyle.Flat
            b.FlatAppearance.BorderColor = Color.FromArgb(180, 200, 230)
            b.FlatAppearance.MouseOverBackColor = Color.FromArgb(200, 215, 240)
        Next

        ' Pulsante Legenda — nella barra filtri
        BtnLegenda.BackColor = navy
        BtnLegenda.ForeColor = Color.White
        BtnLegenda.FlatStyle = FlatStyle.Flat
        BtnLegenda.FlatAppearance.BorderColor = navyDark
        BtnLegenda.FlatAppearance.MouseOverBackColor = navyHover

        ' Pulsante Cerca
        Button_CERCA.BackColor = navy
        Button_CERCA.ForeColor = Color.White
        Button_CERCA.FlatStyle = FlatStyle.Flat
        Button_CERCA.FlatAppearance.BorderColor = navyDark
        Button_CERCA.FlatAppearance.MouseOverBackColor = navyHover

        ' GroupBox label — navy su sfondo chiaro = leggibile
        For Each gb As GroupBox In New GroupBox() {
            GroupBox2, GroupBox3, GroupBox10, GroupBox4, GroupBox22,
            GroupBox6, GroupBox14, GroupBox12, GroupBox13, GroupBox15, GroupBox27}
            gb.ForeColor = navy
            gb.BackColor = Color.White
        Next

        ' DataGridView header navy
        DataGridView_SAP.ColumnHeadersDefaultCellStyle.BackColor = navy
        DataGridView_SAP.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        DataGridView_SAP.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(238, 242, 247)
        DataGridView_SAP.BackgroundColor = Color.FromArgb(250, 252, 255)
        DataGridView1.ColumnHeadersDefaultCellStyle.BackColor = navy
        DataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        DataGridView1.BackgroundColor = Color.FromArgb(250, 252, 255)

        ' Pannello AI — header più prominente con blu vivace
        Dim accentBlue As Color = Color.FromArgb(0, 120, 212)
        Dim accentDark As Color = Color.FromArgb(0, 90, 160)
        PanelAI.BackColor = bgApp
        PanelAI.BorderStyle = BorderStyle.FixedSingle
        PanelAI_Header.BackColor = accentBlue
        PanelAI_Header.Height = 50
        LabelAI_Title.ForeColor = Color.White
        LabelAI_Title.Font = New Font("Segoe UI", 11.0!, FontStyle.Bold)
        LabelAI_Title.Text = "Assistente AI — Ricerca Articoli"
        RtbChat.BackColor = Color.White
        RtbChat.Font = New Font("Segoe UI", 9.5!)
        PanelAI_Btns.BackColor = accentDark
        BtnAiInvia.BackColor = accentBlue
        BtnAiInvia.ForeColor = Color.White
        BtnAiInvia.FlatAppearance.BorderColor = accentDark
        BtnAiInvia.FlatAppearance.MouseOverBackColor = accentDark
        BtnAiInvia.Location = New Point(80 + 95, 0)
        BtnAiInvia.Size = New Size(250, 32)

        _btnCopiaSql.Text = "Copia SQL"
        _btnCopiaSql.Location = New Point(80, 0)
        _btnCopiaSql.Size = New Size(95, 32)
        _btnCopiaSql.FlatStyle = FlatStyle.Flat
        _btnCopiaSql.BackColor = Color.FromArgb(0, 90, 160)
        _btnCopiaSql.ForeColor = Color.White
        _btnCopiaSql.Font = New Font("Segoe UI", 8.0!)
        _btnCopiaSql.FlatAppearance.BorderColor = Color.FromArgb(0, 60, 120)
        _btnCopiaSql.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 70, 140)
        PanelAI_Btns.Controls.Add(_btnCopiaSql)
        AddHandler _btnCopiaSql.Click, Sub(s, ev)
                                           If _lastSqlImpegni <> "" Then
                                               Clipboard.SetText(_lastSqlImpegni)
                                               SetStatus("SQL impegni copiato negli appunti.")
                                           Else
                                               SetStatus("Nessuna query impegni generata ancora.", True)
                                           End If
                                       End Sub

        ' Allarga colonna AI dal 30% al 34%
        Dim tl = TryCast(PanelAI.Parent, TableLayoutPanel)
        If tl IsNot Nothing Then
            tl.ColumnStyles(0).Width = 34
            tl.ColumnStyles(1).Width = 66
        End If

        AggiungiTabImpegni()
        AggiungiTabODP()
        RiorganizzaFiltri()
        AggiungiPanelRicercaImmagine()
        AggiungiBottoneCopiaSqlStatus()
    End Sub

    Private Sub AggiungiTabImpegni()
        If TabControl1.TabPages.ContainsKey("TabImpegniODP") Then Return

        Dim navy As Color = Color.FromArgb(22, 45, 84)

        tabImpegni = New TabPage()
        tabImpegni.Name = "TabImpegniODP"
        tabImpegni.Text = "Impegni ODP"
        tabImpegni.UseVisualStyleBackColor = True

        dgvImpegni = New DataGridView()
        dgvImpegni.Dock = DockStyle.Fill
        dgvImpegni.AllowUserToAddRows = False
        dgvImpegni.AllowUserToDeleteRows = False
        dgvImpegni.ReadOnly = True
        dgvImpegni.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvImpegni.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        dgvImpegni.BackgroundColor = Color.FromArgb(250, 252, 255)
        dgvImpegni.GridColor = Color.WhiteSmoke
        dgvImpegni.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvImpegni.RowHeadersVisible = False
        dgvImpegni.ColumnHeadersDefaultCellStyle.BackColor = navy
        dgvImpegni.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvImpegni.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.0!, FontStyle.Bold)
        dgvImpegni.EnableHeadersVisualStyles = False

        Dim colDefs As (nome As String, header As String)() = {
            ("ODP", "N. ODP"),
            ("Status", "Stato"),
            ("Commessa", "Commessa"),
            ("Sottocommessa", "Sottocom."),
            ("Matricola", "Matricola"),
            ("DescCommessa", "Desc. Commessa"),
            ("Codice", "Codice"),
            ("Descrizione", "Descrizione"),
            ("Disegno", "Disegno"),
            ("Quantita", "Quantità"),
            ("Costo", "Costo"),
            ("CostoTot", "Costo TOT")
        }
        For Each c In colDefs
            Dim col As New DataGridViewTextBoxColumn()
            col.Name = c.nome
            col.HeaderText = c.header
            dgvImpegni.Columns.Add(col)
        Next

        tabImpegni.Controls.Add(dgvImpegni)
        TabControl1.TabPages.Add(tabImpegni)
    End Sub

    Private Sub AggiungiTabODP()
        If TabControl1.TabPages.ContainsKey("TabODP") Then Return

        Dim navy As Color = Color.FromArgb(22, 45, 84)

        tabODP = New TabPage()
        tabODP.Name = "TabODP"
        tabODP.Text = "Ordini Produzione"
        tabODP.UseVisualStyleBackColor = True

        ' ── Griglia ODP ─────────────────────────────────────────────
        dgvODP = New DataGridView()
        dgvODP.Dock = DockStyle.Fill
        dgvODP.AllowUserToAddRows = False
        dgvODP.AllowUserToDeleteRows = False
        dgvODP.ReadOnly = True
        dgvODP.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        dgvODP.BackgroundColor = Color.FromArgb(250, 252, 255)
        dgvODP.GridColor = Color.WhiteSmoke
        dgvODP.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgvODP.RowHeadersVisible = False
        dgvODP.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255)
        dgvODP.ColumnHeadersDefaultCellStyle.BackColor = navy
        dgvODP.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgvODP.ColumnHeadersDefaultCellStyle.Font = New Font("Segoe UI", 9.0!, FontStyle.Bold)
        dgvODP.EnableHeadersVisualStyles = False

        Dim colDefs As (nome As String, header As String)() = {
            ("ODP", "N. ODP"),
            ("Codice", "Codice art."),
            ("Descrizione", "Descrizione"),
            ("Commessa", "Commessa"),
            ("Sottocommessa", "Sottocom."),
            ("Matricola", "Matricola"),
            ("MagVer", "Mag. Ver."),
            ("QPia", "Q. Pianif."),
            ("QRes", "Q. Res."),
            ("DataScad", "Data scad.")
        }
        For Each c In colDefs
            Dim col As New DataGridViewTextBoxColumn()
            col.Name = c.nome
            col.HeaderText = c.header
            dgvODP.Columns.Add(col)
        Next

        tabODP.Controls.Add(dgvODP)
        TabControl1.TabPages.Add(tabODP)
    End Sub

    Private Sub RiorganizzaFiltri()
        Dim navy As Color = Color.FromArgb(22, 45, 84)
        Dim navyHover As Color = Color.FromArgb(30, 63, 122)

        ' ── Switcher bar ─────────────────────────────────────────────
        _pnlFiltriSwitcher = New Panel()
        _pnlFiltriSwitcher.Dock = DockStyle.Top
        _pnlFiltriSwitcher.Height = 30
        _pnlFiltriSwitcher.BackColor = Color.FromArgb(30, 30, 60)

        Dim flowSwitch As New FlowLayoutPanel()
        flowSwitch.Dock = DockStyle.Fill
        flowSwitch.FlowDirection = FlowDirection.LeftToRight
        flowSwitch.Padding = New Padding(2, 1, 2, 1)
        flowSwitch.WrapContents = False

        Dim creaBtnSwitch = Function(testo As String) As Button
                                Dim b As New Button()
                                b.Text = testo
                                b.Width = 130
                                b.Height = 28
                                b.BackColor = navy
                                b.ForeColor = Color.White
                                b.FlatStyle = FlatStyle.Flat
                                b.FlatAppearance.BorderColor = navyHover
                                b.Font = New Font("Segoe UI", 9.0!, FontStyle.Bold)
                                b.Margin = New Padding(2, 0, 2, 0)
                                Return b
                            End Function

        _btnSwitchAna = creaBtnSwitch("📋 Anagrafica")
        _btnSwitchImp = creaBtnSwitch("🔗 Impegni")
        _btnSwitchOdp = creaBtnSwitch("🏭 ODP")

        AddHandler _btnSwitchAna.Click, Sub(s, ev) ImpostaModeFiltri("anagrafica")
        AddHandler _btnSwitchImp.Click, Sub(s, ev) ImpostaModeFiltri("impegni")
        AddHandler _btnSwitchOdp.Click, Sub(s, ev) ImpostaModeFiltri("odp")

        flowSwitch.Controls.Add(_btnSwitchAna)
        flowSwitch.Controls.Add(_btnSwitchImp)
        flowSwitch.Controls.Add(_btnSwitchOdp)
        _pnlFiltriSwitcher.Controls.Add(flowSwitch)

        ' ── Panel filtri Impegni ─────────────────────────────────────
        _pnlFiltriImpegni = New Panel()
        _pnlFiltriImpegni.Dock = DockStyle.Top
        _pnlFiltriImpegni.Height = 62
        _pnlFiltriImpegni.BackColor = Color.FromArgb(245, 248, 255)
        _pnlFiltriImpegni.Padding = New Padding(4, 4, 4, 4)

        Dim tlpImp As New TableLayoutPanel()
        tlpImp.Dock = DockStyle.Fill
        tlpImp.ColumnCount = 7
        tlpImp.RowCount = 1
        tlpImp.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        For i = 0 To 4
            tlpImp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 16.5!))
        Next
        tlpImp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 8.75!))
        tlpImp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 8.75!))

        Dim creaFiltroImp = Function(etichetta As String) As GroupBox
                                Dim gb As New GroupBox()
                                gb.Text = etichetta
                                gb.Dock = DockStyle.Fill
                                gb.Font = New Font("Segoe UI", 7.5!, FontStyle.Regular)
                                gb.Margin = New Padding(2, 0, 2, 0)
                                Dim tb As New TextBox()
                                tb.Dock = DockStyle.Fill
                                tb.Font = New Font("Segoe UI", 9.0!)
                                tb.BorderStyle = BorderStyle.None
                                tb.BackColor = Color.White
                                gb.Controls.Add(tb)
                                Return gb
                            End Function

        Dim gbImpCommessa = creaFiltroImp("Commessa")
        Dim gbImpSotto = creaFiltroImp("Sottocommessa")
        Dim gbImpCodice = creaFiltroImp("Codice art.")
        Dim gbImpDesc = creaFiltroImp("Descrizione")
        Dim gbImpUltimi = creaFiltroImp("Ultimi N.")

        _txtImpCommessa = CType(gbImpCommessa.Controls(0), TextBox)
        _txtImpSottocommessa = CType(gbImpSotto.Controls(0), TextBox)
        _txtImpCodice = CType(gbImpCodice.Controls(0), TextBox)
        _txtImpDescrizione = CType(gbImpDesc.Controls(0), TextBox)
        _txtImpUltimi = CType(gbImpUltimi.Controls(0), TextBox)
        _txtImpUltimi.Text = "200"

        Dim btnImpCerca As New Button()
        btnImpCerca.Text = "🔍 Cerca"
        btnImpCerca.Dock = DockStyle.Fill
        btnImpCerca.BackColor = navy
        btnImpCerca.ForeColor = Color.White
        btnImpCerca.FlatStyle = FlatStyle.Flat
        btnImpCerca.FlatAppearance.BorderColor = navyHover
        btnImpCerca.Font = New Font("Segoe UI", 8.5!, FontStyle.Bold)
        btnImpCerca.Margin = New Padding(2, 2, 2, 2)
        AddHandler btnImpCerca.Click, Sub(s, ev) RicercaImpegni()

        Dim btnImpReset As New Button()
        btnImpReset.Text = "✕ Reset"
        btnImpReset.Dock = DockStyle.Fill
        btnImpReset.BackColor = Color.FromArgb(180, 60, 60)
        btnImpReset.ForeColor = Color.White
        btnImpReset.FlatStyle = FlatStyle.Flat
        btnImpReset.FlatAppearance.BorderColor = Color.FromArgb(140, 40, 40)
        btnImpReset.Font = New Font("Segoe UI", 8.5!, FontStyle.Bold)
        btnImpReset.Margin = New Padding(2, 2, 2, 2)
        AddHandler btnImpReset.Click, Sub(s, ev)
                                          _txtImpCommessa.Clear()
                                          _txtImpSottocommessa.Clear()
                                          _txtImpCodice.Clear()
                                          _txtImpDescrizione.Clear()
                                          _txtImpUltimi.Text = "200"
                                          dgvImpegni.Rows.Clear()
                                          tabImpegni.Text = "Impegni ODP"
                                          SetStatus("Filtri impegni azzerati")
                                      End Sub

        For Each tb As TextBox In {_txtImpCommessa, _txtImpSottocommessa, _txtImpCodice, _txtImpDescrizione, _txtImpUltimi}
            AddHandler tb.KeyDown, Sub(s, ev)
                                       If ev.KeyCode = Keys.Return Then
                                           ev.SuppressKeyPress = True
                                           RicercaImpegni()
                                       End If
                                   End Sub
        Next

        tlpImp.Controls.Add(gbImpCommessa, 0, 0)
        tlpImp.Controls.Add(gbImpSotto, 1, 0)
        tlpImp.Controls.Add(gbImpCodice, 2, 0)
        tlpImp.Controls.Add(gbImpDesc, 3, 0)
        tlpImp.Controls.Add(gbImpUltimi, 4, 0)
        tlpImp.Controls.Add(btnImpCerca, 5, 0)
        tlpImp.Controls.Add(btnImpReset, 6, 0)
        _pnlFiltriImpegni.Controls.Add(tlpImp)

        ' ── Panel filtri ODP ─────────────────────────────────────────
        _pnlFiltriODP = New Panel()
        _pnlFiltriODP.Dock = DockStyle.Top
        _pnlFiltriODP.Height = 62
        _pnlFiltriODP.BackColor = Color.FromArgb(245, 248, 255)
        _pnlFiltriODP.Padding = New Padding(4, 4, 4, 4)

        Dim tlpOdp As New TableLayoutPanel()
        tlpOdp.Dock = DockStyle.Fill
        tlpOdp.ColumnCount = 7
        tlpOdp.RowCount = 1
        tlpOdp.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        For i = 0 To 4
            tlpOdp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 16.5!))
        Next
        tlpOdp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 8.75!))
        tlpOdp.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 8.75!))

        Dim creaFiltroOdp = Function(etichetta As String) As GroupBox
                                Dim gb As New GroupBox()
                                gb.Text = etichetta
                                gb.Dock = DockStyle.Fill
                                gb.Font = New Font("Segoe UI", 7.5!, FontStyle.Regular)
                                gb.Margin = New Padding(2, 0, 2, 0)
                                Dim tb As New TextBox()
                                tb.Dock = DockStyle.Fill
                                tb.Font = New Font("Segoe UI", 9.0!)
                                tb.BorderStyle = BorderStyle.None
                                tb.BackColor = Color.White
                                gb.Controls.Add(tb)
                                Return gb
                            End Function

        Dim gbOdpCommessa = creaFiltroOdp("Commessa")
        Dim gbOdpSotto = creaFiltroOdp("Sottocommessa")
        Dim gbOdpMatricola = creaFiltroOdp("Matricola")
        Dim gbOdpCodice = creaFiltroOdp("Codice art.")
        Dim gbOdpDesc = creaFiltroOdp("Descrizione")

        _txtOdpCommessa = CType(gbOdpCommessa.Controls(0), TextBox)
        _txtOdpSottocommessa = CType(gbOdpSotto.Controls(0), TextBox)
        _txtOdpMatricola = CType(gbOdpMatricola.Controls(0), TextBox)
        _txtOdpCodice = CType(gbOdpCodice.Controls(0), TextBox)
        _txtOdpDescrizione = CType(gbOdpDesc.Controls(0), TextBox)

        Dim btnOdpCerca As New Button()
        btnOdpCerca.Text = "🔍 Cerca"
        btnOdpCerca.Dock = DockStyle.Fill
        btnOdpCerca.BackColor = navy
        btnOdpCerca.ForeColor = Color.White
        btnOdpCerca.FlatStyle = FlatStyle.Flat
        btnOdpCerca.FlatAppearance.BorderColor = navyHover
        btnOdpCerca.Font = New Font("Segoe UI", 8.5!, FontStyle.Bold)
        btnOdpCerca.Margin = New Padding(2, 2, 2, 2)
        AddHandler btnOdpCerca.Click, Sub(s, ev) RicercaODP()

        Dim btnOdpReset As New Button()
        btnOdpReset.Text = "✕ Reset"
        btnOdpReset.Dock = DockStyle.Fill
        btnOdpReset.BackColor = Color.FromArgb(180, 60, 60)
        btnOdpReset.ForeColor = Color.White
        btnOdpReset.FlatStyle = FlatStyle.Flat
        btnOdpReset.FlatAppearance.BorderColor = Color.FromArgb(140, 40, 40)
        btnOdpReset.Font = New Font("Segoe UI", 8.5!, FontStyle.Bold)
        btnOdpReset.Margin = New Padding(2, 2, 2, 2)
        AddHandler btnOdpReset.Click, Sub(s, ev)
                                          _txtOdpCommessa.Clear()
                                          _txtOdpSottocommessa.Clear()
                                          _txtOdpMatricola.Clear()
                                          _txtOdpCodice.Clear()
                                          _txtOdpDescrizione.Clear()
                                          dgvODP.Rows.Clear()
                                          tabODP.Text = "Ordini Produzione"
                                          SetStatus("Filtri ODP azzerati")
                                      End Sub

        For Each tb As TextBox In {_txtOdpCommessa, _txtOdpSottocommessa, _txtOdpMatricola, _txtOdpCodice, _txtOdpDescrizione}
            AddHandler tb.KeyDown, Sub(s, ev)
                                       If ev.KeyCode = Keys.Return Then
                                           ev.SuppressKeyPress = True
                                           RicercaODP()
                                       End If
                                   End Sub
        Next

        tlpOdp.Controls.Add(gbOdpCommessa, 0, 0)
        tlpOdp.Controls.Add(gbOdpSotto, 1, 0)
        tlpOdp.Controls.Add(gbOdpMatricola, 2, 0)
        tlpOdp.Controls.Add(gbOdpCodice, 3, 0)
        tlpOdp.Controls.Add(gbOdpDesc, 4, 0)
        tlpOdp.Controls.Add(btnOdpCerca, 5, 0)
        tlpOdp.Controls.Add(btnOdpReset, 6, 0)
        _pnlFiltriODP.Controls.Add(tlpOdp)

        ' ── Aggiunge a Panel13 (Dock=Top viene messo in cima nell'ordine inverso) ──
        Me.Panel13.Controls.Add(_pnlFiltriImpegni)
        Me.Panel13.Controls.Add(_pnlFiltriODP)
        Me.Panel13.Controls.Add(_pnlFiltriSwitcher)

        _pnlFiltriImpegni.Visible = False
        _pnlFiltriODP.Visible = False

        ImpostaModeFiltri("anagrafica")
    End Sub

    Private Sub ImpostaModeFiltri(mode As String)
        Me.Panel3.Visible = (mode = "anagrafica")
        Me.Panel26.Visible = (mode = "anagrafica")
        If _pnlFiltriImpegni IsNot Nothing Then _pnlFiltriImpegni.Visible = (mode = "impegni")
        If _pnlFiltriODP IsNot Nothing Then _pnlFiltriODP.Visible = (mode = "odp")

        Dim colAttivo As Color = Color.FromArgb(60, 120, 200)
        Dim colNorm As Color = Color.FromArgb(22, 45, 84)
        If _btnSwitchAna IsNot Nothing Then _btnSwitchAna.BackColor = If(mode = "anagrafica", colAttivo, colNorm)
        If _btnSwitchImp IsNot Nothing Then _btnSwitchImp.BackColor = If(mode = "impegni", colAttivo, colNorm)
        If _btnSwitchOdp IsNot Nothing Then _btnSwitchOdp.BackColor = If(mode = "odp", colAttivo, colNorm)
    End Sub

    Private Sub AggiungiBottoneCopiaSqlStatus()
        If Me.Controls.ContainsKey("BtnCopiaSqlStatus") Then Return

        Const BTN_W As Integer = 90

        ' Scollega txbStatus dal Dock e reimposta come Anchor lasciando spazio a destra
        txbStatus.Dock = DockStyle.None
        txbStatus.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        txbStatus.Location = New Point(0, Me.ClientSize.Height - txbStatus.Height)
        txbStatus.Width = Me.ClientSize.Width - BTN_W

        Dim btn As New Button()
        btn.Name = "BtnCopiaSqlStatus"
        btn.Text = "Copia SQL"
        btn.Anchor = AnchorStyles.Bottom Or AnchorStyles.Right
        btn.Size = New Size(BTN_W, txbStatus.Height)
        btn.Location = New Point(Me.ClientSize.Width - BTN_W, Me.ClientSize.Height - txbStatus.Height)
        btn.FlatStyle = FlatStyle.Flat
        btn.BackColor = Color.FromArgb(0, 90, 160)
        btn.ForeColor = Color.White
        btn.Font = New Font("Segoe UI", 7.5!)
        btn.FlatAppearance.BorderColor = Color.FromArgb(0, 60, 120)
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 120, 212)
        btn.Cursor = Cursors.Hand
        Me.Controls.Add(btn)
        btn.BringToFront()
        AddHandler btn.Click, Sub(s, ev)
                                  If _lastSqlImpegni <> "" Then
                                      Clipboard.SetText(_lastSqlImpegni)
                                      SetStatus("SQL copiato negli appunti.")
                                  Else
                                      SetStatus("Nessuna query SQL disponibile.", True)
                                  End If
                              End Sub
    End Sub

    Private Sub AggiungiPanelRicercaImmagine()
        If PanelAI.Controls.ContainsKey("PanelImgRicerca") Then Return

        Dim accentBlue As Color = Color.FromArgb(0, 120, 212)
        Dim accentDark As Color = Color.FromArgb(0, 90, 160)

        ' ── Contenitore principale ──────────────────────────────────────
        Dim pnl As New Panel()
        pnl.Name = "PanelImgRicerca"
        pnl.Dock = DockStyle.Bottom
        pnl.Height = 145
        pnl.BackColor = Color.FromArgb(230, 240, 255)
        pnl.Padding = New Padding(6)

        ' ── Separatore visivo ───────────────────────────────────────────
        Dim sep As New Panel()
        sep.Dock = DockStyle.Top
        sep.Height = 2
        sep.BackColor = accentBlue
        pnl.Controls.Add(sep)

        ' ── Titolino ────────────────────────────────────────────────────
        Dim lblTit As New Label()
        lblTit.Text = "Ricerca per immagine"
        lblTit.Dock = DockStyle.Top
        lblTit.Height = 20
        lblTit.Font = New Font("Segoe UI", 8.25!, FontStyle.Bold)
        lblTit.ForeColor = accentDark
        lblTit.TextAlign = ContentAlignment.MiddleLeft
        pnl.Controls.Add(lblTit)

        ' ── PictureBox ──────────────────────────────────────────────────
        _picRicercaImg.SizeMode = PictureBoxSizeMode.Zoom
        _picRicercaImg.BorderStyle = BorderStyle.FixedSingle
        _picRicercaImg.BackColor = Color.White
        _picRicercaImg.Width = 110
        _picRicercaImg.Height = 100
        _picRicercaImg.Location = New Point(6, 26)
        _picRicercaImg.Cursor = Cursors.Hand
        _picRicercaImg.TabStop = True
        pnl.Controls.Add(_picRicercaImg)
        AddHandler _picRicercaImg.Click, AddressOf PicRicercaImg_Click

        ' ── Hint sul PictureBox ─────────────────────────────────────────
        _lblPastaHint.Text = "Clicca qui" & vbCrLf & "e premi Ctrl+V" & vbCrLf & "per incollare"
        _lblPastaHint.AutoSize = False
        _lblPastaHint.Size = New Size(110, 100)
        _lblPastaHint.Location = New Point(6, 26)
        _lblPastaHint.TextAlign = ContentAlignment.MiddleCenter
        _lblPastaHint.ForeColor = Color.Gray
        _lblPastaHint.Font = New Font("Segoe UI", 8.0!, FontStyle.Italic)
        _lblPastaHint.BackColor = Color.Transparent
        _lblPastaHint.Cursor = Cursors.Hand
        pnl.Controls.Add(_lblPastaHint)
        AddHandler _lblPastaHint.Click, AddressOf PicRicercaImg_Click

        ' ── Pulsante cerca per immagine ─────────────────────────────────
        _btnCercaImmagine.Text = "Cerca per immagine"
        _btnCercaImmagine.Location = New Point(124, 26)
        _btnCercaImmagine.Size = New Size(145, 36)
        _btnCercaImmagine.FlatStyle = FlatStyle.Flat
        _btnCercaImmagine.BackColor = accentBlue
        _btnCercaImmagine.ForeColor = Color.White
        _btnCercaImmagine.Font = New Font("Segoe UI", 8.5!, FontStyle.Bold)
        _btnCercaImmagine.FlatAppearance.BorderColor = accentDark
        _btnCercaImmagine.Enabled = False
        pnl.Controls.Add(_btnCercaImmagine)
        AddHandler _btnCercaImmagine.Click, AddressOf BtnCercaImmagine_Click

        ' ── Label "Svuota" ──────────────────────────────────────────────
        Dim btnSvuota As New Button()
        btnSvuota.Text = "Svuota"
        btnSvuota.Location = New Point(124, 70)
        btnSvuota.Size = New Size(145, 26)
        btnSvuota.FlatStyle = FlatStyle.Flat
        btnSvuota.BackColor = Color.FromArgb(230, 240, 255)
        btnSvuota.ForeColor = accentDark
        btnSvuota.Font = New Font("Segoe UI", 8.0!)
        btnSvuota.FlatAppearance.BorderColor = Color.FromArgb(180, 200, 230)
        pnl.Controls.Add(btnSvuota)
        AddHandler btnSvuota.Click, Sub(s, ev)
                                        _picRicercaImg.Image = Nothing
                                        _lblPastaHint.Visible = True
                                        _btnCercaImmagine.Enabled = False
                                    End Sub

        PanelAI.Controls.Add(pnl)
        pnl.BringToFront()
        PanelAI_Bottom.BringToFront()
    End Sub

    Private Sub BRB_Click(sender As Object, e As EventArgs) Handles BRB.Enter
        visualizzazione = "BRB"
    End Sub

    Private Sub TIRELLI_Click(sender As Object, e As EventArgs) Handles Tirelli.Enter
        visualizzazione = "TIRELLI"
    End Sub

    Private Sub Button_CERCA_Click_1(sender As Object, e As EventArgs) Handles Button_CERCA.Click
        If visualizzazione = "TIRELLI" Then
            Dim codGruppo As String = If(ComboBox1.SelectedIndex > 0, Elenco_gruppi(ComboBox1.SelectedIndex), "")
            cerca(TextBox_codice_SAP_RICERCA.Text.ToUpper, TextBox_Descrizione_ricerca.Text.ToUpper, TextBox3.Text.ToUpper, TextBox_disegno_SAP.Text.ToUpper, TextBox6.Text.ToUpper, codGruppo, ComboBox4.Text.ToUpper, TextBox2.Text.ToUpper, ComboBox6.Text.ToUpper, TextBox4.Text.ToUpper)
        Else
            cerca_BRB()
        End If
    End Sub

    Private Sub TextBox_disegno_SAP_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox_disegno_SAP.TextChanged
        filtro_disegno = CostruisciCondizioneSAP("T0.[u_disegno]", TextBox_disegno_SAP.Text)
        AggiornaEvidenziazioneFiltri()
    End Sub

    Private Sub TextBox6_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        filtro_fam_disegno = CostruisciCondizioneSAP("t0.u_famiglia_disegno", TextBox6.Text)
        AggiornaEvidenziazioneFiltri()
    End Sub

    Private Sub DataGridView_SAP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_SAP.CellContentClick
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = DataGridView_SAP.Columns.IndexOf(Codice) Then
                Magazzino.Codice_SAP = DataGridView_SAP.Rows(e.RowIndex).Cells(columnName:="Codice").Value
                Magazzino.Show()
                Magazzino.BringToFront()
                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)

            ElseIf e.ColumnIndex = DataGridView_SAP.Columns.IndexOf(Disegno_) Then
                Magazzino.visualizza_disegno(DataGridView_SAP.Rows(e.RowIndex).Cells(columnName:="Disegno_").Value)

            ElseIf DataGridView_SAP.Columns.Contains("Immagine") AndAlso
                   e.ColumnIndex = DataGridView_SAP.Columns("Immagine").Index Then
                Dim disegnoCell = DataGridView_SAP.Rows(e.RowIndex).Cells("Disegno_").Value?.ToString()
                If Not String.IsNullOrWhiteSpace(disegnoCell) Then
                    Form_visualizza_picture.Show()
                    Magazzino.visualizza_picture(disegnoCell, Form_visualizza_picture.PictureBox1)
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView_SAP_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_SAP.CellFormatting
        If e.RowIndex < 0 Then Return
        Dim val = DataGridView_SAP.Rows(e.RowIndex).Cells("Attivo").Value?.ToString()
        If val = "Y" OrElse val = "I" Then
            DataGridView_SAP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightSlateGray
            DataGridView_SAP.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.DimGray
        Else
            DataGridView_SAP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Empty
            DataGridView_SAP.Rows(e.RowIndex).DefaultCellStyle.ForeColor = Color.Empty
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Codice_Tirelli) Then
                Magazzino.Codice_SAP = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice_tirelli").Value
                Magazzino.Show()
                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)
            End If
        End If
    End Sub

    Private Sub TextBox_Descrizione_ricerca_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox_Descrizione_ricerca.TextChanged
        If visualizzazione = "TIRELLI" Then
            filtro_descrizione = CostruisciCondizioneSAP("t0.itemname", TextBox_Descrizione_ricerca.Text)
        Else
            filtro_descrizione = CostruisciCondizioneOrSAP(
                New String() {"t10.descrizione_brb", "t10.itemname"},
                TextBox_Descrizione_ricerca.Text)
        End If
        AggiornaEvidenziazioneFiltri()
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If visualizzazione = "TIRELLI" Then
            filtro_descrizione_supp = CostruisciCondizioneSAP("T0.[FrgnName]", TextBox3.Text)
        Else
            filtro_descrizione_supp = CostruisciCondizioneOrSAP(
                New String() {"t10.descrizione_supp", "t10.frgnname"},
                TextBox3.Text)
        End If
        AggiornaEvidenziazioneFiltri()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        filtro_catalogo = CostruisciCondizioneSAP("T0.[suppcatnum]", TextBox2.Text)
        AggiornaEvidenziazioneFiltri()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Button20_Click_1(sender As Object, e As EventArgs) Handles Button20.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub DataGridView_SAP_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_SAP.CellClick

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        AggiornaEvidenziazioneFiltri()
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        AggiornaEvidenziazioneFiltri()
    End Sub

    Private Sub TextBox_codice_SAP_RICERCA_TextChanged(sender As Object, e As EventArgs) Handles TextBox_codice_SAP_RICERCA.TextChanged
        AggiornaEvidenziazioneFiltri()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        AggiornaEvidenziazioneFiltri()
    End Sub

    ' ──────────────────────────────────────────────────────────────────
    '  FUNZIONI FILTRO — logica asterischi
    '  *testo    → LIKE 'testo%'   (inizia con)
    '  testo*    → LIKE '%testo'   (finisce con)
    '  *testo*   → LIKE '%testo%'  (contiene, = senza asterischi)
    '  t1*t2     → LIKE '%t1%' AND LIKE '%t2%'  (multi-termine)
    '  **t1**t2  → idem con doppio separatore
    ' ──────────────────────────────────────────────────────────────────

    ''' <summary>Condizione WHERE SQL Server per un singolo campo.</summary>
    Private Function CostruisciCondizioneSAP(campo As String, testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then Return ""
        Dim t As String = testo.Trim()

        ' Inizia con: un solo * all'inizio, nessun altro *
        If t.StartsWith("*") AndAlso t.IndexOf("*"c, 1) < 0 Then
            Return " and " & campo & " LIKE '" & t.Substring(1) & "%'"
        End If
        ' Finisce con: un solo * alla fine
        If t.EndsWith("*") AndAlso t.LastIndexOf("*"c, t.Length - 2) < 0 Then
            Return " and " & campo & " LIKE '%" & t.Substring(0, t.Length - 1) & "'"
        End If
        ' Multi-termine o contiene
        Dim terms = t.Split({"*"c}, StringSplitOptions.RemoveEmptyEntries)
        Return String.Join("", terms.Where(Function(x) x.Trim() <> "").Select(Function(x) " and " & campo & " LIKE '%" & x.Trim() & "%'"))
    End Function

    ''' <summary>Condizione WHERE SQL Server con OR su più campi (es. BRB).</summary>
    Private Function CostruisciCondizioneOrSAP(campi As String(), testo As String) As String
        If String.IsNullOrWhiteSpace(testo) OrElse campi.Length = 0 Then Return ""
        Dim t As String = testo.Trim()

        Dim OredCond = Function(likeExpr As String) As String
                           Return " and (" & String.Join(" or ", campi.Select(Function(c) c & " LIKE '" & likeExpr & "'")) & ")"
                       End Function

        If t.StartsWith("*") AndAlso t.IndexOf("*"c, 1) < 0 Then Return OredCond(t.Substring(1) & "%")
        If t.EndsWith("*") AndAlso t.LastIndexOf("*"c, t.Length - 2) < 0 Then Return OredCond("%" & t.Substring(0, t.Length - 1))

        Dim terms = t.Split({"*"c}, StringSplitOptions.RemoveEmptyEntries)
        Return String.Join("", terms.Where(Function(x) x.Trim() <> "").Select(Function(x) OredCond("%" & x.Trim() & "%")))
    End Function

    ''' <summary>Condizione WHERE AS400/OPENQUERY per un singolo campo (escape '' per le virgolette).</summary>
    Private Function CostruisciCondizioneAS400(campo As String, testo As String) As String
        If String.IsNullOrWhiteSpace(testo) Then Return ""
        Dim t As String = testo.Trim().ToUpper()

        If t.StartsWith("*") AndAlso t.IndexOf("*"c, 1) < 0 Then
            Return " and upper(" & campo & ") LIKE ''" & t.Substring(1) & "%''"
        End If
        If t.EndsWith("*") AndAlso t.LastIndexOf("*"c, t.Length - 2) < 0 Then
            Return " and upper(" & campo & ") LIKE ''%" & t.Substring(0, t.Length - 1) & "''"
        End If
        Dim terms = t.Split({"*"c}, StringSplitOptions.RemoveEmptyEntries)
        Return String.Join("", terms.Where(Function(x) x.Trim() <> "").Select(Function(x) " and upper(" & campo & ") LIKE ''%" & x.Trim() & "%''"))
    End Function

    ' ──────────────────────────────────────────────────────────────────
    '  PANNELLO AI
    ' ──────────────────────────────────────────────────────────────────

    Private Sub BtnAiInvia_Click(sender As Object, e As EventArgs) Handles BtnAiInvia.Click
        InviaMessaggioAI()
    End Sub

    Private Sub TxtAiInput_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtAiInput.KeyDown
        ' Enter invia, Shift+Enter va a capo
        If e.KeyCode = Keys.Enter AndAlso Not e.Shift Then
            e.SuppressKeyPress = True
            InviaMessaggioAI()
        End If
    End Sub

    Private Async Sub InviaMessaggioAI()
        Dim domanda = TxtAiInput.Text.Trim()
        If String.IsNullOrEmpty(domanda) Then Return

        Dim det = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO)
        Dim nomeUtente = (det.cognome & " " & det.nome).Trim()
        If String.IsNullOrEmpty(nomeUtente) Then nomeUtente = Homepage.ID_SALVATO

        Dim conferma = MessageBox.Show(
            "La chiamata all'AI è un servizio a pagamento e verrà addebitata all'azienda." & vbCrLf &
            "La domanda verrà registrata nel sistema a nome di: " & nomeUtente & vbCrLf & vbCrLf &
            "Vuoi procedere?",
            "Conferma chiamata AI",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

        If conferma = DialogResult.No Then Return

        AppendChat("Tu:  " & domanda, False)
        TxtAiInput.Clear()
        BtnAiInvia.Enabled = False
        BtnAiInvia.Text = "..."
        txbStatus.Text = "AI in elaborazione..."
        txbStatus.ForeColor = Color.FromArgb(130, 100, 0)
        txbStatus.BackColor = Color.FromArgb(255, 250, 180)
        txbStatus.Refresh()

        Try
            Dim risposta = Await ChiediAIAsync(domanda)
            AppendChat("AI:  " & risposta, True)
            SalvaChiamataApi(nomeUtente, domanda, risposta)
        Catch ex As Exception
            AppendChat("[Errore] " & ex.Message, True)
            SetStatus("Errore AI: " & ex.Message, True)
        Finally
            txbStatus.BackColor = SystemColors.Window
            BtnAiInvia.Enabled = True
            BtnAiInvia.Text = "Invia  (Invio)"
        End Try
    End Sub

    Private Sub SalvaChiamataApi(nomeUtente As String, domanda As String, risposta As String)
        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Dim cmdCreate As New SqlCommand(
                    "IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'Chiamate_Api')
                     CREATE TABLE Chiamate_Api (
                         ID INT IDENTITY(1,1) PRIMARY KEY,
                         DataOra DATETIME NOT NULL,
                         Utente NVARCHAR(200) NOT NULL,
                         Domanda NVARCHAR(MAX) NOT NULL,
                         Risposta NVARCHAR(MAX) NOT NULL
                     )", cnn)
                cmdCreate.ExecuteNonQuery()

                Dim cmdInsert As New SqlCommand(
                    "INSERT INTO Chiamate_Api (DataOra, Utente, Domanda, Risposta) VALUES (@DataOra, @Utente, @Domanda, @Risposta)", cnn)
                cmdInsert.Parameters.AddWithValue("@DataOra", DateTime.Now)
                cmdInsert.Parameters.AddWithValue("@Utente", nomeUtente)
                cmdInsert.Parameters.AddWithValue("@Domanda", domanda)
                cmdInsert.Parameters.AddWithValue("@Risposta", risposta)
                cmdInsert.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            ' Logging fallito silenziosamente — non bloccare l'utente per un errore di tracciamento
        End Try
    End Sub

    ''' <summary>
    ''' Chiama l'API Anthropic, interpreta la risposta JSON e gestisce l'azione.
    ''' Restituisce il testo "messaggio" da mostrare in chat.
    ''' Se azione="cerca" precompila i filtri ed esegue la ricerca.
    ''' </summary>
    Private Async Function ChiediAIAsync(domanda As String) As Threading.Tasks.Task(Of String)
        Dim keyPath = IO.Path.Combine(Application.StartupPath, "anthropic_key.txt")
        If Not IO.File.Exists(keyPath) Then
            Return "Chiave API non trovata. Crea il file 'anthropic_key.txt' nella cartella dell'applicazione."
        End If
        Dim apiKey = IO.File.ReadAllText(keyPath).Trim()

        _aiHistory.Add(New JObject From {{"role", "user"}, {"content", domanda}})

        Dim requestBody = New JObject From {
            {"model", "claude-haiku-4-5-20251001"},
            {"max_tokens", 1500},
            {"system", AI_SYSTEM_PROMPT},
            {"messages", _aiHistory}
        }

        Using client As New HttpClient()
            client.DefaultRequestHeaders.Add("x-api-key", apiKey)
            client.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01")
            client.Timeout = TimeSpan.FromSeconds(60)

            Dim content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
            Dim response = Await client.PostAsync("https://api.anthropic.com/v1/messages", content)
            Dim body = Await response.Content.ReadAsStringAsync()

            If Not response.IsSuccessStatusCode Then
                Throw New Exception("Errore API " & CInt(response.StatusCode) & ": " & body)
            End If

            Dim apiParsed = JObject.Parse(body)
            Dim tokenVal = apiParsed.SelectToken("content[0].text")
            Dim rawTesto As String = If(tokenVal IsNot Nothing, tokenVal.ToString().Trim(), "(risposta vuota)")

            _aiHistory.Add(New JObject From {{"role", "assistant"}, {"content", rawTesto}})

            ' Interpreta il JSON di risposta del modello
            Return InterpretaRispostaAI(rawTesto)
        End Using
    End Function

    ''' <summary>
    ''' Analizza il JSON restituito dal modello.
    ''' Se azione=cerca, precompila i filtri e avvia la ricerca (su UI thread).
    ''' Restituisce il testo messaggio da mostrare.
    ''' </summary>
    Private Function InterpretaRispostaAI(rawTesto As String) As String
        Try
            ' Rimuove eventuali blocchi ```json ... ``` lasciati dal modello
            Dim jsonPulito = rawTesto
            Dim mdStart = jsonPulito.IndexOf("{")
            Dim mdEnd = jsonPulito.LastIndexOf("}")
            If mdStart >= 0 AndAlso mdEnd > mdStart Then
                jsonPulito = jsonPulito.Substring(mdStart, mdEnd - mdStart + 1)
            End If

            Dim risposta = JObject.Parse(jsonPulito)
            Dim azione = If(risposta("azione") IsNot Nothing, risposta("azione").ToString().ToLower(), "info")
            Dim messaggio = If(risposta("messaggio") IsNot Nothing, risposta("messaggio").ToString(), rawTesto)

            If azione = "cerca" AndAlso risposta("filtri") IsNot Nothing Then
                ApplicaFiltriAI(CType(risposta("filtri"), JObject))
                Return messaggio & vbCrLf & "[Ricerca avviata — vedi risultati nella griglia]"
            End If

            If azione = "impegni" AndAlso risposta("filtri") IsNot Nothing Then
                EseguiQueryImpegni(CType(risposta("filtri"), JObject))
                Return messaggio & vbCrLf & "[Impegni caricati — vedi risultati nella griglia]"
            End If

            If azione = "odp" AndAlso risposta("filtri") IsNot Nothing Then
                EseguiQueryODP(CType(risposta("filtri"), JObject))
                Return messaggio & vbCrLf & "[ODP caricati — vedi risultati nella griglia]"
            End If

            If azione = "ordina" Then
                Dim colonna = If(risposta("colonna") IsNot Nothing, risposta("colonna").ToString().Trim(), "")
                Dim direzione = If(risposta("direzione") IsNot Nothing, risposta("direzione").ToString().Trim().ToUpper(), "ASC")
                OrdinaGriglia(DataGridView_SAP, colonna, direzione)
                Return messaggio
            End If

            Return messaggio

        Catch ex As Exception
            ' Se il JSON è malformato, mostra il testo grezzo
            Return rawTesto
        End Try
    End Function

    ''' <summary>
    ''' Precompila i campi filtro dall'oggetto JSON "filtri" e avvia la ricerca.
    ''' </summary>
    Private Sub ApplicaFiltriAI(filtri As JObject)
        ' Pulisce tutti i filtri
        TextBox_codice_SAP_RICERCA.Text = ""
        TextBox_Descrizione_ricerca.Text = ""
        TextBox3.Text = ""
        TextBox_disegno_SAP.Text = ""
        TextBox6.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        ComboBox4.Text = ""
        ComboBox6.Text = ""
        _filtroPrezzoMin = Decimal.MinValue
        _filtroPrezzoMax = Decimal.MaxValue
        _filtroStato = ""
        _filtroTipoParte = ""

        ' Applica i valori ricevuti dall'AI (solo quelli non vuoti)
        Dim prendi = Function(chiave As String) As String
                         Dim tok = filtri(chiave)
                         Return If(tok IsNot Nothing, tok.ToString().Trim(), "")
                     End Function

        TextBox_codice_SAP_RICERCA.Text = prendi("codice")
        TextBox_Descrizione_ricerca.Text = prendi("descrizione")
        TextBox3.Text = prendi("desc_supp")
        TextBox_disegno_SAP.Text = prendi("disegno")
        TextBox2.Text = prendi("catalogo")
        TextBox4.Text = prendi("ubicazione")

        ' ComboBox (testo libero per ricerca parziale)
        Dim prod = prendi("produttore")
        If prod <> "" Then ComboBox4.Text = prod

        Dim forn = prendi("fornitore")
        If forn <> "" Then
            ComboBox6.Text = forn
            Dim idx = ComboBox6.FindString(forn)
            If idx >= 0 Then ComboBox6.SelectedIndex = idx
        End If

        Dim grp = prendi("gruppo")
        If grp <> "" Then
            Dim idxGrp = ComboBox1.FindString(grp)
            If idxGrp >= 0 Then
                ComboBox1.SelectedIndex = idxGrp
            Else
                ComboBox1.SelectedIndex = 0
            End If
        Else
            ComboBox1.SelectedIndex = 0
        End If

        ' Filtri numerici e stato
        Dim prezzoMinStr = prendi("prezzo_min")
        Dim prezzoMaxStr = prendi("prezzo_max")
        _filtroPrezzoMin = If(prezzoMinStr <> "" AndAlso Decimal.TryParse(prezzoMinStr, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, Nothing),
                              Decimal.Parse(prezzoMinStr, Globalization.CultureInfo.InvariantCulture),
                              Decimal.MinValue)
        _filtroPrezzoMax = If(prezzoMaxStr <> "" AndAlso Decimal.TryParse(prezzoMaxStr, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, Nothing),
                              Decimal.Parse(prezzoMaxStr, Globalization.CultureInfo.InvariantCulture),
                              Decimal.MaxValue)
        _filtroStato = prendi("stato").ToUpper()
        _filtroTipoParte = prendi("tipo_parte").ToUpper()

        ' Avvia la ricerca
        Button_CERCA.PerformClick()
    End Sub

    Private Sub OrdinaGriglia(dgv As DataGridView, nomeColonna As String, direzione As String)
        ' Mappa alias AI → nome colonna reale nella griglia
        Dim mappa As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"Descrizione", "descrizione"}, {"Nome", "descrizione"},
            {"Costo", "Costo"}, {"Prezzo", "Costo"},
            {"Giacenza", "Giagenza"}, {"Giagenza", "Giagenza"},
            {"Disp", "Disp"}, {"Disponibile", "Disp"},
            {"Fornitore", "Fornitore_preferito"}, {"Produttore", "Produttore"},
            {"Gruppo", "Gruppo_articoli"}, {"Gruppo_articoli", "Gruppo_articoli"},
            {"Codice", "Codice"}
        }
        Dim colReale As String = ""
        mappa.TryGetValue(nomeColonna, colReale)
        If colReale = "" Then colReale = nomeColonna
        If Not dgv.Columns.Contains(colReale) Then
            SetStatus("Colonna '" & nomeColonna & "' non trovata per l'ordinamento.", True)
            Return
        End If
        Dim asc = (direzione <> "DESC")
        ' Estrae le righe, le ordina, le reinserisce
        Dim righe As New List(Of DataGridViewRow)
        For Each r As DataGridViewRow In dgv.Rows
            righe.Add(r)
        Next
        righe = righe.OrderBy(Function(r)
                                  Dim v = r.Cells(colReale).Value
                                  If v Is Nothing OrElse v.ToString() = "" Then Return If(asc, Decimal.MaxValue, Decimal.MinValue)
                                  Dim d As Decimal
                                  If Decimal.TryParse(v.ToString(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then Return d
                                  Return v.ToString()
                              End Function).ToList()
        If Not asc Then righe.Reverse()
        dgv.Rows.Clear()
        For Each r In righe
            dgv.Rows.Add(r)
        Next
        SetStatus("Ordinato per " & colReale & " " & If(asc, "crescente", "decrescente") & " — " & dgv.Rows.Count & " righe")
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        If DataGridView_SAP.Rows.Count = 0 Then
            SetStatus("Nessun dato da esportare.", True)
            Return
        End If
        Using dlg As New SaveFileDialog()
            dlg.Filter = "CSV (Excel)|*.csv"
            dlg.FileName = "Ricerca_UT_" & DateTime.Now.ToString("yyyyMMdd_HHmm") & ".csv"
            If dlg.ShowDialog() <> DialogResult.OK Then Return
            Try
                Using sw As New IO.StreamWriter(dlg.FileName, False, System.Text.Encoding.UTF8)
                    ' Intestazioni (solo colonne visibili, esclusa immagine)
                    Dim headers As New List(Of String)
                    For Each col As DataGridViewColumn In DataGridView_SAP.Columns
                        If col.Visible AndAlso col.Name <> "Immagine" Then
                            headers.Add("""" & col.HeaderText.Replace("""", """""") & """")
                        End If
                    Next
                    sw.WriteLine(String.Join(";", headers))
                    ' Righe visibili
                    For Each row As DataGridViewRow In DataGridView_SAP.Rows
                        If Not row.Visible Then Continue For
                        Dim vals As New List(Of String)
                        For Each col As DataGridViewColumn In DataGridView_SAP.Columns
                            If col.Visible AndAlso col.Name <> "Immagine" Then
                                Dim v = If(row.Cells(col.Name).Value?.ToString(), "")
                                vals.Add("""" & v.Replace("""", """""") & """")
                            End If
                        Next
                        sw.WriteLine(String.Join(";", vals))
                    Next
                End Using
                SetStatus("Esportato: " & dlg.FileName)
            Catch ex As Exception
                SetStatus("Errore export: " & ex.Message, True)
            End Try
        End Using
    End Sub

    Private Sub AppendChat(testo As String, isAI As Boolean)
        RtbChat.SelectionStart = RtbChat.TextLength
        RtbChat.SelectionLength = 0
        If isAI Then
            RtbChat.SelectionColor = Color.FromArgb(22, 45, 84)
            RtbChat.SelectionFont = New Font(RtbChat.Font, FontStyle.Regular)
        Else
            RtbChat.SelectionColor = Color.FromArgb(50, 50, 50)
            RtbChat.SelectionFont = New Font(RtbChat.Font, FontStyle.Bold)
        End If
        RtbChat.AppendText(testo & vbCrLf & vbCrLf)
        RtbChat.SelectionStart = RtbChat.TextLength
        RtbChat.ScrollToCaret()
    End Sub

    Private Sub BtnAiPulisci_Click(sender As Object, e As EventArgs) Handles BtnAiPulisci.Click
        _aiHistory = New JArray()
        RtbChat.Clear()
    End Sub

    Private Sub BtnStorico_Click(sender As Object, e As EventArgs) Handles BtnStorico.Click
        MostraStoricoChiamate()
    End Sub

    Private Sub MostraStoricoChiamate()
        Dim dt As New DataTable
        Try
            Using cnn As New SqlConnection(Homepage.sap_tirelli)
                cnn.Open()
                Dim cmd As New SqlCommand(
                    "SELECT TOP 200 DataOra, Utente, Domanda, Risposta
                     FROM Chiamate_Api
                     ORDER BY DataOra DESC", cnn)
                Using da As New SqlDataAdapter(cmd)
                    da.Fill(dt)
                End Using
            End Using
        Catch ex As Exception
            MessageBox.Show("Errore caricamento storico: " & ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

        ' Finestra di dialogo con griglia
        Dim frm As New Form
        frm.Text = "Storico chiamate AI (" & dt.Rows.Count & " ultime)"
        frm.Size = New Size(1200, 650)
        frm.StartPosition = FormStartPosition.CenterParent
        frm.MinimumSize = New Size(800, 400)

        Dim dgv As New DataGridView
        dgv.Dock = DockStyle.Fill
        dgv.ReadOnly = True
        dgv.AllowUserToAddRows = False
        dgv.RowHeadersVisible = False
        dgv.AutoGenerateColumns = False
        dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 248, 255)
        dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv.EnableHeadersVisualStyles = False
        dgv.ColumnHeadersHeight = 26
        dgv.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.DisplayedCells

        Dim addCol = Sub(nome As String, header As String, w As Integer, wrap As Boolean)
                         Dim c As New DataGridViewTextBoxColumn
                         c.Name = nome
                         c.HeaderText = header
                         c.DataPropertyName = nome
                         c.Width = w
                         c.ReadOnly = True
                         If wrap Then c.DefaultCellStyle.WrapMode = DataGridViewTriState.True
                         dgv.Columns.Add(c)
                     End Sub

        addCol("DataOra", "Data / Ora", 130, False)
        addCol("Utente", "Utente", 120, False)
        addCol("Domanda", "Domanda", 360, True)
        addCol("Risposta", "Risposta", 500, True)

        ' Formatta data
        For Each row As DataRow In dt.Rows
            If Not IsDBNull(row("DataOra")) Then
                row("DataOra") = CDate(row("DataOra")).ToString("dd/MM/yyyy HH:mm:ss")
            End If
        Next

        dgv.DataSource = dt

        frm.Controls.Add(dgv)
        frm.ShowDialog(Me)
    End Sub

    ' ──────────────────────────────────────────────────────────────────
    '  IMMAGINI NELLA GRIGLIA
    ' ──────────────────────────────────────────────────────────────────

    Private Sub CaricaImmagineRiga(rowIdx As Integer, disegno As String)
        If String.IsNullOrWhiteSpace(disegno) Then Return
        Dim percorso As String = Homepage.percorso_disegni_generico & "PNG no sfondo\" & disegno & ".PNG"
        If Not IO.File.Exists(percorso) Then Return
        Try
            Using fs As New IO.FileStream(percorso, IO.FileMode.Open, IO.FileAccess.Read)
                Using tmp As Image = Image.FromStream(fs)
                    DataGridView_SAP.Rows(rowIdx).Cells("Immagine").Value = New Bitmap(tmp)
                End Using
            End Using
        Catch
        End Try
    End Sub

    ' ──────────────────────────────────────────────────────────────────
    '  RICERCA PER IMMAGINE — paste e invio a Claude Vision
    ' ──────────────────────────────────────────────────────────────────

    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        If keyData = (Keys.Control Or Keys.V) Then
            If _picRicercaImg IsNot Nothing AndAlso _picRicercaImg.Focused Then
                PicRicercaImg_IncollaClipboard()
                Return True
            End If
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Private Sub PicRicercaImg_Click(sender As Object, e As EventArgs)
        _picRicercaImg.Focus()
        PicRicercaImg_IncollaClipboard()
    End Sub

    Private Sub PicRicercaImg_IncollaClipboard()
        If Clipboard.ContainsImage() Then
            _picRicercaImg.Image = Clipboard.GetImage()
            _lblPastaHint.Visible = False
            _btnCercaImmagine.Enabled = True
        ElseIf Clipboard.ContainsData("PNG") OrElse Clipboard.ContainsData("Bitmap") Then
            _picRicercaImg.Image = DirectCast(Clipboard.GetData(DataFormats.Bitmap), Image)
            _lblPastaHint.Visible = False
            _btnCercaImmagine.Enabled = True
        Else
            AppendChat("AI:  Nessuna immagine negli appunti. Copia prima uno screenshot con Stamp o con lo strumento di cattura.", True)
        End If
    End Sub

    Private Async Sub BtnCercaImmagine_Click(sender As Object, e As EventArgs)
        If _picRicercaImg.Image Is Nothing Then Return
        _btnCercaImmagine.Enabled = False
        _btnCercaImmagine.Text = "..."
        AppendChat("Tu:  [ricerca per immagine]", False)
        txbStatus.Text = "AI analizza immagine..."
        txbStatus.ForeColor = Color.FromArgb(130, 100, 0)
        txbStatus.BackColor = Color.FromArgb(255, 250, 180)
        txbStatus.Refresh()
        Try
            Dim imgBase64 As String
            Using ms As New IO.MemoryStream()
                _picRicercaImg.Image.Save(ms, Imaging.ImageFormat.Png)
                imgBase64 = Convert.ToBase64String(ms.ToArray())
            End Using
            Dim risposta = Await ChiediAIConImmagineAsync(imgBase64)
            AppendChat("AI:  " & risposta, True)
        Catch ex As Exception
            AppendChat("[Errore] " & ex.Message, True)
        Finally
            txbStatus.BackColor = SystemColors.Window
            _btnCercaImmagine.Enabled = True
            _btnCercaImmagine.Text = "Cerca per immagine"
        End Try
    End Sub

    Private Async Function ChiediAIConImmagineAsync(imgBase64 As String) As Threading.Tasks.Task(Of String)
        Dim keyPath = IO.Path.Combine(Application.StartupPath, "anthropic_key.txt")
        If Not IO.File.Exists(keyPath) Then Return "Chiave API non trovata."
        Dim apiKey = IO.File.ReadAllText(keyPath).Trim()

        Dim userContent = New JArray(
            New JObject From {
                {"type", "image"},
                {"source", New JObject From {
                    {"type", "base64"},
                    {"media_type", "image/png"},
                    {"data", imgBase64}
                }}
            },
            New JObject From {
                {"type", "text"},
                {"text", "Analizza questa immagine di un componente meccanico/articolo. Identifica di che tipo di pezzo si tratta e imposta i filtri per trovarlo nel database Tirelli. Usa descrizione, produttore, gruppo articoli e gli altri campi appropriati."}
            }
        )

        Dim requestBody = New JObject From {
            {"model", "claude-haiku-4-5-20251001"},
            {"max_tokens", 1500},
            {"system", AI_SYSTEM_PROMPT},
            {"messages", New JArray(New JObject From {{"role", "user"}, {"content", userContent}})}
        }

        Using client As New HttpClient()
            client.DefaultRequestHeaders.Add("x-api-key", apiKey)
            client.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01")
            client.Timeout = TimeSpan.FromSeconds(60)
            Dim content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
            Dim response = Await client.PostAsync("https://api.anthropic.com/v1/messages", content)
            Dim body = Await response.Content.ReadAsStringAsync()
            If Not response.IsSuccessStatusCode Then
                Throw New Exception("Errore API " & CInt(response.StatusCode) & ": " & body)
            End If
            Dim apiParsed = JObject.Parse(body)
            Dim tokenVal = apiParsed.SelectToken("content[0].text")
            Dim rawTesto As String = If(tokenVal IsNot Nothing, tokenVal.ToString().Trim(), "(risposta vuota)")
            _aiHistory.Add(New JObject From {{"role", "user"}, {"content", userContent}})
            _aiHistory.Add(New JObject From {{"role", "assistant"}, {"content", rawTesto}})
            Return InterpretaRispostaAI(rawTesto)
        End Using
    End Function

    ' ──────────────────────────────────────────────────────────────────
    '  QUERY IMPEGNI DI PRODUZIONE
    ' ──────────────────────────────────────────────────────────────────

    Private Sub ImpostaModalitaImpegni(attiva As Boolean)
        If attiva = _modalitaImpegni Then Return
        _modalitaImpegni = attiva

        ' Colonne SOLO per modalità normale → nascoste in impegni
        ' (non includere colonne riutilizzate con header diverso in impegni)
        Dim colNormali = {"Fam_disegno", "Disegno_",
                          "Fornitore_preferito", "Ubic", "Costo", "Attivo", "Immagine"}
        ' Remapping header per modalità impegni
        Dim mappaImpegni As New Dictionary(Of String, String) From {
            {"descrizione", "Descrizione"},
            {"Disegno", "N.ODP"},
            {"Gruppo_articoli", "Commessa"},
            {"Produttore", "Data ODP"},
            {"Catalogo_fornitore", "Sottocommessa"},
            {"Giagenza", "Q.Pianif"},
            {"Mag_a_Wip", "Q.Trasf"},
            {"Disp", "Q.Da trasf"}
        }

        If attiva Then
            _headersOriginali.Clear()
            For Each kv In mappaImpegni
                If DataGridView_SAP.Columns.Contains(kv.Key) Then
                    _headersOriginali(kv.Key) = DataGridView_SAP.Columns(kv.Key).HeaderText
                    DataGridView_SAP.Columns(kv.Key).HeaderText = kv.Value
                End If
            Next
            For Each nome In colNormali
                If DataGridView_SAP.Columns.Contains(nome) Then
                    DataGridView_SAP.Columns(nome).Visible = False
                End If
            Next
        Else
            For Each kv In _headersOriginali
                If DataGridView_SAP.Columns.Contains(kv.Key) Then
                    DataGridView_SAP.Columns(kv.Key).HeaderText = kv.Value
                End If
            Next
            _headersOriginali.Clear()
            For Each nome In colNormali
                If DataGridView_SAP.Columns.Contains(nome) AndAlso nome <> "Fam_disegno" Then
                    DataGridView_SAP.Columns(nome).Visible = True
                End If
            Next
        End If
    End Sub

    Private Sub EseguiQueryImpegni(filtri As JObject)
        If _txtImpCommessa Is Nothing Then Return
        Dim prendi = Function(chiave As String) As String
                         Dim tok = filtri(chiave)
                         Return If(tok IsNot Nothing, tok.ToString().Trim(), "")
                     End Function
        _txtImpCommessa.Text = prendi("commessa").ToUpper()
        _txtImpSottocommessa.Text = prendi("sottocommessa").ToUpper()
        _txtImpCodice.Text = prendi("codice_articolo").ToUpper()
        _txtImpDescrizione.Text = prendi("descrizione_articolo")
        _txtImpUltimi.Text = If(prendi("ultimi_odp") <> "", prendi("ultimi_odp"), "200")
        ImpostaModeFiltri("impegni")
        RicercaImpegni()
    End Sub

    Private Sub RicercaImpegni()
        If dgvImpegni Is Nothing OrElse _txtImpCommessa Is Nothing Then Return

        Dim commessa = _txtImpCommessa.Text.Trim().ToUpper()
        Dim sottocommessa = _txtImpSottocommessa.Text.Trim().ToUpper()
        Dim codiceArt = _txtImpCodice.Text.Trim().ToUpper()
        Dim descArt = _txtImpDescrizione.Text.Trim().ToUpper()
        Dim ultimiStr = _txtImpUltimi.Text.Trim()

        Dim n As Integer = 200
        Dim modoUltimi As Boolean = (ultimiStr <> "" AndAlso ultimiStr <> "0")
        If modoUltimi Then
            If Not Integer.TryParse(ultimiStr, n) Then n = 200
        End If

        If commessa = "" AndAlso sottocommessa = "" AndAlso codiceArt = "" AndAlso descArt = "" AndAlso Not modoUltimi Then
            SetStatus("Specificare almeno un filtro (commessa, sottocommessa, codice, descrizione o ultimi ODP).", True)
            Return
        End If

        Dim whereImp As String = " WHERE 1=1"
        If commessa <> "" Then
            whereImp &= " AND upper(trim(t1.cod_commessa)) = ''" & commessa.Replace("'", "''") & "''"
        End If
        If sottocommessa <> "" Then
            whereImp &= " AND upper(trim(t1.cod_sottocommessa)) = ''" & sottocommessa.Replace("'", "''") & "''"
        End If
        If codiceArt <> "" Then
            whereImp &= " AND upper(trim(t1.codart)) LIKE ''%" & codiceArt.Replace("'", "''") & "%''"
        End If
        If descArt <> "" Then
            whereImp &= " AND upper(t1.itemname) LIKE ''%" & descArt.Replace("'", "''") & "%''"
        End If

        Dim fetchClause As String = ""
        If modoUltimi OrElse codiceArt <> "" OrElse commessa = "" Then
            If n <= 0 OrElse n > 2000 Then n = 200
            fetchClause = " FETCH FIRST " & n & " ROWS ONLY"
        End If

        Dim sqlImp As String = "SELECT
    TRIM(CAST(i.codart AS VARCHAR(50)))            AS Codice,
    i.itemname                                     AS Descrizione,
    TRIM(CAST(i.odp AS VARCHAR(50)))               AS ODP,
    TRIM(CAST(i.status AS VARCHAR(10)))            AS Status,
    TRIM(CAST(i.cod_commessa AS VARCHAR(50)))      AS Commessa,
    TRIM(CAST(i.cod_sottocommessa AS VARCHAR(50))) AS Sottocommessa,
    TRIM(CAST(i.matricola AS VARCHAR(50)))         AS Matricola,
    i.desc_commessa                                AS DescCommessa,
    i.qtapia                                       AS Quantita
FROM OPENQUERY(AS400, '
    SELECT codart, odp, itemname, status, cod_commessa, cod_sottocommessa, matricola, desc_commessa, dtasca, qtapia
    FROM TIR90VIS.JGALIMP t1" & whereImp & " AND documento=''ODP''
    ORDER BY dtasca DESC, codart" & fetchClause & "
') i"

        Dim whereArt As String
        If codiceArt <> "" Then
            whereArt = "upper(trim(code)) LIKE ''%" & codiceArt.Replace("'", "''") & "%''"
        Else
            whereArt = "stat_code = ''A''"
        End If
        Dim sqlArt As String = "SELECT
    TRIM(CAST(code AS VARCHAR(50)))    AS Codice,
    TRIM(CAST(disegno AS VARCHAR(50))) AS Disegno,
    costo_std                          AS Costo
FROM OPENQUERY(AS400, 'SELECT code, disegno, costo_std FROM TIR90VIS.JGALART WHERE " & whereArt & "')"

        Dim sql As String = "SELECT i.ODP, i.Status, i.Commessa, i.Sottocommessa, i.Matricola, i.DescCommessa,
    i.Codice, i.Descrizione,
    ISNULL(a.Disegno, '')                       AS Disegno,
    i.Quantita,
    ISNULL(CAST(a.Costo AS DECIMAL(18,6)), 0)   AS Costo
FROM (" & sqlImp & ") i
LEFT JOIN (" & sqlArt & ") a ON i.Codice = a.Codice"

        _lastSqlImpegni = sql
        dgvImpegni.Rows.Clear()

        Dim parti As New List(Of String)
        If commessa <> "" Then parti.Add("Comm.: " & commessa)
        If sottocommessa <> "" Then parti.Add("Sotto.: " & sottocommessa)
        If codiceArt <> "" Then parti.Add("Cod.: " & codiceArt)
        If descArt <> "" Then parti.Add("Desc.: " & descArt)
        If modoUltimi AndAlso commessa = "" AndAlso codiceArt = "" Then parti.Add("Ultimi " & n)

        Dim Cnn As New SqlConnection(Homepage.sap_tirelli)
        Try
            Cnn.Open()
            Dim cmd As New SqlCommand(sql, Cnn)
            cmd.CommandTimeout = 120
            Dim rdr = cmd.ExecuteReader()
            Do While rdr.Read()
                Dim qta As Decimal = 0
                Decimal.TryParse(rdr("Quantita").ToString(), qta)
                Dim costo As Decimal = 0
                Decimal.TryParse(rdr("Costo").ToString(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, costo)
                Dim costoTot As Decimal = qta * costo
                dgvImpegni.Rows.Add(
                    rdr("ODP"),
                    rdr("Status"),
                    rdr("Commessa"),
                    rdr("Sottocommessa"),
                    rdr("Matricola"),
                    rdr("DescCommessa"),
                    rdr("Codice"),
                    rdr("Descrizione"),
                    rdr("Disegno"),
                    qta.ToString("N2"),
                    costo.ToString("N4"),
                    costoTot.ToString("N2"))
            Loop
            rdr.Close()
            Dim labelFiltri = If(parti.Count > 0, String.Join("  |  ", parti), "ultimi ODP")
            SetStatus("Impegni " & labelFiltri & " — " & dgvImpegni.Rows.Count & " righe")
            If tabImpegni IsNot Nothing Then TabControl1.SelectedTab = tabImpegni
        Catch ex As Exception
            SetStatus("Errore impegni: " & ex.Message, True)
            Try
                Clipboard.SetText(sql)
                AppendChat("AI:  [SQL copiato negli appunti — incollalo per vedere la query completa]", True)
            Catch
            End Try
        Finally
            Cnn.Close()
        End Try
    End Sub

    ' ──────────────────────────────────────────────────────────────────
    '  QUERY ORDINI DI PRODUZIONE (JGALODP)
    ' ──────────────────────────────────────────────────────────────────

    Private Sub ImpostaModalitaODP(attiva As Boolean)
        If attiva = _modalitaODP Then Return
        _modalitaODP = attiva

        Dim mappaODP As New Dictionary(Of String, String) From {
            {"descrizione", "Descrizione"},
            {"Disegno", "N.ODP"},
            {"Gruppo_articoli", "Commessa"},
            {"Produttore", "Data scad"},
            {"Catalogo_fornitore", "Sottocommessa"},
            {"Fornitore_preferito", "Matricola"},
            {"Giagenza", "Q.Pianif"},
            {"Mag_a_Wip", "Q.Res"}
        }
        Dim colNascODP = {"Fam_disegno", "Disegno_", "Ubic", "Disp", "Costo", "Attivo", "Immagine"}

        If attiva Then
            _headersOriginali.Clear()
            For Each kv In mappaODP
                If DataGridView_SAP.Columns.Contains(kv.Key) Then
                    _headersOriginali(kv.Key) = DataGridView_SAP.Columns(kv.Key).HeaderText
                    DataGridView_SAP.Columns(kv.Key).HeaderText = kv.Value
                End If
            Next
            For Each nome In colNascODP
                If DataGridView_SAP.Columns.Contains(nome) Then
                    DataGridView_SAP.Columns(nome).Visible = False
                End If
            Next
        Else
            For Each kv In _headersOriginali
                If DataGridView_SAP.Columns.Contains(kv.Key) Then
                    DataGridView_SAP.Columns(kv.Key).HeaderText = kv.Value
                End If
            Next
            _headersOriginali.Clear()
            For Each nome In colNascODP
                If DataGridView_SAP.Columns.Contains(nome) AndAlso nome <> "Fam_disegno" Then
                    DataGridView_SAP.Columns(nome).Visible = True
                End If
            Next
        End If
    End Sub

    ''' <summary>Chiamato dall'AI: imposta i TextBox filtro e avvia la ricerca.</summary>
    Private Sub EseguiQueryODP(filtri As JObject)
        If _txtOdpCommessa Is Nothing Then Return
        Dim prendi = Function(chiave As String) As String
                         Dim tok = filtri(chiave)
                         Return If(tok IsNot Nothing, tok.ToString().Trim(), "")
                     End Function
        _txtOdpCommessa.Text = prendi("commessa").ToUpper()
        _txtOdpSottocommessa.Text = prendi("sottocommessa").ToUpper()
        _txtOdpMatricola.Text = prendi("matricola").ToUpper()
        _txtOdpCodice.Text = prendi("codice_articolo").ToUpper()
        _txtOdpDescrizione.Text = prendi("descrizione_articolo")
        ImpostaModeFiltri("odp")
        RicercaODP()
    End Sub

    ''' <summary>Legge i TextBox filtro ed esegue la query ODP.</summary>
    Private Sub RicercaODP()
        If dgvODP Is Nothing OrElse _txtOdpCommessa Is Nothing Then Return

        Dim commessa = _txtOdpCommessa.Text.Trim().ToUpper()
        Dim sottocommessa = _txtOdpSottocommessa.Text.Trim().ToUpper()
        Dim matricola = _txtOdpMatricola.Text.Trim().ToUpper()
        Dim codiceArt = _txtOdpCodice.Text.Trim().ToUpper()
        Dim descArt = _txtOdpDescrizione.Text.Trim().ToUpper()

        If commessa = "" AndAlso sottocommessa = "" AndAlso matricola = "" AndAlso codiceArt = "" AndAlso descArt = "" Then
            SetStatus("ODP: specificare almeno un filtro.", True)
            Return
        End If

        Dim whereOdp As String = " WHERE 1=1"
        If commessa <> "" Then
            whereOdp &= " AND upper(trim(t0.cod_commessa)) LIKE ''%" & commessa.Replace("'", "''") & "%''"
        End If
        If sottocommessa <> "" Then
            whereOdp &= " AND upper(trim(t0.cod_sottocommessa)) LIKE ''%" & sottocommessa.Replace("'", "''") & "%''"
        End If
        If matricola <> "" Then
            whereOdp &= " AND upper(trim(t0.matricola)) LIKE ''%" & matricola.Replace("'", "''") & "%''"
        End If
        If codiceArt <> "" Then
            whereOdp &= " AND upper(trim(t0.codart)) LIKE ''%" & codiceArt.Replace("'", "''") & "%''"
        End If
        If descArt <> "" Then
            whereOdp &= " AND upper(t0.dscodart_odp) LIKE ''%" & descArt.Replace("'", "''") & "%''"
        End If

        ' Etichetta filtri attivi per tab e status bar
        Dim parti As New List(Of String)
        If commessa <> "" Then parti.Add("Comm.: " & commessa)
        If sottocommessa <> "" Then parti.Add("Sotto.: " & sottocommessa)
        If matricola <> "" Then parti.Add("Matr.: " & matricola)
        If codiceArt <> "" Then parti.Add("Cod.: " & codiceArt)
        If descArt <> "" Then parti.Add("Desc.: " & descArt)
        Dim labelFiltri As String = String.Join("  |  ", parti)

        Dim sql As String = "SELECT
    trim(t10.numodp)                                                   AS ODP,
    trim(t10.codart)                                                   AS Codice,
    t10.dscodart_odp                                                   AS Descrizione,
    trim(t10.cod_commessa)                                             AS Commessa,
    trim(t10.cod_sottocommessa)                                        AS Sottocommessa,
    trim(t10.matricola)                                                AS Matricola,
    trim(t10.mag_ver)                                                  AS MagVer,
    t10.qta_pia                                                        AS QPia,
    t10.qta_res                                                        AS QRes,
    CASE WHEN t10.data_scad IS NOT NULL AND t10.data_scad > 0
         THEN CONVERT(DATE, CAST(t10.data_scad AS CHAR(8)), 112)
         ELSE NULL END                                                 AS DataScad
FROM OPENQUERY(AS400, '
    SELECT
        trim(t0.numodp) as numodp,
        trim(t0.codart) as codart,
        t0.dscodart_odp,
        trim(t0.cod_commessa) as cod_commessa,
        trim(t0.cod_sottocommessa) as cod_sottocommessa,
        trim(t0.matricola) as matricola,
        trim(t0.mag_ver) as mag_ver,
        t0.qta_pia, t0.qta_res, t0.data_scad
    FROM TIR90VIS.JGALODP t0" & whereOdp & "
    ORDER BY t0.data_scad DESC, t0.numodp
    FETCH FIRST 500 ROWS ONLY
') AS t10"

        _lastSqlImpegni = sql
        dgvODP.Rows.Clear()
        tabODP.Text = "ODP  [" & labelFiltri & "]"
        SetStatus("Ricerca ODP in corso…")
        Application.DoEvents()

        Dim Cnn As New SqlConnection(Homepage.sap_tirelli)
        Try
            Cnn.Open()
            Dim cmd As New SqlCommand(sql, Cnn)
            cmd.CommandTimeout = 60
            Dim rdr = cmd.ExecuteReader()
            Do While rdr.Read()
                Dim dataScad As String = ""
                If Not IsDBNull(rdr("DataScad")) Then
                    dataScad = CDate(rdr("DataScad")).ToString("dd/MM/yyyy")
                End If
                dgvODP.Rows.Add(
                    rdr("ODP"),
                    rdr("Codice"),
                    rdr("Descrizione"),
                    rdr("Commessa"),
                    rdr("Sottocommessa"),
                    rdr("Matricola"),
                    rdr("MagVer"),
                    rdr("QPia"),
                    rdr("QRes"),
                    dataScad)
            Loop
            rdr.Close()
            SetStatus("ODP  " & labelFiltri & "  —  " & dgvODP.Rows.Count & " righe")
            If tabODP IsNot Nothing Then TabControl1.SelectedTab = tabODP
        Catch ex As Exception
            SetStatus("Errore ODP: " & ex.Message, True)
            Try
                Clipboard.SetText(sql)
                AppendChat("AI:  [SQL copiato negli appunti — incollalo per vedere la query completa]", True)
            Catch
            End Try
        Finally
            Cnn.Close()
        End Try
    End Sub

    Private Sub BtnLegenda_Click(sender As Object, e As EventArgs) Handles BtnLegenda.Click
        MostraLegenda()
    End Sub

    Private Sub MostraLegenda()
        Dim msg As String =
            "LEGENDA CRITERI DI RICERCA" & vbCrLf &
            String.Empty.PadRight(40, "─"c) & vbCrLf & vbCrLf &
            "Sintassi nei campi di testo:" & vbCrLf & vbCrLf &
            "  parola          →  contiene 'parola'" & vbCrLf &
            "  *parola         →  inizia con 'parola'" & vbCrLf &
            "  parola*         →  finisce con 'parola'" & vbCrLf &
            "  *parola*        →  contiene 'parola'  (= senza *)" & vbCrLf &
            "  par1*par2       →  contiene 'par1'  E  contiene 'par2'" & vbCrLf &
            "  **p1**p2        →  contiene 'p1'  E  contiene 'p2'" & vbCrLf & vbCrLf &
            "Esempi:" & vbCrLf &
            "  *cuscinetto    →  LIKE 'cuscinetto%'" & vbCrLf &
            "  cuscinetto*    →  LIKE '%cuscinetto'" & vbCrLf &
            "  SKF*6205       →  LIKE '%SKF%' AND LIKE '%6205%'" & vbCrLf & vbCrLf &
            "Nota: le ricerche non distinguono maiuscole/minuscole."

        MessageBox.Show(msg, "Legenda criteri di ricerca",
                        MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

End Class
