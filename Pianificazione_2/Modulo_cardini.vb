Imports System.Data.SqlClient
Imports Npgsql
Imports Tirelli.Scheda_tecnica

Public Class Modulo_cardini
    Inherits UserControl

    ' Inizializza il modulo senza dati pesanti
    Public Sub inizializza_modulo_vuoto(par_commessa As String)
        Label1.Text = par_commessa
        Label2.Text = "Caricamento..."
        DataGridView1.Rows.Clear()
    End Sub

    ' Funzione che carica dati pesanti (senza toccare UI)
    Public Function CaricaDati(par_commessa As String, conn_string As String) As List(Of Object())

        Dim dati As New List(Of Object())

        Dim connString As String = conn_string

        Using conn As New NpgsqlConnection(connString)
            conn.Open()
            Dim STRINGA_QUERY As String = "SELECT 
--J.jobcod,
P.PRJcod,
P.PrjStsUid
,T1.STSCOD
,p.PrjTssTrg,
--T.TSKCOD As PRJ_ATT,
--T.UID AS ATT_KEY,
W.WBSLVLCOD As ATT_WBS,
T.TSKKND,
--substring(W.WBSLVLCOD,1,3) AS Prime_3_WBS,
L.TSKDSC AS ATT_DES,
t.TskStsUid,
--c.CAUMOVTMPCOD As CODCATEGORIA,
--P.UID AS ATT_PRJ_KEY,
---CASE WHEN T.TSKKND = '1' THEN 'MACRO' WHEN T.TSKKND = '0' THEN 'TASK' ELSE 'MILESTONE' END AS ATT_TIPO,
S.STSCOD AS ATT_STATO,
coalesce(SL.STSDSC,'') As ATT_STATO_DESC,
--TD.TSKDCLTSSSTR AS ATT_DTCONSINI,
--TD.TSKDCLTSSEND As ATT_DTCONSFIN,
T.TSKRSDTSSSTR AS ATT_DTPIAINI,
T.TSKRSDTSSEND As ATT_DTPIAFIN,
gptb.BSLTSSSTR as BSL_DTINIZIO,
gptb.BSLTSSEND as BSL_DTFINE,
--CST.CSTCOD AS ATT_TPVINCOLO,
--T.CSTTSS As ATT_DTVINCOLO,
--J.JOBCOD AS ATT_CODCOM,
--T.TSKNOT As ATT_NOTE,
T.TSKRSDTSSEND-gptb.BSLTSSEND as Delta

From PRJTSK T
Left Join PRJTSKDET TD ON TD.TSKUID = T.UID
Left Join PRJTSKLNG L ON L.RECUID = T.UID And L.LNGUID = 1
Left Join ANGWBSLVL W ON T.WBSLVLUID = W.UID
Left Join ANGCAUMOVTMP C ON T.CAUMOVTMPUID = C.UID
Left Join ANGPRJTSKSTS S ON T.TSKSTSUID = S.UID
Left Join ANGPRJTSKSTSLNG SL ON S.UID = SL.RECUID And SL.LNGUID = 1
Left Join PRJ P ON T.PRJUID = P.UID
Left Join ANGPRJCSTTYP CST ON T.CSTUID = CST.UID
Left Join ANGJOB J ON T.JOBUID = J.UID
LEFT JOIN AngPrjSts T1 ON T1.UID=p.PRJSTSUID
-- BASELINE
left join AngPrjBsl apb on (apb.bslcod = 'pre-pianificazione')
left join GenPrjBsl gpb on (gpb.bsluid = apb.uid and gpb.prjuid = p.uid)
left join GenPrjTskBsl gptb on (gptb.genprjbsluid = gpb.uid and gptb.tskuid = t.uid)
-- FINE BASELINE
WHERE P.PRJcod='" & par_commessa & "'
--and W.WBSLVLCOD='1.1.7'
AND T.LOGDEL = 0 
and T.TSKKND = '2' -- = 2 SE SE CARDINE, = 1 SE MACRO
and p.PrjStsUid<>3 AND p.PrjStsUid<>4 AND p.PrjStsUid<>5 -- gestione dello stato del progetto, tolgo quelli chiusi
order by P.PRJcod , W.WBSLVLCOD
--LIMIT 5
--And L.TSKDSC = 'APPROVAZIONE C.O. E LAYOUT PRE-PROGETTO'
--And P.PRJcod = '" & par_commessa & "' and W.WBSLVLCOD like '2%' and L.TSKDSC = 'IMPREVISTI DESIGN REVIEW'-- and lower(L.TSKDSC) like '%progettazione%'
--And L.TSKDSC Like '%APPROVAZ%' --and L.TSKDSC = 'APPROVAZIONE C.O. E LAYOUT PRE-PROGETTO' " ' la tua query completa
            Dim cmd As New NpgsqlCommand(STRINGA_QUERY, conn)
            Dim reader As NpgsqlDataReader = cmd.ExecuteReader()

            Do While reader.Read()
                Dim deltaVal As Object = reader("Delta")
                Dim deltaGiorni As Double = 0

                If TypeOf deltaVal Is TimeSpan Then
                    deltaGiorni = DirectCast(deltaVal, TimeSpan).TotalDays
                ElseIf IsDate(deltaVal) Then
                    deltaGiorni = (CDate(deltaVal) - CDate(reader("BSL_DTFINE"))).TotalDays
                ElseIf IsNumeric(deltaVal) Then
                    deltaGiorni = CDbl(deltaVal)
                End If

                dati.Add(New Object() {
                    reader("PRJcod"),
                    reader("ATT_WBS"),
                    reader("TSKKND"),
                    reader("ATT_DES"),
                    reader("TskStsUid"),
                    reader("ATT_STATO_DESC"),
                    reader("BSL_DTFINE"),
                    reader("ATT_DTPIAFIN"),
                    Math.Round(deltaGiorni, 1)
                })
            Loop
            reader.Close()
        End Using

        Return dati
    End Function






    ' Funzione che aggiorna il DataGridView sul thread UI
    Public Sub VisualizzaDati(dati As List(Of Object()), par_nome_commessa As String)
        DataGridView1.Rows.Clear()
        For Each riga In dati
            DataGridView1.Rows.Add(riga)
        Next
        Label2.Text = par_nome_commessa
    End Sub



    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Dim par_datagridview As DataGridView = DataGridView1

        ' Nomi colonne
        Dim nome_colonna_stato As String = "att_stato_desc"
        Dim nome_colonna_delta As String = "delta"

        ' Gestione colori e stile per la colonna Stato
        Dim valoreStato As String = par_datagridview.Rows(e.RowIndex).Cells(nome_colonna_stato).Value?.ToString().ToLower()

        Select Case valoreStato
            Case "chiuso"
                par_datagridview.Rows(e.RowIndex).Cells(nome_colonna_stato).Style.BackColor = Color.Lime

            Case "annullato", "non applicabile"
                For Each cell As DataGridViewCell In par_datagridview.Rows(e.RowIndex).Cells
                    ' cell.Style.Font = New Font(par_datagridview.Font, FontStyle.Strikeout)
                    cell.Style.ForeColor = Color.Gray
                Next

            Case "rilasciato"
                par_datagridview.Rows(e.RowIndex).Cells(nome_colonna_stato).Style.BackColor = Color.Yellow
        End Select

        ''Gestione colore testo per la colonna Delta
        'Dim valoreDelta As Object = par_datagridview.Rows(e.RowIndex).Cells(nome_colonna_delta).Value

        'If TypeOf valoreDelta Is TimeSpan Then
        '    Dim delta As TimeSpan = DirectCast(valoreDelta, TimeSpan)
        '    If delta < TimeSpan.Zero Then
        '        par_datagridview.Rows(e.RowIndex).Cells(nome_colonna_delta).Style.ForeColor = Color.Lime
        '    ElseIf delta > TimeSpan.Zero Then
        '        par_datagridview.Rows(e.RowIndex).Cells(nome_colonna_delta).Style.ForeColor = Color.Red
        '    End If

        'ElseIf IsNumeric(valoreDelta) Then
        '    Dim deltaNum As Double = CDbl(valoreDelta)
        '    If deltaNum < 0 Then
        '        par_datagridview.Rows(e.RowIndex).Cells(nome_colonna_delta).Style.ForeColor = Color.Lime
        '    ElseIf deltaNum > 0 Then
        '        par_datagridview.Rows(e.RowIndex).Cells(nome_colonna_delta).Style.ForeColor = Color.Red
        '    End If
        'End If

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
