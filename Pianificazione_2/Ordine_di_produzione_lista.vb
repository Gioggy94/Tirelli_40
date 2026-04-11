Imports System.Data.SqlClient
Imports System.IO
Imports System.Security.Policy
Imports Microsoft.Office.Interop



Public Class Ordine_di_produzione_lista

    Public filtro_n_odp As String
    Public filtro_codice As String
    Public filtro_descrizione As String
    Public filtro_commessa As String
    Public filtro_cliente As String
    Public filtro_Fase As String
    Public filtro_produzione As String
    Public filtro_stato As String
    Public filtro_ferretto As String
    Public filtro_commessa_ As String
    Public filtro_sottocommessa As String

    Public isShiftKeyDown As Boolean = False
    Private startIndex As Integer = -1
    Public ID_lotto_di_prelievo As Integer
    Private filtro_n_oc As String
    Private filtro_cliente_OC As String
    Private filtro_causcons_oc As String
    Private filtro_completi As String
    Private filtro_magazzino_destinazione As String
    Private filtro_ferretto_oc As String

    Sub inizializzazione_ordine_di_produzione_lista()
        ID_lotto_di_prelievo = Trova_nuovo_lotto_di_prelievo()
        Txt_DocNum.Text = ID_lotto_di_prelievo
        RIEMPI_Odp(TextBox10.Text)
    End Sub

    Private Sub DataGridView_ODP_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView_ODP.CellMouseDown
        If e.Button = MouseButtons.Left AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If ModifierKeys = Keys.Shift AndAlso startIndex >= 0 Then
                Dim endIndex As Integer = e.RowIndex
                Dim minIndex As Integer = Math.Min(startIndex, endIndex)
                Dim maxIndex As Integer = Math.Max(startIndex, endIndex)

                ' Escludi la riga cliccata dal blocco (verrà gestita dal click normale)
                If endIndex > startIndex Then
                    maxIndex -= 1
                ElseIf endIndex < startIndex Then
                    minIndex += 1
                Else
                    ' Se stessa riga, nessuna azione
                    Return
                End If

                ' Controlla se tutte le righe nell'intervallo sono selezionate
                Dim allSelected As Boolean = True
                For i As Integer = minIndex To maxIndex
                    If Not Convert.ToBoolean(DataGridView_ODP.Rows(i).Cells(0).Value) Then
                        allSelected = False
                        Exit For
                    End If
                Next

                ' Applica lo stato opposto a tutto l'intervallo
                For i As Integer = minIndex To maxIndex
                    DataGridView_ODP.Rows(i).Cells(0).Value = Not allSelected
                Next
            Else
                startIndex = e.RowIndex
            End If
        End If
    End Sub

    Private Sub DataGridView_ODP_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView_ODP.KeyDown
        ' Controlla se il tasto Shift è stato premuto
        isShiftKeyDown = (e.KeyCode = Keys.ShiftKey)
    End Sub

    Private Sub DataGridView_ODP_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView_ODP.KeyUp
        ' Controlla se il tasto Shift è stato rilasciato
        isShiftKeyDown = (e.KeyCode <> Keys.ShiftKey)
    End Sub
    Sub RIEMPI_Odp(par_numero_risultati)
        Dim Cnn1 As New SqlConnection
        DataGridView_ODP.Rows.Clear()

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader

        CMD_SAP_2.Connection = Cnn1

        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_2.CommandText = "select top " & par_numero_risultati & " t0.docnum, t0.itemcode, t1.itemname
, coalesce(t0.u_disegno,'') as 'u_disegno'
,t0.PlannedQty
,t0.CmpltQty
, t0.U_PRG_AZS_Commessa, t2.U_Final_customer_name, coalesce(t3.ResName,t0.u_fase) as 'Resname', t0.U_PRODUZIONE,t0.Status, t0.warehouse , t0.StartDate, case when a.n_codici is null then 0 else a.n_codici end as 'n_codici',case when a.N_da_trasferire is null then 0 else a.n_da_trasferire end as 'N_da_trasferire',case when a.N_da_trasferire is null then 0 else a.n_da_trasferire end - B.[01]-B.[03]-B.FERRETTO-B.SCA- B.CAP2- B.MUT - B.[02] as 'Mancanti',  coalesce(B.[01],0) as '01',coalesce(B.[03],0) as '03',coalesce(B.FERRETTO,0) as 'FERRETTO',coalesce(B.SCA,0) as 'SCA', coalesce(B.CAP2,0) as 'CAP2', coalesce(B.MUT,0) as 'MUT' , coalesce(B.[02],0) as '02'
,coalesce(c.id,'') as 'Lotto'
from owor t0 left join oitm t1 on t0.itemcode=t1.itemcode
left join oitm t2 on t2.itemcode=t0.U_PRG_AZS_Commessa
left join orsc t3 on t3.VisResCode=t0.U_Fase
left join 
    (select t10.docnum, count(t10.itemcode) as 'N_codici', sum(case when t10.U_PRG_WIP_QtaDaTrasf >0 then 1 else 0 end ) as 'N_da_trasferire' , sum(case when t10.onhand>=U_PRG_WIP_QtaDaTrasf  then 1 else 0 end ) as 'N_trasferibili'
    from (select t0.docnum, t1.itemcode, case when t1.U_PRG_WIP_QtaDaTrasf is null then 0 else t1.U_PRG_WIP_QtaDaTrasf end as 'U_PRG_WIP_QtaDaTrasf', case when t2.onhand is null then 0 else t2.onhand end as 'onhand'
            from owor t0 inner join wor1 t1 on t0.DocEntry=t1.docentry
            left join oitw t2 on t2.WhsCode=t1.wareHouse and t2.itemcode=t1.itemcode
            where (t0.status='P' or t0.status='R') and t1.ItemType=4) as t10
            group by t10.docnum) A on a.docnum=t0.docnum
left join
    (select t20.docnum, sum( case when t20.warehouse='01' then N_Trasferibili else 0 end) as '01'
    , sum( case when t20.warehouse='FERRETTO' then N_Trasferibili else 0 end) as 'FERRETTO'
    , sum( case when t20.warehouse='BSCA' then N_Trasferibili else 0 end) as 'BSCA'
    , sum( case when t20.warehouse='03' then N_Trasferibili else 0 end) as '03'
    , sum( case when t20.warehouse='SCA' then N_Trasferibili else 0 end) as 'SCA'
    , sum( case when t20.warehouse='CAP2' then N_Trasferibili else 0 end) as 'CAP2'
    , sum( case when t20.warehouse='MUT' then N_Trasferibili else 0 end) as 'MUT'
    , sum( case when t20.warehouse='02' then N_Trasferibili else 0 end) as '02'
        from (select t10.docnum, t10.wareHouse, sum(case when t10.ONHAND>= t10.U_PRG_WIP_QtaDaTrasf AND t10.U_PRG_WIP_QtaDaTrasf>0 then 1 else 0 end) as 'N_Trasferibili'
            from (select t0.docnum, case when t1.U_PRG_WIP_QtaDaTrasf is null then 0 else t1.U_PRG_WIP_QtaDaTrasf end as 'U_PRG_WIP_QtaDaTrasf', t1.wareHouse, CASE WHEN T2.ONHAND IS NULL THEN 0 ELSE T2.ONHAND END AS 'ONHAND'
                from owor t0 inner join wor1 t1 on t0.DocEntry=t1.docentry
                left join oitw t2 on t2.WhsCode=t1.wareHouse and t2.itemcode=t1.itemcode
                where (t0.status='P' or t0.status='R') and (substring(t1.itemcode,1,1)='0' or substring(t1.itemcode,1,1)='c' or substring(t1.itemcode,1,1)='d' or substring(t1.itemcode,1,1)='f')) as t10
                group by t10.docnum, t10.wareHouse) as t20
        GROUP BY t20.docnum) B on t0.docnum=B.docnum
left join 
    (select max(t0.id) as 'ID', t0.docnum 
    from [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 inner join owor t1 on t0.docnum=t1.docnum
    where t1.status='P' or t1.status='R'
    group by t0.docnum) C on c.docnum=t0.docnum
where " & filtro_stato & " " & filtro_n_odp & " " & filtro_codice & " " & filtro_descrizione & " " & filtro_commessa & " " & filtro_cliente & " " & filtro_produzione & " " & filtro_ferretto & " " & filtro_Fase & filtro_completi & " " & filtro_magazzino_destinazione & " AND T1.ItemName NOT Like '%%ANTICIP%%' 
order by t0.StartDate"

        Else
            ' ── AS400 ── alias allineati ai nomi usati nel Rows.Add ──────────────
            CMD_SAP_2.CommandText = "select
  t10.numodp                                          as 'docnum'
, trim(t10.codart)                                    as 'itemcode'
, trim(t10.disegno)                                   as 'u_disegno'
, t10.dscodart_odp                                    as 'itemname'
, T10.QTA_PIA                                         as 'plannedqty'
, T10.QTA_PIA - T10.QTA_RES                           as 'CmpltQty'
, trim(t10.matricola)                                       as 'U_PRG_AZS_Commessa'
, trim(t10.cod_sottocommessa)                                       as 'Sottocommessa'
, trim(t10.commessa_odp)                                       as 'Commessa'
, t10.cliente                                         as 'U_Final_customer_name'
, ''                                                  as 'ResName'
, ''                                                  as 'U_PRODUZIONE'
, T10.PIANIFICATO                                     as 'status'
, ''                                                  as 'Lotto'
, t10.mag_ver                                         as 'warehouse'
, CONVERT(DATETIME, CAST(t10.data_iniz AS CHAR(8)), 112) as 'startdate'
, 0                                                   as 'n_codici'
, 0                                                   as 'N_da_trasferire'
, 0                                                   as 'Mancanti'
, 0                                                   as '01'
, 0                                                   as 'FERRETTO'
, 0                                                   as '03'
, 0                                                   as 'SCA'
, 0                                                   as 'MUT'
, 0                                                   as 'CAP2'
, 0                                                   as '02'
FROM OPENQUERY([AS400], '
    SELECT numodp, qta_pia, qta_res, pianificato, codart, dscodart_odp, disegno,
           cod_commessa, matricola, cod_sottocommessa,
           commessa as commessa_odp,
           cliente, data_iniz, data_immissione, data_scad,
           mag_ver, posizione
    FROM S786FAD1.TIR90VIS.JGALODP
where 0=0 " & filtro_stato & " " & filtro_n_odp & " " & filtro_codice & " " & filtro_descrizione & " " & filtro_commessa & " " & filtro_cliente & " " & filtro_produzione & " " & filtro_ferretto & " " & filtro_Fase & filtro_completi & " " & filtro_magazzino_destinazione & filtro_commessa_ & filtro_sottocommessa & "
    limit " & par_numero_risultati & "
') AS t10
order by t10.numodp DESC"
        End If

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            ' Calcolo % completamento: gestione divisione per zero
            Dim pct As Double = 0
            Dim n_cod As Integer = cmd_SAP_reader_2("n_codici")
            Dim n_da As Integer = cmd_SAP_reader_2("N_da_trasferire")
            If n_cod > 0 Then
                pct = (n_cod - n_da) / n_cod * 100
            End If

            DataGridView_ODP.Rows.Add(
            False,
            cmd_SAP_reader_2("docnum"),
            cmd_SAP_reader_2("itemcode"),
            cmd_SAP_reader_2("u_disegno"),
            cmd_SAP_reader_2("itemname"),
            cmd_SAP_reader_2("plannedqty"),
            cmd_SAP_reader_2("CmpltQty"),
            cmd_SAP_reader_2("commessa"),
                cmd_SAP_reader_2("Sottocommessa"),
                 cmd_SAP_reader_2("U_PRG_AZS_Commessa"),
            cmd_SAP_reader_2("U_Final_customer_name"),
            cmd_SAP_reader_2("ResName"),
            cmd_SAP_reader_2("U_PRODUZIONE"),
            cmd_SAP_reader_2("status"),
            cmd_SAP_reader_2("Lotto"),
            cmd_SAP_reader_2("warehouse"),
            cmd_SAP_reader_2("startdate"),
            pct,
            cmd_SAP_reader_2("Mancanti"),
            cmd_SAP_reader_2("01"),
            cmd_SAP_reader_2("FERRETTO"),
            cmd_SAP_reader_2("03"),
            cmd_SAP_reader_2("SCA"),
            cmd_SAP_reader_2("MUT"),
            cmd_SAP_reader_2("CAP2"),
            cmd_SAP_reader_2("02")
        )
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        DataGridView_ODP.ClearSelection()
    End Sub



    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text = Nothing Then
            filtro_n_odp = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_n_odp = " and t0.docnum Like '%%" & TextBox1.Text & "%%' "
            Else
                filtro_n_odp = " and numodp Like ''%%" & TextBox1.Text & "%%'' "
            End If
        End If
        RIEMPI_Odp(TextBox10.Text)
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        If TextBox9.Text = Nothing Then
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_stato = " (t0.status='P' or t0.status='R')"
            Else
                filtro_stato = " and (pianificato='P' or pianificato='R')"
            End If
        Else
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_stato = " t0.status Like '%%" & TextBox9.Text & "%%'"
            Else
                filtro_stato = " and pianificato Like ''%%" & TextBox9.Text & "%%'' "
            End If
        End If
        RIEMPI_Odp(TextBox10.Text)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        If TextBox2.Text = Nothing Then
            filtro_codice = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_codice = " and t0.itemcode Like '%%" & TextBox2.Text & "%%'"
            Else
                filtro_codice = " and codart Like ''%%" & TextBox2.Text & "%%'' "
            End If
        End If
        RIEMPI_Odp(TextBox10.Text)
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        If TextBox3.Text = Nothing Then
            filtro_descrizione = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_descrizione = " and t1.itemname Like '%%" & TextBox3.Text & "%%'"
            Else
                filtro_descrizione = " and dscodart_odp Like ''%%" & TextBox3.Text & "%%'' "
            End If
        End If
        RIEMPI_Odp(TextBox10.Text)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = Nothing Then
            filtro_commessa = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then
                Dim filtro As String = TextBox4.Text
                If filtro.Contains("_") Then
                    filtro = filtro.Replace("_", "\_")
                    filtro_commessa = " and t0.U_PRG_AZS_Commessa Like '%" & filtro & "%' ESCAPE '\'"
                Else
                    filtro_commessa = " and matricola Like '%" & filtro & "%'"
                End If
            Else
                ' Dentro OPENQUERY non si può usare ESCAPE, l'underscore in AS400 è raro
                filtro_commessa = " and matricola Like ''%%" & TextBox4.Text & "%%'' "
            End If
        End If
        RIEMPI_Odp(TextBox10.Text)
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If TextBox5.Text = Nothing Then
            filtro_cliente = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_cliente = " and t2.U_Final_customer_name Like '%%" & TextBox5.Text & "%%'"
            Else
                filtro_cliente = " and cliente Like ''%%" & TextBox5.Text & "%%'' "
            End If
        End If
        RIEMPI_Odp(TextBox10.Text)
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If TextBox6.Text = Nothing Then
            filtro_Fase = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_Fase = " and coalesce(t3.ResName,t0.u_fase) Like '%%" & TextBox6.Text & "%%'"
            Else
                ' Fase non disponibile in AS400: filtro ignorato silenziosamente
                filtro_Fase = ""
            End If
        End If
        RIEMPI_Odp(TextBox10.Text)
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If TextBox7.Text = Nothing Then
            filtro_produzione = ""
        Else
            If Homepage.ERP_provenienza = "SAP" Then
                filtro_produzione = " and t0.u_produzione Like '%%" & TextBox7.Text & "%%'"
            Else
                ' U_PRODUZIONE non disponibile in AS400: filtro ignorato silenziosamente
                filtro_produzione = ""
            End If
        End If
        RIEMPI_Odp(TextBox10.Text)
    End Sub



    Private Sub DataGridView_ODP_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellClick
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = DataGridView_ODP.Columns.IndexOf(N_ODP) Then




                Dim new_form_odp_form = New ODP_Form
                new_form_odp_form.docnum_odp = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="N_ODP").Value
                new_form_odp_form.Show()
                new_form_odp_form.inizializza_form(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="N_ODP").Value)


            End If

            If e.ColumnIndex = DataGridView_ODP.Columns.IndexOf(Codice) Then
                Magazzino.Codice_SAP = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Codice").Value
                Magazzino.Show()

                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)
            End If

            If e.ColumnIndex = DataGridView_ODP.Columns.IndexOf(Disegno) Then
                Try
                    Process.Start(Homepage.percorso_disegni_generico & "PDF\" & DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Disegno").Value & ".PDF")
                Catch ex As Exception
                    MsgBox("Il disegno " & DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Disegno").Value & " non è ancora stato processato")
                End Try


            End If

        End If

    End Sub

    Private Sub DataGridView_ODP_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_ODP.CellFormatting

        If DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="trasferito").Value = 100 Then

            DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Trasferito").Style.BackColor = Color.Lime

        End If

        If DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="trasferito").Value = 100 Then
            DataGridView_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime

        End If

    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        If TextBox8.Text = Nothing Then
            filtro_ferretto = ""
        Else
            filtro_ferretto = " and b.ferretto > " & TextBox8.Text & ""
        End If
        RIEMPI_Odp(TextBox10.Text)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click





        ODP_Form.stampa_etichetta = "NO"


        For Each Row As DataGridViewRow In DataGridView1.Rows





            'If CheckBox1.Checked = True Then

            '    ODP_Form.stampa_etichetta = "YES"
            '    FORM6.ODP = Row.Cells("Numero_ODP").Value
            '    ODP_Form.docnum_odp = Row.Cells("Numero_ODP").Value
            '    ODP_Form.percorso_documento = Homepage.PERCORSO_DOCUMENTO_ODP_ETICHETTA
            '    ODP_Form.Genera_ordine()

            'ElseIf CheckBox2.Checked = True Then
            ODP_Form.testata_odp(Row.Cells("Numero_ODP").Value)
                ODP_Form.Fun_Stampa()
            'End If
            'If CheckBox3.Checked And File.Exists(Homepage.percorso_disegni_generico & "PDF\" & Row.Cells("disegno_odp").Value & ".PDF") = True Then


            '    'AxFoxitCtl1.OpenFile(Homepage.percorso_disegni_generico & "PDF\"  & Row.Cells("disegno_odp").Value & ".PDF")
            '    'AxFoxitCtl1.PrintFile()

            'End If




        Next
        MsgBox("FINE STAMPE")

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        cambia_stato()
    End Sub

    Sub cambia_stato()

        Dim nuovo_stato As String



        For Each row As DataGridViewRow In DataGridView1.Rows


            'If DataGridView_ODP.Rows(contatore).Cells(columnName:="Stato").Value = "R" Then
            '        nuovo_stato = "P"
            '    Else
            '        nuovo_stato = "R"
            '    End If

            nuovo_stato = "R"

            row.Cells("Stato_odp").Value = nuovo_stato
            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()
            Dim CMD_SAP As New SqlCommand
            CMD_SAP.Connection = Cnn

            CMD_SAP.CommandText = "UPDATE owor SET STATUS='" & nuovo_stato & "' WHERE DOCNUM ='" & row.Cells("Numero_ODP").Value & "'"
            CMD_SAP.ExecuteNonQuery()
            Cnn.Close()


            ' End If

        Next
    End Sub

    Private Sub button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Itera attraverso le righe della DataGridView "datagridview_odp"
        For Each row As DataGridViewRow In DataGridView_ODP.Rows
            ' Verifica se la cella della colonna "seleziona" è flaggata
            If CBool(row.Cells("seleziona").Value) = True Then
                ' Crea una nuova riga nella DataGridView "datagridview1"
                Dim index As Integer = DataGridView1.Rows.Add()

                ' Copia i valori dalle colonne necessarie
                DataGridView1.Rows(index).Cells("Numero_ODP").Value = row.Cells("N_ODP").Value
                DataGridView1.Rows(index).Cells("Commessa_odp").Value = row.Cells("Commessa").Value
                DataGridView1.Rows(index).Cells("disegno_odp").Value = row.Cells("Disegno").Value
                DataGridView1.Rows(index).Cells("Stato_odp").Value = row.Cells("stato").Value
                DataGridView1.Rows(index).Cells("Tipo").Value = "ODP"
                row.Cells("seleziona").Value = False
            End If
        Next

        For Each row As DataGridViewRow In DataGridView2.Rows
            ' Verifica se la cella della colonna "seleziona" è flaggata
            If CBool(row.Cells("DataGridViewCheckBoxColumn1").Value) = True Then
                ' Crea una nuova riga nella DataGridView "datagridview1"
                Dim index As Integer = DataGridView1.Rows.Add()

                ' Copia i valori dalle colonne necessarie
                DataGridView1.Rows(index).Cells("Numero_ODP").Value = row.Cells("Docnum").Value
                DataGridView1.Rows(index).Cells("Commessa_odp").Value = row.Cells("Cardname").Value
                DataGridView1.Rows(index).Cells("disegno_odp").Value = ""
                DataGridView1.Rows(index).Cells("Stato_odp").Value = "O"
                DataGridView1.Rows(index).Cells("Tipo").Value = "OC"
                row.Cells("DataGridViewCheckBoxColumn1").Value = False
            End If
        Next

    End Sub

    Private Sub button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Itera all'indietro attraverso le righe selezionate nella DataGridView "datagridview1"
        For i As Integer = DataGridView1.SelectedRows.Count - 1 To 0 Step -1
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(i)
            DataGridView1.Rows.Remove(row)
        Next
    End Sub

    Public Function Trova_nuovo_lotto_di_prelievo()
        Dim id_lotto As Integer
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select max(id)+1 as 'ID' from [Tirelli_40].[dbo].[Lotto_prelievo_testata]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                id_lotto = cmd_SAP_reader_2("ID")
            Else
                id_lotto = 1
            End If
        Else
            id_lotto = 1
        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        Return id_lotto
    End Function

    Sub check_esistenza_lotto_di_prelievo(id_lotto_prelievo As Integer, par_datagridview As DataGridView)
        Dim AZIONE As String
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader


        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "select *
from [Tirelli_40].[dbo].[Lotto_prelievo_testata]
where id='" & id_lotto_prelievo & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then
            AZIONE = "aggiorno"

        Else
            AZIONE = "creo"


        End If
        cmd_SAP_reader.Close()
        Cnn.Close()

        'If AZIONE = "aggiorno" Then
        '    aggiorna_lotto_di_prelievo(id_lotto_prelievo, par_datagridview)
        '    MsgBox("Lotto di prelievo aggiornato")
        'ElseIf AZIONE = "creo" Then
        '    
        '    crea_lotto_di_prelievo(id_lotto_prelievo, par_datagridview)
        '    MsgBox("Lotto di prelievo creato")
        'End If


        cancella_lotto_di_prelievo(id_lotto_prelievo)
        crea_lotto_di_prelievo(id_lotto_prelievo, par_datagridview, AZIONE)
        MsgBox("Lotto di prelievo creato")


    End Sub

    Sub aggiorna_lotto_di_prelievo(id_lotto_prelievo As Integer, par_datagridview As DataGridView)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "

delete [Tirelli_40].[dbo].[lotto_prelievo_riga] where id = " & id_lotto_prelievo & ""
        CMD_SAP.ExecuteNonQuery()


        For Each row As DataGridViewRow In par_datagridview.Rows


            CMD_SAP.CommandText = "insert into [Tirelli_40].[dbo].[lotto_prelievo_riga] (id,tipo_doc,docnum)
values (" & id_lotto_prelievo & ",'" & row.Cells("Tipo").Value & "','" & row.Cells("Numero_ODP").Value & "')"
            CMD_SAP.ExecuteNonQuery()



        Next
        Cnn.Close()
    End Sub

    Sub cancella_lotto_di_prelievo(id_lotto_prelievo As Integer)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "

delete [Tirelli_40].[dbo].[lotto_prelievo_riga] where id = " & id_lotto_prelievo & ""
        CMD_SAP.ExecuteNonQuery()



        Cnn.Close()
    End Sub

    Sub crea_lotto_di_prelievo(id_lotto_prelievo As Integer, par_datagridview As DataGridView, par_Azione As String)
        If par_Azione = "creo" Then
            crea_testata_lotto_prelievo(id_lotto_prelievo, "")
        End If
        If par_Azione = "TRASFERIBILI_ODP" Then
            crea_righe_lotto_prelievo_TRASFERIBILI_ODP(id_lotto_prelievo)
        ElseIf par_Azione = "TRASFERIBILI_OC" Then
            crea_righe_lotto_prelievo_TRASFERIBILI_OC(id_lotto_prelievo)
        Else

            For Each row As DataGridViewRow In par_datagridview.Rows
                crea_righe_lotto_prelievo(id_lotto_prelievo, row.Cells("Numero_ODP").Value, row.Cells("Tipo").Value)
            Next
        End If

    End Sub


    Sub crea_testata_lotto_prelievo(id_lotto_prelievo As Integer, par_commento As String)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "INSERT INTO [Tirelli_40].[dbo].[Lotto_prelievo_testata] (ID,DATA,ORA,DIP,COMMENTO) VALUES
(" & id_lotto_prelievo & ",GETDATE(),convert(varchar, getdate(), 108)," & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato & ",'" & par_commento & "')"
        Cmd_SAP.ExecuteNonQuery()



        cnn.Close()
    End Sub

    Sub crea_righe_lotto_prelievo(id_lotto_prelievo As Integer, par_numero_odp As Integer, par_tipo_documento As String)

        Dim Cnn6 As New SqlConnection
        Cnn6.ConnectionString = Homepage.sap_tirelli
        Cnn6.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn6

        CMD_SAP.CommandText = "insert into [Tirelli_40].[dbo].[lotto_prelievo_riga] (id,tipo_doc,docnum)
values (" & id_lotto_prelievo & ",'" & par_tipo_documento & "','" & par_numero_odp & "')"
        CMD_SAP.ExecuteNonQuery()
        Cnn6.Close()



    End Sub

    Sub crea_righe_lotto_prelievo_TRASFERIBILI_ODP(id_lotto_prelievo As Integer)

        Dim Cnn6 As New SqlConnection
        Cnn6.ConnectionString = Homepage.sap_tirelli
        Cnn6.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn6

        CMD_SAP.CommandText = "insert into [Tirelli_40].[dbo].[lotto_prelievo_riga] (id,tipo_doc,docnum)
SELECT 
    " & id_lotto_prelievo & ",
		'ODP',
    t0.docnum
	

from owor t0 left join oitm t1 on t0.itemcode=t1.itemcode
left join oitm t2 on t2.itemcode=t0.U_PRG_AZS_Commessa
left join orsc t3 on t3.VisResCode=t0.U_Fase


left join 
    (
    select t10.docnum, count(t10.itemcode) as 'N_codici', sum(case when t10.U_PRG_WIP_QtaDaTrasf >0 then 1 else 0 end ) as 'N_da_trasferire' , sum(case when t10.onhand>=U_PRG_WIP_QtaDaTrasf  then 1 else 0 end ) as 'N_trasferibili'
    from
            (
            select t0.docnum, t1.itemcode, case when t1.U_PRG_WIP_QtaDaTrasf is null then 0 else t1.U_PRG_WIP_QtaDaTrasf end as 'U_PRG_WIP_QtaDaTrasf', case when t2.onhand is null then 0 else t2.onhand end as 'onhand'
            from owor t0 inner join wor1 t1 on t0.DocEntry=t1.docentry
            left join oitw t2 on t2.WhsCode=t1.wareHouse and t2.itemcode=t1.itemcode
            where (t0.status='P' or t0.status='R') and t1.ItemType=4
            )
            as t10
            group by t10.docnum
    ) A on a.docnum=t0.docnum


left join

        (select t20.docnum, sum( case when t20.warehouse='01' then N_Trasferibili else 0 end) as '01'
		, sum( case when t20.warehouse='FERRETTO' then N_Trasferibili else 0 end) as 'FERRETTO'
		, sum( case when t20.warehouse='BSCA' then N_Trasferibili else 0 end) as 'BSCA'
		, sum( case when t20.warehouse='03' then N_Trasferibili else 0 end) as '03', sum( case when t20.warehouse='SCA' then N_Trasferibili else 0 end) as 'SCA', sum( case when t20.warehouse='CAP2' then N_Trasferibili else 0 end) as 'CAP2', sum( case when t20.warehouse='MUT' then N_Trasferibili else 0 end) as 'MUT' , sum( case when t20.warehouse='02' then N_Trasferibili else 0 end) as '02'
        from
            (
            select t10.docnum, t10.wareHouse, sum(case when t10.ONHAND>= t10.U_PRG_WIP_QtaDaTrasf AND t10.U_PRG_WIP_QtaDaTrasf>0 then 1 else 0 end) as 'N_Trasferibili'
            from
                (
                select t0.docnum,  case when t1.U_PRG_WIP_QtaDaTrasf is null then 0 else t1.U_PRG_WIP_QtaDaTrasf end as 'U_PRG_WIP_QtaDaTrasf', t1.wareHouse, CASE WHEN T2.ONHAND IS NULL THEN 0 ELSE T2.ONHAND END AS 'ONHAND'
                from owor t0 inner join wor1 t1 on t0.DocEntry=t1.docentry
                left join oitw t2 on t2.WhsCode=t1.wareHouse and t2.itemcode=t1.itemcode
                where (t0.status='P' or t0.status='R') and (substring(t1.itemcode,1,1)='0' or substring(t1.itemcode,1,1)='c' or substring(t1.itemcode,1,1)='d' or substring(t1.itemcode,1,1)='f')

                )
                as t10
                group by t10.docnum, t10.wareHouse
            )
            as t20
            GROUP BY t20.docnum) B on t0.docnum=B.docnum

left join 
(select  max(t0.id) as 'ID', t0.docnum 
from [Tirelli_40].[dbo].[lotto_prelievo_riga] t0 inner join owor t1 on t0.docnum=t1.docnum
where t1.status='P' or t1.status='R'


group by t0.docnum) C on c.docnum=t0.docnum

left join 
(
SELECT  
      T0.[COMMESSA]
    
	 
      ,MIN(case when (T0.RISORSA='P01501' or T0.RISORSA='P02001') then  T0.[DATA_I]  end) AS 'inizio_montaggio'
	  ,MIN(case when (T0.RISORSA='P04001' or T0.RISORSA='P03001') then  T0.[DATA_I]  end) AS 'inizio_collaudo'
      
  FROM [Tirelli_40].[dbo].[PIANIFICAZIONE] T0 LEFT JOIN ORSC T1 ON T1.VisResCode=T0.RISORSA
  WHERE (T0.RISORSA='P01501' OR T0.RISORSA='P02001' OR T0.RISORSA='P04001' OR T0.RISORSA='P03001')
  group BY T0.[COMMESSA]
  ) D on D.COMMESSA=T0.U_PRG_AZS_COMMESSA

where 



(t0.status='P' or t0.status='R')  and (t0.u_produzione Like '%%ASS%%' or t0.u_produzione Like '%%EST%%')  and a.N_da_trasferire>0 and (B.[03]+B.FERRETTO+B.SCA+B.BSCA)>0  AND T1.ItemName   NOT Like '%%ANTICIP%%' 

AND ((t0.StartDate<GETDATE()+14 AND SUBSTRING(T0.U_PRG_AZS_COMMESSA,1,1)<>'M') OR (coalesce(d.inizio_montaggio,getdate()+1) <=GETDATE() AND (t0.u_FASE='P01501' or t0.u_FASE='P02001') OR coalesce(d.inizio_collaudo,getdate()+6) <=GETDATE()+5 and (t0.u_FASE<>'P01501' and t0.u_FASE<>'P02001') ))"

        CMD_SAP.ExecuteNonQuery()
        Cnn6.Close()



    End Sub
    Sub crea_righe_lotto_prelievo_TRASFERIBILI_OC(id_lotto_prelievo As Integer)

        Dim Cnn6 As New SqlConnection
        Cnn6.ConnectionString = Homepage.sap_tirelli
        Cnn6.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn6

        CMD_SAP.CommandText = "insert into [Tirelli_40].[dbo].[lotto_prelievo_riga] (id,tipo_doc,docnum)
select " & id_lotto_prelievo & ",'OC',T10.DOCNUM 
from
(
select t1.docentry,t1.docnum,t1.cardcode, t1.cardname,coalesce(t2.cardcode,'') as 'Codice_CF', coalesce(t2.CardName,'') as 'Cliente_f'
,t1.u_causcons
, t1.CreateDate,t1.DocDueDate
, sum(case when t0.U_Datrasferire>0 then 1 else 0 end) as 'Codici_da_trasf'
,sum(case when t0.U_Datrasferire>0 and t3.onhand>0 then 1 else 0 end) as 'Codici_trasferibili'
,sum(case when t0.U_Datrasferire>0 and t7.onhand>0 then 1 else 0 end) as 'SCA'
,sum(case when t0.U_Datrasferire>0 and t4.onhand>0 then 1 else 0 end) as 'Ferretto'
,sum(case when t0.U_Datrasferire>0 and t6.onhand>0 then 1 else 0 end) as '03'
,sum(case when t0.U_Datrasferire>0 and t5.onhand>0 then 1 else 0 end) as 'BSCA'
,sum(case when t0.U_Datrasferire>0 and t8.onhand>0 then 1 else 0 end) as '01'


from rdr1 t0 
inner join ordr t1 on t0.docentry=t1.docentry
left join ocrd t2 on t2.cardcode=t1.U_codicebp
left join oitw t3 on t3.whscode=t0.WhsCode and t3.itemcode=t0.itemcode
left join oitw t4 on t4.whscode=t0.WhsCode and t4.itemcode=t0.itemcode and t4.WhsCode='FERRETTO'
left join oitw t5 on t5.whscode=t0.WhsCode and t5.itemcode=t0.itemcode and t5.WhsCode='BSCA'
left join oitw t6 on t6.whscode=t0.WhsCode and t6.itemcode=t0.itemcode and t6.WhsCode='03'
left join oitw t7 on t7.whscode=t0.WhsCode and t7.itemcode=t0.itemcode and t7.WhsCode='SCA'
left join oitw t8 on t8.whscode=t0.WhsCode and t8.itemcode=t0.itemcode and t7.WhsCode='01'

where t1.docstatus='O'and t0.OpenCreQty>0 and (substring(t0.itemcode,1,1)='0' or substring(t0.itemcode,1,1)='C' or substring(t0.itemcode,1,1)='d' or substring(t0.itemcode,1,1)='F')
  
group by t1.docentry,t1.docnum,t1.cardcode, t1.cardname,t2.cardcode , t2.CardName,t1.CreateDate,t1.DocDueDate,t1.u_causcons
)
as t10
where t10.Codici_trasferibili>0 and t10.docduedate<=getdate()+10
GROUP BY T10.DOCNUM 
"
        CMD_SAP.ExecuteNonQuery()
        Cnn6.Close()



    End Sub


    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        check_esistenza_lotto_di_prelievo(CInt(Txt_DocNum.Text), DataGridView1)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)
        Form_lotto_di_prelievo.Show()
        Form_lotto_di_prelievo.id_lotto = Txt_DocNum.Text
        Form_lotto_di_prelievo.inizializzazione_lotto_di_prelievo()
    End Sub

    Private Sub Cmd_Indietro_Click(sender As Object, e As EventArgs) Handles Cmd_Indietro.Click
        Txt_DocNum.Text = Int(Txt_DocNum.Text) - 1
        RIEMPI_datagridview_documenti_lotto(Txt_DocNum.Text, DataGridView1, DataGridView1)
    End Sub

    Private Sub Cmd_Avanti_Click(sender As Object, e As EventArgs) Handles Cmd_Avanti.Click
        Txt_DocNum.Text = Int(Txt_DocNum.Text) + 1
        RIEMPI_datagridview_documenti_lotto(Txt_DocNum.Text, DataGridView1, DataGridView1)
    End Sub

    Sub RIEMPI_datagridview_documenti_lotto(id_lotto_prelievo As Integer, par_datagridview_aggiunta As DataGridView, par_datagridview_eliminazione As DataGridView)

        par_datagridview_eliminazione.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "declare @id_lotto_prelievo as integer
set @id_lotto_prelievo=' " & id_lotto_prelievo & "'


select t0.tipo_doc, t1.docnum, t1.itemcode, t1.ProdName, t2.U_Disegno, t1.status, case when t1.u_prg_azs_commessa is null then '' else t1.u_prg_azs_commessa end as 'u_prg_azs_commessa'
, sum(case when t3.u_prg_wip_qtadatrasf > 0 and t3.wareHousE ='03' then 1 else 0 end) as 'Mag_03'
, sum(case when t3.u_prg_wip_qtadatrasf > 0 and t3.wareHousE ='SCA' then 1 else 0 end) as 'SCA'
, sum(case when t3.u_prg_wip_qtadatrasf > 0 and t3.wareHousE ='FERRETTO' then 1 else 0 end) as 'FERRETTO'
, sum(case when t3.u_prg_wip_qtadatrasf > 0 then 1 else 0 end) as 'Da_trasf'
, sum(case when t3.u_prg_wip_qtadatrasf = 0 then 1 else 0 end) as 'TRASFERITO'

from 
[Tirelli_40].[dbo].[lotto_prelievo_riga] t0 
left join owor t1 on t0.docnum=t1.docnum and t0.tipo_DOC='ODP'
left join oitm t2 on t2.itemcode=t1.itemcode
left join wor1 t3 on t3.docentry=t1.docentry AND (SUBSTRING(T3.ITEMCODE,1,1)='0' OR SUBSTRING(T3.ITEMCODE,1,1)='C' OR SUBSTRING(T3.ITEMCODE,1,1)='D' OR SUBSTRING(T3.ITEMCODE,1,1)='F')
where t0.id=@id_lotto_prelievo
group by t0.tipo_doc, t1.docnum, t1.itemcode, t1.ProdName, t2.U_Disegno, t1.status, t1.u_prg_azs_commessa"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            par_datagridview_aggiunta.Rows.Add(cmd_SAP_reader_2("docnum"), cmd_SAP_reader_2("u_disegno"), cmd_SAP_reader_2("u_prg_azs_commessa"), cmd_SAP_reader_2("status"), cmd_SAP_reader_2("Tipo_doc"))
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        DataGridView1.ClearSelection()
    End Sub

    Sub RIEMPI_datagridview_OC(par_datagridview As DataGridView)

        par_datagridview.Rows.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "
select *
from
(
select t1.docentry,t1.docnum,t1.cardcode, t1.cardname,coalesce(t2.cardcode,'') as 'Codice_CF', coalesce(t2.CardName,'') as 'Cliente_f'
,t1.u_causcons
, t1.CreateDate,t1.DocDueDate
, sum(case when t0.U_Datrasferire>0 then 1 else 0 end) as 'Codici_da_trasf'
,sum(case when t0.U_Datrasferire>0 and t3.onhand>0 then 1 else 0 end) as 'Codici_trasferibili'
,sum(case when t0.U_Datrasferire>0 and t4.onhand>0 then 1 else 0 end) as 'Ferretto'

from rdr1 t0 
inner join ordr t1 on t0.docentry=t1.docentry
left join ocrd t2 on t2.cardcode=t1.U_codicebp
left join oitw t3 on t3.whscode=t0.WhsCode and t3.itemcode=t0.itemcode
left join oitw t4 on t4.whscode=t0.WhsCode and t4.itemcode=t0.itemcode and t4.WhsCode='FERRETTO'
where t1.docstatus='O'and t0.OpenCreQty>0 and (substring(t0.itemcode,1,1)='0' or substring(t0.itemcode,1,1)='C' or substring(t0.itemcode,1,1)='d' or substring(t0.itemcode,1,1)='F')
" & filtro_n_oc & " " & filtro_cliente_OC & " " & filtro_causcons_oc & "
group by t1.docentry,t1.docnum,t1.cardcode, t1.cardname,t2.cardcode , t2.CardName,t1.CreateDate,t1.DocDueDate,t1.u_causcons
)
as t10
where 0=0 " & filtro_ferretto_oc & "
order by t10.docduedate"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(False, cmd_SAP_reader_2("docentry"), cmd_SAP_reader_2("docnum"), cmd_SAP_reader_2("cardcode"), cmd_SAP_reader_2("cardname"), cmd_SAP_reader_2("codice_CF"), cmd_SAP_reader_2("cliente_f"), cmd_SAP_reader_2("u_causcons"), cmd_SAP_reader_2("createdate"), cmd_SAP_reader_2("docduedate"), cmd_SAP_reader_2("Codici_da_trasf"), cmd_SAP_reader_2("Codici_trasferibili"), cmd_SAP_reader_2("Ferretto"))
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        par_datagridview.ClearSelection()
    End Sub

    Private Sub tabpage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter

        RIEMPI_datagridview_OC(DataGridView2)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        RIEMPI_datagridview_documenti_lotto(Txt_DocNum.Text, DataGridView1, DataGridView1)
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        If TextBox13.Text = Nothing Then
            filtro_n_oc = ""
        Else
            filtro_n_oc = " and t1.docnum    Like '%%" & TextBox13.Text & "%%'  "
        End If

        RIEMPI_datagridview_OC(DataGridView2)
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        If TextBox14.Text = Nothing Then
            filtro_cliente_OC = ""
        Else
            filtro_cliente_OC = " and (t1.cardname Like '%%" & TextBox14.Text & "%%' or t2.cardname Like '%%" & TextBox14.Text & "%%' )   "
        End If

        RIEMPI_datagridview_OC(DataGridView2)
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        If TextBox15.Text = Nothing Then
            filtro_causcons_oc = ""
        Else
            filtro_causcons_oc = " and t1.u_causcons    Like '%%" & TextBox15.Text & "%%'  "
        End If

        RIEMPI_datagridview_OC(DataGridView2)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        inizializzazione_ordine_di_produzione_lista()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            filtro_completi = " and a.N_da_trasferire>0 and (B.[01]+B.[03]+B.FERRETTO+B.SCA +B.BSCA)>0"
        Else
            filtro_completi = " "
        End If

        RIEMPI_Odp(TextBox10.Text)
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

    End Sub

    Private Sub DataGridView_ODP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellContentClick

    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        If TextBox11.Text = Nothing Then
            filtro_magazzino_destinazione = ""
        Else
            filtro_magazzino_destinazione = " and t0.warehouse ='" & TextBox11.Text & "'"
        End If
        RIEMPI_Odp(TextBox10.Text)
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

        If TextBox12.Text = Nothing Then
            filtro_ferretto_oc = ""
        Else
            filtro_ferretto_oc = " and t10.ferretto > " & TextBox12.Text & ""
        End If
        RIEMPI_datagridview_OC(DataGridView2)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim par_datagridview As DataGridView = DataGridView_ODP
        For Each row As DataGridViewRow In par_datagridview.Rows
            If row.Cells("Seleziona").Value = True Then
                Trasferimento_magazzino.docentry_odp = ODP_Form.ottieni_informazioni_odp("Numero", 0, row.Cells("N_ODP").Value).docentry
                Trasferimento_magazzino.docnum_odp = row.Cells("N_ODP").Value

                Trasferimento_magazzino.Text = "Trasferimento"
                Trasferimento_magazzino.inizializzazione_trasferimento(Trasferimento_magazzino.docentry_odp, 0, "Trasferimento", "ODP")

            End If



        Next
        Trasferimento_magazzino.Show()

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        cancella_lotto_di_prelievo(1)
        crea_lotto_di_prelievo(1, DataGridView_ODP, "TRASFERIBILI_ODP")
        cancella_lotto_di_prelievo(2)
        crea_lotto_di_prelievo(2, DataGridView_ODP, "TRASFERIBILI_OC")
        MsgBox("Lotto di prelievo TRASFERIBILI aggiornato")
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub GroupBox19_Enter(sender As Object, e As EventArgs) Handles GroupBox19.Enter

    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        trova_dato_da_excel_pEr_importazionE("\\tirfs01\07-Warehouse\Magazzino\Andrea\Import lotto di prelievo.xlsx", "Sheet1", 2)
    End Sub

    Sub trova_dato_da_excel_pEr_importazionE(par_percorso_file As String, par_nome_foglio As String, par_riga_inizio As Integer)

        Dim colonna1 As String
        Dim colonna2 As String
        Dim colonna3 As String
        Dim colonna4 As String
        Dim colonna5 As String
        Dim colonna6 As String

        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True

        ' ciclo fino a quando la colonna A (1) non è vuota
        Do While Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).Value <> "" AndAlso
             Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).Value IsNot Nothing

            colonna1 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).Value
            colonna2 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).Value
            colonna3 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 3).Value
            colonna4 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).Value
            colonna5 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 5).Value
            colonna6 = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 6).Value

            ' crea_codice_articolo(colonna1, colonna2, colonna3, colonna4, colonna5, colonna6)
            insert_into_lotto_prelievo_riga(colonna1, colonna2, colonna3, colonna4, colonna5, colonna6)
            'UPDATE(colonna1, colonna2, colonna3, colonna4, colonna5)

            par_riga_inizio += 1
        Loop

        Beep()
        MsgBox("Importazione effettuata con successo")

    End Sub

    Sub insert_into_lotto_prelievo_riga(par_Colonna_1 As String, par_Colonna_2 As String, par_Colonna_3 As String, par_Colonna_4 As String, par_colonna_5 As String, par_colonna_6 As String)

        par_Colonna_1 = Replace(par_Colonna_1, "'", " ")
        par_Colonna_2 = Replace(par_Colonna_2, "'", " ")
        par_Colonna_3 = Replace(par_Colonna_3, "'", " ")
        par_Colonna_4 = Replace(par_Colonna_4, "'", " ")
        par_colonna_5 = Replace(par_colonna_5, "'", " ")
        par_colonna_6 = Replace(par_colonna_6, ",", ".")


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()
        Dim Cmd_SAP As New SqlCommand
        Cmd_SAP.Connection = Cnn1
        Cmd_SAP.CommandText = "INSERT INTO [tirelli_40].[dbo].[lotto_prelievo_riga]
           ([Tipo_doc]
,[docnum]
,[id]
         
           
           )
     VALUES
           ('" & par_Colonna_1 & "'
           ," & par_Colonna_2 & "
           ," & par_Colonna_3 & "
           )
          "

        Cmd_SAP.ExecuteNonQuery()
        Cnn1.Close()

    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged

        If TextBox16.Text = Nothing Then
            filtro_commessa_ = ""
        Else

            If Homepage.ERP_provenienza = "SAP" Then

                filtro_commessa_ = ""
            Else
                filtro_commessa = " and cod_commessa Like ''%" & TextBox16.Text & "%''"
            End If


        End If

        RIEMPI_Odp(TextBox10.Text)
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged

        If TextBox17.Text = Nothing Then
            filtro_sottocommessa = ""
        Else

            If Homepage.ERP_provenienza = "SAP" Then

                filtro_sottocommessa = ""
            Else
                filtro_sottocommessa = " and cod_sottocommessa Like ''%" & TextBox17.Text & "%''"
            End If


        End If

        RIEMPI_Odp(TextBox10.Text)
    End Sub
End Class