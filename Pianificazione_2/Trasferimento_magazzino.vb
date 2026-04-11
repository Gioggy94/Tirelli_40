Imports System.Data.SqlClient

Public Class Trasferimento_magazzino
    Public docnum_odp As String
    Public docentry_odp As Integer = 0
    Public documento As String = 0
    Public docentry_oc As Integer = 0
    Public docnum_oc As Integer = 0
    Private filtro_magazzino As String
    Public isShiftKeyDown As Boolean = False
    Private startIndex As Integer

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Trasferimento_magazzino_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Sub inizializzazione_trasferimento(par_docentry_odp As Integer, par_docentry_oc As Integer, par_stringa_trasferimento As String, par_documento As String)
        Me.Text = par_stringa_trasferimento
        If par_stringa_trasferimento = "Trasferimento" Then
            magazzini_disponibili(ComboBox1, par_docentry_odp, par_docentry_oc, par_stringa_trasferimento, par_documento)

            trova_trasferibili(DataGridView1, par_docentry_odp, par_docentry_oc, par_stringa_trasferimento, par_documento)
        ElseIf par_stringa_trasferimento = "Reso" Then




            trova_rendibili(DataGridView1, par_docentry_odp, par_docentry_oc, par_stringa_trasferimento, par_documento)
        End If

    End Sub

    Sub magazzini_disponibili(par_combobox As ComboBox, par_docentry_odp As Integer, par_docentry_oc As Integer, par_stringa_Trasferimento As String, PAR_DOCUMENTO As String)

        documento = PAR_DOCUMENTO
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        If PAR_DOCUMENTO = "ODP" Then



            CMD_SAP_2.CommandText = "select T1.[wareHouse] as 'Whscode'

from owor t0 
inner join wor1 t1 on t0.docentry=t1.docentry
inner join oitm t2 on t2.itemcode=t1.itemcode
inner join oitw t3 on t3.itemcode=t1.itemcode and t3.whscode=t1.wareHouse


where t1.U_PRG_WIP_QtaDaTrasf>0 and t1.itemtype=4 and t1.docentry=" & par_docentry_odp & " and t3.onhand>0
group by T1.[wareHouse]"
        ElseIf PAR_DOCUMENTO = "OC" Then
            CMD_SAP_2.CommandText = "select T1.[whscode]

from ordr t0 
inner join rdr1 t1 on t0.docentry=t1.docentry
inner join oitm t2 on t2.itemcode=t1.itemcode
inner join oitw t3 on t3.itemcode=t1.itemcode and t3.whscode=t1.whscode


where t1.U_datrasferire>0 and t1.itemtype=4 and t1.docentry=" & par_docentry_oc & " and t3.onhand>0
group by T1.[whscode]"

        End If

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()

            par_combobox.Items.Add(cmd_SAP_reader_2("whscode"))

        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()


    End Sub

    Sub trova_trasferibili(par_datagridview As DataGridView, par_docentry_odp As Integer, par_docentry_oc As Integer, par_stringa_Trasferimento As String, PAR_DOCUMENTO As String)
        Dim magazzino_arrivo As String

        ComboBox4.Items.Clear()


        '  par_datagridview.Rows.Clear()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        If PAR_DOCUMENTO = "ODP" Then



            CMD_SAP_2.CommandText = "select t0.docnum, t1.itemcode,t2.itemname
,coalesce(t1.u_prg_wip_qtaspedita,0) as 'Trasf'
,t1.U_PRG_WIP_QtaDaTrasf as 'Da_trasf', t1.PlannedQty,t3.onhand,
case when t3.onhand<t1.U_PRG_WIP_QtaDaTrasf then t3.onhand else t1.U_PRG_WIP_QtaDaTrasf end as 'Q'
,t1.IssuedQty,T1.[wareHouse] as 'whscode',t1.U_PRG_WIP_QtaRichMagAuto,t1.U_Qta_richiesta_wip,t1.docentry as 'Docentry_odp',0 as 'docentry_oc',t1.linenum as 'Linenum_ODP', 0 as 'Linenum_OC', t0.u_prg_azs_commessa

from owor t0 
inner join wor1 t1 on t0.docentry=t1.docentry
inner join oitm t2 on t2.itemcode=t1.itemcode
inner join oitw t3 on t3.itemcode=t1.itemcode and t3.whscode=t1.wareHouse


where t1.U_PRG_WIP_QtaDaTrasf>0 and t1.itemtype=4 and t1.docentry=" & par_docentry_odp & " and t3.onhand>0 " & filtro_magazzino & ""

        ElseIf PAR_DOCUMENTO = "OC" Then

            CMD_SAP_2.CommandText = "select t0.docnum, t1.itemcode,t2.itemname
,coalesce(t1.u_trasferito,0) as 'Trasf',
t1.U_Datrasferire as 'Da_trasf', t1.OpenQty,t3.onhand,
case when t3.onhand<t1.U_Datrasferire then t3.onhand else t1.U_Datrasferire end as 'Q',
T1.[whscode],t1.U_PRG_WIP_QtaRichMagAuto,t1.docentry as 'docentry_oc' ,0 as 'docentry_odp', 0 as 'Linenum_Odp',t1.linenum as 'Linenum_OC', t0.u_prg_azs_commessa, 0 as 'U_Qta_richiesta_wip'

from ordr t0 
inner join rdr1 t1 on t0.docentry=t1.docentry
inner join oitm t2 on t2.itemcode=t1.itemcode
inner join oitw t3 on t3.itemcode=t1.itemcode and t3.whscode=t1.WhsCode


where t1.U_Datrasferire>0 and t1.itemtype=4 and t1.docentry=" & par_docentry_oc & " and t3.onhand>0 " & filtro_magazzino & ""
        End If

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Dim contatore As Integer = 0
        Do While cmd_SAP_reader_2.Read()



            magazzino_arrivo = "WIP"


            par_datagridview.Rows.Add(False, PAR_DOCUMENTO, cmd_SAP_reader_2("itemcode"), cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("Trasf"), cmd_SAP_reader_2("Da_trasf"), cmd_SAP_reader_2("Onhand"), Replace(cmd_SAP_reader_2("Q"), ",", "."), cmd_SAP_reader_2("whscode"), magazzino_arrivo, cmd_SAP_reader_2("U_PRG_WIP_QtaRichMagAuto"), cmd_SAP_reader_2("U_Qta_richiesta_wip"), cmd_SAP_reader_2("docentry_odp"), cmd_SAP_reader_2("docentry_oc"), cmd_SAP_reader_2("linenum_ODP"), cmd_SAP_reader_2("linenum_OC"), cmd_SAP_reader_2("u_prg_azs_commessa"))
            contatore += 1
        Loop



        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        par_datagridview.ClearSelection()
        If contatore > 0 Then
            Me.Show()
            ComboBox4.Items.Add(magazzino_arrivo)
            ComboBox4.Text = magazzino_arrivo
        Else
            MsgBox("Non sono presenti elementi trasferibili")
        End If

        par_datagridview.ClearSelection()
    End Sub

    Sub trova_rendibili(par_datagridview As DataGridView, par_docentry_odp As Integer, par_docentry_oc As Integer, par_stringa_Trasferimento As String, PAR_DOCUMENTO As String)
        Dim magazzino_arrivo As String = ""
        ComboBox4.Items.Clear()

        par_datagridview.Rows.Clear()
        documento = PAR_DOCUMENTO
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1

        If PAR_DOCUMENTO = "ODP" Then



            CMD_SAP_2.CommandText = "select t0.docnum, t1.itemcode,t2.itemname,
coalesce(t1.u_prg_wip_Qtaspedita,0) as 'Trasf'
,t1.U_PRG_WIP_QtaDaTrasf as 'Da_trasf', t1.PlannedQty,t3.onhand,
case when t3.onhand<coalesce(t1.u_prg_wip_Qtaspedita,0) then t3.onhand else coalesce(t1.u_prg_wip_Qtaspedita,0) end as 'Q'
,t1.IssuedQty,T1.[wareHouse] as 'whscode',t1.U_PRG_WIP_QtaRichMagAuto,t1.U_Qta_richiesta_wip,t1.docentry as 'Docentry_odp',0 as 'docentry_oc',t1.linenum as 'Linenum_ODP', 0 as 'Linenum_OC', t0.u_prg_azs_commessa

from owor t0 
inner join wor1 t1 on t0.docentry=t1.docentry
inner join oitm t2 on t2.itemcode=t1.itemcode
inner join oitw t3 on t3.itemcode=t1.itemcode and t3.whscode=t1.wareHouse


where t1.U_PRG_WIP_Qtaspedita>0 and t1.itemtype=4 and t1.docentry=" & par_docentry_odp & " and t3.onhand>0 and (T1.[wareHouse]='WIP' OR T1.[wareHouse]='BWIP')"

        ElseIf PAR_DOCUMENTO = "OC" Then

            CMD_SAP_2.CommandText = "select t0.docnum, t1.itemcode,t2.itemname,coalesce(t1.u_trasferito,0) as 'Trasf',t1.U_Datrasferire as 'Da_trasf', t1.OpenQty,t3.onhand,
case when t3.onhand<t1.U_TRASFERITO then t3.onhand else t1.U_TRASFERITO end as 'Q',
T1.[whscode],t1.U_PRG_WIP_QtaRichMagAuto,t1.docentry as 'docentry_oc' ,0 as 'docentry_odp',0 as 'Linenum_Odp',t1.linenum as 'Linenum_OC', t0.u_prg_azs_commessa, 0 as 'U_Qta_richiesta_wip'

from ordr t0 
inner join rdr1 t1 on t0.docentry=t1.docentry
inner join oitm t2 on t2.itemcode=t1.itemcode
inner join oitw t3 on t3.itemcode=t1.itemcode and t3.whscode=t1.WhsCode


where t1.U_trasferito>0 and t1.itemtype=4 and t1.docentry=" & par_docentry_oc & " and t3.onhand>0 and (T1.[whscode]='WIP' OR T1.[whscode]='BWIP') "
        End If

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Dim contatore As Integer = 0
        Do While cmd_SAP_reader_2.Read()
            If Form_Entrate_Merci.Trova_business_unit_magazzino(cmd_SAP_reader_2("whscode")) = 13 Then
                magazzino_arrivo = "B01"
            Else
                magazzino_arrivo = "01"
            End If
            par_datagridview.Rows.Add(False, PAR_DOCUMENTO, cmd_SAP_reader_2("itemcode"), cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("Trasf"), cmd_SAP_reader_2("Da_trasf"), cmd_SAP_reader_2("Onhand"), Replace(cmd_SAP_reader_2("Q"), ",", "."), cmd_SAP_reader_2("whscode"), magazzino_arrivo, cmd_SAP_reader_2("U_PRG_WIP_QtaRichMagAuto"), cmd_SAP_reader_2("U_Qta_richiesta_wip"), cmd_SAP_reader_2("docentry_odp"), cmd_SAP_reader_2("docentry_oc"), cmd_SAP_reader_2("linenum_ODP"), cmd_SAP_reader_2("linenum_OC"), cmd_SAP_reader_2("u_prg_azs_commessa"))
            If contatore = 0 Then
                ComboBox1.Items.Add(cmd_SAP_reader_2("whscode"))
            End If
            contatore += 1
        Loop
        If magazzino_arrivo = "" Then
            MsgBox("Non sono presenti elementi rendibili")
            End
        Else
            Me.Show()
        End If
        ComboBox4.Items.Add(magazzino_arrivo)
        ComboBox4.Text = magazzino_arrivo

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        par_datagridview.ClearSelection()
        If contatore > 0 Then
            Me.Show()
        Else
            MsgBox("Non sono presenti elementi rendibili")
        End If

        par_datagridview.ClearSelection()
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Dim linenum As Integer
        Dim contatore_odp As Integer = 0
        Dim contatore_oc As Integer = 0
        Dim magazzino_destinazione As String


        magazzino_destinazione = "WIP"




        If ComboBox1.SelectedIndex < 0 Then
            MsgBox("SELEZIONARE UN MAGAZZINO DI PARTENZA")
        Else
            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.Visible AndAlso Convert.ToBoolean(row.Cells("Sel").Value) = True Then


                    If row.Cells("DOC").Value = "ODP" Then
                        linenum = row.Cells("linenum_odp").Value
                        contatore_odp += 1
                    ElseIf row.Cells("DOC").Value = "OC" Then
                        linenum = row.Cells("linenum_oc").Value
                        contatore_oc += 1
                    Else linenum = 0
                    End If

                    If (Form_Entrate_Merci.Trova_business_unit_magazzino(magazzino_destinazione) = 13 And Form_Entrate_Merci.Trova_business_unit_magazzino(row.Cells("Dal_mag").Value) <> 13) Or (Form_Entrate_Merci.Trova_business_unit_magazzino(row.Cells("al_mag").Value) = 13 And Form_Entrate_Merci.Trova_business_unit_magazzino(magazzino_destinazione) <> 13 Or (Form_Entrate_Merci.Trova_business_unit_magazzino(row.Cells("dal_mag").Value) = 13 And Form_Entrate_Merci.Trova_business_unit_magazzino(row.Cells("al_mag").Value) <> 13)) Then
                        MsgBox("non è possibile trasferire materiale da Tirelli per  WIP BRB o viceversa")
                        Return
                    End If
                    ' Verifica se la cella della colonna "seleziona" è flaggata

                    If row.Cells("q_trasf").Value.ToString().Contains(",") Then
                        MsgBox("Attenzione alle ',' nella quantità ")
                        Return
                    End If
                    If Form_Entrate_Merci.check_giacenza_per_trasferimento(row.Cells("codice").Value, row.Cells("Dal_mag").Value, row.Cells("q_trasf").Value) = 1 Then


                        Beep()
                        MsgBox("ERRORE Si sta cercando di trasferire una quantità maggiore della giacenza OITW")


                        Return

                    End If

                    If Form_Entrate_Merci.check_giacenza_per_trasferimento_trasferimenti_oivl(row.Cells("codice").Value, row.Cells("Dal_mag").Value, row.Cells("q_trasf").Value) = 1 Then
                        Beep()
                        MsgBox("ERRORE Si sta cercando di trasferire una quantità maggiore della giacenza Ovl")
                        Return
                    End If

                    If row.Cells("dal_mag").Value = "FERRETTO" Or row.Cells("al_mag").Value = "FERRETTO" Then
                        Beep()
                        MsgBox("Non è possibile effettuare trasferimenti Da/Per Ferretto")
                        Return
                    End If

                    Form_Entrate_Merci.Trasferimento_in_WIP(row.Cells("DOC").Value, row.Cells("codice").Value, ODP_Form.ottieni_informazioni_odp("Docentry", row.Cells("docentry_dop").Value, 0).docnum, docnum_oc, Replace(row.Cells("q_trasf").Value, ",", "."), row.Cells("dal_mag").Value, row.Cells("al_mag").Value, linenum, row.Cells("al_mag").Value, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, "Manuale", 0, 0, Me.Text)
                    docnum_odp = ODP_Form.ottieni_informazioni_odp("Docentry", row.Cells("docentry_DOP").Value, 0).docnum

                    If row.Cells("al_mag").Value = "WIP" And row.Cells("Dal_mag").Value <> "02" Then

                        Form_Entrate_Merci.stampa_per_wip("ODP", row.Cells("codice").Value, ODP_Form.ottieni_informazioni_odp("Docentry", row.Cells("docentry_dop").Value, 0).docnum, Replace(row.Cells("q_trasf").Value, ",", "."), row.Cells("al_mag").Value)

                    End If



                End If
            Next
            If contatore_odp > 0 Then
                'MsgBox("Trasferimenti eseguiti con successo")
                ODP_Form.Show()
                ODP_Form.inizializza_form(docnum_odp)
                Me.Close()
            ElseIf contatore_oc > 0 Then
                'MsgBox("Trasferimenti eseguiti con successo")
                Form_nuova_offerta.Show()

                Form_nuova_offerta.TextBox10.Text = docnum_oc
                Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                Form_nuova_offerta.inizializzazione_form(docnum_oc, "ORDR", "RDR1", "OC")
                MsgBox("Trasferimento eseguito con successo")
                Me.Close()
            Else
                MsgBox("Non sono stati selezionati trasferimenti")
            End If



        End If




    End Sub

    Sub stampa_scontrino_da_trasf(par_n_trasferimento As Integer, par_itemcode As String)

        trova_info_trasferimento_mag(par_n_trasferimento, par_itemcode)
    End Sub

    Sub trova_info_trasferimento_mag(par_n_trasferimento As Integer, par_itemcode As String)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1




        CMD_SAP_2.CommandText = "

SELECT top 1 T1.[ItemCode], T1.quantity,T1.FROMWHSCOD,T1.WHSCODE, coalesce(T1.U_PRG_AZS_OPDOCENTRY,0) as 'docentry_odp'
, T1.U_PRG_AZS_OCDOCENTRY
, coalesce(T2.DOCNUM,0) as 'ODP', coalesce(T3.DOCNUM,0) as 'OC', T2.U_PRG_AZS_COMMESSA, COALESCE(T3.CARDNAME,'') AS 'CLIENTE' 

FROM OWTR T0  INNER JOIN WTR1 T1 ON T0.[DocEntry] = T1.[DocEntry]
LEFT JOIN OWOR T2 ON T2.DOCENTRY=T1.U_PRG_AZS_OPDOCENTRY
LEFT JOIN ORDR T3 ON T3.DOCENTRY=T1.U_PRG_AZS_OCDOCENTRY 
LEFT JOIN OCRD T4 ON T4.CARDCODE=T3.U_CodiceBP

WHERE T0.[DocNum] =" & par_n_trasferimento & " and t1.itemcode='" & par_itemcode & "'"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Dim contatore As Integer = 0
        If cmd_SAP_reader_2.Read() Then

            Dim tipo_doc As String = ""
            Dim numero_doc As Integer = 0
            If cmd_SAP_reader_2("docentry_odp") = 0 Then
                tipo_doc = "OC"
                numero_doc = cmd_SAP_reader_2("OC")

            Else
                tipo_doc = "ODP"
                numero_doc = cmd_SAP_reader_2("ODP")
            End If
            If cmd_SAP_reader_2("whscode") = "WIP" Then

                Form_Entrate_Merci.stampa_per_wip(tipo_doc, par_itemcode, numero_doc, cmd_SAP_reader_2("quantity"), cmd_SAP_reader_2("whscode"))
            Else
                Form_Entrate_Merci.compila_scontrino(par_itemcode, Magazzino.OttieniDettagliAnagrafica(par_itemcode).Descrizione, cmd_SAP_reader_2("quantity"), "")
                Form_Entrate_Merci.Fun_Stampa("Trasferimento interno", False, Form_Entrate_Merci.Stampante_Selezionata, Form_Entrate_Merci.Scontrino, cmd_SAP_reader_2("whscode"), Magazzino.OttieniDettagliAnagrafica(par_itemcode).Ubicazione, par_itemcode)
            End If

            End If

            cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedItem IsNot Nothing Then
            Dim filtro As String = ComboBox1.Text

            ' Itera attraverso le righe della DataGridView
            For Each riga As DataGridViewRow In DataGridView1.Rows
                ' Verifica se il valore nella colonna "dal_mag" è uguale al testo selezionato nella ComboBox
                If riga.Cells("dal_mag").Value IsNot Nothing AndAlso riga.Cells("dal_mag").Value.ToString() = filtro Then
                    riga.Visible = True ' Rendi visibile la riga se il filtro corrisponde
                Else
                    riga.Visible = False ' Nascondi la riga se il filtro non corrisponde
                End If
            Next
        Else
            ' Se la ComboBox non ha un elemento selezionato, mostra tutte le righe
            For Each riga As DataGridViewRow In DataGridView1.Rows
                riga.Visible = True
            Next
        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Codice) Then

                Magazzino.Codice_SAP = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice").Value
                ' Ripristina la finestra se è minimizzata
                If Magazzino.WindowState = FormWindowState.Minimized Then
                    Magazzino.WindowState = FormWindowState.Normal
                End If

                ' Porta la finestra in primo piano
                Magazzino.BringToFront()
                Magazzino.Activate()
                Magazzino.Show()

                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)


            End If
        End If
    End Sub

    Private Sub DataGridView1_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDown
        Dim par_datagridview As DataGridView = CType(sender, DataGridView)

        If e.Button = MouseButtons.Left AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If ModifierKeys = Keys.Shift AndAlso startIndex >= 0 Then
                Dim endIndex As Integer = e.RowIndex
                Dim minIndex As Integer = Math.Min(startIndex, endIndex)
                Dim maxIndex As Integer = Math.Max(startIndex, endIndex)

                ' Escludi la riga cliccata dall'intervallo
                If endIndex > startIndex Then
                    maxIndex -= 1
                ElseIf endIndex < startIndex Then
                    minIndex += 1
                Else
                    ' Se si clicca sulla stessa riga, non fare nulla
                    Return
                End If

                ' Verifica se tutte le righe dell'intervallo sono selezionate
                Dim allSelected As Boolean = True
                For i As Integer = minIndex To maxIndex
                    If Not Convert.ToBoolean(par_datagridview.Rows(i).Cells(0).Value) Then
                        allSelected = False
                        Exit For
                    End If
                Next

                ' Applica selezione o deselezione in blocco
                For i As Integer = minIndex To maxIndex
                    par_datagridview.Rows(i).Cells(0).Value = Not allSelected
                Next
            Else
                startIndex = e.RowIndex
            End If
        End If
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        ' Controlla se il tasto Shift è stato premuto
        isShiftKeyDown = (e.KeyCode = Keys.ShiftKey)
    End Sub

    Private Sub DataGridView1_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyUp
        ' Controlla se il tasto Shift è stato rilasciato
        isShiftKeyDown = (e.KeyCode <> Keys.ShiftKey)
    End Sub

End Class