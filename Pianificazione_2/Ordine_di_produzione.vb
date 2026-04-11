Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports Word = Microsoft.Office.Interop.Word


Public Class Ordine_di_produzione
    Public riga_ As Integer
    Public riga_1 As Integer
    Public riga_2 As Integer
    Public linenum_ODP AS Integer
    Public linenum_ODP_1 As Integer
    Public linenum_ODP_2 As Integer
    Public Visorder_ODP As Integer
    Public Visorder_ODP_1 As Integer
    Public Visorder_ODP_2 As Integer
    Public cognome As String
    Public quantità_riga As String
    Public destinatario_mu_1 As String = "vanniponti@tirelli.net"
    Public destinatario_mu_2 As String = "macchineutensili@tirelli.net"
    Public percorso_documento As String
    Public docentry_odp As Integer



    Public testata_odp_itemcode As String
    Public testata_odp_prodname As String
    Public testata_odp_u_disegno As String
    Public testata_odp_data As String
    Public testata_odp_plannedqty As Integer
    Public testata_odp_resname As String
    Public testata_odp_u_produzione As String
    Public testata_odp_commessa As String
    Public testata_odp_max_righe As String
    Public testata_odp_docnum As String
    Public stampa_etichetta As String
    public testata_odp_cardname As String
    public testata_odp_U_Clientefinale As String
    Public testata_odp_docnum_oc As String
    Public testata_odp_cliente_eti As String
    Public testata_odp_Itemname_commessa As String
    Public magazzino_riga As String

    Public oWord As Word.Application
    Public oDoc As Word.Document
    Public oTable As Word.Table


    Public codice_riga As String





    Private Sub DataGridView_ODP_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_ODP.CellFormatting
        Try

            If DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Azione").Value = "Trasferibile" Then
                DataGridView_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
            ElseIf DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Azione").Value = "Trasferibile/Da ordinare" Then
                DataGridView_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.MediumSpringGreen
            ElseIf DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Azione").Value = "IN APPROV/DA ORDINARE" Then
                DataGridView_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Orange
            ElseIf DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Azione").Value = "IN APPROV" Then
                DataGridView_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Khaki
            ElseIf DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Azione").Value = "Da ordinare" Then
                DataGridView_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Red
            End If

            If DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Stato").Value = "C" Then
                DataGridView_ODP.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
            End If
        Catch ex As Exception

        End Try

    End Sub


    Private Sub DataGridView_ODP_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellClick
        If e.RowIndex >= 0 Then
            riga_ = e.RowIndex
            linenum_ODP = DataGridView_ODP.Rows(riga_).Cells(columnName:="Linenum").Value
            quantità_riga = DataGridView_ODP.Rows(riga_).Cells(columnName:="Quantità").Value
            magazzino_riga = DataGridView_ODP.Rows(riga_).Cells(columnName:="MAG").Value
            codice_riga = DataGridView_ODP.Rows(riga_).Cells(columnName:="Codice").Value

            If e.ColumnIndex = DataGridView_ODP.Columns.IndexOf(Codice) Then

                Magazzino.Codice_SAP = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Codice").Value






                Magazzino.Show()

                Magazzino.TextBox2.Text = Magazzino.Codice_SAP
                Magazzino.dettagli_anagrafica(Magazzino.Codice_SAP)
                Me.WindowState = FormWindowState.Minimized

            End If
            If e.ColumnIndex = DataGridView_ODP.Columns.IndexOf(Disegno) Then

                Try
                    Process.Start("\\192.168.0.150\k\Tecnico\Disegni Meccanici\PDF-DXF\PDF\" & DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Disegno").Value & ".PDF")
                Catch ex As Exception
                    MsgBox("Il disegno " & DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Disegno").Value & " non è ancora stato processato")
                End Try


            End If

            If e.ColumnIndex = DataGridView_ODP.Columns.IndexOf(ODP) Then
                FORM6.ODP = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="ODP").Value
                If FORM6.ODP = Nothing Then
                    MsgBox("Scegliere un ordine di produzione")
                Else

                    riempi_ODP()
                    intestazioni_ordine_di_produzione()
                    Dashboard_MU_New.mu = 0
                    Trasferibili.trasferibili = 0

                End If


            End If

            Try


                Inventario.Codice_SAP = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Codice").Value
                TextBox_codice_riga.Text = Inventario.Codice_SAP
                Inventario.dettagli_anagrafica()
                Label3.Text = Inventario.Descrizione
                Label4.Text = Inventario.disegno
                TextBox_quantità.Text = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Quantità").Value
                Label5.Text = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="linenum").Value
                Label8.Text = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Trasferito").Value
                ComboBox1.Text = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Stato").Value
            Catch ex As Exception

            End Try

            Try

                If File.Exists("\\192.168.0.150\k\Tecnico\Disegni Meccanici\PDF-DXF\PDF\" & DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Disegno").Value & ".PDF") Then
                    AxFoxitCtl2.OpenFile("\\192.168.0.150\k\Tecnico\Disegni Meccanici\PDF-DXF\PDF\" & DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="Disegno").Value & ".PDF")
                    Label1.Hide()
                    AxFoxitCtl2.Show()
                Else
                    AxFoxitCtl2.Hide()
                    Label1.Show()
                End If

            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Process.Start("\\192.168.0.150\k\Tecnico\Disegni Meccanici\PDF-DXF\PDF\" & Button1.Text & ".PDF")
        Catch ex As Exception
            MsgBox("Il disegno " & Button1.Text & " non è ancora stato processato")
        End Try
    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click



        If Dashboard_MU_New.mu = 1 Then
            Try


                Carico_macchine.Check_Lavorazioni_aperte_macchina()
            Catch ex As Exception

            End Try
        End If
        Me.Close()

    End Sub

    Sub CHECK_ARTICOLo()
        Form_principale.Cnn.ConnectionString = Form_principale.SAP
        Form_principale.Cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand
        Dim cmd_SAP_reader_7 As SqlDataReader
        CMD_SAP_7.Connection = Form_principale.Cnn

        CMD_SAP_7.CommandText = "SELECT T0.VALIDFOR AS 'Valido' FROM OITM T0 WHERE T0.[ITEMCODE]= '" & Inventario.Codice_SAP & "' AND T1.VALIDFOR='N'"

        cmd_SAP_reader_7 = CMD_SAP_7.ExecuteReader
        If cmd_SAP_reader_7.Read() = True Then

            MsgBox("il codice " & Inventario.Codice_SAP & " è inattivo ")
        Else



        End If
        cmd_SAP_reader_7.Close()
        Form_principale.Cnn.Close()


    End Sub

    Private Sub TextBox_codice_riga_TextChanged(sender As Object, e As EventArgs) Handles TextBox_codice_riga.TextChanged
        Inventario.Codice_SAP = TextBox_codice_riga.Text
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        cognome = InputBox("Inserire proprio cognome")
        If UCase(cognome) = "PONTI" Or UCase(cognome) = "GIANGI" Then
            AGGIORNA_ODP_riga()
            riempi_ODP()

        Else

            AGGIORNA_ODP_riga()
            riempi_ODP()
            invio_mail_VARIAZIONE_QUANTITà()
        End If
    End Sub

    Sub invio_mail_VARIAZIONE_QUANTITà()
        Dim TESTO_MAIL As String



        TESTO_MAIL = TESTO_MAIL & "Variazione quantità " & TextBox_codice_riga.Text & " " & Label3.Text & " Da " & quantità_riga & " a " & TextBox_quantità.Text


        TESTO_MAIL = TESTO_MAIL & "</tr>"
        TESTO_MAIL = TESTO_MAIL & "</table>"

        TESTO_MAIL = TESTO_MAIL & "</BODY>"
        ' Invio E-Mail


        Dim mySmtp As New SmtpClient
        Dim myMail As New MailMessage()



        mySmtp.UseDefaultCredentials = False
        mySmtp.Credentials = New Net.NetworkCredential(Form_principale.sender_mail, Form_principale.Password_Mail)
        mySmtp.Host = "smtp.office365.com"
        mySmtp.Port = 25
        mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network
        mySmtp.EnableSsl = True


        myMail = New MailMessage()

        myMail.From = New MailAddress(Form_principale.sender_mail)

        myMail.To.Add(destinatario_mu_1)

        myMail.To.Add(destinatario_mu_2)


        myMail.Subject = "VARIAZIONE RIGHE ODP Da " & UCase(cognome) & " ODP N° " & Label_numero_ODP_F.Text & " " & Label_descrizione_F.Text
        myMail.IsBodyHtml = True
        myMail.Body = TESTO_MAIL

        Try
            mySmtp.Send(myMail)
        Catch ex As Exception

        End Try

        TESTO_MAIL = Nothing

    End Sub



    Sub invio_mail_aggiungere_riga()
        Dim TESTO_MAIL As String


        TESTO_MAIL = TESTO_MAIL & "Aggiunta riga " & TextBox1.Text & " " & Label6.Text & " quantità " & TextBox2.Text


        TESTO_MAIL = TESTO_MAIL & "</tr>"
        TESTO_MAIL = TESTO_MAIL & "</table>"

        TESTO_MAIL = TESTO_MAIL & "</BODY>"
        ' Invio E-Mail


        Dim mySmtp As New SmtpClient
        Dim myMail As New MailMessage()



        mySmtp.UseDefaultCredentials = False
        mySmtp.Credentials = New Net.NetworkCredential(Form_principale.sender_mail, Form_principale.Password_Mail)
        mySmtp.Host = "smtp.office365.com"
        mySmtp.Port = 25
        mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network
        mySmtp.EnableSsl = True


        myMail = New MailMessage()

        myMail.From = New MailAddress(Form_principale.sender_mail)

        myMail.To.Add(destinatario_mu_1)

        myMail.To.Add(destinatario_mu_2)


        myMail.Subject = "AGGIUNTA RIGA ODP Da " & UCase(cognome) & " ODP N° " & Label_numero_ODP_F.Text & " " & Label_descrizione_F.Text
        myMail.IsBodyHtml = True
        myMail.Body = TESTO_MAIL

        Try
            mySmtp.Send(myMail)
        Catch ex As Exception

        End Try

        TESTO_MAIL = Nothing

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If DataGridView_ODP.Rows(riga_).Cells(columnName:="Trasferito").Value > 0 Then

            MsgBox("Trasferito >0 , rendere il pezzo prima di eliminare la riga")

        Else

            Dim Question
            Question = MsgBox("Sei sicuro di voler eliminare il codice " & DataGridView_ODP.Rows(riga_).Cells(columnName:="Codice").Value & " ?", vbYesNo)
            If Question = vbYes Then

                Analisi_riga_magazzino.CODICE_confermato = DataGridView_ODP.Rows(riga_).Cells(columnName:="Codice").Value
                linenum_ODP = DataGridView_ODP.Rows(riga_).Cells(columnName:="Linenum").Value
                elimina_riga()
                Analisi_riga_magazzino.ripara_confermati()

            End If
        End If
    End Sub

    Sub elimina_riga()
        Form_principale.Cnn3.ConnectionString = Form_principale.SAP
        Form_principale.Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Form_principale.Cnn3


        CMD_SAP_3.CommandText = "DELETE T0 FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] WHERE T1.[DocNum] =" & FORM6.ODP & " AND  T0.[ItemCode] ='" & Analisi_riga_magazzino.CODICE_confermato & "' AND  T0.[LineNum] =" & linenum_ODP & ""

        CMD_SAP_3.ExecuteNonQuery()
        Form_principale.Cnn3.Close()
        For Each Riga As DataGridViewRow In DataGridView_ODP.SelectedRows
            DataGridView_ODP.Rows.RemoveAt(riga_)
        Next

    End Sub

    Sub inserisci_riga()
        Form_principale.Cnn3.ConnectionString = Form_principale.SAP
        Form_principale.Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Form_principale.Cnn3


        CMD_SAP_3.CommandText = "insert into WOR1 (WOR1.DocEntry, WOR1.LineNum, WOR1.ItemCode, WOR1.BaseQty, WOR1.PlannedQty, WOR1.IssuedQty, WOR1.IssueType, WOR1.wareHouse, WOR1.VisOrder, WOR1.WipActCode, WOR1.CompTotal, WOR1.OcrCode, WOR1.OcrCode2, WOR1.OcrCode3, WOR1.OcrCode4, WOR1.OcrCode5, WOR1.LocCode, WOR1.Project, WOR1.UomEntry, WOR1.UomCode, WOR1.ItemType, WOR1.AdditQty, WOR1.LineText, WOR1.PickStatus, WOR1.PickQty, WOR1.PickIdNo, WOR1.ReleaseQty, WOR1.ResAlloc, WOR1.StartDate, WOR1.EndDate, WOR1.StageId, WOR1.BaseQtyNum, WOR1.BaseQtyDen, WOR1.ReqDays, WOR1.RtCalcProp, WOR1.Status, WOR1.ItemName, WOR1.AlwProcDoc, WOR1.PoDocType, WOR1.PoDocNum, WOR1.PoDocEntry, WOR1.PoLineNum, WOR1.PoQuantity, WOR1.U_TEMPOME, WOR1.U_UBIMAG, WOR1.U_Prezzolis, WOR1.U_CodDis, WOR1.U_PRG_AZS_Terzista, WOR1.U_PRG_AZS_StatoAv, WOR1.U_PRG_AZS_PhanFat, WOR1.U_PRG_CLV_Ris_Orig, WOR1.U_PRG_CLV_Qta_Trasf, WOR1.U_IEO_LPN_QTASPED, WOR1.U_DisponiibileTOT, WOR1.U_PRG_WIP_QtaSpedita, WOR1.U_PRG_WIP_QtaDaTrasf, WOR1.U_Data_ora_inizio_fase, WOR1.U_Data_ora_fine_fase, WOR1.U_Dipendente, WOR1.U_Ordinato_TOT, WOR1.U_Confermato_TOT, WOR1.U_PRG_WIP_QtaRichMagAuto, WOR1.U_PRG_WMS_Exp, WOR1.U_PRG_WMS_ExpDate, WOR1.U_PRG_WMS_MdMovQty, WOR1.U_Stato_lavorazione)
                                                    SELECT t1.docentry, max(t2.linenum)+1,'" & TextBox1.Text & "' , " & TextBox2.Text & " / T1.PLANNEDQTY, " & TextBox2.Text & ",0,'B',CASE WHEN SUBSTRING('" & TextBox1.Text & "',1,1)='R' THEN 'RIS' ELSE case when T3.dfltwh is null then '01' else T3.dfltwh end END,max(t2.visorder)+1,'',0,'','','','','','','','','',CASE WHEN SUBSTRING('" & TextBox1.Text & "',1,1)='R' THEN 290 ELSE 4 END,0,'','N',0,0,0,CASE WHEN SUBSTRING('" & TextBox1.Text & "',1,1)='R' then 'F' else null end,T1.[STARTDate], T1.[DueDate],NULL,0,0,0,100,T1.STATUS,T3.itemname,'N','','','','',0,0,'',T4.PRICE,'','','X','','',0,0,0,0,CASE WHEN SUBSTRING('" & TextBox1.Text & "',1,1)='R' THEN 0 ELSE " & TextBox2.Text & " END ,'','','',0,0,0,'N','',0,'O'
FROM OWOR T1 inner join wor1 t2 on t1.docentry=t2.docentry
inner join oitm t3 on '" & TextBox1.Text & "' =t3.itemcode
inner join itm1 t4 on '" & TextBox1.Text & "' =t4.itemcode
WHERE (T1.STATUS ='P' OR T1.STATUS ='R') AND T1.DOCNUM=" & FORM6.ODP & " and t4.pricelist=2

group by t1.docentry, T1.PLANNEDQTY,T1.[STARTDate], T1.[DueDate],T1.STATUS,T3.itemname,T4.PRICE,t3.dfltwh"

        CMD_SAP_3.ExecuteNonQuery()
        Form_principale.Cnn3.Close()


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        cognome = InputBox("Inserire proprio cognome")
        If UCase(cognome) = "-" Or UCase(cognome) = "." Then
            inserisci_riga_odp()
        Else
            inserisci_riga_odp()
            invio_mail_aggiungere_riga()
        End If
    End Sub

    Sub inserisci_riga_odp()
        If TextBox2.Text <> Nothing Then
            If TextBox2.Text > 0 Then

                TextBox1.Text = UCase(TextBox1.Text)

                Form_principale.Cnn.ConnectionString = Form_principale.SAP
                Form_principale.Cnn.Open()

                Dim CMD_SAP_7 As New SqlCommand
                Dim cmd_SAP_reader_7 As SqlDataReader
                CMD_SAP_7.Connection = Form_principale.Cnn

                CMD_SAP_7.CommandText = "SELECT T1.VALIDFOR AS 'Valido', t1.itemcode as 'Codice' FROM OITM T1 WHERE T1.[itemcode]= '" & TextBox1.Text & "'"

                cmd_SAP_reader_7 = CMD_SAP_7.ExecuteReader
                If cmd_SAP_reader_7.Read() = True Then
                    If cmd_SAP_reader_7("Valido") = "N" Then
                        MsgBox("Il codice " & TextBox1.Text & " è inattivo ")
                    Else
                        inserisci_riga()
                        riempi_ODP()
                        Analisi_riga_magazzino.CODICE_confermato = TextBox1.Text
                        Analisi_riga_magazzino.ripara_confermati()
                    End If

                Else
                    MsgBox("Il codice " & TextBox1.Text & " non esiste ")
                End If
                Form_principale.Cnn.Close()
            Else
                MsgBox("Inserire un numero >0")
            End If
        Else
            MsgBox("Inserire una quantità esatta")
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Magazzino.Codice_SAP = TextBox1.Text
        Magazzino.dettagli_anagrafica(Magazzino.Codice_SAP)
        Label7.Text = Magazzino.disegno
        Label6.Text = Magazzino.Descrizione
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button_down.Click

        If riga_ < DataGridView_ODP.RowCount - 1 Then


            riga_1 = riga_ + 1
            linenum_ODP = DataGridView_ODP.Rows(riga_).Cells(columnName:="Linenum").Value
            linenum_ODP_1 = DataGridView_ODP.Rows(riga_1).Cells(columnName:="Linenum").Value
            Visorder_ODP = DataGridView_ODP.Rows(riga_).Cells(columnName:="Visorder").Value
            Visorder_ODP_1 = DataGridView_ODP.Rows(riga_1).Cells(columnName:="Visorder").Value
            switch_riga()
            riempi_ODP()
            DataGridView_ODP.Rows(riga_).Selected = False
            riga_ = riga_ + 1
            'DataGridView_ODP.SelectedRows.Clear()
            DataGridView_ODP.Rows(riga_).Selected = True
        End If
    End Sub

    Sub switch_riga()
        Form_principale.Cnn3.ConnectionString = Form_principale.SAP
        Form_principale.Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Form_principale.Cnn3


        CMD_SAP_3.CommandText = "UPDATE T0 SET T0.LINENUM = case when t0.linenum= " & linenum_ODP & " then " & linenum_ODP_1 & " else " & linenum_ODP & " end , t0.visorder = case when t0.visorder= " & Visorder_ODP & " then " & Visorder_ODP_1 & " else " & Visorder_ODP & " end  from wor1 t0 inner join owor t1 on t0.docentry=t1.docentry WHERE (T0.LINENUM= " & linenum_ODP_1 & " or T0.LINENUM= " & linenum_ODP & ") and t1.docnum=" & FORM6.ODP & " "

        CMD_SAP_3.ExecuteNonQuery()



        Form_principale.Cnn3.Close()


    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If riga_ > 0 Then
            riga_1 = riga_ - 1
            linenum_ODP = DataGridView_ODP.Rows(riga_).Cells(columnName:="Linenum").Value
            linenum_ODP_1 = DataGridView_ODP.Rows(riga_1).Cells(columnName:="Linenum").Value
            Visorder_ODP = DataGridView_ODP.Rows(riga_).Cells(columnName:="Visorder").Value
            Visorder_ODP_1 = DataGridView_ODP.Rows(riga_1).Cells(columnName:="Visorder").Value
            switch_riga()
            riempi_ODP()
            DataGridView_ODP.Rows(riga_).Selected = False
            riga_ = riga_ - 1
            ' DataGridView_ODP.SelectedRows.Clear()
            DataGridView_ODP.Rows(riga_).Selected = True

        End If
    End Sub


    Sub AGGIORNA_ODP_riga()


        Form_principale.Cnn3.ConnectionString = Form_principale.SAP
            Form_principale.Cnn3.Open()

            Dim CMD_SAP_3 As New SqlCommand

            CMD_SAP_3.Connection = Form_principale.Cnn3


        CMD_SAP_3.CommandText = "UPDATE T0 SET T0.[PlannedQty]=" & Replace(TextBox_quantità.Text, ",", ".") & ", T0.[U_PRG_WIP_QtaDaTrasf]= CASE WHEN SUBSTRING('" & TextBox_codice_riga.Text & "',1,1)='R' THEN 0 ELSE " & Replace(TextBox_quantità.Text, ",", ".") & "- " & Replace(Label8.Text, ",", ".") & " END, T0.U_STATO_LAVORAZIONE= '" & ComboBox1.Text & "'  
from wor1 t0 inner join owor t1 on t0.docentry=t1.docentry WHERE T0.LINENUM= " & linenum_ODP & " and t1.docnum=" & FORM6.ODP & " "

        CMD_SAP_3.ExecuteNonQuery()

            Form_principale.Cnn3.Close()

    End Sub
    Sub AGGIORNA_ODP_doc()
        Form_principale.Cnn.ConnectionString = Form_principale.SAP
        Form_principale.Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Form_principale.Cnn

        CMD_SAP.CommandText = "UPDATE owor SET STATUS='" & ComboBox_stato_odp.Text & "', u_lavorazione=" & TextBox3.Text & ", u_stato='" & ComboBox2.Text & "', u_aggiorna_db='" & ComboBox3.Text & "' WHERE DOCNUM ='" & FORM6.ODP & "'"
        CMD_SAP.ExecuteNonQuery()
        CMD_SAP.CommandText = "UPDATE wor1 SET wor1.STATUS='" & ComboBox_stato_odp.Text & "' from owor inner join wor1 on owor.docentry=wor1.docentry WHERE DOCNUM ='" & FORM6.ODP & "'"
        CMD_SAP.ExecuteNonQuery()
        Form_principale.Cnn.Close()
    End Sub



    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        AGGIORNA_ODP_doc()
        MsgBox("ODP aggiornato")
    End Sub



    Private Sub Button13_Click(sender As Object, e As EventArgs)
        Form_nuovo_ticket.Show()
        Form_nuovo_ticket.ComboBox2.Text = Homepage.business
        Form_nuovo_ticket.Inserimento_dipendenti()
        Me.Hide()
        Form_nuovo_ticket.Owner = Me
        Form_nuovo_ticket.Reparto = 2
        Form_nuovo_ticket.Administrator = 0
        Form_nuovo_ticket.Startup()
        Form_nuovo_ticket.Txt_Commessa.Text = Form_principale.commessa
    End Sub

    Private Sub Button13_Click_1(sender As Object, e As EventArgs) Handles Button13.Click
        Form_nuovo_ticket.Show()
        Form_nuovo_ticket.ComboBox2.Text = Homepage.business
        Form_nuovo_ticket.Inserimento_dipendenti()
        Me.Hide()
        Form_nuovo_ticket.Owner = Me
        Form_nuovo_ticket.Administrator = 1
        Form_nuovo_ticket.Startup()
        Form_nuovo_ticket.Txt_Commessa.Text = Form_principale.commessa
        Form_nuovo_ticket.Combo_Riferimenti.SelectedIndex = 1
        Form_nuovo_ticket.Txt_Nuovo_Riferimento.Text = FORM6.ODP
        Form_nuovo_ticket.Cmd_Aggiungi_Riferimento.PerformClick()
    End Sub

    Sub Genera_ordine()

        testata_odp()


        oWord = CreateObject("Word.Application")

        oDoc = oWord.Documents.Add("" & Percorso_documento & "")

        segnalibri_testata_odp()
        If stampa_etichetta = "YES" Then
            segnalibri_etichetta_cassetta()
        End If


        oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("Tabella").Range, testata_odp_max_righe + 1, 7)

        oTable.Cell(1, 1).Range.Text = "COD"
        oTable.Cell(1, 2).Range.Text = "Descrizione"
        oTable.Cell(1, 3).Range.Text = "U.M."
        oTable.Cell(1, 4).Range.Text = "Q.TA"
        oTable.Cell(1, 5).Range.Text = "T"
        oTable.Cell(1, 6).Range.Text = "MAG"
        oTable.Cell(1, 7).Range.Text = "UBIC"



        Form_principale.Cnn1.ConnectionString = Form_principale.SAP
        Form_principale.Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader
        Dim i As Integer = 2

        CMD_SAP_2.Connection = Form_principale.Cnn1
        CMD_SAP_2.CommandText = "Select T10.[ItemCode], T10.[ItemName], T10.[InvntryUom], T10.[PlannedQty], case when t10.u_prg_wip_qtaspedita is null then 0 else t10.u_prg_wip_qtaspedita end as 'u_prg_wip_qtaspedita' , T10.[wareHouse], t10.u_ubicazione , t10.Trasferibile
from
(
SELECT T1.[ItemCode], T2.[ItemName], case when T2.[InvntryUom] is null then '' else T2.[InvntryUom] end as 'Invntryuom', T1.[PlannedQty],  t1.u_prg_wip_qtaspedita  , T1.[wareHouse], case when t2.u_ubicazione is null then '' else t2.u_ubicazione end as 'U_ubicazione' , case when t3.onhand>=T1.[PlannedQty]- case when t1.u_prg_wip_qtaspedita is null then 0 else t1.u_prg_wip_qtaspedita end then 1 else 2 end as 'Trasferibile'

FROM OWOR T0  INNER JOIN WOR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 
inner join OITM T2 on t2.itemcode=t1.itemcode 
inner join oitw t3 on t3.itemcode=t1.itemcode and t3.whscode=T1.[wareHouse]
WHERE T0.[DocNum] ='" & FORM6.ODP & "' and t1.itemtype=4
)
as t10
order by t10.Trasferibile,T10.[wareHouse],t10.u_ubicazione, t10.itemcode"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            oTable.Cell(i, 1).Range.Text = cmd_SAP_reader_2("ItemCode")
            oTable.Cell(i, 2).Range.Text = cmd_SAP_reader_2("ItemName")
            oTable.Cell(i, 3).Range.Text = cmd_SAP_reader_2("InvntryUom")

            oTable.Cell(i, 4).Range.Text = FormatNumber(cmd_SAP_reader_2("PlannedQty"), 1, , , TriState.True)



            If cmd_SAP_reader_2("u_prg_wip_qtaspedita") = 0 Then
                    oTable.Cell(i, 5).Range.Text = ""
                Else
                    oTable.Cell(i, 5).Range.Text = FormatNumber(cmd_SAP_reader_2("u_prg_wip_qtaspedita"), 1, , , TriState.True)
                End If


            oTable.Cell(i, 6).Range.Text = cmd_SAP_reader_2("wareHouse")
            oTable.Cell(i, 7).Range.Text = cmd_SAP_reader_2("u_ubicazione")


            oTable.Cell(i, 2).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify
            oTable.Rows(i).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter



            i = i + 1
        Loop



        Form_principale.Cnn1.Close()

        'oTable.Range.ParagraphFormat.SpaceAfter = 6

        'oTable.Cell(r, 1).Range.Text = r - 1
        'oTable.Cell(r, 1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
        'oTable.Cell(r, 2).Range.Text = Codice
        'ble.Cell(r, 3).Range.Text = Descrizione
        'With oTable.Cell(r, 3).Range.Font
        '.Name = "Arial"
        '.Bold = 1
        ' If Len(Descrizione) < 35 Then
        '.Size = 9
        'Else
        '.Size = 7
        'End If

        'End With '
        'oTable.Cell(r, 4).Range.Text = Note
        'oTable.Cell(r, 5).Range.Text = Quantità
        'oTable.Cell(r, 5).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
        'oTable.Cell(r, 6).Range.Text = Valuta & " " & FormatNumber(Prezzo, 2, , , TriState.True)
        'oTable.Cell(r, 6).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight
        'oTable.Cell(r, 7).Range.Text = Valuta & " " & FormatNumber(Totale, 2, , , TriState.True)
        ' oTable.Cell(r, 7).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight

        'Next

        oTable.AutoFormat(ApplyColor:=False, ApplyBorders:=False)
        oTable.Rows.Item(1).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter
        oTable.Rows.Item(1).Range.Font.Bold = True
        'oTable.Columns.Item(1).Width = oWord.InchesToPoints(1)   'Change width of columns 1 & 2
        oTable.Rows(1).Cells.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        oTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        oTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle
        oTable.AutoFitBehavior(Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow)


        'oWord.Visible = True
        'oWord.ShowMe()

        oWord.PrintOut()



        'oWord.Quit()

        oWord.Documents.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
        If oWord.Documents.Count = 0 Then
            oWord.Application.Quit()
        End If

        'oWord.Application.Quit

    End Sub

    Sub testata_odp()


        Form_principale.Cnn1.Close()
        Form_principale.Cnn1.ConnectionString = Form_principale.SAP
        Form_principale.Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Form_principale.Cnn1
        CMD_SAP_2.CommandText = "SELECT T0.[DocNum], T0.[ItemCode], substring(t0.prodname,1, 50) as 'prodname', case when t0.u_disegno is null then '' else t0.u_disegno end as 'u_disegno', getdate() as 'Data', T0.[PlannedQty], case when t1.resname is null then '' else t1.resname end as 'resname', t0.u_produzione , t0.u_prg_azs_commessa, count(t2.itemcode) as 'Max_righe', case when t3.cardname is null then '' else t3.cardname end as 'Cardname', case when t3.U_Clientefinale is null then '' else t3.U_Clientefinale end as 'U_Clientefinale',  case when t3.Docnum is null then '' else t3.Docnum end as 'Docnum_OC', case when t4.itemname is null then '' else t4.itemname end as 'itemname_Commessa', case when t3.u_clientefinale is null then case when t3.cardname is null then '' else t3.cardname end else t3.U_clientefinale end  as 'Cliente_Eti' 
FROM OWOR T0 left join orsc t1 on t1.visrescode=t0.u_fase left join wor1 t2 on t2.docentry=t0.docentry and t2.itemtype=4 left join ordr t3 on t3.docnum=t0.originnum left join oitm t4 on t4.itemcode=t0.u_prg_azs_commessa WHERE T0.[DocNum] ='" & FORM6.ODP & "' group by T0.[DocNum], T0.[ItemCode], t0.prodname, t0.u_disegno , T0.[PlannedQty], t1.resname, t0.u_produzione , t0.u_prg_azs_commessa, t3.docnum, t3.cardname, t3.u_clientefinale, t3.Docnum, t4.itemname"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() = True Then
            testata_odp_docnum = cmd_SAP_reader_2("DOCNUM")
            testata_odp_itemcode = cmd_SAP_reader_2("ItemCode")
            testata_odp_prodname = cmd_SAP_reader_2("prodname")

            testata_odp_u_disegno = cmd_SAP_reader_2("u_disegno")
            testata_odp_data = cmd_SAP_reader_2("data")
            testata_odp_plannedqty = cmd_SAP_reader_2("plannedqty")
            testata_odp_resname = cmd_SAP_reader_2("resname")
            testata_odp_u_produzione = cmd_SAP_reader_2("u_produzione")
            testata_odp_commessa = cmd_SAP_reader_2("u_prg_azs_commessa")
            testata_odp_max_righe = cmd_SAP_reader_2("Max_righe")

            testata_odp_cardname = cmd_SAP_reader_2("Cardname")
            testata_odp_U_Clientefinale = cmd_SAP_reader_2("U_clientefinale")
            testata_odp_docnum_oc = cmd_SAP_reader_2("Docnum_oc")

            testata_odp_cliente_eti = cmd_SAP_reader_2("Cliente_eti")
            testata_odp_Itemname_commessa = cmd_SAP_reader_2("Itemname_Commessa")


        End If
        cmd_SAP_reader_2.Close()
        Form_principale.Cnn1.Close()



    End Sub

    Sub segnalibri_testata_odp()

        oDoc.Bookmarks.Item("cod_f").Range.Text = testata_odp_itemcode
        oDoc.Bookmarks.Item("commessa").Range.Text = testata_odp_commessa
        oDoc.Bookmarks.Item("descrizione_f").Range.Text = testata_odp_prodname
        oDoc.Bookmarks.Item("data").Range.Text = testata_odp_data
        oDoc.Bookmarks.Item("disegno").Range.Text = testata_odp_u_disegno

        oDoc.Bookmarks.Item("odp").Range.Text = testata_odp_docnum
        oDoc.Bookmarks.Item("qta_f").Range.Text = Math.Round(testata_odp_plannedqty, 1)
        oDoc.Bookmarks.Item("Fase").Range.Text = testata_odp_resname
        oDoc.Bookmarks.Item("Produzione").Range.Text = testata_odp_u_produzione
        oDoc.Bookmarks.Item("Cliente").Range.Text = testata_odp_cardname
        oDoc.Bookmarks.Item("Utilizzatore").Range.Text = testata_odp_U_Clientefinale
        oDoc.Bookmarks.Item("OC").Range.Text = testata_odp_docnum_oc

    End Sub

    Sub segnalibri_etichetta_cassetta()
        oDoc.Bookmarks.Item("ODP_eti").Range.Text = FORM6.ODP
        oDoc.Bookmarks.Item("cod_eti").Range.Text = testata_odp_itemcode
        oDoc.Bookmarks.Item("fase_eti").Range.Text = testata_odp_resname
        oDoc.Bookmarks.Item("data_eti").Range.Text = testata_odp_data
        oDoc.Bookmarks.Item("Desc_eti").Range.Text = testata_odp_prodname
        oDoc.Bookmarks.Item("prod_eti").Range.Text = testata_odp_u_produzione

        oDoc.Bookmarks.Item("commessa_eti").Range.Text = testata_odp_commessa
        oDoc.Bookmarks.Item("cliente_eti").Range.Text = testata_odp_cliente_eti
        oDoc.Bookmarks.Item("oc_eti").Range.Text = testata_odp_docnum_oc
        oDoc.Bookmarks.Item("modello").Range.Text = testata_odp_Itemname_commessa
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        stampa_etichetta = "no"


        If CheckBox1.Checked = True And CheckBox2.Checked = False Then
            percorso_documento = "W:\Report aziendali\Visual basics\Word stampa automatica\ordine di produzione.docx"
            Genera_ordine()
        ElseIf CheckBox2.Checked = True Then
            percorso_documento = "W:\Report aziendali\Visual basics\Word stampa automatica\ordine di produzione + etichetta.docx"
            stampa_etichetta = "YES"
            Genera_ordine()
        End If
        If CheckBox3.Checked = True Then
            testata_odp()
            If testata_odp_u_disegno = "" Then
                MsgBox("Disegno non presente")
            Else
                AxFoxitCtl1.OpenFile("\\192.168.0.150\k\Tecnico\Disegni Meccanici\PDF-DXF\PDF\" & testata_odp_u_disegno & ".PDF")
                AxFoxitCtl1.PrintFile()
            End If

        End If



    End Sub

    Private Sub Cmd_Materiale_Click(sender As Object, e As EventArgs) Handles Cmd_Materiale.Click
        Form_Richiesta_Materiale.Show()
        Form_Richiesta_Materiale.Owner = Me
        Form_Richiesta_Materiale.Txt_Commessa.Text = Label_commessa_F.Text
        Form_Richiesta_Materiale.TXT_ODP.Text = Label_numero_ODP_F.Text
        Form_Richiesta_Materiale.Home_Lista()
        Me.Hide()
    End Sub


















    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Richiesta_trasferimento_materiale.docnum_odp = Label_numero_ODP_F.Text
        Richiesta_trasferimento_materiale.Inizializzazione_rt()

        Richiesta_trasferimento_materiale.riempi_datagridview_rt()
        Richiesta_trasferimento_materiale.Owner = Me

        Me.Hide()
        Richiesta_trasferimento_materiale.Show()


    End Sub

    Sub riempi_ODP()
        DataGridView_ODP.Rows.Clear()

        Form_principale.Cnn1.Close()
        Form_principale.Cnn1.ConnectionString = Form_principale.SAP
        Form_principale.Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Form_principale.Cnn1
        If Dashboard_MU_New.mu = 0 Then


            CMD_SAP_2.CommandText = "declare @odp as integer

set @odp=" & FORM6.ODP & "

select t1.itemcode, t1.ItemName,  t2.u_disegno, t1.PlannedQty, t1.U_PRG_wip_Qtaspedita, t1.U_PRG_WIP_QtaDaTrasf, t1.wareHouse, case when t1.U_PRG_WIP_QtaDaTrasf=0 then 'OK' when a.giacenza>=t1.U_PRG_WIP_QtaDaTrasf then 'Trasferibile' when a.giacenza+a.CAP2>=t1.U_PRG_WIP_QtaDaTrasf then 'CAP2' when a.giacenza+a.CAP2+a.[CQ-Clavter]>=t1.U_PRG_WIP_QtaDaTrasf then 'CQ-Clavter' when a.giacenza+a.CAP2+a.[CQ-Clavter]+a.ordinato>=t1.U_PRG_WIP_QtaDaTrasf then 'IN APPROV' else'Da ordinare' end as 'Azione',b.odp,b.U_PRG_AZS_Commessa,b.U_PRODUZIONE,b.Cons_odp,b.oa,b.cardname,b.Shipdate, t1.linenum, t1.VisOrder, t1.U_Stato_lavorazione
from owor t0 inner join wor1 t1 on t0.DocEntry=t1.docentry
left join oitm t2 on t2.itemcode=t1.itemcode

inner join

(
select t30.docnum, t30.itemcode, t30.linenum, t30.Giacenza,t30.CAP2, t30.[CQ-Clavter], sum(case when t31.onorder is null then 0 else t31.onorder end) as 'Ordinato' , sum(case when t31.onhand is null then 0 else t31.onhand end) + sum(case when t31.onorder is null then 0 else t31.onorder end) -sum(case when t31.iscommited is null then 0 else t31.iscommited end)  as 'disponibile'
from
(
select t20.docnum, t20.itemcode, t20.linenum, t20.Giacenza,t20.CAP2, sum(case when t21.onhand is null then 0 else t21.onhand end) as 'CQ-Clavter'
from
(
select t10.docnum, t10.itemcode, t10.linenum, t10.Giacenza, case when t11.onhand is null then 0 else t11.onhand end as 'CAP2'
from
(
select t0.docnum, t1.itemcode, t1.linenum, sum(t2.onhand) as 'Giacenza'
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
inner join oitw t2 on t2.itemcode=t1.itemcode
where t0.docnum=@odp and t2.whscode<>'WIP' and t2.whscode<>'CQ' and t2.whscode<>'CAP2' and t2.whscode<>'clavter'

group by t0.docnum, t1.itemcode, t1.linenum

)
as t10 left join oitw t11 on t11.itemcode=t10.itemcode and t11.whscode='CAP2'
)
as t20 left join oitw t21 on t21.itemcode=t20.itemcode and (t21.whscode='CQ' or t21.whscode='Clavter')
group by t20.docnum, t20.itemcode, t20.linenum, t20.Giacenza,t20.CAP2
)
as t30 left join oitw t31 on t31.itemcode=t30.itemcode
group by t30.docnum, t30.itemcode, t30.linenum, t30.Giacenza,t30.CAP2, t30.[CQ-Clavter]

) A on t0.docnum=a.docnum and t1.linenum=a.linenum

inner join 

(
select t40.docnum, t40.itemcode, t40.linenum, t40.ODP, t41.U_PRG_AZS_Commessa,t41.U_PRODUZIONE, case when substring(t41.u_produzione,1,3)='INT' then t41.U_Data_cons_MES else t41.DueDate end as 'Cons_odp', t42.docnum as 'OA',t42.cardname, t40.shipdate
from
(
select t30.docnum, t30.itemcode, t30.linenum, t30.ODP, t30.Shipdate,min(t31.docentry) as 'Docentry'
from
(
select t20.docnum, t20.itemcode, t20.linenum, t20.ODP, min(t21.ShipDate) as 'Shipdate'
from
(
select t10.docnum, t10.itemcode, t10.linenum, min(t10.ODP) as 'ODP'
from
(
select t0.docnum, t1.itemcode, t1.linenum, t2.docnum as 'ODP', min(case when substring(t2.u_produzione,1,3)='INT' then t2.U_Data_cons_MES else t2.duedate end)as 'Consegna'
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry
left join owor t2 on t2.itemcode=t1.itemcode and (t2.status='P' or t2.status='R') 

where t0.docnum=@odp
group by t0.docnum, t1.itemcode, t1.linenum, t2.docnum
)
as t10 left join owor t11 on t11.docnum=t10.odp and t11.duedate=t10.consegna
group by t10.docnum, t10.itemcode, t10.linenum
)
as t20 left join por1 t21 on t21.OpenQty>0 and t20.ItemCode=t21.itemcode
group by t20.docnum, t20.itemcode, t20.linenum, t20.ODP
)
as t30 left join por1 t31 on t31.shipdate=t30.shipdate and t30.ItemCode=t31.itemcode and t31.OpenQty>0
group by t30.docnum, t30.itemcode, t30.linenum, t30.ODP, t30.Shipdate
)
as t40 left join owor t41 on t41.docnum=t40.odp
left join opor t42 on t42.docentry=t40.docentry

) B on t0.docnum=B.docnum and t1.linenum=B.linenum



where t0.docnum=@odp and t1.itemtype=4 and (substring(T1.[ITEMCODE],1,1)='0' or substring(T1.[ITEMCODE],1,1)='C' or substring(T1.[ITEMCODE],1,1)='D')"


        Else

            CMD_SAP_2.CommandText = "SELECT t40.linenum, t40.visorder, t40.articolo, t40.[Desc articolo], t40.Disegno, t40.Quantita , t40.Trasferito, t40.[Da trasferire]  ,t40.warehouse, t40.azione  , min(t40.ODP) as 'ODP', t40.Commessa, t40.Reparto, t40.[Cons ODP], t40.OA,t40.Fornitore, t40.[Cons OA] ,t40.u_stato_lavorazione
FROM
(
Select t30.linenum, t30.visorder, t30.articolo as 'Articolo', t30.[Desc articolo] as 'Desc articolo', t30.Disegno as 'Disegno', t30.Quantita as 'Quantita', t30.Trasferito as 'Trasferito', t30.[Da trasferire] as 'Da trasferire' ,t30.warehouse, t30.azione as 'Azione' , t31.docnum as 'ODP', t31.[U_PRG_AZS_Commessa] as 'Commessa', t31.U_produzione as 'Reparto', CASE WHEN SUBSTRING(t31.U_produzione,1,3)='INT' THEN T31.[U_Data_cons_MES] ELSE t31.duedate END as 'Cons ODP', t30.OA as'OA',t30.Fornitore as 'Fornitore', t30.[Cons OA] as 'Cons OA',t30.u_stato_lavorazione
from
(
Select t20.linenum, t20.visorder, t20.articolo, t20.[Desc articolo], t20.Disegno, t20.Quantita, t20.Trasferito, t20.[Da trasferire],t20.warehouse, t20.azione, min(case when t20.[Cons ODP] is null then '' else cast(t20.[Cons ODP] as date) end) as 'Cons ODP', t22.docnum as 'OA' , t22.cardname as 'Fornitore', t21.shipdate as 'Cons OA',t20.u_stato_lavorazione
from
(
Select t10.linenum, t10.visorder, t10.articolo, t10.[Desc articolo], t10.Disegno, t10.Quantita, t10.Trasferito, t10.[Da trasferire],t10.warehouse, t10.azione, t10.ODP, t10.[Cons ODP], t10.Commessa, t10.Reparto, min(t10.[Cons OA]) as 'Cons OA',t10.u_stato_lavorazione
from
(
Select t100.linenum, t100.visorder, T100.Articolo, t100.[Desc articolo] , t100.Disegno, t100.Quantita,t100.Trasferito, t100.[Da trasferire], t100.warehouse,
case when t100.[Da trasferire]=0 then 'OK' when (t100.giacenza>=t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 )  then 'Trasferibile/Da ordinare' when (t100.giacenza>=t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)>=0 )  then 'Trasferibile' when t100.giacenza<t100.[Da trasferire] and t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 and sum(t106.onorder)>=t100.[Da trasferire] then 'IN APPROV/DA ORDINARE' when t100.[Da trasferire]=0 then 'OK' when (t100.[Da trasferire]>0 and sum(t106.onhand-T106.[IsCommited]+t106.onorder)>=0 and t100.giacenza<t100.[Da trasferire]) then 'IN APPROV'   when sum(t106.onhand-T106.[IsCommited]+t106.onorder)<0 then 'Da ordinare' end as 'Azione', case when t100.[Da trasferire]=0 then '' else t102.docnum end as 'ODP', cast(case when t100.[Da trasferire]=0 then '' else cast(T102.[DueDate] as varchar) end as VARCHAR)as 'Cons ODP' , case when t100.[Da trasferire]=0 then '' else t102.U_PRG_AZS_commessa end as 'Commessa' ,case when t100.[Da trasferire]=0 then '' else  t102.U_produzione end as 'Reparto', case when t100.[Da trasferire]=0 then '' else t107.docnum  end  as 'OA', 
case when t100.[Da trasferire]=0 then '' else t107.cardname end as 'Fornitore', cast(case when t100.[Da trasferire]=0 then '' else cast(t103.[ShipDate] as varchar)  end as varchar) as 'Cons OA',t100.u_stato_lavorazione

from
(
SELECT  t1.linenum, t1.visorder,  T9.[ITEMCODE] as 'Articolo', case when t1.itemtype = -18 then CAST (t1.linetext as varchar) else t9.itemname end as 'Desc articolo' , t9.u_disegno as 'Disegno', t1.plannedqty as 'Quantita',case when t1.U_prg_wip_qtaspedita is null then 0 else t1.U_prg_wip_qtaspedita end as 'Trasferito', t1.u_prg_wip_qtadatrasf as 'Da trasferire', t1.warehouse, sum (t20.onhand) as 'giacenza', t1.docentry, t1.u_stato_lavorazione

from wor1 t1 inner join owor t0 on t0.docentry=t1.docentry
left join oitm t9 on t9.itemcode=t1.itemcode
left join oitw t20 on t20.itemcode=t1.itemcode

WHERE t0.docnum=" & FORM6.ODP & " and  (t20.whscode='01' or t20.whscode='03' or t20.whscode='SCA' or t20.whscode='FERRETTO' or t20.whscode='MUT' OR t1.itemtype=290 OR t1.itemtype=-18)

group by 
t1.linenum, t1.itemtype, CAST (t1.linetext as varchar), t1.visorder, T9.[ITEMCODE] , t9.itemname  , t9.u_disegno , t1.plannedqty, t1.U_prg_wip_qtaspedita , t1.u_prg_wip_qtadatrasf, t1.docentry,t1.u_stato_lavorazione,t1.warehouse
)
as t100 left join wor1 t101 on t101.itemcode=t100.articolo and t101.docentry=t100.docentry and t100.linenum=t101.linenum
left join owor t102 on t101.itemcode=t102.itemcode and (T102.Status ='P' or T102.Status ='R' )
left join por1 t103 on t103.itemcode=t101.itemcode and t103.opencreqty >0
LEFT OUTER JOIN ITT1 T104 on T101.itemCode = T104.Father
left join oitw t105 on t105.itemcode=t104.code and t105.[WhsCode]='01'
left join oitw t106 on t106.itemcode=t101.itemcode
left join opor t107 on t107.docentry=t103.docentry

group by
 T100.[articolo], t100.trasferito, T100.[DESC articolo], t100.linenum, t100.visorder, t100.quantita,  t100.disegno, t100.giacenza,t100.[da trasferire], t102.docnum, T102.[DueDate],t102.U_PRG_AZS_commessa,t102.U_produzione,t107.docnum,t107.cardname,t103.[ShipDate],t100.u_stato_lavorazione, t100.warehouse
)
as t10
group by t10.linenum, t10.visorder, t10.articolo, t10.[Desc articolo], t10.[Desc articolo], t10.Disegno, t10.Quantita, t10.Trasferito, t10.[Da trasferire],t10.azione, t10.ODP, t10.[Cons ODP], t10.Commessa, t10.Reparto,t10.u_stato_lavorazione,t10.warehouse
) 
as t20
left join por1 t21 on t21.itemcode=t20.articolo and t21.shipdate=t20.[Cons OA] and t21.opencreqty >0
left join opor t22 on t22.docentry=t21.docentry
group by
t20.linenum, t20.visorder, t20.articolo, t20.[Desc articolo], t20.Disegno, t20.Quantita, t20.Trasferito, t20.[Da trasferire],t20.azione,   t22.docnum ,t22.cardname, t21.shipdate ,t20.u_stato_lavorazione,t20.warehouse
)
as t30
left join owor t31 on t31.itemcode=t30.articolo and (T31.Status <> N'L' )  AND  (T31.Status <> N'C' ) and T31.[DueDate]=t30.[Cons ODP]
)
AS T40
group by
t40.linenum, t40.visorder, t40.articolo, t40.[Desc articolo], t40.Disegno, t40.Quantita , t40.Trasferito, t40.[Da trasferire]  ,t40.warehouse, t40.azione , t40.Commessa, t40.Reparto, t40.[Cons ODP], t40.OA,t40.Fornitore, t40.[Cons OA] ,t40.u_stato_lavorazione
order by t40.linenum"


        End If
        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        If Dashboard_MU_New.mu = 0 Then


            Do While cmd_SAP_reader_2.Read()

                DataGridView_ODP.Rows.Add(cmd_SAP_reader_2("itemcode"), cmd_SAP_reader_2("itemname"), cmd_SAP_reader_2("u_Disegno"), cmd_SAP_reader_2("plannedqty"), cmd_SAP_reader_2("U_PRG_wip_Qtaspedita"), cmd_SAP_reader_2("U_PRG_WIP_QtaDaTrasf"), cmd_SAP_reader_2("wareHouse"), cmd_SAP_reader_2("Azione"), cmd_SAP_reader_2("ODP"), cmd_SAP_reader_2("U_PRG_AZS_Commessa"), cmd_SAP_reader_2("u_produzione"), cmd_SAP_reader_2("Cons_ODP"), cmd_SAP_reader_2("OA"), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("shipdate"), cmd_SAP_reader_2("linenum"), cmd_SAP_reader_2("visorder"))

            Loop

        Else
            Do While cmd_SAP_reader_2.Read()

                DataGridView_ODP.Rows.Add(cmd_SAP_reader_2("Articolo"), cmd_SAP_reader_2("Desc articolo"), cmd_SAP_reader_2("Disegno"), cmd_SAP_reader_2("Quantita"), cmd_SAP_reader_2("Trasferito"), cmd_SAP_reader_2("Da trasferire"), cmd_SAP_reader_2("warehouse"), cmd_SAP_reader_2("Azione"), cmd_SAP_reader_2("ODP"), cmd_SAP_reader_2("Commessa"), cmd_SAP_reader_2("Reparto"), cmd_SAP_reader_2("Cons ODP"), cmd_SAP_reader_2("OA"), cmd_SAP_reader_2("Fornitore"), cmd_SAP_reader_2("Cons OA"), cmd_SAP_reader_2("linenum"), cmd_SAP_reader_2("visorder"), cmd_SAP_reader_2("u_stato_lavorazione"))

            Loop
        End If

        cmd_SAP_reader_2.Close()
        Form_principale.Cnn1.Close()

        DataGridView_ODP.ClearSelection()

    End Sub

    Sub intestazioni_ordine_di_produzione()



        Form_principale.Cnn1.ConnectionString = Form_principale.SAP
        Form_principale.Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Form_principale.Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.docentry, T0.[DocNum] AS 'docnum', T0.PLANNEDQTY, case when t0.u_stato is null then '' else t0.u_stato end as 'u_stato', T0.STATUS AS 'Stato', t0.u_lavorazione as 'lavorazione', T0.[ItemCode] as 'Itemcode', T1.[ItemName] as 'Itemname', case when T1.[U_Disegno] is null then '' else t1.u_disegno end as 'Disegno', case when T0.[U_PRG_AZS_Commessa] is null then '' else T0.[U_PRG_AZS_Commessa] end as 'Commessa', case when T3.[resname] is null then '' else t3.resname end as 'Fase' , T2.[ItmsGrpNam] as 'Gruppo articolo', case when T0.U_AGGIORNA_DB is null then '' else t0.u_aggiorna_db end as 'u_aggiorna_db', T0.POSTDATE, T0.DUEDATE

FROM OWOR T0 INNER JOIN OITM T1 ON T0.ITEMCODE=T1.ITEMCODE
INNER JOIN OITB T2 ON T1.[ItmsGrpCod] = T2.[ItmsGrpCod] 
left join orsc t3 on t3.visrescode =t0.u_fase
WHERE T0.[DocNum]='" & FORM6.ODP & "'"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() = True Then
            Label_numero_ODP_F.Text = cmd_SAP_reader_2("docnum")
            ComboBox_stato_odp.Text = cmd_SAP_reader_2("stato")
            Label_Codice_ODP_F.Text = cmd_SAP_reader_2("Itemcode")
            Label_descrizione_F.Text = cmd_SAP_reader_2("Itemname")
            Label_commessa_F.Text = cmd_SAP_reader_2("Commessa")
            Button1.Text = cmd_SAP_reader_2("Disegno")
            Label_fase_F.Text = cmd_SAP_reader_2("Fase")
            Label_gruppo_articolo_F.Text = cmd_SAP_reader_2("Gruppo articolo")
            TextBox3.Text = cmd_SAP_reader_2("Lavorazione")
            ComboBox2.Text = cmd_SAP_reader_2("u_stato")
            Label9.Text = Math.Round(cmd_SAP_reader_2("PLANNEDQTY"))
            docentry_odp = cmd_SAP_reader_2("docentry")
            ComboBox3.Text = cmd_SAP_reader_2("u_aggiorna_db")
            ComboBox3.Text = cmd_SAP_reader_2("u_aggiorna_db")
            Label10.Text = cmd_SAP_reader_2("POSTDATE")
            Label11.Text = cmd_SAP_reader_2("DUEDATE")

            cmd_SAP_reader_2.Close()
        End If
        cmd_SAP_reader_2.Close()
        Form_principale.Cnn1.Close()



        If File.Exists("\\192.168.0.150\k\Tecnico\Disegni Meccanici\PDF-DXF\PDF\" & Button1.Text & ".PDF") Then

            AxFoxitCtl1.OpenFile("\\192.168.0.150\k\Tecnico\Disegni Meccanici\PDF-DXF\PDF\" & Button1.Text & ".PDF")
            Label2.Hide()
            AxFoxitCtl1.Show()
        Else
            AxFoxitCtl1.Hide()
            Label2.Show()
        End If

    End Sub

    Private Sub DataGridView_ODP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellContentClick

    End Sub
End Class