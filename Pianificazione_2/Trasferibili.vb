Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop


Public Class Trasferibili
    Public riga As Integer
    Public trasferibili = 0
    Private isShiftKeyDown As Boolean = False
    Private startIndex As Integer = -1
    Private ID_lotto_di_prelievo As Integer

    Private Sub DataGridView_ODP_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView_ODP.CellMouseDown
        If e.Button = MouseButtons.Left AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If ModifierKeys = Keys.Shift AndAlso startIndex >= 0 Then
                'Se è premuto Shift, cambia il flag per le righe comprese tra startIndex ed e.RowIndex
                Dim endIndex As Integer = e.RowIndex
                Dim minIndex As Integer = Math.Min(startIndex, endIndex) + 1
                Dim maxIndex As Integer = Math.Max(startIndex, endIndex) - 1

                For i As Integer = minIndex To maxIndex
                    DataGridView_ODP.Rows(i).SetValues(True)
                Next i
            Else
                '  Altrimenti, imposta startIndex alla riga corrente
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

    Private Sub DataGridView_ODP_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView_ODP.CellFormatting
        If DataGridView_ODP.Rows(e.RowIndex).Cells(12).Value > 0 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(12).Style.BackColor = Color.Green
        End If

        If DataGridView_ODP.Rows(e.RowIndex).Cells(8).Value = 100 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(8).Style.BackColor = Color.Green
        ElseIf DataGridView_ODP.Rows(e.RowIndex).Cells(8).Value < 100 And DataGridView_ODP.Rows(e.RowIndex).Cells(8).Value > 90 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(8).Style.BackColor = Color.Yellow
        Else
            DataGridView_ODP.Rows(e.RowIndex).Cells(8).Style.BackColor = Color.Red
        End If
        If DataGridView_ODP.Rows(e.RowIndex).Cells(9).Value > 0 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(9).Style.BackColor = Color.Green
        End If

        If DataGridView_ODP.Rows(e.RowIndex).Cells(10).Value > 0 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(10).Style.BackColor = Color.Green
        End If

        If DataGridView_ODP.Rows(e.RowIndex).Cells(11).Value > 0 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(11).Style.BackColor = Color.Green
        End If

        If DataGridView_ODP.Rows(e.RowIndex).Cells(12).Value > 0 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(12).Style.BackColor = Color.Green
        End If

        If DataGridView_ODP.Rows(e.RowIndex).Cells(13).Value > 0 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(13).Style.BackColor = Color.Green
        End If
        If DataGridView_ODP.Rows(e.RowIndex).Cells(14).Value > 0 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(14).Style.BackColor = Color.Green
        End If
        If DataGridView_ODP.Rows(e.RowIndex).Cells(15).Value > 0 Then
            DataGridView_ODP.Rows(e.RowIndex).Cells(15).Style.BackColor = Color.Green
        End If


    End Sub

    Private Sub DataGridView_ODP_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellClick

        If e.RowIndex >= 0 Then
            riga = e.RowIndex
            If e.ColumnIndex > 0 Then
                If File.Exists(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_ODP.Rows(e.RowIndex).Cells(4).Value & ".PDF") Then

                    AxFoxitCtl1.OpenFile(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_ODP.Rows(e.RowIndex).Cells(4).Value & ".PDF")

                    AxFoxitCtl1.Show()

                Else
                    AxFoxitCtl1.Hide()

                End If

                If File.Exists(Homepage.percorso_DWF & DataGridView_ODP.Rows(riga).Cells(4).Value & ".iam.dwf") Then
                    Button7.BackColor = Color.Lime
                Else
                    Button7.BackColor = Color.Red
                End If

                If File.Exists(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_ODP.Rows(riga).Cells(4).Value & ".PDF") Then
                    Button8.BackColor = Color.Lime
                Else
                    Button8.BackColor = Color.Red
                End If


                If e.ColumnIndex = 1 Then



                    ODP_Form.docnum_odp = DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="N_ODP").Value
                    ODP_Form.Show()
                    ODP_Form.inizializza_form(DataGridView_ODP.Rows(e.RowIndex).Cells(columnName:="N_ODP").Value)



                End If

                If e.ColumnIndex = 4 Then

                    Try
                        Process.Start(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_ODP.Rows(e.RowIndex).Cells(4).Value & ".PDF")
                    Catch ex As Exception
                        MsgBox("Il disegno " & DataGridView_ODP.Rows(e.RowIndex).Cells(4).Value & " non è ancora stato processato")
                    End Try


                End If
            End If
        End If

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If File.Exists(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_ODP.Rows(riga).Cells(3).Value & ".PDF") Then
            Process.Start(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_ODP.Rows(riga).Cells(3).Value & ".PDF")
        Else
            MsgBox("PDF non presente")
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If File.Exists(Homepage.percorso_DWF & DataGridView_ODP.Rows(riga).Cells(3).Value & ".iam.dwf") Then
            Process.Start(Homepage.percorso_DWF & DataGridView_ODP.Rows(riga).Cells(3).Value & ".iam.dwf")
        Else
            MsgBox("3D non presente")
        End If
    End Sub




    Sub filtra()
        Dim i = 0
        Do While i < DataGridView_ODP.RowCount
            Dim parola0 As String
            Dim parola1 As String
            Dim parola2 As String
            Dim parola5 As String
            Dim parola6 As String

            parola0 = UCase(DataGridView_ODP.Rows(i).Cells(0).Value)
            parola1 = UCase(DataGridView_ODP.Rows(i).Cells(1).Value)
            parola2 = UCase(DataGridView_ODP.Rows(i).Cells(2).Value)
            parola5 = UCase(DataGridView_ODP.Rows(i).Cells(5).Value)
            parola6 = UCase(DataGridView_ODP.Rows(i).Cells(6).Value)

            If parola0.Contains(UCase(TextBox1.Text)) Then
                DataGridView_ODP.Rows(i).Visible = True

                If parola1.Contains(UCase(TextBox3.Text)) Then
                    DataGridView_ODP.Rows(i).Visible = True

                    If parola2.Contains(UCase(TextBox2.Text)) Then
                        DataGridView_ODP.Rows(i).Visible = True

                        If parola5.Contains(UCase(TextBox5.Text)) Then
                            DataGridView_ODP.Rows(i).Visible = True


                            If parola6.Contains(UCase(TextBox4.Text)) Then
                                DataGridView_ODP.Rows(i).Visible = True

                            Else
                                DataGridView_ODP.Rows(i).Visible = False


                            End If

                        Else
                            DataGridView_ODP.Rows(i).Visible = False


                        End If

                    Else
                        DataGridView_ODP.Rows(i).Visible = False

                    End If
                Else
                    DataGridView_ODP.Rows(i).Visible = False


                End If
            Else
                DataGridView_ODP.Rows(i).Visible = False

            End If
            i = i + 1
        Loop
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Me.Hide()
        Commesse_magazzino.Owner = Me
        Commesse_magazzino.Show()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Analisi_riga_magazzino.TextBox7.Text = pianificazione.commessa
        Analisi_riga_magazzino.Materiale_mancante()
        Analisi_riga_magazzino.Owner = Me
        Analisi_riga_magazzino.Show()
        Me.Hide()
        Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa)
        Analisi_riga_magazzino.Button_commessa.Text = pianificazione.commessa
        Analisi_riga_magazzino.Label_descrizione.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Descrizione_commessa
        Analisi_riga_magazzino.Label_cliente.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_commessa
        Analisi_riga_magazzino.Label_cliente_finale.Text = Commesse_MES.SCHEDA_COMMESSA(Pianificazione.commessa).Cliente_finale_commessa
    End Sub

    Private Sub Button_commessa_Click(sender As Object, e As EventArgs) Handles Button_commessa.Click
        Homepage.mostra_dashboard()
        Mostra.Show()
        Mostra.Owner = Me
        Me.Hide()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click

        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        filtra()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs)
        Commesse_magazzino.elenco_ODP_commessa()

    End Sub

    Private Sub DataGridView_ODP_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView_ODP.CellContentClick

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'Dim contatore As Integer = 1

        'ODP_Form.stampa_etichetta = "NO"
        'Do While contatore < DataGridView_ODP.Rows.Count
        '    If DataGridView_ODP.Rows(contatore).Cells(0).Value = True Then


        '        If CheckBox3.Checked = True Then
        '            'testata_odp()
        '            AxFoxitCtl1.OpenFile(Homepage.percorso_disegni_generico & "PDF\"  & DataGridView_ODP.Rows(contatore).Cells(4).Value & ".PDF")
        '            AxFoxitCtl1.PrintFile()

        '        End If

        '        If CheckBox1.Checked = True And CheckBox2.Checked = False Then
        '            FORM6.ODP = DataGridView_ODP.Rows(contatore).Cells(1).Value
        '            ODP_Form.percorso_documento = Homepage.PERCORSO_DOCUMENTO_ODP
        '            ODP_Form.Genera_ordine()
        '        ElseIf CheckBox2.Checked = True Then
        '            ODP_Form.stampa_etichetta = "YES"
        '            FORM6.ODP = DataGridView_ODP.Rows(contatore).Cells(1).Value
        '            ODP_Form.percorso_documento = Homepage.PERCORSO_DOCUMENTO_ODP_ETICHETTA
        '            ODP_Form.Genera_ordine()
        '        End If
        '        DataGridView_ODP.Rows(contatore).Cells(0).Value = False
        '    End If

        '    contatore = contatore + 1
        'Loop
        'contatore = 0




        ODP_Form.stampa_etichetta = "NO"


        For Each Row As DataGridViewRow In DataGridView_ODP.Rows

            If Row.Cells("Stampa").Value = True Then



                If CheckBox1.Checked = True And CheckBox2.Checked = False Then
                    FORM6.ODP = Row.Cells("N_ODP").Value
                    ODP_Form.docnum_odp = Row.Cells("N_ODP").Value
                    ODP_Form.percorso_documento = Homepage.PERCORSO_DOCUMENTO_ODP
                    ODP_Form.Genera_ordine()
                ElseIf CheckBox2.Checked = True Then
                    ODP_Form.stampa_etichetta = "YES"
                    FORM6.ODP = Row.Cells("N_ODP").Value
                    ODP_Form.docnum_odp = Row.Cells("N_ODP").Value
                    ODP_Form.percorso_documento = Homepage.PERCORSO_DOCUMENTO_ODP_ETICHETTA
                    ODP_Form.Genera_ordine()
                End If
                If CheckBox3.Checked And File.Exists(Homepage.percorso_disegni_generico & "PDF\"  & Row.Cells("disegno").Value & ".PDF") = True Then


                    AxFoxitCtl1.OpenFile(Homepage.percorso_disegni_generico & "PDF\"  & Row.Cells("disegno").Value & ".PDF")
                    AxFoxitCtl1.PrintFile()

                End If

            End If


        Next
        MsgBox("FINE STAMPE")


    End Sub








End Class