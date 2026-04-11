Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Imports System.Windows.Controls
Imports Tirelli.Form_gestione_campioni
Imports Outlook = Microsoft.Office.Interop.Outlook

Public Class Movimento_campioni

    Private isShiftKeyDown As Boolean = False
    Private startIndex As Integer = -1
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged


        Dim par_combobox As System.Windows.Forms.ComboBox
        par_combobox = CType(ComboBox5, System.Windows.Forms.ComboBox)
        If par_combobox.Text = "Entrata merce" Then
            For Each row As DataGridViewRow In DataGridView1.Rows
                ' Assicurati che il nome della colonna corrisponda esattamente al nome nella DataGridView
                row.Cells("DOC").Value = "EM"
                row.Cells("Dal_mag").Value = "-"
                row.Cells("al_mag").Value = "01"
            Next
            ComboBox1.Enabled = False
            ComboBox4.Enabled = False
        ElseIf par_combobox.Text = "Trasferimento di magazzino" Then
            For Each row As DataGridViewRow In DataGridView1.Rows
                ' Assicurati che il nome della colonna corrisponda esattamente al nome nella DataGridView
                row.Cells("DOC").Value = "TR"
                row.Cells("Dal_mag").Value = "-"
                row.Cells("al_mag").Value = "-"
                ComboBox1.Enabled = True
                ComboBox4.Enabled = True

            Next
        End If

    End Sub

    Public Function trova_giacenza_campione(PAR_campione As String, par_magazzino As String)

        Dim giacenza As Integer = 0
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT TOP (1000) [id_campione]
      ,[mag]
      ,coalesce([Q_IN],0) as 'Q_IN'
  FROM [TIRELLI_40].[DBO].[coll_campioni_giacenze]
where [id_campione]='" & PAR_campione & "' and [mag] = '" & par_magazzino & "' "

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then

            giacenza = cmd_SAP_reader("Q_IN")
        Else

            giacenza = 0

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()
        Return giacenza
    End Function 'Inserisco le risorse nella combo box

    Sub Inserimento_da_mag(par_combobox As System.Windows.Forms.ComboBox)
        par_combobox.Items.Clear()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.WhsCode, T0.WhsName
 FROM OWHS T0
where t0.inactive='N' "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 1
        Do While cmd_SAP_reader.Read()

            par_combobox.Items.Add(cmd_SAP_reader("WhsCode"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Inserisco le risorse nella combo box
    Sub Inserimento_a_mag(par_combobox As System.Windows.Forms.ComboBox)
        par_combobox.Items.Clear()

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.WhsCode, T0.WhsName
 FROM OWHS T0
where t0.inactive='N' "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 1
        Do While cmd_SAP_reader.Read()

            par_combobox.Items.Add(cmd_SAP_reader("WhsCode"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()
    End Sub 'Inserisco le risorse nella combo box

    Private Sub Movimento_campioni_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Inserimento_da_mag(ComboBox1)
        Inserimento_da_mag(ComboBox4)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Assicurati che il nome della colonna corrisponda esattamente al nome nella DataGridView

            row.Cells("Dal_mag").Value = ComboBox1.Text
            row.Cells("Giacenza").Value = trova_giacenza_campione(row.Cells("ID_campione").Value, row.Cells("Dal_mag").Value)


        Next
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        For Each row As DataGridViewRow In DataGridView1.Rows
            ' Assicurati che il nome della colonna corrisponda esattamente al nome nella DataGridView

            row.Cells("Al_mag").Value = ComboBox4.Text

        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox5.SelectedIndex < 0 Then
            MsgBox("Selezionare un tipo di trasferimento")
            Return
        End If

        ' Verifica se Outlook è già in esecuzione
        Dim outlookApp As Outlook.Application = Nothing
        Try
            outlookApp = Marshal.GetActiveObject("Outlook.Application")
        Catch ex As COMException
            ' Se non è in esecuzione, lo avvia
            outlookApp = New Outlook.Application()
        End Try

        Dim par_datagridview As DataGridView = DataGridView1
        Dim riepilogoHTML As String = "<html><body><h3>Campioni movimentati:</h3><table border='1' cellpadding='5'>"
        Dim immaginiInEmail As New List(Of Tuple(Of String, String)) ' Tuple: (immaginePath, contentId)
        Dim destinatariUnici As New HashSet(Of String)
        Dim primoCardName As String = ""

        For i As Integer = par_datagridview.Rows.Count - 1 To 0 Step -1
            Dim row As DataGridViewRow = par_datagridview.Rows(i)
            If row.Cells("SEl").Value = True Then
                If row.Cells("q_trasf").Value > 0 Then
                    Dim idCampione = row.Cells("ID_campione").Value

                    ' Cast dell'oggetto restituito a Dettagli_Campione
                    Dim info As Dettagli_Campione = CType(Form_gestione_campioni.ottieni_informazioni_campione(idCampione), Dettagli_Campione)
                    ' Prendi solo il primo cardname valido
                    If String.IsNullOrEmpty(primoCardName) Then
                        primoCardName = info.cardname
                    End If
                    ' Aggiunge l'email se presente
                    If Not String.IsNullOrWhiteSpace(info.Email) Then
                        destinatariUnici.Add(info.Email.ToLower.Trim)
                    End If

                    Dim descrizione As String = info.iniziale_sigla & info.nome
                    Dim immaginePath = Homepage.Percorso_immagini & info.immagine
                    Dim contentId As String = "img" & idCampione.ToString()

                    If row.Cells("DOC").Value = "TR" Then
                        If row.Cells("q_trasf").Value < row.Cells("giacenza").Value Then
                            If row.Cells("dal_mag").Value = "-" Or row.Cells("al_mag").Value = "-" Then
                                MsgBox("Selezionare un magazzino per il campione " & idCampione)
                            Else
                                riepilogoHTML &= "<tr><td><b>" & descrizione & "</b><br>ID: " & idCampione &
                                         "<br>TR da " & row.Cells("dal_mag").Value & " a " & row.Cells("al_mag").Value &
                                         "<br>Q.tà: " & row.Cells("q_trasf").Value & "</td>"
                                If IO.File.Exists(immaginePath) Then
                                    immaginiInEmail.Add(Tuple.Create(immaginePath, contentId))
                                    riepilogoHTML &= "<td><img src='cid:" & contentId & "' width='100'></td></tr>"
                                Else
                                    riepilogoHTML &= "<td>(Immagine non trovata)</td></tr>"
                                End If

                                ' Movimentazioni (commentate)
                                movimenta_campione(idCampione, row.Cells("doc").Value, "-", row.Cells("dal_mag").Value, row.Cells("q_trasf").Value, Homepage.ID_SALVATO, 0)
                                movimenta_campione(idCampione, row.Cells("doc").Value, "+", row.Cells("al_mag").Value, row.Cells("q_trasf").Value, Homepage.ID_SALVATO, 0)
                                aggiusta_giacenza(idCampione, row.Cells("dal_mag").Value)
                                aggiusta_giacenza(idCampione, row.Cells("al_mag").Value)

                                par_datagridview.Rows.RemoveAt(i)
                                Form_gestione_campioni.riempi_movimentazioni_campioni("", "", Form_campione_visualizza.DataGridView4, "", "", "", idCampione)
                            End If
                        Else
                            MsgBox("Il campione " & idCampione & " ha giacenza insufficiente per essere trasferito")
                        End If
                    Else
                        riepilogoHTML &= "<tr><td><b>" & descrizione & "</b><br>ID: " & idCampione &
                                 "<br>" & row.Cells("doc").Value & " in " & row.Cells("al_mag").Value &
                                 "<br>Q.tà: " & row.Cells("q_trasf").Value & "</td>"
                        If IO.File.Exists(immaginePath) Then
                            immaginiInEmail.Add(Tuple.Create(immaginePath, contentId))
                            riepilogoHTML &= "<td><img src='cid:" & contentId & "' width='100'></td></tr>"
                        Else
                            riepilogoHTML &= "<td>(Immagine non trovata)</td></tr>"
                        End If

                        ' Movimentazioni (commentate)
                        movimenta_campione(idCampione, row.Cells("doc").Value, "+", row.Cells("al_mag").Value, row.Cells("q_trasf").Value, Homepage.ID_SALVATO, 0)
                        aggiusta_giacenza(idCampione, row.Cells("al_mag").Value)

                        par_datagridview.Rows.RemoveAt(i)
                    End If
                Else
                    MsgBox("Selezionare una quantità > 0 per il campione " & row.Cells("ID_campione").Value & " per il trasferimento")
                End If
            End If
        Next

        riepilogoHTML &= "</table></body></html>"

        MsgBox("Campioni selezionati trasferiti con successo")
        par_datagridview.ClearSelection()

        ' Invio email riepilogo con immagini inline
        If immaginiInEmail.Count > 0 Then
            Try
                '  Dim outlookApp As New Outlook.Application
                Dim mail As Outlook.MailItem = outlookApp.CreateItem(Outlook.OlItemType.olMailItem)

                If destinatariUnici.Count = 0 Then
                    mail.To = "projectmanagers@tirelli.net"
                    mail.CC = "giovanni.tirelli@tirelli.net;mattia.rossi@tirelli.net"
                Else
                    mail.To = String.Join(";", destinatariUnici)
                    mail.CC = "projectmanagers@tirelli.net;giovanni.tirelli@tirelli.net;mattia.rossi@tirelli.net"
                End If

                mail.Subject = "Trasferimento campioni " & primoCardName
                mail.HTMLBody = riepilogoHTML

                Dim attachments = mail.Attachments
                For Each img In immaginiInEmail
                    Dim attachment = attachments.Add(img.Item1, Outlook.OlAttachmentType.olByValue, Nothing, "Image")
                    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", img.Item2)
                Next

                mail.Send()
            Catch ex As Exception
                MsgBox("Errore durante l'invio dell'email: " & ex.Message)
            End Try
        End If
    End Sub

    Sub movimenta_campione(par_id_campione As String, par_tipo_movimentazione As String, par_segno As String, par_mag As String, par_q As Integer, par_id_dipendente As Integer, par_id_richiesta As Integer)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn
        'CONVERT(DATETIME, '" & ComboBox7.Text & ComboBox6.Text & ComboBox5.Text & "', 112)
        CMD_SAP.CommandText = "
INSERT INTO [TIRELLI_40].[DBO].[coll_campioni_movimentazioni]
           ([tipo_movimentazione]
           ,[Id_richiesta]
           ,[id_campione]
           ,[Insertdate]
           ,[Owner]
,[segno]
           ,[Q]

           ,[mag]
)
     VALUES
           ('" & par_tipo_movimentazione & "'
           ," & par_id_richiesta & "
           ,'" & par_id_campione & "'
           ,getdate()
           ," & par_id_dipendente & "
,'" & par_segno & "' 
           ,'" & par_q & "' 
           ,'" & par_mag & "')"
        CMD_SAP.ExecuteNonQuery()

        Cnn.Close()
    End Sub

    Sub aggiusta_giacenza(par_id_campione As String, par_mag As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn
        'CONVERT(DATETIME, '" & ComboBox7.Text & ComboBox6.Text & ComboBox5.Text & "', 112)
        CMD_SAP.CommandText = " DELETE [TIRELLI_40].[DBO].[coll_campioni_giacenze] 
WHERE [id_campione] ='" & par_id_campione & "' AND [mag]= '" & par_mag & "' "
        CMD_SAP.ExecuteNonQuery()


        CMD_SAP.CommandText = " INSERT INTO [TIRELLI_40].[DBO].[coll_campioni_GIACENZE] ([id_campione]
      ,[mag]
      ,[Q_IN])
SELECT 

     '" & par_id_campione & "'
	 ,'" & par_mag & "'
	  , SUM(CASE WHEN [segno]='+' THEN Q ELSE -Q END) AS 'Q'


  FROM [TIRELLI_40].[DBO].[coll_campioni_movimentazioni]
  WHERE [mag]='" & par_mag & "' AND [id_campione]='" & par_id_campione & "'
  GROUP BY [id_campione]
 "
        CMD_SAP.ExecuteNonQuery()

        Cnn.Close()
    End Sub

    Private Sub DataGridView1_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDown
        If e.Button = MouseButtons.Left AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If ModifierKeys = Keys.Shift AndAlso startIndex >= 0 Then
                'Se è premuto Shift, cambia il flag per le righe comprese tra startIndex ed e.RowIndex
                Dim endIndex As Integer = e.RowIndex
                Dim minIndex As Integer = Math.Min(startIndex, endIndex) + 1
                Dim maxIndex As Integer = Math.Max(startIndex, endIndex) - 1

                For i As Integer = minIndex To maxIndex
                    DataGridView1.Rows(i).SetValues(True)
                Next i
            Else
                '  Altrimenti, imposta startIndex alla riga corrente
                startIndex = e.RowIndex
            End If
        End If
    End Sub

    Private Sub DataGridView_1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        ' Controlla se il tasto Shift è stato premuto
        isShiftKeyDown = (e.KeyCode = Keys.ShiftKey)
    End Sub

    Private Sub DataGridView_1_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyUp
        ' Controlla se il tasto Shift è stato rilasciato
        isShiftKeyDown = (e.KeyCode <> Keys.ShiftKey)
    End Sub
End Class