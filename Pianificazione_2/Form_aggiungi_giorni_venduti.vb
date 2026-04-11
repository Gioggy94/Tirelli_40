Imports System.Data.SqlClient

Public Class Form_aggiungi_giorni_venduti

    Public Elenco_owner(1000) As String
    Private Codice As String

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Sub inizializza_form()
        Inserimento_owner()
    End Sub

    Sub Inserimento_owner()
        ComboBox1.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code inner join [TIRELLI_40].[DBO].COLL_Reparti t2 on (t2.sap_id_reparto =t1.code or t2.sap_id_reparto_2 =t1.code)   where t0.active='Y' AND (T0.POSITION<>3 OR T0.POSITION IS NULL) and t2.id_reparto='" & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto & "'  order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_owner(Indice) = cmd_SAP_reader("Codice dipendenti")
            ComboBox1.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Private Sub Form_aggiungi_giorni_venduti_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inizializza_form()
    End Sub

    Private Sub TextBox3KeyPress(sender As Object, e As KeyPressEventArgs)
        ' Accetto solo l'immissione di numeri interi e decimali

        ' Recupero il codice ascii del tasto digitato
        ' il tasto digitato è memorizzato nella proprietà "KeyChar"
        ' dell'oggetto System.Windows.Forms.KeyPressEventArgs

        Dim KeyAscii As Short = Asc(e.KeyChar)

        ' In questo caso oltre a consentire numeri, tasto Canc
        ' e tasto BackSpace, devo consentire anche l'immissione
        ' del punto e della virgola
        If KeyAscii < 48 And KeyAscii <> 24 And KeyAscii <> 8 And e.KeyChar <> "." And e.KeyChar <> "," Then
            KeyAscii = 0
        ElseIf KeyAscii > 57 Then
            KeyAscii = 0
        End If

        ' Faccio in modo che se l'utente digita la virgola
        ' mi appaia il punto
        If e.KeyChar = "," Then
            KeyAscii = 46 ' 46 è il codice ascii del punto
        End If

        ' Il punto è si consentito
        ' ma non come primo carattere
        If TextBox3.TextLength = 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If

        ' ovviamente se c'è già un punto
        ' non è consentito digitarne altri
        If (KeyAscii = 46) And
            TextBox3.Text.IndexOf(".") > 0 Then
            KeyAscii = 0
        End If

        ' Reimposto il keychar
        e.KeyChar = Chr(KeyAscii)
    End Sub 'permetto solo numeri come input

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox1.SelectedIndex < 0 Then
            MsgBox("Seleziona un PM")

        Else
            If ComboBox2.SelectedIndex < 0 Then
                MsgBox("Seleziona un codice")
            Else

                If TextBox1.Text = "" Then
                    MsgBox("scegliere una commessa")
                Else

                    If TextBox3.Text = "" Then
                        MsgBox("scegliere quanti giorni")
                    Else

                        If RichTextBox1.Text = "" Then
                            MsgBox("scegliere una motivazione")
                        Else


                            inserisci_giorni()
                            MsgBox("giorni aggiunti con successo")
                            Form_giorni_venduti_4_0.inizializza_form()
                            Me.Close()
                        End If




                    End If


                End If

            End If
        End If
    End Sub

    Sub inserisci_giorni()
        Dim CNN3 As New SqlConnection
        CNN3.ConnectionString = Homepage.sap_tirelli
        cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = CNN3


        CMD_SAP_3.CommandText = "insert into [Tirelli_40].[dbo].giorni_venduti_4_0 (itemcode,giorni,matricola,motivazione,dipendente,data_registrazione) values ('" & Codice & "'," & TextBox3.Text & ",'" & TextBox1.Text & "','" & RichTextBox1.Text & "'," & Elenco_owner(ComboBox1.SelectedIndex) & ",getdate())"

        CMD_SAP_3.ExecuteNonQuery()
        cnn3.Close()


    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex = 0 Then
            Codice = "L00508"
        ElseIf ComboBox2.SelectedIndex = 1 Then
            Codice = "L00540"
        End If
    End Sub
End Class