Public Class Form107
    Public Quantità_pezzi As Integer
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label_quantità_modifica.Click

    End Sub

    Private Sub Button_conferma_Click(sender As Object, e As EventArgs) Handles Button_conferma.Click

        If TextBox_quantità.Text <> Nothing Then
            Quantità_pezzi = TextBox_quantità.Text

            Dashboard_pianificazione.Show()
            Me.Hide()
        End If
    End Sub


    Private Sub TextBox_quantità_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_quantità.KeyPress
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
        If TextBox_quantità.TextLength = 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If

        ' ovviamente se c'è già un punto
        ' non è consentito digitarne altri
        If (KeyAscii = 46) And
            TextBox_quantità.Text.IndexOf(".") > 0 Then
            KeyAscii = 0
        End If

        ' Reimposto il keychar
        e.KeyChar = Chr(KeyAscii)
    End Sub 'permetto solo numeri come input

    Private Sub Quantità_modifica_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub TextBox_quantità_TextChanged(sender As Object, e As EventArgs) Handles TextBox_quantità.TextChanged

    End Sub
End Class