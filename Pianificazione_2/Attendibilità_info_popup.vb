
Public Class Attendibilità_info_popup
    Public valore_attendibilità As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        valore_attendibilità = 0

        funzione()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        valore_attendibilità = 1

        funzione()
    End Sub



    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        valore_attendibilità = 2

        funzione()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        valore_attendibilità = 3

        funzione()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        valore_attendibilità = 4

        funzione()
    End Sub

    Sub funzione()
        Scheda_commessa_documentazione.Inserisci_record()

        Me.Close()
        'MsgBox("Campo aggiornato con successo")
    End Sub
End Class