Public Class CQ
    Private Sub Button1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        CQ_AttivitaAperte.N_attivita = Nothing


        CQ_AttivitaAperte.Show()
        CQ_AttivitaAperte.Attività_aperte()
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        CQ_Password.Show()
        Me.Hide()
        CQ_Password.Owner = Me
        CQ_ModificaBP.Owner = Me
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        CQ_Tabelle.Show()
        CQ_Tabelle.inizializzazione_form()


    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Procedure.Show()

    End Sub

    Private Sub CQ_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class