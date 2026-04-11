Public Class Sales_disambiguazione
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text = "[]" Then

            Me.WindowState = FormWindowState.Maximized
            Button1.Text = "Riduci"
        ElseIf Button1.Text = "Riduci" Then
            Me.WindowState = FormWindowState.Normal
            Button1.Text = "[]"
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        'Sales.lISTA_MACCHINE()

        Sales.Inserimento_dipendenti()
        If Homepage.totem = "N" Then
            Sales.ComboBox2.Text = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).cognome & " " & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).nome
        End If
        Sales.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Cambio_BP.Show()

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Layout_documenti.Show()

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Pianificazione_offerte.Show()
        Pianificazione_offerte.inizializza_form()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Opportunità.Show()
        Opportunità.inizializza_opportunità()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        'MsgBox("Funzione non più disponibile, usare Galileo")
        'Return
        form_Spare_Parts.Show()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Process.Start("\\tirfs01\00-BRB\Layout\checklist layout.xlsx")
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Dashboard.Show()
    End Sub

    Private Sub TableLayoutPanel2_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel2.Paint

    End Sub
End Class