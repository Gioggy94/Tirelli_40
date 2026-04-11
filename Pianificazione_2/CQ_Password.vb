Public Class CQ_Password
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = False Then
            MsgBox("Selezionare modulo a cui si vuole avere accesso")
        ElseIf RadioButton1.Checked = True And RadioButton2.Checked = False And RadioButton3.Checked = False Then
            MsgBox("Inserire modulo: Gestione autocontrollo")
        ElseIf RadioButton1.Checked = False And RadioButton2.Checked = False And RadioButton3.Checked = True Then

            CQ_Modifica_Autocontrollo.Show()


            CQ_Modifica_Autocontrollo.carica_controlli()
        ElseIf RadioButton1.Checked = False And RadioButton2.Checked = True And RadioButton3.Checked = False Then
            CQ_ModificaBP.Show()

            CQ_ModificaBP.TextBox1.Enabled = True
        Else
            TextBox2.Text = Nothing
            MsgBox("Credenziali errate")
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        RadioButton2.Checked = False
        RadioButton3.Checked = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        RadioButton1.Checked = False
        RadioButton3.Checked = False
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        RadioButton1.Checked = False
        RadioButton2.Checked = False
    End Sub

    Private Sub CQ_Password_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class