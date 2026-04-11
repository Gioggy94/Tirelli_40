Public Class ConfigPassword
    Private Sub BTN_OK_Click(sender As Object, e As EventArgs) Handles BTN_OK.Click
        If TXTPAssword.Text = "T1r3l11@4zero!?" Then
            LBLError.Text = ""
            TXTPAssword.Text = ""
            DialogResult = DialogResult.OK
            Close()
        Else
            LBLError.Text = "Password Errata"
            TXTPAssword.Text = ""
        End If
    End Sub

    Private Sub BTNCancel_Click(sender As Object, e As EventArgs) Handles BTNCancel.Click
        LBLError.Text = ""
        TXTPAssword.Text = ""
        Close()
    End Sub

    Private Sub ConfigPassword_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TXTPAssword.Focus()
    End Sub
End Class