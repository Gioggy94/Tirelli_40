Public Class Discriminante_Totem_PC
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If RadioButton1.Checked = False And RadioButton2.Checked = False Then
            MsgBox("Selezionare un'opzione")

        Else
            If RadioButton1.Checked = True Then
                Homepage.totem = "Y"

            Else
                Homepage.totem = "N"

            End If
            Homepage.Aggiorna_INI_COMPUTER()

            If Homepage.totem = "Y" And Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto = Nothing Then

                Form_Cambia_Reparto.Show()
                Homepage.ID_SALVATO = Nothing
                Homepage.Label1.Text = "TOTEM"
                Me.Hide()
            ElseIf Homepage.totem = "N" And Homepage.ID_SALVATO = Nothing Then
                Form_gestione_utente.Show()
                Homepage.Label1.Text = "PC"
                Me.Hide()

            Else

                Homepage.Enabled = True
                Me.Hide()

            End If




        End If
        Me.Hide()
    End Sub
End Class