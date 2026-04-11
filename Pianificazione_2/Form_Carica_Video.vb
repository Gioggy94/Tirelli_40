Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Threading
Imports System.Text.RegularExpressions


Public Class Form_Carica_Video
    Public Commessa As String
    Public Combinazione As String
    Public Cartella As String

    Public Sub Aggiorna_Video()
        If Directory.Exists(Homepage.percorso_cartelle_macchine & Cartella & "\Video\") = False Then
            Directory.CreateDirectory(Homepage.percorso_cartelle_macchine & Cartella & "\Video\")
        End If
        Lbl_Nome_File.Text = Cartella & "\Video\" & Commessa & Combinazione & "_" & Now.ToString("yyyy_MM_dd_HH_mm")
    End Sub

    Private Sub Cmd_Esci_Click(sender As Object, e As EventArgs) Handles Cmd_Esci.Click
        Form_Scheda_Collaudi.Show()
        Me.Close()
    End Sub

    Private Sub Cmd_Sfoglia_Click(sender As Object, e As EventArgs) Handles Cmd_Sfoglia.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.InitialDirectory = "c:\"
        openFileDialog1.Filter = "Video|*.mpg;*.mp4;*.mov;*.wmv;*.avi;"
        openFileDialog1.FilterIndex = 1
        openFileDialog1.RestoreDirectory = True

        If openFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            TXT_Input.Text = openFileDialog1.FileName
        End If
    End Sub



    Private Sub Cmd_Salva_Video_Click(sender As Object, e As EventArgs) Handles Cmd_Salva_Video.Click
        If Form_Scheda_Collaudi.LinkLabel1.Text = "-" Then
            MsgBox("Farsi creare la cartella macchina dal PM")
        Else
            Dim File_Output As String
            Console.WriteLine(Homepage.percorso_cartelle_macchine & Lbl_Nome_File.Text)
            File_Output = Homepage.percorso_cartelle_macchine & Lbl_Nome_File.Text
            If TXT_Info.Text.Length > 0 Then
                File_Output = File_Output & "_" & TXT_Info.Text & ".mp4"
            Else
                File_Output = File_Output & ".mp4"
            End If
            If TXT_Input.Text.Length > 1 Then
                Label1.Visible = True
                MsgBox("Il caricamento può durare alcuni minuti. Attendere il messaggio di completamento")

                ' Dim Comando As String
                ' Comando = "ffmpeg -i """ & TXT_Input.Text & """ -an -b:v 1000k -c:v libx264 """ & File_Output & """"
                'Shell("ffmpeg -i """ & TXT_Input.Text & """ -vf scale=640:360 -an -b:v 1000k -c:v libx264 """ & File_Output & """")
                'Shell("ffmpeg -i """ & TXT_Input.Text & """ -vf scale=640:360 -an -b:v " & TextBox1.Text & "k -c:v libx264 """ & File_Output & """")
                'Shell("ffmpeg -i """ & TXT_Input.Text & """ -vf scale=640:360 -an -b:v " & TextBox1.Text & "k -crf 20 -c:v libx264 -preset fast """ & File_Output & """")
                ' Shell("ffmpeg -i """ & TXT_Input.Text & """ -vf scale=640:360 -an -b:v " & TextBox1.Text & "k -crf 18 -c:v libx264 -preset slow """ & File_Output & """")



                'ultima compressione utilizzata
                ' Shell("ffmpeg -i """ & TXT_Input.Text & """ -an -b:v 2000k -c:v libx264 """ & File_Output & """", AppWinStyle.Hide, True)



                System.IO.File.Copy(TXT_Input.Text, File_Output, True)




                Cmd_Sfoglia.Enabled = False
                Cmd_Salva_Video.Enabled = False
                Cmd_Guarda_Video.Enabled = False

                MsgBox("Video Caricato con Successo")
                Label1.Visible = False
                Form_Scheda_Collaudi.mostra_video()


            Else
                MsgBox("Selezionare un Video da Caricare")
            End If
        End If
    End Sub


    Private Sub Process_Exited(sender As Object, e As EventArgs)
        ' Riabilita i pulsanti e nasconde la ProgressBar
        Invoke(Sub()
                   ResetControls()
                   MsgBox("Video Caricato con Successo")
               End Sub)
    End Sub

    Private Sub ResetControls()

        Cmd_Sfoglia.Enabled = True
        Cmd_Salva_Video.Enabled = True
        Cmd_Guarda_Video.Enabled = True
    End Sub

    Private Sub Form_Carica_Video_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Private Sub Cmd_Guarda_Video_Click(sender As Object, e As EventArgs) Handles Cmd_Guarda_Video.Click
        If TXT_Input.Text.Length > 1 Then


            Process.Start(TXT_Input.Text)
        Else
            MsgBox("Selezionare un Video da Guardare")
        End If
    End Sub
End Class