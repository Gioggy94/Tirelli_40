Imports System
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports Newtonsoft.Json
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word
Imports MailMessage = System.Net.Mail.MailMessage

Public Class Stampante_3D
    Public id As Integer
    Public file_name As String
    Public number_of_layer As Integer
    Public total_time As Integer
    Public file_status As String

    Public ultimo_file_name As String
    Public ultimo_file_status As String


    Public percorso_documento As String

    Private Sub Stampante_3D_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        risposta_stampante()
        Timer1.Start()
    End Sub

    Sub risposta_stampante()
        'Create parameters for login call
        Dim totMs As Long = DateTimeOffset.Now.ToUnixTimeMilliseconds()

        Dim loginString As String = String.Format("password={0}&timestamp={1}",
            "377956", 'printer password
            totMs)

        Dim sha1Obj As New Security.Cryptography.SHA1CryptoServiceProvider
        Dim sha1Hash() As Byte = sha1Obj.ComputeHash(System.Text.Encoding.ASCII.GetBytes(loginString))
        Dim strResultSha1 As String = BitConverter.ToString(sha1Hash).Replace("-", "").ToLower()

        Dim hasher As MD5 = MD5.Create()
        Dim md5Hash As Byte() = hasher.ComputeHash(System.Text.Encoding.UTF8.GetBytes(strResultSha1))
        Dim strResultMd5 As String = BitConverter.ToString(md5Hash).Replace("-", "").ToLower()

        Dim loginResponse As LoginResponse = CallRaise3d(Of LoginResponse)(String.Format("http://10.7.111.21:10800/v1/login?sign={0}&timestamp={1}",
                            strResultMd5,
                            totMs))

        If Not IsNothing(loginResponse) And loginResponse.status = 1 And Not IsNothing(loginResponse.data) And Not String.IsNullOrEmpty(loginResponse.data.token) Then
            Dim currentPrintJob As GetCurrentJobResponse = CallRaise3d(Of GetCurrentJobResponse)(String.Format("http://10.7.111.21:10800/v1/job/currentjob?token={0}",
                                loginResponse.data.token))
            file_name = currentPrintJob.data.file_name
            Label1.Text = file_name
            file_status = currentPrintJob.data.job_status
            Label2.Text = file_status
            Label3.Text = currentPrintJob.data.printed_layer
            Try
                Label5.Text = Math.Round(currentPrintJob.data.print_progress, 2) & " %"
            Catch ex As Exception
                Label5.Text = "0 %"
            End Try


            ProgressBar1.Value = currentPrintJob.data.print_progress
            number_of_layer = currentPrintJob.data.total_layer
            Label6.Text = number_of_layer

            Try
                ProgressBar2.Value = Math.Round(currentPrintJob.data.printed_layer / currentPrintJob.data.total_layer * 100, 2)
            Catch ex As Exception
                ProgressBar2.Value = 0
            End Try

            Try
                Label8.Text = Math.Round(currentPrintJob.data.printed_layer / currentPrintJob.data.total_layer * 100, 2) & " %"
            Catch ex As Exception
                Label8.Text = "0 %"
            End Try



            Label4.Text = New TimeSpan(0, 0, currentPrintJob.data.printed_time).ToString
            total_time = currentPrintJob.data.total_time
            Label7.Text = New TimeSpan(0, 0, currentPrintJob.data.total_time).ToString


            If Not IsNothing(currentPrintJob) And currentPrintJob.status = 1 And Not IsNothing(currentPrintJob.data) And Not String.IsNullOrEmpty(currentPrintJob.data.job_id) Then
                'There is an active print job, let's use its data

            End If
        End If
        compila_datagridview()
        If Homepage.ID_SALVATO = 221 Then
            Trova_ultimo_record()
            If ultimo_file_name <> file_name Or ultimo_file_status <> file_status Then

                insert_into_stampante_3d_log()

                ' Invia_Mail("Avanzamento stampante file: " & file_name, "lo stato è: " & file_status, "fabiopassirani@tirelli.net", "davidebalasini@tirelli.net", "Notifica stato avanzamento stampante 3D")

            End If
        End If


    End Sub

    Public Function CallRaise3d(Of T)(ByVal requestUri As String) As T
        Dim retVal As T = Nothing

        Dim httpRequest As HttpWebRequest = WebRequest.CreateHttp(requestUri)
        httpRequest.Method = "GET"

        Using httpResponse As HttpWebResponse = httpRequest.GetResponse()
            Dim responseStream As Stream = httpResponse.GetResponseStream()
            Dim sr As New StreamReader(responseStream)
            Dim result As String = sr.ReadToEnd()
            retVal = JsonConvert.DeserializeObject(Of T)(result)
        End Using

        Return retVal
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        risposta_stampante()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
    End Sub

    Sub insert_into_stampante_3d_log()
        Trova_ID()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "insert into [Tirelli_40].[dbo].[Stampante_3D_log] (id,date,Ora,File_name,N_layer,Tempo_totale,Status) values (" & id & ", getdate(),convert(varchar, getdate(), 108),'" & file_name & "'," & number_of_layer & "," & total_time & ",'" & file_status & "')"
        CMD_SAP.ExecuteNonQuery()


        cnn.Close()

    End Sub

    Sub Trova_ID()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn
        CMD_SAP_2.CommandText = "select max(id)+1 as 'ID' from [Tirelli_40].[dbo].[Stampante_3D_log]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                id = cmd_SAP_reader_2("ID")
            Else
                id = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        cnn.Close()
    End Sub

    Sub Trova_ultimo_record()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn
        CMD_SAP_2.CommandText = "select t0.id, t0.File_name,t0.status
from [Tirelli_40].[dbo].[Stampante_3D_log] t0 inner join (
select max(id) as 'ID' from [Tirelli_40].[dbo].[Stampante_3D_log]) A on a.id=t0.id"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            ultimo_file_name = cmd_SAP_reader_2("File_name")
            ultimo_file_status = cmd_SAP_reader_2("status")
            cmd_SAP_reader_2.Close()
        End If
        cnn.Close()
    End Sub

    Sub compila_datagridview()
        DataGridView.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn
        CMD_SAP_2.CommandText = "select top 100 t0.date, t0.ora, t0.File_name,t0.status, t0.N_layer, t0.Tempo_totale
from [Tirelli_40].[dbo].[Stampante_3D_log] t0 
order by t0.id desc"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            DataGridView.Rows.Add(cmd_SAP_reader_2("date"), cmd_SAP_reader_2("ora"), cmd_SAP_reader_2("File_name"), cmd_SAP_reader_2("status"), cmd_SAP_reader_2("N_layer"), New TimeSpan(0, 0, cmd_SAP_reader_2("tempo_totale").ToString))


        Loop
        cmd_SAP_reader_2.Close()
        cnn.Close()
    End Sub


    Public Sub Invia_Mail(body As String, testo_mail_1 As String, destinatario_1 As String, destinatario_2 As String, subject As String)

        Dim Testo_Mail As String
        Testo_Mail = "<BODY><H3>" & body & "</h3><P>"


        Testo_Mail = Testo_Mail & "<BR><BR>" & testo_mail_1 & ""

        Testo_Mail = Testo_Mail & "</P></BODY>"


        Dim mySmtp As New SmtpClient
        Dim myMail As New MailMessage()
        mySmtp.UseDefaultCredentials = False
        mySmtp.Credentials = New Net.NetworkCredential(Homepage.Mittente_Mail, Pianificazione_Tickets.Password_Mail)
        mySmtp.Host = "tirelli-net.mail.protection.outlook.com"

        mySmtp.Port = 25


        mySmtp.EnableSsl = True
        mySmtp.DeliveryMethod = SmtpDeliveryMethod.Network


        myMail = New MailMessage()
        myMail.From = New MailAddress(Homepage.Mittente_Mail)
        myMail.To.Add(destinatario_1)
        myMail.To.Add(destinatario_2)

        myMail.Bcc.Add("report@tirelli.net")
        myMail.Subject = subject
        myMail.IsBodyHtml = True
        myMail.Body = Testo_Mail

        Try
            mySmtp.Send(myMail)
        Catch ex As Exception
            MsgBox("Errore Invio Mail" & ex.ToString)
        End Try

    End Sub


    Sub InviaEmailConAllegato()

        Dim objOutlook As Object
        Dim objMail As Object
        Dim strEmail As String
        Dim strSubject As String
        Dim strBody As String
        Dim strAttachmentPath As String

        'Imposta i valori dei campi email, oggetto, corpo e percorso dell'allegato
        strEmail = "indirizzoemail@destinatario.com"
        strSubject = "Oggetto della mail"
        strBody = "Buongiorno, "
        strAttachmentPath = Layout_documenti.percorso_documento_PDF

        'Crea un oggetto Outlook e una nuova email
        objOutlook = CreateObject("Outlook.Application")
        objMail = objOutlook.CreateItem(0)

        'Imposta i campi della nuova email
        With objMail
            .To = strEmail
            .Subject = strSubject
            .HTMLBody = strBody
            .Display 'Apre la mail in anteprima
            .Attachments.Add(strAttachmentPath)

        End With

        'Rilascia gli oggetti creati
        objMail = Nothing
        objOutlook = Nothing

    End Sub


End Class