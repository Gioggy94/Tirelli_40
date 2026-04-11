Imports System
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography
Imports Newtonsoft.Json

Module Program
    Sub Main(args As String())
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

        Dim loginResponse As LoginResponse = CallRaise3d(Of LoginResponse)(String.Format("http://192.168.11.50:10800/v1/login?sign={0}&timestamp={1}",
                            strResultMd5,
                            totMs))

        If Not IsNothing(loginResponse) And loginResponse.status = 1 And Not IsNothing(loginResponse.data) And Not String.IsNullOrEmpty(loginResponse.data.token) Then
            Dim currentPrintJob As GetCurrentJobResponse = CallRaise3d(Of GetCurrentJobResponse)(String.Format("http://192.168.11.50:10800/v1/job/currentjob?token={0}",
                                loginResponse.data.token))

            If Not IsNothing(currentPrintJob) And currentPrintJob.status = 1 And Not IsNothing(currentPrintJob.data) And Not String.IsNullOrEmpty(currentPrintJob.data.job_id) Then
                'There is an active print job, let's use its data
                'TODO
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
End Module
