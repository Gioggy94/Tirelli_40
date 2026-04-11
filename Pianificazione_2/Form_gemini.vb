Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class Form_gemini

    ' ---- FORM LOAD ----
    Private Sub Form_gemini_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Forza TLS 1.2 e 1.3
        System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls13

        ' Bypass certificato (solo test, rimuovere in produzione)
        System.Net.ServicePointManager.ServerCertificateValidationCallback =
            Function(sender2, cert, chain, sslPolicyErrors) True
    End Sub

    ' ---- BOTTONE ----
    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim domanda As String = TextBox1.Text.Trim()
        If domanda = "" Then
            MsgBox("Scrivi una domanda!")
            Exit Sub
        End If

        RichTextBox1.Text = "Sto pensando..."

        ' ⛔ Inserisci qui la tua API Key di Gemini
        Dim apiKey As String = "AIzaSyAEF9YRonOgtYuyWZHM3rEIoG4yEexGnLI"

        Dim risposta As String = Await GeminiHelper.ChiediGemini(domanda, apiKey)

        RichTextBox1.Text = risposta
    End Sub

End Class

' ---- GEMINI HELPER ----
Public Class GeminiHelper

    Public Shared Async Function ChiediGemini(prompt As String, apiKey As String) As Task(Of String)

        ' Endpoint modello text-bison-001
        Dim url As String = "https://generativelanguage.googleapis.com/v1beta2/models/text-bison-001:generate"

        Try
            Using client As New HttpClient()
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)

                ' Corpo richiesta
                Dim body = New With {
                    .prompt = New With {.text = prompt},
                    .temperature = 0.7,
                    .maxOutputTokens = 500
                }

                Dim json = JsonConvert.SerializeObject(body)
                Dim content = New StringContent(json, Encoding.UTF8, "application/json")

                ' Invia POST
                Dim response = Await client.PostAsync(url, content)
                Dim result As String = Await response.Content.ReadAsStringAsync()

                ' ---- DEBUG: mostra risposta grezza ----
                ' MsgBox("Risposta grezza API:" & vbCrLf & result)

                ' Parsing JSON robusto
                Try
                    Dim parsed As JObject = JObject.Parse(result)
                    Dim text As String = Nothing

                    ' Controlla percorsi comuni del testo generato
                    If parsed.SelectToken("candidates[0].content[0].text") IsNot Nothing Then
                        text = parsed.SelectToken("candidates[0].content[0].text").ToString()
                    ElseIf parsed.SelectToken("candidates[0].content.parts[0].text") IsNot Nothing Then
                        text = parsed.SelectToken("candidates[0].content.parts[0].text").ToString()
                    End If

                    If String.IsNullOrEmpty(text) Then
                        Return "Errore: risposta JSON vuota o non valida."
                    End If

                    Return text

                Catch exJson As Exception
                    ' Se JSON non valido
                    Return "Errore: JSON non valido. Contenuto: " & result
                End Try

            End Using
        Catch ex As Exception
            Return "Errore: " & ex.Message
        End Try

    End Function

End Class