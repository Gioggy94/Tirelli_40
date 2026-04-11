Imports System.Net.Http
Imports System.Text
Imports Newtonsoft.Json
Imports PdfSharp.Quality
Imports Newtonsoft.Json.Linq




Public Class Form_IA


    Public Class AIHelper

        Public Shared Async Function ChiediAI(prompt As String) As Task(Of String)
            Dim apiKey As String = ""  ' <-- inserisci la tua API key
            Dim url As String = "https://api.openai.com/v1/chat/completions"

            Dim client As New HttpClient()
            client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)

            ' Corpo della richiesta
            Dim body = New With {
            .model = "gpt-4o-mini",
            .messages = {
                New With {.role = "user", .content = prompt}
            }
        }

            Dim json = JsonConvert.SerializeObject(body)
            Dim content = New StringContent(json, Encoding.UTF8, "application/json")

            ' Invio richiesta
            Dim response = Await client.PostAsync(url, content)
            Dim result = Await response.Content.ReadAsStringAsync()

            ' Parsing sicuro della risposta
            Dim parsed As JObject = JObject.Parse(result)
            Dim contentToken = parsed.SelectToken("choices[0].message.content")

            ' Se non c'è, prova choices[0].text
            If contentToken Is Nothing Then
                contentToken = parsed.SelectToken("choices[0].text")
            End If

            If contentToken IsNot Nothing Then
                Return contentToken.ToString().Trim()
            Else
                Return "Errore: risposta vuota o struttura JSON diversa." & vbCrLf & "JSON ricevuto: " & result
            End If
        End Function

    End Class

    Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim domanda = TextBox1.Text.Trim()

        If domanda = "" Then
            MsgBox("Scrivi una domanda!")
            Exit Sub
        End If

        RichTextBox1.Text = "Sto pensando..."

        Dim risposta = Await AIHelper.ChiediAI(domanda)

        RichTextBox1.Text = risposta
    End Sub
End Class