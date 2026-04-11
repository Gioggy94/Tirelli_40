Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Text.Json
Imports System.Threading.Tasks

Public Class MC_AnthropicService

    Private Const MODEL As String = "claude-opus-4-5"
    Private Const API_URL As String = "https://api.anthropic.com/v1/messages"
    Private ReadOnly _client As New HttpClient()

    Public Sub New()
        _client.DefaultRequestHeaders.Add("x-api-key", GetApiKey())
        _client.DefaultRequestHeaders.Add("anthropic-version", "2023-06-01")
        _client.Timeout = TimeSpan.FromMinutes(5)
    End Sub

    Private Shared Function GetApiKey() As String
        ' Cambiamo il nome della variabile in "filePath" o "keyPath"
        Dim filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "anthropic_key.txt")

        If File.Exists(filePath) Then
            Return File.ReadAllText(filePath).Trim()
        End If
        Return ""
    End Function

    ' ──────────────────────────────────────────────
    ' 1. ANALISI FILE SORGENTE PLC
    ' ──────────────────────────────────────────────

    Public Async Function AnalizzaSoftwarePLC(contenutoFile As String,
                                              nomeMacchina As String,
                                              lingua As String) As Task(Of List(Of MC_CodiceErrore))
        Dim prompt = $"Sei un esperto di PLC industriali e documentazione tecnica.
Analizza il seguente codice sorgente PLC/HMI della macchina '{nomeMacchina}'.
Individua TUTTI i codici errore/allarme presenti (es. prefissi ALM_, E_, ERR_, ALARM, ecc.).
Per ciascuno restituisci un oggetto JSON con questi campi:
  - CodiceErrore (string): il codice esatto
  - Titolo (string): titolo breve dell'errore
  - Descrizione (string): descrizione tecnica dettagliata
  - Causa (string): possibili cause
  - Rimedio (string): azioni correttive
  - Gravita (string): uno tra 'Avviso', 'Allarme', 'Blocco'
Lingua di output: {lingua}.
Rispondi SOLO con un array JSON valido, senza testo prima o dopo, senza ```json.

CODICE SORGENTE:
{contenutoFile.Substring(0, Math.Min(contenutoFile.Length, 80000))}"

        Dim risposta = Await CallAPIAsync(prompt)
        Return ParseCodiciErroreJson(risposta)
    End Function

    ' ──────────────────────────────────────────────
    ' 2. ANALISI SCREENSHOT SCHERMATA ERRORE
    ' ──────────────────────────────────────────────

    Public Async Function AnalizzaScreenshotErrore(imagePath As String,
                                                   nomeMacchina As String,
                                                   lingua As String) As Task(Of MC_CodiceErrore)
        Dim imageBytes = File.ReadAllBytes(imagePath)
        Dim base64 = Convert.ToBase64String(imageBytes)
        Dim ext = Path.GetExtension(imagePath).ToLower()
        Dim mediaType = If(ext = ".png", "image/png",
                        If(ext = ".jpg" OrElse ext = ".jpeg", "image/jpeg", "image/png"))

        Dim prompt = $"Sei un esperto di HMI industriali e documentazione tecnica.
Analizza questa schermata del pannello operatore della macchina '{nomeMacchina}'.
Identifica: il codice errore visualizzato, il messaggio di errore, eventuali dettagli aggiuntivi.
Restituisci un oggetto JSON con:
  - CodiceErrore (string): codice esatto visibile a schermo
  - Titolo (string): titolo/messaggio dell'errore
  - Descrizione (string): descrizione tecnica
  - Causa (string): possibili cause dell'errore
  - Rimedio (string): azioni correttive per l'operatore
  - Gravita (string): uno tra 'Avviso', 'Allarme', 'Blocco'
Lingua di output: {lingua}.
Rispondi SOLO con l'oggetto JSON, senza testo prima o dopo, senza ```json."

        Dim payload = New With {
            .model = MODEL,
            .max_tokens = 1500,
            .messages = New Object() {
                New With {
                    .role = "user",
                    .content = New Object() {
                        New With {
                            .type = "image",
                            .source = New With {
                                .type = "base64",
                                .media_type = mediaType,
                                .data = base64
                            }
                        },
                        New With {.type = "text", .text = prompt}
                    }
                }
            }
        }

        Dim json = JsonSerializer.Serialize(payload)
        Dim risposta = Await PostJsonAsync(json)
        Return ParseSingoloCodiceErrore(risposta)
    End Function

    ' ──────────────────────────────────────────────
    ' 3. GENERAZIONE TESTO MANUALE
    ' ──────────────────────────────────────────────

    Public Async Function GeneraCapitoloFotocellule(fotocellule As List(Of MC_Fotocellula),
                                                    macchina As MC_Macchina,
                                                    lingua As String) As Task(Of String)
        Dim elenco As New StringBuilder()
        For Each f In fotocellule
            elenco.AppendLine($"- Codice: {f.Codice} | Marca: {f.Marca} | Modello: {f.Modello} | Tipo: {f.TipoRilevazione} | Posizione: {f.Posizione} | Tensione: {f.TensioneLavoro} | Uscita: {f.UscitaLogica} | Distanza: {f.DistanzaRilev}")
            If Not String.IsNullOrEmpty(f.NoteInstallaz) Then
                elenco.AppendLine($"  Note: {f.NoteInstallaz}")
            End If
        Next

        Dim prompt = $"Sei un redattore tecnico esperto di manuali per macchine da packaging industriale.
Scrivi il capitolo 5.1 'Fotocellule' del manuale di uso e manutenzione per la macchina '{macchina.NomeMacchina}' (matricola {macchina.Matricola}).
Usa un linguaggio tecnico preciso, adatto a un manuale industriale professionale.
Lingua: {lingua}.
Per ogni fotocellula scrivi: posizione, funzione, dati tecnici, note di installazione e manutenzione.
Includi una breve introduzione generale sulle fotocellule presenti nella macchina.
DATI FOTOCELLULE:
{elenco}
Formatta il testo con sotto-paragrafi per ogni fotocellula (5.1.1, 5.1.2, ecc.)."

        Return Await CallAPIAsync(prompt)
    End Function

    Public Async Function GeneraCapitoloErrori(errori As List(Of MC_CodiceErrore),
                                               macchina As MC_Macchina,
                                               lingua As String) As Task(Of String)
        Dim elenco As New StringBuilder()
        For Each e In errori
            elenco.AppendLine($"Codice: {e.Codice} | Gravità: {e.Gravita} | Titolo: {e.Titolo}")
            elenco.AppendLine($"  Descrizione: {e.Descrizione}")
            elenco.AppendLine($"  Causa: {e.Causa}")
            elenco.AppendLine($"  Rimedio: {e.Rimedio}")
            elenco.AppendLine()
        Next

        Dim prompt = $"Sei un redattore tecnico. Scrivi il capitolo 'Messaggi di allarme e codici errore' del manuale della macchina '{macchina.NomeMacchina}'.
Lingua: {lingua}.
Includi un'introduzione, poi presenta ogni errore con: codice, gravità, descrizione, causa, rimedio.
DATI ERRORI:
{elenco}"

        Return Await CallAPIAsync(prompt)
    End Function

    ' ──────────────────────────────────────────────
    ' PRIVATI
    ' ──────────────────────────────────────────────

    Private Async Function CallAPIAsync(prompt As String) As Task(Of String)
        Dim payload = New With {
            .model = MODEL,
            .max_tokens = 4000,
            .messages = New Object() {New With {.role = "user", .content = prompt}}
        }
        Return Await PostJsonAsync(JsonSerializer.Serialize(payload))
    End Function

    Private Async Function PostJsonAsync(jsonPayload As String) As Task(Of String)
        Dim content = New StringContent(jsonPayload, Encoding.UTF8, "application/json")
        Dim resp = Await _client.PostAsync(API_URL, content)
        Dim body = Await resp.Content.ReadAsStringAsync()
        If Not resp.IsSuccessStatusCode Then
            Throw New Exception($"API Anthropic errore {resp.StatusCode}: {body}")
        End If
        Dim doc = JsonDocument.Parse(body)
        Return doc.RootElement.GetProperty("content")(0).GetProperty("text").GetString()
    End Function

    Private Function ParseCodiciErroreJson(json As String) As List(Of MC_CodiceErrore)
        Dim lista As New List(Of MC_CodiceErrore)
        Try
            Dim doc = JsonDocument.Parse(json.Trim())
            For Each el In doc.RootElement.EnumerateArray()
                lista.Add(ParseErroreElement(el))
            Next
        Catch ex As Exception
            lista.Add(New MC_CodiceErrore With {
                .Codice = "PARSE_ERROR",
                .Titolo = "Errore parsing risposta AI",
                .Descrizione = ex.Message & " | " & json.Substring(0, Math.Min(500, json.Length))
            })
        End Try
        Return lista
    End Function

    Private Function ParseSingoloCodiceErrore(json As String) As MC_CodiceErrore
        Try
            Dim doc = JsonDocument.Parse(json.Trim())
            Return ParseErroreElement(doc.RootElement)
        Catch ex As Exception
            Return New MC_CodiceErrore With {.Codice = "PARSE_ERROR", .Titolo = "Errore parsing AI", .Descrizione = ex.Message}
        End Try
    End Function

    Private Function ParseErroreElement(el As JsonElement) As MC_CodiceErrore
        Return New MC_CodiceErrore With {
            .Codice      = GetStr(el, "CodiceErrore"),
            .Titolo      = GetStr(el, "Titolo"),
            .Descrizione = GetStr(el, "Descrizione"),
            .Causa       = GetStr(el, "Causa"),
            .Rimedio     = GetStr(el, "Rimedio"),
            .Gravita     = GetStr(el, "Gravita", "Avviso")
        }
    End Function

    Private Function GetStr(el As JsonElement, key As String, Optional def As String = "") As String
        Try : Return el.GetProperty(key).GetString()
        Catch : Return def
        End Try
    End Function

End Class
