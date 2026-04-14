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
    ' 1. ANALISI COMPLETA PROGRAMMA → CAPITOLI MANUALE
    ' ──────────────────────────────────────────────

    Public Async Function AnalizzaProgrammaCompleto(contenutoFile As String,
                                                    macchina As MC_Macchina,
                                                    lingua As String) As Task(Of Dictionary(Of String, String))
        Dim troncato = contenutoFile.Substring(0, Math.Min(contenutoFile.Length, 120000))

        Dim prompt = $"Sei un ingegnere industriale esperto di PLC/HMI e redattore di manuali tecnici per macchine da packaging.
Hai davanti il programma completo della macchina '{macchina.NomeMacchina}' (matricola {macchina.Matricola}, modello {macchina.Modello}).

COMPITO: Leggi il programma, simula mentalmente il funzionamento della macchina tracciando la logica, poi genera le seguenti sezioni del manuale USO E MANUTENZIONE esattamente nello stile Tirelli (vedi esempio):

=== ESEMPIO STILE CAP 5.1 ===
Bottles arriving on the conveyor enter the filling machine as indicated by the red arrows and accumulate between the worm screw (1) and the minimum load cell B2. After a short delay from the arrival of bottles in front of B2, the machine starts. The B4 cell detects the passage of the correct number of bottles and starts the filling cycle by synchronizing the movement of the sideshifter with the bottles on the belt...

=== ESEMPIO STILE CAP 5.2 ===
B4 - START CYCLE PHOTOCELL
This photocell coordinates the movement of the sideshifter with the bottles in transit and starts the filling cycle.
B2 - MINIMUM BOTTLE LOADING DETECTION PHOTOCELL
This photocell detects the quantity of bottles at the infeed...

=== FINE ESEMPIO ===

Rispondi con un JSON con questa struttura (tutti i campi in lingua {lingua}):
{{
  ""operazione"": ""[testo completo del capitolo 5.1 - Descrizione funzionamento, narrativa del ciclo macchina, menzione sensori/attuatori per codice]"",
  ""comandi"": ""[testo completo capitolo 5.2 - per ogni sensore/fotocellula/fine-corsa/attuatore: CODICE - NOME, descrizione funzione]"",
  ""allarmi"": ""[tabella allarmi in formato testo: CODICE | DESCRIZIONE | CAUSA | RIMEDIO, una riga per allarme]"",
  ""allarmi_json"": ""[array JSON degli allarmi: [{{\""CodiceErrore\"":\""..."",\""Titolo\"":\""..."",\""Descrizione\"":\""..."",\""Causa\"":\""..."",\""Rimedio\"":\""..."",\""Gravita\"":\""Avviso|Allarme|Blocco\""}},...]]""
}}

Rispondi SOLO con il JSON valido, senza testo prima/dopo, senza ```json.

PROGRAMMA MACCHINA:
{troncato}"

        Dim risposta = Await CallAPIAsync(prompt)
        Return ParseRisultatoManuale(risposta)
    End Function

    Public Function ParseAllarmiJson(json As String) As List(Of MC_CodiceErrore)
        Return ParseCodiciErroreJson(json)
    End Function

    Private Function ParseRisultatoManuale(json As String) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)
        Try
            Dim doc = JsonDocument.Parse(json.Trim())
            For Each key In {"operazione", "comandi", "allarmi", "allarmi_json"}
                Try : result(key) = doc.RootElement.GetProperty(key).GetString() : Catch : End Try
            Next
        Catch ex As Exception
            result("operazione") = $"Errore parsing risposta AI: {ex.Message}{vbLf}{vbLf}{json}"
        End Try
        Return result
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
            elenco.AppendLine($"- Codice: {f.Codice} | Tipo: {f.TipoNome}")
        Next

        Dim prompt = $"Sei un redattore tecnico esperto di manuali per macchine da packaging industriale (stile Tirelli).
Scrivi il capitolo 5.2 'Comandi e fotocellule' del manuale di uso e manutenzione per la macchina '{macchina.NomeMacchina}' (matricola {macchina.Matricola}).
Per ogni fotocellula elenca: codice, tipo, funzione tecnica nella macchina.
Usa lo stile: 'B4 - NOME FOTOCELLULA{vbLf}Descrizione della funzione.'
Lingua: {lingua}.
FOTOCELLULE INSTALLATE:
{elenco}"

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
