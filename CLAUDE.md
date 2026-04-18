# CLAUDE.md — Pianificazione 249

## Modulo: Riconoscimento Bolle di Entrata Merce

### File coinvolti
- `Pianificazione_2/Entrate_merci.vb` — form VB.NET principale
- `Pianificazione_2/Entrate_merci.Designer.vb` — designer del form
- `Pianificazione_2/entrate_merci_ocr.py` — script Python che chiama Claude API per analizzare il PDF
- `Pianificazione_2/bin/Debug/entrate_merci_ocr.py` — copia in output (da tenere allineata manualmente con il sorgente)

---

### Architettura generale

1. L'utente seleziona un PDF di bolla DDT fornitore
2. Il VB.NET lancia lo script Python passando il path del PDF e la chiave API Claude
3. Lo script Python converte ogni pagina PDF in immagine (PyMuPDF/fitz) e chiede a Claude Opus di estrarre le righe articolo come JSON
4. Il JSON viene restituito su stdout e parsato dal VB.NET
5. Le righe vengono mostrate nella griglia `dgvRighe`
6. La funzione `VerificaOrdiniAS400()` confronta ogni riga con gli ordini aperti su AS400 tramite OPENQUERY su SQL Server collegato (linked server `AS400`)

---

### Estrazione PDF — script Python

**Prompt a Claude:** chiede di restituire un array JSON con i campi:
`ddt_numero`, `ddt_data`, `fornitore`, `ordine`, `codice`, `descrizione`, `um`, `quantita`

**Regola ordine:** cercare "Vs. Ordine" / "Vs. ord." / "Vostro Ordine" — quello è il numero ordine Tirelli.
- "OVC" = numero conferma ordine del fornitore → IGNORARE
- "Ns. Ordine" = numero interno fornitore → IGNORARE
- Il nostro ordine ha tipicamente 6 cifre (es. 202211, 202841)

**Codice articolo — regola:**
- Se il DDT ha una colonna "Vs. Codice" / "Vostro Codice" / "Cod. Cliente" → usare QUELLA (è il nostro codice Tirelli)
- Altrimenti usare il codice alfanumerico principale del fornitore
- Motivo: alcuni fornitori (es. ZM Automazione srl, cod forn 1410005560) riportano il proprio codice interno come colonna principale e il nostro codice in "Vs. Codice"

**Quantità — regola critica:**
- Il campo `quantita` viene richiesto come **stringa grezza** (testo esatto dal PDF, es. `"1.000"`, `"2,5"`)
- I fornitori usano punto O virgola **sempre e solo come separatore decimale**, mai come separatore delle migliaia
- Dopo aver ricevuto la risposta da Claude, lo script Python sanitizza:
  ```python
  q = str(item["quantita"]).strip().replace(",", ".")
  item["quantita"] = float(q)
  ```
- Questo garantisce: `"1.000"` → 1.0, `"1,000"` → 1.0, `"2,5"` → 2.5, `"1000"` → 1000.0
- **Motivo:** se si chiede a Claude di restituire direttamente un numero, interpreta "1.000" come 1000 (notazione italiana migliaia)

---

### Query AS400 — `VerificaOrdiniAS400()`

La query OPENQUERY verso AS400 usa SQL Server come tramite (linked server `AS400`):

```sql
SELECT * FROM OPENQUERY(AS400,
'SELECT trim(codart) as codart, trim(disegno) as disegno,
 t0.numdoc as n_documento, t0.qta_ord as Q, data_richiesta, evaso
 FROM TIR90VIS.JGALord t0
 WHERE DOC = ''OA''
   and (evaso <> ''S''
   or data_richiesta >= INTEGER(TO_CHAR(CURRENT DATE - 100 DAYS, ''YYYYMMDD'')))')
```

**Regola escaping virgolette in OPENQUERY:**
- In VB.NET string: `''S''` (2 apici per lato)
- SQL Server vede nel literal: `''S''` → AS400 riceve: `'S'` ✓
- Con `''''S''''` (4 apici) → AS400 riceve `''S''` → errore di sintassi DB2

**Campi restituiti:**
- `codart` — codice articolo AS400
- `disegno` — numero di disegno AS400
- `n_documento` — numero ordine di acquisto
- `Q` — quantità ordinata (formato numerico AS400 con virgola decimale)

---

### Confronto ordini — logica di matching

#### 1. Matching numero ordine — `OrdiniCorrispondono(a, b)`
Il numero ordine nel DDT può avere meno cifre del numero in AS400:
- DDT: `202211` — AS400: `A0202211`
- Logica: se i due stringhe hanno lunghezze diverse, il più lungo deve terminare con il più corto
- `"A0202211".EndsWith("202211")` → True ✓

#### 2. Matching codice/disegno — `CodiciCorrispondono(a, b)` + `NormalizzaCodice(s)`
Alcuni fornitori (es. Lasergi) usano `#` per indicare la revisione, mentre AS400 usa `-R`:
- DDT: `D115918#02` — AS400: `D115918-R02`
- Normalizzazione: `Regex.Replace(s, "#(\d+)", "-R$1")`
- `D115918#02` → `D115918-R02` = `D115918-R02` ✓
- Il confronto avviene sia su `codart` che su `disegno` dell'AS400

#### 3. Parsing quantità AS400 — tre interpretazioni
L'AS400 restituisce le quantità con separatori variabili per fornitore:

| Fornitore | Formato AS400 | Interpr. A (InvariantCulture) | Interpr. B (it-IT) | Interpr. C (rimuovi sep.) |
|-----------|--------------|-------------------------------|---------------------|---------------------------|
| G&G Service | `1,0000` | 1 ✓ | 1 | 10000 |
| MB Meccanica | `1.000` | 1 | 1000 | 1000 |

Logica (in ordine di priorità):
```vbnet
If Math.Abs(qtaDDT - qtaOC_A) < 0.001D Then qtaOC = qtaOC_A
ElseIf Math.Abs(qtaDDT - qtaOC_B) < 0.001D Then qtaOC = qtaOC_B
ElseIf Math.Abs(qtaDDT - qtaOC_C) < 0.001D Then qtaOC = qtaOC_C
Else qtaOC = qtaOC_A  ' mostra scostamento
```

- **A:** `qRaw.Replace(",", ".")` → InvariantCulture
- **B:** `qRaw` → it-IT (punto=migliaia, virgola=decimale)
- **C:** `qRaw.Replace(".", "").Replace(",", "")` → InvariantCulture (rimuove tutti i separatori)

#### 4. Fallback codice da descrizione — `EstraiCodiceDaDescrizione()`
Alcuni fornitori (es. Thenar, cod 1410000541) non mettono il codice Tirelli nella colonna codice, ma lo scrivono nella **descrizione** della riga con la dicitura "Vs. Codice:", "Vs.Codice:", "Vostro Codice:" ecc.

Se il primo matching fallisce (ordine trovato, codice non trovato), `VerificaOrdiniAS400()` cerca nella colonna `colDescrizione` un pattern:
```
(?:vs\.?\s*codice(?:\s+prodotto)?|vostro\s+codice(?:\s+prodotto)?)\s*[:\s]\s*([A-Z0-9][A-Z0-9\-\.#_]*)
```
e ritenta il matching con il codice estratto. Funziona per **tutti i fornitori** in modo trasparente, senza configurazione specifica.

#### 5. Stati risultato
- Verde — **OK**: quantità DDT = quantità ordine AS400
- Giallo — **scostamento**: ordine trovato, codice trovato, ma quantità diverse → mostra `Q DDT=X  OA=Y`
- Arancio — **Codice non in questo ordine**: ordine trovato ma il codice/disegno non è presente (nemmeno nella descrizione)
- Rosso — **Ordine non trovato**: nessun ordine aperto in AS400 con quel numero

---

### Registro peculiarità fornitore — `RegoleFornitore`

Dizionario statico a livello di classe in `Entrate_merci.vb`:
- **Chiave:** `COD_forn` AS400 (trimmed, case-insensitive)
- **Valore:** array di coppie `(pattern regex, sostituzione)` da applicare al codice/disegno estratto dal DDT prima del confronto con AS400

```vbnet
Private Shared ReadOnly RegoleFornitore As New Dictionary(Of String, String())(StringComparer.OrdinalIgnoreCase) From {
    {"LASERGI", New String() {"#(\d+)", "-R$1"}}
}
```

**Come aggiungere un nuovo fornitore:**
```vbnet
{"COD_FORN", New String() {"pattern1", "repl1", "pattern2", "repl2"}}
```

**Fornitori registrati:**

| COD_forn | Peculiarità | Pattern | Sostituzione |
|----------|-------------|---------|--------------|
| 1410002492 (Lasergi) | Usa `#N` invece di `-RN` (revisione) | `#(\d+)` | `-R$1` |
| 1410002492 (Lasergi) | Scrive `-` invece di `_` | `_` | `-` |
| 1410005129 (Vanoni SRL) | Aggiunge prefisso `TIR` davanti al codice Tirelli (es. `TIRD122578` → `D122578`) | `^TIR` | `` |

**Peculiarità estratzione ordine per fornitore (solo prompt OCR, nessun RegoleFornitore VB.NET):**

| COD_forn | Fornitore | Peculiarità formato DDT |
|----------|-----------|-------------------------|
| 1410000465 | BETT Sistemi Srl | Campo "VS.RIF.ORDINE / YOUR ORDER REF." su due righe: prima riga = numero ordine (es. `203144`), seconda riga = `del 270326` (data 27/03/26). La data va ignorata, solo il primo numero è l'ordine. |
| 1410000541 | Thenar | Il codice Tirelli non è nella colonna codice ma scritto nella descrizione riga come "Vs. Codice: XXXXX". Il VB gestisce questo in automatico via fallback `EstraiCodiceDaDescrizione()`. |

**Nota:** esiste anche una regola universale `#N → -RN` applicata a tutti i fornitori (prima delle regole specifiche), nel caso in cui il fornitore non sia nel registro.

Il campo `cod_forn` è estratto dalla query AS400 (`trim(cod_forn) as cod_forn`) e passato a `CodiciCorrispondono(a, b, codForn)` → `NormalizzaCodice(s, codForn)`.

---

### Connessioni database
- `Homepage.sap_tirelli` — connection string SQL Server usata per OPENQUERY verso AS400
- Il linked server si chiama `AS400`
- La tabella ordini è `TIR90VIS.JGALord`
- Campi usati: `codart`, `disegno`, `numdoc`, `qta_ord`, `data_richiesta`, `evaso`, `cod_forn`

---

### Dipendenze Python
```
pip install pymupdf anthropic
```

---

### Note operative
- Lo script Python in `bin/Debug/` deve essere sempre allineato con il sorgente in `Pianificazione_2/` — aggiornarli entrambi ad ogni modifica
- La chiave API Claude viene letta da `anthropic_key.txt` nella cartella di avvio applicazione
