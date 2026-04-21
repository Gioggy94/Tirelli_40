#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script per il riconoscimento automatico delle bolle di entrata merce (DDT)
Uso: python entrate_merci_ocr.py <pdf_path> <anthropic_api_key>
Output: JSON su stdout con la lista degli articoli estratti
Errori: JSON con campo "errore" su stdout
"""

import sys
import os
import json
import base64

def converti_pagina_in_base64(page, zoom=2.0):
    import fitz
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    return base64.standard_b64encode(pix.tobytes("png")).decode("utf-8")

def analizza_pagina_con_claude(client, img_base64):
    import anthropic

    prompt = """Analizza questo documento DDT (Documento di Trasporto) italiano e estrai TUTTE le righe di articoli.

Per ogni riga articolo restituisci un oggetto JSON con questi campi:
- ddt_numero: numero del documento DDT/bolla (campo "N. DOCUMENTO" o "N" sul documento)
- ddt_data: data del documento in formato GG/MM/AAAA
- fornitore: ragione sociale del fornitore mittente (non il destinatario Tirelli)
- ordine: numero dell'ORDINE DI ACQUISTO TIRELLI. Regole precise:
  * Cerca la dicitura "Vs. Ordine", "Vs. ord.", "Vostro Ordine", "Your Order", "VS.RIF.ORDINE", "Riferimento" seguita da un numero di circa 6 cifre: QUELLO e il nostro ordine.
  * "OVC" e un numero di CONFERMA ORDINE del fornitore, NON e il nostro ordine di acquisto: IGNORALO completamente.
  * "Ns. Ordine" o "Nostro Ordine" e il numero interno del fornitore: IGNORALO.
  * Il nostro ordine ha tipicamente 6 cifre (es. 202211, 201229, 202841).
  * ATTENZIONE: alcuni fornitori (es. BETT) scrivono l'ordine su due righe nel campo riferimento, nella forma "203144 / del 270326": il numero ordine e SOLO il primo numero (203144), mentre "del 270326" e la DATA dell'ordine (27/03/26) e va IGNORATA. Non includere mai la data nel campo ordine.
  * Se ci sono piu ordini Tirelli nel documento, associa ogni articolo al suo ordine corretto.
- codice: il NOSTRO codice articolo Tirelli. Regole precise:
  * Se il documento ha una colonna "Vs. Codice", "Vostro Codice", "Cod. Cliente", "Codice Cliente", "Vs. Art." o simili, usa QUELLO: e il nostro codice.
  * Se la colonna si chiama "CODICE ARTICOLO/VS CODICE" (o simile combinazione), ogni cella contiene DUE codici sovrapposti: il PRIMO e il codice del fornitore, il SECONDO (solitamente in corsivo o piu piccolo, sotto) e il NOSTRO codice Tirelli. Usa SEMPRE il secondo. Esempio concreto: Terreni Industriale (fornitore) scrive nella stessa cella il proprio codice in alto (es. "252-1606A4") e il codice Tirelli in basso (es. "C23212", inizia sempre per "C"): devi usare "C23212", NON "252-1606A4".
  * Se nella descrizione o nel corpo della riga appare una delle seguenti diciture seguite da un codice alfanumerico, quel codice e il NOSTRO codice Tirelli: usa QUELLO e ignora il codice nella colonna principale (che e il codice del fornitore).
    Diciture da riconoscere (tutte le varianti, con o senza punto, con o senza spazio, con o senza due punti):
    "Vs.Codice:", "Vs. Codice:", "Vs.Codice", "Vs. Codice",
    "Vs.Art.:", "Vs. Art.:",
    "Vs cod:", "Vs. cod:", "Vs cod", "Vs.cod:",
    "Vostro Codice:", "Vostro Codice Prodotto:", "Vs. Codice Prodotto:", "Vs.Codice Prodotto:"
    Esempio Thenar: descrizione contiene "Vs.Codice: C00369" -> codice=C00369
    Esempio DETAS: descrizione contiene "Vs cod:C14413" -> codice=C14413, NON "CN40169300" (che e il Cod.Taric/doganale, da ignorare)
  * ATTENZIONE colonne da NON usare come codice articolo: "Cod.Taric", "Codice Taric", "HS Code", "Codice Doganale", "CIG", "CUP" — questi non sono mai codici articolo.
  * Solo se non esiste nessuna delle indicazioni sopra usa il codice alfanumerico principale del fornitore (es. D20036, 10667-005, D131379-L).
- descrizione: descrizione dell'articolo. IMPORTANTE: se nella riga e presente una dicitura tipo "Vs cod:", "Vs.Codice:" o simili, includi TUTTO il testo della cella descrizione cosi com'e (inclusa la riga "Vs cod:XXXXX"), in modo che il sistema possa usarla come riferimento di backup.
- um: unita di misura (PZ, NR, KG, MT ecc.)
- quantita: il TESTO ESATTO della quantita come appare nel documento, come stringa (non come numero). Non interpretare, non convertire. Se il PDF mostra "1.000" restituisci "1.000", se mostra "1,000" restituisci "1,000", se mostra "1000" restituisci "1000", se mostra "2,5" restituisci "2,5".

Regole importanti:
1. Il numero ordine e SEMPRE quello dopo "Vs." (Vostro = ordine di Tirelli), mai quello dopo "Ns." (Nostro = ordine del fornitore)
2. OVC NON e mai il numero ordine Tirelli, ignoralo sempre
3. Il codice articolo e SEMPRE il nostro (Tirelli): cerca "Vs cod:", "Vs.Codice:", "Vs. Codice:", "Vs.Art.:", "Codice Cliente", "Vs. Codice Prodotto:", "Vostro Codice" nel testo della riga (in tutte le varianti con/senza punto, spazio, due punti), o la seconda riga in celle con doppio codice sovrapposto (es. colonna "CODICE ARTICOLO/VS CODICE") prima di usare il codice del fornitore. MAI usare Cod.Taric/HS Code come codice articolo.
4. Ignora righe vuote, totali, note e testi generici
5. Restituisci SOLO un array JSON valido, senza testo aggiuntivo, senza markdown

Esempio output corretto (G&G: "Vs. Ord. 202211 ... (OVC 2286)" -> ordine=202211, NON 2286):
[{"ddt_numero":"1196","ddt_data":"27/03/2026","fornitore":"G&G SERVICE SRL","ordine":"202211","codice":"D20036","descrizione":"GUIDA","um":"PZ","quantita":"1.0"},{"ddt_numero":"1196","ddt_data":"27/03/2026","fornitore":"G&G SERVICE SRL","ordine":"202841","codice":"D131379-L","descrizione":"PIASTRA","um":"PZ","quantita":"2.0"}]

Se la pagina non contiene righe articolo (es. e solo una pagina con timbri, firme o testo generico), restituisci: []"""

    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": img_base64
                        }
                    },
                    {
                        "type": "text",
                        "text": prompt
                    }
                ]
            }
        ]
    )

    testo = message.content[0].text.strip()

    # Rimuove eventuali blocchi markdown ```json ... ```
    if "```" in testo:
        inizio = testo.find("[")
        fine = testo.rfind("]") + 1
        if inizio >= 0 and fine > inizio:
            testo = testo[inizio:fine]

    articoli = json.loads(testo)

    # Sanitizza la quantita: punto e virgola sono sempre e solo separatori decimali
    for item in articoli:
        if "quantita" in item:
            q = str(item["quantita"]).strip().replace(",", ".")
            try:
                item["quantita"] = float(q)
            except ValueError:
                item["quantita"] = 0.0

    return articoli

def main():
    if len(sys.argv) < 3:
        print(json.dumps({"errore": "Uso: entrate_merci_ocr.py <pdf_path> <api_key>"}, ensure_ascii=False))
        sys.exit(1)

    pdf_path = sys.argv[1]
    api_key = sys.argv[2]

    if not os.path.exists(pdf_path):
        print(json.dumps({"errore": f"File non trovato: {pdf_path}"}, ensure_ascii=False))
        sys.exit(1)

    try:
        import fitz
    except ImportError:
        print(json.dumps({"errore": "Modulo PyMuPDF non installato. Eseguire: pip install pymupdf"}, ensure_ascii=False))
        sys.exit(1)

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
    except ImportError:
        print(json.dumps({"errore": "Modulo anthropic non installato. Eseguire: pip install anthropic"}, ensure_ascii=False))
        sys.exit(1)
    except Exception as e:
        print(json.dumps({"errore": f"Errore inizializzazione client Claude: {str(e)}"}, ensure_ascii=False))
        sys.exit(1)

    tutti_articoli = []
    errori_pagine = []

    try:
        doc = fitz.open(pdf_path)
        num_pagine = len(doc)

        for i in range(num_pagine):
            try:
                page = doc[i]
                img_base64 = converti_pagina_in_base64(page)
                articoli_pagina = analizza_pagina_con_claude(client, img_base64)
                tutti_articoli.extend(articoli_pagina)
            except json.JSONDecodeError as e:
                errori_pagine.append(f"Pagina {i+1}: risposta Claude non parsabile come JSON")
            except Exception as e:
                errori_pagine.append(f"Pagina {i+1}: {str(e)}")

        doc.close()
    except Exception as e:
        print(json.dumps({"errore": f"Errore apertura PDF: {str(e)}"}, ensure_ascii=False))
        sys.exit(1)

    # Carry-forward del numero ordine: se una riga non ha ordine,
    # usa l'ultimo ordine visto per lo stesso numero DDT
    last_ordine_per_ddt = {}
    for item in tutti_articoli:
        ddt_key = str(item.get("ddt_numero", "")).strip()
        ordine = str(item.get("ordine", "")).strip()
        if ordine:
            last_ordine_per_ddt[ddt_key] = ordine
        elif ddt_key in last_ordine_per_ddt:
            item["ordine"] = last_ordine_per_ddt[ddt_key]

    # Se ci sono errori ma anche articoli, li includo come warning
    risultato = tutti_articoli
    if errori_pagine and not tutti_articoli:
        print(json.dumps({"errore": "; ".join(errori_pagine)}, ensure_ascii=False))
        sys.exit(1)

    print(json.dumps(risultato, ensure_ascii=False))

if __name__ == "__main__":
    main()
