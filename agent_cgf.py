import os
import re
import requests
import pandas as pd
from bs4 import BeautifulSoup
from pypdf import PdfReader
from io import BytesIO
from datetime import datetime

# File di input
EXCEL_FILE = 'COND_CLIENTI.xlsx'

def fetch_content(url):
    """Scarica il contenuto del link (HTML o PDF)."""
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code == 200:
            content_type = response.headers.get('Content-Type', '').lower()
            if 'application/pdf' in content_type or url.endswith('.pdf'):
                # Estrazione testo da PDF
                pdf = PdfReader(BytesIO(response.content))
                text = " ".join([page.extract_text() or '' for page in pdf.pages])
                return text, "PDF"
            else:
                # Estrazione testo da HTML
                soup = BeautifulSoup(response.text, 'html.parser')
                return soup.get_text(separator=' '), "HTML"
        else:
            return None, f"HTTP Error {response.status_code}"
    except Exception as e:
        return None, str(e)

def find_new_link_via_ia(cliente, old_url, text_context=""):
    """
    Funzione di fallback: Se il link non è più valido o la revisione è cambiata,
    questa funzione simula/esegue la ricerca della nuova pagina ufficiale.
    """
    # Integrazione API di ricerca / Modello LLM per individuare la nuova URL
    prompt = f"Trova la pagina ufficiale Condizioni Generali di Fornitura per {cliente}."
    # Esempio di valore restituito dall'Agente IA
    return f"Nessun link automatico generato - Richiesta verifica manuale per {cliente}"

def process_clienti():
    df = pd.read_excel(EXCEL_FILE)
    anomalies = []

    for index, row in df.iterrows():
        cliente = str(row['CLIENTE']).strip()
        revisione_attesa = str(row['REVISIONE']).strip()
        data_attesa = str(row['DATA']).strip()
        link = str(row['LINK']).strip()

        print(f"\n[+] Verifica in corso per: {cliente}")
        content, status = fetch_content(link)

        if not content:
            print(f"[-] Link non raggiungibile ({status}). Avvio ricerca nuovo link...")
            nuovo_link = find_new_link_via_ia(cliente, link)
            anomalies.append({
                "cliente": cliente,
                "errore": f"Link non valido / 404 ({status})",
                "revisione_attuale": revisione_attesa,
                "data_attuale": data_attesa,
                "nuovo_link_suggerito": nuovo_link
            })
            continue

        # Verifica presenza della Revisione o della Data all'interno del contenuto estratto
        rev_found = False
        data_found = False

        if revisione_attesa.upper() != 'NO INDEX':
            # Cerca il pattern della revisione (es. "Rev. C", "Revisione 1", etc.)
            if re.search(r'\b(Rev|Revisione)?\s*' + re.escape(revisione_attesa) + r'\b', content, re.IGNORECASE):
                rev_found = True
        else:
            rev_found = True # Ignorato se NO INDEX

        # Verifica presenza data nel testo
        if data_attesa != 'nan' and data_attesa in content:
            data_found = True

        # Se non si trovano i riferimenti attesi, segnala la discrepanza
        if not rev_found:
            print(f"[!] Attenzione: Revisione '{revisione_attesa}' non trovata nel documento linkato!")
            nuovo_link = find_new_link_via_ia(cliente, link, content[:500])
            anomalies.append({
                "cliente": cliente,
                "errore": f"Revisione {revisione_attesa} non riscontrata sul sito",
                "revisione_attuale": revisione_attesa,
                "data_attuale": data_attesa,
                "nuovo_link_suggerito": nuovo_link
            })
        else:
            print(f"[OK] Documento per {cliente} allineato.")

    # Se ci sono anomalie, crea un Report / GitHub Issue
    if anomalies:
        generate_github_issue(anomalies)

def generate_github_issue(anomalies):
    print("\n--- GENERAZIONE REPORT ANOMALIE ---")
    for a in anomalies:
        print(f"Cliente: {a['cliente']} | Errore: {a['errore']} | Suggerimento: {a['nuovo_link_suggerito']}")

if __name__ == "__main__":
    process_clienti()
