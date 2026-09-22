import os
import re
import requests
import pandas as pd
from bs4 import BeautifulSoup
from pypdf import PdfReader
from io import BytesIO
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google import genai

# Variabili di ambiente e Secrets
EXCEL_FILE = 'COND_CLIENTI.xlsx'
GEMINI_API_KEY = os.environ.get("GEMINI_APP_KEY")
EMAIL_MITTENTE = os.environ.get("EMAIL_MITTENTE")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
EMAIL_DESTINATARIO = os.environ.get("EMAIL_DESTINATARIO")

# Inizializzazione Client Gemini
client_gemini = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None

def fetch_content(url):
    """Scarica il contenuto del link (HTML o PDF)."""
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    try:
        response = requests.get(url, headers=headers, timeout=15)
        if response.status_code == 200:
            content_type = response.headers.get('Content-Type', '').lower()
            if 'application/pdf' in content_type or url.lower().endswith('.pdf'):
                pdf = PdfReader(BytesIO(response.content))
                text = " ".join([page.extract_text() or '' for page in pdf.pages])
                return text, "PDF"
            else:
                soup = BeautifulSoup(response.text, 'html.parser')
                return soup.get_text(separator=' '), "HTML"
        else:
            return None, f"HTTP Error {response.status_code}"
    except Exception as e:
        return None, str(e)

def analyze_with_gemini(cliente, url, text_content):
    """Analizza il testo estratto tramite Gemini per ricavare Revisione, Data e Note."""
    if not client_gemini:
        return "N/A", "N/A", "API Gemini non configurata (GEMINI_APP_KEY/GEMINI_API_KEY mancante)"

    prompt = (
        f"Sei un assistente legale specializzato nel monitoraggio contrattuale per il cliente '{cliente}'.\n"
        f"URL: {url}\n"
        f"Estratto testo del documento o pagina web:\n{text_content[:3000]}\n\n"
        "Esegui l'analisi ed estrai esattamente in questo formato:\n"
        "REVISIONE: <numero/indice di revisione trovato, oppure 'Non specificata'>\n"
        "DATA: <data pubblicazione/revisione trovata, oppure 'Non specificata'>\n"
        "NOTE: <breve sintesi dello stato e di eventuali anomalie o suggerimenti per il link>"
    )
    try:
        response = client_gemini.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
        )
        res_text = response.text.strip()

        rev_match = re.search(r"REVISIONE:\s*(.*)", res_text, re.IGNORECASE)
        data_match = re.search(r"DATA:\s*(.*)", res_text, re.IGNORECASE)
        note_match = re.search(r"NOTE:\s*(.*)", res_text, re.IGNORECASE)

        rev_found = rev_match.group(1).strip() if rev_match else "N/A"
        data_found = data_match.group(1).strip() if data_match else "N/A"
        note_found = note_match.group(1).strip() if note_match else res_text

        return rev_found, data_found, note_found
    except Exception as e:
        return "N/A", "N/A", f"Errore analisi Gemini: {str(e)}"

def send_email_table_report(results):
    """Genera e invia un'email contenente la tabella riepilogativa dell'analisi."""
    if not all([EMAIL_MITTENTE, EMAIL_PASSWORD, EMAIL_DESTINATARIO]):
        print("[-] Credenziali Email non completamente configurate nei Secrets. Email non inviata.")
        return

    anomalies_count = sum(1 for r in results if r['anomalia'])
    
    msg = MIMEMultipart()
    msg['From'] = EMAIL_MITTENTE
    msg['To'] = EMAIL_DESTINATARIO

    if anomalies_count > 0:
        msg['Subject'] = f"⚠️ [ALERT] Monitoraggio CGF - {anomalies_count} anomalie rilevate"
    else:
        msg['Subject'] = "✅ [OK] Monitoraggio Condizioni Generali di Fornitura - Allineato"

    body = f"""
    <h2>Report Monitoraggio Condizioni Generali di Fornitura</h2>
    <p>Di seguito la tabella di confronto tra i dati registrati nel file Excel e i riscontri rilevati dall'Agente AI:</p>
    
    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px;">
        <thead>
            <tr style="background-color: #2c3e50; color: #ffffff; text-align: left;">
                <th>CLIENTE</th>
                <th>REV. ATTESA</th>
                <th>DATA ATTESA</th>
                <th>LINK REGISTRATO</th>
                <th>REV. RILEVATA</th>
                <th>DATA RILEVATA</th>
                <th>NOTE / ANOMALIE</th>
            </tr>
        </thead>
        <tbody>
    """

    for r in results:
        bg_color = "#fdf2f2" if r['anomalia'] else "#f2f9f2"
        status_icon = "⚠️ " if r['anomalia'] else "✅ "

        body += f"""
        <tr style="background-color: {bg_color};">
            <td><b>{r['cliente']}</b></td>
            <td>{r['rev_attesa']}</td>
            <td>{r['data_attesa']}</td>
            <td><a href="{r['link']}" target="_blank">Apri Link</a></td>
            <td><b>{r['rev_rilevata']}</b></td>
            <td><b>{r['data_rilevata']}</b></td>
            <td>{status_icon}{r['note']}</td>
        </tr>
        """

    body += """
        </tbody>
    </table>
    <br>
    <p><i>Report generato automaticamente dal servizio AI di monitoraggio su GitHub.</i></p>
    """

    msg.attach(MIMEText(body, 'html'))

    try:
        print("[+] Connessione al server SMTP per invio mail...")
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_MITTENTE, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print(f"[OK] Email di report inviata con successo a: {EMAIL_DESTINATARIO}")
    except Exception as e:
        print(f"[-] Errore durante l'invio dell'email SMTP: {str(e)}")

def process_clienti():
    df = pd.read_excel(EXCEL_FILE)
    results = []

    for index, row in df.iterrows():
        cliente = str(row['CLIENTE']).strip()
        revisione_attesa = str(row['REVISIONE']).strip()
        data_attesa = str(row['DATA']).strip()
        link = str(row['LINK']).strip()

        print(f"\n[+] Verifica in corso per: {cliente}")
        content, status = fetch_content(link)

        if not content:
            print(f"[-] Link non raggiungibile ({status}).")
            results.append({
                'cliente': cliente,
                'rev_attesa': revisione_attesa,
                'data_attesa': data_attesa,
                'link': link,
                'rev_rilevata': 'N/A',
                'data_rilevata': 'N/A',
                'note': f"Link non raggiungibile / Errore ({status}). Necessaria verifica manuale del nuovo URL.",
                'anomalia': True
            })
            continue

        rev_found, data_found, note = analyze_with_gemini(cliente, link, content)

        # Controllo anomalie tra revisione attesa e rilevata
        anomalia = False
        if revisione_attesa.upper() != 'NO INDEX':
            if revisione_attesa.lower() not in rev_found.lower() and rev_found.upper() != 'N/A':
                anomalia = True
                note = f"Discrepanza revisione! Attesa '{revisione_attesa}', trovata '{rev_found}'. " + note
        
        if "errore" in note.lower() or "discrepanza" in note.lower() or "mancante" in note.lower():
            anomalia = True

        results.append({
            'cliente': cliente,
            'rev_attesa': revisione_attesa,
            'data_attesa': data_attesa,
            'link': link,
            'rev_rilevata': rev_found,
            'data_rilevata': data_found,
            'note': note,
            'anomalia': anomalia
        })

    # Invia l'email formattata a tabella
    send_email_table_report(results)

if __name__ == "__main__":
    process_clienti()
