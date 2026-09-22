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

def find_new_link_via_ia(cliente, old_url, text_context=""):
    """Usa Google Gemini per analizzare il contesto o identificare nuovi riferimenti."""
    if not client_gemini:
        return "API Gemini non configurata (GEMINI_APP_KEY mancante)"
    
    prompt = (
        f"Sei un assistente legale/acquisti. Analizza la situazione per il cliente '{cliente}'.\n"
        f"L'URL registrato era: {old_url}.\n"
        f"Contesto o testo parziale rilevato: {text_context[:1000]}\n"
        f"Fornisci una breve indicazione su come reperire la nuova revisione o se il link appare errato."
    )
    try:
        response = client_gemini.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt,
        )
        return response.text.strip()
    except Exception as e:
        return f"Errore durante l'analisi Gemini: {str(e)}"

def send_email_report(anomalies, total_clienti):
    """Invia un'email di report al destinatario specificato nei Secrets."""
    if not all([EMAIL_MITTENTE, EMAIL_PASSWORD, EMAIL_DESTINATARIO]):
        print("[-] Credenziali Email non completamente configurate nei Secrets. Email non inviata.")
        return

    msg = MIMEMultipart()
    msg['From'] = EMAIL_MITTENTE
    msg['To'] = EMAIL_DESTINATARIO

    if anomalies:
        msg['Subject'] = f"⚠️ [ALERT] Variazioni/Anomalie Condizioni Fornitura ({len(anomalies)} rilevate)"
        
        body = f"<h2>Report Monitoraggio Condizioni Generali di Fornitura</h2>"
        body += f"<p>Sono state rilevate <b>{len(anomalies)}</b> anomalie su un totale di {total_clienti} clienti analizzati:</p><hr>"
        
        for a in anomalies:
            body += f"""
            <div style='margin-bottom: 15px; padding: 10px; border-left: 4px solid #e74c3c; background-color: #fdf2f2;'>
                <p><b>Cliente:</b> {a['cliente']}</p>
                <p><b>Anomalia / Errore:</b> {a['errore']}</p>
                <p><b>Revisione Attesa:</b> {a['revisione_attuale']} | <b>Data Attesa:</b> {a['data_attuale']}</p>
                <p><b>Link Registrato:</b> <a href='{a['link']}'>{a['link']}</a></p>
                <p><b>Analisi Agente AI:</b> {a['nuovo_link_suggerito']}</p>
            </div>
            """
    else:
        msg['Subject'] = f"✅ [OK] Monitoraggio Condizioni Fornitura - Nessuna Anomalia"
        body = f"""
        <h2>Report Monitoraggio Condizioni Generali di Fornitura</h2>
        <p>Tutti i <b>{total_clienti}</b> clienti sotto controllo risultano allineati con i link, le date e gli indici di revisione indicati nel file Excel.</p>
        <p><i>Nessuna azione richiesta.</i></p>
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
    anomalies = []
    total_clienti = len(df)

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
                "link": link,
                "nuovo_link_suggerito": nuovo_link
            })
            continue

        rev_found = False
        if revisione_attesa.upper() != 'NO INDEX':
            if re.search(r'\b(Rev|Revisione)?\s*' + re.escape(revisione_attesa) + r'\b', content, re.IGNORECASE):
                rev_found = True
        else:
            rev_found = True

        if not rev_found:
            print(f"[!] Attenzione: Revisione '{revisione_attesa}' non trovata nel documento linkato!")
            nuovo_link = find_new_link_via_ia(cliente, link, content[:1000])
            anomalies.append({
                "cliente": cliente,
                "errore": f"Revisione {revisione_attesa} non riscontrata sul sito",
                "revisione_attuale": revisione_attesa,
                "data_attuale": data_attesa,
                "link": link,
                "nuovo_link_suggerito": nuovo_link
            })
        else:
            print(f"[OK] Documento per {cliente} allineato.")

    # Invia sempre la mail di feedback
    send_email_report(anomalies, total_clienti)

if __name__ == "__main__":
    process_clienti()
