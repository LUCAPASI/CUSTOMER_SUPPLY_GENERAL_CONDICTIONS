import os
import re
import requests
import pandas as pd
from bs4 import BeautifulSoup
from pypdf import PdfReader
from io import BytesIO
from urllib.parse import urljoin
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

def get_session():
    """Crea una sessione HTTP con headers browser standard e bypass controlli di sicurezza base."""
    session = requests.Session()
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
        'Accept-Language': 'it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7'
    })
    return session

def extract_text_from_url(url, session, depth=0):
    """Scarica il contenuto del link. Se trova pagine HTML, analizza anche i link PDF interni."""
    if depth > 1:
        return "", "Max Depth Reached"

    try:
        response = session.get(url, timeout=20, verify=False)
        if response.status_code != 200:
            return "", f"HTTP Error {response.status_code}"

        content_type = response.headers.get('Content-Type', '').lower()
        
        # Caso 1: È un file PDF
        if 'application/pdf' in content_type or url.lower().endswith('.pdf'):
            pdf = PdfReader(BytesIO(response.content))
            text = " ".join([page.extract_text() or '' for page in pdf.pages])
            return text, "PDF"

        # Caso 2: È una pagina HTML
        soup = BeautifulSoup(response.text, 'html.parser')
        main_text = soup.get_text(separator=' ')

        # Cerca eventuali sub-link PDF pertinenti nella pagina (es. "Condizioni", "Fornitura", "Terms", "Purchase")
        pdf_texts = []
        for a_tag in soup.find_all('a', href=True):
            href = a_tag['href']
            link_text = a_tag.get_text().strip().lower()
            if href.lower().endswith('.pdf') or 'condizion' in link_text or 'terms' in link_text or 'purchase' in link_text:
                full_url = urljoin(url, href)
                sub_text, sub_type = extract_text_from_url(full_url, session, depth=depth+1)
                if sub_text:
                    pdf_texts.append(f"\n--- CONTENUTO DA LINK INTERNO ({full_url}) ---\n" + sub_text)

        combined_text = main_text + "\n" + "\n".join(pdf_texts)
        return combined_text, "HTML"

    except Exception as e:
        return "", str(e)

def analyze_with_gemini(cliente, url, text_content):
    """Analizza il testo estratto tramite Gemini (modello aggiornato) per ricavare Revisione, Data e Note."""
    if not client_gemini:
        return "N/A", "N/A", "API Gemini non configurata (GEMINI_APP_KEY mancante)"

    prompt = (
        f"Sei un assistente legale/acquisti. Analizza il seguente testo estratto dal sito/documento per il cliente '{cliente}' (URL: {url}).\n\n"
        f"TESTO ESTRATTO:\n{text_content[:4000]}\n\n"
        "Compito:\n"
        "1. Identifica se sono presenti le Condizioni Generali di Fornitura / Acquisto.\n"
        "2. Estrai il numero o l'indice di Revisione (es. Rev. 1, Rev. C, 0, ecc.). Se non presente, scrivi 'NO INDEX'.\n"
        "3. Estrai la Data delle condizioni/revisione (es. 27/04/2026, Gennaio 2024, ecc.). Se non presente, scrivi 'Non specificata'.\n"
        "4. Fornisci un'analisi sintetica in NOTE.\n\n"
        "Rispondi ESATTAMENTE con questa struttura:\n"
        "REVISIONE: <valore>\n"
        "DATA: <valore>\n"
        "NOTE: <breve nota>"
    )
    
    try:
        # Usa il nome modello 'gemini-2.5-flash' o l'alias standard
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
    """Invia un'email contenente la tabella riepilogativa dell'analisi."""
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
    <p>Di seguito la tabella di confronto tra i dati registrati nel file Excel e i riscontri rilevati dall'Agente AI su web/PDF:</p>
    
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
    # Disabilita gli avvisi SSL per siti con certificati non conformi
    requests.packages.urllib3.disable_warnings()
    
    df = pd.read_excel(EXCEL_FILE)
    results = []
    session = get_session()

    for index, row in df.iterrows():
        cliente = str(row['CLIENTE']).strip()
        revisione_attesa = str(row['REVISIONE']).strip()
        data_attesa = str(row['DATA']).strip()
        link = str(row['LINK']).strip()

        print(f"\n[+] Verifica in corso per: {cliente}")
        content, status = extract_text_from_url(link, session)

        if not content:
            print(f"[-] Link non raggiungibile ({status}).")
            results.append({
                'cliente': cliente,
                'rev_attesa': revisione_attesa,
                'data_attesa': data_attesa,
                'link': link,
                'rev_rilevata': 'N/A',
                'data_rilevata': 'N/A',
                'note': f"Link non raggiungibile ({status}). Necessaria verifica del nuovo URL sul sito del cliente.",
                'anomalia': True
            })
            continue

        rev_found, data_found, note = analyze_with_gemini(cliente, link, content)

        # Verifica discrepanza
        anomalia = False
        if revisione_attesa.upper() != 'NO INDEX':
            if revisione_attesa.lower() not in rev_found.lower() and rev_found.upper() != 'N/A':
                anomalia = True

        if "errore" in note.lower() or "non raggiungibile" in note.lower() or "discrepanza" in note.lower():
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

    send_email_table_report(results)

if __name__ == "__main__":
    process_clienti()
