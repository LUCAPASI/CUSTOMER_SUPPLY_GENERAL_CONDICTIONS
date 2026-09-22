import os
import requests
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configurazioni e Secrets
EXCEL_FILE = 'COND_CLIENTI.xlsx'
EMAIL_MITTENTE = os.environ.get("EMAIL_MITTENTE")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD")
EMAIL_DESTINATARIO = os.environ.get("EMAIL_DESTINATARIO")

def get_session():
    """Crea una sessione HTTP con Header da reale browser per evitare blocchi anti-bot."""
    session = requests.Session()
    session.headers.update({
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'it-IT,it;q=0.9,en-US;q=0.8,en;q=0.7'
    })
    return session

def check_link_status(url, session):
    """Verifica se il link esiste e restituisce codice HTTP 200."""
    try:
        response = session.get(url, timeout=15, verify=False, stream=True)
        return response.status_code == 200
    except Exception:
        return False

def send_email_report(results):
    """Invia il report email riassuntivo in formato tabella HTML."""
    if not all([EMAIL_MITTENTE, EMAIL_PASSWORD, EMAIL_DESTINATARIO]):
        print("[-] Credenziali email non configurate nei Secrets. Invio saltato.")
        return

    aggiornamenti_count = sum(1 for r in results if r['aggiornamento'])

    msg = MIMEMultipart()
    msg['From'] = EMAIL_MITTENTE
    msg['To'] = EMAIL_DESTINATARIO

    if aggiornamenti_count > 0:
        msg['Subject'] = f"⚠️ [ALERT] CGF Clienti - {aggiornamenti_count} Possibili Aggiornamenti Rilevati"
    else:
        msg['Subject'] = "✅ [OK] Monitoraggio CGF Clienti - Tutti i link sono Verificati e Invariati"

    body = f"""
    <h2>Report Monitoraggio Condizioni Generali di Fornitura</h2>
    <p>Di seguito l'esito della verifica delle condizioni generali di fornitura:</p>

    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-family: Arial, sans-serif; font-size: 13px;">
        <thead>
            <tr style="background-color: #2c3e50; color: #ffffff; text-align: left;">
                <th>CLIENTE</th>
                <th>REVISIONE EXCEL</th>
                <th>DATA EXCEL</th>
                <th>LINK VERIFICATO</th>
                <th>STATO</th>
            </tr>
        </thead>
        <tbody>
    """

    for r in results:
        bg_color = "#fdf2f2" if r['aggiornamento'] else "#f2f9f2"
        status_label = "⚠️ RILEVATO AGGIORNAMENTO" if r['aggiornamento'] else "✅ INVARIATO"

        body += f"""
        <tr style="background-color: {bg_color};">
            <td><b>{r['cliente']}</b></td>
            <td>{r['revisione']}</td>
            <td>{r['data']}</td>
            <td><a href="{r['link']}" target="_blank">Apri Link Documento</a></td>
            <td><b>{status_label}</b></td>
        </tr>
        """

    body += """
        </tbody>
    </table>
    <br>
    <p><i>Nota: Se lo stato indica 'RILEVATO AGGIORNAMENTO', il vecchio link non risponde più (es. 404), indicando che il cliente potrebbe aver pubblicato una nuova revisione con un URL diverso.</i></p>
    """

    msg.attach(MIMEText(body, 'html'))

    try:
        print("[+] Inizio invio email SMTP...")
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_MITTENTE, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print(f"[OK] Email inviata con successo a: {EMAIL_DESTINATARIO}")
    except Exception as e:
        print(f"[-] Errore invio email SMTP: {str(e)}")

def main():
    requests.packages.urllib3.disable_warnings()

    df = pd.read_excel(EXCEL_FILE)
    results = []
    session = get_session()

    for index, row in df.iterrows():
        cliente = str(row['CLIENTE']).strip()
        revisione = str(row['REVISIONE']).strip()
        data_doc = str(row['DATA']).strip()
        link = str(row['LINK']).strip()

        print(f"[+] Verifica link per {cliente}...")
        is_ok = check_link_status(link, session)

        # Se il link NON existe (is_ok == False) -> C'è stato un aggiornamento
        aggiornamento_rilevato = not is_ok

        results.append({
            'cliente': cliente,
            'revisione': revisione,
            'data': data_doc,
            'link': link,
            'aggiornamento': aggiornamento_rilevato
        })

    send_email_report(results)

if __name__ == "__main__":
    main()
