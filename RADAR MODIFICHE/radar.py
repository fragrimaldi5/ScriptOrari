import requests
from bs4 import BeautifulSoup
import hashlib
import time
import os
import re

# ================= CONFIGURAZIONE =================
# Se il file è locale e non hai un server attivo, usa il percorso del file
URL = "http://localhost:8000/indice.html" 
FILE_HASH = "orario_hash.txt"
INTERVALLO = 5 

def get_timetable_data(url):
    try:
        if url.startswith("http"):
            response = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=10)
            response.raise_for_status()
            html_content = response.text
        else:
            with open(url, "r", encoding="utf-8") as f:
                html_content = f.read()

        soup = BeautifulSoup(html_content, 'html.parser')
        grid = soup.find('table', class_='timegrid')
        
        if not grid:
            print("[DEBUG] Errore: Tabella 'timegrid' non trovata nell'HTML!")
            return None

        table_snapshot = ""
        lezione_trovata = False

        rows = grid.find_all('tr')
        for row_idx, row in enumerate(rows):
            cells = row.find_all(['td', 'th'])
            for col_idx, cell in enumerate(cells):
                
                # Cerchiamo il blocco lezione (EasyCourse usa ID parent_ o classi subject_pos)
                lesson_box = cell.find(['table', 'div'], id=re.compile(r'^parent_'))
                if not lesson_box:
                    # Se non c'è il parent, cerchiamo se la cella stessa ha classi di lezione
                    lesson_box = cell.find(class_=re.compile(r'subject_pos'))

                if lesson_box:
                    lezione_trovata = True
                    # Estraiamo i testi dei sottotag
                    materia = lesson_box.select_one('.subject_pos1')
                    docente = lesson_box.select_one('.subject_pos2')
                    aula = lesson_box.select_one('.subject_pos3')

                    m_txt = materia.get_text(strip=True) if materia else "MAT-VUOTA"
                    d_txt = docente.get_text(strip=True) if docente else "DOC-VUOTO"
                    a_txt = aula.get_text(strip=True) if aula else "AULA-VUOTA"

                    # Firma della cella: posizione + contenuto
                    table_snapshot += f"[{row_idx},{col_idx}]|{m_txt}|{d_txt}|{a_txt}###"
                else:
                    # Cella vuota: importante per rilevare cancellazioni o spostamenti
                    table_snapshot += f"[{row_idx},{col_idx}]|EMPTY###"

        if not lezione_trovata:
            print("[DEBUG] Attenzione: Tabella trovata ma nessuna lezione rilevata con le classi 'subject_pos'!")
            # Forniamo comunque l'hash della griglia vuota per monitorare se appare qualcosa
        
        return hashlib.sha256(table_snapshot.encode('utf-8')).hexdigest()

    except Exception as e:
        print(f"[DEBUG] Errore critico: {e}")
        return None

def monitor():
    print(f"--- Radar Modifiche EasyCourse Attivo ---")
    print(f"Target: {URL}\n")
    
    while True:
        new_hash = get_timetable_data(URL)
        
        if new_hash:
            if os.path.exists(FILE_HASH):
                with open(FILE_HASH, "r") as f:
                    old_hash = f.read()
                
                if new_hash != old_hash:
                    print(f"\a[@] MODIFICA RILEVATA! L'orario è cambiato: {time.ctime()}")
                    with open(FILE_HASH, "w") as f:
                        f.write(new_hash)
                else:
                    print(f"[-] Nessuna modifica... ({time.ctime()})")
            else:
                with open(FILE_HASH, "w") as f:
                    f.write(new_hash)
                print("[!] Prima scansione effettuata. Stato salvato.")
        
        time.sleep(INTERVALLO)

if __name__ == "__main__":
    try:
        monitor()
    except KeyboardInterrupt:
        print("\n[!] Chiusura radar.")