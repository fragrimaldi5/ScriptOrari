import os
import pandas as pd

# ================= CONFIGURAZIONE =================
# Cartella che contiene i file orari (entrerà anche nelle sottocartelle)
INPUT_DIR = "input_excel" 
OUTPUT_FILE = "Orario_per_Aule_Completo.xlsx"

GIORNI = ["lunedì", "martedì", "mercoledì", "giovedì", "venerdì"]
ORARI = [
    "08:30-09:30", "09:30-10:30", "10:30-11:30", "11:30-12:30", 
    "12:30-13:30", "13:30-14:30", "14:30-15:30", "15:30-16:30", 
    "16:30-17:30", "17:30-18:30"
]

def trova_tabella_dati(df_raw):
    """
    Cerca la riga di intestazione nel foglio 'Lista'.
    Restituisce il DataFrame con le colonne corrette partendo dalla riga trovata.
    """
    for i, row in df_raw.iterrows():
        # Converte i valori della riga in stringhe minuscole per il confronto
        row_values = [str(val).strip().lower() for val in row.values]
        
        # Cerchiamo la riga che contiene i capisaldi della tabella
        if 'giorno' in row_values and 'ora' in row_values and 'aula' in row_values:
            # Abbiamo trovato la riga degli header
            headers = [str(h).strip() for h in row.values]
            df_pulito = df_raw.iloc[i+1:].copy()
            df_pulito.columns = headers
            return df_pulito
    return None

def genera_orario():
    # Struttura: {giorno: {ora: {aula: nome_corso}}}
    database = {g: {o: {} for o in ORARI} for g in GIORNI}
    tutte_le_aule = set()

    print(f"Inizio scansione cartella: {INPUT_DIR}")

    for root, dirs, files in os.walk(INPUT_DIR):
        for file in files:
            if file.lower().endswith(('.xls', '.xlsx')):
                path_completo = os.path.join(root, file)
                
                try:
                    df_raw = pd.read_excel(path_completo, sheet_name="Lista", header=None)
                    df = trova_tabella_dati(df_raw)
                    
                    if df is None:
                        continue

                    for _, row in df.iterrows():
                        giorno_raw = str(row.get('Giorno', '')).lower().strip()
                        ora_raw = str(row.get('Ora', '')).strip()
                        aula_raw = str(row.get('Aula', '')).strip()
                        
                        # --- MODIFICA QUI ---
                        # Prendiamo il valore, lo convertiamo in stringa e 
                        # teniamo solo la prima riga prima del carattere \n
                        cella_intero = str(row.get('Nome insegnamento', ''))
                        nome_corso = cella_intero.split('\n')[0].strip()
                        # --------------------

                        if giorno_raw in database and ora_raw in ORARI:
                            if aula_raw not in ["nan", "", "None"]:
                                tutte_le_aule.add(aula_raw)
                                
                                if aula_raw not in database[giorno_raw][ora_raw]:
                                    database[giorno_raw][ora_raw][aula_raw] = nome_corso
                                else:
                                    corso_esistente = database[giorno_raw][ora_raw][aula_raw]
                                    if corso_esistente != nome_corso:
                                        print(f"  [Conflitto] {aula_raw} il {giorno_raw} alle {ora_raw}: {nome_corso} vs {corso_esistente}")

                except Exception as e:
                    print(f"  [ERRORE] Impossibile leggere {file}: {e}")

    # 2. Scrittura del file Excel finale
    if not tutte_le_aule:
        print("\nATTENZIONE: Nessun dato trovato. Controlla che i file abbiano un foglio chiamato 'Lista'.")
        return

    # Ordiniamo le aule alfabeticamente per le colonne
    aule_ordinate = sorted(list(tutte_le_aule))
    
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for giorno in GIORNI:
            # Creiamo il DataFrame per ogni foglio
            df_giorno = pd.DataFrame(index=ORARI, columns=aule_ordinate)
            
            for ora in ORARI:
                for aula in aule_ordinate:
                    # Recupera il nome originale del corso o lascia vuoto
                    df_giorno.at[ora, aula] = database[giorno][ora].get(aula, "")
            
            # Scrittura nel foglio (es. Lunedì, Martedì...)
            df_giorno.to_excel(writer, sheet_name=giorno.capitalize())

    print(f"\n{'-'*50}")
    print(f"COMPLETATO!")
    print(f"File generato: {OUTPUT_FILE}")
    print(f"File totali censiti: {len(aule_ordinate)}")
    print(f"{'-'*50}")

if __name__ == "__main__":
    # Verifica esistenza cartella input
    if not os.path.exists(INPUT_DIR):
        os.makedirs(INPUT_DIR)
        print(f"Creata cartella '{INPUT_DIR}'. Inserisci i file Excel e riavvia lo script.")
    else:
        genera_orario()