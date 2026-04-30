import os
import pandas as pd

# ================= CONFIGURAZIONE =================
INPUT_DIR = "input_excel" 
OUTPUT_FILE = "Orario_per_Aule_Completo.xlsx"

GIORNI = ["lunedì", "martedì", "mercoledì", "giovedì", "venerdì"]
ORARI = [
    "08:30-09:30", "09:30-10:30", "10:30-11:30", "11:30-12:30", 
    "12:30-13:30", "13:30-14:30", "14:30-15:30", "15:30-16:30", 
    "16:30-17:30", "17:30-18:30"
]

def trova_tabella_dati(df_raw):
    for i, row in df_raw.iterrows():
        row_values = [str(val).strip().lower() for val in row.values]
        if 'giorno' in row_values and 'ora' in row_values and 'aula' in row_values:
            headers = [str(h).strip() for h in row.values]
            df_pulito = df_raw.iloc[i+1:].copy()
            df_pulito.columns = headers
            return df_pulito
    return None

def genera_orario():
    # Struttura modificata: {giorno: {ora: {aula: (nome_corso, nome_file)}}}
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
                        
                        cella_intero = str(row.get('Nome insegnamento', ''))
                        if cella_intero.lower() in ['nan', 'none', '']:
                            continue
                            
                        nome_corso = cella_intero.split('\n')[0].strip()

                        if giorno_raw in database and ora_raw in ORARI:
                            if aula_raw not in ["nan", "", "None"]:
                                tutte_le_aule.add(aula_raw)
                                
                                # Verifichiamo se l'aula è già occupata
                                if aula_raw not in database[giorno_raw][ora_raw]:
                                    # Salviamo una tupla con (Nome Corso, Nome File)
                                    database[giorno_raw][ora_raw][aula_raw] = (nome_corso, file)
                                else:
                                    corso_esistente, file_esistente = database[giorno_raw][ora_raw][aula_raw]
                                    
                                    if corso_esistente != nome_corso:
                                        print(f"\n⚠️ [CONFLITTO] Aula: {aula_raw} | {giorno_raw} {ora_raw}")
                                        print(f"   - GIÀ OCCUPATA DA: '{corso_esistente}' (File: {file_esistente})")
                                        print(f"   - TENTATIVO DI:    '{nome_corso}' (File: {file})")

                except Exception as e:
                    print(f"   [ERRORE] Impossibile leggere {file}: {e}")

    if not tutte_le_aule:
        print("\nATTENZIONE: Nessun dato trovato.")
        return

    aule_ordinate = sorted(list(tutte_le_aule))
    
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        for giorno in GIORNI:
            df_giorno = pd.DataFrame(index=ORARI, columns=aule_ordinate)
            for ora in ORARI:
                for aula in aule_ordinate:
                    # Estraiamo solo il nome del corso (indice 0 della tupla)
                    dati = database[giorno][ora].get(aula)
                    df_giorno.at[ora, aula] = dati[0] if dati else ""
            
            df_giorno.to_excel(writer, sheet_name=giorno.capitalize())

    print(f"\n{'-'*50}")
    print(f"COMPLETATO!")
    print(f"File generato: {OUTPUT_FILE}")
    print(f"{'-'*50}")

if __name__ == "__main__":
    if not os.path.exists(INPUT_DIR):
        os.makedirs(INPUT_DIR)
        print(f"Creata cartella '{INPUT_DIR}'.")
    else:
        genera_orario()