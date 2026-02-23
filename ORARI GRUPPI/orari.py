import os
import pandas as pd
import re
import shutil
from PIL import Image, ImageDraw, ImageFont

# ================= CONFIGURAZIONE =================
INPUT_DIR = "input_excel"
OUTPUT_DIR = "output_img"
LOGO_PATH = "logo_studentingegneria.png"
GIORNI_SETTIMANA = ["Lunedì", "Martedì", "Mercoledì", "Giovedì", "Venerdì"]

SCALE = 3
BASE_CELL_WIDTH = int(460 * SCALE)
BASE_CELL_HEIGHT = int(260 * SCALE)
MARGIN = int(50 * SCALE)
BORDER_WIDTH = 5 * SCALE 

PADDING_STD = 0.07 
PADDING_MULTI = 0.07

FONT_BOLD = os.path.join("fonts", "arialbd.ttf")
FONT_REG = os.path.join("fonts", "arial.ttf")

SIZE_HEADER = int(48 * SCALE)
SIZE_ORARIO = int(55 * SCALE)
SIZE_CORSO_MAX, SIZE_CORSO_MIN = int(55 * SCALE), int(38 * SCALE)
SIZE_AULA_MAX, SIZE_AULA_MIN = int(45 * SCALE), int(36 * SCALE)
SIZE_MULTI_CORSO_MAX, SIZE_MULTI_CORSO_MIN = int(45 * SCALE), int(30 * SCALE)
SIZE_MULTI_AULA_MAX, SIZE_MULTI_AULA_MIN = int(38 * SCALE), int(5 * SCALE)

COLORS = [
    (240, 128, 128), # Corallo chiaro (ex rosso)
    (144, 238, 144), # Verde chiaro
    (150, 190, 230), # Carta da zucchero (ex azzurro)
    (245, 245, 150), # Giallo paglierino
    (245, 190, 130), # Pesce arancio
    (180, 160, 230), # Lavanda (ex viola)
    (150, 225, 225), # Turchese pallido
    (230, 160, 210), # Rosa antico
    (190, 220, 160), # Salvia (ex verde acido)
    (235, 200, 170), # Carne/Beige
    (180, 230, 150), # Verde mela soft
    (160, 230, 200), # Menta fredda
    (230, 170, 230), # Lilla chiaro
    (245, 235, 160), # Crema/Oro pallido
    (160, 180, 240)  # Perwinkle (Blu-viola soft)
]

# ================= FUNZIONI DI SUPPORTO =================

def svuota_cartella(directory):
    if os.path.exists(directory):
        for elemento in os.listdir(directory):
            percorso_elemento = os.path.join(directory, elemento)
            try:
                if os.path.isfile(percorso_elemento) or os.path.islink(percorso_elemento):
                    os.unlink(percorso_elemento)
                elif os.path.isdir(percorso_elemento):
                    shutil.rmtree(percorso_elemento)
            except Exception as e:
                print(f"Errore eliminazione {percorso_elemento}: {e}")

def clean_string(text):
    if not text or str(text).lower() == "nan": return ""
    text = re.sub(r'[,\-\*]', ' ', str(text))
    text = re.sub(r'[^a-zA-Z0-9\s\.]', '', text)
    return " ".join(text.split()).upper()

def get_font(path, size):
    try: return ImageFont.truetype(path, int(size))
    except: return ImageFont.load_default()

def dividi_greedy(parole, draw, font, max_w):
    righe, riga_corr = [], []
    for p in parole:
        test = " ".join(riga_corr + [p]) if riga_corr else p
        if draw.textlength(test, font=font) <= max_w:
            riga_corr.append(p)
        else:
            if riga_corr: righe.append(" ".join(riga_corr))
            riga_corr = [p]
    if riga_corr: righe.append(" ".join(riga_corr))
    return righe

def abbrevia_riga_mirata(riga, draw, font, max_w):
    parole = riga.split()
    while draw.textlength(" ".join(parole), font=font) > max_w:
        lunghezze = [len(p.replace(".", "")) for p in parole]
        if not lunghezze or max(lunghezze) <= 3: break 
        idx = lunghezze.index(max(lunghezze))
        parola_da_ridurre = parole[idx].replace(".", "")
        parole[idx] = parola_da_ridurre[:-1] + "."
    return " ".join(parole)

def estrai_dati_cella(cella):
    val = str(cella).strip()
    if any(g.lower() in val.lower() for g in GIORNI_SETTIMANA) or val.lower() == "nan" or not val:
        return []
    blocchi = [b.strip() for b in val.split("--------------") if b.strip()] if "--------------" in val else [val]
    risultati = []
    for blocco in blocchi:
        righe = [r.strip() for r in blocco.splitlines() if r.strip()]
        if not righe: continue
        gruppo = "AH" if "(A-H)" in blocco else "IZ" if "(I-Z)" in blocco else None
        corso = clean_string(righe[0].replace("(A-H)", "").replace("(I-Z)", ""))
        aula_raw = righe[-1].upper() if len(righe) > 1 else ""
        
        # Sostituzioni speciali
        if any(t in aula_raw for t in ["SPAZIO PER ATTIVITÀ COMPLEMENTARI", "SPAZIO PER ATTIVITA COMPLEMENTARI", "SPAZIO PER ATTIVITA' COMPLEMENTARI"]):
            aula_raw = f"AULA {aula_raw.split('_')[-1]}" if "_" in aula_raw else "AULA"
        elif any(t in aula_raw for t in ["LAB DI ING SANITARIA", "LAB. DI ING SANITARIA", "LAB. DI ING. SANITARIA"]):
            aula_raw = "LAB. ING. SANITARIA"
        elif "CENTRO LINGUISTICO DI ATENEO" in aula_raw: aula_raw = "CLA"
        elif "DIDATTICA A DISTANZA" in aula_raw: aula_raw = "DAD"
        if any(t in aula_raw for t in ["LAB STRUTTURE", "LAB. STRUTTURE"]): aula_raw = "LAB. STRUTTURE"

        parole_aula = aula_raw.split()[:3]
        if parole_aula and parole_aula[0] == "LABORATORIO": parole_aula[0] = "LAB."
        aula = clean_string(" ".join(parole_aula))
        if "LAB " in aula or aula == "LAB": aula = aula.replace("LAB", "LAB.")
        risultati.append({"corso": corso, "aula": aula, "gruppo": gruppo})
    return risultati

def formatta_blocco_definitivo(testo, aula, draw, max_w, max_h, is_multi):
    # Font limits
    c_max, c_min = (SIZE_MULTI_CORSO_MAX, SIZE_MULTI_CORSO_MIN) if is_multi else (SIZE_CORSO_MAX, SIZE_CORSO_MIN)
    a_max, a_min = (SIZE_MULTI_AULA_MAX, SIZE_MULTI_AULA_MIN) if is_multi else (SIZE_AULA_MAX, SIZE_AULA_MIN)
    a_mid = (a_max + a_min) // 2
    
    # --- GESTIONE CORSO (Invariata) ---
    parole_c = testo.split()
    final_righe_c, final_font_c = [], None
    for fs in range(c_max, c_min - 1, -2):
        f_test = get_font(FONT_BOLD, fs)
        tentativo = dividi_greedy(parole_c, draw, f_test, max_w)
        if any(draw.textlength(p, font=f_test) > max_w for p in parole_c): continue
        h_stimata = len(tentativo) * fs * 1.15
        if len(tentativo) <= 3 and h_stimata < (max_h * 0.65):
            final_righe_c, final_font_c = tentativo, f_test
            break
    if not final_righe_c:
        final_font_c = get_font(FONT_BOLD, c_min)
        k = (len(parole_c) + 2) // 3
        raw_righe = [" ".join(parole_c[i:i+k]) for i in range(0, len(parole_c), k)][:3]
        final_righe_c = [abbrevia_riga_mirata(r, draw, final_font_c, max_w) for r in raw_righe]

    # --- GESTIONE AULA (LOGICA RICHIESTA: NO ABBREVIAZIONI) ---
    parole_a = aula.split()
    final_righe_a = [aula]
    final_font_a = get_font(FONT_REG, a_max)
    success = False

    # STEP 1: Rimpicciolimento granulare (-1) fino a MID su riga SINGOLA
    for fs in range(a_max, a_mid - 1, -1):
        f_test = get_font(FONT_REG, fs)
        if draw.textlength(aula, font=f_test) <= max_w:
            final_font_a, final_righe_a = f_test, [aula]
            success = True
            break

    # STEP 2: Se non entra, proviamo a dividere in DUE RIGHE
    # Testiamo ogni font (da MAX a MIN) e per ogni font cerchiamo la divisione migliore
    if not success:
        # Partiamo da MAX per avere il font più grande possibile su due righe
        for fs in range(a_max, a_min - 1, -1):
            f_test = get_font(FONT_REG, fs)
            # Proviamo a spostare l'ultima parola, poi le ultime due, ecc.
            for i in range(len(parole_a) - 1, 0, -1):
                r1, r2 = " ".join(parole_a[:i]), " ".join(parole_a[i:])
                # Se entrambe le righe intere entrano nel max_w, abbiamo vinto
                if draw.textlength(r1, font=f_test) <= max_w and draw.textlength(r2, font=f_test) <= max_w:
                    final_font_a, final_righe_a = f_test, [r1, r2]
                    success = True
                    break
            if success: break

    # STEP 3: Fallback (se ancora non entra, usa il font minimo e due righe)
    if not success:
        final_font_a = get_font(FONT_REG, a_min)
        if len(parole_a) > 1:
            split_point = max(1, len(parole_a) // 2)
            final_righe_a = [" ".join(parole_a[:split_point]), " ".join(parole_a[split_point:])]
        else:
            final_righe_a = [aula]

    return final_righe_c, final_font_c, final_righe_a, final_font_a

# ================= CORE GRAFICA =================

def genera_immagine_quadrata(excel_path, output_subfolder, filtro_gruppo=None):
    df = pd.read_excel(excel_path, header=None)
    try: 
        start_row = next(i for i, row in df.iterrows() if 'lunedì' in str(row.values).lower())
    except: 
        return
        
    df_tabella = df.iloc[start_row+1:].copy()
    df_tabella = df_tabella[df_tabella.iloc[:, 0].astype(str).str.contains('-', na=False)]
    orari, matrice_dati = df_tabella.iloc[:, 0].values, df_tabella.iloc[:, 1:6]
    
    table_w, table_h = (len(GIORNI_SETTIMANA) + 1) * BASE_CELL_WIDTH, (len(orari) + 1) * BASE_CELL_HEIGHT
    logo_area_h = int(220 * SCALE)
    content_w, content_h = table_w + MARGIN * 2, table_h + MARGIN * 2 + logo_area_h
    lato_quadrato = max(content_w, content_h)
    
    img = Image.new("RGB", (lato_quadrato, lato_quadrato), "white")
    draw = ImageDraw.Draw(img)
    offset_x, offset_y = (lato_quadrato - content_w) // 2, (lato_quadrato - content_h) // 2
    
    color_map, next_color = {}, 0
    for r_idx in range(len(orari)):
        for c_idx in range(len(GIORNI_SETTIMANA)):
            # Estrazione e filtraggio corsi
            corsi_raw = estrai_dati_cella(matrice_dati.iloc[r_idx, c_idx])
            corsi = [c for c in corsi_raw if not filtro_gruppo or not c['gruppo'] or c['gruppo'] == filtro_gruppo]
            if not corsi: continue
            
            # --- LOGICA ORDINAMENTO INTELLIGENTE ---
            # Ordiniamo i corsi per lunghezza del nome: i più corti prima, il più lungo per ultimo
            corsi_ordinati = sorted(corsi, key=lambda x: len(x['corso']))
            num = len(corsi_ordinati)
            
            xb, yb = offset_x + MARGIN + (c_idx+1) * BASE_CELL_WIDTH, offset_y + MARGIN + (r_idx+1) * BASE_CELL_HEIGHT
            w, h = BASE_CELL_WIDTH, BASE_CELL_HEIGHT
            
            # Definizione geometrica delle sotto-celle
            sotto_celle = []
            if num == 1:
                sotto_celle = [(xb, yb, w, h)]
            elif num == 2:
                sotto_celle = [(xb, yb, w//2, h), (xb + w//2, yb, w//2, h)]
            elif num == 3:
                # I due più corti sopra (w/2), il più lungo sotto (w piena)
                sotto_celle = [
                    (xb, yb, w//2, h//2),         # Corso corto 1
                    (xb + w//2, yb, w//2, h//2),  # Corso corto 2
                    (xb, yb + h//2, w, h//2)      # Corso lungo (base intera)
                ]
            else: # 4 o più corsi
                sotto_celle = [
                    (xb, yb, w//2, h//2), (xb + w//2, yb, w//2, h//2),
                    (xb, yb + h//2, w//2, h//2), (xb + w//2, yb + h//2, w//2, h//2)
                ]
            
            pad_val = PADDING_MULTI if num > 1 else PADDING_STD
            
            # Disegno dei corsi (limitato a 4 per layout)
            for i, dati in enumerate(corsi_ordinati[:4]):
                sx, sy, sw, sh = sotto_celle[i]
                
                if dati['corso'] not in color_map:
                    color_map[dati['corso']] = COLORS[next_color % len(COLORS)]
                    next_color += 1
                
                # Sfondo e bordo cella
                draw.rectangle([sx, sy, sx+sw, sy+sh], fill=color_map[dati['corso']])
                draw.rectangle([sx, sy, sx+sw, sy+sh], outline="black", width=1*SCALE)
                
                # Area utile interna (padding)
                area_w, area_h = sw * (1 - pad_val * 2), sh * (1 - pad_val * 2)
                righe_c, f_c, righe_a, f_a = formatta_blocco_definitivo(dati['corso'], dati['aula'], draw, area_w, area_h, num > 1)
                
                # Calcolo altezze per centratura verticale
                h_corso = len(righe_c) * f_c.size * 1.15
                h_aula = len(righe_a) * f_a.size * 1.15
                total_h_text = h_corso + h_aula + (4 * SCALE)
                curr_y = sy + (sh - total_h_text) / 2
                
                # Scrittura Corso
                for line in righe_c:
                    draw.text((sx + (sw - draw.textlength(line, f_c))/2, curr_y), line, "black", font=f_c)
                    curr_y += f_c.size * 1.15
                
                # Scrittura Aula
                curr_y += 4 * SCALE
                for line_a in righe_a:
                    draw.text((sx + (sw - draw.textlength(line_a, f_a))/2, curr_y), line_a, "black", font=f_a)
                    curr_y += f_a.size * 1.15

    # --- DISEGNO GRIGLIA, HEADER E ORARI ---
    f_h, f_o = get_font(FONT_BOLD, SIZE_HEADER), get_font(FONT_BOLD, SIZE_ORARIO)
    for r in range(len(orari) + 1):
        for c in range(len(GIORNI_SETTIMANA) + 1):
            x, y = offset_x + MARGIN + c * BASE_CELL_WIDTH, offset_y + MARGIN + r * BASE_CELL_HEIGHT
            draw.rectangle([x, y, x+BASE_CELL_WIDTH, y+BASE_CELL_HEIGHT], outline="black", width=BORDER_WIDTH)
            
            if r == 0 and c > 0: # Intestazione Giorni
                t = GIORNI_SETTIMANA[c-1].upper().strip()
                draw.text((x + (BASE_CELL_WIDTH - draw.textlength(t, f_h))/2, y + (BASE_CELL_HEIGHT - SIZE_HEADER)/2), t, "black", f_h)
            elif c == 0 and r > 0: # Colonna Orari
                t = str(orari[r-1]).splitlines()[0].replace(".", ":").strip()
                t = re.sub(r'[^\d: \-]', '', t)
                draw.text((x + (BASE_CELL_WIDTH - draw.textlength(t, f_o))/2, y + (BASE_CELL_HEIGHT - SIZE_ORARIO)/2), t, "black", f_o)

    # Inserimento Logo
    if os.path.exists(LOGO_PATH):
        logo = Image.open(LOGO_PATH).convert("RGBA")
        max_h = int(logo_area_h * 0.75)
        logo = logo.resize((int(logo.size[0]*(max_h/logo.size[1])), max_h), Image.Resampling.LANCZOS)
        img.paste(logo, ((lato_quadrato-logo.size[0])//2, offset_y+MARGIN+table_h+(logo_area_h-max_h)//2), logo)
    
    os.makedirs(output_subfolder, exist_ok=True)
    nome_f = os.path.basename(excel_path).rsplit('.', 1)[0] + (f"_{filtro_gruppo}" if filtro_gruppo else "_FULL") + ".png"
    img.save(os.path.join(output_subfolder, nome_f))

# ... [Il resto del MAIN rimane invariato] ...
def merge_and_save(cartella_sorgente, nome_destinazione):
    files = [f for f in os.listdir(cartella_sorgente) if f.endswith(('.xlsx', '.xls'))]
    if not files: return False
    merged_df = pd.read_excel(os.path.join(cartella_sorgente, files[0]), header=None).astype(str)
    merged_df = merged_df.replace(['nan', 'None'], '')
    for file in files[1:]:
        current_df = pd.read_excel(os.path.join(cartella_sorgente, file), header=None).astype(str)
        current_df = current_df.replace(['nan', 'None'], '')
        for r in range(min(len(merged_df), len(current_df))):
            for c in range(min(merged_df.shape[1], current_df.shape[1])):
                val_new = current_df.iloc[r, c].strip()
                if not val_new or val_new in merged_df.iloc[r, c] or any(g.lower() in val_new.lower() for g in GIORNI_SETTIMANA):
                    continue
                sep = "\n--------------\n" if merged_df.iloc[r, c].strip() else ""
                merged_df.iloc[r, c] = f"{merged_df.iloc[r, c]}{sep}{val_new}"
    merged_df.to_excel(os.path.join(INPUT_DIR, nome_destinazione), index=False, header=False)
    return True

if __name__ == "__main__":
    os.system("cls" if os.name == "nt" else "clear")
    svuota_cartella(OUTPUT_DIR)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    if not os.path.exists(INPUT_DIR): os.makedirs(INPUT_DIR)
    
    madri = [d for d in os.listdir(INPUT_DIR) if os.path.isdir(os.path.join(INPUT_DIR, d))]
    for madre in madri:
        madre_path = os.path.join(INPUT_DIR, madre)
        gruppi = [g for g in os.listdir(madre_path) if os.path.isdir(os.path.join(madre_path, g))]
        for gruppo in gruppi:
            path_gruppo = os.path.join(madre_path, gruppo)
            nome_unificato = f"{madre}_{gruppo}.xlsx"
            if merge_and_save(path_gruppo, nome_unificato):
                print(f"[MERGE OK] {nome_unificato}")

    excel_pronti = [f for f in os.listdir(INPUT_DIR) if os.path.isfile(os.path.join(INPUT_DIR, f)) and f.endswith('.xlsx')]
    for excel in excel_pronti:
        full_path = os.path.join(INPUT_DIR, excel)
        prefix = excel.split('_')[0]
        folder_output = os.path.join(OUTPUT_DIR, prefix)
        df_check = pd.read_excel(full_path).astype(str)
        testo_totale = " ".join(df_check.values.flatten())
        base_name = excel.rsplit('.', 1)[0]

        if "(A-H)" in testo_totale or "(I-Z)" in testo_totale:
            genera_immagine_quadrata(full_path, folder_output, "AH")
            print(f"[GRAFICA OK] {base_name}_AH.png")

            genera_immagine_quadrata(full_path, folder_output, "IZ")
            print(f"[GRAFICA OK] {base_name}_IZ.png")
        else:
            genera_immagine_quadrata(full_path, folder_output)
            print(f"[GRAFICA OK] {base_name}_FULL.png")
            
    print("\nCompletato!")