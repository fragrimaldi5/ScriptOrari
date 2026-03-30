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

SCALE = 2
BASE_CELL_WIDTH = int(460 * SCALE)
BASE_CELL_HEIGHT = int(260 * SCALE)
MARGIN = int(50 * SCALE)
BORDER_WIDTH = 5 * SCALE 

AULA_MIN_WIDTH_PERC = 0.60
PADDING_STD = 0.07
PADDING_MULTI = 0.05
MIN_LENGTH_WORD_ABBREVIATED = 4

FONT_BOLD = os.path.join("fonts", "arialbd.ttf")
FONT_REG = os.path.join("fonts", "arial.ttf")

SIZE_HEADER = int(48 * SCALE)
SIZE_ORARIO = int(55 * SCALE)
SIZE_CORSO_MAX, SIZE_CORSO_MIN = int(65 * SCALE), int(38 * SCALE)
SIZE_AULA_MAX, SIZE_AULA_MIN = int(65 * SCALE), int(36 * SCALE)
SIZE_MULTI_CORSO_MAX, SIZE_MULTI_CORSO_MIN = int(55 * SCALE), int(25 * SCALE)
SIZE_MULTI_AULA_MAX, SIZE_MULTI_AULA_MIN = int(38 * SCALE), int(10 * SCALE)

COLORS = [
    (240, 128, 128), # 1. Light Coral (Rosso)
    (144, 238, 144), # 2. Light Green (Verde)
    (150, 190, 230), # 3. Sky Blue (Azzurro)
    (245, 245, 150), # 4. Pale Goldenrod (Giallo)
    (245, 190, 130), # 5. Sandy Brown (Arancio)
    (180, 160, 230), # 6. Soft Purple (Viola)
    (150, 225, 225), # 7. Pale Turquoise (Teal)
    (230, 160, 210), # 8. Pink Orchid (Rosa)
    (190, 220, 160), # 9. Olive Pastel
    (235, 200, 170), # 10. Apricot
    (160, 230, 200), # 11. Mint Green
    (210, 180, 140), # 12. Tan/Beige
    (180, 210, 255), # 13. Baby Blue
    (255, 180, 180), # 14. Melon/Salmon
    (200, 250, 150), # 15. Lime Pastel
    (220, 220, 220), # 16. Silver/Gray
    (250, 210, 250), # 17. Lavender Blush
    (170, 200, 200), # 18. Cadet Blue Gray
    (255, 225, 130), # 19. Bright Maize (Giallo carico)
    (200, 180, 200)  # 20. Muted Plum
]

# ================= FUNZIONI DI SUPPORTO =================

def svuota_excel(directory):
    if os.path.exists(directory):
        for elem in os.listdir(directory):
            path = os.path.join(directory, elem)

            try:
                if (os.path.isfile(path) or os.path.islink(path)) and path.endswith(('.xls', '.xlsx')):
                    os.unlink(path)
            except Exception as e:
                print(f"Errore eliminazione {path}: {e}")

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
    text = re.sub(r"[^a-zA-Z0-9àèéìòùÀÈÉÌÒÙ\s\.']", " ", str(text))
    text = re.sub(r"\s+[-']\s+", " ", text)
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

def abbrevia_per_lunghezza_dinamica(parole_originali, draw, font, max_w, max_righe):
    """
    Abbrevia iterativamente le parole più lunghe finché il testo non rientra nel max_w 
    o non ci sono più parole >= 5 lettere.
    """
    parole_lavoro = [p.replace(".", "") for p in parole_originali]
    
    while True:
        righe_attuali = dividi_greedy(parole_lavoro, draw, font, max_w)
        
        # Se rientriamo nei limiti di righe e ogni riga è entro max_w, abbiamo finito
        if len(righe_attuali) <= max_righe and all(draw.textlength(r, font=font) <= max_w for r in righe_attuali):
            return righe_attuali
            
        # Altrimenti, cerchiamo la parola più lunga (tra quelle in righe che sforano o per ridurre il numero righe)
        # Per semplicità cerchiamo la più lunga in assoluto >= 5 lettere
        candidati = [(len(p), i) for i, p in enumerate(parole_lavoro) if len(p.replace(".", "")) >= MIN_LENGTH_WORD_ABBREVIATED]
        
        if not candidati:
            # Non ci sono più parole da abbreviare, restituiamo il wrap attuale (verrà scalato col font dopo)
            return righe_attuali
            
        candidati.sort(key=lambda x: x[0], reverse=True)
        idx_da_accorciare = candidati[0][1]
        parole_lavoro[idx_da_accorciare] = parole_lavoro[idx_da_accorciare].replace(".", "")[:-1] + "."

def formatta_blocco_definitivo(testo, aula, draw, max_w, max_h, num_corsi):
    is_multi = num_corsi > 1
    
    # 1. Definizione Limiti Righe
    if num_corsi == 2: lim_r_c, lim_r_a = 4, 2
    elif num_corsi >= 3: lim_r_c, lim_r_a = 2, 1
    else: lim_r_c, lim_r_a = 3, 1

    # 2. Configurazione Range Font e Padding
    if is_multi: 
        c_max, c_min = SIZE_MULTI_CORSO_MAX, SIZE_MULTI_CORSO_MIN
        a_max, a_min = SIZE_MULTI_AULA_MAX, SIZE_MULTI_AULA_MIN
        pad_perc = PADDING_MULTI
    else: 
        c_max, c_min = SIZE_CORSO_MAX, SIZE_CORSO_MIN
        a_max, a_min = SIZE_AULA_MAX, SIZE_AULA_MIN
        pad_perc = PADDING_STD

    # Calcolo spazio usabile rigoroso
    usable_w = max_w * (1 - (pad_perc * 2))
    usable_h = max_h * (1 - (pad_perc * 2))
    interlinea, gap_fisso = 1.10, 4 * SCALE
    soglia_critica = c_min + (c_max - c_min) // 2
    
    parole_c_orig = testo.split()
    parole_a_orig = aula.split()

    best_fs_c, best_righe_c = None, []
    successo_corso = False

    # --- FASE 1: RICERCA A ONDATE (Vincoli Usable Area) ---
    for n_righe in range(1, lim_r_c + 1):
        fs_c = c_max  # Reset font al massimo per ogni nuova ondata
        while fs_c >= soglia_critica:
            f_test = get_font(FONT_BOLD, fs_c)
            tentativo = dividi_greedy(parole_c_orig, draw, f_test, usable_w)
            
            if len(tentativo) <= n_righe and all(draw.textlength(r, font=f_test) <= usable_w for r in tentativo):
                # Verifica altezza (con aula al minimo e gap fisso)
                h_c = len(tentativo) * fs_c * interlinea
                h_a_min = a_min * interlinea
                if (h_c + h_a_min + gap_fisso) <= usable_h:
                    best_fs_c, best_righe_c = fs_c, tentativo
                    successo_corso = True
                    break
            fs_c -= 1
        if successo_corso: break

    # --- FASE 2: PROTOCOLLO ABBREVIAZIONE CHIRURGICA ---
    if not successo_corso:
        best_fs_c = soglia_critica
        f_critico = get_font(FONT_BOLD, best_fs_c)
        # Abbreviazione mirata sulla riga che sfora
        best_righe_c = abbrevia_per_lunghezza_dinamica(parole_c_orig, draw, f_critico, usable_w, lim_r_c)
        
        # Riduzione font extra-soglia se necessario
        while best_fs_c > c_min:
            f_curr = get_font(FONT_BOLD, best_fs_c)
            if len(best_righe_c) <= lim_r_c and all(draw.textlength(r, f_curr) <= usable_w for r in best_righe_c):
                break
            best_fs_c -= 1
            best_righe_c = dividi_greedy([p for r in best_righe_c for p in r.split()], draw, get_font(FONT_BOLD, best_fs_c), usable_w)

    # --- FASE 3: FITTING AULA E BILANCIAMENTO (Su usable_w) ---
    fs_a = a_max
    f_a = get_font(FONT_REG, fs_a)
    stringa_aula = " ".join(parole_a_orig)
    
    # Adattamento iniziale aula su riga singola
    while (draw.textlength(stringa_aula, f_a) > usable_w) and fs_a > a_min:
        fs_a -= 1
        f_a = get_font(FONT_REG, fs_a)
    
    best_righe_a = [stringa_aula] if not (num_corsi == 2 and draw.textlength(stringa_aula, f_a) > usable_w) else dividi_greedy(parole_a_orig, draw, f_a, usable_w)
    best_fs_a = fs_a

    # DITTATURA DELL'AULA: Bilanciamento basato su usable_w
    while draw.textlength(best_righe_a[0], get_font(FONT_REG, best_fs_a)) < (usable_w * AULA_MIN_WIDTH_PERC) and best_fs_a < a_max:
        if best_fs_c <= c_min: break
        best_fs_a += 1
        best_fs_c -= 1
        f_a_curr, f_c_curr = get_font(FONT_REG, best_fs_a), get_font(FONT_BOLD, best_fs_c)
        best_righe_c = dividi_greedy([p for r in best_righe_c for p in r.split()], draw, f_c_curr, usable_w)
        if len(best_righe_c) > lim_r_c: best_righe_c = best_righe_c[:lim_r_c]
        if len(best_righe_a) > 1:
            best_righe_a = dividi_greedy(parole_a_orig, draw, f_a_curr, usable_w)

    # --- FASE 4: SATURAZIONE PARALLELA FINALE (Controlli Spaziali Rigorosi) ---
    for _ in range(20): # Tentativi di crescita
        n_fs_c, n_fs_a = best_fs_c + 1, best_fs_a + 1
        if n_fs_c > c_max or n_fs_a > a_max: break
        
        nf_c, nf_a = get_font(FONT_BOLD, n_fs_c), get_font(FONT_REG, n_fs_a)
        nr_c = dividi_greedy([p for r in best_righe_c for p in r.split()], draw, nf_c, usable_w)
        nr_a = [best_righe_a[0]] if len(best_righe_a) == 1 else dividi_greedy(parole_a_orig, draw, nf_a, usable_w)
        
        h_t = (len(nr_c) * n_fs_c * interlinea) + (len(nr_a) * n_fs_a * interlinea) + gap_fisso
        if h_t <= usable_h and len(nr_c) <= lim_r_c and len(nr_a) <= lim_r_a:
            if all(draw.textlength(r, nf_c) <= usable_w for r in nr_c) and all(draw.textlength(r, nf_a) <= usable_w for r in nr_a):
                best_fs_c, best_fs_a, best_righe_c, best_righe_a = n_fs_c, n_fs_a, nr_c, nr_a
                continue
        break

    return best_righe_c, get_font(FONT_BOLD, best_fs_c), best_righe_a, get_font(FONT_REG, best_fs_a)

def estrai_dati_cella(cella):
    val = str(cella).strip()
    if any(g.lower() in val.lower() for g in GIORNI_SETTIMANA) or val.lower() == "nan" or not val:
        return []
    
    blocchi = [b.strip() for b in val.split("--------------") if b.strip()] if "--------------" in val else [val]
    risultati = []

    # Per tenere traccia delle aule già inserite in questa cella
    aule_gestite = set()

    for blocco in blocchi:
        righe = [r.strip() for r in blocco.splitlines() if r.strip()]
        if not righe: continue
        
        gruppo = "AH" if "(A-H)" in blocco else "IZ" if "(I-Z)" in blocco else None
        corso = clean_string(righe[0].replace("(A-H)", "").replace("(I-Z)", ""))
        aula_raw = righe[-1].upper() if len(righe) > 1 else ""
        
        # --- LOGICA DI PULIZIA AULA ---
        if any(t in aula_raw for t in ["SPAZIO PER ATTIVITÀ COMPLEMENTARI", "SPAZIO PER ATTIVITA COMPLEMENTARI", "SPAZIO PER ATTIVITA' COMPLEMENTARI"]):
            aula_raw = f"AULA {aula_raw.split('_')[-1]}" if "_" in aula_raw else "AULA"
        elif any(t in aula_raw for t in ["LAB DI ING SANITARIA", "LAB. DI ING SANITARIA", "LAB. DI ING. SANITARIA"]):
            aula_raw = "LAB. ING. SANITARIA"
        elif "CENTRO LINGUISTICO DI ATENEO" in aula_raw: aula_raw = "CLA"
        elif "DIDATTICA A DISTANZA" in aula_raw: aula_raw = "DAD"
        if any(t in aula_raw for t in ["LAB STRUTTURE", "LAB. STRUTTURE"]): aula_raw = "LAB. STRUTTURE"
        if "LABORATORIO" in aula_raw: aula_raw = aula_raw.replace("LABORATORIO", "LAB.")

        parole_aula = aula_raw.split()[:3]
        aula = clean_string(" ".join(parole_aula))
        if "LAB " in aula or aula == "LAB": aula = aula.replace("LAB", "LAB.")
        
        # --- LOGICA DI FILTRAGGIO ---
        # Se l'aula è già presente nella lista 'aule_gestite', saltiamo questo blocco
        if aula in aule_gestite:
            continue
        
        # Aggiungiamo l'aula al set e il risultato alla lista finale
        aule_gestite.add(aula)
        risultati.append({"corso": corso, "aula": aula, "gruppo": gruppo})
        
    return risultati

# ================= CORE GRAFICA =================

def genera_immagine_quadrata(excel_path, output_subfolder, filtro_gruppo=None):
    df = pd.read_excel(excel_path, header=None)
    try: start_row = next(i for i, row in df.iterrows() if 'lunedì' in str(row.values).lower())
    except: return
        
    df_tabella = df.iloc[start_row+1:].copy()
    df_tabella = df_tabella[df_tabella.iloc[:, 0].astype(str).str.contains('-', na=False)]
    orari, matrice_dati = df_tabella.iloc[:, 0].values, df_tabella.iloc[:, 1:6]
    
    table_w, table_h = (len(GIORNI_SETTIMANA) + 1) * BASE_CELL_WIDTH, (len(orari) + 1) * BASE_CELL_HEIGHT
    logo_area_h, content_w = int(220 * SCALE), table_w + MARGIN * 2
    content_h = table_h + MARGIN * 2 + logo_area_h
    lato_quadrato = max(content_w, content_h)
    
    img = Image.new("RGB", (lato_quadrato, lato_quadrato), "white")
    draw = ImageDraw.Draw(img)
    offset_x, offset_y = (lato_quadrato - content_w) // 2, (lato_quadrato - content_h) // 2
    
    color_map, next_color = {}, 0
    print(f"Generazione orario: {os.path.basename(excel_path)}")

    for r_idx in range(len(orari)):
        for c_idx in range(len(GIORNI_SETTIMANA)):
            corsi_raw = estrai_dati_cella(matrice_dati.iloc[r_idx, c_idx])
            corsi = [c for c in corsi_raw if not filtro_gruppo or not c['gruppo'] or c['gruppo'] == filtro_gruppo]
            if not corsi: continue
            
            corsi_ordinati = sorted(corsi, key=lambda x: len(x['corso']))
            num = len(corsi_ordinati)
            xb, yb = offset_x+MARGIN+(c_idx+1)*BASE_CELL_WIDTH, offset_y+MARGIN+(r_idx+1)*BASE_CELL_HEIGHT
            w, h = BASE_CELL_WIDTH, BASE_CELL_HEIGHT
            
            if num == 1: s_cells = [(xb, yb, w, h)]
            elif num == 2: s_cells = [(xb, yb, w//2, h), (xb+w//2, yb, w//2, h)]
            elif num == 3: s_cells = [(xb, yb, w//2, h//2), (xb+w//2, yb, w//2, h//2), (xb, yb+h//2, w, h//2)]
            else: s_cells = [(xb, yb, w//2, h//2), (xb+w//2, yb, w//2, h//2), (xb, yb+h//2, w//2, h//2), (xb+w//2, yb+h//2, w//2, h//2)]
            
            for i, d in enumerate(corsi_ordinati[:4]):
                sx, sy, sw, sh = s_cells[i]
                if d['corso'] not in color_map:
                    color_map[d['corso']] = COLORS[next_color % len(COLORS)]
                    next_color += 1
                
                draw.rectangle([sx, sy, sx+sw, sy+sh], fill=color_map[d['corso']])
                draw.rectangle([sx, sy, sx+sw, sy+sh], outline="black", width=1*SCALE)
                
                aw, ah = sw * (1 - (PADDING_MULTI if num > 1 else PADDING_STD) * 2), sh * (1 - (PADDING_MULTI if num > 1 else PADDING_STD) * 2)
                righe_c, f_c, righe_a, f_a = formatta_blocco_definitivo(d['corso'], d['aula'], draw, aw, ah, num)
                
                total_h_text = (len(righe_c)*f_c.size + len(righe_a)*f_a.size)*1.10 + (4*SCALE)
                curr_y = sy + (sh - total_h_text) / 2
                
                for line in righe_c:
                    draw.text((sx + (sw - draw.textlength(line, f_c))/2, curr_y), line, "black", font=f_c)
                    curr_y += f_c.size * 1.10
                curr_y += 4 * SCALE
                for line_a in righe_a:
                    draw.text((sx + (sw - draw.textlength(line_a, f_a))/2, curr_y), line_a, "black", font=f_a)
                    curr_y += f_a.size * 1.10

    # Griglia Headers
    f_h, f_o = get_font(FONT_BOLD, SIZE_HEADER), get_font(FONT_BOLD, SIZE_ORARIO)
    for r in range(len(orari) + 1):
        for c in range(len(GIORNI_SETTIMANA) + 1):
            x, y = offset_x + MARGIN + c * BASE_CELL_WIDTH, offset_y + MARGIN + r * BASE_CELL_HEIGHT
            draw.rectangle([x, y, x+BASE_CELL_WIDTH, y+BASE_CELL_HEIGHT], outline="black", width=BORDER_WIDTH)
            if r == 0 and c > 0:
                t = GIORNI_SETTIMANA[c-1].upper()
                draw.text((x + (BASE_CELL_WIDTH - draw.textlength(t, f_h))/2, y + (BASE_CELL_HEIGHT - SIZE_HEADER)/2), t, "black", f_h)
            elif c == 0 and r > 0:
                t = str(orari[r-1]).splitlines()[0].replace(".", ":").strip()
                t = re.sub(r'[^\d: \-]', '', t)
                draw.text((x + (BASE_CELL_WIDTH - draw.textlength(t, f_o))/2, y + (BASE_CELL_HEIGHT - SIZE_ORARIO)/2), t, "black", f_o)

    if os.path.exists(LOGO_PATH):
        logo = Image.open(LOGO_PATH).convert("RGBA")
        max_lh = int(logo_area_h * 0.75)
        logo = logo.resize((int(logo.size[0]*(max_lh/logo.size[1])), max_lh), Image.Resampling.LANCZOS)
        img.paste(logo, ((lato_quadrato-logo.size[0])//2, offset_y+MARGIN+table_h+(logo_area_h-max_lh)//2), logo)
    
    os.makedirs(output_subfolder, exist_ok=True)
    nome_f = os.path.basename(excel_path).rsplit('.', 1)[0] + (f"_{filtro_gruppo}" if filtro_gruppo else "_FULL") + ".png"
    img.save(os.path.join(output_subfolder, nome_f))

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
        if "(A-H)" in testo_totale or "(I-Z)" in testo_totale:
            genera_immagine_quadrata(full_path, folder_output, "AH")
            genera_immagine_quadrata(full_path, folder_output, "IZ")
        else:
            genera_immagine_quadrata(full_path, folder_output)
    
    svuota_excel(INPUT_DIR)

    print("\nCompletato!")