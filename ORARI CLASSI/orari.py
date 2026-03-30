import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import re

# ================= CONFIGURAZIONE =================
INPUT_DIR = "input_excel" 
OUTPUT_DIR = "output_pdf"
PIE_PATH = "pie_di_pagina.png"
GIORNI_SETTIMANA = ["lunedì", "martedì", "mercoledì", "giovedì", "venerdì"]

# --- WHITELIST: Inserisci qui i nomi puliti delle aule che vuoi tenere ---
# Esempio: ["AULA 107", "AULA 24"]. Se lasci la lista vuota [], le terrà TUTTE.
# --- CARICAMENTO WHITELIST DA FILE TXT ---
AULE_DESIDERATE = []
file_whitelist = "whitelist.txt"
    
if os.path.exists(file_whitelist):
    with open(file_whitelist, "r", encoding="utf-8") as f:
        # Legge le righe, toglie gli spazi bianchi e ignora le righe vuote
        AULE_DESIDERATE = [line.strip().upper() for line in f if line.strip()]
    print(f"Caricate {len(AULE_DESIDERATE)} aule da {file_whitelist}")
else:
    print(f"Avviso: {file_whitelist} non trovato. Verranno elaborate TUTTE le aule.")

SCALE = 2
A4_WIDTH, A4_HEIGHT = int(827 * SCALE), int(1170 * SCALE) 
MARGIN = int(40 * SCALE)
HEADER_H = int(140 * SCALE)
FOOTER_AREA_H = int(200 * SCALE) 
CELL_PADDING_PCT = 0.10 

FONT_MAX = 22 * SCALE
FONT_MIN = 10 * SCALE
FONT_GIORNI_SIZE = 19 * SCALE 
THICK_BORDER = 2 * SCALE

FONT_BOLD = os.path.join("fonts", "arialbd.ttf")
FONT_REG  = os.path.join("fonts", "arial.ttf")
FONT_AULA = os.path.join("fonts", "American Captain.ttf")

DARK_COLORS = [
    (180, 40, 40), (40, 120, 40), (40, 60, 150), (130, 80, 20), 
    (100, 40, 140), (20, 100, 110), (150, 70, 20), (60, 60, 60),
    (0, 51, 102), (0, 102, 51), (102, 0, 0), (102, 51, 0),
    (51, 0, 102), (0, 102, 102), (130, 0, 80), (153, 102, 0)
]

# ================= LOGICA DI PULIZIA E FITTING =================

def pulisci_nome_aula(nome):
    if not isinstance(nome, str): return ""
    u_nome = nome.upper().replace("SPAZIO PER ATTIVITA' COMPLEMENTARI", "AULA").replace("SPAZIO PER ATTIVITÀ COMPLEMENTARI", "AULA")
    parole = u_nome.split()[:3]
    nome_pulito = re.sub(r'[^A-Z0-9\s]', ' ', " ".join(parole))
    return " ".join(nome_pulito.split()).strip()

def clean_string_corso(text):
    if not text or str(text).lower() == "nan": return ""
    text = re.sub(r'[,\*]', ' ', str(text))
    text = re.sub(r'[^a-zA-Z0-9\s\.\(\)\-]', '', text)
    return " ".join(text.split()).upper()

def abbrevia_parola_piu_lunga(parole):
    idx_target = -1
    max_len = 0
    for i, p in enumerate(parole):
        p_senza_punto = p.rstrip('.')
        if len(p_senza_punto) > max_len and len(p_senza_punto) > 3:
            max_len = len(p_senza_punto)
            idx_target = i
    if idx_target != -1:
        parola_target = parole[idx_target].rstrip('.')
        parole[idx_target] = parola_target[:-1] + "."
        return True, parole
    return False, parole

def wrap_una_parola_in_piu(parole, draw, font, max_w):
    linee = []
    curr_parole = list(parole)
    while curr_parole:
        temp_line = []
        limit_fisico = 0
        for p in curr_parole:
            if draw.textlength(" ".join(temp_line + [p]), font=font) <= max_w:
                limit_fisico += 1
                temp_line.append(p)
            else: break
        cnt = limit_fisico + 1
        linee.append(" ".join(curr_parole[:cnt]))
        curr_parole = curr_parole[cnt:]
    return linee

def fit_course_text(draw, testo, max_w, max_h):
    parole_orig = testo.split()
    line_spacing = 1.1
    for size in range(FONT_MAX, FONT_MIN - 1, -1):
        f = ImageFont.truetype(FONT_BOLD, size)
        if any(draw.textlength(p, font=f) > max_w for p in parole_orig): continue
        linee = []
        curr_p = list(parole_orig)
        while curr_p:
            temp_l = []
            for p in curr_p:
                if draw.textlength(" ".join(temp_l + [p]), font=f) <= max_w: temp_l.append(p)
                else: break
            if not temp_l: temp_l = [curr_p[0]]
            linee.append(" ".join(temp_l))
            curr_p = curr_p[len(temp_l):]
        if (len(linee) * size * line_spacing) <= max_h: return linee, f, size
    parole_abr = list(parole_orig)
    f_min = ImageFont.truetype(FONT_BOLD, FONT_MIN)
    while True:
        success, parole_abr = abbrevia_parola_piu_lunga(parole_abr)
        if any(draw.textlength(p, font=f_min) > max_w for p in parole_abr):
            if success: continue
            else: break
        linee = []
        curr_p = list(parole_abr)
        while curr_p:
            temp_l = []
            for p in curr_p:
                if draw.textlength(" ".join(temp_l + [p]), font=f_min) <= max_w: temp_l.append(p)
                else: break
            if not temp_l: temp_l = [curr_p[0]]
            linee.append(" ".join(temp_l))
            curr_p = curr_p[len(temp_l):]
        if (len(linee) * FONT_MIN * line_spacing) <= max_h: return linee, f_min, FONT_MIN
        if not success: break
    return wrap_una_parola_in_piu(parole_abr, draw, f_min, max_w), f_min, FONT_MIN

# ================= GENERAZIONE GRAFICA =================

def crea_orario_aula(nome_aula, dati_aula):
    img = Image.new("RGB", (A4_WIDTH, A4_HEIGHT), "white")
    draw = ImageDraw.Draw(img)
    titolo = nome_aula # Il nome arriva già pulito dal main
    try: f_titolo = ImageFont.truetype(FONT_AULA, int(85 * SCALE))
    except: f_titolo = ImageFont.truetype(FONT_BOLD, int(85 * SCALE))
    draw.text(((A4_WIDTH - draw.textlength(titolo, f_titolo)) / 2, MARGIN), titolo, "black", font=f_titolo)
    
    orari_std = ["08:30-09:30", "09:30-10:30", "10:30-11:30", "11:30-12:30", "12:30-13:30", 
                  "13:30-14:30", "14:30-15:30", "15:30-16:30", "16:30-17:30", "17:30-18:30"]
    col_w, row_h = (A4_WIDTH - 2 * MARGIN) // 6, (A4_HEIGHT - HEADER_H - FOOTER_AREA_H - MARGIN) // (len(orari_std) + 1)
    
    for i in range(7): draw.line([MARGIN + i*col_w, HEADER_H, MARGIN + i*col_w, HEADER_H + (len(orari_std)+1)*row_h], "black", THICK_BORDER)
    for j in range(len(orari_std) + 2): draw.line([MARGIN, HEADER_H + j*row_h, MARGIN + 6*col_w, HEADER_H + j*row_h], "black", THICK_BORDER)

    f_g = ImageFont.truetype(FONT_BOLD, int(FONT_GIORNI_SIZE))
    for i, g in enumerate(GIORNI_SETTIMANA):
        draw.text((MARGIN + (i+1)*col_w + (col_w - draw.textlength(g.upper(), f_g))/2, HEADER_H + (row_h - FONT_GIORNI_SIZE)/2), g.upper(), "black", font=f_g)

    mappa_colori = {}
    for r_idx, ora_s in enumerate(orari_std):
        y_c = HEADER_H + (r_idx + 1) * row_h
        f_ora = ImageFont.truetype(FONT_BOLD, int(16 * SCALE))
        draw.text((MARGIN + (col_w - draw.textlength(ora_s, f_ora))/2, y_c + (row_h - 16*SCALE)/2), ora_s, "black", font=f_ora)
        for c_idx, giorno in enumerate(GIORNI_SETTIMANA):
            testo_raw = dati_aula.get(ora_s, {}).get(giorno, "")
            materia = clean_string_corso(re.split(r' - ', testo_raw)[0])
            if materia:
                if materia not in mappa_colori: mappa_colori[materia] = DARK_COLORS[len(mappa_colori) % len(DARK_COLORS)]
                p_w, p_h = col_w * (1 - 2*CELL_PADDING_PCT), row_h * (1 - 2*CELL_PADDING_PCT)
                linee, font_f, sz = fit_course_text(draw, materia, p_w, p_h)
                v_spacing, curr_y = sz * 1.1, y_c + (row_h - (len(linee) * (sz * 1.1)))/2
                for l in linee:
                    draw.text((MARGIN + (c_idx+1)*col_w + (col_w - draw.textlength(l, font_f))/2, curr_y), l, mappa_colori[materia], font=font_f)
                    curr_y += v_spacing

    if os.path.exists(PIE_PATH):
        pie = Image.open(PIE_PATH).convert("RGBA")
        nuova_w = int(A4_WIDTH * 0.85)
        pie = pie.resize((nuova_w, int(nuova_w * (pie.height / pie.width))), Image.Resampling.LANCZOS)
        img.paste(pie, ((A4_WIDTH - nuova_w) // 2, A4_HEIGHT - pie.height - 20), pie)
    return img

# ================= MAIN =================

def main():
    os.system("cls" if os.name == "nt" else "clear")
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    aule_data = {}
    
    # Preparazione Whitelist
    filtro_attivo = len(AULE_DESIDERATE) > 0
    white_list = AULE_DESIDERATE
    
    files = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(('.xlsx', '.xls'))]
    
    for f in files:
        file_path = os.path.join(INPUT_DIR, f)
        dict_df = pd.read_excel(file_path, sheet_name=None, header=None)
        for sheet_name, df in dict_df.items():
            giorno = next((g for g in GIORNI_SETTIMANA if g in sheet_name.lower()), None)
            if not giorno: continue
            
            idx_830 = -1
            for i, r in df.iterrows():
                row_str = " ".join(r.astype(str).str.lower().values)
                if "08:30" in row_str or "08.30" in row_str or "8:30" in row_str:
                    idx_830 = i
                    break
            if idx_830 <= 0: continue 

            df_dati = df.iloc[idx_830:].reset_index(drop=True)
            df_dati.columns = df.iloc[idx_830 - 1]
            
            col_ora = df_dati.columns[0]
            col_aule = [c for c in df_dati.columns[1:] if pd.notna(c) and "unnamed" not in str(c).lower()]
            
            for _, row in df_dati.iterrows():
                ora_raw = str(row[col_ora]).strip()
                if not ora_raw or '-' not in ora_raw or ora_raw.lower() == 'nan': continue
                ora_n = re.sub(r'[^0-9\-:]', '', ora_raw).replace('.', ':')
                
                for a_col in col_aule:
                    valore = row[a_col]
                    if pd.isna(valore) or str(valore).strip().lower() in ['nan', '']: continue
                    
                    # LOGICA WHITELIST BASATA SU NOME PULITO (GRAFICA)
                    nome_pulito = pulisci_nome_aula(str(a_col).strip())
                    
                    if filtro_attivo and nome_pulito not in white_list:
                        continue
                    
                    if nome_pulito not in aule_data: aule_data[nome_pulito] = {}
                    if ora_n not in aule_data[nome_pulito]: aule_data[nome_pulito][ora_n] = {}
                    aule_data[nome_pulito][ora_n][giorno] = str(valore)

    pagine = []
    for a_nome in sorted(aule_data.keys()):
        print(f"Rendering: {a_nome}")
        try:
            pagine.append(crea_orario_aula(a_nome, aule_data[a_nome]).convert("RGB"))
        except Exception as e: print(f"Errore rendering {a_nome}: {e}")

    if pagine:
        nome_file = "Orario_Aule_Selezionate.pdf" if filtro_attivo else "Orario_Aule_Completo.pdf"
        pagine[0].save(os.path.join(OUTPUT_DIR, nome_file), save_all=True, append_images=pagine[1:])
        print(f"\nSuccesso! Creato {nome_file} con {len(pagine)} pagine.")
    else:
        print("\nErrore: Nessuna aula trovata. Controlla i nomi nella Whitelist.")

if __name__ == "__main__":
    main()