# Script per le generazioni degli orari

Questi due script servono per poter generare gli orari per le foto profilo dei gruppi e per l'affissione davanti alle aule, a partire dai fogli excel che si possono scaricare su EasyCourse. 

## 1. Prerequisiti (Importante!)

ℹ️ Prima di installare, assicurati di avere **Python 3.7+**.

    1.  **Installa Python (se non l'hai mai fatto):**

        * Verifica se è già installato aprendo 'terminale' (cmd su windows) e digitando l'istruzione: python --version. Se viene visualizzata la versione vorrà dire che è già installato.
        * Se non è installato, vai su [python.org](https://www.python.org/downloads/) e scarica l'installer per il tuo sistema operativo.
        * Esegui l'installer.
        * **IMPORTANTE:** Nella prima schermata, spunta la casella **"Add Python to PATH"** in basso.

    2. **Installa le dipendenze**

        * Per installare i moduli necessari si puo' digitare il comando da terminale:
        -------> pip install pandas openpyxl Pillow
        * Dopo questo processo, potrebbe succedere che i moduli non siano installati correttamente inibendo così l'utilizzo. Potrebbe essere utile cambiare versione di Python

## 2. Come si usa?

In entrambi i casi, bisogna avere a disposizione i fogli excel da posizionare opportunamente nella cartella 'input_excel'. 

ORARIO GRUPPI
Nel caso dei gruppi classe, c'è la possibilità di indicare più folgi excel per una singola classe (da posizionare nell'opportuna cartella). La fusione avviene in modo automatico, senza però tollerare eventuali sovrapposizioni di corsi uguali nello stesso orario (se voglio fondere due excel che il lunedì alla terza ora hanno entrambi lo stesso corso, vengono trattati come diversi). Nel caso in cui ci sia un sovraffollamento particolare (sovrapposizione di molti corsi nello stesso orario) potrebbe capitare di dover modificare manualmente i fogli excel per abbreviare a mano i nomi dei corsi per ottenere il risultato desiderato.

ORARIO CLASSI
In questo caso è sufficiente fornire solo i fogli excel degli edifici desiderati poichè la generazione degli orari avviene in una modalità più semplice 