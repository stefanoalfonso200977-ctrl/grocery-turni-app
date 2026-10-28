import os
import glob
import joblib
import pandas as pd
import numpy as np
from sklearn.ensemble import HistGradientBoostingRegressor, HistGradientBoostingClassifier
from sklearn.metrics import roc_auc_score, average_precision_score
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==============================================================================
# CONFIGURAZIONE PERCORSI FILE E AUTO-RILEVAMENTO ULTIMO EXCEL CARICATO
# ==============================================================================
FILE_EXCEL_INPUT = "Distinta_Colli_Effettivi_TUTTI_GLI_ORDINI_CARICATI (7).xlsx"
FILE_EXCEL_OUTPUT = "Report_Previsionale_Bisettimanale_Colli_Addetti_Costi_Produttivita.xlsx"
FILE_MODELLO_ML = "modello_grocery_ml.joblib"


class GroceryForecastingUnifiedSystem:
    def __init__(self, input_excel_path=None, output_excel_path=FILE_EXCEL_OUTPUT):
        self.input_excel_path = self._risolvi_file_input(input_excel_path)
        self.output_excel_path = output_excel_path
        
        # Sottoreparti e mappatura categorica
        self.dept_list = ['ALIMENTARI', 'CHIMICA', 'LIQUIDI', 'PETFOOD', 'PROFUMERIA']
        self.dept_cat_map = {d: i for i, d in enumerate(self.dept_list)}
        self.giorni_ita = ['Lunedì', 'Martedì', 'Mercoledì', 'Giovedì', 'Venerdì', 'Sabato', 'Domenica']
        self.giorni_abbr = ['Lun', 'Mar', 'Mer', 'Gio', 'Ven', 'Sab', 'Dom']
        self.settimane_list = ['Settimana 1', 'Settimana 2']
        
        # Mappatura mesi italiani (per lettura e formattazione output)
        self.mesi_map = {
            'Gennaio': '01', 'Febbraio': '02', 'Marzo': '03', 'Aprile': '04',
            'Maggio': '05', 'Giugno': '06', 'Luglio': '07', 'Agosto': '08',
            'Settembre': '09', 'Ottobre': '10', 'Novembre': '11', 'Dicembre': '12'
        }
        self.mesi_inv_map = {
            1: 'Gennaio', 2: 'Febbraio', 3: 'Marzo', 4: 'Aprile',
            5: 'Maggio', 6: 'Giugno', 7: 'Luglio', 8: 'Agosto',
            9: 'Settembre', 10: 'Ottobre', 11: 'Novembre', 12: 'Dicembre'
        }
        
        # Rese di allestimento target (colli/ora) e durata turno (3 ore)
        self.rese_orarie = {
            'ALIMENTARI': 40.0,
            'CHIMICA': 45.0,
            'LIQUIDI': 35.0,
            'PROFUMERIA': 30.0,
            'PETFOOD': 40.0
        }
        self.ore_turno = 3.0
        
        # Parametri economici (Allestitore 4° livello)
        self.costo_orario = 12.80  # €/ora
        self.costo_turno = round(self.ore_turno * self.costo_orario, 2)  # 38.40 €/turno da 3h

    @staticmethod
    def _risolvi_file_input(percorso_specificato):
        """
        Individua automaticamente l'ultimo file Excel caricato nella directory
        se sono presenti più versioni di 'Distinta_Colli_Effettivi*.xlsx',
        così da aggiornare sempre le date all'ultimo file disponibile.
        """
        candidati = glob.glob("Distinta_Colli_Effettivi*.xlsx")
        if candidati:
            ultimo_file = max(candidati, key=os.path.getmtime)
            if percorso_specificato and os.path.exists(percorso_specificato):
                if os.path.getmtime(ultimo_file) > os.path.getmtime(percorso_specificato):
                    return ultimo_file
                return percorso_specificato
            return ultimo_file
        return percorso_specificato

    @staticmethod
    def arrotonda_soglia_050(val):
        """
        Regola di arrotondamento sottoreparto:
        - Se parte decimale < 0.50 -> arrotondamento per DIFETTO (floor)
        - Se parte decimale >= 0.50 -> arrotondamento per ECCESSO (ceil)
        """
        if val <= 0:
            return 0
        parte_decimale = round(val - np.floor(val), 4)
        if parte_decimale < 0.50:
            return int(np.floor(val))
        else:
            return int(np.ceil(val))

    def _parse_ita_date(self, val):
        if pd.isna(val):
            return pd.NaT
        if isinstance(val, (pd.Timestamp, np.datetime64)):
            return pd.to_datetime(val)
        testo = str(val).strip()
        parts = testo.split()
        # Formato tipo: "Lunedì 12 Ottobre 2026"
        if len(parts) >= 4 and parts[2].capitalize() in self.mesi_map:
            giorno = parts[1].zfill(2)
            mese = self.mesi_map[parts[2].capitalize()]
            anno = parts[3]
            return pd.to_datetime(f"{anno}-{mese}-{giorno}", errors='coerce')
        # Fallback per date standard (es. YYYY-MM-DD o DD/MM/YYYY)
        return pd.to_datetime(testo, dayfirst=True, errors='coerce')

    def _formatta_data_estesa_ita(self, dt):
        """Restituisce stringa 'Lunedì 19 Ottobre 2026'."""
        g_nome = self.giorni_ita[dt.dayofweek]
        m_nome = self.mesi_inv_map[dt.month]
        return f"{g_nome} {dt.day:02d} {m_nome} {dt.year}"

    def _formatta_data_breve_ita(self, dt):
        """Restituisce stringa 'Lun 19/10/2026'."""
        g_abbr = self.giorni_abbr[dt.dayofweek]
        return f"{g_abbr} {dt.strftime('%d/%m/%Y')}"

    # --------------------------------------------------------------------------
    # 1. CARICAMENTO DATI, VINCOLO COLONNA L E CALCOLO CALENDARIO DINAMICO 14 GG
    # --------------------------------------------------------------------------
    def carica_e_pulisci_dati(self):
        print(f"[1/5] Caricamento file Excel: {self.input_excel_path}")
        xls = pd.ExcelFile(self.input_excel_path)
        sheet_principale = xls.sheet_names[0]  # 'Distinta Tutti gli Ordini'
        df = pd.read_excel(self.input_excel_path, sheet_name=sheet_principale)

        # Aggancio posizionale rigoroso:
        # Colonna B (idx 1) = Data Allestimento
        # Colonna K (idx 10) = Colli Ordinati
        # Colonna L (idx 11) = Colli Allestimento Effettivi (TARGET PRINCIPALE)
        col_b_data = df.columns[1]
        col_k_ordinati = df.columns[10]
        col_l_effettivi = df.columns[11]

        print(f"      -> Colonna B agganciata per le date: '{col_b_data}'")
        print(f"      -> Colonna L agganciata per i calcoli effettivi: '{col_l_effettivi}'")

        df['Colli_Ordinati_K'] = pd.to_numeric(df[col_k_ordinati], errors='coerce').fillna(0)
        df['Colli_Effettivi_L'] = pd.to_numeric(df[col_l_effettivi], errors='coerce').fillna(0)

        # Rimozione date mancanti e righe escluse
        df = df.dropna(subset=[col_b_data]).copy()
        if 'Destinazione' in df.columns:
            df = df[df['Destinazione'] != "Eliminata dall'Ordine (Esclusa)"].copy()

        df['Data'] = df[col_b_data].apply(self._parse_ita_date)
        df = df.dropna(subset=['Data']).reset_index(drop=True)
        df['Giorno_Num'] = df['Data'].dt.dayofweek
        df['Giorno_Nome'] = df['Giorno_Num'].map(lambda x: self.giorni_ita[x])

        # Target binario per il modello di previsione tagli/rotture
        df['Taglio'] = (df['Colli_Effettivi_L'] < df['Colli_Ordinati_K']).astype(int)

        # --- CALCOLO AUTOMATICO DELLE DATE DAL FILE EXCEL CARICATO ---
        self.data_min_storico = df['Data'].min().normalize()
        self.data_max_storico = df['Data'].max().normalize()

        # La pianificazione bisettimanale (14 giorni) parte dal Lunedì immediatamente successivo
        # all'ultima data presente nel file Excel caricato:
        giorni_al_prossimo_lunedi = (7 - self.data_max_storico.dayofweek) % 7
        if giorni_al_prossimo_lunedi == 0:
            giorni_al_prossimo_lunedi = 7
        self.data_inizio_piano_2w = self.data_max_storico + pd.Timedelta(days=giorni_al_prossimo_lunedi)
        self.data_fine_piano_2w = self.data_inizio_piano_2w + pd.Timedelta(days=13)

        self.calendario_14gg = pd.date_range(start=self.data_inizio_piano_2w, periods=14, freq='D')
        self.inizio_w1 = self.calendario_14gg[0]
        self.fine_w1 = self.calendario_14gg[6]
        self.inizio_w2 = self.calendario_14gg[7]
        self.fine_w2 = self.calendario_14gg[13]

        self.range_storico_str = f"{self.data_min_storico.strftime('%d/%m/%Y')} - {self.data_max_storico.strftime('%d/%m/%Y')}"
        self.range_w1_str = f"{self.inizio_w1.strftime('%d/%m/%Y')} - {self.fine_w1.strftime('%d/%m/%Y')}"
        self.range_w2_str = f"{self.inizio_w2.strftime('%d/%m/%Y')} - {self.fine_w2.strftime('%d/%m/%Y')}"
        self.range_2w_str = f"{self.data_inizio_piano_2w.strftime('%d/%m/%Y')} - {self.data_fine_piano_2w.strftime('%d/%m/%Y')}"

        print(f"      -> Storico rilevato nel file Excel: dal {self.data_min_storico.strftime('%d/%m/%Y')} al {self.data_max_storico.strftime('%d/%m/%Y')}")
        print(f"      -> Rilascio Turni Settimana 1: dal {self.inizio_w1.strftime('%d/%m/%Y')} al {self.fine_w1.strftime('%d/%m/%Y')}")
        print(f"      -> Rilascio Turni Settimana 2: dal {self.inizio_w2.strftime('%d/%m/%Y')} al {self.fine_w2.strftime('%d/%m/%Y')}")

        self.df = df
        return df

    # --------------------------------------------------------------------------
    # 2. STADIO 1: CLASSIFICAZIONE GRADIENT BOOSTING PER TAGLI / ROTTURE ORDINE
    # --------------------------------------------------------------------------
    def addestra_modello_tagli(self):
        print("[2/5] Addestramento Stadio 1: Classificatore Rischio Tagli/Rotture...")
        df_c = self.df.copy()
        df_c['Colli_Ord_Log'] = np.log1p(df_c['Colli_Ordinati_K'])
        
        tasso_cat = df_c.groupby('Codice Categoria')['Taglio'].mean().rename('Tasso_Taglio_Cat')
        df_c = df_c.join(tasso_cat, on='Codice Categoria')
        df_c['Tasso_Taglio_Cat'] = df_c['Tasso_Taglio_Cat'].fillna(df_c['Taglio'].mean())

        cat_cols = ['Codice Categoria', 'Sottoreparto', 'Tipo Carico', 'Piattaforma Consegna', 'Destinazione']
        for c in cat_cols:
            df_c[c] = df_c[c].astype('category')

        features_clf = cat_cols + ['Colli_Ordinati_K', 'Colli_Ord_Log', 'Giorno_Num', 'Tasso_Taglio_Cat']
        cat_idx = [features_clf.index(c) for c in cat_cols]

        clf = HistGradientBoostingClassifier(
            class_weight='balanced',
            categorical_features=cat_idx,
            max_iter=150,
            learning_rate=0.05,
            max_depth=6,
            random_state=42
        )
        clf.fit(df_c[features_clf], df_c['Taglio'])
        
        prob_taglio = clf.predict_proba(df_c[features_clf])[:, 1]
        self.df['Prob_Rischio_Taglio'] = prob_taglio
        self.clf_tagli = clf

        auc = roc_auc_score(df_c['Taglio'], prob_taglio)
        pr_auc = average_precision_score(df_c['Taglio'], prob_taglio)
        print(f"      -> Performance Classificatore Tagli: ROC-AUC = {auc:.4f} | PR-AUC = {pr_auc:.4f}")

    # --------------------------------------------------------------------------
    # 3. STADIO 2: REGRESSORE COLLI E PROIEZIONE BISETTIMANALE CON DATE DINAMICHE
    # --------------------------------------------------------------------------
    def addestra_e_prevedi_due_settimane(self):
        print(f"[3/5] Addestramento Stadio 2 e Previsione Calendario Dinamico ({self.range_2w_str})...")
        daily = self.df.groupby(['Data', 'Giorno_Num', 'Giorno_Nome', 'Sottoreparto']).agg(
            Colli_Effettivi_L=('Colli_Effettivi_L', 'sum'),
            Righe_Tagliate=('Taglio', 'sum'),
            Rischio_Medio_Taglio=('Prob_Rischio_Taglio', 'mean')
        ).reset_index()

        daily['Dept_Code'] = daily['Sottoreparto'].map(self.dept_cat_map)
        daily = daily.sort_values(['Sottoreparto', 'Data']).reset_index(drop=True)

        # Feature temporali su Colonna L
        daily['Lag_1'] = daily.groupby('Sottoreparto')['Colli_Effettivi_L'].shift(1)
        daily['Lag_7'] = daily.groupby('Sottoreparto')['Colli_Effettivi_L'].shift(6)
        daily['Rolling_3'] = daily.groupby('Sottoreparto')['Colli_Effettivi_L'].shift(1).rolling(3).mean()

        for c in ['Lag_1', 'Lag_7', 'Rolling_3']:
            daily[c] = daily[c].fillna(daily.groupby('Sottoreparto')['Colli_Effettivi_L'].transform('mean'))

        features_reg = ['Giorno_Num', 'Dept_Code', 'Lag_1', 'Lag_7', 'Rolling_3']

        reg = HistGradientBoostingRegressor(
            loss='poisson',
            categorical_features=[0, 1],
            max_iter=150,
            learning_rate=0.05,
            random_state=42
        )
        reg.fit(daily[features_reg], daily['Colli_Effettivi_L'])
        self.reg_colli = reg

        # Inizializzazione memoria storica per la simulazione sui 14 giorni del calendario reale
        recent = daily.groupby('Sottoreparto')[['Lag_1', 'Lag_7', 'Rolling_3']].mean()
        
        rows = []
        storia_previsioni_w1 = {dept: {} for dept in self.dept_list}
        ultimi_colli_attivi = {dept: [
            recent.loc[dept, 'Lag_1'],
            recent.loc[dept, 'Rolling_3'],
            recent.loc[dept, 'Lag_7']
        ] for dept in self.dept_list}

        self.slots_14gg = []

        for idx_14, data_corr in enumerate(self.calendario_14gg):
            w_idx = 1 if idx_14 < 7 else 2
            w_label = f"Settimana {w_idx}"
            range_w_corr = self.range_w1_str if w_idx == 1 else self.range_w2_str
            g_num = data_corr.dayofweek
            g_nome = self.giorni_ita[g_num]
            giorno_prog = idx_14 + 1

            data_str = data_corr.strftime('%d/%m/%Y')
            data_estesa = self._formatta_data_estesa_ita(data_corr)
            data_breve = self._formatta_data_breve_ita(data_corr)
            etichetta_slot = f"S{w_idx} - {data_breve}"
            self.slots_14gg.append(etichetta_slot)

            for dept in self.dept_list:
                if g_num == 6:  # Domenica chiuso / nessun allestimento
                    colli_pred = 0.0
                else:
                    if w_idx == 1:
                        lag_1_val = recent.loc[dept, 'Lag_1']
                        lag_7_val = recent.loc[dept, 'Lag_7']
                        roll_3_val = recent.loc[dept, 'Rolling_3']
                    else:
                        lag_1_val = 0.5 * recent.loc[dept, 'Lag_1'] + 0.5 * ultimi_colli_attivi[dept][-1]
                        lag_7_val = 0.5 * recent.loc[dept, 'Lag_7'] + 0.5 * storia_previsioni_w1[dept].get(g_num, recent.loc[dept, 'Lag_7'])
                        roll_3_val = 0.5 * recent.loc[dept, 'Rolling_3'] + 0.5 * np.mean(ultimi_colli_attivi[dept][-3:])

                    X_p = pd.DataFrame([{
                        'Giorno_Num': g_num,
                        'Dept_Code': self.dept_cat_map[dept],
                        'Lag_1': lag_1_val,
                        'Lag_7': lag_7_val,
                        'Rolling_3': roll_3_val
                    }])
                    colli_pred = max(0.0, reg.predict(X_p)[0])
                    storia_previsioni_w1[dept][g_num] = colli_pred
                    ultimi_colli_attivi[dept].append(colli_pred)

                resa_target = self.rese_orarie[dept]
                capacita_turno_target = resa_target * self.ore_turno
                ore_nec = colli_pred / resa_target
                fte_3h = round(ore_nec / self.ore_turno, 2)
                
                # Applicazione regola di arrotondamento (< 0.50 difetto, >= 0.50 eccesso)
                addetti_soglia_050 = self.arrotonda_soglia_050(fte_3h)
                ore_presidiate_050 = round(addetti_soglia_050 * self.ore_turno, 2)
                bilancio_ore = round((fte_3h - addetti_soglia_050) * self.ore_turno, 2)

                # Metriche di Produttività per Sottoreparto
                capacita_colli_presidio_050 = round(ore_presidiate_050 * resa_target, 1)
                prod_oraria_eff_050 = round(colli_pred / ore_presidiate_050, 2) if ore_presidiate_050 > 0 else 0.0
                prod_turno_eff_050 = round(colli_pred / addetti_soglia_050, 1) if addetti_soglia_050 > 0 else 0.0
                indice_saturazione_050 = round((ore_nec / ore_presidiate_050) * 100, 1) if ore_presidiate_050 > 0 else 0.0

                # Metriche Costo del Lavoro (€ 12,80/h - 4° Livello)
                costo_teorico_euro = round(ore_nec * self.costo_orario, 2)
                costo_presidio_050_euro = round(ore_presidiate_050 * self.costo_orario, 2)
                costo_per_collo_teorico = round(costo_teorico_euro / colli_pred, 4) if colli_pred > 0 else 0.0
                costo_per_collo_050 = round(costo_presidio_050_euro / colli_pred, 4) if colli_pred > 0 else 0.0

                rows.append({
                    'Settimana_Num': w_idx,
                    'Settimana': w_label,
                    'Periodo_Settimana': range_w_corr,
                    'Giorno_Progressivo_14gg': giorno_prog,
                    'Data_Dt': data_corr,
                    'Data': data_str,
                    'Data_Estesa': data_estesa,
                    'Giorno_Num': g_num,
                    'Giorno': g_nome,
                    'Slot_Bisettimanale': etichetta_slot,
                    'Sottoreparto': dept,
                    'Colli_L': round(colli_pred, 1),
                    'Resa_Target_Oraria': resa_target,
                    'Capacita_Target_Turno_3h': capacita_turno_target,
                    'Ore_Richieste': round(ore_nec, 2),
                    'FTE_Teorici_3h': fte_3h,
                    'Addetti_Soglia_050': addetti_soglia_050,
                    'Ore_Presidiate_050': ore_presidiate_050,
                    'Bilancio_Ore_Mobilita': bilancio_ore,
                    'Capacita_Colli_Presidio_050': capacita_colli_presidio_050,
                    'Prod_Effettiva_Colli_Ora_050': prod_oraria_eff_050,
                    'Prod_Effettiva_Colli_Turno_050': prod_turno_eff_050,
                    'Saturazione_Presidio_Pct': indice_saturazione_050,
                    'Costo_Teorico_Euro': costo_teorico_euro,
                    'Costo_Presidio_050_Euro': costo_presidio_050_euro,
                    'Costo_Unitario_Collo_Teorico': costo_per_collo_teorico,
                    'Costo_Unitario_Collo_050': costo_per_collo_050
                })

        self.df_staff_2w = pd.DataFrame(rows)

        # Costruzione tabelle pivot sui 14 giorni con le date aggiornate
        self.piv_colli_2w = self.df_staff_2w.pivot(index='Slot_Bisettimanale', columns='Sottoreparto', values='Colli_L').reindex(self.slots_14gg)
        self.piv_colli_2w['Totale_Colli_L'] = self.piv_colli_2w.sum(axis=1)

        self.piv_fte_2w = self.df_staff_2w.pivot(index='Slot_Bisettimanale', columns='Sottoreparto', values='FTE_Teorici_3h').reindex(self.slots_14gg)
        self.piv_arr050_2w = self.df_staff_2w.pivot(index='Slot_Bisettimanale', columns='Sottoreparto', values='Addetti_Soglia_050').reindex(self.slots_14gg)
        self.piv_ore_ric_2w = self.df_staff_2w.pivot(index='Slot_Bisettimanale', columns='Sottoreparto', values='Ore_Richieste').reindex(self.slots_14gg)
        self.piv_costo_050_2w = self.df_staff_2w.pivot(index='Slot_Bisettimanale', columns='Sottoreparto', values='Costo_Presidio_050_Euro').reindex(self.slots_14gg)
        self.piv_prod_ora_2w = self.df_staff_2w.pivot(index='Slot_Bisettimanale', columns='Sottoreparto', values='Prod_Effettiva_Colli_Ora_050').reindex(self.slots_14gg)

        # Riepilogo giornaliero sui 14 giorni (con colonne Data e Giorno)
        meta_slot = self.df_staff_2w[[
            'Slot_Bisettimanale', 'Settimana_Num', 'Settimana', 'Periodo_Settimana',
            'Data_Dt', 'Data', 'Data_Estesa', 'Giorno_Num', 'Giorno'
        ]].drop_duplicates().set_index('Slot_Bisettimanale').reindex(self.slots_14gg)
        
        riepilogo_14gg = meta_slot.copy()
        riepilogo_14gg['Totale_Colli_Giorno'] = self.piv_colli_2w['Totale_Colli_L'].round(1)
        riepilogo_14gg['Somma_Addetti_Reparto_050'] = self.piv_arr050_2w.sum(axis=1)
        riepilogo_14gg['Turni_Teorici_Totali_FTE'] = self.piv_fte_2w.sum(axis=1).round(2)
        riepilogo_14gg['Ore_Totali_Richieste'] = self.piv_ore_ric_2w.sum(axis=1).round(2)
        riepilogo_14gg['Ore_Coperte_da_Arrotondati_050'] = (riepilogo_14gg['Somma_Addetti_Reparto_050'] * self.ore_turno).round(2)
        riepilogo_14gg['Ore_Nette_Scoperte'] = np.maximum(0.0, riepilogo_14gg['Ore_Totali_Richieste'] - riepilogo_14gg['Ore_Coperte_da_Arrotondati_050']).round(2)

        # Squadra ottimale in mobilità interna (12 Lun-Gio, 11 Ven-Sab per ciascuna settimana)
        riepilogo_14gg['Squadra_Ottimale_Reale'] = np.ceil(riepilogo_14gg['Turni_Teorici_Totali_FTE']).astype(int)
        riepilogo_14gg.loc[riepilogo_14gg['Giorno_Num'] == 6, 'Squadra_Ottimale_Reale'] = 0
        riepilogo_14gg['Addetti_Jolly_Mobilita'] = riepilogo_14gg['Squadra_Ottimale_Reale'] - riepilogo_14gg['Somma_Addetti_Reparto_050']
        riepilogo_14gg['Ore_Reali_Squadra'] = (riepilogo_14gg['Squadra_Ottimale_Reale'] * self.ore_turno).round(2)

        riepilogo_14gg['Produttivita_Media_Squadra_Colli_Ora'] = np.where(
            riepilogo_14gg['Ore_Reali_Squadra'] > 0,
            (riepilogo_14gg['Totale_Colli_Giorno'] / riepilogo_14gg['Ore_Reali_Squadra']).round(2),
            0.0
        )

        # Costo del Lavoro sui 14 giorni (€ 12,80/h)
        riepilogo_14gg['Costo_Lavoro_Teorico_Euro'] = (riepilogo_14gg['Ore_Totali_Richieste'] * self.costo_orario).round(2)
        riepilogo_14gg['Costo_Lavoro_Presidio_050_Euro'] = (riepilogo_14gg['Ore_Coperte_da_Arrotondati_050'] * self.costo_orario).round(2)
        riepilogo_14gg['Costo_Lavoro_Jolly_Euro'] = (riepilogo_14gg['Addetti_Jolly_Mobilita'] * self.ore_turno * self.costo_orario).round(2)
        riepilogo_14gg['Costo_Lavoro_Squadra_Reale_Euro'] = (riepilogo_14gg['Ore_Reali_Squadra'] * self.costo_orario).round(2)
        riepilogo_14gg['Costo_Medio_per_Collo_Reale'] = np.where(
            riepilogo_14gg['Totale_Colli_Giorno'] > 0,
            (riepilogo_14gg['Costo_Lavoro_Squadra_Reale_Euro'] / riepilogo_14gg['Totale_Colli_Giorno']).round(4),
            0.0
        )

        self.riepilogo_14gg = riepilogo_14gg

        # --- RIEPILOGO PRODUTTIVITÀ E COSTI PER SOTTOREPARTO (SETT. 1, SETT. 2 E TOTALE 2 SETTIMANE) ---
        df_attivi = self.df_staff_2w[self.df_staff_2w['Giorno_Num'] < 6]

        prod_per_sett = df_attivi.groupby(['Settimana', 'Sottoreparto']).agg(
            Colli=('Colli_L', 'sum'),
            Resa_Target_Oraria=('Resa_Target_Oraria', 'first'),
            Capacita_Target_Turno_3h=('Capacita_Target_Turno_3h', 'first'),
            Ore_Richieste=('Ore_Richieste', 'sum'),
            Turni_Teorici=('FTE_Teorici_3h', 'sum'),
            Turni_Presidio_050=('Addetti_Soglia_050', 'sum'),
            Ore_Presidio_050=('Ore_Presidiate_050', 'sum'),
            Bilancio_Ore_Mobilita=('Bilancio_Ore_Mobilita', 'sum'),
            Costo_Teorico_Euro=('Costo_Teorico_Euro', 'sum'),
            Costo_Presidio_050_Euro=('Costo_Presidio_050_Euro', 'sum')
        ).reset_index()

        prod_2w = df_attivi.groupby('Sottoreparto').agg(
            Resa_Target_Oraria=('Resa_Target_Oraria', 'first'),
            Capacita_Target_Turno_3h=('Capacita_Target_Turno_3h', 'first'),
            Colli_Bisettimanali=('Colli_L', 'sum'),
            Ore_Richieste_2W=('Ore_Richieste', 'sum'),
            Turni_Teorici_2W=('FTE_Teorici_3h', 'sum'),
            Turni_Presidio_050_2W=('Addetti_Soglia_050', 'sum'),
            Ore_Presidio_050_2W=('Ore_Presidiate_050', 'sum'),
            Bilancio_Ore_Mobilita_2W=('Bilancio_Ore_Mobilita', 'sum'),
            Capacita_Colli_Presidio_050_2W=('Capacita_Colli_Presidio_050', 'sum'),
            Costo_Teorico_2W_Euro=('Costo_Teorico_Euro', 'sum'),
            Costo_Presidio_050_2W_Euro=('Costo_Presidio_050_Euro', 'sum')
        ).reindex(self.dept_list)

        w1_sub = prod_per_sett[prod_per_sett['Settimana'] == 'Settimana 1'].set_index('Sottoreparto').reindex(self.dept_list)
        w2_sub = prod_per_sett[prod_per_sett['Settimana'] == 'Settimana 2'].set_index('Sottoreparto').reindex(self.dept_list)

        prod_2w['Colli_Sett_1'] = w1_sub['Colli'].round(1)
        prod_2w['Turni_050_Sett_1'] = w1_sub['Turni_Presidio_050'].astype(int)
        prod_2w['Costo_050_Sett_1_Euro'] = w1_sub['Costo_Presidio_050_Euro'].round(2)

        prod_2w['Colli_Sett_2'] = w2_sub['Colli'].round(1)
        prod_2w['Turni_050_Sett_2'] = w2_sub['Turni_Presidio_050'].astype(int)
        prod_2w['Costo_050_Sett_2_Euro'] = w2_sub['Costo_Presidio_050_Euro'].round(2)

        prod_2w['Media_Colli_Giorno_Lav'] = (prod_2w['Colli_Bisettimanali'] / 12.0).round(1)
        prod_2w['Media_Colli_Settimana'] = (prod_2w['Colli_Bisettimanali'] / 2.0).round(1)

        prod_2w['Prod_Effettiva_Oraria_050'] = np.where(
            prod_2w['Ore_Presidio_050_2W'] > 0,
            (prod_2w['Colli_Bisettimanali'] / prod_2w['Ore_Presidio_050_2W']).round(2),
            0.0
        )
        prod_2w['Prod_Effettiva_Turno_050'] = np.where(
            prod_2w['Turni_Presidio_050_2W'] > 0,
            (prod_2w['Colli_Bisettimanali'] / prod_2w['Turni_Presidio_050_2W']).round(1),
            0.0
        )
        prod_2w['Saturazione_Media_Presidio_Pct'] = np.where(
            prod_2w['Ore_Presidio_050_2W'] > 0,
            ((prod_2w['Ore_Richieste_2W'] / prod_2w['Ore_Presidio_050_2W']) * 100).round(1),
            0.0
        )
        prod_2w['Costo_per_Collo_Teorico_Euro'] = (prod_2w['Costo_Teorico_2W_Euro'] / prod_2w['Colli_Bisettimanali']).round(4)
        prod_2w['Costo_per_Collo_Presidio_050_Euro'] = np.where(
            prod_2w['Colli_Bisettimanali'] > 0,
            (prod_2w['Costo_Presidio_050_2W_Euro'] / prod_2w['Colli_Bisettimanali']).round(4),
            0.0
        )

        self.prod_2w = prod_2w
        return self.df_staff_2w, self.riepilogo_14gg, self.prod_2w

    # --------------------------------------------------------------------------
    # 4. ESPORTAZIONE REPORT EXCEL BISETTIMANALE CON DATE DINAMICHE (5 FOGLI)
    # --------------------------------------------------------------------------
    def esporta_excel_completo(self):
        print(f"[4/5] Generazione report Excel Bisettimanale con Date Dinamiche: {self.output_excel_path}")
        wb = openpyxl.Workbook()

        ws_mob = wb.active
        ws_mob.title = "1_Turni_2_Settimane_e_Costi"
        ws_orizzonti = wb.create_sheet(title="2_Sintesi_1W_2W_Mese_Trim")
        ws_colli = wb.create_sheet(title="3_Previsione_Colli_14gg_Col_L")
        ws_bilancio = wb.create_sheet(title="4_Dettaglio_14gg_Ore_Costi")
        ws_prod = wb.create_sheet(title="5_Produttivita_e_Costi_2_Sett")

        # Stili grafici professionali
        header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
        sub_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        w2_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        jolly_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        opt_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        cost_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

        white_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        bold_font = Font(name="Calibri", size=11, bold=True)
        reg_font = Font(name="Calibri", size=11)
        title_font = Font(name="Calibri", size=13, bold=True, color="1F497D")
        border = Border(
            left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
            top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF')
        )

        # ======================================================================
        # FOGLIO 1: PIANIFICAZIONE RILASCIO TURNI 2 SETTIMANE CON DATE REALI
        # ======================================================================
        ws_mob.cell(
            row=2, column=2,
            value=f"PIANIFICAZIONE RILASCIO TURNI BISETTIMANALE ({self.range_2w_str}): ADDETTI, MOBILITÀ E COSTO LAVORO (€ 12,80/h)"
        ).font = title_font
        ws_mob.cell(
            row=3, column=2,
            value=f"File Sorgente: {os.path.basename(self.input_excel_path)} (Storico: {self.range_storico_str}) | Settimana 1: {self.range_w1_str} | Settimana 2: {self.range_w2_str}"
        ).font = Font(italic=True, size=10)

        headers_m = [
            "Settimana",
            "Data",
            "Giorno",
            "ALIMENTARI (Soglia 0,50 | FTE)",
            "CHIMICA (Soglia 0,50 | FTE)",
            "LIQUIDI (Soglia 0,50 | FTE)",
            "PETFOOD (Soglia 0,50 | FTE)",
            "PROFUMERIA (Soglia 0,50 | FTE)",
            "Somma Addetti (Soglia 0,50)",
            "Ore Nette Scoperte",
            "Addetti Jolly Mobilità",
            "Turni Teorici Totali (FTE)",
            "SQUADRA OTTIMALE REALE (Turni 3h)",
            "Ore Reali Squadra",
            "Costo Lavoro Teorico (€)",
            "Costo Lavoro Reale Squadra (€)"
        ]
        for i, h in enumerate(headers_m, 2):
            c = ws_mob.cell(row=5, column=i, value=h)
            c.fill, c.font, c.alignment = header_fill, white_font, Alignment(horizontal="center", wrap_text=True)

        r = 6
        for w_idx, w_label in enumerate(self.settimane_list, start=1):
            sub_14 = self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == w_idx]
            range_w_corr = self.range_w1_str if w_idx == 1 else self.range_w2_str

            for slot_idx, row_g in sub_14.iterrows():
                c_w = ws_mob.cell(row=r, column=2, value=w_label)
                c_w.font, c_w.alignment, c_w.border = bold_font, Alignment(horizontal="center"), border

                c_dt = ws_mob.cell(row=r, column=3, value=row_g['Data'])
                c_dt.font, c_dt.alignment, c_dt.border = bold_font, Alignment(horizontal="center"), border

                c_gg = ws_mob.cell(row=r, column=4, value=row_g['Giorno'])
                c_gg.font, c_gg.border = bold_font, border

                for j, d in enumerate(self.dept_list, 5):
                    arr_v = int(self.piv_arr050_2w.loc[slot_idx, d])
                    fte_v = float(self.piv_fte_2w.loc[slot_idx, d])
                    c = ws_mob.cell(row=r, column=j, value=f"{arr_v}  ({fte_v:.2f})")
                    c.alignment, c.border, c.font = Alignment(horizontal="center"), border, reg_font

                s_050 = int(row_g['Somma_Addetti_Reparto_050'])
                o_res = float(row_g['Ore_Nette_Scoperte'])
                jolly = int(row_g['Addetti_Jolly_Mobilita'])
                teo = float(row_g['Turni_Teorici_Totali_FTE'])
                opt = int(row_g['Squadra_Ottimale_Reale'])
                ore_reali = float(row_g['Ore_Reali_Squadra'])
                costo_teo = float(row_g['Costo_Lavoro_Teorico_Euro'])
                costo_reale = float(row_g['Costo_Lavoro_Squadra_Reale_Euro'])

                c_b = ws_mob.cell(row=r, column=10, value=s_050)
                c_b.alignment, c_b.font, c_b.fill, c_b.border = Alignment(horizontal="center"), bold_font, sub_fill, border
                c_o = ws_mob.cell(row=r, column=11, value=o_res)
                c_o.number_format, c_o.border = '0.00 "h"', border
                c_j = ws_mob.cell(row=r, column=12, value=jolly)
                c_j.alignment, c_j.font, c_j.fill, c_j.border = Alignment(horizontal="center"), bold_font, jolly_fill, border
                c_t = ws_mob.cell(row=r, column=13, value=teo)
                c_t.number_format, c_t.border = '0.00', border
                c_opt = ws_mob.cell(row=r, column=14, value=opt)
                c_opt.alignment, c_opt.font, c_opt.fill, c_opt.border = Alignment(horizontal="center"), bold_font, opt_fill, border
                c_or = ws_mob.cell(row=r, column=15, value=ore_reali)
                c_or.number_format, c_or.border, c_or.alignment = '0.00 "h"', border, Alignment(horizontal="center")
                c_ct = ws_mob.cell(row=r, column=16, value=costo_teo)
                c_ct.number_format, c_ct.border = '#,##0.00 €', border
                c_cr = ws_mob.cell(row=r, column=17, value=costo_reale)
                c_cr.number_format, c_cr.font, c_cr.fill, c_cr.border = '#,##0.00 €', bold_font, cost_fill, border
                r += 1

            # Subtotale Singola Settimana con Range Date
            slots_w = sub_14.index
            ws_mob.cell(row=r, column=2, value=f"SUBTOTALE {w_label.upper()}").font = bold_font
            ws_mob.cell(row=r, column=2).fill, ws_mob.cell(row=r, column=2).border = sub_fill, border
            ws_mob.cell(row=r, column=3, value=range_w_corr).font = bold_font
            ws_mob.cell(row=r, column=3).fill, ws_mob.cell(row=r, column=3).border, ws_mob.cell(row=r, column=3).alignment = sub_fill, border, Alignment(horizontal="center")
            ws_mob.cell(row=r, column=4, value="6 gg lav. su 7").font = bold_font
            ws_mob.cell(row=r, column=4).fill, ws_mob.cell(row=r, column=4).border = sub_fill, border

            for j, d in enumerate(self.dept_list, 5):
                c = ws_mob.cell(row=r, column=j, value=f"{int(self.piv_arr050_2w.loc[slots_w, d].sum())}  ({float(self.piv_fte_2w.loc[slots_w, d].sum()):.2f})")
                c.alignment, c.font, c.fill, c.border = Alignment(horizontal="center"), bold_font, sub_fill, border

            ws_mob.cell(row=r, column=10, value=int(sub_14['Somma_Addetti_Reparto_050'].sum())).font = bold_font
            ws_mob.cell(row=r, column=10).fill, ws_mob.cell(row=r, column=10).border, ws_mob.cell(row=r, column=10).alignment = sub_fill, border, Alignment(horizontal="center")

            ws_mob.cell(row=r, column=11, value=float(sub_14['Ore_Nette_Scoperte'].sum())).number_format = '0.00 "h"'
            ws_mob.cell(row=r, column=11).font, ws_mob.cell(row=r, column=11).fill, ws_mob.cell(row=r, column=11).border = bold_font, sub_fill, border

            ws_mob.cell(row=r, column=12, value=int(sub_14['Addetti_Jolly_Mobilita'].sum())).font = bold_font
            ws_mob.cell(row=r, column=12).fill, ws_mob.cell(row=r, column=12).border, ws_mob.cell(row=r, column=12).alignment = jolly_fill, border, Alignment(horizontal="center")

            ws_mob.cell(row=r, column=13, value=float(sub_14['Turni_Teorici_Totali_FTE'].sum())).number_format = '0.00'
            ws_mob.cell(row=r, column=13).font, ws_mob.cell(row=r, column=13).fill, ws_mob.cell(row=r, column=13).border = bold_font, sub_fill, border

            ws_mob.cell(row=r, column=14, value=int(sub_14['Squadra_Ottimale_Reale'].sum())).font = bold_font
            ws_mob.cell(row=r, column=14).fill, ws_mob.cell(row=r, column=14).border, ws_mob.cell(row=r, column=14).alignment = opt_fill, border, Alignment(horizontal="center")

            ws_mob.cell(row=r, column=15, value=float(sub_14['Ore_Reali_Squadra'].sum())).number_format = '0.00 "h"'
            ws_mob.cell(row=r, column=15).font, ws_mob.cell(row=r, column=15).fill, ws_mob.cell(row=r, column=15).border = bold_font, sub_fill, border

            ws_mob.cell(row=r, column=16, value=float(sub_14['Costo_Lavoro_Teorico_Euro'].sum())).number_format = '#,##0.00 €'
            ws_mob.cell(row=r, column=16).font, ws_mob.cell(row=r, column=16).fill, ws_mob.cell(row=r, column=16).border = bold_font, sub_fill, border

            ws_mob.cell(row=r, column=17, value=float(sub_14['Costo_Lavoro_Squadra_Reale_Euro'].sum())).number_format = '#,##0.00 €'
            ws_mob.cell(row=r, column=17).font, ws_mob.cell(row=r, column=17).fill, ws_mob.cell(row=r, column=17).border = bold_font, cost_fill, border
            r += 1

        # Riga TOTALE GENERALE BISETTIMANALE (2 SETTIMANE CONSECUTIVE - 14 GIORNI)
        ws_mob.cell(row=r, column=2, value="TOTALE 2 SETTIMANE (RILASCIO TURNI)").font = bold_font
        ws_mob.cell(row=r, column=2).fill, ws_mob.cell(row=r, column=2).border = opt_fill, border
        ws_mob.cell(row=r, column=3, value=self.range_2w_str).font = bold_font
        ws_mob.cell(row=r, column=3).fill, ws_mob.cell(row=r, column=3).border, ws_mob.cell(row=r, column=3).alignment = opt_fill, border, Alignment(horizontal="center")
        ws_mob.cell(row=r, column=4, value="12 gg lav. su 14").font = bold_font
        ws_mob.cell(row=r, column=4).fill, ws_mob.cell(row=r, column=4).border = opt_fill, border

        for j, d in enumerate(self.dept_list, 5):
            c = ws_mob.cell(row=r, column=j, value=f"{int(self.piv_arr050_2w[d].sum())}  ({float(self.piv_fte_2w[d].sum()):.2f})")
            c.alignment, c.font, c.fill, c.border = Alignment(horizontal="center"), bold_font, opt_fill, border

        ws_mob.cell(row=r, column=10, value=int(self.riepilogo_14gg['Somma_Addetti_Reparto_050'].sum())).font = bold_font
        ws_mob.cell(row=r, column=10).fill, ws_mob.cell(row=r, column=10).border, ws_mob.cell(row=r, column=10).alignment = opt_fill, border, Alignment(horizontal="center")

        ws_mob.cell(row=r, column=11, value=float(self.riepilogo_14gg['Ore_Nette_Scoperte'].sum())).number_format = '0.00 "h"'
        ws_mob.cell(row=r, column=11).font, ws_mob.cell(row=r, column=11).fill, ws_mob.cell(row=r, column=11).border = bold_font, opt_fill, border

        ws_mob.cell(row=r, column=12, value=int(self.riepilogo_14gg['Addetti_Jolly_Mobilita'].sum())).font = bold_font
        ws_mob.cell(row=r, column=12).fill, ws_mob.cell(row=r, column=12).border, ws_mob.cell(row=r, column=12).alignment = jolly_fill, border, Alignment(horizontal="center")

        ws_mob.cell(row=r, column=13, value=float(self.riepilogo_14gg['Turni_Teorici_Totali_FTE'].sum())).number_format = '0.00'
        ws_mob.cell(row=r, column=13).font, ws_mob.cell(row=r, column=13).fill, ws_mob.cell(row=r, column=13).border = bold_font, opt_fill, border

        ws_mob.cell(row=r, column=14, value=int(self.riepilogo_14gg['Squadra_Ottimale_Reale'].sum())).font = bold_font
        ws_mob.cell(row=r, column=14).fill, ws_mob.cell(row=r, column=14).border, ws_mob.cell(row=r, column=14).alignment = opt_fill, border, Alignment(horizontal="center")

        ws_mob.cell(row=r, column=15, value=float(self.riepilogo_14gg['Ore_Reali_Squadra'].sum())).number_format = '0.00 "h"'
        ws_mob.cell(row=r, column=15).font, ws_mob.cell(row=r, column=15).fill, ws_mob.cell(row=r, column=15).border = bold_font, opt_fill, border

        ws_mob.cell(row=r, column=16, value=float(self.riepilogo_14gg['Costo_Lavoro_Teorico_Euro'].sum())).number_format = '#,##0.00 €'
        ws_mob.cell(row=r, column=16).font, ws_mob.cell(row=r, column=16).fill, ws_mob.cell(row=r, column=16).border = bold_font, opt_fill, border

        ws_mob.cell(row=r, column=17, value=float(self.riepilogo_14gg['Costo_Lavoro_Squadra_Reale_Euro'].sum())).number_format = '#,##0.00 €'
        ws_mob.cell(row=r, column=17).font, ws_mob.cell(row=r, column=17).fill, ws_mob.cell(row=r, column=17).border = bold_font, cost_fill, border

        # ======================================================================
        # FOGLIO 2: SINTESI 1 SETTIMANA, 2 SETTIMANE (CON DATE), MESE E TRIMESTRE
        # ======================================================================
        ws_orizzonti.cell(
            row=2, column=2,
            value=f"SINTESI MULTI-ORIZZONTE: SETTIMANA 1 ({self.range_w1_str}), 2 SETTIMANE ({self.range_2w_str}), MESE E TRIMESTRE"
        ).font = title_font
        headers_hz = [
            "Sottoreparto / Voce", "Dotazione Giornaliera (Soglia 0,50)",
            f"Sett. 1 ({self.range_w1_str}): Colli", "Sett. 1: Turni Reali", "Sett. 1: Costo Reale (€)",
            f"2 Settimane ({self.range_2w_str}): Colli", "2 Settimane: Turni Teorici", "2 Settimane: Turni Reali", "2 Settimane: Costo Reale (€)",
            "Mese (30gg): Turni Reali", "Mese (30gg): Costo Reale (€)",
            "Trimestre (13 Sett.): Turni Reali", "Trimestre: Costo Reale (€)"
        ]
        for i, h in enumerate(headers_hz, 2):
            c = ws_orizzonti.cell(row=4, column=i, value=h)
            c.fill, c.font, c.alignment = header_fill, white_font, Alignment(horizontal="center", wrap_text=True)

        r = 5
        for d in self.dept_list:
            ps = self.prod_2w.loc[d]
            sub_d = self.df_staff_2w[(self.df_staff_2w['Sottoreparto'] == d) & (self.df_staff_2w['Giorno_Num'] < 6)]
            min_d = int(sub_d['Addetti_Soglia_050'].min())
            max_d = int(sub_d['Addetti_Soglia_050'].max())
            dot_str = f"{min_d} - {max_d} addetti" if min_d != max_d else f"{min_d} addetti"

            colli_1w = round(float(ps['Colli_Sett_1']))
            turni_1w = int(ps['Turni_050_Sett_1'])
            costo_1w = float(ps['Costo_050_Sett_1_Euro'])

            colli_2w = round(float(ps['Colli_Bisettimanali']))
            teo_2w = round(float(ps['Turni_Teorici_2W']), 2)
            turni_2w = int(ps['Turni_Presidio_050_2W'])
            costo_2w = float(ps['Costo_Presidio_050_2W_Euro'])

            turni_m = round(turni_2w * (30.0 / 14.0))
            costo_m = round(turni_m * self.costo_turno, 2)
            turni_t = round(turni_2w * (13.0 / 2.0))
            costo_t = round(turni_t * self.costo_turno, 2)

            vals = [
                d, dot_str,
                colli_1w, turni_1w, costo_1w,
                colli_2w, teo_2w, turni_2w, costo_2w,
                turni_m, costo_m,
                turni_t, costo_t
            ]
            for col_i, v in enumerate(vals, 2):
                c = ws_orizzonti.cell(row=r, column=col_i, value=v)
                c.border = border
                if col_i in [7, 8, 9, 10]:
                    c.fill = w2_fill
                    if col_i in [9, 10]:
                        c.font = bold_font
                if col_i in [6, 10, 12, 14]:
                    c.number_format = '#,##0.00 €'
            r += 1

        jolly_1w = int(self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == 1]['Addetti_Jolly_Mobilita'].sum())
        costo_j_1w = round(jolly_1w * self.costo_turno, 2)
        jolly_2w = int(self.riepilogo_14gg['Addetti_Jolly_Mobilita'].sum())
        teo_res_2w = round(self.riepilogo_14gg['Ore_Nette_Scoperte'].sum() / self.ore_turno, 2)
        costo_j_2w = round(jolly_2w * self.costo_turno, 2)
        jolly_m = round(jolly_2w * (30.0 / 14.0))
        costo_j_m = round(jolly_m * self.costo_turno, 2)
        jolly_t = round(jolly_2w * (13.0 / 2.0))
        costo_j_t = round(jolly_t * self.costo_turno, 2)

        vals_j = [
            "Integrazione Jolly Mobilità Interna (Resti < 0,50)", "1 - 2 Jolly/giorno",
            "-", jolly_1w, costo_j_1w,
            "-", teo_res_2w, jolly_2w, costo_j_2w,
            jolly_m, costo_j_m,
            jolly_t, costo_j_t
        ]
        for col_i, v in enumerate(vals_j, 2):
            c = ws_orizzonti.cell(row=r, column=col_i, value=v)
            c.fill, c.font, c.border = jolly_fill, bold_font, border
            if col_i in [6, 10, 12, 14]:
                c.number_format = '#,##0.00 €'
        r += 1

        colli_tot_1w = round(self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == 1]['Totale_Colli_Giorno'].sum())
        turni_tot_1w = int(self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == 1]['Squadra_Ottimale_Reale'].sum())
        costo_tot_1w = round(turni_tot_1w * self.costo_turno, 2)

        colli_tot_2w = round(self.riepilogo_14gg['Totale_Colli_Giorno'].sum())
        teo_tot_2w = round(float(self.riepilogo_14gg['Turni_Teorici_Totali_FTE'].sum()), 2)
        turni_tot_2w = int(self.riepilogo_14gg['Squadra_Ottimale_Reale'].sum())
        costo_tot_2w = round(turni_tot_2w * self.costo_turno, 2)

        turni_tot_m = round(turni_tot_2w * (30.0 / 14.0))
        costo_tot_m = round(turni_tot_m * self.costo_turno, 2)
        turni_tot_t = round(turni_tot_2w * (13.0 / 2.0))
        costo_tot_t = round(turni_tot_t * self.costo_turno, 2)

        vals_tot = [
            "TOTALE SQUADRA OTTIMALE", "12 (Lun-Gio) / 11 (Ven-Sab)",
            colli_tot_1w, turni_tot_1w, costo_tot_1w,
            colli_tot_2w, teo_tot_2w, turni_tot_2w, costo_tot_2w,
            turni_tot_m, costo_tot_m,
            turni_tot_t, costo_tot_t
        ]
        for col_i, v in enumerate(vals_tot, 2):
            c = ws_orizzonti.cell(row=r, column=col_i, value=v)
            c.fill, c.font, c.border = opt_fill, bold_font, border
            if col_i in [6, 10, 12, 14]:
                c.number_format = '#,##0.00 €'

        # ======================================================================
        # FOGLIO 3: PREVISIONE COLLI EFFETTIVI SU 2 SETTIMANE CON DATE REALI
        # ======================================================================
        ws_colli.cell(
            row=2, column=2,
            value=f"PREVISIONE BISETTIMANALE ({self.range_2w_str}) COLLI EFFETTIVI DA ALLESTIRE (DA COLONNA L)"
        ).font = title_font
        headers_c = ["Settimana", "Data", "Giorno"] + self.dept_list + ["TOTALE COLLI GIORNO"]
        for i, h in enumerate(headers_c, 2):
            c = ws_colli.cell(row=4, column=i, value=h)
            c.fill, c.font, c.alignment = header_fill, white_font, Alignment(horizontal="center")

        r = 5
        for w_idx, w_label in enumerate(self.settimane_list, start=1):
            sub_14 = self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == w_idx]
            range_w_corr = self.range_w1_str if w_idx == 1 else self.range_w2_str
            for slot_idx, row_g in sub_14.iterrows():
                ws_colli.cell(row=r, column=2, value=w_label).font = bold_font
                ws_colli.cell(row=r, column=2).border = border
                c_dt = ws_colli.cell(row=r, column=3, value=row_g['Data'])
                c_dt.font, c_dt.alignment, c_dt.border = bold_font, Alignment(horizontal="center"), border
                ws_colli.cell(row=r, column=4, value=row_g['Giorno']).font = bold_font
                ws_colli.cell(row=r, column=4).border = border
                for j, d in enumerate(self.dept_list, 5):
                    c = ws_colli.cell(row=r, column=j, value=float(self.piv_colli_2w.loc[slot_idx, d]))
                    c.number_format, c.border = '#,##0.0', border
                ct = ws_colli.cell(row=r, column=10, value=float(self.piv_colli_2w.loc[slot_idx, 'Totale_Colli_L']))
                ct.number_format, ct.font, ct.fill, ct.border = '#,##0.0', bold_font, sub_fill, border
                r += 1

            # Subtotale Settimana Colli
            slots_w = sub_14.index
            ws_colli.cell(row=r, column=2, value=f"SUBTOTALE {w_label.upper()}").font = bold_font
            ws_colli.cell(row=r, column=2).fill, ws_colli.cell(row=r, column=2).border = sub_fill, border
            ws_colli.cell(row=r, column=3, value=range_w_corr).font = bold_font
            ws_colli.cell(row=r, column=3).fill, ws_colli.cell(row=r, column=3).border, ws_colli.cell(row=r, column=3).alignment = sub_fill, border, Alignment(horizontal="center")
            ws_colli.cell(row=r, column=4, value="6 gg lav.").font = bold_font
            ws_colli.cell(row=r, column=4).fill, ws_colli.cell(row=r, column=4).border = sub_fill, border
            for j, d in enumerate(self.dept_list, 5):
                c = ws_colli.cell(row=r, column=j, value=float(self.piv_colli_2w.loc[slots_w, d].sum()))
                c.number_format, c.font, c.fill, c.border = '#,##0.0', bold_font, sub_fill, border
            ct_w = ws_colli.cell(row=r, column=10, value=float(self.piv_colli_2w.loc[slots_w, 'Totale_Colli_L'].sum()))
            ct_w.number_format, ct_w.font, ct_w.fill, ct_w.border = '#,##0.0', bold_font, sub_fill, border
            r += 1

        # Totale 2 Settimane Colli
        ws_colli.cell(row=r, column=2, value="TOTALE 2 SETTIMANE (14 GG)").font = bold_font
        ws_colli.cell(row=r, column=2).fill, ws_colli.cell(row=r, column=2).border = opt_fill, border
        ws_colli.cell(row=r, column=3, value=self.range_2w_str).font = bold_font
        ws_colli.cell(row=r, column=3).fill, ws_colli.cell(row=r, column=3).border, ws_colli.cell(row=r, column=3).alignment = opt_fill, border, Alignment(horizontal="center")
        ws_colli.cell(row=r, column=4, value="12 gg lav.").font = bold_font
        ws_colli.cell(row=r, column=4).fill, ws_colli.cell(row=r, column=4).border = opt_fill, border
        for j, d in enumerate(self.dept_list, 5):
            c = ws_colli.cell(row=r, column=j, value=float(self.piv_colli_2w[d].sum()))
            c.number_format, c.font, c.fill, c.border = '#,##0.0', bold_font, opt_fill, border
        ct_tot = ws_colli.cell(row=r, column=10, value=float(self.piv_colli_2w['Totale_Colli_L'].sum()))
        ct_tot.number_format, ct_tot.font, ct_tot.fill, ct_tot.border = '#,##0.0', bold_font, opt_fill, border

        # ======================================================================
        # FOGLIO 4: DETTAGLIO GIORNALIERO SUI 14 GIORNI CON DATA E DATA ESTESA
        # ======================================================================
        ws_bilancio.cell(
            row=2, column=2,
            value=f"DETTAGLIO GIORNALIERO BISETTIMANALE ({self.range_2w_str}) PER SOTTOREPARTO: ORE, PRODUTTIVITÀ E COSTI"
        ).font = title_font
        headers_b = [
            "Settimana", "Data", "Giorno", "Data Estesa", "Sottoreparto", "Colli Previsti (Col. L)",
            "Resa Target (c/h)", "Capacità Target Turno (c/3h)",
            "Ore Richieste", "FTE Teorici (3h)", "Addetti (Soglia 0,50)", "Ore Presidiate (Soglia 0,50)",
            "Bilancio Ore (+ Deficit / - Avanzo)",
            "Produttività Effettiva (c/h su Presidio)", "Produttività Effettiva (c/turno su Presidio)",
            "Costo Lavoro Teorico (€)", "Costo Lavoro Presidio 0,50 (€)", "Costo per Collo Teorico (€)"
        ]
        for i, h in enumerate(headers_b, 2):
            c = ws_bilancio.cell(row=4, column=i, value=h)
            c.fill, c.font, c.alignment = header_fill, white_font, Alignment(horizontal="center", wrap_text=True)

        r = 5
        for _, row_s in self.df_staff_2w[self.df_staff_2w['Giorno_Num'] < 6].iterrows():
            vals = [
                row_s['Settimana'], row_s['Data'], row_s['Giorno'], row_s['Data_Estesa'],
                row_s['Sottoreparto'], row_s['Colli_L'],
                int(row_s['Resa_Target_Oraria']), int(row_s['Capacita_Target_Turno_3h']),
                row_s['Ore_Richieste'], row_s['FTE_Teorici_3h'],
                int(row_s['Addetti_Soglia_050']), row_s['Ore_Presidiate_050'],
                row_s['Bilancio_Ore_Mobilita'],
                row_s['Prod_Effettiva_Colli_Ora_050'], row_s['Prod_Effettiva_Colli_Turno_050'],
                row_s['Costo_Teorico_Euro'], row_s['Costo_Presidio_050_Euro'], row_s['Costo_Unitario_Collo_Teorico']
            ]
            for col_i, v in enumerate(vals, 2):
                c = ws_bilancio.cell(row=r, column=col_i, value=v)
                c.border = border
                if col_i == 3:
                    c.alignment = Alignment(horizontal="center")
                    c.font = bold_font
                elif col_i in [17, 18]:
                    c.number_format = '#,##0.00 €'
                elif col_i == 19:
                    c.number_format = '0.0000 €'
            r += 1

        # ======================================================================
        # FOGLIO 5: PRODUTTIVITÀ E COSTO DEL LAVORO BISETTIMANALE CON DATE
        # ======================================================================
        ws_prod.cell(
            row=2, column=2,
            value=f"ANALISI PRODUTTIVITÀ E COSTO DEL LAVORO SU 2 SETTIMANE CONSECUTIVE ({self.range_2w_str} - € 12,80/ORA)"
        ).font = title_font
        ws_prod.cell(
            row=3, column=2,
            value=f"Sezione A: Riepilogo Settimana 1 ({self.range_w1_str}), Settimana 2 ({self.range_w2_str}) e Totale 2 Settimane | Sezione B: Matrice Giornaliera con Date"
        ).font = Font(italic=True, size=10)

        headers_p1 = [
            "Sottoreparto / Voce",
            "Resa Target (c/h)",
            "Media Colli / Giorno Lav.",
            f"Sett. 1 ({self.range_w1_str}): Colli",
            "Sett. 1: Turni (0,50)",
            "Sett. 1: Costo (€)",
            f"Sett. 2 ({self.range_w2_str}): Colli",
            "Sett. 2: Turni (0,50)",
            "Sett. 2: Costo (€)",
            f"TOTALE 2 SETT. ({self.range_2w_str}): Colli",
            "TOTALE 2 SETT.: Ore Richieste",
            "TOTALE 2 SETT.: Turni Teorici",
            "TOTALE 2 SETT.: Turni Reali",
            "TOTALE 2 SETT.: Ore Presidiate",
            "Prod. Effettiva (Colli/h)",
            "Prod. Effettiva (Colli/Turno)",
            "Saturazione Presidio (%)",
            "TOTALE 2 SETT.: Costo Teorico (€)",
            "TOTALE 2 SETT.: Costo Reale (€)",
            "Costo Medio per Collo (€/collo)"
        ]
        for i, h in enumerate(headers_p1, 2):
            c = ws_prod.cell(row=5, column=i, value=h)
            c.fill, c.font, c.alignment = header_fill, white_font, Alignment(horizontal="center", wrap_text=True)

        r = 6
        for d in self.dept_list:
            ps = self.prod_2w.loc[d]
            vals = [
                d,
                int(ps['Resa_Target_Oraria']),
                float(ps['Media_Colli_Giorno_Lav']),
                float(ps['Colli_Sett_1']),
                int(ps['Turni_050_Sett_1']),
                float(ps['Costo_050_Sett_1_Euro']),
                float(ps['Colli_Sett_2']),
                int(ps['Turni_050_Sett_2']),
                float(ps['Costo_050_Sett_2_Euro']),
                float(ps['Colli_Bisettimanali']),
                float(ps['Ore_Richieste_2W']),
                round(float(ps['Turni_Teorici_2W']), 2),
                int(ps['Turni_Presidio_050_2W']),
                float(ps['Ore_Presidio_050_2W']),
                float(ps['Prod_Effettiva_Oraria_050']),
                float(ps['Prod_Effettiva_Turno_050']),
                float(ps['Saturazione_Media_Presidio_Pct']) / 100.0,
                float(ps['Costo_Teorico_2W_Euro']),
                float(ps['Costo_Presidio_050_2W_Euro']),
                float(ps['Costo_per_Collo_Teorico_Euro'])
            ]
            for col_i, v in enumerate(vals, 2):
                c = ws_prod.cell(row=r, column=col_i, value=v)
                c.border = border
                if col_i == 2:
                    c.font = bold_font
                elif col_i in [4, 5, 8, 11, 17]:
                    c.number_format = '#,##0.0'
                elif col_i in [12, 13, 15, 16]:
                    c.number_format = '0.00'
                elif col_i == 18:
                    c.number_format = '0.0%'
                elif col_i in [7, 10, 19, 20]:
                    c.number_format = '#,##0.00 €'
                    if col_i == 20:
                        c.fill = cost_fill
                        c.font = bold_font
                elif col_i == 21:
                    c.number_format = '0.0000 €'
            r += 1

        # Riga Jolly Mobilità Interna nel Foglio 5
        ore_scoperte_2w = float(self.riepilogo_14gg['Ore_Nette_Scoperte'].sum())
        turni_j_w1 = int(self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == 1]['Addetti_Jolly_Mobilita'].sum())
        costo_j_w1 = round(turni_j_w1 * self.costo_turno, 2)
        turni_j_w2 = int(self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == 2]['Addetti_Jolly_Mobilita'].sum())
        costo_j_w2 = round(turni_j_w2 * self.costo_turno, 2)
        turni_j_2w = turni_j_w1 + turni_j_w2
        ore_j_2w = round(turni_j_2w * self.ore_turno, 2)
        costo_j_2w = round(ore_j_2w * self.costo_orario, 2)

        vals_j_p = [
            "JOLLY MOBILITÀ INTERNA (Copertura Resti)",
            "-", "-",
            "-", turni_j_w1, costo_j_w1,
            "-", turni_j_w2, costo_j_w2,
            "-", round(ore_scoperte_2w, 2), round(ore_scoperte_2w / self.ore_turno, 2),
            turni_j_2w, ore_j_2w,
            "-", "-",
            round(ore_scoperte_2w / ore_j_2w, 3) if ore_j_2w > 0 else 0.0,
            "-", costo_j_2w, "-"
        ]
        for col_i, v in enumerate(vals_j_p, 2):
            c = ws_prod.cell(row=r, column=col_i, value=v)
            c.fill, c.font, c.border = jolly_fill, bold_font, border
            if col_i == 18 and isinstance(v, float):
                c.number_format = '0.0%'
            elif col_i in [7, 10, 20]:
                c.number_format = '#,##0.00 €'
        r += 1

        # Riga Totale Squadra Ottimale su 2 Settimane nel Foglio 5
        tot_colli_w1 = float(self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == 1]['Totale_Colli_Giorno'].sum())
        tot_turni_w1 = int(self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == 1]['Squadra_Ottimale_Reale'].sum())
        tot_costo_w1 = float(self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == 1]['Costo_Lavoro_Squadra_Reale_Euro'].sum())

        tot_colli_w2 = float(self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == 2]['Totale_Colli_Giorno'].sum())
        tot_turni_w2 = int(self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == 2]['Squadra_Ottimale_Reale'].sum())
        tot_costo_w2 = float(self.riepilogo_14gg[self.riepilogo_14gg['Settimana_Num'] == 2]['Costo_Lavoro_Squadra_Reale_Euro'].sum())

        tot_colli_2w = float(self.prod_2w['Colli_Bisettimanali'].sum())
        tot_ore_ric_2w = float(self.prod_2w['Ore_Richieste_2W'].sum())
        tot_fte_2w = float(self.prod_2w['Turni_Teorici_2W'].sum())
        tot_turni_reali_2w = int(self.riepilogo_14gg['Squadra_Ottimale_Reale'].sum())
        tot_ore_reali_2w = float(self.riepilogo_14gg['Ore_Reali_Squadra'].sum())
        prod_media_oraria_2w = round(tot_colli_2w / tot_ore_reali_2w, 2) if tot_ore_reali_2w > 0 else 0.0
        prod_media_turno_2w = round(tot_colli_2w / tot_turni_reali_2w, 1) if tot_turni_reali_2w > 0 else 0.0
        sat_tot_2w = (tot_ore_ric_2w / tot_ore_reali_2w) if tot_ore_reali_2w > 0 else 0.0
        tot_costo_teo_2w = float(self.prod_2w['Costo_Teorico_2W_Euro'].sum())
        tot_costo_reale_2w = float(self.riepilogo_14gg['Costo_Lavoro_Squadra_Reale_Euro'].sum())
        costo_collo_reale_2w = round(tot_costo_reale_2w / tot_colli_2w, 4) if tot_colli_2w > 0 else 0.0

        vals_tot_p = [
            f"TOTALE SQUADRA GROCERY ({self.range_2w_str})",
            round(tot_colli_2w / tot_ore_ric_2w, 2),
            round(tot_colli_2w / 12.0, 1),
            tot_colli_w1, tot_turni_w1, tot_costo_w1,
            tot_colli_w2, tot_turni_w2, tot_costo_w2,
            tot_colli_2w, tot_ore_ric_2w, round(tot_fte_2w, 2),
            tot_turni_reali_2w, tot_ore_reali_2w,
            prod_media_oraria_2w, prod_media_turno_2w,
            sat_tot_2w,
            tot_costo_teo_2w, tot_costo_reale_2w, costo_collo_reale_2w
        ]
        for col_i, v in enumerate(vals_tot_p, 2):
            c = ws_prod.cell(row=r, column=col_i, value=v)
            c.fill, c.font, c.border = opt_fill, bold_font, border
            if col_i in [4, 5, 8, 11, 17]:
                c.number_format = '#,##0.0'
            elif col_i in [3, 12, 13, 15, 16]:
                c.number_format = '0.00'
            elif col_i == 18:
                c.number_format = '0.0%'
            elif col_i in [7, 10, 19, 20]:
                c.number_format = '#,##0.00 €'
            elif col_i == 21:
                c.number_format = '0.0000 €'

        # Sezione B nel Foglio 5: Matrice Giornaliera sui 14 Giorni con Date Reali
        r += 3
        ws_prod.cell(
            row=r, column=2,
            value=f"MATRICE GIORNALIERA BISETTIMANALE ({self.range_2w_str}): PRODUTTIVITÀ EFFETTIVA (COLLI/ORA) E COSTO LAVORO PER SOTTOREPARTO"
        ).font = title_font
        r += 1
        headers_p2 = ["Settimana", "Data", "Giorno"] + [f"{d} (Colli | c/h eff. | Costo 0,50)" for d in self.dept_list] + [
            "Costo Presidi 0,50 (€)", "Costo Jolly Mobilità (€)", "COSTO TOTALE SQUADRA REALE (€)", "Prod. Media Squadra (c/h)"
        ]
        for i, h in enumerate(headers_p2, 2):
            c = ws_prod.cell(row=r, column=i, value=h)
            c.fill, c.font, c.alignment = header_fill, white_font, Alignment(horizontal="center", wrap_text=True)

        r += 1
        for slot_idx, row_g in self.riepilogo_14gg.iterrows():
            ws_prod.cell(row=r, column=2, value=row_g['Settimana']).font = bold_font
            ws_prod.cell(row=r, column=2).border = border
            c_dt = ws_prod.cell(row=r, column=3, value=row_g['Data'])
            c_dt.font, c_dt.alignment, c_dt.border = bold_font, Alignment(horizontal="center"), border
            ws_prod.cell(row=r, column=4, value=row_g['Giorno']).font = bold_font
            ws_prod.cell(row=r, column=4).border = border
            for j, d in enumerate(self.dept_list, 5):
                colli_g = float(self.piv_colli_2w.loc[slot_idx, d])
                prod_h_g = float(self.piv_prod_ora_2w.loc[slot_idx, d])
                costo_050_g = float(self.piv_costo_050_2w.loc[slot_idx, d])
                val_str = f"{colli_g:.0f} colli | {prod_h_g:.1f} c/h | € {costo_050_g:.2f}" if colli_g > 0 else "-"
                c = ws_prod.cell(row=r, column=j, value=val_str)
                c.alignment, c.border, c.font = Alignment(horizontal="center"), border, reg_font

            c_p050 = ws_prod.cell(row=r, column=10, value=float(row_g['Costo_Lavoro_Presidio_050_Euro']))
            c_p050.number_format, c_p050.border = '#,##0.00 €', border
            c_pj = ws_prod.cell(row=r, column=11, value=float(row_g['Costo_Lavoro_Jolly_Euro']))
            c_pj.number_format, c_pj.fill, c_pj.border = '#,##0.00 €', jolly_fill, border
            c_ptot = ws_prod.cell(row=r, column=12, value=float(row_g['Costo_Lavoro_Squadra_Reale_Euro']))
            c_ptot.number_format, c_ptot.font, c_ptot.fill, c_ptot.border = '#,##0.00 €', bold_font, cost_fill, border
            c_pmed = ws_prod.cell(row=r, column=13, value=float(row_g['Produttivita_Media_Squadra_Colli_Ora']))
            c_pmed.number_format, c_pmed.font, c_pmed.border, c_pmed.alignment = '0.00 "c/h"', bold_font, border, Alignment(horizontal="center")
            r += 1

        # Adattamento larghezza colonne su tutti i fogli
        for ws in [ws_mob, ws_orizzonti, ws_colli, ws_bilancio, ws_prod]:
            for col in ws.columns:
                max_len = max(len(str(c.value or '')) for c in col)
                ws.column_dimensions[get_column_letter(col[0].column)].width = max(min(max_len + 4, 34), 14)

        wb.save(self.output_excel_path)
        print(f"      -> Report Excel salvato in: {self.output_excel_path}")

    # --------------------------------------------------------------------------
    # 5. ESPORTAZIONE MODELLI MACHINE LEARNING (.JOBLIB) PER WEB APP
    # --------------------------------------------------------------------------
    def esporta_modelli_ml(self, file_path=FILE_MODELLO_ML):
        print(f"[5/5] Esportazione modelli Machine Learning in: {file_path}")
        pacchetto_ml = {
            'clf_tagli': self.clf_tagli,
            'reg_colli': self.reg_colli,
            'rese_orarie': self.rese_orarie,
            'costo_orario': self.costo_orario,
            'ore_turno': self.ore_turno,
            'dept_cat_map': self.dept_cat_map,
            'range_storico': self.range_storico_str,
            'range_previsione_2w': self.range_2w_str
        }
        joblib.dump(pacchetto_ml, file_path)
        print(f"      -> Modello ML esportato con successo!")
        return file_path


# ==============================================================================
# AVVIO AUTOMATICO DELLA PIPELINE COMPLETA
# ==============================================================================
if __name__ == "__main__":
    sistema = GroceryForecastingUnifiedSystem(
        input_excel_path=FILE_EXCEL_INPUT,
        output_excel_path=FILE_EXCEL_OUTPUT
    )
    sistema.carica_e_pulisci_dati()
    sistema.addestra_modello_tagli()
    df_dettaglio_2w, df_riepilogo_14gg, df_prod_2w = sistema.addestra_e_prevedi_due_settimane()
    sistema.esporta_excel_completo()
    sistema.esporta_modelli_ml()

    print(f"\n--- 1. RIEPILOGO TURNI BISETTIMANALI CON DATE DINAMICHE ({sistema.range_2w_str}) ---")
    cols_14gg = [
        'Settimana', 'Data', 'Giorno', 'Totale_Colli_Giorno',
        'Somma_Addetti_Reparto_050', 'Addetti_Jolly_Mobilita',
        'Squadra_Ottimale_Reale', 'Produttivita_Media_Squadra_Colli_Ora',
        'Costo_Lavoro_Squadra_Reale_Euro'
    ]
    print(df_riepilogo_14gg[cols_14gg].to_string(index=False))

    print(f"\n--- 2. PRODUTTIVITÀ E COSTO DEL LAVORO BISETTIMANALE ({sistema.range_w1_str} + {sistema.range_w2_str}) ---")
    cols_2w = [
        'Resa_Target_Oraria', 'Colli_Sett_1', 'Turni_050_Sett_1',
        'Colli_Sett_2', 'Turni_050_Sett_2', 'Colli_Bisettimanali',
        'Turni_Presidio_050_2W', 'Prod_Effettiva_Oraria_050',
        'Costo_Teorico_2W_Euro', 'Costo_Presidio_050_2W_Euro'
    ]
    print(df_prod_2w[cols_2w].to_string())