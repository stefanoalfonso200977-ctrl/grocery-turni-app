import streamlit as st
import pandas as pd
import tempfile
import os
import joblib
from io import BytesIO

# Importiamo la tua classe dal file del modello v3
from modello_grocery_costi_produttivita_v3 import GroceryForecastingUnifiedSystem

# ==============================================================================
# CONFIGURAZIONE PAGINA WEB APP
# ==============================================================================
st.set_page_config(
    page_title="Pianificazione Turni & Costi Grocery - AI",
    page_icon="🛒",
    layout="wide"
)

st.title("🛒 Sistema Previsionale Bisettimanale Colli, Turni e Costo del Lavoro")
st.markdown(
    "**Modello Gradient Boosting a 2 Stadi** | Pianificazione Rilascio Turni su **2 Settimane Consecutive (14 gg)** "
    "con aggiornamento automatico delle date dal file Excel caricato."
)

# ==============================================================================
# SIDEBAR: PARAMETRI OPERATIVI E SIMULAZIONE WHAT-IF
# ==============================================================================
st.sidebar.header("⚙️ Parametri Operativi & Costi")

costo_orario_input = st.sidebar.number_input(
    "Costo Orario 4° Livello (€/h)", min_value=8.0, max_value=25.0, value=12.80, step=0.10, format="%.2f"
)
ore_turno_input = st.sidebar.number_input(
    "Durata Turno Allestitori (ore)", min_value=2.0, max_value=8.0, value=3.0, step=0.5
)

st.sidebar.subheader("📦 Rese Orarie Target (Colli/h)")
resa_alim = st.sidebar.number_input("ALIMENTARI (c/h)", min_value=15.0, max_value=80.0, value=40.0, step=1.0)
resa_chim = st.sidebar.number_input("CHIMICA (c/h)", min_value=15.0, max_value=80.0, value=45.0, step=1.0)
resa_liq = st.sidebar.number_input("LIQUIDI (c/h)", min_value=15.0, max_value=80.0, value=35.0, step=1.0)
resa_pet = st.sidebar.number_input("PETFOOD (c/h)", min_value=15.0, max_value=80.0, value=40.0, step=1.0)
resa_prof = st.sidebar.number_input("PROFUMERIA (c/h)", min_value=15.0, max_value=80.0, value=30.0, step=1.0)

# ==============================================================================
# CARICAMENTO FILE EXCEL DISTINTA ORDINI
# ==============================================================================
uploaded_file = st.file_uploader(
    "📂 Carica l'ultimo file Excel delle distinte ordini",
    type=["xlsx"]
)

if uploaded_file is not None:
    with st.spinner("Elaborazione Machine Learning e calcolo calendario bisettimanale in corso..."):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
            tmp_in.write(uploaded_file.getbuffer())
            tmp_input_path = tmp_in.name

        tmp_output_path = os.path.join(tempfile.gettempdir(), "Report_Previsionale_Bisettimanale_Grocery.xlsx")

        # Inizializzazione della classe con i parametri della sidebar
        sistema = GroceryForecastingUnifiedSystem(
            input_excel_path=tmp_input_path,
            output_excel_path=tmp_output_path
        )
        sistema.costo_orario = costo_orario_input
        sistema.ore_turno = ore_turno_input
        sistema.costo_turno = round(sistema.ore_turno * sistema.costo_orario, 2)
        sistema.rese_orarie = {
            'ALIMENTARI': resa_alim,
            'CHIMICA': resa_chim,
            'LIQUIDI': resa_liq,
            'PETFOOD': resa_pet,
            'PROFUMERIA': resa_prof
        }

        # Esecuzione pipeline
        sistema.carica_e_pulisci_dati()
        sistema.addestra_modello_tagli()
        df_dettaglio_2w, df_riepilogo_14gg, df_prod_2w = sistema.addestra_e_prevedi_due_settimane()
        sistema.esporta_excel_completo()

    # ==========================================================================
    # KPI CARDS E RIEPILOGO DIREZIONALE
    # ==========================================================================
    st.success(
        f"✅ **Storico Rilevato nel File:** `{sistema.range_storico_str}`  |  "
        f"📅 **Settimana 1:** `{sistema.range_w1_str}`  |  "
        f"📅 **Settimana 2:** `{sistema.range_w2_str}`"
    )

    tot_colli_2w = df_riepilogo_14gg['Totale_Colli_Giorno'].sum()
    tot_turni_reali_2w = int(df_riepilogo_14gg['Squadra_Ottimale_Reale'].sum())
    tot_fte_teo_2w = df_riepilogo_14gg['Turni_Teorici_Totali_FTE'].sum()
    tot_ore_reali_2w = df_riepilogo_14gg['Ore_Reali_Squadra'].sum()
    tot_costo_reale_2w = df_riepilogo_14gg['Costo_Lavoro_Squadra_Reale_Euro'].sum()
    tot_costo_teo_2w = df_riepilogo_14gg['Costo_Lavoro_Teorico_Euro'].sum()
    prod_media_squadra = round(tot_colli_2w / tot_ore_reali_2w, 2) if tot_ore_reali_2w > 0 else 0.0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("📦 Colli Previsti (2 Sett.)", f"{tot_colli_2w:,.0f} colli", f"Media: {tot_colli_2w/12:,.0f} colli/gg lav.")
    col2.metric("👥 Turni Rilasciati (14 gg)", f"{tot_turni_reali_2w} turni ({tot_ore_reali_2w:.0f} h)", f"Teorici: {tot_fte_teo_2w:.2f} FTE")
    col3.metric("💶 Budget Costo Lavoro (2 Sett.)", f"€ {tot_costo_reale_2w:,.2f}", f"Teorico: € {tot_costo_teo_2w:,.2f}")
    col4.metric("⚡ Produttività Media Squadra", f"{prod_media_squadra:.2f} colli/h", f"€ {tot_costo_reale_2w/tot_colli_2w:.4f} / collo")

    # ==========================================================================
    # PULSANTI DI DOWNLOAD
    # ==========================================================================
    st.markdown("---")
    dl_col1, dl_col2 = st.columns(2)

    with open(tmp_output_path, "rb") as f_excel:
        dl_col1.download_button(
            label=f"📥 Scarica Report Excel Completo 5 Fogli ({sistema.range_2w_str})",
            data=f_excel.read(),
            file_name=f"Report_Turni_2_Settimane_{sistema.inizio_w1.strftime('%Y%m%d')}_{sistema.fine_w2.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    buffer_ml = BytesIO()
    joblib.dump({
        'clf_tagli': sistema.clf_tagli,
        'reg_colli': sistema.reg_colli,
        'rese_orarie': sistema.rese_orarie,
        'costo_orario': sistema.costo_orario,
        'range_storico': sistema.range_storico_str,
        'range_previsione_2w': sistema.range_2w_str
    }, buffer_ml)
    buffer_ml.seek(0)

    dl_col2.download_button(
        label="🧠 Esporta Modello Machine Learning Addestrato (.joblib)",
        data=buffer_ml,
        file_name="modello_gradient_boosting_grocery.joblib",
        mime="application/octet-stream",
        use_container_width=True
    )

    # ==========================================================================
    # TABS DI CONSULTAZIONE
    # ==========================================================================
    tab1, tab2, tab3, tab4 = st.tabs([
        "📅 1. Turni & Costi 14 Giorni",
        "📊 2. Produttività & Costi per Sottoreparto",
        "📦 3. Previsione Colli (Colonna L)",
        "🔍 4. Esploso Giornaliero Sottoreparti"
    ])

    with tab1:
        st.subheader(f"Tabellone Rilascio Turni Bisettimanale ({sistema.range_2w_str})")
        cols_show_14 = [
            'Settimana', 'Data', 'Giorno', 'Totale_Colli_Giorno',
            'Somma_Addetti_Reparto_050', 'Ore_Nette_Scoperte', 'Addetti_Jolly_Mobilita',
            'Turni_Teorici_Totali_FTE', 'Squadra_Ottimale_Reale', 'Ore_Reali_Squadra',
            'Produttivita_Media_Squadra_Colli_Ora', 'Costo_Lavoro_Teorico_Euro', 'Costo_Lavoro_Squadra_Reale_Euro'
        ]
        st.dataframe(df_riepilogo_14gg[cols_show_14], use_container_width=True, hide_index=True)

    with tab2:
        st.subheader(f"Analisi Produttività e Costo del Lavoro per Sottoreparto ({sistema.range_w1_str} + {sistema.range_w2_str})")
        st.dataframe(df_prod_2w, use_container_width=True)

    with tab3:
        st.subheader("Previsione Colli Effettivi da Allestire (Target Colonna L)")
        df_colli_view = pd.concat([
            df_riepilogo_14gg[['Settimana', 'Data', 'Giorno']],
            sistema.piv_colli_2w
        ], axis=1)
        st.dataframe(df_colli_view, use_container_width=True, hide_index=True)
        st.bar_chart(sistema.piv_colli_2w[sistema.dept_list])

    with tab4:
        st.subheader("Dettaglio Giornaliero Ore, Produttività e Costi per Singolo Sottoreparto")
        filtro_sett = st.multiselect("Filtra Settimana", options=['Settimana 1', 'Settimana 2'], default=['Settimana 1', 'Settimana 2'])
        filtro_dept = st.multiselect("Filtra Sottoreparto", options=sistema.dept_list, default=sistema.dept_list)
        df_filtrato = df_dettaglio_2w[
            (df_dettaglio_2w['Giorno_Num'] < 6) &
            (df_dettaglio_2w['Settimana'].isin(filtro_sett)) &
            (df_dettaglio_2w['Sottoreparto'].isin(filtro_dept))
        ]
        st.dataframe(df_filtrato.drop(columns=['Data_Dt', 'Slot_Bisettimanale']), use_container_width=True, hide_index=True)

else:
    st.info("👆 Carica un file Excel delle distinte ordini per generare automaticamente il calendario delle 2 settimane successive, i turni e il report dei costi.")