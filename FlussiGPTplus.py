import pandas as pd
import streamlit as st
from datetime import date
import numpy as np
import openpyxl
import altair as alt
from tabulate import tabulate
import calendar

# leggi il file Excel
df = pd.read_excel('registro_turisti copy.xlsx', sheet_name=None)

# combina i dati da tutti i fogli in un unico DataFrame
combined_df = pd.concat(df.values(), ignore_index=True)

# sostituire i valori NaN con zeri
combined_df.fillna(0, inplace=True)

# convertire la colonna data in formato datetime
combined_df['DATA'] = pd.to_datetime(combined_df['DATA'], format='%d/%m/%Y', errors='coerce')

# converte solo i valori numerici in interi, ignorando quelli non numerici
numeric_cols = combined_df.columns[combined_df.dtypes == np.number]
combined_df[numeric_cols] = combined_df[numeric_cols].apply(pd.to_numeric, errors='coerce', downcast='integer')

import streamlit as st

# Creazione del menu a tendina nella barra laterale
menu = ['Dashboard', 'Analisi dati', 'Report mensile']
choice = st.sidebar.selectbox("Seleziona un'opzione", menu)

# Creazione delle diverse schermate in base alla selezione dell'utente
if choice == 'Dashboard':
    
    st.write("""## Flussi Turistici""")

    # crea un widget per selezionare la data
    selected_date = st.date_input('Seleziona data', date.today())

    # filtra il DataFrame per mostrare solo i dati della data selezionata
    filtered_df = combined_df[combined_df['DATA'] == selected_date.strftime('%Y-%m-%d')]

    # crea un widget per selezionare la nazionalità da aggiornare
    selected_nationality = st.selectbox('Seleziona la nazionalità da aggiornare', filtered_df.columns[1:])

    # ottiene il valore del conteggio attuale
    current_count = filtered_df[selected_nationality].iloc[0]

    # crea un widget per aggiornare il conteggio
    container = st.container()
    col1, col2 = container.columns([1, 2])

    # Aggiunge una barra per inserire manualmente il nuovo valore del conteggio
    new_count = col1.number_input(f'Aggiorna il conteggio per {selected_nationality}', value=current_count, min_value=0)

    if new_count != current_count:
        # Aggiorna il conteggio nel DataFrame combined_df
        combined_df.loc[(combined_df['DATA'] == pd.to_datetime(selected_date)), selected_nationality] = new_count

        # Aggiorna il valore di current_count
        current_count = new_count

        # Aggiorna il DataFrame filtered_df
        filtered_df = combined_df[combined_df['DATA'] == selected_date.strftime('%Y-%m-%d')]

    # verifica se ci sono dati per la data selezionata
    if filtered_df.drop('DATA', axis=1).eq(0).all().all():
        st.write('Nessun dato inserito per questa data')
    else:
        # Inizializziamo un dizionario con i link alle icone di bandiera
        flag_icons = {
            'Italia': 'https://upload.wikimedia.org/wikipedia/commons/0/03/Flag_of_Italy.svg',
            'Francia': 'https://upload.wikimedia.org/wikipedia/commons/c/c3/Flag_of_France.svg',
            'Germania': 'https://upload.wikimedia.org/wikipedia/commons/b/ba/Flag_of_Germany.svg',
            'Regno Unito': 'https://upload.wikimedia.org/wikipedia/en/a/ae/Flag_of_the_United_Kingdom.svg',
            'Svizzera': 'https://upload.wikimedia.org/wikipedia/commons/0/08/Flag_of_Switzerland_%28Pantone%29.svg',
            'Olanda': 'https://upload.wikimedia.org/wikipedia/commons/2/20/Flag_of_the_Netherlands.svg',
            'Spagna': 'https://upload.wikimedia.org/wikipedia/commons/9/9a/Flag_of_Spain.svg',
            'USA': 'https://upload.wikimedia.org/wikipedia/commons/a/a4/Flag_of_the_United_States.svg',
            'Altre': 'https://upload.wikimedia.org/wikipedia/commons/e/ef/International_Flag_of_Planet_Earth.svg',
        }
        # mostra i dati del DataFrame
        st.write('\tConteggio \tnazionalità')
        for column in filtered_df.columns[1:]:  # cicla attraverso tutte le colonne tranne la colonna "DATA"
            nationality = column
            flag_icon = flag_icons.get(nationality, '')
            count = filtered_df[column].iloc[0]  # ottieni il valore del conteggio dalla colonna corrispondente e dalla prima riga del DataFrame
            if count != 0:  # verifica se il conteggio è diverso da 0
                col1, col2 = st.columns([0.05, 1])  # Dividiamo la riga in due colonne e rendiamo la prima colonna più stretta
                with col1:
                    st.image(flag_icon, width=32, use_column_width=False)  # Aggiungiamo l'immagine della bandiera nella prima colonna e adattiamo la larghezza alla colonna
                with col2:
                    st.write(f'{nationality}\t\t{count}')  # Aggiungiamo il testo della nazionalità nella seconda colonna

    # crea un pulsante per salvare il DataFrame aggiornato in un nuovo file Excel
    if st.button('Salva'):
        # verifica se la colonna 'DATA' esiste e se è di tipo datetime
        if 'DATA' in combined_df.columns and pd.api.types.is_datetime64_any_dtype(combined_df['DATA']):
            # salva il DataFrame in un nuovo file Excel
            combined_df.to_excel('registro_turisti copy.xlsx', sheet_name='Dati', index=False)

            # conferma il salvataggio
            st.write('Il file Excel è stato aggiornato con successo')
        else:
            st.write('La colonna "DATA" non esiste o non è di tipo datetime')

if choice == 'Analisi dati':
    
    # crea un widget per selezionare il periodo di interesse
    st.write("""## Seleziona periodo di interesse""")
    start_date = pd.to_datetime(st.date_input("Data di inizio", value=pd.to_datetime(combined_df["DATA"].min())))
    end_date = pd.to_datetime(st.date_input("Data di fine", value=pd.to_datetime(combined_df["DATA"].max())))


    # filtra il DataFrame per mostrare solo i dati del periodo di interesse
    filtered_df = combined_df[(combined_df['DATA'] >= start_date) & (combined_df['DATA'] <= end_date)]

    # crea un widget per selezionare le nazionalità da visualizzare
    st.write("""## Seleziona nazionalità""")
    nationalities = st.multiselect("Seleziona le nazionalità da visualizzare", filtered_df.columns[1:])

    # crea il grafico a linee
    st.write("""## Andamento flussi turistici""")
    if len(nationalities) > 0:
        # filtra il DataFrame per mostrare solo le nazionalità selezionate
        filtered_df = filtered_df[nationalities + ['DATA']]
        
        # munge il DataFrame per prepararlo per il grafico
        melted_df = filtered_df.melt(id_vars=['DATA'], var_name='Nazionalità', value_name='Conteggio')
        
        # crea il grafico a linee interattivo
        chart = alt.Chart(melted_df).mark_line().encode(
            x='DATA',
            y='Conteggio',
            color='Nazionalità'
        ).interactive()
        
        # visualizza il grafico
        st.altair_chart(chart, use_container_width=True)
    else:
        st.write("Seleziona almeno una nazionalità")
    # crea il grafico a torta
    st.write("""## Distribuzione nazionalità""")

    # filtra il DataFrame per mostrare solo le righe del periodo di interesse
    filtered_df = combined_df[(combined_df['DATA'] >= start_date) & (combined_df['DATA'] <= end_date)]

    # calcola la somma dei conteggi per ogni nazionalità
    nationality_counts = filtered_df.drop('DATA', axis=1).sum()

    # munge il DataFrame per prepararlo per il grafico a torta
    nationality_counts_df = nationality_counts.reset_index()
    nationality_counts_df.columns = ['Nazionalità', 'Conteggio']

    # crea il grafico a torta interattivo
    pie_chart = alt.Chart(nationality_counts_df).mark_arc(innerRadius=50, outerRadius=100, cornerRadius=5, padAngle=0.01).encode(
        theta='Conteggio:Q',
        color='Nazionalità:N',
        tooltip=['Nazionalità', 'Conteggio']
    ).interactive()

    # visualizza il grafico a torta
    st.altair_chart(pie_chart, use_container_width=True)

if choice == 'Report mensile':
    st.write("""## Report mensile""")
    # chiedi all'utente di selezionare un mese e un anno
    selected_month = st.selectbox('Seleziona il mese', list(enumerate(calendar.month_name[1:], 1)), format_func=lambda x: x[1], key='selected_month')[0]
    selected_year = st.selectbox('Seleziona l\'anno', combined_df['DATA'].dt.year.unique(), key='selected_year')
   
    # aggiungi un pulsante "Genera tabella"
    if st.button("Genera tabella"):
        # crea un filtro per il mese e l'anno selezionati dall'utente
        month_filter = (combined_df['DATA'].dt.month == selected_month) & (combined_df['DATA'].dt.year == selected_year)

        # crea un filtro per il mese precedente
        previous_month = (selected_month - 1) if (selected_month > 1) else 12
        previous_month_filter = (combined_df['DATA'].dt.month == previous_month) & (combined_df['DATA'].dt.year == selected_year)

        # crea un filtro per lo stesso mese dell'anno precedente
        previous_year_filter = (combined_df['DATA'].dt.month == selected_month) & (combined_df['DATA'].dt.year == (selected_year - 1))

        # crea un filtro per tutti gli anni precedenti al mese selezionato
        all_previous_years_filter = (combined_df['DATA'].dt.month == selected_month) & (combined_df['DATA'].dt.year < selected_year)

        # crea i sotto-dataframe con i dati filtrati
        selected_month_df = combined_df[month_filter]
        previous_month_df = combined_df[previous_month_filter]
        previous_year_df = combined_df[previous_year_filter]
        all_previous_years_df = combined_df[all_previous_years_filter]

        # Calcola la somma di tutte le colonne per ogni DataFrame
        selected_month_totals = selected_month_df.sum(axis=0)
        previous_month_totals = previous_month_df.sum(axis=0)
        previous_year_totals = previous_year_df.sum(axis=0)
        all_previous_years_totals = all_previous_years_df.sum(axis=0)

        # calcola la media degli anni precedenti a quello selezionato
        num_previous_years = selected_year - combined_df['DATA'].dt.year.min()
        all_previous_years_totals = all_previous_years_totals / num_previous_years

        # Calcola la somma dei turisti stranieri per ogni DataFrame
        selected_month_foreign = selected_month_df.iloc[:, 1:].sum(axis=0)
        previous_month_foreign = previous_month_df.iloc[:, 1:].sum(axis=0)
        previous_year_foreign = previous_year_df.iloc[:, 1:].sum(axis=0)
        all_previous_years_foreign = all_previous_years_df.iloc[:, 1:].sum(axis=0)

        # Calcola la somma di tutti i turisti per ogni DataFrame
        selected_month_total = selected_month_totals.sum()
        previous_month_total = previous_month_totals.sum()
        previous_year_total = previous_year_totals.sum()
        all_previous_years_total = all_previous_years_totals.sum()

        def calculate_percent_change_vectorized(current, previous):
            non_zero_mask = (previous != 0)
            percent_change = np.zeros_like(current)
            percent_change[non_zero_mask] = (current[non_zero_mask] - previous[non_zero_mask]) / previous[non_zero_mask] * 100
            return percent_change

        # Calcola la variazione percentuale rispetto al mese precedente
        percent_change_prev_month = calculate_percent_change_vectorized(selected_month_totals, previous_month_totals)

        # Calcola la variazione percentuale rispetto all'anno precedente
        percent_change_prev_year = calculate_percent_change_vectorized(selected_month_totals, previous_year_totals)

        # Calcola la variazione percentuale rispetto a tutti gli anni precedenti
        percent_change_all_prev_years = calculate_percent_change_vectorized(selected_month_totals, all_previous_years_totals)

        # Prepara i dati per la tabella
        percent_changes_data = {
            'Nazionalità': selected_month_totals.index,
            'Mese precedente': percent_change_prev_month,
            'Anno precedente': percent_change_prev_year,
            'Media anni precedenti': percent_change_all_prev_years,
        }

        # Crea un DataFrame con i dati delle variazioni percentuali
        percent_changes_df = pd.DataFrame(percent_changes_data)

        # Trasponi il DataFrame
        percent_changes_df = percent_changes_df.set_index('Nazionalità').transpose()

        # Resetta l'indice per avere una colonna con i titoli "Variazione rispetto a..."
        percent_changes_df.reset_index(inplace=True)
        percent_changes_df.rename(columns={'index': 'Variazione rispetto a'}, inplace=True)

        # Definisce una funzione per colorare i valori negativi in rosso e i positivi in verde
        def color_negative_red(val):
            if isinstance(val, float):
                color = 'red' if val < 0 else 'green'
                return f'color: {color}'
            return ''

        # Applica lo stile alla tabella
        styled_table = (
            percent_changes_df.style
            .applymap(color_negative_red, subset=pd.IndexSlice[:, percent_changes_df.columns[1:]])
            .format({col: '{:.2f}%' for col in percent_changes_df.columns[1:]})
            .hide_index()
        )

        # Mostra la tabella all'utente utilizzando Streamlit
        st.write("Variazioni percentuali:", styled_table)
