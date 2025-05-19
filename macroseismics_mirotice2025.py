import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os


# --- Konfigurace ---
DATA_FILE_PATH = 'makroseis2025.xlsx'
SHEET_NAME = 0
OUTPUT_DIR = 'analyza_vysledky_2025'

EQ_YEAR_TARGET = 2025
EQ_MONTH_TARGET = 4
EQ_DAY_TARGET = 24
EQ_DATETIME_UTC_STR = f'{EQ_YEAR_TARGET}-{EQ_MONTH_TARGET:02d}-{EQ_DAY_TARGET:02d} 17:32:47.3'

EQ_LAT = 49.422
EQ_LON = 14.043
EQ_MAGNITUDE = 3.1
EQ_GFU_ID = 1773

# --- Definice názvů sloupců z Excelu ---
COL_OBS_DATETIME = 'eqdatetime' 
COL_LAT = 'lat'
COL_LON = 'lon'
COL_IN_BUILDING = 'pozorovaniodkud'     
COL_FEAR = 'reakcepanika'
COL_TREMOR_TYPE = 'popispohybu' 
COL_FELT_BY = 'kolikpozorvenku'
COLS_OBJECT_MOVEMENT_DETAILS = [
    'nabytektezky',           
    'okna',       
    'dvere',        
    'zavespredmety',   
    'nadobi',        
    'malepredmety',       
    'kapalina'     
]
COL_DAMAGE_OVERALL = 'poskozomitka' 

try:
    df = pd.read_excel(DATA_FILE_PATH, sheet_name=SHEET_NAME, na_values=['NULL', 'null', ''])
except FileNotFoundError:
    print(f"CHYBA: Soubor {DATA_FILE_PATH} nebyl nalezen.")
    exit()
except Exception as e:
    print(f"CHYBA při načítání souboru Excel '{DATA_FILE_PATH}': {e}")
    exit()
print(f"Načteno {len(df)} řádků z {DATA_FILE_PATH}")

try:
    df[COL_OBS_DATETIME] = pd.to_datetime(df[COL_OBS_DATETIME], dayfirst=True, errors='coerce')
    df.dropna(subset=[COL_OBS_DATETIME], inplace=True)
    print(f"Po konverzi času a odstranění neplatných časů zbývá: {len(df)} řádků")
    if df[COL_OBS_DATETIME].dt.tz is None:
        df[COL_OBS_DATETIME] = df[COL_OBS_DATETIME].dt.tz_localize('Europe/Prague', ambiguous='infer').dt.tz_convert('UTC')
    else:
        df[COL_OBS_DATETIME] = df[COL_OBS_DATETIME].dt.tz_convert('UTC')
except KeyError:
    print(f"CHYBA: Klíčový sloupec (např. '{COL_OBS_DATETIME}', '{COL_LAT}', '{COL_LON}') nebyl nalezen. Zkontrolujte definice COL_... ve skriptu.")
    exit()
except Exception as e:
    print(f"CHYBA při konverzi sloupce '{COL_OBS_DATETIME}' nebo přiřazování časové zóny: {e}")
    exit()
df[COL_LAT] = pd.to_numeric(df[COL_LAT], errors='coerce')
df[COL_LON] = pd.to_numeric(df[COL_LON], errors='coerce')
df.dropna(subset=[COL_LAT, COL_LON], inplace=True)
print(f"Po konverzi souřadnic a odstranění neplatných zbývá: {len(df)} řádků")
df_filtered_year_month = df[
    (df[COL_OBS_DATETIME].dt.year == EQ_YEAR_TARGET) &
    (df[COL_OBS_DATETIME].dt.month == EQ_MONTH_TARGET)
].copy()
print(f"Nalezeno {len(df_filtered_year_month)} záznamů pro {EQ_MONTH_TARGET}/{EQ_YEAR_TARGET}.")
if df_filtered_year_month.empty:
    print(f"Nebyly nalezeny žádné záznamy pro {EQ_MONTH_TARGET}/{EQ_YEAR_TARGET}. Analýza končí.")
    exit()
target_eq_datetime_utc = pd.Timestamp(EQ_DATETIME_UTC_STR, tz='UTC')
TIME_WINDOW_HOURS_FILTER = 1.5
time_delta_filter = pd.Timedelta(hours=TIME_WINDOW_HOURS_FILTER)
start_time_filter = target_eq_datetime_utc - time_delta_filter
end_time_filter = target_eq_datetime_utc + time_delta_filter
df_event = df_filtered_year_month.loc[
    (df_filtered_year_month[COL_OBS_DATETIME] >= start_time_filter) &
    (df_filtered_year_month[COL_OBS_DATETIME] <= end_time_filter)
].copy()
print(f"\n--- ANALÝZA PRO ZEMĚTŘESENÍ {EQ_GFU_ID} ({target_eq_datetime_utc.strftime('%Y-%m-%d %H:%M:%S %Z')}) ---")
print(f"Filtrováno pro časové okno od {start_time_filter.strftime('%Y-%m-%d %H:%M:%S %Z')} do {end_time_filter.strftime('%Y-%m-%d %H:%M:%S %Z')}")
print(f"Nalezeno {len(df_event)} pozorování v tomto okně.")
if df_event.empty:
    print("Nebyly nalezeny žádné záznamy v definovaném časovém okně pro událost. Analýza končí.")
    exit()
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
    print(f"Vytvořen adresář pro výstupy: {OUTPUT_DIR}")



# ---  Mapa pozorování ---
print("\n---  Mapa pozorování ---")
hover_data_map = {COL_OBS_DATETIME: True}
color_map = None
CENTER_LAT_CR = 49.8175  # latitude
CENTER_LON_CR = 15.4730  # longitude
ZOOM_LEVEL_CR = 5.8      

fig_map = px.scatter_mapbox(
    df_event, lat=COL_LAT, lon=COL_LON, hover_name=df_event.index,
    hover_data=hover_data_map,
    color=color_map,
    size_max=10,
    zoom=ZOOM_LEVEL_CR,             # Použití nové úrovně zoomu
    center={"lat": CENTER_LAT_CR, "lon": CENTER_LON_CR}, # Použití nového středu
    title=f"Mapa pozorování zemětřesení {EQ_GFU_ID} (časové okno +/- {TIME_WINDOW_HOURS_FILTER}h)"
)
fig_map.add_trace(go.Scattermapbox(
    lat=[EQ_LAT], lon=[EQ_LON], mode='markers', # Epicentrum zemětřesení
    marker=go.scattermapbox.Marker(size=15,color='red',symbol='star'),
    name=f'Epicentrum (Mag: {EQ_MAGNITUDE})',
    text=[f'Epicentrum Magnituda: {EQ_MAGNITUDE}'],
    hoverinfo='text'
))
fig_map.update_layout(mapbox_style="open-street-map", margin={"r":0,"t":50,"l":0,"b":0})
try:
    fig_map.show()
    fig_map.write_html(os.path.join(OUTPUT_DIR, "mapa_pozorovani_2025.html"))
    print(f"Mapa uložena do souboru: {os.path.join(OUTPUT_DIR, 'mapa_pozorovani_2025.html')}")
except Exception as e:
    print(f"CHYBA při zobrazování nebo ukládání mapy: {e}")





# --- Pozorování doma vs. venku  ---
if COL_IN_BUILDING and COL_IN_BUILDING in df_event.columns:
    print(f"\n--- Pozorování doma vs. venku (sloupec: {COL_IN_BUILDING}) ---")
    # Hodnota 'budova' znamená uvnitř
    # Cokoli jiného (včetně prázdné buňky/NaN je převedeno na prázdný string) a znamená venku
    df_event.loc[:, 'Mist_Pozorovani_Kat'] = df_event[COL_IN_BUILDING].fillna('').astype(str).str.lower().apply(
        lambda x: 'Doma (uvnitř)' if x == 'budova' else 'Venku/Nezadáno v budově'
    )
    indoor_outdoor_counts = df_event['Mist_Pozorovani_Kat'].value_counts()
    print("  Počty:")
    print(indoor_outdoor_counts)
    print(f"\n  Procenta (ze všech {len(df_event)} pozorování):")
    for category, count in indoor_outdoor_counts.items():
        percentage_of_total = (count / len(df_event)) * 100
        print(f"    - {category}: {count} ({percentage_of_total:.2f} %)")
else:
    print(f"\nINFO: Sloupec '{COL_IN_BUILDING}' pro Úkol 3 nebyl nalezen nebo není definován.")




# --- Typ pocítění (jeden vs. většina) ---
if COL_FELT_BY and COL_FELT_BY in df_event.columns:
    print(f"\n--- Typ pocítění (sloupec: {COL_FELT_BY}) ---")
    df_event.loc[:, 'Pocit_Kategorie_Text'] = df_event[COL_FELT_BY].astype(str).str.lower().str.strip().map({
        'pouze vy': 'Pocítil(a) pouze respondent / někdo', 
        'většina ano': 'Pocítila většina přítomných'       
    }).fillna('Nezadáno/Jiná odpověď') 
    felt_by_counts = df_event['Pocit_Kategorie_Text'].value_counts()

    print("\n  Počty odpovědí:")
    print(felt_by_counts)

    print(f"\n  Procentuální zastoupení (ze všech {len(df_event)} pozorování v časovém okně):")
    for category, count in felt_by_counts.items():
        percentage_of_total = (count / len(df_event)) * 100
        print(f"    - {category}: {count} ({percentage_of_total:.2f} %)")

else:
    print(f"\nINFO: Sloupec '{COL_FELT_BY}' pro Úkol 4 (Typ pocítění) nebyl nalezen v datech nebo není definován.")
    print("      Analýza typu pocítění bude přeskočena.")





# --- Otřesy (intenzita) ---
if COL_TREMOR_TYPE and COL_TREMOR_TYPE in df_event.columns:
    print(f"\n--- Intenzita otřesů (sloupec: {COL_TREMOR_TYPE}) ---")
    df_event.loc[:, 'Intenzita_Kat'] = df_event[COL_TREMOR_TYPE].astype(str).str.lower().fillna('Nezadáno')
    tremor_counts = df_event['Intenzita_Kat'].value_counts()
    print("  Počty:")
    print(tremor_counts)
    print(f"\n  Procenta (ze všech {len(df_event)} pozorování):")
    for category, count in tremor_counts.items():
        percentage_of_total = (count / len(df_event)) * 100
        print(f"    - {category}: {count} ({percentage_of_total:.2f} %)")
else:
    print(f"\nINFO: Sloupec '{COL_TREMOR_TYPE}' pro Úkol 5 (intenzita) nebyl nalezen nebo není definován.")




# --- Strach (panika) ---
if COL_FEAR and COL_FEAR in df_event.columns:
    print(f"\n--- Strach/Panika (sloupec: {COL_FEAR}) ---")
    # Hodnota 1 ve sloupci COL_FEAR ('reakcepanika') znamená "Ano, strach/panika"
    df_event.loc[:, 'Strach_Pocit_Kat'] = pd.to_numeric(df_event[COL_FEAR], errors='coerce') == 1
    
    fear_counts = df_event['Strach_Pocit_Kat'].map({True: 'Ano (strach/panika)', False: 'Ne/Nezadáno'}).value_counts()
    print("  Počty:")
    print(fear_counts)
    print(f"\n  Procenta (ze všech {len(df_event)} pozorování):")
    if True in fear_counts.index: 
         print(f"    - Ano (strach/panika): {fear_counts.get(True, 0)} ({(fear_counts.get(True, 0) / len(df_event) * 100):.2f} %)")
    if False in fear_counts.index: 
         print(f"    - Ne/Nezadáno: {fear_counts.get(False, 0)} ({(fear_counts.get(False, 0) / len(df_event) * 100):.2f} %)")
else:
    print(f"\nINFO: Sloupec '{COL_FEAR}' pro Úkol 6 (strach/panika) nebyl nalezen nebo není definován.")




# --- Pohyb předmětů ---
print("\n--- Pohyb předmětů ---")
actual_movement_detail_cols = [col for col in COLS_OBJECT_MOVEMENT_DETAILS if col in df_event.columns]
if actual_movement_detail_cols:
    df_event.loc[:, 'Pohyb_Predmetu_Agregovany'] = df_event[actual_movement_detail_cols].apply(
        lambda row: any(pd.notna(val) and str(val).strip() != '' for val in row), axis=1
    )
    movement_agg_counts = df_event['Pohyb_Predmetu_Agregovany'].value_counts()
    print("Agregovaně (zda byl pozorován jakýkoli pohyb předmětů):")
    if True in movement_agg_counts.index:
        print(f"  - Ano (nějaký pohyb): {movement_agg_counts.get(True, 0)} pozorování ({(movement_agg_counts.get(True, 0) / len(df_event) * 100):.2f} %)")
    if False in movement_agg_counts.index:
         print(f"  - Ne (žádný pohyb): {movement_agg_counts.get(False, 0)} pozorování ({(movement_agg_counts.get(False, 0) / len(df_event) * 100):.2f} %)")

    print(f"\nProcentuální zastoupení jednotlivých typů pozorovaných pohybů (ze všech {len(df_event)} pozorování):")
    for col_name in actual_movement_detail_cols:
        observed_count = (df_event[col_name].notna() & (df_event[col_name].astype(str).str.strip() != '')).sum()
        if observed_count > 0:
            percentage = (observed_count / len(df_event)) * 100
            print(f"  - Typ pohybu (sloupec '{col_name}'): {observed_count} krát ({percentage:.2f} %)")
else:
    print(f"INFO: Nebyly nalezeny žádné definované sloupce pro detailní pohyb předmětů ({COLS_OBJECT_MOVEMENT_DETAILS}).")


# --- Poškození budov ---
print("\n--- Poškození budovy (bylo/nebylo) ---")
if COL_DAMAGE_OVERALL and COL_DAMAGE_OVERALL in df_event.columns:
    print(f"Analýza na základě sloupce: '{COL_DAMAGE_OVERALL}'")
    df_event.loc[:, 'Poskozeni_Obecne_Kat'] = df_event[COL_DAMAGE_OVERALL].astype(str).str.lower().str.strip().map({
        'bylo': 'Ano (poškození hlášeno)',
        'nebylo': 'Ne (poškození nehlášeno)'
    }).fillna('Nezadáno/Jiná hodnota') 
    damage_overall_counts = df_event['Poskozeni_Obecne_Kat'].value_counts()

    print("\n  Počty odpovědí:")
    print(damage_overall_counts)

    print(f"\n  Procentuální zastoupení (ze všech {len(df_event)} pozorování v časovém okně):")
    for category, count in damage_overall_counts.items():
        percentage_of_total = (count / len(df_event)) * 100
        print(f"    - {category}: {count} ({percentage_of_total:.2f} %)")
else:
    print(f"\nINFO: Sloupec '{COL_DAMAGE_OVERALL}' pro analýzu poškození (bylo/nebylo) nebyl nalezen v datech nebo není definován.")
    print("      Analýza poškození budovy (bylo/nebylo) bude přeskočena.")


print("\n--- Konec analýzy ---")