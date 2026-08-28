import os
import math
import gc
import re
import json
from datetime import datetime, timedelta
from flask import Flask, request, jsonify, send_from_directory, make_response
from flask_cors import CORS
import pandas as pd
import numpy as np

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(BASE_DIR, 'forecast_cache.json')
CONFIG_FILE = os.path.join(BASE_DIR, 'wfm_config.json') 
EXCEL_DEFAULT = os.path.join(BASE_DIR, 'historico.xlsx')

app = Flask(__name__)
CORS(app)

VENTANAS_SERVICIO = {
    'ambulancia servicios': {'inicio': 0 * 60, 'fin': 24 * 60},
    'asignación hogar': {'inicio': 0 * 60, 'fin': 24 * 60},
    'asignacion hogar': {'inicio': 0 * 60, 'fin': 24 * 60},
    'asignación vial': {'inicio': 0 * 60, 'fin': 24 * 60},
    'asignacion vial': {'inicio': 0 * 60, 'fin': 24 * 60},
    'coppel servicios': {'inicio': 0 * 60, 'fin': 24 * 60},
    'liverpool servicios': {'inicio': 0 * 60, 'fin': 24 * 60},
    'multicampañas': {'inicio': 0 * 60, 'fin': 24 * 60},
    'multicampanas': {'inicio': 0 * 60, 'fin': 24 * 60},
    'seguimiento hogar': {'inicio': 0 * 60, 'fin': 24 * 60},
    'seguimiento vial': {'inicio': 0 * 60, 'fin': 24 * 60},
    'suburbia servicios': {'inicio': 0 * 60, 'fin': 24 * 60},
    'experiencias liverpool': {'inicio': 9 * 60, 'fin': 21 * 60},
    'experiencias suburbia': {'inicio': 9 * 60, 'fin': 21 * 60},
    'retenciones suburbia': {'inicio': 9 * 60, 'fin': 20 * 60},
    'retenciones liverpool': {'inicio': 9 * 60, 'fin': 20 * 60}
}

def buscar_archivo_excel():
    if os.path.exists(EXCEL_DEFAULT): return EXCEL_DEFAULT
    try:
        archivos = [f for f in os.listdir(BASE_DIR) if f.lower().endswith('.xlsx') and not f.startswith('~')]
        if not archivos: return None
        for f in archivos:
            if 'historico' in f.lower(): return os.path.join(BASE_DIR, f)
        return os.path.join(BASE_DIR, archivos[0])
    except: return None

@app.route('/')
@app.route('/index.html')
def serve_index():
    rutas_a_buscar = [BASE_DIR, os.getcwd(), os.path.dirname(BASE_DIR)]
    for ruta in rutas_a_buscar:
        target_path = os.path.join(ruta, 'index.html')
        if os.path.exists(target_path):
            response = make_response(send_from_directory(ruta, 'index.html'))
            response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
            response.headers['Pragma'] = 'no-cache'
            response.headers['Expires'] = '0'
            return response
    return jsonify({"error": "ALERTA CRITICA: No se encontro el archivo index.html."}), 404

@app.route('/favicon.ico')
def favicon(): return '', 204

@app.route('/api/config', methods=['GET', 'POST'])
def manage_config():
    if request.method == 'POST':
        try:
            new_config = request.get_json(force=True)
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(new_config, f)
            return jsonify({'status': 'Guardado'}), 200
        except Exception as e: return jsonify({'error': str(e)}), 500
    else:
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f: return jsonify(json.load(f)), 200
            except: pass
        return jsonify({'targetSl': 80, 'targetTime': 20, 'merma': 30, 'duracionJornada': 8, 'chkNocturno': False, 'chkPicos': False}), 200

def clean_num(val, default=0.0):
    if pd.isna(val) or val is None: return default
    try:
        val_str = str(val).strip().replace(',', '.')
        val_str = re.sub(r'[^0-9.]', '', val_str)
        return float(val_str) if val_str else default
    except: return default

def parse_aht_to_seconds(val):
    if pd.isna(val) or val is None: return 180.0
    if isinstance(val, (int, float)): return float(val) if float(val) > 15 else float(val) * 60.0
    val_str = str(val).strip()
    if ':' in val_str:
        p = val_str.split(':')
        try:
            if len(p) == 3: return int(p[0]) * 3600 + int(p[1]) * 60 + float(p[2])
            elif len(p) == 2: return int(p[0]) * 60 + float(p[1])
        except: pass
    return clean_num(val_str, 180.0)

def format_aht_str(seconds):
    if pd.isna(seconds) or seconds is None or seconds <= 0: return "00:00:00"
    secs = int(round(seconds))
    return f"{secs // 3600:02d}:{(secs % 3600) // 60:02d}:{secs % 60:02d}"

def clean_interval_str(val):
    try:
        if pd.isna(val): return "00:00"
        val_str = str(val).strip()
        if hasattr(val, 'hour') and hasattr(val, 'minute'): hh, mm = val.hour, val.minute
        else:
            m = re.search(r'(\d{1,2}):(\d{2})', val_str)
            if m: hh, mm = int(m.group(1)), int(m.group(2))
            else: return "00:00"
        if mm < 15: mm_round = 0
        elif mm < 45: mm_round = 30
        else:
            mm_round = 0
            hh = (hh + 1) % 24
        return f"{hh:02d}:{mm_round:02d}"
    except: return "00:00"

ERLANG_CACHE = {}
def erlang_c_sl_optimizado(A, N, AHT, target_time):
    if N <= A or A <= 0 or N <= 0: return 0.0
    key = (round(A, 2), N, round(AHT, 1), target_time)
    if key in ERLANG_CACHE: return ERLANG_CACHE[key]
    try:
        sum_terms, current_term = 1.0, 1.0
        int_N = min(int(N), 1000)
        for k in range(1, int_N):
            current_term *= (A / k)
            sum_terms += current_term
        last_term = current_term * (A / N) / (1.0 - (A / N))
        pw = last_term / (sum_terms + last_term)
        sl = 1.0 - (pw * math.exp(-(N - A) * (target_time / AHT)))
        resultado = round(max(0.0, min(100.0, sl * 100.0)), 1)
        ERLANG_CACHE[key] = resultado
        return resultado
    except: return 0.0

def calcular_agentes_requeridos_erlang_c(A, aht, target_time, target_sl):
    if A <= 0 or aht <= 0: return 0
    n = max(1, int(math.floor(A)) + 1)
    if A > 50: n = max(n, int(math.floor(A + math.sqrt(A))))
    while n < 3000:
        if erlang_c_sl_optimizado(A, n, aht, target_time) >= target_sl: return n
        n += 1
    return n

def parse_time_str(t_str):
    if not t_str: return None
    t = re.sub(r'[^\d:]', '', str(t_str).lower())
    if not t: return None
    if ':' not in t: t += ':00'
    try:
        p = t.split(':')
        return int(p[0]) * 60 + int(p[1])
    except: return None

def esta_en_ventana_servicio(campana, intervalo_str):
    camp_key = str(campana).strip().lower()
    min_in = parse_time_str(intervalo_str)
    if min_in is None: return True
    for key, ventana in VENTANAS_SERVICIO.items():
        if key in camp_key or camp_key in key:
            return ventana['inicio'] <= min_in < ventana['fin']
    return True

def encontrar_columna(df, posibles):
    for p in posibles:
        for c in df.columns:
            if p.strip().lower() in str(c).strip().lower(): return c
    return None

def generar_intervalos_cobertura(start_min, end_min):
    intervals = []
    curr = start_min
    if start_min < end_min:
        while curr < end_min:
            intervals.append(f"{int(curr // 60):02d}:{int(curr % 60):02d}")
            curr += 30
    else: 
        while curr < 24 * 60:
            intervals.append(f"{int(curr // 60):02d}:{int(curr % 60):02d}")
            curr += 30
        curr = 0
        while curr < end_min:
            intervals.append(f"{int(curr // 60):02d}:{int(curr % 60):02d}")
            curr += 30
    return intervals

def procesar_hoja_roster(df_roster):
    dias_map = {'lunes': 'Lunes', 'martes': 'Martes', 'miércoles': 'Miércoles', 'miercoles': 'Miércoles', 
                'jueves': 'Jueves', 'viernes': 'Viernes', 'sábado': 'Sábado', 'sabado': 'Sábado', 'domingo': 'Domingo'}
    roster_cov, roster_total_camp, roster_total_dia_camp = {}, {}, {}
    col_camp = encontrar_columna(df_roster, ['campaña', 'campana', 'skill', 'servicio'])
    if not col_camp: return roster_cov, roster_total_camp, roster_total_dia_camp
        
    for idx, row in df_roster.iterrows():
        camp = str(row[col_camp]).strip().title()
        if camp == 'Nan' or camp == '': continue
        roster_total_camp[camp] = roster_total_camp.get(camp, 0) + 1
        for col in df_roster.columns:
            c_lower = str(col).lower().strip()
            if c_lower in dias_map:
                dia_real = dias_map[c_lower]
                horario = str(row[col]).strip().upper()
                if horario != 'DD-DD' and 'NAN' not in horario and '-' in horario:
                    key_dia = (camp, dia_real)
                    roster_total_dia_camp[key_dia] = roster_total_dia_camp.get(key_dia, 0) + 1
                    parts = horario.split('-')
                    if len(parts) == 2:
                        s_min = parse_time_str(parts[0].strip())
                        e_min = parse_time_str(parts[1].strip())
                        if s_min is not None and e_min is not None:
                            for inv in generar_intervalos_cobertura(s_min, e_min):
                                roster_cov[(camp, dia_real, inv)] = roster_cov.get((camp, dia_real, inv), 0) + 1
    return roster_cov, roster_total_camp, roster_total_dia_camp

def procesar_archivo_excel(file_source, target_sl=80.0, target_time=20.0, merma=0.20, dias_futuros=45):
    xls_file = pd.ExcelFile(file_source, engine='openpyxl')
    sheet_calls = xls_file.sheet_names[0]
    for s in xls_file.sheet_names:
        if 'llam' in s.lower() or 'hist' in s.lower() or 'datos' in s.lower(): sheet_calls = s; break
            
    sheet_roster = None
    for s in xls_file.sheet_names:
        if 'roster' in s.lower() or 'plantilla' in s.lower() or 'horario' in s.lower(): sheet_roster = s; break

    roster_coverage, roster_total_camp, roster_total_dia_camp = {}, {}, {}
    if sheet_roster:
        try:
            df_roster = pd.read_excel(xls_file, sheet_name=sheet_roster, engine='openpyxl')
            roster_coverage, roster_total_camp, roster_total_dia_camp = procesar_hoja_roster(df_roster)
        except: pass

    df_raw = pd.read_excel(xls_file, sheet_name=sheet_calls, engine='openpyxl')
    col_calls = encontrar_columna(df_raw, ['recibidas', 'llamadas', 'calls', 'volumen', 'ofrecidas', 'entrada'])
    col_aht = encontrar_columna(df_raw, ['aht', 'tmo', 'handle', 'duracion'])
    col_camp = encontrar_columna(df_raw, ['campaña', 'campana', 'skill', 'servicio', 'ring group'])
    col_inter = encontrar_columna(df_raw, ['intervalo', 'hora', 'time'])
    col_fecha = encontrar_columna(df_raw, ['fecha', 'date'])

    if not col_camp: col_camp = df_raw.columns[0]
    if not col_fecha: col_fecha = df_raw.columns[1]
    if not col_inter: col_inter = df_raw.columns[2]
    if not col_calls: col_calls = df_raw.columns[3]

    df_raw[col_camp] = df_raw[col_camp].astype(str).str.strip().str.title()
    df_raw[col_fecha] = pd.to_datetime(df_raw[col_fecha], dayfirst=True, errors='coerce').dt.normalize()
    df_raw = df_raw.dropna(subset=[col_fecha])
    df_raw[col_calls] = [clean_num(x, 0.0) for x in df_raw[col_calls]]

    df_valido = df_raw[df_raw[col_calls] > 0]
    if df_valido.empty: raise ValueError("El archivo no tiene volumen mayor a cero.")
    
    max_fecha_real = df_valido[col_fecha].max()
    df_raw = df_raw[df_raw[col_fecha] <= max_fecha_real]

    if col_aht: df_raw[col_aht] = [parse_aht_to_seconds(x) for x in df_raw[col_aht]]
    else: df_raw['AHT_Calc'] = 180.0; col_aht = 'AHT_Calc'

    df_raw['Inter_Clean'] = df_raw[col_inter].apply(clean_interval_str)
    df_raw['Total_Segundos_Handle'] = df_raw[col_calls] * df_raw[col_aht]

    df = df_raw.groupby([col_fecha, col_camp, 'Inter_Clean']).agg({col_calls: 'sum', 'Total_Segundos_Handle': 'sum'}).reset_index()
    df[col_aht] = np.where(df[col_calls] > 0, df['Total_Segundos_Handle'] / df[col_calls], 180.0)
    df = df.drop(columns=['Total_Segundos_Handle'])

    dias_espanol = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']
    meses_espanol = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    df['Dia_Semana_Clean'] = df[col_fecha].dt.weekday.apply(lambda w: dias_espanol[w])

    fecha_inicio_forecast = max_fecha_real + timedelta(days=1)
    aht_global_campana = df.groupby(col_camp)[col_aht].apply(lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 180.0).to_dict()

    df_diario = df.groupby([col_fecha, col_camp])[col_calls].sum().reset_index()
    campanas_unicas = df[col_camp].unique()

    predicciones_futuras, factores_ui = {}, {}

    for camp in campanas_unicas:
        sub = df_diario[df_diario[col_camp] == camp].sort_values(col_fecha).reset_index(drop=True)
        if sub.empty: continue
        
        dow_avg = {}
        for i in range(7):
            vols_dow = sub[sub[col_fecha].dt.weekday == i][col_calls]
            vols_dow = vols_dow[vols_dow > 0]
            if len(vols_dow) >= 3: dow_avg[i] = vols_dow.tail(4).mean()
            elif len(vols_dow) > 0: dow_avg[i] = vols_dow.mean()
            else: dow_avg[i] = sub[col_calls].mean()

        n = len(sub)
        
        # Factor de tendencia
        avg_julio = sub[(sub[col_fecha].dt.month == 7)][col_calls].mean()
        avg_agosto = sub[(sub[col_fecha].dt.month == 8)][col_calls].mean()
        
        if avg_julio > 0 and avg_agosto > 0:
            trend_rate = avg_agosto / avg_julio
        else:
            trend_rate = 1.0
            
        trend_rate = max(0.92, min(1.08, trend_rate))

        preds_finales = []
        for d in range(dias_futuros):
            fecha_futura = fecha_inicio_forecast + timedelta(days=d)
            wd = fecha_futura.weekday()
            day_m = fecha_futura.day
            
            vol_base = dow_avg.get(wd, sub[col_calls].mean())
            
            month_idx = (fecha_futura.year - fecha_inicio_forecast.year) * 12 + (fecha_futura.month - fecha_inicio_forecast.month)
            monthly_trend = math.pow(trend_rate, month_idx)
            quincena_factor = 1.10 if day_m in [1, 15, 16, 30, 31] else 1.0
            
            vol_final = vol_base * monthly_trend * quincena_factor
            
            # === CANDADO ANTI-ZOMBIES ===
            if sub[col_calls].mean() < 1.0:
                vol_final = 0.0
                
            preds_finales.append(max(0.0, vol_final))

        predicciones_futuras[camp] = preds_finales
        factores_ui[camp] = round(trend_rate, 2)

    df['En_Ventana'] = [esta_en_ventana_servicio(c, i) for c, i in zip(df[col_camp], df['Inter_Clean'])]
    df_filtrado = df[df['En_Ventana']].copy()

    df_reciente = df_filtrado[df_filtrado[col_fecha] >= (max_fecha_real - timedelta(days=28))]
    if df_reciente.empty: df_reciente = df_filtrado.copy()
    
    perfil_intradia = df_reciente.groupby([col_camp, 'Dia_Semana_Clean', 'Inter_Clean']).agg(
        avg_calls=(col_calls, 'mean'),
        avg_aht=(col_aht, lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 0)
    ).reset_index()

    totales_dia = perfil_intradia.groupby([col_camp, 'Dia_Semana_Clean'])['avg_calls'].transform('sum')
    perfil_intradia['weight'] = [(c / t) if t > 0 else 0 for c, t in zip(perfil_intradia['avg_calls'], totales_dia)]

    mapa_perfil = {(r[col_camp], r['Dia_Semana_Clean'], r['Inter_Clean']): {'weight': r['weight'], 'aht': r['avg_aht']} for _, r in perfil_intradia.iterrows()}
    todos_los_intervalos_crudos = [f"{int(h):02d}:{int(m):02d}" for h in range(24) for m in (0, 30)]
    intervalos_operativos_por_camp = {camp: [i for i in todos_los_intervalos_crudos if esta_en_ventana_servicio(camp, i)] for camp in campanas_unicas}

    del df_raw, df, df_diario, df_filtrado, df_reciente
    gc.collect()

    factor_asistencia = max(0.01, 1.0 - merma)
    data_processed = []

    for d in range(dias_futuros):
        fecha_actual = fecha_inicio_forecast + timedelta(days=d)
        str_fecha = fecha_actual.strftime('%Y-%m-%d')
        str_mes = f"{meses_espanol[fecha_actual.month]} {fecha_actual.year}"
        nombre_dia = dias_espanol[fecha_actual.weekday()]

        for camp in campanas_unicas:
            vol_diario = predicciones_futuras.get(camp, [0]*dias_futuros)[d]
            factor_visual_ui = factores_ui.get(camp, 1.0)
            intervalos_validos = intervalos_operativos_por_camp.get(camp, [])

            pesos_crudos = [mapa_perfil.get((camp, nombre_dia, inter), {}).get('weight', 0.0) for inter in intervalos_validos]
            suma_pesos = sum(pesos_crudos)
            if suma_pesos > 0: pesos_norm = [p / suma_pesos for p in pesos_crudos]
            elif len(intervalos_validos) > 0: pesos_norm = [1.0 / len(intervalos_validos)] * len(intervalos_validos)
            else: pesos_norm = []

            exact_calls = [vol_diario * p for p in pesos_norm]
            floor_calls = [int(math.floor(c)) for c in exact_calls]
            remainders = [(exact_calls[i] - floor_calls[i], i) for i in range(len(exact_calls))]
            remainders.sort(reverse=True, key=lambda x: x[0])
            
            diff = int(round(vol_diario)) - sum(floor_calls)
            for i in range(diff):
                if i < len(remainders): floor_calls[remainders[i][1]] += 1

            for idx_inter, inter in enumerate(intervalos_validos):
                calls_int = floor_calls[idx_inter]
                calls_float = exact_calls[idx_inter] 

                info_p = mapa_perfil.get((camp, nombre_dia, inter), {})
                aht_real = info_p.get('aht', 0.0)
                aht = aht_real if (aht_real > 0 and not pd.isna(aht_real)) else aht_global_campana.get(camp, 180.0)

                a_erlang = (calls_float * aht) / 1800.0 if (aht > 0 and calls_float > 0) else 0.0
                req_ftes = calcular_agentes_requeridos_erlang_c(a_erlang, aht, target_time, target_sl) if calls_float > 0 else 0
                req_hc = math.ceil(req_ftes / factor_asistencia) if req_ftes > 0 else 0
                
                hc_roster = roster_coverage.get((str(camp), nombre_dia.capitalize(), inter), 0)
                tot_camp = roster_total_camp.get(str(camp), 0)
                tot_camp_dia = roster_total_dia_camp.get((str(camp), nombre_dia.capitalize()), 0)

                data_processed.append({
                    'Campaña': str(camp),
                    'Fecha': str_fecha,
                    'Mes': str_mes,
                    'Día_Semana': nombre_dia.capitalize(),
                    'Intervalo': inter,
                    'Llamadas': calls_int,
                    'AHT': format_aht_str(aht),
                    'AHT_Segundos': int(round(aht)),
                    'Agentes_Requeridos': req_hc,
                    'HC_Actual_Roster': hc_roster,
                    'Total_Roster_Campana': tot_camp,
                    'Total_Roster_Dia': tot_camp_dia,
                    'Factor_Correccion': factor_visual_ui
                })

    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f: json.dump(data_processed, f)
    except: pass
    return data_processed

def resolver_turnos_optimos(intervalos, campanas_activas, llamadas_vec=None, aht_vec=None, req_vec=None, target_sl=80.0, target_time=20.0, merma=0.20, duracion_jornada=8.0, es_nocturno=False):
    m = len(intervalos)
    if m == 0: return [], [0]*m, 0, 0, 100.0, [100.0]*m, 100.0, 100.0, [0]*m

    try:
        llamadas_arr = np.array([float(x) if (x is not None and str(x).lower() != 'nan') else 0.0 for x in llamadas_vec], dtype=float)
        aht_arr = np.array([float(x) if (x is not None and str(x).lower() != 'nan') else 180.0 for x in aht_vec], dtype=float)
        req_hc_base = np.array([int(x) if (x is not None and str(x).lower() != 'nan') else 0 for x in req_vec], dtype=int)
    except:
        llamadas_arr = np.zeros(m)
        aht_arr = np.full(m, 180.0)
        req_hc_base = np.zeros(m)
    
    tot_llamadas = float(np.sum(llamadas_arr))
    factor_asistencia = max(0.01, 1.0 - merma)
    target_sl_dinamico = float(target_sl)
    req_hc_pooled = req_hc_base.tolist()
    cob_hc = np.zeros(m, dtype=float)
    x_turnos_dict = {}

    agentes_nocturnos_totales_hc = 0
    agentes_diurnos_totales_hc = 0

    if es_nocturno:
        label_jornada_noc = "9.0 hrs (Nocturno 5x2)"
        indices_nocturnos = [j for j in range(m) if parse_time_str(intervalos[j]) is not None and (parse_time_str(intervalos[j]) >= (22 * 60) or parse_time_str(intervalos[j]) < (7 * 60))]

        if len(indices_nocturnos) > 0 and sum([llamadas_arr[idx] for idx in indices_nocturnos]) > 0:
            agentes_noc_hc = 1
            while agentes_noc_hc <= 200:
                cob_temp_ftes = agentes_noc_hc * factor_asistencia
                sl_acum, llamadas_noc = 0.0, 0.0
                for idx in indices_nocturnos:
                    c, aht_s = llamadas_arr[idx], aht_arr[idx]
                    a_erl = (c * aht_s) / 1800.0 if (c > 0 and aht_s > 0) else 0.0
                    sl_v = erlang_c_sl_optimizado(a_erl, cob_temp_ftes, aht_s, target_time) if c > 0 else 100.0
                    sl_acum += (c * sl_v); llamadas_noc += c
                if (sl_acum / llamadas_noc if llamadas_noc > 0 else 100.0) >= target_sl_dinamico: break
                agentes_noc_hc += 1
            x_turnos_dict[("22:00", "07:00", label_jornada_noc)] = agentes_noc_hc
            agentes_nocturnos_totales_hc = agentes_noc_hc
            for idx in indices_nocturnos: cob_hc[idx] += agentes_noc_hc

    duracion_minutos = int(round(float(duracion_jornada) * 60))
    SHIFT_BLOCKS = int(round(float(duracion_jornada) * 2))
    label_jornada_diurna = f"{float(duracion_jornada):.1f} hrs".replace('.0', '')

    valid_starts = []
    for j in range(m):
        t_min = parse_time_str(intervalos[j])
        if t_min is not None:
            # Restricción diurna: El turno debe empezar a las 07:00 o después (>= 420 mins) 
            # y terminar a las 22:00 o antes (<= 1320 mins)
            if t_min >= 420 and (t_min + duracion_minutos) <= 1320:
                valid_starts.append(j)

    def calc_current_global_sl(current_cob):
        if tot_llamadas <= 0: return 100.0
        sl_acum = sum([c * erlang_c_sl_optimizado((c * aht_arr[i]) / 1800.0, current_cob[i] * factor_asistencia, aht_arr[i], target_time) for i, c in enumerate(llamadas_arr) if c > 0])
        return sl_acum / tot_llamadas

    if len(valid_starts) > 0:
        for _ in range(5000):
            current_sl = calc_current_global_sl(cob_hc)
            if current_sl >= target_sl_dinamico: break

            deficit = req_hc_base - cob_hc
            best_start_idx, best_cov, best_pen = -1, -1, 999999
            
            for s_idx in valid_starts:
                sub_def = deficit[s_idx : s_idx + SHIFT_BLOCKS] if s_idx + SHIFT_BLOCKS <= m else np.concatenate((deficit[s_idx:], deficit[:(s_idx + SHIFT_BLOCKS) - m]))
                cov, pen = np.sum(np.maximum(0, sub_def)), np.sum(np.maximum(0, -sub_def))
                if cov > best_cov or (cov == best_cov and pen < best_pen): 
                    best_cov, best_pen, best_start_idx = cov, pen, s_idx

            if best_start_idx == -1 or best_cov <= 0: break
                
            min_in = parse_time_str(intervalos[best_start_idx])
            if min_in is None: min_in = 0 
            
            min_out = (min_in + duracion_minutos) % (24 * 60)
            key_turno = (f"{(int(min_in // 60)):02d}:{(int(min_in % 60)):02d}", f"{(int(min_out // 60)):02d}:{(int(min_out % 60)):02d}", label_jornada_diurna)
            x_turnos_dict[key_turno] = x_turnos_dict.get(key_turno, 0) + 1
            
            if best_start_idx + SHIFT_BLOCKS <= m: cob_hc[best_start_idx : best_start_idx + SHIFT_BLOCKS] += 1
            else: cob_hc[best_start_idx:] += 1; cob_hc[:(best_start_idx + SHIFT_BLOCKS) - m] += 1

    sl_optimo_vector = [float(erlang_c_sl_optimizado((llamadas_arr[i] * aht_arr[i]) / 1800.0 if (llamadas_arr[i] > 0 and aht_arr[i] > 0) else 0.0, cob_hc[i] * factor_asistencia, aht_arr[i], target_time) if llamadas_arr[i] > 0 else 100.0) for i in range(m)]
    sl_optimo_global = float(np.sum(llamadas_arr * np.array(sl_optimo_vector)) / tot_llamadas) if tot_llamadas > 0 else 100.0

    cobertura_hc_entera = [int(x) for x in np.round(cob_hc)]
    turnos_sugeridos = []
    total_agentes_diarios_hc = 0

    for (h_in, h_out, label_dur), qty in x_turnos_dict.items():
        if qty > 0:
            turnos_sugeridos.append({'horario_entrada': h_in, 'horario_salida': h_out, 'agentes_a_programar': int(qty), 'duracion': label_dur})
            total_agentes_diarios_hc += int(qty)
            if "Nocturno" not in label_dur: agentes_diurnos_totales_hc += int(qty)

    turnos_sugeridos = sorted(turnos_sugeridos, key=lambda x: parse_time_str(x['horario_entrada']) or 0)
    hc_nocturno = math.ceil(agentes_nocturnos_totales_hc * (7.0 / 5.0))
    hc_diurno = math.ceil(agentes_diurnos_totales_hc * (7.0 / 6.0))
    total_req_hc_pooled = float(np.sum(req_hc_pooled))
    total_prog_hc = float(np.sum(cob_hc))

    return turnos_sugeridos, cobertura_hc_entera, total_agentes_diarios_hc, int(hc_nocturno + hc_diurno), float(min(100.0, (total_req_hc_pooled / total_prog_hc) * 100.0)) if total_prog_hc > 0 else 100.0, sl_optimo_vector, sl_optimo_global, float((total_prog_hc / total_req_hc_pooled) * 100.0) if total_req_hc_pooled > 0 else 100.0, req_hc_pooled

@app.route('/api/latest', methods=['GET'])
def get_latest_forecast():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
                if isinstance(cache_data, list) and len(cache_data) > 0: 
                    return jsonify(cache_data), 200
        except Exception as e:
            print(f"Error leyendo caché: {e}")
            pass
            
    return jsonify([]), 200

@app.route('/api/optimize-schedules', methods=['POST'])
def api_optimize_schedules():
    try:
        body = request.get_json(force=True)
        turnos, cob_optima, total_diario, total_hc, eficiencia, sl_vec, sl_global, staff_level, req_hc_pooled = resolver_turnos_optimos(
            body.get('intervalos', []), body.get('campanas', []), body.get('llamadas', []), body.get('ahts', []), body.get('requeridos', []),
            float(body.get('target_sl', 80.0)), float(body.get('target_time', 20.0)), float(body.get('merma', 30.0)) / 100.0, float(body.get('duracion_jornada', 8.0)), bool(body.get('es_nocturno', False))
        )
        return jsonify({
            'turnos': turnos, 'cobertura_optima': [int(x) for x in cob_optima], 'total_agentes_diarios': int(total_diario),
            'headcount_semanal_6x1': int(total_hc), 'eficiencia_cobertura': float(eficiencia), 'sl_optimo_vector': [float(x) for x in sl_vec],
            'sl_optimo_global': float(sl_global), 'staffing_level_optimo': float(staff_level), 'req_hc_pooled': [int(x) for x in req_hc_pooled]
        }), 200
    except Exception as e: return jsonify({'error': f'Error optimizando turnos: {str(e)}'}), 500

@app.route('/api/process', methods=['POST', 'GET'])
def process_data():
    if request.method == 'GET': return jsonify({'status': 'API activa'}), 200
    excel_path = buscar_archivo_excel()
    if not excel_path: return jsonify({'error': 'No se encontro Excel (.xlsx).'}), 400
    try:
        data = procesar_archivo_excel(excel_path, float(clean_num(request.form.get('target_sl'), 80.0)), float(clean_num(request.form.get('target_time'), 20.0)), float(clean_num(request.form.get('merma'), 20.0)) / 100.0, int(clean_num(request.form.get('dias'), 45)))
        gc.collect()
        return jsonify(data)
    except Exception as e:
        gc.collect()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
