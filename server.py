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

# --- BUSCADOR UNIVERSAL DE ARCHIVO EXCEL ---
def buscar_archivo_excel():
    if os.path.exists(EXCEL_DEFAULT):
        return EXCEL_DEFAULT
    try:
        archivos = [f for f in os.listdir(BASE_DIR) if f.lower().endswith('.xlsx') and not f.startswith('~')]
        if not archivos:
            return None
        for f in archivos:
            if 'historico' in f.lower():
                return os.path.join(BASE_DIR, f)
        return os.path.join(BASE_DIR, archivos[0])
    except Exception:
        return None

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
    return jsonify({"error": "ALERTA CRÍTICA: No se encontró el archivo index.html en el servidor."}), 404

@app.route('/favicon.ico')
def favicon():
    return '', 204

@app.route('/logo.png')
def serve_logo():
    return send_from_directory(BASE_DIR, 'logo.png')

@app.route('/api/config', methods=['GET', 'POST'])
def manage_config():
    if request.method == 'POST':
        try:
            new_config = request.get_json(force=True)
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(new_config, f)
            return jsonify({'status': 'Configuración guardada exitosamente'}), 200
        except Exception as e:
            return jsonify({'error': str(e)}), 500
    else:
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return jsonify(json.load(f)), 200
            except:
                pass
        return jsonify({
            'targetSl': 80,
            'targetTime': 20,
            'merma': 30,
            'duracionJornada': 8,
            'chkNocturno': False,
            'chkPicos': False
        }), 200

def clean_num(val, default=0.0):
    if pd.isna(val) or val is None: return default
    try:
        val_str = str(val).strip().replace(',', '.')
        val_str = re.sub(r'[^0-9.]', '', val_str)
        return float(val_str) if val_str else default
    except Exception:
        return default

def parse_aht_to_seconds(val):
    if pd.isna(val) or val is None: return 180.0
    secs = 180.0
    if isinstance(val, (int, float)):
        secs = float(val)
    elif hasattr(val, 'hour') and hasattr(val, 'minute') and hasattr(val, 'second'):
        secs = val.hour * 3600 + val.minute * 60 + val.second
    else:
        val_str = str(val).strip()
        if ':' in val_str:
            parts = val_str.split(':')
            try:
                if len(parts) == 3: secs = int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
                elif len(parts) == 2: secs = int(parts[0]) * 60 + float(parts[1])
            except: pass
        else:
            try: secs = float(val_str)
            except: pass
    if 0 < secs <= 15: secs = secs * 60.0
    return secs if secs > 0 else 180.0

def format_aht_str(seconds):
    if pd.isna(seconds) or seconds is None or seconds <= 0: return "00:00:00"
    secs = int(round(seconds))
    hrs = secs // 3600
    mins = (secs % 3600) // 60
    s = secs % 60
    return f"{hrs:02d}:{mins:02d}:{s:02d}"

ERLANG_CACHE = {}

def erlang_c_sl_optimizado(A, N, AHT, target_time):
    if N <= A or A <= 0 or N <= 0: return 0.0
    key = (round(A, 2), N, round(AHT, 1), target_time)
    if key in ERLANG_CACHE:
        return ERLANG_CACHE[key]
        
    try:
        sum_terms, current_term = 1.0, 1.0
        int_N = min(int(N), 1000)
        for k in range(1, int_N):
            current_term *= (A / k)
            sum_terms += current_term
        last_term = current_term * (A / N) / (1.0 - (A / N))
        pw = last_term / (sum_terms + last_term)
        intensity = N - A
        sl = 1.0 - (pw * math.exp(-intensity * (target_time / AHT)))
        resultado = round(max(0.0, min(100.0, sl * 100.0)), 1)
        ERLANG_CACHE[key] = resultado
        return resultado
    except: return 0.0

def calcular_agentes_requeridos_erlang_c(A, aht, target_time, target_sl):
    if A <= 0 or aht <= 0: return 0
    n = max(1, int(math.floor(A)) + 1)
    while n < 1000:
        if erlang_c_sl_optimizado(A, n, aht, target_time) >= target_sl: return n
        n += 1
    return n

def parse_time_str(t_str):
    if not t_str: return None
    t = str(t_str).lower().replace('hrs', '').replace(' ', '')
    is_pm = 'pm' in t
    is_am = 'am' in t
    t = t.replace('am', '').replace('pm', '')
    t = re.sub(r'[^\d:]', '', t)
    if not t: return None
    if ':' not in t: t += ':00'
    try:
        parts = t.split(':')
        hh, mm = int(parts[0]), int(parts[1])
        if is_pm and hh < 12: hh += 12
        if is_am and hh == 12: hh = 0
        return hh * 60 + mm
    except: return None

def esta_en_ventana_servicio(campana, intervalo_str):
    camp_key = str(campana).strip().lower()
    minutos_inter = parse_time_str(intervalo_str)
    
    if minutos_inter is None: 
        return True
        
    for key, ventana in VENTANAS_SERVICIO.items():
        if key in camp_key or camp_key in key:
            return ventana['inicio'] <= minutos_inter < ventana['fin']
            
    return True

def encontrar_columna(df, posibles_nombres):
    for pos in posibles_nombres:
        for col_orig in df.columns:
            col_clean = str(col_orig).strip().lower()
            if pos.strip().lower() == col_clean or pos.strip().lower() in col_clean:
                return col_orig
    return None

def calc_mae(y_true, y_pred):
    return float(np.mean(np.abs(np.array(y_true) - np.array(y_pred))))

def entrenar_ridge_ml(X, y, l2_reg=10.0):
    X_b = np.c_[np.ones((X.shape[0], 1)), X]
    mean = np.mean(X_b[:, 1:], axis=0)
    std = np.std(X_b[:, 1:], axis=0) + 1e-8
    X_norm = X_b.copy()
    X_norm[:, 1:] = (X_b[:, 1:] - mean) / std
    I = np.eye(X_norm.shape[1])
    I[0, 0] = 0.0
    try:
        weights = np.linalg.inv(X_norm.T @ X_norm + l2_reg * I) @ X_norm.T @ y
    except np.linalg.LinAlgError:
        weights = np.linalg.pinv(X_norm.T @ X_norm + l2_reg * I) @ X_norm.T @ y
    return weights, mean, std

def predecir_ridge_ml(weights, mean, std, X_new):
    n_rows = X_new.shape[0] if hasattr(X_new, 'shape') else len(X_new)
    X_b = np.c_[np.ones((n_rows, 1)), X_new]
    X_norm = X_b.copy()
    X_norm[:, 1:] = (X_b[:, 1:] - mean) / std
    pred = X_norm @ weights
    return float(pred[0])

def extraer_features_fecha(fecha, volumenes_hist, trend_idx):
    day_of_week = fecha.weekday()
    day_of_month = fecha.day
    is_weekend = 1.0 if day_of_week >= 5 else 0.0
    is_quincena = 1.0 if day_of_month in [1, 15, 16, 30, 31] else 0.0
    lag_1 = volumenes_hist[-1] if len(volumenes_hist) >= 1 else 100.0
    lag_7 = volumenes_hist[-7] if len(volumenes_hist) >= 7 else lag_1
    lag_14 = volumenes_hist[-14] if len(volumenes_hist) >= 14 else lag_7
    dow_encoded = [1.0 if day_of_week == i else 0.0 for i in range(7)]
    return [lag_1, lag_7, lag_14, float(day_of_month), is_weekend, is_quincena, float(trend_idx)] + dow_encoded

def holt_winters_fit_predict(series, season_len=7, alpha=0.2, beta=0.1, gamma=0.3, n_preds=30):
    n = len(series)
    if n < season_len * 2:
        return [np.mean(series) if len(series) > 0 else 100.0] * n_preds
    level = np.mean(series[:season_len])
    trend = (np.mean(series[season_len:2*season_len]) - np.mean(series[:season_len])) / season_len
    seasonals = [series[i] - level for i in range(season_len)]
    for i in range(n):
        val = series[i]
        last_level, last_trend = level, trend
        st_prev = seasonals[i % season_len]
        level = alpha * (val - st_prev) + (1 - alpha) * (last_level + last_trend)
        trend = beta * (level - last_level) + (1 - beta) * last_trend
        seasonals[i % season_len] = gamma * (val - level) + (1 - gamma) * st_prev
    preds = []
    for m in range(1, n_preds + 1):
        p = level + m * trend + seasonals[(n + m - 1) % season_len]
        preds.append(max(0.0, float(p)))
    return preds

def grid_search_auto_hw(series, n_preds=30):
    return holt_winters_fit_predict(series, season_len=7, alpha=0.2, beta=0.1, gamma=0.3, n_preds=n_preds)

def limpiar_outliers_iqr(series_list):
    if len(series_list) < 14:
        return list(series_list)
    arr = np.array(series_list)
    q25, q75 = np.percentile(arr, 25), np.percentile(arr, 75)
    iqr = q75 - q25
    lower, upper = q25 - 1.5 * iqr, q75 + 1.5 * iqr
    return np.clip(arr, lower, upper).tolist()

def generar_intervalos_cobertura(start_min, end_min):
    intervals = []
    if start_min < end_min:
        curr = start_min
        while curr < end_min:
            hh = curr // 60
            mm = curr % 60
            intervals.append(f"{int(hh):02d}:{int(mm):02d}")
            curr += 30
    else: 
        curr = start_min
        while curr < 24 * 60:
            hh = curr // 60
            mm = curr % 60
            intervals.append(f"{int(hh):02d}:{int(mm):02d}")
            curr += 30
        curr = 0
        while curr < end_min:
            hh = curr // 60
            mm = curr % 60
            intervals.append(f"{int(hh):02d}:{int(mm):02d}")
            curr += 30
    return intervals

def procesar_hoja_roster(df_roster):
    dias_map = {'lunes': 'Lunes', 'martes': 'Martes', 'miércoles': 'Miércoles', 'miercoles': 'Miércoles', 
                'jueves': 'Jueves', 'viernes': 'Viernes', 'sábado': 'Sábado', 'sabado': 'Sábado', 'domingo': 'Domingo'}
    
    roster_cov = {} 
    roster_total_camp = {}
    roster_total_dia_camp = {} 
    
    col_camp = encontrar_columna(df_roster, ['campaña', 'campana', 'skill', 'servicio'])
    if not col_camp:
        return roster_cov, roster_total_camp, roster_total_dia_camp
        
    for idx, row in df_roster.iterrows():
        camp = str(row[col_camp]).strip().title()
        if camp == 'Nan' or camp == '': continue
        
        roster_total_camp[camp] = roster_total_camp.get(camp, 0) + 1
        
        for col in df_roster.columns:
            col_lower = str(col).lower().strip()
            if col_lower in dias_map:
                dia_real = dias_map[col_lower]
                horario = str(row[col]).strip().upper()
                
                if horario != 'DD-DD' and 'NAN' not in horario and horario != '' and '-' in horario:
                    key_dia = (camp, dia_real)
                    roster_total_dia_camp[key_dia] = roster_total_dia_camp.get(key_dia, 0) + 1
                    
                    parts = horario.split('-')
                    if len(parts) == 2:
                        start_min = parse_time_str(parts[0].strip())
                        end_min = parse_time_str(parts[1].strip())
                        
                        if start_min is not None and end_min is not None:
                            intervals = generar_intervalos_cobertura(start_min, end_min)
                            for inv in intervals:
                                key = (camp, dia_real, inv)
                                roster_cov[key] = roster_cov.get(key, 0) + 1
                                
    return roster_cov, roster_total_camp, roster_total_dia_camp

def procesar_archivo_excel(file_source, target_sl=80.0, target_time=20.0, merma=0.20, dias_futuros=30):
    xls_file = pd.ExcelFile(file_source, engine='openpyxl')
    
    sheet_calls = xls_file.sheet_names[0]
    for s in xls_file.sheet_names:
        if 'llam' in s.lower() or 'hist' in s.lower() or 'datos' in s.lower():
            sheet_calls = s
            break
            
    sheet_roster = None
    for s in xls_file.sheet_names:
        if 'roster' in s.lower() or 'plantilla' in s.lower() or 'horario' in s.lower():
            sheet_roster = s
            break

    roster_coverage = {}
    roster_total_camp = {}
    roster_total_dia_camp = {}
    
    if sheet_roster:
        try:
            df_roster = pd.read_excel(xls_file, sheet_name=sheet_roster, engine='openpyxl')
            roster_coverage, roster_total_camp, roster_total_dia_camp = procesar_hoja_roster(df_roster)
        except Exception as e:
            pass

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
    
    df_raw[col_fecha] = pd.to_datetime(df_raw[col_fecha], errors='coerce', dayfirst=True).dt.normalize()
    df_raw = df_raw.dropna(subset=[col_fecha])
    if df_raw.empty:
        raise ValueError("Error Crítico: No se pudieron leer las fechas. Revisa que la columna de fechas en tu Excel.")
    
    df_raw[col_calls] = [clean_num(x, 0.0) for x in df_raw[col_calls]]

    df_valido = df_raw[df_raw[col_calls] > 0]
    if df_valido.empty:
        raise ValueError("Error Crítico: El archivo no tiene volumen de llamadas mayor a cero.")
    
    max_fecha_real = df_valido[col_fecha].max()
    df_raw = df_raw[df_raw[col_fecha] <= max_fecha_real]

    if col_aht:
        df_raw[col_aht] = [parse_aht_to_seconds(x) for x in df_raw[col_aht]]
    else:
        df_raw['AHT_Calc'] = 180.0
        col_aht = 'AHT_Calc'

    df_raw['Total_Segundos_Handle'] = df_raw[col_calls] * df_raw[col_aht]
    df_raw['Inter_Clean'] = [':'.join(str(x).strip().split(':')[:2]) if len(str(x).strip().split(':')) == 3 else str(x).strip() for x in df_raw[col_inter]]

    df = df_raw.groupby([col_fecha, col_camp, 'Inter_Clean']).agg({
        col_calls: 'sum',
        'Total_Segundos_Handle': 'sum'
    }).reset_index()

    df[col_aht] = np.where(df[col_calls] > 0, df['Total_Segundos_Handle'] / df[col_calls], 180.0)
    df = df.drop(columns=['Total_Segundos_Handle'])
    df[col_inter] = df['Inter_Clean']

    dias_espanol = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']
    meses_espanol = ['', 'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

    df['Dia_Semana_Clean'] = df[col_fecha].dt.weekday.apply(lambda w: dias_espanol[w])

    fecha_maxima = df[col_fecha].max()
    if pd.isna(fecha_maxima):
        fecha_maxima = pd.to_datetime(datetime.now()).normalize()
        
    fecha_inicio_forecast = fecha_maxima + timedelta(days=1)
    
    aht_global_campana = df.groupby(col_camp)[col_aht].apply(lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 180.0).to_dict()

    df_diario = df.groupby([col_fecha, col_camp])[col_calls].sum().reset_index()
    campanas_unicas = df[col_camp].unique()

    modelos_ml, historial_volumenes, hw_forecasts, pesos_campana = {}, {}, {}, {}

    for camp in campanas_unicas:
        sub = df_diario[df_diario[col_camp] == camp].sort_values(col_fecha).reset_index(drop=True)
        fechas_list = sub[col_fecha].tolist()
        volumenes_list = limpiar_outliers_iqr(sub[col_calls].tolist())
        historial_volumenes[camp] = list(volumenes_list)
        
        vols = list(volumenes_list)
        n = len(vols)
        
        peso_hw = 0.65
        peso_ridge = 0.35
        factor_ajuste = 1.0

        if n >= 21:
            train_vols = vols[:-7]
            val_vols = vols[-7:]
            
            hw_val_preds = grid_search_auto_hw(train_vols, n_preds=7)
            
            X_train_bt, y_train_bt = [], []
            for i in range(14, len(train_vols)):
                feat = extraer_features_fecha(fechas_list[i], train_vols[:i], trend_idx=i)
                X_train_bt.append(feat)
                y_train_bt.append(train_vols[i])
            
            if len(X_train_bt) > 5:
                w_bt, m_bt, s_bt = entrenar_ridge_ml(np.array(X_train_bt), np.array(y_train_bt), l2_reg=10.0)
                ridge_val_preds = []
                for i in range(7):
                    f_idx = len(train_vols) + i
                    feat = extraer_features_fecha(fechas_list[f_idx], train_vols + val_vols[:i], trend_idx=f_idx)
                    pred_r = predecir_ridge_ml(w_bt, m_bt, s_bt, np.array([feat]))
                    ridge_val_preds.append(max(0, pred_r))
                
                error_hw = calc_mae(val_vols, hw_val_preds) + 1e-5
                error_ridge = calc_mae(val_vols, ridge_val_preds) + 1e-5
                
                total_inv_error = (1.0 / error_hw) + (1.0 / error_ridge)
                peso_hw = (1.0 / error_hw) / total_inv_error
                peso_ridge = (1.0 / error_ridge) / total_inv_error
                
                ensemble_val_preds = [(peso_hw * hw_val_preds[i] + peso_ridge * ridge_val_preds[i]) for i in range(7)]
                sum_actual = sum(val_vols)
                sum_pred = sum(ensemble_val_preds)
                
                if sum_pred > 0:
                    drift = sum_actual / sum_pred
                    factor_ajuste = max(0.5, min(1.5, drift))
                    
        pesos_campana[camp] = {
            'w_hw': peso_hw,
            'w_ridge': peso_ridge,
            'factor_ajuste': factor_ajuste
        }

        hw_forecasts[camp] = grid_search_auto_hw(vols, n_preds=dias_futuros)

        X_data, y_data = [], []
        for i in range(14, len(vols)):
            f = fechas_list[i]
            feat = extraer_features_fecha(f, vols[:i], trend_idx=i)
            X_data.append(feat)
            y_data.append(vols[i])

        if len(X_data) > 10:
            X_arr, y_arr = np.array(X_data), np.array(y_data)
            weights, mean, std = entrenar_ridge_ml(X_arr, y_arr, l2_reg=10.0)
            modelos_ml[camp] = {'weights': weights, 'mean': mean, 'std': std, 'promedio_base': np.mean(y_arr)}
        else:
            modelos_ml[camp] = None

    df['En_Ventana'] = [esta_en_ventana_servicio(c, i) for c, i in zip(df[col_camp], df['Inter_Clean'])]
    df_filtrado = df[df['En_Ventana']].copy()

    max_date_hist = df_filtrado[col_fecha].max()
    
    if pd.isna(max_date_hist):
        df_reciente = df_filtrado.copy()
    else:
        df_reciente = df_filtrado[df_filtrado[col_fecha] >= (max_date_hist - timedelta(days=21))]

    perfil_intradia = df_reciente.groupby([col_camp, 'Dia_Semana_Clean', 'Inter_Clean']).agg(
        avg_calls=(col_calls, 'mean'),
        avg_aht=(col_aht, lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 0)
    ).reset_index()

    totales_dia = perfil_intradia.groupby([col_camp, 'Dia_Semana_Clean'])['avg_calls'].transform('sum')
    perfil_intradia['weight'] = [(c / t) if t > 0 else 0 for c, t in zip(perfil_intradia['avg_calls'], totales_dia)]

    mapa_perfil = {}
    for _, r in perfil_intradia.iterrows():
        key = (r[col_camp], r['Dia_Semana_Clean'], r['Inter_Clean'])
        mapa_perfil[key] = {'weight': r['weight'], 'aht': r['avg_aht']}

    TODOS_LOS_INTERVALOS = [f"{int(h):02d}:{int(m):02d}" for h in range(24) for m in (0, 30)]
    intervalos_operativos_por_camp = {}
    for camp in campanas_unicas:
        intervalos_operativos_por_camp[camp] = sorted([i for i in TODOS_LOS_INTERVALOS if esta_en_ventana_servicio(camp, i)])

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
            hist_vol = historial_volumenes[camp]
            feat_futuras = np.array([extraer_features_fecha(fecha_actual, hist_vol, len(hist_vol))])

            model_info = modelos_ml.get(camp)
            if model_info:
                vol_ridge = predecir_ridge_ml(model_info['weights'], model_info['mean'], model_info['std'], feat_futuras)
                vol_ridge = max(vol_ridge, model_info['promedio_base'] * 0.15)
            else:
                vol_ridge = np.mean(hist_vol[-7:]) if hist_vol else 100.0

            w_info = pesos_campana.get(camp, {'w_hw': 0.65, 'w_ridge': 0.35, 'factor_ajuste': 1.0})
            vol_hw = hw_forecasts[camp][d] if d < len(hw_forecasts[camp]) else vol_ridge
            
            volumen_predicho_diario = (w_info['w_hw'] * vol_hw + w_info['w_ridge'] * vol_ridge) * w_info['factor_ajuste']
            
            historial_volumenes[camp].append(volumen_predicho_diario)

            intervalos_validos = intervalos_operativos_por_camp.get(camp, [])

            for inter in intervalos_validos:
                key_p = (camp, nombre_dia, inter)
                info_p = mapa_perfil.get(key_p, {'weight': 0.0, 'aht': 0.0})
                calls = volumen_predicho_diario * info_p['weight']
                calls_int = int(round(calls))

                if calls_int == 0:
                    aht = 0.0
                else:
                    aht = info_p['aht'] if (info_p['aht'] > 0 and not pd.isna(info_p['aht'])) else aht_global_campana.get(camp, 180.0)

                a_erlang = (calls * aht) / 1800.0 if (aht > 0 and calls > 0) else 0.0
                req_ftes = calcular_agentes_requeridos_erlang_c(a_erlang, aht, target_time, target_sl) if calls > 0 else 0
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
                    'Factor_Correccion': round(float(w_info['factor_ajuste']), 2)
                })

    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(data_processed, f)
    except Exception as err:
        pass

    return data_processed

def resolver_turnos_optimos(intervalos, campanas_activas, llamadas_vec=None, aht_vec=None, req_vec=None, target_sl=80.0, target_time=20.0, merma=0.20, duracion_jornada=8.0, es_nocturno=False):
    m = len(intervalos)
    if m == 0:
        return [], [0]*m, 0, 0, 100.0, [100.0]*m, 100.0, 100.0, [0]*m

    llamadas_arr = np.nan_to_num(np.array(llamadas_vec, dtype=float), nan=0.0) if llamadas_vec is not None else np.zeros(m)
    aht_arr = np.nan_to_num(np.array(aht_vec, dtype=float), nan=180.0) if aht_vec is not None else np.full(m, 180.0)
    
    tot_llamadas = float(np.sum(llamadas_arr))
    factor_asistencia = max(0.01, 1.0 - merma)
    target_sl_dinamico = float(target_sl)
    
    req_hc_pooled = []
    req_hc_base = np.zeros(m)
    
    for i in range(m):
        if req_vec is not None and i < len(req_vec) and req_vec[i] > 0:
            req_hc_i = int(req_vec[i])
        else:
            c = llamadas_arr[i]
            aht_s = aht_arr[i]
            a_erl = (c * aht_s) / 1800.0 if (c > 0 and aht_s > 0) else 0.0
            req_ftes_i = calcular_agentes_requeridos_erlang_c(a_erl, aht_s, target_time, target_sl_dinamico) if c > 0 else 0
            req_hc_i = math.ceil(req_ftes_i / factor_asistencia) if req_ftes_i > 0 else 0
            
        req_hc_pooled.append(int(req_hc_i))
        req_hc_base[i] = req_hc_i

    cob_hc = np.zeros(m, dtype=float)
    x_turnos_dict = {}

    agentes_nocturnos_totales_hc = 0
    agentes_diurnos_totales_hc = 0

    if es_nocturno:
        label_jornada_noc = "9.0 hrs (Nocturno 5x2)"
        indices_nocturnos = []
        for j in range(m):
            min_in = parse_time_str(intervalos[j])
            if min_in is not None:
                if min_in >= (22 * 60) or min_in < (7 * 60):
                    indices_nocturnos.append(j)

        if len(indices_nocturnos) > 0:
            agentes_noc_hc = 1
            while agentes_noc_hc <= 200:
                cob_temp_ftes = agentes_noc_hc * factor_asistencia
                sl_acum, llamadas_noc = 0.0, 0.0
                for idx in indices_nocturnos:
                    c = llamadas_arr[idx]
                    aht_s = aht_arr[idx]
                    a_erl = (c * aht_s) / 1800.0 if (c > 0 and aht_s > 0) else 0.0
                    sl_v = erlang_c_sl_optimizado(a_erl, cob_temp_ftes, aht_s, target_time) if c > 0 else 100.0
                    sl_acum += (c * sl_v)
                    llamadas_noc += c
                sl_prom_noc = (sl_acum / llamadas_noc) if llamadas_noc > 0 else 100.0
                if sl_prom_noc >= target_sl_dinamico:
                    break
                agentes_noc_hc += 1
            key_turno_noc = ("22:00", "07:00", label_jornada_noc)
            x_turnos_dict[key_turno_noc] = agentes_noc_hc
            agentes_nocturnos_totales_hc = agentes_noc_hc
            for idx in indices_nocturnos:
                cob_hc[idx] += agentes_noc_hc

    duracion_jornada = float(duracion_jornada)
    SHIFT_BLOCKS = int(round(duracion_jornada * 2))
    duracion_minutos = int(round(duracion_jornada * 60))
    label_jornada_diurna = f"{duracion_jornada:.1f} hrs".replace('.0', '')

    # --- NUEVO CANDADO DE LÍMITE DE HORARIOS (07:00 a 22:00) ---
    valid_starts = []
    for j in range(m):
        min_in_val = parse_time_str(intervalos[j])
        if min_in_val is not None:
            min_out_val = min_in_val + duracion_minutos
            # Todo turno dinámico diurno debe comenzar >= 07:00 y terminar <= 22:00
            if min_in_val >= (7 * 60) and min_out_val <= (22 * 60):
                valid_starts.append(j)
    # ----------------------------------------------------------

    def calc_current_global_sl(current_cob):
        if tot_llamadas <= 0: return 100.0
        sl_acum = 0.0
        for i in range(m):
            c = llamadas_arr[i]
            if c > 0:
                a_erl = (c * aht_arr[i]) / 1800.0
                n_opt = current_cob[i] * factor_asistencia
                sl_v = erlang_c_sl_optimizado(a_erl, n_opt, aht_arr[i], target_time)
                sl_acum += c * sl_v
        return sl_acum / tot_llamadas

    if len(valid_starts) > 0:
        max_iterations = 5000
        iteration = 0
        while iteration < max_iterations:
            iteration += 1
            
            if calc_current_global_sl(cob_hc) >= target_sl_dinamico:
                break

            deficit = req_hc_base - cob_hc
            if np.max(deficit) <= 0:
                break 

            best_start_idx = -1
            best_score = -999999

            for s_idx in valid_starts:
                if s_idx + SHIFT_BLOCKS <= m:
                    sub_deficit = deficit[s_idx : s_idx + SHIFT_BLOCKS]
                else:
                    sub_deficit = np.concatenate((deficit[s_idx:], deficit[:(s_idx + SHIFT_BLOCKS) - m]))
                
                score = np.sum(np.maximum(0, sub_deficit)) - np.sum(np.maximum(0, -sub_deficit)) * 0.001
                
                if score > best_score:
                    best_score = score
                    best_start_idx = s_idx

            if best_start_idx == -1 or best_score <= 0.0001:
                break
                
            min_in_val = parse_time_str(intervalos[best_start_idx])
            min_out_val = min_in_val + duracion_minutos
            min_out_val = min_out_val % (24 * 60) 
            
            h_in_str = f"{(int(min_in_val // 60)):02d}:{(int(min_in_val % 60)):02d}"
            h_out_str = f"{(int(min_out_val // 60)):02d}:{(int(min_out_val % 60)):02d}"
            
            key_turno = (h_in_str, h_out_str, label_jornada_diurna)
            x_turnos_dict[key_turno] = x_turnos_dict.get(key_turno, 0) + 1
            
            if best_start_idx + SHIFT_BLOCKS <= m:
                cob_hc[best_start_idx : best_start_idx + SHIFT_BLOCKS] += 1
            else:
                cob_hc[best_start_idx:] += 1
                cob_hc[:(best_start_idx + SHIFT_BLOCKS) - m] += 1

    sl_optimo_vector = []
    for i in range(m):
        c = llamadas_arr[i]
        aht_s = aht_arr[i]
        n_opt_ftes = cob_hc[i] * factor_asistencia
        a_erl = (c * aht_s) / 1800.0 if (c > 0 and aht_s > 0) else 0.0
        sl_val = erlang_c_sl_optimizado(a_erl, n_opt_ftes, aht_s, target_time) if c > 0 else 100.0
        sl_optimo_vector.append(float(sl_val))

    sl_arr = np.array(sl_optimo_vector)
    sl_optimo_global = float(np.sum(llamadas_arr * sl_arr) / tot_llamadas) if tot_llamadas > 0 else 100.0

    cobertura_hc_entera = [int(x) for x in np.round(cob_hc)]
    turnos_sugeridos = []
    total_agentes_diarios_hc = 0

    for (h_in, h_out, label_dur), qty in x_turnos_dict.items():
        if qty > 0:
            turnos_sugeridos.append({
                'horario_entrada': h_in,
                'horario_salida': h_out,
                'agentes_a_programar': int(qty),
                'duracion': label_dur
            })
            total_agentes_diarios_hc += int(qty)
            if "Nocturno" not in label_dur:
                agentes_diurnos_totales_hc += int(qty)

    turnos_sugeridos = sorted(turnos_sugeridos, key=lambda x: parse_time_str(x['horario_entrada']) or 0)

    hc_nocturno = math.ceil(agentes_nocturnos_totales_hc * (7.0 / 5.0))
    hc_diurno = math.ceil(agentes_diurnos_totales_hc * (7.0 / 6.0))
    headcount_semanal_requerido = int(hc_nocturno + hc_diurno)

    total_req_hc_pooled = float(np.sum(req_hc_pooled))
    total_prog_hc = float(np.sum(cob_hc))
    
    if total_req_hc_pooled > 0:
        staffing_level_optimo = float((total_prog_hc / total_req_hc_pooled) * 100.0)
        eficiencia = float(min(100.0, (total_req_hc_pooled / total_prog_hc) * 100.0)) if total_prog_hc > 0 else 100.0
    else:
        staffing_level_optimo = 100.0
        eficiencia = 100.0

    return turnos_sugeridos, cobertura_hc_entera, total_agentes_diarios_hc, headcount_semanal_requerido, eficiencia, sl_optimo_vector, sl_optimo_global, staffing_level_optimo, req_hc_pooled

@app.route('/api/latest', methods=['GET'])
def get_latest_forecast():
    use_cache = False
    excel_path = buscar_archivo_excel()
    
    if os.path.exists(CACHE_FILE) and excel_path:
        if os.path.getmtime(CACHE_FILE) >= os.path.getmtime(excel_path):
            use_cache = True
        else:
            try: os.remove(CACHE_FILE)
            except: pass
    elif os.path.exists(CACHE_FILE):
        use_cache = True

    if use_cache:
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                return jsonify(json.load(f)), 200
        except Exception as e:
            pass
            
    if excel_path:
        try:
            data = procesar_archivo_excel(excel_path)
            return jsonify(data), 200
        except Exception as e:
            return jsonify({'error': f'Error procesando Excel: {str(e)}'}), 500
            
    return jsonify({'error': 'No se encontró ningún archivo Excel en el servidor.'}), 404

@app.route('/api/optimize-schedules', methods=['POST'])
def api_optimize_schedules():
    try:
        body = request.get_json(force=True)
        intervalos = body.get('intervalos', [])
        campanas = body.get('campanas', [])
        llamadas = body.get('llamadas', [])
        ahts = body.get('ahts', [])
        requeridos = body.get('requeridos', [])
        target_sl = float(body.get('target_sl', 80.0))
        target_time = float(body.get('target_time', 20.0))
        merma = float(body.get('merma', 30.0)) / 100.0
        duracion_jornada = float(body.get('duracion_jornada', 8.0))
        es_nocturno = bool(body.get('es_nocturno', False))

        turnos, cob_optima, total_diario, total_hc, eficiencia, sl_vec, sl_global, staff_level, req_hc_pooled = resolver_turnos_optimos(
            intervalos, campanas, llamadas_vec=llamadas, aht_vec=ahts, req_vec=requeridos,
            target_sl=target_sl, target_time=target_time, merma=merma, 
            duracion_jornada=duracion_jornada, es_nocturno=es_nocturno
        )
        return jsonify({
            'turnos': turnos,
            'cobertura_optima': [int(x) for x in cob_optima],
            'total_agentes_diarios': int(total_diario),
            'headcount_semanal_6x1': int(total_hc),
            'eficiencia_cobertura': float(eficiencia),
            'sl_optimo_vector': [float(x) for x in sl_vec],
            'sl_optimo_global': float(sl_global),
            'staffing_level_optimo': float(staff_level),
            'req_hc_pooled': [int(x) for x in req_hc_pooled]
        }), 200
    except Exception as e:
        return jsonify({'error': f'Error optimizando turnos: {str(e)}'}), 500

@app.route('/api/process', methods=['POST', 'GET'])
def process_data():
    if request.method == 'GET':
        return jsonify({'status': 'API predictiva activa'}), 200

    target_sl = float(clean_num(request.form.get('target_sl'), 80.0))
    target_time = float(clean_num(request.form.get('target_time'), 20.0))
    merma = float(clean_num(request.form.get('merma'), 20.0)) / 100.0
    dias_futuros = int(clean_num(request.form.get('dias'), 30))

    excel_path = buscar_archivo_excel()

    if not excel_path:
        return jsonify({'error': 'No se encontró ningún archivo Excel (.xlsx) en el repositorio.'}), 400

    try:
        data_processed = procesar_archivo_excel(excel_path, target_sl, target_time, merma, dias_futuros)
        gc.collect()
        return jsonify(data_processed)
    except Exception as e:
        gc.collect()
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
