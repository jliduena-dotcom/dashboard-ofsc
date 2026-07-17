#!/usr/bin/env python3
"""
Dashboard OFSC — Generador automático
--------------------------------------
Modos:
  python generar_dashboard.py                              -> genera una vez con datos actuales
  python generar_dashboard.py --refresh                   -> refresca Power Query, genera una vez
  python generar_dashboard.py --refresh --publish         -> refresca, genera y publica en GitHub Pages
  python generar_dashboard.py --refresh --publish --watch -> refresca, publica y queda monitoreando

Flags:
  --refresh   Abre Excel en segundo plano, ejecuta Power Query, guarda y cierra
  --publish   Hace commit y push del HTML a GitHub Pages
  --watch     Queda monitoreando el Excel; regenera y publica cada vez que se guarda

Requisitos:
  pip install pandas openpyxl pyxlsb watchdog xlwings
"""

import pandas as pd
import json, sys, os, time
from datetime import datetime, timedelta
from pathlib import Path

_modo_publish = False  # se activa con flag --publish

# ── CONFIGURACIÓN ──────────────────────────────────────────────────────────────
EXCEL_FILE     = "Repositorio_OFSC_Diario_Avance.xlsx"
RM_FILE        = "RM_Comercial.xlsb"          # puede estar ausente; el dashboard funciona igual
# Sin template separado: el script lee y reescribe dashboard_avance_operativo.html
# Los marcadores <!--DATOS_INICIO--> / <!--DATOS_FIN--> delimitan la zona de datos.
OUTPUT_FILE    = "dashboard_avance_operativo.html"
INTERVALO_MIN  = 30
COMP_TEMPLATE_FILE = "comparativa_diaria_template.html"  # template independiente comparativa
COMP_OUTPUT_FILE   = "comparativa_diaria.html"            # output para GitHub Pages
# Carpeta donde caen las bases que alimentan Power Query del Repositorio
# Power Query en el Repositorio debe apuntar a esta carpeta
BASES_OFSC_DIR = Path(r"D:\OneDrive - Comunicacion Celular S.A.- Comcel S.A"
                     r"\CCOT gestiones\Base Equipo Despacho CCOT"
                     r"\Sinergia Mill-Claro\Bases OFSC")
# ──────────────────────────────────────────────────────────────────────────────

MARKER_A  = "<!--DATOS_INICIO-->"
MARKER_B  = "<!--DATOS_FIN-->"
MARKER_CA = "<!--COMPARATIVA_INICIO-->"
MARKER_CB = "<!--COMPARATIVA_FIN-->"

# Nombres cortos para compañías (se actualiza automáticamente si entra una nueva)
COMP_SHORT_MAP = {
    'ADSM INGENIEROS':                               'ADSM',
    'CLARO COLOMBIA':                                'Claro Colombia',
    'COMCEL CELUMAX JT COMUNICACIONES SAS':          'Celumax JT',
    'COMCEL CELUNORTE COMUNICACIONES S.A.S':         'Celunorte',
    'COMCEL COMTEL':                                 'Comtel',
    'COMCEL COMUNICACIONES DEL SUR':                 'Com. del Sur',
    'COMCEL DANCELL':                                'Dancell',
    'COMCEL INVERSIONES GERA S.A.S.':                'Inv. Gera',
    'COMCEL MAGANET.COM SAS':                        'Maganet',
    'COMCEL MC MULTICELL S.A.S':                     'MC Multicell',
    'COMCEL RYL TELECOMUNICACIONES SAS':             'RYL Telecom',
    'COMCEL SOLUCIONES EN COMUNICACIONES JJL S.A.S.':'Sol. JJL',
    'COMCEL SUPERCEL E.U.':                          'Supercel',
    'COMCEL WINCELL':                                'Wincell',
    'DOMINION':                                      'Dominion',
    'GEISHI TRAVEL SAS':                             'Geishi Travel',
    'INMEL INGENIERIA':                              'Inmel Ing.',
    'INVERSIONES BMS S.A.S':                         'Inv. BMS',
    'TELECELL LORICA SAS':                           'Telecell Lorica',
}


# ── CLASIFICADORES ─────────────────────────────────────────────────────────────

def clasificar_tipo(tipo):
    t = str(tipo).strip()
    if t in ['Instalaciones','INSTALACIONES DTH','INSTALACIONES FTTH',
             'INSTALACIONES FWA','Orden Especial']:
        return 'Instalaciones'
    if t in ['Arreglos','MANTENIMIENTO FTTH','MANTENIMIENTOS DTH']:
        return 'Arreglos'
    if t in ['Blindaje.','CAMBIO EQUIPOS FWA','Post Venta',
             'POSTVENTA  FTTH','POSTEVENTA DTH','POSTEVENTA FTTH']:
        return 'Postventa'
    if t in ['Traslados','TRASLADO  FTTH']:
        return 'Traslados'
    if t in ['Desconexiones']:
        return 'Desconexiones'
    if t in ['Almuerzo','Actividades de Almacen','Capacitacion','Supervision',
             'Recursos Humanos','Pre-Turno','Ausentismo','Vehiculo con Fallas',
             'Apoyo Caso VIP','Supervisión_','Gestión con Administraciones',
             'VENTAS TECNICO']:
        return None
    return 'Otros'


def clasificar_franja(intervalo):
    i = str(intervalo).strip()
    if i in ['07-09','07-10','07-12','07-13','08-13','09-11','10-13','11-13']:
        return 'AM'
    if i in ['13-17','13-18','14-16','14-17','14-18','14-20','16-18']:
        return 'PM'
    try:
        return 'AM' if int(i.split('-')[0]) < 13 else 'PM'
    except Exception:
        return 'Otro'


# ── COMPARATIVA DIARIA ────────────────────────────────────────────────────────

TIPOS_EXCL_COMP = {
    'Almuerzo','Actividades de Almacen','Capacitacion','Supervision',
    'Recursos Humanos','Pre-Turno','Ausentismo','Vehiculo con Fallas',
    'Apoyo Caso VIP','Supervisión_','Gestión con Administraciones','VENTAS TECNICO'
}
ESTADOS_GESTIONADOS = {'Completado','No completado','Cancelado','Iniciado'}
ESTADOS_OK          = {'Completado'}
HORAS_COMP          = list(range(7, 23))   # 7h a 22h


def calcular_comparativa(df_raw: 'pd.DataFrame') -> dict:
    """
    Recibe el DataFrame crudo del Excel (antes de filtrar por franja/grupo)
    y devuelve el diccionario CD_PRECOMP listo para inyectar en el HTML.

    Estructura por fecha:
      {
        "YYYY-MM-DD": {
          total   : int,          # total órdenes operativas del día
          ok      : int,          # órdenes Completadas
          acum    : {h: pct},     # avance acumulado % hasta la hora h
          ritmo   : {h: pct},     # % gestionado EN esa hora exacta
          am      : float,        # % del total gestionado en franja AM (7-13h)
          pm      : float,        # % del total gestionado en franja PM (14-21h)
          efect   : float,        # Completado/(Completado+NoCompletado) × 100
          atipico : bool,         # True si total < UMBRAL_ATIPICO
          parcial : bool          # True si el día es hoy y aún hay horas sin datos
        }
      }
    También devuelve un 'benchmark' calculado sobre los días completos normales.
    """
    import datetime as dt

    df = df_raw.copy()

    # ── 1. Filtrar órdenes operativas ────────────────────────────────
    df = df[df['Tipo de Actividad'].notna()].copy()
    df = df[~df['Tipo de Actividad'].isin(TIPOS_EXCL_COMP)].copy()
    if df.empty:
        return {}

    df['gestionado'] = df['Estado'].isin(ESTADOS_GESTIONADOS)
    df['ok_estado']  = df['Estado'].isin(ESTADOS_OK)
    df['Fecha_str']  = pd.to_datetime(df['Fecha'], errors='coerce').dt.strftime('%Y-%m-%d')
    df = df[df['Fecha_str'].notna()].copy()

    # ── 2. Extraer hora real de fin ──────────────────────────────────
    def get_hour(v):
        if isinstance(v, dt.time):
            return v.hour
        return None

    df['Fin_hora_int'] = df['Fin'].apply(get_hour)   # int o None, nunca float NaN
    df['Fin_obj']      = df['Fin'].apply(lambda x: x if isinstance(x, dt.time) else None)

    # ── 3. Calcular por fecha ────────────────────────────────────────
    UMBRAL_ATIPICO  = 500    # días con menos órdenes se marcan atípicos
    hoy_str = dt.date.today().strftime('%Y-%m-%d')

    fechas_disponibles = sorted(df['Fecha_str'].dropna().unique().tolist())
    resultado = {}

    for fecha in fechas_disponibles:
        sub = df[df['Fecha_str'] == fecha].copy()
        total = len(sub)
        if total == 0:
            continue

        ok_cnt  = int(sub['ok_estado'].sum())
        comp    = int(sub['Estado'].isin({'Completado'}).sum())
        no_comp = int(sub['Estado'].isin({'No completado'}).sum())
        denom_efect = comp + no_comp
        efect = round(comp / denom_efect * 100, 1) if denom_efect > 0 else 0.0

        # Avance acumulado y ritmo por hora
        acum  = {}
        ritmo = {}
        for h in HORAS_COMP:
            lim = dt.time(h, 59, 59)
            gest_hasta_h = sub['gestionado'] & sub['Fin_obj'].apply(
                lambda x: x is not None and x <= lim)
            gest_en_h = sub['gestionado'] & sub['Fin_hora_int'].apply(
                lambda x: x is not None and x == h)
            acum[h]  = round(float(gest_hasta_h.sum()) / total * 100, 1)
            ritmo[h] = round(float(gest_en_h.sum())    / total * 100, 1)

        # Balance AM (7–13h) / PM (14–21h)
        am_gest = sub['gestionado'] & sub['Fin_hora_int'].apply(
            lambda x: x is not None and 7 <= x <= 13)
        pm_gest = sub['gestionado'] & sub['Fin_hora_int'].apply(
            lambda x: x is not None and 14 <= x <= 21)
        am_pct = round(float(am_gest.sum()) / total * 100, 1)
        pm_pct = round(float(pm_gest.sum()) / total * 100, 1)

        atipico = total < UMBRAL_ATIPICO
        parcial = (fecha == hoy_str) and (acum.get(18, 0) < 70)

        resultado[fecha] = {
            'total':   total,
            'ok':      ok_cnt,
            'acum':    acum,
            'ritmo':   ritmo,
            'am':      am_pct,
            'pm':      pm_pct,
            'efect':   efect,
            'atipico': atipico,
            'parcial': parcial,
        }

    # ── 4. Benchmark: promedio de días completos normales ────────────
    dias_bench = [f for f, v in resultado.items()
                  if not v['atipico'] and not v['parcial']]
    bench_acum  = {}
    bench_ritmo = {}
    if dias_bench:
        for h in HORAS_COMP:
            vals_a = [resultado[f]['acum'][h]  for f in dias_bench]
            vals_r = [resultado[f]['ritmo'][h] for f in dias_bench]
            bench_acum[h]  = round(sum(vals_a) / len(vals_a), 1)
            bench_ritmo[h] = round(sum(vals_r) / len(vals_r), 1)

    return {
        'dias':         resultado,
        'benchmark':    bench_acum,
        'bench_ritmo':  bench_ritmo,
        'dias_bench':   dias_bench,
    }


# ── GENERADOR ─────────────────────────────────────────────────────────────────

def generar():
    base        = Path(__file__).resolve().parent
    excel_path  = base / EXCEL_FILE
    rm_path     = base / RM_FILE
    output_path = base / OUTPUT_FILE

    ts = datetime.now().strftime('%H:%M:%S')
    print(f"\n[{ts}] Procesando {EXCEL_FILE}...")

    # El dashboard_avance_operativo.html actúa como template y output a la vez.
    # Los marcadores <!--DATOS_INICIO--> / <!--DATOS_FIN--> delimitan la zona
    # de datos generados; todo lo demás (HTML/CSS/JS) permanece intacto.
    for p, label in [(excel_path, EXCEL_FILE), (output_path, OUTPUT_FILE)]:
        if not p.exists():
            print(f"[ERROR] No encontré: {p}")
            return False

    with open(output_path, encoding='utf-8') as f:
        tmpl = f.read()
    if MARKER_A not in tmpl or MARKER_B not in tmpl:
        print("[ERROR] dashboard_avance_operativo.html no tiene los marcadores DATOS.")
        print("        Asegúrate de usar la versión refactorizada del dashboard.")
        return False

    tiene_comp = MARKER_CA in tmpl and MARKER_CB in tmpl

    html_before = tmpl[:tmpl.index(MARKER_A) + len(MARKER_A)]
    html_after  = tmpl[tmpl.index(MARKER_B):]

    # ── Leer OFSC ──────────────────────────────────────────────────
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        print(f"[ERROR] No pude leer el Excel: {e}")
        return False

    if 'Razón' in df.columns:
        df = df.rename(columns={'Razón': 'Razon'})

    cols_req = ['Tipo de Actividad','Intervalos de tiempo','Estado',
                'Compañia','Ciudad','Tipo de Red','Razon','Fecha',
                'Código Asesor comercial']
    faltantes = [c for c in cols_req if c not in df.columns]
    if faltantes:
        print(f"[ERROR] Columnas faltantes en OFSC: {faltantes}")
        return False

    df = df.copy()
    for c in ['Tipo de Actividad', 'Intervalos de tiempo', 'Estado',
              'Compañia', 'Ciudad', 'Tipo de Red', 'Razon',
              'Código Asesor comercial']:
        if c in df.columns:
            df[c] = df[c].astype('string').str.strip()

    df['grupo']  = df['Tipo de Actividad'].apply(clasificar_tipo)
    df = df[df['grupo'].notna()].copy()
    df['franja'] = df['Intervalos de tiempo'].apply(clasificar_franja)
    df = df[df['franja'].isin(['AM','PM'])].copy()
    df['gestionado'] = df['Estado'].isin(['Completado','No completado','Cancelado','Iniciado'])
    # Subtipo de orden: Brownfield vs Resto
    BROWNFIELD_VALS = {'BROWNFIELD','BROWNFIELD PYMES','BROWNFIELD FLASH',
                       'Brownfield Pymes','Brownfield Flash'}
    def cls_subtipo(s):
        if pd.isna(s): return 'Resto de ordenes'
        return 'Brownfield' if str(s).strip().upper() in {v.upper() for v in BROWNFIELD_VALS} else 'Resto de ordenes'
    if 'Subtipo de la Orden de Trabajo' in df.columns:
        df['subtipo_grp'] = df['Subtipo de la Orden de Trabajo'].apply(cls_subtipo)
    else:
        df['subtipo_grp'] = 'Resto de ordenes'

    # Técnico limpio (quitar prefijo EMPRESA_ si existe)
    if 'Técnico' in df.columns:
        df['tec_key'] = df['Técnico'].astype(str).str.strip()
        # Quitar prefijo tipo "MILLENIUM_" o "CLARO_"
        df['tec_name'] = df['tec_key'].str.replace(r'^[A-Z]+_', '', regex=True).str.strip()
    else:
        df['tec_key'] = ''
        df['tec_name'] = ''

    # Red Neutra: tiene valor en columna Nap
    if 'Nap' in df.columns:
        df['nap_rn'] = df['Nap'].notna().astype(int)
    else:
        df['nap_rn'] = 0
    # Convertir Fecha a string (evita error Timestamp not JSON serializable)
    df['Fecha'] = df['Fecha'].astype(str).str.strip()

    if df.empty:
        print("[AVISO] No hay registros operativos válidos.")
        return False

    # ── Cruce con RM Comercial ──────────────────────────────────────
    canal_col = 'CANAL2'
    df['canal_val'] = ''

    if rm_path.exists():
        try:
            df_rm = pd.read_excel(rm_path, engine='pyxlsb',
                                  usecols=['TCARGU', canal_col])
            # Normalizar: quitar ceros a la izquierda para match
            df_rm['tcargu_key'] = df_rm['TCARGU'].astype(str).str.strip().str.lstrip('0')
            df['asesor_key'] = df['Código Asesor comercial'].astype(str).str.strip().str.lstrip('0')

            canal_dict = df_rm.set_index('tcargu_key')[canal_col].to_dict()
            df['canal_val'] = df['asesor_key'].map(canal_dict).fillna('')
            matched = df['canal_val'].notna() & (df['canal_val'] != '')
            print(f"[{ts}] RM cruzado: {matched.sum()} registros con canal asignado")
        except Exception as e:
            print(f"[AVISO] No pude leer RM_Comercial: {e} — canal desactivado")
    else:
        print(f"[AVISO] {RM_FILE} no encontrado — canal desactivado")

    # ── Catálogos ──────────────────────────────────────────────────
    cities  = sorted(df['Ciudad'].dropna().unique().tolist())
    comps   = sorted(df['Compañia'].dropna().unique().tolist())
    reds    = sorted(df['Tipo de Red'].dropna().unique().tolist())
    estados = sorted(df['Estado'].dropna().unique().tolist())
    grupos  = sorted(df['grupo'].dropna().unique().tolist())
    razones = sorted(df['Razon'].dropna().unique().tolist())
    fechas   = sorted(df['Fecha'].dropna().astype(str).str.strip().unique().tolist())
    canales  = sorted([c for c in df['canal_val'].dropna().unique().tolist() if c])
    subtipos = ['Brownfield', 'Resto de ordenes']
    tecnicos = sorted(df['tec_key'].dropna().unique().tolist())
    tec_names= {t: df[df['tec_key']==t]['tec_name'].iloc[0] if len(df[df['tec_key']==t])>0 else t
                for t in tecnicos}

    # Nombres cortos de compañías (dinámico)
    comp_short = [COMP_SHORT_MAP.get(c, c[:14]) for c in comps]

    city_map   = {c: i for i, c in enumerate(cities)}
    comp_map   = {c: i for i, c in enumerate(comps)}
    red_map    = {c: i for i, c in enumerate(reds)}
    estado_map = {c: i for i, c in enumerate(estados)}
    grupo_map  = {c: i for i, c in enumerate(grupos)}
    razon_map  = {c: i for i, c in enumerate(razones)}
    fecha_map  = {c: i for i, c in enumerate(fechas)}
    canal_map   = {c: i for i, c in enumerate(canales)}
    subtipo_map = {c: i for i, c in enumerate(subtipos)}
    tec_map     = {c: i for i, c in enumerate(tecnicos)}

    # ── Codificar filas ────────────────────────────────────────────
    # [ciudad(0),comp(1),franja(2),grupo(3),estado(4),red(5),razon(6),gestionado(7),fecha_idx(8),canal_idx(9),nap_rn(10),subtipo_idx(11),tec_idx(12)]
    encoded = []
    for _, r in df.iterrows():
        encoded.append([
            city_map.get(r['Ciudad'],    -1) if pd.notna(r['Ciudad'])           else -1,
            comp_map.get(r['Compañia'],  -1) if pd.notna(r['Compañia'])         else -1,
            0 if r['franja'] == 'AM' else 1,
            grupo_map.get(r['grupo'],    -1),
            estado_map.get(r['Estado'],  -1) if pd.notna(r['Estado'])           else -1,
            red_map.get(r['Tipo de Red'],-1) if pd.notna(r['Tipo de Red'])      else -1,
            razon_map.get(r['Razon'],    -1) if pd.notna(r['Razon'])            else -1,
            1 if r['gestionado'] else 0,
            fecha_map.get(str(r['Fecha']).strip(), -1) if pd.notna(r['Fecha'])  else -1,
            canal_map.get(r['canal_val'],-1) if r['canal_val']                  else -1,
            int(r['nap_rn']),
            subtipo_map.get(r['subtipo_grp'], -1),
            tec_map.get(r['tec_key'], -1),
        ])

    raw = dict(
        cities=cities, comps=comps, comp_short=comp_short,
        reds=reds, estados=estados, grupos=grupos, razones=razones,
        fechas=fechas, canales=canales,
        subtipos=subtipos,
        tecnicos=tecnicos,
        tec_names=tec_names
    )
    raw_js  = json.dumps(raw,     ensure_ascii=False, separators=(',',':'))
    rows_js = json.dumps(encoded, ensure_ascii=False, separators=(',',':'))

    datos_block = (
        "<script>\n"
        f"const RAW = {raw_js};\n"
        "</script>\n"
        "<script>\n"
        f"const ROWS_DATA = {rows_js};\n"
        "RAW.rows = ROWS_DATA;\n"
        "</script>\n"
    )

    comp_template_path = base / COMP_TEMPLATE_FILE
    comp_output_path   = base / COMP_OUTPUT_FILE

    # ── Bloque comparativa diaria ──────────────────────────────────
    comp_block = ""
    if tiene_comp or comp_template_path.exists():
        try:
            df_raw_full = pd.read_excel(excel_path)
            comp_data   = calcular_comparativa(df_raw_full)

            if comp_data and comp_data.get('dias'):
                dias_js       = json.dumps(comp_data['dias'],        ensure_ascii=False, separators=(',',':'))
                bench_js      = json.dumps(comp_data['benchmark'],   ensure_ascii=False, separators=(',',':'))
                bench_ritmo_js= json.dumps(comp_data['bench_ritmo'], ensure_ascii=False, separators=(',',':'))
                n_dias        = len(comp_data['dias'])
                n_bench       = len(comp_data['dias_bench'])

                comp_block = (
                    "\n<script>\n"
                    f"// Comparativa Diaria — generado {ts}\n"
                    f"CD_PRECOMP = {dias_js};\n"
                    f"const CD_BENCH_GEN = {bench_js};\n"
                    f"const CD_BENCH_RITMO_GEN = {bench_ritmo_js};\n"
                    "// Sobreescribir constantes estáticas del template con datos frescos\n"
                    "Object.assign(CD_BENCH, CD_BENCH_GEN);\n"
                    "Object.assign(CD_BENCH_RITMO, CD_BENCH_RITMO_GEN);\n"
                    "Object.keys(CD_PRECOMP).forEach(k => { const d = CD_PRECOMP[k]; "
                    "if (d.acum && typeof d.acum === 'object') { "
                    "const na = {}; Object.keys(d.acum).forEach(h => na[parseInt(h)] = d.acum[h]); d.acum = na; "
                    "const nr = {}; Object.keys(d.ritmo).forEach(h => nr[parseInt(h)] = d.ritmo[h]); d.ritmo = nr; "
                    "} });\n"
                    "</script>"
                )
                print(f"[{ts}] Comparativa: {n_dias} días calculados "
                      f"({n_bench} para benchmark)")
            else:
                print(f"[AVISO] Comparativa: sin datos válidos, se usarán valores estáticos del template")
        except Exception as e:
            print(f"[AVISO] No pude calcular comparativa: {e} — se usarán valores estáticos")

    # ── Ensamblar HTML final ───────────────────────────────────────
    if tiene_comp and comp_block:
        html_comp_before = tmpl[:tmpl.index(MARKER_CA) + len(MARKER_CA)]
        html_comp_after  = tmpl[tmpl.index(MARKER_CB):]
        # Reemplazar DATOS primero, luego COMPARATIVA
        html_tmp = html_before + datos_block + html_after
        # Ahora reemplazar bloque comparativa en el html_tmp resultante
        if MARKER_CA in html_tmp and MARKER_CB in html_tmp:
            part1 = html_tmp[:html_tmp.index(MARKER_CA) + len(MARKER_CA)]
            part2 = html_tmp[html_tmp.index(MARKER_CB):]
            html_final = part1 + comp_block + part2
        else:
            html_final = html_tmp
    else:
        html_final = html_before + datos_block + html_after
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_final)

    # ── Generar también la Comparativa Diaria independiente ──
    if comp_template_path.exists():
        try:
            with open(comp_template_path, encoding='utf-8') as f:
                comp_tmpl = f.read()
            if MARKER_CA in comp_tmpl and MARKER_CB in comp_tmpl:
                cp1 = comp_tmpl[:comp_tmpl.index(MARKER_CA) + len(MARKER_CA)]
                cp2 = comp_tmpl[comp_tmpl.index(MARKER_CB):]
                comp_html = cp1 + (comp_block if comp_block else '\n') + cp2
                with open(comp_output_path, 'w', encoding='utf-8') as f:
                    f.write(comp_html)
                print(f"[{ts}] OK — comparativa_diaria.html generado")
        except Exception as e:
            print(f"[AVISO] No pude generar comparativa_diaria.html: {e}")
    else:
        print(f"[AVISO] {COMP_TEMPLATE_FILE} no encontrado — omitiendo comparativa independiente")

    proxima = (datetime.now() + timedelta(minutes=INTERVALO_MIN)).strftime('%H:%M')
    print(f"[{ts}] OK — {len(encoded)} órdenes, {len(fechas)} fecha(s), "
          f"{len(canales)} canales → {output_path.name}  |  próxima: {proxima}h")
    return True


# ── PUBLICAR EN GITHUB PAGES ──────────────────────────────────────────────────

def publicar():
    """Hace commit y push del HTML generado a GitHub Pages."""
    import subprocess

    base = Path(__file__).resolve().parent
    html_file = str(base / OUTPUT_FILE)
    ts_commit = datetime.now().strftime('%d/%m/%Y %H:%M')

    def run(cmd, check=True):
        result = subprocess.run(cmd, shell=True, cwd=str(base),
                                capture_output=True, text=True)
        if check and result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or result.stdout.strip())
        return result.stdout.strip()

    try:
        # Verificar que estamos en un repo git
        run('git status')

        # ── NUEVO: detectar la rama actual automáticamente ──────────────────
        rama = run('git rev-parse --abbrev-ref HEAD')
        # ────────────────────────────────────────────────────────────────────

        # Añadir solo el HTML (nunca los Excel)
        run(f'git add {OUTPUT_FILE}')

        # Ver si hay cambios staged para commitear.
        # IMPORTANTE: --porcelain incluye archivos untracked con "??" y
        # produce falsos positivos. git diff --cached solo muestra lo
        # que está en el índice listo para commit — es lo correcto aquí.
        staged = run('git diff --cached --name-only', check=False)
        if not staged:
            print(f"[OFSC] Sin cambios para publicar.")
            return True

        run(f'git commit -m "Dashboard actualizado {ts_commit}"')

        # ── CORREGIDO: usa la rama detectada, no 'main' hardcodeado ─────────
        run(f'git push origin {rama}')
        # ────────────────────────────────────────────────────────────────────

        ts = datetime.now().strftime('%H:%M:%S')
        print(f"[{ts}] ✓ Publicado en GitHub Pages")
        return True

    except RuntimeError as e:
        ts = datetime.now().strftime('%H:%M:%S')
        print(f"[{ts}] ERROR al publicar: {e}")
        print("      Verifica que git esté configurado y tengas conexión.")
        return False
    except FileNotFoundError:
        print("[ERROR] Git no encontrado. Instálalo desde: https://git-scm.com")
        return False


# ── REFRESCAR EXCEL (Power Query automático) ───────────────────────────────────

def refrescar_excel():
    """
    Abre el Excel con xlwings, refresca todas las conexiones de Power Query,
    espera a que terminen, guarda y cierra. Sin intervención del usuario.

    Requiere: pip install xlwings
    Solo funciona en Windows con Excel instalado.
    """
    try:
        import xlwings as xw
    except Exception as e:
        import sys
        print(f"[ERROR] No pude importar xlwings: {type(e).__name__}: {e}")
        print(f"        Python usado: {sys.executable}")
        print("        Asegúrate de ejecutar el script con el Python que tiene xlwings/pywin32 instalados.")
        print("        Instala xlwings y pywin32 en ese entorno: pip install xlwings pywin32")
        if 'win32com' in str(e) or 'COM support' in str(e):
            print("        Si el error menciona win32com, reinstala pywin32: python -m pip install --upgrade pywin32")
        return False

    base       = Path(__file__).resolve().parent
    excel_path = base / EXCEL_FILE

    if not excel_path.exists():
        print(f"[ERROR] No encontré: {excel_path}")
        return False

    ts = datetime.now().strftime('%H:%M:%S')
    print(f"[{ts}] Abriendo Excel para refrescar Power Query...")

    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        wb = app.books.open(str(excel_path))

        ts = datetime.now().strftime('%H:%M:%S')
        print(f"[{ts}] Refrescando conexiones Power Query...")

        wb.api.RefreshAll()

        # Esperar a que todas las consultas terminen
        timeout = 120
        inicio  = time.time()
        while True:
            time.sleep(3)
            ocupadas = 0
            try:
                for conn in wb.api.Connections:
                    if hasattr(conn, 'OLEDBConnection'):
                        if conn.OLEDBConnection.Refreshing:
                            ocupadas += 1
                    elif hasattr(conn, 'ODBCConnection'):
                        if conn.ODBCConnection.Refreshing:
                            ocupadas += 1
            except Exception:
                pass

            if ocupadas == 0:
                break
            if time.time() - inicio > timeout:
                print(f"[AVISO] Timeout esperando Power Query ({timeout}s) — guardando igual")
                break

            restante = int(timeout - (time.time() - inicio))
            print(f"\r  Esperando Power Query... ({restante}s restantes)   ",
                  end='', flush=True)

        print()

        wb.save()
        wb.close()

        ts = datetime.now().strftime('%H:%M:%S')
        print(f"[{ts}] Excel actualizado y guardado correctamente")
        return True

    except Exception as e:
        print(f"[ERROR] No pude refrescar Excel: {e}")
        return False

    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass


# ── MODO WATCH ─────────────────────────────────────────────────────────────────

def watch():
    try:
        from watchdog.observers import Observer
        from watchdog.events import FileSystemEventHandler
    except ImportError:
        print("[ERROR] Instala watchdog:  pip install watchdog")
        sys.exit(1)

    base = Path(__file__).resolve().parent

    class Handler(FileSystemEventHandler):
        def __init__(self):
            self._ultimo = 0
            self._busy = False
        def on_modified(self, event):
            nombre = Path(event.src_path).name
            if nombre not in (EXCEL_FILE, RM_FILE):
                return
            ahora = time.time()
            if ahora - self._ultimo < 2:
                return
            self._ultimo = ahora
            time.sleep(1.2)
            if self._busy:
                return
            self._busy = True
            try:
                ts = datetime.now().strftime('%H:%M:%S')
                print(f"\n[{ts}] Cambio detectado en '{nombre}' — refrescando y regenerando...")
                # Al detectar cambios, abrimos el Excel del repositorio,
                # refrescamos Power Query, guardamos y cerramos, luego generamos.
                ok_refresh = refrescar_excel()
                if not ok_refresh:
                    print(f"[{ts}] [AVISO] Falló refresco automático — se intentará generar con datos actuales")

                ok = generar()
                if ok and '--publish' in sys.argv:
                    publicar()
            finally:
                # pequeña pausa para evitar eventos en cascada
                time.sleep(0.5)
                self._busy = False

    # Preferir monitorear la carpeta BASES_OFSC_DIR (entradas manuales).
    # Si no existe, usar el folder del script como fallback (monitorea el Excel).
    obs = Observer()
    if BASES_OFSC_DIR.exists():
        class BasesHandler(FileSystemEventHandler):
            def __init__(self):
                self._ultimo = 0
                self._busy = False
            def on_created(self, event):
                self._on_event(event)
            def on_modified(self, event):
                self._on_event(event)
            def on_moved(self, event):
                self._on_event(event)
            def _on_event(self, event):
                if Path(event.src_path).is_dir():
                    return
                ahora = time.time()
                if ahora - self._ultimo < 2:
                    return
                self._ultimo = ahora
                time.sleep(1.0)
                if self._busy:
                    return
                self._busy = True
                try:
                    ts = datetime.now().strftime('%H:%M:%S')
                    print(f"\n[{ts}] Cambio detectado en Bases OFSC ('{Path(event.src_path).name}') — refrescando repositorio...")
                    ok_refresh = refrescar_excel()
                    if not ok_refresh:
                        print(f"[{ts}] [AVISO] Falló refresco automático tras cambio en Bases OFSC — se intentará generar con datos actuales")
                    ok = generar()
                    if ok and '--publish' in sys.argv:
                        publicar()
                finally:
                    time.sleep(0.5)
                    self._busy = False

        obs.schedule(BasesHandler(), str(BASES_OFSC_DIR), recursive=True)
        obs.start()
        print(f"[OFSC] Watch activo — monitoreando carpeta Bases: {BASES_OFSC_DIR}")
    else:
        # Fallback: monitorizar el Excel en el directorio del script
        class RepoHandler(FileSystemEventHandler):
            def __init__(self):
                self._ultimo = 0
                self._busy = False
            def on_modified(self, event):
                nombre = Path(event.src_path).name
                if nombre not in (EXCEL_FILE, RM_FILE):
                    return
                ahora = time.time()
                if ahora - self._ultimo < 2:
                    return
                self._ultimo = ahora
                time.sleep(1.2)
                if self._busy:
                    return
                self._busy = True
                try:
                    ts = datetime.now().strftime('%H:%M:%S')
                    print(f"\n[{ts}] Cambio detectado en '{nombre}' — refrescando y regenerando...")
                    ok_refresh = refrescar_excel()
                    if not ok_refresh:
                        print(f"[{ts}] [AVISO] Falló refresco automático — se intentará generar con datos actuales")

                    ok = generar()
                    if ok and '--publish' in sys.argv:
                        publicar()
                finally:
                    time.sleep(0.5)
                    self._busy = False

        obs.schedule(RepoHandler(), str(base), recursive=False)
        obs.start()
        print(f"[OFSC] Watch activo — monitoreando '{EXCEL_FILE}' y '{RM_FILE}' (fallback)")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        obs.stop()
    obs.join()


# ── ENTRADA ────────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    modo_watch    = '--watch'   in sys.argv
    modo_refresh  = '--refresh' in sys.argv  # refresca Power Query antes de generar
    _modo_publish = '--publish' in sys.argv

    # 1. Refrescar Excel (Power Query) si se pidió
    if modo_refresh:
        ok_refresh = refrescar_excel()
        if not ok_refresh:
            print("[AVISO] Fallo el refresco de Excel — se intentará generar con datos actuales")

    # 2. Generar el dashboard HTML
    ok = generar()

    # 3. Publicar en GitHub Pages si se pidió
    if ok and _modo_publish:
        publicar()

    # 4. Modo watch: monitorear cambios futuros en el Excel
    if modo_watch:
        print()
        watch()
