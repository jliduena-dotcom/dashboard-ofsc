"""
agente_alertas.py
-----------------
Lee index.html del dashboard OFSC, aplica criterios de alerta
y envía mensajes WhatsApp via CallMeBot.

Configuración: config/criterios.json
Secretos GitHub: WA_NUMEROS, WA_APIKEYS  (coma-separados)
"""

import json
import re
import os
import sys
import requests
from datetime import datetime, timezone, timedelta

# Colombia: UTC-5
TZ_COL = timezone(timedelta(hours=-5))


def ahora_col():
    return datetime.now(TZ_COL)


def cargar_config():
    ruta = os.path.join(os.path.dirname(__file__), "..", "config", "criterios.json")
    with open(ruta, encoding="utf-8") as f:
        return json.load(f)


def extraer_datos_html(ruta):
    """Extrae RAW y ROWS_DATA embebidos como JSON en el HTML."""
    with open(ruta, encoding="utf-8") as f:
        html = f.read()

    decoder = json.JSONDecoder()

    m = re.search(r"const\s+RAW\s*=\s*", html)
    if not m:
        raise ValueError("No se encontró 'const RAW' en el HTML")
    raw, _ = decoder.raw_decode(html, m.end())

    m2 = re.search(r"const\s+ROWS_DATA\s*=\s*", html)
    if not m2:
        raise ValueError("No se encontró 'const ROWS_DATA' en el HTML")
    rows, _ = decoder.raw_decode(html, m2.end())

    return raw, rows


def decodificar_filas(raw, rows):
    """Convierte las filas indexadas en dicts legibles."""
    resultado = []
    for r in rows:
        resultado.append({
            "ciudad":  raw["cities"][r[0]]   if r[0]  is not None else "",
            "comp":    raw["comps"][r[1]]    if r[1]  is not None else "",
            "franja":  r[2] or "",
            "grupo":   raw["grupos"][r[3]]   if r[3]  is not None else "",
            "estado":  raw["estados"][r[4]]  if r[4]  is not None else "",
            "razon":   raw["razones"][r[6]]  if r[6]  is not None else "",
            "fecha":   raw["fechas"][r[8]]   if r[8]  is not None else "",
            "subtipo": raw["subtipos"][r[11]] if r[11] is not None else "",
            "tecnico": raw["tecnicos"][r[12]] if r[12] is not None else "",
        })
    return resultado


def filtrar_hoy(filas, ahora):
    hoy = ahora.strftime("%Y-%m-%d")
    return [f for f in filas if f["fecha"] == hoy]


# ── ANÁLISIS 1: TÉCNICOS RETRASADOS ─────────────────────────────────────────

def analizar_retrasos(filas_hoy, cfg, ahora):
    """Técnico con actividades activas en franja AM pasadas las 13h, o PM pasadas las 18h."""
    alertas = []
    cr = cfg["retraso_franja"]
    estados_activos = set(cr["estados_activos"])
    lim_am = datetime.strptime(cr["AM"]["hora_limite"], "%H:%M").time()
    lim_pm = datetime.strptime(cr["PM"]["hora_limite"], "%H:%M").time()
    hora = ahora.time()

    retrasados = {}
    for f in filas_hoy:
        if f["estado"] not in estados_activos or not f["tecnico"]:
            continue
        franja = f["franja"]
        if "AM" in franja and hora >= lim_am:
            clave = (f["tecnico"], "AM")
        elif "PM" in franja and hora >= lim_pm:
            clave = (f["tecnico"], "PM")
        else:
            continue
        if clave not in retrasados:
            retrasados[clave] = {"actividades": 0, "razones": set()}
        retrasados[clave]["actividades"] += 1
        if f["razon"]:
            retrasados[clave]["razones"].add(f["razon"])

    for (tec, franja), info in retrasados.items():
        razones = ", ".join(info["razones"]) or "Sin razón registrada"
        alertas.append({
            "tipo": "TECNICO_RETRASADO",
            "tecnico": tec,
            "franja": franja,
            "cantidad": info["actividades"],
            "razones": razones,
        })
    return alertas


# ── ANÁLISIS 2: TÉCNICOS SOBRECARGADOS ──────────────────────────────────────

def analizar_sobrecarga(filas_hoy, cfg):
    """
    Arreglos + Postventa: alerta si > max_actividades (default 6)
    Instalaciones + Traslados + Brownfield: alerta si > max_actividades (default 4)
    """
    alertas = []
    cs = cfg["sobrecarga"]
    grupos_ap = set(cs["arreglos_postventa"]["grupos"])
    max_ap = cs["arreglos_postventa"]["max_actividades"]
    grupos_it = set(cs["instalaciones_traslados"]["grupos"])
    max_it = cs["instalaciones_traslados"]["max_actividades"]
    estados_activos = {"Pendiente", "Iniciado", "en ruta"}

    conteo = {}
    for f in filas_hoy:
        if f["estado"] not in estados_activos or not f["tecnico"]:
            continue
        tec = f["tecnico"]
        grupo = f["grupo"]
        if tec not in conteo:
            conteo[tec] = {"ap": 0, "it": 0}
        if grupo in grupos_ap:
            conteo[tec]["ap"] += 1
        elif grupo in grupos_it:
            conteo[tec]["it"] += 1

    for tec, cnt in conteo.items():
        if cnt["ap"] > max_ap:
            alertas.append({
                "tipo": "SOBRECARGA",
                "tecnico": tec,
                "categoria": "Arreglos / Postventa",
                "cantidad": cnt["ap"],
                "limite": max_ap,
            })
        if cnt["it"] > max_it:
            alertas.append({
                "tipo": "SOBRECARGA",
                "tecnico": tec,
                "categoria": "Instalaciones / Traslados / Brownfield",
                "cantidad": cnt["it"],
                "limite": max_it,
            })
    return alertas


# ── ANÁLISIS 3: AVANCE DEL DÍA ──────────────────────────────────────────────

def analizar_avance(filas_hoy, cfg, ahora):
    """Alerta si el % de completadas está por debajo del umbral a la hora de evaluación."""
    ca = cfg.get("avance_dia", {})
    umbral = ca.get("umbral_pct", 80)
    hora_eval = datetime.strptime(ca.get("hora_evaluacion", "15:00"), "%H:%M").time()

    if ahora.time() < hora_eval or not filas_hoy:
        return []

    total = len(filas_hoy)
    completadas = sum(1 for f in filas_hoy if f["estado"] == "Completado")
    pct = round(completadas / total * 100, 1)

    if pct >= umbral:
        return []

    return [{
        "tipo": "AVANCE_RETRASADO",
        "total": total,
        "completadas": completadas,
        "pct": pct,
        "umbral": umbral,
        "hora_eval": ca.get("hora_evaluacion", "15:00"),
    }]


# ── FORMATO MENSAJES ─────────────────────────────────────────────────────────

def formatear(alerta, ahora):
    ts = ahora.strftime("%d/%m/%Y %H:%M")
    tipo = alerta["tipo"]

    if tipo == "TECNICO_RETRASADO":
        return (
            f"⚠️ *ALERTA RETRASO — OFSC*\n"
            f"🕐 {ts}\n\n"
            f"Técnico: *{alerta['tecnico']}*\n"
            f"Franja: {alerta['franja']}\n"
            f"Act. sin completar: *{alerta['cantidad']}*\n"
            f"Razón: {alerta['razones']}"
        )
    if tipo == "SOBRECARGA":
        return (
            f"🔴 *SOBRECARGA TÉCNICO — OFSC*\n"
            f"🕐 {ts}\n\n"
            f"Técnico: *{alerta['tecnico']}*\n"
            f"Categoría: {alerta['categoria']}\n"
            f"Asignadas: *{alerta['cantidad']}* (máx permitido: {alerta['limite']})"
        )
    if tipo == "AVANCE_RETRASADO":
        return (
            f"📊 *AVANCE DÍA RETRASADO — OFSC*\n"
            f"🕐 {ts}\n\n"
            f"Avance actual: *{alerta['pct']}%* (meta: ≥{alerta['umbral']}%)\n"
            f"Completadas: {alerta['completadas']} de {alerta['total']}\n"
            f"Evaluado a las {alerta['hora_eval']}h Colombia"
        )
    return str(alerta)


# ── ENVÍO WHATSAPP (CallMeBot) ───────────────────────────────────────────────

def enviar_whatsapp(numero, apikey, mensaje):
    try:
        r = requests.get(
            "https://api.callmebot.com/whatsapp.php",
            params={"phone": numero, "text": mensaje, "apikey": apikey},
            timeout=15,
        )
        return r.status_code == 200
    except Exception as e:
        print(f"    Error red: {e}")
        return False


# ── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    cfg = cargar_config()
    ahora = ahora_col()

    # Verificar horario de operación
    inicio = datetime.strptime(cfg["whatsapp"]["hora_inicio"], "%H:%M").time()
    fin_h  = datetime.strptime(cfg["whatsapp"]["hora_fin"],   "%H:%M").time()
    if not (inicio <= ahora.time() <= fin_h):
        print(f"⏸ Fuera de horario ({cfg['whatsapp']['hora_inicio']}–{cfg['whatsapp']['hora_fin']} Col)")
        sys.exit(0)

    print(f"🔍 Agente OFSC — {ahora.strftime('%d/%m/%Y %H:%M')} (Colombia)")

    html_path = os.path.join(os.path.dirname(__file__), "..", "index.html")
    raw, rows = extraer_datos_html(html_path)
    filas = decodificar_filas(raw, rows)
    filas_hoy = filtrar_hoy(filas, ahora)

    print(f"   Actividades hoy ({ahora.strftime('%Y-%m-%d')}): {len(filas_hoy)}")
    if not filas_hoy:
        print("   Sin actividades para hoy en el dashboard.")
        sys.exit(0)

    # Aplicar los tres criterios
    alertas = []
    alertas += analizar_retrasos(filas_hoy, cfg, ahora)
    alertas += analizar_sobrecarga(filas_hoy, cfg)
    alertas += analizar_avance(filas_hoy, cfg, ahora)

    if not alertas:
        print("✅ Sin alertas en este momento.")
        sys.exit(0)

    print(f"🚨 Alertas: {len(alertas)}")

    # Destinatarios: desde secrets de GitHub o desde config
    numeros_env = os.environ.get("WA_NUMEROS", "").strip()
    apikeys_env = os.environ.get("WA_APIKEYS", "").strip()

    if numeros_env and apikeys_env:
        numeros = [n.strip() for n in numeros_env.split(",") if n.strip()]
        apikeys = [k.strip() for k in apikeys_env.split(",") if k.strip()]
        destinatarios = [
            {"numero": n, "apikey": k, "nombre": n}
            for n, k in zip(numeros, apikeys)
        ]
    else:
        destinatarios = cfg["whatsapp"].get("destinatarios", [])

    if not destinatarios:
        print("❌ Sin destinatarios configurados. Agrega WA_NUMEROS y WA_APIKEYS en GitHub Secrets.")
        sys.exit(1)

    # Enviar cada alerta a cada destinatario
    for alerta in alertas:
        msg = formatear(alerta, ahora)
        print(f"\n{'─'*50}\n{msg}")
        for dest in destinatarios:
            ok = enviar_whatsapp(dest["numero"], dest["apikey"], msg)
            print(f"  📱 {dest.get('nombre', dest['numero'])}: {'✅ enviado' if ok else '❌ error'}")


if __name__ == "__main__":
    main()
