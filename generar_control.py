#!/usr/bin/env python3
"""Genera index.html del panel "Control semanal DCE" a partir de BASE_DCE_TIRO.xlsx.

Uso: python3 generar_control.py <BASE_DCE_TIRO.xlsx> <index.html> [AAAA-MM-DD fecha de corte]

Todo el histórico se recalcula desde la base en cada corrida (no hay estado que
mantener): cada semana lunes–viernes que tenga jornadas cargadas aparece en el
selector; la semana "actual" es la que contiene la fecha de corte.
"""
import sys, json, unicodedata, collections, datetime as dt
from pathlib import Path
import openpyxl

BASE = Path(sys.argv[1])
OUT = Path(sys.argv[2])
HOY = dt.date.fromisoformat(sys.argv[3]) if len(sys.argv) > 3 else dt.date.today()

DEP_ALIAS = {
    "DIVISION PATRULLAJE DE EMERGENCIAS": "DIVISION PATRULLAJE DE EMERGENCIA",
    "ANILLO DIGITAL": "DIVISION ANILLO DIGITAL",
}
JER_ALIAS = {
    "OF.MAYOR": "Oficial Mayor", "OF MAYOR": "Oficial Mayor", "OF.PRIMERO": "Oficial Primero", "OF PRIMERO": "Oficial Primero",
    "OF.": "Oficial", "SUBCOMISARIO": "Subcomisario", "COMISARIO": "Comisario", "PRINCIPAL": "Principal", "INSPECTOR": "Inspector",
}
def norm_dep(s):
    s = (s or "").strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    s = " ".join(s.split())
    return DEP_ALIAS.get(s, s)

def norm_jer(s):
    s = " ".join((s or "").strip().upper().split())
    return JER_ALIAS.get(s, cap(s))

def cap(s):
    return " ".join(w.capitalize() for w in (s or "").strip().split())

wb = openpyxl.load_workbook(BASE, data_only=True)
def rows(sheet):
    return [r for r in list(wb[sheet].iter_rows(values_only=True))[3:] if r[0] is not None]

J = rows("JORNADAS")
P = rows("PARTICIPACION")
PI = rows("PART_INSTRUCTORES")
A = rows("ACTIVIDADES")
PER = {str(r[0]).strip(): r for r in rows("PERSONAL")}

def lunes(d):
    return d - dt.timedelta(days=d.weekday())

# --- agrupar jornadas por semana (lunes) ---
semanas = collections.OrderedDict()
for j in sorted(J, key=lambda r: r[1]):
    f = dt.date.fromisoformat(str(j[1])[:10])
    if f.weekday() > 4:      # sábado/domingo: se suma a la semana de ese lunes igual
        pass
    lu = lunes(f)
    act = [a for a in A if a[1] == j[0]]
    ins = sorted({(i[4] or "").strip().upper() for i in PI if i[1] == j[0]})
    parts = [p for p in P if p[1] == j[0] and (p[7] or "PRESENTE").strip().upper() == "PRESENTE"]
    deps = collections.defaultdict(list)
    for p in parts:
        leg = str(p[3] or "").strip()
        per = PER.get(leg)
        deps[norm_dep(p[6])].append({
            "jer": norm_jer(per[3]) if per and per[3] else "",
            "ap": (p[4] or "").strip().upper(),
            "nom": cap(p[5]),
            "leg": leg,
        })
    for d in deps.values():
        d.sort(key=lambda x: (x["ap"], x["nom"]))
    modulos = ", ".join(f"{a[3]} ({a[4]} hs)" if a[4] else str(a[3]) for a in act) or (j[3] or "")
    horas = sum((a[4] or 0) for a in act)
    semanas.setdefault(lu, []).append({
        "id": j[0],
        "fecha": f.isoformat(),
        "lugar": j[2] or "",
        "tipo": j[3] or "",
        "modulos": modulos,
        "horas": horas,
        "aCargo": cap(j[5]),
        "instructores": ins,
        "estado": (j[7] or "").upper(),
        "obs": (j[6] or "").strip(),
        "asistentes": len(parts),
        "dependencias": [{"nombre": k, "personas": v} for k, v in sorted(deps.items(), key=lambda kv: (-len(kv[1]), kv[0]))],
    })

def resumen(lu, jornadas):
    personas = {}
    deps = set()
    ins = set()
    for j in jornadas:
        ins.update(j["instructores"])
        for d in j["dependencias"]:
            deps.add(d["nombre"])
            for p in d["personas"]:
                personas[(p["leg"] or p["ap"] + p["nom"])] = 1
    # listado consolidado por dependencia (una persona una vez por semana)
    cons = collections.defaultdict(dict)
    for j in jornadas:
        for d in j["dependencias"]:
            for p in d["personas"]:
                k = p["leg"] or p["ap"] + p["nom"]
                e = cons[d["nombre"]].setdefault(k, dict(p, fechas=[]))
                e["fechas"].append(j["fecha"])
    listado = [{"nombre": k, "personas": sorted(v.values(), key=lambda x: (x["ap"], x["nom"]))}
               for k, v in sorted(cons.items(), key=lambda kv: (-len(kv[1]), kv[0]))]
    return {
        "lunes": lu.isoformat(),
        "viernes": (lu + dt.timedelta(days=4)).isoformat(),
        "jornadas": jornadas,
        "tot": {"jornadas": len(jornadas), "personas": len(personas), "deps": len(deps), "ins": len(ins),
                "horas": sum(j["horas"] for j in jornadas)},
        "listado": listado,
    }

lista = [resumen(lu, js) for lu, js in semanas.items()]
lu_hoy = lunes(HOY)
if lu_hoy not in semanas:
    lista.append(resumen(lu_hoy, []))
lista.sort(key=lambda s: s["lunes"], reverse=True)

state = {
    "generatedAt": dt.datetime.now().strftime("%Y-%m-%dT%H:%M"),
    "semanaActual": lu_hoy.isoformat(),
    "semanas": lista,
}

tpl = Path(__file__).with_name("plantilla.html").read_text(encoding="utf-8")
html = tpl.replace("/*__STATE__*/", json.dumps(state, ensure_ascii=False))
OUT.write_text(html, encoding="utf-8")
print(f"OK: {len(lista)} semanas, semana actual {lu_hoy}, -> {OUT}")
