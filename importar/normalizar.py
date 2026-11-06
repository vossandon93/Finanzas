#!/usr/bin/env python3
"""
Convierte una cartola (Excel o CSV del banco) al formato único que lee el script.

Lo usa Claude: Vicente le pasa las cartolas y Claude corre este script, revisa el
resultado y deja el CSV en la carpeta de Drive "Finanzas – importar". Las cartolas
y los CSV generados NO se suben a GitHub (ver .gitignore).

Formato de salida (movimientos):
    fecha,banco,producto,descripcion,monto,cuotas,tipo,centro,cuenta,nota
    - fecha: yyyy-mm-dd
    - monto: negativo = sale plata, positivo = entra
    - cuotas: "3/12" si es la cuota 3 de 12 (se detecta sola en la descripción)
    - tipo/centro/cuenta/nota: opcionales (el script los completa con las reglas)

Ejemplos:
    # Cartola con columnas Cargo y Abono separadas
    python3 normalizar.py movimientos cartola.xlsx --banco Santander \
        --producto "Cuenta corriente Santander" \
        --fecha "Fecha" --descripcion "Descripción" --cargo "Cargo" --abono "Abono"

    # Tarjeta: una sola columna Monto donde las compras vienen en positivo
    python3 normalizar.py movimientos cmr.xlsx --banco "Banco Falabella" \
        --producto "CMR Falabella" --fecha "Fecha" --descripcion "Descripción" \
        --monto "Monto" --invertir

    # Estado de cuenta de una tarjeta (deuda, monto a pagar, vencimiento)
    python3 normalizar.py tarjeta --producto "CMR Falabella" --fecha-estado 2026-10-25 \
        --deuda 820000 --pagar 310000 --fecha-pago 2026-11-10 --cupo 2000000

Cuando veamos las cartolas reales de cada banco se agregan aquí sus formatos fijos
(FORMATOS), para no tener que escribir las columnas cada vez.
"""
import argparse
import csv
import re
import sys
import unicodedata
from datetime import date, datetime
from pathlib import Path

import pandas as pd

SALIDA = Path(__file__).resolve().parent / "salida"
COLUMNAS = ["fecha", "banco", "producto", "descripcion", "monto", "cuotas", "tipo", "centro", "cuenta", "nota"]

# Formatos conocidos por banco y producto: nombre -> argumentos por defecto.
# Se completan cuando lleguen las cartolas reales.
FORMATOS = {
    # "santander-cc": {"banco": "Santander", "producto": "Cuenta corriente Santander",
    #                  "fecha": "Fecha", "descripcion": "Descripción", "cargo": "Cargos", "abono": "Abonos"},
}


def sin_tildes(t):
    t = unicodedata.normalize("NFD", str(t))
    return "".join(c for c in t if unicodedata.category(c) != "Mn").strip().lower()


def leer_tabla(archivo, hoja=None):
    """Lee Excel o CSV y busca sola la fila de títulos (la primera que dice 'fecha')."""
    archivo = Path(archivo)
    if archivo.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        crudo = pd.read_excel(archivo, sheet_name=hoja or 0, header=None, dtype=object)
    else:
        texto = archivo.read_bytes()
        for cod in ("utf-8-sig", "latin-1"):
            try:
                contenido = texto.decode(cod)
                break
            except UnicodeDecodeError:
                continue
        sep = ";" if contenido.splitlines()[0].count(";") > contenido.splitlines()[0].count(",") else ","
        crudo = pd.read_csv(archivo, sep=sep, header=None, dtype=object, encoding=cod)
    for i, fila in crudo.iterrows():
        if any(sin_tildes(v).startswith("fecha") for v in fila if pd.notna(v)):
            tabla = crudo.iloc[i + 1:].copy()
            tabla.columns = [str(v).strip() if pd.notna(v) else f"col{j}" for j, v in enumerate(fila)]
            return tabla.dropna(how="all")
    raise SystemExit("No encontré la fila de títulos (ninguna celda dice 'Fecha').")


def columna(tabla, nombre):
    if not nombre:
        return None
    for c in tabla.columns:
        if sin_tildes(c) == sin_tildes(nombre):
            return c
    raise SystemExit(f'No existe la columna "{nombre}". Columnas: {", ".join(map(str, tabla.columns))}')


def a_fecha(v):
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    t = str(v).strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", t)
    if m:
        return date(int(m[1]), int(m[2]), int(m[3])).isoformat()
    m = re.match(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})", t)
    if m:
        anio = int(m[3]) + (2000 if int(m[3]) < 100 else 0)
        return date(anio, int(m[2]), int(m[1])).isoformat()
    return None


def a_monto(v):
    """'$ -12.345' / '12.345,00' / 12345.0 / '(5.000)' -> entero de pesos (o None)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return round(v)
    t = re.sub(r"[\s$]", "", str(v))
    if not t:
        return None
    negativo = t.startswith("-") or t.endswith("-") or (t.startswith("(") and t.endswith(")"))
    t = re.sub(r",\d{1,2}$", "", re.sub(r"[()\-+]", "", t))
    t = re.sub(r"[.,]", "", t)
    if not t.isdigit():
        return None
    return -int(t) if negativo else int(t)


def cuotas_de(texto):
    m = re.search(r"\b(\d{1,2})\s*/\s*(\d{1,2})\b", str(texto))
    if m and 0 < int(m[1]) <= int(m[2]) <= 48 and int(m[2]) > 1:
        return f"{int(m[1])}/{int(m[2])}"
    return ""


def movimientos(a):
    if a.formato:
        for k, v in FORMATOS[a.formato].items():
            if getattr(a, k, None) in (None, False):
                setattr(a, k, v)
    if not (a.banco and a.producto and a.fecha and a.descripcion and (a.monto or a.cargo or a.abono)):
        raise SystemExit("Faltan datos: --banco, --producto, --fecha, --descripcion y --monto (o --cargo/--abono).")
    tabla = leer_tabla(a.archivo, a.hoja)
    cf, cd = columna(tabla, a.fecha), columna(tabla, a.descripcion)
    cm, cc, ca = columna(tabla, a.monto), columna(tabla, a.cargo), columna(tabla, a.abono)
    ccu = columna(tabla, a.cuotas)
    filas, saltadas = [], 0
    for _, r in tabla.iterrows():
        f = a_fecha(r[cf])
        desc = re.sub(r"\s+", " ", str(r[cd] if pd.notna(r[cd]) else "")).strip()
        if cm:
            monto = a_monto(r[cm])
            if monto is not None and a.invertir:
                monto = -monto
        else:
            cargo = a_monto(r[cc]) if cc else None
            abono = a_monto(r[ca]) if ca else None
            monto = (abs(abono) if abono else 0) - (abs(cargo) if cargo else 0) if (cargo or abono) else None
        if not f or monto is None or monto == 0 or not desc:
            saltadas += 1
            continue
        cuotas = cuotas_de(r[ccu]) if ccu else cuotas_de(desc)
        filas.append({"fecha": f, "banco": a.banco, "producto": a.producto, "descripcion": desc,
                      "monto": monto, "cuotas": cuotas})
    if not filas:
        raise SystemExit("No salió ningún movimiento: revisa las columnas.")
    filas.sort(key=lambda x: x["fecha"])
    nombre = a.salida or f"{sin_tildes(a.producto).replace(' ', '-')}_{filas[0]['fecha']}_{filas[-1]['fecha']}.csv"
    escribir(nombre, filas)
    ing = sum(x["monto"] for x in filas if x["monto"] > 0)
    gas = -sum(x["monto"] for x in filas if x["monto"] < 0)
    print(f"✓ {len(filas)} movimientos ({filas[0]['fecha']} a {filas[-1]['fecha']}) · "
          f"entra ${ing:,.0f} · sale ${gas:,.0f} · filas saltadas: {saltadas}".replace(",", "."))


def tarjeta(a):
    fila = {"producto": a.producto, "fecha_estado": a.fecha_estado, "deuda_total": a.deuda,
            "monto_a_pagar": a.pagar or "", "fecha_pago": a.fecha_pago or "", "cupo_total": a.cupo or ""}
    nombre = a.salida or f"estado_{sin_tildes(a.producto).replace(' ', '-')}_{a.fecha_estado}.csv"
    escribir(nombre, [fila], list(fila.keys()))
    print("✓ Estado de cuenta de", a.producto)


def escribir(nombre, filas, columnas=COLUMNAS):
    SALIDA.mkdir(exist_ok=True)
    ruta = SALIDA / nombre
    with open(ruta, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=columnas, extrasaction="ignore")
        w.writeheader()
        for x in filas:
            w.writerow({c: x.get(c, "") for c in columnas})
    print("Archivo:", ruta)


def main():
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)
    m = sub.add_parser("movimientos", help="cartola de cuenta o tarjeta")
    m.add_argument("archivo")
    m.add_argument("--formato", choices=sorted(FORMATOS))
    m.add_argument("--hoja", help="nombre de la hoja del Excel (por defecto la primera)")
    for c in ("banco", "producto", "fecha", "descripcion", "monto", "cargo", "abono", "cuotas", "salida"):
        m.add_argument("--" + c)
    m.add_argument("--invertir", action="store_true", help="la cartola trae las compras en positivo")
    t = sub.add_parser("tarjeta", help="estado de cuenta de una tarjeta")
    t.add_argument("--producto", required=True)
    t.add_argument("--fecha-estado", required=True)
    t.add_argument("--deuda", required=True, type=int)
    t.add_argument("--pagar", type=int)
    t.add_argument("--fecha-pago")
    t.add_argument("--cupo", type=int)
    t.add_argument("--salida")
    a = p.parse_args()
    movimientos(a) if a.cmd == "movimientos" else tarjeta(a)


if __name__ == "__main__":
    sys.exit(main())
