"""
Convierte la exportación de afiliados de Berrly (data.xlsx) a un CSV
importable en Admidio.

Uso:
    python main.py
"""

import csv

import pandas as pd

INPUT_FILE = "data.xlsx"
OUTPUT_FILE = "data.csv"

CSV_HEADER = [
    "Apellido", "Nombre", "Calle", "Código Postal", "Ciudad", "Pais",
    "Teléfono", "Teléfono 2", "WhatsApp", "E-Mail",
    "Numero de afiliado", "Fecha de Afiliacion", "Tipo de Documento",
    "Documento de Identidad", "Genero", "Fecha de Nacimiento", "Notas",
    "¿Permite el almacenamiento y tratamiento de datos personales?",
]


# ---------------------------------------------------------------------------
# Utilidades de limpieza de datos
# ---------------------------------------------------------------------------

def val(v) -> str:
    """Convierte un valor de pandas a texto limpio, o cadena vacía si es NaN."""
    if pd.isna(v):
        return ""
    return str(v).strip().strip('"')


def telf_val(v) -> str:
    """Normaliza un número de teléfono: sin decimales, prefijo ni espacios."""
    if pd.isna(v):
        return ""
    t = str(v).strip()
    t = t.replace(".0", "")   # quita el decimal que añade pandas
    t = t.replace("+34", "")  # quita prefijo de país
    t = t.replace(" ", "")    # quita espacios
    return t.strip()


def juntar_telf(*campos) -> str:
    """Une varios campos de teléfono no vacíos separados por '/'."""
    partes = [telf_val(c) for c in campos if telf_val(c)]
    return "/".join(partes)


def fecha(v) -> str:
    """Formatea una fecha al estilo d.m.aaaa (vacío si no es una fecha válida)."""
    if pd.isna(v):
        return ""
    try:
        return pd.to_datetime(v).strftime("%#d.%#m.%Y")  # Windows
        # return pd.to_datetime(v).strftime("%-d.%-m.%Y")  # Linux/macOS
    except (ValueError, TypeError):
        return ""


def juntar(*campos, sep=" ") -> str:
    """Une varios campos no vacíos con el separador indicado."""
    return sep.join(p for p in (val(c) for c in campos) if p)


# ---------------------------------------------------------------------------
# Transformación de cada fila del Excel
# ---------------------------------------------------------------------------

def transformar_fila(fila) -> list[str]:
    """Convierte una fila del Excel en la lista de campos del CSV de salida."""
    apellido = val(fila["Apellidos"]).title()
    nombre = val(fila["Nombre"]).title()
    calle = val(fila["Dirección"]).title()
    cp = val(fila["Código postal"])
    ciudad = val(fila["Población"]).title()
    pais = val(fila["Pais"]).title()
    telf = telf_val(fila["Teléfono - Móvil"])
    telf2 = juntar_telf(fila["Teléfono"], fila["Teléfono - Emergencia"])
    whatsapp = "No"
    email = juntar(fila["Email"], fila["Email 2"], fila["Email 3"], sep=", ").lower()
    num_afil = val(fila["NÚM SOCIO"])
    fecha_alta = fecha(fila["Fecha de alta"])
    # "Tipo de documento" aparece dos veces en el Excel; pandas renombra la
    # segunda columna con el sufijo ".1", así que aquí solo se usa la primera.
    tipo_doc = juntar(fila["Tipo de documento"])
    num_doc = val(fila["Número de documento"])
    genero = val(fila["Género"]).capitalize()
    fec_nac = fecha(fila["Cumpleaños"])
    notas = val(fila["OBSERVACIONES"])
    tratamiento = "Si"

    return [
        apellido, nombre, calle, cp, ciudad, pais,
        telf, telf2, whatsapp, email, num_afil, fecha_alta,
        tipo_doc, num_doc, genero, fec_nac, notas, tratamiento,
    ]


# ---------------------------------------------------------------------------
# Programa principal
# ---------------------------------------------------------------------------

def main() -> None:
    df = pd.read_excel(INPUT_FILE)

    with open(OUTPUT_FILE, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_ALL)
        writer.writerow(CSV_HEADER)
        for _, fila in df.iterrows():
            writer.writerow(transformar_fila(fila))


if __name__ == "__main__":
    main()
